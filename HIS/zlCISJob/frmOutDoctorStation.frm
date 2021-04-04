VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlIDKind.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOutDoctorStation 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "门诊医生工作站"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13650
   Icon            =   "frmOutDoctorStation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8280
   ScaleMode       =   0  'User
   ScaleWidth      =   13650
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picYy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   3960
      ScaleHeight     =   825
      ScaleWidth      =   690
      TabIndex        =   60
      Top             =   6240
      Width           =   690
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   675
         Index           =   4
         Left            =   0
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   0
         Width           =   435
         _Version        =   589884
         _ExtentX        =   767
         _ExtentY        =   1191
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
   End
   Begin VB.PictureBox picHUIZ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   2415
      ScaleHeight     =   825
      ScaleWidth      =   690
      TabIndex        =   58
      Top             =   4875
      Width           =   690
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   675
         Index           =   2
         Left            =   0
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   0
         Width           =   435
         _Version        =   589884
         _ExtentX        =   767
         _ExtentY        =   1191
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
   End
   Begin VB.PictureBox picMore 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDFDFD&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   2010
      ScaleHeight     =   1575
      ScaleWidth      =   11190
      TabIndex        =   51
      Top             =   2700
      Visible         =   0   'False
      Width           =   11190
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   7
         Left            =   900
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Top             =   30
         Width           =   9630
      End
      Begin zl9CISJob.UCPatiVitalSigns UCPatiVitalSigns 
         Height          =   285
         Left            =   90
         TabIndex        =   53
         Top             =   810
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   503
         ControlLock     =   -1  'True
         TextBackColor   =   -2147483633
         LblBackColor    =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         ShowMode        =   0
         Style           =   1
         XDis            =   100
         YDis            =   200
         LabToTxt        =   -90
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘要:"
         Height          =   180
         Index           =   7
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Width           =   450
      End
   End
   Begin VB.PictureBox picHZ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   1245
      ScaleHeight     =   870
      ScaleWidth      =   675
      TabIndex        =   47
      Top             =   2640
      Width           =   675
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   675
         Index           =   0
         Left            =   0
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   0
         Width           =   435
         _Version        =   589884
         _ExtentX        =   767
         _ExtentY        =   1191
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
   End
   Begin VB.PictureBox picJZ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   1425
      ScaleHeight     =   885
      ScaleWidth      =   705
      TabIndex        =   46
      Top             =   4485
      Width           =   705
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   675
         Index           =   1
         Left            =   0
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   0
         Width           =   435
         _Version        =   589884
         _ExtentX        =   767
         _ExtentY        =   1191
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
   End
   Begin VB.PictureBox picMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   540
      ScaleHeight     =   915
      ScaleWidth      =   1335
      TabIndex        =   45
      Top             =   6465
      Width           =   1335
      Begin XtremeReportControl.ReportControl rptNotify 
         Height          =   675
         Left            =   0
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   0
         Width           =   1095
         _Version        =   589884
         _ExtentX        =   1931
         _ExtentY        =   1191
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
   End
   Begin VB.PictureBox picBasisNew 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FDFDFD&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   1275
      Left            =   2670
      ScaleHeight     =   1275
      ScaleWidth      =   11265
      TabIndex        =   20
      Top             =   615
      Width           =   11265
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   8
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   720
         Width           =   1680
      End
      Begin VB.Frame fraBillType 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   9990
         TabIndex        =   43
         Top             =   750
         Width           =   1860
         Begin VB.ComboBox cboBillType 
            BackColor       =   &H00FDFDFD&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   -30
            Width           =   1110
         End
      End
      Begin VB.PictureBox picPatient 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   30
         ScaleHeight     =   780
         ScaleWidth      =   1050
         TabIndex        =   37
         Top             =   195
         Width           =   1050
         Begin VB.Image imgPatient 
            Height          =   705
            Left            =   30
            Picture         =   "frmOutDoctorStation.frx":058A
            Stretch         =   -1  'True
            Top             =   15
            Width           =   975
         End
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   1755
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "测试"
         Top             =   120
         Width           =   1620
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   1
         Left            =   3240
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "男"
         Top             =   165
         Width           =   465
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   2
         Left            =   3795
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "27岁"
         Top             =   165
         Width           =   720
      End
      Begin VB.Frame fraPayType 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   8325
         TabIndex        =   24
         Top             =   180
         Width           =   1860
         Begin VB.ComboBox cboPayType 
            BackColor       =   &H00FDFDFD&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   -30
            Width           =   1845
         End
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   5715
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "1988-11-11"
         Top             =   150
         Width           =   1080
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   5685
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "65546578"
         Top             =   705
         Width           =   1080
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   7905
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "456455885"
         Top             =   705
         Width           =   1080
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "号类:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   8
         Left            =   2040
         TabIndex        =   57
         Top             =   705
         Width           =   495
      End
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "隐藏其它信息↑"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   4
         Left            =   8580
         MouseIcon       =   "frmOutDoctorStation.frx":1454
         MousePointer    =   99  'Custom
         TabIndex        =   55
         ToolTipText     =   "其它信息:摘要和生命体征。"
         Top             =   495
         Width           =   1260
      End
      Begin VB.Line linBillType 
         X1              =   9900
         X2              =   10680
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   6
         Left            =   9060
         TabIndex        =   42
         Top             =   870
         Width           =   495
      End
      Begin VB.Label lblUrg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "急"
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
         Left            =   11175
         TabIndex        =   41
         Top             =   75
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblMore 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "当日多科就诊"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9915
         TabIndex        =   40
         Top             =   660
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblRec 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "记"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   390
         Left            =   10620
         TabIndex        =   39
         Top             =   90
         Width           =   405
      End
      Begin VB.Line linPayType 
         X1              =   8430
         X2              =   9810
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "文件"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   1140
         MouseIcon       =   "frmOutDoctorStation.frx":2DD6
         MousePointer    =   99  'Custom
         TabIndex        =   36
         Top             =   90
         Width           =   360
      End
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "采集"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   1155
         MouseIcon       =   "frmOutDoctorStation.frx":4758
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   465
         Width           =   360
      End
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "清除"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   2
         Left            =   1155
         MouseIcon       =   "frmOutDoctorStation.frx":60DA
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   990
         Width           =   360
      End
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "修改病人基本信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   3
         Left            =   1665
         MouseIcon       =   "frmOutDoctorStation.frx":7A5C
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   990
         Width           =   1440
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   4605
         TabIndex        =   32
         Top             =   150
         Width           =   885
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "付费方式:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   11
         Left            =   7020
         TabIndex        =   31
         Top             =   120
         Width           =   885
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "就诊卡号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   4560
         TabIndex        =   30
         Top             =   705
         Width           =   885
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医保卡号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   6960
         TabIndex        =   29
         Top             =   705
         Width           =   885
      End
   End
   Begin VB.PictureBox picRegist 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2820
      Left            =   3420
      ScaleHeight     =   2820
      ScaleWidth      =   5265
      TabIndex        =   19
      Top             =   2850
      Width           =   5265
      Begin XtremeSuiteControls.TabControl tbcSub 
         Height          =   1875
         Left            =   240
         TabIndex        =   38
         Top             =   75
         Width           =   2580
         _Version        =   589884
         _ExtentX        =   4551
         _ExtentY        =   3307
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picTmpH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   1125
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   16
      Top             =   495
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picTmpH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   750
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4635
      Left            =   0
      ScaleHeight     =   4635
      ScaleWidth      =   1125
      TabIndex        =   8
      Top             =   1905
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox picFind 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   0
         ScaleHeight     =   270
         ScaleWidth      =   495
         TabIndex        =   12
         Top             =   0
         Width           =   495
         Begin VB.Label lblFind 
            Caption         =   "查找:"
            Height          =   255
            Left            =   60
            TabIndex        =   13
            Top             =   30
            Width           =   495
         End
      End
      Begin VB.Frame fraPatiUD 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   150
         MousePointer    =   7  'Size N S
         TabIndex        =   9
         Top             =   3105
         Width           =   6975
      End
      Begin XtremeSuiteControls.TabControl tbcWait 
         Height          =   435
         Left            =   570
         TabIndex        =   10
         Top             =   315
         Width           =   390
         _Version        =   589884
         _ExtentX        =   688
         _ExtentY        =   767
         _StockProps     =   64
      End
      Begin XtremeSuiteControls.TabControl tbcInTreat 
         Height          =   435
         Left            =   555
         TabIndex        =   11
         Top             =   3480
         Width           =   390
         _Version        =   589884
         _ExtentX        =   688
         _ExtentY        =   767
         _StockProps     =   64
      End
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   270
         Left            =   555
         TabIndex        =   14
         Top             =   0
         Width           =   2895
         _ExtentX        =   5106
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
         IDKindStr       =   $"frmOutDoctorStation.frx":93DE
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
   End
   Begin VB.PictureBox picYZ 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   1830
      ScaleHeight     =   1170
      ScaleWidth      =   1380
      TabIndex        =   4
      Top             =   5925
      Width           =   1380
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   675
         Index           =   3
         Left            =   180
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   465
         Width           =   435
         _Version        =   589884
         _ExtentX        =   767
         _ExtentY        =   1191
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CommandButton cmdOtherFilter 
         Caption         =   "更多条件"
         Height          =   300
         Left            =   2400
         TabIndex        =   7
         Top             =   0
         Width           =   1100
      End
      Begin VB.ComboBox cboSelectTime 
         Height          =   300
         Left            =   795
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   15
         Width           =   1230
      End
      Begin VB.Label lblSeeTim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "就诊时间"
         Height          =   180
         Left            =   30
         TabIndex        =   6
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   1485
      ScaleHeight     =   165
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   555
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Timer timRefresh 
      Interval        =   1000
      Left            =   2055
      Top             =   105
   End
   Begin VB.Frame fraRoom 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   9210
      TabIndex        =   1
      Top             =   7545
      Width           =   300
      Begin VB.Label lblRoom 
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   300
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7920
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmOutDoctorStation.frx":94BB
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17145
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "当天接诊20人"
            TextSave        =   "当天接诊20人"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   1843
            MinWidth        =   1843
            Text            =   "诊室闲"
            TextSave        =   "诊室闲"
            Object.ToolTipText     =   "诊室状态(鼠标点击可设置)"
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
   Begin MSComctlLib.ImageList imgPati 
      Left            =   2880
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":9D4D
            Key             =   "候诊"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":A2E7
            Key             =   "就诊"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":A881
            Key             =   "已诊"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":AE1B
            Key             =   "转诊"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":B3B5
            Key             =   "拒绝"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":B94F
            Key             =   "暂停"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":BEE9
            Key             =   "提醒"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tbcRegist 
      Height          =   915
      Left            =   3090
      TabIndex        =   18
      Top             =   1875
      Width           =   6015
      _Version        =   589884
      _ExtentX        =   10610
      _ExtentY        =   1614
      _StockProps     =   64
   End
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   3870
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgLoad 
      Height          =   705
      Left            =   1905
      Picture         =   "frmOutDoctorStation.frx":C23B
      Stretch         =   -1  'True
      Top             =   1635
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgDefual 
      Height          =   705
      Left            =   2250
      Picture         =   "frmOutDoctorStation.frx":D105
      Stretch         =   -1  'True
      Top             =   1110
      Visible         =   0   'False
      Width           =   975
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmOutDoctorStation.frx":DFCF
      Left            =   960
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
      ScaleMode       =   1
   End
End
Attribute VB_Name = "frmOutDoctorStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const COLOR_FREE As Long = &HC000&
Private Const COLOR_BUSY As Long = &HFF&

Private Enum PatiType
    pt候诊 = 0
    pt就诊 = 1
    pt已诊 = 2
    pt转诊 = 3
    pt预约 = 4
    pt回诊 = 5
    pt排队叫号 = 6
End Enum

Private Enum PATI_RPT_LIST
    PATI_RPT候诊 = 0
    PATI_RPT就诊 = 1
    PATI_RPT回诊 = 2
    PATI_RPT已诊 = 3
    PATI_RPT预约 = 4
End Enum

Private Enum m_Ctl_ID    '一定要连续编号
    txtInfo姓名 = 0
    txtInfo性别 = 1
    txtInfo年龄 = 2
    txtInfo出生日期 = 3
    txtInfo就诊卡号 = 4
    txtInfo医保卡号 = 5
    txtInfo费别 = 6
    txtInfo摘要 = 7
    txtInfo付费方式 = 11
    txtInfo号类 = 8

    lblLink文件 = 0
    lblLink采集 = 1
    lblLink清除 = 2
    lblLink修改 = 3
    lblLink显示 = 4

    '在诊:3人，完成:45人，回诊:15人
    t在诊 = 0
    t完成 = 1
    t回诊 = 2
End Enum

Private Enum PATI_COL_候诊
    COL_HZ_标识 = 0
    COL_HZ_门诊号
    COL_HZ_姓名
    COL_HZ_挂号时间
    COL_HZ_性别
    COL_HZ_年龄
    COL_HZ_急
    COL_HZ_复
    COL_HZ_NO
    COL_HZ_社区
    COL_HZ_就诊诊室
    COL_HZ_就诊医生
    COL_HZ_序号
    COL_HZ_分诊时间
    COL_HZ_就诊卡号
    COL_HZ_病人类型
    COL_HZ_转诊状态
    COL_HZ_预约医生
    COL_HZ_预约时间
    COL_HZ_身份证号
    COL_HZ_号类
    COL_HZ_病人科室
    
'隐藏行
    COL_HZ_病人ID
    COL_HZ_发生时间
    COL_HZ_执行部门ID
    COL_HZ_执行人
    COL_HZ_状态 '转诊状态标志
    COL_HZ_IC卡号
    COL_HZ_记录标志
    COL_HZ_执行状态
End Enum

Private Enum PATI_COL_就诊 '就诊列表和回诊列表共用
    COL_JZ_标识 = 0
    COL_JZ_门诊号
    COL_JZ_姓名
    COL_JZ_就诊时间
    COL_JZ_性别
    COL_JZ_年龄
    COL_JZ_急
    COL_JZ_复
    COL_JZ_NO
    COL_JZ_社区
    COL_JZ_就诊卡号
    COL_JZ_病人类型
    COL_JZ_转诊状态
    COL_JZ_传染病
    COL_JZ_号类
    COL_JZ_病人科室
    
'隐藏行
    COL_JZ_病人ID
    COL_JZ_发生时间
    COL_JZ_执行部门ID
    COL_JZ_执行人
    COL_JZ_状态 '转诊状态标志
    COL_JZ_身份证号
    COL_JZ_IC卡号
    COL_JZ_记录标志
End Enum

Private Enum PATI_COL_已诊
    COL_YZ_NO
    COL_YZ_门诊号
    COL_YZ_姓名
    COL_YZ_性别
    COL_YZ_年龄
    COL_YZ_急
    COL_YZ_复
    COL_YZ_社区
    COL_YZ_时间
    COL_YZ_就诊医生
    COL_YZ_就诊卡号
    COL_YZ_病人类型
    COL_YZ_号类
    COL_YZ_病人科室
    COL_YZ_西医诊断
    COL_YZ_中医诊断

'隐藏行
    COL_YZ_病人ID
    COL_YZ_发生时间
    COL_YZ_执行部门ID
    COL_YZ_执行人
    COL_YZ_身份证号
    COL_YZ_IC卡号
    COL_YZ_记录标志
End Enum

Private Enum Msg_Type '消息提醒类别
    m危机值 = 1
    m医嘱安排 = 2
    m处方审查 = 3
    m传染病 = 4
    m备血完成 = 5
    m用血审核 = 6
    m输血反应 = 6
End Enum
 
Private Enum NOTIFYREPORT_COLUMN
    c_图标 = 0
    C_病人ID = 1
    C_No = 2
    c_姓名 = 3
    C_门诊号 = 4
    C_就诊时间 = 5
    C_状态 = 6
    '隐藏列
    C_消息 = 7
    C_序号 = 8
    C_日期 = 9
    C_业务 = 10
    C_挂号Id = 11
    C_Id = 12
End Enum

Private Type PatiInfo
    类型 As PatiType
    门诊号 As String
    挂号ID As Long
    挂号单 As String
    科室ID As Long
    诊室 As String
    社区 As Integer
    社区号 As String
    挂号时间 As Date
    数据转出 As Boolean
    病人ID As Long
    保存人 As String
    是否签名 As Boolean
    性别 As String
    婚姻状况 As String
    民族 As String
    国籍 As String
    区域 As String
    出生地点 As String
    传染病上传 As Long
    家庭地址邮编 As String
    单位邮编 As String
    其他证件 As String
    户口地址 As String
    户口地址邮编 As String
    籍贯  As String
    Email As String
    QQ As String
    复诊 As Integer
    急诊 As Integer
    路径状态 As Integer
End Type

Private Type ty_Queue
    strQueuePrivs As String '排队叫号虚拟模块权限
    str呼叫站点 As String     '呼叫的站点:空为本站点;否则为其他站点
    byt排队叫号模式 As Byte '排队叫号处理模式:1.代表分诊台分诊呼叫或医生主动呼叫;2-先分诊呼叫,再医生呼叫就诊.0-不排队叫号
    int呼叫人数 As Integer  '0-不限制,>0表示限制人数
    bln呼叫含回诊 As Boolean   '呼叫是否含回诊人数
    bln医生主动呼叫 As Boolean  'true:表示医生主动呼叫;False-医生非主动呼叫
    strCurrQueueName As String '当前队列名称
    lngcurr挂号ID As Long '当前挂号ID
End Type
Private mty_Queue As ty_Queue

'已诊过滤条件
Private Type COND_FILTER
    Begin As Date
    End As Date
    科室ID As Long
    医生 As String
    挂号单 As String
    门诊号 As String
    就诊卡 As String
    姓名 As String
End Type
Private mvCondFilter As COND_FILTER

'子窗体对象定义
Private mclsEMR As Object  '新版病历zlRichEMR.clsDockEMR
Private mclsDisease As zlRichEPR.cDockDisease
Private WithEvents mclsDis As zl9Disease.clsDisease
Attribute mclsDis.VB_VarHelpID = -1
Private WithEvents mclsAdvices As zlPublicAdvice.clsDockOutAdvices
Attribute mclsAdvices.VB_VarHelpID = -1
Private WithEvents mclsPath As zlPublicPath.clsDockOutPath
Attribute mclsPath.VB_VarHelpID = -1
Private WithEvents mclsEPRs As zlRichEPR.cDockOutEPRs
Attribute mclsEPRs.VB_VarHelpID = -1
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private WithEvents mobjEPRDoc As zlRichEPR.cEPRDocument
Attribute mobjEPRDoc.VB_VarHelpID = -1
Private WithEvents mobjQueue As zlQueueManage.clsQueueManage
Attribute mobjQueue.VB_VarHelpID = -1
Private WithEvents mobjPati As frmDockPatiInfo
Attribute mobjPati.VB_VarHelpID = -1
Private mfrmView As frmOutDoctorView
Private mclsReg As zlPublicExpense.clsRegist
Private mcolSubForm As Collection
Private mfrmActive As Form
Private mblnShowLeavePati As Boolean
Private mobjPatient As Object
Private mobjMsg As Object '公共部件气泡对象

'参数设置变量
Private mint接诊范围 As Integer '1-本人,2-本诊室,3-本科室
Private mlng接诊科室ID As Long
Private mstr接诊诊室 As String
Private mstr接诊医生 As String
Private mbln要求分诊 As Boolean
Private mintRefresh As Integer '候诊病人刷新间隔(s)
Private mbln自动接诊 As Boolean
Private mlng自动进行 As Long
Private mbln呼叫后接诊 As Boolean
Private mbln危急值弹窗 As Boolean
 
Private mlng接诊控制 As Long '0-不控制 1-禁止 2-提示 问题号:57566
Private mlng提前接收时间 As Long  '当需要对预约号接收进行控制时,该值表明预约号可以提前接收的分钟数 问题号:57566
Private mblnAutoHandle As Boolean '参数："接诊时自动处理完成就诊"；接诊病人时自动处理上一个病人完成就诊或需回诊。
Private mblnUseTYT As Boolean '使用太元通接口
Private mint过敏输入来源 As Integer '医生站的过敏输入来源
Private mintOutPreTime As Integer
Private mbyt本次就诊 As Byte    '记录 【本次就诊】 页签下标值 0-没有就诊一览页签 1-存在就诊一览页签
Private mbln免挂号模式 As Boolean

'---------------排队叫号相关
'呼叫列宽初始化
Private Const C_STR_QUEUECALL = "0,0,0,0,50,0,90,0,60,0,0,60,60,0,0,60,0,0,125"
'排队列宽初始化
Private Const C_STR_QUEUEQUEUE = "0,0,0,30,50,0,90,40,60,60,0,60,60,50,125,0,120,60,0"

Private Enum mCol
    队列名称 = 0: ID: 病人ID: 排队标记: 排队号码:  排队序号: 患者姓名: 优先: 回诊序号: 回诊排序号: 科室ID: 诊室: 医生姓名: 排队状态: 排队时间: 呼叫医生: 业务类型: 业务ID: 呼叫时间: 排序名称: ORD
End Enum
Private mlngQueueGroupType As Long
Private mstrShowCalledColumnInf As String
Private mstrShowColumnInf As String
Private mlngOrderStyle As Long
Private mlng回诊病人优先 As Long
Private mlngMaxLen As Long
Private mobjQueueList As Object
Private mobjCallList As Object
'------------------

'其它窗体变量
Private mrsAller As ADODB.Recordset '病人过敏记录
Private mstrIDCard As String '最近自动刷出来的身份证号
Private WithEvents mobjIDCard As clsIDCard '身份证对象
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object 'IC卡对象
Private mblnUnRefresh As Boolean
Private mstrPrivs As String
Private mlngModul As Long

Private mPatiInfo As PatiInfo '历史就诊记录中的,不一定为当前的

'-----列表中选择行的相关的信息。
Private mintActive As PatiType '病人类型
Private mintRPTIndex As PATI_RPT_LIST '选中的病人列表索引值，默认值为-1
Private mPr As Long          '选中的病人列表中选择的行号，默认值为-1，通过 mintRPTIndex和mPr来标定当前选中的列表行，不使用RPT控件的 SelectedRows 属性

Private mlng病人ID As Long
Private mstr挂号单 As String
Private mlng挂号ID As Long
Private mlng科室ID As Long

Private mintFindType As Integer '0-就诊卡,1-标识号（即门诊号）,2-挂号单,3-姓名,4-二代身份证,5-IC卡
Private mstrFindType As String '用来存储当查找前类型的名称"就诊卡，标识号，挂号单，姓名，二代身份证，IC卡"
Private mblnFindTypeEnabled As Boolean
Private mstr挂号IDs As String ' 病人挂号记录.ID  串逗号分割，记录当前点击了那些病人
Private mblnIsInit As Boolean 'PatiIdentify初始化标志
Private mobjKernel As zlPublicAdvice.clsPublicAdvice          '临床核心部件

'医疗卡
Private mobjSquareCard As Object      '卡结算对象
Private mstrCardKind As String        '卡结算对象返回的可用的医疗卡

Private mstrPrePati As String
Private mintPreTime As Integer
Private mlngCommunityID As Long '自动执行的社区功能
Private mbytSize As Byte '字体 0-小字体（9号字体），1-大字体（12号字体）
Private mblnTabTmp As Boolean
Private mstrPreSubTab As String ' tbsSub前一次选中的页签
Private mblnSizeTmp As Boolean

Private mblnMsgOk As Boolean '是否有消息来过
Private mblnFirstMsg As Boolean 'mblnFirstMsg=false 表示打开医生站后的第一条消息
Private mintNotify As Integer '医嘱提醒自动刷新间隔(分钟)
Private mintNotifyDay As Integer '提醒多少天内的医嘱
Private mstrNotifyAdvice As String '提醒的医嘱类型
Private mstrPreNotify As String
Private mblnPatiDetail As Boolean
Private mclsMsg As clsCISMsg
Private mrsMsg As ADODB.Recordset
Private mbln消息语音 As Boolean
Private mlng当日接诊人数 As Long
Private mbln危急值 As Boolean '处危急值的权限
Private mbln显示预约病人 As Boolean
Private mint预约列表 As Integer

Private Sub cboPayType_Click()
    Dim strTmp As String
    Dim strSQL As String
    If mstr挂号单 = "" Then Exit Sub
    strTmp = Split(cboPayType.Text, "-")(1)
    If cboPayType.ToolTipText <> strTmp Then
        strTmp = Split(cboPayType.Text, "-")(1)
        strSQL = "Zl_病人挂号记录_更新费别('" & mstr挂号单 & "','" & strTmp & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        cboPayType.Tag = ""
        cboPayType.ToolTipText = strTmp
    End If
End Sub

Private Sub cboBillType_Click()
    Dim strSQL As String
    Dim strTmp As String
    If mlng病人ID = 0 Then Exit Sub
    If cboBillType.ToolTipText <> cboBillType.Text Then
        strTmp = cboBillType.Text
        strSQL = "Zl_病人信息_更新信息(" & mlng病人ID & ",'费别','" & strTmp & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        cboBillType.Tag = ""
        cboBillType.ToolTipText = cboBillType.Text
    End If
End Sub

Private Sub cboSelectTime_Click()
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    intDateCount = cboSelectTime.ItemData(cboSelectTime.ListIndex)
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If Me.Visible Then
        If intDateCount = -1 Then
            If Not frmSelectTime.ShowMe(Me, mvCondFilter.Begin, mvCondFilter.End, cboSelectTime) Then
                '取消时恢复原来的选择
                Call Cbo.SetIndex(cboSelectTime.hwnd, mintOutPreTime)
                Exit Sub
            End If
        ElseIf intDateCount = 0 Then
            '今天  86114
            mvCondFilter.Begin = Format(datCurr, "yyyy-MM-dd 00:00:00")
            mvCondFilter.End = Format(datCurr, "yyyy-MM-dd 23:59:59")
        Else
            mvCondFilter.End = Format(datCurr, "yyyy-MM-dd 23:59:59")
            mvCondFilter.Begin = Format(mvCondFilter.End - intDateCount, "yyyy-MM-dd 00:00:00")
        End If
    End If
    '选择了时间之后，清除挂号单条件
    mvCondFilter.挂号单 = ""
    mvCondFilter.就诊卡 = ""
    mvCondFilter.门诊号 = ""
    mvCondFilter.姓名 = ""
    '保存参数，保证每个地方提取的出院病人都是在同一时间范围内（72783）
    Call zlDatabase.SetPara("已诊病人结束间隔", DateDiff("d", datCurr, mvCondFilter.End), glngSys, p门诊医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
    Call zlDatabase.SetPara("已诊病人开始间隔", DateDiff("d", mvCondFilter.Begin, datCurr), glngSys, p门诊医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
    cboSelectTime.ToolTipText = Format(mvCondFilter.Begin, "yyyy-MM-dd") & " - " & Format(mvCondFilter.End, "yyyy-MM-dd")
    lblSeeTim.ToolTipText = cboSelectTime.ToolTipText
    mintOutPreTime = cboSelectTime.ListIndex
    Call LoadPatients已诊
End Sub

Private Sub cmdOtherFilter_Click()
    Dim datCurr As Date
    
    With mvCondFilter
        .科室ID = IIf(.科室ID = 0, mlng接诊科室ID, .科室ID)
        If frmPatiFilter.ShowMe(Me, .Begin, .End, .科室ID, .医生, .挂号单, .门诊号, .就诊卡, .姓名, mstrPrivs) Then
            datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
            Call Cbo.SetIndex(cboSelectTime.hwnd, 3)
            '保存参数，保证每个地方提取的出院病人都是在同一时间范围内（72783）
            Call zlDatabase.SetPara("已诊病人结束间隔", DateDiff("d", datCurr, mvCondFilter.End), glngSys, p门诊医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
            Call zlDatabase.SetPara("已诊病人开始间隔", DateDiff("d", mvCondFilter.Begin, datCurr), glngSys, p门诊医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
            cboSelectTime.ToolTipText = Format(mvCondFilter.Begin, "yyyy-MM-dd") & " - " & Format(mvCondFilter.End, "yyyy-MM-dd")
            lblSeeTim.ToolTipText = cboSelectTime.ToolTipText
            mintOutPreTime = cboSelectTime.ListIndex
            Call LoadPatients已诊
        End If
    End With
End Sub

Private Sub mclspath_RequestRefresh(ByVal lngPathState As Long)
'功能：临床路径中刷新病人信息列表中的状态,-1表示未导入状态
    mPatiInfo.路径状态 = lngPathState
End Sub

Private Sub picPati_Resize()
    On Error Resume Next
    
    picFind.Top = 0
    picFind.Left = 0
    lblFind.Top = 30
    lblFind.Width = picFind.Width
    PatiIdentify.Left = picFind.Left + picFind.Width + 10
    PatiIdentify.Top = 0
    PatiIdentify.Width = picPati.ScaleWidth - 500
    
    tbcWait.Left = 0
    tbcWait.Width = picPati.ScaleWidth
    tbcWait.Top = PatiIdentify.Top + PatiIdentify.Height
    
    fraPatiUD.Left = -20
    fraPatiUD.Width = picPati.ScaleWidth + 20
    
    tbcInTreat.Left = 0
    tbcInTreat.Width = picPati.ScaleWidth
    
    tbcWait.Height = fraPatiUD.Top - tbcWait.Top
    tbcInTreat.Top = fraPatiUD.Top + 45
    tbcInTreat.Height = picPati.ScaleHeight - fraPatiUD.Top - 45
End Sub

Private Sub fraPatiUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tbcWait.Height + Y < 1000 Or tbcInTreat.Height - Y < 1000 Then Exit Sub
        fraPatiUD.Top = fraPatiUD.Top + Y
        tbcWait.Height = tbcWait.Height + Y
        tbcInTreat.Top = tbcInTreat.Top + Y
        tbcInTreat.Height = tbcInTreat.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub Form_Activate()
    If Check排队叫号 Then
        DoEvents
        mobjQueue.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '读卡
    PatiIdentify.ActiveFastKey
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim blnTmp As Boolean
    
    If InStr("[|']", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    '在编译病人信息的时候不允许自动定位查找控件，否则影响信息的填写。
    If tbcRegist.Selected.Caption = "本次就诊" Then
        If tbcSub.Visible Then
            If tbcSub.Selected.Tag = "病人" Then
                blnTmp = True
            End If
        End If
    End If
    
    If Not blnTmp Then
        If Me.ActiveControl Is UCPatiVitalSigns Then
            If UCPatiVitalSigns.ControlLock = False Then
                blnTmp = True
            End If
        End If
    End If
    
    If Not blnTmp Then
        If (InStr("0123456789", Chr(KeyAscii)) > 0 Or UCase(Chr(KeyAscii)) >= "A" And UCase(Chr(KeyAscii)) <= "Z") _
            And Not Me.ActiveControl Is PatiIdentify And mstrFindType = "就诊卡" And PatiIdentify.Enabled And PatiIdentify.Visible Then
            PatiIdentify.Text = UCase(Chr(KeyAscii))
            PatiIdentify.NotAutoSel = True
            PatiIdentify.SetFocus
        End If
    End If
End Sub

Private Sub mclsAdvices_VSKeyPress(KeyAscii As Integer)
    If (InStr("0123456789", Chr(KeyAscii)) > 0 Or UCase(Chr(KeyAscii)) >= "A" And UCase(Chr(KeyAscii)) <= "Z") _
        And mstrFindType = "就诊卡" And PatiIdentify.Enabled And PatiIdentify.Visible Then
        picFind.SetFocus
        PatiIdentify.Text = UCase(Chr(KeyAscii))
        PatiIdentify.NotAutoSel = True
        PatiIdentify.SetFocus
    End If
End Sub

Private Sub InitQueuePara()
'功能：初始化排队叫号参数
'排队叫号模式:1.代表分诊台分诊呼叫或医生主动呼叫;2-先分诊呼叫,再医生呼叫就诊.0-不排队叫号
    
    mty_Queue.strQueuePrivs = ";" & GetPrivFunc(glngSys, p排队叫号虚拟模块) & ";"
    mty_Queue.byt排队叫号模式 = Val(zlDatabase.GetPara("排队叫号模式", glngSys, p门诊分诊管理))
 
    If mty_Queue.byt排队叫号模式 = 1 Then
        If 0 = Val(zlDatabase.GetPara("叫号方式", glngSys, p排队叫号虚拟模块)) Then
            mty_Queue.bln医生主动呼叫 = Val(zlGetLocaleComputerNamePara("排队呼叫站点", glngSys, p门诊分诊管理, "0")) = 1
        Else
            mty_Queue.str呼叫站点 = zlDatabase.GetPara("远端呼叫站点", glngSys, p排队叫号虚拟模块)
            mty_Queue.bln医生主动呼叫 = Val(zlGetLocaleComputerNamePara("排队呼叫站点", glngSys, p门诊分诊管理, "0", mty_Queue.str呼叫站点)) = 1
        End If
    Else
        mty_Queue.bln医生主动呼叫 = False
    End If
    
    If mty_Queue.bln医生主动呼叫 Then
        mty_Queue.int呼叫人数 = Val(zlDatabase.GetPara("医生就诊人数", glngSys, p门诊医生站))
    Else
        mty_Queue.int呼叫人数 = 0
    End If
    mty_Queue.bln呼叫含回诊 = Val(zlDatabase.GetPara("就诊人数含回诊", glngSys, p门诊医生站, "1")) = 1
End Sub

Private Sub Form_Load()
    Dim objPane As Pane, strTab As String, intIdx As Integer
    Dim intType As Integer, blnHave As Boolean, blnTmp As Boolean
    Dim i As Integer, arrType() As String
    Dim objControl As CommandBarControl, objTabItem As TabControlItem
    Dim arrTmp As Variant, strTmp As String
    
    mstrPrivs = ";" & gstrPrivs & ";"
    mlngModul = glngModul
    mblnShowLeavePati = False
    Call GetLocalSetting '本地参数
    
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, p门诊医生站, GetInsidePrivs(p门诊医生站))
    Call AddMipModule(mclsMipModule)
    
    Set mclsReg = New zlPublicExpense.clsRegist
    Call mclsReg.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
    Call mclsReg.zlInitData(1)
    
    Set mclsDis = New zl9Disease.clsDisease
    Call mclsDis.InitDisease(gcnOracle, Me, glngSys, glngModul, mstrPrivs, mclsMipModule)
    
    Set mobjKernel = New zlPublicAdvice.clsPublicAdvice

    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)
    
    Call InitQueuePara
 
    '一卡通部件初始，须在tbcSub_SelectedChanged之前，以便传递给医嘱部件
     'zlGetIDKindStr中会自动补齐为至少8位属性
    mstrCardKind = "就|就诊卡|0|0|8|0|0|0;门|标识号|0|0|0|0|0|0;挂|挂号单|0|0|0|0|0|0;姓|姓名|0|0|0|0|0|0;身|二代身份证|0|0|0|0|0|0;ＩＣ|ＩＣ卡|1|0|0|0|0|0"
    If Check排队叫号 = True Then mstrCardKind = mstrCardKind & ";排|排队号|0|0|0|0|0|0;医|医保号|0|0|0|0|0|0"
    On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    err.Clear: On Error GoTo 0
    If Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
            Set mobjSquareCard = Nothing
            MsgBox "医疗卡部件（zl9CardSquare）初始化失败!", vbInformation, gstrSysName
        Else
            mstrCardKind = mobjSquareCard.zlGetIDKindStr(mstrCardKind)
        End If
    End If
    Call PatiIdentify.zlInit(Me, glngSys, p门诊医生站, gcnOracle, gstrDBUser, mobjSquareCard, mstrCardKind, "zl9CISJob")
    PatiIdentify.objIDKind.AllowAutoICCard = True
    PatiIdentify.objIDKind.AllowAutoIDCard = True
    mblnIsInit = True

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
    objPane.Title = "门诊病人列表"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    Set objPane = Me.dkpMain.CreatePane(2, 310, 100, DockBottomOf, objPane)
    objPane.Title = "消息提醒"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    'TabControl
    '-----------------------------------------------------
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, True)
 
    '候诊列表
    With Me.tbcWait
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem(0, "候诊病人", picHZ.hwnd, 0).Tag = "候诊病人"
        .InsertItem(1, "排队叫号", picTmpH(0).hwnd, 0).Tag = "排队叫号"
        .InsertItem(2, "预约病人", picYy.hwnd, 0).Tag = "预约病人"
        
        If Not mbln显示预约病人 Then .Item(2).Selected = True
        .Item(2).Visible = Not mbln显示预约病人
        .Item(1).Selected = True
        .Item(0).Selected = True
        
        Call .RemoveItem(1)
        
        If Check排队叫号 Then
            .InsertItem(1, "排队叫号", mobjQueue.zlGetForm.hwnd, 0).Tag = "排队叫号"
        End If
        
        mint预约列表 = IIf(Check排队叫号, 2, 1)
    End With
    
    '就诊列表
    With Me.tbcInTreat
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem(0, "在诊", picJZ.hwnd, 0).Tag = "正在就诊"
        .InsertItem(1, "完成", picYZ.hwnd, 0).Tag = "已诊病人"
        .InsertItem(2, "回诊", picHUIZ.hwnd, 0).Tag = "需回诊病人"
        .Item(2).Selected = True
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    
    Set mclsAdvices = New zlPublicAdvice.clsDockOutAdvices
    Set mclsEPRs = New zlRichEPR.cDockOutEPRs
    Set mclsDisease = New zlRichEPR.cDockDisease
    Set mobjPati = New frmDockPatiInfo
    Set mclsPath = New zlPublicPath.clsDockOutPath
    Call mclsAdvices.zlInitPath(mclsPath)
    If GetInsidePrivs(p新版门诊病历, True) <> "" Then
        Set mclsEMR = DynamicCreate("zlRichEMR.clsDockEMR", "电子病历")
        If Not mclsEMR Is Nothing Then
            If Not mclsEMR.Init(gobjEmr, gcnOracle, glngSys) Then
                Set mclsEMR = Nothing
            End If
        End If
    End If
    
    If InStr(";" & gstrPrivs & ";", ";诊疗一览;") > 0 Then
        Set mfrmView = New frmOutDoctorView
        mbyt本次就诊 = 1
    Else
        mbyt本次就诊 = 0
    End If
    
    Set mcolSubForm = New Collection
    If mbyt本次就诊 = 1 Then
        mcolSubForm.Add mfrmView, "_诊疗一览"
    End If
    mcolSubForm.Add mobjPati, "_病人"
    mcolSubForm.Add mclsPath.zlGetForm, "_路径"
    mcolSubForm.Add mclsAdvices.zlGetForm, "_医嘱"
    mcolSubForm.Add mclsEPRs.zlGetForm, "_病历"
    mcolSubForm.Add mclsDisease.zlGetForm, "_疾病报告"
    If Not mclsEMR Is Nothing Then
        mcolSubForm.Add mclsEMR.zlGetForm, "_新病历"
    End If
    
    
    '---------------------------------------------------
    '历次就诊列表
    With Me.tbcRegist
        Set tbcRegist.Icons = zlCommFun.GetPubIcons
        With .PaintManager
            .Appearance = xtpTabAppearanceStateButtons
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        If mbyt本次就诊 = 1 Then
             .InsertItem(0, "诊疗一览", mcolSubForm("_诊疗一览").hwnd, 0).Tag = "诊疗一览"
        End If
        .InsertItem(mbyt本次就诊, "本次就诊", picRegist.hwnd, 0).Tag = "-1"
            intIdx = mbyt本次就诊 + 1
        .InsertItem(intIdx, "历史就诊1", picRegist.hwnd, 0).Tag = "-1"
            .Item(intIdx).Visible = False: intIdx = intIdx + 1
        .InsertItem(intIdx, "历史就诊2", picRegist.hwnd, 0).Tag = "-1"
            .Item(intIdx).Visible = False: intIdx = intIdx + 1
        .InsertItem(intIdx, "历史就诊3", picRegist.hwnd, 0).Tag = "-1"
            .Item(intIdx).Visible = False: intIdx = intIdx + 1
        .InsertItem(intIdx, "更多", picRegist.hwnd, 0).Tag = "更多"
            .Item(intIdx).Visible = False
        intIdx = 0
        tbcSub.Visible = True
        picRegist.Visible = True
    End With
    
    '内部卡片
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem(intIdx, "病人信息", picTmp.hwnd, 0).Tag = "病人": intIdx = intIdx + 1
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        If GetInsidePrivs(P门诊路径应用, True) <> "" Then
            .InsertItem(intIdx, "临床路径", picTmp.hwnd, 0).Tag = "路径": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p门诊医嘱下达) <> "" Then
            '先加载医嘱的原因:在启用美康接口，但客户端没有美康部件时。如果先加载排队叫号后加载医嘱的时候，
            '从“病历信息”切换到“医嘱信息”会因弹出Msgbox报错 问题号:67995
            .InsertItem(intIdx, "医嘱信息", mcolSubForm("_医嘱").hwnd, 0).Tag = "医嘱": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p门诊病历管理) <> "" Then
            .InsertItem(intIdx, "病历信息", picTmp.hwnd, 0).Tag = "病历": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p新版门诊病历, True) <> "" And Not mclsEMR Is Nothing Then
            .InsertItem(intIdx, "电子病历", picTmp.hwnd, 0).Tag = "新病历": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p疾病报告填写, True) <> "" Then
            Set objTabItem = .InsertItem(intIdx, "疾病报告", picTmp.hwnd, 0): objTabItem.Tag = "疾病报告": objTabItem.Visible = False: intIdx = intIdx + 1
        End If
        
        '外挂提供的卡片
        Call CreatePlugInOK(p门诊医生站)
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strTmp = gobjPlugIn.GetFormCaption(glngSys, p门诊医生站)
            Call zlPlugInErrH(err, "GetFormCaption")
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ",")
                For i = 0 To UBound(arrTmp)
                    strTmp = arrTmp(i)
                    mcolSubForm.Add gobjPlugIn.GetForm(glngSys, p门诊医生站, strTmp), "_" & strTmp
                    .InsertItem(intIdx, strTmp, mcolSubForm("_" & strTmp).hwnd, 0).Tag = strTmp: intIdx = intIdx + 1
                    Call zlPlugInErrH(err, "GetForm")
                Next
            End If
            err.Clear: On Error GoTo 0
        End If
        
        If .ItemCount = 0 Then
            MsgBox "你没有使用门诊医生工作站的权限。", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        '恢复上次选择的卡片
        strTab = zlDatabase.GetPara("医护功能", glngSys, p门诊医生站)
        
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
    End With
    
    tbcRegist.Item(mbyt本次就诊).Selected = True
    
    Call tbcRegist_SelectedChanged(tbcRegist.Selected)
    mstrPreSubTab = ""
    '只加载选择的子窗体
    Call tbcSub_SelectedChanged(tbcSub.Selected)
            
    '读取界面数据
    '-----------------------------------------------------
    mblnUnRefresh = True
    mstrPrePati = ""
    mintPreTime = -1
    mintActive = -1
    mPr = -1
    
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hwnd)
    On Error Resume Next
    Set mobjICCard = CreateObject("zlICCard.clsICCard")
    If Not mobjICCard Is Nothing Then
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    err.Clear: On Error GoTo 0
    
    Call InitCboData
    Call InitReportColumn
    Call InitCondFilter '已诊病人过滤条件
    
    Call LoadPatients '显示数据
    Call LoadNotify '消息提醒
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        '会恢复Panne的标题,Tag被清除
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, "")
    End If
 
    '设置缺省查找方式
    arrType = Split(mstrCardKind, ";")
    For i = 1 To UBound(arrType) + 1
        If i = mintFindType Then
            PatiIdentify.objIDKind.IDKind = i
            Exit For
        End If
    Next
    
    Call RestoreWinState(Me, App.ProductName, , True)
    
    'ReportControl控件用了数组无法恢复要单独处理
    For i = 0 To rptPati.Count - 1
        strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\ReportControl", "rptPati" & "_" & i, "")
        rptPati(i).LoadSettings strTmp
    Next
    
    Me.WindowState = vbMaximized
    Call SetFixedCommandBar(cbsMain(2).Controls)
    fraPatiUD.Top = 3000
    If Check排队叫号 = True Then
                fraPatiUD.Top = 5000
        '检查是否存在排队叫号
        Call ReshDataQueue
        tbcWait.Item(1).Selected = True
    End If
    
    Call RefreshPass

    If ISPassShowCard Then
        rptPati(PATI_RPT候诊).Columns(COL_HZ_就诊卡号).Visible = False
        rptPati(PATI_RPT预约).Columns(COL_HZ_就诊卡号).Visible = False
        rptPati(PATI_RPT就诊).Columns(COL_JZ_就诊卡号).Visible = False
        rptPati(PATI_RPT回诊).Columns(COL_JZ_就诊卡号).Visible = False
        rptPati(PATI_RPT已诊).Columns(COL_YZ_就诊卡号).Visible = False
    End If
    If InStr(";" & gstrPrivs & ";", ";修改医疗付款方式;") = 0 Then
        cboPayType.Locked = True
    End If
    If InStr(";" & gstrPrivs & ";", ";修改费别;") = 0 Then
        cboBillType.Locked = True
    End If
    Call SetReceiveToday(True, 0)
    If mblnPatiDetail Then
        lblLink(lblLink显示).Caption = "隐藏其它信息↑"
    Else
        lblLink(lblLink显示).Caption = "显示其它信息↓"
    End If
    mblnUnRefresh = False
End Sub

Private Sub RefreshPass()
    '是否调用太元通接口部件
    mblnUseTYT = False
    If gbytPass = 3 Then
        If gint过敏输入来源 = 0 Then
            mint过敏输入来源 = Val(zlDatabase.GetPara("过敏输入来源", glngSys, p门诊医生站, "0"))
        End If
        mblnUseTYT = gint过敏输入来源 = 0 And mint过敏输入来源 = 1 Or gint过敏输入来源 = 2
    End If
    '创建太元通接口对象，创建失败，则不启用太元通
    If gbytPass = 3 Then
        On Error Resume Next
        If gobjPass Is Nothing Then
            Set gobjPass = DynamicCreate("zlPassInterface.clsPass", "合理用药监测", True)
            If Not gobjPass Is Nothing Then
                Call gobjPass.zlPassInit(gcnOracle, glngSys, 5)
                If gobjPass.PassType = 0 Then
                    Set gobjPass = Nothing
                End If
            End If
        End If
        If err.Number <> 0 Then err.Clear: gbytPass = 0
        If gobjPass Is Nothing Then gbytPass = 0
        On Error GoTo 0
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    Dim str挂号单 As String, strCardNO As String
    Dim rsTmp As Recordset
    Dim str疾病ID As String, str诊断ID As String
    Dim intFindTypeTmp As Integer
    Dim strPictureFile As String
    
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
            mbytSize = 0
            Call zlDatabase.SetPara("字体", mbytSize, glngSys, p门诊医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_FontSize_L '大字体
        If mbytSize <> 1 Then
            mbytSize = 1
            Call zlDatabase.SetPara("字体", mbytSize, glngSys, p门诊医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
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
        If PatiIdentify.Text = "" And mstrIDCard = "" Then
            PatiIdentify.SetFocus
        Else
            Call ExecuteFindPati(True, IIf(PatiIdentify.Text = "", mstrIDCard, ""))
        End If
    Case conMenu_View_Busy '诊室状态
        Call SetRoomState(lblRoom.BackColor = COLOR_FREE)
    Case conMenu_View_Refresh '刷新
        Call LoadPatients("11001")
        Call LoadNotify
    Case conMenu_View_Jump '跳转
        If Me.tbcSub.Selected.Index + 1 <= Me.tbcSub.ItemCount - 1 Then
            Me.tbcSub.Item(Me.tbcSub.Selected.Index + 1).Selected = True
        Else
            Me.tbcSub.Item(0).Selected = True
        End If
    Case conMenu_File_Parameter '参数设置
        frmOutStationSetup.mstrPrivs = mstrPrivs
        frmOutStationSetup.Show 1, Me

        If gblnOK Then
            intFindTypeTmp = mintFindType
            Call GetLocalSetting
            mintFindType = intFindTypeTmp
            Call LoadPatients
            Call InitQueuePara
        End If
        If Check排队叫号 Then
            Call ReshDataQueue
        End If
        Me.tbcWait.Item(mint预约列表).Visible = Not mbln显示预约病人
        If Me.tbcWait.Item(mint预约列表).Visible = False Then
             Me.tbcWait.Item(0).Selected = True
        End If
    Case conMenu_Tool_KssAudit '抗菌用药审核
        On Error Resume Next
        Call frmExamineKSS.Show(0, Me)
    Case conMenu_Tool_CISMed  '临床自管药
        Call Set临床自管药(Me)
     Case conMenu_Tool_TransAudit '输血审核管理
        On Error Resume Next
        Call frmExamineTransfuse.ShowMe(Me, 1)
    Case conMenu_Tool_Archive '电子病案查阅
        Call frmArchiveView.ShowArchive(Me, mPatiInfo.病人ID, mPatiInfo.挂号ID)
    Case conMenu_Tool_ExaReport
        '调用陈福荣提供的接口 OpenExaminationReport
    Case conMenu_Tool_Reference_1 '疾病诊断参考
        Call gobjKernel.ShowDiagHelp(vbModeless, Me)
    Case conMenu_Tool_Reference_2 '诊疗措施参考
        Call gobjKernel.ShowClincHelp(vbModeless, Me)
    Case conMenu_Manage_FeeItemSet  '诊疗项目费用设置
        Call Set诊疗项目费用设置
        
    Case conMenu_Tool_Community * 100# + 1 '社区身份验证
        Call ExecuteCommunityIdentify
    Case conMenu_Tool_Community * 100# + 2 To conMenu_Tool_Community * 100# + 99 '社区其他功能
        If Not gobjCommunity Is Nothing And mPatiInfo.社区 <> 0 And mPatiInfo.挂号ID <> 0 Then
            If gobjCommunity.CommunityFunc(glngSys, mlngModul, Val(Control.Parameter), mPatiInfo.社区, mPatiInfo.社区号, mPatiInfo.病人ID, mPatiInfo.挂号ID) Then
                Call LoadPatients
            End If
        End If
    Case conMenu_Manage_Regist '病人挂号
        Control.Enabled = False
        Call mclsReg.zlExecuteCommandBars(Me, Control, str挂号单)
        Call ExecuteRegist(str挂号单)
        If str挂号单 <> "" Then Call SetReceiveToday(False, 1): Call ReceiveAfterExec
        Control.Enabled = True
    Case conMenu_Manage_Bespeak '预约挂号
        Control.Enabled = False
        Call mclsReg.zlExecuteCommandBars(Me, Control, "", mlng病人ID)
        Control.Enabled = True
    Case conMenu_Edit_AppRequest
        Control.Enabled = False
        Call mclsReg.zlExecuteCommandBars(Me, Control, "", mlng病人ID)
        Control.Enabled = True
    Case conMenu_Edit_AppRequestManage
        Control.Enabled = False
        Call mclsReg.zlExecuteCommandBars(Me, Control, "")
        Control.Enabled = True
    Case conMenu_View_Option '"挂号选项设置"
        Control.Enabled = False
        Call mclsReg.zlExecuteCommandBars(Me, Control, "")
        Control.Enabled = True
    Case conMenu_File_Print_Bespeak '重打预约挂号单
        Control.Enabled = False
        Call ExecuteBespeakPrint
        Control.Enabled = True
    Case conMenu_Manage_Transfer_Send '病人转诊
        Call ExecuteTransferSend
    Case conMenu_Manage_Transfer_Cancel '取消转诊
        Call ExecuteTransferCancel
    Case conMenu_Manage_Transfer_Incept '接收转诊
        Call ExecuteTransferIncept
    Case conMenu_Manage_Transfer_Refuse '转诊拒绝
        Call ExecuteTransferRefuse
    Case conMenu_Manage_Transfer_Force '强制续诊
        str挂号单 = frmForceGet.ShowMe(Me, mstrPrivs, mlng接诊科室ID, mobjSquareCard)
        If str挂号单 <> "" Then
            If rptPati(PATI_RPT就诊).Visible Then
                Call LoadPatients("11001", PATI_RPT就诊, str挂号单)
            Else
                Call LoadPatients("11001")
            End If
        End If
    Case conMenu_Manage_Receive '病人接诊
        Call ExecuteReceive
    Case conMenu_Manage_Cancel '取消接诊
        Call ExecuteCancel
    Case conMenu_Manage_Finish '完成接诊
        Call ExecuteFinish
    Case conMenu_Manage_Redo '恢复接诊
        Call ExecuteRedo
    Case conMenu_Manage_ReBack '暂停就诊
        Call ExecuteStopAndReuse(False)
    Case conMenu_Manage_ReBackCancel '恢复暂停就诊
          Call ExecuteStopAndReuse(True)
   Case conmenu_View_Leave  '显示不就诊病人
         mblnShowLeavePati = Not mblnShowLeavePati
         Control.Checked = mblnShowLeavePati
        Call LoadPatients("10001")
    Case conmenu_Edit_Leave     '病人不就诊
        If Set病人挂号状态(-1) Then
            Call LoadPatients("10001")
            Call ReshDataQueue
        End If
    Case conmenu_Edit_Wait      '病人就诊
        If Set病人挂号状态(0) Then
            Call LoadPatients("10001")
            Call ReshDataQueue
        End If
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_Help '帮助
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '退出
        Unload Me
    Case conMenu_Tool_HealthCard  '居民健康卡
        If Not mobjSquareCard Is Nothing Then
            Call mobjSquareCard.zlHealthArchivesShow(Me, p门诊医生站, mlng病人ID, "")
        End If
    Case conMenu_Edit_TraReactionRecord '输血反应
        Call FuncTraReactionRecord(Me, 0, p门诊医嘱下达)
    Case conMenu_Edit_NewItemQAdvice
        Call ExecuteTabChange("医嘱")
    Case conMenu_Edit_NewItemQEpr
        Call ExecuteTabChange("病历")
    Case conMenu_Tool_Positive '阳性结果查看
        Call mclsDis.ShowRegistByPati(Me, ByVal 1, mlng病人ID, , mstr挂号单)
    Case conMenu_Tool_Critical '危急值查看处理
        Call ExecuteCritical
    Case Else
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            '执行发布到当前模块的报表
            With mPatiInfo
                If Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_2" Or Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_3" Then
                    If mlng接诊科室ID = 0 Then
                        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
                    Else
                        Set rsTmp = zlDatabase.OpenSQLRecord("Select 名称 From 部门表 Where ID=[1]", Me.Caption, mlng接诊科室ID)
                        If rsTmp.EOF Then Exit Sub
                        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                            "开嘱科室=" & rsTmp!名称 & "|=" & mlng接诊科室ID)
                    End If
                Else
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                        "病人ID=" & mPatiInfo.病人ID, "门诊号=" & .门诊号, "挂号单=" & .挂号单, "诊室=" & .诊室)
                End If
            End With
        Else
            If Check排队叫号 = True Then
                mobjQueue.zlExecuteCommandBars Control
            End If
            Select Case Me.tbcSub.Selected.Tag
            Case "路径"
                Call mclsPath.zlExecuteCommandBars(Control)
            Case "医嘱"
                Call mclsAdvices.zlExecuteCommandBars(Control)
            Case "病历"
                Call mclsEPRs.zlExecuteCommandBars(Control)
            Case "新病历"
                Call mclsEMR.zlExecuteCommandBars(Control)
            Case "疾病报告"
                Call mclsDisease.zlExecuteCommandBars(Control)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.ExeButtomClick(glngSys, p门诊医生站, mcolSubForm("_" & tbcSub.Selected.Tag), tbcSub.Selected.Tag, Control.Caption, mlng病人ID, 0, mstr挂号单)
                    Call zlPlugInErrH(err, "ExeButtomClick")
                    err.Clear: On Error GoTo 0
                End If
            End Select
        End If
    End Select
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl
    Dim strFunc As String, arrFunc As Variant
    Dim i As Long
    Dim arrKind() As String
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID

    Case conMenu_Manage_Transfer
        With CommandBar.Controls
            If .Count = 0 Then
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Send, "转诊病人(&S)", -1, False)
                objControl.IconId = conMenu_Manage_Transfer
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Cancel, "取消转诊(&C)", -1, False)
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Incept, "转诊接收(&I)", -1, False)
                objControl.IconId = conMenu_Manage_Receive
                objControl.BeginGroup = True
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Refuse, "转诊拒绝(&R)", -1, False)
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Force, "强制续诊(&F)", -1, False)
                objControl.BeginGroup = True
            End If
        End With
    Case conMenu_Tool_Community '社区功能
        mlngCommunityID = 0
        With CommandBar.Controls
            .DeleteAll
            If Not gobjCommunity Is Nothing Then
                '补充验证
                If mPatiInfo.社区 = 0 Then
                    Set objControl = .Add(xtpControlButton, conMenu_Tool_Community * 100# + 1, "身份验证(&V)")
                End If
                
                '其他功能
                If mPatiInfo.社区 <> 0 Then
                    strFunc = gobjCommunity.GetCommunityFunc(glngSys, p门诊医生站, mPatiInfo.社区)
                    If strFunc <> "" Then
                        arrFunc = Split(strFunc, ";")
                        For i = 0 To UBound(arrFunc)
                            Set objControl = .Add(xtpControlButton, conMenu_Tool_Community * 100# + i + 2, Split(arrFunc(i), ",")(1))
                            If i < 9 Then objControl.Caption = objControl.Caption & "(&" & i + 1 & ")"
                            
                            If UCase(arrFunc(i)) Like UCase("Auto:*") Then
                                objControl.Parameter = Mid(Split(arrFunc(i), ",")(0), 6)
                                mlngCommunityID = objControl.ID
                            Else
                                objControl.Parameter = Split(arrFunc(i), ",")(0)
                            End If
                            objControl.ToolTipText = Split(arrFunc(i), ",")(2)
                        Next
                    End If
                End If
            End If
        End With
    Case Else
       Select Case tbcSub.Selected.Tag
       Case "医嘱"
           Call mclsAdvices.zlPopupCommandBars(CommandBar)
       Case "路径"
            Call mclsPath.zlPopupCommandBars(CommandBar)
       Case "病历"
       End Select
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    Dim strTmp As String
 
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
    Case conMenu_View_Busy '诊室状态
        Control.Checked = lblRoom.BackColor = COLOR_BUSY
    Case conMenu_Tool_KssAudit  '抗菌用药审核
        If GetInsidePrivs(p抗菌用药审核) = "" Then
            Control.Visible = False
        End If
    Case conMenu_Tool_TransAudit '输血分级管理
        If GetInsidePrivs(p输血审核管理) = "" Or Not gbln输血分级管理 Then
            Control.Visible = False
        End If
    Case conMenu_Tool_CISMed  '临床自管药
        If InStr(GetInsidePrivs(p门诊医生站), ";临床自管药;") = 0 Then
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
    Case conMenu_Tool_Community '社区菜单
        If gobjCommunity Is Nothing Then
            Control.Visible = False
        End If
    Case conMenu_Edit_TraReactionRecord '输血反应
        Control.Visible = InStr(1, GetInsidePrivs(9005, , 2200), "输血反应登记") <> 0
        Control.Enabled = Control.Visible And gbln血库系统
    Case conMenu_Manage_FeeItemSet '诊疗项目费用设置,没有权限时可查看
                
    Case conMenu_Tool_Community * 100# + 1 '社区身份验证
        Control.Enabled = mlng病人ID <> 0 And mPatiInfo.社区 = 0 And (mPatiInfo.类型 = pt就诊 Or mPatiInfo.类型 = pt回诊) And InStr(mstrPrivs, "病人接诊") > 0
    Case conMenu_Tool_Community * 100# + 2 To conMenu_Tool_Community * 100# + 99 '社区其他功能
        Control.Enabled = mlng病人ID <> 0 And mPatiInfo.社区 <> 0

    Case conMenu_File_MedRec '首页打印
        If InStr(mstrPrivs, "打印首页") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mlng病人ID <> 0
        End If
    Case conMenu_ManagePopup '“接诊”菜单
        If InStr(mstrPrivs, ";病人接诊;") = 0 Then Control.Visible = False
    Case conMenu_File_Print_Bespeak
        Control.Visible = InStr(mstrPrivs, ";预约挂号单;") > 0 And (rptPati(PATI_RPT候诊).Visible Or rptPati(PATI_RPT预约).Visible)
        blnEnabled = False
        If mPr <> -1 Then
            If mintRPTIndex = PATI_RPT预约 Or mintRPTIndex = PATI_RPT候诊 Then
                strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_HZ_标识).Value
                blnEnabled = (strTmp = "预")
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Manage_Transfer '转诊处理
        If InStr(mstrPrivs, "病人接诊") = 0 _
            And InStr(mstrPrivs, "病人转诊") = 0 _
                And InStr(mstrPrivs, "续诊病人") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_Manage_Transfer_Send '病人转诊
        If InStr(mstrPrivs, "病人转诊") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = (mintActive = pt候诊 Or mintActive = pt就诊)
            If blnEnabled Then
                If mPr <> -1 And rptPati(mintRPTIndex).Visible Then
                    If mintRPTIndex = PATI_RPT候诊 Or mintRPTIndex = PATI_RPT预约 Then
                        strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_HZ_状态).Value
                    ElseIf mintRPTIndex = PATI_RPT就诊 Or mintRPTIndex = PATI_RPT回诊 Then
                        strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_JZ_状态).Value
                    End If
                    blnEnabled = (strTmp = "" Or Val(strTmp) = 1)
                Else
                    blnEnabled = False
                End If
            End If
            
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Manage_Transfer_Cancel '取消转诊
        If InStr(mstrPrivs, "病人转诊") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = (mintActive = pt候诊 Or mintActive = pt就诊)
            If blnEnabled Then
                If mPr <> -1 And rptPati(mintRPTIndex).Visible Then
                    If mintRPTIndex = PATI_RPT候诊 Or mintRPTIndex = PATI_RPT预约 Then
                        strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_HZ_状态).Value
                    ElseIf mintRPTIndex = PATI_RPT就诊 Or mintRPTIndex = PATI_RPT回诊 Then
                        strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_JZ_状态).Value
                    End If
                    blnEnabled = (strTmp <> "" And Val(strTmp) = 0 Or Val(strTmp) = -1)
                Else
                    blnEnabled = False
                End If
            End If
            Control.Enabled = blnEnabled
        End If
    Case conmenu_View_Leave  '显示不就诊病人
        Control.Checked = mblnShowLeavePati
    Case conmenu_Edit_Leave
            blnEnabled = (mintActive = pt候诊)
            If blnEnabled Then
                If mPr <> -1 And rptPati(mintRPTIndex).Visible Then
                    strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_HZ_执行状态).Value
                    blnEnabled = Val(strTmp) = 0
                Else
                    blnEnabled = False
                End If
            End If
            Control.Enabled = blnEnabled
    Case conmenu_Edit_Wait
        blnEnabled = mintActive = pt候诊
        If blnEnabled Then
            If mPr <> -1 And rptPati(mintRPTIndex).Visible Then
                strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_HZ_执行状态).Value
                blnEnabled = Val(strTmp) = -1
            Else
                blnEnabled = False
            End If
            
        End If
        Control.Enabled = blnEnabled
        
    Case conMenu_Manage_Transfer_Incept, conMenu_Manage_Transfer_Refuse '转诊接收,转诊拒绝
        If InStr(mstrPrivs, "病人接诊") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = (mintActive = pt转诊 And mPr <> -1 And rptPati(mintRPTIndex).Visible)
            Control.Enabled = blnEnabled
        End If
        
    Case conMenu_Manage_Transfer_Force '强制续诊
        If InStr(mstrPrivs, "病人接诊") = 0 Or InStr(mstrPrivs, "续诊病人") = 0 Then Control.Visible = False
    Case conMenu_Manage_ReBack '暂停就诊:需回诊
        If InStr(mstrPrivs, "病人接诊") = 0 Then
            Control.Visible = False
        Else
            If mintActive = pt就诊 And mPr <> -1 And rptPati(mintRPTIndex).Visible Then
                strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_JZ_记录标志).Value
                blnEnabled = Val(strTmp) < 2
            Else
                blnEnabled = False
            End If
            
            If blnEnabled Then
                If mstr接诊医生 <> UserInfo.姓名 Then
                    blnEnabled = False
                    If InStr(GetInsidePrivs(p门诊医生站), ";操作其他医生的病人;") > 0 Then
                        blnEnabled = True
                    End If
                End If
            End If
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Manage_ReBackCancel '恢复暂停就诊
        If InStr(mstrPrivs, "病人接诊") = 0 Then
            Control.Visible = False
        Else
            
            If mintActive = pt回诊 And mPr <> -1 And rptPati(mintRPTIndex).Visible Then
                strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_JZ_记录标志).Value
                Control.Enabled = Val(strTmp) = 2
            Else
                Control.Enabled = False
            End If
            
        End If
    Case conMenu_Manage_Receive '病人接诊
        If InStr(mstrPrivs, "病人接诊") = 0 Or (mty_Queue.bln医生主动呼叫 And mbln呼叫后接诊) Then
            Control.Enabled = False
            Control.Visible = False
        Else
            Control.Visible = True
            '候诊，预约挂号病人可以直接接诊，转诊病人不通过这个功能
            blnEnabled = False
            
            If (mintRPTIndex = PATI_RPT候诊 Or mintRPTIndex = PATI_RPT预约) And rptPati(mintRPTIndex).Visible Then
                blnEnabled = mPr <> -1
            End If
            Control.Enabled = blnEnabled    '不用再判断当前是否为转诊病人列表，因为如果是转诊列表的话，blnEnabled已经是False
             
        End If
    Case conMenu_Manage_Cancel '取消接诊
        If InStr(mstrPrivs, "病人接诊") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mintActive = pt就诊 And mPr <> -1 And rptPati(mintRPTIndex).Visible
        End If
    Case conMenu_Manage_Finish '完成就诊
        If InStr(mstrPrivs, "病人接诊") = 0 Then
            Control.Visible = False
        ElseIf mintRPTIndex = PATI_RPT就诊 Or mintRPTIndex = PATI_RPT回诊 Then
            blnEnabled = mPr <> -1 And rptPati(mintRPTIndex).Visible
            If mstr接诊医生 <> UserInfo.姓名 And blnEnabled Then
                blnEnabled = False
                If InStr(GetInsidePrivs(p门诊医生站), ";操作其他医生的病人;") > 0 Then
                    blnEnabled = True
                End If
            End If
            Control.Enabled = blnEnabled
        Else
            Control.Enabled = False
        End If
    Case conMenu_Manage_Redo '恢复接诊
        If InStr(mstrPrivs, "病人接诊") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = mintActive = pt已诊 And mPr <> -1 And rptPati(mintRPTIndex).Visible
            If blnEnabled Then '只能恢复接诊自已的病人(否则有权限可用强制续诊)
                blnEnabled = rptPati(mintRPTIndex).Rows(mPr).Record(COL_YZ_执行人).Value = UserInfo.姓名
            End If
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Edit_NewItemQAdvice
        If tbcSub.Selected.Tag = "医嘱" And tbcRegist.Selected.Tag <> "诊疗一览" Then
            Control.Visible = False
        Else
            Control.Visible = True
            blnEnabled = True
            If mstr接诊医生 <> UserInfo.姓名 And blnEnabled Then
                blnEnabled = False
                If InStr(GetInsidePrivs(p门诊医生站), ";操作其他医生的病人;") > 0 Then
                    blnEnabled = True
                End If
            End If
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Edit_NewItemQEpr
        If tbcSub.Selected.Tag = "病历" And tbcRegist.Selected.Tag <> "诊疗一览" Then
            Control.Visible = False
        Else
            Control.Visible = True
        End If
    Case conMenu_Tool_Positive
        Control.Enabled = IIf(Val(mlng病人ID) = 0, False, True)
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
        If Check排队叫号 Then mobjQueue.zlUpdateCommandBars Control
        mclsReg.zlUpdateCommandBars Control
        Select Case tbcSub.Selected.Tag
        Case "路径"
            Call mclsPath.zlUpdateCommandBars(Control)
        Case "医嘱"
            Call mclsAdvices.zlUpdateCommandBars(Control)
        Case "病历"
            Call mclsEPRs.zlUpdateCommandBars(Control)
        Case "新病历"
            Call mclsEMR.zlUpdateCommandBars(Control)
        Case "疾病报告"
            Call mclsDisease.zlUpdateCommandBars(Control)
        End Select
        '抗菌用药报表
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_2" Or Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_3" Then
                If gblnKSSStrict Then
                    Control.Visible = True
                Else
                    Control.Visible = False
                End If
            End If
        End If
    End Select
End Sub

Private Sub SubWinDefCommandBar(ByVal objItem As TabControlItem, Optional ByVal intType As Integer)
'功能：刷新子窗体菜单及工具条
'参数：intType 0－内部TabControl,1-点就诊TabControl
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
        
    Me.Caption = "门诊医生工作站 - " & objItem.Caption & "(当前用户：" & UserInfo.姓名 & ")"
    
    '删除现在的工具栏及顶级菜单项
    For lngCount = cbsMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.Count To 2 Step -1
        cbsMain(lngCount).Delete
    Next
    
    '主窗口重新加入
    Call MainDefCommandBar
    
    If Not mclsReg Is Nothing Then Call mclsReg.zlDefCommandBars(Me, Me.cbsMain, True)
    
    '子窗口重新加入
    Select Case objItem.Tag
    Case "诊疗一览", "病人"
        '诊疗一览/病人信息页 页签不用加载菜单
    Case "医嘱"
        Call mclsAdvices.zlDefCommandBars(Me, Me.cbsMain, 0, gobjPlugIn, mobjSquareCard)
    Case "病历"
        Call mclsEPRs.zlDefCommandBars(Me.cbsMain)
    Case "新病历"
        Call mclsEMR.zlDefCommandBars(Me.cbsMain)
    Case "疾病报告"
        Call mclsDisease.zlDefCommandBars(Me.cbsMain)
    Case "路径"
        Call mclsPath.zlDefCommandBars(Me, Me.cbsMain)
    Case Else
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strName = gobjPlugIn.GetButtomName(glngSys, p门诊医生站, mcolSubForm("_" & objItem.Tag), objItem.Tag)
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

Private Sub SubWinRefreshData(ByVal objItem As TabControlItem, Optional ByVal intType As Integer)
'功能：刷新子窗体数据及状态
'参数：intType 0－内部TabControl,1-点就诊TabControl
Dim i As Integer, blnDis As Boolean
    If mlng病人ID = 0 Or (mintActive = pt候诊 And mPatiInfo.挂号单 = mstr挂号单) Then
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

        '候诊和预约病人，本次就诊没有医嘱和病历数据
        '要求子窗体按无数据处理界面
        Select Case objItem.Tag
        Case "病人"
            Call mobjPati.zlRefresh(0, 0, False, False)
        Case "路径"
            Call mclsPath.zlRefresh(0, 0, "", 0, 0, False)
        Case "医嘱"
            Call mclsAdvices.zlRefresh(0, "", False)
        Case "病历"
            Call mclsEPRs.zlRefresh(0, 0, 0, False, False)
        Case "新病历"
            Call mclsEMR.zlRefresh(0, 0, 0, 0, 1)
        Case "疾病报告"
            Call mclsDisease.zlRefresh(0, 0, 1, 0, False, False)
        Case "诊疗一览"
            Call mfrmView.zlRefresh(Me, 0, 0)
        Case Else
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                Call gobjPlugIn.RefreshForm(glngSys, p门诊医生站, mcolSubForm("_" & objItem.Tag), objItem.Tag, 0, "", 0, False)
                Call zlPlugInErrH(err, "RefreshForm")
                err.Clear: On Error GoTo 0
            End If
        End Select
    Else
        With mPatiInfo
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
            
            Select Case objItem.Tag
            Case "病人"
                Call mobjPati.zlRefresh(.病人ID, .挂号ID, Not tbcRegist.Item(mbyt本次就诊).Selected Or .类型 = pt已诊 Or mstr挂号单 <> .挂号单, .数据转出)
            Case "路径"
                Call mclsPath.zlRefresh(.病人ID, .挂号ID, .挂号单, mlng科室ID, .类型, .数据转出, True, mclsMipModule)
            Case "医嘱"
                Call mclsAdvices.zlRefresh(.病人ID, .挂号单, mstr挂号单 = .挂号单 And (.类型 = pt就诊 Or .类型 = pt回诊) And mlng病人ID <> 0, .数据转出, , , mclsMipModule, , mPatiInfo.路径状态, .类型)
            Case "病历"
                Call mclsEPRs.zlRefresh(.病人ID, .挂号ID, mlng科室ID, mstr挂号单 = .挂号单 And mlng科室ID = .科室ID And (.类型 = pt就诊 Or .类型 = pt回诊) And mlng病人ID <> 0, .数据转出, True)
            Case "新病历"
                Call mclsEMR.zlRefresh(.病人ID, .挂号ID, mlng科室ID, .类型, 1)
            Case "疾病报告"
                If objItem.Visible Then
                    Call mclsDisease.zlRefresh(.病人ID, .挂号ID, 1, mlng科室ID, .数据转出, mstr挂号单 = .挂号单 And mlng科室ID = .科室ID And mlng病人ID <> 0)
                End If
            Case "诊疗一览"
                Call mfrmView.zlRefresh(Me, .病人ID, mlng科室ID)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.RefreshForm(glngSys, p门诊医生站, mcolSubForm("_" & objItem.Tag), objItem.Tag, .病人ID, .挂号单, 0, .数据转出, 0, 0)
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
            .Add xtpControlButton, conMenu_File_MedRecPreview, "打印预览(&V)", -1, False
            .Add xtpControlButton, conMenu_File_MedRecPrint, "打印首页(&P)", -1, False
        End With
        '56274
        Set objControl = .Add(xtpControlButton, conMenu_File_Print_Bespeak, "重打预约挂号单(&P)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "接诊(&C)", -1, False)
    objMenu.ID = conMenu_ManagePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conmenu_Edit_Leave, "病人不就诊(&L)", -1, False): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conmenu_Edit_Wait, "病人待诊(&W)", -1, False)
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Manage_Transfer, "转诊处理(&C)"): objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Receive, "病人接诊(&Z)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Cancel, "取消接诊(&Q)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Finish, "完成接诊(&O)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Redo, "恢复接诊(&R)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBack, "需回诊(&S)"): objControl.BeginGroup = True
        objControl.IconId = conMenu_Edit_Pause
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBackCancel, "取消回诊(&R)")
        objControl.IconId = conMenu_Edit_Reuse
 
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
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)") '固有
        objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FontSize, "字体大小(&N)") '固有
        With objPopup.CommandBar.Controls
             .Add xtpControlButton, conMenu_View_FontSize_S, "小字体(&S)", -1, False '固有(小字体对应小卡片，大字体对应大卡片)
             .Add xtpControlButton, conMenu_View_FontSize_L, "大字体(&L)", -1, False '固有
        End With
        objPopup.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conmenu_View_Leave, "显示不就诊病人(&4)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "查找下一个(&N)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Busy, "诊室忙(&M)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True '固有
        Set objControl = .Add(xtpControlButton, conMenu_View_Jump, "窗格跳转(&J)")
        
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", -1, False)
    objMenu.ID = conMenu_ToolPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Community, "社区功能(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_KssAudit, "抗菌用药审核(&K)")
        objControl.IconId = 3551
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_TransAudit, "输血审核管理(&M)")
        objControl.IconId = 3551
        Set objControl = .Add(xtpControlButton, conMenu_Tool_CISMed, "临床自管药(&J)")
        objControl.IconId = 3901
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "电子病案查阅(&I)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_ExaReport, "查阅体检总检报告")
            objControl.IconId = conMenu_File_Preview
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "资料参考(&R)"): objPopup.BeginGroup = True
        objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_Reference_1, "疾病诊断参考(&D)", -1, False
            .Add xtpControlButton, conMenu_Tool_Reference_2, "诊疗措施参考(&C)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Manage_FeeItemSet, "诊疗项目费用设置(&C)")
        
        If gbln血库系统 = True Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_TraReactionRecord, "输血反应记录"): objControl.BeginGroup = True
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Positive, "阳性结果")
            objControl.IconId = 3551
        If mbln危急值 Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Critical, "危急值")
                objControl.IconId = 4113
        End If
            
        On Error Resume Next
        If mobjSquareCard.zlHealthArchiveIsSHow(Me, p门诊医生站, strFunName, "") Then
            If err.Number = 0 Then
                Set objControl = .Add(xtpControlButton, conMenu_Tool_HealthCard, strFunName)
                objControl.BeginGroup = True
                objControl.IconId = 3208
            Else
                strFunName = ""
            End If
        Else
            strFunName = ""
        End If
        On Error GoTo 0
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
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览") '固有
        Set objPopup = .Add(xtpControlPopup, conMenu_Manage_Transfer, "转诊")
        
        objPopup.ID = conMenu_Manage_Transfer
        objPopup.IconId = conMenu_Manage_Transfer
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Receive, "接诊")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Finish, "完成"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBack, "需回诊")
        objControl.IconId = conMenu_Edit_Pause
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBackCancel, "取消")
        objControl.IconId = conMenu_Edit_Reuse
        objControl.ToolTipText = "取消回诊"
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItemQAdvice, "医嘱")
        objControl.IconId = conMenu_Edit_NewItem
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItemQEpr, "病历")
        objControl.IconId = conMenu_Edit_NewParent

        Set objPopup = .Add(xtpControlPopup, conMenu_Tool_Community, "社区")
        objPopup.ID = conMenu_Tool_Community
        objPopup.IconId = conMenu_Tool_Community
        Set objControl = .Add(xtpControlButton, conMenu_Tool_TransAudit, "输血审核")
        objControl.IconId = 3551
         
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "病案查阅")
            objControl.ToolTipText = "电子病案查阅"
        
        If strFunName <> "" Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_HealthCard, strFunName)
                objControl.ToolTipText = strFunName
                objControl.IconId = 3208
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助") '固有
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出") '固有
    End With
    
    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyH, conMenu_Manage_Regist '挂号
        .Add 0, vbKeyF7, conMenu_Manage_Receive '接诊
        .Add 0, vbKeyF8, conMenu_Manage_Finish '完成就诊
        .Add FCONTROL, vbKeyB, conMenu_View_Busy '诊室状态
        
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
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL1_INSIDE_1260_2")
End Sub

Private Sub mclsAdvices_Activate()
    mblnUnRefresh = False
End Sub

Private Sub mclsAdvices_CheckInfectDisease(ByVal blnOnChek As Boolean, ByVal str疾病ID As String, ByVal str诊断ID As String, ByRef blnNo As Boolean)
'功能：根据诊断与疾病编码 得到病历编辑器
'      blnOnChek    是否只进行传染病报告卡书写检查
'      str疾病ID    疾病ID
'      str诊断ID   诊断ID
'blnNO 是否要填写传染病报告卡
    Call mclsDisease.EditDiseaseDoc(Me, mlng病人ID, mlng挂号ID, 1, mlng科室ID, str疾病ID, str诊断ID)
End Sub

Private Sub mclsAdvices_EditDiagnose(ParentForm As Object, ByVal 挂号单 As String, Succeed As Boolean)
'功能：要求输入门诊诊断
    Succeed = False
End Sub

Private Sub mclsAdvices_RequestRefresh()
'功能：医嘱子窗体要求刷新
    Call LoadPatients
End Sub

Private Sub mclsAdvices_StatusTextUpdate(ByVal Text As String)
'功能：医嘱子窗体要求更新状态栏
    Me.stbThis.Panels(2).Text = Text
End Sub

Private Sub mclsAdvices_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
'功能：查看电子病历报告
    Call gobjRichEPR.ViewDocument(Me, 报告ID, CanPrint)
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

Private Function CheckIsAskNextQueue(Optional str业务ID As String = "") As Boolean
   '------------------------------------------------------------------------------------------------------------------------
    '功能：检查医生是否允许呼叫下一个队列
    '编制：刘兴洪
    '返回:允许,返回true,否则返回False
    '日期：2010-06-09 16:48:30
    '说明：检查标准:以实际已呼叫为准(只有完成后，才能再叫)(问题:37442)
    '   取掉:候诊人数(不包含不就诊的)+已接诊的+转的<呼叫人数
    '------------------------------------------------------------------------------------------------------------------------
    Dim objItem As ListItem, lngCount As Long, rsTemp As ADODB.Recordset
    Dim strSQL As String, strLimit As String, strResult As String, arrCheck As Variant
    
    If Val(str业务ID) <> 0 Then
           strSQL = "Select Zl_QueuedateCheck([1]) as Chk From Dual"
           Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(str业务ID))
           strResult = NVL(rsTemp!chk) & "|"
           arrCheck = Split(strResult, "|")
           If Val(arrCheck(0)) <> 0 Then
              If Val(arrCheck(0)) = 1 Then
                If MsgBox(CStr(arrCheck(1)) & vbCrLf & "是否继续?", vbDefaultButton2 + vbYesNo + vbQuestion, Me.Caption) = vbNo Then
                    Exit Function
                End If
              Else
                 MsgBox CStr(arrCheck(1)), vbCritical, Me.Caption
                 Exit Function
              End If
              
           End If
    End If
    
    If mty_Queue.bln医生主动呼叫 = False Or mty_Queue.int呼叫人数 <= 0 Then
        CheckIsAskNextQueue = True: Exit Function
    End If
    '0:排队中，1:呼叫中，2：已弃号，3：暂停，4：完成就诊，6：回诊，7：已呼叫
    'mty_Queue.bln呼叫含回诊
    
    '问题:44250
    strLimit = ",0,4," & IIf(mty_Queue.bln呼叫含回诊, "", ",6,")
    strSQL = "" & _
    "   Select Count(distinct B.ID) as Count From 病人挂号记录 B ,排队叫号队列 A" & _
    "   Where A.业务ID=B.ID And A.业务类型=0  " & _
    "               And instr([4],','||A.排队状态||',')=0   And B.记录性质=1 And B.记录状态=1" & _
    "               And A.医生姓名||''=[1]   " & IIf(mty_Queue.bln呼叫含回诊, " And nvl(A.回诊序号,0) = 0", "") & _
    "               And (  (nvl(B.急诊,0)=1  and B.发生时间>=Sysdate-[3] ) or   (nvl(B.急诊,0)<>1  and B.发生时间>=Sysdate-[2] )) "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名, IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数), IIf(gint急诊挂号天数 = 0, 1, gint急诊挂号天数), strLimit)
    lngCount = Val(NVL(rsTemp!Count))

    If lngCount >= mty_Queue.int呼叫人数 Then
            MsgBox "最多只能有" & mty_Queue.int呼叫人数 & "个候诊病人,不能再进行呼叫！", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
    End If
    CheckIsAskNextQueue = True
End Function

Private Sub mobjEPRDoc_AfterSaved(lngRecordId As Long)
    With mPatiInfo
        Call mclsEPRs.zlRefresh(mlng病人ID, mlng挂号ID, mlng科室ID, mlng科室ID = .科室ID And (.类型 = pt就诊 Or .类型 = pt回诊) And mlng病人ID <> 0, .数据转出, True)
    End With
End Sub

Private Sub mobjQueue_OnQueueExecuteAfter(ByVal str业务ID As String, ByVal byt操作类型 As Byte)
    '------------------------------------------------------------------------------------------------------------------------
    '入参：byt操作类型-0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
    '------------------------------------------------------------------------------------------------------------------------
    If mty_Queue.bln医生主动呼叫 = False Then Exit Sub
    If byt操作类型 <> 1 Then Exit Sub
    
    '重新刷新病人信息
    Call LoadPatients("10001")
End Sub

Private Sub mobjQueue_OnQueueExecuteBefore(ByVal str业务ID As String, ByVal byt操作类型 As Byte, blnCancel As Boolean, strNewQueueName As String)
    Dim strSQL As String, rsTemp As ADODB.Recordset
   ' byt操作类型 -0 - 复诊, 1 - 直呼, 2 - 弃号, 3 - 暂停, 4 - 完成就诊, 5 - 广播
   
    If InStr(1, "15", byt操作类型) = 0 Then Exit Sub
    If CheckIsAskNextQueue(str业务ID) = False Then blnCancel = True: Exit Sub
    
    strSQL = "SELECT a.ID,a.No,a.病人ID,a.执行部门ID,A.执行状态 From 病人挂号记录 A,排队叫号队列 B  " & _
        "  where  a.ID=b.业务id and b.业务类型=0 and a.ID=[1] and nvl(b.排队状态,0)=0 And a.记录性质 in(1,2) And a.记录状态=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(str业务ID))
    If rsTemp.EOF Then Exit Sub
    
    '68736:刘尔旋,2014-02-18,转诊病人没有诊室信息
    If byt操作类型 = 1 Then
        If Is转诊病人(str业务ID) Then
            If CheckTransferDetail(str业务ID) = False Then
                strSQL = "ZL_病人挂号记录_更新诊室 ('" & NVL(rsTemp!NO) & "'," & Val(NVL(rsTemp!病人ID)) & ",'" & mstr接诊诊室 & "','" & UserInfo.姓名 & "',to_Date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),2)"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
            Exit Sub
        End If
    End If
    
    If InStr(1, "12", Val(NVL(rsTemp!执行状态))) > 0 Then
        '1-完成就诊,2-正在就诊:主要是第二次呼叫
        '应用于:如果已经分诊后,医生接诊后,叫病人去检查后,再复诊来呼叫
        Exit Sub
    End If
    
    '更新诊室_In Integer := 1
    strSQL = "ZL_病人挂号记录_更新诊室 ('" & NVL(rsTemp!NO) & "'," & Val(NVL(rsTemp!病人ID)) & ",'" & mstr接诊诊室 & "','" & UserInfo.姓名 & "',to_Date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),0)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
End Sub

Private Function CheckTransferDetail(strID As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------
'功能:检查该转诊病人是否有诊室信息
'入参:strID-str业务ID
'返回:True 代表转诊病人有诊室信息 False 代表转诊病人无诊室信息
'编制:刘尔旋
'日期:2014-02-18
'备注:
'-----------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHand:
    
    strSQL = "Select 诊室 From 排队叫号队列 Where 业务Id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strID)
    '排队叫号队列没有记录,不更新
    If rsTemp.EOF Then CheckTransferDetail = True: Exit Function
    If NVL(rsTemp!诊室) = "" Then CheckTransferDetail = False: Exit Function
    CheckTransferDetail = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Is转诊病人(str业务ID As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能:检查该病人是否是转诊病人并且未接收
    '入参:str业务ID
    '返回:True 代表为转诊病人 False 代表为普通病人
    '编制:王吉
    '编制日期:2012-9-14
    '问题号:51514
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHand:
    strSQL = _
    "   Select Count(ID) as 是否为转诊病人 From 病人挂号记录 Where ID=[1] And Nvl(转诊科室ID,0) <> 0 And Nvl(转诊状态,0)=0  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str业务ID)
    If rsTemp.EOF Then Is转诊病人 = False
    Is转诊病人 = rsTemp!是否为转诊病人 > 0
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mobjQueue_OnRecevieDiagnose(ByVal str业务ID As String, ByVal lng业务类型 As Long)
    '接诊:
    Dim objControl As CommandBarControl
    Dim strNO As String, strSQL As String, rsTmp As ADODB.Recordset
    Dim bln回诊 As Boolean, arrCheck As Variant, strResult As String
    Dim bln转诊病人 As Boolean '问题号:51514
    Dim datCurr As Date
        Dim blnTran As Boolean, colsql As New Collection, i As Long, intOut As Integer, str结果 As String
    
    If lng业务类型 <> 0 Then Exit Sub
    On Error GoTo errH
     If Val(str业务ID) <> 0 Then
           strSQL = "Select Zl_QueuedateCheck([1]) as Chk From Dual"
           Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(str业务ID))
           strResult = NVL(rsTmp!chk) & "|"
           arrCheck = Split(strResult, "|")
           If Val(arrCheck(0)) <> 0 Then
              If Val(arrCheck(0)) = 1 Then
                If MsgBox(CStr(arrCheck(1)) & vbCrLf & "是否继续?", vbDefaultButton2 + vbYesNo + vbQuestion, Me.Caption) = vbNo Then
                    Exit Sub
                End If
              Else
                 MsgBox CStr(arrCheck(1)), vbCritical, Me.Caption
                 Exit Sub
              End If
           End If
    End If
    strSQL = "Select 病人ID,执行人,NO,记录标志,执行状态,记录性质,姓名,门诊号,id as 挂号id,复诊,急诊 From 病人挂号记录 Where  ID=[1]  "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str业务ID)
    If rsTmp.EOF Then
        MsgBox "该病人没有挂号记录不能接诊。", vbInformation, gstrSysName
        Call LoadPatients("10101"): Exit Sub
    End If
    
    '问题号:57566
    If Check接诊控制("接诊", rsTmp!NO) = False Then Exit Sub
    
    '0-等待接诊,1-完成就诊,2-正在就诊,-1标记为不就诊
    If Val(rsTmp!执行状态) = 1 Then
        MsgBox "该病人已经完成就诊,不能再进行就诊操作。", vbInformation, gstrSysName
        Call LoadPatients("10101"): Exit Sub
    ElseIf Val(rsTmp!执行状态) = -1 Then
        MsgBox "该病人已经标记为不就诊,不能再进行就诊操作。", vbInformation, gstrSysName
        Call LoadPatients("10101"): Exit Sub
    End If
    strNO = NVL(rsTmp!NO)
    
    '转诊接收 问题号:51514
    bln转诊病人 = Is转诊病人(str业务ID)
    If bln转诊病人 Then
        strSQL = "Zl_病人挂号记录_转诊('" & strNO & "',1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        '刷新并定位病人
        If rptPati(PATI_RPT就诊).Visible Then
            Call LoadPatients("11001", PATI_RPT就诊, strNO)
        Else
            Call LoadPatients("11001")
        End If
    End If
    
    '接收预约挂号单
    datCurr = zlDatabase.Currentdate
    If Val("" & rsTmp!记录性质) = 2 Then
        If Val(zlDatabase.GetPara("挂号模式", glngSys, 9000, 1)) <> 1 And Not mobjSquareCard Is Nothing Then
            If Not mobjSquareCard.zlRegisterIncept(Me, mlngModul, strNO, mstr接诊诊室, 0, "") Then Exit Sub
        Else
            strSQL = "Zl_病人预约挂号_接收('" & strNO & "','" & mstr接诊诊室 & "',NULL,NULL,NULL,NULL,NULL,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    Else
        If Val(NVL(rsTmp!执行状态)) = 0 Then
            '正常挂号接诊
            strSQL = "zl_病人接诊(" & Val(NVL(rsTmp!病人ID)) & ",'" & strNO & "',Null,'" & UserInfo.姓名 & "','" & mstr接诊诊室 & "',0,0,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        Else
            'Zl_病人接诊
            strSQL = "Zl_病人接诊("
            '  病人id_In     病人信息.病人id%Type,
            strSQL = strSQL & "" & Val(NVL(rsTmp!病人ID)) & ","
            '  No_In         病人挂号记录.NO%Type,
            strSQL = strSQL & "'" & strNO & "',"
            '  执行部门id_In 病人挂号记录.执行部门id%Type,
            strSQL = strSQL & "" & IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID) & ","
            '  执行人_In     病人挂号记录.执行人%Type,
            strSQL = strSQL & "'" & IIf(mstr接诊医生 = "", UserInfo.姓名, mstr接诊医生) & "',"
            '  诊室_In       病人挂号记录.诊室%Type := Null,
            strSQL = strSQL & "'" & mstr接诊诊室 & "',"
            '  标记急诊_In   病人挂号记录.急诊%Type := 0,
            strSQL = strSQL & "0,"
            '  回诊_In Integer:=0
            strSQL = strSQL & "1,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            bln回诊 = True
        End If
        If mbln免挂号模式 Then
            '先判断是回诊，是则提示是否产生费用
            If Val(rsTmp!执行状态 & "") = 1 Or Val(rsTmp!执行状态 & "") = 2 Then
                str结果 = zlCommFun.ShowMsgBox("请选择", "当前病人为回诊病人，请选择该病人就诊模式？", _
                        "!继续就诊(&Y),新增就诊(&N)", Me, vbQuestion)
            Else
                '先判断是否收过费
                Call mclsReg.zlCheckRegisterNoIsCharge(strNO, intOut)
                'intOut=:-1-未找到对应的单据,0-未收费;1-挂号单已收;2-免挂号模式下，还未产生划价记录;
        '                      3-挂号单对应的收费划价单已全收费(存在多张划价单时，必须全收的);
        '                      4-挂号单对应的划价单存在部分收费)
                If intOut = 2 Then
                    str结果 = "新增就诊"
                End If
            End If
            If str结果 = "新增就诊" Then
                If mclsReg.zlBulidingPriceDataFromRegistNo(strNO, IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID), IIf(mstr接诊医生 = "", UserInfo.姓名, mstr接诊医生), , True, colsql) = False Then
                    MsgBox "挂号划价单未正确生成，不能进行接诊。", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        End If
        gcnOracle.BeginTrans: blnTran = True
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        For i = 1 To colsql.Count
            Call zlDatabase.ExecuteProcedure(colsql(i), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTran = False
    End If
        
    mstr挂号单 = strNO
    mlng病人ID = Val(NVL(rsTmp!病人ID))
    
    '门诊患者接诊消息发送
    Call ZLHIS_CIS_009(mclsMipModule, mlng病人ID, NVL(rsTmp!姓名), NVL(rsTmp!门诊号), 0, 0, NVL(rsTmp!挂号ID), NVL(rsTmp!复诊, 0), NVL(rsTmp!急诊, 0), datCurr, mlng接诊科室ID, , mstr接诊诊室, UserInfo.姓名)
    
    '刷新并定位病人
    On Error GoTo 0
    If rptPati(PATI_RPT就诊).Visible Then
        Call LoadPatients("11001", PATI_RPT就诊, strNO)
    Else
        Call LoadPatients("11001")
    End If
    '社区病人自动调用功能
    If Not gobjCommunity Is Nothing And mlngCommunityID <> 0 And mlng病人ID <> 0 And mPatiInfo.社区 <> 0 Then
        Set objControl = cbsMain.FindControl(, mlngCommunityID, , True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then objControl.Execute
        End If
    End If
    
    Call ReceiveAfterExec(bln回诊)
    '处理排队叫号队列(重新刷新)
    Call ReshDataQueue
    Call SetReceiveToday(False, 1)
    Exit Sub
errH:
        If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mobjQueue_OnSelectionChanged(ByVal blnIsCallingList As Boolean, objReportRow As Object, cbrMain As Object)
    If mty_Queue.bln医生主动呼叫 Then
        mobjQueue.zlCommandBarSet 7, blnIsCallingList Or Not mbln呼叫后接诊
    End If
     
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    Dim lngPatiID As Long
    
    If objHisPati Is Nothing Then
        lngPatiID = 0
    Else
        lngPatiID = objHisPati.病人ID
    End If
    
    Call ExecuteFindPati(False, , blnCard, lngPatiID)
End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnIsInit Then mintFindType = Index: mstrFindType = objCard.名称
End Sub

Private Sub rptPati_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    
    If Button = 2 And InStr(mstrPrivs, "病人接诊") > 0 Then
        If mPr <> -1 And Index = mintRPTIndex Then
            Set objPopup = cbsMain.ActiveMenuBar.Controls(2)
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub rptPati_RowDblClick(Index As Integer, ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
'功能：双击自动接诊或完成接诊
    Dim objControl As CommandBarControl
    
    If Index = PATI_RPT候诊 Or Index = PATI_RPT预约 Or Index = PATI_RPT就诊 Then
        If InStr(mstrPrivs, "病人接诊") > 0 Then
            If Index = PATI_RPT候诊 Or Index = PATI_RPT预约 Then
                Set objControl = cbsMain.FindControl(, conMenu_Manage_Receive, True, True)
            ElseIf Index = PATI_RPT就诊 Then
                Set objControl = cbsMain.FindControl(, conMenu_Manage_Finish, True, True)
            End If
            If Not objControl Is Nothing Then
                If objControl.Enabled Then Call cbsMain_Update(objControl) '首次执行，没有显示菜单前，事件没有执行
                If objControl.Enabled Then objControl.Execute
            End If
        End If
    End If
End Sub

Private Sub rptPati_SelectionChanged(Index As Integer)
    Call RptItemClick(Index)
End Sub
 
Private Sub tbcRegist_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'功能：就诊选择
'说明：tbcRegist.Tag 中记录上一次卡片的选择情况
    Dim objControl As CommandBarControl
    
    If Item.Tag = "" Then Exit Sub '初始添卡时,还没赋值
    If Item.Tag = "诊疗一览" Then
        If tbcRegist.Tag <> "诊疗一览" And tbcSub.Selected.Tag <> "病人" Then
            Call SubWinDefCommandBar(Item)
        Else
            Set mfrmActive = mcolSubForm("_" & Item.Tag)
        End If
        Call SubWinRefreshData(Item)
        Call UCPatiVitalSigns.ClearTxtToolTipText
        UCPatiVitalSigns.ControlLock = True
        UCPatiVitalSigns.TextBackColor = vbButtonFace
        Call UCPatiVitalSigns.SetUseType(False)
        If Visible Then mfrmActive.SetFocus
    ElseIf Item.Tag = "更多" Then
        tbcRegist.Item(mbyt本次就诊).Selected = True
        Set objControl = cbsMain.FindControl(, conMenu_Tool_Archive, True, True) '电子病案查阅
        If Not objControl Is Nothing Then
            If objControl.Enabled Then Call cbsMain_Update(objControl) '首次执行，没有显示菜单前，事件没有执行
            If objControl.Enabled Then objControl.Execute
        End If
    Else
        If Val(Item.Tag) <> 0 Then
            If tbcRegist.Tag = "诊疗一览" And tbcSub.Selected.Tag <> "病人" Then
                Call SubWinDefCommandBar(tbcSub.Selected)
            Else
                Set mfrmActive = mcolSubForm("_" & tbcSub.Selected.Tag)
            End If
            If Val(Item.Tag) > 0 Then
                Call LoadRegist(Val(Item.Tag))
            End If
        End If
    End If
    tbcRegist.Tag = Item.Tag
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtInfo(Index))
End Sub

Private Sub LoadPatiInfo(ByVal lng挂号id As Long)
'功能：选择某次历史就诊记录时，读取相关的病人信息
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim lngidx As Long
    Dim i As Long
    On Error GoTo errH
    strSQL = "Select B.Id,B.NO,B.门诊号,B.姓名,B.性别,B.年龄,A.出生日期,b.医疗付款方式,A.职业," & _
        "   a.就诊卡号,A.费别,A.险类,A.医保号,B.急诊,A.结算模式,B.发生时间,B.执行人,B.执行状态,B.执行时间," & _
        "   B.执行部门ID as 科室ID,B.诊室,B.社区,D.社区号,C.名称 as 科室,B.复诊,B.摘要," & _
        "   A.身份证号,A.监护人,A.家庭地址,A.家庭电话,A.工作单位,A.合同单位id,A.单位电话,B.发病时间,B.发病地址," & _
        "   A.民族,A.国籍,A.区域,A.婚姻状况,A.家庭地址邮编,A.单位邮编,A.出生地点,B.传染病上传,A.其他证件,a.户口地址," & _
        "   A.户口地址邮编,a.籍贯,a.email,a.qq,A.病人类型,a.病人ID,B.路径状态,nvl(g.号类,E.号类) as 号类 " & _
        " From 病人信息 A,病人挂号记录 B,部门表 C,病人社区信息 D,挂号安排 E,临床出诊记录 f,临床出诊号源 g" & _
        " Where A.病人ID=B.病人ID And B.ID=[1] And B.执行部门ID=C.ID" & _
        " And B.病人ID=D.病人ID(+) And B.社区=D.社区(+) And B.号别=E.号码(+) and b.出诊记录id=f.id(+) and f.号源id=g.id(+)"
        '按ID读取挂号记录，不用加记录性质、状态的条件
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng挂号id)
    
    For i = 0 To lblLink清除
        lblLink(i).ForeColor = &HC00000
    Next
    
    lblLink(lblLink修改).ForeColor = IIf(InStr(GetInsidePrivs(9003), "基本信息调整") > 0, &HC00000, &HC0C0C0)  'p病人信息公共部件
    
    Call ReadPatPricture(Val(rsTmp!病人ID & ""), imgLoad)
    
    If imgLoad.Picture = 0 Then
        imgPatient.Picture = imgDefual.Picture
        picPatient.Tag = ""
    Else
        imgPatient.Picture = imgLoad.Picture
        picPatient.Tag = "有"
    End If
    
    txtInfo(txtInfo姓名).Text = rsTmp!姓名 & ""
    txtInfo(txtInfo姓名).ToolTipText = rsTmp!姓名 & ""
    '显示病人颜色
    If Not IsNull(rsTmp!险类) And NVL(rsTmp!病人类型) = "" Then
        txtInfo(txtInfo姓名).ForeColor = &HC0&
    Else
        txtInfo(txtInfo姓名).ForeColor = zlDatabase.GetPatiColor(NVL(rsTmp!病人类型))
    End If
    
    txtInfo(txtInfo性别).Text = rsTmp!性别 & ""
    txtInfo(txtInfo性别).ToolTipText = txtInfo(txtInfo性别).Text
    txtInfo(txtInfo年龄).Text = rsTmp!年龄 & ""
    txtInfo(txtInfo年龄).ToolTipText = txtInfo(txtInfo年龄).Text
    txtInfo(txtInfo出生日期).Text = Format(rsTmp!出生日期 & "", "yyyy-MM-dd")
    txtInfo(txtInfo出生日期).ToolTipText = txtInfo(txtInfo出生日期).Text
    txtInfo(txtInfo号类).Text = rsTmp!号类 & ""
    txtInfo(txtInfo号类).ToolTipText = txtInfo(txtInfo号类).Text
    txtInfo(txtInfo就诊卡号).Text = rsTmp!就诊卡号 & ""
    txtInfo(txtInfo就诊卡号).ToolTipText = txtInfo(txtInfo就诊卡号).Text
    txtInfo(txtInfo医保卡号).Text = rsTmp!医保号 & ""
    txtInfo(txtInfo医保卡号).ToolTipText = txtInfo(txtInfo医保卡号).Text
    txtInfo(txtInfo摘要).Text = rsTmp!摘要 & ""
    txtInfo(txtInfo摘要).ToolTipText = txtInfo(txtInfo摘要).Text
    
    With cboBillType
        lngidx = -1
        For i = 0 To .ListCount
            If InStr(.List(i) & "", rsTmp!费别 & "") > 0 Then
                .ToolTipText = rsTmp!费别 & ""
                lngidx = i
                Exit For
            End If
        Next
    End With
    If lngidx <> -1 Then
        Call Cbo.SetIndex(cboBillType.hwnd, lngidx)
    End If
    
    With cboPayType
        lngidx = -1
        For i = 0 To .ListCount
            If InStr(.List(i) & "-", "-" & rsTmp!医疗付款方式 & "-") > 0 Then
                .ToolTipText = rsTmp!医疗付款方式 & ""
                lngidx = i
                Exit For
            End If
        Next
    End With
    
    If lngidx <> -1 Then
        Call Cbo.SetIndex(cboPayType.hwnd, lngidx)
    End If
    
    With rsTmp
        '病人信息
        If mintActive = pt转诊 Then
            mPatiInfo.类型 = pt转诊
        Else
            mPatiInfo.类型 = Decode(NVL(!执行状态, 0), 0, 0, 2, 1, 1, 2)
        End If
        
        mPatiInfo.门诊号 = NVL(!门诊号)
        mPatiInfo.病人ID = !病人ID
        mPatiInfo.挂号ID = !ID
        mPatiInfo.挂号单 = !NO
        mPatiInfo.科室ID = !科室ID
        mPatiInfo.诊室 = NVL(!诊室)
        mPatiInfo.社区 = NVL(!社区, 0)
        mPatiInfo.社区号 = NVL(!社区号)
        mPatiInfo.挂号时间 = !发生时间
        mPatiInfo.性别 = "" & !性别
        mPatiInfo.婚姻状况 = "" & !婚姻状况
        
        mPatiInfo.民族 = "" & !民族
        mPatiInfo.国籍 = "" & !国籍
        mPatiInfo.区域 = "" & !区域
        mPatiInfo.出生地点 = "" & !出生地点
        mPatiInfo.传染病上传 = Val("" & !传染病上传)
        mPatiInfo.家庭地址邮编 = "" & !家庭地址邮编
        mPatiInfo.单位邮编 = "" & !单位邮编
        mPatiInfo.其他证件 = "" & !其他证件
        mPatiInfo.户口地址 = "" & !户口地址
        mPatiInfo.户口地址邮编 = "" & !户口地址邮编
        mPatiInfo.籍贯 = "" & !籍贯
        mPatiInfo.Email = "" & !Email
        mPatiInfo.QQ = "" & !QQ
        mPatiInfo.复诊 = Val(!复诊 & "")
        mPatiInfo.急诊 = Val(!急诊 & "")
        mPatiInfo.路径状态 = Val(!路径状态 & "")
        lblUrg.Visible = Val(!急诊 & "") <> 0
        lblRec.Visible = Val(!结算模式 & "") <> 0
        
        If mPatiInfo.类型 = pt已诊 Then
            mPatiInfo.数据转出 = zlDatabase.NOMoved("病人挂号记录", !NO)
        Else
            mPatiInfo.数据转出 = False
        End If
    End With
    If mintRPTIndex = PATI_RPT就诊 Or mintRPTIndex = PATI_RPT回诊 Then
        If InStr("," & mstr挂号IDs & ",", "," & rsTmp!ID & ",") = 0 Then
            mstr挂号IDs = mstr挂号IDs & "," & rsTmp!ID
        End If
    End If
    If tbcRegist.Selected.Index = 0 And mbyt本次就诊 = 0 Or tbcRegist.Selected.Index = 1 And mbyt本次就诊 = 1 Then
        UCPatiVitalSigns.ControlLock = False
        UCPatiVitalSigns.TextBackColor = vbWindowBackground
        Call UCPatiVitalSigns.SetUseType(True)
    Else
        Call UCPatiVitalSigns.ClearTxtToolTipText
        UCPatiVitalSigns.ControlLock = True
        UCPatiVitalSigns.TextBackColor = vbButtonFace
        Call UCPatiVitalSigns.SetUseType(False)
    End If
    Call UCPatiVitalSigns.LoadPatiVitalSigns(mPatiInfo.病人ID, lng挂号id)
    Call UCPatiVitalSigns.TxtAlignment(2)
    txtInfo(txtInfo摘要).Locked = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadRegist(ByVal lng挂号id As Long)
'功能：选择某次历史就诊记录时，读取相关的病人信息
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim lngidx As Long
    Dim i As Long
    
    If lng挂号id <= 0 Then
        '按当前列表无数据刷新子窗体
        Call ClearPatiInfo
        '刷新子窗体数据
        Call SubWinRefreshData(tbcSub.Selected)
        
        Exit Sub
    End If
    On Error GoTo errH
    Call LoadPatiInfo(lng挂号id)
    
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

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
'功能：身份证识别成功后激活
    mstrIDCard = strID
    If mstrFindType = "二代身份证" Then
        PatiIdentify.Text = mstrIDCard
    Else
        PatiIdentify.Text = "" '否则清除(目前是在已清除情况下才能激活)。
    End If
    Call ExecuteFindPati(False, mstrIDCard)
End Sub

Private Function CheckHaveAdvice(ByVal lng病人ID As Long, ByVal str挂号单 As String) As Boolean
'功能：判断病人是否开了医嘱
    Dim strSQL As String
    Dim rsTmp As Recordset
    
    On Error GoTo errH
    strSQL = "select 1 from 病人医嘱记录 where 病人ID=[1] and 挂号单=[2] and rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, str挂号单)
    CheckHaveAdvice = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'功能：刷新子窗体界面及数据
'说明：仅在人为切换界面卡片激活
    Dim objControl As CommandBarControl
    Dim Index As Long, objItem As TabControlItem
    
    If mblnTabTmp Then Exit Sub
    
    If Item.Tag = "" Then Exit Sub '初始添卡时,还没赋值
     
    If Item.Handle = picTmp.hwnd Then
        Index = Item.Index
        mblnTabTmp = True
        Screen.MousePointer = 11
        On Error GoTo errH
        Select Case Item.Tag
            Case "病人"
                Set objItem = tbcSub.InsertItem(Index, "病人信息", mcolSubForm("_病人").hwnd, 0)
                objItem.Tag = "病人"
                mstrPreSubTab = "病人"
            Case "路径"
                Set objItem = tbcSub.InsertItem(Index, "临床路径", mcolSubForm("_路径").hwnd, 0)
                objItem.Tag = "路径"
            Case "病历"
                Set objItem = tbcSub.InsertItem(Index, "病历信息", mcolSubForm("_病历").hwnd, 0)
                objItem.Tag = "病历"
            Case "新病历"
                Set objItem = tbcSub.InsertItem(Index, "电子病历", mcolSubForm("_新病历").hwnd, 0)
                objItem.Tag = "新病历"
            Case "疾病报告"
                Set objItem = tbcSub.InsertItem(Index, "疾病报告", mcolSubForm("_疾病报告").hwnd, 0)
                objItem.Tag = "疾病报告"
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
    
    '自动新增一份门诊/急诊/复诊病历/如果是医嘱，则新增医嘱，先判断没有医嘱再新增
    If Item.Tag = "病历" And mlng自动进行 = 1 Then
        mblnUnRefresh = True
        Call mclsEPRs.zlOpenDefaultEPR(mstr挂号单)
        '因为执行命令的是非模态窗体，所以在mclsAdvices和mclsEPRs的active中设置 mblnUnRefresh = False
    ElseIf Item.Tag = "医嘱" And mlng自动进行 = 2 Then
        If CheckHaveAdvice(mlng病人ID, mstr挂号单) = False Then
            cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
            Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
            If Not objControl Is Nothing Then
                If objControl.Enabled Then objControl.Execute
            End If
        End If
    End If
    
    If mstrPreSubTab = "病人" And Not mobjPati Is Nothing Then
        Call mobjPati.UpdateLastItem
    End If
    mstrPreSubTab = Item.Tag
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picPati.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picMsg.hwnd
    End If
End Sub

Private Sub picHZ_Resize()
    Call RPTResize(picHZ, 0)
End Sub

Private Sub picJZ_Resize()
    Call RPTResize(picJZ, 1)
End Sub

Private Sub picMsg_Resize()
    Call RPTResize(picMsg, 2)
End Sub

Private Sub picHUIZ_Resize()
    Call RPTResize(picHUIZ, 3)
End Sub

Private Sub picYy_Resize()
    Call RPTResize(picYy, 4)
End Sub


Private Sub picYZ_Resize()
    Dim lngTmp As Long
    On Error Resume Next
    lblSeeTim.Left = 100
    cboSelectTime.Left = lblSeeTim.Left + lblSeeTim.Width + 15
    cmdOtherFilter.Left = cboSelectTime.Left + cboSelectTime.Width + 50
    rptPati(PATI_RPT已诊).Top = cboSelectTime.Top + cboSelectTime.Height + 30
    rptPati(PATI_RPT已诊).Left = 0
    rptPati(PATI_RPT已诊).Width = picYZ.Width
    lngTmp = picYZ.Height - rptPati(PATI_RPT已诊).Top
    If mbytSize = 0 Then
        If lngTmp < 1010 Then
            lngTmp = 1010
        End If
    Else
        If lngTmp < 1130 Then
            lngTmp = 1130
        End If
    End If
    rptPati(PATI_RPT已诊).Height = lngTmp
End Sub

Private Sub RPTResize(ByVal objC As Object, ByVal lngID As Long)
'功能：设置表格控件大小
'参数：objC 上层容器控件  lngId 0-候诊，1-就诊，2-消息，3-回诊，4-预约
    Dim lngTmp As Long
    Dim objRpt As Object
    On Error Resume Next
    
    Select Case lngID
    Case 0
        Set objRpt = rptPati(0)
    Case 1
        Set objRpt = rptPati(1)
    Case 2
        Set objRpt = rptNotify
    Case 3
        Set objRpt = rptPati(2)
    Case 4
        Set objRpt = rptPati(4)
    End Select
    
    lngTmp = objC.Height
    
    If mbytSize = 0 Then
        If lngTmp < 1010 Then
            lngTmp = 1010
        End If
    Else
        If lngTmp < 1130 Then
            lngTmp = 1130
        End If
    End If
    
    objRpt.Top = 0
    objRpt.Left = 0
    objRpt.Width = objC.Width
    objRpt.Height = lngTmp
    
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long, lngTopPanelHeight As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
 
    With picBasisNew
        .Visible = True
        .Height = IIf(mbytSize = 0, 1000, 1080)
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        lngTopPanelHeight = .Height
    End With
    
    With picMore
        .BackColor = vbButtonFace
        .Visible = mblnPatiDetail
        .Left = lngLeft
        .Top = lngTop + lngTopPanelHeight
        .Width = lngRight - lngLeft
        .Height = IIf(mbytSize = 0, 600, 800)
    End With
    
    txtInfo(txtInfo摘要).BackColor = vbButtonFace
    UCPatiVitalSigns.BackColor = vbButtonFace
    UCPatiVitalSigns.LblBackColor = vbButtonFace
    
    lngTopPanelHeight = IIf(mblnPatiDetail, picMore.Height, 0) + picBasisNew.Height
    
    With Me.tbcRegist
        .Left = lngLeft: .Width = lngRight - lngLeft
        .Top = lngTop + lngTopPanelHeight: .Height = lngBottom - lngTop - lngTopPanelHeight
    End With
    
    With Me.tbcSub
        .Left = 0
        .Top = 0
        .Height = picRegist.Height
        .Width = picRegist.Width
    End With
    
    With Me.fraRoom
        .Visible = Me.stbThis.Visible
        .Left = Me.stbThis.Panels(4).Left + 60: .Top = Me.stbThis.Top + 60
    End With
End Sub

Private Sub picRegist_Resize()
    On Error Resume Next
    With Me.tbcSub
        .Left = 0
        .Top = 0
        .Height = picRegist.ScaleHeight ' Height
        .Width = picRegist.ScaleWidth
    End With
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
 
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    Dim blnSetup As Boolean
    
    If Not mobjMsg Is Nothing Then
        Call mobjMsg.CloseAirBubble
    End If
    
    mblnMsgOk = False: mblnFirstMsg = False
    mblnIsInit = False 'PatiIdentify初始化标志
    
    blnSetup = InStr(";" & mstrPrivs & ";", ";参数设置;") > 0
    Call zlDatabase.SetPara("病人查找方式", mintFindType, glngSys, p门诊医生站, blnSetup)

    If Not tbcSub.Selected Is Nothing Then
        Call zlDatabase.SetPara("医护功能", tbcSub.Selected.Tag, glngSys, p门诊医生站, blnSetup)
    End If
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, dkpMain.SaveStateToString)
    End If

    '公共部件固定按第一个控件的样式保存，工作站部件如果第一个是打印，则固定是图标样式,所以需恢复为其它按钮的样式
    If Me.Visible Then  'Form_load中退出时不处理
        cbsMain(2).Controls(1).Style = cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style
        Call SaveWinState(Me, App.ProductName)
    End If
    
    '单独存一次，用的控件数组
    For i = 0 To rptPati.Count - 1
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\ReportControl", "rptPati" & "_" & i, rptPati(i).SaveSettings)
    Next
    
    mstrIDCard = ""
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled False
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    Set mobjSquareCard = Nothing

    '--关闭所有排队的窗体
    If Not mobjQueue Is Nothing Then
        Call mobjQueue.CloseWindows
        Set mobjQueue = Nothing
    End If
    Set mobjQueueList = Nothing
    Set mobjCallList = Nothing

    '强行Unload,不然不会激活子窗体的事件
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next
    Set mclsAdvices = Nothing
    Set mobjKernel = Nothing
    Set mclsEMR = Nothing
    Set mclsPath = Nothing
    Set mclsEPRs = Nothing
    Set mrsAller = Nothing
    Set mobjEPRDoc = Nothing
    If Not mfrmActive Is Nothing Then
        Unload mfrmActive
    End If
    Set mfrmActive = Nothing
    Set gobjPublicPacs = Nothing
    
    If Not mfrmView Is Nothing Then
        Unload mfrmView
    End If
    Set mfrmView = Nothing
    Set mclsReg = Nothing
    mPatiInfo.挂号ID = 0
    '问题号:57566
    mlng接诊控制 = 0
    mlng提前接收时间 = 0
    If Not (mclsMipModule Is Nothing) Then
        mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    Set mobjMsg = Nothing
    Set mobjPatient = Nothing
    Set mclsDis = Nothing
    Set mclsMsg = Nothing
    Set mrsMsg = Nothing
    
    mbln危急值 = False
    Set mclsDisease = Nothing
    
    If Not mobjPati Is Nothing Then
        Unload mobjPati
    End If
    Set mobjPati = Nothing
End Sub

Private Sub lblRoom_Click()
    Call SetRoomState(lblRoom.BackColor = COLOR_FREE)
End Sub

Private Sub RptItemClick(ByVal Index As Integer)
'功能:数据处理列表中病人的切换
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim intCount As Integer
    Dim i As Long, j As Long, k As Long
    Dim strNO As String
    Dim str发生时间 As String
    Dim lng挂号id As Long
    Dim strCaption As String
    Dim objItem As TabControlItem
    Dim strRegTag As String
    Dim blnDo As Boolean
    Dim str身份证号 As String
    Dim strTmp As String
    Dim str病人IDs As String
        
    On Error GoTo errH
    
    If rptPati(Index).SelectedRows.Count <= 0 Then Exit Sub
    For k = PATI_RPT候诊 To PATI_RPT预约
        For i = 0 To rptPati(k).Rows.Count - 1
            For j = 0 To rptPati(k).Columns.Count - 1
                If rptPati(k).Rows(i).Record.Item(j).Bold Then
                    rptPati(k).Rows(i).Record.Item(j).Bold = False
                    rptPati(k).Rows(i).Record.Item(j).BackColor = rptPati(k).PaintManager.BackColor
                    blnDo = True
                End If
            Next
        Next
        If blnDo Then
            blnDo = False
            rptPati(k).Redraw
        End If
    Next
    blnDo = False
    With rptPati(Index).SelectedRows(0).Record
        Select Case Index
        Case PATI_RPT候诊, PATI_RPT预约
            strNO = .Item(COL_HZ_NO).Value
            mstr挂号单 = strNO
            mlng病人ID = Val(.Item(COL_HZ_病人ID).Value)
            mlng科室ID = Val(.Item(COL_HZ_执行部门ID).Value)
            str发生时间 = .Item(COL_HZ_发生时间).Value
            str身份证号 = .Item(COL_HZ_身份证号).Value
            strNO = .Item(COL_HZ_标识).Value
            If strNO = "预" Then
                mintActive = pt预约
            ElseIf strNO = "转" Then
                mintActive = pt转诊
            Else
                mintActive = pt候诊
            End If
             
        Case PATI_RPT就诊, PATI_RPT回诊
            strNO = .Item(COL_JZ_NO).Value
            mstr挂号单 = strNO
            mlng病人ID = Val(.Item(COL_JZ_病人ID).Value)
            mlng科室ID = Val(.Item(COL_JZ_执行部门ID).Value)
            str发生时间 = .Item(COL_JZ_发生时间).Value
            str身份证号 = .Item(COL_JZ_身份证号).Value
            If Index = PATI_RPT回诊 Then
                mintActive = pt回诊
            Else
                mintActive = pt就诊
            End If
        Case PATI_RPT已诊
            strNO = .Item(COL_YZ_NO).Value
            mstr挂号单 = strNO
            mlng病人ID = Val(.Item(COL_YZ_病人ID).Value)
            mlng科室ID = Val(.Item(COL_YZ_执行部门ID).Value)
            str发生时间 = .Item(COL_YZ_发生时间).Value
            str身份证号 = .Item(COL_YZ_身份证号).Value
            mintActive = pt已诊
        End Select
        mintRPTIndex = Index
        For i = 0 To rptPati(Index).Columns.Count - 1
           .Item(i).Bold = True
           .Item(i).BackColor = rptPati(Index).PaintManager.HighlightBackColor
        Next
        mPr = rptPati(Index).SelectedRows(0).Index
    End With
    
    For i = 0 To 4
        If rptPati(i).SelectedRows.Count > 0 Then
            rptPati(i).SelectedRows(0).Selected = False
        End If
    Next
    
    If mstr挂号单 = mstrPrePati Then Exit Sub
    mstrPrePati = mstr挂号单
            
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
        '以前的单病人模式
        '读取"历史的"就诊记录
        strSQL = "Select A.ID,A.NO,A.发生时间 as 时间,B.名称 as 科室,a.执行人 From 病人挂号记录 A,部门表 B" & _
            " Where A.执行部门ID=B.ID And A.病人ID=[1] And A.发生时间<=[2] And A.记录性质=1 And A.记录状态=1 Order by A.发生时间 Desc,a.接收时间 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, CDate(str发生时间))
    Else
        '通过身份证号找出多个病的情况
        strSQL = "Select A.ID,A.NO,A.发生时间 as 时间,B.名称 as 科室,a.执行人 From 病人挂号记录 A,部门表 B" & _
            " Where A.执行部门ID=B.ID And A.病人ID In (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X)" & _
            " And A.发生时间<=[2] And A.记录性质=1 And A.记录状态=1 Order by A.发生时间 Desc,a.接收时间 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str病人IDs, CDate(str发生时间))
    End If
    strRegTag = tbcRegist.Selected.Caption
    
    '先隐藏卡片
    For i = mbyt本次就诊 + 1 To tbcRegist.ItemCount - 1
        tbcRegist.Item(i).Visible = False
    Next
    
    i = 0: blnDo = False
    For j = 1 To rsTmp.RecordCount
        If Not blnDo Then
            If mstr挂号单 = rsTmp!NO & "" Then
                mlng挂号ID = Val(rsTmp!ID & "")
                blnDo = True
            End If
        End If
        If blnDo Then
            i = i + 1
            If i = 5 Then
                tbcRegist.Item(mbyt本次就诊 + i - 1).Visible = True
            ElseIf i > 1 And i < 5 Then
                strCaption = Format(rsTmp!时间, "YYMMdd") & "/" & rsTmp!科室 & "/" & rsTmp!执行人
                tbcRegist.Item(mbyt本次就诊 + i - 1).Caption = strCaption
                tbcRegist.Item(mbyt本次就诊 + i - 1).Tag = Val(rsTmp!ID & "")
                tbcRegist.Item(mbyt本次就诊 + i - 1).Visible = True
            End If
            If rsTmp!NO = mstr挂号单 Then
                lng挂号id = Val(rsTmp!ID & "")
            Else
                If lng挂号id = 0 Then
                    lng挂号id = Val(rsTmp!ID & "")
                End If
            End If
            '当日多科就诊
            If Format(rsTmp!时间, "yyyy-MM-dd") = Format(str发生时间, "yyyy-MM-dd") Then
                intCount = intCount + 1
            End If
        End If
        rsTmp.MoveNext
    Next
   
    If strRegTag = "诊疗一览" Then
        tbcRegist.Item(0).Selected = True
        tbcRegist.Item(0).Tag = "诊疗一览"
        Call LoadPatiInfo(lng挂号id)
    Else
        If mbyt本次就诊 = 1 Then
            tbcRegist.Item(0).Tag = "诊疗一览"
        End If
        tbcRegist.Item(mbyt本次就诊).Tag = ""
        tbcRegist.Item(mbyt本次就诊).Selected = True
        tbcRegist.Item(mbyt本次就诊).Tag = mlng挂号ID
    End If

    Call tbcRegist_SelectedChanged(tbcRegist.Selected)
    
    lblMore.Visible = intCount > 1 And mintActive = pt就诊
    rptPati(Index).SetFocus
    LockWindowUpdate 0
    
    Exit Sub
errH:
    LockWindowUpdate 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Index = 4 Then
        Call SetRoomState(lblRoom.BackColor = COLOR_FREE)
    End If
End Sub

Private Sub tbcSub_GotFocus()
    On Error Resume Next
    If Not mfrmActive Is Nothing Then mfrmActive.SetFocus
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngidx As Long, i As Long
    For i = 0 To 1
        If i = 0 Then
            lngidx = PATI_RPT候诊
        Else
            lngidx = PATI_RPT预约
        End If
        With rptPati(lngidx)
            Set objCol = .Columns.Add(COL_HZ_标识, "", 18, True)
            Set objCol = .Columns.Add(COL_HZ_门诊号, "门诊号", 60, True)
            Set objCol = .Columns.Add(COL_HZ_姓名, "姓名", 60, True)
            Set objCol = .Columns.Add(COL_HZ_挂号时间, "挂号时间", 80, True)
            Set objCol = .Columns.Add(COL_HZ_性别, "性别", 30, True)
            Set objCol = .Columns.Add(COL_HZ_年龄, "年龄", 40, True)
            Set objCol = .Columns.Add(COL_HZ_急, "急", 20, True)
            Set objCol = .Columns.Add(COL_HZ_复, "复", 20, True)
            Set objCol = .Columns.Add(COL_HZ_NO, "挂号单", 60, True)
            Set objCol = .Columns.Add(COL_HZ_社区, "社区", 30, True)
            Set objCol = .Columns.Add(COL_HZ_就诊诊室, "就诊诊室", 60, True)
            Set objCol = .Columns.Add(COL_HZ_就诊医生, "就诊医生", 60, True)
            Set objCol = .Columns.Add(COL_HZ_序号, "序号", 60, True)
            Set objCol = .Columns.Add(COL_HZ_分诊时间, "分诊时间", 80, True)
            Set objCol = .Columns.Add(COL_HZ_就诊卡号, "就诊卡号", 60, True)
            Set objCol = .Columns.Add(COL_HZ_病人类型, "病人类型", 60, True)
            Set objCol = .Columns.Add(COL_HZ_转诊状态, "转诊状态", 60, True)
            Set objCol = .Columns.Add(COL_HZ_预约医生, "预约医生", 60, True)
            Set objCol = .Columns.Add(COL_HZ_预约时间, "预约时间", 80, True)
            Set objCol = .Columns.Add(COL_HZ_身份证号, "身份证号", 60, True)
            Set objCol = .Columns.Add(COL_HZ_号类, "号类", 30, True)
            Set objCol = .Columns.Add(COL_HZ_病人科室, "病人科室", 60, True)
            
            Set objCol = .Columns.Add(COL_HZ_病人ID, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_发生时间, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_执行部门ID, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_执行人, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_状态, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_IC卡号, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_记录标志, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_执行状态, "", 0, False): objCol.Visible = False
            
            
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
        End With
    Next
    
    
    For i = 0 To 1
        If i = 0 Then
            lngidx = PATI_RPT就诊
        Else
            lngidx = PATI_RPT回诊
        End If
        With rptPati(lngidx)
            Set objCol = .Columns.Add(COL_JZ_标识, "", 18, True)
            Set objCol = .Columns.Add(COL_JZ_门诊号, "门诊号", 60, True)
            Set objCol = .Columns.Add(COL_JZ_姓名, "姓名", 60, True)
            Set objCol = .Columns.Add(COL_JZ_就诊时间, "就诊时间", 80, True)
            Set objCol = .Columns.Add(COL_JZ_性别, "性别", 30, True)
            Set objCol = .Columns.Add(COL_JZ_年龄, "年龄", 40, True)
            Set objCol = .Columns.Add(COL_JZ_急, "急", 20, True)
            Set objCol = .Columns.Add(COL_JZ_复, "复", 20, True)
            Set objCol = .Columns.Add(COL_JZ_NO, "挂号单", 60, True)
            Set objCol = .Columns.Add(COL_JZ_社区, "社区", 30, True)
            Set objCol = .Columns.Add(COL_JZ_就诊卡号, "就诊卡号", 60, True)
            Set objCol = .Columns.Add(COL_JZ_病人类型, "病人类型", 60, True)
            Set objCol = .Columns.Add(COL_JZ_转诊状态, "转诊状态", 60, True)
            Set objCol = .Columns.Add(COL_JZ_传染病, "传染病", 60, True)
            Set objCol = .Columns.Add(COL_JZ_号类, "号类", 30, True)
            Set objCol = .Columns.Add(COL_JZ_病人科室, "病人科室", 60, True)
            
            Set objCol = .Columns.Add(COL_JZ_病人ID, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_JZ_发生时间, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_JZ_执行部门ID, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_JZ_执行人, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_JZ_状态, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_JZ_身份证号, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_JZ_IC卡号, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_JZ_记录标志, "", 0, False): objCol.Visible = False
            
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
        End With
    Next
    
    With rptPati(PATI_RPT已诊)
        Set objCol = .Columns.Add(COL_YZ_NO, "挂号单", 60, True)
        Set objCol = .Columns.Add(COL_YZ_门诊号, "门诊号", 60, True)
        Set objCol = .Columns.Add(COL_YZ_姓名, "姓名", 60, True)
        Set objCol = .Columns.Add(COL_YZ_性别, "性别", 30, True)
        Set objCol = .Columns.Add(COL_YZ_年龄, "年龄", 40, True)
        Set objCol = .Columns.Add(COL_YZ_急, "急", 20, True)
        Set objCol = .Columns.Add(COL_YZ_复, "复", 20, True)
        Set objCol = .Columns.Add(COL_YZ_社区, "社区", 30, True)
        Set objCol = .Columns.Add(COL_YZ_时间, "时间", 120, True)
        Set objCol = .Columns.Add(COL_YZ_就诊医生, "就诊医生", 60, True)
        Set objCol = .Columns.Add(COL_YZ_就诊卡号, "就诊卡号", 60, True)
        Set objCol = .Columns.Add(COL_YZ_病人类型, "病人类型", 60, True)
        Set objCol = .Columns.Add(COL_YZ_号类, "号类", 30, True)
        Set objCol = .Columns.Add(COL_YZ_病人科室, "病人科室", 60, True)
        Set objCol = .Columns.Add(COL_YZ_西医诊断, "西医诊断", 120, True)
        Set objCol = .Columns.Add(COL_YZ_中医诊断, "中医诊断", 120, True)
        
        Set objCol = .Columns.Add(COL_YZ_病人ID, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_YZ_发生时间, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_YZ_执行部门ID, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_YZ_执行人, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_YZ_身份证号, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_YZ_IC卡号, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_YZ_记录标志, "", 0, False): objCol.Visible = False
        
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
    End With
    
    With rptNotify
        Set objCol = .Columns.Add(c_图标, "", 18, True): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(C_病人ID, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_No, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_姓名, "姓名", 60, True)
        Set objCol = .Columns.Add(C_门诊号, "门诊号", 62, True)
        Set objCol = .Columns.Add(C_就诊时间, "就诊时间", 60, True)
        Set objCol = .Columns.Add(C_状态, "状态", 150, True)
         
        Set objCol = .Columns.Add(C_消息, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_序号, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_日期, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_业务, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_挂号Id, "", 0, False): objCol.Visible = False
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
    
End Sub

Private Sub InitCondFilter()
    Dim curDate As Date, intDay As Long
    Dim intStart As Long
    
    cboSelectTime.Clear
    
    With cboSelectTime
        .AddItem "今天"
        .ItemData(.NewIndex) = 0
        .AddItem "昨天(含今天)"
        .ItemData(.NewIndex) = 1
        .AddItem "一周内"
        .ItemData(.NewIndex) = 7
        .AddItem "[指定...]"
        .ItemData(.NewIndex) = -1
    End With
    
    '已诊病人时间范围
    curDate = zlDatabase.Currentdate
    
    intStart = Val(zlDatabase.GetPara("已诊病人结束间隔", glngSys, p门诊医生站, "0", Array(lblSeeTim, cboSelectTime), InStr(";" & mstrPrivs & ";", ";参数设置;") > 0))
    If lblSeeTim.ForeColor <> vbBlue Then
        '私有参数
        mvCondFilter.End = Format(curDate, "yyyy-MM-dd 23:59:59")
        mvCondFilter.Begin = Format(mvCondFilter.End, "yyyy-MM-dd 00:00:00")
        If cboSelectTime.ListCount > 0 Then cboSelectTime.ListIndex = 0
    Else
        '系统参数(恢复成管理员设置的值，防止通方)
        mvCondFilter.End = Format(curDate + intStart, "yyyy-MM-dd 23:59:59")
        intDay = Val(zlDatabase.GetPara("已诊病人开始间隔", glngSys, p门诊医生站, "7", Array(lblSeeTim, cboSelectTime), InStr(";" & mstrPrivs & ";", ";参数设置;") > 0))
        If intDay > 7 Then intDay = 7
        mvCondFilter.Begin = Format(mvCondFilter.End - intDay, "yyyy-MM-dd 00:00:00")
        cboSelectTime.ToolTipText = Format(mvCondFilter.Begin, "yyyy-MM-dd") & " - " & Format(mvCondFilter.End, "yyyy-MM-dd")
        lblSeeTim.ToolTipText = cboSelectTime.ToolTipText
        If intDay = 7 And intStart = 0 Then
            cboSelectTime.ListIndex = 1
        ElseIf intDay = 0 And intStart = 0 Then
            cboSelectTime.ListIndex = 0
        Else
            cboSelectTime.ListIndex = cboSelectTime.ListCount - 1
        End If
    End If
    
    '缺省医生本人
    mvCondFilter.医生 = UserInfo.姓名
    
    '其他不缺省
    mvCondFilter.挂号单 = ""
    mvCondFilter.就诊卡 = ""
    mvCondFilter.科室ID = 0
    mvCondFilter.门诊号 = ""
    mvCondFilter.姓名 = ""
    
End Sub

Private Sub GetLocalSetting()
'功能：从注册表读取出院病人的时间范围
    '接诊范围：1=挂本人号的病人,2=本诊室病人,3=本科室病人
    Dim strSQL As String, rsTmp As Recordset, intType As Integer
    Dim str病人接诊控制 As String '问题号:57566
    
    mint接诊范围 = Val(zlDatabase.GetPara("接诊范围", glngSys, p门诊医生站, "2"))
    mstr接诊诊室 = zlDatabase.GetPara("本地诊室", glngSys, p门诊医生站)
    mlng接诊科室ID = Val(zlDatabase.GetPara("接诊科室", glngSys, p门诊医生站))
    On Error GoTo errH
    strSQL = "Select Distinct B.ID,B.编码,B.名称,A.缺省" & _
        " From 部门人员 A,部门表 B,部门性质说明 C" & _
        " Where A.部门ID=B.ID And B.ID=C.部门ID And C.服务对象 In(1,3) And C.工作性质='临床'" & _
        " And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is Null)" & _
        " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) And A.人员ID=[1] And b.ID=[2]" & _
        " Order by B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, mlng接诊科室ID)
    If rsTmp.RecordCount = 0 Then mlng接诊科室ID = 0
    mbln要求分诊 = Val(zlDatabase.GetPara("只接收已经分诊的病人", glngSys, p门诊医生站)) <> 0
    
    '续诊病人
    If InStr(mstrPrivs, "续诊病人") > 0 Then
        mstr接诊医生 = zlDatabase.GetPara("接诊医生", glngSys, p门诊医生站, UserInfo.姓名)
    Else
        mstr接诊医生 = UserInfo.姓名
    End If
    
    '自动化参数
    mbln自动接诊 = Val(zlDatabase.GetPara("找到病人后自动接诊", glngSys, p门诊医生站)) <> 0
    mlng自动进行 = Val(zlDatabase.GetPara("接诊后自动进行", glngSys, p门诊医生站))
    
    '医生主动呼叫后才允许接诊
    mbln呼叫后接诊 = Val(zlDatabase.GetPara("医生主动呼叫后才允许接诊", glngSys, p门诊医生站)) <> 0
    '字体设置
    mbytSize = zlDatabase.GetPara("字体", glngSys, p门诊医生站, "0")

    
    mintFindType = Val(zlDatabase.GetPara("病人查找方式", glngSys, p门诊医生站, "1", , , intType))
    mblnFindTypeEnabled = Not ((intType = 3 Or intType = 15) And InStr(mstrPrivs, "参数设置") = 0)
    
    '问题号:57566
    str病人接诊控制 = CStr(zlDatabase.GetPara("病人接诊控制", glngSys, p门诊医生站))
    If str病人接诊控制 <> "" Then
        mlng接诊控制 = Val(Left(str病人接诊控制, 1))
        If UBound(Split(str病人接诊控制, "|")) >= 1 Then
            mlng提前接收时间 = Val(Split(str病人接诊控制, "|")(1))
        End If
    End If
    
    mblnAutoHandle = Val(zlDatabase.GetPara("接诊时自动处理完成就诊", glngSys, p门诊医生站)) = 1
    
    '医嘱提醒刷新设置
    mstrNotifyAdvice = zlDatabase.GetPara("自动刷新内容", glngSys, p门诊医生站, "0000")
    mintNotifyDay = Val(zlDatabase.GetPara("自动刷新病历审阅天数", glngSys, p门诊医生站, 1))
    mintNotify = Val(zlDatabase.GetPara("自动刷新病历审阅间隔", glngSys, p门诊医生站))
    mbln消息语音 = Val(zlDatabase.GetPara("启用语音提示", glngSys, p门诊医生站)) = 1
    mblnPatiDetail = Val(zlDatabase.GetPara("显示病人详细信息", glngSys, p门诊医生站, 0)) = 1
    mbln危急值 = InStr(GetInsidePrivs(p门诊医生站), ";危急值处理;") > 0
    
    mbln显示预约病人 = Val(zlDatabase.GetPara("显示预约病人", glngSys, p门诊医生站, "1"))
    mbln免挂号模式 = Val(zlDatabase.GetPara(290, glngSys)) = 1
    mbln危急值弹窗 = Val(zlDatabase.GetPara("门诊危急值弹窗提醒", glngSys, p门诊医生站, 1)) = 1
    
    '设置自动刷新
    Call SetTimer
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadPatients候诊()
'功能：加载候诊病人列表
    Dim strSQL As String, strSQLTest As String
    Dim strTime As String
    Dim str标识 As String
    Dim str转诊状态 As String
    Dim i As Long, j As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim rsPati As ADODB.Recordset
    Dim intType As Integer '1候诊、2预约、3转诊
    Dim strTmp As String
    Dim lngColor As Long
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    mblnUnRefresh = True
    
    rptPati(PATI_RPT候诊).Records.DeleteAll
    
    For intType = 1 To 3
        Select Case intType
        Case 1 '候诊病人
            If mint接诊范围 = 1 Then
                strSQL = " And B.执行人||''=[2]" '挂本人号
                If mbln要求分诊 Then strSQL = strSQL & " And B.诊室 is Not NULL"
            ElseIf mint接诊范围 = 2 Then
                '本诊室
                If mlng接诊科室ID <> 0 Then
                    strSQL = " And B.诊室=[3] And b.执行部门id+0 =[4] And (B.执行人||''=[2] Or B.执行人 Is Null) "
                Else    '10.28以前选诊室时没有定科室
                    strSQL = " And B.诊室=[3] And (B.执行人||''=[2] Or B.执行人 Is Null) " & _
                        "And Exists (Select 科室id" & vbNewLine & _
                        " From 挂号安排 F, 部门人员 D" & vbNewLine & _
                        " Where D.人员id = [6] And F.科室id = D.部门id And b.执行部门id = F.科室id)"
                End If
            ElseIf mint接诊范围 = 3 Then
                strSQL = " And B.执行部门ID+0=[4] And (B.执行人||''=[2] Or B.执行人 Is Null)" '本科室
                If mbln要求分诊 Then strSQL = strSQL & " And B.诊室 is Not NULL"
            End If
            strSQL = " Select /*+ Rule*/B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,B.复诊,B.急诊,B.社区," & _
                "       B.发生时间 as 时间,A.就诊卡号,A.身份证号,A.IC卡号,A.险类,nvl(g.号类,E.号类) as 号类,D.名称 as 病人科室," & _
                "       B.号序,B.诊室,B.分诊时间,B.发生时间,B.执行部门ID,B.执行人," & _
                "       B.转诊状态,C.名称 as 转诊科室,B.转诊诊室,B.转诊医生,B.执行状态,B.记录标志,A.病人类型" & _
                " From 病人信息 A,病人挂号记录 B,部门表 C,部门表 D,挂号安排 E, 临床出诊记录 f,临床出诊号源 g" & _
                " Where B.病人ID=A.病人ID And (Nvl(B.执行状态,0)=0 or nvl(B.执行状态,0)=[5]) And B.转诊科室ID=C.ID(+) and b.出诊记录id=f.id(+) and f.号源id=g.id(+) And B.记录性质=1 And B.记录状态=1" & _
                "      And B.号别=E.号码(+) And B.执行部门ID=D.id And B.执行时间 is Null And B.发生时间 <= Trunc(Sysdate)+1-1/24/60/60 " & strSQL & _
                IIf(gint普通挂号天数 = gint急诊挂号天数, " And B.发生时间>=Sysdate-" & IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数), _
                " And B.发生时间 >= Sysdate-" & IIf(gint普通挂号天数 > gint急诊挂号天数, gint普通挂号天数, gint急诊挂号天数) & " And B.发生时间>=Sysdate-Decode(B.急诊,1," & IIf(gint急诊挂号天数 = 0, 1, gint急诊挂号天数) & "," & IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数) & ")") & _
                " Order By Decode(B.分诊时间,NULL,2,1),B.分诊时间,B.NO"
            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "未用", UserInfo.姓名, mstr接诊诊室, IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID), IIf(mblnShowLeavePati, -1, 0), UserInfo.ID)
            
            str标识 = " "
        Case 2 '预约病人
            Set rsPati = Nothing
            If mbln显示预约病人 Then
                If gbln挂号按排 Then
                    '新版挂号出诊安排模式
                    strSQL = "Select B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,B.复诊,B.急诊,B.社区,B.发生时间 as 时间,A.就诊卡号,A.身份证号,A.IC卡号,A.险类,d.号类," & vbNewLine & _
                        "e.名称 as 病人科室,B.号序,B.诊室,B.分诊时间,B.发生时间,B.执行部门ID,B.执行人,B.转诊状态,f.名称 as 转诊科室,B.转诊诊室,B.转诊医生,B.执行状态,B.记录标志,A.病人类型" & vbNewLine & _
                        "From 病人信息 A,病人挂号记录 B, 临床出诊记录 C,临床出诊号源 D,部门表 E,部门表 f" & vbNewLine & _
                        "Where B.病人ID=A.病人ID And  b.出诊记录id = c.Id And c.号源id = d.Id And B.执行部门ID=E.ID And B.转诊科室ID=f.ID(+) and b.记录性质=2 and b.记录状态=1" & vbNewLine & _
                        "And b.发生时间 Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 And  Sysdate Between c.开始时间 And c.终止时间"
                    If mint接诊范围 = 1 Then
                        strSQL = strSQL & " And B.执行人||''=[1]" '挂本人号
                    ElseIf mint接诊范围 = 2 Or mint接诊范围 = 3 Then '本诊室（预约挂号的发药窗口填的是预号，没有诊室），本科室
                        strSQL = strSQL & " And B.执行部门ID+0=[2] And (B.执行人||''=[1] Or B.执行人 Is Null)"
                    End If
                Else
                    If mint接诊范围 = 1 Then
                        strSQL = " And A.执行人||''=[1]" '挂本人号
                    ElseIf mint接诊范围 = 2 Or mint接诊范围 = 3 Then '本诊室（预约挂号的发药窗口填的是预号，没有诊室），本科室
                        strSQL = " And A.执行部门ID+0=[2] And (A.执行人||''=[1] Or A.执行人 Is Null)"
                    End If
                    '现在现在的时间段：用表联接的方式会很慢。
                    strTime = _
                        "Select 时间段 From 时间段 Where" & _
                        " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
                        " Between" & _
                        " Decode(Sign(开始时间-终止时间),1,'3000-01-09 '||To_Char(开始时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(开始时间,'HH24:MI:SS'))" & _
                        " And" & _
                        " '3000-01-10 '||To_Char(终止时间,'HH24:MI:SS'))" & _
                        " Or" & _
                        " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
                        " Between" & _
                        " '3000-01-10 '||To_Char(开始时间,'HH24:MI:SS')" & _
                        " And" & _
                        " Decode(Sign(开始时间-终止时间),1,'3000-01-11 '||To_Char(终止时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(终止时间,'HH24:MI:SS')))"
                    '取现在的星期数对应安排的时间段
                    strTime = " And Decode(To_Char(SysDate,'D'),'1',B.周日,'2',B.周一,'3',B.周二,'4',B.周三,'5',B.周四,'6',B.周五,'7',B.周六,NULL) IN(" & strTime & ")"
                    strSQL = "Select A.NO,A.病人ID,A.标识号 as 门诊号,A.姓名,A.性别,A.年龄,A.加班标志 as 急诊,A.执行人,B.号类,D.名称 as 病人科室," & _
                        " A.发生时间 as 时间,C.就诊卡号,C.身份证号,C.IC卡号,C.险类,A.发生时间,A.执行部门ID,0 as 执行状态,0 as 记录标志,C.病人类型,null as 转诊状态" & _
                        " From 门诊费用记录 A,挂号安排 B,病人信息 C,部门表 D" & _
                        " Where A.计算单位=B.号码 And A.病人ID=C.病人ID(+) And A.执行部门ID=D.id And A.序号=1" & _
                        " And A.记录性质=4 And A.记录状态=0 " & strTime & strSQL & _
                        " And A.发生时间 Between Trunc(Sysdate) And Trunc(Sysdate)+1-1/24/60/60"
                End If
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名, IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID))
            End If
            str标识 = "预"
        Case 3 '转诊病人
            If mint接诊范围 = 1 Then
                strSQL = " And B.转诊医生=[2]" '转本人号
            ElseIf mint接诊范围 = 2 Then
                '转本诊室：不是自已转的，接收医生是自已或者未指定接收医生
                strSQL = " And B.转诊诊室=[3] And B.转诊科室ID=[4] And Nvl(B.执行人,'无')<>[2] And (B.转诊医生=[2] Or B.转诊医生 Is NULL)"
            ElseIf mint接诊范围 = 3 Then
                '转本科室：不是自已转的，接收医生是自已或者未指定接收医生
                strSQL = " And B.转诊科室ID=[4] And Nvl(B.执行人,'无')<>[2] And (B.转诊医生=[2] Or B.转诊医生 Is NULL)"
            End If
            strSQL = _
                " Select /*+ Rule*/B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,B.复诊,B.急诊,B.社区,B.执行人,nvl(g.号类,E.号类) as 号类,D.名称 as 病人科室," & _
                " B.发生时间 as 时间,A.就诊卡号,A.身份证号,A.IC卡号,A.险类,B.发生时间,B.转诊科室ID as 执行部门ID," & _
                " B.转诊状态,C.名称 as 转诊科室,B.诊室 as 转诊诊室,B.执行人 as 转诊医生,B.执行状态,B.记录标志,A.病人类型" & _
                " From 病人信息 A,病人挂号记录 B,部门表 C,部门表 D,挂号安排 E, 临床出诊记录 f,临床出诊号源 g" & _
                " Where B.病人ID=A.病人ID And B.转诊状态=0 And B.执行部门ID=C.ID And B.记录性质=1 And B.记录状态=1 And B.号别=E.号码(+)  and b.出诊记录id=f.id(+) and f.号源id=g.id(+) And B.转诊科室ID=D.id " & strSQL & _
                IIf(gint普通挂号天数 = gint急诊挂号天数, " And B.发生时间>=Sysdate-" & IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数), _
                " And B.发生时间 >= Sysdate-" & IIf(gint普通挂号天数 > gint急诊挂号天数, gint普通挂号天数, gint急诊挂号天数) & " And B.发生时间>=Sysdate-Decode(B.急诊,1," & IIf(gint急诊挂号天数 = 0, 1, gint急诊挂号天数) & "," & IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数) & ")") & _
                " Order By B.NO"
            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "未用", UserInfo.姓名, mstr接诊诊室, IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID), 0, 0)
            str标识 = "转"
        End Select
        
        If Not rsPati Is Nothing Then
            For i = 1 To rsPati.RecordCount
                Set objRecord = rptPati(PATI_RPT候诊).Records.Add()
                For j = 0 To rptPati(PATI_RPT候诊).Columns.Count - 1
                    objRecord.AddItem ""
                Next
                
                With objRecord
                    .Item(COL_HZ_标识).Value = str标识
                    .Item(COL_HZ_门诊号).Value = rsPati!门诊号 & ""
                    .Item(COL_HZ_姓名).Value = rsPati!姓名 & ""
                    .Item(COL_HZ_性别).Value = rsPati!性别 & ""
                    .Item(COL_HZ_年龄).Value = rsPati!年龄 & ""
                    .Item(COL_HZ_急).Value = IIf(Val(rsPati!急诊 & "") <> 0, "急", "")
                    .Item(COL_HZ_就诊卡号).Value = rsPati!就诊卡号 & ""
                    .Item(COL_HZ_病人类型).Value = rsPati!病人类型 & ""
                    .Item(COL_HZ_NO).Value = rsPati!NO & ""
                    .Item(COL_HZ_病人ID).Value = rsPati!病人ID & ""
                    .Item(COL_HZ_发生时间).Value = CStr(Format(rsPati!发生时间, "yyyy-MM-dd HH:mm:ss"))
                    .Item(COL_HZ_执行部门ID).Value = Val(rsPati!执行部门ID & "")
                    .Item(COL_HZ_执行人).Value = rsPati!执行人 & ""
                    .Item(COL_HZ_状态).Value = NVL(rsPati!转诊状态)
                    .Item(COL_HZ_IC卡号).Value = rsPati!IC卡号 & ""
                    .Item(COL_HZ_记录标志).Value = rsPati!记录标志 & ""
                    .Item(COL_HZ_号类).Value = rsPati!号类 & ""
                    .Item(COL_HZ_病人科室).Value = rsPati!病人科室 & ""
                    
                    If intType = 1 Then '候诊
                        .Item(COL_HZ_就诊诊室).Value = rsPati!诊室 & ""
                        .Item(COL_HZ_就诊医生).Value = rsPati!执行人 & ""
                        .Item(COL_HZ_序号).Value = zlStr.Lpad(NVL(rsPati!号序), 5)
                        .Item(COL_HZ_分诊时间).Value = CStr(Format(rsPati!分诊时间, "yyyy-MM-dd HH:mm"))
                        .Item(COL_HZ_执行状态).Value = rsPati!执行状态 & ""
                    End If
                    
                    If intType = 1 Or intType = 3 Then '候诊、转诊
                        .Item(COL_HZ_社区).Value = IIf(Val(rsPati!社区 & "") <> 0, "√", "")
                        .Item(COL_HZ_复).Value = IIf(Val(rsPati!复诊 & "") <> 0, "复", "")
                        .Item(COL_HZ_挂号时间).Value = Format(rsPati!时间, "yyyy-MM-dd HH:mm")
                    End If
                    
                    '转诊状态
                    str转诊状态 = ""
                    If intType = 1 Then
                        If Not IsNull(rsPati!转诊状态) Then
                            If rsPati!转诊状态 = 0 Then
                                '已经转诊
                                str转诊状态 = "待对方接收,科室:" & rsPati!转诊科室 & _
                                    IIf(Not IsNull(rsPati!转诊诊室), ",诊室:" & NVL(rsPati!转诊诊室), "") & _
                                    IIf(Not IsNull(rsPati!转诊医生), ",医生:" & NVL(rsPati!转诊医生), "")
                            ElseIf rsPati!转诊状态 = -1 Then
                                '已拒绝转诊
                                str转诊状态 = "对方已拒绝,科室:" & rsPati!转诊科室 & _
                                    IIf(Not IsNull(rsPati!转诊诊室), ",诊室:" & NVL(rsPati!转诊诊室), "") & _
                                    IIf(Not IsNull(rsPati!转诊医生), ",医生:" & NVL(rsPati!转诊医生), "")
                            End If
                        End If
                    ElseIf intType = 3 Then
                        '转诊病人
                        str转诊状态 = "待接收转诊,科室:" & rsPati!转诊科室 & _
                            IIf(Not IsNull(rsPati!转诊诊室), ",诊室:" & NVL(rsPati!转诊诊室), "") & _
                            IIf(Not IsNull(rsPati!转诊医生), ",医生:" & NVL(rsPati!转诊医生), "")
                    End If
                    .Item(COL_HZ_转诊状态).Value = str转诊状态
                    
                    If intType = 2 Then '预约
                        .Item(COL_HZ_预约医生).Value = rsPati!执行人 & ""
                        .Item(COL_HZ_预约时间).Value = CStr(Format(rsPati!时间 & "", "yyyy-MM-dd HH:mm"))
                    End If
                    .Item(COL_HZ_身份证号).Value = rsPati!身份证号 & ""
                                    
                    '保险病人用红色显示
                    If Not IsNull(rsPati!险类) And rsPati!病人类型 & "" = "" Then
                        .Item(COL_HZ_门诊号).ForeColor = &HC0&
                        .Item(COL_HZ_病人类型).ForeColor = &HC0&
                    Else
                        '病人颜色
                        lngColor = zlDatabase.GetPatiColor(NVL(rsPati!病人类型))
                        .Item(COL_HZ_门诊号).ForeColor = lngColor
                        .Item(COL_HZ_病人类型).ForeColor = lngColor
                    End If
                    
                    '急诊标志红色突出显示
                    If NVL(rsPati!急诊, 0) <> 0 Then
                        .Item(COL_HZ_急).ForeColor = vbRed
                    End If
                    
                    '不就诊病人灰色
                    If Val(rsPati!执行状态 & "") = -1 Then
                        For j = 0 To rptPati(PATI_RPT候诊).Columns.Count - 1
                            .Item(j).ForeColor = &H808080
                        Next
                    End If
                    
                End With
                rsPati.MoveNext
            Next
        End If
        If intType = 1 Then
            Call SetRoomState(rsPati.RecordCount > 0)
        End If
    Next
    
    rptPati(PATI_RPT候诊).Populate
    i = rptPati(PATI_RPT候诊).Records.Count
    tbcWait.Item(0).Caption = "候诊病人" & IIf(i = 0, "", ":" & i & "人")
    mblnUnRefresh = False
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadPatients预约()
'功能：加载候诊病人列表
    Dim strSQL As String, strSQLTest As String
    Dim strTime As String
    Dim i As Long, j As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim rsPati As ADODB.Recordset
    Dim strTmp As String
    Dim lngColor As Long
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    mblnUnRefresh = True
    
    rptPati(PATI_RPT预约).Records.DeleteAll
    

    Set rsPati = Nothing

    If gbln挂号按排 Then
        '新版挂号出诊安排模式
        strSQL = "Select B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,B.复诊,B.急诊,B.社区,B.发生时间 as 时间,A.就诊卡号,A.身份证号,A.IC卡号,A.险类,d.号类," & vbNewLine & _
            "e.名称 as 病人科室,B.号序,B.诊室,B.分诊时间,B.发生时间,B.执行部门ID,B.执行人,B.转诊状态,f.名称 as 转诊科室,B.转诊诊室,B.转诊医生,B.执行状态,B.记录标志,A.病人类型" & vbNewLine & _
            "From 病人信息 A,病人挂号记录 B, 临床出诊记录 C,临床出诊号源 D,部门表 E,部门表 f" & vbNewLine & _
            "Where B.病人ID=A.病人ID And  b.出诊记录id = c.Id And c.号源id = d.Id And B.执行部门ID=E.ID And B.转诊科室ID=f.ID(+) and b.记录性质=2 and b.记录状态=1" & vbNewLine & _
            "And b.发生时间 Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 And  Sysdate Between c.开始时间 And c.终止时间"
        If mint接诊范围 = 1 Then
            strSQL = strSQL & " And B.执行人||''=[1]" '挂本人号
        ElseIf mint接诊范围 = 2 Or mint接诊范围 = 3 Then '本诊室（预约挂号的发药窗口填的是预号，没有诊室），本科室
            strSQL = strSQL & " And B.执行部门ID+0=[2] And (B.执行人||''=[1] Or B.执行人 Is Null)"
        End If
    Else
        If mint接诊范围 = 1 Then
            strSQL = " And A.执行人||''=[1]" '挂本人号
        ElseIf mint接诊范围 = 2 Or mint接诊范围 = 3 Then '本诊室（预约挂号的发药窗口填的是预号，没有诊室），本科室
            strSQL = " And A.执行部门ID+0=[2] And (A.执行人||''=[1] Or A.执行人 Is Null)"
        End If
        '现在现在的时间段：用表联接的方式会很慢。
        strTime = _
            "Select 时间段 From 时间段 Where" & _
            " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
            " Between" & _
            " Decode(Sign(开始时间-终止时间),1,'3000-01-09 '||To_Char(开始时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(开始时间,'HH24:MI:SS'))" & _
            " And" & _
            " '3000-01-10 '||To_Char(终止时间,'HH24:MI:SS'))" & _
            " Or" & _
            " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
            " Between" & _
            " '3000-01-10 '||To_Char(开始时间,'HH24:MI:SS')" & _
            " And" & _
            " Decode(Sign(开始时间-终止时间),1,'3000-01-11 '||To_Char(终止时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(终止时间,'HH24:MI:SS')))"
        '取现在的星期数对应安排的时间段
        strTime = " And Decode(To_Char(SysDate,'D'),'1',B.周日,'2',B.周一,'3',B.周二,'4',B.周三,'5',B.周四,'6',B.周五,'7',B.周六,NULL) IN(" & strTime & ")"
        strSQL = "Select A.NO,A.病人ID,A.标识号 as 门诊号,A.姓名,A.性别,A.年龄,A.加班标志 as 急诊,A.执行人,B.号类,D.名称 as 病人科室," & _
            " A.发生时间 as 时间,C.就诊卡号,C.身份证号,C.IC卡号,C.险类,A.发生时间,A.执行部门ID,0 as 执行状态,0 as 记录标志,C.病人类型,null as 转诊状态" & _
            " From 门诊费用记录 A,挂号安排 B,病人信息 C,部门表 D" & _
            " Where A.计算单位=B.号码 And A.病人ID=C.病人ID(+) And A.执行部门ID=D.id And A.序号=1" & _
            " And A.记录性质=4 And A.记录状态=0 " & strTime & strSQL & _
            " And A.发生时间 Between Trunc(Sysdate) And Trunc(Sysdate)+1-1/24/60/60"
    End If
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名, IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID))
        
    If Not rsPati Is Nothing Then
        For i = 1 To rsPati.RecordCount
            Set objRecord = rptPati(PATI_RPT预约).Records.Add()
            For j = 0 To rptPati(PATI_RPT预约).Columns.Count - 1
                objRecord.AddItem ""
            Next
            
            With objRecord
                .Item(COL_HZ_标识).Value = "预"
                .Item(COL_HZ_门诊号).Value = rsPati!门诊号 & ""
                .Item(COL_HZ_姓名).Value = rsPati!姓名 & ""
                .Item(COL_HZ_性别).Value = rsPati!性别 & ""
                .Item(COL_HZ_年龄).Value = rsPati!年龄 & ""
                .Item(COL_HZ_急).Value = IIf(Val(rsPati!急诊 & "") <> 0, "急", "")
                .Item(COL_HZ_就诊卡号).Value = rsPati!就诊卡号 & ""
                .Item(COL_HZ_病人类型).Value = rsPati!病人类型 & ""
                .Item(COL_HZ_NO).Value = rsPati!NO & ""
                .Item(COL_HZ_病人ID).Value = rsPati!病人ID & ""
                .Item(COL_HZ_发生时间).Value = CStr(Format(rsPati!发生时间, "yyyy-MM-dd HH:mm:ss"))
                .Item(COL_HZ_执行部门ID).Value = Val(rsPati!执行部门ID & "")
                .Item(COL_HZ_执行人).Value = rsPati!执行人 & ""
                .Item(COL_HZ_状态).Value = NVL(rsPati!转诊状态)
                .Item(COL_HZ_IC卡号).Value = rsPati!IC卡号 & ""
                .Item(COL_HZ_记录标志).Value = rsPati!记录标志 & ""
                .Item(COL_HZ_号类).Value = rsPati!号类 & ""
                .Item(COL_HZ_病人科室).Value = rsPati!病人科室 & ""
                .Item(COL_HZ_转诊状态).Value = ""
                .Item(COL_HZ_预约医生).Value = rsPati!执行人 & ""
                .Item(COL_HZ_预约时间).Value = CStr(Format(rsPati!时间 & "", "yyyy-MM-dd HH:mm"))
                .Item(COL_HZ_身份证号).Value = rsPati!身份证号 & ""
                                
                '保险病人用红色显示
                If Not IsNull(rsPati!险类) And rsPati!病人类型 & "" = "" Then
                    .Item(COL_HZ_门诊号).ForeColor = &HC0&
                    .Item(COL_HZ_病人类型).ForeColor = &HC0&
                Else
                    '病人颜色
                    lngColor = zlDatabase.GetPatiColor(NVL(rsPati!病人类型))
                    .Item(COL_HZ_门诊号).ForeColor = lngColor
                    .Item(COL_HZ_病人类型).ForeColor = lngColor
                End If
                
                '急诊标志红色突出显示
                If NVL(rsPati!急诊, 0) <> 0 Then
                    .Item(COL_HZ_急).ForeColor = vbRed
                End If
                
                '不就诊病人灰色
                If Val(rsPati!执行状态 & "") = -1 Then
                    For j = 0 To rptPati(PATI_RPT预约).Columns.Count - 1
                        .Item(j).ForeColor = &H808080
                    Next
                End If
                
            End With
            rsPati.MoveNext
        Next
    End If
    
    rptPati(PATI_RPT预约).Populate
    i = rptPati(PATI_RPT预约).Records.Count
    tbcWait.Item(mint预约列表).Caption = "预约病人" & IIf(i = 0, "", ":" & i & "人")
    mblnUnRefresh = False
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub LoadPatients就诊()
'功能：加载候诊就诊列表
    Dim strSQL As String
    Dim strTime As String
    Dim i As Long, j As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim rsPati As ADODB.Recordset
    Dim strTmp As String
    Dim lngColor As Long
    Dim rs传染病状态 As ADODB.Recordset
    Dim blnDo传染病状态 As Boolean
 
    On Error GoTo errH
    Screen.MousePointer = 11
    mblnUnRefresh = True
    
    strSQL = _
        " Select B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,B.复诊,B.急诊,B.社区,nvl(g.号类,E.号类) as 号类,D.名称 as 病人科室," & _
        " B.执行时间 as 时间,A.就诊卡号,A.身份证号,A.IC卡号,A.险类,B.发生时间,B.执行部门ID,B.执行人," & _
        " B.转诊状态,C.名称 as 转诊科室,B.转诊诊室,B.转诊医生,B.执行状态,B.记录标志,A.病人类型" & _
        " From 病人信息 A,病人挂号记录 B,部门表 C,部门表 D,挂号安排 E,临床出诊记录 f,临床出诊号源 g" & _
        " Where B.病人ID=A.病人ID And B.转诊科室ID=C.ID(+) and B.号别=E.号码(+) and B.执行部门ID=d.id and b.出诊记录id=f.id(+) and f.号源id=g.id(+)" & _
        " And B.执行状态=2 And B.执行人||''=[1] And B.记录性质=1 And B.记录状态=1 and nvl(B.记录标志,0)<=1" & _
        " Order By B.NO"
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr接诊医生)
    
    strSQL = "select m.病人id,m.id,m.no,max(m.记录) as 记录,max(m.填写) as 填写,max(m.状态) as 状态 from" & vbNewLine & _
        "(select a.病人id,a.id, a.no,1 as 记录,0 as 填写,0 as 状态 from 病人挂号记录 a,疾病阳性记录 b" & vbNewLine & _
        "where a.no=b.挂号单 and a.执行状态=2 And a.执行人||''=[1] And a.记录性质=1 And a.记录状态=1 and nvl(a.记录标志,0)<=1" & vbNewLine & _
        "union all" & vbNewLine & _
        "Select  a.病人id,a.id, a.no,0 as 记录,1 as 填写,0 as 状态" & vbNewLine & _
        "From 病人挂号记录 A, 电子病历记录 C, 病历文件列表 D" & vbNewLine & _
        "Where c.文件id = d.Id And d.种类 = 5  and c.病历名称 like '%传染病%' And a.病人id = c.病人id And a.id = c.主页id and a.执行状态=2 And a.执行人||''=[1] And a.记录性质=1 And a.记录状态=1 and nvl(a.记录标志,0)<=1" & vbNewLine & _
        "union all" & vbNewLine & _
        "Select  a.病人id,a.id, a.no,0 as 记录,1 as 填写,e.处理状态 as 状态" & vbNewLine & _
        "From 病人挂号记录 A,电子病历记录 C,病历文件列表 D,疾病申报记录 E" & vbNewLine & _
        "Where a.病人id = c.病人id And a.id = c.主页id and c.id=e.文件id and d.种类=5 and c.病历名称 like '%传染病%' and e.文件id =d.id and a.执行状态=2 And a.执行人||''=[1] And a.记录性质=1 And a.记录状态=1 and nvl(a.记录标志,0)<=1) M" & vbNewLine & _
        "group by m.病人id,m.id,m.no"
    Set rs传染病状态 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr接诊医生)
    If rs传染病状态.RecordCount > 0 Then blnDo传染病状态 = True
 
    rptPati(PATI_RPT就诊).Records.DeleteAll
    For i = 1 To rsPati.RecordCount
        Set objRecord = rptPati(PATI_RPT就诊).Records.Add()
        For j = 0 To rptPati(PATI_RPT就诊).Columns.Count - 1
            objRecord.AddItem ""
        Next
        With objRecord
            .Item(COL_JZ_标识).Value = ""
            .Item(COL_JZ_门诊号).Value = rsPati!门诊号 & ""
            .Item(COL_JZ_姓名).Value = rsPati!姓名 & ""
            .Item(COL_JZ_就诊时间).Value = Format(rsPati!时间, "yyyy-MM-dd HH:mm")
            .Item(COL_JZ_性别).Value = rsPati!性别 & ""
            .Item(COL_JZ_年龄).Value = rsPati!年龄 & ""
            .Item(COL_JZ_急).Value = IIf(Val(rsPati!急诊 & "") <> 0, "急", "")
            .Item(COL_JZ_复).Value = IIf(Val(rsPati!复诊 & "") <> 0, "复", "")
            .Item(COL_JZ_NO).Value = rsPati!NO & ""
            .Item(COL_JZ_社区).Value = IIf(Val(rsPati!社区 & "") <> 0, "√", "")
            .Item(COL_JZ_就诊卡号).Value = rsPati!就诊卡号 & ""
            .Item(COL_JZ_病人类型).Value = rsPati!病人类型 & ""
            .Item(COL_JZ_病人ID).Value = rsPati!病人ID & ""
            .Item(COL_JZ_发生时间).Value = CStr(Format(rsPati!发生时间, "yyyy-MM-dd HH:mm:ss"))
            .Item(COL_JZ_执行部门ID).Value = rsPati!执行部门ID & ""
            .Item(COL_JZ_执行人).Value = rsPati!执行人 & ""
            .Item(COL_JZ_身份证号).Value = rsPati!身份证号 & ""
            .Item(COL_JZ_IC卡号).Value = rsPati!IC卡号 & ""
            .Item(COL_JZ_记录标志).Value = rsPati!记录标志 & ""
            .Item(COL_JZ_号类).Value = rsPati!号类 & ""
            .Item(COL_JZ_病人科室).Value = rsPati!病人科室 & ""
            
            '转诊状态:显示在最后一列
            .Item(COL_JZ_状态).Value = NVL(rsPati!转诊状态)
            If Not IsNull(rsPati!转诊状态) Then
                If rsPati!转诊状态 = 0 Then
                    .Item(COL_JZ_转诊状态).Value = "待对方接收,科室:" & rsPati!转诊科室 & _
                        IIf(Not IsNull(rsPati!转诊诊室), ",诊室:" & NVL(rsPati!转诊诊室), "") & _
                        IIf(Not IsNull(rsPati!转诊医生), ",医生:" & NVL(rsPati!转诊医生), "")
                ElseIf rsPati!转诊状态 = -1 Then
                    '已拒绝转诊
                    .Item(COL_JZ_转诊状态).Value = "对方已拒绝,科室:" & rsPati!转诊科室 & _
                        IIf(Not IsNull(rsPati!转诊诊室), ",诊室:" & NVL(rsPati!转诊诊室), "") & _
                        IIf(Not IsNull(rsPati!转诊医生), ",医生:" & NVL(rsPati!转诊医生), "")
                End If
            End If
            
            '保险病人用红色显示
            If Not IsNull(rsPati!险类) And rsPati!病人类型 & "" = "" Then
                .Item(COL_JZ_门诊号).ForeColor = &HC0&
                .Item(COL_JZ_病人类型).ForeColor = &HC0&
            Else
                '病人颜色
                lngColor = zlDatabase.GetPatiColor(NVL(rsPati!病人类型))
                .Item(COL_JZ_门诊号).ForeColor = lngColor
                .Item(COL_JZ_病人类型).ForeColor = lngColor
            End If
            
            '急诊标志红色突出显示
            If NVL(rsPati!急诊, 0) <> 0 Then
                .Item(COL_JZ_急).ForeColor = vbRed
            End If
            
            '添加传染病状态
            strSQL = ""
            If blnDo传染病状态 Then
                rs传染病状态.Filter = "no='" & rsPati!NO & "'"
                If Not rs传染病状态.EOF Then strSQL = Get传染病状态(Val(rs传染病状态!记录 & ""), Val(rs传染病状态!填写 & ""), Val(rs传染病状态!状态 & ""))
            End If
            .Item(COL_JZ_传染病).Value = strSQL
        End With
        rsPati.MoveNext
    Next
    rptPati(PATI_RPT就诊).Populate
    i = rptPati(PATI_RPT就诊).Records.Count
    tbcInTreat.Item(t在诊).Caption = "在诊" & IIf(i = 0, "", ":" & i & "人")
    mblnUnRefresh = False
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadPatients已诊()
'功能：加载已诊病人列表
    Dim strSQL As String
    Dim strTime As String
    Dim i As Long, j As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim rsPati As ADODB.Recordset
    Dim lngColor As Long
    Dim bln中医 As Boolean
    
    On Error GoTo errH
    Screen.MousePointer = 11
    mblnUnRefresh = True
    rptPati(PATI_RPT已诊).Records.DeleteAll
    
    strSQL = "Select /*+ Rule*/" & vbNewLine & _
        " Distinct(b.No), b.病人id, b.门诊号, b.姓名, b.性别, b.年龄, b.复诊, b.急诊, b.社区, b.执行时间 As 时间, a.就诊卡号, a.身份证号, a.Ic卡号, a.险类, b.发生时间, b.执行部门id," & vbNewLine & _
        " b.执行人, b.执行状态, b.记录标志, a.病人类型,nvl(i.号类,E.号类) as 号类,F.名称 as 病人科室," & vbNewLine & _
        "First_Value(Decode(Sign(h.诊断类型 - 10), -1, h.诊断描述, '')) Over(Partition By h.病人id, h.主页id Order By Sign(h.诊断类型 - 10), Decode(h.记录来源, 4, 0, h.记录来源) Desc, Decode(h.诊断类型, 1, 1, 0) Desc, h.诊断次序) As 西医诊断," & vbNewLine & _
        "First_Value(Decode(Sign(h.诊断类型 - 10), 1, h.诊断描述, '')) Over(Partition By h.病人id, h.主页id Order By -Sign(h.诊断类型 - 10), Decode(h.记录来源, 4, 0, h.记录来源) Desc, Decode(h.诊断类型,11,11, 0) Desc, h.诊断次序) As 中医诊断" & vbNewLine & _
        "From 病人信息 A, 病人挂号记录 B" & IIf(mvCondFilter.就诊卡 <> "", ",病人医疗卡信息 C, 医疗卡类别 D", "") & ",挂号安排 E,部门表 F, 病人诊断记录 H, 临床出诊记录 g,临床出诊号源 I" & vbNewLine & _
        "Where b.病人id = a.病人id And h.病人id(+) = b.病人id And h.主页id(+) = b.id And b.执行状态 + 0 = 1 And b.记录性质 = 1 And b.记录状态 = 1 and B.号别=E.号码(+) and b.执行部门id=f.id and b.出诊记录id=g.id(+) and g.号源id=i.id(+)" & _
         IIf(mvCondFilter.就诊卡 <> "", " And c.病人id = a.病人id And c.卡类别id = d.Id And d.是否固定 = 1 And d.名称 = '就诊卡' ", "")

    If mvCondFilter.挂号单 <> "" Then
        strSQL = strSQL & " And B.NO=[5]"
    ElseIf mvCondFilter.门诊号 <> "" Then
        strSQL = strSQL & " And A.门诊号=[6]"
    ElseIf mvCondFilter.就诊卡 <> "" Then
        strSQL = strSQL & " And C.卡号=[7]"
    
    Else
        strSQL = strSQL & " And B.执行时间 Between To_Date('" & Format(mvCondFilter.Begin, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(mvCondFilter.End, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        strSQL = strSQL & IIf(mvCondFilter.医生 = "", "", " And B.执行人||''=[3]")
        If mvCondFilter.科室ID <> 0 Then strSQL = strSQL & " And B.执行部门ID+0=[4]"
                If mvCondFilter.姓名 <> "" Then strSQL = strSQL & " And A.姓名=[8]"
    End If
    
    If zlDatabase.DateMoved(mvCondFilter.Begin) Then
        strSQL = strSQL & " Union ALL " & Replace(strSQL, "病人挂号记录", "H病人挂号记录")
    End If

    strSQL = strSQL & " Order By NO Desc"
    
    With mvCondFilter
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "未用", "未用", .医生, .科室ID, .挂号单, .门诊号, .就诊卡, .姓名)
    End With
    
    For i = 1 To rsPati.RecordCount
        Set objRecord = rptPati(PATI_RPT已诊).Records.Add()
        For j = 0 To rptPati(PATI_RPT已诊).Columns.Count - 1
            objRecord.AddItem ""
        Next
        
        With objRecord
            .Item(COL_YZ_门诊号).Value = rsPati!门诊号 & ""
            .Item(COL_YZ_姓名).Value = rsPati!姓名 & ""
            .Item(COL_YZ_性别).Value = rsPati!性别 & ""
            .Item(COL_YZ_年龄).Value = rsPati!年龄 & ""
            .Item(COL_YZ_急).Value = IIf(Val(rsPati!急诊 & "") <> 0, "急", "")
            .Item(COL_YZ_复).Value = IIf(Val(rsPati!复诊 & "") <> 0, "复", "")
            .Item(COL_YZ_社区).Value = IIf(Val(rsPati!社区 & "") <> 0, "√", "")
            .Item(COL_YZ_时间).Value = CStr(Format(rsPati!时间 & "", "yyyy-MM-dd HH:mm"))
            .Item(COL_YZ_就诊医生).Value = rsPati!执行人 & ""
            .Item(COL_YZ_就诊卡号).Value = rsPati!就诊卡号 & ""
            .Item(COL_YZ_病人类型).Value = rsPati!病人类型 & ""
            .Item(COL_YZ_号类).Value = rsPati!号类 & ""
            .Item(COL_YZ_病人科室).Value = rsPati!病人科室 & ""
            .Item(COL_YZ_NO).Value = rsPati!NO & ""
            .Item(COL_YZ_病人ID).Value = rsPati!病人ID & ""
            .Item(COL_YZ_发生时间).Value = CStr(Format(rsPati!发生时间, "yyyy-MM-dd HH:mm:ss"))
            .Item(COL_YZ_执行部门ID).Value = Val(rsPati!执行部门ID & "")
            .Item(COL_YZ_执行人).Value = rsPati!执行人 & ""
            .Item(COL_YZ_身份证号).Value = rsPati!身份证号 & ""
            .Item(COL_YZ_IC卡号).Value = rsPati!IC卡号 & ""
            .Item(COL_YZ_记录标志).Value = rsPati!记录标志 & ""
            .Item(COL_YZ_西医诊断).Value = rsPati!西医诊断 & ""
            .Item(COL_YZ_中医诊断).Value = rsPati!中医诊断 & ""
            If rsPati!中医诊断 & "" <> "" Then bln中医 = True
            
            '保险病人用红色显示
            If Not IsNull(rsPati!险类) And rsPati!病人类型 & "" = "" Then
                .Item(COL_YZ_门诊号).ForeColor = &HC0&
                .Item(COL_YZ_病人类型).ForeColor = &HC0&
            Else
                '病人颜色
                lngColor = zlDatabase.GetPatiColor(NVL(rsPati!病人类型))
                .Item(COL_YZ_门诊号).ForeColor = lngColor
                .Item(COL_YZ_病人类型).ForeColor = lngColor
            End If
            
            '急诊标志红色突出显示
            If NVL(rsPati!急诊, 0) <> 0 Then
                .Item(COL_YZ_急).ForeColor = vbRed
            End If
            
        End With
        rsPati.MoveNext
    Next
    
    rptPati(PATI_RPT已诊).Columns(COL_YZ_中医诊断).Visible = bln中医
    rptPati(PATI_RPT已诊).Populate
    i = rptPati(PATI_RPT已诊).Records.Count
    tbcInTreat.Item(t完成).Caption = "完成" & IIf(i = 0, "", ":" & i & "人")
    mblnUnRefresh = False
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadPatients(Optional ByVal strRefesh As String = "11111", Optional ByVal intActive As PATI_RPT_LIST = -1, Optional ByVal strActNO As String) As Boolean
'功能：读取病人列表
'参数：strActNO=刷新后想要定位的列表索引和病人挂号单(如果有)
'      注意其中如果指定了intActive,则必须要包含strRefesh刷新列表中
'      strRefesh=分别是否刷新指定的列表，分别为 第1位－"候诊/转诊/预约"，第2位－"就诊"，第3位－"已诊"，第4位-"回诊"，第5位-"预约"
    Dim strPrePati As String
    Dim i As Long, j As Long
    Dim blnFinded As Boolean
    Dim intIdx As Long
    Dim lngCol As Long
    Dim objRpt As ReportControl
    
    strPrePati = mstrPrePati '因为要破坏,因此临时记录
    
    If strActNO <> "" Then strPrePati = strActNO
    
    Screen.MousePointer = 11
    On Error GoTo errH
    mblnUnRefresh = True
    
    For i = 1 To 5
        If Mid(strRefesh, i, 1) = "1" Then
            If i = 1 Then
                Call LoadPatients候诊
                If mbln显示预约病人 Then
                    rptPati(PATI_RPT预约).Records.DeleteAll
                    rptPati(PATI_RPT预约).Populate
                End If
            ElseIf i = 2 Then
                Call LoadPatients就诊
            ElseIf i = 3 Then
                Call LoadPatients已诊
            ElseIf i = 4 Then
                Call LoadPatients回诊
            ElseIf i = 5 Then
                If Not mbln显示预约病人 Then
                    Call LoadPatients预约
                End If
            End If
        End If
    Next
    i = 0
    For intIdx = 0 To 4
        Set objRpt = rptPati(intIdx)
        If objRpt.Visible Then
            lngCol = IIf(intIdx = PATI_RPT候诊 Or intIdx = PATI_RPT预约, COL_HZ_NO, IIf(intIdx = PATI_RPT已诊, COL_YZ_NO, COL_JZ_NO))
            For i = i To objRpt.Rows.Count - 1
                With objRpt.Rows(i)
                    If CStr(.Record(lngCol).Value) = strPrePati Then
                        Exit For
                    End If
                End With
            Next
            If i <= objRpt.Rows.Count - 1 Then Exit For
            i = 0
        End If
    Next
    If intIdx <= 4 Then
        For lngCol = 0 To 3
            If rptPati(lngCol).SelectedRows.Count > 0 Then
                rptPati(lngCol).SelectedRows(0).Selected = False
            End If
        Next
        mstrPrePati = ""
        '该行选中且显示在可见区域,并引发SelectionChanged事件
        Set objRpt.FocusedRow = objRpt.Rows(i)
        If objRpt.Visible Then objRpt.SetFocus
        Call RptItemClick(intIdx)
    Else
        '按当前列表无数据刷新子窗体
        Call ClearPatiInfo
        Call SubWinRefreshData(tbcSub.Selected)
    End If
    Screen.MousePointer = 0
    LoadPatients = True
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
    mstr挂号单 = ""
    mlng科室ID = 0
    mlng挂号ID = 0
    mstrPrePati = ""
    mPatiInfo.类型 = 0
    mPatiInfo.门诊号 = ""
    mPatiInfo.挂号单 = ""
    mPatiInfo.病人ID = 0
    mPatiInfo.挂号ID = 0
    mPatiInfo.科室ID = 0
    mPatiInfo.诊室 = ""
    mPatiInfo.社区 = 0
    mPatiInfo.社区号 = ""
    mPatiInfo.挂号时间 = CDate(0)
    mPatiInfo.数据转出 = False
    mPatiInfo.是否签名 = False
    mPatiInfo.保存人 = ""
    mPatiInfo.婚姻状况 = ""
    mPatiInfo.性别 = ""
    mPatiInfo.民族 = ""
    mPatiInfo.国籍 = ""
    mPatiInfo.区域 = ""
    mPatiInfo.出生地点 = ""
    mPatiInfo.传染病上传 = 0
    mPatiInfo.家庭地址邮编 = ""
    mPatiInfo.单位邮编 = ""
    mPatiInfo.其他证件 = ""
    
    imgPatient.Picture = imgDefual.Picture

    txtInfo(txtInfo姓名).Text = ""
    txtInfo(txtInfo性别).Text = ""
    txtInfo(txtInfo年龄).Text = ""
    txtInfo(txtInfo出生日期).Text = ""
    txtInfo(txtInfo就诊卡号).Text = ""
    txtInfo(txtInfo医保卡号).Text = ""
    txtInfo(txtInfo摘要).Text = ""
    txtInfo(txtInfo摘要).ToolTipText = ""
    
    lblMore.Visible = False
    lblRec.Visible = False
    lblUrg.Visible = False
    
    cboPayType.ListIndex = -1
    cboBillType.ListIndex = -1
    
    For i = 0 To lblLink修改
        lblLink(i).ForeColor = &HC0C0C0
    Next
    mPr = -1
End Sub

Private Sub ExecuteRegist(ByVal strNO As String)
'功能：病人挂号
    Dim objControl As CommandBarControl
    mblnUnRefresh = True
    '刷新并定位到刚挂号的病人上
    If strNO <> "" And rptPati(PATI_RPT就诊).Visible Then
        Call LoadPatients("11001", PATI_RPT就诊, strNO)
    Else
        Call LoadPatients("10001")
    End If
    mblnUnRefresh = False
End Sub

Private Sub ExecuteBespeakPrint()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重打预约挂号单
    '编制:刘兴洪
    '日期:2012-12-24 10:55:39
    '说明:
    '问题:56274
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCommon As String, intAtom As Integer, strNO As String
    On Error Resume Next
    If gobjRegist Is Nothing Then
        Set gobjRegist = CreateObject("zl9RegEvent.clsRegEvent")
        If err <> 0 Then
            err = 0: On Error GoTo 0
        End If
        If gobjRegist Is Nothing Then Exit Sub
    End If
    On Error GoTo errHandle
 
    strNO = mstr挂号单
 
    If strNO = "" Then Exit Sub
    '部件调用(处理合法性设置)
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    strNO = gobjRegist.zlPrintBespeak(Me, gcnOracle, glngSys, gstrDBUser, mstrPrivs, strNO)
    Call GlobalDeleteAtom(intAtom)
    mblnUnRefresh = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ExecuteTransferSend()
'功能：病人转诊
    Dim rsTmp As New ADODB.Recordset
    Dim lng科室ID As Long, str诊室 As String
    Dim str医生 As String, lng医生ID As Long
    Dim strSQL As String
     
    If mstr挂号单 = "" Then
        MsgBox "请先选择病人。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mintActive = pt已诊 Then
        If zlDatabase.NOMoved("门诊费用记录", mstr挂号单, "记录性质=", "4") Then
            MsgBox "该病人的挂号费用已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '检查挂号单时限
    If BillExpend(mstr挂号单) Then
        MsgBox "该病人挂号已超过有效天数，不能再进行转诊。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errH
    
    '对正在就诊的病人的检查
    If mintActive = pt就诊 Or mintActive = pt回诊 Then
        If InStr(GetInsidePrivs(p门诊医生站), "已下医嘱转诊") > 0 Then
            '检查是否还有未发送的医嘱
            strSQL = "Select ID From 病人医嘱记录 Where 病人ID+0=[1] And 挂号单=[2] And 医嘱状态=1  And NVL(执行标记,0) <> -1 And Nvl(执行性质,0)<>0 And Rownum = 1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
            If Not rsTmp.EOF Then
                MsgBox "该病人还有未发送医嘱，只有将所有医嘱发送后才能进行转诊。", vbInformation, gstrSysName
                Exit Sub
            End If
        Else    '只要下过医嘱(不含已作废的)，说明就诊行为已发生，不允许转诊，须重新挂号
            strSQL = "Select ID From 病人医嘱记录 Where 病人ID+0=[1] And 挂号单=[2] And 医嘱状态 <> 4 And Rownum = 1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
            If Not rsTmp.EOF Then
                MsgBox "已经对该病人下过医嘱，不允许转诊，请删除或作废医嘱后再进行，或者重新挂号。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    If Not frmRegistPlan.ShowMe(Me, mstr挂号单, lng科室ID, str诊室, str医生, lng医生ID) Then mblnUnRefresh = False: Exit Sub
    
    '执行转诊
    strSQL = "Zl_病人挂号记录_转诊('" & mstr挂号单 & "',0," & lng科室ID & ",'" & str诊室 & "','" & str医生 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    '门诊患者转诊消息发送
    Call ZLHIS_CIS_007(mclsMipModule, mlng病人ID, Trim(txtInfo(txtInfo姓名).Text), mPatiInfo.门诊号, mlng挂号ID, mlng接诊科室ID, , lng科室ID, , lng医生ID, str医生, str诊室, UserInfo.姓名)
    
    Call zlShowQuence(mstr挂号单)
    '刷新界面
    Call LoadPatients("11001")
    Call SetReceiveToday(False, -1)
  
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlShowQuence(ByVal strNO As String)
    '功能:显示排队叫号队列的新号
    Dim strSQL As String, rsTemp As ADODB.Recordset
    If Check排队叫号 = False Then Exit Sub
    '95637:李南春,2016/7/20,存在有效队列才提示
    strSQL = "Select 排队号码 From 排队叫号队列 Where 业务类型=0 And 排队状态 In (0,1,7) and 业务ID in (Select ID From 病人挂号记录 where NO=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Sub
    MsgBox "注意:" & vbCrLf & "    该病人重新进行了排队处理,队号为:[ " & NVL(rsTemp!排队号码) & " ]", vbInformation + vbOKOnly, gstrSysName
End Sub

Private Sub ExecuteTransferRefuse()
'功能：转诊拒绝
    Dim strSQL As String
        
    On Error GoTo errH
    
    If mPr <> -1 Then
        If MsgBox("确实要拒绝该转诊病人""" & rptPati(mintRPTIndex).Rows(mPr).Record(COL_JZ_姓名).Value & """吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        strSQL = "Zl_病人挂号记录_转诊('" & rptPati(mintRPTIndex).Rows(mPr).Record(COL_JZ_NO).Value & "',-1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    '刷新界面
    Call LoadPatients("11001")
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteTransferCancel(Optional ByVal blnMsg As Boolean = True)
'功能：取消转诊
    Dim strSQL As String
    
    On Error GoTo errH
 
    With rptPati(mintRPTIndex).Rows(mPr)
        If blnMsg Then
            If MsgBox("确实要取消病人""" & .Record(COL_JZ_姓名).Value & """的转诊吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        End If
        strSQL = "Zl_病人挂号记录_转诊('" & .Record(COL_JZ_NO).Value & "',Null)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End With
    
    '刷新界面
    Call LoadPatients("11011")
    Call ReshDataQueue
    Call SetReceiveToday(False, 1)
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteTransferIncept()
'功能：接收转诊
    Dim strSQL As String
    
    On Error GoTo errH
    
    With rptPati(mintRPTIndex).Rows(mPr)
        If MsgBox(.Record(COL_JZ_转诊状态).Value & vbCrLf & vbCrLf & "确认接收该转诊病人吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        strSQL = "Zl_病人挂号记录_转诊('" & .Record(COL_JZ_NO).Value & "',1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        If HaveRIS Then
            If gobjRis.HISModPati(1, mlng病人ID, mlng挂号ID) <> 1 Then
                MsgBox "当前启用了影像信息系统接口， 但由于影像信息系统接口(HISModPati)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
            End If
        ElseIf gbln启用影像信息系统接口 = True Then
            MsgBox "当前启用了影像信息系统接口，但由于RIS接口创建失败未调用(HISModPati)接口，请与系统管理员联系。", vbInformation, gstrSysName
        End If
        Call mclsAdvices.zlRefresh(0, "", False) '87707
        '刷新并定位病人
        If rptPati(PATI_RPT就诊).Visible Then
            Call LoadPatients("11001", PATI_RPT就诊, .Record(COL_JZ_NO).Value)
        Else
            Call LoadPatients("11001")
        End If
        Call SetReceiveToday(False, 1)
    End With
    Call ReshDataQueue
    
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteReceive(Optional ByVal blnIsCard As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人接诊
    '参数:blnIsCard-是否是刷卡调用接收预约病人
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim objControl As CommandBarControl
    Dim strSQL As String, strNO As String
    Dim blnReserve As Boolean
    Dim datCurr As Date
   
    On Error GoTo errH

    datCurr = zlDatabase.Currentdate
    
    If (mintRPTIndex = PATI_RPT候诊 Or mintRPTIndex = PATI_RPT预约) And mPr <> -1 Then
        If rptPati(mintRPTIndex).Rows(mPr).Record(COL_HZ_标识).Value = "预" Then
            blnReserve = True
        End If
    Else
        Exit Sub
    End If
    
    If blnReserve Then
        '对预约挂号病人进行接诊
        '问题号:57566
        If Check接诊控制("接诊", mstr挂号单) = False Then Exit Sub
        
        '门诊医生站预约接收时调用挂号部件的接收接口进行扣费的功能
        If Val(zlDatabase.GetPara("挂号模式", glngSys, 9000, 1)) <> 1 And Not mobjSquareCard Is Nothing Then
            If Not mobjSquareCard.zlRegisterIncept(Me, mlngModul, mstr挂号单, mstr接诊诊室, PatiIdentify.objIDKind.GetCurCard.接口序号, PatiIdentify.Text) Then Exit Sub
        Else
            strSQL = "Zl_病人预约挂号_接收('" & mstr挂号单 & "','" & mstr接诊诊室 & "',NULL,NULL,NULL,NULL,NULL,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    Else
        '问题号:57566
        If Check接诊控制("接诊", mstr挂号单) = False Then Exit Sub
                '转诊病人直接调用转诊接收
        If mintActive = pt转诊 Then
            ExecuteTransferIncept
            Exit Sub
        End If
        '对正常挂号病人进行接诊
        strSQL = "Select 执行人 From 病人挂号记录 Where 病人ID+0=[1] And NO=[2] And Nvl(执行状态,0)<>0 And 记录性质=1 And 记录状态=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
        If Not rsTmp.EOF Then
            MsgBox "该病人已由" & IIf(IsNull(rsTmp!执行人), "其他医生", "医生：" & rsTmp!执行人 & " ") & "接诊。", vbInformation, gstrSysName
            Call LoadPatients("10001"): Exit Sub
        End If
        
        strSQL = "Select 执行人 From 病人挂号记录 Where 病人ID+0=[1] And NO=[2] And Nvl(执行状态,0)=0 And 记录性质=1 And 记录状态=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
        If rsTmp.EOF Then
            MsgBox "该病人已退号，不能接诊。", vbInformation, gstrSysName
            Call LoadPatients("10001"): Exit Sub
        End If
        
        strSQL = "zl_病人接诊(" & mlng病人ID & ",'" & mstr挂号单 & "',Null,'" & UserInfo.姓名 & "','" & mstr接诊诊室 & "',0,0,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    End If
    
    If mblnAutoHandle Then Call Tip病人自动完成
    
    '刷新并定位病人
    On Error GoTo 0
    
    If rptPati(PATI_RPT就诊).Visible Then
        tbcInTreat.Item(t在诊).Selected = True
        Call LoadPatients("11001", PATI_RPT就诊, mstr挂号单)
    Else
        tbcInTreat.Item(t在诊).Selected = True
    End If
    
    '门诊患者接诊消息发送
    Call ZLHIS_CIS_009(mclsMipModule, mlng病人ID, Trim(txtInfo(txtInfo姓名).Text), mPatiInfo.门诊号, 0, 0, mlng挂号ID, mPatiInfo.复诊, mPatiInfo.急诊, datCurr, mlng接诊科室ID, , mstr接诊诊室, UserInfo.姓名)

    '社区病人自动调用功能
    If Not gobjCommunity Is Nothing And mlngCommunityID <> 0 And mlng病人ID <> 0 And mPatiInfo.社区 <> 0 Then
        Set objControl = cbsMain.FindControl(, mlngCommunityID, , True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then objControl.Execute
        End If
    End If
    
    Call ReceiveAfterExec
    
    '处理排队叫号队列(重新刷新)
    Call ReshDataQueue
    Call SetReceiveToday(False, 1)
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ReceiveAfterExec(Optional ByVal bln回诊 As Boolean)
'功能：接诊后需要调用的部分
    Dim objControl As CommandBarControl
    
    Call CreatePlugInOK(p门诊医生站)
    '接诊后调用外挂接口
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.ClinicReceive(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID)
        Call zlPlugInErrH(err, "ClinicReceive")
        err.Clear: On Error GoTo errH
    End If
    
    '接诊之后自动进行医嘱下达状态
    If mlng自动进行 = 1 And bln回诊 = False Then
        Call LocatedCard("医嘱")
        cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
        Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then objControl.Execute
        End If
    ElseIf mlng自动进行 = 2 And bln回诊 = False Then
        If GetInsidePrivs(p新版门诊病历, True) <> "" And Not mclsEMR Is Nothing Then
            Call LocatedCard("新病历")
            cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
        Else
            Call LocatedCard("病历")
            cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
            mblnUnRefresh = True
            Call mclsEPRs.zlOpenDefaultEPR(mstr挂号单)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteCancel()
'功能：取消接诊
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        Dim blnTran As Boolean, colsql As New Collection, i As Long, bytOut As Byte
        
    If BillExpend(mstr挂号单) Then
        MsgBox "该病人挂号已超过有效天数，不允许再取消接诊。", vbInformation, gstrSysName
        Exit Sub
    End If
        
    On Error GoTo errH
    
    '只能取消自己接诊的病人
    strSQL = "Select 执行人 From 病人挂号记录 Where id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mPatiInfo.挂号ID)
    If rsTmp!执行人 <> UserInfo.姓名 Then
        MsgBox "只能取消自己接诊的病人。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'ToDo:取消接诊时病历数据的检查
    '医嘱数据的检查
    strSQL = "Select Count(*) as 医嘱 From 病人医嘱记录 Where 医嘱状态 IN(1,8) And 病人ID+0=[1] And 挂号单=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
    If NVL(rsTmp!医嘱, 0) > 0 Then
        MsgBox "该病人已有新开或已发送的医嘱，不能取消接诊。" & vbCrLf & _
            "如果确实要取消接诊，请先将这些医嘱删除或作废。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mbln免挂号模式 Then
        If mclsReg.zlRegisterPriceDeleteFromNO(mstr挂号单, IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID), IIf(mstr接诊医生 = "", UserInfo.姓名, mstr接诊医生), True, bytOut, colsql) = False Then
            'bytOut:错误类型:0-正常执行;1-未找到挂号单;2-未生成划价单;3-未找到符合条件的划价单;4-存在已经收费的单据
            '不管是否交费，允许取消接诊，如果已经缴费，则费用病人自行去退费
        End If
    End If
    gcnOracle.BeginTrans: blnTran = True
    
    strSQL = "Zl_病人接诊_Cancel(" & mlng病人ID & ",'" & mstr挂号单 & "'," & IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID) & ",'" & IIf(mstr接诊医生 = "", UserInfo.姓名, mstr接诊医生) & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    For i = 1 To colsql.Count
        Call zlDatabase.ExecuteProcedure(colsql(i), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTran = False
    
    '刷新并定位病人
    If rptPati(PATI_RPT候诊).Visible Then
        Call LoadPatients("11001", PATI_RPT候诊, mstr挂号单)
    ElseIf rptPati(PATI_RPT预约).Visible Then
        Call LoadPatients("11001", PATI_RPT预约, mstr挂号单)
    Else
        Call LoadPatients("11001")
    End If
    Call ReshDataQueue
    Call SetReceiveToday(False, -1)
    Exit Sub
errH:
        If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteFinish()
'功能：完成接诊
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, blnTran As Boolean
    Dim str疾病IDs As String, str诊断IDs As String
    Dim lng挂号id As Long
    Dim str姓名 As String
    Dim str状态 As String
    Dim lngSelectedIndex As Long
    Dim rptRow As ReportRow
    
    On Error GoTo errH
 
    If (mintRPTIndex = PATI_RPT就诊 Or mintRPTIndex = PATI_RPT回诊) And mPr <> -1 Then
        With rptPati(mintRPTIndex).Rows(mPr)
            str姓名 = .Record(COL_JZ_姓名).Value
            str状态 = .Record(COL_JZ_状态).Value
            lngSelectedIndex = .Record.Index
        End With
    Else
        Exit Sub
    End If
    
    '如果列表长时间不刷新并发操作检查
    strSQL = "select 1 from 病人挂号记录 where no=[1] and 执行人=[2] And 执行状态=2 And 记录性质=1 And 记录状态=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号单, mstr接诊医生)
    If rsTmp.EOF Then
        MsgBox """" & str姓名 & """可能被其他医生强制续诊接收，请重试。", vbInformation, gstrSysName
        Call LoadPatients
        Call ReshDataQueue
        Exit Sub
    End If
    
    'ToDo:完成接诊时病历数据的检查
    If str状态 = "0" Then
        If MsgBox("当前病人""" & str姓名 & """已经转诊，是否要取消转诊后再完成接诊？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        Else
            Call ExecuteTransferCancel(False)
            Call ExecuteFinish
            Exit Sub
        End If
    End If
    
    '检查是否存在有效医嘱
    strSQL = "Select Count(*) as 医嘱 From 病人医嘱记录 Where 病人ID+0=[1] And 挂号单=[2] And 医嘱状态<>4"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
    If NVL(rsTmp!医嘱, 0) = 0 Then
        If MsgBox("未对""" & str姓名 & """下达任何有效的医嘱，确实要完成接诊吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    '检查是否存在未发送的医嘱
    strSQL = "Select Count(*) as 医嘱 From 病人医嘱记录 Where 病人ID+0=[1] And 挂号单=[2] And 医嘱状态=1 And NVL(执行标记,0) <> -1 And Nvl(执行性质,0)<>0 And Nvl(皮试结果,'无')<>'免试'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
    If NVL(rsTmp!医嘱, 0) > 0 Then
        MsgBox """" & str姓名 & """还有未发送的医嘱，不能完成接诊。", vbInformation, gstrSysName
        Exit Sub
    End If
    '检查未填写的疾病证明报告
    strSQL = "Select 主页ID,疾病ID,诊断ID From 病人诊断记录 Where 取消时间 is Null And 病人ID=[1] And 主页ID=(Select ID From 病人挂号记录 Where NO=[2] And 记录性质=1 And 记录状态=1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
    Do While Not rsTmp.EOF
        If lng挂号id = 0 Then lng挂号id = rsTmp!主页ID
        If Not IsNull(rsTmp!疾病id) Then str疾病IDs = str疾病IDs & "," & rsTmp!疾病id
        If Not IsNull(rsTmp!诊断id) Then str诊断IDs = str诊断IDs & "," & rsTmp!诊断id
        rsTmp.MoveNext
    Loop
    If str疾病IDs <> "" Or str诊断IDs <> "" Then
        If Not CheckDiseaseFile(Me, mlng病人ID, lng挂号id, mlng接诊科室ID, Mid(str疾病IDs, 2), Mid(str诊断IDs, 2), , True, , 1) Then Exit Sub
    End If
    
    If lng挂号id = 0 Then lng挂号id = mPatiInfo.挂号ID
    
    If Not ExecuteFinishInSide(mstr挂号单, mlng病人ID, lng挂号id) Then
        Exit Sub
    End If

    '刷新:不定位到已诊列表
    Call LoadPatients
    Call ReshDataQueue
    
     '完成接诊之后，自动定位到下一行
    If rptPati(mintRPTIndex).Rows.Count > lngSelectedIndex Then
        For Each rptRow In rptPati(mintRPTIndex).Rows
            If rptRow.GroupRow = False Then
                If rptRow.Record.Index = lngSelectedIndex Then
                    Set rptPati(mintRPTIndex).FocusedRow = rptRow
                End If
            End If
        Next
    ElseIf rptPati(mintRPTIndex).Rows.Count = lngSelectedIndex And lngSelectedIndex <> 0 Then
        For Each rptRow In rptPati(mintRPTIndex).Rows
            If rptRow.GroupRow = False Then
                If rptRow.Record.Index = lngSelectedIndex - 1 Then
                    Set rptPati(mintRPTIndex).FocusedRow = rptRow
                End If
            End If
        Next
    End If
    
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ExecuteFinishInSide(ByVal strNO As String, ByVal lng病人ID As Long, ByVal lng挂号id As Long) As Boolean
'功能：完成就诊
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim blnTran As Boolean
    
    On Error GoTo errH
    
    '读取必要的信息供社区接口调用:以左边就诊病人本次就诊为准,右边可能当前选择的历史就诊
    strSQL = "Select A.ID,A.社区,B.社区号 From 病人挂号记录 A,病人社区信息 B Where A.病人ID=B.病人ID(+) And A.记录性质=1 And A.记录状态=1 And A.社区=B.社区(+) And A.NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    
    '执行过程
    '-----------------------------------
    gcnOracle.BeginTrans: blnTran = True
    
    strSQL = "Zl_病人接诊完成(" & lng病人ID & ",'" & strNO & "','" & mstr接诊诊室 & "','" & UserInfo.姓名 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
         
    If Not gobjCommunity Is Nothing And NVL(rsTmp!社区, 0) <> 0 Then
        '调用社区病人信息提交
        If Not gobjCommunity.ClinicSubmit(glngSys, mlngModul, rsTmp!社区, NVL(rsTmp!社区号), lng病人ID, rsTmp!ID) Then
            gcnOracle.RollbackTrans: blnTran = False: Exit Function
        End If
    End If
    gcnOracle.CommitTrans: blnTran = False

    '接诊后调用外挂接口
    Call CreatePlugInOK(p门诊医生站)
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.ClinicFinish(glngSys, p门诊医生站, lng病人ID, lng挂号id)
        Call zlPlugInErrH(err, "ClinicFinish")
        err.Clear: On Error GoTo errH
    End If
    
    '一卡通数据上传
    If Not mobjICCard Is Nothing Then
        strSQL = "Select 1 From 一卡通目录 Where 启用=2 And Rownum=1"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            mobjICCard.UploadSwap lng病人ID, ""
        End If
    End If
        ExecuteFinishInSide = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ExecuteRedo()
'恢复接诊
    Dim strSQL As String
    
    '只检查在线数据表中的
    If BillExpend(mstr挂号单) Then
        MsgBox "该病人挂号已超过有效天数，不允许再恢复接诊。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mintRPTIndex = PATI_RPT已诊 Then
        If zlDatabase.NOMoved("病人挂号记录", mstr挂号单) Then
            MsgBox "该挂号记录已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '当前医生完成的病人才可以直接恢复(否则有权限可用强制续诊)
    If mintRPTIndex = PATI_RPT已诊 Then
        With rptPati(PATI_RPT已诊).Rows(mPr)
            If .Record(COL_YZ_执行人).Value <> UserInfo.姓名 Then
                MsgBox "该病人不是由你完成就诊的，不能直接恢复接诊。", vbInformation, gstrSysName
                Exit Sub
            End If
        End With
    End If
    
    On Error GoTo errH
    strSQL = "zl_病人接诊完成_Cancel(" & mlng病人ID & ",'" & mstr挂号单 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    '刷新并定位病人，如果是病历卡片要手动刷新一下否则有的按钮不可用
    If tbcSub.Selected.Tag = "病历" Then Call mclsEPRs.zlRefresh(0, 0, 0, False, False)
    If tbcInTreat.Item(t在诊).Visible Then
        tbcInTreat.Item(t在诊).Selected = True
    End If
    If rptPati(PATI_RPT就诊).Visible Then
        Call LoadPatients("011", PATI_RPT就诊, mstr挂号单)
    Else
        Call LoadPatients("011")
    End If
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteCommunityIdentify()
'功能：补充社区身份验证
    Dim arrSQL As Variant, i As Long
    Dim colInfo As New Collection
    Dim int社区 As Integer, str社区号 As String
    Dim str出生日期 As String
        
    If gobjCommunity Is Nothing Or mPatiInfo.病人ID = 0 Or mPatiInfo.挂号ID = 0 Or mPatiInfo.社区 <> 0 Then Exit Sub
    
    If Not gobjCommunity.Identify(glngSys, p门诊医生站, int社区, str社区号, colInfo, mPatiInfo.病人ID, mPatiInfo.挂号ID) Then Exit Sub
    
    arrSQL = Array()
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_病人社区信息_Insert(" & mPatiInfo.病人ID & "," & int社区 & ",'" & str社区号 & "',1,Sysdate)"
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    str出生日期 = GetColItem(colInfo, "出生日期")
    If IsDate(str出生日期) Then
        str出生日期 = "To_Date('" & Format(str出生日期, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
    Else
        str出生日期 = "Null"
    End If
    arrSQL(UBound(arrSQL)) = "Zl_病人挂号记录_社区验证(" & mPatiInfo.病人ID & "," & mPatiInfo.挂号ID & "," & int社区 & "," & _
        "'" & GetColItem(colInfo, "姓名") & "','" & GetColItem(colInfo, "性别") & "','" & GetColItem(colInfo, "年龄") & "'," & _
        str出生日期 & ",'" & GetColItem(colInfo, "出生地点") & "','" & GetColItem(colInfo, "身份证号") & "'," & _
        "'" & GetColItem(colInfo, "民族") & "','" & GetColItem(colInfo, "国籍") & "','" & GetColItem(colInfo, "婚姻状况") & "'," & _
        "'" & GetColItem(colInfo, "职业") & "','" & GetColItem(colInfo, "家庭地址") & "','" & GetColItem(colInfo, "家庭电话") & "'," & _
        "'" & GetColItem(colInfo, "家庭地址邮编") & "','" & GetColItem(colInfo, "工作单位") & "','" & GetColItem(colInfo, "单位电话") & "'," & _
        "'" & GetColItem(colInfo, "单位邮编") & "','" & GetColItem(colInfo, "联系人姓名") & "','" & GetColItem(colInfo, "联系人关系") & "'," & _
        "'" & GetColItem(colInfo, "联系人电话") & "','" & GetColItem(colInfo, "联系人地址") & "','" & GetColItem(colInfo, "户口地址") & "','" & GetColItem(colInfo, "户口地址邮编") & "')"
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), "ExecuteCommunityIdentify"
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
    Call LoadPatients候诊
    If Not mbln显示预约病人 Then
        Call LoadPatients预约
    End If
    Call ReshDataQueue
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetColItem(colInfo As Collection, strItem As String) As String
    If colInfo Is Nothing Then Exit Function
    
    err.Clear: On Error Resume Next
    GetColItem = colInfo("_" & strItem)
    err.Clear: On Error GoTo 0
End Function

Private Sub SetRoomState(ByVal blnBusy As Boolean)
'功能：设置诊室忙闲状态
    On Error GoTo DBError
    gcnOracle.Execute "Update 门诊诊室 Set 缺省标志=" & IIf(blnBusy, 1, 0) & " Where 名称='" & mstr接诊诊室 & "' And 缺省标志<>" & IIf(blnBusy, 1, 0)
    On Error GoTo 0
    
    Me.stbThis.Panels(4).Text = "诊室" + IIf(blnBusy, "忙", "闲")
    Me.lblRoom.BackColor = IIf(blnBusy, COLOR_BUSY, COLOR_FREE)
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetReceiveToday(ByVal blnDo As Boolean, ByVal intStep As Integer)
'功能：当日接诊人数
'参数：blnDo true-访问数据库，false 不访问数据库。intStep 步长，－1或1
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If blnDo Then
        strSQL = "select count(1) as 人数 from 病人挂号记录 a where a.记录状态=1 and a.执行人=[1] and  a.执行时间 between Trunc(Sysdate) and Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名)
        mlng当日接诊人数 = Val(rsTmp!人数 & "")
    Else
        mlng当日接诊人数 = mlng当日接诊人数 + intStep
        If mlng当日接诊人数 < 0 Then mlng当日接诊人数 = 0
    End If
    
    Me.stbThis.Panels(3).Text = "当日接诊" & mlng当日接诊人数 & "人"
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetTimer()
    mintRefresh = Val(zlDatabase.GetPara("候诊刷新间隔", glngSys, p门诊医生站, 180))
    If mintRefresh <> 0 And mintRefresh < 30 Then mintRefresh = 30
    If mintRefresh = 0 Then
        timRefresh.Enabled = False
    Else
        timRefresh.Interval = 1000 '固定为1秒钟
        timRefresh.Enabled = True
    End If
End Sub

Private Sub timRefresh_Timer()
    Static lngSecond As Long
    Static strPreTime1 As String
    Dim curTime As Date
    
    If mbln消息语音 Then
        If Not mrsMsg Is Nothing Then
            If mrsMsg.RecordCount > 0 Then
                timRefresh.Enabled = False
                Call mclsMsg.PlayMsgSound(mrsMsg)
                Set mrsMsg = Nothing
                timRefresh.Enabled = True
            End If
        End If
    End If
    
    If Not mclsMipModule Is Nothing Then
        If mclsMipModule.IsConnect Then '使用了消息平台用新的刷新策略
            lngSecond = lngSecond + 1
            If lngSecond Mod 180 = 0 Then
                lngSecond = 0
                Call RefeshByMsg
            End If
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
    
    If mintRefresh = 0 Or mblnUnRefresh Or Me.hwnd <> GetForegroundWindow Then Exit Sub
    lngSecond = lngSecond + 1 '秒数
    If lngSecond Mod mintRefresh = 0 Then
        lngSecond = 0
        Call LoadPatients候诊
        If Not mbln显示预约病人 Then
            Call LoadPatients预约
        End If
        Call ReshDataQueue
    End If
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean, Optional ByVal strIDCard As String, Optional ByVal blnIsCard As Boolean _
                            , Optional ByVal lngPatiID As Long)
'功能：查找(下一个)病人
'参数：blnNext=是否查找下一个
'      strIDCard=当有值时，表示固定按身份证号查找
'      blnIsCard=是否是刷卡调用接收预约病人
    Static blnReStart As Boolean
    Dim intIdx As PatiType, i As Long
    Dim objControl As CommandBarControl
    Dim blnQueueFind As Boolean
    Dim objRpt As ReportControl
    Dim lngCol As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    If mintRPTIndex = -1 Then PatiIdentify.Text = "": Exit Sub
    
    '按其他方式查找后，自动刷身份证的继续查找则取消
    If strIDCard = "" And PatiIdentify.Text <> "" Then mstrIDCard = ""
    
    If Not blnNext And mstrFindType = "挂号单" Then
        PatiIdentify.Text = GetFullNO(PatiIdentify.Text, 12)
    End If
    PatiIdentify.SetFocus
    
    Set objRpt = rptPati(mintRPTIndex)
    
    '开始查找行
    If Not blnNext Or blnReStart Or mPr = -1 Then
        intIdx = 0: i = 0
    Else
        intIdx = mintRPTIndex
        If mPr <> -1 Then
            i = mPr + 1
        Else
            i = 0
        End If
    End If
    
     '查找病人
    If lngPatiID = 0 And Not mobjSquareCard Is Nothing And mstrFindType <> "就诊卡" And mstrFindType <> "标识号" And mstrFindType <> "挂号单" And mstrFindType <> "姓名" And mstrFindType <> "二代身份证" Then
        If mstrFindType = "IC卡" Then
            Call mobjSquareCard.zlGetPatiID("IC卡", PatiIdentify.Text, , lngPatiID)
        Else
            Call mobjSquareCard.zlGetPatiID(Val(PatiIdentify.objIDKind.GetCurCard.接口序号), PatiIdentify.Text, , lngPatiID)
        End If
    End If
    
    '查找病人
    If Check排队叫号 = True Then
        blnQueueFind = mobjQueue.FindQueue(IIf(PatiIdentify.objIDKind.GetCurCard.接口序号 > 0, _
                            PatiIdentify.objIDKind.GetCurCard.接口序号, _
                            IIf(PatiIdentify.objIDKind.GetCurCard.名称 = "标识号", "门诊号", PatiIdentify.objIDKind.GetCurCard.名称)), _
                            PatiIdentify.Text)
    End If
    If blnQueueFind = False Then
        For intIdx = intIdx To 4
            Set objRpt = rptPati(intIdx)
            For i = i To objRpt.Rows.Count - 1
                With objRpt.Rows(i)
                    If strIDCard <> "" Then '身份证自动识别强制优先
                        lngCol = IIf(intIdx = PATI_RPT候诊 Or intIdx = PATI_RPT预约, COL_HZ_身份证号, IIf(intIdx = PATI_RPT已诊, COL_YZ_身份证号, COL_JZ_身份证号))
                        If UCase(.Record(lngCol).Value) = UCase(strIDCard) Then Exit For
                    Else
                        lngCol = IIf(intIdx = PATI_RPT候诊 Or intIdx = PATI_RPT预约, COL_HZ_病人ID, IIf(intIdx = PATI_RPT已诊, COL_YZ_病人ID, COL_JZ_病人ID))
                        If Val(.Record(lngCol).Value) = lngPatiID And lngPatiID <> 0 Then Exit For
                        Select Case mstrFindType
                            Case "就诊卡"
                                lngCol = IIf(intIdx = PATI_RPT候诊 Or intIdx = PATI_RPT预约, COL_HZ_就诊卡号, IIf(intIdx = PATI_RPT已诊, COL_YZ_就诊卡号, COL_JZ_就诊卡号))
                                If .Record(lngCol).Value = PatiIdentify.Text Then Exit For
                            Case "标识号"
                                lngCol = IIf(intIdx = PATI_RPT候诊 Or intIdx = PATI_RPT预约, COL_HZ_门诊号, IIf(intIdx = PATI_RPT已诊, COL_YZ_门诊号, COL_JZ_门诊号))
                                If .Record(lngCol).Value = PatiIdentify.Text Then Exit For
                            Case "挂号单"
                                lngCol = IIf(intIdx = PATI_RPT候诊 Or intIdx = PATI_RPT预约, COL_HZ_NO, IIf(intIdx = PATI_RPT已诊, COL_YZ_NO, COL_JZ_NO))
                                If UCase(.Record(lngCol).Value) = UCase(PatiIdentify.Text) Then Exit For
                            Case "姓名"
                                lngCol = IIf(intIdx = PATI_RPT候诊 Or intIdx = PATI_RPT预约, COL_HZ_姓名, IIf(intIdx = PATI_RPT已诊, COL_YZ_姓名, COL_JZ_姓名))
                                If .Record(lngCol).Value Like "*" & PatiIdentify.Text & "*" Then Exit For
                            Case "二代身份证"
                                lngCol = IIf(intIdx = PATI_RPT候诊 Or intIdx = PATI_RPT预约, COL_HZ_身份证号, IIf(intIdx = PATI_RPT已诊, COL_YZ_身份证号, COL_JZ_身份证号))
                                If UCase(.Record(lngCol).Value) = UCase(PatiIdentify.Text) Then Exit For
                            Case "IC卡"
                                lngCol = IIf(intIdx = PATI_RPT候诊 Or intIdx = PATI_RPT预约, COL_HZ_IC卡号, IIf(intIdx = PATI_RPT已诊, COL_YZ_IC卡号, COL_JZ_IC卡号))
                                If UCase(.Record(lngCol).Value) = UCase(PatiIdentify.Text) Then Exit For
                            Case Else
                                lngCol = IIf(intIdx = PATI_RPT候诊 Or intIdx = PATI_RPT预约, COL_HZ_病人ID, IIf(intIdx = PATI_RPT已诊, COL_YZ_病人ID, COL_JZ_病人ID))
                                If Val(.Record(lngCol).Value) = lngPatiID Then Exit For
                        End Select
                    End If
                End With
            Next
            If i <= objRpt.Rows.Count - 1 Then Exit For
            i = 0
        Next
    
        If intIdx <= 4 Then
            blnReStart = False
            If Not objRpt.Visible Then
                If intIdx = PATI_RPT就诊 Then
                    tbcInTreat.Item(t在诊).Selected = True
                ElseIf intIdx = PATI_RPT回诊 Then
                    tbcInTreat.Item(t回诊).Selected = True
                ElseIf intIdx = PATI_RPT已诊 Then
                    tbcInTreat.Item(t完成).Selected = True
                ElseIf intIdx = PATI_RPT候诊 Then
                    tbcWait.Item(0).Selected = True
                ElseIf intIdx = PATI_RPT预约 Then
                    If tbcWait.Item(mint预约列表).Visible Then
                        tbcWait.Item(mint预约列表).Selected = True
                    End If
                End If
            End If

            '该行选中且显示在可见区域,并引发SelectionChanged事件
            Set objRpt.FocusedRow = objRpt.Rows(i)
            If objRpt.Visible Then objRpt.SetFocus
            
            '找到后自动进行接诊,预约病人自动接收
            If mbln自动接诊 And (intIdx = PATI_RPT候诊 Or intIdx = PATI_RPT预约) Then
                cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
                strTmp = objRpt.Rows(i).Record(COL_HZ_标识).Value
                If strTmp = "预" Then
                    If mstrFindType = "标识号" Or mstrFindType = "挂号单" Or mstrFindType = "姓名" Or mstrFindType = "二代身份证" Then Exit Sub
                    Call ExecuteReceive(blnIsCard)
                Else
                    Set objControl = cbsMain.FindControl(, conMenu_Manage_Receive, True, True)
                    If Not objControl Is Nothing Then
                        If objControl.Enabled Then Call cbsMain_Update(objControl) '首次执行，没有显示菜单前，事件没有执行
                        If objControl.Enabled Then objControl.Execute
                    End If
                End If
            End If
        Else
            blnReStart = True
            MsgBox IIf(blnNext, "后面已", "") & "找不到符合条件的病人。", vbInformation, gstrSysName
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
      
Private Function Check排队叫号() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查和创建排队叫号功能
    '返回：排队叫号功能所有的都合法,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-06-06 10:19:43
    '说明：需检查: 权限合法检查;启用了排队叫号的;创建排队叫号成功!
    '------------------------------------------------------------------------------------------------------------------------
    '排队叫号处理模式:1.代表分诊台分诊呼叫或医生主动呼叫;2-先分诊呼叫,再医生呼叫就诊.0-不排队叫号
    If mty_Queue.byt排队叫号模式 = 0 Then GoTo GOEND:
    If Not (InStr(mty_Queue.strQueuePrivs, ";基本;") > 0) Then GoTo GOEND:
    If mty_Queue.bln医生主动呼叫 = False And mty_Queue.byt排队叫号模式 = 1 Then GoTo GOEND:
    
    err = 0: On Error GoTo GOEND:
    If mobjQueue Is Nothing Then
        Set mobjQueue = CreateObject("zlQueueManage.clsQueueManage")
        err = 0: On Error GoTo ErrHand:
        mobjQueue.zlInitVar gcnOracle, glngSys, 0, IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数), mty_Queue.strQueuePrivs, CStr(mlngModul), False
        mobjQueue.zlSetToolIcon 24, True
        mobjQueue.IsShowFindTools = False
    End If
    Check排队叫号 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
GOEND:
    If Not mobjQueue Is Nothing Then mobjQueue.CloseWindows
    Set mobjQueue = Nothing

End Function

Private Sub ReshDataQueue()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：刷新排队叫号数据
    '编制：刘兴洪
    '日期：2010-06-07 15:27:57
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim varQueue() As String, strTemp As String, rsTemp As ADODB.Recordset, strSQL As String
    Dim str诊室 As String, str医生 As String, str科室 As String
    Dim intType As Integer
    
    If mobjQueue Is Nothing Then Exit Sub
    If Check排队叫号 = False Then Exit Sub
    '获取相关的队列名称
    '接诊范围：1=挂本人号的病人,2=本诊室病人,3=本科室病人
    mint接诊范围 = Val(zlDatabase.GetPara("接诊范围", glngSys, p门诊医生站, "2"))
    Dim strQueue() As String
    
    ReDim Preserve strQueue(1 To 1) As String
    str科室 = IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID)
    strQueue(1) = str科室
    str医生 = IIf(mstr接诊医生 = "", UserInfo.姓名, mstr接诊医生)
    str诊室 = mstr接诊诊室
    intType = 1
    Select Case mint接诊范围
    Case 1   '1=挂本人号的病人
        If Not mty_Queue.bln医生主动呼叫 Then
           str医生 = UserInfo.姓名  '64696,刘尔旋,2014-01-08,用登录人员的姓名过滤排队叫号队列
        End If
        If mlng接诊科室ID = 0 Then strQueue(1) = ""
        intType = 3
    Case 2  '2=本诊室病人
        If Not mty_Queue.bln医生主动呼叫 Then
           str诊室 = mstr接诊诊室
        End If
        If mlng接诊科室ID = 0 Then strQueue(1) = ""
        intType = 2
    Case 3  '3=本科室病人
    End Select
    
    '需要排队没有建档的病人
    strSQL = "" & _
    "   Select distinct  /*+ Rule*/  c.业务ID From 病人挂号记录 A ,排队叫号队列  C" & _
    "   Where A.id=C.业务ID and C.队列名称=[1]  and nvl(C.业务类型,0)=0 and nvl(A.病人ID,0) =0 And a.记录性质=1 And a.记录状态=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str科室)
    With rsTemp
        strTemp = ""
        Do While Not .EOF
            strTemp = strTemp & "," & Val(NVL(rsTemp!业务ID))
            .MoveNext
        Loop
        If strTemp <> "" Then strTemp = "0|" & Mid(strTemp, 2)
    End With
    Call mobjQueue.zlRefresh(strQueue, mty_Queue.strCurrQueueName, mty_Queue.lngcurr挂号ID, str诊室, str医生, strTemp, intType)
End Sub
 
Private Sub zlQueueStartus(intType As Integer, strNO As String, lng病人ID As Long)
  '------------------------------------------------------------------------------------------------------------------------
    '功能：功能操作后,
    '入参：2-换号;3-病人不就诊;4-病人待诊;5-病人完成就诊;6-病人取消就诊
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-06-03 14:15:46
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strQueueName As String, lngID As Long
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim i As Byte
    If Check排队叫号 = False Then Exit Sub
    
    strSQL = "SELECT ID,执行部门ID,诊室,执行人 From 病人挂号记录 where NO=[1] And 记录性质=1 And 记录状态=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Sub
    
    strQueueName = NVL(rsTemp!执行部门ID)
    If NVL(rsTemp!执行人) <> "" Then
        strQueueName = strQueueName & ":" & NVL(rsTemp!执行人)
    ElseIf NVL(rsTemp!诊室) <> "" Then
        strQueueName = strQueueName & ":" & NVL(rsTemp!诊室)
    End If
    
    lngID = Val(NVL(rsTemp!ID))
    Select Case intType
    Case 3   ' 病人不就诊;
        ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 3
    Case 4, 6   '病人待诊,'病人取消就诊
        ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 0
    Case 5  '病人完成就诊
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 4
    End Select
End Sub

Private Function Set病人挂号状态(ByVal lngState As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置病人挂号状态
    '入参：lngState : -1- 病人不就诊
    '                         0-病人待诊
    '出参：
    '返回：是否设置成功，病人不就诊时可以删除划价单据，当再次设置待诊时会设置不成功 返回False ,其他情况返回True
    '编制：刘兴洪
    '日期：2010-06-03 15:24:48
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim str划价NO As String
    
    If mstr挂号单 = "" Then Exit Function
    
    On Error GoTo errH
    
    If lngState = -1 Then
        '检查病人是否存在有效的医嘱
        strSQL = "Select 1 From 病人医嘱记录 Where 病人id = [1] And 挂号单 = [2]  And 医嘱状态 <> -1 And 医嘱状态 <> 1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
        If Not rsTmp.EOF Then
            MsgBox "该病人存在有效医嘱,不能设置为不就诊!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    End If

    '获取挂号划价单信息
    strSQL = "Select 摘要 From 门诊费用记录 Where NO = [1] And 记录性质 = 4 And 记录状态 = 1 And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号单)
    If Not rsTmp.EOF Then
        If rsTmp!摘要 & "" <> "" And InStr(rsTmp!摘要 & "", "划价:") <> 0 Then
            '获取挂号划价单信息,判断挂号划价单是否存在，不存在，则不允许将病人状态设置为待诊
            str划价NO = Mid(rsTmp!摘要 & "", Len("划价:") + 1)
            strSQL = "Select 1 From 门诊费用记录 Where NO = [1] And Mod(记录性质,10) = 1 And 记录状态 = 0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str划价NO)
            If rsTmp.EOF Then
                If lngState = 0 Then '设置为待诊
                    MsgBox "该挂号单的划价费用不存在，请退号后重新挂号!", vbInformation + vbDefaultButton1, gstrSysName
                    Exit Function
                End If
            Else
                If lngState = -1 Then '设置为不就诊
                    If MsgBox("该病人存在挂号单的划价费用，设置为不就诊时将删除该挂号单的划价费用，" & vbCrLf & "并且不能再恢复为待诊,是否继续?。", vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    
    gcnOracle.BeginTrans
        strSQL = "Zl_病人挂号记录_状态 ('" & mstr挂号单 & "'," & lngState & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Call zlQueueStartus(IIf(lngState = -1, 3, 4), mstr挂号单, mlng病人ID)
        'intType:intType:1-分诊;2-换号;3-病人不就诊;4-病人待诊;5-病人完成就诊;6-恢复就诊
        ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
    gcnOracle.CommitTrans
    MsgBox "操作成功!", vbInformation, gstrSysName
    
    Set病人挂号状态 = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ExecuteStopAndReuse(ByVal bln启用 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对就诊病人进行暂停就诊或启用诊断
    '入参:bln启用-true:启用已经停用的就诊病人
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-12-08 20:26:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, bln暂停 As Boolean
    Dim strNO As String, rsTemp As ADODB.Recordset
    Dim lngSelectedIndex As Long
    Dim rptRow As ReportRow
    
    lngSelectedIndex = mPr
    With rptPati(mintRPTIndex).Rows(mPr)
        bln暂停 = Val(.Record(COL_JZ_记录标志).Value) = 2
        If bln启用 Then
            If bln暂停 = False Then
                MsgBox "注意:" & vbCrLf & "    该病人还未暂停就诊,不能进行恢复暂停就诊!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Sub
            End If
        Else
            If bln暂停 Then
                MsgBox "注意:" & vbCrLf & "    该病人还启用暂停就诊,不能进行暂停就诊!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Sub
            End If
        End If
        strNO = .Record(COL_JZ_NO).Value
        
        strSQL = "Select ID From 病人挂号记录 where NO=[1] And 记录性质=1 And 记录状态=1"
        On Error GoTo errHandle
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTemp.EOF Then
            Exit Sub
        End If
    End With
    
    If Not bln启用 Then
        '回诊
        strSQL = "Zl_病人挂号记录_回诊(" & rsTemp!ID & ",NULL,NULL,NULL,1)"
    Else
        '取消回诊
        strSQL = "Zl_病人挂号记录_取消回诊(" & rsTemp!ID & ",1)"
    End If
    
    On Error GoTo errHandle
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If bln启用 Then
        '取消回诊时转到就诊列表中
        If tbcInTreat.Item(t在诊).Visible Then
            tbcInTreat.Item(t在诊).Selected = True
        End If
        If rptPati(PATI_RPT就诊).Visible Then
            Call LoadPatients("0101", PATI_RPT就诊, mstr挂号单)
        Else
            Call LoadPatients("0101")
        End If
    Else
        '标记回诊后自动定位下一个病人
        Call LoadPatients("0101")
        '回诊诊之后，自动定位到下一行
        If rptPati(mintRPTIndex).Rows.Count > lngSelectedIndex Then
            For Each rptRow In rptPati(mintRPTIndex).Rows
                If rptRow.GroupRow = False Then
                    If rptRow.Record.Index = lngSelectedIndex Then
                        Set rptPati(mintRPTIndex).FocusedRow = rptRow
                    End If
                End If
            Next
        ElseIf rptPati(mintRPTIndex).Rows.Count = lngSelectedIndex And lngSelectedIndex <> 0 Then
            For Each rptRow In rptPati(mintRPTIndex).Rows
                If rptRow.GroupRow = False Then
                    If rptRow.Record.Index = lngSelectedIndex - 1 Then
                        Set rptPati(mintRPTIndex).FocusedRow = rptRow
                    End If
                End If
            Next
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Set诊疗项目费用设置()
    Dim lng科室ID As Long
    
     On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "诊疗基础部件(ZLCISBase)没有正确安装，该功能无法执行。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    If mlng科室ID = 0 Then
        lng科室ID = mPatiInfo.科室ID
    Else
        lng科室ID = mlng科室ID
    End If
    If lng科室ID = 0 Then
        lng科室ID = UserInfo.部门ID
    End If
        
    Call gobjCISBase.CallSetClinicCharge(lng科室ID, 1, Me, gcnOracle, glngSys, gstrDBUser, E门诊调用, InStr(GetInsidePrivs(p门诊医生站), ";诊疗项目费用设置;") = 0)
End Sub
Private Sub SetFontSize(ByVal blnSetMainFont As Boolean)
'功能: 进行界面字体的统一设置
'参数: blnSetMainFont 是否设置主界面字体(用以区分子界面切换)
    Dim objFont As Object
    Dim lngFontSize As Long
    
    lngFontSize = IIf(mbytSize = 0, 9, 12)
    
    If blnSetMainFont Then
        Call zlControl.SetPubFontSize(Me, mbytSize)
        Set objFont = UCPatiVitalSigns.Font
        objFont.Size = lngFontSize
        Set UCPatiVitalSigns.Font = objFont
        Call picBasisNew_Resize
        Call picYZ_Resize
    End If

    Select Case tbcSub.Selected.Tag
        Case "病人"
            Call mobjPati.SetFontSize(mbytSize)
        Case "路径"
            Call mclsPath.SetFontSize(mbytSize)
        Case "医嘱"
            Call mclsAdvices.SetFontSize(mbytSize)
        Case "病历"
            Call mclsEPRs.SetFontSize(mbytSize)
                Case "新病历"
            On Error Resume Next
            Call mclsEMR.SetFontSize(mbytSize)
            err.Clear: On Error GoTo 0
    End Select
    
    If tbcRegist.Selected.Tag = "诊疗一览" Then
        Call mfrmView.SetFontSize(mbytSize)
    End If
        
End Sub

Private Function Check接诊控制(str操作 As String, strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人接诊控制
    '入参:str操作 -当前操作 strNo - 挂号单据号
    '出参:
    '返回:
    '编制:王吉
    '日期:2013-1-17 20:26:59
    '问题号:57566
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHanl:
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim rs预约时间 As Recordset
    Dim strMsg As String
    
    If mlng接诊控制 = 0 Then Check接诊控制 = True: Exit Function
    
    strSQL = "" & _
    "   Select  Nvl(A.预约时间,nvl(发生时间,sysdate)) - " & mlng提前接收时间 & "/24/60 as 挂号时间  " & _
    "   From 病人挂号记录 A " & _
    "   Where No=[1] And Nvl(A.预约时间,nvl(发生时间,sysdate))- " & mlng提前接收时间 & "*1/24/60>sysdate"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Check接诊控制 = True: Exit Function
    strMsg = "该病人需要在" & Format(rsTemp!挂号时间, "yyyy-mm-dd HH:MM:SS") & "后才允许进行" & str操作
    If mlng接诊控制 = 2 Then
        Check接诊控制 = (MsgBox(strMsg & ",您确定要进行" & str操作 & "吗？", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes)
    Else
        MsgBox strMsg & ",不允许" & str操作, vbInformation, gstrSysName
    End If
    Exit Function
ErrHanl:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mclsMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
'功能：处理门诊医生站接收到的消息
    Dim objXML As zl9ComLib.clsXML
    Dim rsTmp As ADODB.Recordset
    Dim rsMsg As ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim byt刷新 As Byte  '刷新方式：1-候诊列表，2－转诊列表
    Dim bln刷新 As Boolean
    
    On Error GoTo errH
    
    If strMsgItemIdentity = "ZLHIS_RECIPEAUDIT_001" Then
        If Mid(mstrNotifyAdvice, m处方审查, 1) = "1" Then
            '将消息加到消息列表中
            Set rsMsg = zlDatabase.ParseXMLToRecord(strMsgItemIdentity, strMsgContent)
            If rsMsg Is Nothing Then Exit Sub
            Call AddMsgToLis(rsMsg)
        End If
    End If
    
    strTmp = ",ZLHIS_REGIST_001,ZLHIS_REGIST_002,ZLHIS_CIS_007,ZLHIS_LIS_003,ZLHIS_PACS_005,"
    
    If InStr(strTmp, "," & strMsgItemIdentity & ",") = 0 Then Exit Sub
    
    strTmp = ""
    Set objXML = New zl9ComLib.clsXML
    Call objXML.OpenXMLDocument(strMsgContent)
    Select Case strMsgItemIdentity '获取挂号记录id
        Case "ZLHIS_REGIST_001", "ZLHIS_REGIST_002" '门诊患者挂号，门诊分诊通知。采取一分钟刷新一次方式，如果是第一条消息则立即刷新。
            byt刷新 = 1
            Call objXML.GetSingleNodeValue("register_id", strTmp)
        Case "ZLHIS_CIS_007" '门诊患者转诊。即时刷新，消息到来的时候就刷新，只刷新转诊列表
            byt刷新 = 2
            Call objXML.GetSingleNodeValue("clinic_id", strTmp)
    End Select
    
    If strTmp = "" Then Exit Sub
    
    strSQL = "Select 执行人,诊室,执行部门id,转诊医生,转诊诊室,转诊科室id From 病人挂号记录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strTmp))
    
    Set rsMsg = zlDatabase.ParseXMLToRecord(strMsgItemIdentity, strMsgContent)
    If Not rsMsg Is Nothing Then
        Call AddMsgToLis(rsMsg)
    End If
    
    If byt刷新 = 1 Then
        If mint接诊范围 = 1 And rsTmp!执行人 & "" = UserInfo.姓名 And (Not mbln要求分诊 Or mbln要求分诊 And rsTmp!诊室 & "" <> "") Then
            bln刷新 = True
        Else
            If (mint接诊范围 = 2 And rsTmp!诊室 & "" = mstr接诊诊室 Or mint接诊范围 = 3 And (Not mbln要求分诊 Or mbln要求分诊 And rsTmp!诊室 & "" <> "")) And _
                Val(rsTmp!执行部门ID & "") = IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID) And _
                (rsTmp!执行人 & "" = "" Or rsTmp!执行人 & "" = UserInfo.姓名) Then
                
                bln刷新 = True
            End If
        End If
        
        If bln刷新 Then
            mblnMsgOk = True
            If Not mblnFirstMsg Then     '是第一条消息
                mblnFirstMsg = True
                Call RefeshByMsg
            End If
        End If
    ElseIf byt刷新 = 2 Then
        If mint接诊范围 = 1 And rsTmp!转诊医生 & "" = UserInfo.姓名 Then
            bln刷新 = True
        Else
            If (mint接诊范围 = 2 And rsTmp!转诊诊室 & "" = mstr接诊诊室 Or mint接诊范围 = 3) And _
                Val(rsTmp!转诊科室ID & "") = IIf(mlng接诊科室ID = 0, UserInfo.部门ID, mlng接诊科室ID) And _
                UserInfo.姓名 <> IIf("" = rsTmp!执行人 & "", "无", rsTmp!执行人) And _
                (rsTmp!转诊医生 & "" = "" Or rsTmp!转诊医生 & "" = UserInfo.姓名) Then
                
                bln刷新 = True
            End If
        End If
        
        If bln刷新 Then
            Call LoadPatients("10001")
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefeshByMsg()
'功能：启用消息平台后使用的刷新方式
    If Not mblnMsgOk Then Exit Sub
    Call LoadPatients候诊
    If Not mbln显示预约病人 Then
        Call LoadPatients预约
    End If
    Call ReshDataQueue
    mblnMsgOk = False
End Sub

Private Function LoadNotify() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strSQL As String
    Dim strTmp As String
    Dim i As Long, j As Long
    Dim blnDo As Boolean
    Dim strTag As String
    
    mstrPreNotify = ""
    rptNotify.Records.DeleteAll
    If Mid(mstrNotifyAdvice, m危机值, 1) = "1" Then strTmp = strTmp & ",ZLHIS_LIS_003,ZLHIS_PACS_005"
    If Mid(mstrNotifyAdvice, m医嘱安排, 1) = "1" Then strTmp = strTmp & ",ZLHIS_OPER_001,ZLHIS_CIS_015,ZLHIS_CIS_005"
    If Mid(mstrNotifyAdvice, m处方审查, 1) = "1" Then strTmp = strTmp & ",ZLHIS_RECIPEAUDIT_001"
    If Mid(mstrNotifyAdvice, m传染病, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CIS_032,ZLHIS_CIS_033"
    If Mid(mstrNotifyAdvice, m备血完成, 1) = "1" And gbln血库系统 Then strTmp = strTmp & ",ZLHIS_BLOOD_001"   '启用血库流程才有此消息和参数
    If Mid(mstrNotifyAdvice, m用血审核, 1) = "1" And gbln血库系统 Then strTmp = strTmp & ",ZLHIS_BLOOD_004"  '启用血库才有此消息和参数
    If Mid(mstrNotifyAdvice, m输血反应, 1) = "1" And gbln血库系统 Then strTmp = strTmp & ",ZLHIS_BLOOD_006"  '启用血库才有此消息和参数
        
    strTmp = Mid(strTmp, 2)
    If strTmp = "" Then LoadNotify = True: Exit Function
       
    strSQL = "Select b.id,a.病人id,a.NO,a.id as 挂号ID,a.门诊号,a.姓名,a.执行时间 as 就诊时间,b.消息内容,b.类型编码, b.业务标识, b.优先程度, b.登记时间,a.险类,b.病人来源" & _
        " From 业务消息清单 B, 病人挂号记录 A" & _
        " Where b.就诊id=a.Id And a.执行人||''=[1]  And b.登记时间>=Trunc(Sysdate-" & (mintNotifyDay - 1) & ")" & _
        " And Nvl(b.是否已阅,0)=0 And instr(','||[2]||',',','||b.类型编码||',')>0 " & _
        " Order By b.优先程度 Desc, b.登记时间 Desc"

    Screen.MousePointer = 11

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mstr接诊医生, strTmp)

    For i = 1 To rsTmp.RecordCount
        Select Case rsTmp!类型编码
        Case "ZLHIS_CIS_032", "ZLHIS_CIS_033"
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
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!病人ID & "," & rsTmp!挂号ID & "," & rsTmp!类型编码 & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!病人ID & "," & rsTmp!挂号ID & "," & rsTmp!类型编码
                blnDo = True
            End If
        End Select
        
        If blnDo Then
            Call AddReportRow(rsTmp!病人ID & "," & rsTmp!挂号ID, rsTmp!病人ID, rsTmp!NO, NVL(rsTmp!姓名), NVL(rsTmp!门诊号), Format(rsTmp!就诊时间 & "", "yyyy-MM-dd HH:mm"), _
                 NVL(rsTmp!消息内容), rsTmp!类型编码 & "", rsTmp!优先程度 & "", Format(rsTmp!登记时间 & "", "yyyy-MM-dd HH:mm:ss"), NVL(rsTmp!业务标识), rsTmp!病人来源 & "", _
                 NVL(rsTmp!险类, 0), rsTmp!挂号ID, rsTmp!ID)
                        blnDo = False
        End If
        rsTmp.MoveNext
    Next
    rptNotify.Populate '缺省不选中任何行
    rptNotify.TabStop = rptNotify.Rows.Count > 0
    Screen.MousePointer = 0
    LoadNotify = True
    If mbln消息语音 Then
        If mclsMsg Is Nothing Then
            Set mclsMsg = New clsCISMsg
            Call mclsMsg.InitCISMsg(0)
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
    objRecord.Tag = arrInput(Index): Index = Index + 1         'Tag值 病人ID,挂号ID
    Set objItem = objRecord.AddItem(""): objItem.Icon = 6
    Set objItemIcon = objItem
    
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '病人id
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  'NO
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1 '姓名
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index))) '门诊号
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index))) '就诊时间
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))     '状态，内容
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    strNO = arrInput(Index)                            '消息编号
    objRecord.AddItem strNO: Index = Index + 1
    
    int优先级 = Val(arrInput(Index))                     '优先级
    objRecord.AddItem int优先级: Index = Index + 1
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  '登记日期
    
    str业务 = arrInput(Index): Index = Index + 1              '业务标识
    str病人来源 = arrInput(Index): Index = Index + 1          '病人来源
    
    int险类 = arrInput(Index): Index = Index + 1
    objRecord.AddItem str业务
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))     '挂号ID
    
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

Private Sub AddMsgToLis(ByVal rsMsg As ADODB.Recordset)
'功能：将接收到的消息加入提醒列表中
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim i As Long
    
    On Error GoTo errH
    
    strSQL = "select a.NO,a.姓名,a.执行人,a.门诊号,a.执行时间,a.险类 from 病人挂号记录 a where a.id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMsg!就诊id & ""))

    If mstr接诊医生 = rsTmp!执行人 & "" Then
        '判断列表是否已经有这类消息了
        For i = 0 To rptNotify.Rows.Count - 1
            If Not rptNotify.Rows(i).GroupRow Then
                If rptNotify.Rows(i).Record(C_消息).Value = rsMsg!类型编码 And rptNotify.Rows(i).Record.Tag = CStr(rsMsg!病人ID & "," & rsMsg!就诊id) Then
                    Exit Sub
                End If
            End If
        Next
        
        Call AddReportRow(rsMsg!病人ID & "," & rsMsg!就诊id, rsMsg!病人ID, rsMsg!NO, rsTmp!姓名, NVL(rsTmp!门诊号), Format(rsTmp!执行时间 & "", "yyyy-MM-dd HH:mm"), NVL(rsMsg!消息内容), _
             rsMsg!类型编码 & "", rsMsg!优先程度 & "", Format(rsMsg!登记时间 & "", "yyyy-MM-dd HH:mm:ss"), rsMsg!业务标识 & "", rsMsg!病人来源 & "", NVL(rsTmp!险类, 0), rsMsg!就诊id, 0)
        
        rptNotify.Populate
         
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub rptNotify_KeyUp(KeyCode As Integer, Shift As Integer)
'功能：自动进入医嘱校对、确认停止的执行界面
    Dim objControl As CommandBarControl
    Dim lngIndex As Long, lng病人ID As Long
    Dim lng医嘱ID As Long, lng挂号id As Long, lng消息ID As Long
    Dim str业务 As String, blnOk As Boolean
    Dim blnFinded As Boolean
    Dim strTmp As String, str姓名 As String, str门诊号 As String
    Dim strNO As String
    Dim str挂号单 As String
    Dim str消息内容 As String
    Dim i As Long
    Dim strPatis As String
    Dim blnOnePati As Boolean
    Dim blnTmp As Boolean
    
    If KeyCode = vbKeyReturn Then
        If rptNotify.SelectedRows.Count > 0 Then
            With rptNotify.SelectedRows(0).Record
                strNO = .Item(C_消息).Value
                str业务 = .Item(C_业务).Value
                str挂号单 = .Item(C_No).Value
                str消息内容 = .Item(C_状态).Value
                lng病人ID = Val(.Item(C_病人ID).Value)
                lng挂号id = Val(.Item(C_挂号Id).Value)
                lng消息ID = Val(.Item(C_Id).Value)
                str姓名 = .Item(c_姓名).Value
                str门诊号 = .Item(C_门诊号).Value
                lngIndex = .Index
            End With
    
            blnTmp = True
            
            If str挂号单 <> mstr挂号单 Then blnTmp = LocatePati(str挂号单)
            If str业务 <> "" Then      '找到病人后
                lng医嘱ID = Val(str业务)
            End If
            '如果是危机值消息，阅读后发送消息
            strTmp = ""
            If strNO = "ZLHIS_LIS_003" Then '检验
                strTmp = "ZLHIS_CIS_014"
            ElseIf strNO = "ZLHIS_PACS_005" Then '检查
                strTmp = "ZLHIS_CIS_025"
            End If
            If strTmp <> "" Then
                If Not (mclsMipModule Is Nothing) Then
                    If mclsMipModule.IsConnect Then
                        Call ZLHIS_CIS_MsgReadAfter(mclsMipModule, strTmp, lng病人ID, str姓名, , str门诊号, 1, lng挂号id, , mlng科室ID, , lng医嘱ID)
                    End If
                End If
            End If
            
            If strNO = "ZLHIS_RECIPEAUDIT_001" Then
                If str业务 = "合理用药审方" Then
                    blnTmp = CheckZLPass(Me, lng病人ID, lng挂号id)
                    If blnTmp Then
                        str消息内容 = "处方审查合格。"
                    Else
                        str消息内容 = ""
                    End If
                End If
                '先将卡片切换到医嘱卡片方便查找菜单
                Call LocatedCard("医嘱")
                cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
                If str消息内容 = "处方审查合格。" Then
                    '弹出消息发送窗体
                    Set objControl = cbsMain.FindControl(, conMenu_Edit_Send, True, True)
                    If Not objControl Is Nothing Then
                        If objControl.Enabled Then objControl.Execute
                    End If
                Else
                    '医嘱编辑窗体
                    Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
                    If Not objControl Is Nothing Then
                        If objControl.Enabled Then objControl.Execute
                    End If
                End If
            End If
            
            If strNO = "ZLHIS_CIS_032" Then
                Call mclsDis.ShowDisRegist(Me, 1, Val(str业务), lng病人ID, 0, str挂号单)
            End If

            If strNO = "ZLHIS_BLOOD_006" Then
                If gobjPublicBlood Is Nothing And gbln血库系统 Then InitObjBlood
                blnOk = gobjPublicBlood.zlIsBloodMessageDone(2, lng病人ID, lng挂号id, 1, mlng科室ID)
                If blnOk Then
                    Call rptNotify.Records.RemoveAt(lngIndex)
                Else
                    If FuncTraReaction(Val(str业务), mlngModul, False, IIf(InStr(1, str业务, ":") > 0, Val(Split(str业务, ":")(1)), 0)) Then
                        If gobjPublicBlood.zlIsBloodMessageDone(2, lng病人ID, lng挂号id, 1, mlng科室ID) Then
                            Call rptNotify.Records.RemoveAt(lngIndex)
            End If
                    End If
                End If
            End If
            If strNO = "ZLHIS_CIS_033" Then
            '传染病报告反修改消息阅读
                blnOk = ReadMsgCIS033(lng病人ID, lng挂号id, str业务, lng消息ID)
                If blnOk Then Call rptNotify.Records.RemoveAt(lngIndex)
            End If
            
            If strNO <> "ZLHIS_CIS_033" And strNO <> "ZLHIS_BLOOD_006" Then
                blnOk = ReadMsg(lng病人ID, lng挂号id, strNO, str业务, lng消息ID, str挂号单)
                If blnOk Then Call rptNotify.Records.RemoveAt(lngIndex)
            End If
            Call rptNotify.Populate
        End If
    End If
End Sub

Private Sub rptNotify_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptNotify_KeyUp(vbKeyReturn, 0)
End Sub

Private Sub rptNotify_SelectionChanged()
    Dim str挂号单 As String
    
    If rptNotify.SelectedRows.Count = 0 Then Exit Sub  '非正常情况
    
    str挂号单 = rptNotify.SelectedRows(0).Record.Item(C_No).Value
 
    If str挂号单 <> mstr挂号单 Then Call LocatePati(str挂号单)
    
End Sub

Private Function ReadMsg(ByVal lng病人ID As Long, ByVal lng挂号id As Long, ByVal strNO As String, ByVal str业务 As String, ByVal lng消息ID As Long, ByVal str挂号单 As String) As Boolean
'功能：阅读消息
'说明：消息阅读方式目前有3种：按消息编译码阅读，消息ID阅读，按业务标识阅读
    Dim strSQL As String
    Dim lng科室ID As Long
    Dim str医嘱ID As String
    Dim blnDo As Boolean
    Dim lng危急值ID As Long  '本次处理的危急值记录ID
    Dim strSQLReadMsg As String
    Dim blnHis危急值 As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim objControl As CommandBarControl
    
    If mlng接诊科室ID = 0 Then
        lng科室ID = UserInfo.部门ID
    Else
        lng科室ID = mlng接诊科室ID
    End If
    blnDo = True
    
    On Error GoTo errH
    
    strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng挂号id & ",'" & strNO & "',1,'" & UserInfo.姓名 & "'," & lng科室ID
    Select Case strNO
    Case "ZLHIS_LIS_003", "ZLHIS_PACS_005"
        strSQL = strSQL & ",null,null,'" & str业务 & "'"
    Case "ZLHIS_CIS_032", "ZLHIS_CIS_033"
        strSQL = strSQL & ",null," & lng消息ID
    End Select
    strSQL = strSQL & ")"
    
    strSQLReadMsg = strSQL
    
    If strNO = "ZLHIS_LIS_003" Or strNO = "ZLHIS_PACS_005" Then
        If mbln危急值 Then
            '危急值消息相关处理
            Call mobjKernel.ShowDealCritical(Me, lng病人ID, 0, str挂号单, lng危急值ID)
            
            If lng危急值ID <> 0 Then
                strSQL = "select a.标本id,a.处理情况,a.确认人 from 病人危急值记录 a where a.id=[1] and a.确认人 is not null"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng危急值ID)
                If Not rsTmp.EOF Then
                    '将消息设置为已阅
                    Call zlDatabase.ExecuteProcedure(strSQLReadMsg, Me.Caption)
                    
                    '如果是LIS危急值调用LIS接口
                    If strNO = "ZLHIS_LIS_003" Then
                        Call InitObjLis(p门诊医生站)
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
                Call InitObjLis(p门诊医生站)
                If Not gobjLIS Is Nothing Then
                    blnDo = gobjLIS.GetReadNotify(Me, str医嘱ID, UserInfo.姓名)
                End If
            End If
        End If
        If strNO = "ZLHIS_BLOOD_004" Then
            '用血审核消息的阅读状态设置在血库部件内部，临床不用执行阅读消息过程
            strSQL = "select 1 from 病人医嘱记录 a where a.挂号单=[1] and a.医嘱状态=1 and a.诊疗类别='K' and a.检查方法='1' and a.审核状态=1 and rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str挂号单)
            If Not rsTmp.EOF Then
                '如果有数据，则弹出医嘱修改界面，本过程中不执行消息阅读SQL语句
                '先将卡片切换到医嘱卡片方便查找菜单
                Call LocatedCard("医嘱")
                cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
                '医嘱编辑窗体
                Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
                If Not objControl Is Nothing Then
                    If objControl.Enabled Then objControl.Execute
                End If
                ReadMsg = True
                Exit Function
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

Private Function LocatePati(ByVal strTag As String) As Boolean
'功能：通过挂号单定位，当前可以见的列表
    Dim lngCol As Long
    Dim objRow As ReportRow
    Dim objRpt As ReportControl
    Dim blnEnabled  As Boolean
    
    Dim i As Long
    For i = 0 To 4
        Set objRpt = rptPati(i)
        If objRpt.Visible Then
            lngCol = IIf(i = PATI_RPT候诊 Or i = PATI_RPT预约, COL_HZ_NO, IIf(i = PATI_RPT已诊, COL_YZ_NO, COL_JZ_NO))
            For Each objRow In objRpt.Rows
                If objRow.GroupRow Then objRow.Expanded = True
                If Not objRow.GroupRow Then
                    If objRow.Record(lngCol).Value = strTag Then
                        blnEnabled = timRefresh.Enabled
                        timRefresh.Enabled = False '避免连锁引起刷新提醒内容
                        Set objRpt.FocusedRow = objRow '选中,显示,[激活Change事件]
                        timRefresh.Enabled = blnEnabled
                        LocatePati = True: Exit Function
                    End If
                End If
            Next
        End If
    Next
End Function

Private Sub picMore_Resize()
    On Error Resume Next
    lblEdit(txtInfo摘要).Left = 10
    lblEdit(txtInfo摘要).Top = 10
    txtInfo(txtInfo摘要).Left = lblEdit(txtInfo摘要).Left + 20 + lblEdit(txtInfo摘要).Width
    txtInfo(txtInfo摘要).Width = picMore.Width - txtInfo(txtInfo摘要).Left
    UCPatiVitalSigns.Top = txtInfo(txtInfo摘要).Top + txtInfo(txtInfo摘要).Height + 60
    UCPatiVitalSigns.Left = 10
End Sub

Private Sub picBasisNew_Resize()
    On Error Resume Next
    '此处可以固定高度
    
    lblUrg.FontName = "黑体"
    lblUrg.FontSize = IIf(mbytSize = 0, 14, 18)
    
    lblRec.FontName = "黑体"
    lblRec.FontSize = IIf(mbytSize = 0, 14, 18)
    
    If err.Number <> 0 Then err.Clear

    picPatient.Left = 10
    picPatient.Top = 10
    
    picPatient.Height = picBasisNew.Height - picPatient.Top - 60
    picPatient.Width = picPatient.Height
    imgPatient.Height = picPatient.Height
    imgPatient.Width = picPatient.Width
    
    lblLink(lblLink文件).Left = picPatient.Left + picPatient.Width + 80
    lblLink(lblLink文件).Top = picPatient.Top
    
    lblLink(lblLink采集).Left = lblLink(lblLink文件).Left
    lblLink(lblLink采集).Top = picBasisNew.Height / 2 - 120
    
    lblLink(lblLink清除).Left = lblLink(lblLink文件).Left
    lblLink(lblLink清除).Top = picPatient.Height + picPatient.Top - lblLink(lblLink清除).Height
    
    lblLink(lblLink修改).Left = lblLink(lblLink清除).Left + lblLink(lblLink清除).Width + 180
    lblLink(lblLink修改).Top = lblLink(lblLink清除).Top
    
    txtInfo(txtInfo姓名).Top = 100
    txtInfo(txtInfo姓名).Left = lblLink(lblLink修改).Left
    txtInfo(txtInfo姓名).FontSize = IIf(mbytSize = 0, 12, 15)
    txtInfo(txtInfo姓名).Width = IIf(mbytSize = 0, 1400, 1800)
    
    txtInfo(txtInfo性别).Top = txtInfo(txtInfo姓名).Top + txtInfo(txtInfo姓名).Height - txtInfo(txtInfo性别).Height + 160
    txtInfo(txtInfo性别).Left = txtInfo(txtInfo姓名).Left + txtInfo(txtInfo姓名).Width + 50
    
    txtInfo(txtInfo年龄).Top = txtInfo(txtInfo性别).Top + txtInfo(txtInfo性别).Height - txtInfo(txtInfo年龄).Height
    txtInfo(txtInfo年龄).Left = txtInfo(txtInfo性别).Left + txtInfo(txtInfo性别).Width + 100
    
    Call zlControl.SetPubCtrlPos(False, -1, txtInfo(txtInfo年龄), 250, lblEdit(txtInfo出生日期), 30, txtInfo(txtInfo出生日期), 250, lblEdit(txtInfo付费方式), 30, fraPayType)
    fraPayType.Top = lblEdit(txtInfo付费方式).Top - 30
    
    fraPayType.Width = cboPayType.Width
    fraPayType.Height = cboPayType.Height - 60
    
    linPayType.x1 = fraPayType.Left - 20
    linPayType.y1 = fraPayType.Top + fraPayType.Height
    linPayType.x2 = linPayType.x1 + fraPayType.Width
    linPayType.y2 = linPayType.y1
    
    lblEdit(12).Left = lblEdit(10).Left
    lblEdit(12).Top = lblLink(lblLink修改).Top + 10
    
    txtInfo(txtInfo就诊卡号).Width = 1300
    Call zlControl.SetPubCtrlPos(False, -1, lblLink(lblLink修改), 250, lblEdit(txtInfo号类), 30, txtInfo(txtInfo号类), 150, lblEdit(txtInfo就诊卡号), 30, txtInfo(txtInfo就诊卡号), 150, lblEdit(txtInfo医保卡号), 30, txtInfo(txtInfo医保卡号), 150, lblEdit(txtInfo费别), 30, fraBillType)
    fraBillType.Top = lblEdit(txtInfo费别).Top - 30
    
    fraBillType.Width = cboBillType.Width
    fraBillType.Height = cboBillType.Height - 60
    
    linBillType.x1 = fraBillType.Left - 20
    linBillType.y1 = fraBillType.Top + fraBillType.Height
    linBillType.x2 = linBillType.x1 + fraBillType.Width
    linBillType.y2 = linBillType.y1
    
    lblMore.Top = lblEdit(txtInfo医保卡号).Top
    lblMore.Left = picBasisNew.Width - lblMore.Width - 40

    lblUrg.Top = 200
    lblUrg.Left = picBasisNew.Width - lblUrg.Width - 40
    lblRec.Top = lblUrg.Top
    lblRec.Left = lblUrg.Left - lblRec.Width - 20
    
    lblLink(lblLink显示).Left = fraBillType.Left + fraBillType.Width + 200
    lblLink(lblLink显示).Top = txtInfo(txtInfo医保卡号).Top
End Sub

Private Sub lblLink_Click(Index As Integer)
    Dim strPictureFile As String
    Dim objControl As CommandBarControl
    
    On Error GoTo errH
    
    If lblLink(Index).ForeColor <> &HC00000 Then Exit Sub
    
    Select Case Index
    Case lblLink文件
        With cmdialog
            .CancelError = False
            .Flags = cdlOFNHideReadOnly
            .Filter = "(*.bmp)|*.bmp"
            .FilterIndex = 2
            .ShowOpen
            strPictureFile = .FileName
            If strPictureFile = "" Then Exit Sub
            imgPatient.Picture = LoadPicture(strPictureFile)
            picPatient.Tag = strPictureFile
        End With
        Call SetPatPicture(mPatiInfo.病人ID, False)
    Case lblLink清除
        If picPatient.Tag <> "" Then
            If SetPatPicture(mPatiInfo.病人ID, True) Then
                imgPatient.Picture = imgDefual.Picture
                picPatient.Tag = ""
            End If
        End If
    Case lblLink采集, lblLink修改
        If mobjPatient Is Nothing Then
            On Error Resume Next
            Set mobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
            err.Clear: On Error GoTo 0
        End If
        If mobjPatient Is Nothing Then
            MsgBox "创建病人信息公共部件（zlPublicPatient.clsPublicPatient）失败！", vbInformation, Me.Caption
            Exit Sub
        End If
        Call mobjPatient.zlInitCommon(gcnOracle, glngSys, UserInfo.用户名)
        
        If lblLink采集 = Index Then
            If mobjPatient.PatiImageGatherer(Me, strPictureFile) = False Then Exit Sub
            Set imgPatient.Picture = LoadPicture(strPictureFile)
            picPatient.Tag = strPictureFile
            Call SetPatPicture(mPatiInfo.病人ID, False)
        Else
            If mobjPatient.ModiPatiBaseInfo(Me, "门诊医生工作站", mPatiInfo.病人ID, mPatiInfo.挂号ID, 1, False) Then
                '修改成功后刷新，整个界面统一刷新
                Set objControl = cbsMain.FindControl(, conMenu_View_Refresh, , True)
                If Not objControl Is Nothing Then
                    If objControl.Enabled Then objControl.Execute
                End If
            End If
        End If
    Case lblLink显示
        If mblnPatiDetail Then
            mblnPatiDetail = False
            lblLink(Index).Caption = "显示其它信息↓"
        Else
            mblnPatiDetail = True
            lblLink(Index).Caption = "隐藏其它信息↑"
        End If
        Call cbsMain_Resize
        Call zlDatabase.SetPara("显示病人详细信息", IIf(mblnPatiDetail, 1, 0), glngSys, p门诊医生站, InStr(";" & mstrPrivs & ";", ";参数设置;") > 0)
    End Select
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    err.Clear
End Sub

Private Function SetPatPicture(ByVal lng病人ID As Long, ByVal blnDel As Boolean) As Boolean
'功能:设置病人照片
'入参:lng病人ID - 病人ID，blnDel true 删除照片，false 保存照片
    Dim strFile As String, strSQL As String
    On Error GoTo errH

    If blnDel Then
        If MsgBox("病人" & txtInfo(txtInfo姓名).Text & "的照片将被删除，是否继续？", vbDefaultButton2 + vbYesNo + vbQuestion, Me.Caption) = vbNo Then
            Exit Function
        End If
        strSQL = strSQL & "Zl_病人照片_Delete(" & lng病人ID & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Else
        '图片没有被清除，则重新插入图片
        If picPatient.Tag <> "" Then
            strFile = picPatient.Tag
            If Sys.SaveLob(glngSys, 27, lng病人ID, strFile) = False Then
                MsgBox "保存照片有误,请确认文件是否被删除!", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    SetPatPicture = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mobjPati_EditFullDoc(ByVal lngEPRFileID As Long, ByVal lngFileID As Long, ByVal strDoctor As String, ByVal strIn As String)
'功能：病人信息卡片保存了病历后
    Dim blnDoc As Boolean
    
    If InStr(";" & GetPrivFunc(glngSys, p门诊病历管理) & ";", ";病历书写;") > 0 Then
        blnDoc = mlng科室ID <> 0 And mlng科室ID = mPatiInfo.科室ID And mstr挂号单 = mPatiInfo.挂号单 And _
                 (lngFileID = 0 And lngEPRFileID <> 0 Or lngFileID <> 0) And (mintActive = pt就诊 Or mintActive = pt回诊)
        If blnDoc And lngFileID <> 0 And strIn = "0" Then  '没有修改他人病历的权限
            blnDoc = strDoctor = UserInfo.姓名
        End If
        
        If blnDoc Then
            If mobjEPRDoc Is Nothing Then
                Set mobjEPRDoc = New zlRichEPR.cEPRDocument
            End If
            If lngFileID = 0 And lngEPRFileID <> 0 Then '如果没有新建则新建
                Call mobjEPRDoc.InitEPRDoc(0, 2, lngEPRFileID, 1, mPatiInfo.病人ID, mPatiInfo.挂号ID, , mPatiInfo.科室ID, , False)
            Else
                Call mobjEPRDoc.InitEPRDoc(1, 2, lngFileID, 1, mPatiInfo.病人ID, mPatiInfo.挂号ID, , mPatiInfo.科室ID, , False)
            End If
            Call mobjEPRDoc.ShowEPREditor(Me)
        Else
            MsgBox "当前病历不能修改。", vbInformation, Me.Caption
        End If
    Else
        MsgBox "您没有病历书写的权限。", vbInformation, Me.Caption
    End If
End Sub

Private Sub mobjPati_EPRRefresh()
    With mPatiInfo
        Call mclsEPRs.zlRefresh(mlng病人ID, mlng挂号ID, mlng科室ID, mlng科室ID = .科室ID And (.类型 = pt就诊 Or .类型 = pt回诊) And mlng病人ID <> 0, .数据转出, True)
    End With
End Sub

Private Sub mobjPati_UpdatePatiInfo(ByVal strBirthday As String, ByVal strAge As String, ByVal strSex As String, ByVal strTag As String)
'功能：更新病人信息
    Dim strTmp As String
    If strBirthday <> "" Then
        txtInfo(txtInfo出生日期).Text = Format(strBirthday, "yyyy-MM-dd")
    End If
    If strAge <> "" Then
        txtInfo(txtInfo年龄).Text = strAge
    End If
    If strSex <> "" Then
        txtInfo(txtInfo性别).Text = strSex
    End If
    If (strAge <> "" Or strSex <> "") And (mintRPTIndex = PATI_RPT就诊 Or mintRPTIndex = PATI_RPT回诊) Then
        If mPr <> -1 And rptPati(mintRPTIndex).Visible Then
            With rptPati(mintRPTIndex).Rows(mPr)
                .Record(COL_JZ_性别).Value = strSex
                .Record(COL_JZ_年龄).Value = strAge
            End With
            rptPati(mintRPTIndex).Populate
        End If
    End If
    If strTag <> "" Then
        txtInfo(txtInfo摘要).Text = IIf("NULL" = strTag, "", strTag)
    End If
End Sub

Private Sub mobjPati_UpdateDiagInfo(ByVal str疾病ID As String, ByVal str诊断ID As String, ByVal strTag As String)
'功能：传染病检查
    Dim blnNo As Boolean
    Dim rsTmp  As ADODB.Recordset
    Dim blnNotView As Boolean
    
    If InStr(";" & GetPrivFunc(glngSys, p疾病报告填写) & ";", ";病历书写;") > 0 Then
        Set rsTmp = mclsDisease.SatisfyEditDiseaseDoc(mlng病人ID, mlng挂号ID, mlng科室ID, str疾病ID, str诊断ID)
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then
                If Not mclsDis.ShowDiseaseStation(Me, mlng病人ID, mlng挂号ID, 1, mlng科室ID, str疾病ID, str诊断ID, blnNotView) Then
                    Call mclsDisease.EditDiseaseReport(Me, rsTmp, mlng病人ID, mlng挂号ID, 1, mlng科室ID, blnNo)
                    If blnNo Then
                        Call mclsDis.EditNotFillReason(Me, mlng病人ID, mlng挂号ID, 1)
                    End If
                ElseIf blnNotView Then
                    Call mclsDisease.EditDiseaseReport(Me, rsTmp, mlng病人ID, mlng挂号ID, 1, mlng科室ID, blnNo)
                    If blnNo Then
                        Call mclsDis.EditNotFillReason(Me, mlng病人ID, mlng挂号ID, 1)
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub LocatedCard(ByVal strTag As String)
'功能：定位到指定的页签卡片，内部页签
    Dim i As Long
    '1.先定位到本次就诊
    If tbcRegist.Selected.Caption <> "本次就诊" Then
        tbcRegist.Item(mbyt本次就诊).Selected = True
    End If
    If tbcSub.Selected.Tag <> strTag Then
        For i = 0 To tbcSub.ItemCount - 1
            If tbcSub.Item(i).Visible Then
                If tbcSub.Item(i).Tag = strTag Then
                    tbcSub.Item(i).Selected = True
                    Exit For
                End If
            End If
        Next
    End If
End Sub

Private Sub InitCboData()
'功能：加载下拉列表值
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    
    strSQL = "Select 编码, 名称 From 医疗付款方式 Order By 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With cboPayType
        .Clear
        For i = 1 To rsTmp.RecordCount
            .AddItem rsTmp!编码 & "-" & rsTmp!名称
            .ItemData(.NewIndex) = Val(rsTmp!编码 & "")
            rsTmp.MoveNext
        Next
    End With
    
    strSQL = "Select 编码, 名称 From 费别 Order By 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With cboBillType
        .Clear
        For i = 1 To rsTmp.RecordCount
            .AddItem rsTmp!名称 & ""
            .ItemData(.NewIndex) = Val(rsTmp!编码 & "")
            rsTmp.MoveNext
        Next
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Tip病人自动完成() As Boolean
'功能：将当前病人设为完成就诊回诊
    Dim objMsg As New zl9ComLib.clsAirBubble
    Dim varPatis As Variant
    Dim lng挂号id As Long
    Dim i As Long, j As Long
    Dim strSQL As String
    Dim rsPati As ADODB.Recordset
    Dim strInfo As String
    Dim blnDo As Boolean
    Dim str病人姓名1 As String
    Dim str病人姓名2 As String
    Dim intType As Integer
    If mstr挂号IDs = "" Then Exit Function
    On Error GoTo errH
    strInfo = mstr挂号IDs
    varPatis = Split(strInfo, ",")
    For i = 0 To UBound(varPatis)
        lng挂号id = Val(varPatis(i))
        If lng挂号id <> 0 And lng挂号id <> mPatiInfo.挂号ID Then
            blnDo = False
            intType = 0
            '首先判断当前病人是不是已经完成就诊和回诊了
            strSQL = "select ID,NO,姓名,病人ID,执行状态,转诊状态,decode(记录标志,2,1,3,1,0) as 回诊 from 病人挂号记录 where 记录性质=1 And 记录状态=1 and id=[1]"
            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng挂号id)
            If rsPati.RecordCount = 1 Then
                If 2 = Val(rsPati!执行状态 & "") Then
                    blnDo = CanAutoFinish(Val(rsPati!病人ID & ""), rsPati!NO & "", lng挂号id, intType)
                End If
            End If
            
            If blnDo Then
                If intType <> 2 Then
                    str病人姓名1 = str病人姓名1 & "，" & rsPati!姓名
                    '完成就诊之前如果已经转诊的先取消转诊
                    If Not IsNull(rsPati!转诊状态) Then
                        If Val(rsPati!转诊状态 & "") = 0 Then
                            strSQL = "Zl_病人挂号记录_转诊('" & rsPati!NO & "',Null)"
                            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                        End If
                    End If
                    Call ExecuteFinishInSide(rsPati!NO & "", Val(rsPati!病人ID & ""), lng挂号id)
                ElseIf intType <> 0 Then
                    If Val(rsPati!回诊 & "") = 0 Then
                        strSQL = "Zl_病人挂号记录_回诊(" & lng挂号id & ",NULL,NULL,NULL,1)"
                        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                        '只提示未标为回诊的，已经标记了的就不提示了
                        str病人姓名2 = str病人姓名2 & "，" & rsPati!姓名
                    End If
                End If
            End If
        End If
    Next
    strInfo = ""
    If str病人姓名1 <> "" Then
        str病人姓名1 = Mid(str病人姓名1, 2) & " 病人自动完成就诊，可在已诊列表中查看。"
    End If
    
    If str病人姓名2 <> "" Then
        str病人姓名2 = Mid(str病人姓名2, 2) & " 病人自动标记回诊，可在就诊列表中查看。"
    End If
    strInfo = IIf("" = str病人姓名1, "", str病人姓名1 & vbCrLf) & IIf("" = str病人姓名2, "", str病人姓名2)
    If strInfo <> "" Then
        Call objMsg.OpenTransparentAirBubble(Me, strInfo, 2, 2, 15, &HFF8080, &HFFFFFF, , 3, , , 无)
        Set mobjMsg = objMsg
    End If
    mstr挂号IDs = ""
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CanAutoFinish(ByVal lng病人ID As Long, ByVal strNO As String, ByVal lng挂号id As Long, ByRef intType As Integer) As Boolean
'功能：当前病人是否可以自动进入下一个环节，完成就诊或者回诊
'参数：intType 1-完成就诊，2－回诊
    Dim i As Long, j As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsPati As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim blnDo As Boolean, blnSigned As Boolean, blnOK病历 As Boolean
    Dim objEmr As Object
    Dim str医嘱IDs As String
    Dim lngTmp As Long, lngTmp1 As Long
    Dim str病人姓名1 As String
    Dim str病人姓名2 As String
    
    On Error GoTo errH
    intType = 1
    '1.病历检查
    strSQL = "select 1 from 电子病历记录 where 病人ID=[1] and 主页ID=[2] and 签名级别<>1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng挂号id)
    If rsTmp.EOF Then
        blnSigned = True
        If GetInsidePrivs(p新版门诊病历, True) <> "" Then
            On Error Resume Next
            Set objEmr = CreateObject("zl9EmrInterface.ClsEmrInterface")
            If Not objEmr Is Nothing Then
                Call objEmr.CheckOutEPRIsAllSign(lng挂号id, blnSigned)
            End If
            err.Clear: On Error GoTo 0
            On Error GoTo errH
        End If
        If blnSigned Then
            blnOK病历 = True
        End If
    Else
        Exit Function
    End If
    
    '2.医嘱检查
    If blnOK病历 Then
        strSQL = "select a.id,a.相关id,a.序号,a.医嘱状态,a.诊疗类别," & _
            " NVL(a.执行标记,0) as 执行标记, Nvl(a.执行性质,0) as 执行性质,Nvl(a.皮试结果,'无') as 皮试结果 from 病人医嘱记录 a where a.医嘱状态<>4 and a.挂号单=[1] and a.病人ID+0=[2]"
        Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, lng病人ID)
        If rsAdvice.RecordCount = 0 Then
            '无有效医嘱
            Exit Function
        End If
        
        rsAdvice.Filter = "医嘱状态=1 And 执行标记<>-1 And 执行性质<>0 And 皮试结果<>'免试'"
        If rsAdvice.RecordCount <> 0 Then
            '未发送的医嘱
            Exit Function
        End If
        
        '已经发送的检查检验医嘱
        rsAdvice.Filter = "(医嘱状态=8 and 诊疗类别='D') or (医嘱状态=8 and 诊疗类别='C')"
        str医嘱IDs = ""
        For j = 1 To rsAdvice.RecordCount
            lngTmp = Val(rsAdvice!ID & "")
            lngTmp1 = Val(rsAdvice!相关ID & "")
            
            If InStr("," & str医嘱IDs & ",", "," & lngTmp & ",") = 0 Then
                str医嘱IDs = str医嘱IDs & "," & lngTmp
            End If
            
            If InStr("," & str医嘱IDs & ",", "," & lngTmp1 & ",") = 0 Then
                str医嘱IDs = str医嘱IDs & "," & lngTmp1
            End If
            rsAdvice.MoveNext
        Next
        
        If str医嘱IDs <> "" Then
            str医嘱IDs = Mid(str医嘱IDs, 2)
            strSQL = "select 1 from 病人医嘱发送 a where a.医嘱id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) and a.执行状态<>1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str医嘱IDs)
            If Not rsTmp.EOF Then
                intType = 2
            End If
        End If
    End If
    CanAutoFinish = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadMsgCIS033(ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal str标识 As String, ByVal lng消息ID As Long) As Boolean
'功能：传染病报告反修改消息阅读
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim lng文件ID As Long
    Dim lng科室ID As Long
    Dim objControl As CommandBarControl
    
    On Error GoTo errH
    'conMenu_Edit_Modify 3003 修改按钮。
    lng文件ID = Val(Split(str标识, ",")(0))
    
    strSQL = "Select 1 From 疾病申报记录 where 文件ID=[1] and 处理状态=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng文件ID, 4)
    If rsTmp.RecordCount = 0 Then
    '把消息标记为已读
        If mlng接诊科室ID = 0 Then
            lng科室ID = UserInfo.部门ID
        Else
            lng科室ID = mlng接诊科室ID
        End If
        strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng就诊ID & ",'ZLHIS_CIS_033',1,'" & UserInfo.姓名 & "'," & lng科室ID & ",null," & lng消息ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        ReadMsgCIS033 = True
        Exit Function
    End If
    
    If "中华人民共和国传染病报告卡" = Sys.RowValue("电子病历记录", lng文件ID, "病历名称") Then
        '弹出来修改报告
        '先将卡片切换到医嘱卡片方便查找菜单
        Call LocatedCard("疾病报告")
        cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
        If tbcSub.Selected.Tag = "疾病报告" And tbcSub.Selected.Visible = True Then
            Set objControl = cbsMain.FindControl(, conMenu_Edit_Modify, True, True)
            If Not objControl Is Nothing Then
                If objControl.Enabled Then objControl.Execute
            End If
        End If
    Else
        '弹出来修改报告
        Call mclsDis.ModifyDiseaseDoc(Me, lng文件ID, mlng病人ID, mlng挂号ID, 1, mlng科室ID)
    End If
    
    strSQL = "Select 1 From 疾病申报记录 where 文件ID=[1] and 处理状态=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng文件ID, 4)
    If rsTmp.RecordCount = 0 Then
    '把消息标记为已读
        If mlng接诊科室ID = 0 Then
            lng科室ID = UserInfo.部门ID
        Else
            lng科室ID = mlng接诊科室ID
        End If
        strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng就诊ID & ",'ZLHIS_CIS_033',1,'" & UserInfo.姓名 & "'," & lng科室ID & ",null," & lng消息ID & ")"
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

Private Sub mclsDis_PatiTransfer(ByVal lng病人ID As Long, ByVal str挂号No As String)
'功能：传染病阳性界面触发事件转诊。
    Call ExecuteTransferSend
End Sub

Private Sub mobjPati_SetEdit()
    picBasisNew.SetFocus
    tbcSub.SetFocus
End Sub

Private Sub ExecuteCritical()
'功能：危急值相关处理
    Dim lng危急值ID As Long  '本次处理的危急值记录ID
    
    Call mobjKernel.ShowDealCritical(Me, mlng病人ID, 0, mstr挂号单, lng危急值ID)
    
    Call SetCriticalAdvice(lng危急值ID)
End Sub

Private Sub mobjQueue_OnInitQueueList(objQueueList As Object, objCallList As Object, blnIsCustom As Boolean)
'功能：排队叫列表的初始化操作

    Dim Column As ReportColumn
    Dim str排队列宽 As String
    Dim str呼叫列宽 As String
    Dim strReg As String
    
    On Error GoTo errH
    
    Set mobjQueueList = objQueueList
    Set mobjCallList = objCallList
 
    strReg = "公共全局\自定义排队叫号" & CStr(mlngModul)
    str排队列宽 = GetSetting("ZLSOFT", strReg, "排队列宽度配置", C_STR_QUEUEQUEUE)
    str呼叫列宽 = GetSetting("ZLSOFT", strReg, "呼叫列宽度配置", C_STR_QUEUECALL)
    If UBound(Split(str排队列宽, ",")) <> 18 Then
        str排队列宽 = C_STR_QUEUEQUEUE
    End If
    If UBound(Split(str呼叫列宽, ",")) <> 18 Then
        str呼叫列宽 = C_STR_QUEUECALL
    End If
    mlngQueueGroupType = zlDatabase.GetPara("排队分组类型", glngSys, p排队叫号虚拟模块, "0")
    mstrShowColumnInf = zlDatabase.GetPara("数据显示列", glngSys, p排队叫号虚拟模块, "号码,患者姓名,排队状态")
    mstrShowColumnInf = Replace(mstrShowColumnInf, "，", ",")
    mstrShowColumnInf = "," & mstrShowColumnInf & ","
    mstrShowCalledColumnInf = zlDatabase.GetPara("呼叫数据显示列", glngSys, p排队叫号虚拟模块, "号码,患者姓名")
    mstrShowCalledColumnInf = Replace(mstrShowCalledColumnInf, "，", ",")
    mstrShowCalledColumnInf = "," & mstrShowCalledColumnInf & ","
    mlngOrderStyle = zlDatabase.GetPara("使用数据原始顺序排序", glngSys, p排队叫号虚拟模块, "0")
    mlng回诊病人优先 = zlDatabase.GetPara("回诊病人是否优先", glngSys, p排队叫号虚拟模块, "1")
    mlngQueueGroupType = zlDatabase.GetPara("排队分组类型", glngSys, p排队叫号虚拟模块, "0")
    mlng回诊病人优先 = zlDatabase.GetPara("回诊病人是否优先", glngSys, p排队叫号虚拟模块, "1")

    '原来的流程
    With objCallList.Columns
        objCallList.AllowColumnRemove = False
        objCallList.ShowItemsInGroups = False
        objCallList.SkipGroupsFocus = True
        objCallList.MultipleSelection = False
        objCallList.AutoColumnSizing = False
        With objCallList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "将列标题拖动到此,可按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
        Set Column = .Add(mCol.队列名称, IIf(mlngQueueGroupType = 0, "", "队列"), Val(Split(str排队列宽, ",")(0)), False)
        If mlngQueueGroupType = 0 Then
            Column.Groupable = True
        Else
            Column.Visible = False
        End If

        Set Column = .Add(mCol.ID, "ID", Val(Split(str呼叫列宽, ",")(1)), False)
        Column.Visible = False

        Set Column = .Add(mCol.病人ID, "病人ID", Val(Split(str呼叫列宽, ",")(2)), False)
        Column.Visible = False

        Set Column = .Add(mCol.排队标记, "标记", Val(Split(str呼叫列宽, ",")(3)), False)
        Column.Visible = False

        Set Column = .Add(mCol.排队号码, "号码", Val(Split(str呼叫列宽, ",")(4)), True)
        Column.Visible = True

        Set Column = .Add(mCol.排队序号, "排队序号", Val(Split(str呼叫列宽, ",")(5)), False)
        Column.Visible = False

        Set Column = .Add(mCol.患者姓名, "患者姓名", Val(Split(str呼叫列宽, ",")(6)), True)
        Column.Visible = True

        Set Column = .Add(mCol.优先, "优先", Val(Split(str呼叫列宽, ",")(7)), False)
        Column.Visible = False

        Set Column = .Add(mCol.回诊序号, "回诊序号", Val(Split(str呼叫列宽, ",")(8)), True)
        Column.Visible = IIf(InStr(1, mstrShowCalledColumnInf, ",回诊序号,") > 0, True, False)

        Set Column = .Add(mCol.回诊排序号, "回诊排序号", Val(Split(str呼叫列宽, ",")(9)), False)
        Column.Visible = False

        Set Column = .Add(mCol.科室ID, "科室ID", Val(Split(str呼叫列宽, ",")(10)), False)
        Column.Visible = False

        Set Column = .Add(mCol.诊室, IIf(mlngQueueGroupType = 2, "", "诊室"), Val(Split(str呼叫列宽, ",")(11)), True)
        If mlngQueueGroupType = 2 Then
            Column.Groupable = True
            Column.Visible = False
        Else
            Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",诊室,") > 0, True, False)
        End If

        Set Column = .Add(mCol.医生姓名, IIf(mlngQueueGroupType = 1, "", "医生姓名"), Val(Split(str呼叫列宽, ",")(12)), True)
        If mlngQueueGroupType = 1 Then
            Column.Groupable = True
            Column.Visible = False
        Else
            Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",医生姓名,") > 0, True, False)
        End If

        Set Column = .Add(mCol.排队状态, "排队状态", Val(Split(str呼叫列宽, ",")(13)), False)
        Column.Visible = False

        Set Column = .Add(mCol.排队时间, "排队时间", Val(Split(str呼叫列宽, ",")(14)), False)
        Column.Visible = False

        Set Column = .Add(mCol.呼叫医生, "呼叫人", Val(Split(str呼叫列宽, ",")(15)), True)
        Column.Visible = IIf(InStr(1, mstrShowCalledColumnInf, ",呼叫医生,") > 0, True, False)

        Set Column = .Add(mCol.业务类型, "业务类型", Val(Split(str呼叫列宽, ",")(16)), False)
        Column.Visible = False

        Set Column = .Add(mCol.业务ID, "业务ID", Val(Split(str呼叫列宽, ",")(17)), False)
        Column.Visible = False

        Set Column = .Add(mCol.呼叫时间, "呼叫时间", Val(Split(str呼叫列宽, ",")(18)), True)
        Column.Visible = IIf(InStr(1, mstrShowCalledColumnInf, ",呼叫时间,") > 0, True, False)

        Set Column = .Add(mCol.排序名称, "排序名称", 0, False)
        Column.Visible = False

        Set Column = .Add(mCol.ORD, "ORD", 0, False)
        Column.Visible = False

    End With

    With objCallList
        Set .Icons = zlCommFun.GetPubIcons
        .GroupsOrder.DeleteAll
        If mlngQueueGroupType = 0 Then
            .GroupsOrder.Add .Columns(mCol.排序名称)
        ElseIf mlngQueueGroupType = 1 Then
            .GroupsOrder.Add .Columns(mCol.医生姓名)
        Else
            .GroupsOrder.Add .Columns(mCol.诊室)
        End If
        .GroupsOrder(0).SortAscending = True '分组之后,如果分组列不显示,分组列的排序是不变的
        .SortOrder.DeleteAll
        If mlngOrderStyle = 1 Then
            .SortOrder.Add .Columns(mCol.ORD)
            .SortOrder(0).SortAscending = True
        Else

            .SortOrder.Add .Columns(mCol.排队状态)
            .SortOrder(0).SortAscending = False

            .SortOrder.Add .Columns(mCol.排队序号)
            .SortOrder(1).SortAscending = True

            .SortOrder.Add .Columns(mCol.呼叫时间)
            .SortOrder(2).SortAscending = True

            .SortOrder.Add .Columns(mCol.排队号码)
            .SortOrder(3).SortAscending = True
        End If
    End With

    '初始化排队队列显示字段
    Call objQueueList.Columns.DeleteAll
    With objQueueList.Columns
        objQueueList.AllowColumnRemove = False
        objQueueList.ShowItemsInGroups = False
        objQueueList.SkipGroupsFocus = True
        objQueueList.MultipleSelection = False
        objQueueList.AutoColumnSizing = False

        With objQueueList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "将列标题拖动到此,可按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With

        Set Column = .Add(mCol.队列名称, IIf(mlngQueueGroupType = 0, "", "队列"), Val(Split(str排队列宽, ",")(0)), False)

        If mlngQueueGroupType = 0 Then
            Column.Groupable = True
        Else
            Column.Visible = False
        End If
        
        Set Column = .Add(mCol.ID, "ID", Val(Split(str排队列宽, ",")(1)), False)
        Column.Visible = False

        Set Column = .Add(mCol.病人ID, "病人ID", Val(Split(str排队列宽, ",")(2)), False)
        Column.Visible = False

        Set Column = .Add(mCol.排队标记, "标记", Val(Split(str排队列宽, ",")(3)), False)
        Column.Visible = False

        Set Column = .Add(mCol.排队号码, "号码", Val(Split(str排队列宽, ",")(4)), True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",号码,") > 0, True, False)

        Set Column = .Add(mCol.排队序号, "排队序号", Val(Split(str排队列宽, ",")(5)), False)
        Column.Visible = False

        Set Column = .Add(mCol.患者姓名, "患者姓名", Val(Split(str排队列宽, ",")(6)), True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",患者姓名,") > 0, True, False)

        Set Column = .Add(mCol.优先, "优先", Val(Split(str排队列宽, ",")(7)), True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, "优先") > 0, True, False)

        Set Column = .Add(mCol.回诊序号, "回诊序号", Val(Split(str排队列宽, ",")(8)), True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",回诊序号,") > 0, True, False)

        Set Column = .Add(mCol.回诊排序号, "回诊排序号", Val(Split(str排队列宽, ",")(9)), True)
        Column.Visible = False

        Set Column = .Add(mCol.科室ID, "科室ID", Val(Split(str排队列宽, ",")(10)), False)
        Column.Visible = False

        Set Column = .Add(mCol.诊室, IIf(mlngQueueGroupType = 2, "", "诊室"), Val(Split(str排队列宽, ",")(11)), True)
        If mlngQueueGroupType = 2 Then
            Column.Groupable = True
            Column.Visible = False
        Else
            Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",诊室,") > 0, True, False)
        End If

        Set Column = .Add(mCol.医生姓名, IIf(mlngQueueGroupType = 1, "", "医生姓名"), Val(Split(str排队列宽, ",")(12)), True)
        If mlngQueueGroupType = 1 Then
            Column.Groupable = True
            Column.Visible = False
        Else
            Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",医生姓名,") > 0, True, False)
        End If
        Set Column = .Add(mCol.排队状态, "排队状态", Val(Split(str排队列宽, ",")(13)), True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",排队状态,") > 0, True, False)

        Set Column = .Add(mCol.排队时间, "排队时间", Val(Split(str排队列宽, ",")(14)), True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",排队时间,") > 0, True, False)

        Set Column = .Add(mCol.呼叫医生, "呼叫人", Val(Split(str排队列宽, ",")(15)), False)
        Column.Visible = False

        Set Column = .Add(mCol.业务类型, "业务类型", Val(Split(str排队列宽, ",")(16)), False)
        Column.Visible = False

        Set Column = .Add(mCol.业务ID, "业务ID", Val(Split(str排队列宽, ",")(17)), False)
        Column.Visible = False

        Set Column = .Add(mCol.呼叫时间, "呼叫时间", Val(Split(str排队列宽, ",")(18)), False)
        Column.Visible = False

        Set Column = .Add(mCol.排序名称, "排序名称", 0, False)
        Column.Visible = False

        Set Column = .Add(mCol.ORD, "ORD", 0, False)
        Column.Visible = False
    End With

    With objQueueList
        Set .Icons = zlCommFun.GetPubIcons

        .GroupsOrder.DeleteAll

        If mlngQueueGroupType = 0 Then
            .GroupsOrder.Add .Columns(mCol.排序名称)
        ElseIf mlngQueueGroupType = 1 Then
            .GroupsOrder.Add .Columns(mCol.医生姓名)
        Else
            .GroupsOrder.Add .Columns(mCol.诊室)
        End If

        .GroupsOrder(0).SortAscending = True '分组之后,如果分组列不显示,分组列的排序是不变的

        '队列名称 = 0: Id:排队标记: 排队号码: 优先: 患者姓名: 科室ID:  诊室: 医生姓名:排队状态 : 排队时间: 业务ID
        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.DeleteAll
        If mlngOrderStyle = 1 Then
            .SortOrder.Add .Columns(mCol.ORD)
            .SortOrder(0).SortAscending = True
        Else
            .SortOrder.Add .Columns(mCol.排队状态)
            .SortOrder(0).SortAscending = True

            .SortOrder.Add .Columns(mCol.排队序号)
            .SortOrder(1).SortAscending = True

            .SortOrder.Add .Columns(mCol.优先)
            .SortOrder(2).SortAscending = False

            .SortOrder.Add .Columns(mCol.回诊排序号)
            .SortOrder(3).SortAscending = True

            .SortOrder.Add .Columns(mCol.排队时间)
            .SortOrder(4).SortAscending = True

            .SortOrder.Add .Columns(mCol.排队号码)
            .SortOrder(5).SortAscending = True
        End If
    End With
    blnIsCustom = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mobjQueue_OnRefresh(str队列名称() As String, ByVal strCur队列名称 As String, ByVal strCur业务ID As String, ByVal strMustCols As String, ByVal str诊室 As String, ByVal str医生 As String, ByVal strExcludeData As String, ByVal intViewDataType As Integer, ByVal str执行状态 As String, blnIsCustom As Boolean)
'功能：排队叫号刷新
 
    Dim strFilter As String
    Dim j As Long, i As Long
    Dim strValues(0 To 10) As String, strValue As String, strUninTable As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsLocal As ADODB.Recordset
    Dim rptCalling As ReportRecord
    Dim rptRecord As ReportRecord
 

    On Error GoTo errH
    
    '非自定义流程，保持113794前的处理方式
    strFilter = "": strValue = "": j = 0: strUninTable = ""
    If SafeArrayGetDim(str队列名称) > 0 Then
        For i = 1 To UBound(str队列名称)
            If Trim(str队列名称(i)) <> "" Then
                If j > 10 Then
                    strFilter = strFilter & " Or A.队列名称 ='" & str队列名称(i) & "'"
                Else
                    If zlCommFun.ActualLen(strValue) > 2000 Then
                         strValues(j) = Mid(strValue, 2)
                         strUninTable = strUninTable & " Union ALL  Select  Column_Value as 队列名称 From Table(Cast(f_Str2list([" & j + 4 & "]) As zlTools.t_Strlist))  " & vbCrLf
                         strValue = "": j = j + 1
                    End If
                    strValue = strValue & "," & str队列名称(i)
                End If
            End If
        Next i
        If strValue <> "" Then
            strValues(j) = Mid(strValue, 2)
            strUninTable = strUninTable & " Union ALL  Select  Column_Value as 队列名称 From Table(Cast(f_Str2list([" & j + 4 & "]) As zlTools.t_Strlist))  " & vbCrLf
        End If
    End If
    
    If strUninTable <> "" Then
        strUninTable = Mid(strUninTable, 11)
    Else
'        labError.Caption = "没有可显示的叫号队列信息，请检查相关排队科室设置"
        Exit Sub
    End If
    
    If strFilter <> "" Then strFilter = "( " & Mid(strFilter, 4) & ")"
     
    '为了支持复制，需要将number类型的字段进行转换，可以使用to_Number方式
    strSQL = "" & _
    "   Select /*+ Rule*/  to_Number(A.ID) as ID, to_Number(a.病人id) as 病人id, A.队列名称, A.排队序号, to_Number(A.业务类型) as 业务类型, to_Number(A.业务ID) as 业务ID," & _
    "           to_Number(A.科室ID) as 科室ID, x.名称 as 部门名称, A.排队号码 , A.排队标记,A.患者姓名||decode(e.预约,1,'(预)',null) as 患者姓名,A.诊室,A.医生姓名," & _
    "            (select j.姓名 from 人员表 J,上机人员表 K where J.ID=K.人员ID and K.用户名=A.呼叫医生) as 呼叫医生, " & _
    "           to_Number(A.优先) as 优先, to_Number(A.回诊序号) as 回诊序号, To_Char(A.排队时间, 'yyyy-mm-dd hh24:mi:ss') as 排队时间, To_Char(A.呼叫时间, 'yyyy-mm-dd hh24:mi:ss') as 呼叫时间,to_Number(A.排队状态) as 排队状态, " & _
                IIf(mlng回诊病人优先 = 1, "to_number(nvl(A.回诊序号, 9999999999)) as 回诊排序号", "0 as 回诊排序号") & _
    "   From 排队叫号队列 a, 部门表 x " & IIf(strUninTable <> "", ", (" & strUninTable & ") b ", "") & _
                IIf(intViewDataType = 1, ", Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D ", "") & _
                IIf(intViewDataType = 2, ", Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D ", "") & _
                IIf(intViewDataType = 3, ", Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D", "") & " , 病人挂号记录 E" & _
    "   Where To_Number(a.业务id) = e.Id and  (nvl(a.是否分时点, 0)=0 and A.排队时间 <= trunc(sysdate + 1) - 1/24/60/60 or nvl(a.是否分时点, 0)=1 and sysdate>a.排队时间) " & IIf(strUninTable <> "", " and a.队列名称=b.队列名称 ", "") & " and instr([3],A.排队状态)=0  and x.ID=a.科室ID  " & _
                IIf(intViewDataType = 1, " and  ((a.诊室=C.Column_Value and a.医生姓名 is null) or a.医生姓名=D.Column_Value or (a.诊室 is null and a.医生姓名 is null))", "") & _
                IIf(intViewDataType = 2, " and (a.诊室=C.Column_Value and (a.医生姓名 is Null or a.医生姓名=D.Column_Value)) ", "") & _
                IIf(intViewDataType = 3, " and a.医生姓名=D.Column_Value", "") & _
    "           " & strFilter & _
    "   Order by  排队状态 desc, 排队序号,优先 Desc, 回诊排序号, 排队时间, 排队号码 "
    

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询队列", str诊室, str医生, str执行状态, strValues(0), strValues(1), strValues(2), strValues(3), strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10))
    Set rsLocal = zlDatabase.CopyNewRec(rsTemp)
    
    '删除需要排除的数据,并获取实际排队号码值得最长度
    If rsLocal.RecordCount > 0 Then rsLocal.MoveFirst
    While Not rsLocal.EOF
        If InStr(1, strExcludeData, rsLocal!业务类型 & ":" & rsLocal!业务ID) > 0 Then
            rsLocal.Delete
        End If
        If LenB(StrConv(Trim(NVL(rsLocal("排队号码"))), vbFromUnicode)) > mlngMaxLen Then
            mlngMaxLen = LenB(StrConv(Trim(NVL(rsLocal("排队号码"))), vbFromUnicode))
        End If
        rsLocal.MoveNext
    Wend

    rsLocal.Sort = "队列名称, 排队状态 desc, 排队序号, 优先 Desc, 回诊排序号, 排队时间, 排队号码"
    If rsLocal.RecordCount > 0 Then rsLocal.MoveFirst
    Call mobjQueueList.Records.DeleteAll
    Call mobjCallList.Records.DeleteAll
    While Not rsLocal.EOF
        If rsLocal("排队状态") = 7 Or rsLocal("排队状态") = 1 Then
            Set rptCalling = mobjCallList.Records.Add
            For j = 0 To mobjCallList.Columns.Count - 1
                rptCalling.AddItem ""
            Next
            Call SetReportRecordItem(rptCalling, rsLocal)
        Else
            Set rptRecord = mobjQueueList.Records.Add
            For j = 0 To mobjQueueList.Columns.Count - 1
                rptRecord.AddItem ""
            Next
            Call SetReportRecordItem(rptRecord, rsLocal)
        End If
        rsLocal.MoveNext
    Wend
    mobjQueueList.Populate
    mobjCallList.Populate
    
    blnIsCustom = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetReportRecordItem(rriItem As ReportRecord, rsData As ADODB.Recordset)
    Dim i As Integer
    
    On Error GoTo errHandle
    rriItem(mCol.ID).Value = rsData("id")
    rriItem(mCol.病人ID).Value = NVL(rsData("病人ID"))
    
    rriItem(mCol.队列名称).Caption = rsData("部门名称") & ":" & IIf(InStr(1, NVL(rsData("队列名称")), ":") <= 0, "", Mid(NVL(rsData("队列名称")), InStr(1, NVL(rsData("队列名称")), ":") + 1))
    rriItem(mCol.队列名称).Value = NVL(rsData("队列名称"))

    rriItem(mCol.患者姓名).Value = NVL(rsData("患者姓名"))
    rriItem(mCol.科室ID).Value = NVL(rsData("科室ID"))
    rriItem(mCol.排队标记).Value = NVL(rsData("排队标记"))
    rriItem(mCol.排队序号).Value = zlStr.Lpad(NVL(rsData("排队序号")), 20)
    rriItem(mCol.排队号码).Value = zlStr.Lpad(NVL(rsData("排队号码")), mlngMaxLen)
    rriItem(mCol.排队时间).Value = NVL(rsData("排队时间"))
    rriItem(mCol.呼叫时间).Value = NVL(rsData("呼叫时间"))
    rriItem(mCol.回诊序号).Value = NVL(rsData("回诊序号"))
    rriItem(mCol.回诊排序号).Value = NVL(rsData("回诊排序号"))
    rriItem(mCol.呼叫医生).Value = NVL(rsData("呼叫医生"))
    rriItem(mCol.排序名称).Value = DeptNametransform(NVL(rsData("部门名称")))
    rriItem(mCol.排序名称).Caption = (NVL(rsData("部门名称")))
    rriItem(mCol.ORD).Value = Format(rsData.AbsolutePosition, "00000000")
    
    If NVL(rsData("回诊序号")) = "" Then
        rriItem(mCol.患者姓名).Icon = 807
    Else
        rriItem(mCol.患者姓名).Icon = 3504
    End If
    
    
    If NVL(rsData("排队状态")) = 1 Then
        rriItem(mCol.排队状态).Value = "呼叫中"
        For i = 0 To mobjQueueList.Columns.Count - 1
            rriItem(i).BackColor = &HC0C0FF
        Next
    ElseIf NVL(rsData("排队状态")) = 0 Then
        rriItem(mCol.排队状态).Value = "排队中"
        For i = 0 To mobjQueueList.Columns.Count - 1
            rriItem(i).BackColor = ColorConstants.vbWhite
        Next
    ElseIf NVL(rsData("排队状态")) = 3 Then
        rriItem(mCol.排队状态).Value = "暂停"
        For i = 0 To mobjQueueList.Columns.Count - 1
            rriItem(i).BackColor = ColorConstants.vbYellow
        Next
    ElseIf NVL(rsData("排队状态")) = 4 Then
        rriItem(mCol.排队状态).Value = "完成"
        For i = 0 To mobjQueueList.Columns.Count - 1
            rriItem(i).BackColor = ColorConstants.vbGreen
        Next
    ElseIf NVL(rsData("排队状态")) = 7 Then
        rriItem(mCol.排队状态).Value = "已呼叫"
    Else
        rriItem(mCol.排队状态).Value = "已弃号"
        For i = 0 To mobjQueueList.Columns.Count - 1
            rriItem(i).BackColor = &HC0C0C0
        Next
    End If
    
    If mlngQueueGroupType = 1 Then
        rriItem(mCol.医生姓名).Value = NVL(rsData("部门名称")) & ":" & NVL(rsData("医生姓名"))
    Else
        rriItem(mCol.医生姓名).Value = NVL(rsData("医生姓名"))
    End If

    rriItem(mCol.业务类型).Value = NVL(rsData("业务类型"))
    rriItem(mCol.业务ID).Value = NVL(rsData("业务ID"))

    rriItem(mCol.优先).Value = IIf(NVL(rsData("优先")) = 1, "优先", "")
    
    If mlngQueueGroupType = 2 Then
        rriItem(mCol.诊室).Value = NVL(rsData("部门名称")) & ":" & NVL(rsData("诊室"))
    Else
        rriItem(mCol.诊室).Value = NVL(rsData("诊室"))
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function DeptNametransform(ByVal strOldName) As String
'功能：排队叫号方法，部门名称转化，目前只支持 一到十的处理 将小写数字转化为 abc 这种形式便于排序
    Dim strWord As String '单个字符
    Dim intCount As Integer
    Dim i As Integer
    
    On Error GoTo errH
    
    DeptNametransform = strOldName
    intCount = 0
    For i = 1 To Len(strOldName)
        strWord = Mid(strOldName, i, 1)
        If strWord = "一" Or strWord = "二" Or strWord = "三" Or strWord = "四" Or strWord = "五" Or strWord = "六" Or _
           strWord = "七" Or strWord = "八" Or strWord = "九" Or strWord = "十" Then
            intCount = intCount + 1
        End If
    Next
    If intCount = 1 Then
        DeptNametransform = Replace(strOldName, "一", "a")
        DeptNametransform = Replace(DeptNametransform, "二", "b")
        DeptNametransform = Replace(DeptNametransform, "三", "c")
        DeptNametransform = Replace(DeptNametransform, "四", "d")
        DeptNametransform = Replace(DeptNametransform, "五", "e")
        DeptNametransform = Replace(DeptNametransform, "六", "f")
        DeptNametransform = Replace(DeptNametransform, "七", "g")
        DeptNametransform = Replace(DeptNametransform, "八", "h")
        DeptNametransform = Replace(DeptNametransform, "九", "i")
        DeptNametransform = Replace(DeptNametransform, "十", "j")
    End If

    Exit Function
errH:
    DeptNametransform = strOldName
End Function

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

Private Sub LoadPatients回诊()
'功能：加载候诊回诊列表
    Dim strSQL As String
    Dim strTime As String
    Dim i As Long, j As Long, k As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim rsPati As ADODB.Recordset
    Dim strTmp As String
    Dim lngColor As Long
    Dim rs传染病状态 As ADODB.Recordset
    Dim blnDo传染病状态 As Boolean
    
    On Error GoTo errH
    Screen.MousePointer = 11
    mblnUnRefresh = True
     
    strSQL = _
        " Select B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,B.复诊,B.急诊,B.社区,nvl(g.号类,E.号类) as 号类,D.名称 as 病人科室," & _
        " B.执行时间 as 时间,A.就诊卡号,A.身份证号,A.IC卡号,A.险类,B.发生时间,B.执行部门ID,B.执行人," & _
        " B.转诊状态,C.名称 as 转诊科室,B.转诊诊室,B.转诊医生,B.执行状态,B.记录标志,A.病人类型" & _
        " From 病人信息 A,病人挂号记录 B,部门表 C,部门表 D,挂号安排 E, 临床出诊记录 f,临床出诊号源 g" & _
        " Where B.病人ID=A.病人ID And B.转诊科室ID=C.ID(+) and B.号别=E.号码(+) and B.执行部门ID=d.id and b.出诊记录id=f.id(+) and f.号源id=g.id(+)" & _
        " And B.执行状态=2 And B.执行人||''=[1] And B.记录性质=1 And B.记录状态=1 and nvl(B.记录标志,0) in (2,3)" & _
        " Order By B.NO"
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr接诊医生)
    
    strSQL = "select m.病人id,m.id,m.no,max(m.记录) as 记录,max(m.填写) as 填写,max(m.状态) as 状态 from" & vbNewLine & _
        "(select a.病人id,a.id, a.no,1 as 记录,0 as 填写,0 as 状态 from 病人挂号记录 a,疾病阳性记录 b" & vbNewLine & _
        "where a.no=b.挂号单 and a.执行状态=2 And a.执行人||''=[1] And a.记录性质=1 And a.记录状态=1 and nvl(a.记录标志,0) in (2,3)" & vbNewLine & _
        "union all" & vbNewLine & _
        "Select  a.病人id,a.id, a.no,0 as 记录,1 as 填写,0 as 状态" & vbNewLine & _
        "From 病人挂号记录 A, 电子病历记录 C, 病历文件列表 D" & vbNewLine & _
        "Where c.文件id = d.Id And d.种类 = 5  and c.病历名称 like '%传染病%' And a.病人id = c.病人id And a.id = c.主页id and a.执行状态=2 And a.执行人||''=[1] And a.记录性质=1 And a.记录状态=1 and nvl(a.记录标志,0) in (2,3)" & vbNewLine & _
        "union all" & vbNewLine & _
        "Select  a.病人id,a.id, a.no,0 as 记录,1 as 填写,e.处理状态 as 状态" & vbNewLine & _
        "From 病人挂号记录 A,电子病历记录 C,病历文件列表 D,疾病申报记录 E" & vbNewLine & _
        "Where a.病人id = c.病人id And a.id = c.主页id and c.id=e.文件id and d.种类=5 and c.病历名称 like '%传染病%' and e.文件id =d.id and a.执行状态=2 And a.执行人||''=[1] And a.记录性质=1 And a.记录状态=1 and nvl(a.记录标志,0) in (2,3)) M" & vbNewLine & _
        "group by m.病人id,m.id,m.no"
    Set rs传染病状态 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr接诊医生)
    If rs传染病状态.RecordCount > 0 Then blnDo传染病状态 = True
    
    rptPati(PATI_RPT回诊).Records.DeleteAll
    For i = 1 To rsPati.RecordCount
        Set objRecord = rptPati(PATI_RPT回诊).Records.Add()
        For j = 0 To rptPati(PATI_RPT回诊).Columns.Count - 1
            objRecord.AddItem ""
        Next
        With objRecord
            .Item(COL_JZ_标识).Value = "回"
            .Item(COL_JZ_门诊号).Value = rsPati!门诊号 & ""
            .Item(COL_JZ_姓名).Value = rsPati!姓名 & ""
            .Item(COL_JZ_就诊时间).Value = Format(rsPati!时间, "yyyy-MM-dd HH:mm")
            .Item(COL_JZ_性别).Value = rsPati!性别 & ""
            .Item(COL_JZ_年龄).Value = rsPati!年龄 & ""
            .Item(COL_JZ_急).Value = IIf(Val(rsPati!急诊 & "") <> 0, "急", "")
            .Item(COL_JZ_复).Value = IIf(Val(rsPati!复诊 & "") <> 0, "复", "")
            .Item(COL_JZ_NO).Value = rsPati!NO & ""
            .Item(COL_JZ_社区).Value = IIf(Val(rsPati!社区 & "") <> 0, "√", "")
            .Item(COL_JZ_就诊卡号).Value = rsPati!就诊卡号 & ""
            .Item(COL_JZ_病人类型).Value = rsPati!病人类型 & ""
            .Item(COL_JZ_病人ID).Value = rsPati!病人ID & ""
            .Item(COL_JZ_发生时间).Value = CStr(Format(rsPati!发生时间, "yyyy-MM-dd HH:mm:ss"))
            .Item(COL_JZ_执行部门ID).Value = rsPati!执行部门ID & ""
            .Item(COL_JZ_执行人).Value = rsPati!执行人 & ""
            .Item(COL_JZ_身份证号).Value = rsPati!身份证号 & ""
            .Item(COL_JZ_IC卡号).Value = rsPati!IC卡号 & ""
            .Item(COL_JZ_记录标志).Value = rsPati!记录标志 & ""
            .Item(COL_JZ_号类).Value = rsPati!号类 & ""
            .Item(COL_JZ_病人科室).Value = rsPati!病人科室 & ""
            
            '保险病人用红色显示
            If Not IsNull(rsPati!险类) And rsPati!病人类型 & "" = "" Then
                .Item(COL_JZ_门诊号).ForeColor = &HC0&
                .Item(COL_JZ_病人类型).ForeColor = &HC0&
            Else
                '病人颜色
                lngColor = zlDatabase.GetPatiColor(NVL(rsPati!病人类型))
                .Item(COL_JZ_门诊号).ForeColor = lngColor
                .Item(COL_JZ_病人类型).ForeColor = lngColor
            End If
            
            '急诊标志红色突出显示
            If NVL(rsPati!急诊, 0) <> 0 Then
                .Item(COL_JZ_急).ForeColor = vbRed
            End If
            
            '添加传染病状态
            strSQL = ""
            If blnDo传染病状态 Then
                rs传染病状态.Filter = "no='" & rsPati!NO & "'"
                If Not rs传染病状态.EOF Then strSQL = Get传染病状态(Val(rs传染病状态!记录 & ""), Val(rs传染病状态!填写 & ""), Val(rs传染病状态!状态 & ""))
            End If
            .Item(COL_JZ_传染病).Value = strSQL
        End With
        rsPati.MoveNext
    Next
    rptPati(PATI_RPT回诊).Populate
    i = rptPati(PATI_RPT回诊).Records.Count
    tbcInTreat.Item(t回诊).Caption = "回诊" & IIf(i = 0, "", ":" & i & "人")
    mblnUnRefresh = False
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ExecuteTabChange(ByVal strTab As String)
'功能：病历/医嘱 页签快速切换，调用新增病历/医嘱
    Dim lngidx As Long
    Dim j As Long
    Dim objControl As CommandBarControl
    
    lngidx = -1
    For j = 0 To tbcSub.ItemCount - 1
        If tbcSub.Item(j).Tag = strTab Then
            lngidx = j
            Exit For
        End If
    Next
    
    If lngidx <> -1 Then
        If strTab = "医嘱" Then
            If tbcRegist.Selected.Tag = "诊疗一览" Then tbcRegist.Item(mbyt本次就诊).Selected = True
            If tbcSub.Selected.Tag <> "医嘱" Then tbcSub.Item(lngidx).Selected = True
            cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
            Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
            If Not objControl Is Nothing Then
                If objControl.Enabled Then objControl.Execute
            End If
        ElseIf strTab = "病历" Then
            If tbcRegist.Selected.Tag = "诊疗一览" Then tbcRegist.Item(mbyt本次就诊).Selected = True
            If tbcSub.Selected.Tag <> "病历" Then tbcSub.Item(lngidx).Selected = True
            cbsMain.RecalcLayout: Me.Refresh '可能未来得及刷新
            mblnUnRefresh = True
            Call mclsEPRs.zlOpenDefaultEPR(mstr挂号单)
        End If
    End If
End Sub

Private Sub ReadMsgAuto()
'功能：危急值消息处理自动弹出
    Dim i As Long
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    Dim strNO As String
    Dim str业务 As String
    Dim lng消息ID As Long
    Dim lng挂号id As String
    Dim str挂号单 As String
    Dim blnRs As Boolean
    
    On Error GoTo errH

    For i = i To rptNotify.Rows.Count - 1
        With rptNotify.Rows(i)
            If Not .GroupRow Then
                strNO = .Record(C_消息).Value
                If strNO = "ZLHIS_LIS_003" Or strNO = "ZLHIS_PACS_005" Then
                    lng病人ID = Val(.Record(C_病人ID).Value)
                    lng挂号id = Val(.Record(C_挂号Id).Value)
                    str挂号单 = .Record(C_No).Value
                    str业务 = .Record(C_业务).Value
                    
                    lng消息ID = Val(.Record(C_Id).Value)
                    blnRs = ReadMsg(lng病人ID, lng挂号id, strNO, str业务, lng消息ID, str挂号单)
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
