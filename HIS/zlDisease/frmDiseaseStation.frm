VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#7.0#0"; "zlIDKind.ocx"
Begin VB.Form frmDiseaseStation 
   Caption         =   "传染病管理工作站"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18240
   Icon            =   "frmDiseaseStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   18240
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   600
      ScaleHeight     =   6735
      ScaleWidth      =   6735
      TabIndex        =   38
      Top             =   1200
      Width           =   6735
      Begin VB.PictureBox picReportList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5415
         Left            =   360
         ScaleHeight     =   5415
         ScaleWidth      =   6495
         TabIndex        =   40
         Top             =   720
         Width           =   6495
         Begin XtremeReportControl.ReportControl rptList 
            Height          =   450
            Left            =   120
            TabIndex        =   41
            Top             =   4560
            Width           =   1140
            _Version        =   589884
            _ExtentX        =   2011
            _ExtentY        =   794
            _StockProps     =   0
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
         Begin VB.PictureBox picState 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   120
            ScaleHeight     =   315
            ScaleWidth      =   4095
            TabIndex        =   68
            Top             =   120
            Width           =   4095
            Begin VB.CheckBox chkAduitState 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "待审核"
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   0
               Left            =   720
               TabIndex        =   71
               Top             =   0
               Width           =   920
            End
            Begin VB.CheckBox chkAduitState 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "待返修"
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   1
               Left            =   1680
               TabIndex        =   70
               Top             =   0
               Width           =   920
            End
            Begin VB.CheckBox chkAduitState 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "返修待审核"
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   2
               Left            =   2640
               TabIndex        =   69
               Top             =   0
               Width           =   1275
            End
            Begin VB.Label lblAuditState 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "状态(S):"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   72
               Top             =   25
               Width           =   735
            End
         End
         Begin VB.PictureBox picState 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   120
            ScaleHeight     =   315
            ScaleWidth      =   4095
            TabIndex        =   64
            Top             =   720
            Width           =   4095
            Begin VB.CheckBox chkSendState 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "待上报"
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   0
               Left            =   720
               TabIndex        =   66
               Top             =   0
               Width           =   920
            End
            Begin VB.CheckBox chkSendState 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "已上报"
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   1
               Left            =   1680
               TabIndex        =   65
               Top             =   0
               Width           =   920
            End
            Begin VB.Label lblSendState 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "状态(S):"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   67
               Top             =   25
               Width           =   735
            End
         End
         Begin VB.PictureBox picState 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1095
            Index           =   4
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   4935
            TabIndex        =   55
            Top             =   2400
            Visible         =   0   'False
            Width           =   4935
            Begin VB.TextBox txtDiagnose 
               Height          =   315
               Left            =   600
               MaxLength       =   100
               TabIndex        =   59
               Top             =   700
               Width           =   2535
            End
            Begin VB.TextBox txtName 
               Height          =   300
               Left            =   600
               MaxLength       =   20
               TabIndex        =   58
               Top             =   50
               Width           =   2535
            End
            Begin VB.CommandButton cmdDuplicateCheck 
               Caption         =   "查找"
               Height          =   360
               Left            =   3480
               TabIndex        =   57
               Top             =   650
               Width           =   900
            End
            Begin VB.TextBox txtProfession 
               Height          =   300
               Left            =   600
               MaxLength       =   20
               TabIndex        =   56
               Top             =   370
               Width           =   2535
            End
            Begin VB.Label lblDiagnose 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "诊断"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   63
               Top             =   720
               Width           =   495
            End
            Begin VB.Label lblProfession 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "职业"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   62
               Top             =   400
               Width           =   495
            End
            Begin VB.Label lblName 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "*姓名"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   40
               TabIndex        =   61
               Top             =   50
               Width           =   480
            End
            Begin VB.Label lblNotice 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "*姓名必须填写！"
               ForeColor       =   &H000040C0&
               Height          =   255
               Left            =   3120
               TabIndex        =   60
               Top             =   50
               Visible         =   0   'False
               Width           =   1575
            End
         End
         Begin VB.PictureBox picState 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   400
            Index           =   3
            Left            =   120
            ScaleHeight     =   405
            ScaleWidth      =   5415
            TabIndex        =   52
            Top             =   1920
            Width           =   5415
            Begin VB.ComboBox cboSelectTime 
               Height          =   300
               Index           =   1
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   53
               Top             =   50
               Width           =   1695
            End
            Begin VB.Label lblDate 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "报告完成时间"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   54
               Top             =   75
               Width           =   1215
            End
         End
         Begin VB.PictureBox picState 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   400
            Index           =   5
            Left            =   0
            ScaleHeight     =   405
            ScaleWidth      =   5415
            TabIndex        =   49
            Top             =   3600
            Width           =   5415
            Begin VB.ComboBox cboSelectTime 
               Height          =   300
               Index           =   2
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   50
               Top             =   50
               Width           =   1695
            End
            Begin VB.Label lblReportDate 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "报告完成时间"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   51
               Top             =   75
               Width           =   1215
            End
         End
         Begin VB.PictureBox picState 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   715
            Index           =   0
            Left            =   0
            ScaleHeight     =   720
            ScaleWidth      =   5055
            TabIndex        =   42
            Top             =   1080
            Width           =   5055
            Begin VB.CheckBox chkDisState 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "非传染病记录"
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   1
               Left            =   3480
               TabIndex        =   46
               Top             =   400
               Width           =   1380
            End
            Begin VB.CheckBox chkDisState 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "待填写报告卡"
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   0
               Left            =   2100
               TabIndex        =   45
               Top             =   400
               Width           =   1375
            End
            Begin VB.ComboBox cboSelectTime 
               Height          =   300
               Index           =   0
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   44
               Top             =   50
               Width           =   1695
            End
            Begin VB.CheckBox chkDisState 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "待处理反馈单"
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   2
               Left            =   720
               TabIndex        =   43
               Top             =   400
               Width           =   1375
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "状态(S):"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   48
               Top             =   423
               Width           =   735
            End
            Begin VB.Label lblFinishDate 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "反馈单登记时间"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   47
               Top             =   75
               Width           =   1575
            End
         End
      End
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   255
         Left            =   720
         TabIndex        =   39
         Top             =   90
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmDiseaseStation.frx":6852
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindAppearance=   0
         ShowSortName    =   -1  'True
         DefaultCardType =   "就诊卡"
         IDKindWidth     =   555
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowAutoCommCard=   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
      End
      Begin XtremeSuiteControls.TabControl tbcReportList 
         Height          =   795
         Left            =   120
         TabIndex        =   73
         Top             =   5760
         Width           =   1935
         _Version        =   589884
         _ExtentX        =   3413
         _ExtentY        =   1402
         _StockProps     =   64
      End
      Begin VB.Label lblFind 
         BackColor       =   &H80000005&
         Caption         =   "查找:"
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   100
         Width           =   615
      End
   End
   Begin VB.PictureBox picBasisNew 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   6735
      Left            =   7800
      ScaleHeight     =   6735
      ScaleWidth      =   18345
      TabIndex        =   11
      Top             =   1200
      Width           =   18345
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   800
         Index           =   0
         Left            =   1680
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   12000
         Begin VB.TextBox txtInfo 
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
            Height          =   255
            Index           =   4
            Left            =   7920
            TabIndex        =   32
            Top             =   120
            Width           =   2295
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   3
            Left            =   4560
            TabIndex        =   31
            Top             =   120
            Width           =   1815
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   2
            Left            =   2520
            TabIndex        =   30
            Text            =   "27岁"
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   29
            Text            =   "男"
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Text            =   "测试"
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox txtInfo 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   5
            Left            =   960
            TabIndex        =   27
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox txtInfo 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   6
            Left            =   4620
            TabIndex        =   26
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtInfo 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   7
            Left            =   7920
            TabIndex        =   25
            Top             =   480
            Width           =   2955
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "科室:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   7320
            TabIndex        =   37
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "标识号:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   3720
            TabIndex        =   36
            Top             =   120
            Width           =   855
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "职    业:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   35
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "电    话:"
            Height          =   255
            Index           =   6
            Left            =   3720
            TabIndex        =   34
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   " 家庭地址:"
            Height          =   255
            Index           =   7
            Left            =   6960
            TabIndex        =   33
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.TextBox txtState 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   23
         Text            =   "等待。。。。。。"
         Top             =   840
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "750"
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   1
         Left            =   1680
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   12000
         Begin VB.TextBox txtInfo 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   11
            Left            =   7920
            TabIndex        =   17
            Top             =   480
            Width           =   2880
         End
         Begin VB.TextBox txtInfo 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   10
            Left            =   7920
            TabIndex        =   16
            Top             =   120
            Width           =   2880
         End
         Begin VB.TextBox txtInfo 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   9
            Left            =   4620
            TabIndex        =   15
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   8
            Left            =   4620
            TabIndex        =   14
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox txtInfo 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   12
            Left            =   960
            TabIndex        =   13
            Text            =   "结核"
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "诊断描述2:"
            Height          =   255
            Index           =   11
            Left            =   6960
            TabIndex        =   22
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "诊断描述1:"
            Height          =   255
            Index           =   10
            Left            =   6960
            TabIndex        =   21
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "确诊日期:"
            Height          =   255
            Index           =   9
            Left            =   3720
            TabIndex        =   20
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "发病日期:"
            Height          =   255
            Index           =   8
            Left            =   3720
            TabIndex        =   19
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "疑似疾病:"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Image imgPatiPhoto 
         Enabled         =   0   'False
         Height          =   1000
         Left            =   3720
         Picture         =   "frmDiseaseStation.frx":6905
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Image imgPati 
         Height          =   1365
         Left            =   120
         Picture         =   "frmDiseaseStation.frx":72F5
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1290
      End
   End
   Begin VB.PictureBox picRegist 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   3495
      Left            =   1080
      ScaleHeight     =   3495
      ScaleWidth      =   6855
      TabIndex        =   4
      Top             =   4320
      Width           =   6855
      Begin VB.PictureBox PicSendContent 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   240
         ScaleHeight     =   1455
         ScaleWidth      =   2295
         TabIndex        =   7
         Top             =   1920
         Width           =   2295
         Begin XtremeReportControl.ReportControl rptSendContent 
            Height          =   1455
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1935
            _Version        =   589884
            _ExtentX        =   3413
            _ExtentY        =   2566
            _StockProps     =   0
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
      End
      Begin VB.PictureBox PicAuditContent 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   1920
         ScaleHeight     =   1935
         ScaleWidth      =   2775
         TabIndex        =   5
         Top             =   120
         Width           =   2775
         Begin XtremeReportControl.ReportControl rptAuditContent 
            Height          =   1455
            Left            =   600
            TabIndex        =   6
            Top             =   120
            Width           =   2415
            _Version        =   589884
            _ExtentX        =   4260
            _ExtentY        =   2566
            _StockProps     =   0
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
      End
      Begin XtremeSuiteControls.TabControl tabContent 
         Height          =   1575
         Left            =   2760
         TabIndex        =   9
         Top             =   1920
         Width           =   2775
         _Version        =   589884
         _ExtentX        =   4895
         _ExtentY        =   2778
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   5760
      ScaleHeight     =   4935
      ScaleWidth      =   12015
      TabIndex        =   0
      Top             =   3480
      Width           =   12015
      Begin VB.PictureBox picDis 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   5640
         ScaleHeight     =   3255
         ScaleWidth      =   6255
         TabIndex        =   1
         Top             =   1440
         Width           =   6255
         Begin XtremeSuiteControls.TabControl tbcDis 
            Height          =   2175
            Left            =   360
            TabIndex        =   2
            Top             =   120
            Width           =   4575
            _Version        =   589884
            _ExtentX        =   8070
            _ExtentY        =   3836
            _StockProps     =   64
         End
      End
      Begin XtremeSuiteControls.TabControl tbcMain 
         Height          =   1575
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   4575
         _Version        =   589884
         _ExtentX        =   8070
         _ExtentY        =   2778
         _StockProps     =   64
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgTemp 
      Height          =   1215
      Left            =   7800
      TabIndex        =   10
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
      _cx             =   2566
      _cy             =   2143
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   75
      Top             =   8280
      Width           =   18240
      _ExtentX        =   32173
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiseaseStation.frx":A702
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   29263
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   2160
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":AF94
            Key             =   ""
            Object.Tag             =   "已删除"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":117F6
            Key             =   ""
            Object.Tag             =   "待填写"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":18058
            Key             =   ""
            Object.Tag             =   "待返修"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":1E8BA
            Key             =   ""
            Object.Tag             =   "待上报"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":2511C
            Key             =   ""
            Object.Tag             =   "待审核"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":2B97E
            Key             =   ""
            Object.Tag             =   "返修待审核"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":321E0
            Key             =   ""
            Object.Tag             =   "已上报"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":38A42
            Key             =   ""
            Object.Tag             =   "待处理"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":3F2A4
            Key             =   ""
            Object.Tag             =   "非传染病"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmDiseaseStation.frx":45B06
      Left            =   840
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDiseaseStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    图标 = 0
    ID = 1
    状态 = 2
    报告 = 3
    来源 = 4
    科室 = 5
    就诊号 = 6
    姓名 = 7
    性别 = 8
    年龄 = 9
    填报时间 = 10
    填报人 = 11
    报卡类别 = 12
    疑似疾病 = 13
    登记人 = 14
    登记时间 = 15
    填报备注 = 16
    删除人 = 17
    删除时间 = 18
    报送人 = 19
    报送时间 = 20
    报送单位 = 21
    报送备注 = 22
    数据转出 = 23
    病人ID = 24
    主页ID = 25
    文件ID = 26
    编辑方式 = 27
    信息 = 28
End Enum

Private Enum mRptCol
    ID = 0
    次数 = 1
    反馈内容 = 2
    处理说明 = 3
    登记人 = 4
    登记时间 = 5
    处理人 = 6
    处理时间 = 7
End Enum

Private Enum mSendRptCol
    ID = 0
    次数 = 1
    反馈内容 = 2
    处理说明 = 3
    登记人 = 4
    登记时间 = 5
End Enum

Private Enum mCtlID
    txt姓名 = 0
    txt性别 = 1
    txt年龄 = 2
    txt标识号 = 3
    txt科室 = 4
    txt职业 = 5
    txt电话 = 6
    txt地址 = 7
    txt发病日期 = 8
    txt确诊日期 = 9
    txt诊断描述1 = 10
    txt诊断描述2 = 11
    txt疑似疾病 = 12

    chk待审核 = 0
    chk待返修 = 1
    chk返修待审核 = 2
    chk待填写 = 0
    chk非传染病 = 1
    chk待处理 = 2
    chk待上报 = 0
    chk已上报 = 1
End Enum

Private Enum mTcbID
    tcb未填写 = 0
    tcb审核 = 1
    tcb上报 = 2
    tcb已删除 = 3
    tcb查重工作 = 4
End Enum

Private Const conPane_Reports = 1
Private Const conPane_AppInfo = 2
Private Const conPane_Preview = 3
Private Const conPane_Feedback = 4

Private mTcbSelectID As mTcbID           '选中了TabControl的页面
Private mstrState As String              '筛选时报告的处理状态

Private mstrPrivs As String              '当前使用者权限串
Private mstrFiles As String              '本机管理的报告文件

Private mintWaitIndex As Integer  '未填写时时间选择的项
Private mintDelIndex As Integer   '已删除时时间选择的项
Private mintIndex As Integer      '审核和上报工作时时间选择的项

'审核工作和上报工作 查看报告的记录天数与范围
Private mintDates As Integer              '默认查看最近记录的天数；为0时，说明进入程序执行参数设置，要求按日期范围查看
Private mdtFrom As Date, mdtTo As Date    '按范围查看开始日期，在mintDates=0时有效;按范围查看截止日期，在mintDates=0时有效

'未填写 查看报告的记录天数与范围
Private mintWaitDays As Integer
Private mdtWaitBegin As Date, mdtWaitEnd As Date

'已删除 查看报告的记录天数与范围
Private mintDelDays As Integer
Private mdtDelBegin As Date, mdtDelEnd As Date

Private mfrmPreview As frmDockEPRContent                 '报告内容预览窗格
Private mfrmPreFeedBack As frmDockEPRContent             '已关联的反馈单报告内容预览窗格
Private mobjInfection As Object                          '中华人民共和国传染病报告卡

Private mstrCurId As String               '当前记录ID EMR库的ID是字符型
Private mstrContent As String             '新病历的XML内容
Private mIntState As Integer              '0-待接收；-1-已拒收；1-待审核；4-待返修；3-待上报；2-已上报；5-返修待审；6-已删除
Private mblnCurMoved As Boolean           '当前记录转出状态 0-未转出 1-已转出
Private mstrFindType As String            '查找病人时查找的类型
Private mdatTime As Date                  '查看选中的上报报告的反馈说明记录的处理时间

Private mblnReportCheck As Boolean        '是否只是显示查重的页面，医生站调用接口显示一年内是否有传入病人的重复报告时用
Private mrsOld As ADODB.Recordset         '查重时老版电子病历查询到的数据
Private mblnReport As Boolean             '当前显示的是否是报告卡（true-报告卡；false-反馈单）
Private mlngID As Long                    '当前显示的报告卡或者反馈单的ID


'查看阳性结果反馈单时调用该界面时的一些变量
Private mstrName As String
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mIntPatiFrom As Integer
Private mlng科室ID As Long
Private mstr疾病ID As String
Private mstr诊断ID As String

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl, strInfo As String
    Dim strSqlFile As String, rsTmp As ADODB.Recordset
    
    If mblnCurMoved And (Control.ID = conMenu_File_Open Or Control.ID = conMenu_Edit_Reuse Or Control.ID = conMenu_Edit_Send Or Control.ID = conMenu_Edit_Untread) Then
        MsgBox "该病人的本次数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                        "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
    Select Case Control.ID
        Case conMenu_File_PrintSet: Call zlPrintSet
        Case conMenu_File_Preview: Call zlEPRPrint(True)
        Case conMenu_File_Print: Call zlEPRPrint(False)
        Case conMenu_File_RowPrint: Call zlRptPrint(1)
        Case conMenu_File_Parameter
            Call frmDiseaseStationSet.ShowMe(Me, InStr(1, mstrPrivs, "范围设置") > 0, mstrFiles)
            Call zlRefList
        Case conMenu_File_Exit
             Unload Me
        Case conMenu_Edit_Audit
                Dim intAduitState As Integer
                 With rptList
                    strInfo = .FocusedRow.Record.Item(mCol.报告).Value
                    strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.科室).Value
                    strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.姓名).Value
                    strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.性别).Value
                    strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.年龄).Value
                    strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.病人ID).Value
                    strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.主页ID).Value
                    strInfo = strInfo & "|" & IIf(.FocusedRow.Record.Item(mCol.来源).Value = "住院", "2", "1")
                    strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.填报人).Value
                    strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.填报时间).Value
                End With
                If frmDiseaseAduit.ShowDiseaseAudit(Me, mstrCurId, strInfo, intAduitState) Then
                    '(intAduitState = 3)审核通过，处理状态=3;要求返修，处理状态=4；发送消息到门诊/住院医生站
                    If intAduitState = 4 Then
                        Call SendMsg    '发送消息
                    End If
                    Call zlRefList(mstrCurId)
                End If
        Case conMenu_Edit_Delete
            If mstrCurId <> "" Then
                If MsgBox("您确定要删除该报告吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    If DeleteReport(mstrCurId) Then
                        If rptList.SelectedRows.Count > 0 Then
                            Call rptList.Records.RemoveAt(rptList.FocusedRow.Record.Index)
                            Call rptList.Populate
                        End If
                        If Me.rptList.Rows.Count > 0 Then
                            If Me.rptList.FocusedRow Is Nothing Then
                                Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
                                Call rptList_SelectionChanged
                            End If
                        Else
                            mstrCurId = ""
                            Call rptList_SelectionChanged
                        End If
                    End If
                End If
            End If
        Case conMenu_Edit_Send
            'strInfo=报告|科室|姓名|性别|年龄|就诊号|填报人|填报时间|病人ID|主页ID
            With rptList
                strInfo = .FocusedRow.Record.Item(mCol.报告).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.科室).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.姓名).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.性别).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.年龄).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.就诊号).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.填报人).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.填报时间).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.病人ID).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.主页ID).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.文件ID).Value
            End With
            If frmDiseaseReportSend.ShowMe(Me, mstrCurId, strInfo, txtInfo(txt诊断描述1).Text, txtInfo(txt诊断描述2).Text) Then Call zlRefList(mstrCurId)
        Case conMenu_Edit_Untread  '收回
            Dim strMsg As String
            Dim strSQL As String
            Select Case mIntState
                Case 2
                    If CheckUntread() Then
                        strMsg = "真的取消该疾病报告的“申报登记”吗？"
                    Else
                        MsgBox "该上报的报告已经进行了后续处理，不允许取消。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Case 3:  strMsg = "真的取消该疾病报告的“审核通过”吗？"
                Case 4:  strMsg = "真的取消该疾病报告的“要求返修”吗？"
                Case 6:  strMsg = "真的取消该疾病报告的“删除处理”吗？"
                Case Else: Exit Sub
            End Select
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            strSQL = "Zl_疾病申报记录_Untread('" & mstrCurId & "',1)"
            Err = 0: On Error GoTo ErrHand
            Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            Call zlRefList(mstrCurId)
        Case conMenu_View_ToolBar_Button
            Me.cbsMain(2).Visible = Not Me.cbsMain(2).Visible
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each cbrControl In Me.cbsMain(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Size
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        Case conMenu_View_StatusBar
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsMain.RecalcLayout
        Case conMenu_View_Refresh
            Call zlRefList(mstrCurId)
        Case conMenu_Edit_NewTable
            If WriteReport Then
                Unload Me
            End If
        Case conMenu_Edit_Add
            Call EditSendInfo(1)
        Case conMenu_Edit_Modify
            Call EditSendInfo(2)
        Case conMenu_Edit_EditInfo
            strSqlFile = "select t.病人id,t.主页id,t.科室id,t.婴儿,t.病人来源 from 电子病历记录 t where t.id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSqlFile, "文件信息查询", mstrCurId)
            If rsTmp.RecordCount <> 0 Then
                 mobjInfection.OpenDoc Me, 1, rsTmp!病人ID, rsTmp!主页ID, rsTmp!病人来源, Val(rsTmp!婴儿 & ""), rsTmp!科室ID, mstrCurId
            End If
        Case conMenu_Help_Web_Home: Call gobjComlib.zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Forum '中联论坛
            Call gobjComlib.zlWebForum(Me.hwnd)
        Case conMenu_Help_Web_Mail: Call gobjComlib.zlMailTo(Me.hwnd)
        Case conMenu_Help_About:    Call gobjComlib.ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case Else
            '执行发布到当前模块的报表
            Dim lng报告ID As Long
            If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                If rptList.SelectedRows.Count > 0 Then
                    If Not rptList.SelectedRows(0).GroupRow Then
                        lng报告ID = Val(rptList.SelectedRows(0).Record(mCol.ID).Value)
                    End If
                End If
                If lng报告ID <> 0 Then
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "报告ID=" & lng报告ID)
                Else
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
                End If
            End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    Err = 0: On Error Resume Next
    Select Case Control.ID
        Case conMenu_File_Preview
            Control.Enabled = (mlngID > 0)
            If Control.Enabled Then Control.Enabled = (Me.rptList.FocusedRow.Record.Item(mCol.编辑方式).Value <> 2)
        Case conMenu_File_Print
             Control.Enabled = (mlngID > 0)
        Case conMenu_File_RowPrint
             Control.Enabled = (Me.rptList.Records.Count <> "")
        Case conMenu_Edit_Audit
            Control.Visible = (InStr(1, mstrPrivs, "审核") > 0)
            If Control.Visible Then Control.Visible = (mTcbSelectID = tcb审核)
            If Control.Visible Then Control.Visible = chkAduitState(chk待审核).Value = 1
            If Control.Visible Then Control.Visible = (tbcMain.Item(0).Selected)
            Control.Enabled = (mstrCurId <> "" And (mIntState = -1 Or mIntState = 0 Or mIntState = 1 Or mIntState = 5))
        Case conMenu_Edit_Send
            Control.Visible = (InStr(1, mstrPrivs, "报送") > 0)
            If Control.Visible Then Control.Visible = (mTcbSelectID = tcb上报)
            If Control.Visible Then Control.Visible = chkSendState(chk待上报).Value = 1
            If Control.Visible Then Control.Visible = (tbcMain.Item(0).Selected)
            Control.Enabled = (mstrCurId <> "" And mIntState = 3)
        Case conMenu_Edit_Untread
            Control.Visible = (InStr(1, mstrPrivs, "回退") > 0)
            If Control.Visible Then Control.Visible = (mTcbSelectID <> tcb查重工作 And mTcbSelectID <> tcb未填写)
            If Control.Visible Then Control.Visible = (tbcMain.Item(0).Selected)
            Control.Enabled = (mstrCurId <> "" And (mIntState = 2 Or mIntState = 3 Or mIntState = 4 Or mIntState = 6))
        Case conMenu_Edit_Delete
            Control.Visible = (InStr(1, mstrPrivs, "删除") > 0)
            If Control.Visible Then Control.Visible = (mTcbSelectID = tcb查重工作)
            If Control.Visible Then Control.Visible = (tbcMain.Item(0).Selected)
            Control.Enabled = (mstrCurId <> "" And (mIntState = -1 Or mIntState = 1 Or mIntState = 2 Or mIntState = 3 Or mIntState = 5))
        Case conMenu_Edit_NewTable
            Control.Visible = mIntPatiFrom <> 0 And mblnReportCheck
        Case conMenu_Edit_Add
            Control.Visible = (InStr(1, mstrPrivs, "报送") > 0)
            If Control.Visible Then Control.Visible = (mTcbSelectID = tcb上报)
            If Control.Visible Then Control.Visible = chkSendState(chk已上报).Value = 1
            If Control.Visible Then Control.Visible = (tbcMain.Item(0).Selected)
            Control.Enabled = (mstrCurId <> "" And mIntState = 2)
            
        Case conMenu_Edit_EditInfo           '修改报告卡
            Control.Visible = (InStr(1, mstrPrivs, "审核") > 0) And mblnReport And Me.rptList.FocusedRow.Record.Item(mCol.编辑方式).Value = 2
            If Control.Visible Then Control.Visible = (mTcbSelectID = tcb审核)
            If Control.Visible Then Control.Visible = (tbcMain.Item(0).Selected)
            Control.Enabled = (mstrCurId <> "" And (mIntState = -1 Or mIntState = 0 Or mIntState = 1 Or mIntState = 5)) And Me.rptList.FocusedRow.Record.Item(mCol.编辑方式).Value = 2
            
        Case conMenu_Edit_Modify
            Control.Visible = (InStr(1, mstrPrivs, "报送") > 0)
            If Control.Visible Then Control.Visible = (mTcbSelectID = tcb上报)
            If Control.Visible Then Control.Visible = chkSendState(chk已上报).Value = 1
            If Control.Visible Then Control.Visible = (tbcMain.Item(0).Selected)
            Control.Enabled = (mstrCurId <> "" And mIntState = 2 And mdatTime <> 0)
        Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsMain(2).Visible
        Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
        Case conMenu_View_Refresh
            Control.Visible = (Not mblnReportCheck And mTcbSelectID <> tcb查重工作)
            Control.Enabled = (Trim(mstrFiles) <> "")
    End Select
End Sub

Private Sub chkAduitState_Click(Index As Integer)
'审核工作的条件过滤
    Dim i As Integer
    Dim strState As String, strTemp As String
    mstrState = ""
    For i = chkAduitState.LBound To chkAduitState.UBound
        If chkAduitState(i).Value = 1 Then
            Select Case i
                Case chk待审核
                    strState = strState & ",-1,0,1"
                    strTemp = " or S.处理状态 is null "
                Case chk待返修
                    strState = strState & ", 4"
                Case chk返修待审核
                    strState = strState & ", 5"
            End Select
        End If
    Next
    If strState <> "" Then
        If strTemp <> "" Then
            mstrState = " and (S.处理状态 in (" & Mid(strState, 2) & ")" & strTemp & ") "
        Else
            mstrState = " and S.处理状态 in (" & Mid(strState, 2) & ") "
        End If
    End If
    If Me.Visible And (Index <> -1) Then Call zlRefList
End Sub

Private Sub chkDisState_Click(Index As Integer)
'未填写的条件过滤
    Dim i As Integer
    Dim strState As String
    For i = chkDisState.LBound To chkDisState.UBound
        If chkDisState(i).Value = 1 Then
            Select Case i
                Case chk待填写
                    strState = strState & ", 2, 4"
                Case chk非传染病
                    strState = strState & ", 3"
                Case chk待处理
                    strState = strState & ", 1"
            End Select
        End If
    Next
    mstrState = IIf(strState <> "", " and A.记录状态 in (" & Mid(strState, 2) & ") ", "")
    If Me.Visible And (Index <> -1) Then Call zlRefList
End Sub

Private Sub chkSendState_Click(Index As Integer)
''上报工作的条件过滤
    Dim i As Integer
    Dim strState As String
    For i = chkSendState.LBound To chkSendState.UBound
        If chkSendState(i).Value = 1 Then
            Select Case i
                Case chk待上报
                    strState = strState & ", 3"
                Case chk已上报
                    strState = strState & ", 2"
            End Select
        End If
    Next
    mstrState = IIf(strState <> "", " and S.处理状态 in (" & Mid(strState, 2) & ") ", "")
    If Me.Visible And (Index <> -1) Then Call zlRefList
End Sub

Private Sub Form_Load()
    Dim strState As String
    Dim arrayState() As String
    Dim dtCurDate As Date
    Dim strBegin As String, strEnd As String
On Error Resume Next
    If mblnReportCheck Then
        Me.Caption = "传染病报告卡重复提醒"
    Else
         Me.Caption = "传染病管理工作站"
    End If

    Call PatiIdentify.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser)
    PatiIdentify.objIDKind.AllowAutoICCard = True
    PatiIdentify.objIDKind.AllowAutoIDCard = True
    PatiIdentify.ShowSortName = True

    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs

    mstrFiles = Trim(gobjComlib.zlDatabase.GetPara("本工作站可管理文件", glngSys, 1278))
    strState = CStr(gobjComlib.zlDatabase.GetPara("传染病系统查看状态范围", glngSys, 1278))
    
    If strState <> "" Then
        arrayState = Split(strState, ",")
        chkAduitState(chk待审核).Value = Val(arrayState(0))
        chkAduitState(chk待返修).Value = Val(arrayState(1))
        chkAduitState(chk返修待审核).Value = Val(arrayState(2))
        chkSendState(chk待上报).Value = Val(arrayState(3))
        chkSendState(chk已上报).Value = Val(arrayState(4))
        chkDisState(chk待填写).Value = Val(arrayState(5))
        chkDisState(chk非传染病).Value = Val(arrayState(6))
        chkDisState(chk待处理).Value = Val(arrayState(7))
    End If

      '查询参数
    dtCurDate = gobjComlib.zlDatabase.Currentdate
    mintDates = Val(gobjComlib.zlDatabase.GetPara("审核与上报工作状态下查看最近天数的报告", glngSys, 1278))

    If mintDates = -1 Then
        strBegin = gobjComlib.zlDatabase.GetPara("审核与上报工作状态下查看指定天数的报告的起始天数", glngSys, 1278)
        strEnd = gobjComlib.zlDatabase.GetPara("审核与上报工作状态下查看指定天数的报告的结束天数", glngSys, 1278)
        If IsDate(strEnd) Then
            mdtTo = strEnd
        Else
            mdtTo = dtCurDate
        End If
        If IsDate(strBegin) Then
            mdtFrom = strBegin
        Else
            mdtFrom = CDate(mdtTo - 7)
        End If
    Else
        mdtFrom = CDate(dtCurDate - 7)
        mdtTo = dtCurDate
    End If

    mintWaitDays = Val(gobjComlib.zlDatabase.GetPara("未填写状态下查看最近天数的报告", glngSys, 1278))
    If mintWaitDays = -1 Then
        strBegin = gobjComlib.zlDatabase.GetPara("未填写状态下查看指定天数的报告的起始天数", glngSys, 1278)
        strEnd = gobjComlib.zlDatabase.GetPara("未填写状态下查看指定天数的报告的结束天数", glngSys, 1278)
        If IsDate(strEnd) Then
            mdtWaitEnd = strEnd
        Else
            mdtWaitEnd = dtCurDate
        End If
        If IsDate(strBegin) Then
            mdtWaitBegin = strBegin
        Else
            mdtWaitBegin = CDate(mdtWaitEnd - 7)
        End If
    Else
        mdtWaitBegin = CDate(dtCurDate - 7)
        mdtWaitEnd = dtCurDate
    End If

    mintDelDays = Val(gobjComlib.zlDatabase.GetPara("已删除状态下查看最近天数的报告", glngSys, 1278))
    If mintDelDays = -1 Then
        strBegin = gobjComlib.zlDatabase.GetPara("已删除状态下查看指定天数的报告的起始天数", glngSys, 1278)
        strEnd = gobjComlib.zlDatabase.GetPara("已删除状态下查看指定天数的报告的结束天数", glngSys, 1278)
        If IsDate(strEnd) Then
            mdtDelEnd = strEnd
        Else
            mdtDelEnd = dtCurDate
        End If
        If IsDate(strBegin) Then
            mdtDelBegin = strBegin
        Else
            mdtDelBegin = CDate(mdtDelEnd - 7)
        End If
    Else
        mdtDelBegin = CDate(dtCurDate - 7)
        mdtDelEnd = dtCurDate
    End If

    Call gobjComlib.ZLCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)

     '设置词句显示停靠窗格
    Set mfrmPreview = New frmDockEPRContent
    Set mfrmPreFeedBack = New frmDockEPRContent
    Set mobjInfection = DynamicCreate("zlDisReportCard.clsDisReportCard", "传染病报告卡", True)
    If Not mobjInfection Is Nothing Then
        mobjInfection.Init gcnOracle, glngSys
    End If

    Call InitCommandBar
    Call InitDkpMain
    Call InitReportControl
    Call InitTabContol
    Call InitCboSelectTime

    If Not mblnReportCheck Then
        mTcbSelectID = Val(gobjComlib.zlDatabase.GetPara("当前查看报告的工作状态", glngSys, 1278))
        Me.tbcReportList.Item(mTcbSelectID).Selected = True
    End If

'     数据装入
    If mblnReportCheck Then
        Call SetDuplicateReportData(mrsOld)
    Else
        If mstrFiles = "" Then
            Me.stbThis.Panels(2).Text = "未设置本工作站的疾病报告范围"
        Else
            Call zlRefList
        End If
        '界面恢复
        Call gobjComlib.RestoreWinState(Me, App.ProductName)
    End If
    Me.WindowState = vbMaximized
End Sub

Private Sub InitReportControl()
    Dim rptCol As ReportColumn

    With Me.rptList
        Set rptCol = .Columns.Add(mCol.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.状态, "状态", 90, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.报告, "报告", 0, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.来源, "来源", 40, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.科室, "科室", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.就诊号, "标识号", 75, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.姓名, "姓名", 60, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.性别, "性别", 40, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.年龄, "年龄", 40, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.填报时间, "填报时间", 100, True): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.填报人, "填报人", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.报卡类别, "报卡类别", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.疑似疾病, "疑似疾病", 100, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.登记人, "登记人", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.登记时间, "登记时间", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.填报备注, "填报备注", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.删除人, "删除人", 50, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.删除时间, "删除时间", 50, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.报送人, "报送人", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.报送时间, "报送时间", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.报送单位, "报送单位", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.报送备注, "报送备注", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.数据转出, "数据转出", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.病人ID, "病人ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.主页ID, "主页ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.文件ID, "文件ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.编辑方式, "编辑方式", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.信息, "信息", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
    
        .GroupsOrder.Add .Columns.Find(mCol.状态)
        .GroupsOrder(0).SortAscending = True
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With

    With Me.rptAuditContent
        Set rptCol = .Columns.Add(mRptCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mRptCol.次数, "次数", 300, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mRptCol.反馈内容, "反馈内容", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mRptCol.处理说明, "处理说明", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mRptCol.登记人, "登记人", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mRptCol.登记时间, "登记时间", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mRptCol.处理人, "处理人", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mRptCol.处理时间, "处理时间", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        .GroupsOrder.Add .Columns(1)
        .GroupsOrder(0).SortAscending = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
    
    With Me.rptSendContent
        Set rptCol = .Columns.Add(mSendRptCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mSendRptCol.次数, "次数", 300, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mSendRptCol.反馈内容, "反馈内容", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mSendRptCol.处理说明, "处理说明", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mSendRptCol.登记人, "登记人", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mSendRptCol.登记时间, "登记时间", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        
        .GroupsOrder.Add .Columns(1)
        .GroupsOrder(0).SortAscending = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
End Sub

Private Sub InitTabContol()
    With Me.tbcDis
        With .PaintManager
            .Appearance = xtpTabAppearanceStateButtons
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
    End With
            
    With Me.tbcMain
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem(0, "报告卡", mfrmPreview.hwnd, 0).Tag = "报告卡"
        .InsertItem(1, "反馈单", picDis.hwnd, 0).Tag = "反馈单"
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    
    With Me.tbcReportList
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With

        If Not mblnReportCheck Then
            .InsertItem(mTcbID.tcb未填写, "未填写", picReportList.hwnd, 0).Tag = "未填写"
            .InsertItem(mTcbID.tcb审核, "审核工作", picReportList.hwnd, 0).Tag = "审核工作"
            .InsertItem(mTcbID.tcb上报, "上报工作", picReportList.hwnd, 0).Tag = "上报工作"
            .InsertItem(mTcbID.tcb已删除, "已删除", picReportList.hwnd, 0).Tag = "已删除"
            .InsertItem(mTcbID.tcb查重工作, "查重工作", picReportList.hwnd, 0).Tag = "查重工作"
            .Item(1).Selected = True
         Else
            .InsertItem(0, mstrName & " 报告列表", picReportList.hwnd, 0).Tag = "传染病报告"
         End If
    End With
    
    With tabContent
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With

        .InsertItem(0, "审核反馈说明", PicAuditContent.hwnd, 0).Tag = "审核反馈说明"
        .InsertItem(1, "上报反馈说明", PicSendContent.hwnd, 0).Tag = "上报反馈说明"
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
End Sub

Private Sub SetFeedbackContent(ByVal strID As String, ByVal intState As Integer, Optional ByVal intType As Integer = 0)
'功能：显示选中报告的反馈情况
'参数: strID 选中报告的ID
'      intState 选中报告的处理状态
'      intType:0-更新审核反馈和上报反馈，切换到审核反馈页面；1-更新审核反馈，切换到审核反馈页面；2-更新上报反馈，切换到上报反馈页面
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rptRcd As ReportRecord
    Dim paneFeedBack As Pane
    Dim blnAudit As Boolean, blnSend As Boolean
On Error GoTo errH:
    Set paneFeedBack = dkpMain.FindPane(conPane_Feedback)
    If (intState = 1 Or intState = 2 Or intState = 3 Or intState = 4 Or intState = 5) Then
        mdatTime = 0
        If (intType = 0 Or intType = 1) And IsNumeric(strID) Then
            strSQL = "Select Rownum As 次数, 文件id, 登记人, 登记时间, 记录状态, 反馈内容, 处理人, 处理时间, 处理情况说明" & vbNewLine & _
                    "From (Select 文件id, 登记人, 登记时间, 记录状态, 反馈内容, 处理人, 处理时间, 处理情况说明" & vbNewLine & _
                    "       From 疾病报告反馈 Where 文件id = [1] Order By 登记时间)"
            Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strID))
            If rsTemp.RecordCount > 0 Then
                paneFeedBack.Closed = False
                blnAudit = True
                Me.rptAuditContent.Records.DeleteAll
                Do While Not rsTemp.EOF
                    Set rptRcd = Me.rptAuditContent.Records.Add()
                    rptRcd.AddItem CStr(NVL(rsTemp!文件ID))
                    rptRcd.AddItem "第" & CStr(NVL(rsTemp!次数)) & "次反馈"
                    rptRcd.AddItem CStr(NVL(rsTemp!反馈内容))
                    rptRcd.AddItem CStr(NVL(rsTemp!处理情况说明))
                    rptRcd.AddItem CStr(NVL(rsTemp!登记人))
                    rptRcd.AddItem CStr(NVL(rsTemp!登记时间))
                    rptRcd.AddItem CStr(NVL(rsTemp!处理人))
                    rptRcd.AddItem CStr(NVL(rsTemp!处理时间))
                    rsTemp.MoveNext
                Loop
                rptAuditContent.Populate
            End If
        End If
        
        If (intType = 0 Or intType = 2) And intState = 2 Then
            If IsNumeric(strID) Then
                strSQL = "Select Rownum As 次数, 申报id, 反馈信息, 登记人, 登记时间, 处理情况说明" & vbNewLine & _
                            "From (Select 申报id, 反馈信息, 登记人, 登记时间, 处理情况说明 From 疾病申报反馈 Where 申报id = [1] Order By 登记时间)"
    
                Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strID))
                If rsTemp.RecordCount > 0 Then
                    blnSend = True
                    paneFeedBack.Closed = False
                    Me.rptSendContent.Records.DeleteAll
                    Do While Not rsTemp.EOF
                        Set rptRcd = Me.rptSendContent.Records.Add()
                        rptRcd.AddItem CStr(NVL(rsTemp!申报ID))
                        rptRcd.AddItem "第" & CStr(NVL(rsTemp!次数)) & "次反馈"
                        rptRcd.AddItem CStr(NVL(rsTemp!反馈信息))
                        rptRcd.AddItem CStr(NVL(rsTemp!处理情况说明))
                        rptRcd.AddItem CStr(NVL(rsTemp!登记人))
                        rptRcd.AddItem CStr(NVL(rsTemp!登记时间))
                        rsTemp.MoveNext
                    Loop
                    rptSendContent.Populate
                End If
            End If
        End If
        If blnAudit And blnSend Then
            tabContent.Item(1).Visible = True
            tabContent.Item(0).Selected = True
        ElseIf Not blnAudit And blnSend Then
            tabContent.Item(1).Visible = True
            tabContent.Item(0).Selected = False
        ElseIf blnAudit And Not blnSend Then
            tabContent.Item(1).Visible = False
            tabContent.Item(0).Selected = True
        ElseIf Not blnAudit And Not blnSend Then
            paneFeedBack.Close
        End If
    Else
        paneFeedBack.Close
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitDkpMain()
    Dim objPane As Pane

    Set objPane = Me.dkpMain.CreatePane(conPane_Reports, 300, 400, DockLeftOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    objPane.Title = "病人报告列表"
    objPane.MinTrackSize.Width = 350
    objPane.MaxTrackSize.Width = 360

    Set objPane = Me.dkpMain.CreatePane(conPane_Feedback, 300, 250, DockBottomOf, objPane)
    objPane.Options = PaneNoFloatable Or PaneNoHideable
    objPane.Title = "反馈说明"

    Set objPane = Me.dkpMain.CreatePane(conPane_AppInfo, ScaleX(Me.ScaleWidth - picPati.Width, vbTwips, vbPixels), ScaleY(360, vbTwips, vbPixels), DockRightOf, Nothing)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    objPane.Title = ""
    objPane.MinTrackSize.Height = 100
    objPane.MaxTrackSize.Height = 110

    Set objPane = Me.dkpMain.CreatePane(conPane_Preview, ScaleX(Me.ScaleWidth - picPati.Width, vbTwips, vbPixels), 100, DockBottomOf, objPane)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    objPane.Title = ""

    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.HideClient = True
End Sub

Private Sub InitCommandBar()
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim rptCol As ReportColumn
    Dim lngCount As Long

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbsMain.Icons = gobjComlib.ZLCommFun.GetPubIcons
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .UseDisabledIcons = True
        .UseFadedIcons = True           '图标显示为褪色效果
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsMain.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    Me.cbsMain.ActiveMenuBar.Title = "菜单"

    Me.cbsMain.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        If Not mblnReportCheck Then
            Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
            Set cbrControl = .Add(xtpControlButton, conMenu_File_RowPrint, "清单打印(&L)…"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&M)…"): cbrControl.BeginGroup = True
        End If
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        If Not mblnReportCheck Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审核(&A)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_EditInfo, "修改报告卡(&E)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "报送(&S)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "回退(&B)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Add, "新增备注(&X)"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改备注(&M)")
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewTable, "填写报告卡(&N)")
        End If
        cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With

    '快键绑定
    With Me.cbsMain.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, vbKeyF12, conMenu_File_Parameter
        .Add FCONTROL, Asc("A"), conMenu_Edit_Audit
        .Add FCONTROL, Asc("S"), conMenu_Edit_Send
        .Add FCONTROL, Asc("U"), conMenu_Edit_Untread
        .Add 0, VK_F5, conMenu_View_Refresh
    End With

    '工具栏定义
    Set cbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ContextMenuPresent = False                   '工具栏上点击鼠标右键时不弹出设置菜单
    cbrToolBar.ShowTextBelowIcons = False                   '工具栏中的按钮文字显示在图标右侧
    cbrToolBar.EnableDocking xtpFlagHideWrap                '工具栏宽度不足时也不换行
    With cbrToolBar.Controls
        If Not mblnReportCheck Then
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
            Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_EditInfo, "修改报告卡"): cbrControl.Visible = False
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审核")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "报送")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Add, "新增说明"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改说明")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "回退")
            Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewTable, "填写报告卡")
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
            cbrControl.BeginGroup = True
        End If
    End With

    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = conPane_Reports Then
        Item.Handle = picPati.hwnd
    ElseIf Item.ID = conPane_AppInfo Then
        Item.Handle = picBasisNew.hwnd
     ElseIf Item.ID = conPane_Preview Then
        Item.Handle = picMain.hwnd
    ElseIf Item.ID = conPane_Feedback Then
        Item.Handle = picRegist.hwnd
    End If
End Sub

Private Sub Form_Resize()
    Dim paneReports As Pane
    On Error Resume Next
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Set paneReports = dkpMain.FindPane(conPane_Reports)
        paneReports.MinTrackSize.Width = 355
        paneReports.MaxTrackSize.Width = 355
        dkpMain.RecalcLayout
        dkpMain.NormalizeSplitters
        paneReports.MinTrackSize.Width = 0
        paneReports.MaxTrackSize.Width = Me.ScaleWidth
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strChkState As String
On Error Resume Next
    If Not mblnReportCheck Then
        strChkState = chkAduitState(chk待审核).Value & "," & chkAduitState(chk待返修).Value & "," & chkAduitState(chk返修待审核).Value & "," & chkSendState(chk待上报).Value & "," & chkSendState(chk已上报).Value & "," & chkDisState(chk待填写).Value & "," & chkDisState(chk非传染病).Value & "," & chkDisState(chk待处理).Value
        Call gobjComlib.zlDatabase.SetPara("传染病系统查看状态范围", strChkState, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("审核与上报工作状态下查看最近天数的报告", mintDates, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("审核与上报工作状态下查看指定天数的报告的起始天数", mdtFrom, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("审核与上报工作状态下查看指定天数的报告的结束天数", mdtTo, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("未填写状态下查看最近天数的报告", mintWaitDays, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("未填写状态下查看指定天数的报告的起始天数", mdtWaitBegin, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("未填写状态下查看指定天数的报告的结束天数", mdtWaitEnd, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("已删除状态下查看最近天数的报告", mintDelDays, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("已删除状态下查看指定天数的报告的起始天数", mdtDelBegin, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("已删除状态下查看指定天数的报告的结束天数", mdtDelEnd, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("当前查看报告的工作状态", mTcbSelectID, glngSys, 1278)
    End If
    If Not mfrmPreview Is Nothing Then
        Unload mfrmPreview
        Set mfrmPreview = Nothing
    End If
    If Not mfrmPreFeedBack Is Nothing Then
        Unload mfrmPreFeedBack
        Set mfrmPreFeedBack = Nothing
    End If
    Set mobjInfection = Nothing
    mblnReportCheck = False
    Set mrsOld = Nothing
    If Not mblnReportCheck Then Call gobjComlib.SaveWinState(Me, App.ProductName)
End Sub

Private Sub PicAuditContent_Resize()
On Error Resume Next
    rptAuditContent.Move PicAuditContent.ScaleLeft, PicAuditContent.ScaleTop, PicAuditContent.ScaleWidth, PicAuditContent.ScaleHeight
End Sub

Private Sub picDis_Resize()
On Error Resume Next
    tbcDis.Move picDis.ScaleLeft, picDis.ScaleTop, picDis.ScaleWidth, picDis.ScaleHeight
End Sub

Private Sub picMain_Resize()
 On Error Resume Next
    tbcMain.Move picMain.ScaleLeft, picMain.ScaleTop, picMain.ScaleWidth, picMain.ScaleHeight
End Sub

Private Sub picPati_Resize()
    On Error Resume Next
    tbcReportList.Move picPati.ScaleLeft, lblFind.Top + lblFind.Height + 10, picPati.ScaleWidth, picPati.ScaleHeight - lblFind.Top - lblFind.Height - 20
End Sub

Private Sub picRegist_Resize()
    On Error Resume Next
    tabContent.Move picRegist.ScaleLeft, picRegist.ScaleTop, picRegist.ScaleWidth, picRegist.ScaleHeight
End Sub

Private Sub picReportList_Resize()
On Error Resume Next
    If Not mblnReportCheck Then
        If mTcbSelectID = tcb上报 Or mTcbSelectID = tcb审核 Then
            rptList.Move picReportList.ScaleLeft, picState(5).Top + picState(5).Height + picState(mTcbSelectID).Height, picReportList.ScaleWidth, picReportList.ScaleHeight - (picState(5).Top + picState(5).Height + picState(mTcbSelectID).Height + 10)
        Else
            rptList.Move picReportList.ScaleLeft, picState(mTcbSelectID).Top + picState(mTcbSelectID).Height, picReportList.ScaleWidth, picReportList.ScaleHeight - (picState(mTcbSelectID).Top + picState(mTcbSelectID).Height + 10)
        End If
    Else
        rptList.Move picReportList.ScaleLeft, picReportList.ScaleTop, picReportList.ScaleWidth, picReportList.ScaleHeight
    End If
End Sub

Private Sub PicSendContent_Resize()
    On Error Resume Next
    rptSendContent.Move PicSendContent.ScaleLeft, PicSendContent.ScaleTop, PicSendContent.ScaleWidth, PicSendContent.ScaleHeight
End Sub

Private Sub rptList_SelectionChanged()
'功能：选中的报告发生变化
    Dim strInfo As String
    If Not Me.Visible Then Exit Sub

    '清除掉界面上的病人信息
    Call ClearTxtInfo

    mstrContent = "": rptList.Tag = ""
    With Me.rptList
        If .FocusedRow Is Nothing Then
            mstrCurId = "": mIntState = 0: strInfo = "": mblnCurMoved = False
        ElseIf .FocusedRow.GroupRow = True Then
            mstrCurId = "":  mIntState = 0: strInfo = "": mblnCurMoved = False
        Else
            mstrCurId = .FocusedRow.Record.Item(mCol.ID).Value
            mIntState = .FocusedRow.Record.Item(mCol.图标).Value
            mblnCurMoved = (.FocusedRow.Record.Item(mCol.数据转出).Value = 1)
        End If
    End With

'   在界面上显示病人的基本信息
    Call SetPatiInfo

'   显示选中报告的反馈情况
    Call SetFeedbackContent(mstrCurId, mIntState)
    Call GetFeedbackIDs(Val(mstrCurId))
    tbcMain.Item(0).Selected = True
    
    If IsNumeric(mstrCurId) Then
        Call RefreshReport(Val(mstrCurId), mblnCurMoved, NVL(rptList.FocusedRow.Record.Item(mCol.编辑方式).Value, 0))
    Else
        mlngID = 0
    End If
End Sub

Private Function RefreshReport(ByVal lngID As Long, ByVal blnMoved As Boolean, ByVal lngEditType As Long, Optional ByVal lnfType As Long) As Boolean
    dkpMain.FindPane(conPane_Preview).Handle = picMain.hwnd
    If lnfType = 1 Then
        Call mfrmPreFeedBack.zlRefresh(lngID, "", , blnMoved, , lngEditType)
    Else
        Call mfrmPreview.zlRefresh(lngID, "", , blnMoved, , lngEditType)
    End If
    mblnReport = lngEditType <> 3
    mlngID = lngID
End Function

Private Sub ClearTxtInfo()
'功能:清除界面上病人的基本信息
    txtInfo(txt姓名).Text = ""
    txtInfo(txt性别).Text = ""
    txtInfo(txt年龄).Text = ""
    txtInfo(txt标识号).Text = ""
    txtInfo(txt科室).Text = ""
    txtInfo(txt职业).Text = ""
    txtInfo(txt地址).Text = ""
    txtInfo(txt电话).Text = ""
    txtInfo(txt发病日期).Text = ""
    txtInfo(txt确诊日期).Text = ""
    txtInfo(txt诊断描述1).Text = ""
    txtInfo(txt诊断描述2).Text = ""
    fraInfo(0).Visible = False
    fraInfo(1).Visible = False
    imgPati.Visible = False
    txtState.Visible = False
End Sub

Public Sub zlRptPrint(ByVal bytMode As Byte)
'功能:将数据复制到可打印的对象，调用打印
'参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub
    '-------------------------------------------------
    '复制数据表格
    If zlReportToVSFlexGrid(Me.vfgTemp, Me.rptList) = False Then Exit Sub
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow

    Set objPrint.Body = Me.vfgTemp
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

'-------------------------------------------------------
'功能：  报告预览及打印
'参数：  blnPreview  :是否是预览模式
'-------------------------------------------------------
Private Sub zlEPRPrint(blnPreview As Boolean)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim clsPrint As zlRichEPR.clsDockAduits
    
    If mstrCurId = "" Then Exit Sub

    Err = 0: On Error GoTo ErrHand
    If mblnReport And IsNumeric(mstrCurId) Then
        strSQL = "Select l.病人来源, l.病人id, l.主页id,l.编辑方式, f.页面 From 电子病历记录 l, 病历文件列表 f Where l.文件id = f.Id And l.Id = [1]"
        If mblnCurMoved Then
            strSQL = Replace(strSQL, "电子病历记录", "H电子病历记录")
        End If
        Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(mstrCurId))
        With rsTemp
            If .RecordCount <= 0 Then MsgBox "该疾病报告可能已经被临床删除！", vbExclamation, gstrSysName: Exit Sub
            If NVL(rsTemp!编辑方式, 0) = 0 Or NVL(rsTemp!编辑方式, 0) = 1 Then
                Set clsPrint = New zlRichEPR.clsDockAduits
                Call clsPrint.zlPrintDocument(3, IIf(blnPreview, 1, 2), CLng(mstrCurId))
                Set clsPrint = Nothing
            ElseIf rsTemp!编辑方式 = 2 Then
                mobjInfection.PrintDoc Me, !病人ID, !主页ID, CLng(mstrCurId), ""
            End If
        End With
    ElseIf Not mblnReport And mlngID > 0 Then
        Call PrintDiseaseRegist(IIf(blnPreview, 1, 2), mlngID, Me)
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetPatiInfo()
'功能： 在界面上显示病人的基本信息
    Dim strInfo As String, aryInfo() As String

    With Me.rptList
        If .FocusedRow Is Nothing Then
            If Not (mfrmPreview Is Nothing) Then
                dkpMain.FindPane(conPane_Preview).Handle = picMain.hwnd
                Call mfrmPreview.zlRefresh(0, "", , False, , 0)
            End If
            Exit Sub
        ElseIf .FocusedRow.GroupRow = True Then
            If Not (mfrmPreview Is Nothing) Then
                dkpMain.FindPane(conPane_Preview).Handle = picMain.hwnd
                Call mfrmPreview.zlRefresh(0, "", , False, , 0)
            End If
            Exit Sub
        Else
            strInfo = .FocusedRow.Record.Item(mCol.信息).Value
            aryInfo = Split(strInfo, "|")
            imgPati.Visible = True
            txtInfo(txt姓名).Text = .FocusedRow.Record.Item(mCol.姓名).Value
            txtInfo(txt性别).Text = .FocusedRow.Record.Item(mCol.性别).Value
            txtInfo(txt年龄).Text = .FocusedRow.Record.Item(mCol.年龄).Value & "岁"
            txtInfo(txt年龄).Text = Replace(txtInfo(txt年龄).Text, "岁岁", "岁")
            txtInfo(txt标识号).Text = .FocusedRow.Record.Item(mCol.就诊号).Value
            txtInfo(txt科室).Text = .FocusedRow.Record.Item(mCol.科室).Value
            If .FocusedRow.Record.Item(mCol.来源).Value = "门诊" Then
                lblInfo(txt标识号) = "门诊号:"
            ElseIf .FocusedRow.Record.Item(mCol.来源).Value = "住院" Then
                lblInfo(txt标识号) = "住院号:"
            Else
                lblInfo(txt标识号) = "标识号:"
            End If
            txtInfo(txt职业).Text = aryInfo(0)
            txtInfo(txt电话).Text = aryInfo(1)
            txtInfo(txt地址).Text = aryInfo(2)

            Call ReadPatPricture(Val(.FocusedRow.Record.Item(mCol.病人ID).Value), imgPati)
            fraInfo(0).Visible = True
            txtInfo(txt发病日期).Text = aryInfo(3)
            txtInfo(txt确诊日期).Text = aryInfo(4)
            txtInfo(txt诊断描述1).Text = aryInfo(5)
            txtInfo(txt诊断描述2).Text = aryInfo(6)
            txtInfo(txt疑似疾病).Text = .FocusedRow.Record.Item(mCol.疑似疾病).Value
            fraInfo(1).Visible = True
        End If
    End With
End Sub

Public Sub ReadPatPricture(ByVal lng病人ID As Long, ByRef imgPatient As Image)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人照片
    '参数：lng病人ID=读取指定病人的照片
    '           imgPatient=照片加载位置
    '           strFile=照片的本地路径
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFile As String
    On Error GoTo ErrHand
    If lng病人ID = 0 Then Exit Sub
    strFile = ""
    strFile = gobjComlib.sys.Readlob(glngSys, 27, lng病人ID, strFile)
    If strFile <> "" Then
        imgPatient.Picture = Nothing
        imgPatient.Picture = LoadPicture(strFile)
        Kill strFile
    Else
        imgPatient.Picture = imgPatiPhoto.Picture
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
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

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
     mstrFindType = objCard.名称
End Sub

Private Sub PatiIdentify_KeyPress(KeyAscii As Integer)
    Select Case mstrFindType
        Case "住院号"
            If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        Case "门诊号"
            If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        Case "就诊卡"
            If InStr(":：';；?？", Chr(KeyAscii)) > 0 Then
                KeyAscii = 0
            Else
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End If
    End Select
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean, Optional ByVal lngPatiID As Long)
'功能：查找(下一个)病人
'参数：blnNext=是否查找下一个
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long

    '开始查找行
    If rptList.SelectedRows.Count > 0 Then
        If Not rptList.SelectedRows(0).GroupRow Then
            If Val(rptList.SelectedRows(0).Record(mCol.病人ID).Value) <> 0 Then blnHave = True
        End If
    End If

    If Not blnNext Or blnReStart Or Not blnHave Then
        i = 0       'ReportControl的索引从是0开始
    Else
        i = rptList.SelectedRows(0).Index + 1
    End If

    '查找病人
    For i = i To rptList.Rows.Count - 1
        With rptList.Rows(i)
            If Not .GroupRow Then
                If Val(.Record(mCol.病人ID).Value) = lngPatiID And lngPatiID <> 0 Then Exit For

                If mstrFindType = "住院号" Then '住院号
                    If UCase(Trim(.Record(mCol.就诊号).Value)) = UCase(PatiIdentify.Text) And .Record(mCol.来源).Value <> "门诊" Then Exit For
                ElseIf mstrFindType = "门诊号" Then
                    If UCase(Trim(.Record(mCol.就诊号).Value)) = UCase(PatiIdentify.Text) And .Record(mCol.来源).Value <> "住院" Then Exit For
                ElseIf mstrFindType = "姓名" OR mstrFindType = "姓名或就诊卡" Then '姓名
                    If .Record(mCol.姓名).Value Like "*" & PatiIdentify.Text & "*" Then Exit For
                End If
            End If
        End With
    Next

    If i <= rptList.Rows.Count - 1 Then
        blnReStart = False
        '该行选中且显示在可见区域,并引发SelectionChanged事件
        Set rptList.FocusedRow = rptList.Rows(i)
        If rptList.Visible Then rptList.SetFocus
    Else
        MsgBox IIf(blnNext, "后面已", "") & "找不到符合条件的病人。", vbInformation, gstrSysName
    End If
End Sub

Private Function DeleteReport(ByVal strID As String) As Boolean
'功能：删除报告
'参数：strID 所要删除的报告的ID
    Dim strSQL As String
On Error GoTo ErrHand

    strSQL = "Zl_疾病申报记录_Delete(" & strID & ")"
    Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    DeleteReport = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub SetRptListData(ByVal rsTemp As ADODB.Recordset)
'功能：将查询出来的报告的信息加载到ReportControl控件上
    Dim strPatiInfo As String
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem

    If rsTemp Is Nothing Then Exit Sub
    If rsTemp.RecordCount > 0 Then
        Do While Not rsTemp.EOF
            strPatiInfo = CStr(NVL(rsTemp!职业)) & "|" & CStr(NVL(rsTemp!家庭电话)) & "|" & CStr(NVL(rsTemp!家庭地址)) & "|" & CStr(NVL(rsTemp!发病日期)) & "|" & CStr(NVL(rsTemp!确诊日期)) & "|" & CStr(NVL(rsTemp!诊断描述1)) & "|" & CStr(NVL(rsTemp!诊断描述2))
            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(CStr(NVL(rsTemp!状态)))
            Select Case rptItem.Value
                Case -1: rptItem.Icon = 4
                Case 0: rptItem.Icon = 4
                Case 1: rptItem.Icon = 4
                Case 2: rptItem.Icon = 6
                Case 3: rptItem.Icon = 3
                Case 4: rptItem.Icon = 2
                Case 5: rptItem.Icon = 5
                Case 6: rptItem.Icon = 1
                Case 7: rptItem.Icon = 8
                Case 8: rptItem.Icon = 0
                Case 9: rptItem.Icon = 7
            End Select
            rptRcd.AddItem CStr(rsTemp!ID)
            Select Case rsTemp!状态
                Case -1: rptRcd.AddItem CStr("待审核")
                Case 0: rptRcd.AddItem CStr("待审核")
                Case 1: rptRcd.AddItem CStr("待审核")
                Case 2: rptRcd.AddItem CStr("已报送")
                Case 3: rptRcd.AddItem CStr("待上报")
                Case 4: rptRcd.AddItem CStr("待返修")
                Case 5: rptRcd.AddItem CStr("返修待审核")
                Case 6: rptRcd.AddItem CStr("待填写")
                Case 7: rptRcd.AddItem CStr("非传染病")
                Case 8: rptRcd.AddItem CStr("已删除")
                Case 9: rptRcd.AddItem CStr("待处理")
                Case Else: rptRcd.AddItem ""
            End Select
            rptRcd.AddItem CStr(NVL(rsTemp!报告))
            rptRcd.AddItem IIf(Val(rsTemp!病人来源) = 1, "门诊", IIf(Val(rsTemp!病人来源) = 2, "住院", ""))
            rptRcd.AddItem CStr(NVL(rsTemp!科室))
            rptRcd.AddItem CStr(NVL(rsTemp!就诊号))
            rptRcd.AddItem CStr(NVL(rsTemp!姓名))
            rptRcd.AddItem CStr(NVL(rsTemp!性别))
            rptRcd.AddItem CStr(NVL(rsTemp!年龄))
            rptRcd.AddItem CStr(NVL(rsTemp!填报时间))
            rptRcd.AddItem CStr(NVL(rsTemp!填报人))
            rptRcd.AddItem CStr(NVL(rsTemp!报卡类型))
            rptRcd.AddItem CStr(NVL(rsTemp!疑似疾病))
            rptRcd.AddItem CStr(NVL(rsTemp!登记人))
            rptRcd.AddItem CStr(NVL(rsTemp!登记时间))
            rptRcd.AddItem CStr(NVL(rsTemp!填报备注))
            rptRcd.AddItem CStr(NVL(rsTemp!撤档人))
            rptRcd.AddItem CStr(NVL(rsTemp!撤档时间))
            rptRcd.AddItem CStr(NVL(rsTemp!报送人))
            rptRcd.AddItem CStr(NVL(rsTemp!报送时间))
            rptRcd.AddItem CStr(NVL(rsTemp!报送单位))
            rptRcd.AddItem CStr(NVL(rsTemp!报送备注))
            rptRcd.AddItem CStr(NVL(rsTemp!数据转出, 0))
            rptRcd.AddItem CStr(NVL(rsTemp!病人ID, 0))
            rptRcd.AddItem CStr(NVL(rsTemp!主页ID, 0))
            rptRcd.AddItem CStr(NVL(rsTemp!文件ID, 0))
            rptRcd.AddItem CStr(NVL(rsTemp!编辑方式, 0))
            rptRcd.AddItem CStr(strPatiInfo)
            rsTemp.MoveNext
        Loop
    End If
End Sub

Private Function GetReportFiles() As String
'功能：获取管理的报告文件
    Dim i As Integer, strFiles As String

    If Trim(mstrFiles) = "" Then Exit Function

    For i = 0 To UBound(Split(mstrFiles, ","))
        If IsNumeric(Split(mstrFiles, ",")(i)) Then
            strFiles = strFiles & "," & Split(mstrFiles, ",")(i)
        End If
    Next
    If strFiles <> "" Then
        strFiles = Mid(strFiles, 2)
    End If
    GetReportFiles = strFiles
End Function

Private Function zlRefList(Optional strCurId As String) As Long
'功能：刷新显示报告文件
'参数：strCurId 所要定位到的报告文件ID
    Dim rptRow As ReportRow
On Error GoTo ErrHand
    If mblnReportCheck Then Exit Function
    Me.rptList.Records.DeleteAll
    If Trim(mstrFiles) = "" Then Exit Function
    
    dkpMain.FindPane(conPane_Feedback).Close
    If mTcbSelectID = tcb未填写 Then
        Call zlRefWaitFullData
    ElseIf mTcbSelectID = tcb上报 Or mTcbSelectID = tcb审核 Then
        Call zlRefOldData(mintDates, mdtFrom, mdtTo)
    ElseIf mTcbSelectID = tcb已删除 Then
        Call zlRefOldData(mintDelDays, mdtDelBegin, mdtDelEnd)
    End If

    Me.rptList.Populate

    If strCurId <> "" Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If CStr(rptRow.Record(mCol.ID).Value) = strCurId Then
                    Set Me.rptList.FocusedRow = rptRow: Exit For
                End If
            End If
        Next
    End If

    If Me.rptList.Rows.Count > 0 Then
        If Me.rptList.FocusedRow Is Nothing Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        Else
            rptList_SelectionChanged
        End If
    Else
        strCurId = ""
        rptList_SelectionChanged
    End If

    Me.stbThis.Panels(2).Text = "共有" & Me.rptList.Records.Count & "份疾病报告。"
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
End Function

Private Sub zlRefWaitFullData()
'功能：查询未填写的报告文件
    Dim blnMoved As Boolean, strTemp As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strSqlDate As String
On Error GoTo ErrHand

    If mTcbSelectID <> tcb未填写 Then Exit Sub
    If mstrState = "" Then Exit Sub

    If mintWaitDays <> -1 Then
        strSqlDate = " And a.登记时间 >= trunc(Sysdate - [1])"
    Else
        blnMoved = MovedByDate(CDate(mdtWaitBegin))
        strSqlDate = " And a.登记时间 Between [2] And [3]"
    End If
    
    strSQL = "Select A.ID as ID, null As 文件id, a.病人id, a.主页id, '传染病阳性结果反馈单' As 报告, Nvl2(A.挂号单, 1, 2) as 病人来源 , e.名称 As 科室," & vbNewLine & _
           "Nvl(c.住院号, b.门诊号) As 就诊号, Nvl(c.姓名, b.姓名) As 姓名, Nvl(c.性别, b.性别) As 性别, Nvl(c.年龄, b.年龄) As 年龄, Null As 填报时间," & vbNewLine & _
           "Null As 填报人, Null As 报卡类型, 3 As 编辑方式, Decode(A.记录状态,1,9,3,7,6) As 状态, Null As 报送人,Null as 撤档人 ,Null as 撤档时间, " & vbNewLine & _
           "Null As 报送时间, Null As 报送单位, Null As 报送备注, Null As 登记人, Null As 登记时间, F.职业," & vbNewLine & _
           "F.家庭地址, F.家庭电话, Null As 发病日期, Null As 确诊日期,  Null As 诊断描述1, Null As 诊断描述2," & vbNewLine & _
           "Null As 填报备注, a.传染病名称 as  疑似疾病 , 0 As 数据转出 " & vbNewLine & _
           "From 疾病阳性记录 A, 病人挂号记录 B, 病案主页 C, 部门表 E,病人信息 F " & vbNewLine & _
           "Where a.病人id = c.病人id(+) And a.主页id = c.主页id(+) And a.病人id = b.病人id(+) And a.挂号单 = b.No(+) " & vbNewLine & _
           mstrState & " And Nvl(b.执行部门id, c.出院科室id) = e.Id and A.病人ID =F.病人ID  And a.文件id Is Null " & strSqlDate
    If blnMoved Then
        strTemp = Replace(strTemp, "疾病阳性记录", "H疾病阳性记录")
        strSQL = strSQL & " Union All " & strTemp
    End If
    Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mintWaitDays, mdtWaitBegin, mdtWaitEnd)

    Call SetRptListData(rsTemp)
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlRefOldData(ByVal intDates As Integer, ByVal dtFrom As Date, ByVal dtTo As Date)
'功能：查询老版电子病历的报告文件
'参数：intDates     查询最近天数的报告文件
'      dtFrom  查询指定时间段报告文件的起始日期
'      dtTo    查询指定时间段报告文件的终止日期
    Dim blnMoved As Boolean, strTemp As String, strFiles As String
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strDelSql As String
On Error GoTo ErrHand

    strFiles = GetReportFiles()

    If Trim(strFiles) = "" Then Exit Sub
    If mstrState = "" Then Exit Sub
    If mTcbSelectID = tcb上报 Or mTcbSelectID = tcb审核 Then
        strDelSql = " and S.撤档人 Is Null And S.撤档时间 Is Null "
    ElseIf mTcbSelectID = tcb已删除 Then
        strDelSql = " and S.撤档人 Is Not Null And S.撤档时间 Is Not Null "
    End If

    If intDates <> -1 Then
        strSQL = " And l.完成时间 >= trunc(Sysdate - [1])  "
    Else
        blnMoved = MovedByDate(CDate(dtFrom))
        strSQL = " And l.完成时间 Between [2] And [3]"
    End If

    strSQL = "Select l.Id, l.文件id, l.病人id, l.主页id, l.病历名称 As 报告,l.病人来源, d.名称 As 科室," & vbNewLine & _
            " Decode(l.病人来源, 1, p.门诊号, 2, p.住院号) As 就诊号, Nvl(l.姓名, p.姓名) As 姓名, Nvl(l.性别, p.性别) As 性别, Nvl(l.年龄, p.年龄) As 年龄," & vbNewLine & _
            " To_Char(l.完成时间, 'yyyy-mm-dd hh24:mi') As 填报时间, l.保存人 As 填报人, l.报卡类型, l.编辑方式," & vbNewLine & _
            " l.状态 As 状态, l.报送人," & vbNewLine & _
            " To_Char(l.报送时间, 'yyyy-mm-dd hh24:mi') as 报送时间, l.报送单位,l.报送备注,l.登记人," & vbNewLine & _
            " To_Char(l.登记时间, 'yyyy-mm-dd hh24:mi') as 登记时间, Nvl(l.职业,P.职业) as 职业 , Nvl(l.家庭地址,P.家庭地址) as 家庭地址, Nvl(l.家庭电话,P.家庭电话) as 家庭电话 ," & vbNewLine & _
            " To_Char(l.发病日期, 'yyyy-mm-dd') as 发病日期,To_Char(l.确诊日期, 'yyyy-mm-dd') as 确诊日期 , l.诊断描述1 ,l.诊断描述2," & vbNewLine & _
            " l.填报备注 ,L.撤档人,L.撤档时间 , null as  疑似疾病,0 As 数据转出" & vbNewLine & _
            " From (Select l.Id, l.文件id, l.病人id, l.主页id, l.病历名称, l.病人来源, l.科室id, l.完成时间, l.保存人, l.保存时间, l.编辑方式, decode(S.撤档人, null ,Nvl(s.处理状态, 0),8) As 状态," & vbNewLine & _
            "       s.报送人, s.报送时间, s.报送单位, s.报送备注, s.登记人, s.登记时间, s.姓名, s.性别, s.年龄, s.职业, s.家庭地址, s.家庭电话," & vbNewLine & _
            "        s.发病日期, s.确诊日期, s.诊断描述1, s.诊断描述2, s.填报备注, s.报卡类型,s.撤档人,S.撤档时间 " & vbNewLine & _
            "       From 电子病历记录 L, 疾病申报记录 S " & vbNewLine & _
            "       Where l.Id = s.文件id(+) And l.病历种类 = 5 And l.文件id In (" & strFiles & ") " & strSQL & _
               mstrState & strDelSql & " ) L, 病人信息 P, 部门表 D" & vbNewLine & _
            "Where l.病人id = p.病人id And l.科室id = d.Id "

    If blnMoved Then
        strTemp = Replace(strSQL, "0 as 数据转出", "1 as 数据转出")
        strTemp = Replace(strTemp, "电子病历记录", "H电子病历记录")
        strTemp = Replace(strTemp, "疾病申报记录", "H疾病申报记录")
        strSQL = strSQL & " Union All " & strTemp
    End If
    Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, intDates, dtFrom, dtTo)

    Call SetRptListData(rsTemp)
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub rptSendContent_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
'双击修改反馈说明
    With rptSendContent
        If .FocusedRow Is Nothing Then
            mdatTime = 0
            Exit Sub
        ElseIf .FocusedRow.GroupRow = True Then
            mdatTime = 0
            Exit Sub
        Else
            mdatTime = Format(.FocusedRow.Record.Item(mSendRptCol.登记时间).Value, "yyyy-mm-dd HH:MM:SS")
            Call EditSendInfo(2)
        End If
    End With
End Sub

Private Sub rptSendContent_SelectionChanged()
    With rptSendContent
        If .FocusedRow Is Nothing Then
            mdatTime = 0
            Exit Sub
        ElseIf .FocusedRow.GroupRow = True Then
            mdatTime = 0
            Exit Sub
        Else
            mdatTime = Format(.FocusedRow.Record.Item(mSendRptCol.登记时间).Value, "yyyy-mm-dd HH:MM:SS")
        End If
    End With
End Sub

Private Sub tbcDis_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim lngFeedbackID As Long
    If Me.Visible Then
        lngFeedbackID = Val(tbcDis.Selected.Tag)
        Call RefreshReport(lngFeedbackID, False, 3, 1)
        mblnReport = False
        mlngID = lngFeedbackID
    End If
End Sub

Private Sub tbcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Me.Visible Then
        dkpMain.FindPane(conPane_Preview).Handle = picMain.hwnd
        If tbcMain.Selected.Tag = "报告卡" Then
            picDis.Visible = False
            If IsNumeric(mstrCurId) Then
                Call RefreshReport(Val(mstrCurId), mblnCurMoved, NVL(rptList.FocusedRow.Record.Item(mCol.编辑方式).Value, 0))
            End If
        ElseIf tbcMain.Selected.Tag = "反馈单" Then
            picDis.Visible = True
            tbcDis.Item(0).Selected = True
            mblnReport = False
            mlngID = CLng(Val(tbcDis.Selected.Tag))
        End If
    End If
End Sub

Private Sub tbcReportList_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim i As Long

    For i = 0 To picState.Count - 1
        picState(i).Visible = False
    Next

    mTcbSelectID = Item.Index
    Me.rptList.Columns(mCol.删除人).Visible = False
    Me.rptList.Columns(mCol.删除时间).Visible = False

    picState(mTcbSelectID).Visible = True
    txtInfo(txt发病日期).Visible = True
    txtInfo(txt确诊日期).Visible = True
    txtInfo(txt诊断描述1).Visible = True
    txtInfo(txt诊断描述2).Visible = True
    lblInfo(txt发病日期).Visible = True
    lblInfo(txt确诊日期).Visible = True
    lblInfo(txt诊断描述1).Visible = True
    lblInfo(txt诊断描述2).Visible = True

    If Item.Tag = "审核工作" Then
        picState(5).Visible = True
        picState(5).Move 100, 50
        picState(mTcbSelectID).Move 100, 450
        Call chkAduitState_Click(-1)
    ElseIf Item.Tag = "上报工作" Then
        picState(5).Visible = True
        picState(5).Move 100, 50
        picState(mTcbSelectID).Move 100, 450
        Call chkSendState_Click(-1)
    ElseIf Item.Tag = "未填写" Then
        picState(mTcbSelectID).Move 100, 50
        txtInfo(txt发病日期).Visible = False
        txtInfo(txt确诊日期).Visible = False
        txtInfo(txt诊断描述1).Visible = False
        txtInfo(txt诊断描述2).Visible = False
        lblInfo(txt发病日期).Visible = False
        lblInfo(txt确诊日期).Visible = False
        lblInfo(txt诊断描述1).Visible = False
        lblInfo(txt诊断描述2).Visible = False
        Call chkDisState_Click(-1)
    ElseIf Item.Tag = "已删除" Then
        Me.rptList.Columns(mCol.删除人).Visible = True
        Me.rptList.Columns(mCol.删除时间).Visible = True
        picState(mTcbSelectID).Move 100, 50
        mstrState = " and S.撤档人 is not null "
    ElseIf Item.Tag = "查重工作" Then
        picState(mTcbSelectID).Move 50, 50
        lblNotice.Visible = False
    End If

    Call picReportList_Resize
    tbcMain.Item(1).Visible = False
    tbcMain.PaintManager.ButtonMargin.SetRect -20, -20, 0, 0
    tbcMain.Item(0).Selected = True
    If Me.Visible Then
        Call zlRefList
    End If
End Sub

Private Sub InitCboSelectTime()
    cboSelectTime(0).Clear         '未填写
    With cboSelectTime(0)
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
        .AddItem "30天内"
        .ItemData(.NewIndex) = 30
        .AddItem "[指定...]"
        .ItemData(.NewIndex) = -1
    End With
    If cboSelectTime(0).ListCount > 0 Then
        Select Case mintWaitDays
            Case 0
                cboSelectTime(0).ListIndex = 0
            Case 1
                cboSelectTime(0).ListIndex = 1
            Case 2
                cboSelectTime(0).ListIndex = 2
            Case 7
                cboSelectTime(0).ListIndex = 3
            Case 15
                cboSelectTime(0).ListIndex = 4
            Case 30
                cboSelectTime(0).ListIndex = 5
            Case -1
                cboSelectTime(0).ListIndex = 6
        End Select
        mintWaitIndex = cboSelectTime(0).ListIndex
    End If

    cboSelectTime(1).Clear         '已删除
    With cboSelectTime(1)
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
         .AddItem "30天内"
        .ItemData(.NewIndex) = 30
        .AddItem "[指定...]"
        .ItemData(.NewIndex) = -1
    End With
    If cboSelectTime(1).ListCount > 0 Then
        Select Case mintDelDays
            Case 0
                cboSelectTime(1).ListIndex = 0
            Case 1
                cboSelectTime(1).ListIndex = 1
            Case 2
                cboSelectTime(1).ListIndex = 2
            Case 7
                cboSelectTime(1).ListIndex = 3
            Case 15
                cboSelectTime(1).ListIndex = 4
            Case 30
                cboSelectTime(1).ListIndex = 5
            Case -1
                cboSelectTime(1).ListIndex = 6
        End Select
        mintDelIndex = cboSelectTime(1).ListIndex
    End If

    cboSelectTime(2).Clear         '审核与上报工作
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
         .AddItem "30天内"
        .ItemData(.NewIndex) = 30
        .AddItem "[指定...]"
        .ItemData(.NewIndex) = -1
    End With
    If cboSelectTime(2).ListCount > 0 Then
        Select Case mintDates
            Case 0
                cboSelectTime(2).ListIndex = 0
            Case 1
                cboSelectTime(2).ListIndex = 1
            Case 2
                cboSelectTime(2).ListIndex = 2
            Case 7
                cboSelectTime(2).ListIndex = 3
            Case 15
                cboSelectTime(2).ListIndex = 4
            Case 30
                cboSelectTime(2).ListIndex = 5
            Case -1
                cboSelectTime(2).ListIndex = 6
        End Select
        mintIndex = cboSelectTime(2).ListIndex
    End If
End Sub

Private Sub cboSelectTime_Click(Index As Integer)
'参数：Index 0未填写 1已删除 2审核与上报工作
    Dim intOldIndex As Integer

    If Index = 0 Then
        intOldIndex = mintWaitIndex
        mintWaitIndex = cboSelectTime(Index).ListIndex
        mintWaitDays = cboSelectTime(Index).ItemData(cboSelectTime(Index).ListIndex)
        If mintWaitIndex = intOldIndex And mintWaitDays <> -1 Then Exit Sub
        If mintWaitDays = -1 Then
            If Me.Visible Then
                If Not frmSelectTime.ShowMe(Me, mdtWaitBegin, mdtWaitEnd, cboSelectTime(Index)) Then
                    Call gobjComlib.zlControl.CboSetIndex(cboSelectTime(Index).hwnd, intOldIndex)
                    Exit Sub
                End If
            End If
        End If
        If mdtWaitBegin = CDate(0) Or mdtWaitEnd = CDate(0) Then
            cboSelectTime(Index).ToolTipText = ""
        Else
            cboSelectTime(Index).ToolTipText = "范围：" & Format(mdtWaitBegin, "yyyy-MM-dd") & " 至 " & Format(mdtWaitEnd, "yyyy-MM-dd")
        End If
        If Me.Visible = True Then Call zlRefList
    ElseIf Index = 1 Then
        intOldIndex = mintDelIndex
        mintDelIndex = cboSelectTime(Index).ListIndex
        mintDelDays = cboSelectTime(Index).ItemData(cboSelectTime(Index).ListIndex)

        If intOldIndex = mintDelIndex And mintDelDays <> -1 Then Exit Sub
        If mintDelDays = -1 Then
            If Me.Visible Then
                If Not frmSelectTime.ShowMe(Me, mdtDelBegin, mdtDelEnd, cboSelectTime(Index)) Then
                    '取消时恢复原来的选择
                    Call gobjComlib.zlControl.CboSetIndex(cboSelectTime(Index).hwnd, intOldIndex)
                    Exit Sub
                End If
            End If
        End If
        If mdtDelBegin = CDate(0) Or mdtDelEnd = CDate(0) Then
            cboSelectTime(Index).ToolTipText = ""
        Else
            cboSelectTime(Index).ToolTipText = "范围：" & Format(mdtDelBegin, "yyyy-MM-dd") & " 至 " & Format(mdtDelEnd, "yyyy-MM-dd")
        End If
        If Me.Visible = True Then Call zlRefList
     ElseIf Index = 2 Then
        intOldIndex = mintIndex
        mintIndex = cboSelectTime(Index).ListIndex
        mintDates = cboSelectTime(Index).ItemData(cboSelectTime(Index).ListIndex)

        If intOldIndex = mintIndex And mintDates <> -1 Then Exit Sub
        If mintDates = -1 Then
            If Me.Visible Then
                If Not frmSelectTime.ShowMe(Me, mdtFrom, mdtTo, cboSelectTime(Index)) Then
                    '取消时恢复原来的选择
                    Call gobjComlib.zlControl.CboSetIndex(cboSelectTime(Index).hwnd, intOldIndex)
                    Exit Sub
                End If
            End If
        End If

        If mdtFrom = CDate(0) Or mdtTo = CDate(0) Then
            cboSelectTime(Index).ToolTipText = ""
        Else
            cboSelectTime(Index).ToolTipText = "范围：" & Format(mdtFrom, "yyyy-MM-dd") & " 至 " & Format(mdtTo, "yyyy-MM-dd")
        End If
        If Me.Visible = True Then Call zlRefList
    End If
End Sub

Private Sub cmdDuplicateCheck_Click()
'功能：显示查重的所有报告
    Dim strName As String
    Dim strProfession As String
    Dim strDiagnose As String
    Dim rsOld As ADODB.Recordset

On Error GoTo ErrHand
    strName = Trim(txtName.Text)
    strProfession = Trim(txtProfession.Text)
    strDiagnose = Trim(txtDiagnose.Text)

    If strName = "" Then
        lblNotice.Visible = True
        txtName.SetFocus
        Exit Sub
    End If

    Call ZlRefDuplicateReport(rsOld, strName, strProfession, strDiagnose)
    Call SetDuplicateReportData(rsOld)

    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ZlRefDuplicateReport(ByRef rsOld As ADODB.Recordset, ByVal strName As String, Optional ByVal strProfession As String, Optional ByVal strDiagnose As String)
'功能：报告查重
'参数：rsOld 老版电子病历查询出来的数据集
'      strName 要查寻的人的姓名
'      strProfession  要查寻的人的职业
'      strDiagnose    要查寻的人的诊断
    Dim strSQL As String
    Dim blnMoved As Boolean
    Dim strTemp As String

On Error GoTo ErrHand
    If strProfession <> "" Then
        strSQL = " And a.职业 = [2] "
    End If
    If strDiagnose <> "" Then
        strSQL = strSQL & " And (a.诊断描述1 Like [3] Or a.诊断描述2 Like [3]) "
    End If
    blnMoved = MovedByDate(DateAdd("m", -12, CDate(gobjComlib.zlDatabase.Currentdate)))

    strSQL = "Select a.文件id as ID, l.病历名称 As 报告, Null As 就诊号, a.处理状态 As 状态, a.报送人, a.报送时间, a.报送单位, a.报送备注," & vbNewLine & _
            "       a.登记人, a.登记时间, a.姓名, a.性别, a.年龄, a.职业, a.家庭地址, a.家庭电话, a.发病日期, a.确诊日期, a.诊断描述1, a.诊断描述2, a.填报备注, a.报卡类型," & vbNewLine & _
            "       a.报告医生, l.保存人 As 填报人, l.保存时间 As 填报时间, l.病人来源, b.名称 As 科室, NULL As 疑似疾病, Null As 就诊号, l.编辑方式 As 编辑方式, a.撤档人, a.撤档时间," & vbNewLine & _
            "       l.病人id, l.主页id, l.文件ID, 0 As 数据转出" & vbNewLine & _
            "From 疾病申报记录 A, 电子病历记录 L, 部门表 B" & vbNewLine & _
            "Where a.撤档人 Is Null And a.撤档时间 Is Null And a.文件id = l.Id And l.病历种类 = 5 and a.处理状态= 3 And b.Id = l.科室id " & vbNewLine & _
            " and a.姓名 Like [1] " & strSQL

    If blnMoved Then
        strTemp = Replace(strSQL, "0 As 数据转出", "1 as 数据转出")
        strTemp = Replace(strTemp, "电子病历记录", "H电子病历记录")
        strTemp = Replace(strTemp, "疾病申报记录", "H疾病申报记录")
        strSQL = strSQL & " Union All " & strTemp
    End If

    Set rsOld = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strName & "%", strProfession, "%" & strDiagnose & "%")

    ZlRefDuplicateReport = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SetDuplicateReportData(ByVal rsOld As ADODB.Recordset)
'功能：报告查重显示
'参数：rsOld 老版电子病历查询出来的数据集
  On Error GoTo ErrHand
    Me.rptList.Records.DeleteAll
    dkpMain.FindPane(conPane_Feedback).Close
    Call SetRptListData(rsOld)
    Me.rptList.Populate

    If Me.rptList.Rows.Count > 0 Then
        If Me.rptList.FocusedRow Is Nothing Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        Else
            rptList_SelectionChanged
        End If
    Else
        mstrCurId = ""
        rptList_SelectionChanged
    End If

    If Me.rptList.Records.Count > 0 Then
        Me.stbThis.Panels(2).Text = "共查找到" & Me.rptList.Records.Count & "份符合条件疾病报告。"
    Else
        Me.stbThis.Panels(2).Text = "没有查找到符合条件的疾病报告。"
    End If
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtName_Change()
    If Trim(txtName.Text) <> "" Then
        lblNotice.Visible = False
    End If
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    If Trim(txtName.Text) = "" Then
        lblNotice.Visible = True
    End If
End Sub

Private Sub txtName_GotFocus()
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 20
    Call gobjComlib.ZLCommFun.OpenIme(True)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdDuplicateCheck.SetFocus
        Call cmdDuplicateCheck_Click
        Exit Sub
    End If
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtName_LostFocus()
    Me.txtName.Text = Trim(Me.txtName)
    Call gobjComlib.ZLCommFun.OpenIme(False)
End Sub

Private Sub txtProfession_GotFocus()
    Me.txtProfession.SelStart = 0: Me.txtProfession.SelLength = 20
    Call gobjComlib.ZLCommFun.OpenIme(True)
End Sub

Private Sub txtProfession_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdDuplicateCheck.SetFocus
        Call cmdDuplicateCheck_Click
        Exit Sub
    End If
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtProfession_LostFocus()
    Me.txtProfession.Text = Trim(Me.txtProfession)
    Call gobjComlib.ZLCommFun.OpenIme(False)
End Sub

Private Sub txtDiagnose_GotFocus()
    Me.txtDiagnose.SelStart = 0: Me.txtDiagnose.SelLength = 100
    Call gobjComlib.ZLCommFun.OpenIme(True)
End Sub

Private Sub txtDiagnose_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdDuplicateCheck.SetFocus
        Call cmdDuplicateCheck_Click
    End If
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Public Function ShowDiseaseStation(ByVal frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
                ByVal intPatiFrom As Integer, ByVal lng科室ID As Long, ByVal str疾病ID As String, ByVal str诊断ID As String, ByRef blnNotView As Boolean) As Boolean
'功能：查询指定人员一年内是否填写过传染病报告卡
'参数：lng病人ID    病人ID

'      lng主页ID    住院为 主页ID，门诊为 挂号ID
'      intPatiFrom  病人来源 住院为 2， 门诊为 1
'      lng科室ID    科室 ID
'      str疾病ID    疾病ID
'      str诊断ID    诊断ID
    Dim blnMoved As Boolean, strTemp As String
    Dim strSQL As String
    Dim blnHasData As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim str主页ids As String
    Dim blnDiag As Boolean
    Dim vMsg As String
    On Error GoTo errHand

    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mIntPatiFrom = intPatiFrom
    mlng科室ID = lng科室ID
    mstr疾病ID = str疾病ID
    mstr诊断ID = str诊断ID

    If mstr诊断ID = "" And mstr疾病ID = "" Then Exit Function

    mblnReportCheck = True
    blnMoved = MovedByDate(DateAdd("m", -12, CDate(gobjComlib.zlDatabase.Currentdate)))
    

    strSQL = "Select distinct l.Id, l.文件id, l.病人id, l.主页id, l.病历名称 As 报告, l.病人来源,  d.名称 As 科室," & vbNewLine & _
            " Decode(l.病人来源, 1, p.门诊号, 2, p.住院号) As 就诊号, Nvl(l.姓名, p.姓名) As 姓名, Nvl(l.性别, p.性别) As 性别, Nvl(l.年龄, p.年龄) As 年龄," & vbNewLine & _
            " To_Char(l.完成时间, 'yyyy-mm-dd hh24:mi') As 填报时间, l.保存人 As 填报人, l.报卡类型, l.编辑方式," & vbNewLine & _
            " l.状态 As 状态,l.报送人," & vbNewLine & _
            " To_Char(l.报送时间, 'yyyy-mm-dd hh24:mi') as 报送时间, l.报送单位,l.报送备注,l.登记人," & vbNewLine & _
            " To_Char(l.登记时间, 'yyyy-mm-dd hh24:mi') as 登记时间, Nvl(l.职业,P.职业) as 职业 , Nvl(l.家庭地址,P.家庭地址) as 家庭地址, Nvl(l.家庭电话,P.家庭电话) as 家庭电话 ," & vbNewLine & _
            " To_Char(l.发病日期, 'yyyy-mm-dd') as 发病日期,To_Char(l.确诊日期, 'yyyy-mm-dd') as 确诊日期 , l.诊断描述1 ,l.诊断描述2," & vbNewLine & _
            " l.填报备注 ,L.撤档人,L.撤档时间 , Q.传染病名称 as  疑似疾病,0 As 数据转出" & vbNewLine & _
            " From (Select l.Id, l.文件id, l.病人id, l.主页id, l.病历名称, l.病人来源, l.科室id, l.完成时间, l.保存人, l.保存时间, l.编辑方式, decode(S.撤档人, null ,Nvl(s.处理状态, 0),7) As 状态," & vbNewLine & _
            "      s.报送人, s.报送时间, s.报送单位, s.报送备注, s.登记人, s.登记时间, s.姓名, s.性别, s.年龄, s.职业, s.家庭地址, s.家庭电话," & vbNewLine & _
            "        s.发病日期, s.确诊日期, s.诊断描述1, s.诊断描述2, s.填报备注, s.报卡类型,s.撤档人,S.撤档时间 " & vbNewLine & _
            "       From 电子病历记录 L, 疾病申报记录 S " & vbNewLine & _
            "       Where l.Id = s.文件id(+) And l.病历种类 = 5 And l.病人ID = [1] And trunc(l.完成时间) >=trunc(ADD_MONTHS(sysdate,-12)) " & vbNewLine & _
            " and S.撤档人 Is Null And S.撤档时间 Is Null ) L, 病人信息 P, 部门表 D,疾病阳性记录 Q " & vbNewLine & _
            "Where l.病人id = p.病人id And l.科室id = d.Id and l.id =Q.文件ID(+) and l.病人id =Q.病人ID(+)"
    If blnMoved Then
        strTemp = Replace(strSQL, "0 as 数据转出", "1 as 数据转出")
        strTemp = Replace(strTemp, "电子病历记录", "H电子病历记录")
        strTemp = Replace(strTemp, "疾病申报记录", "H疾病申报记录")
        strTemp = Replace(strTemp, "疾病阳性记录", "H疾病阳性记录")
        strSQL = strSQL & " Union All " & strTemp
    End If
    Set mrsOld = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "查询该病人一年内的传染病报告卡", lng病人ID)
    blnDiag = True
    If mrsOld.RecordCount > 0 Then
        '查询该病人一年内不同的诊断
        Do While Not mrsOld.EOF
            If mrsOld!主页id & "" <> "" Then str主页ids = str主页ids & "," & mrsOld!主页id & ""
            mrsOld.MoveNext
        Loop
		mrsOld.MoveFirst 
        If str主页ids <> "" Then
            str主页ids = Mid(str主页ids, 2)
            If str疾病ID <> "" Then
                strSQL = " Union Select 文件ID,疾病id,诊断id From 疾病报告前提 Where 疾病ID IN (Select Column_Value From Table(f_Num2list([2])))"
            End If
            If str诊断ID <> "" Then
                strSQL = strSQL & " Union Select 文件ID,疾病id,诊断id From 疾病报告前提 Where 诊断ID IN (Select Column_Value From Table(f_Num2list([3])))"
            End If
            strSQL = "(" & Mid(strSQL, 8) & ")"
            strSQL = "Select /*+ Rule*/ b.疾病id,b.诊断id,0 as 数据转出" & vbNewLine & _
                    "From 病历文件列表 A ,(" & strSQL & ") B Where A.ID=B.文件ID  And" & vbNewLine & _
                    "(a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 病历应用科室 C Where c.文件id = a.Id And c.科室id = [4]))"
            strSQL = "(" & strSQL & ") Minus Select A.疾病id, A.诊断id,0 as 数据转出 From 病人诊断记录 A" & vbNewLine & _
                    "Where a.病人id = [1] and a.主页id in (Select Column_Value From Table(f_Num2list([5]))) And Trunc(a.记录日期) >= Trunc(Add_Months(Sysdate, -12)) And a.编码序号 = 1"
        
            If blnMoved Then
                strTemp = Replace(strSQL, "0 as 数据转出", "1 as 数据转出")
                strTemp = Replace(strTemp, "病人诊断记录", "H病人诊断记录")
                strSQL = strSQL & " Union All " & strTemp
            End If
        
            Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "查询该病人一年内的传染病诊断", lng病人ID, str疾病ID, str诊断ID, lng科室ID, str主页ids)
            blnDiag = rsTmp.RecordCount = 0
        End If
    End If
    
    If mrsOld.RecordCount > 0 And blnDiag Then
        blnHasData = True
        ShowDiseaseStation = True
        vMsg = zlCommFun.ShowMsgBox("传染病报告卡提示", "需要填写传染病报告卡，但是该病人在过去的一年内已经填写过传染病报告卡，是否查看填写过的报告卡？", "!查看(&R),重填(&W),?忽略(&Q)", frmParent, vbQuestion)
        If vMsg = "重填" Then
            blnNotView = True
            Exit Function
        ElseIf vMsg = "" Then
            Exit Function
        End If
    Else
        blnHasData = False
        ShowDiseaseStation = False
        mblnReportCheck = False
        Exit Function
    End If
    Me.Show 1, frmParent
    mblnReportCheck = False

    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function WriteReport() As Boolean
'功能：填写报告
    Dim clsDisease As New cDockDisease
    WriteReport = clsDisease.EditDiseaseDoc(Me, mlng病人ID, mlng主页ID, mIntPatiFrom, mlng科室ID, mstr疾病ID, mstr诊断ID)
    Set clsDisease = Nothing
End Function

Private Function GetFeedbackIDs(ByVal lngID As Long) As Boolean
'功能：查询选中的报告的阳性反馈单ID
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strDis As String
    Dim i As Long
On Error GoTo ErrHand
    If lngID = 0 Then Exit Function
    
    strSQL = "select A.ID,A.登记时间,B.名称 as 科室,A.传染病名称 from  疾病阳性记录 A, 部门表 B where A.文件ID = [1] and a.登记科室id = B.ID "
    Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "查询该报告对应的阳性反馈单", lngID)
    
    If rsTemp.RecordCount > 0 Then
         With Me.tbcDis
            .removeAll
            For i = 0 To rsTemp.RecordCount - 1
                If Not InStr(strDis, rsTemp!传染病名称 & "") > 0 Then
                    strDis = ";" & rsTemp!传染病名称 & strDis
                End If
                 .InsertItem(i, rsTemp!科室 & "(" & rsTemp!登记时间 & ")", mfrmPreFeedBack.hwnd, 0).Tag = rsTemp!ID
                rsTemp.MoveNext
            Next
            strDis = Mid(strDis, 2)
            .InsertItem(i, "i", mfrmPreFeedBack.hwnd, 0).Tag = i
            .Item(1).Selected = True
            .Item(0).Selected = True
            .Item(i).Visible = False
            txtInfo(txt疑似疾病).Text = strDis
             If Not rptList.FocusedRow Is Nothing Then
                If Not rptList.FocusedRow.GroupRow Then
                    rptList.FocusedRow.Record.Item(mCol.疑似疾病).Value = strDis
                End If
             End If
        End With
        tbcMain.PaintManager.ButtonMargin.SetRect 0, 0, 0, 0
        tbcMain.Item(1).Visible = True
    Else
        tbcMain.PaintManager.ButtonMargin.SetRect -20, -20, 0, 0
        tbcMain.Item(1).Visible = False
    End If
    
     Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SendMsg()
'功能：发送消息 传染病报告返修
    Dim strSQL As String, strXML As String
    Dim rsPati As ADODB.Recordset
    Dim lng病人ID As Long, lng就诊ID As Long
    Dim strTmp As String
    Dim lng病人来源 As Long
    On Error GoTo errH

    lng病人ID = Val(rptList.FocusedRow.Record.Item(mCol.病人ID).Value)
    lng就诊ID = Val(rptList.FocusedRow.Record.Item(mCol.主页ID).Value)
    lng病人来源 = IIf(rptList.FocusedRow.Record.Item(mCol.来源).Value = "住院", 2, 1)
    strTmp = rptList.FocusedRow.Record.Item(mCol.科室).Value

    If lng病人来源 = 1 Then
        strSQL = "select a.姓名,null as 住院号,a.门诊号,null as 病区ID,a.执行部门id as 科室ID,null as 床号,a.诊室 from 病人挂号记录 a where a.ID=[2]"
    ElseIf lng病人来源 = 2 Then
        strSQL = "select a.姓名,a.住院号,null as 门诊号,a.当前病区id as 病区ID,a.出院科室id as 科室ID,a.出院病床 as 床号,null as 诊室 from 病案主页 a where a.病人ID=[1] and a.主页id=[2]"
    End If
    Set rsPati = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng就诊ID)
    strXML = "<patient_info><patient_id>" & lng病人ID & "</patient_id><patient_name>" & rsPati!姓名 & "</patient_name>"
    strXML = strXML & "<in_number>" & rsPati!住院号 & "</in_number>"
    strXML = strXML & "<out_number>" & rsPati!门诊号 & "</out_number>"
    strXML = strXML & "</patient_info><patient_clinic><patient_source>" & lng病人来源 & "</patient_source>"
    strXML = strXML & "<clinic_id>" & lng就诊ID & "</clinic_id>"
    strXML = strXML & "<clinic_area_id>" & rsPati!病区ID & "</clinic_area_id>"
    strTmp = "" & gobjComlib.sys.RowValue("部门表", Val("" & rsPati!病区ID), "名称")
    strXML = strXML & "<clinic_area_title>" & strTmp & "</clinic_area_title>"
    strXML = strXML & "<clinic_dept_id>" & rsPati!科室ID & "</clinic_dept_id>"
    strTmp = "" & gobjComlib.sys.RowValue("部门表", Val("" & rsPati!科室ID), "名称")
    strXML = strXML & "<clinic_dept_title>" & strTmp & "</clinic_dept_title>"
    strXML = strXML & "<clinic_room>" & rsPati!诊室 & "</clinic_room>"
    strXML = strXML & "<clinic_bed>" & rsPati!床号 & "</clinic_bed>"
    strXML = strXML & "</patient_clinic><disease_report><file_id>" & mstrCurId & "</file_id><doc_id>" & mstrCurId & "</doc_id>"
    strXML = strXML & "<report_name>" & rptList.FocusedRow.Record.Item(mCol.报告).Value & "</report_name>"
    strXML = strXML & "<create_time>  </create_time>"   '接收时间，为空
    strXML = strXML & "<create_doctor>" & UserInfo.姓名 & "</create_doctor>"
    strXML = strXML & "</disease_report>"
    Call gobjComlib.zlDatabase.SendMsg("ZLHIS_CIS_033", strXML)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckUntread() As Boolean
'功能：检查上报报告是否可以撤销上报
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngID  As Long
On Error GoTo ErrHand
    If Val(mstrCurId) = 0 Then Exit Function
    If IsNumeric(mstrCurId) Then
        lngID = CLng(Val(mstrCurId))
        strSQL = "select count(1) as num from 疾病申报反馈 where 申报ID = [1] "
        Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "查询该报告对应的阳性反馈单", lngID)
        
        If rsTemp.RecordCount > 0 Then
            CheckUntread = IIf(rsTemp!num > 0, False, True)
            Exit Function
        End If
        CheckUntread = True
    End If
    
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub EditSendInfo(ByVal intType As Integer)
'功能：修改或者新增上报备注说明
'参数： intType：1-新增；2-修改
    If Val(mstrCurId) <> 0 And IsNumeric(mstrCurId) Then
        If intType = 1 Then
            Call frmSendInfo.ShowMe(Me, intType, CLng(Val(mstrCurId)))
            Call SetFeedbackContent(mstrCurId, mIntState, 2)
        ElseIf intType = 2 And mdatTime <> CDate(0) Then
            Call frmSendInfo.ShowMe(Me, intType, CLng(Val(mstrCurId)), mdatTime)
            Call SetFeedbackContent(mstrCurId, mIntState, 2)
        End If
    End If
End Sub
