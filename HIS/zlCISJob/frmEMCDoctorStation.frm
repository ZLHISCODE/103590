VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJOCK.COMMANDBARS.UNICODE.9600.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJOCK.DOCKINGPANE.UNICODE.9600.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "CODEJOCK.REPORTCONTROL.UNICODE.9600.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJOCK.SUITECTRLS.9600.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "*\A..\ZLIDKIND\zlIDKind.vbp"
Begin VB.Form frmEMCDoctorStation 
   BackColor       =   &H00FFFFFF&
   Caption         =   "����ҽ������վ"
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13545
   Icon            =   "frmEMCDoctorStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   13545
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picCharge 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   66
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picMore 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFFEFE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   3720
      ScaleHeight     =   1575
      ScaleWidth      =   11190
      TabIndex        =   60
      Top             =   1920
      Visible         =   0   'False
      Width           =   11190
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFFEFE&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   10
         Left            =   840
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   64
         Top             =   360
         Width           =   6870
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFFEFE&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   7
         Left            =   900
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   61
         Top             =   30
         Width           =   9630
      End
      Begin zl9CISJob.UCPatiVitalSigns UCPatiVitalSigns 
         Height          =   285
         Left            =   90
         TabIndex        =   62
         Top             =   810
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   503
         ControlLock     =   -1  'True
         TextBackColor   =   -2147483633
         LblBackColor    =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BackColor       =   15728382
         ShowMode        =   0
         Style           =   1
         XDis            =   100
         YDis            =   200
         LabToTxt        =   -90
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ϣ:"
         Height          =   180
         Index           =   10
         Left            =   0
         TabIndex        =   65
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ:"
         Height          =   180
         Index           =   7
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Width           =   450
      End
   End
   Begin VB.Timer timRefresh 
      Interval        =   1000
      Left            =   2055
      Top             =   60
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
      TabIndex        =   57
      Top             =   510
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picYZ 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFFEFE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   1920
      ScaleHeight     =   1170
      ScaleWidth      =   1380
      TabIndex        =   52
      Top             =   5595
      Width           =   1380
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   675
         Index           =   3
         Left            =   180
         TabIndex        =   53
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
      Begin VB.ComboBox cboSelectTime 
         Height          =   300
         Left            =   795
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   15
         Width           =   1230
      End
      Begin VB.CommandButton cmdOtherFilter 
         Caption         =   "��������"
         Height          =   300
         Left            =   2400
         TabIndex        =   54
         Top             =   0
         Width           =   1100
      End
      Begin VB.Label lblSeeTim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   30
         TabIndex        =   56
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFFEFE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4635
      Left            =   0
      ScaleHeight     =   4635
      ScaleWidth      =   1845
      TabIndex        =   45
      Top             =   1860
      Visible         =   0   'False
      Width           =   1845
      Begin VB.Frame fraPatiUD 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   150
         MousePointer    =   7  'Size N S
         TabIndex        =   48
         Top             =   3105
         Width           =   6975
      End
      Begin VB.PictureBox picFind 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   0
         ScaleHeight     =   270
         ScaleWidth      =   495
         TabIndex        =   46
         Top             =   0
         Width           =   495
         Begin VB.Label lblFind 
            Caption         =   "����:"
            Height          =   255
            Left            =   60
            TabIndex        =   47
            Top             =   30
            Width           =   495
         End
      End
      Begin XtremeSuiteControls.TabControl tbcWait 
         Height          =   435
         Left            =   570
         TabIndex        =   49
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
         TabIndex        =   50
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
         TabIndex        =   51
         Top             =   0
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmEMCDoctorStation.frx":08CA
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         DefaultCardType =   "���￨"
         IDKindWidth     =   555
         FindPatiShowName=   0   'False
         HiddenMoseRightKey=   0   'False
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
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
      TabIndex        =   44
      Top             =   435
      Visible         =   0   'False
      Width           =   240
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
      TabIndex        =   43
      Top             =   450
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   360
      ScaleHeight     =   915
      ScaleWidth      =   1335
      TabIndex        =   41
      Top             =   5835
      Width           =   1335
      Begin XtremeReportControl.ReportControl rptNotify 
         Height          =   675
         Left            =   0
         TabIndex        =   42
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
   Begin VB.PictureBox picJZ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   1425
      ScaleHeight     =   885
      ScaleWidth      =   705
      TabIndex        =   39
      Top             =   4440
      Width           =   705
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   675
         Index           =   1
         Left            =   0
         TabIndex        =   40
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
   Begin VB.PictureBox picHZ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   1245
      ScaleHeight     =   870
      ScaleWidth      =   675
      TabIndex        =   37
      Top             =   2595
      Width           =   675
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   675
         Index           =   0
         Left            =   0
         TabIndex        =   38
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
      TabIndex        =   35
      Top             =   4830
      Width           =   690
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   675
         Index           =   2
         Left            =   0
         TabIndex        =   36
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
   Begin VB.PictureBox picYy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   3960
      ScaleHeight     =   825
      ScaleWidth      =   690
      TabIndex        =   33
      Top             =   6195
      Width           =   690
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   675
         Index           =   4
         Left            =   0
         TabIndex        =   34
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
   Begin VB.PictureBox picRegist 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFFEFE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2460
      Left            =   3720
      ScaleHeight     =   2460
      ScaleWidth      =   5265
      TabIndex        =   31
      Top             =   4395
      Width           =   5265
      Begin XtremeSuiteControls.TabControl tbcSub 
         Height          =   1875
         Left            =   240
         TabIndex        =   32
         Top             =   75
         Width           =   2580
         _Version        =   589884
         _ExtentX        =   4551
         _ExtentY        =   3307
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picBasisNew 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFFEFE&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   1275
      Left            =   2160
      ScaleHeight     =   1275
      ScaleWidth      =   12225
      TabIndex        =   2
      Top             =   360
      Width           =   12225
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00EFFEFE&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   5
         Left            =   7905
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "456455885"
         Top             =   705
         Width           =   1080
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00EFFEFE&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   4
         Left            =   5685
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "65546578"
         Top             =   705
         Width           =   1080
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EFFEFE&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   3
         Left            =   5715
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "1988-11-11"
         Top             =   150
         Width           =   1080
      End
      Begin VB.Frame fraPayType 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   8325
         TabIndex        =   12
         Top             =   180
         Width           =   1860
         Begin VB.ComboBox cboPayType 
            BackColor       =   &H00EFFEFE&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   -30
            Width           =   1845
         End
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EFFEFE&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   2
         Left            =   3795
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "27��"
         Top             =   165
         Width           =   720
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EFFEFE&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   1
         Left            =   3240
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "��"
         Top             =   165
         Width           =   465
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00EFFEFE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
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
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "����"
         Top             =   120
         Width           =   1620
      End
      Begin VB.PictureBox picPatient 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   30
         ScaleHeight     =   780
         ScaleWidth      =   1050
         TabIndex        =   8
         Top             =   195
         Width           =   1050
         Begin VB.Image imgPatient 
            Height          =   705
            Left            =   30
            Picture         =   "frmEMCDoctorStation.frx":09A7
            Stretch         =   -1  'True
            Top             =   15
            Width           =   975
         End
      End
      Begin VB.Frame fraBillType 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   9990
         TabIndex        =   6
         Top             =   750
         Width           =   1860
         Begin VB.ComboBox cboBillType 
            BackColor       =   &H00EFFEFE&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   -30
            Width           =   1110
         End
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00EFFEFE&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   8
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   720
         Width           =   1680
      End
      Begin VB.TextBox txtPhone 
         BackColor       =   &H00EFFEFE&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   6960
         MaxLength       =   20
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "13800138000"
         Top             =   960
         Width           =   1440
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00EFFEFE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   9
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������:"
         Height          =   180
         Index           =   5
         Left            =   6960
         TabIndex        =   30
         Top             =   705
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���￨��:"
         Height          =   180
         Index           =   4
         Left            =   4560
         TabIndex        =   29
         Top             =   705
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ѷ�ʽ:"
         Height          =   180
         Index           =   11
         Left            =   7020
         TabIndex        =   28
         Top             =   120
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   180
         Index           =   3
         Left            =   4605
         TabIndex        =   27
         Top             =   150
         Width           =   810
      End
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�޸Ļ�����Ϣ"
         BeginProperty Font 
            Name            =   "����"
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
         MouseIcon       =   "frmEMCDoctorStation.frx":1871
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   990
         Width           =   1080
      End
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����"
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
         MouseIcon       =   "frmEMCDoctorStation.frx":31F3
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   990
         Width           =   360
      End
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ɼ�"
         BeginProperty Font 
            Name            =   "����"
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
         MouseIcon       =   "frmEMCDoctorStation.frx":4B75
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   465
         Width           =   360
      End
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ļ�"
         BeginProperty Font 
            Name            =   "����"
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
         MouseIcon       =   "frmEMCDoctorStation.frx":64F7
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   90
         Width           =   360
      End
      Begin VB.Line linPayType 
         X1              =   8400
         X2              =   9780
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label lblMore 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "���ն�ƾ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9915
         TabIndex        =   22
         Top             =   660
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�:"
         Height          =   180
         Index           =   6
         Left            =   9060
         TabIndex        =   21
         Top             =   870
         Width           =   450
      End
      Begin VB.Line linBillType 
         X1              =   9840
         X2              =   10620
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         Height          =   180
         Index           =   8
         Left            =   2040
         TabIndex        =   20
         Top             =   705
         Width           =   450
      End
      Begin VB.Label lblPhone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ֻ���:"
         Height          =   180
         Left            =   6240
         TabIndex        =   19
         Top             =   960
         Width           =   630
      End
      Begin VB.Line LinPhone 
         X1              =   6960
         X2              =   8340
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblPhysical 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���������:"
         Height          =   180
         Left            =   3360
         TabIndex        =   18
         Top             =   960
         Width           =   990
      End
      Begin VB.Label lblRec 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   390
         Left            =   10800
         TabIndex        =   17
         Top             =   90
         Width           =   405
      End
   End
   Begin VB.Frame fraRoom 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9240
      TabIndex        =   0
      Top             =   5595
      Width           =   300
      Begin VB.Label lblRoom 
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   300
      End
   End
   Begin MSComctlLib.ImageList imgPati 
      Left            =   2880
      Top             =   0
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
            Picture         =   "frmEMCDoctorStation.frx":7E79
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEMCDoctorStation.frx":8413
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEMCDoctorStation.frx":89AD
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEMCDoctorStation.frx":8F47
            Key             =   "ת��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEMCDoctorStation.frx":94E1
            Key             =   "�ܾ�"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEMCDoctorStation.frx":9A7B
            Key             =   "��ͣ"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEMCDoctorStation.frx":A015
            Key             =   "����"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   3870
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeSuiteControls.TabControl tbcRegist 
      Height          =   915
      Left            =   3720
      TabIndex        =   58
      Top             =   3435
      Width           =   6015
      _Version        =   589884
      _ExtentX        =   10610
      _ExtentY        =   1614
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   59
      Top             =   6630
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEMCDoctorStation.frx":A367
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16960
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "�������20��"
            TextSave        =   "�������20��"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   1843
            MinWidth        =   1843
            Text            =   "������"
            TextSave        =   "������"
            Object.ToolTipText     =   "����״̬(�����������)"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgDefual 
      Height          =   705
      Left            =   1680
      Picture         =   "frmEMCDoctorStation.frx":ABF9
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLoad 
      Height          =   705
      Left            =   480
      Picture         =   "frmEMCDoctorStation.frx":BAC3
      Stretch         =   -1  'True
      Top             =   885
      Visible         =   0   'False
      Width           =   975
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   75
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmEMCDoctorStation.frx":C98D
      Left            =   960
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
      ScaleMode       =   1
   End
End
Attribute VB_Name = "frmEMCDoctorStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Private Const COLOR_FREE As Long = &HC000&
Private Const COLOR_BUSY As Long = &HFF&
Private Const COLOR_RPTSelRow = &H4040&
Private Const COLOR_RPTHeadBack = &HC0E0F0
Private Const COLOR_RPTHeadBackSel = &HAEEAEA
Private Const COLOR_Back = &HEFFEFE

Private Enum PatiType
    pt���� = 0
    pt���� = 1
    pt���� = 2
    ptת�� = 3
    ptԤԼ = 4
    pt���� = 5
    pt�Ŷӽк� = 6
End Enum

Private Enum PATI_RPT_LIST
    PATI_RPT���� = 0
    PATI_RPT���� = 1
    PATI_RPT���� = 2
    PATI_RPT���� = 3
    PATI_RPTԤԼ = 4
End Enum

Private Enum m_Ctl_ID    'һ��Ҫ�������
    txtInfo���� = 0
    txtInfo�Ա� = 1
    txtInfo���� = 2
    txtInfo�������� = 3
    txtInfo���￨�� = 4
    txtInfoҽ������ = 5
    txtInfo�ѱ� = 6
    txtInfoժҪ = 7
    txtInfo���� = 8
    txtInfo������ = 9
    txtInfo������Ϣ = 10
    txtInfo���ѷ�ʽ = 11
    
    lblLink�ļ� = 0
    lblLink�ɼ� = 1
    lblLink��� = 2
    lblLink�޸� = 3
    
    '����:3�ˣ����:45�ˣ�����:15��
    t���� = 0
    t���� = 1
    t���� = 2
End Enum

Private Enum PATI_COL_����
    COL_HZ_��ʶ = 0
    COL_HZ_����
    COL_HZ_�����
    COL_HZ_����
    COL_HZ_�Һ�ʱ��
    COL_HZ_�Ա�
    COL_HZ_����
    COL_HZ_��ɫͨ��
    COL_HZ_��
    COL_HZ_NO
    
    COL_HZ_��������
    COL_HZ_����ҽ��
    COL_HZ_���
    COL_HZ_����ʱ��
    COL_HZ_��������
    
    COL_HZ_ת��״̬
    COL_HZ_����
    COL_HZ_���˿���
    COL_HZ_ԤԼҽ��
    COL_HZ_ԤԼʱ��
    
'������
    COL_HZ_����ID
    COL_HZ_����ʱ��
    COL_HZ_ִ�в���ID
    COL_HZ_ִ����
    COL_HZ_״̬ 'ת��״̬��־
    COL_HZ_IC����
    COL_HZ_���￨��
    COL_HZ_���֤��
    COL_HZ_��¼��־
    COL_HZ_ִ��״̬
End Enum

Private Enum PATI_COL_���� '�����б�ͻ����б���
    COL_JZ_��ʶ = 0
    COL_JZ_����
    COL_JZ_�����
    COL_JZ_����
    COL_JZ_����ʱ��
    COL_JZ_�Ա�
    COL_JZ_����
    COL_JZ_��ɫͨ��
    COL_JZ_��
    COL_JZ_NO
    
    COL_JZ_���￨��
    COL_JZ_��������
    COL_JZ_ת��״̬
    COL_JZ_��Ⱦ��
    COL_JZ_����
    COL_JZ_���˿���
    
'������
    COL_JZ_����ID
    COL_JZ_����ʱ��
    COL_JZ_ִ�в���ID
    COL_JZ_ִ����
    COL_JZ_״̬ 'ת��״̬��־
    COL_JZ_���֤��
    COL_JZ_IC����
    COL_JZ_��¼��־
End Enum

Private Enum PATI_COL_����
    COL_YZ_��ʶ = 0
    COL_YZ_����
    COL_YZ_�����
    COL_YZ_����
    COL_YZ_����ʱ��
    COL_YZ_�Ա�
    COL_YZ_����
    COL_YZ_��ɫͨ��
    COL_YZ_��
    COL_YZ_NO
    COL_YZ_����ҽ��
    COL_YZ_���￨��
    COL_YZ_��������
    COL_YZ_����
    COL_YZ_���˿���
    COL_YZ_��ҽ���
    COL_YZ_��ҽ���

'������
    COL_YZ_����ID
    COL_YZ_����ʱ��
    COL_YZ_ִ�в���ID
    COL_YZ_ִ����
    COL_YZ_���֤��
    COL_YZ_IC����
    COL_YZ_��¼��־
End Enum

Private Enum Msg_Type '��Ϣ�������
    mΣ��ֵ = 1
    mҽ������ = 2
    m������� = 3
    m��Ⱦ�� = 4
    m��Ѫ��� = 5
    m��Ѫ��� = 6
    m��Ѫ��Ӧ = 6
End Enum
 
Private Enum NOTIFYREPORT_COLUMN
    c_ͼ�� = 0
    C_����ID = 1
    C_No = 2
    C_����� = 4
    c_���� = 3
    C_����ʱ�� = 5
    C_״̬ = 6
    '������
    C_��Ϣ = 7
    C_��� = 8
    C_���� = 9
    C_ҵ�� = 10
    C_�Һ�Id = 11
    C_Id = 12
    C_������Ϣ = 13
End Enum

Private Type PatiInfo
    ���� As PatiType
    ���� As String
    ����� As String
    �Һ�ID As Long
    �Һŵ� As String
    ����ID As Long
    ���� As String
    ���� As Integer
    ������ As String
    �Һ�ʱ�� As Date
    ����ת�� As Boolean
    ����ID As Long
    ������ As String
    �Ƿ�ǩ�� As Boolean
    �Ա� As String
    ����״�� As String
    ���� As String
    ���� As String
    ���� As String
    �����ص� As String
    ��Ⱦ���ϴ� As Long
    ��ͥ��ַ�ʱ� As String
    ��λ�ʱ� As String
    ����֤�� As String
    ���ڵ�ַ As String
    ���ڵ�ַ�ʱ� As String
    ����  As String
    Email As String
    QQ As String
    ���� As Integer
    ���鼶�� As String
    �Ƿ���ɫͨ�� As Integer
End Type

Private Type ty_Queue
    strQueuePrivs As String '�Ŷӽк�����ģ��Ȩ��
    str����վ�� As String     '���е�վ��:��Ϊ��վ��;����Ϊ����վ��
    byt�Ŷӽк�ģʽ As Byte '�ŶӽкŴ���ģʽ:1.�������̨������л�ҽ����������;2-�ȷ������,��ҽ�����о���.0-���Ŷӽк�
    int�������� As Integer  '0-������,>0��ʾ��������
    bln���к����� As Boolean   '�����Ƿ񺬻�������
    blnҽ���������� As Boolean  'true:��ʾҽ����������;False-ҽ������������
    strCurrQueueName As String '��ǰ��������
    lngcurr�Һ�ID As Long '��ǰ�Һ�ID
End Type
Private mty_Queue As ty_Queue

'�����������
Private Type COND_FILTER
    Begin As Date
    End As Date
    ����ID As Long
    ҽ�� As String
    ����ID As Long
    ���� As String
    �ı� As String
End Type
Private mvCondFilter As COND_FILTER

'�Ӵ��������
Private mclsEMR As Object  '�°没��zlRichEMR.clsDockEMR
Private mclsDisease As zlRichEPR.cDockDisease
Private WithEvents mclsDis As zl9Disease.clsDisease
Attribute mclsDis.VB_VarHelpID = -1
Private WithEvents mclsAdvices As zlPublicAdvice.clsDockOutAdvices
Attribute mclsAdvices.VB_VarHelpID = -1

Private WithEvents mclsEPRs As zlRichEPR.cDockOutEPRs
Attribute mclsEPRs.VB_VarHelpID = -1

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
Private mobjMsg As Object '�����������ݶ���
Private mobjPeis As Object '���ӿڲ���
Private mobjDocMsg As Object '��Ϣ����

'�������ñ���
Private mint���ﷶΧ As Integer '1-����,2-������,3-������
Private mlng�������ID As Long
Private mstr�������� As String
Private mstr����ҽ�� As String
Private mstr����ҽ����� As String
Private mblnҪ����� As Boolean
Private mintRefresh As Integer '���ﲡ��ˢ�¼��(s)
Private mbln�Զ����� As Boolean
Private mlng�Զ����� As Long
Private mbln���к���� As Boolean
Private mblnΣ��ֵ���� As Boolean
 
Private mlng������� As Long '0-������ 1-��ֹ 2-��ʾ �����:57566
Private mlng��ǰ����ʱ�� As Long  '����Ҫ��ԤԼ�Ž��ս��п���ʱ,��ֵ����ԤԼ�ſ�����ǰ���յķ����� �����:57566
Private mblnAutoHandle As Boolean '������"����ʱ�Զ�������ɾ���"�����ﲡ��ʱ�Զ�������һ��������ɾ��������
Private mblnUseTYT As Boolean 'ʹ��̫Ԫͨ�ӿ�
Private mint����������Դ As Integer 'ҽ��վ�Ĺ���������Դ
Private mintOutPreTime As Integer
Private mbyt���ξ��� As Byte    '��¼ �����ξ�� ҳǩ�±�ֵ 0-û�о���һ��ҳǩ 1-���ھ���һ��ҳǩ
Private mbln��Һ�ģʽ As Boolean

'---------------�Ŷӽк����
'�����п��ʼ��
Private Const C_STR_QUEUECALL = "0,0,0,0,50,0,90,0,60,0,0,60,60,0,0,60,0,0,125"
'�Ŷ��п��ʼ��
Private Const C_STR_QUEUEQUEUE = "0,0,0,30,50,0,90,40,60,60,0,60,60,50,125,0,120,60,0"

Private Enum mCol
    �������� = 0: ID: ����ID: �Ŷӱ��: �ŶӺ���:  �Ŷ����: ��������: ����: �������: ���������: ����ID: ����: ҽ������: �Ŷ�״̬: �Ŷ�ʱ��: ����ҽ��: ҵ������: ҵ��ID: ����ʱ��: ��������: ORD
End Enum
Private mlngQueueGroupType As Long
Private mstrShowCalledColumnInf As String
Private mstrShowColumnInf As String
Private mlngOrderStyle As Long
Private mlng���ﲡ������ As Long
Private mlngMaxLen As Long
Private mobjQueueList As Object
Private mobjCallList As Object
'------------------

'�����������
Private mrsAller As ADODB.Recordset '���˹�����¼
Private mstrIDCard As String '����Զ�ˢ���������֤��
Private WithEvents mobjIDCard As clsIDCard '���֤����
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object 'IC������
Private mblnUnRefresh As Boolean
Private mstrPrivs As String
Private mlngModul As Long
Private mPatiInfo As PatiInfo '��ʷ�����¼�е�,��һ��Ϊ��ǰ��

'-----�б���ѡ���е���ص���Ϣ��
Private mintActive As PatiType '��������
Private mintRPTIndex As PATI_RPT_LIST 'ѡ�еĲ����б�����ֵ��Ĭ��ֵΪ-1
Private mPr As Long          'ѡ�еĲ����б���ѡ����кţ�Ĭ��ֵΪ-1��ͨ�� mintRPTIndex��mPr���궨��ǰѡ�е��б��У���ʹ��RPT�ؼ��� SelectedRows ����

Private mlng����ID As Long
Private mstr�Һŵ� As String
Private mlng�Һ�ID As Long
Private mlng����ID As Long
 
Private mintFindType As Integer '0-���￨,1-��ʶ�ţ�������ţ�,2-�Һŵ�,3-����,4-�������֤,5-IC��
Private mstrFindType As String '�����洢������ǰ���͵�����"���￨����ʶ�ţ��Һŵ����������������֤��IC��"
Private mblnFindTypeEnabled As Boolean
Private mstr�Һ�IDs As String ' ���˹Һż�¼.ID  �����ŷָ��¼��ǰ�������Щ����
Private mblnIsInit As Boolean 'PatiIdentify��ʼ����־

'ҽ�ƿ�
Private mstrCardKind As String        '��������󷵻صĿ��õ�ҽ�ƿ�

Private mstrPrePati As String
Private mintPreTime As Integer
Private mlngCommunityID As Long '�Զ�ִ�е���������
Private mbytSize As Byte '���� 0-С���壨9�����壩��1-�����壨12�����壩
Private mblnTabTmp As Boolean
Private mstrPreSubTab As String ' tbsSubǰһ��ѡ�е�ҳǩ

Private mblnMsgOk As Boolean '�Ƿ�����Ϣ����
Private mblnFirstMsg As Boolean 'mblnFirstMsg=false ��ʾ��ҽ��վ��ĵ�һ����Ϣ
Private mintNotify As Integer 'ҽ�������Զ�ˢ�¼��(����)
Private mintNotifyDay As Integer '���Ѷ������ڵ�ҽ��
Private mstrNotifyAdvice As String '���ѵ�ҽ������
Private mstrPreNotify As String

Private mclsMsg As clsCISMsg
Private mrsMsg As ADODB.Recordset
Private mbln��Ϣ���� As Boolean
Private mlng���ս������� As Long
Private mblnΣ��ֵ As Boolean '��Σ��ֵ��Ȩ��
Private mbln��ʾԤԼ���� As Boolean
Private mintԤԼ�б� As Integer
Private mblnΣ��ֵshow As Boolean 'Σ��ֵ�Ƿ��


Private Sub cboPayType_Click()
    Dim strTmp As String

    If mstr�Һŵ� = "" Then Exit Sub
    On Error GoTo errH
    strTmp = Split(cboPayType.Text, "-")(1)
    If cboPayType.ToolTipText <> strTmp Then
        strTmp = Split(cboPayType.Text, "-")(1)
        
        '���·ѱ�
        If Update���¹Һŷѱ�(mstr�Һŵ�, strTmp, 0, p����ҽ��վ) = False Then Exit Sub
        
        cboPayType.Tag = ""
        cboPayType.ToolTipText = strTmp
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboBillType_Click()
    Dim strTmp As String
    If mlng����ID = 0 Then Exit Sub
    On Error GoTo errH
    If cboBillType.ToolTipText <> cboBillType.Text Then
        strTmp = cboBillType.Text

        '���·ѱ�
        If Update���¹Һŷѱ�(mstr�Һŵ�, strTmp, 1, p����ҽ��վ) = False Then Exit Sub
        
        If Update������Ϣ(mlng����ID, "�ѱ�", strTmp, p����ҽ��վ) = False Then Exit Sub
        
        cboBillType.Tag = ""
        cboBillType.ToolTipText = cboBillType.Text
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboSelectTime_Click()
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    intDateCount = cboSelectTime.ItemData(cboSelectTime.ListIndex)
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If Me.Visible Then
        If intDateCount = -1 Then
            If Not frmSelectTime.ShowMe(Me, mvCondFilter.Begin, mvCondFilter.End, cboSelectTime) Then
                'ȡ��ʱ�ָ�ԭ����ѡ��
                Call Cbo.SetIndex(cboSelectTime.hwnd, mintOutPreTime)
                Exit Sub
            End If
        ElseIf intDateCount = 0 Then
            '����  86114
            mvCondFilter.Begin = Format(datCurr, "yyyy-MM-dd 00:00:00")
            mvCondFilter.End = Format(datCurr, "yyyy-MM-dd 23:59:59")
        Else
            mvCondFilter.End = Format(datCurr, "yyyy-MM-dd 23:59:59")
            mvCondFilter.Begin = Format(mvCondFilter.End - intDateCount, "yyyy-MM-dd 00:00:00")
        End If
    End If
    'ѡ����ʱ��֮������Һŵ�����
    mvCondFilter.����ID = 0
    mvCondFilter.���� = ""
    mvCondFilter.�ı� = ""
    '�����������֤ÿ���ط���ȡ�ĳ�Ժ���˶�����ͬһʱ�䷶Χ�ڣ�72783��
    Call zlDatabase.SetPara("���ﲡ�˽������", DateDiff("d", datCurr, mvCondFilter.End), glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
    Call zlDatabase.SetPara("���ﲡ�˿�ʼ���", DateDiff("d", mvCondFilter.Begin, datCurr), glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
    cboSelectTime.ToolTipText = Format(mvCondFilter.Begin, "yyyy-MM-dd") & " - " & Format(mvCondFilter.End, "yyyy-MM-dd")
    lblSeeTim.ToolTipText = cboSelectTime.ToolTipText
    mintOutPreTime = cboSelectTime.ListIndex
    Call LoadPatients����
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long, lngTopPanelHeight As Long
        
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
        
    With picBasisNew
        .Visible = True
        .Height = IIf(mbytSize = 0, 1000, 1080)
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        
        lngTopPanelHeight = .Height
    End With
    
    With picMore
        .Visible = True
        .Left = lngLeft
        .Top = lngTop + lngTopPanelHeight
        .Width = lngRight - lngLeft
        .Height = IIf(mbytSize = 0, 850, 1050)
    End With
        
    lngTopPanelHeight = picMore.Height + picBasisNew.Height
    
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

Private Sub cmdOtherFilter_Click()
    Dim datCurr As Date
    
    With mvCondFilter
        .����ID = IIf(.����ID = 0, mlng�������ID, .����ID)
        If frmPatiFilter.ShowMe(Me, .Begin, .End, .����ID, .ҽ��, .����ID, .����, .�ı�, mstrPrivs, p����ҽ��վ) Then
            datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
            Call Cbo.SetIndex(cboSelectTime.hwnd, 3)
            '�����������֤ÿ���ط���ȡ�ĳ�Ժ���˶�����ͬһʱ�䷶Χ�ڣ�72783��
            Call zlDatabase.SetPara("���ﲡ�˽������", DateDiff("d", datCurr, mvCondFilter.End), glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
            Call zlDatabase.SetPara("���ﲡ�˿�ʼ���", DateDiff("d", mvCondFilter.Begin, datCurr), glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
            cboSelectTime.ToolTipText = Format(mvCondFilter.Begin, "yyyy-MM-dd") & " - " & Format(mvCondFilter.End, "yyyy-MM-dd")
            lblSeeTim.ToolTipText = cboSelectTime.ToolTipText
            mintOutPreTime = cboSelectTime.ListIndex
            Call LoadPatients����
        End If
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not mobjDocMsg Is Nothing Then mobjDocMsg.isUnload = Cancel = 0
End Sub


Private Sub mclsAdvices_UpdatePatiInfo(intType As Integer, strInfo As String)
    If intType = 1 Then
        Call UpdatePhysical(strInfo)
    End If
End Sub

Private Sub picHZ_GotFocus()
    If rptPati(PATI_RPT_LIST.PATI_RPT����).Visible Then rptPati(PATI_RPT_LIST.PATI_RPT����).SetFocus
End Sub

Private Sub picJZ_GotFocus()
    If rptPati(PATI_RPT_LIST.PATI_RPT����).Visible Then rptPati(PATI_RPT_LIST.PATI_RPT����).SetFocus
End Sub

Private Sub picPati_Resize()
    On Error Resume Next
    
    picFind.Top = 0
    picFind.Left = 0
    picFind.Width = IIf(mbytSize = 0, 500, 650)
    lblFind.Top = 45
    lblFind.Width = picFind.Width
    PatiIdentify.Left = picFind.Left + picFind.Width
    PatiIdentify.Top = 0
    PatiIdentify.Width = picPati.ScaleWidth - picFind.Width
    
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
    If Check�Ŷӽк� Then
        DoEvents
        mobjQueue.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '����
    PatiIdentify.ActiveFastKey
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim blnTmp As Boolean
    
    If InStr("[|']", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    '�ڱ��벡����Ϣ��ʱ�������Զ���λ���ҿؼ�������Ӱ����Ϣ����д��
    If tbcRegist.Selected.Caption = "���ξ���" Then
        If tbcSub.Visible Then
            If tbcSub.Selected.Tag = "����" Then
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
            And Not Me.ActiveControl Is PatiIdentify And mstrFindType = "���￨" And PatiIdentify.Enabled And PatiIdentify.Visible Then
            PatiIdentify.Text = UCase(Chr(KeyAscii))
            PatiIdentify.NotAutoSel = True
            PatiIdentify.SetFocus
        End If
    End If
End Sub

Private Sub mclsAdvices_VSKeyPress(KeyAscii As Integer)
    If (InStr("0123456789", Chr(KeyAscii)) > 0 Or UCase(Chr(KeyAscii)) >= "A" And UCase(Chr(KeyAscii)) <= "Z") _
        And mstrFindType = "���￨" And PatiIdentify.Enabled And PatiIdentify.Visible Then
        picFind.SetFocus
        PatiIdentify.Text = UCase(Chr(KeyAscii))
        PatiIdentify.NotAutoSel = True
        PatiIdentify.SetFocus
    End If
End Sub

Private Sub InitQueuePara()
'���ܣ���ʼ���ŶӽкŲ���
'�Ŷӽк�ģʽ:1.�������̨������л�ҽ����������;2-�ȷ������,��ҽ�����о���.0-���Ŷӽк�
    
    mty_Queue.strQueuePrivs = ";" & GetPrivFunc(glngSys, p�Ŷӽк�����ģ��) & ";"
    mty_Queue.byt�Ŷӽк�ģʽ = Val(zlDatabase.GetPara("�Ŷӽк�ģʽ", glngSys, p����������))
 
    If mty_Queue.byt�Ŷӽк�ģʽ = 1 Then
        mty_Queue.blnҽ���������� = Val(zlDatabase.GetPara("�ŶӺ���վ��", glngSys, p����������, "0")) = 1
    Else
        mty_Queue.blnҽ���������� = False
    End If
    
    If mty_Queue.blnҽ���������� Then
        mty_Queue.int�������� = Val(zlDatabase.GetPara("ҽ����������", glngSys, p����ҽ��վ))
    Else
        mty_Queue.int�������� = 0
    End If
    mty_Queue.bln���к����� = Val(zlDatabase.GetPara("��������������", glngSys, p����ҽ��վ, "1")) = 1
End Sub

Private Sub Form_Load()
    Dim objPane As Pane, strTab As String, intIdx As Integer
    Dim i As Integer, arrType() As String
    Dim objTabItem As TabControlItem
    Dim arrTmp As Variant, strTmp As String
    
    mstrPrivs = ";" & gstrPrivs & ";"
    mlngModul = glngModul
    mblnShowLeavePati = False
    Call GetLocalSetting '���ز���
        
    Set mclsReg = New zlPublicExpense.clsRegist
    Call mclsReg.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
    Call mclsReg.zlInitData(4)
    
    Set mclsDis = New zl9Disease.clsDisease
    Call mclsDis.InitDisease(gcnOracle, Me, glngSys, glngModul, mstrPrivs, Nothing)

    Call ZLCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)
    
    Call InitQueuePara
    'һ��ͨ������ʼ������tbcSub_SelectedChanged֮ǰ���Ա㴫�ݸ�ҽ������
     'zlGetIDKindStr�л��Զ�����Ϊ����8λ����
    mstrCardKind = "��|���￨|0|0|8|0|0|0;��|��ʶ��|0|0|0|0|0|0;��|�Һŵ�|0|0|0|0|0|0;��|����|0|0|0|0|0|0;��|�������֤|0|0|0|0|0|0;�ɣ�|�ɣÿ�|1|0|0|0|0|0"
    If Check�Ŷӽк� = True Then mstrCardKind = mstrCardKind & ";��|�ŶӺ�|0|0|0|0|0|0;ҽ|ҽ����|0|0|0|0|0|0"
    If InitObjOneCardComLib(Me, p����ҽ��վ) Then
        mstrCardKind = gobjOneCardComLib.zlGetIDKindStr(mstrCardKind)
    End If
    Call PatiIdentify.zlInit(Me, glngSys, p����ҽ��վ, gcnOracle, gstrDBUser, gobjOneCardComLib, mstrCardKind, "zl9CISJob")
    PatiIdentify.objIDKind.AllowAutoICCard = True
    PatiIdentify.objIDKind.AllowAutoIDCard = True
    mblnIsInit = True

    '��ʼ�����ӿڲ���
    If InStr(GetInsidePrivs(P����ڲ��ӿ�, , 2100), ";��챨�����;") > 0 Then
        On Error Resume Next
        Set mobjPeis = CreateObject("zlPublicPeis.clsPublicPeis")
        Err.Clear: On Error GoTo 0
        If Not mobjPeis Is Nothing Then
            If mobjPeis.Initialize(gcnOracle) = False Then
                Set mobjPeis = Nothing
                MsgBox "���ӿڲ�����zlPublicPeis����ʼ��ʧ��!", vbInformation, gstrSysName
            End If
        End If
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
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = ZLCommFun.GetPubIcons
    
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, IIf(mbytSize = 0, 310, 320), 400, DockLeftOf, Nothing)
    objPane.Title = "���ﲡ���б�"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set objPane = Me.dkpMain.CreatePane(2, 310, 100, DockBottomOf, objPane)
    objPane.Title = "��Ϣ����"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable

    'TabControl
    '-----------------------------------------------------
    Call ZLCommFun.SetWindowsInTaskBar(Me.hwnd, True)
 
    '�����б�
    With Me.tbcWait
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem(0, "���ﲡ��", picHZ.hwnd, 0).Tag = "���ﲡ��"
        .InsertItem(1, "�Ŷӽк�", picTmpH(0).hwnd, 0).Tag = "�Ŷӽк�"
        .InsertItem(2, "ԤԼ����", picYy.hwnd, 0).Tag = "ԤԼ����"
        
        If Not mbln��ʾԤԼ���� Then .Item(2).Selected = True
        .Item(2).Visible = Not mbln��ʾԤԼ����
        .Item(1).Selected = True
        .Item(0).Selected = True
        
        Call .RemoveItem(1)
        
        If Check�Ŷӽк� Then
            .InsertItem(1, "�Ŷӽк�", mobjQueue.zlGetForm.hwnd, 0).Tag = "�Ŷӽк�"
        End If
        
        mintԤԼ�б� = IIf(Check�Ŷӽк�, 2, 1)
    End With
    
    '�����б�
    With Me.tbcInTreat
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem(0, "����", picJZ.hwnd, 0).Tag = "���ھ���"
        .InsertItem(1, "����", picYZ.hwnd, 0).Tag = "���ﲡ��"
        .InsertItem(2, "����", picHUIZ.hwnd, 0).Tag = "����ﲡ��"
        
        .Item(2).Selected = True
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    
    Set mclsAdvices = New zlPublicAdvice.clsDockOutAdvices
    Set mclsEPRs = New zlRichEPR.cDockOutEPRs
    Set mclsDisease = New zlRichEPR.cDockDisease
    Set mobjPati = New frmDockPatiInfo
    mobjPati.mintҽ��վģ��� = p����ҽ��վ
    
    If GetInsidePrivs(p�°����ﲡ��, True) <> "" Then
        Set mclsEMR = DynamicCreate("zlRichEMR.clsDockEMR", "���Ӳ���")
        If Not mclsEMR Is Nothing Then
            If Not mclsEMR.Init(gobjEmr, gcnOracle, glngSys) Then
                Set mclsEMR = Nothing
            End If
        End If
    End If
    
    If InStr(";" & gstrPrivs & ";", ";����һ��;") > 0 Then
        Set mfrmView = New frmOutDoctorView
        mbyt���ξ��� = 1
    Else
        mbyt���ξ��� = 0
    End If
    
    Set mcolSubForm = New Collection
    If mbyt���ξ��� = 1 Then
        mcolSubForm.Add mfrmView, "_����һ��"
    End If
    mcolSubForm.Add mobjPati, "_����"
    
    mcolSubForm.Add mclsAdvices.zlGetForm, "_ҽ��"
    mcolSubForm.Add mclsEPRs.zlGetForm, "_����"
    mcolSubForm.Add mclsDisease.zlGetForm, "_��������"
    If Not mclsEMR Is Nothing Then
        mcolSubForm.Add mclsEMR.zlGetForm, "_�²���"
    End If
    
    
    '---------------------------------------------------
    '���ξ����б�
    With Me.tbcRegist
        Set tbcRegist.Icons = ZLCommFun.GetPubIcons
        With .PaintManager
            .Appearance = xtpTabAppearanceStateButtons
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        If mbyt���ξ��� = 1 Then
             .InsertItem(0, "����һ��", mcolSubForm("_����һ��").hwnd, 0).Tag = "����һ��"
        End If
        .InsertItem(mbyt���ξ���, "���ξ���", picRegist.hwnd, 0).Tag = "-1"
            intIdx = mbyt���ξ��� + 1
        .InsertItem(intIdx, "��ʷ����1", picRegist.hwnd, 0).Tag = "-1"
            .Item(intIdx).Visible = False: intIdx = intIdx + 1
        .InsertItem(intIdx, "��ʷ����2", picRegist.hwnd, 0).Tag = "-1"
            .Item(intIdx).Visible = False: intIdx = intIdx + 1
        .InsertItem(intIdx, "��ʷ����3", picRegist.hwnd, 0).Tag = "-1"
            .Item(intIdx).Visible = False: intIdx = intIdx + 1
        .InsertItem(intIdx, "����", picRegist.hwnd, 0).Tag = "����"
            .Item(intIdx).Visible = False
        intIdx = 0
        tbcSub.Visible = True
        picRegist.Visible = True
    End With
    
    '�ڲ���Ƭ
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem(intIdx, "������Ϣ", picTmp.hwnd, 0).Tag = "����": intIdx = intIdx + 1
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
       
        If GetInsidePrivs(p����ҽ���´�) <> "" Then
            '�ȼ���ҽ����ԭ��:�����������ӿڣ����ͻ���û����������ʱ������ȼ����Ŷӽкź����ҽ����ʱ��
            '�ӡ�������Ϣ���л�����ҽ����Ϣ�����򵯳�Msgbox���� �����:67995
            .InsertItem(intIdx, "ҽ����Ϣ", mcolSubForm("_ҽ��").hwnd, 0).Tag = "ҽ��": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p���ﲡ������) <> "" Then
            .InsertItem(intIdx, "������Ϣ", picTmp.hwnd, 0).Tag = "����": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p�°����ﲡ��, True) <> "" And Not mclsEMR Is Nothing Then
            .InsertItem(intIdx, "���Ӳ���", picTmp.hwnd, 0).Tag = "�²���": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p����������д, True) <> "" Then
            Set objTabItem = .InsertItem(intIdx, "��������", picTmp.hwnd, 0): objTabItem.Tag = "��������": objTabItem.Visible = False: intIdx = intIdx + 1
        End If
        
        '����ṩ�Ŀ�Ƭ
        Call CreatePlugInOK(p����ҽ��վ)
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strTmp = gobjPlugIn.GetFormCaption(glngSys, p����ҽ��վ)
            Call zlPlugInErrH(Err, "GetFormCaption")
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ",")
                For i = 0 To UBound(arrTmp)
                    strTmp = arrTmp(i)
                    mcolSubForm.Add gobjPlugIn.GetForm(glngSys, p����ҽ��վ, strTmp), "_" & strTmp
                    .InsertItem(intIdx, strTmp, mcolSubForm("_" & strTmp).hwnd, 0).Tag = strTmp: intIdx = intIdx + 1
                    Call zlPlugInErrH(Err, "GetForm")
                Next
            End If
            Err.Clear: On Error GoTo 0
        End If
        
        If .ItemCount = 0 Then
            MsgBox "��û��ʹ�ü���ҽ������վ��Ȩ�ޡ�", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        '�ָ��ϴ�ѡ��Ŀ�Ƭ
        strTab = zlDatabase.GetPara("ҽ������", glngSys, p����ҽ��վ)
        
        For intIdx = 0 To tbcSub.ItemCount - 1
            If tbcSub(intIdx).Visible And tbcSub(intIdx).Tag = strTab Then Exit For
        Next
        If intIdx <= tbcSub.ItemCount - 1 Then
            strTab = .Item(intIdx).Tag
            .Item(intIdx).Tag = "" '���⼤���¼�
            .Item(intIdx).Selected = True
            .Item(intIdx).Tag = strTab
        Else
            .Item(0).Selected = True '�½�ʱ���Զ�ѡ�������,�����ټ����¼�
        End If
    End With
    
    tbcRegist.Item(mbyt���ξ���).Selected = True
    
    Call tbcRegist_SelectedChanged(tbcRegist.Selected)
    mstrPreSubTab = ""
    'ֻ����ѡ����Ӵ���
    Call tbcSub_SelectedChanged(tbcSub.Selected)
            
    '��ȡ��������
    '-----------------------------------------------------
    mblnUnRefresh = True
    mblnΣ��ֵshow = False
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
    Err.Clear: On Error GoTo 0
    
    Call InitCboData
    Call InitReportColumn
    Call InitCondFilter '���ﲡ�˹�������
    
    Call LoadPatients '��ʾ����
    Call LoadNotify '��Ϣ����
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        '��ָ�Panne�ı���,Tag�����
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, "")
    End If
 
    '����ȱʡ���ҷ�ʽ
    arrType = Split(mstrCardKind, ";")
    For i = 1 To UBound(arrType) + 1
        If i = mintFindType Then
            PatiIdentify.objIDKind.IDKind = i
            Exit For
        End If
    Next
    
    Call RestoreWinState(Me, App.ProductName, , True)
    Call Get����վ�������(Me, tbcSub.Selected.Caption)
    'ReportControl�ؼ����������޷��ָ�Ҫ��������
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        For i = 0 To rptPati.Count - 1
            strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\ReportControl", "rptPati" & "_" & i, "")
            rptPati(i).LoadSettings strTmp
        Next
    End If
    
    Me.WindowState = vbMaximized
    Call SetFixedCommandBar(cbsMain(2).Controls)
    fraPatiUD.Top = 3000
    If Check�Ŷӽк� = True Then
                fraPatiUD.Top = 5000
        '����Ƿ�����Ŷӽк�
        Call ReshDataQueue
        tbcWait.Item(1).Selected = True
    End If
    
    Call RefreshPass

    If InStr(";" & gstrPrivs & ";", ";�޸�ҽ�Ƹ��ʽ;") = 0 Then
        cboPayType.Locked = True
    End If
    If InStr(";" & gstrPrivs & ";", ";�޸ķѱ�;") = 0 Then
        cboBillType.Locked = True
    End If
    Call SetReceiveToday(True, 0)
    
    '����Ϣ����
    If gbln��ϵͳ Then
        Set mobjDocMsg = New frmDocMsg
        mobjDocMsg.ShowMe Me, 1
    End If
    
    dkpMain.RecalcLayout
    mblnUnRefresh = False
End Sub

Private Sub RefreshPass()
    '�Ƿ����̫Ԫͨ�ӿڲ���
    mblnUseTYT = False
    If gbytPass = 3 Then
        If gint����������Դ = 0 Then
            mint����������Դ = Val(zlDatabase.GetPara("����������Դ", glngSys, p����ҽ��վ, "0"))
        End If
        mblnUseTYT = gint����������Դ = 0 And mint����������Դ = 1 Or gint����������Դ = 2
    End If
    '����̫Ԫͨ�ӿڶ��󣬴���ʧ�ܣ�������̫Ԫͨ
    On Error Resume Next
    If gobjPass Is Nothing Then
        Set gobjPass = DynamicCreate("zlPassInterface.clsPass", "������ҩ���", True)
        If Not gobjPass Is Nothing Then
             If Not gobjPass.zlPassInit(gcnOracle, glngSys, 5) Then gbytPass = 0
        End If
    Else
        gbytPass = gobjPass.PassType
    End If
    If Err.Number <> 0 Then Err.Clear: gbytPass = 0
    If gobjPass Is Nothing Then gbytPass = 0
    On Error GoTo 0
    
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim i As Long
    Dim str�Һŵ� As String
    Dim rsTmp As Recordset
    Dim intFindTypeTmp As Integer
    
    If Control.ID <> 0 Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '������
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '��ť����
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
    Case conMenu_View_ToolBar_Size '��ͼ��
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_FontSize_S 'С����
        If mbytSize <> 0 Then
            mbytSize = 0
            Call zlDatabase.SetPara("����", mbytSize, glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_FontSize_L '������
        If mbytSize <> 1 Then
            mbytSize = 1
            Call zlDatabase.SetPara("����", mbytSize, glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_Find '����
        If Me.ActiveControl Is PatiIdentify Then
            PatiIdentify.SetFocus '��ʱ��Ҫ��λһ��
            If PatiIdentify.Text <> "" Then
                Call ExecuteFindPati
            End If
        Else
            PatiIdentify.SetFocus
        End If
    Case conMenu_View_FindNext '������һ��
        If PatiIdentify.Text = "" And mstrIDCard = "" Then
            PatiIdentify.SetFocus
        Else
            Call ExecuteFindPati(True, IIf(PatiIdentify.Text = "", mstrIDCard, ""))
        End If
    Case conMenu_View_Busy '����״̬
        Call SetRoomState(lblRoom.BackColor = COLOR_FREE)
    Case conMenu_View_Refresh 'ˢ��
        Call LoadPatients
        Call LoadNotify
    Case conMenu_View_Jump '��ת
        If Me.tbcSub.Selected.Index + 1 <= Me.tbcSub.ItemCount - 1 Then
            Me.tbcSub.Item(Me.tbcSub.Selected.Index + 1).Selected = True
        Else
            Me.tbcSub.Item(0).Selected = True
        End If
    Case conMenu_File_Parameter '��������
        frmEMCStationSetup.mstrPrivs = mstrPrivs
        frmEMCStationSetup.Show 1, Me

        If gblnOK Then
            intFindTypeTmp = mintFindType
            Call GetLocalSetting
            mintFindType = intFindTypeTmp
            Call LoadPatients
            Call InitQueuePara
        End If
        If Check�Ŷӽк� Then
            Call ReshDataQueue
        End If
        Me.tbcWait.Item(mintԤԼ�б�).Visible = Not mbln��ʾԤԼ����
        If Me.tbcWait.Item(mintԤԼ�б�).Visible = False Then
             Me.tbcWait.Item(0).Selected = True
        End If
    Case conMenu_Tool_KssAudit '������ҩ���
        On Error Resume Next
        Call frmExamineKSS.Show(0, Me)
     Case conMenu_Tool_TransAudit '��Ѫ��˹���
        On Error Resume Next
        Call frmExamineTransfuse.ShowMe(Me, 1)
    Case conMenu_Tool_Archive '���Ӳ�������
        Call frmArchiveView.ShowArchive(Me, mPatiInfo.����ID, mPatiInfo.�Һ�ID)
    Case conMenu_Tool_ExaReport
        '���ó¸����ṩ�Ľӿ�
        If Not mobjPeis Is Nothing And mlng����ID <> 0 Then
        
            If Not OpenExaReportNew Then
                If mobjPeis.HasExaminationReport(mlng����ID) = True Then
                    Call mobjPeis.OpenExaminationReport(Me, mlng����ID)
                Else
                    MsgBox "��ǰ������������ܼ챨�档", vbInformation, gstrSysName
                End If
            End If
            
        End If
    Case conMenu_Tool_Reference_1 '������ϲο�
        Call gobjKernel.ShowDiagHelp(vbModeless, Me)
    Case conMenu_Tool_Reference_2 '���ƴ�ʩ�ο�
        Call ShowClinicHelp(vbModeless, Me, p����ҽ��վ, 0, 1, mPatiInfo.����ID, mPatiInfo.�Һ�ID)
        
    Case conMenu_Tool_Community * 100# + 1 '���������֤
        Call ExecuteCommunityIdentify
    Case conMenu_Tool_Community * 100# + 2 To conMenu_Tool_Community * 100# + 99 '������������
        If Not gobjCommunity Is Nothing And mPatiInfo.���� <> 0 And mPatiInfo.�Һ�ID <> 0 Then
            If gobjCommunity.CommunityFunc(glngSys, mlngModul, Val(Control.Parameter), mPatiInfo.����, mPatiInfo.������, mPatiInfo.����ID, mPatiInfo.�Һ�ID) Then
                Call LoadPatients
            End If
        End If
    Case conMenu_Manage_Regist '���˹Һ�
        Control.Enabled = False
        Call mclsReg.zlExecuteCommandBars(Me, Control, str�Һŵ�)
        Call ExecuteRegist(str�Һŵ�)
        If str�Һŵ� <> "" Then Call SetReceiveToday(False, 1): Call ReceiveAfterExec
        Control.Enabled = True
    Case conMenu_Manage_Bespeak 'ԤԼ�Һ�
        Control.Enabled = False
        Call mclsReg.zlExecuteCommandBars(Me, Control, "", mlng����ID)
        Control.Enabled = True
    Case conMenu_Edit_AppRequest
        Control.Enabled = False
        Call mclsReg.zlExecuteCommandBars(Me, Control, "", mlng����ID)
        Control.Enabled = True
    Case conMenu_Edit_AppRequestManage
        Control.Enabled = False
        Call mclsReg.zlExecuteCommandBars(Me, Control, "")
        Control.Enabled = True
    Case conMenu_View_Option '"�Һ�ѡ������"
        Control.Enabled = False
        Call mclsReg.zlExecuteCommandBars(Me, Control, "")
        Control.Enabled = True
    Case conMenu_File_Print_Bespeak '�ش�ԤԼ�Һŵ�
        Control.Enabled = False
        Call ExecuteBespeakPrint
        Control.Enabled = True
    Case conMenu_Manage_Transfer_Send '����ת��
        Call ExecuteTransferSend
    Case conMenu_Manage_Transfer_Cancel 'ȡ��ת��
        Call ExecuteTransferCancel
    Case conMenu_Manage_Transfer_Incept '����ת��
        Call ExecuteTransferIncept
    Case conMenu_Manage_Transfer_Refuse 'ת��ܾ�
        Call ExecuteTransferRefuse
    Case conMenu_Manage_Transfer_Force 'ǿ������
        str�Һŵ� = frmForceGet.ShowMe(Me, mstrPrivs, mlng�������ID, gobjOneCardComLib, p����ҽ��վ)
        If str�Һŵ� <> "" Then
            If rptPati(PATI_RPT����).Visible Then
                Call LoadPatients("11001", PATI_RPT����, str�Һŵ�)
            Else
                Call LoadPatients("11001")
            End If
        End If
    Case conMenu_Manage_Receive '���˽���
        Call ExecuteReceive
    Case conMenu_Manage_Cancel 'ȡ������
        Call ExecuteCancel
    Case conMenu_Manage_Finish '��ɽ���
        Call ExecuteFinish
    Case conMenu_Manage_Redo '�ָ�����
        Call ExecuteRedo
    Case conMenu_Manage_ReBack '��ͣ����
        Call ExecuteStopAndReuse(False)
    Case conMenu_Manage_ReBackCancel '�ָ���ͣ����
          Call ExecuteStopAndReuse(True)
   Case conmenu_View_Leave  '��ʾ�����ﲡ��
         mblnShowLeavePati = Not mblnShowLeavePati
         Control.Checked = mblnShowLeavePati
        Call LoadPatients("10001")
    Case conmenu_Edit_Leave     '���˲�����
        If Set���˹Һ�״̬(-1) Then
            Call LoadPatients("10001")
            Call ReshDataQueue
        End If
    Case conmenu_Edit_Wait      '���˾���
        If Set���˹Һ�״̬(0) Then
            Call LoadPatients("10001")
            Call ReshDataQueue
        End If
    Case conMenu_Manage_AdjustGrade  '�������鼶��
        Call ExecAdjustGrade
    Case conMenu_Manage_Green         '�����ɫͨ��
        Call ExecTagGreen
        
        
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_Help '����
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '�˳�
        Unload Me
    Case conMenu_Tool_HealthCard  '���񽡿���
        If InitObjOneCardComLib(Me, p����ҽ��վ) Then
            Call gobjOneCardComLib.zlHealthArchivesShow(Me, p����ҽ��վ, mlng����ID, "")
        End If
    Case conMenu_Edit_TraReactionRecord '��Ѫ��Ӧ
        Call FuncTraReactionRecord(Me, 0, p����ҽ���´�)
    Case conMenu_Edit_NewItemQAdvice
        Call ExecuteTabChange("ҽ��")
    Case conMenu_Edit_NewItemQEpr
        Call ExecuteTabChange("����")
    Case conMenu_Tool_Positive '���Խ���鿴
        Call mclsDis.ShowRegistByPati(Me, ByVal 1, mlng����ID, , mstr�Һŵ�)
    Case conMenu_Tool_Critical 'Σ��ֵ�鿴����
        Call ExecuteCritical
    Case Else
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            'ִ�з�������ǰģ��ı���
            With mPatiInfo
                If Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_2" Or Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_3" Then
                    If mlng�������ID = 0 Then
                        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
                    Else
                        Set rsTmp = zlDatabase.OpenSQLRecord("Select ���� From ���ű� Where ID=[1]", Me.Caption, mlng�������ID)
                        If rsTmp.EOF Then Exit Sub
                        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                            "��������=" & rsTmp!���� & "|=" & mlng�������ID)
                    End If
                Else
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                        "����ID=" & mPatiInfo.����ID, "�����=" & .�����, "�Һŵ�=" & .�Һŵ�, "����=" & .����)
                End If
            End With
        Else
            If Check�Ŷӽк� = True Then
                mobjQueue.zlExecuteCommandBars Control
            End If
            Select Case Me.tbcSub.Selected.Tag
            Case "ҽ��"
                Call mclsAdvices.zlExecuteCommandBars(Control)
            Case "����"
                Call mclsEPRs.zlExecuteCommandBars(Control)
            Case "�²���"
                Call mclsEMR.zlExecuteCommandBars(Control)
            Case "��������"
                Call mclsDisease.zlExecuteCommandBars(Control)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.ExeButtomClick(glngSys, p����ҽ��վ, mcolSubForm("_" & tbcSub.Selected.Tag), tbcSub.Selected.Tag, Control.Caption, mlng����ID, 0, mstr�Һŵ�)
                    Call zlPlugInErrH(Err, "ExeButtomClick")
                    Err.Clear: On Error GoTo 0
                End If
            End Select
        End If
    End Select
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl
    Dim strFunc As String, arrFunc As Variant
    Dim i As Long
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID

    Case conMenu_Manage_Transfer
        With CommandBar.Controls
            If .Count = 0 Then
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Send, "ת�ﲡ��(&S)", -1, False)
                objControl.IconId = conMenu_Manage_Transfer
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Cancel, "ȡ��ת��(&C)", -1, False)
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Incept, "ת�����(&I)", -1, False)
                objControl.IconId = conMenu_Manage_Receive
                objControl.BeginGroup = True
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Refuse, "ת��ܾ�(&R)", -1, False)
                Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Force, "ǿ������(&F)", -1, False)
                objControl.BeginGroup = True
            End If
        End With
    Case conMenu_Tool_Community '��������
        mlngCommunityID = 0
        With CommandBar.Controls
            .DeleteAll
            If Not gobjCommunity Is Nothing Then
                '������֤
                If mPatiInfo.���� = 0 Then
                    Set objControl = .Add(xtpControlButton, conMenu_Tool_Community * 100# + 1, "�����֤(&V)")
                End If
                
                '��������
                If mPatiInfo.���� <> 0 Then
                    strFunc = gobjCommunity.GetCommunityFunc(glngSys, p����ҽ��վ, mPatiInfo.����)
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
       Case "ҽ��"
           Call mclsAdvices.zlPopupCommandBars(CommandBar)
       Case "����"
       End Select
       
       'ˢ�²˵���
       cbsMain.RecalcLayout
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean, bln�쳣���� As Boolean
    Dim strTmp As String
 
     If mPr > 0 And (mintRPTIndex = PATI_RPT���� Or mintRPTIndex = PATI_RPTԤԼ) Then
        If rptPati(mintRPTIndex).Rows(mPr).Record.Tag = "��" Then
             bln�쳣���� = True
        End If
    End If
 
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '������
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text 'ͼ������
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '��ͼ��
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Me.stbThis.Visible
    Case conMenu_View_FontSize_S 'С����
        Control.Checked = Not (mbytSize = 1)
    Case conMenu_View_FontSize_L '������
        Control.Checked = (mbytSize = 1)
    Case conMenu_View_Busy '����״̬
        Control.Checked = lblRoom.BackColor = COLOR_BUSY
    Case conMenu_Tool_KssAudit  '������ҩ���
        If GetInsidePrivs(p������ҩ���) = "" Then
            Control.Visible = False
        End If
    Case conMenu_Tool_TransAudit '��Ѫ�ּ�����
        If GetInsidePrivs(p��Ѫ��˹���) = "" Or Not gbln��Ѫ�ּ����� Then
            Control.Visible = False
        End If
    Case conMenu_Tool_Archive '���Ӳ�������
        If GetInsidePrivs(p���Ӳ�������) = "" Then
            Control.Visible = False
        Else
            Control.Enabled = mlng����ID <> 0
        End If
    Case conMenu_Tool_ExaReport
        Control.Enabled = mlng����ID <> 0 And (Not mobjPeis Is Nothing)
    Case conMenu_Tool_HealthCard  '���񽡿���
        Control.Enabled = mlng����ID <> 0
    Case conMenu_Tool_Reference_1 '������ϲο�
        If GetInsidePrivs(p������ϲο�) = "" Then Control.Visible = False
    Case conMenu_Tool_Reference_2 'ҩƷ�����Ʋο�
        If GetInsidePrivs(pҩƷ���Ʋο�) = "" Then Control.Visible = False
    Case conMenu_Tool_Community '�����˵�
        If gobjCommunity Is Nothing Then
            Control.Visible = False
        End If
    Case conMenu_Edit_TraReactionRecord '��Ѫ��Ӧ
        Control.Visible = InStr(1, GetInsidePrivs(9005, , 2200), "��Ѫ��Ӧ�Ǽ�") <> 0
        Control.Enabled = Control.Visible And gblnѪ��ϵͳ
                
    Case conMenu_Tool_Community * 100# + 1 '���������֤
        Control.Enabled = mlng����ID <> 0 And mPatiInfo.���� = 0 And (mPatiInfo.���� = pt���� Or mPatiInfo.���� = pt����) And InStr(mstrPrivs, "���˽���") > 0
    Case conMenu_Tool_Community * 100# + 2 To conMenu_Tool_Community * 100# + 99 '������������
        Control.Enabled = mlng����ID <> 0 And mPatiInfo.���� <> 0

    Case conMenu_File_MedRec '��ҳ��ӡ
        If InStr(mstrPrivs, "��ӡ��ҳ") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mlng����ID <> 0
        End If
    Case conMenu_ManagePopup '������˵�
        If InStr(mstrPrivs, ";���˽���;") = 0 Then Control.Visible = False
    Case conMenu_File_Print_Bespeak
        Control.Visible = InStr(mstrPrivs, ";ԤԼ�Һŵ�;") > 0 And (rptPati(PATI_RPT����).Visible Or rptPati(PATI_RPTԤԼ).Visible)
        blnEnabled = False
        If mPr <> -1 Then
            If mintRPTIndex = PATI_RPTԤԼ Or mintRPTIndex = PATI_RPT���� Then
                strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_HZ_��ʶ).Value
                blnEnabled = (strTmp = "Ԥ")
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Manage_Transfer 'ת�ﴦ��
        If InStr(mstrPrivs, "���˽���") = 0 _
            And InStr(mstrPrivs, "����ת��") = 0 _
                And InStr(mstrPrivs, "���ﲡ��") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_Manage_Transfer_Send '����ת��
        If InStr(mstrPrivs, "����ת��") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = (mintActive = pt���� Or mintActive = pt����)
            If blnEnabled Then
                If mPr <> -1 And rptPati(mintRPTIndex).Visible Then
                    If mintRPTIndex = PATI_RPT���� Or mintRPTIndex = PATI_RPTԤԼ Then
                        If rptPati(mintRPTIndex).Rows(mPr).Record.Tag = "��" Then '�쳣���ݲ���ת��
                             strTmp = "-1"
                        Else
                            strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_HZ_״̬).Value
                        End If
                        
                    ElseIf mintRPTIndex = PATI_RPT���� Or mintRPTIndex = PATI_RPT���� Then
                        strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_JZ_״̬).Value
                    End If
                    blnEnabled = (strTmp = "" Or Val(strTmp) = 1)
                Else
                    blnEnabled = False
                End If
            End If
            
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Manage_Transfer_Cancel 'ȡ��ת��
        If InStr(mstrPrivs, "����ת��") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = (mintActive = pt���� Or mintActive = pt����)
            If blnEnabled Then
                If mPr <> -1 And rptPati(mintRPTIndex).Visible Then
                    If mintRPTIndex = PATI_RPT���� Or mintRPTIndex = PATI_RPTԤԼ Then
                        strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_HZ_״̬).Value
                    ElseIf mintRPTIndex = PATI_RPT���� Or mintRPTIndex = PATI_RPT���� Then
                        strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_JZ_״̬).Value
                    End If
                    blnEnabled = (strTmp <> "" And Val(strTmp) = 0 Or Val(strTmp) = -1)
                Else
                    blnEnabled = False
                End If
            End If
            Control.Enabled = blnEnabled
        End If
    Case conmenu_View_Leave  '��ʾ�����ﲡ��
        Control.Checked = mblnShowLeavePati
    Case conmenu_Edit_Leave
        If bln�쳣���� = True Then
            blnEnabled = False
        Else
            blnEnabled = (mintActive = pt����)
            If blnEnabled Then
                If mPr <> -1 And rptPati(mintRPTIndex).Visible Then
                    strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_HZ_ִ��״̬).Value
                    blnEnabled = Val(strTmp) = 0
                Else
                    blnEnabled = False
                End If
            End If
        End If
        Control.Enabled = blnEnabled
            
    Case conmenu_Edit_Wait
        If bln�쳣���� = True Then
            blnEnabled = False
        Else
            blnEnabled = mintActive = pt����
            If blnEnabled Then
                If mPr <> -1 And rptPati(mintRPTIndex).Visible Then
                    strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_HZ_ִ��״̬).Value
                    blnEnabled = Val(strTmp) = -1
                Else
                    blnEnabled = False
                End If
                
            End If
        End If
        Control.Enabled = blnEnabled
        
    Case conMenu_Manage_Transfer_Incept, conMenu_Manage_Transfer_Refuse 'ת�����,ת��ܾ�
        If InStr(mstrPrivs, "���˽���") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = (mintActive = ptת�� And mPr <> -1 And rptPati(mintRPTIndex).Visible)
            Control.Enabled = blnEnabled
        End If
        
    Case conMenu_Manage_Transfer_Force 'ǿ������
        If InStr(mstrPrivs, "���˽���") = 0 Or InStr(mstrPrivs, "���ﲡ��") = 0 Then Control.Visible = False
    Case conMenu_Manage_ReBack '��ͣ����:�����
        If InStr(mstrPrivs, "���˽���") = 0 Then
            Control.Visible = False
        Else
            If mintActive = pt���� And mPr <> -1 And rptPati(mintRPTIndex).Visible Then
                strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_JZ_��¼��־).Value
                blnEnabled = Val(strTmp) < 2
            Else
                blnEnabled = False
            End If
            
            If blnEnabled Then
                If mstr����ҽ�� <> UserInfo.���� Then
                    blnEnabled = False
                    If InStr(GetInsidePrivs(p����ҽ��վ), ";��������ҽ���Ĳ���;") > 0 Then
                        blnEnabled = True
                    End If
                End If
            End If
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Manage_ReBackCancel '�ָ���ͣ����
        If InStr(mstrPrivs, "���˽���") = 0 Then
            Control.Visible = False
        Else
            
            If mintActive = pt���� And mPr <> -1 And rptPati(mintRPTIndex).Visible Then
                strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_JZ_��¼��־).Value
                Control.Enabled = Val(strTmp) = 2
            Else
                Control.Enabled = False
            End If
            
        End If
    Case conMenu_Manage_Receive '���˽���
        If InStr(mstrPrivs, "���˽���") = 0 Or (mty_Queue.blnҽ���������� And mbln���к����) Then
            Control.Enabled = False
            Control.Visible = False
        Else
            Control.Visible = True
            '���ԤԼ�ҺŲ��˿���ֱ�ӽ��ת�ﲡ�˲�ͨ���������
            blnEnabled = False
            
            If (mintRPTIndex = PATI_RPT���� Or mintRPTIndex = PATI_RPTԤԼ) And rptPati(mintRPTIndex).Visible Then
                blnEnabled = mPr <> -1
            End If
            Control.Enabled = blnEnabled    '�������жϵ�ǰ�Ƿ�Ϊת�ﲡ���б���Ϊ�����ת���б�Ļ���blnEnabled�Ѿ���False
             
        End If
    Case conMenu_Manage_Cancel 'ȡ������
        If InStr(mstrPrivs, "���˽���") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mintActive = pt���� And mPr <> -1 And rptPati(mintRPTIndex).Visible
        End If
    Case conMenu_Manage_Finish '��ɾ���
        If InStr(mstrPrivs, "���˽���") = 0 Then
            Control.Visible = False
        ElseIf mintRPTIndex = PATI_RPT���� Or mintRPTIndex = PATI_RPT���� Then
            blnEnabled = mPr <> -1 And rptPati(mintRPTIndex).Visible
            If mstr����ҽ�� <> UserInfo.���� And blnEnabled Then
                blnEnabled = False
                If InStr(GetInsidePrivs(p����ҽ��վ), ";��������ҽ���Ĳ���;") > 0 Then
                    blnEnabled = True
                End If
            End If
            Control.Enabled = blnEnabled
        Else
            Control.Enabled = False
        End If
    Case conMenu_Manage_Redo '�ָ�����
        If InStr(mstrPrivs, "���˽���") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = mintActive = pt���� And mPr <> -1 And rptPati(mintRPTIndex).Visible
            If blnEnabled Then 'ֻ�ָܻ��������ѵĲ���(������Ȩ�޿���ǿ������)
                blnEnabled = rptPati(mintRPTIndex).Rows(mPr).Record(COL_YZ_ִ����).Value = UserInfo.����
            End If
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Edit_NewItemQAdvice
        If tbcSub.Selected.Tag = "ҽ��" And tbcRegist.Selected.Tag <> "����һ��" Then
            Control.Visible = False
        Else
            Control.Visible = True
            blnEnabled = True
            If mstr����ҽ�� <> UserInfo.���� And blnEnabled Then
                blnEnabled = False
                If InStr(GetInsidePrivs(p����ҽ��վ), ";��������ҽ���Ĳ���;") > 0 Then
                    blnEnabled = True
                End If
            End If
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Edit_NewItemQEpr
        If tbcSub.Selected.Tag = "����" And tbcRegist.Selected.Tag <> "����һ��" Then
            Control.Visible = False
        Else
            Control.Visible = True
        End If
    Case conMenu_Tool_Positive
        Control.Enabled = mlng����ID <> 0
    Case conMenu_Manage_AdjustGrade  '�������鼶��
        If InStr(mstrPrivs, "���˽���") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mlng����ID <> 0 And mintActive <> pt����
        End If
    Case conMenu_Manage_Green         '�����ɫͨ��
        If InStr(mstrPrivs, "���˽���") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mlng����ID <> 0 And mintActive <> pt����
        End If
    Case Else
        '60075:������,2013-04-03,���ⲿ��ҽ����ӡ��Ԥ���˵��Ĵ�����ֲ���˴�,��ǰ�ķ�ʽ�����޷���������ģ��ĸ����¼�
        If (Control.ID = conMenu_File_Print Or Control.ID = conMenu_File_Preview Or Control.ID = conMenu_Help_Help) Then
            If tbcSub.Selected.Tag = "ҽ��" Then
                Control.Visible = False
                Exit Sub
            Else
                Control.Visible = True
            End If
        End If
        If Check�Ŷӽк� Then mobjQueue.zlUpdateCommandBars Control
        mclsReg.zlUpdateCommandBars Control
        Select Case tbcSub.Selected.Tag
        Case "ҽ��"
            Call mclsAdvices.zlUpdateCommandBars(Control)
        Case "����"
            Call mclsEPRs.zlUpdateCommandBars(Control)
        Case "�²���"
            Call mclsEMR.zlUpdateCommandBars(Control)
        Case "��������"
            Call mclsDisease.zlUpdateCommandBars(Control)
        End Select
        '������ҩ����
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
'���ܣ�ˢ���Ӵ���˵���������
'������intType 0���ڲ�TabControl,1-�����TabControl
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long, idx As Long
    Dim strName As String
    
    '��¼���в˵���ʽ
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsMain.Count >= 2 Then
        idx = GetFirstCommandBar(cbsMain(2).Controls)
        If idx > 0 Then
            blnShowBar = cbsMain(2).Visible
            bytStyle = cbsMain(2).Controls(idx).Style
        End If
    End If
    
    'ˢ���Ӵ��ڲ˵�
    Call LockWindowUpdate(Me.hwnd)
    Call Get����վ�������(Me, objItem.Caption)
    'ɾ�����ڵĹ������������˵���
    For lngCount = cbsMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.Count To 2 Step -1
        cbsMain(lngCount).Delete
    Next
    
    '���������¼���
    Call MainDefCommandBar
    
    If Not mclsReg Is Nothing Then Call mclsReg.zlDefCommandBars(Me, Me.cbsMain, True)
    
    '�Ӵ������¼���
    Select Case objItem.Tag
    Case "����һ��", "����"
        '����һ��/������Ϣҳ ҳǩ���ü��ز˵�
    Case "ҽ��"
        Call mclsAdvices.zlDefCommandBars(Me, Me.cbsMain, 0, gobjPlugIn, gobjOneCardComLib)
    Case "����"
        Call mclsEPRs.zlDefCommandBars(Me.cbsMain)
    Case "�²���"
        Call mclsEMR.zlDefCommandBars(Me.cbsMain)
    Case "��������"
        Call mclsDisease.zlDefCommandBars(Me.cbsMain)
    Case Else
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strName = gobjPlugIn.GetButtomName(glngSys, p����ҽ��վ, mcolSubForm("_" & objItem.Tag), objItem.Tag)
            Call zlPlugInErrH(Err, "GetButtomName")
            '�����˵�
            If strName <> "" Then Call PlugInInSideBar(cbsMain, strName)
            Err.Clear: On Error GoTo 0
        End If
    End Select
    
    '�ָ����̶���һЩ�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
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
    
    '�������RecalcLayout����������
    Call LockWindowUpdate(0)
    
    Set mfrmActive = mcolSubForm("_" & objItem.Tag)
End Sub

Private Sub SubWinRefreshData(ByVal objItem As TabControlItem, Optional ByVal intType As Integer)
'���ܣ�ˢ���Ӵ������ݼ�״̬
'������intType 0���ڲ�TabControl,1-�����TabControl
Dim i As Integer, blnDis As Boolean
    If mlng����ID = 0 Or (mintActive = pt���� And mPatiInfo.�Һŵ� = mstr�Һŵ�) Then
        For i = 0 To tbcSub.ItemCount - 1 'Ĭ���������Ⱦ�����濨����ʾ
            If tbcSub.Item(i).Tag = "��������" Then
                blnDis = tbcSub.Item(i).Selected
                tbcSub.Item(i).Visible = False
                If blnDis Then '�����ǰѡ�е��Ǵ�Ⱦ�����濨����������ѡ�е�0��TAB
                    tbcSub.Item(0).Selected = True: Exit Sub
                End If
                Exit For
            End If
        Next

        '�����ԤԼ���ˣ����ξ���û��ҽ���Ͳ�������
        'Ҫ���Ӵ��尴�����ݴ������
        Select Case objItem.Tag
        Case "����"
            Call mobjPati.zlRefresh(0, IIf(mintActive = pt����, mPatiInfo.�Һ�ID, 0), False, False, , , mintActive, p����ҽ��վ)
        Case "ҽ��"
            Call mclsAdvices.zlRefresh(0, "", False, , , , , , , , p����ҽ��վ)
        Case "����"
            Call mclsEPRs.zlRefresh(0, 0, 0, False, False)
        Case "�²���"
            Call mclsEMR.zlRefresh(0, 0, 0, 0, 1)
        Case "��������"
            Call mclsDisease.zlRefresh(0, 0, 1, 0, False, False)
        Case "����һ��"
            Call mfrmView.zlRefresh(Me, 0, 0)
        Case Else
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                Call gobjPlugIn.RefreshForm(glngSys, p����ҽ��վ, mcolSubForm("_" & objItem.Tag), objItem.Tag, 0, "", 0, False)
                Call zlPlugInErrH(Err, "RefreshForm")
                Err.Clear: On Error GoTo 0
            End If
        End Select
    Else
        With mPatiInfo
            For i = 0 To tbcSub.ItemCount - 1 'Ĭ���������Ⱦ�����濨����ʾ
                If tbcSub.Item(i).Tag = "��������" Then
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
            Case "����"
                Call mobjPati.zlRefresh(.����ID, .�Һ�ID, Not tbcRegist.Item(mbyt���ξ���).Selected Or .���� = pt���� Or mstr�Һŵ� <> .�Һŵ�, .����ת��, , , mintActive, p����ҽ��վ)
            Case "ҽ��"
                Call mclsAdvices.zlRefresh(.����ID, .�Һŵ�, mstr�Һŵ� = .�Һŵ� And (.���� = pt���� Or .���� = pt����) And mlng����ID <> 0, .����ת��, , , Nothing, , , .����, p����ҽ��վ)
            Case "����"
                Call mclsEPRs.zlRefresh(.����ID, .�Һ�ID, mlng����ID, mstr�Һŵ� = .�Һŵ� And mlng����ID = .����ID And (.���� = pt���� Or .���� = pt����) And mlng����ID <> 0, .����ת��, True)
            Case "�²���"
                Call mclsEMR.zlRefresh(.����ID, .�Һ�ID, mlng����ID, .����, 1)
            Case "��������"
                If objItem.Visible Then
                    Call mclsDisease.zlRefresh(.����ID, .�Һ�ID, 1, mlng����ID, .����ת��, mstr�Һŵ� = .�Һŵ� And mlng����ID = .����ID And (.���� = pt���� Or .���� = pt����) And mlng����ID <> 0)
                End If
            Case "����һ��"
                Call mfrmView.zlRefresh(Me, .����ID, mlng����ID)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.RefreshForm(glngSys, p����ҽ��վ, mcolSubForm("_" & objItem.Tag), objItem.Tag, .����ID, .�Һŵ�, 0, .����ת��, 0, 0)
                    Call zlPlugInErrH(Err, "RefreshForm")
                    Err.Clear: On Error GoTo 0
                End If
            End Select
        End With
    End If
    Call SetFontSize(Not Me.Visible)
End Sub

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim strFunName As String

    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False) '����
    objMenu.ID = conMenu_FilePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��") '����
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_File_MedRec, "��ҳ��ӡ(&R)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_File_MedRecSetup, "��ӡ����(&S)", -1, False
            .Add xtpControlButton, conMenu_File_MedRecPreview, "��ӡԤ��(&V)", -1, False
            .Add xtpControlButton, conMenu_File_MedRecPrint, "��ӡ��ҳ(&P)", -1, False
        End With
        '56274
        Set objControl = .Add(xtpControlButton, conMenu_File_Print_Bespeak, "�ش�ԤԼ�Һŵ�(&P)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True '����
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "����(&C)", -1, False)
    objMenu.ID = conMenu_ManagePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conmenu_Edit_Leave, "���˲�����(&L)", -1, False): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conmenu_Edit_Wait, "���˴���(&W)", -1, False)
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Manage_Transfer, "ת�ﴦ��(&C)"): objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Receive, "���˽���(&Z)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Cancel, "ȡ������(&Q)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Finish, "��ɽ���(&O)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Redo, "�ָ�����(&R)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBack, "�����(&S)"): objControl.BeginGroup = True
        objControl.IconId = conMenu_Edit_Pause
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBackCancel, "ȡ������(&R)")
        objControl.IconId = conMenu_Edit_Reuse
 
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Green, "��ɫͨ��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_AdjustGrade, "�������鼶��")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False) '����
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)") '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)") '����
        objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FontSize, "�����С(&N)") '����
        With objPopup.CommandBar.Controls
             .Add xtpControlButton, conMenu_View_FontSize_S, "С����(&S)", -1, False '����(С�����ӦС��Ƭ���������Ӧ��Ƭ)
             .Add xtpControlButton, conMenu_View_FontSize_L, "������(&L)", -1, False '����
        End With
        objPopup.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conmenu_View_Leave, "��ʾ�����ﲡ��(&4)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "������һ��(&N)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Busy, "����æ(&M)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True '����
        Set objControl = .Add(xtpControlButton, conMenu_View_Jump, "������ת(&J)")
        
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", -1, False)
    objMenu.ID = conMenu_ToolPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Community, "��������(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_KssAudit, "������ҩ���(&K)")
        objControl.IconId = 3551
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_TransAudit, "��Ѫ��˹���(&M)")
        objControl.IconId = 3551
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "���Ӳ�������(&I)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_ExaReport, "��������ܼ챨��")
            objControl.IconId = conMenu_File_Preview
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "���ϲο�(&R)"): objPopup.BeginGroup = True
        objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_Reference_1, "������ϲο�(&D)", -1, False
            .Add xtpControlButton, conMenu_Tool_Reference_2, "���ƴ�ʩ�ο�(&C)", -1, False
        End With
        
        If gblnѪ��ϵͳ = True Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_TraReactionRecord, "��Ѫ��Ӧ��¼"): objControl.BeginGroup = True
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Positive, "���Խ��")
            objControl.IconId = 3551
        If mblnΣ��ֵ Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Critical, "Σ��ֵ")
                objControl.IconId = 4113
        End If
            
        On Error Resume Next
        If gobjOneCardComLib.zlHealthArchiveIsSHow(Me, p����ҽ��վ, strFunName, "") Then
            If Err.Number = 0 Then
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

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False) '����
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)") '����
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName) '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False '����
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False '����
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True '����
    End With
    
    '����������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��") '����
        Set objPopup = .Add(xtpControlPopup, conMenu_Manage_Transfer, "ת��")
        
        objPopup.ID = conMenu_Manage_Transfer
        objPopup.IconId = conMenu_Manage_Transfer
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Receive, "����")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Finish, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBack, "�����")
        objControl.IconId = conMenu_Edit_Pause
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItemQAdvice, "ҽ��")
        objControl.IconId = conMenu_Edit_NewItem
         
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "��������")
            objControl.ToolTipText = "���Ӳ�������"
        
        If strFunName <> "" Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_HealthCard, strFunName)
                objControl.ToolTipText = strFunName
                objControl.IconId = 3208
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����") '����
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�") '����
    End With
    
    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyH, conMenu_Manage_Regist '�Һ�
        .Add 0, vbKeyF7, conMenu_Manage_Receive '����
        .Add 0, vbKeyF8, conMenu_Manage_Finish '��ɾ���
        .Add FCONTROL, vbKeyB, conMenu_View_Busy '����״̬
        
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend 'չ��������
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '�۵�������
        .Add FCONTROL, vbKeyF, conMenu_View_Find '���Ҳ���
        .Add 0, vbKeyF3, conMenu_View_FindNext '������һ��
        .Add 0, vbKeyF12, conMenu_File_Parameter '��������
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        .Add 0, vbKeyF6, conMenu_View_Jump '��ת
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With
    
    '����һЩ�����Ĳ���������
    '-----------------------------------------------------
    With cbsMain.Options
'        .AddHiddenCommand conMenu_File_PrintSet '��ӡ����
'        .AddHiddenCommand conMenu_File_Excel '�����Excel
'        .AddHiddenCommand conMenu_View_Jump '��ת
    End With
    
    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "")
End Sub

Private Sub mclsAdvices_Activate()
    mblnUnRefresh = False
End Sub

Private Sub mclsAdvices_CheckInfectDisease(ByVal blnOnChek As Boolean, ByVal str����ID As String, ByVal str���ID As String, ByRef blnNo As Boolean)
'���ܣ���������뼲������ �õ������༭��
'      blnOnChek    �Ƿ�ֻ���д�Ⱦ�����濨��д���
'      str����ID    ����ID
'      str���ID   ���ID
'blnNO �Ƿ�Ҫ��д��Ⱦ�����濨
    Call mclsDisease.EditDiseaseDoc(Me, mlng����ID, mlng�Һ�ID, 1, mlng����ID, str����ID, str���ID, blnNo)
End Sub

Private Sub mclsAdvices_EditDiagnose(ParentForm As Object, ByVal �Һŵ� As String, Succeed As Boolean)
'���ܣ�Ҫ�������������
    Succeed = False
End Sub

Private Sub mclsAdvices_RequestRefresh()
'���ܣ�ҽ���Ӵ���Ҫ��ˢ��
    Call LoadPatients
End Sub

Private Sub mclsAdvices_StatusTextUpdate(ByVal Text As String)
'���ܣ�ҽ���Ӵ���Ҫ�����״̬��
    Me.stbThis.Panels(2).Text = Text
End Sub

Private Sub mclsAdvices_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
'���ܣ��鿴���Ӳ�������
    Call gobjRichEPR.ViewDocument(Me, ����ID, CanPrint)
End Sub

Private Sub mclsAdvices_PrintEPRReport(ByVal ����ID As Long, ByVal Preview As Boolean)
'���ܣ����༭��ʽ��ӡ����
    Call gobjRichEPR.PrintOrPreviewDoc(Me, cpr���Ʊ���, ����ID, Not Preview, True)
End Sub

Private Sub mclsAdvices_ViewPACSImage(ByVal ҽ��ID As Long)
'���ܣ�PACS��Ƭ����
    If CreateObjectPacs(gobjPublicPacs) Then
        Call gobjPublicPacs.ShowImage(ҽ��ID, Me, mPatiInfo.����ת��)
    End If
End Sub

Private Function CheckIsAskNextQueue(Optional strҵ��ID As String = "") As Boolean
   '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ҽ���Ƿ����������һ������
    '���ƣ����˺�
    '����:����,����true,���򷵻�False
    '���ڣ�2010-06-09 16:48:30
    '˵��������׼:��ʵ���Ѻ���Ϊ׼(ֻ����ɺ󣬲����ٽ�)(����:37442)
    '   ȡ��:��������(�������������)+�ѽ����+ת��<��������
    '------------------------------------------------------------------------------------------------------------------------
    Dim lngCount As Long
    Dim strResult As String, arrCheck As Variant
    
    On Error GoTo errH
    If Val(strҵ��ID) <> 0 Then
           strResult = ExseSvrQueuedatecheck(Val(strҵ��ID), p����ҽ��վ) & "|"
           arrCheck = Split(strResult, "|")
           If Val(arrCheck(0)) <> 0 Then
              If Val(arrCheck(0)) = 1 Then
                If MsgBox(CStr(arrCheck(1)) & vbCrLf & "�Ƿ����?", vbDefaultButton2 + vbYesNo + vbQuestion, Me.Caption) = vbNo Then
                    Exit Function
                End If
              Else
                 MsgBox CStr(arrCheck(1)), vbCritical, Me.Caption
                 Exit Function
              End If
              
           End If
    End If
    
    If mty_Queue.blnҽ���������� = False Or mty_Queue.int�������� <= 0 Then
        CheckIsAskNextQueue = True: Exit Function
    End If
    '0:�Ŷ��У�1:�����У�2�������ţ�3����ͣ��4����ɾ��6�����7���Ѻ���
    'mty_Queue.bln���к�����
    
    '����:44250
    lngCount = ExseSvrQueuecallcount(UserInfo.����, 1, p����ҽ��վ)

    If lngCount >= mty_Queue.int�������� Then
            MsgBox "���ֻ����" & mty_Queue.int�������� & "�����ﲡ��,�����ٽ��к��У�", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
    End If
    CheckIsAskNextQueue = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mobjEPRDoc_AfterSaved(lngRecordId As Long)
    With mPatiInfo
        Call mclsEPRs.zlRefresh(mlng����ID, mlng�Һ�ID, mlng����ID, mlng����ID = .����ID And (.���� = pt���� Or .���� = pt����) And mlng����ID <> 0, .����ת��, True)
    End With
End Sub

Private Sub mobjQueue_OnQueueExecuteAfter(ByVal strҵ��ID As String, ByVal byt�������� As Byte)
    '------------------------------------------------------------------------------------------------------------------------
    '��Σ�byt��������-0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
    '------------------------------------------------------------------------------------------------------------------------
    If mty_Queue.blnҽ���������� = False Then Exit Sub
    If byt�������� <> 1 Then Exit Sub
    
    '����ˢ�²�����Ϣ
    Call LoadPatients("10001")
End Sub

Private Sub mobjQueue_OnQueueExecuteBefore(ByVal strҵ��ID As String, ByVal byt�������� As Byte, blnCancel As Boolean, strNewQueueName As String)
    Dim colList As Collection
   ' byt�������� -0 - ����, 1 - ֱ��, 2 - ����, 3 - ��ͣ, 4 - ��ɾ���, 5 - �㲥
   
    On Error GoTo errH
    If InStr(1, "15", byt��������) = 0 Then Exit Sub
    If CheckIsAskNextQueue(strҵ��ID) = False Then blnCancel = True: Exit Sub
    
    Set colList = ExseSvrQueuereginfo(1, strҵ��ID, "0", "", "", p����ҽ��վ)
    
    If colList Is Nothing Then Exit Sub
    If colList.Count = 0 Then Exit Sub
    
    '68736:������,2014-02-18,ת�ﲡ��û��������Ϣ
    If byt�������� = 1 Then
        If Isת�ﲡ��(strҵ��ID) Then
            If GetColVal(colList(1), "_outp_room") = "" Then
                If Update���˹Һ�����(GetColVal(colList(1), "_reg_no"), Val(GetColVal(colList(1), "_pati_id")), mstr��������, UserInfo.����, Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS"), 2, "", p����ҽ��վ) = False Then Exit Sub
            End If
            Exit Sub
        End If
    End If
    
    If InStr(1, "12", Val(GetColVal(colList(1), "_exec_state"))) > 0 Then
        '1-��ɾ���,2-���ھ���:��Ҫ�ǵڶ��κ���
        'Ӧ����:����Ѿ������,ҽ�������,�в���ȥ����,�ٸ���������
        Exit Sub
    End If
    
    '��������_In Integer := 1
    If Update���˹Һ�����(GetColVal(colList(1), "_reg_no"), Val(GetColVal(colList(1), "_pati_id")), mstr��������, UserInfo.����, Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS"), 0, "", p����ҽ��վ) = False Then Exit Sub
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Isת�ﲡ��(strҵ��ID As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '����:���ò����Ƿ���ת�ﲡ�˲���δ����
    '���:strҵ��ID
    '����:True ����Ϊת�ﲡ�� False ����Ϊ��ͨ����
    '����:����
    '��������:2012-9-14
    '�����:51514
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As Recordset
    On Error GoTo errHand:
    strSql = _
    "   Select Count(ID) as �Ƿ�Ϊת�ﲡ�� From ���˹Һż�¼ Where ID=[1] And Nvl(ת�����ID,0) <> 0 And Nvl(ת��״̬,0)=0  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strҵ��ID)
    If rsTemp.EOF Then Isת�ﲡ�� = False
    Isת�ﲡ�� = rsTemp!�Ƿ�Ϊת�ﲡ�� > 0
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mobjQueue_OnRecevieDiagnose(ByVal strҵ��ID As String, ByVal lngҵ������ As Long)
    '����:
    Dim objControl As CommandBarControl
    Dim strNO As String, strSql As String, rsTmp As ADODB.Recordset
    Dim bln���� As Boolean, arrCheck As Variant, strResult As String
    Dim blnת�ﲡ�� As Boolean '�����:51514
    Dim datCurr As Date
    Dim i As Long, intOut As Integer, str��� As String
    Dim int�Һ�ģʽ As Integer
    Dim colPati As Collection, str���֤�� As String, str�������� As String, cllVisitInfo As Collection
    
    If lngҵ������ <> 0 Then Exit Sub
    On Error GoTo errH
     If Val(strҵ��ID) <> 0 Then
           strResult = ExseSvrQueuedatecheck(Val(strҵ��ID), p����ҽ��վ) & "|"
           arrCheck = Split(strResult, "|")
           If Val(arrCheck(0)) <> 0 Then
              If Val(arrCheck(0)) = 1 Then
                If MsgBox(CStr(arrCheck(1)) & vbCrLf & "�Ƿ����?", vbDefaultButton2 + vbYesNo + vbQuestion, Me.Caption) = vbNo Then
                    Exit Sub
                End If
              Else
                 MsgBox CStr(arrCheck(1)), vbCritical, Me.Caption
                 Exit Sub
              End If
           End If
    End If
    strSql = "Select ����ID,ִ����,NO,��¼��־,ִ��״̬,��¼����,����,�Ա�,����,�����,id as �Һ�id,����,���� From ���˹Һż�¼ Where  ID=[1]  "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strҵ��ID)
    If rsTmp.EOF Then
        MsgBox "�ò���û�йҺż�¼���ܽ��", vbInformation, gstrSysName
        Call LoadPatients("10101"): Exit Sub
    End If
    
    '�����:57566
    If Check�������("����", rsTmp!NO) = False Then Exit Sub
    
    '0-�ȴ�����,1-��ɾ���,2-���ھ���,-1���Ϊ������
    If Val(rsTmp!ִ��״̬) = 1 Then
        MsgBox "�ò����Ѿ���ɾ���,�����ٽ��о��������", vbInformation, gstrSysName
        Call LoadPatients("10101"): Exit Sub
    ElseIf Val(rsTmp!ִ��״̬) = -1 Then
        MsgBox "�ò����Ѿ����Ϊ������,�����ٽ��о��������", vbInformation, gstrSysName
        Call LoadPatients("10101"): Exit Sub
    End If
    strNO = Nvl(rsTmp!NO)
    
    'ת����� �����:51514
    blnת�ﲡ�� = Isת�ﲡ��(strҵ��ID)
    If blnת�ﲡ�� Then
        If Update���˹Һ�ת��(Val(Nvl(rsTmp!����ID)), strNO, 2, , , IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��), p����ҽ��վ) = False Then Exit Sub

        'ˢ�²���λ����
        If rptPati(PATI_RPT����).Visible Then
            Call LoadPatients("11001", PATI_RPT����, strNO)
        Else
            Call LoadPatients("11001")
        End If
    End If
    
    '����ԤԼ�Һŵ�
    datCurr = zlDatabase.Currentdate
    If Val("" & rsTmp!��¼����) = 2 Then
        Call InitObjPublicExpense
        int�Һ�ģʽ = Val(zlDatabase.GetPara("�Һ�ģʽ", glngSys, 9000, 1))
        If int�Һ�ģʽ = 0 And Not gobjPublicExpense Is Nothing Then
            If Not gobjPublicExpense.zlRegisterIncept(Me, mlngModul, strNO, mstr��������, 0, "") Then Exit Sub
        ElseIf int�Һ�ģʽ = 2 And Not gobjPublicExpense Is Nothing Then
            If ZLCommFun.ShowMsgBox("��ѡ��", "��ѡ���˵�֧����ʽ,����֧�����ߵ��շѴ���֧����", "!����֧��(&Y),?����֧��(&N)", Me, vbQuestion) = "����֧��" Then
                If Not gobjPublicExpense.zlRegisterIncept(Me, mlngModul, strNO, mstr��������, 0, "") Then Exit Sub
            Else
                 If Update����ԤԼ����(strNO, mstr��������, datCurr, IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��), IIf(mstr����ҽ����� = "", UserInfo.���, mstr����ҽ�����), p����ҽ��վ) = False Then Exit Sub
            End If
        Else
            If Update����ԤԼ����(strNO, mstr��������, datCurr, IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��), IIf(mstr����ҽ����� = "", UserInfo.���, mstr����ҽ�����), p����ҽ��վ) = False Then Exit Sub
        End If
    Else
       If mbln��Һ�ģʽ Then
            '���ж��ǻ��������ʾ�Ƿ��������
            If Val(rsTmp!ִ��״̬ & "") = 1 Or Val(rsTmp!ִ��״̬ & "") = 2 Then
                str��� = ZLCommFun.ShowMsgBox("��ѡ��", "��ǰ����Ϊ���ﲡ�ˣ���ѡ��ò��˾���ģʽ��", _
                        "!��������(&Y),��������(&N)", Me, vbQuestion)
            Else
                '���ж��Ƿ��չ���
                Call mclsReg.zlCheckRegisterNoIsCharge(strNO, intOut)
                'intOut=:-1-δ�ҵ���Ӧ�ĵ���,0-δ�շ�;1-�Һŵ�����;2-��Һ�ģʽ�£���δ�������ۼ�¼;
        '                      3-�Һŵ���Ӧ���շѻ��۵���ȫ�շ�(���ڶ��Ż��۵�ʱ������ȫ�յ�);
        '                      4-�Һŵ���Ӧ�Ļ��۵����ڲ����շ�)
                If intOut = 2 Then
                    str��� = "��������"
                End If
            End If
            If str��� = "��������" Then
                Set colPati = PatiSvrGetpatiinfo(0, Val(rsTmp!����ID), p����ҽ��վ)
                If Not colPati Is Nothing Then
                    If colPati.Count > 0 Then
                         str���֤�� = GetColVal(colPati(1), "_pati_idcard")
                         str�������� = GetColVal(colPati(1), "_pati_birthdate")
                    End If
                End If
                        
                Set cllVisitInfo = New Collection
                cllVisitInfo.Add Array("�������ID", IIf(Val(Nvl(rsTmp!ִ��״̬)) = 0, 0, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID)))
                cllVisitInfo.Add Array("��������", mstr��������)
                cllVisitInfo.Add Array("�����־", 0)
                cllVisitInfo.Add Array("�����־", IIf(Val(Nvl(rsTmp!ִ��״̬)) = 0, 0, 1))
                cllVisitInfo.Add Array("ִ��ʱ��", Format(datCurr, "yyyy-MM-dd HH:mm:ss"))
                
                If zlBulidingPriceDataFromRegistNo(mclsReg, strNO, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID), IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��), _
                    Val(rsTmp!����ID), rsTmp!���� & "", rsTmp!�Ա� & "", rsTmp!���� & "", str��������, str���֤��, "", cllVisitInfo) = False Then Exit Sub
            End If
        Else
            If Val(Nvl(rsTmp!ִ��״̬)) = 0 Then
                '�����ҺŽ���
                If Update���˹ҺŽ���(Val(rsTmp!����ID & ""), strNO, 0, UserInfo.����, mstr��������, 0, 0, datCurr, p����ҽ��վ) = False Then Exit Sub
            Else
                If Update���˹ҺŽ���(Val(Nvl(rsTmp!����ID)), strNO, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID), IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��), mstr��������, 0, 1, datCurr, p����ҽ��վ) = False Then Exit Sub
             
                bln���� = True
            End If
        End If
    End If
        
    mstr�Һŵ� = strNO
    mlng����ID = Val(Nvl(rsTmp!����ID))

    'ˢ�²���λ����
    On Error GoTo 0
    If rptPati(PATI_RPT����).Visible Then
        Call LoadPatients("11001", PATI_RPT����, strNO)
    Else
        Call LoadPatients("11001")
    End If
    '���������Զ����ù���
    If Not gobjCommunity Is Nothing And mlngCommunityID <> 0 And mlng����ID <> 0 And mPatiInfo.���� <> 0 Then
        Set objControl = cbsMain.FindControl(, mlngCommunityID, , True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then objControl.Execute
        End If
    End If
    
    Call ReceiveAfterExec(bln����)
    '�����ŶӽкŶ���(����ˢ��)
    Call ReshDataQueue
    Call SetReceiveToday(False, 1)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mobjQueue_OnSelectionChanged(ByVal blnIsCallingList As Boolean, objReportRow As Object, cbrMain As Object)
    If mty_Queue.blnҽ���������� Then
        mobjQueue.zlCommandBarSet 7, blnIsCallingList Or Not mbln���к����
    End If
     
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlOneCardComLib.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlOneCardComLib.clsPatientInfo, objCardData As zlOneCardComLib.clsPatientInfo, strErrMsg As String, blnCancel As Boolean)
    Dim lngPatiID As Long
    
    If objHisPati Is Nothing Then
        lngPatiID = 0
    Else
        lngPatiID = objHisPati.����ID
    End If
    
    Call ExecuteFindPati(False, , blnCard, lngPatiID)
End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlOneCardComLib.Card)
    If mblnIsInit Then mintFindType = Index: mstrFindType = objCard.����
End Sub

Private Sub rptNotify_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    '���뿪��������õ�ǰѡ���е���ɫ���ֲ���
    '��ʹ��������rptPati����me.activecontrolȴ�������������޷��жϵ�ǰ�Ƿ�����ʧȥ����ʱ
    If Me.Visible Then
        If rptNotify.SelectedRows.Count > 0 Then
            If Row.Index = rptNotify.SelectedRows(0).Index Then
                If rptNotify.Columns(Item.Index).Visible Then
                    '���¼��ᱻ�������Σ�δ���ԭ��
                    Metrics.BackColor = COLOR_RPTSelRow
                    Metrics.ForeColor = vbWhite
                End If
            End If
        End If
    End If
End Sub

Private Sub rptPati_BeforeDrawRow(Index As Integer, ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    '���뿪��������õ�ǰѡ���е���ɫ���ֲ���
    If Me.Visible Then
        If rptPati(Index).SelectedRows.Count > 0 Then
            If Row.Index = rptPati(Index).SelectedRows(0).Index Then
                If rptPati(Index).Columns(Item.Index).Visible Then
                    '��ʹ��������rptPati����me.activecontrolȴ�������������޷��жϵ�ǰ�Ƿ�����ʧȥ����ʱ
                    '���¼��ᱻ�������Σ�δ���ԭ��
                    Metrics.BackColor = COLOR_RPTSelRow
                    Metrics.ForeColor = vbWhite
                End If
            End If
        End If
    End If
End Sub

Private Sub rptPati_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    If Button = 2 And InStr(mstrPrivs, "���˽���") > 0 Then
        If mPr <> -1 And Index = mintRPTIndex Then
            Set objPopup = cbsMain.ActiveMenuBar.Controls(2)
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub rptPati_RowDblClick(Index As Integer, ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
'���ܣ�˫���Զ��������ɽ���
    Dim objControl As CommandBarControl
    
    If Index = PATI_RPT���� Or Index = PATI_RPTԤԼ Or Index = PATI_RPT���� Then
        If InStr(mstrPrivs, "���˽���") > 0 Then
            If Index = PATI_RPT���� Or Index = PATI_RPTԤԼ Then
                Set objControl = cbsMain.FindControl(, conMenu_Manage_Receive, True, True)
            ElseIf Index = PATI_RPT���� Then
                Set objControl = cbsMain.FindControl(, conMenu_Manage_Finish, True, True)
            End If
            If Not objControl Is Nothing Then
                If objControl.Enabled Then Call cbsMain_Update(objControl) '�״�ִ�У�û����ʾ�˵�ǰ���¼�û��ִ��
                If objControl.Enabled Then objControl.Execute
            End If
        End If
    End If
End Sub

Private Sub rptPati_SelectionChanged(Index As Integer)
    Call RptItemClick(Index)
End Sub
 
Private Sub tbcRegist_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'���ܣ�����ѡ��
'˵����tbcRegist.Tag �м�¼��һ�ο�Ƭ��ѡ�����
    Dim objControl As CommandBarControl
    
    If Item.Tag = "" Then Exit Sub '��ʼ��ʱ,��û��ֵ
    If Item.Tag = "����һ��" Then
        If tbcRegist.Tag <> "����һ��" And tbcSub.Selected.Tag <> "����" Then
            Call SubWinDefCommandBar(Item)
        Else
            Set mfrmActive = mcolSubForm("_" & Item.Tag)
            Call Get����վ�������(Me, Item.Tag)
        End If
        Call SubWinRefreshData(Item)
        Call UCPatiVitalSigns.ClearTxtToolTipText
        UCPatiVitalSigns.ControlLock = True
        UCPatiVitalSigns.TextBackColor = vbButtonFace
        Call UCPatiVitalSigns.SetUseType(False)
        If Visible Then mfrmActive.SetFocus
    ElseIf Item.Tag = "����" Then
        tbcRegist.Item(mbyt���ξ���).Selected = True
        Set objControl = cbsMain.FindControl(, conMenu_Tool_Archive, True, True) '���Ӳ�������
        If Not objControl Is Nothing Then
            If objControl.Enabled Then Call cbsMain_Update(objControl) '�״�ִ�У�û����ʾ�˵�ǰ���¼�û��ִ��
            If objControl.Enabled Then objControl.Execute
        End If
    Else
        If Val(Item.Tag) <> 0 Then
            If tbcRegist.Tag = "����һ��" And tbcSub.Selected.Tag <> "����" Then
                Call SubWinDefCommandBar(tbcSub.Selected)
            Else
                Set mfrmActive = mcolSubForm("_" & tbcSub.Selected.Tag)
                Call Get����վ�������(Me, tbcSub.Selected.Caption)
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

Private Sub LoadPatiInfo(ByVal lng�Һ�id As Long)
'���ܣ�ѡ��ĳ����ʷ�����¼ʱ����ȡ��صĲ�����Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim colPati As Collection, col���� As Collection
    Dim strSql As String
    Dim lngidx As Long
    Dim i As Long
    Dim str�ѱ� As String
    On Error GoTo errH

    strSql = "Select b.Id, b.����id, b.No, b.�����, b.����, b.�Ա�, b.����, b.�ѱ�, b.����, b.����ʱ��, b.ִ����, b.ִ��״̬, b.ִ��ʱ��, b.ִ�в���id As ����id, b.����," & vbNewLine & _
                "       b.����, c.���� As ����, b.����, b.ժҪ, b.����ʱ��, b.������ַ, b.��Ⱦ���ϴ�, b.����,b.ҽ�Ƹ��ʽ," & vbNewLine & _
                " Nvl(m.����,'δ��д') as ����,Nvl(m.������ʷ,'δ��д') as ������ʷ,Decode(m.�Ƿ񸴺���,1,'��',0,'��','δ��д') �Ƿ񸴺���,Nvl(m.��ע,'��') ��ע,m.�Ƿ���ɫͨ��,n.���� ���鼶�� " & _
                "From ���˹Һż�¼ B, ���ű� C,��������¼ m,���ﲡ�鼶�� n" & vbNewLine & _
                "Where b.Id = [1] And b.ִ�в���id = c.Id And B.id = m.�Һ�ID(+) And m.���鼶��=n.���(+)"
     
        '��ID��ȡ�Һż�¼�����üӼ�¼���ʡ�״̬������
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng�Һ�id)
    If rsTmp Is Nothing Then Exit Sub
    If rsTmp.EOF Then Exit Sub
    Set colPati = PatiSvrGetpatiinfo(3, Val(rsTmp!����ID & ""), p����ҽ��վ)
    Set col���� = PatiSvrGetCommunityInfo(1, Val(rsTmp!����ID & ""), Val(rsTmp!���� & ""), p����ҽ��վ)
    If colPati Is Nothing Then Exit Sub
    If colPati.Count = 0 Then Exit Sub

    For i = 0 To lblLink���
        lblLink(i).ForeColor = &HC00000
    Next
    
    lblLink(lblLink�޸�).ForeColor = IIf(InStr(GetInsidePrivs(9003), "������Ϣ����") > 0, &HC00000, &HC0C0C0)  'p������Ϣ��������
    
    Call ReadPatPricture(Val(rsTmp!����ID & ""), imgLoad)
    
    If imgLoad.Picture = 0 Then
        imgPatient.Picture = imgDefual.Picture
        picPatient.Tag = ""
    Else
        imgPatient.Picture = imgLoad.Picture
        picPatient.Tag = "��"
    End If
    
    txtInfo(txtInfo����).Text = rsTmp!���� & ""
    txtInfo(txtInfo����).ToolTipText = rsTmp!���� & ""
    '��ʾ������ɫ
    If Not Val(GetColVal(colPati(1), "_insurance_type")) = 0 And GetColVal(colPati(1), "_pati_type") = "" Then
        txtInfo(txtInfo����).ForeColor = &HC0&
    Else
        txtInfo(txtInfo����).ForeColor = GetPatiColor(GetColVal(colPati(1), "_pati_type"))
    End If
    
    txtInfo(txtInfo�Ա�).Text = rsTmp!�Ա� & ""
    txtInfo(txtInfo�Ա�).ToolTipText = txtInfo(txtInfo�Ա�).Text
    txtInfo(txtInfo����).Text = rsTmp!���� & ""
    txtInfo(txtInfo����).ToolTipText = txtInfo(txtInfo����).Text
    txtInfo(txtInfo��������).Text = Format(GetColVal(colPati(1), "_pati_birthdate"), "yyyy-MM-dd")
    txtInfo(txtInfo��������).ToolTipText = txtInfo(txtInfo��������).Text
    txtInfo(txtInfo����).Text = rsTmp!���� & ""
    txtInfo(txtInfo����).ToolTipText = txtInfo(txtInfo����).Text
    txtInfo(txtInfo���￨��).Text = GetColVal(colPati(1), "_vcard_no")
    txtInfo(txtInfo���￨��).ToolTipText = txtInfo(txtInfo���￨��).Text
    txtInfo(txtInfoҽ������).Text = GetColVal(colPati(1), "_insurance_num")
    txtInfo(txtInfoҽ������).ToolTipText = txtInfo(txtInfoҽ������).Text
    txtInfo(txtInfo������Ϣ).Text = "�����ߡ�" & rsTmp!���� & "����������ʷ��" & rsTmp!������ʷ & "���������ˡ�" & rsTmp!�Ƿ񸴺��� & "������ע��" & rsTmp!��ע
    txtInfo(txtInfoժҪ).Text = rsTmp!ժҪ & ""
    txtInfo(txtInfoժҪ).ToolTipText = txtInfo(txtInfoժҪ).Text
    If gobjPass Is Nothing Then
        lblPhysical.Caption = ""
        txtInfo(txtInfo������).Text = ""
        txtInfo(txtInfo������).ToolTipText = ""
    Else
        txtInfo(txtInfo������).Text = gobjPass.zlPassPatiPhysical(CLng(rsTmp!����ID), 0)
        txtInfo(txtInfo������).ToolTipText = txtInfo(txtInfo������).Text
        If txtInfo(txtInfo������).Text <> "" Then
            lblPhysical.Caption = "���������:"
        Else
            lblPhysical.Caption = ""
        End If
    End If
    txtInfo(txtInfo������).Visible = Not (lblPhysical.Caption = "")
    txtPhone.Text = GetColVal(colPati(1), "_phone_number")
    txtPhone.ToolTipText = txtPhone.Text
    
    str�ѱ� = IIf(rsTmp!�ѱ� & "" = "", GetColVal(colPati(1), "_fee_category"), rsTmp!�ѱ� & "")
    
    With cboBillType
        lngidx = -1
        For i = 0 To .ListCount
            If InStr(.List(i) & "", str�ѱ�) > 0 Then
                .ToolTipText = str�ѱ�
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
            If InStr(.List(i) & "-", "-" & rsTmp!ҽ�Ƹ��ʽ & "-") > 0 Then
                .ToolTipText = rsTmp!ҽ�Ƹ��ʽ & ""
                lngidx = i
                Exit For
            End If
        Next
    End With
    
    If lngidx <> -1 Then
        Call Cbo.SetIndex(cboPayType.hwnd, lngidx)
    End If
    
    With rsTmp
        '������Ϣ
        If mintActive = ptת�� Then
            mPatiInfo.���� = ptת��
        Else
            mPatiInfo.���� = Decode(Nvl(!ִ��״̬, 0), 0, 0, 2, 1, 1, 2)
        End If
        
        
        mPatiInfo.���� = rsTmp!���� & ""
        mPatiInfo.����� = Nvl(!�����)
        mPatiInfo.����ID = !����ID
        mPatiInfo.�Һ�ID = !ID
        mPatiInfo.�Һŵ� = !NO
        mPatiInfo.����ID = !����ID
        mPatiInfo.���� = Nvl(!����)
        mPatiInfo.���� = Nvl(!����, 0)
        mPatiInfo.������ = ""
        If Not col���� Is Nothing Then
            If col����.Count > 0 Then
                  mPatiInfo.������ = GetColVal(col����(1), "_community_code")
            End If
        End If
        
        mPatiInfo.�Һ�ʱ�� = !����ʱ��
        mPatiInfo.�Ա� = "" & !�Ա�
        mPatiInfo.����״�� = GetColVal(colPati(1), "_pati_marital_cstatus")
        
        mPatiInfo.���� = GetColVal(colPati(1), "_pati_nation")
        mPatiInfo.���� = GetColVal(colPati(1), "_country_name")
        mPatiInfo.���� = GetColVal(colPati(1), "_pati_area")
        mPatiInfo.�����ص� = GetColVal(colPati(1), "_pati_birthplace")
        mPatiInfo.��Ⱦ���ϴ� = Val("" & !��Ⱦ���ϴ�)
        mPatiInfo.��ͥ��ַ�ʱ� = GetColVal(colPati(1), "_pat_home_postcode")
        mPatiInfo.��λ�ʱ� = GetColVal(colPati(1), "_emp_postcode")
        mPatiInfo.����֤�� = GetColVal(colPati(1), "_cert_no_other")
        mPatiInfo.���ڵ�ַ = GetColVal(colPati(1), "_pat_hous_addr")
        mPatiInfo.���ڵ�ַ�ʱ� = GetColVal(colPati(1), "_pat_hous_postcode")
        mPatiInfo.���� = GetColVal(colPati(1), "_ntvplc_name")
        mPatiInfo.Email = GetColVal(colPati(1), "_pati_email")
        mPatiInfo.QQ = GetColVal(colPati(1), "_pati_qq")
        mPatiInfo.�Ƿ���ɫͨ�� = Val(!�Ƿ���ɫͨ�� & "")
        mPatiInfo.���鼶�� = !���鼶�� & ""
        lblRec.Visible = Val(GetColVal(colPati(1), "_balance_mode")) <> 0
  
        
        If mPatiInfo.���� = pt���� Then
            mPatiInfo.����ת�� = zlDatabase.NOMoved("���˹Һż�¼", !NO)
        Else
            mPatiInfo.����ת�� = False
        End If
    End With
    If mintRPTIndex = PATI_RPT���� Or mintRPTIndex = PATI_RPT���� Then
        If InStr("," & mstr�Һ�IDs & ",", "," & rsTmp!ID & ",") = 0 Then
            mstr�Һ�IDs = mstr�Һ�IDs & "," & rsTmp!ID
        End If
    End If
    If tbcRegist.Selected.Index = 0 And mbyt���ξ��� = 0 Or tbcRegist.Selected.Index = 1 And mbyt���ξ��� = 1 Then
        UCPatiVitalSigns.ControlLock = False
        UCPatiVitalSigns.TextBackColor = picMore.BackColor
        UCPatiVitalSigns.LblBackColor = picMore.BackColor
        Call UCPatiVitalSigns.SetUseType(True)
    Else
        Call UCPatiVitalSigns.ClearTxtToolTipText
        UCPatiVitalSigns.ControlLock = True
        UCPatiVitalSigns.TextBackColor = vbButtonFace
        UCPatiVitalSigns.LblBackColor = vbButtonFace
        Call UCPatiVitalSigns.SetUseType(False)
    End If
    Call UCPatiVitalSigns.LoadPatiVitalSigns(mPatiInfo.����ID, lng�Һ�id)
    Call UCPatiVitalSigns.TxtAlignment(2)
    txtInfo(txtInfoժҪ).Locked = True
    txtInfo(txtInfo������Ϣ).Locked = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadRegist(ByVal lng�Һ�id As Long)
'���ܣ�ѡ��ĳ����ʷ�����¼ʱ����ȡ��صĲ�����Ϣ

    If lng�Һ�id <= 0 Then
        '����ǰ�б�������ˢ���Ӵ���
        Call ClearPatiInfo
        'ˢ���Ӵ�������
        Call SubWinRefreshData(tbcSub.Selected)
        
        Exit Sub
    End If
    On Error GoTo errH
    Call LoadPatiInfo(lng�Һ�id)
    
    'ˢ���Ӵ�������
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
'���ܣ����֤ʶ��ɹ��󼤻�
    mstrIDCard = strID
    If mstrFindType = "�������֤" Then
        PatiIdentify.Text = mstrIDCard
    Else
        PatiIdentify.Text = "" '�������(Ŀǰ�������������²��ܼ���)��
    End If
    Call ExecuteFindPati(False, mstrIDCard)
End Sub

Private Function CheckHaveAdvice(ByVal lng����ID As Long, ByVal str�Һŵ� As String) As Boolean
'���ܣ��жϲ����Ƿ���ҽ��
    Dim strSql As String
    Dim rsTmp As Recordset
    
    On Error GoTo errH
    strSql = "select 1 from ����ҽ����¼ where ����ID=[1] and �Һŵ�=[2] and rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, str�Һŵ�)
    CheckHaveAdvice = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'���ܣ�ˢ���Ӵ�����漰����
'˵����������Ϊ�л����濨Ƭ����
    Dim objControl As CommandBarControl
    Dim Index As Long, objItem As TabControlItem
    
    If mblnTabTmp Then Exit Sub
    
    If Item.Tag = "" Then Exit Sub '��ʼ��ʱ,��û��ֵ
     
    If Item.Handle = picTmp.hwnd Then
        Index = Item.Index
        mblnTabTmp = True
        Screen.MousePointer = 11
        On Error GoTo errH
        Select Case Item.Tag
            Case "����"
                Set objItem = tbcSub.InsertItem(Index, "������Ϣ", mcolSubForm("_����").hwnd, 0)
                objItem.Tag = "����"
                mstrPreSubTab = "����"
            Case "����"
                Set objItem = tbcSub.InsertItem(Index, "������Ϣ", mcolSubForm("_����").hwnd, 0)
                objItem.Tag = "����"
            Case "�²���"
                Set objItem = tbcSub.InsertItem(Index, "���Ӳ���", mcolSubForm("_�²���").hwnd, 0)
                objItem.Tag = "�²���"
            Case "��������"
                Set objItem = tbcSub.InsertItem(Index, "��������", mcolSubForm("_��������").hwnd, 0)
                objItem.Tag = "��������"
        End Select
        Call tbcSub.RemoveItem(Index + 1)
        objItem.Selected = True
        Screen.MousePointer = 0
        mblnTabTmp = False
    End If
     
    'ˢ���Ӵ����Ӧ��CommandBar
    Call SubWinDefCommandBar(Item)
    
    'ˢ���Ӵ�������
    Call SubWinRefreshData(Item)
    
    If Visible Then mfrmActive.SetFocus
    
    '�Զ�����һ������/����/���ﲡ��/�����ҽ����������ҽ�������ж�û��ҽ��������
    If Item.Tag = "����" And mlng�Զ����� = 1 Then
        mblnUnRefresh = True
        Call mclsEPRs.zlOpenDefaultEPR(mstr�Һŵ�)
        '��Ϊִ��������Ƿ�ģ̬���壬������mclsAdvices��mclsEPRs��active������ mblnUnRefresh = False
    ElseIf Item.Tag = "ҽ��" And mlng�Զ����� = 2 Then
        If CheckHaveAdvice(mlng����ID, mstr�Һŵ�) = False Then
            cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
            Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
            If Not objControl Is Nothing Then
                If objControl.Enabled Then objControl.Execute
            End If
        End If
    End If
    
    If mstrPreSubTab = "����" And Not mobjPati Is Nothing Then
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
    rptPati(PATI_RPT����).Top = cboSelectTime.Top + cboSelectTime.Height + 30
    rptPati(PATI_RPT����).Left = 0
    rptPati(PATI_RPT����).Width = picYZ.Width
    lngTmp = picYZ.Height - rptPati(PATI_RPT����).Top
    If mbytSize = 0 Then
        If lngTmp < 1010 Then
            lngTmp = 1010
        End If
    Else
        If lngTmp < 1130 Then
            lngTmp = 1130
        End If
    End If
    rptPati(PATI_RPT����).Height = lngTmp
End Sub

Private Sub RPTResize(ByVal objC As Object, ByVal lngID As Long)
'���ܣ����ñ��ؼ���С
'������objC �ϲ������ؼ�  lngId 0-���1-���2-��Ϣ��3-���4-ԤԼ
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
       
    timRefresh.Enabled = False
    If Not mobjMsg Is Nothing Then
        Call mobjMsg.CloseAirBubble
    End If
    
    mblnMsgOk = False: mblnFirstMsg = False
    mblnIsInit = False 'PatiIdentify��ʼ����־
    
    blnSetup = InStr(";" & mstrPrivs & ";", ";��������;") > 0
    Call zlDatabase.SetPara("���˲��ҷ�ʽ", mintFindType, glngSys, p����ҽ��վ, blnSetup)

    If Not tbcSub.Selected Is Nothing Then
        Call zlDatabase.SetPara("ҽ������", tbcSub.Selected.Tag, glngSys, p����ҽ��վ, blnSetup)
    End If
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, dkpMain.SaveStateToString)
    End If

    '���������̶�����һ���ؼ�����ʽ���棬����վ���������һ���Ǵ�ӡ����̶���ͼ����ʽ,������ָ�Ϊ������ť����ʽ
    If Me.Visible Then  'Form_load���˳�ʱ������
        cbsMain(2).Controls(1).Style = cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style
        Call SaveWinState(Me, App.ProductName)
    End If
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
    '������һ�Σ��õĿؼ�����
        For i = 0 To rptPati.Count - 1
            Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\ReportControl", "rptPati" & "_" & i, rptPati(i).SaveSettings)
        Next
    End If
    
    mstrIDCard = ""
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled False
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    Set mobjPeis = Nothing

    '--�ر������ŶӵĴ���
    If Not mobjQueue Is Nothing Then
        Call mobjQueue.CloseWindows
        Set mobjQueue = Nothing
    End If
    Set mobjQueueList = Nothing
    Set mobjCallList = Nothing

    'ǿ��Unload,��Ȼ���ἤ���Ӵ�����¼�
    For i = 1 To mcolSubForm.Count
        If Not mcolSubForm(i) Is Nothing Then
            Unload mcolSubForm(i)
        End If
    Next
    Set mclsAdvices = Nothing
    Set mclsEMR = Nothing
    
    Set mclsEPRs = Nothing
    Set mrsAller = Nothing
    Set mobjEPRDoc = Nothing
    If Not mfrmActive Is Nothing Then
        Unload mfrmActive
    End If
    Set mfrmActive = Nothing
    Set gobjPublicPacs = Nothing
    Set gobjOneCardComLib = Nothing
    Set gobjPublicExpense = Nothing
    Set gobjService = Nothing
    
    If Not mfrmView Is Nothing Then
        Unload mfrmView
    End If
    Set mfrmView = Nothing
    Set mclsReg = Nothing
    mPatiInfo.�Һ�ID = 0
    '�����:57566
    mlng������� = 0
    mlng��ǰ����ʱ�� = 0
    
    Set mobjMsg = Nothing
    Set mobjPatient = Nothing
    Set mclsDis = Nothing
    Set mclsMsg = Nothing
    Set mrsMsg = Nothing
    mblnΣ��ֵ = False
    Set mclsDisease = Nothing
    
    If Not mobjPati Is Nothing Then
        Unload mobjPati
    End If
    Set mobjPati = Nothing

    If Not mobjDocMsg Is Nothing Then
        mobjDocMsg.isUnload = True
        Unload mobjDocMsg
        Set mobjDocMsg = Nothing
    End If
End Sub

Private Sub lblRoom_Click()
    Call SetRoomState(lblRoom.BackColor = COLOR_FREE)
End Sub

Private Sub RptItemClick(ByVal Index As Integer)
'����:���ݴ����б��в��˵��л�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim intCount As Integer
    Dim i As Long, j As Long
    Dim strNO As String
    Dim str����ʱ�� As String
    Dim lng�Һ�id As Long
    Dim strCaption As String
    Dim strRegTag As String
    Dim blnDo As Boolean
    Dim str���֤�� As String
    Dim strTmp As String
    Dim str����ids As String
    Dim varPar(0 To 10) As String '����һ���ַ�������
    Dim n As Long, p As Long
    Dim strTemp As String, strSQLPati As String, strThis As String
    Dim colPati As Collection
        
    On Error GoTo errH
   
    If rptPati(Index).SelectedRows.Count <= 0 Then Exit Sub
    
    blnDo = False
    With rptPati(Index).SelectedRows(0).Record
        Select Case Index
        Case PATI_RPT����, PATI_RPTԤԼ
            strNO = .Item(COL_HZ_NO).Value
            mstr�Һŵ� = strNO
            mlng����ID = Val(.Item(COL_HZ_����ID).Value)
            mlng����ID = Val(.Item(COL_HZ_ִ�в���ID).Value)
            str����ʱ�� = .Item(COL_HZ_����ʱ��).Value
            str���֤�� = .Item(COL_HZ_���֤��).Value
            
            strNO = .Item(COL_HZ_��ʶ).Value
            If strNO = "Ԥ" Then
                mintActive = ptԤԼ
            ElseIf strNO = "ת" Then
                mintActive = ptת��
            Else
                mintActive = pt����
            End If
             
        Case PATI_RPT����, PATI_RPT����
            strNO = .Item(COL_JZ_NO).Value
            mstr�Һŵ� = strNO
            mlng����ID = Val(.Item(COL_JZ_����ID).Value)
            mlng����ID = Val(.Item(COL_JZ_ִ�в���ID).Value)
            str����ʱ�� = .Item(COL_JZ_����ʱ��).Value
            str���֤�� = .Item(COL_JZ_���֤��).Value
            If Index = PATI_RPT���� Then
                mintActive = pt����
            Else
                mintActive = pt����
            End If
        Case PATI_RPT����
            strNO = .Item(COL_YZ_NO).Value
            mstr�Һŵ� = strNO
            mlng����ID = Val(.Item(COL_YZ_����ID).Value)
            mlng����ID = Val(.Item(COL_YZ_ִ�в���ID).Value)
            str����ʱ�� = .Item(COL_YZ_����ʱ��).Value
            str���֤�� = .Item(COL_YZ_���֤��).Value
            mintActive = pt����
        End Select
        mintRPTIndex = Index
        mPr = rptPati(Index).SelectedRows(0).Index
    End With
        
    For i = 0 To 4
        If i = Index Then
            rptPati(i).PaintManager.CaptionBackColor = COLOR_RPTHeadBackSel
        Else
            rptPati(i).PaintManager.CaptionBackColor = COLOR_RPTHeadBack
            If rptPati(i).SelectedRows.Count > 0 Then rptPati(i).SelectedRows.DeleteAll
        End If
    Next
        
    If mstr�Һŵ� = mstrPrePati Then Exit Sub
    mstrPrePati = mstr�Һŵ�
            
    LockWindowUpdate Me.hwnd
    
    '��֤���֤��
    If str���֤�� <> "" Then
        If mobjPatient Is Nothing Then
            On Error Resume Next
            Set mobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
            Err.Clear: On Error GoTo 0
            If mobjPatient Is Nothing Then
                MsgBox "����������Ϣ����������zlPublicPatient.clsPublicPatient��ʧ�ܣ�", vbInformation, Me.Caption
            Else
                Call mobjPatient.zlInitCommon(gcnOracle, glngSys, UserInfo.�û���)
            End If
        End If
        strTmp = ""
        If Not mobjPatient Is Nothing Then
            If mobjPatient.CheckPatiIdcard(str���֤��) Then
                strTmp = str���֤��
            End If
        End If
        str���֤�� = strTmp
    End If
    
    On Error GoTo errH

    'ͨ�����֤�Ų����
    If str���֤�� <> "" Then
        str����ids = PatiSvrGetpatirelate(1, mlng����ID, str���֤��, p����ҽ��վ)
    End If
    
    'ͨ������������
    strTmp = GetPatiRelate(mlng����ID, str���֤��, p����ҽ��վ)
    If strTmp <> "" Then
        If str����ids <> "" Then
            str����ids = str����ids & "," & strTmp
        Else
            str����ids = strTmp
        End If
    End If
    
    If str����ids = "" Then
        '��ǰ�ĵ�����ģʽ
        '��ȡ"��ʷ��"�����¼
        strSql = "Select A.ID,A.NO,A.����ʱ�� as ʱ��,B.���� as ����,a.ִ����,a.����ʱ�� From ���˹Һż�¼ A,���ű� B" & _
            " Where A.ִ�в���ID=B.ID And A.����ID=[1] And A.����ʱ��<=[2] And A.��¼����=1 And A.��¼״̬=1 Order by A.����ʱ�� Desc,a.����ʱ�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, CDate(str����ʱ��))
    Else
        str����ids = mlng����ID & "," & str����ids
    
        '����4000���ȵĲ��
        strTemp = "Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X"
        n = 0
        Do While True
            If Len(str����ids) < 4000 Then
                p = Len(str����ids) + 1
            Else
                p = InStrRev(Mid(str����ids, 1, 4000), ",")
            End If
            strThis = Mid(str����ids, 1, p - 1)
            
            If n > 10 Then
                strSQLPati = strSQLPati & vbNewLine & " Union All " & Replace(strTemp, "[1]", "'" & strThis & "'")
            Else
                varPar(n) = strThis
                strSQLPati = IIf(strSQLPati = "", "", strSQLPati & vbNewLine & " Union All ") & Replace(strTemp, "[1]", "[" & (n + 2) & "]")
            End If
            
            n = n + 1
            str����ids = Mid(str����ids, p + 1)
            If str����ids = "" Then Exit Do
        Loop
        strTmp = " A.����ID In (" & strSQLPati & ")"

        strSql = "Select A.ID,A.NO,A.����ʱ�� as ʱ��,B.���� as ����,a.ִ����,a.����ʱ�� From ���˹Һż�¼ A,���ű� B" & _
                        " Where A.ִ�в���ID=B.ID And " & strTmp & _
                        " And A.����ʱ��<=[1] And A.��¼����=1 And A.��¼״̬=1 Order by ʱ�� Desc,����ʱ�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CDate(str����ʱ��), varPar(0), varPar(1), varPar(2), varPar(3), varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9), varPar(10))
    End If
    
    strRegTag = tbcRegist.Selected.Caption
    
    '�����ؿ�Ƭ
    For i = mbyt���ξ��� + 1 To tbcRegist.ItemCount - 1
        tbcRegist.Item(i).Visible = False
    Next
    
    i = 0: blnDo = False
    For j = 1 To rsTmp.RecordCount
        If Not blnDo Then
            If mstr�Һŵ� = rsTmp!NO & "" Then
                mlng�Һ�ID = Val(rsTmp!ID & "")
                blnDo = True
            End If
        End If
        If blnDo Then
            i = i + 1
            If i = 5 Then
                tbcRegist.Item(mbyt���ξ��� + i - 1).Visible = True
            ElseIf i > 1 And i < 5 Then
                strCaption = Format(rsTmp!ʱ��, "YYMMdd") & "/" & rsTmp!���� & "/" & rsTmp!ִ����
                tbcRegist.Item(mbyt���ξ��� + i - 1).Caption = strCaption
                tbcRegist.Item(mbyt���ξ��� + i - 1).Tag = Val(rsTmp!ID & "")
                tbcRegist.Item(mbyt���ξ��� + i - 1).Visible = True
            End If
            If rsTmp!NO = mstr�Һŵ� Then
                lng�Һ�id = Val(rsTmp!ID & "")
            Else
                If lng�Һ�id = 0 Then
                    lng�Һ�id = Val(rsTmp!ID & "")
                End If
            End If
            '���ն�ƾ���
            If Format(rsTmp!ʱ��, "yyyy-MM-dd") = Format(str����ʱ��, "yyyy-MM-dd") Then
                intCount = intCount + 1
            End If
        End If
        rsTmp.MoveNext
    Next
   
    If strRegTag = "����һ��" Then
        tbcRegist.Item(0).Selected = True
        tbcRegist.Item(0).Tag = "����һ��"
        Call LoadPatiInfo(lng�Һ�id)
    Else
        If mbyt���ξ��� = 1 Then
            tbcRegist.Item(0).Tag = "����һ��"
        End If
        tbcRegist.Item(mbyt���ξ���).Tag = ""
        tbcRegist.Item(mbyt���ξ���).Selected = True
        tbcRegist.Item(mbyt���ξ���).Tag = mlng�Һ�ID
    End If

    Call tbcRegist_SelectedChanged(tbcRegist.Selected)
    
    lblMore.Visible = intCount > 1 And mintActive = pt����
    
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
    
    If Not mfrmActive Is Nothing Then
        If mfrmActive.Visible Then mfrmActive.SetFocus
    End If
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngidx As Long, i As Long

    For i = 0 To 1
        If i = 0 Then
            lngidx = PATI_RPT����
        Else
            lngidx = PATI_RPTԤԼ
        End If
        With rptPati(lngidx)
            Set objCol = .Columns.Add(COL_HZ_��ʶ, "", 18, True)
            Set objCol = .Columns.Add(COL_HZ_����, "����", 30, True)
            Set objCol = .Columns.Add(COL_HZ_�����, "�����", 74, True)
            Set objCol = .Columns.Add(COL_HZ_����, "����", 50, True)
            Set objCol = .Columns.Add(COL_HZ_�Һ�ʱ��, "�Һ�ʱ��", 80, True)
            Set objCol = .Columns.Add(COL_HZ_�Ա�, "�Ա�", 30, True)
            Set objCol = .Columns.Add(COL_HZ_����, "����", 40, True)
            Set objCol = .Columns.Add(COL_HZ_��ɫͨ��, "��ɫͨ��", 20, True)
            Set objCol = .Columns.Add(COL_HZ_��, "��", 20, True)
            Set objCol = .Columns.Add(COL_HZ_NO, "�Һŵ�", 60, True)
            Set objCol = .Columns.Add(COL_HZ_��������, "��������", 60, True)
            Set objCol = .Columns.Add(COL_HZ_����ҽ��, "����ҽ��", 60, True)
            Set objCol = .Columns.Add(COL_HZ_���, "���", 30, True)
            Set objCol = .Columns.Add(COL_HZ_����ʱ��, "����ʱ��", 80, True)
            Set objCol = .Columns.Add(COL_HZ_��������, "��������", 60, True)
            Set objCol = .Columns.Add(COL_HZ_ת��״̬, "ת��״̬", 60, True)
            Set objCol = .Columns.Add(COL_HZ_����, "����", 30, True)
            Set objCol = .Columns.Add(COL_HZ_���˿���, "�������", 60, True)
            
            If lngidx = PATI_RPT���� Then
                Set objCol = .Columns.Add(COL_HZ_ԤԼҽ��, "", 0, False): objCol.Visible = False
                Set objCol = .Columns.Add(COL_HZ_ԤԼʱ��, "", 0, False): objCol.Visible = False
            Else
                Set objCol = .Columns.Add(COL_HZ_ԤԼҽ��, "ԤԼҽ��", 60, True)
                Set objCol = .Columns.Add(COL_HZ_ԤԼʱ��, "ԤԼʱ��", 80, True)
            End If
            
            Set objCol = .Columns.Add(COL_HZ_����ID, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_����ʱ��, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_ִ�в���ID, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_ִ����, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_״̬, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_IC����, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_���￨��, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_���֤��, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_��¼��־, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_ִ��״̬, "", 0, False): objCol.Visible = False
            
            
            With .PaintManager
                .ColumnStyle = xtpColumnFlat
                .MaxPreviewLines = 1
                .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
                .GroupForeColor = &HC00000
                .BackColor = COLOR_Back
                .CaptionBackColor = COLOR_RPTHeadBack
                .HighlightBackColor = COLOR_RPTSelRow  '��ǰѡ���е���ɫ
                .HideSelection = True   'ʧȥ����󣬲�����ǰѡ��������Ϊ��ɫ��
                
                .GridLineColor = RGB(255, 224, 192)
                .VerticalGridStyle = xtpGridSolid
                .NoGroupByText = "�϶��б��⵽����,�����з���..."
                .NoItemsText = "û�п���ʾ�Ĳ���..."
            End With
            .PreviewMode = True
            .AllowColumnRemove = False
            .MultipleSelection = False '������SelectionChanged�¼�
            .ShowItemsInGroups = False
            .SetImageList Me.imgPati
        End With
    Next
    
    
    For i = 0 To 1
        If i = 0 Then
            lngidx = PATI_RPT����
        Else
            lngidx = PATI_RPT����
        End If
        With rptPati(lngidx)
            Set objCol = .Columns.Add(COL_JZ_��ʶ, "", 18, True)
            Set objCol = .Columns.Add(COL_JZ_����, "����", 30, True)
            Set objCol = .Columns.Add(COL_JZ_�����, "�����", 74, True)
            Set objCol = .Columns.Add(COL_JZ_����, "����", 50, True)
            Set objCol = .Columns.Add(COL_JZ_����ʱ��, "����ʱ��", 80, True)
            Set objCol = .Columns.Add(COL_JZ_�Ա�, "�Ա�", 30, True)
            Set objCol = .Columns.Add(COL_JZ_����, "����", 40, True)
            Set objCol = .Columns.Add(COL_JZ_��ɫͨ��, "��ɫͨ��", 20, True)
            Set objCol = .Columns.Add(COL_JZ_��, "��", 20, True)
            Set objCol = .Columns.Add(COL_JZ_NO, "�Һŵ�", 60, True)
            Set objCol = .Columns.Add(COL_JZ_���￨��, "���￨��", 60, True)
            Set objCol = .Columns.Add(COL_JZ_��������, "��������", 60, True)
            Set objCol = .Columns.Add(COL_JZ_ת��״̬, "ת��״̬", 60, True)
            Set objCol = .Columns.Add(COL_JZ_��Ⱦ��, "��Ⱦ��", 60, True)
            Set objCol = .Columns.Add(COL_JZ_����, "����", 30, True)
            Set objCol = .Columns.Add(COL_JZ_���˿���, "�������", 60, True)
            
            Set objCol = .Columns.Add(COL_JZ_����ID, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_JZ_����ʱ��, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_JZ_ִ�в���ID, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_JZ_ִ����, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_JZ_״̬, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_JZ_���֤��, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_JZ_IC����, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_JZ_��¼��־, "", 0, False): objCol.Visible = False
            
            With .PaintManager
                .ColumnStyle = xtpColumnFlat
                .MaxPreviewLines = 1
                .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
                .GroupForeColor = &HC00000
                .BackColor = COLOR_Back
                .CaptionBackColor = COLOR_RPTHeadBack
                .HighlightBackColor = COLOR_RPTSelRow  '��ǰѡ���е���ɫ
                .HideSelection = True   'ʧȥ����󣬲�����ǰѡ��������Ϊ��ɫ��
                
                .GridLineColor = RGB(255, 224, 192)
                .VerticalGridStyle = xtpGridSolid
                .NoGroupByText = "�϶��б��⵽����,�����з���..."
                .NoItemsText = "û�п���ʾ�Ĳ���..."
            End With
            .PreviewMode = True
            .AllowColumnRemove = False
            .MultipleSelection = False '������SelectionChanged�¼�
            .ShowItemsInGroups = False
            .SetImageList Me.imgPati
        End With
    Next
    
    With rptPati(PATI_RPT����)
        Set objCol = .Columns.Add(COL_YZ_��ʶ, "", 18, True)
        Set objCol = .Columns.Add(COL_YZ_����, "����", 30, True)
        Set objCol = .Columns.Add(COL_YZ_�����, "�����", 74, True)
        Set objCol = .Columns.Add(COL_YZ_����, "����", 50, True)
        Set objCol = .Columns.Add(COL_YZ_����ʱ��, "����ʱ��", 120, True)
        Set objCol = .Columns.Add(COL_YZ_�Ա�, "�Ա�", 30, True)
        Set objCol = .Columns.Add(COL_YZ_����, "����", 40, True)
        Set objCol = .Columns.Add(COL_YZ_��ɫͨ��, "��ɫͨ��", 20, True)
        Set objCol = .Columns.Add(COL_YZ_��, "��", 20, True)
        Set objCol = .Columns.Add(COL_YZ_NO, "�Һŵ�", 60, True)
        Set objCol = .Columns.Add(COL_YZ_����ҽ��, "����ҽ��", 60, True)
        Set objCol = .Columns.Add(COL_YZ_���￨��, "���￨��", 60, True)
        Set objCol = .Columns.Add(COL_YZ_��������, "��������", 60, True)
        Set objCol = .Columns.Add(COL_YZ_����, "����", 30, True)
        Set objCol = .Columns.Add(COL_YZ_���˿���, "�������", 60, True)
        Set objCol = .Columns.Add(COL_YZ_��ҽ���, "��ҽ���", 120, True)
        Set objCol = .Columns.Add(COL_YZ_��ҽ���, "��ҽ���", 120, True)
        
        Set objCol = .Columns.Add(COL_YZ_����ID, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_YZ_����ʱ��, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_YZ_ִ�в���ID, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_YZ_ִ����, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_YZ_���֤��, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_YZ_IC����, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_YZ_��¼��־, "", 0, False): objCol.Visible = False
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .BackColor = COLOR_Back
            .CaptionBackColor = COLOR_RPTHeadBack
            .HighlightBackColor = COLOR_RPTSelRow  '��ǰѡ���е���ɫ
            .HideSelection = True   'ʧȥ����󣬲�����ǰѡ��������Ϊ��ɫ��
                
            .GridLineColor = RGB(255, 224, 192)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�Ĳ���..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList Me.imgPati
    End With
    
    '��Ϣ����
    With rptNotify
        Set objCol = .Columns.Add(c_ͼ��, "", 18, True): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(C_����ID, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_No, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_�����, "�����", 74, True)
        Set objCol = .Columns.Add(c_����, "����", 50, True)
        Set objCol = .Columns.Add(C_����ʱ��, "����ʱ��", 60, True)
        Set objCol = .Columns.Add(C_״̬, "״̬", 150, True)
         
        Set objCol = .Columns.Add(C_��Ϣ, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_���, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_����, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_ҵ��, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_�Һ�Id, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_Id, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_������Ϣ, "", 0, False): objCol.Visible = False
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
            If objCol.Index <> C_��� Or objCol.Index <> C_���� Then objCol.Sortable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GroupForeColor = &HC00000
            .BackColor = COLOR_Back
            .CaptionBackColor = COLOR_RPTHeadBack
            .HighlightBackColor = COLOR_RPTSelRow  '��ǰѡ���е���ɫ
            .HideSelection = True   'ʧȥ����󣬲�����ǰѡ��������Ϊ��ɫ��
            
            .GridLineColor = RGB(255, 224, 192)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û����������..."
        End With
        .PreviewMode = False
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgPati
        
        '���� ����
        .SortOrder.Add .Columns(C_���)
        .SortOrder(0).SortAscending = False
        .SortOrder.Add .Columns(C_����)
        .SortOrder(1).SortAscending = False
    End With
    
End Sub

Private Sub InitCondFilter()
    Dim curDate As Date, intDay As Long
    Dim intStart As Long
    
    cboSelectTime.Clear
    
    With cboSelectTime
        .AddItem "����"
        .ItemData(.NewIndex) = 0
        .AddItem "����(������)"
        .ItemData(.NewIndex) = 1
        .AddItem "һ����"
        .ItemData(.NewIndex) = 7
        .AddItem "[ָ��...]"
        .ItemData(.NewIndex) = -1
    End With
    
    '���ﲡ��ʱ�䷶Χ
    curDate = zlDatabase.Currentdate
    
    intStart = Val(zlDatabase.GetPara("���ﲡ�˽������", glngSys, p����ҽ��վ, "0", Array(lblSeeTim, cboSelectTime), InStr(";" & mstrPrivs & ";", ";��������;") > 0))
    If lblSeeTim.ForeColor <> vbBlue Then
        '˽�в���
        mvCondFilter.End = Format(curDate, "yyyy-MM-dd 23:59:59")
        mvCondFilter.Begin = Format(mvCondFilter.End, "yyyy-MM-dd 00:00:00")
        If cboSelectTime.ListCount > 0 Then cboSelectTime.ListIndex = 0
    Else
        'ϵͳ����(�ָ��ɹ���Ա���õ�ֵ����ֹͨ��)
        mvCondFilter.End = Format(curDate + intStart, "yyyy-MM-dd 23:59:59")
        intDay = Val(zlDatabase.GetPara("���ﲡ�˿�ʼ���", glngSys, p����ҽ��վ, "7", Array(lblSeeTim, cboSelectTime), InStr(";" & mstrPrivs & ";", ";��������;") > 0))
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
    
    'ȱʡҽ������
    mvCondFilter.ҽ�� = UserInfo.����
    
    '������ȱʡ
    mvCondFilter.���� = ""
    mvCondFilter.�ı� = ""
    mvCondFilter.����ID = 0
    mvCondFilter.����ID = 0
End Sub

Private Sub GetLocalSetting()
'���ܣ���ע����ȡ��Ժ���˵�ʱ�䷶Χ
    '���ﷶΧ��1=�ұ��˺ŵĲ���,2=�����Ҳ���,3=�����Ҳ���
    Dim strSql As String, rsTmp As Recordset, intType As Integer
    Dim str���˽������ As String '�����:57566
    
    mint���ﷶΧ = Val(zlDatabase.GetPara("���ﷶΧ", glngSys, p����ҽ��վ, "2"))
    mstr�������� = zlDatabase.GetPara("��������", glngSys, p����ҽ��վ)
    mlng�������ID = Val(zlDatabase.GetPara("�������", glngSys, p����ҽ��վ))
    On Error GoTo errH
    strSql = "Select Distinct B.ID,B.����,B.����,A.ȱʡ" & _
        " From ������Ա A,���ű� B,��������˵�� C" & _
        " Where A.����ID=B.ID And B.ID=C.����ID And C.������� In(1,3) And C.��������='�ٴ�'" & _
        " And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is Null)" & _
        " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) And A.��ԱID=[1] And b.ID=[2]" & _
        " Order by B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID, mlng�������ID)
    If rsTmp.RecordCount = 0 Then mlng�������ID = 0
    mblnҪ����� = Val(zlDatabase.GetPara("ֻ�����Ѿ�����Ĳ���", glngSys, p����ҽ��վ)) <> 0
    
    '���ﲡ��
    If InStr(mstrPrivs, "���ﲡ��") > 0 Then
        mstr����ҽ�� = zlDatabase.GetPara("����ҽ��", glngSys, p����ҽ��վ, UserInfo.����)
        If mstr����ҽ�� <> "" Then
            mstr����ҽ����� = Sys.RowValue("��Ա��", mstr����ҽ��, "���", "����")
        End If
    Else
        mstr����ҽ�� = UserInfo.����
        mstr����ҽ����� = UserInfo.���
    End If
    
    '�Զ�������
    mbln�Զ����� = Val(zlDatabase.GetPara("�ҵ����˺��Զ�����", glngSys, p����ҽ��վ)) <> 0
    mlng�Զ����� = Val(zlDatabase.GetPara("������Զ�����", glngSys, p����ҽ��վ))
    
    'ҽ���������к���������
    mbln���к���� = Val(zlDatabase.GetPara("ҽ���������к���������", glngSys, p����ҽ��վ)) <> 0
    '��������
    mbytSize = zlDatabase.GetPara("����", glngSys, p����ҽ��վ, "0")

    
    mintFindType = Val(zlDatabase.GetPara("���˲��ҷ�ʽ", glngSys, p����ҽ��վ, "1", , , intType))
    mblnFindTypeEnabled = Not ((intType = 3 Or intType = 15) And InStr(mstrPrivs, "��������") = 0)
    
    '�����:57566
    str���˽������ = CStr(zlDatabase.GetPara("���˽������", glngSys, p����ҽ��վ))
    If str���˽������ <> "" Then
        mlng������� = Val(Left(str���˽������, 1))
        If UBound(Split(str���˽������, "|")) >= 1 Then
            mlng��ǰ����ʱ�� = Val(Split(str���˽������, "|")(1))
        End If
    End If
    
    mblnAutoHandle = Val(zlDatabase.GetPara("����ʱ�Զ�������ɾ���", glngSys, p����ҽ��վ)) = 1
    
    'ҽ������ˢ������
    mstrNotifyAdvice = zlDatabase.GetPara("�Զ�ˢ������", glngSys, p����ҽ��վ, "0000")
    mintNotifyDay = Val(zlDatabase.GetPara("�Զ�ˢ�²�����������", glngSys, p����ҽ��վ, 1))
    mintNotify = Val(zlDatabase.GetPara("�Զ�ˢ�²������ļ��", glngSys, p����ҽ��վ))
    mbln��Ϣ���� = Val(zlDatabase.GetPara("����������ʾ", glngSys, p����ҽ��վ)) = 1
    
    mblnΣ��ֵ = InStr(GetInsidePrivs(p����ҽ��վ), ";Σ��ֵ����;") > 0
    
    mbln��ʾԤԼ���� = Val(zlDatabase.GetPara("��ʾԤԼ����", glngSys, p����ҽ��վ, "1"))
    mbln��Һ�ģʽ = Val(zlDatabase.GetPara(290, glngSys)) = 1
    mblnΣ��ֵ���� = Val(zlDatabase.GetPara("����Σ��ֵ��������", glngSys, p����ҽ��վ, 1)) = 1
    
    '�����Զ�ˢ��
    Call SetTimer
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub LoadPatients����()
'���ܣ����غ��ﲡ���б�
    Dim strSql As String

    Dim str��ʶ As String
    Dim strת��״̬ As String
    Dim i As Long, j As Long
    Dim objRecord As ReportRecord
    Dim rsPati As ADODB.Recordset
    Dim intType As Integer '1���2ԤԼ��3ת��
    Dim lngColor As Long
    Dim colPati As Collection, colValue As Collection
    Dim colList As Collection
    Dim str����ids As String, btnFindPati As Boolean
    
    Dim str�Һ�ids As String
    Dim rs���� As ADODB.Recordset
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    mblnUnRefresh = True
    
    rptPati(PATI_RPT����).Records.DeleteAll
    
    For intType = 1 To 4
        Select Case intType
        Case 1 '���ﲡ��
            If mint���ﷶΧ = 1 Then
                strSql = " And B.ִ����||''=[2]" '�ұ��˺�
                If mblnҪ����� Then strSql = strSql & " And B.���� is Not NULL"
            ElseIf mint���ﷶΧ = 2 Then
                '������
                If mlng�������ID <> 0 Then
                    strSql = " And B.����=[3] And b.ִ�в���id+0 =[4] And (B.ִ����||''=[2] Or B.ִ���� Is Null) "
                Else    '10.28��ǰѡ����ʱû�ж�����
                    strSql = " And B.����=[3] And (B.ִ����||''=[2] Or B.ִ���� Is Null) "
                End If
            ElseIf mint���ﷶΧ = 3 Then
                strSql = " And B.ִ�в���ID+0=[4] And (B.ִ����||''=[2] Or B.ִ���� Is Null)" '������
                If mblnҪ����� Then strSql = strSql & " And B.���� is Not NULL"
            End If
            
            strSql = " Select /*+ Rule*/B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,B.����,m.�Ƿ���ɫͨ��,n.���� ���鼶��,n.���߱�ʶ��ɫ,B.����,B.����," & _
                "       B.����ʱ�� as ʱ��,b.����,D.���� as ���˿���," & _
                "       B.����,B.����,B.����ʱ��,B.����ʱ��,B.ִ�в���ID,B.ִ����," & _
                "       B.ת��״̬,C.���� as ת�����,B.ת������,B.ת��ҽ��,B.ִ��״̬,B.��¼��־,b.ԤԼ" & _
                " From ���˹Һż�¼ B,���ű� C,���ű� D,��������¼ m,���ﲡ�鼶�� n" & _
                " Where B.����ID is not null And (Nvl(B.ִ��״̬,0)=0 or nvl(B.ִ��״̬,0)=[5]) And B.ת�����ID=C.ID(+) And B.��¼����=1 And B.��¼״̬=1 And Nvl(B.��¼��־,0)<>-1 " & _
                " And B.���� = 1 And B.id = m.�Һ�ID(+) And m.���鼶��=n.���(+) " & _
                "       And B.ִ�в���ID=D.id And B.ִ��ʱ�� is Null And B.����ʱ�� <= Trunc(Sysdate)+1-1/24/60/60 " & strSql & _
                " and B.����ʱ�� >= Sysdate-" & gint����Һ����� & _
                " Order By ���鼶�� Nulls last,B.����ʱ�� Nulls last,B.NO"

            Set rsPati = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "δ��", UserInfo.����, mstr��������, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID), IIf(mblnShowLeavePati, -1, 0), UserInfo.ID)
            
            str��ʶ = " "
        Case 2 'ԤԼ����
            If mbln��ʾԤԼ���� Then
                Set colList = ExseSvrGetRgsApptPatiList(mint���ﷶΧ, UserInfo.����, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID), 1, p����ҽ��վ)
            End If
            str��ʶ = "Ԥ"
        Case 3 'ת�ﲡ��
            If mint���ﷶΧ = 1 Then
                strSql = " And B.ת��ҽ��=[2]" 'ת���˺�
            ElseIf mint���ﷶΧ = 2 Then
                'ת�����ң���������ת�ģ�����ҽ�������ѻ���δָ������ҽ��
                strSql = " And B.ת������=[3] And B.ת�����ID=[4] And Nvl(B.ִ����,'��')<>[2] And (B.ת��ҽ��=[2] Or B.ת��ҽ�� Is NULL)"
            ElseIf mint���ﷶΧ = 3 Then
                'ת�����ң���������ת�ģ�����ҽ�������ѻ���δָ������ҽ��
                strSql = " And B.ת�����ID=[4] And Nvl(B.ִ����,'��')<>[2] And (B.ת��ҽ��=[2] Or B.ת��ҽ�� Is NULL)"
            End If
            strSql = _
                " Select /*+ Rule*/B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,B.����,m.�Ƿ���ɫͨ��,n.���� ���鼶��,n.���߱�ʶ��ɫ,B.����,B.ִ����,b.����,D.���� as ���˿���," & _
                " B.����ʱ�� as ʱ��,B.����ʱ��,B.ת�����ID as ִ�в���ID," & _
                " B.ת��״̬,C.���� as ת�����,B.���� as ת������,B.ִ���� as ת��ҽ��,B.ִ��״̬,B.��¼��־" & _
                " From ���˹Һż�¼ B,���ű� C,���ű� D,��������¼ m,���ﲡ�鼶�� n" & _
                " Where B.����ID is not null And B.ת��״̬=0 And B.ִ�в���ID=C.ID And B.��¼����=1 And B.��¼״̬=1 And B.ת�����ID=D.id And Nvl(B.��¼��־,0)<>-1 And B.id = m.�Һ�ID And m.���鼶��= n.��� " & strSql & _
                 " and B.����ʱ�� >= Sysdate-" & gint����Һ����� & _
                 " Order By B.NO"
            Set rsPati = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "δ��", UserInfo.����, mstr��������, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID), 0, 0)
            str��ʶ = "ת"
        Case 4 'ԤԼ����
            If mbln��ʾԤԼ���� Then
                Set colList = Get�쳣ԤԼ�����б�(mint���ﷶΧ, UserInfo.����, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID), 1, p����ҽ��վ)
            End If
            str��ʶ = "Ԥ"
        End Select
        
        If intType = 1 Or intType = 3 Then
            If Not rsPati Is Nothing Then
            
                str����ids = ""
                For i = 1 To rsPati.RecordCount
                     If InStr("," & str����ids & ",", "," & Val(rsPati!����ID & "") & ",") = 0 Then
                        str����ids = str����ids & "," & Val(rsPati!����ID & "")
                     End If
                     rsPati.MoveNext
                Next
                If rsPati.RecordCount > 0 Then rsPati.MoveFirst
                str����ids = Mid(str����ids, 2)
                If str����ids <> "" Then
                    Set colPati = PatiSvrGetVisitPatis(str����ids, "", p����ҽ��վ)
                End If
                
                For i = 1 To rsPati.RecordCount
                    If Not colPati Is Nothing Then
                        Set colValue = GetColObj(colPati, "_" & rsPati!����ID)
                    End If
                    
                    If Not colValue Is Nothing Then
                        If colValue.Count > 0 Then
                            Set objRecord = rptPati(PATI_RPT����).Records.Add()
                            For j = 0 To rptPati(PATI_RPT����).Columns.Count - 1
                                objRecord.AddItem ""
                            Next

                            With objRecord
                                If intType = 1 Then
                                    .Item(COL_HZ_��ʶ).Value = IIf(Val(rsPati!ԤԼ & "") = 1, "Ԥ ", str��ʶ)
                                Else
                                    .Item(COL_HZ_��ʶ).Value = str��ʶ
                                End If
                                .Item(COL_HZ_����).Value = rsPati!���鼶�� & ""
                                .Item(COL_HZ_�����).Value = rsPati!����� & ""
                                .Item(COL_HZ_����).Value = rsPati!���� & ""
                                .Item(COL_HZ_�Ա�).Value = rsPati!�Ա� & ""
                                .Item(COL_HZ_����).Value = rsPati!���� & ""
                                .Item(COL_HZ_��ɫͨ��).Value = IIf(Val(rsPati!�Ƿ���ɫͨ�� & "") <> 0, "��", "")
                                .Item(COL_HZ_���￨��).Value = GetColVal(colValue, "_vcard_no")
                                .Item(COL_HZ_��������).Value = GetColVal(colValue, "_pati_type")
                                .Item(COL_HZ_NO).Value = rsPati!NO & ""
                                .Item(COL_HZ_����ID).Value = rsPati!����ID & ""
                                .Item(COL_HZ_����ʱ��).Value = CStr(Format(rsPati!����ʱ��, "yyyy-MM-dd HH:mm:ss"))
                                .Item(COL_HZ_ִ�в���ID).Value = Val(rsPati!ִ�в���ID & "")
                                .Item(COL_HZ_ִ����).Value = rsPati!ִ���� & ""
                                .Item(COL_HZ_״̬).Value = Nvl(rsPati!ת��״̬)
                                .Item(COL_HZ_IC����).Value = GetColVal(colValue, "_iccard_no")
                                .Item(COL_HZ_��¼��־).Value = rsPati!��¼��־ & ""
                                .Item(COL_HZ_����).Value = rsPati!���� & ""
                                .Item(COL_HZ_���˿���).Value = rsPati!���˿��� & ""
                                
                                If intType = 1 Then '����
                                    .Item(COL_HZ_��������).Value = rsPati!���� & ""
                                    .Item(COL_HZ_����ҽ��).Value = rsPati!ִ���� & ""
                                    .Item(COL_HZ_���).Value = zlStr.Lpad(Nvl(rsPati!����), 5)
                                    .Item(COL_HZ_����ʱ��).Value = CStr(Format(rsPati!����ʱ��, "MM-dd HH:mm"))
                                    .Item(COL_HZ_ִ��״̬).Value = rsPati!ִ��״̬ & ""
                                End If
                                
                                If intType = 1 Or intType = 3 Then '���ת��
                                    .Item(COL_HZ_��).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
                                    .Item(COL_HZ_�Һ�ʱ��).Value = Format(rsPati!ʱ��, "MM-dd HH:mm")
                                End If
                                
                                'ת��״̬
                                strת��״̬ = ""
                                If intType = 1 Then
                                    If Not IsNull(rsPati!ת��״̬) Then
                                        If rsPati!ת��״̬ = 0 Then
                                            '�Ѿ�ת��
                                            strת��״̬ = "���Է�����,����:" & rsPati!ת����� & _
                                                IIf(Not IsNull(rsPati!ת������), ",����:" & Nvl(rsPati!ת������), "") & _
                                                IIf(Not IsNull(rsPati!ת��ҽ��), ",ҽ��:" & Nvl(rsPati!ת��ҽ��), "")
                                        ElseIf rsPati!ת��״̬ = -1 Then
                                            '�Ѿܾ�ת��
                                            strת��״̬ = "�Է��Ѿܾ�,����:" & rsPati!ת����� & _
                                                IIf(Not IsNull(rsPati!ת������), ",����:" & Nvl(rsPati!ת������), "") & _
                                                IIf(Not IsNull(rsPati!ת��ҽ��), ",ҽ��:" & Nvl(rsPati!ת��ҽ��), "")
                                        End If
                                    End If
                                ElseIf intType = 3 Then
                                    'ת�ﲡ��
                                    strת��״̬ = "������ת��,����:" & rsPati!ת����� & _
                                        IIf(Not IsNull(rsPati!ת������), ",����:" & Nvl(rsPati!ת������), "") & _
                                        IIf(Not IsNull(rsPati!ת��ҽ��), ",ҽ��:" & Nvl(rsPati!ת��ҽ��), "")
                                End If
                                .Item(COL_HZ_ת��״̬).Value = strת��״̬
                                
                                If intType = 2 Then 'ԤԼ
                                    .Item(COL_HZ_ԤԼҽ��).Value = rsPati!ִ���� & ""
                                    .Item(COL_HZ_ԤԼʱ��).Value = CStr(Format(rsPati!ʱ�� & "", "yyyy-MM-dd HH:mm"))
                                End If
                                .Item(COL_HZ_���֤��).Value = GetColVal(colValue, "_pati_idcard")
                                                
                                '���ղ����ú�ɫ��ʾ
                                If Not Val(GetColVal(colValue, "_insurance_type")) = 0 And GetColVal(colValue, "_pati_type") = "" Then
                                    .Item(COL_HZ_�����).ForeColor = &HC0&
                                    .Item(COL_HZ_��������).ForeColor = &HC0&
                                Else
                                    '������ɫ
                                    lngColor = GetPatiColor(GetColVal(colValue, "_pati_type"))
                                    .Item(COL_HZ_�����).ForeColor = lngColor
                                    .Item(COL_HZ_��������).ForeColor = lngColor
                                End If
                                
                                '���Ｖ��
                                If rsPati!���߱�ʶ��ɫ <> "" Then
                                    .Item(COL_HZ_��ʶ).BackColor = GetBGR_FromRGB(rsPati!���߱�ʶ��ɫ)
                                End If
                                
                                '�����ﲡ�˻�ɫ
                                If Val(rsPati!ִ��״̬ & "") = -1 Then
                                    For j = 0 To rptPati(PATI_RPT����).Columns.Count - 1
                                        .Item(j).ForeColor = &H808080
                                    Next
                                End If
                                
                            End With
                            rsPati.MoveNext
                        End If
                    End If
                Next
            End If
       Else
            If Not colList Is Nothing Then
                    If colList.Count > 0 Then
                            
                        str����ids = ""
                        str�Һ�ids = ""
                        
                        For i = 1 To colList.RecordCount
                             If InStr("," & str����ids & ",", "," & Val(GetColVal(colList(i), "_pati_id")) & ",") = 0 And Val(GetColVal(colList(i), "_pati_id")) <> 0 Then
                                str����ids = str����ids & "," & Val(GetColVal(colList(i), "_pati_id"))
                             End If
                             
                            If Val(GetColVal(colList(i), "_reg_id")) <> 0 Then
                               str�Һ�ids = str�Һ�ids & "," & Val(GetColVal(colList(i), "_reg_id"))
                            End If
                        Next
                
                        str�Һ�ids = Mid(str�Һ�ids, 2)
                        str����ids = Mid(str����ids, 2)
                        
                        '��ȡ������Ϣ
                        If str����ids <> "" Then
                            Set colPati = PatiSvrGetVisitPatis(str����ids, "", p����ҽ��վ)
                        End If
                        
                        '���������Ϣ
                        If str�Һ�ids <> "" Then
                            strSql = "Select m.�Һ�id, m.�Ƿ���ɫͨ��, n.���� ���鼶��, n.���߱�ʶ��ɫ" & vbNewLine & _
                                    "From ��������¼ M, ���ﲡ�鼶�� N" & vbNewLine & _
                                    "Where m.�Һ�id In (Select Column_Value From Table(f_Str2list([1]))) And m.���鼶�� = n.���(+)"
                            Set rs���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, str�Һ�ids)
                        End If
                        
                        For i = 1 To colList.RecordCount
                            Set objRecord = rptPati(PATI_RPT����).Records.Add()
                            For j = 0 To rptPati(PATI_RPT����).Columns.Count - 1
                                objRecord.AddItem ""
                            Next
                            
                            btnFindPati = False
                            
                            If Not colPati Is Nothing And Val(GetColVal(colList(i), "_pati_id")) <> 0 Then
                                Set colValue = GetColObj(colPati, "_" & Val(GetColVal(colList(i), "_pati_id")))
                                If Not colValue Is Nothing Then
                                    If colValue.Count > 0 Then
                                         btnFindPati = True
                                    End If
                                End If
                            End If
 
                            With objRecord
                                .Item(COL_HZ_��ʶ).Value = str��ʶ
                                .Item(COL_HZ_�����).Value = GetColVal(colList(i), "_outpatient_num")
                                .Item(COL_HZ_����).Value = GetColVal(colList(i), "_pati_name")
                                .Item(COL_HZ_�Ա�).Value = GetColVal(colList(i), "_pati_sex")
                                .Item(COL_HZ_����).Value = GetColVal(colList(i), "_pati_age")

                                .Item(COL_HZ_NO).Value = GetColVal(colList(i), "_reg_no")
                                .Item(COL_HZ_����ID).Value = GetColVal(colList(i), "_pati_id")
                                .Item(COL_HZ_����ʱ��).Value = CStr(Format(GetColVal(colList(i), "_happen_time"), "yyyy-MM-dd HH:mm:ss"))
                                .Item(COL_HZ_ִ�в���ID).Value = Val(GetColVal(colList(i), "_exe_deptid"))
                                .Item(COL_HZ_ִ����).Value = GetColVal(colList(i), "_exetr")
                                .Item(COL_HZ_״̬).Value = GetColVal(colList(i), "_outp_rfrl_status")
                                .Item(COL_HZ_��¼��־).Value = GetColVal(colList(i), "_record_sign")
                                .Item(COL_HZ_����).Value = GetColVal(colList(i), "_outptyp_name")
                                .Item(COL_HZ_���˿���).Value = GetColVal(colList(i), "_pait_dept")
                                .Item(COL_HZ_ת��״̬).Value = ""
                                
                                .Item(COL_HZ_ת��״̬).Value = ""
                                .Item(COL_HZ_ԤԼҽ��).Value = GetColVal(colList(i), "_exetr")
                                .Item(COL_HZ_ԤԼʱ��).Value = CStr(Format(GetColVal(colList(i), "_happen_time"), "MM-dd HH:mm:ss"))

                                If btnFindPati Then
                                    .Item(COL_HZ_���￨��).Value = GetColVal(colValue, "_vcard_no")
                                    .Item(COL_HZ_��������).Value = GetColVal(colValue, "_pati_type")
                                    .Item(COL_HZ_IC����).Value = GetColVal(colValue, "_iccard_no")
                                    .Item(COL_HZ_���֤��).Value = GetColVal(colValue, "_pati_idcard")
                                    
                                                
                                    '���ղ����ú�ɫ��ʾ
                                    If Not Val(GetColVal(colValue, "_insurance_type")) = 0 And GetColVal(colValue, "_pati_type") = "" Then
                                        .Item(COL_HZ_�����).ForeColor = &HC0&
                                        .Item(COL_HZ_��������).ForeColor = &HC0&
                                    Else
                                        '������ɫ
                                        lngColor = GetPatiColor(GetColVal(colValue, "_pati_type"))
                                        .Item(COL_HZ_�����).ForeColor = lngColor
                                        .Item(COL_HZ_��������).ForeColor = lngColor
                                    End If
                                Else
                                    .Item(COL_HZ_���￨��).Value = ""
                                    .Item(COL_HZ_��������).Value = ""
                                    .Item(COL_HZ_IC����).Value = ""
                                    .Item(COL_HZ_���֤��).Value = ""
                                    .Item(COL_HZ_�����).ForeColor = &HC0&
                                    .Item(COL_HZ_��������).ForeColor = &HC0&
                                End If
                                
                                If Val(GetColVal(colList(i), "_reg_id")) <> 0 Then
                                    If Not rs���� Is Nothing Then
                                        rs����.Filter = "�Һ�id =" & Val(GetColVal(colList(i), "_reg_id"))
                                        
                                        If Not rs����.EOF Then
                                            .Item(COL_HZ_����).Value = rs����!���鼶�� & ""
                                        
                                            If rs����!���߱�ʶ��ɫ <> "" Then
                                                .Item(COL_HZ_��ʶ).BackColor = GetBGR_FromRGB(rs����!���߱�ʶ��ɫ & "")
                                            End If
                                        End If
                                    End If
                                End If
                                
                                '�����ﲡ�˻�ɫ
                                If Val(GetColVal(colList(i), "_exe_status")) = -1 Then
                                    For j = 0 To rptPati(PATI_RPT����).Columns.Count - 1
                                        .Item(j).ForeColor = &H808080
                                    Next
                                End If
                                
                                
                                
                                '�쳣�������ɫ
                                If intType = 4 Then
                                    For j = 0 To rptPati(PATI_RPT����).Columns.Count - 1
                                        .Item(j).ForeColor = &HC0E0FF
                                    Next
                                    .Tag = "��"
                                End If
                                
                                
                            End With
                        Next
                  
                    End If
                End If
       End If
  
        If intType = 1 Then
            Call SetRoomState(rptPati(PATI_RPT����).Records.Count > 0)
        End If
    Next
    
    rptPati(PATI_RPT����).Populate
    i = rptPati(PATI_RPT����).Records.Count
    tbcWait.Item(0).Caption = "���ﲡ��" & IIf(i = 0, "", ":" & i & "��")
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





Private Sub LoadPatientsԤԼ()
'���ܣ�����ԤԼ�����б�
    Dim i As Long, j As Long, lngErr As Long
    Dim objRecord As ReportRecord
    Dim rs���� As ADODB.Recordset
    Dim lngColor As Long
    Dim strSql As String
    Dim str�Һ�ids As String
    
    
    Dim colList As Collection
    Dim colPati As Collection, colPatiValue As Collection
    Dim str����ids As String, btnFindPati As Boolean
    
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    mblnUnRefresh = True
    
    rptPati(PATI_RPTԤԼ).Records.DeleteAll
    
    For lngErr = 0 To 1
        Set rs���� = Nothing

        If lngErr = 0 Then
            Set colList = ExseSvrGetRgsApptPatiList(mint���ﷶΧ, UserInfo.����, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID), 1, p����ҽ��վ)
        Else
            Set colList = Get�쳣ԤԼ�����б�(mint���ﷶΧ, UserInfo.����, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID), 1, p����ҽ��վ)
        End If
        
        
        If Not colList Is Nothing Then
            If colList.Count > 0 Then
                str����ids = ""
                str�Һ�ids = ""
                For i = 1 To colList.Count
                     If InStr("," & str����ids & ",", "," & Val(GetColVal(colList(i), "_pati_id")) & ",") = 0 And Val(GetColVal(colList(i), "_pati_id")) <> 0 Then
                        str����ids = str����ids & "," & Val(GetColVal(colList(i), "_pati_id"))
                     End If
                     If Val(GetColVal(colList(i), "_reg_id")) <> 0 Then
                        str�Һ�ids = str�Һ�ids & "," & Val(GetColVal(colList(i), "_reg_id"))
                     End If
                Next
                
                str�Һ�ids = Mid(str�Һ�ids, 2)
                str����ids = Mid(str����ids, 2)
                
                '��ȡ������Ϣ
                If str����ids <> "" Then
                    Set colPati = PatiSvrGetVisitPatis(str����ids)
                End If
                
                '���������Ϣ
                If str�Һ�ids <> "" Then
                    strSql = "Select m.�Һ�id, m.�Ƿ���ɫͨ��, n.���� ���鼶��, n.���߱�ʶ��ɫ" & vbNewLine & _
                            "From ��������¼ M, ���ﲡ�鼶�� N" & vbNewLine & _
                            "Where m.�Һ�id In (Select Column_Value From Table(f_Str2list([1]))) And m.���鼶�� = n.���(+)"
                    Set rs���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, str�Һ�ids)
                End If
                
                For i = 1 To colList.Count
                    Set objRecord = rptPati(PATI_RPTԤԼ).Records.Add()
                    For j = 0 To rptPati(PATI_RPTԤԼ).Columns.Count - 1
                        objRecord.AddItem ""
                    Next
                    
                    btnFindPati = False
                    
                    If Not colPati Is Nothing And Val(GetColVal(colList(i), "_pati_id")) <> 0 Then
                        Set colPatiValue = GetColObj(colPati, "_" & Val(GetColVal(colList(i), "_pati_id")))
                        If Not colPatiValue Is Nothing Then
                            If colPatiValue.Count > 0 Then
                                 btnFindPati = True
                            End If
                        End If
                    End If
                    
    
                    
                    With objRecord
                        .Item(COL_HZ_��ʶ).Value = "Ԥ"
                        .Item(COL_HZ_�����).Value = GetColVal(colList(i), "_outpatient_num")
                        .Item(COL_HZ_����).Value = GetColVal(colList(i), "_pati_name")
                        .Item(COL_HZ_�Ա�).Value = GetColVal(colList(i), "_pati_sex")
                        .Item(COL_HZ_����).Value = GetColVal(colList(i), "_pati_age")
    
                        .Item(COL_HZ_NO).Value = GetColVal(colList(i), "_reg_no")
                        .Item(COL_HZ_����ID).Value = GetColVal(colList(i), "_pati_id")
                        .Item(COL_HZ_����ʱ��).Value = CStr(Format(GetColVal(colList(i), "_happen_time"), "yyyy-MM-dd HH:mm:ss"))
                        .Item(COL_HZ_ִ�в���ID).Value = Val(GetColVal(colList(i), "_exe_deptid"))
                        .Item(COL_HZ_ִ����).Value = GetColVal(colList(i), "_exetr")
                        .Item(COL_HZ_״̬).Value = GetColVal(colList(i), "_outp_rfrl_status")
    
                        .Item(COL_HZ_��¼��־).Value = GetColVal(colList(i), "_record_sign")
                        .Item(COL_HZ_����).Value = GetColVal(colList(i), "_outptyp_name")
                        .Item(COL_HZ_���˿���).Value = GetColVal(colList(i), "_pait_dept")
    
                        .Item(COL_HZ_ת��״̬).Value = ""
                        .Item(COL_HZ_ԤԼҽ��).Value = GetColVal(colList(i), "_exetr")
                        .Item(COL_HZ_ԤԼʱ��).Value = CStr(Format(GetColVal(colList(i), "_happen_time"), "MM-dd HH:mm:ss"))
    
                        
                        If btnFindPati Then
                            .Item(COL_HZ_���￨��).Value = GetColVal(colPatiValue, "_vcard_no")
                            .Item(COL_HZ_��������).Value = GetColVal(colPatiValue, "_pati_type")
                            .Item(COL_HZ_IC����).Value = GetColVal(colPatiValue, "_iccard_no")
                            .Item(COL_HZ_���֤��).Value = GetColVal(colPatiValue, "_pati_idcard")
                            
                                        
                            '���ղ����ú�ɫ��ʾ
                            If Not Val(GetColVal(colPatiValue, "_insurance_type")) = 0 And GetColVal(colPatiValue, "_pati_type") = "" Then
                                .Item(COL_HZ_�����).ForeColor = &HC0&
                                .Item(COL_HZ_��������).ForeColor = &HC0&
                            Else
                                '������ɫ
                                lngColor = GetPatiColor(GetColVal(colPatiValue, "_pati_type"))
                                .Item(COL_HZ_�����).ForeColor = lngColor
                                .Item(COL_HZ_��������).ForeColor = lngColor
                            End If
                        Else
                            .Item(COL_HZ_���￨��).Value = ""
                            .Item(COL_HZ_��������).Value = ""
                            .Item(COL_HZ_IC����).Value = ""
                            .Item(COL_HZ_���֤��).Value = ""
                            .Item(COL_HZ_�����).ForeColor = &HC0&
                            .Item(COL_HZ_��������).ForeColor = &HC0&
                        End If
                        
                        
                        If Val(GetColVal(colList(i), "_reg_id")) <> 0 Then
                            If Not rs���� Is Nothing Then
                                rs����.Filter = "�Һ�id =" & Val(GetColVal(colList(i), "_reg_id"))
                                
                                If Not rs����.EOF Then
                                    .Item(COL_HZ_����).Value = rs����!���鼶�� & ""
                                
                                    If rs����!���߱�ʶ��ɫ <> "" Then
                                        .Item(COL_HZ_��ʶ).BackColor = GetBGR_FromRGB(rs����!���߱�ʶ��ɫ)
                                    End If
                                End If
                            End If
                        End If
                        
    
                        '�����ﲡ�˻�ɫ
                        If Val(GetColVal(colList(i), "_exe_status")) = -1 Then
                            For j = 0 To rptPati(PATI_RPTԤԼ).Columns.Count - 1
                                .Item(j).ForeColor = &H808080
                            Next
                        End If
                        
                        '�쳣�������ɫ
                        If lngErr = 1 Then
                            For j = 0 To rptPati(PATI_RPT����).Columns.Count - 1
                                .Item(j).ForeColor = &HC0E0FF
                            Next
                            .Tag = "��"
                        End If
                        
                        
                    End With
    
                Next
          
            End If
        End If
    Next
    rptPati(PATI_RPTԤԼ).Populate
    i = rptPati(PATI_RPTԤԼ).Records.Count
    tbcWait.Item(mintԤԼ�б�).Caption = "ԤԼ����" & IIf(i = 0, "", ":" & i & "��")
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


Private Sub LoadPatients����()
'���ܣ����غ�������б�
    Dim strSql As String
    Dim i As Long, j As Long
    Dim objRecord As ReportRecord
    Dim rsPati As ADODB.Recordset
    Dim lngColor As Long
    Dim rs��Ⱦ�������¼ As ADODB.Recordset
    Dim blnDo��Ⱦ��״̬ As Boolean
    Dim colPati As Collection, colValue As Collection
    Dim str����ids As String
    
 
    On Error GoTo errH
    Screen.MousePointer = 11
    mblnUnRefresh = True
    
    strSql = _
        " Select B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,B.����,m.�Ƿ���ɫͨ��,n.���� ���鼶��,n.���߱�ʶ��ɫ,B.����,b.����,D.���� as ���˿���," & _
        " B.ִ��ʱ�� as ʱ��,B.����ʱ��,B.ִ�в���ID,B.ִ����," & _
        " B.ת��״̬,C.���� as ת�����,B.ת������,B.ת��ҽ��,B.ִ��״̬,B.��¼��־" & _
        " From ���˹Һż�¼ B,���ű� C,���ű� D,��������¼ m,���ﲡ�鼶�� n" & _
        " Where B.����ID IS NOT NULL And B.ת�����ID=C.ID(+) and B.ִ�в���ID=d.id " & _
        " And B.ִ��״̬=2 And B.ִ����||''=[1] And B.��¼����=1 And B.��¼״̬=1 and nvl(B.��¼��־,0)<=1  And B.���� = 1 And B.id = m.�Һ�ID(+) And m.���鼶��=n.���(+)" & _
        " Order By B.NO"
    Set rsPati = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstr����ҽ��)
  
    str����ids = ""
    If Not rsPati Is Nothing Then
        For i = 1 To rsPati.RecordCount
             If InStr("," & str����ids & ",", "," & Val(rsPati!����ID & "") & ",") = 0 Then
                str����ids = str����ids & "," & Val(rsPati!����ID & "")
             End If
             rsPati.MoveNext
        Next
        If rsPati.RecordCount > 0 Then rsPati.MoveFirst
    End If
    str����ids = Mid(str����ids, 2)
    If str����ids <> "" Then
        Set colPati = PatiSvrGetVisitPatis(str����ids, "", p����ҽ��վ)
    End If
  
  
    Set rs��Ⱦ�������¼ = Get��Ⱦ�������¼(mstr����ҽ��, PatiType.pt����)
    If rs��Ⱦ�������¼.RecordCount > 0 Then blnDo��Ⱦ��״̬ = True
 
    rptPati(PATI_RPT����).Records.DeleteAll
    For i = 1 To rsPati.RecordCount
        If Not colPati Is Nothing Then
            Set colValue = GetColObj(colPati, "_" & rsPati!����ID)
        End If
        
        If Not colValue Is Nothing Then
            If colValue.Count > 0 Then
                Set objRecord = rptPati(PATI_RPT����).Records.Add()
                For j = 0 To rptPati(PATI_RPT����).Columns.Count - 1
                    objRecord.AddItem ""
                Next
                With objRecord
                    .Item(COL_JZ_��ʶ).Value = ""
                    .Item(COL_JZ_����).Value = rsPati!���鼶�� & ""
                    .Item(COL_JZ_�����).Value = rsPati!����� & ""
                    .Item(COL_JZ_����).Value = rsPati!���� & ""
                    .Item(COL_JZ_����ʱ��).Value = Format(rsPati!ʱ��, "MM-dd HH:mm")
                    .Item(COL_JZ_�Ա�).Value = rsPati!�Ա� & ""
                    .Item(COL_JZ_����).Value = rsPati!���� & ""
                    .Item(COL_JZ_��ɫͨ��).Value = IIf(Val(rsPati!�Ƿ���ɫͨ�� & "") <> 0, "��", "")
                    .Item(COL_JZ_��).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
                    .Item(COL_JZ_NO).Value = rsPati!NO & ""
                    
                    .Item(COL_JZ_���￨��).Value = GetColVal(colValue, "_vcard_no")
                    .Item(COL_JZ_��������).Value = GetColVal(colValue, "_pati_type")
                    .Item(COL_JZ_����ID).Value = rsPati!����ID & ""
                    .Item(COL_JZ_����ʱ��).Value = CStr(Format(rsPati!����ʱ��, "yyyy-MM-dd HH:mm:ss"))
                    .Item(COL_JZ_ִ�в���ID).Value = rsPati!ִ�в���ID & ""
                    .Item(COL_JZ_ִ����).Value = rsPati!ִ���� & ""
                    .Item(COL_JZ_���֤��).Value = GetColVal(colValue, "_pati_idcard")
                    .Item(COL_JZ_IC����).Value = GetColVal(colValue, "_iccard_no")
                    .Item(COL_JZ_��¼��־).Value = rsPati!��¼��־ & ""
                    .Item(COL_JZ_����).Value = rsPati!���� & ""
                    .Item(COL_JZ_���˿���).Value = rsPati!���˿��� & ""
                    
                    'ת��״̬:��ʾ�����һ��
                    .Item(COL_JZ_״̬).Value = Nvl(rsPati!ת��״̬)
                    If Not IsNull(rsPati!ת��״̬) Then
                        If rsPati!ת��״̬ = 0 Then
                            .Item(COL_JZ_ת��״̬).Value = "���Է�����,����:" & rsPati!ת����� & _
                                IIf(Not IsNull(rsPati!ת������), ",����:" & Nvl(rsPati!ת������), "") & _
                                IIf(Not IsNull(rsPati!ת��ҽ��), ",ҽ��:" & Nvl(rsPati!ת��ҽ��), "")
                        ElseIf rsPati!ת��״̬ = -1 Then
                            '�Ѿܾ�ת��
                            .Item(COL_JZ_ת��״̬).Value = "�Է��Ѿܾ�,����:" & rsPati!ת����� & _
                                IIf(Not IsNull(rsPati!ת������), ",����:" & Nvl(rsPati!ת������), "") & _
                                IIf(Not IsNull(rsPati!ת��ҽ��), ",ҽ��:" & Nvl(rsPati!ת��ҽ��), "")
                        End If
                    End If
                    
                    '���ղ����ú�ɫ��ʾ
                   If Not Val(GetColVal(colValue, "_insurance_type")) = 0 And GetColVal(colValue, "_pati_type") = "" Then
                        .Item(COL_JZ_�����).ForeColor = &HC0&
                        .Item(COL_JZ_��������).ForeColor = &HC0&
                    Else
                        '������ɫ
                        lngColor = GetPatiColor(GetColVal(colValue, "_pati_type"))
                        .Item(COL_JZ_�����).ForeColor = lngColor
                        .Item(COL_JZ_��������).ForeColor = lngColor
                    End If
                            
                    '����ּ���ɫ
                    If rsPati!���߱�ʶ��ɫ <> "" Then
                        .Item(COL_JZ_��ʶ).BackColor = GetBGR_FromRGB(rsPati!���߱�ʶ��ɫ)
                    End If
                    
                    '��Ӵ�Ⱦ��״̬
                    strSql = ""
                    If blnDo��Ⱦ��״̬ Then
                        rs��Ⱦ�������¼.Filter = "no='" & rsPati!NO & "'"
                        If Not rs��Ⱦ�������¼.EOF Then strSql = Get��Ⱦ��״̬(Val(rs��Ⱦ�������¼!��¼ & ""), Val(rs��Ⱦ�������¼!��д & ""), Val(rs��Ⱦ�������¼!״̬ & ""))
                    End If
                    .Item(COL_JZ_��Ⱦ��).Value = strSql
                End With
            End If
        End If
        rsPati.MoveNext
    Next
    rptPati(PATI_RPT����).Populate
    i = rptPati(PATI_RPT����).Records.Count
    tbcInTreat.Item(t����).Caption = "����" & IIf(i = 0, "", ":" & i & "��")
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

Private Sub LoadPatients����()
'���ܣ��������ﲡ���б�
    Dim strSql As String
    Dim i As Long, j As Long
    Dim objRecord As ReportRecord
    Dim rsPati As ADODB.Recordset
    Dim lngColor As Long
    Dim bln��ҽ As Boolean
    Dim colPati As Collection, colValue As Collection
    Dim str����ids As String
    
    On Error GoTo errH
    Screen.MousePointer = 11
    mblnUnRefresh = True
    rptPati(PATI_RPT����).Records.DeleteAll
    
    
    strSql = "Select /*+ Rule*/" & vbNewLine & _
        " Distinct(b.No), b.����id, b.�����, b.����, b.�Ա�, b.����, b.����,m.�Ƿ���ɫͨ��,n.���� ���鼶��,n.���߱�ʶ��ɫ, b.����, b.ִ��ʱ�� As ʱ��, b.����ʱ��, b.ִ�в���id," & vbNewLine & _
        " b.ִ����, b.ִ��״̬, b.��¼��־,b.����,F.���� as ���˿���," & vbNewLine & _
        " First_Value(Decode(Sign(h.������� - 10), -1, h.�������, '')) Over(Partition By h.����id, h.��ҳid Order By Sign(h.������� - 10), Decode(h.��¼��Դ, 4, 0, h.��¼��Դ) Desc, Decode(h.�������, 1, 1, 0) Desc, h.��ϴ���) As ��ҽ���," & vbNewLine & _
        " First_Value(Decode(Sign(h.������� - 10), 1, h.�������, '')) Over(Partition By h.����id, h.��ҳid Order By -Sign(h.������� - 10), Decode(h.��¼��Դ, 4, 0, h.��¼��Դ) Desc, Decode(h.�������,11,11, 0) Desc, h.��ϴ���) As ��ҽ���" & vbNewLine & _
        " From ���˹Һż�¼ B,���ű� F, ������ϼ�¼ H,��������¼ m,���ﲡ�鼶�� n " & vbNewLine & _
        " Where b.����id is not null And h.����id(+) = b.����id And h.��ҳid(+) = b.id and b.ִ�в���id=f.id And b.ִ��״̬ + 0 = 1 And b.��¼���� = 1 And b.��¼״̬ = 1 " & _
        " and B.���� = 1 And B.id = m.�Һ�ID(+) And m.���鼶��=n.���(+) "
        

    If mvCondFilter.����ID <> 0 Then
        strSql = strSql & " And B.����id=[5]"
    ElseIf (mvCondFilter.���� = "�Һŵ�" Or mvCondFilter.���� = "���ݺ�") And mvCondFilter.�ı� <> "" Then
        strSql = strSql & " And B.NO=[3]"
    ElseIf mvCondFilter.���� = "�����" And mvCondFilter.�ı� <> "" Then
        strSql = strSql & " And B.�����=[3]"
    ElseIf mvCondFilter.���� = "���￨" And mvCondFilter.�ı� <> "" Then
        strSql = strSql & " And B.����id in (Select Column_Value From Table(f_Str2list([4])))"

        str����ids = ""
        Set colPati = PatiSvrGetVisitPatis("", mvCondFilter.�ı�, p����ҽ��վ)
        If Not colPati Is Nothing Then
            If colPati.Count > 0 Then
                For i = 1 To colPati.Count
                    str����ids = str����ids & "," & GetColVal(colPati(i), "_pati_id")
                Next
            End If
        End If
        str����ids = Mid(str����ids, 2)

    Else
        strSql = strSql & " And B.ִ��ʱ�� Between To_Date('" & Format(mvCondFilter.Begin, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(mvCondFilter.End, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & IIf(mvCondFilter.ҽ�� = "", "", " And B.ִ����||''=[1]")
        If mvCondFilter.����ID <> 0 Then strSql = strSql & " And B.ִ�в���ID+0=[2]"
        
        If mvCondFilter.���� = "����" And mvCondFilter.�ı� <> "" Then
            strSql = strSql & " And B.����=[3]"
        End If
    End If
    
    If zlDatabase.DateMoved(mvCondFilter.Begin) Then
        strSql = strSql & " Union ALL " & Replace(strSql, "���˹Һż�¼", "H���˹Һż�¼")
    End If

    strSql = strSql & " Order By NO Desc"
    
    With mvCondFilter
        Set rsPati = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .ҽ��, .����ID, .�ı�, str����ids, .����ID)
    End With
    
    str����ids = ""
    If (Not rsPati Is Nothing) And (Not (mvCondFilter.���� = "���￨" And mvCondFilter.�ı� <> "")) Then
        For i = 1 To rsPati.RecordCount
             If InStr("," & str����ids & ",", "," & Val(rsPati!����ID & "") & ",") = 0 Then
                str����ids = str����ids & "," & Val(rsPati!����ID & "")
             End If
             rsPati.MoveNext
        Next
        If rsPati.RecordCount > 0 Then rsPati.MoveFirst
    End If
    str����ids = Mid(str����ids, 2)
    If str����ids <> "" Then
        Set colPati = PatiSvrGetVisitPatis(str����ids, "", p����ҽ��վ)
    End If
     
    For i = 1 To rsPati.RecordCount
        If Not colPati Is Nothing Then
            Set colValue = GetColObj(colPati, "_" & rsPati!����ID)
        End If
        
        If Not colValue Is Nothing Then
            If colValue.Count > 0 Then
                Set objRecord = rptPati(PATI_RPT����).Records.Add()
                For j = 0 To rptPati(PATI_RPT����).Columns.Count - 1
                    objRecord.AddItem ""
                Next
            
                With objRecord
                    .Item(COL_YZ_����).Value = rsPati!���鼶�� & ""
                    .Item(COL_YZ_�����).Value = rsPati!����� & ""
                    .Item(COL_YZ_����).Value = rsPati!���� & ""
                    .Item(COL_YZ_�Ա�).Value = rsPati!�Ա� & ""
                    .Item(COL_YZ_����).Value = rsPati!���� & ""
                    .Item(COL_YZ_��ɫͨ��).Value = IIf(Val(rsPati!�Ƿ���ɫͨ�� & "") <> 0, "��", "")
                    .Item(COL_YZ_��).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
                    .Item(COL_YZ_����ʱ��).Value = CStr(Format(rsPati!ʱ�� & "", "yyyy-MM-dd HH:mm"))
                    .Item(COL_YZ_����ҽ��).Value = rsPati!ִ���� & ""
                    .Item(COL_YZ_���￨��).Value = GetColVal(colValue, "_vcard_no")
                    .Item(COL_YZ_��������).Value = GetColVal(colValue, "_pati_type")
                    .Item(COL_YZ_����).Value = rsPati!���� & ""
                    .Item(COL_YZ_���˿���).Value = rsPati!���˿��� & ""
                    .Item(COL_YZ_NO).Value = rsPati!NO & ""
                    .Item(COL_YZ_����ID).Value = rsPati!����ID & ""
                    .Item(COL_YZ_����ʱ��).Value = CStr(Format(rsPati!����ʱ��, "yyyy-MM-dd HH:mm:ss"))
                    .Item(COL_YZ_ִ�в���ID).Value = Val(rsPati!ִ�в���ID & "")
                    .Item(COL_YZ_ִ����).Value = rsPati!ִ���� & ""
                    .Item(COL_YZ_���֤��).Value = GetColVal(colValue, "_pati_idcard")
                    .Item(COL_YZ_IC����).Value = GetColVal(colValue, "_iccard_no")
                    .Item(COL_YZ_��¼��־).Value = rsPati!��¼��־ & ""
                    .Item(COL_YZ_��ҽ���).Value = Replace(rsPati!��ҽ��� & "", "&", "��")
                    .Item(COL_YZ_��ҽ���).Value = Replace(rsPati!��ҽ��� & "", "&", "��")
                    If rsPati!��ҽ��� & "" <> "" Then bln��ҽ = True
                    
                    '���ղ����ú�ɫ��ʾ
                    If Not Val(GetColVal(colValue, "_insurance_type")) = 0 And GetColVal(colValue, "_pati_type") = "" Then
                        .Item(COL_YZ_�����).ForeColor = &HC0&
                        .Item(COL_YZ_��������).ForeColor = &HC0&
                    Else
                        '������ɫ
                        lngColor = GetPatiColor(GetColVal(colValue, "_pati_type"))
                        .Item(COL_YZ_�����).ForeColor = lngColor
                        .Item(COL_YZ_��������).ForeColor = lngColor
                    End If
                    
                    If rsPati!���߱�ʶ��ɫ <> "" Then
                        .Item(COL_YZ_��ʶ).BackColor = GetBGR_FromRGB(rsPati!���߱�ʶ��ɫ)
                    End If
                End With
                
                
            End If
        End If
        rsPati.MoveNext
    Next
    
    rptPati(PATI_RPT����).Columns(COL_YZ_��ҽ���).Visible = bln��ҽ
    rptPati(PATI_RPT����).Populate
    i = rptPati(PATI_RPT����).Records.Count
    tbcInTreat.Item(t����).Caption = "����" & IIf(i = 0, "", ":" & i & "��")
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

Private Function GetBGR_FromRGB(strRGB As String)
'���ܣ���HTML��ʽ��RGB��ɫת��ΪVB��ʽ��BGR��ʽ
    
    GetBGR_FromRGB = "&H" & Mid(strRGB, 5, 2) & Mid(strRGB, 3, 2) & Mid(strRGB, 1, 2)
End Function

Private Function LoadPatients(Optional ByVal strRefesh As String = "11111", Optional ByVal intActive As PATI_RPT_LIST = -1, Optional ByVal strActNO As String) As Boolean
'���ܣ���ȡ�����б�
'������strActNO=ˢ�º���Ҫ��λ���б������Ͳ��˹Һŵ�(�����)
'      ע���������ָ����intActive,�����Ҫ����strRefeshˢ���б���
'      strRefesh=�ֱ��Ƿ�ˢ��ָ�����б��ֱ�Ϊ ��1λ��"����/ת��/ԤԼ"����2λ��"����"����3λ��"����"����4λ-"����"����5λ-"ԤԼ"
    Dim strPrePati As String
    Dim i As Long
    Dim intIdx As Long
    Dim lngCol As Long
    Dim objRpt As ReportControl
    
    strPrePati = mstrPrePati '��ΪҪ�ƻ�,�����ʱ��¼
    
    If strActNO <> "" Then strPrePati = strActNO
    
    Screen.MousePointer = 11
    On Error GoTo errH
    mblnUnRefresh = True
    
    For i = 1 To 5
        If Mid(strRefesh, i, 1) = "1" Then
            If i = 1 Then
                Call LoadPatients����
                If mbln��ʾԤԼ���� Then
                    rptPati(PATI_RPTԤԼ).Records.DeleteAll
                    rptPati(PATI_RPTԤԼ).Populate
                End If
            ElseIf i = 2 Then
                Call LoadPatients����
            ElseIf i = 3 Then
                Call LoadPatients����
            ElseIf i = 4 Then
                Call LoadPatients����
            ElseIf i = 5 Then
                If Not mbln��ʾԤԼ���� Then
                    Call LoadPatientsԤԼ
                End If
            End If
        End If
    Next
    i = 0
    For intIdx = 0 To 4
        Set objRpt = rptPati(intIdx)
        If objRpt.Visible Then
            lngCol = IIf(intIdx = PATI_RPT���� Or intIdx = PATI_RPTԤԼ, COL_HZ_NO, IIf(intIdx = PATI_RPT����, COL_YZ_NO, COL_JZ_NO))
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
'        For lngCol = 0 To 3
'            If rptPati(lngCol).SelectedRows.Count > 0 Then
'                rptPati(lngCol).SelectedRows(0).Selected = False
'            End If
'        Next
        mstrPrePati = ""
        '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        Set objRpt.FocusedRow = objRpt.Rows(i)
        If objRpt.Visible Then objRpt.SetFocus
        Call RptItemClick(intIdx)
    Else
        '����ǰ�б�������ˢ���Ӵ���
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
'���ܣ��������������ص���ʾ��Ϣ
    Dim i As Long
    
 
    mlng����ID = 0
    mstr�Һŵ� = ""
    mlng����ID = 0
    mlng�Һ�ID = 0
    mstrPrePati = ""
    mPatiInfo.���� = 0
    mPatiInfo.���� = ""
    mPatiInfo.����� = ""
    mPatiInfo.�Һŵ� = ""
    mPatiInfo.����ID = 0
    mPatiInfo.�Һ�ID = 0
    mPatiInfo.����ID = 0
    mPatiInfo.���� = ""
    mPatiInfo.���� = 0
    mPatiInfo.������ = ""
    mPatiInfo.�Һ�ʱ�� = CDate(0)
    mPatiInfo.����ת�� = False
    mPatiInfo.�Ƿ�ǩ�� = False
    mPatiInfo.������ = ""
    mPatiInfo.����״�� = ""
    mPatiInfo.�Ա� = ""
    mPatiInfo.���� = ""
    mPatiInfo.���� = ""
    mPatiInfo.���� = ""
    mPatiInfo.�����ص� = ""
    mPatiInfo.��Ⱦ���ϴ� = 0
    mPatiInfo.��ͥ��ַ�ʱ� = ""
    mPatiInfo.��λ�ʱ� = ""
    mPatiInfo.����֤�� = ""
    mPatiInfo.�Ƿ���ɫͨ�� = 0
    mPatiInfo.���鼶�� = ""
        
    
    imgPatient.Picture = imgDefual.Picture

    txtInfo(txtInfo����).Text = ""
    txtInfo(txtInfo�Ա�).Text = ""
    txtInfo(txtInfo����).Text = ""
    txtInfo(txtInfo��������).Text = ""
    txtInfo(txtInfo���￨��).Text = ""
    txtInfo(txtInfoҽ������).Text = ""
    txtInfo(txtInfo������Ϣ).Text = ""
    txtInfo(txtInfoժҪ).Text = ""
    txtInfo(txtInfoժҪ).ToolTipText = ""
    txtPhone.Text = ""
    txtPhone.ToolTipText = ""
    
    lblMore.Visible = False
    lblRec.Visible = False
    
    cboPayType.ListIndex = -1
    cboBillType.ListIndex = -1
    
    For i = 0 To lblLink�޸�
        lblLink(i).ForeColor = &HC0C0C0
    Next
    mPr = -1
End Sub

Private Sub ExecuteRegist(ByVal strNO As String)
'���ܣ����˹Һ�
    mblnUnRefresh = True
    'ˢ�²���λ���չҺŵĲ�����
    If strNO <> "" And rptPati(PATI_RPT����).Visible Then
        Call LoadPatients("11001", PATI_RPT����, strNO)
    Else
        Call LoadPatients("10001")
    End If
    mblnUnRefresh = False
End Sub

Private Sub ExecuteBespeakPrint()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ش�ԤԼ�Һŵ�
    '����:���˺�
    '����:2012-12-24 10:55:39
    '˵��:
    '����:56274
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCommon As String, intAtom As Integer, strNO As String
    On Error Resume Next
    If gobjRegist Is Nothing Then
        Set gobjRegist = CreateObject("zl9RegEvent.clsRegEvent")
        If Err <> 0 Then
            Err = 0: On Error GoTo 0
        End If
        If gobjRegist Is Nothing Then Exit Sub
    End If
    On Error GoTo errHandle
 
    strNO = mstr�Һŵ�
 
    If strNO = "" Then Exit Sub
    '��������(����Ϸ�������)
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
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
'���ܣ�����ת��
    Dim rsTmp As New ADODB.Recordset
    Dim lng����ID As Long, str���� As String
    Dim strҽ�� As String, lngҽ��ID As Long
    Dim strSql As String
    
    If mstr�Һŵ� = "" Then
        MsgBox "����ѡ���ˡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mintActive = pt���� Then
        If zlDatabase.NOMoved("������ü�¼", mstr�Һŵ�, "��¼����=", "4") Then
            MsgBox "�ò��˵ĹҺŷ����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '���Һŵ�ʱ��
    If BillExpend(mstr�Һŵ�) Then
        MsgBox "�ò��˹Һ��ѳ�����Ч�����������ٽ���ת�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errH
    
    '�����ھ���Ĳ��˵ļ��
    If mintActive = pt���� Or mintActive = pt���� Then
        If InStr(GetInsidePrivs(p����ҽ��վ), "����ҽ��ת��") > 0 Then
            '����Ƿ���δ���͵�ҽ��
            strSql = "Select ID From ����ҽ����¼ Where ����ID+0=[1] And �Һŵ�=[2] And ҽ��״̬=1  And NVL(ִ�б��,0) <> -1 And Nvl(ִ������,0)<>0 And Rownum = 1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mstr�Һŵ�)
            If Not rsTmp.EOF Then
                MsgBox "�ò��˻���δ����ҽ����ֻ�н�����ҽ�����ͺ���ܽ���ת�", vbInformation, gstrSysName
                Exit Sub
            End If
        Else    'ֻҪ�¹�ҽ��(���������ϵ�)��˵��������Ϊ�ѷ�����������ת������¹Һ�
            strSql = "Select ID From ����ҽ����¼ Where ����ID+0=[1] And �Һŵ�=[2] And ҽ��״̬ <> 4 And Rownum = 1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mstr�Һŵ�)
            If Not rsTmp.EOF Then
                MsgBox "�Ѿ��Ըò����¹�ҽ����������ת���ɾ��������ҽ�����ٽ��У��������¹Һš�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    If Not frmRegistPlan.ShowMe(Me, mstr�Һŵ�, lng����ID, str����, strҽ��, lngҽ��ID) Then mblnUnRefresh = False: Exit Sub
    
    'ִ��ת��
    If Update���˹Һ�ת��(mlng����ID, mstr�Һŵ�, 1, lng����ID, str����, strҽ��, p����ҽ��վ) = False Then Exit Sub

    Call zlShowQuence(mstr�Һŵ�)
    'ˢ�½���
    Call LoadPatients("11001")
    Call SetReceiveToday(False, -1)
 
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlShowQuence(ByVal strNO As String)
    '����:��ʾ�ŶӽкŶ��е��º�
    On Error GoTo errH
    Dim colList As Collection
    If Check�Ŷӽк� = False Then Exit Sub
    '95637:���ϴ�,2016/7/20,������Ч���в���ʾ
    Set colList = ExseSvrQueuereginfo(2, "", "0,1,7", strNO, "", p����ҽ��վ)
    If colList Is Nothing Then Exit Sub
    If colList.Count = 0 Then Exit Sub
    MsgBox "ע��:" & vbCrLf & "    �ò������½������ŶӴ���,�Ӻ�Ϊ:[ " & GetColVal(colList(1), "_queue_num") & " ]", vbInformation + vbOKOnly, gstrSysName
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteTransferRefuse()
'���ܣ�ת��ܾ�
    On Error GoTo errH
    
    If mPr <> -1 Then
        If MsgBox("ȷʵҪ�ܾ���ת�ﲡ��""" & rptPati(mintRPTIndex).Rows(mPr).Record(COL_JZ_����).Value & """��", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        If Update���˹Һ�ת��(Val(rptPati(mintRPTIndex).Rows(mPr).Record(COL_JZ_����ID).Value), rptPati(mintRPTIndex).Rows(mPr).Record(COL_JZ_NO).Value, 3, 0, "", "", p����ҽ��վ) = False Then Exit Sub
    End If
    'ˢ�½���
    Call LoadPatients("11001")
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteTransferCancel(Optional ByVal blnMsg As Boolean = True)
'���ܣ�ȡ��ת��
    On Error GoTo errH
 
    With rptPati(mintRPTIndex).Rows(mPr)
        If blnMsg Then
            If MsgBox("ȷʵҪȡ������""" & .Record(COL_JZ_����).Value & """��ת����", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        End If
         If Update���˹Һ�ת��(Val(.Record(COL_JZ_����ID).Value), .Record(COL_JZ_NO).Value, 4, 0, "", "", p����ҽ��վ) = False Then Exit Sub
    End With
    
    'ˢ�½���
    Call LoadPatients("11011")
    Call ReshDataQueue
    Call SetReceiveToday(False, 1)
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteTransferIncept()
'���ܣ�����ת��
    On Error GoTo errH
    
    With rptPati(mintRPTIndex).Rows(mPr)
        If MsgBox(.Record(COL_JZ_ת��״̬).Value & vbCrLf & vbCrLf & "ȷ�Ͻ��ո�ת�ﲡ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If Update���˹Һ�ת��(Val(.Record(COL_JZ_����ID).Value), .Record(COL_JZ_NO).Value, 2, 0, "", IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��), p����ҽ��վ) = False Then Exit Sub
        If HaveRIS Then
            If gobjRis.HISModPati(1, mlng����ID, mlng�Һ�ID) <> 1 Then
                MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(HISModPati)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            End If
        ElseIf gbln����Ӱ����Ϣϵͳ�ӿ� = True Then
            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������RIS�ӿڴ���ʧ��δ����(HISModPati)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        End If
        Call mclsAdvices.zlRefresh(0, "", False) '87707
        'ˢ�²���λ����
        If rptPati(PATI_RPT����).Visible Then
            Call LoadPatients("11001", PATI_RPT����, .Record(COL_JZ_NO).Value)
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
    '����:���˽���
    '����:blnIsCard-�Ƿ���ˢ�����ý���ԤԼ����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim objControl As CommandBarControl
    Dim strSql As String
    Dim blnReserve As Boolean
    Dim datCurr As Date
    Dim int�Һ�ģʽ As Integer
    Dim bln�쳣���� As Boolean
   
    On Error GoTo errH

    datCurr = zlDatabase.Currentdate
    
    If (mintRPTIndex = PATI_RPT���� Or mintRPTIndex = PATI_RPTԤԼ) And mPr <> -1 Then
        If rptPati(mintRPTIndex).Rows(mPr).Record(COL_HZ_��ʶ).Value = "Ԥ" Then
            blnReserve = True
            If rptPati(mintRPTIndex).Rows(mPr).Record.Tag = "��" Then
                bln�쳣���� = True
            End If
        End If
    Else
        Exit Sub
    End If
    
    If blnReserve Then
        '��ԤԼ�ҺŲ��˽��н���
        '�����:57566
        If Check�������("����", mstr�Һŵ�) = False Then Exit Sub
        
        
        If mobjPatient Is Nothing Then
            On Error Resume Next
            Set mobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
            Err.Clear: On Error GoTo 0
            On Error GoTo errH
            If mobjPatient Is Nothing Then
                MsgBox "����������Ϣ����������zlPublicPatient.clsPublicPatient��ʧ�ܣ�", vbInformation, Me.Caption
            Else
                Call mobjPatient.zlInitCommon(gcnOracle, glngSys, UserInfo.�û���)
            End If

        End If
        
        
        Call mobjPatient.zlAutoCalcBackLists(mlng����ID)
        If mobjPatient.zlCheckBlackListValied(p����ҽ��վ, mlng����ID, "�Һ�") = False Then Exit Sub
        Call InitObjPublicExpense
        '����ҽ��վԤԼ����ʱ���ùҺŲ����Ľ��սӿڽ��п۷ѵĹ���
        int�Һ�ģʽ = Val(zlDatabase.GetPara("�Һ�ģʽ", glngSys, 9000, 1))
        If int�Һ�ģʽ = 0 And Not gobjPublicExpense Is Nothing Then
            If bln�쳣���� Then
                Call gobjPublicExpense.zlRegisterIncept(Me, mlngModul, mstr�Һŵ�, mstr��������, PatiIdentify.objIDKind.GetCurCard.�ӿ����, PatiIdentify.Text): Exit Sub
            Else
                If Not gobjPublicExpense.zlRegisterIncept(Me, mlngModul, mstr�Һŵ�, mstr��������, PatiIdentify.objIDKind.GetCurCard.�ӿ����, PatiIdentify.Text) Then Exit Sub
            End If
        ElseIf int�Һ�ģʽ = 2 And Not gobjPublicExpense Is Nothing Then
            If bln�쳣���� Then
                Call gobjPublicExpense.zlRegisterIncept(Me, mlngModul, mstr�Һŵ�, mstr��������, PatiIdentify.objIDKind.GetCurCard.�ӿ����, PatiIdentify.Text): Exit Sub
            Else
                If ZLCommFun.ShowMsgBox("��ѡ��", "��ѡ���˵�֧����ʽ,����֧�����ߵ��շѴ���֧����", "!����֧��(&Y),?����֧��(&N)", Me, vbQuestion) = "����֧��" Then
                    If Not gobjPublicExpense.zlRegisterIncept(Me, mlngModul, mstr�Һŵ�, mstr��������, PatiIdentify.objIDKind.GetCurCard.�ӿ����, PatiIdentify.Text) Then Exit Sub
                Else
                    If Update����ԤԼ����(mstr�Һŵ�, mstr��������, datCurr, IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��), IIf(mstr����ҽ����� = "", UserInfo.���, mstr����ҽ�����), p����ҽ��վ) = False Then Exit Sub
                End If
            End If
        Else
            If Update����ԤԼ����(mstr�Һŵ�, mstr��������, datCurr, IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��), IIf(mstr����ҽ����� = "", UserInfo.���, mstr����ҽ�����), p����ҽ��վ) = False Then Exit Sub
        End If
    Else
        '�����:57566
        If Check�������("����", mstr�Һŵ�) = False Then Exit Sub
                'ת�ﲡ��ֱ�ӵ���ת�����
        If mintActive = ptת�� Then
            ExecuteTransferIncept
            Exit Sub
        End If
        '�������ҺŲ��˽��н���
        strSql = "Select ִ���� From ���˹Һż�¼ Where ����ID+0=[1] And NO=[2] And Nvl(ִ��״̬,0)<>0 And ��¼����=1 And ��¼״̬=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mstr�Һŵ�)
        If Not rsTmp.EOF Then
            MsgBox "�ò�������" & IIf(IsNull(rsTmp!ִ����), "����ҽ��", "ҽ����" & rsTmp!ִ���� & " ") & "���", vbInformation, gstrSysName
            Call LoadPatients("10001"): Exit Sub
        End If
        
        strSql = "Select ִ���� From ���˹Һż�¼ Where ����ID+0=[1] And NO=[2] And Nvl(ִ��״̬,0)=0 And ��¼����=1 And ��¼״̬=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mstr�Һŵ�)
        If rsTmp.EOF Then
            MsgBox "�ò������˺ţ����ܽ��", vbInformation, gstrSysName
            Call LoadPatients("10001"): Exit Sub
        End If
        
        If Update���˹ҺŽ���(mlng����ID, mstr�Һŵ�, 0, UserInfo.����, mstr��������, 0, 0, datCurr, p����ҽ��վ) = False Then Exit Sub
    
    End If
    
    If mblnAutoHandle Then Call Tip�����Զ����
    
    'ˢ�²���λ����
    On Error GoTo 0
    
    tbcInTreat.Item(t����).Selected = True
    Call LoadPatients("11001", PATI_RPT����, mstr�Һŵ�)
    

    '���������Զ����ù���
    If Not gobjCommunity Is Nothing And mlngCommunityID <> 0 And mlng����ID <> 0 And mPatiInfo.���� <> 0 Then
        Set objControl = cbsMain.FindControl(, mlngCommunityID, , True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then objControl.Execute
        End If
    End If
    
    Call ReceiveAfterExec
    
    '�����ŶӽкŶ���(����ˢ��)
    Call ReshDataQueue
    Call SetReceiveToday(False, 1)
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ReceiveAfterExec(Optional ByVal bln���� As Boolean)
'���ܣ��������Ҫ���õĲ���
    Dim objControl As CommandBarControl
    
    Call CreatePlugInOK(p����ҽ��վ)
    '����������ҽӿ�
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.ClinicReceive(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID)
        Call zlPlugInErrH(Err, "ClinicReceive")
        Err.Clear: On Error GoTo errH
    End If
    
    '����֮���Զ�����ҽ���´�״̬
    If mlng�Զ����� = 1 And bln���� = False Then
        Call LocatedCard("ҽ��")
        cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
        Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then objControl.Execute
        End If
    ElseIf mlng�Զ����� = 2 And bln���� = False Then
        If GetInsidePrivs(p�°����ﲡ��, True) <> "" And Not mclsEMR Is Nothing Then
            Call LocatedCard("�²���")
            cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
        Else
            Call LocatedCard("����")
            cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
            mblnUnRefresh = True
            Call mclsEPRs.zlOpenDefaultEPR(mstr�Һŵ�)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteCancel()
'���ܣ�ȡ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim i As Long, bytOut As Byte
        
    If BillExpend(mstr�Һŵ�) Then
        MsgBox "�ò��˹Һ��ѳ�����Ч��������������ȡ�����", vbInformation, gstrSysName
        Exit Sub
    End If
        
    On Error GoTo errH
    
    'ֻ��ȡ���Լ�����Ĳ���
    strSql = "Select ִ���� From ���˹Һż�¼ Where id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPatiInfo.�Һ�ID)
    If rsTmp!ִ���� <> UserInfo.���� Then
        MsgBox "ֻ��ȡ���Լ�����Ĳ��ˡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'ToDo:ȡ������ʱ�������ݵļ��
    'ҽ�����ݵļ��
    strSql = "Select Count(*) as ҽ�� From ����ҽ����¼ Where ҽ��״̬ IN(1,8) And ����ID+0=[1] And �Һŵ�=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mstr�Һŵ�)
    If Nvl(rsTmp!ҽ��, 0) > 0 Then
        MsgBox "�ò��������¿����ѷ��͵�ҽ��������ȡ�����" & vbCrLf & _
            "���ȷʵҪȡ��������Ƚ���Щҽ��ɾ�������ϡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mbln��Һ�ģʽ Then
        If zlRegisterPriceDeleteFromNO(mclsReg, mstr�Һŵ�, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID), _
                IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��), mlng����ID) = False Then
            Exit Sub
        End If
    Else
        If Update����ȡ������(mlng����ID, mstr�Һŵ�, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID), IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��), p����ҽ��վ) = False Then Exit Sub
    End If

    
    'ˢ�²���λ����
    If rptPati(PATI_RPT����).Visible Then
        Call LoadPatients("11001", PATI_RPT����, mstr�Һŵ�)
    ElseIf rptPati(PATI_RPTԤԼ).Visible Then
        Call LoadPatients("11001", PATI_RPTԤԼ, mstr�Һŵ�)
    Else
        Call LoadPatients("11001")
    End If
    Call ReshDataQueue
    Call SetReceiveToday(False, -1)
    Exit Sub
errH:

    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteFinish()
'���ܣ���ɽ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, blnTran As Boolean
    Dim str����IDs As String, str���IDs As String
    Dim lng�Һ�id As Long
    Dim str���� As String
    Dim str״̬ As String
    Dim lngSelectedIndex As Long
    Dim rptRow As ReportRow
    
    On Error GoTo errH
 
    If (mintRPTIndex = PATI_RPT���� Or mintRPTIndex = PATI_RPT����) And mPr <> -1 Then
        With rptPati(mintRPTIndex).Rows(mPr)
            str���� = .Record(COL_JZ_����).Value
            str״̬ = .Record(COL_JZ_״̬).Value
            lngSelectedIndex = .Record.Index
        End With
    Else
        Exit Sub
    End If
    
    '����б�ʱ�䲻ˢ�²����������
    strSql = "select 1 from ���˹Һż�¼ where no=[1] and ִ����=[2] And ִ��״̬=2 And ��¼����=1 And ��¼״̬=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstr�Һŵ�, mstr����ҽ��)
    If rsTmp.EOF Then
        MsgBox """" & str���� & """���ܱ�����ҽ��ǿ��������գ������ԡ�", vbInformation, gstrSysName
        Call LoadPatients
        Call ReshDataQueue
        Exit Sub
    End If
    
    'ToDo:��ɽ���ʱ�������ݵļ��
    If str״̬ = "0" Then
        If MsgBox("��ǰ����""" & str���� & """�Ѿ�ת��Ƿ�Ҫȡ��ת�������ɽ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        Else
            Call ExecuteTransferCancel(False)
            Call ExecuteFinish
            Exit Sub
        End If
    End If
    
    '����Ƿ������Чҽ��
    strSql = "Select Count(*) as ҽ�� From ����ҽ����¼ Where ����ID+0=[1] And �Һŵ�=[2] And ҽ��״̬<>4"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mstr�Һŵ�)
    If Nvl(rsTmp!ҽ��, 0) = 0 Then
        If MsgBox("δ��""" & str���� & """�´��κ���Ч��ҽ����ȷʵҪ��ɽ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    '����Ƿ����δ���͵�ҽ��
    strSql = "Select Count(*) as ҽ�� From ����ҽ����¼ Where ����ID+0=[1] And �Һŵ�=[2] And ҽ��״̬=1 And NVL(ִ�б��,0) <> -1 And Nvl(ִ������,0)<>0 And Nvl(Ƥ�Խ��,'��')<>'����'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mstr�Һŵ�)
    If Nvl(rsTmp!ҽ��, 0) > 0 Then
        MsgBox """" & str���� & """����δ���͵�ҽ����������ɽ��", vbInformation, gstrSysName
        Exit Sub
    End If
    '���δ��д�ļ���֤������
    strSql = "Select ��ҳID,����ID,���ID From ������ϼ�¼ Where ȡ��ʱ�� is Null And ����ID=[1] And ��ҳID=(Select ID From ���˹Һż�¼ Where NO=[2] And ��¼����=1 And ��¼״̬=1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mstr�Һŵ�)
    Do While Not rsTmp.EOF
        If lng�Һ�id = 0 Then lng�Һ�id = rsTmp!��ҳID
        If Not IsNull(rsTmp!����id) Then str����IDs = str����IDs & "," & rsTmp!����id
        If Not IsNull(rsTmp!���ID) Then str���IDs = str���IDs & "," & rsTmp!���ID
        rsTmp.MoveNext
    Loop
    If str����IDs <> "" Or str���IDs <> "" Then
        If Not CheckDiseaseFile(Me, mlng����ID, lng�Һ�id, mlng�������ID, Mid(str����IDs, 2), Mid(str���IDs, 2), , True, , 1) Then Exit Sub
    End If
    
    If lng�Һ�id = 0 Then lng�Һ�id = mPatiInfo.�Һ�ID
    
    If Not ExecuteFinishInSide(mstr�Һŵ�, mlng����ID, lng�Һ�id) Then
        Exit Sub
    End If

    'ˢ��:����λ�������б�
    Call LoadPatients
    Call ReshDataQueue
    
     '��ɽ���֮���Զ���λ����һ��
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

Private Function ExecuteFinishInSide(ByVal strNO As String, ByVal lng����ID As Long, ByVal lng�Һ�id As Long) As Boolean
'���ܣ���ɾ���
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim blnTran As Boolean
    Dim str������ As String
    Dim col���� As Collection
    
    On Error GoTo errH
    
    '��ȡ��Ҫ����Ϣ�������ӿڵ���:����߾��ﲡ�˱��ξ���Ϊ׼,�ұ߿��ܵ�ǰѡ�����ʷ����
    strSql = "Select A.ID,A.����ID,A.���� From ���˹Һż�¼ A Where  A.��¼����=1 And A.��¼״̬=1 And A.NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO)
    
    'ִ�й���
    '-----------------------------------
    If Update������ɽ���(lng����ID, strNO, mstr��������, UserInfo.����, "", 0, p����ҽ��վ) = False Then
        Exit Function
    End If

    If Not gobjCommunity Is Nothing And Nvl(rsTmp!����, 0) <> 0 Then
        '��������������Ϣ�ύ
    
        Set col���� = PatiSvrGetCommunityInfo(1, Val(rsTmp!����ID & ""), Val(rsTmp!���� & ""), p����ҽ��վ)
        str������ = ""
        If Not col���� Is Nothing Then
            If col����.Count > 0 Then
                  str������ = GetColVal(col����(1), "_community_code")
            End If
        End If


        If Not gobjCommunity.ClinicSubmit(glngSys, mlngModul, rsTmp!����, str������, lng����ID, rsTmp!ID) Then
            Exit Function
        End If
    End If

    '����������ҽӿ�
    Call CreatePlugInOK(p����ҽ��վ)
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.ClinicFinish(glngSys, p����ҽ��վ, lng����ID, lng�Һ�id)
        Call zlPlugInErrH(Err, "ClinicFinish")
        Err.Clear: On Error GoTo errH
    End If
    
    'һ��ͨ�����ϴ�
    'If Not mobjICCard Is Nothing Then
     '   strSQL = "Select 1 From һ��ͨĿ¼ Where ����=2 And Rownum=1"
      '  Set rsTmp = New ADODB.Recordset
       ' Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        'If Not rsTmp.EOF Then
         '   mobjICCard.UploadSwap lng����ID, ""
        'End If
    'End If
        ExecuteFinishInSide = True
    Exit Function
errH:

    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub ExecuteRedo()
'�ָ�����
    'ֻ����������ݱ��е�
    If BillExpend(mstr�Һŵ�) Then
        MsgBox "�ò��˹Һ��ѳ�����Ч�������������ٻָ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mintRPTIndex = PATI_RPT���� Then
        If zlDatabase.NOMoved("���˹Һż�¼", mstr�Һŵ�) Then
            MsgBox "�ùҺż�¼�Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '��ǰҽ����ɵĲ��˲ſ���ֱ�ӻָ�(������Ȩ�޿���ǿ������)
    If mintRPTIndex = PATI_RPT���� Then
        With rptPati(PATI_RPT����).Rows(mPr)
            If .Record(COL_YZ_ִ����).Value <> UserInfo.���� Then
                MsgBox "�ò��˲���������ɾ���ģ�����ֱ�ӻָ����", vbInformation, gstrSysName
                Exit Sub
            End If
        End With
    End If
    
    On Error GoTo errH
   
    If Update����ȡ����ɽ���(mlng����ID, mstr�Һŵ�, 0, p����ҽ��վ) = False Then
        Exit Sub
    End If
    
    'ˢ�²���λ���ˣ�����ǲ�����ƬҪ�ֶ�ˢ��һ�·����еİ�ť������
    If tbcSub.Selected.Tag = "����" Then Call mclsEPRs.zlRefresh(0, 0, 0, False, False)
    If tbcInTreat.Item(t����).Visible Then
        tbcInTreat.Item(t����).Selected = True
    End If
    If rptPati(PATI_RPT����).Visible Then
        Call LoadPatients("011", PATI_RPT����, mstr�Һŵ�)
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
'���ܣ��������������֤
    Dim colInfo As New Collection
    Dim int���� As Integer, str������ As String
        
    If gobjCommunity Is Nothing Or mPatiInfo.����ID = 0 Or mPatiInfo.�Һ�ID = 0 Or mPatiInfo.���� <> 0 Then Exit Sub
    
    If Not gobjCommunity.Identify(glngSys, p����ҽ��վ, int����, str������, colInfo, mPatiInfo.����ID, mPatiInfo.�Һ�ID) Then Exit Sub

    If Update����������Ϣ(mPatiInfo.����ID, mPatiInfo.�Һ�ID, int����, str������, 1, zlDatabase.Currentdate, colInfo, p����ҽ��վ) = False Then Exit Sub
    On Error GoTo 0
    Call LoadPatients����
    If Not mbln��ʾԤԼ���� Then
        Call LoadPatientsԤԼ
    End If
    Call ReshDataQueue
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function GetColItem(colInfo As Collection, strItem As String) As String
    If colInfo Is Nothing Then Exit Function
    
    Err.Clear: On Error Resume Next
    GetColItem = colInfo("_" & strItem)
    Err.Clear: On Error GoTo 0
End Function

Private Sub SetRoomState(ByVal blnBusy As Boolean)
'���ܣ���������æ��״̬
    On Error GoTo DBError
    gcnOracle.Execute "Update �������� Set ȱʡ��־=" & IIf(blnBusy, 1, 0) & " Where ����='" & mstr�������� & "' And ȱʡ��־<>" & IIf(blnBusy, 1, 0)
    On Error GoTo 0
    
    Me.stbThis.Panels(4).Text = "����" + IIf(blnBusy, "æ", "��")
    Me.lblRoom.BackColor = IIf(blnBusy, COLOR_BUSY, COLOR_FREE)
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetReceiveToday(ByVal blnDo As Boolean, ByVal intStep As Integer)
'���ܣ����ս�������
'������blnDo true-�������ݿ⣬false ���������ݿ⡣intStep ��������1��1
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If blnDo Then
        strSql = "select count(1) as ���� from ���˹Һż�¼ a where a.��¼״̬=1 and a.ִ����=[1] and  a.ִ��ʱ�� between Trunc(Sysdate) and Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.����)
        mlng���ս������� = Val(rsTmp!���� & "")
    Else
        mlng���ս������� = mlng���ս������� + intStep
        If mlng���ս������� < 0 Then mlng���ս������� = 0
    End If
    
    Me.stbThis.Panels(3).Text = "���ս���" & mlng���ս������� & "��"
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetTimer()
    mintRefresh = Val(zlDatabase.GetPara("����ˢ�¼��", glngSys, p����ҽ��վ, 180))
    If mintRefresh <> 0 And mintRefresh < 30 Then mintRefresh = 30
    If mintRefresh = 0 Then
        timRefresh.Enabled = False
    Else
        timRefresh.Interval = 1000 '�̶�Ϊ1����
        timRefresh.Enabled = True
    End If
End Sub

Private Sub timRefresh_Timer()
    Static lngSecond As Long
    Static strPreTime1 As String
    Dim curTime As Date
    
    If mbln��Ϣ���� Then
        If Not mrsMsg Is Nothing Then
            If mrsMsg.RecordCount > 0 Then
                timRefresh.Enabled = False
                Call mclsMsg.PlayMsgSound(mrsMsg)
                Set mrsMsg = Nothing
                timRefresh.Enabled = True
            End If
        End If
    End If
    
        
    curTime = Now
    
    'ˢ�²����������
    If mintNotify > 0 And rptNotify.Visible Then
        If strPreTime1 = "" Then
            strPreTime1 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
        End If
        If DateDiff("s", CDate(strPreTime1), curTime) > mintNotify * CLng(60) Then
            strPreTime1 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
            Call LoadNotify
            If mblnΣ��ֵ���� And mblnΣ��ֵshow = False Then Call ReadMsgAuto
        End If
    End If
    
    If mintRefresh = 0 Or mblnUnRefresh Or Me.hwnd <> GetForegroundWindow Then Exit Sub
    lngSecond = lngSecond + 1 '����
    If lngSecond Mod mintRefresh = 0 Then
        lngSecond = 0
        Call LoadPatients����
        If Not mbln��ʾԤԼ���� Then
            Call LoadPatientsԤԼ
        End If
        Call ReshDataQueue
    End If
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean, Optional ByVal strIDCard As String, Optional ByVal blnIsCard As Boolean _
                            , Optional ByVal lngPatiID As Long)
'���ܣ�����(��һ��)����
'������blnNext=�Ƿ������һ��
'      strIDCard=����ֵʱ����ʾ�̶������֤�Ų���
'      blnIsCard=�Ƿ���ˢ�����ý���ԤԼ����
    Static blnReStart As Boolean
    Dim intIdx As PatiType, i As Long
    Dim objControl As CommandBarControl
    Dim blnQueueFind As Boolean
    Dim objRpt As ReportControl
    Dim lngCol As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    If mintRPTIndex = -1 Then PatiIdentify.Text = "": Exit Sub
    
    '��������ʽ���Һ��Զ�ˢ���֤�ļ���������ȡ��
    If strIDCard = "" And PatiIdentify.Text <> "" Then mstrIDCard = ""
    
    If Not blnNext And mstrFindType = "�Һŵ�" Then
        PatiIdentify.Text = GetFullNO(PatiIdentify.Text, 12)
    End If
    PatiIdentify.SetFocus
    
    Set objRpt = rptPati(mintRPTIndex)
    
    '��ʼ������
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
    
    Call InitObjOneCardComLib(Me, p����ҽ��վ)
    
     '���Ҳ���
    If lngPatiID = 0 And Not gobjOneCardComLib Is Nothing And mstrFindType <> "���￨" And mstrFindType <> "��ʶ��" And mstrFindType <> "�Һŵ�" And mstrFindType <> "����" And mstrFindType <> "�������֤" Then
        If mstrFindType = "IC��" Then
            Call gobjOneCardComLib.zlGetPatiID("IC��", PatiIdentify.Text, , lngPatiID)
        Else
            Call gobjOneCardComLib.zlGetPatiID(Val(PatiIdentify.objIDKind.GetCurCard.�ӿ����), PatiIdentify.Text, , lngPatiID)
        End If
    End If
    
    '���Ҳ���
    If Check�Ŷӽк� = True Then
        blnQueueFind = mobjQueue.FindQueue(IIf(PatiIdentify.objIDKind.GetCurCard.�ӿ���� > 0, _
                            PatiIdentify.objIDKind.GetCurCard.�ӿ����, _
                            IIf(PatiIdentify.objIDKind.GetCurCard.���� = "��ʶ��", "�����", PatiIdentify.objIDKind.GetCurCard.����)), _
                            PatiIdentify.Text)
    End If
    If blnQueueFind = False Then
        For intIdx = intIdx To 4
            Set objRpt = rptPati(intIdx)
            For i = i To objRpt.Rows.Count - 1
                With objRpt.Rows(i)
                    If strIDCard <> "" Then '���֤�Զ�ʶ��ǿ������
                        lngCol = IIf(intIdx = PATI_RPT���� Or intIdx = PATI_RPTԤԼ, COL_HZ_���֤��, IIf(intIdx = PATI_RPT����, COL_YZ_���֤��, COL_JZ_���֤��))
                        If UCase(.Record(lngCol).Value) = UCase(strIDCard) Then Exit For
                    Else
                        lngCol = IIf(intIdx = PATI_RPT���� Or intIdx = PATI_RPTԤԼ, COL_HZ_����ID, IIf(intIdx = PATI_RPT����, COL_YZ_����ID, COL_JZ_����ID))
                        If Val(.Record(lngCol).Value) = lngPatiID And lngPatiID <> 0 Then Exit For
                        Select Case mstrFindType
                            Case "���￨"
                                lngCol = IIf(intIdx = PATI_RPT���� Or intIdx = PATI_RPTԤԼ, COL_HZ_���￨��, IIf(intIdx = PATI_RPT����, COL_YZ_���￨��, COL_JZ_���￨��))
                                If .Record(lngCol).Value = PatiIdentify.Text Then Exit For
                            Case "��ʶ��"
                                lngCol = IIf(intIdx = PATI_RPT���� Or intIdx = PATI_RPTԤԼ, COL_HZ_�����, IIf(intIdx = PATI_RPT����, COL_YZ_�����, COL_JZ_�����))
                                If .Record(lngCol).Value = PatiIdentify.Text Then Exit For
                            Case "�Һŵ�"
                                lngCol = IIf(intIdx = PATI_RPT���� Or intIdx = PATI_RPTԤԼ, COL_HZ_NO, IIf(intIdx = PATI_RPT����, COL_YZ_NO, COL_JZ_NO))
                                If UCase(.Record(lngCol).Value) = UCase(PatiIdentify.Text) Then Exit For
                            Case "����"
                                lngCol = IIf(intIdx = PATI_RPT���� Or intIdx = PATI_RPTԤԼ, COL_HZ_����, IIf(intIdx = PATI_RPT����, COL_YZ_����, COL_JZ_����))
                                If .Record(lngCol).Value Like "*" & PatiIdentify.Text & "*" Then Exit For
                            Case "�������֤"
                                lngCol = IIf(intIdx = PATI_RPT���� Or intIdx = PATI_RPTԤԼ, COL_HZ_���֤��, IIf(intIdx = PATI_RPT����, COL_YZ_���֤��, COL_JZ_���֤��))
                                If UCase(.Record(lngCol).Value) = UCase(PatiIdentify.Text) Then Exit For
                            Case "IC��"
                                lngCol = IIf(intIdx = PATI_RPT���� Or intIdx = PATI_RPTԤԼ, COL_HZ_IC����, IIf(intIdx = PATI_RPT����, COL_YZ_IC����, COL_JZ_IC����))
                                If UCase(.Record(lngCol).Value) = UCase(PatiIdentify.Text) Then Exit For
                            Case Else
                                lngCol = IIf(intIdx = PATI_RPT���� Or intIdx = PATI_RPTԤԼ, COL_HZ_����ID, IIf(intIdx = PATI_RPT����, COL_YZ_����ID, COL_JZ_����ID))
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
                If intIdx = PATI_RPT���� Then
                    tbcInTreat.Item(t����).Selected = True
                ElseIf intIdx = PATI_RPT���� Then
                    tbcInTreat.Item(t����).Selected = True
                ElseIf intIdx = PATI_RPT���� Then
                    tbcInTreat.Item(t����).Selected = True
                ElseIf intIdx = PATI_RPT���� Then
                    tbcWait.Item(0).Selected = True
                ElseIf intIdx = PATI_RPTԤԼ Then
                    If tbcWait.Item(mintԤԼ�б�).Visible Then
                        tbcWait.Item(mintԤԼ�б�).Selected = True
                    End If
                End If
            End If

            '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
            Set objRpt.FocusedRow = objRpt.Rows(i)
            If objRpt.Visible Then objRpt.SetFocus
            
            '�ҵ����Զ����н���,ԤԼ�����Զ�����
            If mbln�Զ����� And (intIdx = PATI_RPT���� Or intIdx = PATI_RPTԤԼ) Then
                cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
                strTmp = objRpt.Rows(i).Record(COL_HZ_��ʶ).Value
                If strTmp = "Ԥ" Then
                    If mstrFindType = "��ʶ��" Or mstrFindType = "�Һŵ�" Or mstrFindType = "����" Or mstrFindType = "�������֤" Then Exit Sub
                    Call ExecuteReceive(blnIsCard)
                Else
                    Set objControl = cbsMain.FindControl(, conMenu_Manage_Receive, True, True)
                    If Not objControl Is Nothing Then
                        If objControl.Enabled Then Call cbsMain_Update(objControl) '�״�ִ�У�û����ʾ�˵�ǰ���¼�û��ִ��
                        If objControl.Enabled Then objControl.Execute
                    End If
                End If
            End If
        Else
            blnReStart = True
            MsgBox IIf(blnNext, "������", "") & "�Ҳ������������Ĳ��ˡ�", vbInformation, gstrSysName
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
      
Private Function Check�Ŷӽк�() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ʹ����ŶӽкŹ���
    '���أ��ŶӽкŹ������еĶ��Ϸ�,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-06-06 10:19:43
    '˵��������: Ȩ�޺Ϸ����;�������Ŷӽкŵ�;�����Ŷӽкųɹ�!
    '------------------------------------------------------------------------------------------------------------------------
    '�ŶӽкŴ���ģʽ:1.�������̨������л�ҽ����������;2-�ȷ������,��ҽ�����о���.0-���Ŷӽк�
    If mty_Queue.byt�Ŷӽк�ģʽ = 0 Then GoTo GOEND:
    If Not (InStr(mty_Queue.strQueuePrivs, ";����;") > 0) Then GoTo GOEND:
    If mty_Queue.blnҽ���������� = False And mty_Queue.byt�Ŷӽк�ģʽ = 1 Then GoTo GOEND:
    
    Err = 0: On Error GoTo GOEND:
    If mobjQueue Is Nothing Then
        Set mobjQueue = CreateObject("zlQueueManage.clsQueueManage")
        Err = 0: On Error GoTo errHand:
        mobjQueue.zlInitVar gcnOracle, glngSys, 0, IIf(gint����Һ����� = 0, 1, gint����Һ�����), mty_Queue.strQueuePrivs, CStr(mlngModul), False
        mobjQueue.zlSetToolIcon 24, True
        mobjQueue.IsShowFindTools = False
    End If
    Check�Ŷӽк� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
GOEND:
    If Not mobjQueue Is Nothing Then mobjQueue.CloseWindows
    Set mobjQueue = Nothing

End Function

Private Sub ReshDataQueue()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ���Ŷӽк�����
    '���ƣ����˺�
    '���ڣ�2010-06-07 15:27:57
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    Dim str���� As String, strҽ�� As String, str���� As String
    Dim intType As Integer
    Dim colList As Collection
    Dim i As Long
    
    On Error GoTo errH
    If mobjQueue Is Nothing Then Exit Sub
    If Check�Ŷӽк� = False Then Exit Sub
    '��ȡ��صĶ�������
    '���ﷶΧ��1=�ұ��˺ŵĲ���,2=�����Ҳ���,3=�����Ҳ���
    mint���ﷶΧ = Val(zlDatabase.GetPara("���ﷶΧ", glngSys, p����ҽ��վ, "2"))
    Dim strQueue() As String
    
    ReDim Preserve strQueue(1 To 1) As String
    str���� = IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID)
    strQueue(1) = str����
    strҽ�� = IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��)
    str���� = mstr��������
    intType = 1
    Select Case mint���ﷶΧ
    Case 1   '1=�ұ��˺ŵĲ���
        If Not mty_Queue.blnҽ���������� Then
           strҽ�� = UserInfo.����  '64696,������,2014-01-08,�õ�¼��Ա�����������ŶӽкŶ���
        End If
        If mlng�������ID = 0 Then strQueue(1) = ""
        intType = 3
    Case 2  '2=�����Ҳ���
        If Not mty_Queue.blnҽ���������� Then
           str���� = mstr��������
        End If
        If mlng�������ID = 0 Then strQueue(1) = ""
        intType = 2
    Case 3  '3=�����Ҳ���
    End Select
    
    '��Ҫ�Ŷ�û�н����Ĳ���
    Set colList = ExseSvrQueuereginfo(3, "", "", "", str����, p����ҽ��վ)
    
    strTemp = ""
    If Not colList Is Nothing Then
        For i = 1 To colList.Count
            strTemp = strTemp & "," & Val(GetColVal(colList(i), "_reg_id"))
        Next
    End If
    
    If strTemp <> "" Then strTemp = "0|" & Mid(strTemp, 2)
    
    Call mobjQueue.zlRefresh(strQueue, mty_Queue.strCurrQueueName, mty_Queue.lngcurr�Һ�ID, str����, strҽ��, strTemp, intType)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

 
Private Sub zlQueueStartus(intType As Integer, strNO As String, lng����ID As Long)
  '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ܲ�����,
    '��Σ�2-����;3-���˲�����;4-���˴���;5-������ɾ���;6-����ȡ������
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-06-03 14:15:46
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strQueueName As String, lngID As Long
    Dim strSql As String, rsTemp As New ADODB.Recordset
    If Check�Ŷӽк� = False Then Exit Sub
    On Error GoTo errH
    strSql = "SELECT ID,ִ�в���ID,����,ִ���� From ���˹Һż�¼ where NO=[1] And ��¼����=1 And ��¼״̬=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Sub
    
    strQueueName = Nvl(rsTemp!ִ�в���ID)
    If Nvl(rsTemp!ִ����) <> "" Then
        strQueueName = strQueueName & ":" & Nvl(rsTemp!ִ����)
    ElseIf Nvl(rsTemp!����) <> "" Then
        strQueueName = strQueueName & ":" & Nvl(rsTemp!����)
    End If
    
    lngID = Val(Nvl(rsTemp!ID))
    Select Case intType
    Case 3   ' ���˲�����;
        ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 3
    Case 4, 6   '���˴���,'����ȡ������
        ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 0
    Case 5  '������ɾ���
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 4
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Set���˹Һ�״̬(ByVal lngState As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ò��˹Һ�״̬
    '��Σ�lngState : -1- ���˲�����
    '                         0-���˴���
    '���Σ�
    '���أ��Ƿ����óɹ������˲�����ʱ����ɾ�����۵��ݣ����ٴ����ô���ʱ�����ò��ɹ� ����False ,�����������True
    '���ƣ����˺�
    '���ڣ�2010-06-03 15:24:48
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim str����NO As String
    
    If mstr�Һŵ� = "" Then Exit Function
    
    On Error GoTo errH
    
    If lngState = -1 Then
        '��鲡���Ƿ������Ч��ҽ��
        strSql = "Select 1 From ����ҽ����¼ Where ����id = [1] And �Һŵ� = [2]  And ҽ��״̬ <> -1 And ҽ��״̬ <> 1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mstr�Һŵ�)
        If Not rsTmp.EOF Then
            MsgBox "�ò��˴�����Чҽ��,��������Ϊ������!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    End If

    
     If Update���˹Һ�״̬(mstr�Һŵ�, lngState, p����ҽ��վ) = False Then Exit Function
    
    Call zlQueueStartus(IIf(lngState = -1, 3, 4), mstr�Һŵ�, mlng����ID)
    'intType:intType:1-����;2-����;3-���˲�����;4-���˴���;5-������ɾ���;6-�ָ�����
    ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
    MsgBox "�����ɹ�!", vbInformation, gstrSysName
    
    Set���˹Һ�״̬ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ExecuteStopAndReuse(ByVal bln���� As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ծ��ﲡ�˽�����ͣ������������
    '���:bln����-true:�����Ѿ�ͣ�õľ��ﲡ��
    '����:
    '����:
    '����:���˺�
    '����:2010-12-08 20:26:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, bln��ͣ As Boolean
    Dim strNO As String, rsTemp As ADODB.Recordset
    Dim lngSelectedIndex As Long
    Dim rptRow As ReportRow
    
    lngSelectedIndex = mPr
    With rptPati(mintRPTIndex).Rows(mPr)
        bln��ͣ = Val(.Record(COL_JZ_��¼��־).Value) = 2
        If bln���� Then
            If bln��ͣ = False Then
                MsgBox "ע��:" & vbCrLf & "    �ò��˻�δ��ͣ����,���ܽ��лָ���ͣ����!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Sub
            End If
        Else
            If bln��ͣ Then
                MsgBox "ע��:" & vbCrLf & "    �ò��˻�������ͣ����,���ܽ�����ͣ����!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Sub
            End If
        End If
        strNO = .Record(COL_JZ_NO).Value
        
        strSql = "Select ID From ���˹Һż�¼ where NO=[1] And ��¼����=1 And ��¼״̬=1"
        On Error GoTo errHandle
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO)
        If rsTemp.EOF Then
            Exit Sub
        End If
    End With
    
    If Not bln���� Then
        '����
        If Update���˻���(Val(rsTemp!ID & ""), 0, "", "", 1, "", p����ҽ��վ) = False Then Exit Sub
    Else
        'ȡ������
        If Update����ȡ������(Val(rsTemp!ID & ""), 1, p����ҽ��վ) = False Then Exit Sub
    End If
    
    On Error GoTo errHandle
    If bln���� Then
        'ȡ������ʱת�������б���
        If tbcInTreat.Item(t����).Visible Then
            tbcInTreat.Item(t����).Selected = True
        End If
        If rptPati(PATI_RPT����).Visible Then
            Call LoadPatients("0101", PATI_RPT����, mstr�Һŵ�)
        Else
            Call LoadPatients("0101")
        End If
    Else
        '��ǻ�����Զ���λ��һ������
        Call LoadPatients("0101")
        '������֮���Զ���λ����һ��
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

Private Sub SetFontSize(ByVal blnSetMainFont As Boolean)
'����: ���н��������ͳһ����
'����: blnSetMainFont �Ƿ���������������(���������ӽ����л�)
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
        Case "����"
            Call mobjPati.SetFontSize(mbytSize)
        Case "ҽ��"
            Call mclsAdvices.SetFontSize(mbytSize)
        Case "����"
            Call mclsEPRs.SetFontSize(mbytSize)
                Case "�²���"
            On Error Resume Next
            Call mclsEMR.SetFontSize(mbytSize)
            Err.Clear: On Error GoTo 0
    End Select
    
    If tbcRegist.Selected.Tag = "����һ��" Then
        Call mfrmView.SetFontSize(mbytSize)
    End If
        
End Sub

Private Function Check�������(str���� As String, strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���˽������
    '���:str���� -��ǰ���� strNo - �Һŵ��ݺ�
    '����:
    '����:
    '����:����
    '����:2013-1-17 20:26:59
    '�����:57566
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHanl:
    Dim strSql As String
    Dim rsTemp As Recordset
    Dim strMsg As String
    
    If mlng������� = 0 Then Check������� = True: Exit Function
    
    strSql = "" & _
    "   Select  Nvl(A.ԤԼʱ��,nvl(����ʱ��,sysdate)) - " & mlng��ǰ����ʱ�� & "/24/60 as �Һ�ʱ��  " & _
    "   From ���˹Һż�¼ A " & _
    "   Where No=[1] And Nvl(A.ԤԼʱ��,nvl(����ʱ��,sysdate))- " & mlng��ǰ����ʱ�� & "*1/24/60>sysdate"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO)
    If rsTemp.EOF Then Check������� = True: Exit Function
    strMsg = "�ò�����Ҫ��" & Format(rsTemp!�Һ�ʱ��, "yyyy-mm-dd HH:MM:SS") & "����������" & str����
    If mlng������� = 2 Then
        Check������� = (MsgBox(strMsg & ",��ȷ��Ҫ����" & str���� & "��", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes)
    Else
        MsgBox strMsg & ",������" & str����, vbInformation, gstrSysName
    End If
    Exit Function
ErrHanl:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function LoadNotify() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim strTmp As String
    Dim i As Long
    Dim blnDo As Boolean
    Dim strTag As String
    
    mstrPreNotify = ""
    rptNotify.Records.DeleteAll
    If Mid(mstrNotifyAdvice, mΣ��ֵ, 1) = "1" Then strTmp = strTmp & ",ZLHIS_LIS_003,ZLHIS_PACS_005"
    If Mid(mstrNotifyAdvice, mҽ������, 1) = "1" Then strTmp = strTmp & ",ZLHIS_OPER_001,ZLHIS_CIS_015,ZLHIS_CIS_005"
    If Mid(mstrNotifyAdvice, m�������, 1) = "1" Then strTmp = strTmp & ",ZLHIS_RECIPEAUDIT_001"
    If Mid(mstrNotifyAdvice, m��Ⱦ��, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CIS_032,ZLHIS_CIS_033"
    If Mid(mstrNotifyAdvice, m��Ѫ���, 1) = "1" And gblnѪ��ϵͳ Then strTmp = strTmp & ",ZLHIS_BLOOD_001"   '����Ѫ�����̲��д���Ϣ�Ͳ���
    If Mid(mstrNotifyAdvice, m��Ѫ���, 1) = "1" And gblnѪ��ϵͳ Then strTmp = strTmp & ",ZLHIS_BLOOD_004"  '����Ѫ����д���Ϣ�Ͳ���
    If Mid(mstrNotifyAdvice, m��Ѫ��Ӧ, 1) = "1" And gblnѪ��ϵͳ Then strTmp = strTmp & ",ZLHIS_BLOOD_006"  '����Ѫ����д���Ϣ�Ͳ���

    strTmp = strTmp & ",ZLHIS_CIS_037" 'Ĭ�Ͽ���ҽ�������������
    
    strTmp = strTmp & ",ZLHIS_EMR_019,ZLHIS_EMR_026" 'Ĭ�Ͽ��� �²���δǩ������ʱ,����ҽ��վ���²���ǩ������ʱ����Ҫ�󶩵�,����ҽ��վ
    
    
    strTmp = Mid(strTmp, 2)
    strSql = "Select b.id,a.����id,a.NO,a.id as �Һ�ID,a.�����,a.����,a.ִ��ʱ�� as ����ʱ��,b.��Ϣ����,b.���ͱ���, b.ҵ���ʶ, b.���ȳ̶�, b.�Ǽ�ʱ��,a.����,b.������Դ,f.�Ƿ�������Ϣ" & _
        " From ҵ����Ϣ�嵥 B, ���˹Һż�¼ A, ҵ����Ϣ���� F" & _
        " Where b.����id=a.Id And a.ִ����||''=[1]  And b.�Ǽ�ʱ��>=Trunc(Sysdate-(" & (mintNotifyDay - 1) & "))" & _
        " And Nvl(b.�Ƿ�����,0)=0 And b.���ͱ��� = f.���� And (f.�Ƿ�������Ϣ = 1 Or instr(','||[2]||',',','||b.���ͱ���||',')>0) AND substr(b.���ѳ���,1,1)='1' " & _
        " Order By b.���ȳ̶� Desc, b.�Ǽ�ʱ�� Desc"

    Screen.MousePointer = 11

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Name, mstr����ҽ��, strTmp)

    For i = 1 To rsTmp.RecordCount
        Select Case rsTmp!���ͱ���
        Case "ZLHIS_CIS_032", "ZLHIS_CIS_033"
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!���ͱ��� & ":" & rsTmp!����ID & ":" & rsTmp!ҵ���ʶ & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!���ͱ��� & ":" & rsTmp!����ID & ":" & rsTmp!ҵ���ʶ
                blnDo = True
            End If
        Case "ZLHIS_LIS_003", "ZLHIS_PACS_005"
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!ҵ���ʶ & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!ҵ���ʶ
                blnDo = True
            End If
        Case "ZLHIS_BLOOD_006"
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!���ͱ��� & ":" & rsTmp!����ID & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!���ͱ��� & ":" & rsTmp!����ID
                blnDo = True
            End If
        Case "ZLHIS_CIS_037", "ZLHIS_EMR_019", "ZLHIS_EMR_026"
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!����ID & "," & rsTmp!�Һ�ID & "," & rsTmp!���ͱ��� & "," & rsTmp!ҵ���ʶ & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!����ID & "," & rsTmp!�Һ�ID & "," & rsTmp!���ͱ��� & "," & rsTmp!ҵ���ʶ
                blnDo = True
            End If
        Case Else
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!����ID & "," & rsTmp!�Һ�ID & "," & rsTmp!���ͱ��� & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!����ID & "," & rsTmp!�Һ�ID & "," & rsTmp!���ͱ���
                blnDo = True
            End If
        End Select
        
        If Val(rsTmp!�Ƿ�������Ϣ & "") = 1 Then
            blnDo = True
        End If
        
        If blnDo Then
            Call AddReportRow(rsTmp!����ID & "," & rsTmp!�Һ�ID, rsTmp!����ID, rsTmp!NO, Nvl(rsTmp!����), Nvl(rsTmp!�����), Format(rsTmp!����ʱ�� & "", "yyyy-MM-dd HH:mm"), _
                 Nvl(rsTmp!��Ϣ����), rsTmp!���ͱ��� & "", rsTmp!���ȳ̶� & "", Format(rsTmp!�Ǽ�ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), Nvl(rsTmp!ҵ���ʶ), rsTmp!������Դ & "", _
                 Nvl(rsTmp!����, 0), rsTmp!�Һ�ID, rsTmp!ID, Val(rsTmp!�Ƿ�������Ϣ & ""))
                        blnDo = False
        End If
        rsTmp.MoveNext
    Next
    rptNotify.Populate 'ȱʡ��ѡ���κ���
    rptNotify.TabStop = rptNotify.Rows.Count > 0
    Screen.MousePointer = 0
    LoadNotify = True
    If mbln��Ϣ���� Then
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
'���ܣ�����Ϣ�����б�������һ��
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim objItemIcon As ReportRecordItem
    Dim strNO As String
    Dim strҵ�� As String
    Dim str������Դ As String
    Dim int���ȼ� As Integer
    Dim int���� As Integer
    Dim Index As Integer
    
    On Error GoTo errH
     
    Set objRecord = Me.rptNotify.Records.Add()
    objRecord.Tag = arrInput(Index): Index = Index + 1         'Tagֵ ����ID,�Һ�ID
    Set objItem = objRecord.AddItem(""): objItem.Icon = 6
    Set objItemIcon = objItem
    
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '����id
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  'NO
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1 '����
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index))) '�����
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index))) '����ʱ��
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))     '״̬������
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    strNO = arrInput(Index)                            '��Ϣ���
    objRecord.AddItem strNO: Index = Index + 1
    
    int���ȼ� = Val(arrInput(Index))                     '���ȼ�
    objRecord.AddItem int���ȼ�: Index = Index + 1
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  '�Ǽ�����
    
    strҵ�� = arrInput(Index): Index = Index + 1              'ҵ���ʶ
    str������Դ = arrInput(Index): Index = Index + 1          '������Դ
    
    int���� = arrInput(Index): Index = Index + 1
    objRecord.AddItem strҵ��
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))     '�Һ�ID
    
    Index = Index + 1
    objRecord.AddItem CStr(arrInput(Index)) '��ϢID��ҵ����Ϣ�嵥.ID
    
    Index = Index + 1
    objRecord.AddItem CStr(arrInput(Index)) '�Ƿ�������Ϣ��ҵ����Ϣ�嵥.�Ƿ�������Ϣ
    
    If int���ȼ� > 1 Then
        For Index = 0 To rptNotify.Columns.Count - 1
            If int���ȼ� = 3 Then
                objRecord.Item(Index).ForeColor = &HC0&
            End If
            objRecord.Item(Index).Bold = True
        Next
    End If
    '���ղ����ú�ɫ��ʾ
    If int���� > 0 And int���ȼ� <> 3 Then
        For Index = 0 To rptNotify.Columns.Count - 1
            objRecord.Item(Index).ForeColor = &HC0&
        Next
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub rptNotify_KeyUp(KeyCode As Integer, Shift As Integer)
'���ܣ��Զ�����ҽ��У�ԡ�ȷ��ֹͣ��ִ�н���
    Dim objControl As CommandBarControl
    Dim lngIndex As Long, lng����ID As Long
    Dim lngҽ��ID As Long, lng�Һ�id As Long, lng��ϢID As Long, lng������Ϣ As Long
    Dim strҵ�� As String, blnOK As Boolean
    Dim str���� As String, str����� As String
    Dim strNO As String
    Dim str�Һŵ� As String
    Dim str��Ϣ���� As String
    Dim blnTmp As Boolean
    Dim strTmp As String
    
    On Error GoTo errH
    
    If KeyCode = vbKeyReturn Then
        If rptNotify.SelectedRows.Count > 0 Then
            With rptNotify.SelectedRows(0).Record
                strNO = .Item(C_��Ϣ).Value
                strҵ�� = .Item(C_ҵ��).Value
                str�Һŵ� = .Item(C_No).Value
                str��Ϣ���� = .Item(C_״̬).Value
                lng����ID = Val(.Item(C_����ID).Value)
                lng�Һ�id = Val(.Item(C_�Һ�Id).Value)
                lng��ϢID = Val(.Item(C_Id).Value)
                str���� = .Item(c_����).Value
                str����� = .Item(C_�����).Value
                lng������Ϣ = Val(.Item(C_������Ϣ).Value)
                lngIndex = .Index
            End With
    
            blnTmp = True
            
            If str�Һŵ� <> mstr�Һŵ� Then blnTmp = LocatePati(str�Һŵ�)
            If strҵ�� <> "" Then      '�ҵ����˺�
                lngҽ��ID = Val(strҵ��)
            End If
            
            
            If strNO = "ZLHIS_RECIPEAUDIT_001" Then
                If strҵ�� = "������ҩ��" Then
                    blnTmp = CheckZLPass(Me, lng����ID, lng�Һ�id)
                    If blnTmp Then
                        str��Ϣ���� = "�������ϸ�"
                    Else
                        str��Ϣ���� = ""
                    End If
                End If
                '�Ƚ���Ƭ�л���ҽ����Ƭ������Ҳ˵�
                Call LocatedCard("ҽ��")
                cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
                If str��Ϣ���� = "�������ϸ�" Then
                    '������Ϣ���ʹ���
                    Set objControl = cbsMain.FindControl(, conMenu_Edit_Send, True, True)
                    If Not objControl Is Nothing Then
                        If objControl.Enabled Then objControl.Execute
                    End If
                Else
                    'ҽ���༭����
                    Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
                    If Not objControl Is Nothing Then
                        If objControl.Enabled Then objControl.Execute
                    End If
                End If
            End If
            
            If strNO = "ZLHIS_CIS_032" Then
                Call mclsDis.ShowDisRegist(Me, 1, Val(strҵ��), lng����ID, 0, str�Һŵ�)
            End If
            
            If strNO = "ZLHIS_BLOOD_006" Then
                If gobjPublicBlood Is Nothing And gblnѪ��ϵͳ Then InitObjBlood
                blnOK = gobjPublicBlood.zlIsBloodMessageDone(2, lng����ID, lng�Һ�id, 1, mlng����ID)
                If blnOK Then
                    Call rptNotify.Records.RemoveAt(lngIndex)
                Else
                    If FuncTraReaction(Val(strҵ��), mlngModul, False, IIf(InStr(1, strҵ��, ":") > 0, Val(Split(strҵ��, ":")(1)), 0)) Then
                        If gobjPublicBlood.zlIsBloodMessageDone(2, lng����ID, lng�Һ�id, 1, mlng����ID) Then
                            Call rptNotify.Records.RemoveAt(lngIndex)
                        End If
                    End If
                End If
            End If
            
            If strNO = "ZLHIS_CIS_033" Then
            '��Ⱦ�����淴�޸���Ϣ�Ķ�
                blnOK = ReadMsgCIS033(lng����ID, lng�Һ�id, strҵ��, lng��ϢID)
                If blnOK Then Call rptNotify.Records.RemoveAt(lngIndex)
            End If
            
            If strNO = "ZLHIS_CIS_037" Then 'ҽ��������¶�λ��ҽ��ҳǩ
                Call LocateMsgPati(lng����ID, lng�Һ�id, Val(strҵ��))
            End If
            
            '�²������¶�λ���²���ҳǩ
            If strNO = "ZLHIS_EMR_019" Or strNO = "ZLHIS_EMR_026" And InStr("," & strҵ�� & ",", "|") > 0 Then
                If GetInsidePrivs(p�°����ﲡ��, True) <> "" And Not mclsEMR Is Nothing Then
                    If mlng����ID = lng����ID And GetTabTag = "�²���" Then
                        Call mclsEMR.zlRefresh(mPatiInfo.����ID, mPatiInfo.�Һ�ID, mlng����ID, mPatiInfo.����, 1)
                    Else
                        '��λ������
                        If mlng����ID <> lng����ID Then
                            Call ExecuteFindPati(False, , , lng����ID)
                        End If
                        
                        '�ҵ����˺��л�ҳǩ
                        If mlng����ID = lng����ID Then
                            '��λ���²�����Ϣҳ
                            Call LocatedCard("�²���")
                            cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
                        End If
                    End If
                    
                    If mlng����ID = lng����ID And GetTabTag = "�²���" Then
                        strTmp = Trim(strҵ��)
                        If strTmp <> "" Then Call mclsEMR.EditDoc(strTmp)
                    End If
                End If
            End If
            
            
            If strNO <> "ZLHIS_CIS_033" And strNO <> "ZLHIS_BLOOD_006" Then
                blnOK = ReadMsg(lng����ID, lng�Һ�id, strNO, strҵ��, lng��ϢID, str�Һŵ�, lng������Ϣ)
                If blnOK Then Call rptNotify.Records.RemoveAt(lngIndex)
            End If
            Call rptNotify.Populate
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub rptNotify_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptNotify_KeyUp(vbKeyReturn, 0)
End Sub

Private Sub rptNotify_SelectionChanged()
    Dim str�Һŵ� As String
    
    If rptNotify.SelectedRows.Count = 0 Then Exit Sub  '���������
    
    str�Һŵ� = rptNotify.SelectedRows(0).Record.Item(C_No).Value
 
    If str�Һŵ� <> mstr�Һŵ� Then Call LocatePati(str�Һŵ�)
    
End Sub

Private Function ReadMsg(ByVal lng����ID As Long, ByVal lng�Һ�id As Long, ByVal strNO As String, ByVal strҵ�� As String, ByVal lng��ϢID As Long, ByVal str�Һŵ� As String, ByVal lng������Ϣ As Long) As Boolean
'���ܣ��Ķ���Ϣ
'˵������Ϣ�Ķ���ʽĿǰ��3�֣�����Ϣ�������Ķ�����ϢID�Ķ�����ҵ���ʶ�Ķ�
    Dim strSql As String
    Dim lng����ID As Long
    Dim strҽ��ID As String
    Dim blnDo As Boolean
    Dim lngΣ��ֵID As Long  '���δ����Σ��ֵ��¼ID
    Dim strSQLReadMsg As String
    Dim blnHisΣ��ֵ As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim objControl As CommandBarControl
    
    If mlng�������ID = 0 Then
        lng����ID = UserInfo.����ID
    Else
        lng����ID = mlng�������ID
    End If
    blnDo = True
    
    On Error GoTo errH
    
    strSql = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng�Һ�id & ",'" & strNO & "',1,'" & UserInfo.���� & "'," & lng����ID
    Select Case strNO
    Case "ZLHIS_LIS_003", "ZLHIS_PACS_005", "ZLHIS_CIS_037", "ZLHIS_EMR_019", "ZLHIS_EMR_026"
        strSql = strSql & ",null,null,'" & strҵ�� & "')"
    Case "ZLHIS_CIS_032", "ZLHIS_CIS_033"
        strSql = strSql & ",null," & lng��ϢID & ")"
    Case Else
        If lng������Ϣ = 1 Then
            strSql = strSql & ",null," & lng��ϢID & ")"
        Else
            strSql = strSql & ")"
        End If
    End Select
    
    strSQLReadMsg = strSql
    
    If strNO = "ZLHIS_LIS_003" Or strNO = "ZLHIS_PACS_005" Then
        If mblnΣ��ֵ Then
            'Σ��ֵ��Ϣ��ش���
            mblnΣ��ֵshow = True
            Call gobjKernel.ShowDealCritical(Me, lng����ID, 0, str�Һŵ�, lngΣ��ֵID)
            mblnΣ��ֵshow = False
            If lngΣ��ֵID <> 0 Then
                strSql = "select a.�걾id,a.�������,a.ȷ���� from ����Σ��ֵ��¼ a where a.id=[1] and a.ȷ���� is not null"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngΣ��ֵID)
                If Not rsTmp.EOF Then
                    '����Ϣ����Ϊ����
                    Call zlDatabase.ExecuteProcedure(strSQLReadMsg, Me.Caption)
                    
                    '�����LISΣ��ֵ����LIS�ӿ�
                    If strNO = "ZLHIS_LIS_003" Then
                        Call InitObjLis(p����ҽ��վ)
                        If Not gobjLIS Is Nothing Then
                            Call gobjLIS.WriteNotifyToLis(Val(rsTmp!�걾ID & ""), rsTmp!ȷ���� & "", rsTmp!������� & "")
                        End If
                    End If
                End If
            End If
            Call SetCriticalAdvice(lngΣ��ֵID)
            blnHisΣ��ֵ = True
        End If
    End If
    
    If Not blnHisΣ��ֵ Then
        If strNO = "ZLHIS_LIS_003" Then
            If strҵ�� <> "" Then
                strҽ��ID = strҵ��
                Call InitObjLis(p����ҽ��վ)
                If Not gobjLIS Is Nothing Then
                    blnDo = gobjLIS.GetReadNotify(Me, strҽ��ID, UserInfo.����)
                End If
            End If
        End If
        
        If strNO = "ZLHIS_BLOOD_004" Then
            '��Ѫ�����Ϣ���Ķ�״̬������Ѫ�ⲿ���ڲ����ٴ�����ִ���Ķ���Ϣ����
            strSql = "select 1 from ����ҽ����¼ a where a.�Һŵ�=[1] and a.ҽ��״̬=1 and a.�������='K' and a.��鷽��='1' and a.���״̬=1 and rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str�Һŵ�)
            If Not rsTmp.EOF Then
                '��������ݣ��򵯳�ҽ���޸Ľ��棬�������в�ִ����Ϣ�Ķ�SQL���
                '�Ƚ���Ƭ�л���ҽ����Ƭ������Ҳ˵�
                Call LocatedCard("ҽ��")
                cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
                'ҽ���༭����
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
'���ܣ�ͨ���Һŵ���λ����ǰ���Լ����б�
    Dim lngCol As Long
    Dim objRow As ReportRow
    Dim objRpt As ReportControl
    Dim blnEnabled  As Boolean
    
    Dim i As Long
    For i = 0 To 4
        Set objRpt = rptPati(i)
        If objRpt.Visible Then
            lngCol = IIf(i = PATI_RPT���� Or i = PATI_RPTԤԼ, COL_HZ_NO, IIf(i = PATI_RPT����, COL_YZ_NO, COL_JZ_NO))
            For Each objRow In objRpt.Rows
                If objRow.GroupRow Then objRow.Expanded = True
                If Not objRow.GroupRow Then
                    If objRow.Record(lngCol).Value = strTag Then
                        blnEnabled = timRefresh.Enabled
                        timRefresh.Enabled = False '������������ˢ����������
                        Set objRpt.FocusedRow = objRow 'ѡ��,��ʾ,[����Change�¼�]
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
    lblEdit(txtInfoժҪ).Left = 10
    lblEdit(txtInfoժҪ).Top = 10
    txtInfo(txtInfoժҪ).Left = lblEdit(txtInfoժҪ).Left + lblEdit(txtInfoժҪ).Width + 30
    txtInfo(txtInfoժҪ).Width = picMore.Width - txtInfo(txtInfoժҪ).Left
    
    lblEdit(txtInfo������Ϣ).Top = lblEdit(txtInfoժҪ).Top + lblEdit(txtInfoժҪ).Height + 60
    lblEdit(txtInfo������Ϣ).Left = lblEdit(txtInfoժҪ).Left
    
    txtInfo(txtInfo������Ϣ).Top = lblEdit(txtInfo������Ϣ).Top
    txtInfo(txtInfo������Ϣ).Left = lblEdit(txtInfo������Ϣ).Left + lblEdit(txtInfo������Ϣ).Width + 30
    txtInfo(txtInfo������Ϣ).Width = picMore.Width - txtInfo(txtInfo������Ϣ).Left
    
    UCPatiVitalSigns.Top = txtInfo(txtInfoժҪ).Top + txtInfo(txtInfoժҪ).Height + txtInfo(txtInfo������Ϣ).Height + 60
    UCPatiVitalSigns.Left = 10
End Sub

Private Sub picBasisNew_Resize()
    On Error Resume Next
    '�˴����Թ̶��߶�
        
    lblRec.FontName = "����"
    lblRec.FontSize = IIf(mbytSize = 0, 14, 18)
    
    If Err.Number <> 0 Then Err.Clear

    picPatient.Left = 10
    picPatient.Top = 10
    
    picPatient.Height = picBasisNew.Height - picPatient.Top - 60
    picPatient.Width = picPatient.Height
    imgPatient.Height = picPatient.Height
    imgPatient.Width = picPatient.Width
    
    lblLink(lblLink�ļ�).Left = picPatient.Left + picPatient.Width + 80
    lblLink(lblLink�ļ�).Top = picPatient.Top
    
    lblLink(lblLink�ɼ�).Left = lblLink(lblLink�ļ�).Left
    
    lblLink(lblLink�ɼ�).Top = picBasisNew.Height / 2 - 120
    
    lblLink(lblLink���).Left = lblLink(lblLink�ļ�).Left
    lblLink(lblLink���).Top = picPatient.Height + picPatient.Top - lblLink(lblLink���).Height
    
    lblLink(lblLink�޸�).Left = lblLink(lblLink���).Left + lblLink(lblLink���).Width + 180
    lblLink(lblLink�޸�).Top = lblLink(lblLink�ɼ�).Top
        
    txtInfo(txtInfo����).Top = 100
    txtInfo(txtInfo����).Left = lblLink(lblLink���).Left + lblLink(lblLink���).Width + 180
    txtInfo(txtInfo����).FontSize = IIf(mbytSize = 0, 12, 15)
    txtInfo(txtInfo����).Width = IIf(mbytSize = 0, 1400, 1800)
    
    txtInfo(txtInfo�Ա�).Top = txtInfo(txtInfo����).Top + txtInfo(txtInfo����).Height - txtInfo(txtInfo�Ա�).Height + 160
    txtInfo(txtInfo�Ա�).Left = txtInfo(txtInfo����).Left + txtInfo(txtInfo����).Width + 50
    
    txtInfo(txtInfo����).Top = txtInfo(txtInfo�Ա�).Top + txtInfo(txtInfo�Ա�).Height - txtInfo(txtInfo����).Height
    txtInfo(txtInfo����).Left = txtInfo(txtInfo�Ա�).Left + txtInfo(txtInfo�Ա�).Width + 100
    
    txtInfo(txtInfo������).Top = txtInfo(txtInfo����).Top
    txtInfo(txtInfo������).Width = IIf(mbytSize = 0, 1800, 2500)
    txtInfo(txtInfo������).Height = txtPhone.Height
    lblPhysical.Caption = "���������:"
    Call zlControl.SetPubCtrlPos(False, -1, txtInfo(txtInfo����), 250, lblEdit(txtInfo��������), 30, txtInfo(txtInfo��������), 250, _
        lblEdit(txtInfo���ѷ�ʽ), 30, fraPayType, 250, lblPhone, 30, txtPhone, 250, lblPhysical, 30, txtInfo(txtInfo������))
    If txtInfo(txtInfo������).Text = "" Then lblPhysical.Caption = ""
    txtInfo(txtInfo������).Visible = Not (lblPhysical.Caption = "")
    
    fraPayType.Top = lblEdit(txtInfo���ѷ�ʽ).Top - 30
    fraPayType.Width = cboPayType.Width
    fraPayType.Height = cboPayType.Height - 60
    
    linPayType.x1 = fraPayType.Left - 20
    linPayType.y1 = fraPayType.Top + fraPayType.Height
    linPayType.x2 = linPayType.x1 + fraPayType.Width
    linPayType.y2 = linPayType.y1
    
    LinPhone.x1 = txtPhone.Left - 20
    LinPhone.y1 = txtPhone.Top + txtPhone.Height
    LinPhone.x2 = LinPhone.x1 + txtPhone.Width
    LinPhone.y2 = LinPhone.y1
    
    
    txtInfo(txtInfo���￨��).Width = 1300
    Call zlControl.SetPubCtrlPos(False, -1, lblEdit(txtInfo����), 30, txtInfo(txtInfo����), 150, lblEdit(txtInfo���￨��), 30, txtInfo(txtInfo���￨��), 150, lblEdit(txtInfoҽ������), 30, txtInfo(txtInfoҽ������), 150, lblEdit(txtInfo�ѱ�), 30, fraBillType)
    fraBillType.Top = lblEdit(txtInfo�ѱ�).Top - 30
    
    fraBillType.Width = cboBillType.Width
    fraBillType.Height = cboBillType.Height - 60
    
    linBillType.x1 = fraBillType.Left - 20
    linBillType.y1 = fraBillType.Top + fraBillType.Height
    linBillType.x2 = linBillType.x1 + fraBillType.Width
    linBillType.y2 = linBillType.y1
    
    lblMore.Top = lblEdit(txtInfoҽ������).Top
    lblMore.Left = picBasisNew.Width - lblMore.Width - 40

    lblRec.Top = 200
    lblRec.Left = picBasisNew.Width - 40 - lblRec.Width - 20

End Sub

Private Sub lblLink_Click(Index As Integer)
    Dim strPictureFile As String
    Dim objControl As CommandBarControl
    
    On Error GoTo errH
    
    If lblLink(Index).ForeColor <> &HC00000 Then Exit Sub
    
    Select Case Index
    Case lblLink�ļ�
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
        Call SetPatPicture(mPatiInfo.����ID, False)
    Case lblLink���
        If picPatient.Tag <> "" Then
            If SetPatPicture(mPatiInfo.����ID, True) Then
                imgPatient.Picture = imgDefual.Picture
                picPatient.Tag = ""
            End If
        End If
    Case lblLink�ɼ�, lblLink�޸�
        If mobjPatient Is Nothing Then
            On Error Resume Next
            Set mobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
            Err.Clear: On Error GoTo 0
        End If
        If mobjPatient Is Nothing Then
            MsgBox "����������Ϣ����������zlPublicPatient.clsPublicPatient��ʧ�ܣ�", vbInformation, Me.Caption
            Exit Sub
        End If
        Call mobjPatient.zlInitCommon(gcnOracle, glngSys, UserInfo.�û���)
        
        If lblLink�ɼ� = Index Then
            If mobjPatient.PatiImageGatherer(Me, strPictureFile) = False Then Exit Sub
            Set imgPatient.Picture = LoadPicture(strPictureFile)
            picPatient.Tag = strPictureFile
            Call SetPatPicture(mPatiInfo.����ID, False)
        Else
            If mobjPatient.ModiPatiBaseInfo(Me, "����ҽ������վ", mPatiInfo.����ID, mPatiInfo.�Һ�ID, 1, False) Then
                '�޸ĳɹ���ˢ�£���������ͳһˢ��
                Set objControl = cbsMain.FindControl(, conMenu_View_Refresh, , True)
                If Not objControl Is Nothing Then
                    If objControl.Enabled Then objControl.Execute
                End If
            End If
        End If
    End Select
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Err.Clear
End Sub


Private Function SetPatPicture(ByVal lng����ID As Long, ByVal blnDel As Boolean) As Boolean
'����:���ò�����Ƭ
'���:lng����ID - ����ID��blnDel true ɾ����Ƭ��false ������Ƭ
    Dim strFile As String, strSql As String
    
    Dim strPhotoNew As String
    
    On Error GoTo errH

    If blnDel Then
        If MsgBox("����" & txtInfo(txtInfo����).Text & "����Ƭ����ɾ�����Ƿ������", vbDefaultButton2 + vbYesNo + vbQuestion, Me.Caption) = vbNo Then
            Exit Function
        End If
        
        Call DeletePatPicture(lng����ID)
    Else
        'ͼƬû�б�����������²���ͼƬ
        If picPatient.Tag <> "" Then
            strFile = picPatient.Tag
            Call ReLoadPicture(strFile, strPhotoNew)
            If SavePatPicture(lng����ID, IIf(strPhotoNew = "", strFile, strPhotoNew)) = False Then
                MsgBox "������Ƭ����,��ȷ���ļ��Ƿ�ɾ��!", vbInformation, gstrSysName
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
'���ܣ�������Ϣ��Ƭ�����˲�����
    Dim blnDoc As Boolean
    
    If InStr(";" & GetPrivFunc(glngSys, p���ﲡ������) & ";", ";������д;") > 0 Then
        blnDoc = mlng����ID <> 0 And mlng����ID = mPatiInfo.����ID And mstr�Һŵ� = mPatiInfo.�Һŵ� And _
                 (lngFileID = 0 And lngEPRFileID <> 0 Or lngFileID <> 0) And (mintActive = pt���� Or mintActive = pt����)
        If blnDoc And lngFileID <> 0 And strIn = "0" Then  'û���޸����˲�����Ȩ��
            blnDoc = strDoctor = UserInfo.����
        End If
        
        If blnDoc Then
            'If mobjEPRDoc Is Nothing Then
                Set mobjEPRDoc = New zlRichEPR.cEPRDocument
            'End If
            If lngFileID = 0 And lngEPRFileID <> 0 Then '���û���½����½�
                Call mobjEPRDoc.InitEPRDoc(0, 2, lngEPRFileID, 1, mPatiInfo.����ID, mPatiInfo.�Һ�ID, , mPatiInfo.����ID, , False)
            Else
                Call mobjEPRDoc.InitEPRDoc(1, 2, lngFileID, 1, mPatiInfo.����ID, mPatiInfo.�Һ�ID, , mPatiInfo.����ID, , False)
            End If
            Call mobjEPRDoc.ShowEPREditor(Me)
        Else
            MsgBox "��ǰ���������޸ġ�", vbInformation, Me.Caption
        End If
    Else
        MsgBox "��û�в�����д��Ȩ�ޡ�", vbInformation, Me.Caption
    End If
End Sub

Private Sub mobjPati_EPRRefresh()
    With mPatiInfo
        Call mclsEPRs.zlRefresh(mlng����ID, mlng�Һ�ID, mlng����ID, mlng����ID = .����ID And (.���� = pt���� Or .���� = pt����) And mlng����ID <> 0, .����ת��, True)
    End With
End Sub

Private Sub mobjPati_UpdatePatiInfo(ByVal strBirthday As String, ByVal strAge As String, ByVal strSex As String, ByVal strTag As String)
'���ܣ����²�����Ϣ
    If strBirthday <> "" Then
        txtInfo(txtInfo��������).Text = Format(strBirthday, "yyyy-MM-dd")
    End If
    If strAge <> "" Then
        txtInfo(txtInfo����).Text = strAge
    End If
    If strSex <> "" Then
        txtInfo(txtInfo�Ա�).Text = strSex
    End If
    If (strAge <> "" Or strSex <> "") And (mintRPTIndex = PATI_RPT���� Or mintRPTIndex = PATI_RPT����) Then
        If mPr <> -1 And rptPati(mintRPTIndex).Visible Then
            With rptPati(mintRPTIndex).Rows(mPr)
                .Record(COL_JZ_�Ա�).Value = strSex
                .Record(COL_JZ_����).Value = strAge
            End With
            rptPati(mintRPTIndex).Populate
        End If
    End If
    If strTag <> "" Then
        txtInfo(txtInfoժҪ).Text = IIf("NULL" = strTag, "", strTag)
    End If
End Sub

Private Sub mobjPati_UpdateDiagInfo(ByVal str����ID As String, ByVal str���ID As String, ByVal strTag As String)
'���ܣ���Ⱦ�����
    Dim blnNo As Boolean
    Dim rsTmp  As ADODB.Recordset
    Dim blnNotView As Boolean
    
    If InStr(";" & GetPrivFunc(glngSys, p����������д) & ";", ";������д;") > 0 Then
        Set rsTmp = mclsDisease.SatisfyEditDiseaseDoc(mlng����ID, mlng�Һ�ID, mlng����ID, str����ID, str���ID)
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then
                If Not mclsDis.ShowDiseaseStation(Me, mlng����ID, mlng�Һ�ID, 1, mlng����ID, str����ID, str���ID, blnNotView) Then
                    Call mclsDisease.EditDiseaseReport(Me, rsTmp, mlng����ID, mlng�Һ�ID, 1, mlng����ID, blnNo)
                    If blnNo Then
                        Call mclsDis.EditNotFillReason(Me, mlng����ID, mlng�Һ�ID, 1)
                    End If
                ElseIf blnNotView Then
                    Call mclsDisease.EditDiseaseReport(Me, rsTmp, mlng����ID, mlng�Һ�ID, 1, mlng����ID, blnNo)
                    If blnNo Then
                        Call mclsDis.EditNotFillReason(Me, mlng����ID, mlng�Һ�ID, 1)
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub LocatedCard(ByVal strTag As String)
'���ܣ���λ��ָ����ҳǩ��Ƭ���ڲ�ҳǩ
    Dim i As Long
    '1.�ȶ�λ�����ξ���
    If tbcRegist.Selected.Caption <> "���ξ���" Then
        tbcRegist.Item(mbyt���ξ���).Selected = True
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
'���ܣ����������б�ֵ
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    
    strSql = "Select ����, ���� From ҽ�Ƹ��ʽ Order By ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    With cboPayType
        .Clear
        For i = 1 To rsTmp.RecordCount
            .AddItem rsTmp!���� & "-" & rsTmp!����
            .ItemData(.NewIndex) = Val(rsTmp!���� & "")
            rsTmp.MoveNext
        Next
    End With
    
    strSql = "Select ����, ���� From �ѱ� Order By ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    With cboBillType
        .Clear
        For i = 1 To rsTmp.RecordCount
            .AddItem rsTmp!���� & ""
            .ItemData(.NewIndex) = Val(rsTmp!���� & "")
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

Private Function Tip�����Զ����() As Boolean
'���ܣ�����ǰ������Ϊ��ɾ������
    Dim objMsg As New zl9ComLib.clsAirBubble
    Dim varPatis As Variant
    Dim lng�Һ�id As Long
    Dim i As Long
    Dim strSql As String
    Dim rsPati As ADODB.Recordset
    Dim strInfo As String
    Dim blnDo As Boolean
    Dim str��������1 As String
    Dim str��������2 As String
    Dim intType As Integer
    If mstr�Һ�IDs = "" Then Exit Function
    On Error GoTo errH
    strInfo = mstr�Һ�IDs
    varPatis = Split(strInfo, ",")
    For i = 0 To UBound(varPatis)
        lng�Һ�id = Val(varPatis(i))
        If lng�Һ�id <> 0 And lng�Һ�id <> mPatiInfo.�Һ�ID Then
            blnDo = False
            intType = 0
            '�����жϵ�ǰ�����ǲ����Ѿ���ɾ���ͻ�����
            strSql = "select ID,NO,����,����ID,ִ��״̬,ת��״̬,decode(��¼��־,2,1,3,1,0) as ���� from ���˹Һż�¼ where ��¼����=1 And ��¼״̬=1 and id=[1]"
            Set rsPati = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng�Һ�id)
            If rsPati.RecordCount = 1 Then
                If 2 = Val(rsPati!ִ��״̬ & "") Then
                    blnDo = CanAutoFinish(Val(rsPati!����ID & ""), rsPati!NO & "", lng�Һ�id, intType)
                End If
            End If
            
            If blnDo Then
                If intType <> 2 Then
                    str��������1 = str��������1 & "��" & rsPati!����
                    '��ɾ���֮ǰ����Ѿ�ת�����ȡ��ת��
                    If Not IsNull(rsPati!ת��״̬) Then
                        If Val(rsPati!ת��״̬ & "") = 0 Then
                            If Update���˹Һ�ת��(Val(rsPati!����ID), rsPati!NO, 4, 0, "", "", p����ҽ��վ) = False Then Exit Function
                        End If
                    End If
                    Call ExecuteFinishInSide(rsPati!NO & "", Val(rsPati!����ID & ""), lng�Һ�id)
                ElseIf intType <> 0 Then
                    If Val(rsPati!���� & "") = 0 Then
                        If Update���˻���(lng�Һ�id, 0, "", "", 1, "", p����ҽ��վ) = False Then Exit Function
                        'ֻ��ʾδ��Ϊ����ģ��Ѿ�����˵ľͲ���ʾ��
                        str��������2 = str��������2 & "��" & rsPati!����
                    End If
                End If
            End If
        End If
    Next
    strInfo = ""
    If str��������1 <> "" Then
        str��������1 = Mid(str��������1, 2) & " �����Զ���ɾ�����������б��в鿴��"
        Call LoadPatients����
    End If
    
    If str��������2 <> "" Then
        str��������2 = Mid(str��������2, 2) & " �����Զ���ǻ�����ڻ����б��в鿴��"
        Call LoadPatients����
    End If
    strInfo = IIf("" = str��������1, "", str��������1 & vbCrLf) & IIf("" = str��������2, "", str��������2)
    If strInfo <> "" Then
        Call objMsg.OpenTransparentAirBubble(Me, strInfo, 2, 2, 15, &HFF8080, &HFFFFFF, , 3, , , ��)
        Set mobjMsg = objMsg
    End If
    mstr�Һ�IDs = ""
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CanAutoFinish(ByVal lng����ID As Long, ByVal strNO As String, ByVal lng�Һ�id As Long, ByRef intType As Integer) As Boolean
'���ܣ���ǰ�����Ƿ�����Զ�������һ�����ڣ���ɾ�����߻���
'������intType 1-��ɾ��2������
    Dim j As Long
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim blnSigned As Boolean, blnOK���� As Boolean
    Dim objEmr As Object
    Dim strҽ��IDs As String
    Dim lngTmp As Long, lngTmp1 As Long
    
    On Error GoTo errH
    intType = 1
    '1.�������
    strSql = "select 1 from ���Ӳ�����¼ where ����ID=[1] and ��ҳID=[2] and ǩ������<>1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, lng�Һ�id)
    If rsTmp.EOF Then
        blnSigned = True
        If GetInsidePrivs(p�°����ﲡ��, True) <> "" Then
            On Error Resume Next
            Set objEmr = CreateObject("zl9EmrInterface.ClsEmrInterface")
            If Not objEmr Is Nothing Then
                Call objEmr.CheckOutEPRIsAllSign(lng�Һ�id, blnSigned)
            End If
            Err.Clear: On Error GoTo 0
            On Error GoTo errH
        End If
        If blnSigned Then
            blnOK���� = True
        End If
    Else
        Exit Function
    End If
    
    '2.ҽ�����
    If blnOK���� Then
        strSql = "select a.id,a.���id,a.���,a.ҽ��״̬,a.�������," & _
            " NVL(a.ִ�б��,0) as ִ�б��, Nvl(a.ִ������,0) as ִ������,Nvl(a.Ƥ�Խ��,'��') as Ƥ�Խ�� from ����ҽ����¼ a where a.ҽ��״̬<>4 and a.�Һŵ�=[1] and a.����ID+0=[2]"
        Set rsAdvice = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, lng����ID)
        If rsAdvice.RecordCount = 0 Then
            '����Чҽ��
            Exit Function
        End If
        
        rsAdvice.Filter = "ҽ��״̬=1 And ִ�б��<>-1 And ִ������<>0 And Ƥ�Խ��<>'����'"
        If rsAdvice.RecordCount <> 0 Then
            'δ���͵�ҽ��
            Exit Function
        End If
        
        '�Ѿ����͵ļ�����ҽ��
        rsAdvice.Filter = "(ҽ��״̬=8 and �������='D') or (ҽ��״̬=8 and �������='C')"
        strҽ��IDs = ""
        For j = 1 To rsAdvice.RecordCount
            lngTmp = Val(rsAdvice!ID & "")
            lngTmp1 = Val(rsAdvice!���ID & "")
            
            If InStr("," & strҽ��IDs & ",", "," & lngTmp & ",") = 0 Then
                strҽ��IDs = strҽ��IDs & "," & lngTmp
            End If
            
            If InStr("," & strҽ��IDs & ",", "," & lngTmp1 & ",") = 0 Then
                strҽ��IDs = strҽ��IDs & "," & lngTmp1
            End If
            rsAdvice.MoveNext
        Next
        
        If strҽ��IDs <> "" Then
            strҽ��IDs = Mid(strҽ��IDs, 2)
            strSql = "select 1 from ����ҽ������ a where a.ҽ��id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) and a.ִ��״̬<>1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strҽ��IDs)
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

Private Function ReadMsgCIS033(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal str��ʶ As String, ByVal lng��ϢID As Long) As Boolean
'���ܣ���Ⱦ�����淴�޸���Ϣ�Ķ�
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim lng�ļ�ID As Long
    Dim lng����ID As Long
    Dim objControl As CommandBarControl
    
    On Error GoTo errH
    'conMenu_Edit_Modify 3003 �޸İ�ť��
    lng�ļ�ID = Val(Split(str��ʶ, ",")(0))
    
    strSql = "Select 1 From �����걨��¼ where �ļ�ID=[1] and ����״̬=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng�ļ�ID, 4)
    If rsTmp.RecordCount = 0 Then
    '����Ϣ���Ϊ�Ѷ�
        If mlng�������ID = 0 Then
            lng����ID = UserInfo.����ID
        Else
            lng����ID = mlng�������ID
        End If
        strSql = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng����ID & ",'ZLHIS_CIS_033',1,'" & UserInfo.���� & "'," & lng����ID & ",null," & lng��ϢID & ")"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        ReadMsgCIS033 = True
        Exit Function
    End If
    
    If "�л����񹲺͹���Ⱦ�����濨" = Sys.RowValue("���Ӳ�����¼", lng�ļ�ID, "��������") Then
        '�������޸ı���
        '�Ƚ���Ƭ�л���ҽ����Ƭ������Ҳ˵�
        Call LocatedCard("��������")
        cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
        If tbcSub.Selected.Tag = "��������" And tbcSub.Selected.Visible = True Then
            Set objControl = cbsMain.FindControl(, conMenu_Edit_Modify, True, True)
            If Not objControl Is Nothing Then
                If objControl.Enabled Then objControl.Execute
            End If
        End If
    Else
        '�������޸ı���
        Call mclsDis.ModifyDiseaseDoc(Me, lng�ļ�ID, mlng����ID, mlng�Һ�ID, 1, mlng����ID)
    End If
    
    strSql = "Select 1 From �����걨��¼ where �ļ�ID=[1] and ����״̬=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng�ļ�ID, 4)
    If rsTmp.RecordCount = 0 Then
    '����Ϣ���Ϊ�Ѷ�
        If mlng�������ID = 0 Then
            lng����ID = UserInfo.����ID
        Else
            lng����ID = mlng�������ID
        End If
        strSql = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng����ID & ",'ZLHIS_CIS_033',1,'" & UserInfo.���� & "'," & lng����ID & ",null," & lng��ϢID & ")"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
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

Private Sub mclsDis_PatiTransfer(ByVal lng����ID As Long, ByVal str�Һ�No As String)
'���ܣ���Ⱦ�����Խ��津���¼�ת�
    Call ExecuteTransferSend
End Sub

Private Sub mobjPati_SetEdit()
    picBasisNew.SetFocus
    tbcSub.SetFocus
End Sub

Private Sub ExecuteCritical()
'���ܣ�Σ��ֵ��ش���
    Dim lngΣ��ֵID As Long  '���δ����Σ��ֵ��¼ID
    
    mblnΣ��ֵshow = True
    Call gobjKernel.ShowDealCritical(Me, mlng����ID, 0, mstr�Һŵ�, lngΣ��ֵID)
    mblnΣ��ֵshow = False
    
    Call SetCriticalAdvice(lngΣ��ֵID)
End Sub

Private Sub mobjQueue_OnInitQueueList(objQueueList As Object, objCallList As Object, blnIsCustom As Boolean)
'���ܣ��Ŷӽ��б�ĳ�ʼ������

    Dim Column As ReportColumn
    Dim str�Ŷ��п� As String
    Dim str�����п� As String
    Dim strReg As String
    
    On Error GoTo errH
    
    Set mobjQueueList = objQueueList
    Set mobjCallList = objCallList
 
    strReg = "����ȫ��\�Զ����Ŷӽк�" & CStr(mlngModul)
    str�Ŷ��п� = GetSetting("ZLSOFT", strReg, "�Ŷ��п������", C_STR_QUEUEQUEUE)
    str�����п� = GetSetting("ZLSOFT", strReg, "�����п������", C_STR_QUEUECALL)
    If UBound(Split(str�Ŷ��п�, ",")) <> 18 Then
        str�Ŷ��п� = C_STR_QUEUEQUEUE
    End If
    If UBound(Split(str�����п�, ",")) <> 18 Then
        str�����п� = C_STR_QUEUECALL
    End If
    mlngQueueGroupType = zlDatabase.GetPara("�Ŷӷ�������", glngSys, p�Ŷӽк�����ģ��, "0")
    mstrShowColumnInf = zlDatabase.GetPara("������ʾ��", glngSys, p�Ŷӽк�����ģ��, "����,��������,�Ŷ�״̬")
    mstrShowColumnInf = Replace(mstrShowColumnInf, "��", ",")
    mstrShowColumnInf = "," & mstrShowColumnInf & ","
    mstrShowCalledColumnInf = zlDatabase.GetPara("����������ʾ��", glngSys, p�Ŷӽк�����ģ��, "����,��������")
    mstrShowCalledColumnInf = Replace(mstrShowCalledColumnInf, "��", ",")
    mstrShowCalledColumnInf = "," & mstrShowCalledColumnInf & ","
    mlngOrderStyle = zlDatabase.GetPara("ʹ������ԭʼ˳������", glngSys, p�Ŷӽк�����ģ��, "0")
    mlng���ﲡ������ = zlDatabase.GetPara("���ﲡ���Ƿ�����", glngSys, p�Ŷӽк�����ģ��, "1")
    mlngQueueGroupType = zlDatabase.GetPara("�Ŷӷ�������", glngSys, p�Ŷӽк�����ģ��, "0")
    mlng���ﲡ������ = zlDatabase.GetPara("���ﲡ���Ƿ�����", glngSys, p�Ŷӽк�����ģ��, "1")

    'ԭ��������
    With objCallList.Columns
        objCallList.AllowColumnRemove = False
        objCallList.ShowItemsInGroups = False
        objCallList.SkipGroupsFocus = True
        objCallList.MultipleSelection = False
        objCallList.AutoColumnSizing = False
        With objCallList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "���б����϶�����,�ɰ����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
        Set Column = .Add(mCol.��������, IIf(mlngQueueGroupType = 0, "", "����"), Val(Split(str�Ŷ��п�, ",")(0)), False)
        If mlngQueueGroupType = 0 Then
            Column.Groupable = True
        Else
            Column.Visible = False
        End If

        Set Column = .Add(mCol.ID, "ID", Val(Split(str�����п�, ",")(1)), False)
        Column.Visible = False

        Set Column = .Add(mCol.����ID, "����ID", Val(Split(str�����п�, ",")(2)), False)
        Column.Visible = False

        Set Column = .Add(mCol.�Ŷӱ��, "���", Val(Split(str�����п�, ",")(3)), False)
        Column.Visible = False

        Set Column = .Add(mCol.�ŶӺ���, "����", Val(Split(str�����п�, ",")(4)), True)
        Column.Visible = True

        Set Column = .Add(mCol.�Ŷ����, "�Ŷ����", Val(Split(str�����п�, ",")(5)), False)
        Column.Visible = False

        Set Column = .Add(mCol.��������, "��������", Val(Split(str�����п�, ",")(6)), True)
        Column.Visible = True

        Set Column = .Add(mCol.����, "����", Val(Split(str�����п�, ",")(7)), False)
        Column.Visible = False

        Set Column = .Add(mCol.�������, "�������", Val(Split(str�����п�, ",")(8)), True)
        Column.Visible = IIf(InStr(1, mstrShowCalledColumnInf, ",�������,") > 0, True, False)

        Set Column = .Add(mCol.���������, "���������", Val(Split(str�����п�, ",")(9)), False)
        Column.Visible = False

        Set Column = .Add(mCol.����ID, "����ID", Val(Split(str�����п�, ",")(10)), False)
        Column.Visible = False

        Set Column = .Add(mCol.����, IIf(mlngQueueGroupType = 2, "", "����"), Val(Split(str�����п�, ",")(11)), True)
        If mlngQueueGroupType = 2 Then
            Column.Groupable = True
            Column.Visible = False
        Else
            Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",����,") > 0, True, False)
        End If

        Set Column = .Add(mCol.ҽ������, IIf(mlngQueueGroupType = 1, "", "ҽ������"), Val(Split(str�����п�, ",")(12)), True)
        If mlngQueueGroupType = 1 Then
            Column.Groupable = True
            Column.Visible = False
        Else
            Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",ҽ������,") > 0, True, False)
        End If

        Set Column = .Add(mCol.�Ŷ�״̬, "�Ŷ�״̬", Val(Split(str�����п�, ",")(13)), False)
        Column.Visible = False

        Set Column = .Add(mCol.�Ŷ�ʱ��, "�Ŷ�ʱ��", Val(Split(str�����п�, ",")(14)), False)
        Column.Visible = False

        Set Column = .Add(mCol.����ҽ��, "������", Val(Split(str�����п�, ",")(15)), True)
        Column.Visible = IIf(InStr(1, mstrShowCalledColumnInf, ",����ҽ��,") > 0, True, False)

        Set Column = .Add(mCol.ҵ������, "ҵ������", Val(Split(str�����п�, ",")(16)), False)
        Column.Visible = False

        Set Column = .Add(mCol.ҵ��ID, "ҵ��ID", Val(Split(str�����п�, ",")(17)), False)
        Column.Visible = False

        Set Column = .Add(mCol.����ʱ��, "����ʱ��", Val(Split(str�����п�, ",")(18)), True)
        Column.Visible = IIf(InStr(1, mstrShowCalledColumnInf, ",����ʱ��,") > 0, True, False)

        Set Column = .Add(mCol.��������, "��������", 0, False)
        Column.Visible = False

        Set Column = .Add(mCol.ORD, "ORD", 0, False)
        Column.Visible = False

    End With

    With objCallList
        Set .Icons = ZLCommFun.GetPubIcons
        .GroupsOrder.DeleteAll
        If mlngQueueGroupType = 0 Then
            .GroupsOrder.Add .Columns(mCol.��������)
        ElseIf mlngQueueGroupType = 1 Then
            .GroupsOrder.Add .Columns(mCol.ҽ������)
        Else
            .GroupsOrder.Add .Columns(mCol.����)
        End If
        .GroupsOrder(0).SortAscending = True '����֮��,��������в���ʾ,�����е������ǲ����
        .SortOrder.DeleteAll
        If mlngOrderStyle = 1 Then
            .SortOrder.Add .Columns(mCol.ORD)
            .SortOrder(0).SortAscending = True
        Else

            .SortOrder.Add .Columns(mCol.�Ŷ�״̬)
            .SortOrder(0).SortAscending = False

            .SortOrder.Add .Columns(mCol.�Ŷ����)
            .SortOrder(1).SortAscending = True

            .SortOrder.Add .Columns(mCol.����ʱ��)
            .SortOrder(2).SortAscending = True

            .SortOrder.Add .Columns(mCol.�ŶӺ���)
            .SortOrder(3).SortAscending = True
        End If
    End With

    '��ʼ���ŶӶ�����ʾ�ֶ�
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
            .NoGroupByText = "���б����϶�����,�ɰ����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With

        Set Column = .Add(mCol.��������, IIf(mlngQueueGroupType = 0, "", "����"), Val(Split(str�Ŷ��п�, ",")(0)), False)

        If mlngQueueGroupType = 0 Then
            Column.Groupable = True
        Else
            Column.Visible = False
        End If
        
        Set Column = .Add(mCol.ID, "ID", Val(Split(str�Ŷ��п�, ",")(1)), False)
        Column.Visible = False

        Set Column = .Add(mCol.����ID, "����ID", Val(Split(str�Ŷ��п�, ",")(2)), False)
        Column.Visible = False

        Set Column = .Add(mCol.�Ŷӱ��, "���", Val(Split(str�Ŷ��п�, ",")(3)), False)
        Column.Visible = False

        Set Column = .Add(mCol.�ŶӺ���, "����", Val(Split(str�Ŷ��п�, ",")(4)), True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",����,") > 0, True, False)

        Set Column = .Add(mCol.�Ŷ����, "�Ŷ����", Val(Split(str�Ŷ��п�, ",")(5)), False)
        Column.Visible = False

        Set Column = .Add(mCol.��������, "��������", Val(Split(str�Ŷ��п�, ",")(6)), True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",��������,") > 0, True, False)

        Set Column = .Add(mCol.����, "����", Val(Split(str�Ŷ��п�, ",")(7)), True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, "����") > 0, True, False)

        Set Column = .Add(mCol.�������, "�������", Val(Split(str�Ŷ��п�, ",")(8)), True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",�������,") > 0, True, False)

        Set Column = .Add(mCol.���������, "���������", Val(Split(str�Ŷ��п�, ",")(9)), True)
        Column.Visible = False

        Set Column = .Add(mCol.����ID, "����ID", Val(Split(str�Ŷ��п�, ",")(10)), False)
        Column.Visible = False

        Set Column = .Add(mCol.����, IIf(mlngQueueGroupType = 2, "", "����"), Val(Split(str�Ŷ��п�, ",")(11)), True)
        If mlngQueueGroupType = 2 Then
            Column.Groupable = True
            Column.Visible = False
        Else
            Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",����,") > 0, True, False)
        End If

        Set Column = .Add(mCol.ҽ������, IIf(mlngQueueGroupType = 1, "", "ҽ������"), Val(Split(str�Ŷ��п�, ",")(12)), True)
        If mlngQueueGroupType = 1 Then
            Column.Groupable = True
            Column.Visible = False
        Else
            Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",ҽ������,") > 0, True, False)
        End If
        Set Column = .Add(mCol.�Ŷ�״̬, "�Ŷ�״̬", Val(Split(str�Ŷ��п�, ",")(13)), True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",�Ŷ�״̬,") > 0, True, False)

        Set Column = .Add(mCol.�Ŷ�ʱ��, "�Ŷ�ʱ��", Val(Split(str�Ŷ��п�, ",")(14)), True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",�Ŷ�ʱ��,") > 0, True, False)

        Set Column = .Add(mCol.����ҽ��, "������", Val(Split(str�Ŷ��п�, ",")(15)), False)
        Column.Visible = False

        Set Column = .Add(mCol.ҵ������, "ҵ������", Val(Split(str�Ŷ��п�, ",")(16)), False)
        Column.Visible = False

        Set Column = .Add(mCol.ҵ��ID, "ҵ��ID", Val(Split(str�Ŷ��п�, ",")(17)), False)
        Column.Visible = False

        Set Column = .Add(mCol.����ʱ��, "����ʱ��", Val(Split(str�Ŷ��п�, ",")(18)), False)
        Column.Visible = False

        Set Column = .Add(mCol.��������, "��������", 0, False)
        Column.Visible = False

        Set Column = .Add(mCol.ORD, "ORD", 0, False)
        Column.Visible = False
    End With

    With objQueueList
        Set .Icons = ZLCommFun.GetPubIcons

        .GroupsOrder.DeleteAll

        If mlngQueueGroupType = 0 Then
            .GroupsOrder.Add .Columns(mCol.��������)
        ElseIf mlngQueueGroupType = 1 Then
            .GroupsOrder.Add .Columns(mCol.ҽ������)
        Else
            .GroupsOrder.Add .Columns(mCol.����)
        End If

        .GroupsOrder(0).SortAscending = True '����֮��,��������в���ʾ,�����е������ǲ����

        '�������� = 0: Id:�Ŷӱ��: �ŶӺ���: ����: ��������: ����ID:  ����: ҽ������:�Ŷ�״̬ : �Ŷ�ʱ��: ҵ��ID
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.DeleteAll
        If mlngOrderStyle = 1 Then
            .SortOrder.Add .Columns(mCol.ORD)
            .SortOrder(0).SortAscending = True
        Else
            .SortOrder.Add .Columns(mCol.�Ŷ�״̬)
            .SortOrder(0).SortAscending = True

            .SortOrder.Add .Columns(mCol.�Ŷ����)
            .SortOrder(1).SortAscending = True

            .SortOrder.Add .Columns(mCol.����)
            .SortOrder(2).SortAscending = False

            .SortOrder.Add .Columns(mCol.���������)
            .SortOrder(3).SortAscending = True

            .SortOrder.Add .Columns(mCol.�Ŷ�ʱ��)
            .SortOrder(4).SortAscending = True

            .SortOrder.Add .Columns(mCol.�ŶӺ���)
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

Private Sub mobjQueue_OnRefresh(str��������() As String, ByVal strCur�������� As String, ByVal strCurҵ��ID As String, ByVal strMustCols As String, ByVal str���� As String, ByVal strҽ�� As String, ByVal strExcludeData As String, ByVal intViewDataType As Integer, ByVal strִ��״̬ As String, blnIsCustom As Boolean)
'���ܣ��Ŷӽк�ˢ��
    Dim j As Long, i As Long
    Dim strValue As String, strUninTable As String
    Dim rsLocal As ADODB.Recordset
    Dim rptCalling As ReportRecord
    Dim rptRecord As ReportRecord
    Dim str���������ַ��� As String
    
    On Error GoTo errH

    '�ݴ���
    If mobjQueueList Is Nothing Or mobjCallList Is Nothing Then
        Set mobjQueueList = mobjQueue.QueueList
        Set mobjCallList = mobjQueue.CallList
    End If
    If mobjQueueList Is Nothing Or mobjCallList Is Nothing Then
        Exit Sub
    End If

    '���Զ������̣�����113794ǰ�Ĵ���ʽ
     strValue = "": j = 0: strUninTable = ""
    If SafeArrayGetDim(str��������) > 0 Then
        For i = 1 To UBound(str��������)
            If Trim(str��������(i)) <> "" Then
                str���������ַ��� = str���������ַ��� & "," & str��������(i)
            End If
        Next i
    End If

    If str���������ַ��� <> "" Then str���������ַ��� = Mid(str���������ַ���, 2)
    Set rsLocal = GetRs�ŶӽкŲ����б�(str����, strҽ��, strִ��״̬, str���������ַ���, intViewDataType, p����ҽ��վ)

    'ɾ����Ҫ�ų�������,����ȡʵ���ŶӺ���ֵ�����
    If rsLocal.RecordCount > 0 Then rsLocal.MoveFirst
    While Not rsLocal.EOF
        If InStr(1, strExcludeData, rsLocal!ҵ������ & ":" & rsLocal!ҵ��ID) > 0 Then
            rsLocal.Delete
        End If
        If LenB(StrConv(Trim(Nvl(rsLocal("�ŶӺ���"))), vbFromUnicode)) > mlngMaxLen Then
            mlngMaxLen = LenB(StrConv(Trim(Nvl(rsLocal("�ŶӺ���"))), vbFromUnicode))
        End If
        rsLocal.MoveNext
    Wend

    rsLocal.Sort = "��������, �Ŷ�״̬ desc, �Ŷ����, ���� Desc, ���������, �Ŷ�ʱ��, �ŶӺ���"
    If rsLocal.RecordCount > 0 Then rsLocal.MoveFirst
    Call mobjQueueList.Records.DeleteAll
    Call mobjCallList.Records.DeleteAll
    While Not rsLocal.EOF
        If rsLocal("�Ŷ�״̬") = 7 Or rsLocal("�Ŷ�״̬") = 1 Then
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
    rriItem(mCol.����ID).Value = Nvl(rsData("����ID"))
    
    rriItem(mCol.��������).Caption = rsData("��������") & ":" & IIf(InStr(1, Nvl(rsData("��������")), ":") <= 0, "", Mid(Nvl(rsData("��������")), InStr(1, Nvl(rsData("��������")), ":") + 1))
    rriItem(mCol.��������).Value = Nvl(rsData("��������"))

    rriItem(mCol.��������).Value = Nvl(rsData("��������"))
    rriItem(mCol.����ID).Value = Nvl(rsData("����ID"))
    rriItem(mCol.�Ŷӱ��).Value = Nvl(rsData("�Ŷӱ��"))
    rriItem(mCol.�Ŷ����).Value = zlStr.Lpad(Nvl(rsData("�Ŷ����")), 20)
    rriItem(mCol.�ŶӺ���).Value = zlStr.Lpad(Nvl(rsData("�ŶӺ���")), mlngMaxLen)
    rriItem(mCol.�Ŷ�ʱ��).Value = Nvl(rsData("�Ŷ�ʱ��"))
    rriItem(mCol.����ʱ��).Value = Nvl(rsData("����ʱ��"))
    rriItem(mCol.�������).Value = Nvl(rsData("�������"))
    rriItem(mCol.���������).Value = Nvl(rsData("���������"))
    rriItem(mCol.����ҽ��).Value = Nvl(rsData("����ҽ��"))
    rriItem(mCol.��������).Value = DeptNametransform(Nvl(rsData("��������")))
    rriItem(mCol.��������).Caption = (Nvl(rsData("��������")))
    rriItem(mCol.ORD).Value = Format(rsData.AbsolutePosition, "00000000")
    
    If Nvl(rsData("�������")) = "" Then
        rriItem(mCol.��������).Icon = 807
    Else
        rriItem(mCol.��������).Icon = 3504
    End If
    
    
    If Nvl(rsData("�Ŷ�״̬")) = 1 Then
        rriItem(mCol.�Ŷ�״̬).Value = "������"
        For i = 0 To mobjQueueList.Columns.Count - 1
            rriItem(i).BackColor = &HC0C0FF
        Next
    ElseIf Nvl(rsData("�Ŷ�״̬")) = 0 Then
        rriItem(mCol.�Ŷ�״̬).Value = "�Ŷ���"
        For i = 0 To mobjQueueList.Columns.Count - 1
            rriItem(i).BackColor = IIf(InStr(rsData("����") & "", "��") > 0 And Val(rsData("����") & "") >= 80, &HC0FFC0, ColorConstants.vbWhite)
            rriItem(i).Bold = (InStr(rsData("����") & "", "��") > 0 And Val(rsData("����") & "") >= 80)
        Next
    ElseIf Nvl(rsData("�Ŷ�״̬")) = 3 Then
        rriItem(mCol.�Ŷ�״̬).Value = "��ͣ"
        For i = 0 To mobjQueueList.Columns.Count - 1
            rriItem(i).BackColor = ColorConstants.vbYellow
        Next
    ElseIf Nvl(rsData("�Ŷ�״̬")) = 4 Then
        rriItem(mCol.�Ŷ�״̬).Value = "����"
        For i = 0 To mobjQueueList.Columns.Count - 1
            rriItem(i).BackColor = ColorConstants.vbGreen
        Next
    ElseIf Nvl(rsData("�Ŷ�״̬")) = 7 Then
        rriItem(mCol.�Ŷ�״̬).Value = "�Ѻ���"
    Else
        rriItem(mCol.�Ŷ�״̬).Value = "������"
        For i = 0 To mobjQueueList.Columns.Count - 1
            rriItem(i).BackColor = &HC0C0C0
        Next
    End If
    
    If mlngQueueGroupType = 1 Then
        rriItem(mCol.ҽ������).Value = Nvl(rsData("��������")) & ":" & Nvl(rsData("ҽ������"))
    Else
        rriItem(mCol.ҽ������).Value = Nvl(rsData("ҽ������"))
    End If

    rriItem(mCol.ҵ������).Value = Nvl(rsData("ҵ������"))
    rriItem(mCol.ҵ��ID).Value = Nvl(rsData("ҵ��ID"))

    rriItem(mCol.����).Value = IIf(Nvl(rsData("����")) = 1, "����", "")
    
    If mlngQueueGroupType = 2 Then
        rriItem(mCol.����).Value = Nvl(rsData("��������")) & ":" & Nvl(rsData("����"))
    Else
        rriItem(mCol.����).Value = Nvl(rsData("����"))
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function DeptNametransform(ByVal strOldName) As String
'���ܣ��Ŷӽкŷ�������������ת����Ŀǰֻ֧�� һ��ʮ�Ĵ��� ��Сд����ת��Ϊ abc ������ʽ��������
    Dim strWord As String '�����ַ�
    Dim intCount As Integer
    Dim i As Integer
    
    On Error GoTo errH
    
    DeptNametransform = strOldName
    intCount = 0
    For i = 1 To Len(strOldName)
        strWord = Mid(strOldName, i, 1)
        If strWord = "һ" Or strWord = "��" Or strWord = "��" Or strWord = "��" Or strWord = "��" Or strWord = "��" Or _
           strWord = "��" Or strWord = "��" Or strWord = "��" Or strWord = "ʮ" Then
            intCount = intCount + 1
        End If
    Next
    If intCount = 1 Then
        DeptNametransform = Replace(strOldName, "һ", "a")
        DeptNametransform = Replace(DeptNametransform, "��", "b")
        DeptNametransform = Replace(DeptNametransform, "��", "c")
        DeptNametransform = Replace(DeptNametransform, "��", "d")
        DeptNametransform = Replace(DeptNametransform, "��", "e")
        DeptNametransform = Replace(DeptNametransform, "��", "f")
        DeptNametransform = Replace(DeptNametransform, "��", "g")
        DeptNametransform = Replace(DeptNametransform, "��", "h")
        DeptNametransform = Replace(DeptNametransform, "��", "i")
        DeptNametransform = Replace(DeptNametransform, "ʮ", "j")
    End If

    Exit Function
errH:
    DeptNametransform = strOldName
End Function

Private Sub SetCriticalAdvice(ByVal lng��¼ID As Long)
'���ܣ�ȷ����Σ��ֵ�󵯳�ҽ���´���棬�ղŵ�ǰ�����ҽ���뱾�εļ�¼������
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim objControl As Object
    
    On Error GoTo errH
    If lng��¼ID = 0 Then Exit Sub
    strSql = "select 1 from ����Σ��ֵ��¼ a where a.id=[1] and a.�Ƿ�Σ��ֵ=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng��¼ID)
    
    If Not rsTmp.EOF Then
        '�����´�ҽ���Ĵ���
        If GetTabTag <> "ҽ��" Then
            For i = 0 To tbcSub.ItemCount - 1
                If tbcSub.Item(i).Visible Then
                    If tbcSub.Item(i).Tag = "ҽ��" Then
                        tbcSub.Item(i).Selected = True
                        cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
                        Exit For
                    End If
                End If
            Next
        End If
        Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then
                objControl.Parameter = lng��¼ID
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

Private Sub LoadPatients����()
'���ܣ����غ�������б�
    Dim strSql As String
    Dim i As Long, j As Long
    Dim objRecord As ReportRecord
    Dim rsPati As ADODB.Recordset
    Dim lngColor As Long
    Dim rs��Ⱦ�������¼ As ADODB.Recordset
    Dim blnDo��Ⱦ��״̬ As Boolean
    Dim colPati As Collection, colValue As Collection
    Dim str����ids As String
    
    On Error GoTo errH
    Screen.MousePointer = 11
    mblnUnRefresh = True
     
    strSql = _
        " Select B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,B.����,m.�Ƿ���ɫͨ��,n.���� ���鼶��,n.���߱�ʶ��ɫ,B.����,b.����,D.���� as ���˿���," & _
        " B.ִ��ʱ�� as ʱ��,B.����ʱ��,B.ִ�в���ID,B.ִ����," & _
        " B.ת��״̬,C.���� as ת�����,B.ת������,B.ת��ҽ��,B.ִ��״̬,B.��¼��־" & _
        " From ���˹Һż�¼ B,���ű� C,���ű� D,��������¼ m,���ﲡ�鼶�� n" & _
        " Where B.����ID is not null And B.ת�����ID=C.ID(+) and B.ִ�в���ID=d.id " & _
        " And B.ִ��״̬=2 And B.ִ����||''=[1] And B.��¼����=1 And B.��¼״̬=1 and nvl(B.��¼��־,0) in (2,3) And B.���� = 1 And b.ID=m.�Һ�ID(+) And m.���鼶��=n.���(+)" & _
        " Order By B.NO"
    Set rsPati = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstr����ҽ��)
    
    str����ids = ""
    If Not rsPati Is Nothing Then
        For i = 1 To rsPati.RecordCount
             If InStr("," & str����ids & ",", "," & Val(rsPati!����ID & "") & ",") = 0 Then
                str����ids = str����ids & "," & Val(rsPati!����ID & "")
             End If
             rsPati.MoveNext
        Next
         If rsPati.RecordCount > 0 Then rsPati.MoveFirst
    End If
    str����ids = Mid(str����ids, 2)
    If str����ids <> "" Then
        Set colPati = PatiSvrGetVisitPatis(str����ids, "", p����ҽ��վ)
    End If
    
    
       
    Set rs��Ⱦ�������¼ = Get��Ⱦ�������¼(mstr����ҽ��, pt����)
    If rs��Ⱦ�������¼.RecordCount > 0 Then blnDo��Ⱦ��״̬ = True
    
    rptPati(PATI_RPT����).Records.DeleteAll
    For i = 1 To rsPati.RecordCount
    
       If Not colPati Is Nothing Then
            Set colValue = GetColObj(colPati, "_" & rsPati!����ID)
        End If
        
        If Not colValue Is Nothing Then
            If colValue.Count > 0 Then
                    
                Set objRecord = rptPati(PATI_RPT����).Records.Add()
                For j = 0 To rptPati(PATI_RPT����).Columns.Count - 1
                    objRecord.AddItem ""
                Next
                With objRecord
                    .Item(COL_JZ_��ʶ).Value = "��"
                    .Item(COL_JZ_����).Value = rsPati!���鼶��
                    .Item(COL_JZ_�����).Value = rsPati!����� & ""
                    .Item(COL_JZ_����).Value = rsPati!���� & ""
                    .Item(COL_JZ_����ʱ��).Value = Format(rsPati!ʱ��, "MM-dd HH:mm")
                    .Item(COL_JZ_�Ա�).Value = rsPati!�Ա� & ""
                    .Item(COL_JZ_����).Value = rsPati!���� & ""
                    .Item(COL_JZ_��ɫͨ��).Value = IIf(Val(rsPati!�Ƿ���ɫͨ�� & "") <> 0, "��", "")
                    .Item(COL_JZ_��).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
                    .Item(COL_JZ_NO).Value = rsPati!NO & ""

                    .Item(COL_JZ_���￨��).Value = GetColVal(colValue, "_vcard_no")
                    .Item(COL_JZ_��������).Value = GetColVal(colValue, "_pati_type")
                    .Item(COL_JZ_����ID).Value = rsPati!����ID & ""
                    .Item(COL_JZ_����ʱ��).Value = CStr(Format(rsPati!����ʱ��, "yyyy-MM-dd HH:mm:ss"))
                    .Item(COL_JZ_ִ�в���ID).Value = rsPati!ִ�в���ID & ""
                    .Item(COL_JZ_ִ����).Value = rsPati!ִ���� & ""
                    .Item(COL_JZ_���֤��).Value = GetColVal(colValue, "_pati_idcard")
                    .Item(COL_JZ_IC����).Value = GetColVal(colValue, "_iccard_no")
                    .Item(COL_JZ_��¼��־).Value = rsPati!��¼��־ & ""
                    .Item(COL_JZ_����).Value = rsPati!���� & ""
                    .Item(COL_JZ_���˿���).Value = rsPati!���˿��� & ""
                    
                    '���ղ����ú�ɫ��ʾ
                    If Not Val(GetColVal(colValue, "_insurance_type")) = 0 And GetColVal(colValue, "_pati_type") = "" Then
                        .Item(COL_JZ_�����).ForeColor = &HC0&
                        .Item(COL_JZ_��������).ForeColor = &HC0&
                    Else
                        '������ɫ
                        lngColor = GetPatiColor(GetColVal(colValue, "_pati_type"))
                        .Item(COL_JZ_�����).ForeColor = lngColor
                        .Item(COL_JZ_��������).ForeColor = lngColor
                    End If
                    
                    If rsPati!���߱�ʶ��ɫ <> "" Then
                        .Item(COL_JZ_��ʶ).BackColor = GetBGR_FromRGB(rsPati!���߱�ʶ��ɫ)
                    End If
            
                    '��Ӵ�Ⱦ��״̬
                    strSql = ""
                    If blnDo��Ⱦ��״̬ Then
                        rs��Ⱦ�������¼.Filter = "no='" & rsPati!NO & "'"
                        If Not rs��Ⱦ�������¼.EOF Then strSql = Get��Ⱦ��״̬(Val(rs��Ⱦ�������¼!��¼ & ""), Val(rs��Ⱦ�������¼!��д & ""), Val(rs��Ⱦ�������¼!״̬ & ""))
                    End If
                    .Item(COL_JZ_��Ⱦ��).Value = strSql
                End With
                
            End If
        End If
        rsPati.MoveNext
    Next
    rptPati(PATI_RPT����).Populate
    i = rptPati(PATI_RPT����).Records.Count
    tbcInTreat.Item(t����).Caption = "����" & IIf(i = 0, "", ":" & i & "��")
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
'���ܣ�����/ҽ�� ҳǩ�����л���������������/ҽ��
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
        If strTab = "ҽ��" Then
            If tbcRegist.Selected.Tag = "����һ��" Then tbcRegist.Item(mbyt���ξ���).Selected = True
            If tbcSub.Selected.Tag <> "ҽ��" Then tbcSub.Item(lngidx).Selected = True
            cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
            Set objControl = cbsMain.FindControl(, conMenu_Edit_NewItem * 10# + 1, True, True)
            If Not objControl Is Nothing Then
                If objControl.Enabled Then objControl.Execute
            End If
        ElseIf strTab = "����" Then
            If tbcRegist.Selected.Tag = "����һ��" Then tbcRegist.Item(mbyt���ξ���).Selected = True
            If tbcSub.Selected.Tag <> "����" Then tbcSub.Item(lngidx).Selected = True
            cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
            mblnUnRefresh = True
            Call mclsEPRs.zlOpenDefaultEPR(mstr�Һŵ�)
        End If
    End If
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ZLCommFun.PressKey(vbKeyTab)
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Exit Sub
        End If
    End If
End Sub

Private Sub txtPhone_GotFocus()
    Call zlControl.TxtSelAll(txtPhone)
End Sub

Private Sub UpdatePhone()
    Dim strTmp As String
    
    If mlng����ID = 0 Then Exit Sub
    On Error GoTo errH
    If txtPhone.ToolTipText <> txtPhone.Text Then
        strTmp = txtPhone.Text
        
        Call Update������Ϣ(mlng����ID, "�ֻ���", strTmp, p����ҽ��վ)
        
        txtPhone.Tag = ""
        txtPhone.ToolTipText = txtPhone.Text
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtPhone_Validate(Cancel As Boolean)
    If (Not PatiIdentify.IsMobileNo(txtPhone.Text)) And txtPhone.Text <> "" Then
        MsgBox "��ǰ¼����ֻ��Ÿ�ʽ����ȷ��������¼��!", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
    Call UpdatePhone
End Sub

Private Sub ReadMsgAuto()
'���ܣ�Σ��ֵ��Ϣ�����Զ�����
    Dim i As Long
    Dim lng����ID As Long
    Dim strNO As String
    Dim strҵ�� As String
    Dim lng��ϢID As Long
    Dim lng�Һ�id As String
    Dim str�Һŵ� As String
    Dim blnRs As Boolean
    
    On Error GoTo errH

    For i = i To rptNotify.Rows.Count - 1
        With rptNotify.Rows(i)
            If Not .GroupRow Then
                strNO = .Record(C_��Ϣ).Value
                If strNO = "ZLHIS_LIS_003" Or strNO = "ZLHIS_PACS_005" Then
                    lng����ID = Val(.Record(C_����ID).Value)
                    lng�Һ�id = Val(.Record(C_�Һ�Id).Value)
                    str�Һŵ� = .Record(C_No).Value
                    strҵ�� = .Record(C_ҵ��).Value
                    
                    lng��ϢID = Val(.Record(C_Id).Value)
                    blnRs = ReadMsg(lng����ID, lng�Һ�id, strNO, strҵ��, lng��ϢID, str�Һŵ�, 0)
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

Private Sub UpdatePhysical(ByVal strInfo As String)
    txtInfo(txtInfo������).Text = strInfo
    txtInfo(txtInfo������).ToolTipText = strInfo
    If strInfo = "" Then
        lblPhysical.Caption = ""
    Else
        lblPhysical.Caption = "���������:"
    End If
    txtInfo(txtInfo������).Visible = Not (lblPhysical.Caption = "")
End Sub

Private Function OpenExaReportNew() As Boolean
'����:�����°����ӿڲ����鿴����
    Dim objHealthITranLib As Object
    On Error GoTo errH
    Set objHealthITranLib = CreateObject("zlHealthITranLib.clsITranLib")
    If objHealthITranLib.Initialize(gcnOracle) And objHealthITranLib.ZLHEC.Initialize(Me) Then
        Call objHealthITranLib.ZLHEC.OpenExaminationReport(mlng����ID)
    End If
    OpenExaReportNew = True
    Exit Function
errH:
    '���Դ��󷵻�false����,���������ϰ�ӿ�
    Err.Clear
End Function

Public Function LocateMsgPati(lng����ID As Long, lng�Һ�id As Long, lngҽ��ID As Long) As Boolean
'���ܣ���λ��ָ���Ĳ���
    Dim i As Integer
    Dim blnFinded As Boolean

    If lng����ID = 0 Then Exit Function
    
    If mlng����ID = lng����ID And mlng�Һ�ID = lng�Һ�id Then
        blnFinded = True
    Else
        Call ExecuteFindPati(False, , , lng����ID)
        If mlng����ID = lng����ID Then blnFinded = True
    End If
    
    If blnFinded Then   '�ҵ����˺��پ����Ƿ�λҽ��
        If GetTabTag <> "ҽ��" Then
            '��λ��ҽ����Ϣҳ
            For i = 0 To tbcSub.ItemCount - 1
                If tbcSub(i).Visible And tbcSub(i).Tag = "ҽ��" Then
                    tbcSub.Item(i).Selected = True
                End If
            Next
        End If
        If lngҽ��ID <> 0 Then
            Call mclsAdvices.LocatedAdviceRow(lngҽ��ID)
        End If
    End If
End Function

Private Sub ExecAdjustGrade()
'���ܣ��������鼶��
    Dim blnOK As Boolean
    
    If mPatiInfo.���鼶�� = "" Then
        MsgBox "û�з���Ĳ��˲��ܵ������鼶��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    blnOK = frmEMCAdjustGrade.ShowMe(Me, mPatiInfo.�Һ�ID, mPatiInfo.���� & "," & mPatiInfo.�Ա�, mPatiInfo.���鼶��)
    
    If blnOK Then
        '�����ԤԼ���˲��ܵ�������
        Call LoadPatients("11010")
        Call ReshDataQueue
    End If
End Sub


Private Sub ExecTagGreen()
'���ܣ���ɫͨ�����
    Dim strPrompt As String, strSql As String
    
    If mPatiInfo.���鼶�� = "" Then
        MsgBox "û�з���Ĳ��˲��ܱ����ɫͨ����", vbInformation, gstrSysName
        Exit Sub
    End If
    On Error GoTo errH
    
    If mPatiInfo.�Ƿ���ɫͨ�� = 0 Then
        strPrompt = "��ȷ��Ҫ�ԡ�" & mPatiInfo.���� & "�������ɫͨ����" & vbCrLf & "��ɫͨ�����˽�ʵ�������ƺ󸶷ѡ�"
    Else
        strPrompt = "��ȷ��Ҫ�ԡ�" & mPatiInfo.���� & "��ȡ����ɫͨ�������"
    End If
    If MsgBox(strPrompt, vbQuestion + vbOKCancel + vbDefaultButton1, "��ʾ") = vbOK Then
               
        strSql = "Zl_������ɫͨ��_Edit(" & mPatiInfo.�Һ�ID & "," & IIf(mPatiInfo.�Ƿ���ɫͨ�� = 1, 0, 1) & ")"
        Call zlDatabase.ExecuteProcedure(strSql, "��ɫͨ�����")
        Call LoadPatients
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Function GetTabTag() As String
'���ܣ���ȡ��ǰѡ��ҳǩTag
    Dim i As Integer
    Dim strTag As String

    For i = 0 To tbcSub.ItemCount - 1
        If tbcSub(i).Visible And tbcSub.Item(i).Selected Then
            strTag = tbcSub(i).Tag
            Exit For
        End If
    Next
    GetTabTag = strTag
End Function


   
Private Function ReLoadPicture(ByVal strFile As String, ByRef strCopyFile As String) As Boolean

    Dim objPic As StdPicture
    Dim blnComPress As Boolean
    Dim w As Long, h As Long
    Dim X As Single
    Dim intSplit As Integer, intMaxLength As Integer
    Dim strFilePath As String, strFileFormat As String
    Dim objFile     As New FileSystemObject
    
    On Error GoTo errHand
    If strFile = "" Then Exit Function
    Set objPic = LoadPicture(strFile)
    If objPic Is Nothing Then Exit Function
    
    intMaxLength = 500
    
    intSplit = InStrRev(strFile, ".")
    strFilePath = Left(strFile, intSplit - 1)
    strFileFormat = Mid(strFile, intSplit + 1)
    
    
    w = Me.picCharge.ScaleX(objPic.Width, vbHimetric, Me.picCharge.ScaleMode)
    h = Me.picCharge.ScaleY(objPic.Height, vbHimetric, Me.picCharge.ScaleMode)
    
    If w > Me.picCharge.ScaleX(intMaxLength, vbPixels, Me.picCharge.ScaleMode) Or h > Me.picCharge.ScaleY(intMaxLength, vbPixels, Me.picCharge.ScaleMode) Then
        If w > h Then
            If w > Me.picCharge.ScaleX(intMaxLength, vbPixels, Me.picCharge.ScaleMode) Then
                X = h / w
                w = Me.picCharge.ScaleX(intMaxLength, vbPixels, Me.picCharge.ScaleMode)
                h = w * X
            End If
        Else
            If h > Me.picCharge.ScaleY(intMaxLength, vbPixels, Me.picCharge.ScaleMode) Then
                X = w / h
                h = Me.picCharge.ScaleY(intMaxLength, vbPixels, Me.picCharge.ScaleMode)
                w = h * X
            End If
        End If
        Me.picCharge.Width = w
        Me.picCharge.Height = h
        
        Me.picCharge.PaintPicture objPic, 0, 0, Me.picCharge.Width, Me.picCharge.Height
        
        Set objPic = Me.picCharge.Image
        blnComPress = True
    End If
    
    If blnComPress Or strFileFormat <> "JPG" Then
        '������ʽͼƬתJPG����ѹ��ͼƬ
        strCopyFile = objFile.GetSpecialFolder(TemporaryFolder) & "\" & objFile.GetTempName
        Call ZLCommFun.ConvertPicture(objPic, strCopyFile)
    End If
    
    ReLoadPicture = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function




