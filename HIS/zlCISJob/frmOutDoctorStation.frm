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
   Caption         =   "����ҽ������վ"
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
            Name            =   "����"
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
         Caption         =   "ժҪ:"
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
            Name            =   "����"
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
               Name            =   "����"
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
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "����"
         Top             =   120
         Width           =   1620
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FDFDFD&
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
         Height          =   200
         Index           =   1
         Left            =   3240
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "��"
         Top             =   165
         Width           =   465
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FDFDFD&
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
         Height          =   200
         Index           =   2
         Left            =   3795
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "27��"
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
               Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "����:"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����������Ϣ��"
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
         Index           =   4
         Left            =   8580
         MouseIcon       =   "frmOutDoctorStation.frx":1454
         MousePointer    =   99  'Custom
         TabIndex        =   55
         ToolTipText     =   "������Ϣ:ժҪ������������"
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
         Caption         =   "�ѱ�:"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "���ն�ƾ���"
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
         Caption         =   "�޸Ĳ��˻�����Ϣ"
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
         MouseIcon       =   "frmOutDoctorStation.frx":7A5C
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   990
         Width           =   1440
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "���ѷ�ʽ:"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "���￨��:"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "ҽ������:"
         BeginProperty Font 
            Name            =   "����"
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
            Caption         =   "����:"
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
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmOutDoctorStation.frx":93DE
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
         Caption         =   "��������"
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
         Caption         =   "����ʱ��"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17145
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
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":A2E7
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":A881
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":AE1B
            Key             =   "ת��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":B3B5
            Key             =   "�ܾ�"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":B94F
            Key             =   "��ͣ"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutDoctorStation.frx":BEE9
            Key             =   "����"
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
    txtInfo���ѷ�ʽ = 11
    txtInfo���� = 8

    lblLink�ļ� = 0
    lblLink�ɼ� = 1
    lblLink��� = 2
    lblLink�޸� = 3
    lblLink��ʾ = 4

    '����:3�ˣ����:45�ˣ�����:15��
    t���� = 0
    t��� = 1
    t���� = 2
End Enum

Private Enum PATI_COL_����
    COL_HZ_��ʶ = 0
    COL_HZ_�����
    COL_HZ_����
    COL_HZ_�Һ�ʱ��
    COL_HZ_�Ա�
    COL_HZ_����
    COL_HZ_��
    COL_HZ_��
    COL_HZ_NO
    COL_HZ_����
    COL_HZ_��������
    COL_HZ_����ҽ��
    COL_HZ_���
    COL_HZ_����ʱ��
    COL_HZ_���￨��
    COL_HZ_��������
    COL_HZ_ת��״̬
    COL_HZ_ԤԼҽ��
    COL_HZ_ԤԼʱ��
    COL_HZ_���֤��
    COL_HZ_����
    COL_HZ_���˿���
    
'������
    COL_HZ_����ID
    COL_HZ_����ʱ��
    COL_HZ_ִ�в���ID
    COL_HZ_ִ����
    COL_HZ_״̬ 'ת��״̬��־
    COL_HZ_IC����
    COL_HZ_��¼��־
    COL_HZ_ִ��״̬
End Enum

Private Enum PATI_COL_���� '�����б�ͻ����б���
    COL_JZ_��ʶ = 0
    COL_JZ_�����
    COL_JZ_����
    COL_JZ_����ʱ��
    COL_JZ_�Ա�
    COL_JZ_����
    COL_JZ_��
    COL_JZ_��
    COL_JZ_NO
    COL_JZ_����
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
    COL_YZ_NO
    COL_YZ_�����
    COL_YZ_����
    COL_YZ_�Ա�
    COL_YZ_����
    COL_YZ_��
    COL_YZ_��
    COL_YZ_����
    COL_YZ_ʱ��
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
    c_���� = 3
    C_����� = 4
    C_����ʱ�� = 5
    C_״̬ = 6
    '������
    C_��Ϣ = 7
    C_��� = 8
    C_���� = 9
    C_ҵ�� = 10
    C_�Һ�Id = 11
    C_Id = 12
End Enum

Private Type PatiInfo
    ���� As PatiType
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
    ���� As Integer
    ·��״̬ As Integer
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
    �Һŵ� As String
    ����� As String
    ���￨ As String
    ���� As String
End Type
Private mvCondFilter As COND_FILTER

'�Ӵ��������
Private mclsEMR As Object  '�°没��zlRichEMR.clsDockEMR
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
Private mobjMsg As Object '�����������ݶ���

'�������ñ���
Private mint���ﷶΧ As Integer '1-����,2-������,3-������
Private mlng�������ID As Long
Private mstr�������� As String
Private mstr����ҽ�� As String
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
Private mobjKernel As zlPublicAdvice.clsPublicAdvice          '�ٴ����Ĳ���

'ҽ�ƿ�
Private mobjSquareCard As Object      '���������
Private mstrCardKind As String        '��������󷵻صĿ��õ�ҽ�ƿ�

Private mstrPrePati As String
Private mintPreTime As Integer
Private mlngCommunityID As Long '�Զ�ִ�е���������
Private mbytSize As Byte '���� 0-С���壨9�����壩��1-�����壨12�����壩
Private mblnTabTmp As Boolean
Private mstrPreSubTab As String ' tbsSubǰһ��ѡ�е�ҳǩ
Private mblnSizeTmp As Boolean

Private mblnMsgOk As Boolean '�Ƿ�����Ϣ����
Private mblnFirstMsg As Boolean 'mblnFirstMsg=false ��ʾ��ҽ��վ��ĵ�һ����Ϣ
Private mintNotify As Integer 'ҽ�������Զ�ˢ�¼��(����)
Private mintNotifyDay As Integer '���Ѷ������ڵ�ҽ��
Private mstrNotifyAdvice As String '���ѵ�ҽ������
Private mstrPreNotify As String
Private mblnPatiDetail As Boolean
Private mclsMsg As clsCISMsg
Private mrsMsg As ADODB.Recordset
Private mbln��Ϣ���� As Boolean
Private mlng���ս������� As Long
Private mblnΣ��ֵ As Boolean '��Σ��ֵ��Ȩ��
Private mbln��ʾԤԼ���� As Boolean
Private mintԤԼ�б� As Integer

Private Sub cboPayType_Click()
    Dim strTmp As String
    Dim strSQL As String
    If mstr�Һŵ� = "" Then Exit Sub
    strTmp = Split(cboPayType.Text, "-")(1)
    If cboPayType.ToolTipText <> strTmp Then
        strTmp = Split(cboPayType.Text, "-")(1)
        strSQL = "Zl_���˹Һż�¼_���·ѱ�('" & mstr�Һŵ� & "','" & strTmp & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        cboPayType.Tag = ""
        cboPayType.ToolTipText = strTmp
    End If
End Sub

Private Sub cboBillType_Click()
    Dim strSQL As String
    Dim strTmp As String
    If mlng����ID = 0 Then Exit Sub
    If cboBillType.ToolTipText <> cboBillType.Text Then
        strTmp = cboBillType.Text
        strSQL = "Zl_������Ϣ_������Ϣ(" & mlng����ID & ",'�ѱ�','" & strTmp & "')"
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
    mvCondFilter.�Һŵ� = ""
    mvCondFilter.���￨ = ""
    mvCondFilter.����� = ""
    mvCondFilter.���� = ""
    '�����������֤ÿ���ط���ȡ�ĳ�Ժ���˶�����ͬһʱ�䷶Χ�ڣ�72783��
    Call zlDatabase.SetPara("���ﲡ�˽������", DateDiff("d", datCurr, mvCondFilter.End), glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
    Call zlDatabase.SetPara("���ﲡ�˿�ʼ���", DateDiff("d", mvCondFilter.Begin, datCurr), glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
    cboSelectTime.ToolTipText = Format(mvCondFilter.Begin, "yyyy-MM-dd") & " - " & Format(mvCondFilter.End, "yyyy-MM-dd")
    lblSeeTim.ToolTipText = cboSelectTime.ToolTipText
    mintOutPreTime = cboSelectTime.ListIndex
    Call LoadPatients����
End Sub

Private Sub cmdOtherFilter_Click()
    Dim datCurr As Date
    
    With mvCondFilter
        .����ID = IIf(.����ID = 0, mlng�������ID, .����ID)
        If frmPatiFilter.ShowMe(Me, .Begin, .End, .����ID, .ҽ��, .�Һŵ�, .�����, .���￨, .����, mstrPrivs) Then
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

Private Sub mclspath_RequestRefresh(ByVal lngPathState As Long)
'���ܣ��ٴ�·����ˢ�²�����Ϣ�б��е�״̬,-1��ʾδ����״̬
    mPatiInfo.·��״̬ = lngPathState
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
        If 0 = Val(zlDatabase.GetPara("�кŷ�ʽ", glngSys, p�Ŷӽк�����ģ��)) Then
            mty_Queue.blnҽ���������� = Val(zlGetLocaleComputerNamePara("�ŶӺ���վ��", glngSys, p����������, "0")) = 1
        Else
            mty_Queue.str����վ�� = zlDatabase.GetPara("Զ�˺���վ��", glngSys, p�Ŷӽк�����ģ��)
            mty_Queue.blnҽ���������� = Val(zlGetLocaleComputerNamePara("�ŶӺ���վ��", glngSys, p����������, "0", mty_Queue.str����վ��)) = 1
        End If
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
    Dim intType As Integer, blnHave As Boolean, blnTmp As Boolean
    Dim i As Integer, arrType() As String
    Dim objControl As CommandBarControl, objTabItem As TabControlItem
    Dim arrTmp As Variant, strTmp As String
    
    mstrPrivs = ";" & gstrPrivs & ";"
    mlngModul = glngModul
    mblnShowLeavePati = False
    Call GetLocalSetting '���ز���
    
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, p����ҽ��վ, GetInsidePrivs(p����ҽ��վ))
    Call AddMipModule(mclsMipModule)
    
    Set mclsReg = New zlPublicExpense.clsRegist
    Call mclsReg.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
    Call mclsReg.zlInitData(1)
    
    Set mclsDis = New zl9Disease.clsDisease
    Call mclsDis.InitDisease(gcnOracle, Me, glngSys, glngModul, mstrPrivs, mclsMipModule)
    
    Set mobjKernel = New zlPublicAdvice.clsPublicAdvice

    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)
    
    Call InitQueuePara
 
    'һ��ͨ������ʼ������tbcSub_SelectedChanged֮ǰ���Ա㴫�ݸ�ҽ������
     'zlGetIDKindStr�л��Զ�����Ϊ����8λ����
    mstrCardKind = "��|���￨|0|0|8|0|0|0;��|��ʶ��|0|0|0|0|0|0;��|�Һŵ�|0|0|0|0|0|0;��|����|0|0|0|0|0|0;��|�������֤|0|0|0|0|0|0;�ɣ�|�ɣÿ�|1|0|0|0|0|0"
    If Check�Ŷӽк� = True Then mstrCardKind = mstrCardKind & ";��|�ŶӺ�|0|0|0|0|0|0;ҽ|ҽ����|0|0|0|0|0|0"
    On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    err.Clear: On Error GoTo 0
    If Not mobjSquareCard Is Nothing Then
        If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
            Set mobjSquareCard = Nothing
            MsgBox "ҽ�ƿ�������zl9CardSquare����ʼ��ʧ��!", vbInformation, gstrSysName
        Else
            mstrCardKind = mobjSquareCard.zlGetIDKindStr(mstrCardKind)
        End If
    End If
    Call PatiIdentify.zlInit(Me, glngSys, p����ҽ��վ, gcnOracle, gstrDBUser, mobjSquareCard, mstrCardKind, "zl9CISJob")
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
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
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
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, True)
 
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
        .InsertItem(1, "���", picYZ.hwnd, 0).Tag = "���ﲡ��"
        .InsertItem(2, "����", picHUIZ.hwnd, 0).Tag = "����ﲡ��"
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
    mcolSubForm.Add mclsPath.zlGetForm, "_·��"
    mcolSubForm.Add mclsAdvices.zlGetForm, "_ҽ��"
    mcolSubForm.Add mclsEPRs.zlGetForm, "_����"
    mcolSubForm.Add mclsDisease.zlGetForm, "_��������"
    If Not mclsEMR Is Nothing Then
        mcolSubForm.Add mclsEMR.zlGetForm, "_�²���"
    End If
    
    
    '---------------------------------------------------
    '���ξ����б�
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
        If GetInsidePrivs(P����·��Ӧ��, True) <> "" Then
            .InsertItem(intIdx, "�ٴ�·��", picTmp.hwnd, 0).Tag = "·��": intIdx = intIdx + 1
        End If
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
            Call zlPlugInErrH(err, "GetFormCaption")
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ",")
                For i = 0 To UBound(arrTmp)
                    strTmp = arrTmp(i)
                    mcolSubForm.Add gobjPlugIn.GetForm(glngSys, p����ҽ��վ, strTmp), "_" & strTmp
                    .InsertItem(intIdx, strTmp, mcolSubForm("_" & strTmp).hwnd, 0).Tag = strTmp: intIdx = intIdx + 1
                    Call zlPlugInErrH(err, "GetForm")
                Next
            End If
            err.Clear: On Error GoTo 0
        End If
        
        If .ItemCount = 0 Then
            MsgBox "��û��ʹ������ҽ������վ��Ȩ�ޡ�", vbInformation, gstrSysName
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
    
    'ReportControl�ؼ����������޷��ָ�Ҫ��������
    For i = 0 To rptPati.Count - 1
        strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\ReportControl", "rptPati" & "_" & i, "")
        rptPati(i).LoadSettings strTmp
    Next
    
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

    If ISPassShowCard Then
        rptPati(PATI_RPT����).Columns(COL_HZ_���￨��).Visible = False
        rptPati(PATI_RPTԤԼ).Columns(COL_HZ_���￨��).Visible = False
        rptPati(PATI_RPT����).Columns(COL_JZ_���￨��).Visible = False
        rptPati(PATI_RPT����).Columns(COL_JZ_���￨��).Visible = False
        rptPati(PATI_RPT����).Columns(COL_YZ_���￨��).Visible = False
    End If
    If InStr(";" & gstrPrivs & ";", ";�޸�ҽ�Ƹ��ʽ;") = 0 Then
        cboPayType.Locked = True
    End If
    If InStr(";" & gstrPrivs & ";", ";�޸ķѱ�;") = 0 Then
        cboBillType.Locked = True
    End If
    Call SetReceiveToday(True, 0)
    If mblnPatiDetail Then
        lblLink(lblLink��ʾ).Caption = "����������Ϣ��"
    Else
        lblLink(lblLink��ʾ).Caption = "��ʾ������Ϣ��"
    End If
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
    If gbytPass = 3 Then
        On Error Resume Next
        If gobjPass Is Nothing Then
            Set gobjPass = DynamicCreate("zlPassInterface.clsPass", "������ҩ���", True)
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
    Dim str�Һŵ� As String, strCardNO As String
    Dim rsTmp As Recordset
    Dim str����ID As String, str���ID As String
    Dim intFindTypeTmp As Integer
    Dim strPictureFile As String
    
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
        Call LoadPatients("11001")
        Call LoadNotify
    Case conMenu_View_Jump '��ת
        If Me.tbcSub.Selected.Index + 1 <= Me.tbcSub.ItemCount - 1 Then
            Me.tbcSub.Item(Me.tbcSub.Selected.Index + 1).Selected = True
        Else
            Me.tbcSub.Item(0).Selected = True
        End If
    Case conMenu_File_Parameter '��������
        frmOutStationSetup.mstrPrivs = mstrPrivs
        frmOutStationSetup.Show 1, Me

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
    Case conMenu_Tool_CISMed  '�ٴ��Թ�ҩ
        Call Set�ٴ��Թ�ҩ(Me)
     Case conMenu_Tool_TransAudit '��Ѫ��˹���
        On Error Resume Next
        Call frmExamineTransfuse.ShowMe(Me, 1)
    Case conMenu_Tool_Archive '���Ӳ�������
        Call frmArchiveView.ShowArchive(Me, mPatiInfo.����ID, mPatiInfo.�Һ�ID)
    Case conMenu_Tool_ExaReport
        '���ó¸����ṩ�Ľӿ� OpenExaminationReport
    Case conMenu_Tool_Reference_1 '������ϲο�
        Call gobjKernel.ShowDiagHelp(vbModeless, Me)
    Case conMenu_Tool_Reference_2 '���ƴ�ʩ�ο�
        Call gobjKernel.ShowClincHelp(vbModeless, Me)
    Case conMenu_Manage_FeeItemSet  '������Ŀ��������
        Call Set������Ŀ��������
        
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
        str�Һŵ� = frmForceGet.ShowMe(Me, mstrPrivs, mlng�������ID, mobjSquareCard)
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
        If Not mobjSquareCard Is Nothing Then
            Call mobjSquareCard.zlHealthArchivesShow(Me, p����ҽ��վ, mlng����ID, "")
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
            Case "·��"
                Call mclsPath.zlExecuteCommandBars(Control)
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
       Case "·��"
            Call mclsPath.zlPopupCommandBars(CommandBar)
       Case "����"
       End Select
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    Dim strTmp As String
 
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
    Case conMenu_Tool_CISMed  '�ٴ��Թ�ҩ
        If InStr(GetInsidePrivs(p����ҽ��վ), ";�ٴ��Թ�ҩ;") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_Tool_Archive '���Ӳ�������
        If GetInsidePrivs(p���Ӳ�������) = "" Then
            Control.Visible = False
        Else
            Control.Enabled = mlng����ID <> 0
        End If
    Case conMenu_Tool_ExaReport
        Control.Enabled = mlng����ID <> 0
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
    Case conMenu_Manage_FeeItemSet '������Ŀ��������,û��Ȩ��ʱ�ɲ鿴
                
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
                        strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_HZ_״̬).Value
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
            blnEnabled = (mintActive = pt����)
            If blnEnabled Then
                If mPr <> -1 And rptPati(mintRPTIndex).Visible Then
                    strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_HZ_ִ��״̬).Value
                    blnEnabled = Val(strTmp) = 0
                Else
                    blnEnabled = False
                End If
            End If
            Control.Enabled = blnEnabled
    Case conmenu_Edit_Wait
        blnEnabled = mintActive = pt����
        If blnEnabled Then
            If mPr <> -1 And rptPati(mintRPTIndex).Visible Then
                strTmp = rptPati(mintRPTIndex).Rows(mPr).Record(COL_HZ_ִ��״̬).Value
                blnEnabled = Val(strTmp) = -1
            Else
                blnEnabled = False
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
        Control.Enabled = IIf(Val(mlng����ID) = 0, False, True)
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
        Case "·��"
            Call mclsPath.zlUpdateCommandBars(Control)
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
        
    Me.Caption = "����ҽ������վ - " & objItem.Caption & "(��ǰ�û���" & UserInfo.���� & ")"
    
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
        Call mclsAdvices.zlDefCommandBars(Me, Me.cbsMain, 0, gobjPlugIn, mobjSquareCard)
    Case "����"
        Call mclsEPRs.zlDefCommandBars(Me.cbsMain)
    Case "�²���"
        Call mclsEMR.zlDefCommandBars(Me.cbsMain)
    Case "��������"
        Call mclsDisease.zlDefCommandBars(Me.cbsMain)
    Case "·��"
        Call mclsPath.zlDefCommandBars(Me, Me.cbsMain)
    Case Else
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strName = gobjPlugIn.GetButtomName(glngSys, p����ҽ��վ, mcolSubForm("_" & objItem.Tag), objItem.Tag)
            Call zlPlugInErrH(err, "GetButtomName")
            '�����˵�
            If strName <> "" Then Call PlugInInSideBar(cbsMain, strName)
            err.Clear: On Error GoTo 0
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
            Call mobjPati.zlRefresh(0, 0, False, False)
        Case "·��"
            Call mclsPath.zlRefresh(0, 0, "", 0, 0, False)
        Case "ҽ��"
            Call mclsAdvices.zlRefresh(0, "", False)
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
                Call zlPlugInErrH(err, "RefreshForm")
                err.Clear: On Error GoTo 0
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
                Call mobjPati.zlRefresh(.����ID, .�Һ�ID, Not tbcRegist.Item(mbyt���ξ���).Selected Or .���� = pt���� Or mstr�Һŵ� <> .�Һŵ�, .����ת��)
            Case "·��"
                Call mclsPath.zlRefresh(.����ID, .�Һ�ID, .�Һŵ�, mlng����ID, .����, .����ת��, True, mclsMipModule)
            Case "ҽ��"
                Call mclsAdvices.zlRefresh(.����ID, .�Һŵ�, mstr�Һŵ� = .�Һŵ� And (.���� = pt���� Or .���� = pt����) And mlng����ID <> 0, .����ת��, , , mclsMipModule, , mPatiInfo.·��״̬, .����)
            Case "����"
                Call mclsEPRs.zlRefresh(.����ID, .�Һ�ID, mlng����ID, mstr�Һŵ� = .�Һŵ� And mlng����ID = .����ID And (.���� = pt���� Or .���� = pt����) And mlng����ID <> 0, .����ת��, True)
            Case "�²���"
                Call mclsEMR.zlRefresh(.����ID, .�Һ�ID, mlng����ID, .����, 1)
            Case "��������"
                If objItem.Visible Then
                    Call mclsDisease.zlRefresh(.����ID, .�Һ�ID, 1, mlng����ID, .����ת��, mstr�Һŵ� = .�Һŵ� And mlng����ID = .����ID And mlng����ID <> 0)
                End If
            Case "����һ��"
                Call mfrmView.zlRefresh(Me, .����ID, mlng����ID)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.RefreshForm(glngSys, p����ҽ��վ, mcolSubForm("_" & objItem.Tag), objItem.Tag, .����ID, .�Һŵ�, 0, .����ת��, 0, 0)
                    Call zlPlugInErrH(err, "RefreshForm")
                    err.Clear: On Error GoTo 0
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
    Dim objCustom As CommandBarControlCustom
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
        Set objControl = .Add(xtpControlButton, conMenu_Tool_CISMed, "�ٴ��Թ�ҩ(&J)")
        objControl.IconId = 3901
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "���Ӳ�������(&I)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_ExaReport, "��������ܼ챨��")
            objControl.IconId = conMenu_File_Preview
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "���ϲο�(&R)"): objPopup.BeginGroup = True
        objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_Reference_1, "������ϲο�(&D)", -1, False
            .Add xtpControlButton, conMenu_Tool_Reference_2, "���ƴ�ʩ�ο�(&C)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Manage_FeeItemSet, "������Ŀ��������(&C)")
        
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
        If mobjSquareCard.zlHealthArchiveIsSHow(Me, p����ҽ��վ, strFunName, "") Then
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
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Finish, "���"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBack, "�����")
        objControl.IconId = conMenu_Edit_Pause
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReBackCancel, "ȡ��")
        objControl.IconId = conMenu_Edit_Reuse
        objControl.ToolTipText = "ȡ������"
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItemQAdvice, "ҽ��")
        objControl.IconId = conMenu_Edit_NewItem
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItemQEpr, "����")
        objControl.IconId = conMenu_Edit_NewParent

        Set objPopup = .Add(xtpControlPopup, conMenu_Tool_Community, "����")
        objPopup.ID = conMenu_Tool_Community
        objPopup.IconId = conMenu_Tool_Community
        Set objControl = .Add(xtpControlButton, conMenu_Tool_TransAudit, "��Ѫ���")
        objControl.IconId = 3551
         
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
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL1_INSIDE_1260_2")
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
    Call mclsDisease.EditDiseaseDoc(Me, mlng����ID, mlng�Һ�ID, 1, mlng����ID, str����ID, str���ID)
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
    Dim objItem As ListItem, lngCount As Long, rsTemp As ADODB.Recordset
    Dim strSQL As String, strLimit As String, strResult As String, arrCheck As Variant
    
    If Val(strҵ��ID) <> 0 Then
           strSQL = "Select Zl_QueuedateCheck([1]) as Chk From Dual"
           Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strҵ��ID))
           strResult = NVL(rsTemp!chk) & "|"
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
    strLimit = ",0,4," & IIf(mty_Queue.bln���к�����, "", ",6,")
    strSQL = "" & _
    "   Select Count(distinct B.ID) as Count From ���˹Һż�¼ B ,�ŶӽкŶ��� A" & _
    "   Where A.ҵ��ID=B.ID And A.ҵ������=0  " & _
    "               And instr([4],','||A.�Ŷ�״̬||',')=0   And B.��¼����=1 And B.��¼״̬=1" & _
    "               And A.ҽ������||''=[1]   " & IIf(mty_Queue.bln���к�����, " And nvl(A.�������,0) = 0", "") & _
    "               And (  (nvl(B.����,0)=1  and B.����ʱ��>=Sysdate-[3] ) or   (nvl(B.����,0)<>1  and B.����ʱ��>=Sysdate-[2] )) "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����, IIf(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����), IIf(gint����Һ����� = 0, 1, gint����Һ�����), strLimit)
    lngCount = Val(NVL(rsTemp!Count))

    If lngCount >= mty_Queue.int�������� Then
            MsgBox "���ֻ����" & mty_Queue.int�������� & "�����ﲡ��,�����ٽ��к��У�", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
    End If
    CheckIsAskNextQueue = True
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
    Dim strSQL As String, rsTemp As ADODB.Recordset
   ' byt�������� -0 - ����, 1 - ֱ��, 2 - ����, 3 - ��ͣ, 4 - ��ɾ���, 5 - �㲥
   
    If InStr(1, "15", byt��������) = 0 Then Exit Sub
    If CheckIsAskNextQueue(strҵ��ID) = False Then blnCancel = True: Exit Sub
    
    strSQL = "SELECT a.ID,a.No,a.����ID,a.ִ�в���ID,A.ִ��״̬ From ���˹Һż�¼ A,�ŶӽкŶ��� B  " & _
        "  where  a.ID=b.ҵ��id and b.ҵ������=0 and a.ID=[1] and nvl(b.�Ŷ�״̬,0)=0 And a.��¼���� in(1,2) And a.��¼״̬=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strҵ��ID))
    If rsTemp.EOF Then Exit Sub
    
    '68736:������,2014-02-18,ת�ﲡ��û��������Ϣ
    If byt�������� = 1 Then
        If Isת�ﲡ��(strҵ��ID) Then
            If CheckTransferDetail(strҵ��ID) = False Then
                strSQL = "ZL_���˹Һż�¼_�������� ('" & NVL(rsTemp!NO) & "'," & Val(NVL(rsTemp!����ID)) & ",'" & mstr�������� & "','" & UserInfo.���� & "',to_Date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),2)"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
            Exit Sub
        End If
    End If
    
    If InStr(1, "12", Val(NVL(rsTemp!ִ��״̬))) > 0 Then
        '1-��ɾ���,2-���ھ���:��Ҫ�ǵڶ��κ���
        'Ӧ����:����Ѿ������,ҽ�������,�в���ȥ����,�ٸ���������
        Exit Sub
    End If
    
    '��������_In Integer := 1
    strSQL = "ZL_���˹Һż�¼_�������� ('" & NVL(rsTemp!NO) & "'," & Val(NVL(rsTemp!����ID)) & ",'" & mstr�������� & "','" & UserInfo.���� & "',to_Date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),0)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
End Sub

Private Function CheckTransferDetail(strID As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------
'����:����ת�ﲡ���Ƿ���������Ϣ
'���:strID-strҵ��ID
'����:True ����ת�ﲡ����������Ϣ False ����ת�ﲡ����������Ϣ
'����:������
'����:2014-02-18
'��ע:
'-----------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHand:
    
    strSQL = "Select ���� From �ŶӽкŶ��� Where ҵ��Id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strID)
    '�ŶӽкŶ���û�м�¼,������
    If rsTemp.EOF Then CheckTransferDetail = True: Exit Function
    If NVL(rsTemp!����) = "" Then CheckTransferDetail = False: Exit Function
    CheckTransferDetail = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Isת�ﲡ��(strҵ��ID As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '����:���ò����Ƿ���ת�ﲡ�˲���δ����
    '���:strҵ��ID
    '����:True ����Ϊת�ﲡ�� False ����Ϊ��ͨ����
    '����:����
    '��������:2012-9-14
    '�����:51514
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHand:
    strSQL = _
    "   Select Count(ID) as �Ƿ�Ϊת�ﲡ�� From ���˹Һż�¼ Where ID=[1] And Nvl(ת�����ID,0) <> 0 And Nvl(ת��״̬,0)=0  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strҵ��ID)
    If rsTemp.EOF Then Isת�ﲡ�� = False
    Isת�ﲡ�� = rsTemp!�Ƿ�Ϊת�ﲡ�� > 0
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mobjQueue_OnRecevieDiagnose(ByVal strҵ��ID As String, ByVal lngҵ������ As Long)
    '����:
    Dim objControl As CommandBarControl
    Dim strNO As String, strSQL As String, rsTmp As ADODB.Recordset
    Dim bln���� As Boolean, arrCheck As Variant, strResult As String
    Dim blnת�ﲡ�� As Boolean '�����:51514
    Dim datCurr As Date
        Dim blnTran As Boolean, colsql As New Collection, i As Long, intOut As Integer, str��� As String
    
    If lngҵ������ <> 0 Then Exit Sub
    On Error GoTo errH
     If Val(strҵ��ID) <> 0 Then
           strSQL = "Select Zl_QueuedateCheck([1]) as Chk From Dual"
           Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strҵ��ID))
           strResult = NVL(rsTmp!chk) & "|"
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
    strSQL = "Select ����ID,ִ����,NO,��¼��־,ִ��״̬,��¼����,����,�����,id as �Һ�id,����,���� From ���˹Һż�¼ Where  ID=[1]  "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strҵ��ID)
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
    strNO = NVL(rsTmp!NO)
    
    'ת����� �����:51514
    blnת�ﲡ�� = Isת�ﲡ��(strҵ��ID)
    If blnת�ﲡ�� Then
        strSQL = "Zl_���˹Һż�¼_ת��('" & strNO & "',1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
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
        If Val(zlDatabase.GetPara("�Һ�ģʽ", glngSys, 9000, 1)) <> 1 And Not mobjSquareCard Is Nothing Then
            If Not mobjSquareCard.zlRegisterIncept(Me, mlngModul, strNO, mstr��������, 0, "") Then Exit Sub
        Else
            strSQL = "Zl_����ԤԼ�Һ�_����('" & strNO & "','" & mstr�������� & "',NULL,NULL,NULL,NULL,NULL,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    Else
        If Val(NVL(rsTmp!ִ��״̬)) = 0 Then
            '�����ҺŽ���
            strSQL = "zl_���˽���(" & Val(NVL(rsTmp!����ID)) & ",'" & strNO & "',Null,'" & UserInfo.���� & "','" & mstr�������� & "',0,0,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        Else
            'Zl_���˽���
            strSQL = "Zl_���˽���("
            '  ����id_In     ������Ϣ.����id%Type,
            strSQL = strSQL & "" & Val(NVL(rsTmp!����ID)) & ","
            '  No_In         ���˹Һż�¼.NO%Type,
            strSQL = strSQL & "'" & strNO & "',"
            '  ִ�в���id_In ���˹Һż�¼.ִ�в���id%Type,
            strSQL = strSQL & "" & IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID) & ","
            '  ִ����_In     ���˹Һż�¼.ִ����%Type,
            strSQL = strSQL & "'" & IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��) & "',"
            '  ����_In       ���˹Һż�¼.����%Type := Null,
            strSQL = strSQL & "'" & mstr�������� & "',"
            '  ��Ǽ���_In   ���˹Һż�¼.����%Type := 0,
            strSQL = strSQL & "0,"
            '  ����_In Integer:=0
            strSQL = strSQL & "1,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            bln���� = True
        End If
        If mbln��Һ�ģʽ Then
            '���ж��ǻ��������ʾ�Ƿ��������
            If Val(rsTmp!ִ��״̬ & "") = 1 Or Val(rsTmp!ִ��״̬ & "") = 2 Then
                str��� = zlCommFun.ShowMsgBox("��ѡ��", "��ǰ����Ϊ���ﲡ�ˣ���ѡ��ò��˾���ģʽ��", _
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
                If mclsReg.zlBulidingPriceDataFromRegistNo(strNO, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID), IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��), , True, colsql) = False Then
                    MsgBox "�ҺŻ��۵�δ��ȷ���ɣ����ܽ��н��", vbInformation, Me.Caption
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
        
    mstr�Һŵ� = strNO
    mlng����ID = Val(NVL(rsTmp!����ID))
    
    '���ﻼ�߽�����Ϣ����
    Call ZLHIS_CIS_009(mclsMipModule, mlng����ID, NVL(rsTmp!����), NVL(rsTmp!�����), 0, 0, NVL(rsTmp!�Һ�ID), NVL(rsTmp!����, 0), NVL(rsTmp!����, 0), datCurr, mlng�������ID, , mstr��������, UserInfo.����)
    
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
        If blnTran Then gcnOracle.RollbackTrans: blnTran = False
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

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    Dim lngPatiID As Long
    
    If objHisPati Is Nothing Then
        lngPatiID = 0
    Else
        lngPatiID = objHisPati.����ID
    End If
    
    Call ExecuteFindPati(False, , blnCard, lngPatiID)
End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnIsInit Then mintFindType = Index: mstrFindType = objCard.����
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
    Dim strSQL As String
    Dim lngidx As Long
    Dim i As Long
    On Error GoTo errH
    strSQL = "Select B.Id,B.NO,B.�����,B.����,B.�Ա�,B.����,A.��������,b.ҽ�Ƹ��ʽ,A.ְҵ," & _
        "   a.���￨��,A.�ѱ�,A.����,A.ҽ����,B.����,A.����ģʽ,B.����ʱ��,B.ִ����,B.ִ��״̬,B.ִ��ʱ��," & _
        "   B.ִ�в���ID as ����ID,B.����,B.����,D.������,C.���� as ����,B.����,B.ժҪ," & _
        "   A.���֤��,A.�໤��,A.��ͥ��ַ,A.��ͥ�绰,A.������λ,A.��ͬ��λid,A.��λ�绰,B.����ʱ��,B.������ַ," & _
        "   A.����,A.����,A.����,A.����״��,A.��ͥ��ַ�ʱ�,A.��λ�ʱ�,A.�����ص�,B.��Ⱦ���ϴ�,A.����֤��,a.���ڵ�ַ," & _
        "   A.���ڵ�ַ�ʱ�,a.����,a.email,a.qq,A.��������,a.����ID,B.·��״̬,nvl(g.����,E.����) as ���� " & _
        " From ������Ϣ A,���˹Һż�¼ B,���ű� C,����������Ϣ D,�ҺŰ��� E,�ٴ������¼ f,�ٴ������Դ g" & _
        " Where A.����ID=B.����ID And B.ID=[1] And B.ִ�в���ID=C.ID" & _
        " And B.����ID=D.����ID(+) And B.����=D.����(+) And B.�ű�=E.����(+) and b.�����¼id=f.id(+) and f.��Դid=g.id(+)"
        '��ID��ȡ�Һż�¼�����üӼ�¼���ʡ�״̬������
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�Һ�id)
    
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
    If Not IsNull(rsTmp!����) And NVL(rsTmp!��������) = "" Then
        txtInfo(txtInfo����).ForeColor = &HC0&
    Else
        txtInfo(txtInfo����).ForeColor = zlDatabase.GetPatiColor(NVL(rsTmp!��������))
    End If
    
    txtInfo(txtInfo�Ա�).Text = rsTmp!�Ա� & ""
    txtInfo(txtInfo�Ա�).ToolTipText = txtInfo(txtInfo�Ա�).Text
    txtInfo(txtInfo����).Text = rsTmp!���� & ""
    txtInfo(txtInfo����).ToolTipText = txtInfo(txtInfo����).Text
    txtInfo(txtInfo��������).Text = Format(rsTmp!�������� & "", "yyyy-MM-dd")
    txtInfo(txtInfo��������).ToolTipText = txtInfo(txtInfo��������).Text
    txtInfo(txtInfo����).Text = rsTmp!���� & ""
    txtInfo(txtInfo����).ToolTipText = txtInfo(txtInfo����).Text
    txtInfo(txtInfo���￨��).Text = rsTmp!���￨�� & ""
    txtInfo(txtInfo���￨��).ToolTipText = txtInfo(txtInfo���￨��).Text
    txtInfo(txtInfoҽ������).Text = rsTmp!ҽ���� & ""
    txtInfo(txtInfoҽ������).ToolTipText = txtInfo(txtInfoҽ������).Text
    txtInfo(txtInfoժҪ).Text = rsTmp!ժҪ & ""
    txtInfo(txtInfoժҪ).ToolTipText = txtInfo(txtInfoժҪ).Text
    
    With cboBillType
        lngidx = -1
        For i = 0 To .ListCount
            If InStr(.List(i) & "", rsTmp!�ѱ� & "") > 0 Then
                .ToolTipText = rsTmp!�ѱ� & ""
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
            mPatiInfo.���� = Decode(NVL(!ִ��״̬, 0), 0, 0, 2, 1, 1, 2)
        End If
        
        mPatiInfo.����� = NVL(!�����)
        mPatiInfo.����ID = !����ID
        mPatiInfo.�Һ�ID = !ID
        mPatiInfo.�Һŵ� = !NO
        mPatiInfo.����ID = !����ID
        mPatiInfo.���� = NVL(!����)
        mPatiInfo.���� = NVL(!����, 0)
        mPatiInfo.������ = NVL(!������)
        mPatiInfo.�Һ�ʱ�� = !����ʱ��
        mPatiInfo.�Ա� = "" & !�Ա�
        mPatiInfo.����״�� = "" & !����״��
        
        mPatiInfo.���� = "" & !����
        mPatiInfo.���� = "" & !����
        mPatiInfo.���� = "" & !����
        mPatiInfo.�����ص� = "" & !�����ص�
        mPatiInfo.��Ⱦ���ϴ� = Val("" & !��Ⱦ���ϴ�)
        mPatiInfo.��ͥ��ַ�ʱ� = "" & !��ͥ��ַ�ʱ�
        mPatiInfo.��λ�ʱ� = "" & !��λ�ʱ�
        mPatiInfo.����֤�� = "" & !����֤��
        mPatiInfo.���ڵ�ַ = "" & !���ڵ�ַ
        mPatiInfo.���ڵ�ַ�ʱ� = "" & !���ڵ�ַ�ʱ�
        mPatiInfo.���� = "" & !����
        mPatiInfo.Email = "" & !Email
        mPatiInfo.QQ = "" & !QQ
        mPatiInfo.���� = Val(!���� & "")
        mPatiInfo.���� = Val(!���� & "")
        mPatiInfo.·��״̬ = Val(!·��״̬ & "")
        lblUrg.Visible = Val(!���� & "") <> 0
        lblRec.Visible = Val(!����ģʽ & "") <> 0
        
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
        UCPatiVitalSigns.TextBackColor = vbWindowBackground
        Call UCPatiVitalSigns.SetUseType(True)
    Else
        Call UCPatiVitalSigns.ClearTxtToolTipText
        UCPatiVitalSigns.ControlLock = True
        UCPatiVitalSigns.TextBackColor = vbButtonFace
        Call UCPatiVitalSigns.SetUseType(False)
    End If
    Call UCPatiVitalSigns.LoadPatiVitalSigns(mPatiInfo.����ID, lng�Һ�id)
    Call UCPatiVitalSigns.TxtAlignment(2)
    txtInfo(txtInfoժҪ).Locked = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadRegist(ByVal lng�Һ�id As Long)
'���ܣ�ѡ��ĳ����ʷ�����¼ʱ����ȡ��صĲ�����Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim lngidx As Long
    Dim i As Long
    
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
    Dim strSQL As String
    Dim rsTmp As Recordset
    
    On Error GoTo errH
    strSQL = "select 1 from ����ҽ����¼ where ����ID=[1] and �Һŵ�=[2] and rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, str�Һŵ�)
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
            Case "·��"
                Set objItem = tbcSub.InsertItem(Index, "�ٴ�·��", mcolSubForm("_·��").hwnd, 0)
                objItem.Tag = "·��"
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
    
    txtInfo(txtInfoժҪ).BackColor = vbButtonFace
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
    
    '������һ�Σ��õĿؼ�����
    For i = 0 To rptPati.Count - 1
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\ReportControl", "rptPati" & "_" & i, rptPati(i).SaveSettings)
    Next
    
    mstrIDCard = ""
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled False
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    Set mobjSquareCard = Nothing

    '--�ر������ŶӵĴ���
    If Not mobjQueue Is Nothing Then
        Call mobjQueue.CloseWindows
        Set mobjQueue = Nothing
    End If
    Set mobjQueueList = Nothing
    Set mobjCallList = Nothing

    'ǿ��Unload,��Ȼ���ἤ���Ӵ�����¼�
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
    mPatiInfo.�Һ�ID = 0
    '�����:57566
    mlng������� = 0
    mlng��ǰ����ʱ�� = 0
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
    
    mblnΣ��ֵ = False
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
'����:���ݴ����б��в��˵��л�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim intCount As Integer
    Dim i As Long, j As Long, k As Long
    Dim strNO As String
    Dim str����ʱ�� As String
    Dim lng�Һ�id As Long
    Dim strCaption As String
    Dim objItem As TabControlItem
    Dim strRegTag As String
    Dim blnDo As Boolean
    Dim str���֤�� As String
    Dim strTmp As String
    Dim str����IDs As String
        
    On Error GoTo errH
    
    If rptPati(Index).SelectedRows.Count <= 0 Then Exit Sub
    For k = PATI_RPT���� To PATI_RPTԤԼ
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
    
    If mstr�Һŵ� = mstrPrePati Then Exit Sub
    mstrPrePati = mstr�Һŵ�
            
    LockWindowUpdate Me.hwnd
    
    '��֤���֤��
    If str���֤�� <> "" Then
        If mobjPatient Is Nothing Then
            On Error Resume Next
            Set mobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
            err.Clear: On Error GoTo 0
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
    
    If str���֤�� <> "" Then
        strSQL = "select a.����id from ������Ϣ a where a.����id<>[1] and a.���֤��=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, str���֤��)
        Do While Not rsTmp.EOF
            str����IDs = str����IDs & "," & rsTmp!����ID
            rsTmp.MoveNext
        Loop
        If str����IDs <> "" Then
            str����IDs = mlng����ID & str����IDs
        End If
    End If
    
    If str����IDs = "" Then
        '��ǰ�ĵ�����ģʽ
        '��ȡ"��ʷ��"�����¼
        strSQL = "Select A.ID,A.NO,A.����ʱ�� as ʱ��,B.���� as ����,a.ִ���� From ���˹Һż�¼ A,���ű� B" & _
            " Where A.ִ�в���ID=B.ID And A.����ID=[1] And A.����ʱ��<=[2] And A.��¼����=1 And A.��¼״̬=1 Order by A.����ʱ�� Desc,a.����ʱ�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, CDate(str����ʱ��))
    Else
        'ͨ�����֤���ҳ�����������
        strSQL = "Select A.ID,A.NO,A.����ʱ�� as ʱ��,B.���� as ����,a.ִ���� From ���˹Һż�¼ A,���ű� B" & _
            " Where A.ִ�в���ID=B.ID And A.����ID In (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X)" & _
            " And A.����ʱ��<=[2] And A.��¼����=1 And A.��¼״̬=1 Order by A.����ʱ�� Desc,a.����ʱ�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����IDs, CDate(str����ʱ��))
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
            lngidx = PATI_RPT����
        Else
            lngidx = PATI_RPTԤԼ
        End If
        With rptPati(lngidx)
            Set objCol = .Columns.Add(COL_HZ_��ʶ, "", 18, True)
            Set objCol = .Columns.Add(COL_HZ_�����, "�����", 60, True)
            Set objCol = .Columns.Add(COL_HZ_����, "����", 60, True)
            Set objCol = .Columns.Add(COL_HZ_�Һ�ʱ��, "�Һ�ʱ��", 80, True)
            Set objCol = .Columns.Add(COL_HZ_�Ա�, "�Ա�", 30, True)
            Set objCol = .Columns.Add(COL_HZ_����, "����", 40, True)
            Set objCol = .Columns.Add(COL_HZ_��, "��", 20, True)
            Set objCol = .Columns.Add(COL_HZ_��, "��", 20, True)
            Set objCol = .Columns.Add(COL_HZ_NO, "�Һŵ�", 60, True)
            Set objCol = .Columns.Add(COL_HZ_����, "����", 30, True)
            Set objCol = .Columns.Add(COL_HZ_��������, "��������", 60, True)
            Set objCol = .Columns.Add(COL_HZ_����ҽ��, "����ҽ��", 60, True)
            Set objCol = .Columns.Add(COL_HZ_���, "���", 60, True)
            Set objCol = .Columns.Add(COL_HZ_����ʱ��, "����ʱ��", 80, True)
            Set objCol = .Columns.Add(COL_HZ_���￨��, "���￨��", 60, True)
            Set objCol = .Columns.Add(COL_HZ_��������, "��������", 60, True)
            Set objCol = .Columns.Add(COL_HZ_ת��״̬, "ת��״̬", 60, True)
            Set objCol = .Columns.Add(COL_HZ_ԤԼҽ��, "ԤԼҽ��", 60, True)
            Set objCol = .Columns.Add(COL_HZ_ԤԼʱ��, "ԤԼʱ��", 80, True)
            Set objCol = .Columns.Add(COL_HZ_���֤��, "���֤��", 60, True)
            Set objCol = .Columns.Add(COL_HZ_����, "����", 30, True)
            Set objCol = .Columns.Add(COL_HZ_���˿���, "���˿���", 60, True)
            
            Set objCol = .Columns.Add(COL_HZ_����ID, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_����ʱ��, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_ִ�в���ID, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_ִ����, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_״̬, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_IC����, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_��¼��־, "", 0, False): objCol.Visible = False
            Set objCol = .Columns.Add(COL_HZ_ִ��״̬, "", 0, False): objCol.Visible = False
            
            
            With .PaintManager
                .ColumnStyle = xtpColumnFlat
                .MaxPreviewLines = 1
                .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
                .GroupForeColor = &HC00000
                .GridLineColor = RGB(225, 225, 225)
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
            Set objCol = .Columns.Add(COL_JZ_�����, "�����", 60, True)
            Set objCol = .Columns.Add(COL_JZ_����, "����", 60, True)
            Set objCol = .Columns.Add(COL_JZ_����ʱ��, "����ʱ��", 80, True)
            Set objCol = .Columns.Add(COL_JZ_�Ա�, "�Ա�", 30, True)
            Set objCol = .Columns.Add(COL_JZ_����, "����", 40, True)
            Set objCol = .Columns.Add(COL_JZ_��, "��", 20, True)
            Set objCol = .Columns.Add(COL_JZ_��, "��", 20, True)
            Set objCol = .Columns.Add(COL_JZ_NO, "�Һŵ�", 60, True)
            Set objCol = .Columns.Add(COL_JZ_����, "����", 30, True)
            Set objCol = .Columns.Add(COL_JZ_���￨��, "���￨��", 60, True)
            Set objCol = .Columns.Add(COL_JZ_��������, "��������", 60, True)
            Set objCol = .Columns.Add(COL_JZ_ת��״̬, "ת��״̬", 60, True)
            Set objCol = .Columns.Add(COL_JZ_��Ⱦ��, "��Ⱦ��", 60, True)
            Set objCol = .Columns.Add(COL_JZ_����, "����", 30, True)
            Set objCol = .Columns.Add(COL_JZ_���˿���, "���˿���", 60, True)
            
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
                .GridLineColor = RGB(225, 225, 225)
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
        Set objCol = .Columns.Add(COL_YZ_NO, "�Һŵ�", 60, True)
        Set objCol = .Columns.Add(COL_YZ_�����, "�����", 60, True)
        Set objCol = .Columns.Add(COL_YZ_����, "����", 60, True)
        Set objCol = .Columns.Add(COL_YZ_�Ա�, "�Ա�", 30, True)
        Set objCol = .Columns.Add(COL_YZ_����, "����", 40, True)
        Set objCol = .Columns.Add(COL_YZ_��, "��", 20, True)
        Set objCol = .Columns.Add(COL_YZ_��, "��", 20, True)
        Set objCol = .Columns.Add(COL_YZ_����, "����", 30, True)
        Set objCol = .Columns.Add(COL_YZ_ʱ��, "ʱ��", 120, True)
        Set objCol = .Columns.Add(COL_YZ_����ҽ��, "����ҽ��", 60, True)
        Set objCol = .Columns.Add(COL_YZ_���￨��, "���￨��", 60, True)
        Set objCol = .Columns.Add(COL_YZ_��������, "��������", 60, True)
        Set objCol = .Columns.Add(COL_YZ_����, "����", 30, True)
        Set objCol = .Columns.Add(COL_YZ_���˿���, "���˿���", 60, True)
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
            .GridLineColor = RGB(225, 225, 225)
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
    
    With rptNotify
        Set objCol = .Columns.Add(c_ͼ��, "", 18, True): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(C_����ID, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_No, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_����, "����", 60, True)
        Set objCol = .Columns.Add(C_�����, "�����", 62, True)
        Set objCol = .Columns.Add(C_����ʱ��, "����ʱ��", 60, True)
        Set objCol = .Columns.Add(C_״̬, "״̬", 150, True)
         
        Set objCol = .Columns.Add(C_��Ϣ, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_���, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_����, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_ҵ��, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_�Һ�Id, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_Id, "", 0, False): objCol.Visible = False
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
            If objCol.Index <> C_��� Or objCol.Index <> C_���� Then objCol.Sortable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .HideSelection = True
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
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
    mvCondFilter.�Һŵ� = ""
    mvCondFilter.���￨ = ""
    mvCondFilter.����ID = 0
    mvCondFilter.����� = ""
    mvCondFilter.���� = ""
    
End Sub

Private Sub GetLocalSetting()
'���ܣ���ע����ȡ��Ժ���˵�ʱ�䷶Χ
    '���ﷶΧ��1=�ұ��˺ŵĲ���,2=�����Ҳ���,3=�����Ҳ���
    Dim strSQL As String, rsTmp As Recordset, intType As Integer
    Dim str���˽������ As String '�����:57566
    
    mint���ﷶΧ = Val(zlDatabase.GetPara("���ﷶΧ", glngSys, p����ҽ��վ, "2"))
    mstr�������� = zlDatabase.GetPara("��������", glngSys, p����ҽ��վ)
    mlng�������ID = Val(zlDatabase.GetPara("�������", glngSys, p����ҽ��վ))
    On Error GoTo errH
    strSQL = "Select Distinct B.ID,B.����,B.����,A.ȱʡ" & _
        " From ������Ա A,���ű� B,��������˵�� C" & _
        " Where A.����ID=B.ID And B.ID=C.����ID And C.������� In(1,3) And C.��������='�ٴ�'" & _
        " And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is Null)" & _
        " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) And A.��ԱID=[1] And b.ID=[2]" & _
        " Order by B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, mlng�������ID)
    If rsTmp.RecordCount = 0 Then mlng�������ID = 0
    mblnҪ����� = Val(zlDatabase.GetPara("ֻ�����Ѿ�����Ĳ���", glngSys, p����ҽ��վ)) <> 0
    
    '���ﲡ��
    If InStr(mstrPrivs, "���ﲡ��") > 0 Then
        mstr����ҽ�� = zlDatabase.GetPara("����ҽ��", glngSys, p����ҽ��վ, UserInfo.����)
    Else
        mstr����ҽ�� = UserInfo.����
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
    mblnPatiDetail = Val(zlDatabase.GetPara("��ʾ������ϸ��Ϣ", glngSys, p����ҽ��վ, 0)) = 1
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
    Dim strSQL As String, strSQLTest As String
    Dim strTime As String
    Dim str��ʶ As String
    Dim strת��״̬ As String
    Dim i As Long, j As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim rsPati As ADODB.Recordset
    Dim intType As Integer '1���2ԤԼ��3ת��
    Dim strTmp As String
    Dim lngColor As Long
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    mblnUnRefresh = True
    
    rptPati(PATI_RPT����).Records.DeleteAll
    
    For intType = 1 To 3
        Select Case intType
        Case 1 '���ﲡ��
            If mint���ﷶΧ = 1 Then
                strSQL = " And B.ִ����||''=[2]" '�ұ��˺�
                If mblnҪ����� Then strSQL = strSQL & " And B.���� is Not NULL"
            ElseIf mint���ﷶΧ = 2 Then
                '������
                If mlng�������ID <> 0 Then
                    strSQL = " And B.����=[3] And b.ִ�в���id+0 =[4] And (B.ִ����||''=[2] Or B.ִ���� Is Null) "
                Else    '10.28��ǰѡ����ʱû�ж�����
                    strSQL = " And B.����=[3] And (B.ִ����||''=[2] Or B.ִ���� Is Null) " & _
                        "And Exists (Select ����id" & vbNewLine & _
                        " From �ҺŰ��� F, ������Ա D" & vbNewLine & _
                        " Where D.��Աid = [6] And F.����id = D.����id And b.ִ�в���id = F.����id)"
                End If
            ElseIf mint���ﷶΧ = 3 Then
                strSQL = " And B.ִ�в���ID+0=[4] And (B.ִ����||''=[2] Or B.ִ���� Is Null)" '������
                If mblnҪ����� Then strSQL = strSQL & " And B.���� is Not NULL"
            End If
            strSQL = " Select /*+ Rule*/B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,B.����,B.����,B.����," & _
                "       B.����ʱ�� as ʱ��,A.���￨��,A.���֤��,A.IC����,A.����,nvl(g.����,E.����) as ����,D.���� as ���˿���," & _
                "       B.����,B.����,B.����ʱ��,B.����ʱ��,B.ִ�в���ID,B.ִ����," & _
                "       B.ת��״̬,C.���� as ת�����,B.ת������,B.ת��ҽ��,B.ִ��״̬,B.��¼��־,A.��������" & _
                " From ������Ϣ A,���˹Һż�¼ B,���ű� C,���ű� D,�ҺŰ��� E, �ٴ������¼ f,�ٴ������Դ g" & _
                " Where B.����ID=A.����ID And (Nvl(B.ִ��״̬,0)=0 or nvl(B.ִ��״̬,0)=[5]) And B.ת�����ID=C.ID(+) and b.�����¼id=f.id(+) and f.��Դid=g.id(+) And B.��¼����=1 And B.��¼״̬=1" & _
                "      And B.�ű�=E.����(+) And B.ִ�в���ID=D.id And B.ִ��ʱ�� is Null And B.����ʱ�� <= Trunc(Sysdate)+1-1/24/60/60 " & strSQL & _
                IIf(gint��ͨ�Һ����� = gint����Һ�����, " And B.����ʱ��>=Sysdate-" & IIf(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����), _
                " And B.����ʱ�� >= Sysdate-" & IIf(gint��ͨ�Һ����� > gint����Һ�����, gint��ͨ�Һ�����, gint����Һ�����) & " And B.����ʱ��>=Sysdate-Decode(B.����,1," & IIf(gint����Һ����� = 0, 1, gint����Һ�����) & "," & IIf(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����) & ")") & _
                " Order By Decode(B.����ʱ��,NULL,2,1),B.����ʱ��,B.NO"
            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "δ��", UserInfo.����, mstr��������, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID), IIf(mblnShowLeavePati, -1, 0), UserInfo.ID)
            
            str��ʶ = " "
        Case 2 'ԤԼ����
            Set rsPati = Nothing
            If mbln��ʾԤԼ���� Then
                If gbln�ҺŰ��� Then
                    '�°�Һų��ﰲ��ģʽ
                    strSQL = "Select B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,B.����,B.����,B.����,B.����ʱ�� as ʱ��,A.���￨��,A.���֤��,A.IC����,A.����,d.����," & vbNewLine & _
                        "e.���� as ���˿���,B.����,B.����,B.����ʱ��,B.����ʱ��,B.ִ�в���ID,B.ִ����,B.ת��״̬,f.���� as ת�����,B.ת������,B.ת��ҽ��,B.ִ��״̬,B.��¼��־,A.��������" & vbNewLine & _
                        "From ������Ϣ A,���˹Һż�¼ B, �ٴ������¼ C,�ٴ������Դ D,���ű� E,���ű� f" & vbNewLine & _
                        "Where B.����ID=A.����ID And  b.�����¼id = c.Id And c.��Դid = d.Id And B.ִ�в���ID=E.ID And B.ת�����ID=f.ID(+) and b.��¼����=2 and b.��¼״̬=1" & vbNewLine & _
                        "And b.����ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 And  Sysdate Between c.��ʼʱ�� And c.��ֹʱ��"
                    If mint���ﷶΧ = 1 Then
                        strSQL = strSQL & " And B.ִ����||''=[1]" '�ұ��˺�
                    ElseIf mint���ﷶΧ = 2 Or mint���ﷶΧ = 3 Then '�����ң�ԤԼ�Һŵķ�ҩ���������Ԥ�ţ�û�����ң���������
                        strSQL = strSQL & " And B.ִ�в���ID+0=[2] And (B.ִ����||''=[1] Or B.ִ���� Is Null)"
                    End If
                Else
                    If mint���ﷶΧ = 1 Then
                        strSQL = " And A.ִ����||''=[1]" '�ұ��˺�
                    ElseIf mint���ﷶΧ = 2 Or mint���ﷶΧ = 3 Then '�����ң�ԤԼ�Һŵķ�ҩ���������Ԥ�ţ�û�����ң���������
                        strSQL = " And A.ִ�в���ID+0=[2] And (A.ִ����||''=[1] Or A.ִ���� Is Null)"
                    End If
                    '�������ڵ�ʱ��Σ��ñ����ӵķ�ʽ�������
                    strTime = _
                        "Select ʱ��� From ʱ��� Where" & _
                        " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
                        " Between" & _
                        " Decode(Sign(��ʼʱ��-��ֹʱ��),1,'3000-01-09 '||To_Char(��ʼʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(��ʼʱ��,'HH24:MI:SS'))" & _
                        " And" & _
                        " '3000-01-10 '||To_Char(��ֹʱ��,'HH24:MI:SS'))" & _
                        " Or" & _
                        " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
                        " Between" & _
                        " '3000-01-10 '||To_Char(��ʼʱ��,'HH24:MI:SS')" & _
                        " And" & _
                        " Decode(Sign(��ʼʱ��-��ֹʱ��),1,'3000-01-11 '||To_Char(��ֹʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(��ֹʱ��,'HH24:MI:SS')))"
                    'ȡ���ڵ���������Ӧ���ŵ�ʱ���
                    strTime = " And Decode(To_Char(SysDate,'D'),'1',B.����,'2',B.��һ,'3',B.�ܶ�,'4',B.����,'5',B.����,'6',B.����,'7',B.����,NULL) IN(" & strTime & ")"
                    strSQL = "Select A.NO,A.����ID,A.��ʶ�� as �����,A.����,A.�Ա�,A.����,A.�Ӱ��־ as ����,A.ִ����,B.����,D.���� as ���˿���," & _
                        " A.����ʱ�� as ʱ��,C.���￨��,C.���֤��,C.IC����,C.����,A.����ʱ��,A.ִ�в���ID,0 as ִ��״̬,0 as ��¼��־,C.��������,null as ת��״̬" & _
                        " From ������ü�¼ A,�ҺŰ��� B,������Ϣ C,���ű� D" & _
                        " Where A.���㵥λ=B.���� And A.����ID=C.����ID(+) And A.ִ�в���ID=D.id And A.���=1" & _
                        " And A.��¼����=4 And A.��¼״̬=0 " & strTime & strSQL & _
                        " And A.����ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate)+1-1/24/60/60"
                End If
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID))
            End If
            str��ʶ = "Ԥ"
        Case 3 'ת�ﲡ��
            If mint���ﷶΧ = 1 Then
                strSQL = " And B.ת��ҽ��=[2]" 'ת���˺�
            ElseIf mint���ﷶΧ = 2 Then
                'ת�����ң���������ת�ģ�����ҽ�������ѻ���δָ������ҽ��
                strSQL = " And B.ת������=[3] And B.ת�����ID=[4] And Nvl(B.ִ����,'��')<>[2] And (B.ת��ҽ��=[2] Or B.ת��ҽ�� Is NULL)"
            ElseIf mint���ﷶΧ = 3 Then
                'ת�����ң���������ת�ģ�����ҽ�������ѻ���δָ������ҽ��
                strSQL = " And B.ת�����ID=[4] And Nvl(B.ִ����,'��')<>[2] And (B.ת��ҽ��=[2] Or B.ת��ҽ�� Is NULL)"
            End If
            strSQL = _
                " Select /*+ Rule*/B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,B.����,B.����,B.����,B.ִ����,nvl(g.����,E.����) as ����,D.���� as ���˿���," & _
                " B.����ʱ�� as ʱ��,A.���￨��,A.���֤��,A.IC����,A.����,B.����ʱ��,B.ת�����ID as ִ�в���ID," & _
                " B.ת��״̬,C.���� as ת�����,B.���� as ת������,B.ִ���� as ת��ҽ��,B.ִ��״̬,B.��¼��־,A.��������" & _
                " From ������Ϣ A,���˹Һż�¼ B,���ű� C,���ű� D,�ҺŰ��� E, �ٴ������¼ f,�ٴ������Դ g" & _
                " Where B.����ID=A.����ID And B.ת��״̬=0 And B.ִ�в���ID=C.ID And B.��¼����=1 And B.��¼״̬=1 And B.�ű�=E.����(+)  and b.�����¼id=f.id(+) and f.��Դid=g.id(+) And B.ת�����ID=D.id " & strSQL & _
                IIf(gint��ͨ�Һ����� = gint����Һ�����, " And B.����ʱ��>=Sysdate-" & IIf(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����), _
                " And B.����ʱ�� >= Sysdate-" & IIf(gint��ͨ�Һ����� > gint����Һ�����, gint��ͨ�Һ�����, gint����Һ�����) & " And B.����ʱ��>=Sysdate-Decode(B.����,1," & IIf(gint����Һ����� = 0, 1, gint����Һ�����) & "," & IIf(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����) & ")") & _
                " Order By B.NO"
            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "δ��", UserInfo.����, mstr��������, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID), 0, 0)
            str��ʶ = "ת"
        End Select
        
        If Not rsPati Is Nothing Then
            For i = 1 To rsPati.RecordCount
                Set objRecord = rptPati(PATI_RPT����).Records.Add()
                For j = 0 To rptPati(PATI_RPT����).Columns.Count - 1
                    objRecord.AddItem ""
                Next
                
                With objRecord
                    .Item(COL_HZ_��ʶ).Value = str��ʶ
                    .Item(COL_HZ_�����).Value = rsPati!����� & ""
                    .Item(COL_HZ_����).Value = rsPati!���� & ""
                    .Item(COL_HZ_�Ա�).Value = rsPati!�Ա� & ""
                    .Item(COL_HZ_����).Value = rsPati!���� & ""
                    .Item(COL_HZ_��).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
                    .Item(COL_HZ_���￨��).Value = rsPati!���￨�� & ""
                    .Item(COL_HZ_��������).Value = rsPati!�������� & ""
                    .Item(COL_HZ_NO).Value = rsPati!NO & ""
                    .Item(COL_HZ_����ID).Value = rsPati!����ID & ""
                    .Item(COL_HZ_����ʱ��).Value = CStr(Format(rsPati!����ʱ��, "yyyy-MM-dd HH:mm:ss"))
                    .Item(COL_HZ_ִ�в���ID).Value = Val(rsPati!ִ�в���ID & "")
                    .Item(COL_HZ_ִ����).Value = rsPati!ִ���� & ""
                    .Item(COL_HZ_״̬).Value = NVL(rsPati!ת��״̬)
                    .Item(COL_HZ_IC����).Value = rsPati!IC���� & ""
                    .Item(COL_HZ_��¼��־).Value = rsPati!��¼��־ & ""
                    .Item(COL_HZ_����).Value = rsPati!���� & ""
                    .Item(COL_HZ_���˿���).Value = rsPati!���˿��� & ""
                    
                    If intType = 1 Then '����
                        .Item(COL_HZ_��������).Value = rsPati!���� & ""
                        .Item(COL_HZ_����ҽ��).Value = rsPati!ִ���� & ""
                        .Item(COL_HZ_���).Value = zlStr.Lpad(NVL(rsPati!����), 5)
                        .Item(COL_HZ_����ʱ��).Value = CStr(Format(rsPati!����ʱ��, "yyyy-MM-dd HH:mm"))
                        .Item(COL_HZ_ִ��״̬).Value = rsPati!ִ��״̬ & ""
                    End If
                    
                    If intType = 1 Or intType = 3 Then '���ת��
                        .Item(COL_HZ_����).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
                        .Item(COL_HZ_��).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
                        .Item(COL_HZ_�Һ�ʱ��).Value = Format(rsPati!ʱ��, "yyyy-MM-dd HH:mm")
                    End If
                    
                    'ת��״̬
                    strת��״̬ = ""
                    If intType = 1 Then
                        If Not IsNull(rsPati!ת��״̬) Then
                            If rsPati!ת��״̬ = 0 Then
                                '�Ѿ�ת��
                                strת��״̬ = "���Է�����,����:" & rsPati!ת����� & _
                                    IIf(Not IsNull(rsPati!ת������), ",����:" & NVL(rsPati!ת������), "") & _
                                    IIf(Not IsNull(rsPati!ת��ҽ��), ",ҽ��:" & NVL(rsPati!ת��ҽ��), "")
                            ElseIf rsPati!ת��״̬ = -1 Then
                                '�Ѿܾ�ת��
                                strת��״̬ = "�Է��Ѿܾ�,����:" & rsPati!ת����� & _
                                    IIf(Not IsNull(rsPati!ת������), ",����:" & NVL(rsPati!ת������), "") & _
                                    IIf(Not IsNull(rsPati!ת��ҽ��), ",ҽ��:" & NVL(rsPati!ת��ҽ��), "")
                            End If
                        End If
                    ElseIf intType = 3 Then
                        'ת�ﲡ��
                        strת��״̬ = "������ת��,����:" & rsPati!ת����� & _
                            IIf(Not IsNull(rsPati!ת������), ",����:" & NVL(rsPati!ת������), "") & _
                            IIf(Not IsNull(rsPati!ת��ҽ��), ",ҽ��:" & NVL(rsPati!ת��ҽ��), "")
                    End If
                    .Item(COL_HZ_ת��״̬).Value = strת��״̬
                    
                    If intType = 2 Then 'ԤԼ
                        .Item(COL_HZ_ԤԼҽ��).Value = rsPati!ִ���� & ""
                        .Item(COL_HZ_ԤԼʱ��).Value = CStr(Format(rsPati!ʱ�� & "", "yyyy-MM-dd HH:mm"))
                    End If
                    .Item(COL_HZ_���֤��).Value = rsPati!���֤�� & ""
                                    
                    '���ղ����ú�ɫ��ʾ
                    If Not IsNull(rsPati!����) And rsPati!�������� & "" = "" Then
                        .Item(COL_HZ_�����).ForeColor = &HC0&
                        .Item(COL_HZ_��������).ForeColor = &HC0&
                    Else
                        '������ɫ
                        lngColor = zlDatabase.GetPatiColor(NVL(rsPati!��������))
                        .Item(COL_HZ_�����).ForeColor = lngColor
                        .Item(COL_HZ_��������).ForeColor = lngColor
                    End If
                    
                    '�����־��ɫͻ����ʾ
                    If NVL(rsPati!����, 0) <> 0 Then
                        .Item(COL_HZ_��).ForeColor = vbRed
                    End If
                    
                    '�����ﲡ�˻�ɫ
                    If Val(rsPati!ִ��״̬ & "") = -1 Then
                        For j = 0 To rptPati(PATI_RPT����).Columns.Count - 1
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
'���ܣ����غ��ﲡ���б�
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
    
    rptPati(PATI_RPTԤԼ).Records.DeleteAll
    

    Set rsPati = Nothing

    If gbln�ҺŰ��� Then
        '�°�Һų��ﰲ��ģʽ
        strSQL = "Select B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,B.����,B.����,B.����,B.����ʱ�� as ʱ��,A.���￨��,A.���֤��,A.IC����,A.����,d.����," & vbNewLine & _
            "e.���� as ���˿���,B.����,B.����,B.����ʱ��,B.����ʱ��,B.ִ�в���ID,B.ִ����,B.ת��״̬,f.���� as ת�����,B.ת������,B.ת��ҽ��,B.ִ��״̬,B.��¼��־,A.��������" & vbNewLine & _
            "From ������Ϣ A,���˹Һż�¼ B, �ٴ������¼ C,�ٴ������Դ D,���ű� E,���ű� f" & vbNewLine & _
            "Where B.����ID=A.����ID And  b.�����¼id = c.Id And c.��Դid = d.Id And B.ִ�в���ID=E.ID And B.ת�����ID=f.ID(+) and b.��¼����=2 and b.��¼״̬=1" & vbNewLine & _
            "And b.����ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 And  Sysdate Between c.��ʼʱ�� And c.��ֹʱ��"
        If mint���ﷶΧ = 1 Then
            strSQL = strSQL & " And B.ִ����||''=[1]" '�ұ��˺�
        ElseIf mint���ﷶΧ = 2 Or mint���ﷶΧ = 3 Then '�����ң�ԤԼ�Һŵķ�ҩ���������Ԥ�ţ�û�����ң���������
            strSQL = strSQL & " And B.ִ�в���ID+0=[2] And (B.ִ����||''=[1] Or B.ִ���� Is Null)"
        End If
    Else
        If mint���ﷶΧ = 1 Then
            strSQL = " And A.ִ����||''=[1]" '�ұ��˺�
        ElseIf mint���ﷶΧ = 2 Or mint���ﷶΧ = 3 Then '�����ң�ԤԼ�Һŵķ�ҩ���������Ԥ�ţ�û�����ң���������
            strSQL = " And A.ִ�в���ID+0=[2] And (A.ִ����||''=[1] Or A.ִ���� Is Null)"
        End If
        '�������ڵ�ʱ��Σ��ñ����ӵķ�ʽ�������
        strTime = _
            "Select ʱ��� From ʱ��� Where" & _
            " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
            " Between" & _
            " Decode(Sign(��ʼʱ��-��ֹʱ��),1,'3000-01-09 '||To_Char(��ʼʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(��ʼʱ��,'HH24:MI:SS'))" & _
            " And" & _
            " '3000-01-10 '||To_Char(��ֹʱ��,'HH24:MI:SS'))" & _
            " Or" & _
            " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
            " Between" & _
            " '3000-01-10 '||To_Char(��ʼʱ��,'HH24:MI:SS')" & _
            " And" & _
            " Decode(Sign(��ʼʱ��-��ֹʱ��),1,'3000-01-11 '||To_Char(��ֹʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(��ֹʱ��,'HH24:MI:SS')))"
        'ȡ���ڵ���������Ӧ���ŵ�ʱ���
        strTime = " And Decode(To_Char(SysDate,'D'),'1',B.����,'2',B.��һ,'3',B.�ܶ�,'4',B.����,'5',B.����,'6',B.����,'7',B.����,NULL) IN(" & strTime & ")"
        strSQL = "Select A.NO,A.����ID,A.��ʶ�� as �����,A.����,A.�Ա�,A.����,A.�Ӱ��־ as ����,A.ִ����,B.����,D.���� as ���˿���," & _
            " A.����ʱ�� as ʱ��,C.���￨��,C.���֤��,C.IC����,C.����,A.����ʱ��,A.ִ�в���ID,0 as ִ��״̬,0 as ��¼��־,C.��������,null as ת��״̬" & _
            " From ������ü�¼ A,�ҺŰ��� B,������Ϣ C,���ű� D" & _
            " Where A.���㵥λ=B.���� And A.����ID=C.����ID(+) And A.ִ�в���ID=D.id And A.���=1" & _
            " And A.��¼����=4 And A.��¼״̬=0 " & strTime & strSQL & _
            " And A.����ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate)+1-1/24/60/60"
    End If
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID))
        
    If Not rsPati Is Nothing Then
        For i = 1 To rsPati.RecordCount
            Set objRecord = rptPati(PATI_RPTԤԼ).Records.Add()
            For j = 0 To rptPati(PATI_RPTԤԼ).Columns.Count - 1
                objRecord.AddItem ""
            Next
            
            With objRecord
                .Item(COL_HZ_��ʶ).Value = "Ԥ"
                .Item(COL_HZ_�����).Value = rsPati!����� & ""
                .Item(COL_HZ_����).Value = rsPati!���� & ""
                .Item(COL_HZ_�Ա�).Value = rsPati!�Ա� & ""
                .Item(COL_HZ_����).Value = rsPati!���� & ""
                .Item(COL_HZ_��).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
                .Item(COL_HZ_���￨��).Value = rsPati!���￨�� & ""
                .Item(COL_HZ_��������).Value = rsPati!�������� & ""
                .Item(COL_HZ_NO).Value = rsPati!NO & ""
                .Item(COL_HZ_����ID).Value = rsPati!����ID & ""
                .Item(COL_HZ_����ʱ��).Value = CStr(Format(rsPati!����ʱ��, "yyyy-MM-dd HH:mm:ss"))
                .Item(COL_HZ_ִ�в���ID).Value = Val(rsPati!ִ�в���ID & "")
                .Item(COL_HZ_ִ����).Value = rsPati!ִ���� & ""
                .Item(COL_HZ_״̬).Value = NVL(rsPati!ת��״̬)
                .Item(COL_HZ_IC����).Value = rsPati!IC���� & ""
                .Item(COL_HZ_��¼��־).Value = rsPati!��¼��־ & ""
                .Item(COL_HZ_����).Value = rsPati!���� & ""
                .Item(COL_HZ_���˿���).Value = rsPati!���˿��� & ""
                .Item(COL_HZ_ת��״̬).Value = ""
                .Item(COL_HZ_ԤԼҽ��).Value = rsPati!ִ���� & ""
                .Item(COL_HZ_ԤԼʱ��).Value = CStr(Format(rsPati!ʱ�� & "", "yyyy-MM-dd HH:mm"))
                .Item(COL_HZ_���֤��).Value = rsPati!���֤�� & ""
                                
                '���ղ����ú�ɫ��ʾ
                If Not IsNull(rsPati!����) And rsPati!�������� & "" = "" Then
                    .Item(COL_HZ_�����).ForeColor = &HC0&
                    .Item(COL_HZ_��������).ForeColor = &HC0&
                Else
                    '������ɫ
                    lngColor = zlDatabase.GetPatiColor(NVL(rsPati!��������))
                    .Item(COL_HZ_�����).ForeColor = lngColor
                    .Item(COL_HZ_��������).ForeColor = lngColor
                End If
                
                '�����־��ɫͻ����ʾ
                If NVL(rsPati!����, 0) <> 0 Then
                    .Item(COL_HZ_��).ForeColor = vbRed
                End If
                
                '�����ﲡ�˻�ɫ
                If Val(rsPati!ִ��״̬ & "") = -1 Then
                    For j = 0 To rptPati(PATI_RPTԤԼ).Columns.Count - 1
                        .Item(j).ForeColor = &H808080
                    Next
                End If
                
            End With
            rsPati.MoveNext
        Next
    End If
    
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
    Dim strSQL As String
    Dim strTime As String
    Dim i As Long, j As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim rsPati As ADODB.Recordset
    Dim strTmp As String
    Dim lngColor As Long
    Dim rs��Ⱦ��״̬ As ADODB.Recordset
    Dim blnDo��Ⱦ��״̬ As Boolean
 
    On Error GoTo errH
    Screen.MousePointer = 11
    mblnUnRefresh = True
    
    strSQL = _
        " Select B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,B.����,B.����,B.����,nvl(g.����,E.����) as ����,D.���� as ���˿���," & _
        " B.ִ��ʱ�� as ʱ��,A.���￨��,A.���֤��,A.IC����,A.����,B.����ʱ��,B.ִ�в���ID,B.ִ����," & _
        " B.ת��״̬,C.���� as ת�����,B.ת������,B.ת��ҽ��,B.ִ��״̬,B.��¼��־,A.��������" & _
        " From ������Ϣ A,���˹Һż�¼ B,���ű� C,���ű� D,�ҺŰ��� E,�ٴ������¼ f,�ٴ������Դ g" & _
        " Where B.����ID=A.����ID And B.ת�����ID=C.ID(+) and B.�ű�=E.����(+) and B.ִ�в���ID=d.id and b.�����¼id=f.id(+) and f.��Դid=g.id(+)" & _
        " And B.ִ��״̬=2 And B.ִ����||''=[1] And B.��¼����=1 And B.��¼״̬=1 and nvl(B.��¼��־,0)<=1" & _
        " Order By B.NO"
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr����ҽ��)
    
    strSQL = "select m.����id,m.id,m.no,max(m.��¼) as ��¼,max(m.��д) as ��д,max(m.״̬) as ״̬ from" & vbNewLine & _
        "(select a.����id,a.id, a.no,1 as ��¼,0 as ��д,0 as ״̬ from ���˹Һż�¼ a,�������Լ�¼ b" & vbNewLine & _
        "where a.no=b.�Һŵ� and a.ִ��״̬=2 And a.ִ����||''=[1] And a.��¼����=1 And a.��¼״̬=1 and nvl(a.��¼��־,0)<=1" & vbNewLine & _
        "union all" & vbNewLine & _
        "Select  a.����id,a.id, a.no,0 as ��¼,1 as ��д,0 as ״̬" & vbNewLine & _
        "From ���˹Һż�¼ A, ���Ӳ�����¼ C, �����ļ��б� D" & vbNewLine & _
        "Where c.�ļ�id = d.Id And d.���� = 5  and c.�������� like '%��Ⱦ��%' And a.����id = c.����id And a.id = c.��ҳid and a.ִ��״̬=2 And a.ִ����||''=[1] And a.��¼����=1 And a.��¼״̬=1 and nvl(a.��¼��־,0)<=1" & vbNewLine & _
        "union all" & vbNewLine & _
        "Select  a.����id,a.id, a.no,0 as ��¼,1 as ��д,e.����״̬ as ״̬" & vbNewLine & _
        "From ���˹Һż�¼ A,���Ӳ�����¼ C,�����ļ��б� D,�����걨��¼ E" & vbNewLine & _
        "Where a.����id = c.����id And a.id = c.��ҳid and c.id=e.�ļ�id and d.����=5 and c.�������� like '%��Ⱦ��%' and e.�ļ�id =d.id and a.ִ��״̬=2 And a.ִ����||''=[1] And a.��¼����=1 And a.��¼״̬=1 and nvl(a.��¼��־,0)<=1) M" & vbNewLine & _
        "group by m.����id,m.id,m.no"
    Set rs��Ⱦ��״̬ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr����ҽ��)
    If rs��Ⱦ��״̬.RecordCount > 0 Then blnDo��Ⱦ��״̬ = True
 
    rptPati(PATI_RPT����).Records.DeleteAll
    For i = 1 To rsPati.RecordCount
        Set objRecord = rptPati(PATI_RPT����).Records.Add()
        For j = 0 To rptPati(PATI_RPT����).Columns.Count - 1
            objRecord.AddItem ""
        Next
        With objRecord
            .Item(COL_JZ_��ʶ).Value = ""
            .Item(COL_JZ_�����).Value = rsPati!����� & ""
            .Item(COL_JZ_����).Value = rsPati!���� & ""
            .Item(COL_JZ_����ʱ��).Value = Format(rsPati!ʱ��, "yyyy-MM-dd HH:mm")
            .Item(COL_JZ_�Ա�).Value = rsPati!�Ա� & ""
            .Item(COL_JZ_����).Value = rsPati!���� & ""
            .Item(COL_JZ_��).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
            .Item(COL_JZ_��).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
            .Item(COL_JZ_NO).Value = rsPati!NO & ""
            .Item(COL_JZ_����).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
            .Item(COL_JZ_���￨��).Value = rsPati!���￨�� & ""
            .Item(COL_JZ_��������).Value = rsPati!�������� & ""
            .Item(COL_JZ_����ID).Value = rsPati!����ID & ""
            .Item(COL_JZ_����ʱ��).Value = CStr(Format(rsPati!����ʱ��, "yyyy-MM-dd HH:mm:ss"))
            .Item(COL_JZ_ִ�в���ID).Value = rsPati!ִ�в���ID & ""
            .Item(COL_JZ_ִ����).Value = rsPati!ִ���� & ""
            .Item(COL_JZ_���֤��).Value = rsPati!���֤�� & ""
            .Item(COL_JZ_IC����).Value = rsPati!IC���� & ""
            .Item(COL_JZ_��¼��־).Value = rsPati!��¼��־ & ""
            .Item(COL_JZ_����).Value = rsPati!���� & ""
            .Item(COL_JZ_���˿���).Value = rsPati!���˿��� & ""
            
            'ת��״̬:��ʾ�����һ��
            .Item(COL_JZ_״̬).Value = NVL(rsPati!ת��״̬)
            If Not IsNull(rsPati!ת��״̬) Then
                If rsPati!ת��״̬ = 0 Then
                    .Item(COL_JZ_ת��״̬).Value = "���Է�����,����:" & rsPati!ת����� & _
                        IIf(Not IsNull(rsPati!ת������), ",����:" & NVL(rsPati!ת������), "") & _
                        IIf(Not IsNull(rsPati!ת��ҽ��), ",ҽ��:" & NVL(rsPati!ת��ҽ��), "")
                ElseIf rsPati!ת��״̬ = -1 Then
                    '�Ѿܾ�ת��
                    .Item(COL_JZ_ת��״̬).Value = "�Է��Ѿܾ�,����:" & rsPati!ת����� & _
                        IIf(Not IsNull(rsPati!ת������), ",����:" & NVL(rsPati!ת������), "") & _
                        IIf(Not IsNull(rsPati!ת��ҽ��), ",ҽ��:" & NVL(rsPati!ת��ҽ��), "")
                End If
            End If
            
            '���ղ����ú�ɫ��ʾ
            If Not IsNull(rsPati!����) And rsPati!�������� & "" = "" Then
                .Item(COL_JZ_�����).ForeColor = &HC0&
                .Item(COL_JZ_��������).ForeColor = &HC0&
            Else
                '������ɫ
                lngColor = zlDatabase.GetPatiColor(NVL(rsPati!��������))
                .Item(COL_JZ_�����).ForeColor = lngColor
                .Item(COL_JZ_��������).ForeColor = lngColor
            End If
            
            '�����־��ɫͻ����ʾ
            If NVL(rsPati!����, 0) <> 0 Then
                .Item(COL_JZ_��).ForeColor = vbRed
            End If
            
            '��Ӵ�Ⱦ��״̬
            strSQL = ""
            If blnDo��Ⱦ��״̬ Then
                rs��Ⱦ��״̬.Filter = "no='" & rsPati!NO & "'"
                If Not rs��Ⱦ��״̬.EOF Then strSQL = Get��Ⱦ��״̬(Val(rs��Ⱦ��״̬!��¼ & ""), Val(rs��Ⱦ��״̬!��д & ""), Val(rs��Ⱦ��״̬!״̬ & ""))
            End If
            .Item(COL_JZ_��Ⱦ��).Value = strSQL
        End With
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
    Dim strSQL As String
    Dim strTime As String
    Dim i As Long, j As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim rsPati As ADODB.Recordset
    Dim lngColor As Long
    Dim bln��ҽ As Boolean
    
    On Error GoTo errH
    Screen.MousePointer = 11
    mblnUnRefresh = True
    rptPati(PATI_RPT����).Records.DeleteAll
    
    strSQL = "Select /*+ Rule*/" & vbNewLine & _
        " Distinct(b.No), b.����id, b.�����, b.����, b.�Ա�, b.����, b.����, b.����, b.����, b.ִ��ʱ�� As ʱ��, a.���￨��, a.���֤��, a.Ic����, a.����, b.����ʱ��, b.ִ�в���id," & vbNewLine & _
        " b.ִ����, b.ִ��״̬, b.��¼��־, a.��������,nvl(i.����,E.����) as ����,F.���� as ���˿���," & vbNewLine & _
        "First_Value(Decode(Sign(h.������� - 10), -1, h.�������, '')) Over(Partition By h.����id, h.��ҳid Order By Sign(h.������� - 10), Decode(h.��¼��Դ, 4, 0, h.��¼��Դ) Desc, Decode(h.�������, 1, 1, 0) Desc, h.��ϴ���) As ��ҽ���," & vbNewLine & _
        "First_Value(Decode(Sign(h.������� - 10), 1, h.�������, '')) Over(Partition By h.����id, h.��ҳid Order By -Sign(h.������� - 10), Decode(h.��¼��Դ, 4, 0, h.��¼��Դ) Desc, Decode(h.�������,11,11, 0) Desc, h.��ϴ���) As ��ҽ���" & vbNewLine & _
        "From ������Ϣ A, ���˹Һż�¼ B" & IIf(mvCondFilter.���￨ <> "", ",����ҽ�ƿ���Ϣ C, ҽ�ƿ���� D", "") & ",�ҺŰ��� E,���ű� F, ������ϼ�¼ H, �ٴ������¼ g,�ٴ������Դ I" & vbNewLine & _
        "Where b.����id = a.����id And h.����id(+) = b.����id And h.��ҳid(+) = b.id And b.ִ��״̬ + 0 = 1 And b.��¼���� = 1 And b.��¼״̬ = 1 and B.�ű�=E.����(+) and b.ִ�в���id=f.id and b.�����¼id=g.id(+) and g.��Դid=i.id(+)" & _
         IIf(mvCondFilter.���￨ <> "", " And c.����id = a.����id And c.�����id = d.Id And d.�Ƿ�̶� = 1 And d.���� = '���￨' ", "")

    If mvCondFilter.�Һŵ� <> "" Then
        strSQL = strSQL & " And B.NO=[5]"
    ElseIf mvCondFilter.����� <> "" Then
        strSQL = strSQL & " And A.�����=[6]"
    ElseIf mvCondFilter.���￨ <> "" Then
        strSQL = strSQL & " And C.����=[7]"
    
    Else
        strSQL = strSQL & " And B.ִ��ʱ�� Between To_Date('" & Format(mvCondFilter.Begin, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(mvCondFilter.End, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        strSQL = strSQL & IIf(mvCondFilter.ҽ�� = "", "", " And B.ִ����||''=[3]")
        If mvCondFilter.����ID <> 0 Then strSQL = strSQL & " And B.ִ�в���ID+0=[4]"
                If mvCondFilter.���� <> "" Then strSQL = strSQL & " And A.����=[8]"
    End If
    
    If zlDatabase.DateMoved(mvCondFilter.Begin) Then
        strSQL = strSQL & " Union ALL " & Replace(strSQL, "���˹Һż�¼", "H���˹Һż�¼")
    End If

    strSQL = strSQL & " Order By NO Desc"
    
    With mvCondFilter
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "δ��", "δ��", .ҽ��, .����ID, .�Һŵ�, .�����, .���￨, .����)
    End With
    
    For i = 1 To rsPati.RecordCount
        Set objRecord = rptPati(PATI_RPT����).Records.Add()
        For j = 0 To rptPati(PATI_RPT����).Columns.Count - 1
            objRecord.AddItem ""
        Next
        
        With objRecord
            .Item(COL_YZ_�����).Value = rsPati!����� & ""
            .Item(COL_YZ_����).Value = rsPati!���� & ""
            .Item(COL_YZ_�Ա�).Value = rsPati!�Ա� & ""
            .Item(COL_YZ_����).Value = rsPati!���� & ""
            .Item(COL_YZ_��).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
            .Item(COL_YZ_��).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
            .Item(COL_YZ_����).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
            .Item(COL_YZ_ʱ��).Value = CStr(Format(rsPati!ʱ�� & "", "yyyy-MM-dd HH:mm"))
            .Item(COL_YZ_����ҽ��).Value = rsPati!ִ���� & ""
            .Item(COL_YZ_���￨��).Value = rsPati!���￨�� & ""
            .Item(COL_YZ_��������).Value = rsPati!�������� & ""
            .Item(COL_YZ_����).Value = rsPati!���� & ""
            .Item(COL_YZ_���˿���).Value = rsPati!���˿��� & ""
            .Item(COL_YZ_NO).Value = rsPati!NO & ""
            .Item(COL_YZ_����ID).Value = rsPati!����ID & ""
            .Item(COL_YZ_����ʱ��).Value = CStr(Format(rsPati!����ʱ��, "yyyy-MM-dd HH:mm:ss"))
            .Item(COL_YZ_ִ�в���ID).Value = Val(rsPati!ִ�в���ID & "")
            .Item(COL_YZ_ִ����).Value = rsPati!ִ���� & ""
            .Item(COL_YZ_���֤��).Value = rsPati!���֤�� & ""
            .Item(COL_YZ_IC����).Value = rsPati!IC���� & ""
            .Item(COL_YZ_��¼��־).Value = rsPati!��¼��־ & ""
            .Item(COL_YZ_��ҽ���).Value = rsPati!��ҽ��� & ""
            .Item(COL_YZ_��ҽ���).Value = rsPati!��ҽ��� & ""
            If rsPati!��ҽ��� & "" <> "" Then bln��ҽ = True
            
            '���ղ����ú�ɫ��ʾ
            If Not IsNull(rsPati!����) And rsPati!�������� & "" = "" Then
                .Item(COL_YZ_�����).ForeColor = &HC0&
                .Item(COL_YZ_��������).ForeColor = &HC0&
            Else
                '������ɫ
                lngColor = zlDatabase.GetPatiColor(NVL(rsPati!��������))
                .Item(COL_YZ_�����).ForeColor = lngColor
                .Item(COL_YZ_��������).ForeColor = lngColor
            End If
            
            '�����־��ɫͻ����ʾ
            If NVL(rsPati!����, 0) <> 0 Then
                .Item(COL_YZ_��).ForeColor = vbRed
            End If
            
        End With
        rsPati.MoveNext
    Next
    
    rptPati(PATI_RPT����).Columns(COL_YZ_��ҽ���).Visible = bln��ҽ
    rptPati(PATI_RPT����).Populate
    i = rptPati(PATI_RPT����).Records.Count
    tbcInTreat.Item(t���).Caption = "���" & IIf(i = 0, "", ":" & i & "��")
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
'���ܣ���ȡ�����б�
'������strActNO=ˢ�º���Ҫ��λ���б������Ͳ��˹Һŵ�(�����)
'      ע���������ָ����intActive,�����Ҫ����strRefeshˢ���б���
'      strRefesh=�ֱ��Ƿ�ˢ��ָ�����б��ֱ�Ϊ ��1λ��"����/ת��/ԤԼ"����2λ��"����"����3λ��"����"����4λ-"����"����5λ-"ԤԼ"
    Dim strPrePati As String
    Dim i As Long, j As Long
    Dim blnFinded As Boolean
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
        For lngCol = 0 To 3
            If rptPati(lngCol).SelectedRows.Count > 0 Then
                rptPati(lngCol).SelectedRows(0).Selected = False
            End If
        Next
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
    
    imgPatient.Picture = imgDefual.Picture

    txtInfo(txtInfo����).Text = ""
    txtInfo(txtInfo�Ա�).Text = ""
    txtInfo(txtInfo����).Text = ""
    txtInfo(txtInfo��������).Text = ""
    txtInfo(txtInfo���￨��).Text = ""
    txtInfo(txtInfoҽ������).Text = ""
    txtInfo(txtInfoժҪ).Text = ""
    txtInfo(txtInfoժҪ).ToolTipText = ""
    
    lblMore.Visible = False
    lblRec.Visible = False
    lblUrg.Visible = False
    
    cboPayType.ListIndex = -1
    cboBillType.ListIndex = -1
    
    For i = 0 To lblLink�޸�
        lblLink(i).ForeColor = &HC0C0C0
    Next
    mPr = -1
End Sub

Private Sub ExecuteRegist(ByVal strNO As String)
'���ܣ����˹Һ�
    Dim objControl As CommandBarControl
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
        If err <> 0 Then
            err = 0: On Error GoTo 0
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
    Dim strSQL As String
     
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
            strSQL = "Select ID From ����ҽ����¼ Where ����ID+0=[1] And �Һŵ�=[2] And ҽ��״̬=1  And NVL(ִ�б��,0) <> -1 And Nvl(ִ������,0)<>0 And Rownum = 1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
            If Not rsTmp.EOF Then
                MsgBox "�ò��˻���δ����ҽ����ֻ�н�����ҽ�����ͺ���ܽ���ת�", vbInformation, gstrSysName
                Exit Sub
            End If
        Else    'ֻҪ�¹�ҽ��(���������ϵ�)��˵��������Ϊ�ѷ�����������ת������¹Һ�
            strSQL = "Select ID From ����ҽ����¼ Where ����ID+0=[1] And �Һŵ�=[2] And ҽ��״̬ <> 4 And Rownum = 1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
            If Not rsTmp.EOF Then
                MsgBox "�Ѿ��Ըò����¹�ҽ����������ת���ɾ��������ҽ�����ٽ��У��������¹Һš�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    If Not frmRegistPlan.ShowMe(Me, mstr�Һŵ�, lng����ID, str����, strҽ��, lngҽ��ID) Then mblnUnRefresh = False: Exit Sub
    
    'ִ��ת��
    strSQL = "Zl_���˹Һż�¼_ת��('" & mstr�Һŵ� & "',0," & lng����ID & ",'" & str���� & "','" & strҽ�� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    '���ﻼ��ת����Ϣ����
    Call ZLHIS_CIS_007(mclsMipModule, mlng����ID, Trim(txtInfo(txtInfo����).Text), mPatiInfo.�����, mlng�Һ�ID, mlng�������ID, , lng����ID, , lngҽ��ID, strҽ��, str����, UserInfo.����)
    
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
    Dim strSQL As String, rsTemp As ADODB.Recordset
    If Check�Ŷӽк� = False Then Exit Sub
    '95637:���ϴ�,2016/7/20,������Ч���в���ʾ
    strSQL = "Select �ŶӺ��� From �ŶӽкŶ��� Where ҵ������=0 And �Ŷ�״̬ In (0,1,7) and ҵ��ID in (Select ID From ���˹Һż�¼ where NO=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Sub
    MsgBox "ע��:" & vbCrLf & "    �ò������½������ŶӴ���,�Ӻ�Ϊ:[ " & NVL(rsTemp!�ŶӺ���) & " ]", vbInformation + vbOKOnly, gstrSysName
End Sub

Private Sub ExecuteTransferRefuse()
'���ܣ�ת��ܾ�
    Dim strSQL As String
        
    On Error GoTo errH
    
    If mPr <> -1 Then
        If MsgBox("ȷʵҪ�ܾ���ת�ﲡ��""" & rptPati(mintRPTIndex).Rows(mPr).Record(COL_JZ_����).Value & """��", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        strSQL = "Zl_���˹Һż�¼_ת��('" & rptPati(mintRPTIndex).Rows(mPr).Record(COL_JZ_NO).Value & "',-1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
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
    Dim strSQL As String
    
    On Error GoTo errH
 
    With rptPati(mintRPTIndex).Rows(mPr)
        If blnMsg Then
            If MsgBox("ȷʵҪȡ������""" & .Record(COL_JZ_����).Value & """��ת����", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        End If
        strSQL = "Zl_���˹Һż�¼_ת��('" & .Record(COL_JZ_NO).Value & "',Null)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
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
    Dim strSQL As String
    
    On Error GoTo errH
    
    With rptPati(mintRPTIndex).Rows(mPr)
        If MsgBox(.Record(COL_JZ_ת��״̬).Value & vbCrLf & vbCrLf & "ȷ�Ͻ��ո�ת�ﲡ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        strSQL = "Zl_���˹Һż�¼_ת��('" & .Record(COL_JZ_NO).Value & "',1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
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
    Dim strSQL As String, strNO As String
    Dim blnReserve As Boolean
    Dim datCurr As Date
   
    On Error GoTo errH

    datCurr = zlDatabase.Currentdate
    
    If (mintRPTIndex = PATI_RPT���� Or mintRPTIndex = PATI_RPTԤԼ) And mPr <> -1 Then
        If rptPati(mintRPTIndex).Rows(mPr).Record(COL_HZ_��ʶ).Value = "Ԥ" Then
            blnReserve = True
        End If
    Else
        Exit Sub
    End If
    
    If blnReserve Then
        '��ԤԼ�ҺŲ��˽��н���
        '�����:57566
        If Check�������("����", mstr�Һŵ�) = False Then Exit Sub
        
        '����ҽ��վԤԼ����ʱ���ùҺŲ����Ľ��սӿڽ��п۷ѵĹ���
        If Val(zlDatabase.GetPara("�Һ�ģʽ", glngSys, 9000, 1)) <> 1 And Not mobjSquareCard Is Nothing Then
            If Not mobjSquareCard.zlRegisterIncept(Me, mlngModul, mstr�Һŵ�, mstr��������, PatiIdentify.objIDKind.GetCurCard.�ӿ����, PatiIdentify.Text) Then Exit Sub
        Else
            strSQL = "Zl_����ԤԼ�Һ�_����('" & mstr�Һŵ� & "','" & mstr�������� & "',NULL,NULL,NULL,NULL,NULL,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
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
        strSQL = "Select ִ���� From ���˹Һż�¼ Where ����ID+0=[1] And NO=[2] And Nvl(ִ��״̬,0)<>0 And ��¼����=1 And ��¼״̬=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
        If Not rsTmp.EOF Then
            MsgBox "�ò�������" & IIf(IsNull(rsTmp!ִ����), "����ҽ��", "ҽ����" & rsTmp!ִ���� & " ") & "���", vbInformation, gstrSysName
            Call LoadPatients("10001"): Exit Sub
        End If
        
        strSQL = "Select ִ���� From ���˹Һż�¼ Where ����ID+0=[1] And NO=[2] And Nvl(ִ��״̬,0)=0 And ��¼����=1 And ��¼״̬=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
        If rsTmp.EOF Then
            MsgBox "�ò������˺ţ����ܽ��", vbInformation, gstrSysName
            Call LoadPatients("10001"): Exit Sub
        End If
        
        strSQL = "zl_���˽���(" & mlng����ID & ",'" & mstr�Һŵ� & "',Null,'" & UserInfo.���� & "','" & mstr�������� & "',0,0,To_Date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    End If
    
    If mblnAutoHandle Then Call Tip�����Զ����
    
    'ˢ�²���λ����
    On Error GoTo 0
    
    If rptPati(PATI_RPT����).Visible Then
        tbcInTreat.Item(t����).Selected = True
        Call LoadPatients("11001", PATI_RPT����, mstr�Һŵ�)
    Else
        tbcInTreat.Item(t����).Selected = True
    End If
    
    '���ﻼ�߽�����Ϣ����
    Call ZLHIS_CIS_009(mclsMipModule, mlng����ID, Trim(txtInfo(txtInfo����).Text), mPatiInfo.�����, 0, 0, mlng�Һ�ID, mPatiInfo.����, mPatiInfo.����, datCurr, mlng�������ID, , mstr��������, UserInfo.����)

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
        Call zlPlugInErrH(err, "ClinicReceive")
        err.Clear: On Error GoTo errH
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
    Dim strSQL As String
        Dim blnTran As Boolean, colsql As New Collection, i As Long, bytOut As Byte
        
    If BillExpend(mstr�Һŵ�) Then
        MsgBox "�ò��˹Һ��ѳ�����Ч��������������ȡ�����", vbInformation, gstrSysName
        Exit Sub
    End If
        
    On Error GoTo errH
    
    'ֻ��ȡ���Լ�����Ĳ���
    strSQL = "Select ִ���� From ���˹Һż�¼ Where id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mPatiInfo.�Һ�ID)
    If rsTmp!ִ���� <> UserInfo.���� Then
        MsgBox "ֻ��ȡ���Լ�����Ĳ��ˡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'ToDo:ȡ������ʱ�������ݵļ��
    'ҽ�����ݵļ��
    strSQL = "Select Count(*) as ҽ�� From ����ҽ����¼ Where ҽ��״̬ IN(1,8) And ����ID+0=[1] And �Һŵ�=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
    If NVL(rsTmp!ҽ��, 0) > 0 Then
        MsgBox "�ò��������¿����ѷ��͵�ҽ��������ȡ�����" & vbCrLf & _
            "���ȷʵҪȡ��������Ƚ���Щҽ��ɾ�������ϡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mbln��Һ�ģʽ Then
        If mclsReg.zlRegisterPriceDeleteFromNO(mstr�Һŵ�, IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID), IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��), True, bytOut, colsql) = False Then
            'bytOut:��������:0-����ִ��;1-δ�ҵ��Һŵ�;2-δ���ɻ��۵�;3-δ�ҵ����������Ļ��۵�;4-�����Ѿ��շѵĵ���
            '�����Ƿ񽻷ѣ�����ȡ���������Ѿ��ɷѣ�����ò�������ȥ�˷�
        End If
    End If
    gcnOracle.BeginTrans: blnTran = True
    
    strSQL = "Zl_���˽���_Cancel(" & mlng����ID & ",'" & mstr�Һŵ� & "'," & IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID) & ",'" & IIf(mstr����ҽ�� = "", UserInfo.����, mstr����ҽ��) & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    For i = 1 To colsql.Count
        Call zlDatabase.ExecuteProcedure(colsql(i), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTran = False
    
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
        If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteFinish()
'���ܣ���ɽ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, blnTran As Boolean
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
    strSQL = "select 1 from ���˹Һż�¼ where no=[1] and ִ����=[2] And ִ��״̬=2 And ��¼����=1 And ��¼״̬=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�, mstr����ҽ��)
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
    strSQL = "Select Count(*) as ҽ�� From ����ҽ����¼ Where ����ID+0=[1] And �Һŵ�=[2] And ҽ��״̬<>4"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
    If NVL(rsTmp!ҽ��, 0) = 0 Then
        If MsgBox("δ��""" & str���� & """�´��κ���Ч��ҽ����ȷʵҪ��ɽ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    '����Ƿ����δ���͵�ҽ��
    strSQL = "Select Count(*) as ҽ�� From ����ҽ����¼ Where ����ID+0=[1] And �Һŵ�=[2] And ҽ��״̬=1 And NVL(ִ�б��,0) <> -1 And Nvl(ִ������,0)<>0 And Nvl(Ƥ�Խ��,'��')<>'����'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
    If NVL(rsTmp!ҽ��, 0) > 0 Then
        MsgBox """" & str���� & """����δ���͵�ҽ����������ɽ��", vbInformation, gstrSysName
        Exit Sub
    End If
    '���δ��д�ļ���֤������
    strSQL = "Select ��ҳID,����ID,���ID From ������ϼ�¼ Where ȡ��ʱ�� is Null And ����ID=[1] And ��ҳID=(Select ID From ���˹Һż�¼ Where NO=[2] And ��¼����=1 And ��¼״̬=1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
    Do While Not rsTmp.EOF
        If lng�Һ�id = 0 Then lng�Һ�id = rsTmp!��ҳID
        If Not IsNull(rsTmp!����id) Then str����IDs = str����IDs & "," & rsTmp!����id
        If Not IsNull(rsTmp!���id) Then str���IDs = str���IDs & "," & rsTmp!���id
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
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim blnTran As Boolean
    
    On Error GoTo errH
    
    '��ȡ��Ҫ����Ϣ�������ӿڵ���:����߾��ﲡ�˱��ξ���Ϊ׼,�ұ߿��ܵ�ǰѡ�����ʷ����
    strSQL = "Select A.ID,A.����,B.������ From ���˹Һż�¼ A,����������Ϣ B Where A.����ID=B.����ID(+) And A.��¼����=1 And A.��¼״̬=1 And A.����=B.����(+) And A.NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    
    'ִ�й���
    '-----------------------------------
    gcnOracle.BeginTrans: blnTran = True
    
    strSQL = "Zl_���˽������(" & lng����ID & ",'" & strNO & "','" & mstr�������� & "','" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
         
    If Not gobjCommunity Is Nothing And NVL(rsTmp!����, 0) <> 0 Then
        '��������������Ϣ�ύ
        If Not gobjCommunity.ClinicSubmit(glngSys, mlngModul, rsTmp!����, NVL(rsTmp!������), lng����ID, rsTmp!ID) Then
            gcnOracle.RollbackTrans: blnTran = False: Exit Function
        End If
    End If
    gcnOracle.CommitTrans: blnTran = False

    '����������ҽӿ�
    Call CreatePlugInOK(p����ҽ��վ)
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.ClinicFinish(glngSys, p����ҽ��վ, lng����ID, lng�Һ�id)
        Call zlPlugInErrH(err, "ClinicFinish")
        err.Clear: On Error GoTo errH
    End If
    
    'һ��ͨ�����ϴ�
    If Not mobjICCard Is Nothing Then
        strSQL = "Select 1 From һ��ͨĿ¼ Where ����=2 And Rownum=1"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            mobjICCard.UploadSwap lng����ID, ""
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
'�ָ�����
    Dim strSQL As String
    
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
    strSQL = "zl_���˽������_Cancel(" & mlng����ID & ",'" & mstr�Һŵ� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
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
    Dim arrSQL As Variant, i As Long
    Dim colInfo As New Collection
    Dim int���� As Integer, str������ As String
    Dim str�������� As String
        
    If gobjCommunity Is Nothing Or mPatiInfo.����ID = 0 Or mPatiInfo.�Һ�ID = 0 Or mPatiInfo.���� <> 0 Then Exit Sub
    
    If Not gobjCommunity.Identify(glngSys, p����ҽ��վ, int����, str������, colInfo, mPatiInfo.����ID, mPatiInfo.�Һ�ID) Then Exit Sub
    
    arrSQL = Array()
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_����������Ϣ_Insert(" & mPatiInfo.����ID & "," & int���� & ",'" & str������ & "',1,Sysdate)"
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    str�������� = GetColItem(colInfo, "��������")
    If IsDate(str��������) Then
        str�������� = "To_Date('" & Format(str��������, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
    Else
        str�������� = "Null"
    End If
    arrSQL(UBound(arrSQL)) = "Zl_���˹Һż�¼_������֤(" & mPatiInfo.����ID & "," & mPatiInfo.�Һ�ID & "," & int���� & "," & _
        "'" & GetColItem(colInfo, "����") & "','" & GetColItem(colInfo, "�Ա�") & "','" & GetColItem(colInfo, "����") & "'," & _
        str�������� & ",'" & GetColItem(colInfo, "�����ص�") & "','" & GetColItem(colInfo, "���֤��") & "'," & _
        "'" & GetColItem(colInfo, "����") & "','" & GetColItem(colInfo, "����") & "','" & GetColItem(colInfo, "����״��") & "'," & _
        "'" & GetColItem(colInfo, "ְҵ") & "','" & GetColItem(colInfo, "��ͥ��ַ") & "','" & GetColItem(colInfo, "��ͥ�绰") & "'," & _
        "'" & GetColItem(colInfo, "��ͥ��ַ�ʱ�") & "','" & GetColItem(colInfo, "������λ") & "','" & GetColItem(colInfo, "��λ�绰") & "'," & _
        "'" & GetColItem(colInfo, "��λ�ʱ�") & "','" & GetColItem(colInfo, "��ϵ������") & "','" & GetColItem(colInfo, "��ϵ�˹�ϵ") & "'," & _
        "'" & GetColItem(colInfo, "��ϵ�˵绰") & "','" & GetColItem(colInfo, "��ϵ�˵�ַ") & "','" & GetColItem(colInfo, "���ڵ�ַ") & "','" & GetColItem(colInfo, "���ڵ�ַ�ʱ�") & "')"
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), "ExecuteCommunityIdentify"
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
    Call LoadPatients����
    If Not mbln��ʾԤԼ���� Then
        Call LoadPatientsԤԼ
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
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If blnDo Then
        strSQL = "select count(1) as ���� from ���˹Һż�¼ a where a.��¼״̬=1 and a.ִ����=[1] and  a.ִ��ʱ�� between Trunc(Sysdate) and Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����)
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
    
    If Not mclsMipModule Is Nothing Then
        If mclsMipModule.IsConnect Then 'ʹ������Ϣƽ̨���µ�ˢ�²���
            lngSecond = lngSecond + 1
            If lngSecond Mod 180 = 0 Then
                lngSecond = 0
                Call RefeshByMsg
            End If
            Exit Sub
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
            If mblnΣ��ֵ���� Then Call ReadMsgAuto
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
    
     '���Ҳ���
    If lngPatiID = 0 And Not mobjSquareCard Is Nothing And mstrFindType <> "���￨" And mstrFindType <> "��ʶ��" And mstrFindType <> "�Һŵ�" And mstrFindType <> "����" And mstrFindType <> "�������֤" Then
        If mstrFindType = "IC��" Then
            Call mobjSquareCard.zlGetPatiID("IC��", PatiIdentify.Text, , lngPatiID)
        Else
            Call mobjSquareCard.zlGetPatiID(Val(PatiIdentify.objIDKind.GetCurCard.�ӿ����), PatiIdentify.Text, , lngPatiID)
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
                    tbcInTreat.Item(t���).Selected = True
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
    
    err = 0: On Error GoTo GOEND:
    If mobjQueue Is Nothing Then
        Set mobjQueue = CreateObject("zlQueueManage.clsQueueManage")
        err = 0: On Error GoTo ErrHand:
        mobjQueue.zlInitVar gcnOracle, glngSys, 0, IIf(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����), mty_Queue.strQueuePrivs, CStr(mlngModul), False
        mobjQueue.zlSetToolIcon 24, True
        mobjQueue.IsShowFindTools = False
    End If
    Check�Ŷӽк� = True
    Exit Function
ErrHand:
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
    Dim varQueue() As String, strTemp As String, rsTemp As ADODB.Recordset, strSQL As String
    Dim str���� As String, strҽ�� As String, str���� As String
    Dim intType As Integer
    
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
    strSQL = "" & _
    "   Select distinct  /*+ Rule*/  c.ҵ��ID From ���˹Һż�¼ A ,�ŶӽкŶ���  C" & _
    "   Where A.id=C.ҵ��ID and C.��������=[1]  and nvl(C.ҵ������,0)=0 and nvl(A.����ID,0) =0 And a.��¼����=1 And a.��¼״̬=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����)
    With rsTemp
        strTemp = ""
        Do While Not .EOF
            strTemp = strTemp & "," & Val(NVL(rsTemp!ҵ��ID))
            .MoveNext
        Loop
        If strTemp <> "" Then strTemp = "0|" & Mid(strTemp, 2)
    End With
    Call mobjQueue.zlRefresh(strQueue, mty_Queue.strCurrQueueName, mty_Queue.lngcurr�Һ�ID, str����, strҽ��, strTemp, intType)
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
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim i As Byte
    If Check�Ŷӽк� = False Then Exit Sub
    
    strSQL = "SELECT ID,ִ�в���ID,����,ִ���� From ���˹Һż�¼ where NO=[1] And ��¼����=1 And ��¼״̬=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Sub
    
    strQueueName = NVL(rsTemp!ִ�в���ID)
    If NVL(rsTemp!ִ����) <> "" Then
        strQueueName = strQueueName & ":" & NVL(rsTemp!ִ����)
    ElseIf NVL(rsTemp!����) <> "" Then
        strQueueName = strQueueName & ":" & NVL(rsTemp!����)
    End If
    
    lngID = Val(NVL(rsTemp!ID))
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
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim str����NO As String
    
    If mstr�Һŵ� = "" Then Exit Function
    
    On Error GoTo errH
    
    If lngState = -1 Then
        '��鲡���Ƿ������Ч��ҽ��
        strSQL = "Select 1 From ����ҽ����¼ Where ����id = [1] And �Һŵ� = [2]  And ҽ��״̬ <> -1 And ҽ��״̬ <> 1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
        If Not rsTmp.EOF Then
            MsgBox "�ò��˴�����Чҽ��,��������Ϊ������!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    End If

    '��ȡ�ҺŻ��۵���Ϣ
    strSQL = "Select ժҪ From ������ü�¼ Where NO = [1] And ��¼���� = 4 And ��¼״̬ = 1 And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�)
    If Not rsTmp.EOF Then
        If rsTmp!ժҪ & "" <> "" And InStr(rsTmp!ժҪ & "", "����:") <> 0 Then
            '��ȡ�ҺŻ��۵���Ϣ,�жϹҺŻ��۵��Ƿ���ڣ������ڣ�����������״̬����Ϊ����
            str����NO = Mid(rsTmp!ժҪ & "", Len("����:") + 1)
            strSQL = "Select 1 From ������ü�¼ Where NO = [1] And Mod(��¼����,10) = 1 And ��¼״̬ = 0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����NO)
            If rsTmp.EOF Then
                If lngState = 0 Then '����Ϊ����
                    MsgBox "�ùҺŵ��Ļ��۷��ò����ڣ����˺ź����¹Һ�!", vbInformation + vbDefaultButton1, gstrSysName
                    Exit Function
                End If
            Else
                If lngState = -1 Then '����Ϊ������
                    If MsgBox("�ò��˴��ڹҺŵ��Ļ��۷��ã�����Ϊ������ʱ��ɾ���ùҺŵ��Ļ��۷��ã�" & vbCrLf & "���Ҳ����ٻָ�Ϊ����,�Ƿ����?��", vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    
    gcnOracle.BeginTrans
        strSQL = "Zl_���˹Һż�¼_״̬ ('" & mstr�Һŵ� & "'," & lngState & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Call zlQueueStartus(IIf(lngState = -1, 3, 4), mstr�Һŵ�, mlng����ID)
        'intType:intType:1-����;2-����;3-���˲�����;4-���˴���;5-������ɾ���;6-�ָ�����
        ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
    gcnOracle.CommitTrans
    MsgBox "�����ɹ�!", vbInformation, gstrSysName
    
    Set���˹Һ�״̬ = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
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
    Dim strSQL As String, bln��ͣ As Boolean
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
        
        strSQL = "Select ID From ���˹Һż�¼ where NO=[1] And ��¼����=1 And ��¼״̬=1"
        On Error GoTo errHandle
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTemp.EOF Then
            Exit Sub
        End If
    End With
    
    If Not bln���� Then
        '����
        strSQL = "Zl_���˹Һż�¼_����(" & rsTemp!ID & ",NULL,NULL,NULL,1)"
    Else
        'ȡ������
        strSQL = "Zl_���˹Һż�¼_ȡ������(" & rsTemp!ID & ",1)"
    End If
    
    On Error GoTo errHandle
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
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

Private Sub Set������Ŀ��������()
    Dim lng����ID As Long
    
     On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "���ƻ�������(ZLCISBase)û����ȷ��װ���ù����޷�ִ�С�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    If mlng����ID = 0 Then
        lng����ID = mPatiInfo.����ID
    Else
        lng����ID = mlng����ID
    End If
    If lng����ID = 0 Then
        lng����ID = UserInfo.����ID
    End If
        
    Call gobjCISBase.CallSetClinicCharge(lng����ID, 1, Me, gcnOracle, glngSys, gstrDBUser, E�������, InStr(GetInsidePrivs(p����ҽ��վ), ";������Ŀ��������;") = 0)
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
        Case "·��"
            Call mclsPath.SetFontSize(mbytSize)
        Case "ҽ��"
            Call mclsAdvices.SetFontSize(mbytSize)
        Case "����"
            Call mclsEPRs.SetFontSize(mbytSize)
                Case "�²���"
            On Error Resume Next
            Call mclsEMR.SetFontSize(mbytSize)
            err.Clear: On Error GoTo 0
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
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim rsԤԼʱ�� As Recordset
    Dim strMsg As String
    
    If mlng������� = 0 Then Check������� = True: Exit Function
    
    strSQL = "" & _
    "   Select  Nvl(A.ԤԼʱ��,nvl(����ʱ��,sysdate)) - " & mlng��ǰ����ʱ�� & "/24/60 as �Һ�ʱ��  " & _
    "   From ���˹Һż�¼ A " & _
    "   Where No=[1] And Nvl(A.ԤԼʱ��,nvl(����ʱ��,sysdate))- " & mlng��ǰ����ʱ�� & "*1/24/60>sysdate"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
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

Private Sub mclsMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
'���ܣ���������ҽ��վ���յ�����Ϣ
    Dim objXML As zl9ComLib.clsXML
    Dim rsTmp As ADODB.Recordset
    Dim rsMsg As ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim bytˢ�� As Byte  'ˢ�·�ʽ��1-�����б�2��ת���б�
    Dim blnˢ�� As Boolean
    
    On Error GoTo errH
    
    If strMsgItemIdentity = "ZLHIS_RECIPEAUDIT_001" Then
        If Mid(mstrNotifyAdvice, m�������, 1) = "1" Then
            '����Ϣ�ӵ���Ϣ�б���
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
    Select Case strMsgItemIdentity '��ȡ�Һż�¼id
        Case "ZLHIS_REGIST_001", "ZLHIS_REGIST_002" '���ﻼ�߹Һţ��������֪ͨ����ȡһ����ˢ��һ�η�ʽ������ǵ�һ����Ϣ������ˢ�¡�
            bytˢ�� = 1
            Call objXML.GetSingleNodeValue("register_id", strTmp)
        Case "ZLHIS_CIS_007" '���ﻼ��ת���ʱˢ�£���Ϣ������ʱ���ˢ�£�ֻˢ��ת���б�
            bytˢ�� = 2
            Call objXML.GetSingleNodeValue("clinic_id", strTmp)
    End Select
    
    If strTmp = "" Then Exit Sub
    
    strSQL = "Select ִ����,����,ִ�в���id,ת��ҽ��,ת������,ת�����id From ���˹Һż�¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strTmp))
    
    Set rsMsg = zlDatabase.ParseXMLToRecord(strMsgItemIdentity, strMsgContent)
    If Not rsMsg Is Nothing Then
        Call AddMsgToLis(rsMsg)
    End If
    
    If bytˢ�� = 1 Then
        If mint���ﷶΧ = 1 And rsTmp!ִ���� & "" = UserInfo.���� And (Not mblnҪ����� Or mblnҪ����� And rsTmp!���� & "" <> "") Then
            blnˢ�� = True
        Else
            If (mint���ﷶΧ = 2 And rsTmp!���� & "" = mstr�������� Or mint���ﷶΧ = 3 And (Not mblnҪ����� Or mblnҪ����� And rsTmp!���� & "" <> "")) And _
                Val(rsTmp!ִ�в���ID & "") = IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID) And _
                (rsTmp!ִ���� & "" = "" Or rsTmp!ִ���� & "" = UserInfo.����) Then
                
                blnˢ�� = True
            End If
        End If
        
        If blnˢ�� Then
            mblnMsgOk = True
            If Not mblnFirstMsg Then     '�ǵ�һ����Ϣ
                mblnFirstMsg = True
                Call RefeshByMsg
            End If
        End If
    ElseIf bytˢ�� = 2 Then
        If mint���ﷶΧ = 1 And rsTmp!ת��ҽ�� & "" = UserInfo.���� Then
            blnˢ�� = True
        Else
            If (mint���ﷶΧ = 2 And rsTmp!ת������ & "" = mstr�������� Or mint���ﷶΧ = 3) And _
                Val(rsTmp!ת�����ID & "") = IIf(mlng�������ID = 0, UserInfo.����ID, mlng�������ID) And _
                UserInfo.���� <> IIf("" = rsTmp!ִ���� & "", "��", rsTmp!ִ����) And _
                (rsTmp!ת��ҽ�� & "" = "" Or rsTmp!ת��ҽ�� & "" = UserInfo.����) Then
                
                blnˢ�� = True
            End If
        End If
        
        If blnˢ�� Then
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
'���ܣ�������Ϣƽ̨��ʹ�õ�ˢ�·�ʽ
    If Not mblnMsgOk Then Exit Sub
    Call LoadPatients����
    If Not mbln��ʾԤԼ���� Then
        Call LoadPatientsԤԼ
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
    If Mid(mstrNotifyAdvice, mΣ��ֵ, 1) = "1" Then strTmp = strTmp & ",ZLHIS_LIS_003,ZLHIS_PACS_005"
    If Mid(mstrNotifyAdvice, mҽ������, 1) = "1" Then strTmp = strTmp & ",ZLHIS_OPER_001,ZLHIS_CIS_015,ZLHIS_CIS_005"
    If Mid(mstrNotifyAdvice, m�������, 1) = "1" Then strTmp = strTmp & ",ZLHIS_RECIPEAUDIT_001"
    If Mid(mstrNotifyAdvice, m��Ⱦ��, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CIS_032,ZLHIS_CIS_033"
    If Mid(mstrNotifyAdvice, m��Ѫ���, 1) = "1" And gblnѪ��ϵͳ Then strTmp = strTmp & ",ZLHIS_BLOOD_001"   '����Ѫ�����̲��д���Ϣ�Ͳ���
    If Mid(mstrNotifyAdvice, m��Ѫ���, 1) = "1" And gblnѪ��ϵͳ Then strTmp = strTmp & ",ZLHIS_BLOOD_004"  '����Ѫ����д���Ϣ�Ͳ���
    If Mid(mstrNotifyAdvice, m��Ѫ��Ӧ, 1) = "1" And gblnѪ��ϵͳ Then strTmp = strTmp & ",ZLHIS_BLOOD_006"  '����Ѫ����д���Ϣ�Ͳ���
        
    strTmp = Mid(strTmp, 2)
    If strTmp = "" Then LoadNotify = True: Exit Function
       
    strSQL = "Select b.id,a.����id,a.NO,a.id as �Һ�ID,a.�����,a.����,a.ִ��ʱ�� as ����ʱ��,b.��Ϣ����,b.���ͱ���, b.ҵ���ʶ, b.���ȳ̶�, b.�Ǽ�ʱ��,a.����,b.������Դ" & _
        " From ҵ����Ϣ�嵥 B, ���˹Һż�¼ A" & _
        " Where b.����id=a.Id And a.ִ����||''=[1]  And b.�Ǽ�ʱ��>=Trunc(Sysdate-" & (mintNotifyDay - 1) & ")" & _
        " And Nvl(b.�Ƿ�����,0)=0 And instr(','||[2]||',',','||b.���ͱ���||',')>0 " & _
        " Order By b.���ȳ̶� Desc, b.�Ǽ�ʱ�� Desc"

    Screen.MousePointer = 11

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mstr����ҽ��, strTmp)

    For i = 1 To rsTmp.RecordCount
        Select Case rsTmp!���ͱ���
        Case "ZLHIS_CIS_032", "ZLHIS_CIS_033"
            strTag = strTag & "<TB>" & rsTmp!���ͱ��� & "," & rsTmp!ID
            blnDo = True
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
        Case Else
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!����ID & "," & rsTmp!�Һ�ID & "," & rsTmp!���ͱ��� & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!����ID & "," & rsTmp!�Һ�ID & "," & rsTmp!���ͱ���
                blnDo = True
            End If
        End Select
        
        If blnDo Then
            Call AddReportRow(rsTmp!����ID & "," & rsTmp!�Һ�ID, rsTmp!����ID, rsTmp!NO, NVL(rsTmp!����), NVL(rsTmp!�����), Format(rsTmp!����ʱ�� & "", "yyyy-MM-dd HH:mm"), _
                 NVL(rsTmp!��Ϣ����), rsTmp!���ͱ��� & "", rsTmp!���ȳ̶� & "", Format(rsTmp!�Ǽ�ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), NVL(rsTmp!ҵ���ʶ), rsTmp!������Դ & "", _
                 NVL(rsTmp!����, 0), rsTmp!�Һ�ID, rsTmp!ID)
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
    Dim strRowID As String '�����б��е�Ψһ��ʶ��"����id,��ҳid,��Ϣ����"
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
    objRecord.AddItem Val(arrInput(Index)) '��ϢID��ҵ����Ϣ�嵥.ID
    
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

Private Sub AddMsgToLis(ByVal rsMsg As ADODB.Recordset)
'���ܣ������յ�����Ϣ���������б���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim i As Long
    
    On Error GoTo errH
    
    strSQL = "select a.NO,a.����,a.ִ����,a.�����,a.ִ��ʱ��,a.���� from ���˹Һż�¼ a where a.id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMsg!����id & ""))

    If mstr����ҽ�� = rsTmp!ִ���� & "" Then
        '�ж��б��Ƿ��Ѿ���������Ϣ��
        For i = 0 To rptNotify.Rows.Count - 1
            If Not rptNotify.Rows(i).GroupRow Then
                If rptNotify.Rows(i).Record(C_��Ϣ).Value = rsMsg!���ͱ��� And rptNotify.Rows(i).Record.Tag = CStr(rsMsg!����ID & "," & rsMsg!����id) Then
                    Exit Sub
                End If
            End If
        Next
        
        Call AddReportRow(rsMsg!����ID & "," & rsMsg!����id, rsMsg!����ID, rsMsg!NO, rsTmp!����, NVL(rsTmp!�����), Format(rsTmp!ִ��ʱ�� & "", "yyyy-MM-dd HH:mm"), NVL(rsMsg!��Ϣ����), _
             rsMsg!���ͱ��� & "", rsMsg!���ȳ̶� & "", Format(rsMsg!�Ǽ�ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), rsMsg!ҵ���ʶ & "", rsMsg!������Դ & "", NVL(rsTmp!����, 0), rsMsg!����id, 0)
        
        rptNotify.Populate
         
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
    Dim lngҽ��ID As Long, lng�Һ�id As Long, lng��ϢID As Long
    Dim strҵ�� As String, blnOk As Boolean
    Dim blnFinded As Boolean
    Dim strTmp As String, str���� As String, str����� As String
    Dim strNO As String
    Dim str�Һŵ� As String
    Dim str��Ϣ���� As String
    Dim i As Long
    Dim strPatis As String
    Dim blnOnePati As Boolean
    Dim blnTmp As Boolean
    
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
                lngIndex = .Index
            End With
    
            blnTmp = True
            
            If str�Һŵ� <> mstr�Һŵ� Then blnTmp = LocatePati(str�Һŵ�)
            If strҵ�� <> "" Then      '�ҵ����˺�
                lngҽ��ID = Val(strҵ��)
            End If
            '�����Σ��ֵ��Ϣ���Ķ�������Ϣ
            strTmp = ""
            If strNO = "ZLHIS_LIS_003" Then '����
                strTmp = "ZLHIS_CIS_014"
            ElseIf strNO = "ZLHIS_PACS_005" Then '���
                strTmp = "ZLHIS_CIS_025"
            End If
            If strTmp <> "" Then
                If Not (mclsMipModule Is Nothing) Then
                    If mclsMipModule.IsConnect Then
                        Call ZLHIS_CIS_MsgReadAfter(mclsMipModule, strTmp, lng����ID, str����, , str�����, 1, lng�Һ�id, , mlng����ID, , lngҽ��ID)
                    End If
                End If
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
                blnOk = gobjPublicBlood.zlIsBloodMessageDone(2, lng����ID, lng�Һ�id, 1, mlng����ID)
                If blnOk Then
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
                blnOk = ReadMsgCIS033(lng����ID, lng�Һ�id, strҵ��, lng��ϢID)
                If blnOk Then Call rptNotify.Records.RemoveAt(lngIndex)
            End If
            
            If strNO <> "ZLHIS_CIS_033" And strNO <> "ZLHIS_BLOOD_006" Then
                blnOk = ReadMsg(lng����ID, lng�Һ�id, strNO, strҵ��, lng��ϢID, str�Һŵ�)
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
    Dim str�Һŵ� As String
    
    If rptNotify.SelectedRows.Count = 0 Then Exit Sub  '���������
    
    str�Һŵ� = rptNotify.SelectedRows(0).Record.Item(C_No).Value
 
    If str�Һŵ� <> mstr�Һŵ� Then Call LocatePati(str�Һŵ�)
    
End Sub

Private Function ReadMsg(ByVal lng����ID As Long, ByVal lng�Һ�id As Long, ByVal strNO As String, ByVal strҵ�� As String, ByVal lng��ϢID As Long, ByVal str�Һŵ� As String) As Boolean
'���ܣ��Ķ���Ϣ
'˵������Ϣ�Ķ���ʽĿǰ��3�֣�����Ϣ�������Ķ�����ϢID�Ķ�����ҵ���ʶ�Ķ�
    Dim strSQL As String
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
    
    strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng�Һ�id & ",'" & strNO & "',1,'" & UserInfo.���� & "'," & lng����ID
    Select Case strNO
    Case "ZLHIS_LIS_003", "ZLHIS_PACS_005"
        strSQL = strSQL & ",null,null,'" & strҵ�� & "'"
    Case "ZLHIS_CIS_032", "ZLHIS_CIS_033"
        strSQL = strSQL & ",null," & lng��ϢID
    End Select
    strSQL = strSQL & ")"
    
    strSQLReadMsg = strSQL
    
    If strNO = "ZLHIS_LIS_003" Or strNO = "ZLHIS_PACS_005" Then
        If mblnΣ��ֵ Then
            'Σ��ֵ��Ϣ��ش���
            Call mobjKernel.ShowDealCritical(Me, lng����ID, 0, str�Һŵ�, lngΣ��ֵID)
            
            If lngΣ��ֵID <> 0 Then
                strSQL = "select a.�걾id,a.�������,a.ȷ���� from ����Σ��ֵ��¼ a where a.id=[1] and a.ȷ���� is not null"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngΣ��ֵID)
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
            strSQL = "select 1 from ����ҽ����¼ a where a.�Һŵ�=[1] and a.ҽ��״̬=1 and a.�������='K' and a.��鷽��='1' and a.���״̬=1 and rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str�Һŵ�)
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
    txtInfo(txtInfoժҪ).Left = lblEdit(txtInfoժҪ).Left + 20 + lblEdit(txtInfoժҪ).Width
    txtInfo(txtInfoժҪ).Width = picMore.Width - txtInfo(txtInfoժҪ).Left
    UCPatiVitalSigns.Top = txtInfo(txtInfoժҪ).Top + txtInfo(txtInfoժҪ).Height + 60
    UCPatiVitalSigns.Left = 10
End Sub

Private Sub picBasisNew_Resize()
    On Error Resume Next
    '�˴����Թ̶��߶�
    
    lblUrg.FontName = "����"
    lblUrg.FontSize = IIf(mbytSize = 0, 14, 18)
    
    lblRec.FontName = "����"
    lblRec.FontSize = IIf(mbytSize = 0, 14, 18)
    
    If err.Number <> 0 Then err.Clear

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
    lblLink(lblLink�޸�).Top = lblLink(lblLink���).Top
    
    txtInfo(txtInfo����).Top = 100
    txtInfo(txtInfo����).Left = lblLink(lblLink�޸�).Left
    txtInfo(txtInfo����).FontSize = IIf(mbytSize = 0, 12, 15)
    txtInfo(txtInfo����).Width = IIf(mbytSize = 0, 1400, 1800)
    
    txtInfo(txtInfo�Ա�).Top = txtInfo(txtInfo����).Top + txtInfo(txtInfo����).Height - txtInfo(txtInfo�Ա�).Height + 160
    txtInfo(txtInfo�Ա�).Left = txtInfo(txtInfo����).Left + txtInfo(txtInfo����).Width + 50
    
    txtInfo(txtInfo����).Top = txtInfo(txtInfo�Ա�).Top + txtInfo(txtInfo�Ա�).Height - txtInfo(txtInfo����).Height
    txtInfo(txtInfo����).Left = txtInfo(txtInfo�Ա�).Left + txtInfo(txtInfo�Ա�).Width + 100
    
    Call zlControl.SetPubCtrlPos(False, -1, txtInfo(txtInfo����), 250, lblEdit(txtInfo��������), 30, txtInfo(txtInfo��������), 250, lblEdit(txtInfo���ѷ�ʽ), 30, fraPayType)
    fraPayType.Top = lblEdit(txtInfo���ѷ�ʽ).Top - 30
    
    fraPayType.Width = cboPayType.Width
    fraPayType.Height = cboPayType.Height - 60
    
    linPayType.x1 = fraPayType.Left - 20
    linPayType.y1 = fraPayType.Top + fraPayType.Height
    linPayType.x2 = linPayType.x1 + fraPayType.Width
    linPayType.y2 = linPayType.y1
    
    lblEdit(12).Left = lblEdit(10).Left
    lblEdit(12).Top = lblLink(lblLink�޸�).Top + 10
    
    txtInfo(txtInfo���￨��).Width = 1300
    Call zlControl.SetPubCtrlPos(False, -1, lblLink(lblLink�޸�), 250, lblEdit(txtInfo����), 30, txtInfo(txtInfo����), 150, lblEdit(txtInfo���￨��), 30, txtInfo(txtInfo���￨��), 150, lblEdit(txtInfoҽ������), 30, txtInfo(txtInfoҽ������), 150, lblEdit(txtInfo�ѱ�), 30, fraBillType)
    fraBillType.Top = lblEdit(txtInfo�ѱ�).Top - 30
    
    fraBillType.Width = cboBillType.Width
    fraBillType.Height = cboBillType.Height - 60
    
    linBillType.x1 = fraBillType.Left - 20
    linBillType.y1 = fraBillType.Top + fraBillType.Height
    linBillType.x2 = linBillType.x1 + fraBillType.Width
    linBillType.y2 = linBillType.y1
    
    lblMore.Top = lblEdit(txtInfoҽ������).Top
    lblMore.Left = picBasisNew.Width - lblMore.Width - 40

    lblUrg.Top = 200
    lblUrg.Left = picBasisNew.Width - lblUrg.Width - 40
    lblRec.Top = lblUrg.Top
    lblRec.Left = lblUrg.Left - lblRec.Width - 20
    
    lblLink(lblLink��ʾ).Left = fraBillType.Left + fraBillType.Width + 200
    lblLink(lblLink��ʾ).Top = txtInfo(txtInfoҽ������).Top
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
            err.Clear: On Error GoTo 0
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
    Case lblLink��ʾ
        If mblnPatiDetail Then
            mblnPatiDetail = False
            lblLink(Index).Caption = "��ʾ������Ϣ��"
        Else
            mblnPatiDetail = True
            lblLink(Index).Caption = "����������Ϣ��"
        End If
        Call cbsMain_Resize
        Call zlDatabase.SetPara("��ʾ������ϸ��Ϣ", IIf(mblnPatiDetail, 1, 0), glngSys, p����ҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
    End Select
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    err.Clear
End Sub

Private Function SetPatPicture(ByVal lng����ID As Long, ByVal blnDel As Boolean) As Boolean
'����:���ò�����Ƭ
'���:lng����ID - ����ID��blnDel true ɾ����Ƭ��false ������Ƭ
    Dim strFile As String, strSQL As String
    On Error GoTo errH

    If blnDel Then
        If MsgBox("����" & txtInfo(txtInfo����).Text & "����Ƭ����ɾ�����Ƿ������", vbDefaultButton2 + vbYesNo + vbQuestion, Me.Caption) = vbNo Then
            Exit Function
        End If
        strSQL = strSQL & "Zl_������Ƭ_Delete(" & lng����ID & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Else
        'ͼƬû�б�����������²���ͼƬ
        If picPatient.Tag <> "" Then
            strFile = picPatient.Tag
            If Sys.SaveLob(glngSys, 27, lng����ID, strFile) = False Then
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
            If mobjEPRDoc Is Nothing Then
                Set mobjEPRDoc = New zlRichEPR.cEPRDocument
            End If
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
    Dim strTmp As String
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
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    
    strSQL = "Select ����, ���� From ҽ�Ƹ��ʽ Order By ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With cboPayType
        .Clear
        For i = 1 To rsTmp.RecordCount
            .AddItem rsTmp!���� & "-" & rsTmp!����
            .ItemData(.NewIndex) = Val(rsTmp!���� & "")
            rsTmp.MoveNext
        Next
    End With
    
    strSQL = "Select ����, ���� From �ѱ� Order By ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
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
    Dim i As Long, j As Long
    Dim strSQL As String
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
            strSQL = "select ID,NO,����,����ID,ִ��״̬,ת��״̬,decode(��¼��־,2,1,3,1,0) as ���� from ���˹Һż�¼ where ��¼����=1 And ��¼״̬=1 and id=[1]"
            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�Һ�id)
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
                            strSQL = "Zl_���˹Һż�¼_ת��('" & rsPati!NO & "',Null)"
                            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                        End If
                    End If
                    Call ExecuteFinishInSide(rsPati!NO & "", Val(rsPati!����ID & ""), lng�Һ�id)
                ElseIf intType <> 0 Then
                    If Val(rsPati!���� & "") = 0 Then
                        strSQL = "Zl_���˹Һż�¼_����(" & lng�Һ�id & ",NULL,NULL,NULL,1)"
                        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
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
    End If
    
    If str��������2 <> "" Then
        str��������2 = Mid(str��������2, 2) & " �����Զ���ǻ�����ھ����б��в鿴��"
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
    Dim i As Long, j As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsPati As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim blnDo As Boolean, blnSigned As Boolean, blnOK���� As Boolean
    Dim objEmr As Object
    Dim strҽ��IDs As String
    Dim lngTmp As Long, lngTmp1 As Long
    Dim str��������1 As String
    Dim str��������2 As String
    
    On Error GoTo errH
    intType = 1
    '1.�������
    strSQL = "select 1 from ���Ӳ�����¼ where ����ID=[1] and ��ҳID=[2] and ǩ������<>1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng�Һ�id)
    If rsTmp.EOF Then
        blnSigned = True
        If GetInsidePrivs(p�°����ﲡ��, True) <> "" Then
            On Error Resume Next
            Set objEmr = CreateObject("zl9EmrInterface.ClsEmrInterface")
            If Not objEmr Is Nothing Then
                Call objEmr.CheckOutEPRIsAllSign(lng�Һ�id, blnSigned)
            End If
            err.Clear: On Error GoTo 0
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
        strSQL = "select a.id,a.���id,a.���,a.ҽ��״̬,a.�������," & _
            " NVL(a.ִ�б��,0) as ִ�б��, Nvl(a.ִ������,0) as ִ������,Nvl(a.Ƥ�Խ��,'��') as Ƥ�Խ�� from ����ҽ����¼ a where a.ҽ��״̬<>4 and a.�Һŵ�=[1] and a.����ID+0=[2]"
        Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, lng����ID)
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
            strSQL = "select 1 from ����ҽ������ a where a.ҽ��id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) and a.ִ��״̬<>1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strҽ��IDs)
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
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim lng�ļ�ID As Long
    Dim lng����ID As Long
    Dim objControl As CommandBarControl
    
    On Error GoTo errH
    'conMenu_Edit_Modify 3003 �޸İ�ť��
    lng�ļ�ID = Val(Split(str��ʶ, ",")(0))
    
    strSQL = "Select 1 From �����걨��¼ where �ļ�ID=[1] and ����״̬=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ļ�ID, 4)
    If rsTmp.RecordCount = 0 Then
    '����Ϣ���Ϊ�Ѷ�
        If mlng�������ID = 0 Then
            lng����ID = UserInfo.����ID
        Else
            lng����ID = mlng�������ID
        End If
        strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng����ID & ",'ZLHIS_CIS_033',1,'" & UserInfo.���� & "'," & lng����ID & ",null," & lng��ϢID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
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
    
    strSQL = "Select 1 From �����걨��¼ where �ļ�ID=[1] and ����״̬=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ļ�ID, 4)
    If rsTmp.RecordCount = 0 Then
    '����Ϣ���Ϊ�Ѷ�
        If mlng�������ID = 0 Then
            lng����ID = UserInfo.����ID
        Else
            lng����ID = mlng�������ID
        End If
        strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng����ID & ",'ZLHIS_CIS_033',1,'" & UserInfo.���� & "'," & lng����ID & ",null," & lng��ϢID & ")"
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
    
    Call mobjKernel.ShowDealCritical(Me, mlng����ID, 0, mstr�Һŵ�, lngΣ��ֵID)
    
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
        Set .Icons = zlCommFun.GetPubIcons
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
        Set .Icons = zlCommFun.GetPubIcons

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
 
    Dim strFilter As String
    Dim j As Long, i As Long
    Dim strValues(0 To 10) As String, strValue As String, strUninTable As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsLocal As ADODB.Recordset
    Dim rptCalling As ReportRecord
    Dim rptRecord As ReportRecord
 

    On Error GoTo errH
    
    '���Զ������̣�����113794ǰ�Ĵ���ʽ
    strFilter = "": strValue = "": j = 0: strUninTable = ""
    If SafeArrayGetDim(str��������) > 0 Then
        For i = 1 To UBound(str��������)
            If Trim(str��������(i)) <> "" Then
                If j > 10 Then
                    strFilter = strFilter & " Or A.�������� ='" & str��������(i) & "'"
                Else
                    If zlCommFun.ActualLen(strValue) > 2000 Then
                         strValues(j) = Mid(strValue, 2)
                         strUninTable = strUninTable & " Union ALL  Select  Column_Value as �������� From Table(Cast(f_Str2list([" & j + 4 & "]) As zlTools.t_Strlist))  " & vbCrLf
                         strValue = "": j = j + 1
                    End If
                    strValue = strValue & "," & str��������(i)
                End If
            End If
        Next i
        If strValue <> "" Then
            strValues(j) = Mid(strValue, 2)
            strUninTable = strUninTable & " Union ALL  Select  Column_Value as �������� From Table(Cast(f_Str2list([" & j + 4 & "]) As zlTools.t_Strlist))  " & vbCrLf
        End If
    End If
    
    If strUninTable <> "" Then
        strUninTable = Mid(strUninTable, 11)
    Else
'        labError.Caption = "û�п���ʾ�ĽкŶ�����Ϣ����������Ŷӿ�������"
        Exit Sub
    End If
    
    If strFilter <> "" Then strFilter = "( " & Mid(strFilter, 4) & ")"
     
    'Ϊ��֧�ָ��ƣ���Ҫ��number���͵��ֶν���ת��������ʹ��to_Number��ʽ
    strSQL = "" & _
    "   Select /*+ Rule*/  to_Number(A.ID) as ID, to_Number(a.����id) as ����id, A.��������, A.�Ŷ����, to_Number(A.ҵ������) as ҵ������, to_Number(A.ҵ��ID) as ҵ��ID," & _
    "           to_Number(A.����ID) as ����ID, x.���� as ��������, A.�ŶӺ��� , A.�Ŷӱ��,A.��������||decode(e.ԤԼ,1,'(Ԥ)',null) as ��������,A.����,A.ҽ������," & _
    "            (select j.���� from ��Ա�� J,�ϻ���Ա�� K where J.ID=K.��ԱID and K.�û���=A.����ҽ��) as ����ҽ��, " & _
    "           to_Number(A.����) as ����, to_Number(A.�������) as �������, To_Char(A.�Ŷ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') as �Ŷ�ʱ��, To_Char(A.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') as ����ʱ��,to_Number(A.�Ŷ�״̬) as �Ŷ�״̬, " & _
                IIf(mlng���ﲡ������ = 1, "to_number(nvl(A.�������, 9999999999)) as ���������", "0 as ���������") & _
    "   From �ŶӽкŶ��� a, ���ű� x " & IIf(strUninTable <> "", ", (" & strUninTable & ") b ", "") & _
                IIf(intViewDataType = 1, ", Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D ", "") & _
                IIf(intViewDataType = 2, ", Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D ", "") & _
                IIf(intViewDataType = 3, ", Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D", "") & " , ���˹Һż�¼ E" & _
    "   Where To_Number(a.ҵ��id) = e.Id and  (nvl(a.�Ƿ��ʱ��, 0)=0 and A.�Ŷ�ʱ�� <= trunc(sysdate + 1) - 1/24/60/60 or nvl(a.�Ƿ��ʱ��, 0)=1 and sysdate>a.�Ŷ�ʱ��) " & IIf(strUninTable <> "", " and a.��������=b.�������� ", "") & " and instr([3],A.�Ŷ�״̬)=0  and x.ID=a.����ID  " & _
                IIf(intViewDataType = 1, " and  ((a.����=C.Column_Value and a.ҽ������ is null) or a.ҽ������=D.Column_Value or (a.���� is null and a.ҽ������ is null))", "") & _
                IIf(intViewDataType = 2, " and (a.����=C.Column_Value and (a.ҽ������ is Null or a.ҽ������=D.Column_Value)) ", "") & _
                IIf(intViewDataType = 3, " and a.ҽ������=D.Column_Value", "") & _
    "           " & strFilter & _
    "   Order by  �Ŷ�״̬ desc, �Ŷ����,���� Desc, ���������, �Ŷ�ʱ��, �ŶӺ��� "
    

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����", str����, strҽ��, strִ��״̬, strValues(0), strValues(1), strValues(2), strValues(3), strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10))
    Set rsLocal = zlDatabase.CopyNewRec(rsTemp)
    
    'ɾ����Ҫ�ų�������,����ȡʵ���ŶӺ���ֵ�����
    If rsLocal.RecordCount > 0 Then rsLocal.MoveFirst
    While Not rsLocal.EOF
        If InStr(1, strExcludeData, rsLocal!ҵ������ & ":" & rsLocal!ҵ��ID) > 0 Then
            rsLocal.Delete
        End If
        If LenB(StrConv(Trim(NVL(rsLocal("�ŶӺ���"))), vbFromUnicode)) > mlngMaxLen Then
            mlngMaxLen = LenB(StrConv(Trim(NVL(rsLocal("�ŶӺ���"))), vbFromUnicode))
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
    rriItem(mCol.����ID).Value = NVL(rsData("����ID"))
    
    rriItem(mCol.��������).Caption = rsData("��������") & ":" & IIf(InStr(1, NVL(rsData("��������")), ":") <= 0, "", Mid(NVL(rsData("��������")), InStr(1, NVL(rsData("��������")), ":") + 1))
    rriItem(mCol.��������).Value = NVL(rsData("��������"))

    rriItem(mCol.��������).Value = NVL(rsData("��������"))
    rriItem(mCol.����ID).Value = NVL(rsData("����ID"))
    rriItem(mCol.�Ŷӱ��).Value = NVL(rsData("�Ŷӱ��"))
    rriItem(mCol.�Ŷ����).Value = zlStr.Lpad(NVL(rsData("�Ŷ����")), 20)
    rriItem(mCol.�ŶӺ���).Value = zlStr.Lpad(NVL(rsData("�ŶӺ���")), mlngMaxLen)
    rriItem(mCol.�Ŷ�ʱ��).Value = NVL(rsData("�Ŷ�ʱ��"))
    rriItem(mCol.����ʱ��).Value = NVL(rsData("����ʱ��"))
    rriItem(mCol.�������).Value = NVL(rsData("�������"))
    rriItem(mCol.���������).Value = NVL(rsData("���������"))
    rriItem(mCol.����ҽ��).Value = NVL(rsData("����ҽ��"))
    rriItem(mCol.��������).Value = DeptNametransform(NVL(rsData("��������")))
    rriItem(mCol.��������).Caption = (NVL(rsData("��������")))
    rriItem(mCol.ORD).Value = Format(rsData.AbsolutePosition, "00000000")
    
    If NVL(rsData("�������")) = "" Then
        rriItem(mCol.��������).Icon = 807
    Else
        rriItem(mCol.��������).Icon = 3504
    End If
    
    
    If NVL(rsData("�Ŷ�״̬")) = 1 Then
        rriItem(mCol.�Ŷ�״̬).Value = "������"
        For i = 0 To mobjQueueList.Columns.Count - 1
            rriItem(i).BackColor = &HC0C0FF
        Next
    ElseIf NVL(rsData("�Ŷ�״̬")) = 0 Then
        rriItem(mCol.�Ŷ�״̬).Value = "�Ŷ���"
        For i = 0 To mobjQueueList.Columns.Count - 1
            rriItem(i).BackColor = ColorConstants.vbWhite
        Next
    ElseIf NVL(rsData("�Ŷ�״̬")) = 3 Then
        rriItem(mCol.�Ŷ�״̬).Value = "��ͣ"
        For i = 0 To mobjQueueList.Columns.Count - 1
            rriItem(i).BackColor = ColorConstants.vbYellow
        Next
    ElseIf NVL(rsData("�Ŷ�״̬")) = 4 Then
        rriItem(mCol.�Ŷ�״̬).Value = "���"
        For i = 0 To mobjQueueList.Columns.Count - 1
            rriItem(i).BackColor = ColorConstants.vbGreen
        Next
    ElseIf NVL(rsData("�Ŷ�״̬")) = 7 Then
        rriItem(mCol.�Ŷ�״̬).Value = "�Ѻ���"
    Else
        rriItem(mCol.�Ŷ�״̬).Value = "������"
        For i = 0 To mobjQueueList.Columns.Count - 1
            rriItem(i).BackColor = &HC0C0C0
        Next
    End If
    
    If mlngQueueGroupType = 1 Then
        rriItem(mCol.ҽ������).Value = NVL(rsData("��������")) & ":" & NVL(rsData("ҽ������"))
    Else
        rriItem(mCol.ҽ������).Value = NVL(rsData("ҽ������"))
    End If

    rriItem(mCol.ҵ������).Value = NVL(rsData("ҵ������"))
    rriItem(mCol.ҵ��ID).Value = NVL(rsData("ҵ��ID"))

    rriItem(mCol.����).Value = IIf(NVL(rsData("����")) = 1, "����", "")
    
    If mlngQueueGroupType = 2 Then
        rriItem(mCol.����).Value = NVL(rsData("��������")) & ":" & NVL(rsData("����"))
    Else
        rriItem(mCol.����).Value = NVL(rsData("����"))
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
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim objControl As Object
    
    On Error GoTo errH
    If lng��¼ID = 0 Then Exit Sub
    strSQL = "select 1 from ����Σ��ֵ��¼ a where a.id=[1] and a.�Ƿ�Σ��ֵ=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID)
    
    If Not rsTmp.EOF Then
        '�����´�ҽ���Ĵ���
        If tbcSub.Tag <> "ҽ��" Then
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
    Dim strSQL As String
    Dim strTime As String
    Dim i As Long, j As Long, k As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim rsPati As ADODB.Recordset
    Dim strTmp As String
    Dim lngColor As Long
    Dim rs��Ⱦ��״̬ As ADODB.Recordset
    Dim blnDo��Ⱦ��״̬ As Boolean
    
    On Error GoTo errH
    Screen.MousePointer = 11
    mblnUnRefresh = True
     
    strSQL = _
        " Select B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,B.����,B.����,B.����,nvl(g.����,E.����) as ����,D.���� as ���˿���," & _
        " B.ִ��ʱ�� as ʱ��,A.���￨��,A.���֤��,A.IC����,A.����,B.����ʱ��,B.ִ�в���ID,B.ִ����," & _
        " B.ת��״̬,C.���� as ת�����,B.ת������,B.ת��ҽ��,B.ִ��״̬,B.��¼��־,A.��������" & _
        " From ������Ϣ A,���˹Һż�¼ B,���ű� C,���ű� D,�ҺŰ��� E, �ٴ������¼ f,�ٴ������Դ g" & _
        " Where B.����ID=A.����ID And B.ת�����ID=C.ID(+) and B.�ű�=E.����(+) and B.ִ�в���ID=d.id and b.�����¼id=f.id(+) and f.��Դid=g.id(+)" & _
        " And B.ִ��״̬=2 And B.ִ����||''=[1] And B.��¼����=1 And B.��¼״̬=1 and nvl(B.��¼��־,0) in (2,3)" & _
        " Order By B.NO"
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr����ҽ��)
    
    strSQL = "select m.����id,m.id,m.no,max(m.��¼) as ��¼,max(m.��д) as ��д,max(m.״̬) as ״̬ from" & vbNewLine & _
        "(select a.����id,a.id, a.no,1 as ��¼,0 as ��д,0 as ״̬ from ���˹Һż�¼ a,�������Լ�¼ b" & vbNewLine & _
        "where a.no=b.�Һŵ� and a.ִ��״̬=2 And a.ִ����||''=[1] And a.��¼����=1 And a.��¼״̬=1 and nvl(a.��¼��־,0) in (2,3)" & vbNewLine & _
        "union all" & vbNewLine & _
        "Select  a.����id,a.id, a.no,0 as ��¼,1 as ��д,0 as ״̬" & vbNewLine & _
        "From ���˹Һż�¼ A, ���Ӳ�����¼ C, �����ļ��б� D" & vbNewLine & _
        "Where c.�ļ�id = d.Id And d.���� = 5  and c.�������� like '%��Ⱦ��%' And a.����id = c.����id And a.id = c.��ҳid and a.ִ��״̬=2 And a.ִ����||''=[1] And a.��¼����=1 And a.��¼״̬=1 and nvl(a.��¼��־,0) in (2,3)" & vbNewLine & _
        "union all" & vbNewLine & _
        "Select  a.����id,a.id, a.no,0 as ��¼,1 as ��д,e.����״̬ as ״̬" & vbNewLine & _
        "From ���˹Һż�¼ A,���Ӳ�����¼ C,�����ļ��б� D,�����걨��¼ E" & vbNewLine & _
        "Where a.����id = c.����id And a.id = c.��ҳid and c.id=e.�ļ�id and d.����=5 and c.�������� like '%��Ⱦ��%' and e.�ļ�id =d.id and a.ִ��״̬=2 And a.ִ����||''=[1] And a.��¼����=1 And a.��¼״̬=1 and nvl(a.��¼��־,0) in (2,3)) M" & vbNewLine & _
        "group by m.����id,m.id,m.no"
    Set rs��Ⱦ��״̬ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr����ҽ��)
    If rs��Ⱦ��״̬.RecordCount > 0 Then blnDo��Ⱦ��״̬ = True
    
    rptPati(PATI_RPT����).Records.DeleteAll
    For i = 1 To rsPati.RecordCount
        Set objRecord = rptPati(PATI_RPT����).Records.Add()
        For j = 0 To rptPati(PATI_RPT����).Columns.Count - 1
            objRecord.AddItem ""
        Next
        With objRecord
            .Item(COL_JZ_��ʶ).Value = "��"
            .Item(COL_JZ_�����).Value = rsPati!����� & ""
            .Item(COL_JZ_����).Value = rsPati!���� & ""
            .Item(COL_JZ_����ʱ��).Value = Format(rsPati!ʱ��, "yyyy-MM-dd HH:mm")
            .Item(COL_JZ_�Ա�).Value = rsPati!�Ա� & ""
            .Item(COL_JZ_����).Value = rsPati!���� & ""
            .Item(COL_JZ_��).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
            .Item(COL_JZ_��).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
            .Item(COL_JZ_NO).Value = rsPati!NO & ""
            .Item(COL_JZ_����).Value = IIf(Val(rsPati!���� & "") <> 0, "��", "")
            .Item(COL_JZ_���￨��).Value = rsPati!���￨�� & ""
            .Item(COL_JZ_��������).Value = rsPati!�������� & ""
            .Item(COL_JZ_����ID).Value = rsPati!����ID & ""
            .Item(COL_JZ_����ʱ��).Value = CStr(Format(rsPati!����ʱ��, "yyyy-MM-dd HH:mm:ss"))
            .Item(COL_JZ_ִ�в���ID).Value = rsPati!ִ�в���ID & ""
            .Item(COL_JZ_ִ����).Value = rsPati!ִ���� & ""
            .Item(COL_JZ_���֤��).Value = rsPati!���֤�� & ""
            .Item(COL_JZ_IC����).Value = rsPati!IC���� & ""
            .Item(COL_JZ_��¼��־).Value = rsPati!��¼��־ & ""
            .Item(COL_JZ_����).Value = rsPati!���� & ""
            .Item(COL_JZ_���˿���).Value = rsPati!���˿��� & ""
            
            '���ղ����ú�ɫ��ʾ
            If Not IsNull(rsPati!����) And rsPati!�������� & "" = "" Then
                .Item(COL_JZ_�����).ForeColor = &HC0&
                .Item(COL_JZ_��������).ForeColor = &HC0&
            Else
                '������ɫ
                lngColor = zlDatabase.GetPatiColor(NVL(rsPati!��������))
                .Item(COL_JZ_�����).ForeColor = lngColor
                .Item(COL_JZ_��������).ForeColor = lngColor
            End If
            
            '�����־��ɫͻ����ʾ
            If NVL(rsPati!����, 0) <> 0 Then
                .Item(COL_JZ_��).ForeColor = vbRed
            End If
            
            '��Ӵ�Ⱦ��״̬
            strSQL = ""
            If blnDo��Ⱦ��״̬ Then
                rs��Ⱦ��״̬.Filter = "no='" & rsPati!NO & "'"
                If Not rs��Ⱦ��״̬.EOF Then strSQL = Get��Ⱦ��״̬(Val(rs��Ⱦ��״̬!��¼ & ""), Val(rs��Ⱦ��״̬!��д & ""), Val(rs��Ⱦ��״̬!״̬ & ""))
            End If
            .Item(COL_JZ_��Ⱦ��).Value = strSQL
        End With
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

Private Sub ReadMsgAuto()
'���ܣ�Σ��ֵ��Ϣ�����Զ�����
    Dim i As Long
    Dim lng����ID As Long
    Dim lng��ҳID As Long
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
                    blnRs = ReadMsg(lng����ID, lng�Һ�id, strNO, strҵ��, lng��ϢID, str�Һŵ�)
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
