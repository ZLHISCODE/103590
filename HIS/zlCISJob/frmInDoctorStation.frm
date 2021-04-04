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
   Caption         =   "סԺҽ������վ"
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
         ToolTipText     =   "ȡ��"
         Top             =   5550
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterOK 
         Height          =   315
         Left            =   3990
         Picture         =   "frmInDoctorStation.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "ȷ��"
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
            Text            =   "�������"
            TextSave        =   "�������"
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
            Text            =   "������ɫ"
            TextSave        =   "������ɫ"
            Key             =   "������ɫ"
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
         Begin VB.ComboBox cbo���� 
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
            Caption         =   "����:"
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
            Caption         =   "ҽ��:"
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
            Caption         =   "��Һ��:"
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
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl���� 
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
         Begin VB.Label lblҽ���� 
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
         Begin VB.Label lbl��Ժ 
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
         Begin VB.Label lbl���� 
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
         Begin VB.Label lbl���� 
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
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            Height          =   180
            Index           =   0
            Left            =   2055
            TabIndex        =   25
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            Height          =   180
            Index           =   0
            Left            =   4065
            TabIndex        =   24
            Top             =   450
            Width           =   450
         End
         Begin VB.Label lbl��Ժ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ:"
            Height          =   180
            Index           =   0
            Left            =   5115
            TabIndex        =   23
            Top             =   450
            Width           =   450
         End
         Begin VB.Label lblҽ���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ����:"
            Height          =   180
            Index           =   0
            Left            =   8595
            TabIndex        =   22
            Top             =   165
            Width           =   630
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            Height          =   180
            Index           =   0
            Left            =   7060
            TabIndex        =   21
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ��:"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   20
            Top             =   150
            Width           =   630
         End
         Begin VB.Label lbl���� 
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
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            Height          =   180
            Index           =   0
            Left            =   4140
            TabIndex        =   18
            Top             =   180
            Width           =   450
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ:"
            Height          =   180
            Index           =   0
            Left            =   3465
            TabIndex        =   17
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҩ��:"
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
            Caption         =   "���:"
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
            Caption         =   "ͼ��:"
            Height          =   180
            Left            =   -45
            TabIndex        =   77
            Top             =   1110
            Width           =   450
         End
         Begin VB.Label lblPatiName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "����:"
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
            Caption         =   "����:"
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
         ToolTipText     =   "����סԺ�ŶԲ��˽��о�ȷ������ʾ"
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
            Caption         =   "���"
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
               Caption         =   "��С����ʾ"
               BeginProperty Font 
                  Name            =   "����"
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
               ToolTipText     =   "�Ƿ�ҽ��С��ģʽ��ʾ�����б�"
               Top             =   23
               Width           =   1280
            End
            Begin VB.CheckBox chk�������� 
               Caption         =   "��"
               Height          =   195
               Index           =   2
               Left            =   2070
               TabIndex        =   36
               ToolTipText     =   "Ctrl+��ѡ������ѡ��"
               Top             =   23
               Value           =   1  'Checked
               Width           =   480
            End
            Begin VB.CheckBox chk�������� 
               Caption         =   "Σ"
               Height          =   195
               Index           =   1
               Left            =   1500
               TabIndex        =   38
               ToolTipText     =   "Ctrl+��ѡ������ѡ��"
               Top             =   23
               Value           =   1  'Checked
               Width           =   465
            End
            Begin VB.CheckBox chk�������� 
               Caption         =   "һ��"
               Height          =   195
               Index           =   0
               Left            =   750
               TabIndex        =   37
               ToolTipText     =   "Ctrl+��ѡ������ѡ��"
               Top             =   23
               Value           =   1  'Checked
               Width           =   660
            End
            Begin VB.Label lbl�������� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�� ��(&S)"
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
               Caption         =   "ˢ��"
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
            Begin VB.Label lblת�� 
               AutoSize        =   -1  'True
               Caption         =   "��ʾ���    ���ת������"
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
               Caption         =   "�ѻ���"
               Height          =   180
               Index           =   1
               Left            =   3225
               TabIndex        =   67
               Top             =   120
               Value           =   1  'Checked
               Width           =   900
            End
            Begin VB.CheckBox chkHZ 
               Caption         =   "δ����"
               Height          =   180
               Index           =   0
               Left            =   2160
               TabIndex        =   66
               Top             =   90
               Value           =   1  'Checked
               Width           =   900
            End
            Begin VB.CheckBox chkOut 
               Caption         =   "������Ժ����"
               Height          =   195
               Left            =   2760
               TabIndex        =   53
               ToolTipText     =   "Ctrl+��ѡ������ѡ��"
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
            Begin VB.Label lbl��ʼʱ�� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ʼʱ��"
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
               Caption         =   "��С����ʾ"
               BeginProperty Font 
                  Name            =   "����"
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
               ToolTipText     =   "�Ƿ�ҽ��С��ģʽ��ʾ�����б�"
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
            Begin VB.Label lbl��Ժʱ�� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժʱ��"
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
      Begin VB.Frame fra��� 
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
         Begin VB.Label lbl��� 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "���� XXX ��δ����Ĳ�����鷴��..."
            BeginProperty Font 
               Name            =   "����"
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
               Key             =   "�ȴ����"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":10CB2
               Key             =   "�ܾ����"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":1124C
               Key             =   "�������"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":117E6
               Key             =   "���ڳ��"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":121F8
               Key             =   "��鷴��"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":12C0A
               Key             =   "��鷴��"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":131A4
               Key             =   "�������"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":13BB6
               Key             =   "�������"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":145C8
               Key             =   "δ����"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":14B62
               Key             =   "�������"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":150FC
               Key             =   "��������"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":15696
               Key             =   "������"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":160A8
               Key             =   "ִ����"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":16642
               Key             =   "Child"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorStation.frx":16BDC
               Key             =   "������"
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
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmInDoctorStation.frx":1DF72
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
      Begin VB.Label lblFind 
         Caption         =   "����(F3)"
         Height          =   255
         Left            =   360
         TabIndex        =   55
         Top             =   525
         Width           =   735
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&D)��"
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
    col_���� = 0
    col_��� = 1
    col_ͼ�� = 2
    col_·��״̬ = 3
    col_����Id = 4
    col_��ҳID = 5
    col_���� = 6
    col_סԺ�� = 7
    col_���� = 8
    col_�໤ = 9
    col_�Ա� = 10
    col_���� = 11
    col_�ѱ� = 12
    col_���� = 13
    col_���� = 14
    col_סԺҽʦ = 15
    col_��Ժ���� = 16  'ȡ���ʱ�� ������ǰ����:���ʱ�� Ϊ��ʱȡ��Ժ����
    col_��Ժ���� = 17
    col_�������� = 18
    
    col_ҽ��ID = 19
    col_���ͺ� = 20
    col_ִ��״̬ = 21
    col_ִ�п���ID = 22
    col_���￨ = 23
    col_סԺ���� = 24
    col_������ = 25
    COL_Ӥ������ID = 26
    COL_Ӥ������ID = 27
    col_��ҽ��� = 28
    col_��ҽ��� = 29
    COL_������� = 30
    COL_��Ⱦ�� = 31
    col_���λ�ʿ = 32
    col_���ۺ� = 33
    col_���֤�� = 34
    col_�Ƿ��� = 35

End Enum

Private Enum NOTIFYREPORT_COLUMN
    c_ͼ�� = 0
    C_����Id = 1
    C_��ҳId = 2
    c_���� = 3
    c_סԺ�� = 4
    c_���� = 5
    C_״̬ = 6
    '������
    C_��Ϣ = 7
    C_��� = 8
    C_���� = 9
    C_ҵ�� = 10
    C_Id = 11
End Enum

' ��ͼ��󵯳�����ѡ����
Private Enum PATI_COLUMN
    CI_ͼ��1 = 0
    CI_ͼ��2
    CI_ͼ��3
    CI_����
    CI_����ID
    CI_��ҳID
    CI_����
    CI_סԺ��
    CI_��Ժ����
    CI_��Ժ����
    CI_��������
End Enum

Private Enum PATI_TYPE
    pt�ҵ� = 1
    pt��Ժ = 2
    ptԤ�� = 3
    pt��Ժ = 4
    pt���� = 5
    pt���� = 6
    pt���ת�� = 7
End Enum

Private Enum Msg_Type '��Ϣ�������
    m�������� = 1
    mҽ������ = 2
    mΣ��ֵ = 3
    m���泷�� = 4
    mҽ����� = 5
    m������� = 6
    m��Ⱦ�� = 7
    m�����ʿ� = 8
    m��Ѫ��� = 9
    mУ������ = 10
    m��Ѫ��� = 11
    m��Ѫ��Ӧ = 12
End Enum

Private Type PatiInfo
    ״̬ As Integer '������ҳ.״̬
    Ӥ�� As Integer
    סԺ�� As String
    ���� As String
    ����ID As Long
    ��ҳID As Long
    ����ID As Long
    ����ID As Long
    ��Ժ���� As Date
    ��Ժ���� As Date
    ��Ŀ���� As Date
    סԺ���� As Long
    rsͼ�� As ADODB.Recordset
    ����ת�� As Boolean
End Type

'�Ӵ��������
Private mclsEMR As Object  '�°没��zlRichEMR.clsDockEMR
Private mclsDisease As zlRichEPR.cDockDisease
Private WithEvents mclsAdvices As zlPublicAdvice.clsDockInAdvices
Attribute mclsAdvices.VB_VarHelpID = -1
Private WithEvents mclsEPRs As zlRichEPR.cDockInEPRs
Attribute mclsEPRs.VB_VarHelpID = -1
Private WithEvents mclsTends As zlRichEPR.cDockInTends
Attribute mclsTends.VB_VarHelpID = -1
Private WithEvents mclsTendsNew As zl9TendFile.clsTendFile    '�°滤ʿ����վ
Attribute mclsTendsNew.VB_VarHelpID = -1
Private mclsTendEPRs As zlRichEPR.cDockInTendEPRs
Private WithEvents mclsPath As zlPublicPath.clsDockPath
Attribute mclsPath.VB_VarHelpID = -1
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mclsWardMonitor As clsWardMonitor     '�໤�ǽӿ�
Private mobjEPRDoc As zlRichEPR.cEPRDocument
Private mclsChildQuestion As zlRichEPR.clsChildQuestion '�������ӿ�
Private mobjKernel As zlPublicAdvice.clsPublicAdvice          '�ٴ����Ĳ���
Private WithEvents mclsDis As zl9Disease.clsDisease
Attribute mclsDis.VB_VarHelpID = -1

Private WithEvents mFrmConsultation As Form
Attribute mFrmConsultation.VB_VarHelpID = -1

Private WithEvents mclsInOutMedRec As zlMedRecPage.clsInOutMedRec  '��ҳ����
Attribute mclsInOutMedRec.VB_VarHelpID = -1
Private WithEvents mfrmResponse As frmAuditResponse '��鷴������
Attribute mfrmResponse.VB_VarHelpID = -1
Private WithEvents mfrmInView As frmInDoctorView    'סԺһ��
Attribute mfrmInView.VB_VarHelpID = -1
Private mcolSubForm As Collection
Private mfrmActive As Form
Private mobjSquareCard As Object      '���������
Private mstrCardKind As String        '��������󷵻صĿ��õ�ҽ�ƿ�
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsReg As zlPublicExpense.clsRegist

'�������ñ���
Private mintChange As Integer
Private mdtOutBegin As Date, mdtOutEnd As Date
Private mdtMeetBegin As Date, mdtMeetEnd As Date
Private mintNotify As Integer '���������Զ�ˢ�¼��(����)
Private mstrNotify As String  '������������
Private mintNotifyDay As Integer '���Ѷ���������ɵĲ���
Private mintDeptView As Integer '0-��������ʾ��1-��������ʾ
Private mintDeptViewBed As Integer '0��1-ֻ��ʾ�д�λ�Ĳ������߿���
Private mblnDeptViewEnabled As Boolean
Private mlngMedRedDay As Long   '������鷴������
Private mintMecStandard As Integer  '������ҳ��ʽ 0-��������׼��1-�Ĵ�ʡ��׼��2-����ʡ��׼,3-����ʡ��׼
Private mclsMsg As clsCISMsg
Private mrsMsg As ADODB.Recordset
Private mbln��Ϣ���� As Boolean
Private mblnΣ��ֵ���� As Boolean

Private mstrAllPatis As String '��ǰ�б��еĲ�����Ϣ����ʽ��"����ID:��ҳID,����ID:��ҳID,..."
Private mrsNotes As ADODB.Recordset '���˸���ͼ�꣬������ͼ��
Private mstrList���� As String
Private mrsPatiNotes As ADODB.Recordset 'ĳһ�����˵�ͼ��
Private mrsPati���� As ADODB.Recordset 'ͼ�����
Private mlngSource As Long '�����壬С����

Private Const conMenu_ͼ�� = 990050                     '��ע��ʹ�õ�ͼ��ID��990050��ʼ,���150��ͼ��
Private Const conMenu_��ע1 = 990200
Private Const conMenu_��ע2 = 990300
Private Const conMenu_��ע3 = 990400
Private Const conMenu_��ע���� = 990500

Private Const conIconAll = 50 '����ͼ����,�����ݶ�Ϊ50��

'�����������
Private mstrPrivs As String
Private mlngModul As Long
Private mPatiInfo As PatiInfo '��ʷסԺ��¼�е�,��һ��Ϊ��ǰ��
Private mlng����ID As Long, mlng��ҳID As Long '�����嵥�е�
Private mrsPati As ADODB.Recordset '������Ϣ���ϣ�����ͬһ���֤�ŵ����в���
Private mobjPatient As Object '������Ϣ������������֤���֤��
Private mintFindType As Integer '0-סԺ��,1-����,2-���￨,3-����
Private mstrFindType As String '�����洢��ǰ�������͵�����
Private mblnFindTypeEnabled As Boolean
Private mblnICU As Boolean '�Ƿ�Ǳ��Ƶ�ICU��
Private mblnOutDept As Boolean '�Ƿ������������Ŀ��ң��������۲�����ʾ����ţ�
Private mstrDiagInfo As String  '��ҳ����ʱ������ĵ�һ�����Ϣ
Private mstr����ID As String   '���ڴ��Ӵ����ȡ����ID
Private mstr���ID As String   '���ڴ��Ӵ����ȡ���ID
Private mblnReturn As Boolean       'cboDept�س�����
Private mintOutPreTime As Integer
Private mintMeetPreTime As Integer

Private mintPreDept As Integer
Private mstrPrePati As String
Private mintPrePage As Integer
Private mstrPreNotify As String
Private mblnUnRefresh As Boolean
Private mblnNoCheck  As Boolean '����ѡ��
Private mfrmParent As Object
Private mblnIsFindAgain As Boolean
Private mstrUserDeps As String '����Ա���������ַ���
Private mlngNewIndex As Long
Private mlngOldIndex As Long
Private mblnIsNot As Boolean
Private mbytSize As Byte '�����С 0-С���壨9�ţ�1-�����壨12�ţ�
Private mblnTabTmp As Boolean
Private mblnIsInit As Boolean
Private mblnInView As Boolean
Private mbln���ܻ��� As Boolean
Private mblnΣ��ֵ As Boolean '��Σ��ֵ��Ȩ��
Private mlng����ҽ��ID As Long
Private mbln�������� As Boolean

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
    '���¶�ȡ����
    Call LoadPatients
End Sub

Private Sub chk��������_Click(Index As Integer)
    Dim i As Integer, k As Integer
    
    If Not Visible Or mblnNoCheck Then Exit Sub
    
    If (GetKeyState(vbKeyControl) And &H8000) <> 0 Then
        'Ctrl������ѡ��
        mblnNoCheck = True
        For i = 0 To chk��������.UBound
            chk��������(i).Value = IIf(i = Index, 1, 0)
        Next
        mblnNoCheck = False
    Else
        '����ѡ��һ��
        For i = 0 To chk��������.UBound
            If chk��������(i).Value = 1 Then k = k + 1
        Next
        If k = 0 Then chk��������(Index).Value = 1
    End If
    
    '���¶�ȡ����
    Call LoadPatients
End Sub

Private Sub chkHZ_Click(Index As Integer)
    Dim i As Integer, k As Integer
    
    If Not Visible Or mblnNoCheck Then Exit Sub
    
    If (GetKeyState(vbKeyControl) And &H8000) <> 0 Then
        'Ctrl������ѡ��
        mblnNoCheck = True
        For i = 0 To chkHZ.UBound
            chkHZ(i).Value = IIf(i = Index, 1, 0)
        Next
        mblnNoCheck = False
    Else
        '����ѡ��һ��
        For i = 0 To chkHZ.UBound
            If chkHZ(i).Value = 1 Then k = k + 1
        Next
        If k = 0 Then chkHZ(Index).Value = 1
    End If
    
    '���¶�ȡ����
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
    
    cboSelectTime(1).Clear '��Ժ
    With cboSelectTime(1)
        .AddItem "������"
        .ItemData(.NewIndex) = 0
        .AddItem "������"
        .ItemData(.NewIndex) = 1
        .AddItem "ǰ����"
        .ItemData(.NewIndex) = 2
        .AddItem "һ����"
        .ItemData(.NewIndex) = 7
        .AddItem "30����"
        .ItemData(.NewIndex) = 30
        .AddItem "60����"
        .ItemData(.NewIndex) = 60
        .AddItem "[ָ��...]"
        .ItemData(.NewIndex) = -1
    End With
    If cboSelectTime(1).ListCount > 0 Then cboSelectTime(1).ListIndex = 0
    
    cboSelectTime(2).Clear '����
    With cboSelectTime(2)
        .AddItem "������"
        .ItemData(.NewIndex) = 0
        .AddItem "������"
        .ItemData(.NewIndex) = 1
        .AddItem "ǰ����"
        .ItemData(.NewIndex) = 2
        .AddItem "һ����"
        .ItemData(.NewIndex) = 7
        .AddItem "15����"
        .ItemData(.NewIndex) = 15
        .AddItem "[ָ��...]"
        .ItemData(.NewIndex) = -1
    End With
    If cboSelectTime(2).ListCount > 0 Then cboSelectTime(2).ListIndex = 1
End Sub

Private Sub cboSelectTime_Click(Index As Integer)
'Index 1��Ժ 2����
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    intDateCount = cboSelectTime(Index).ItemData(cboSelectTime(Index).ListIndex)
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If Index = 1 Then
        If cboSelectTime(Index).ListIndex = mintOutPreTime And intDateCount <> -1 Then Exit Sub
        If intDateCount = -1 Then
            If Not frmSelectTime.ShowMe(Me, mdtOutBegin, mdtOutEnd, cboSelectTime(1)) Then
                'ȡ��ʱ�ָ�ԭ����ѡ��
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
            cboSelectTime(Index).ToolTipText = "��Χ��" & Format(mdtOutBegin, "yyyy-MM-dd") & " �� " & Format(mdtOutEnd, "yyyy-MM-dd")
        End If
        '�����������֤ÿ���ط���ȡ�ĳ�Ժ���˶�����ͬһʱ�䷶Χ�ڣ�72783��
        Call zlDatabase.SetPara("��Ժ���˽������", DateDiff("d", datCurr, mdtOutEnd), glngSys, pסԺҽ��վ)
        Call zlDatabase.SetPara("��Ժ���˿�ʼ���", DateDiff("d", mdtOutBegin, datCurr), glngSys, pסԺҽ��վ)
        mintOutPreTime = cboSelectTime(Index).ListIndex
    ElseIf Index = 2 Then
        If cboSelectTime(Index).ListIndex = mintMeetPreTime And intDateCount <> -1 Then Exit Sub
        If intDateCount = -1 Then
            If Not frmSelectTime.ShowMe(Me, mdtMeetBegin, mdtMeetEnd, cboSelectTime(2)) Then
                'ȡ��ʱ�ָ�ԭ����ѡ��
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
            cboSelectTime(Index).ToolTipText = "��Χ��" & Format(mdtMeetBegin, "yyyy-MM-dd") & " �� " & Format(mdtMeetEnd, "yyyy-MM-dd")
        End If
        mintMeetPreTime = cboSelectTime(Index).ListIndex
    End If
    If Me.Visible = True Then Call LoadPatients
End Sub

Private Sub cmdRef_Click()
'����ת��������ˢ��
    Call txtChange_KeyPress(vbKeyReturn)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '����
    PatiIdentify.ActiveFastKey
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("[']", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If (InStr("0123456789", Chr(KeyAscii)) > 0 Or UCase(Chr(KeyAscii)) >= "A" And UCase(Chr(KeyAscii)) <= "Z") _
        And Not Me.ActiveControl Is PatiIdentify And mstrFindType = "���￨" Then
        PatiIdentify.Text = UCase(Chr(KeyAscii))
        
        PatiIdentify.SetFocus
        Call zlCommFun.PressKey(vbKeyRight)
    End If
End Sub

Private Sub Form_Load()
    Dim objPane As Pane, strTab As String, intIdx As Integer
    Dim blnCol As Boolean, intType As Integer, bln·��״̬ As Boolean
    Dim strTmp As String, i As Integer
    Dim arrTmp As Variant, objTabItem As TabControlItem
    Dim objTimeLine As Object
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    mblnNoCheck = False '����ѡ��
    
    mblnICU = False
    mintPreDept = -1
    mstrPrePati = ""
    mintPrePage = -1
    mstrPreNotify = ""
    
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, pסԺҽ��վ, GetInsidePrivs(pסԺҽ��վ))
    Call AddMipModule(mclsMipModule)
    Set mobjKernel = New zlPublicAdvice.clsPublicAdvice
    
    Set mclsDis = New zl9Disease.clsDisease
    Call mclsDis.InitDisease(gcnOracle, Me, glngSys, glngModul, mstrPrivs, mclsMipModule)
    
    Call GetLocalSetting '���ز���
    
    'ͼ��
    Call SetAllPatiͼ��
    
    '����ָ���Ĭ���������Ͷ�ȡ
    '-----------------------------------------------------
    mintFindType = Val(zlDatabase.GetPara("���˲��ҷ�ʽ", glngSys, pסԺҽ��վ, , , , intType))
    mblnFindTypeEnabled = Not ((intType = 3 Or intType = 15) And InStr(mstrPrivs, "��������") = 0)
    mintDeptViewBed = Val(zlDatabase.GetPara("����ʾ�޴�λ�Ĳ�������", glngSys, pסԺҽ��վ, , , , intType))
    mblnΣ��ֵ = InStr(GetInsidePrivs(pסԺҽ��վ), ";Σ��ֵ����;") > 0
    
    Set mclsReg = New zlPublicExpense.clsRegist
    Call mclsReg.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
    Call mclsReg.zlInitData(2)
    
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)
    mstrCardKind = "ס|סԺ��|0|0|0|0|0|0;��|����|0|0|0|0|0|0;��|���￨|0|0|8|0|0|0;��|����|0|0|0|0|0|0;��|���ۺ�|0|0|0|0|0|0;��|�������|0|0|0|0|0|0"
    On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    err.Clear: On Error GoTo 0
    If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
        Set mobjSquareCard = Nothing
        MsgBox "ҽ�ƿ�������zl9CardSquare����ʼ��ʧ��!", vbInformation, gstrSysName
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
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, IIf(mbytSize = 0, 310, 320), 400, DockLeftOf, Nothing)
    objPane.Title = "סԺ�����б�"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    Set objPane = Me.dkpMain.CreatePane(2, 310, 100, DockBottomOf, objPane)
    objPane.Title = "��Ϣ����"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    'TabControl
    '-----------------------------------------------------
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, True)
    If GetInsidePrivs(p�°�סԺ����, True) <> "" Then
        Set mclsEMR = DynamicCreate("zlRichEMR.clsDockEMR", "���Ӳ���")
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
        mcolSubForm.Add mclsEMR.zlGetForm, "_�²���"
    End If
    mcolSubForm.Add mclsPath.zlGetForm, "_·��"
    mcolSubForm.Add mclsAdvices.zlGetForm, "_ҽ��"
    mcolSubForm.Add mclsEPRs.zlGetForm, "_����"
    mcolSubForm.Add mclsTends.zlGetForm, "_����"
    mcolSubForm.Add mclsTendEPRs.zlGetForm, "_������"
    If mclsWardMonitor.Enabled Then
        mcolSubForm.Add mclsWardMonitor.zlGetForm, "_�໤"
    End If
    mcolSubForm.Add mclsTendsNew.zlGetForm, "_�°滤��"
    mcolSubForm.Add mclsDisease.zlGetForm, "_��������"
    
    If InStr(GetInsidePrivs(pסԺҽ��վ), "סԺһ��") > 0 Then
        Set objTimeLine = DynamicCreate("ZLSoft.BusinessHome.ClientControl.TimeLineBase.Control.TimeLineControl", "ʱ����", False)
        If Not objTimeLine Is Nothing Then
            Set mfrmInView = New frmInDoctorView
             mcolSubForm.Add mfrmInView, "_סԺһ��"
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
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        If Not mfrmInView Is Nothing Then
            .InsertItem(intIdx, "סԺһ��", picTmp.hwnd, 0).Tag = "סԺһ��": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p�ٴ�·��Ӧ��, True) <> "" Then
            .InsertItem(intIdx, "�ٴ�·��", picTmp.hwnd, 0).Tag = "·��": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(pסԺҽ���´�, True) <> "" Then
            .InsertItem(intIdx, "ҽ����Ϣ", picTmp.hwnd, 0).Tag = "ҽ��": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(pסԺ��������, True) <> "" Then
            .InsertItem(intIdx, "������Ϣ", picTmp.hwnd, 0).Tag = "����": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p�°�סԺ����, True) <> "" And Not mclsEMR Is Nothing Then
            .InsertItem(intIdx, "���Ӳ���", picTmp.hwnd, 0).Tag = "�²���": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p�����¼����, True) <> "" Then
            .InsertItem(intIdx, "������Ϣ", picTmp.hwnd, 0).Tag = "����": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
            mlngOldIndex = intIdx - 1
            .InsertItem(intIdx, "������Ϣ", picTmp.hwnd, 0).Tag = "�°滤��": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
            mlngNewIndex = intIdx - 1
            .InsertItem(intIdx, "������", picTmp.hwnd, 0).Tag = "������": intIdx = intIdx + 1
            .Item(intIdx - 1).Visible = False
        End If
        If GetInsidePrivs(p����������д, True) <> "" Then
            Set objTabItem = .InsertItem(intIdx, "��������", picTmp.hwnd, 0): objTabItem.Tag = "��������": objTabItem.Visible = False: intIdx = intIdx + 1
        End If
        If mclsWardMonitor.Enabled Then
            If InStr(GetInsidePrivs(pסԺҽ��վ), "����໤") > 0 Then
                .InsertItem(intIdx, "����໤", picTmp.hwnd, 0).Tag = "�໤": intIdx = intIdx + 1
            End If
        End If
        If gbln�������廤��ӿ� Then
            If InitNurseIntegrate = True Then
                If Not gobjNurseIntegrate Is Nothing Then
                    mcolSubForm.Add gobjNurseIntegrate.GetDocForm, "_������������"
                    .InsertItem(intIdx, "������������", mcolSubForm("_������������").hwnd, 0).Tag = "������������": intIdx = intIdx + 1
                    .Item(intIdx - 1).Visible = False
                End If
            End If
        End If
                
        '����ṩ�Ŀ�Ƭ
        Call CreatePlugInOK(pסԺҽ��վ)
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strTmp = gobjPlugIn.GetFormCaption(glngSys, pסԺҽ��վ)
            Call zlPlugInErrH(err, "GetFormCaption")
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ",")
                For i = 0 To UBound(arrTmp)
                    strTmp = arrTmp(i)
                    
                    mcolSubForm.Add gobjPlugIn.GetForm(glngSys, pסԺҽ��վ, strTmp), "_" & strTmp
                    .InsertItem(intIdx, strTmp, mcolSubForm("_" & strTmp).hwnd, 0).Tag = strTmp: intIdx = intIdx + 1
                    Call zlPlugInErrH(err, "GetForm")
                Next
            End If
            err.Clear: On Error GoTo 0
        End If
        
        If .ItemCount = 0 Then
            MsgBox "��û��ʹ��סԺҽ������վ��Ȩ�ޡ�", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        '�ָ��ϴ�ѡ��Ŀ�Ƭ
        strTab = zlDatabase.GetPara("ҽ������", glngSys, pסԺҽ��վ)
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
        'ֻ����ѡ����Ӵ���
        Call tbcSub_SelectedChanged(.Selected)
    End With
    
    '---------------------------------------------------
    'tbcPati�����б�
    With Me.tbcPati
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        .InsertItem(0, "����ס", picPatiIn.hwnd, 0).Tag = "����ס"
        .InsertItem(1, "��Ժ", picPatiIn.hwnd, 0).Tag = "��Ժ"
        .InsertItem(2, "��Ժ", picPatiIn.hwnd, 0).Tag = "��Ժ"
        .InsertItem(3, "ת��", picPatiIn.hwnd, 0).Tag = "ת��"
        .InsertItem(4, "����", picPatiIn.hwnd, 0).Tag = "����"
        
        .Item(4).Selected = True
        .Item(1).Selected = True
        '��λ����ѡ�
        tbcPati.Item(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(tbcPati), "tbcPati", 1)).Selected = True
    End With
    
    
    '������������
    Call InitReportColumn
    picPati.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    picInfo.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    
    '����ѡ��
    chk��������(0).BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    chk��������(1).BackColor = chk��������(0).BackColor
    chk��������(2).BackColor = chk��������(0).BackColor
    Call Cbo.SetListWidth(cbo����.hwnd, cbo����.Width * 2)
    
    '��ȡ��������
    Call Set��С��������ʾ
    
     'ת����������
    txtChange.Text = mintChange
    mintOutPreTime = -1
    mintMeetPreTime = -1
    Call InitSelectTime
    
    '��ʼ��סԺ����/����
    Call ReLoadDept
    
    'ȡ����Ա��������
    mstrUserDeps = GetUser����IDs(False)
    
    '��ʼ�����˹�������
    strTmp = zlDatabase.GetPara("��ǰ��������", glngSys, pסԺҽ��վ, "111", _
        Array(lbl��������, chk��������(0), chk��������(1), chk��������(2)), InStr(mstrPrivs, "��������") > 0)
    For i = 0 To chk��������.UBound
        chk��������(i).Value = IIf(Mid(strTmp, i + 1, 1) = "1", 1, 0)
    Next
    
    strTmp = zlDatabase.GetPara("���ﲡ�˹���", glngSys, pסԺҽ��վ, "011", Array(chkOut, chkHZ(0), chkHZ(1)), InStr(mstrPrivs, "��������") > 0)
    chkOut.Value = IIf(Mid(strTmp, 1, 1) = "1", 1, 0)
    chkHZ(0).Value = IIf(Mid(strTmp, 2, 1) = "1", 1, 0)
    chkHZ(1).Value = IIf(Mid(strTmp, 3, 1) = "1", 1, 0)
    If chkHZ(0).Value = 0 And chkHZ(1).Value = 0 Then
        chkHZ(0).Value = 1
        chkHZ(1).Value = 1
    End If
    
    '��С����ʾ
    strTmp = zlDatabase.GetPara("��С����ʾ", glngSys, pסԺҽ��վ, "0", Array(chkByTeam), InStr(mstrPrivs, "��������") > 0)
    chkByTeam.Value = IIf(strTmp = "1", 1, 0)
        
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(chkOutByTeam), "chkOutByTeam", "0")
    chkOutByTeam.Value = IIf(strTmp = "1", 1, 0)
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, "")
    End If
    blnCol = rptPati.Columns(col_���).Visible
    bln·��״̬ = rptPati.Columns(col_·��״̬).Visible
    Call RestoreWinState(Me, App.ProductName, , True)
    Me.WindowState = vbMaximized
    Call SetFixedCommandBar(cbsMain(2).Controls)
  
    rptPati.Columns(col_���).Visible = blnCol
    rptPati.Columns(col_·��״̬).Visible = bln·��״̬
    If bln·��״̬ And rptPati.Columns(col_·��״̬).Width = 0 Then rptPati.Columns(col_·��״̬).Width = 18

    '�����е��ã���λ��ָ���Ĳ���
    If Not mfrmParent Is Nothing Then
        If mfrmParent.frmHide Then Call LocatePati
    End If
    Call LoadNotify
End Sub

Private Function LocatePati(Optional ByVal strTag As String) As Boolean
'���ܣ���λ��ָ���Ĳ���
    Dim varCmd As Variant, i As Integer
    Dim lng����ID As Long, lng��ҳID As Long, lng����ID As Long
    Dim blnEnabled As Boolean
    Dim objRow As ReportRow
    Dim lngKey As Long
    
    If strTag <> "" Then
        For Each objRow In rptPati.Rows
            If objRow.GroupRow Then objRow.Expanded = True
            If Not objRow.GroupRow Then
                If InStr("_" & objRow.Record.Tag & "_", "_" & strTag & "_") > 0 Then
                    blnEnabled = timNotify.Enabled
                    timNotify.Enabled = False '������������ˢ����������
                    Set rptPati.FocusedRow = objRow 'ѡ��,��ʾ,[����Change�¼�]
                    timNotify.Enabled = blnEnabled
                    LocatePati = True: Exit Function
                End If
            End If
        Next
        Exit Function
    End If
    
    '��ȡ�����в����еĲ���ID����ҳID
    varCmd = Split(mfrmParent.GetCommand, " ")
    For i = LBound(varCmd) To UBound(varCmd)
        If UCase(varCmd(i)) Like "����ID=*" Then
            lng����ID = Val(Split(varCmd(i), "=")(1))
        ElseIf UCase(varCmd(i)) Like "��ҳID=*" Then
            lng��ҳID = Val(Split(varCmd(i), "=")(1))
        ElseIf UCase(varCmd(i)) Like "SINGLEPATI=*" Then
            lngKey = Val(Split(varCmd(i), "=")(1))
        End If
    Next
    mbln�������� = False
    If lng����ID <> 0 Then
        lng����ID = GetPatiDept(lng����ID, lng��ҳID, mintDeptView)
        If lng����ID <> 0 Then Call Cbo.Locate(cboDept, lng����ID, True)   '�������֮ǰ�Ŀ�����ͬ�����ᴥ��click�¼�
    
        For i = 0 To rptPati.Rows.Count - 1
            With rptPati.Rows(i)
                If Not .GroupRow Then
                    If .Record(col_����Id).Value = lng����ID And .Record(col_��ҳID).Value = lng��ҳID Then Exit For
                End If
            End With
        Next
    
        If i <= rptPati.Rows.Count - 1 Then
            '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
            Set rptPati.FocusedRow = rptPati.Rows(i)
            If rptPati.Visible Then rptPati.SetFocus
            If lngKey = 1 Then
            mbln�������� = True
            dkpMain.Panes(1).Closed = True
            dkpMain.Panes(2).Closed = True
            End If
        End If
        
        '��λ��ҽ����Ϣҳ
        For i = 0 To tbcSub.ItemCount - 1
            If tbcSub(i).Visible And tbcSub(i).Tag = "ҽ��" Then
                tbcSub.Item(i).Selected = True
            End If
        Next
    End If
End Function

Private Sub Set��С��������ʾ()
'���ܣ������Ƿ���ʾ��С����ʾ������
    Dim rsTmp As ADODB.Recordset, strSQL As String
        
    If InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0 Then
        strSQL = "Select 1 From �ٴ�ҽ��С�� Where rownum=1"
    Else
        strSQL = "Select 1 From ҽ��С����Ա Where ��Աid = [1] And Rownum = 1"
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
'���ܣ�������/������ʽ��ȡ���ò���
    lblDept.Caption = IIf(mintDeptView = 0, "����(&D)��", "����(&D)��")
   
    mintPreDept = -1
    Call InitDepts
    Call cboDept_Click
    
    If cboDept.ListIndex = -1 Then
        If InStr(mstrPrivs, "ȫԺ����") > 0 Then
            MsgBox "û�з���סԺ" & IIf(mintDeptView = 0, "����", "����") & "��Ϣ,���ȵ����Ź��������ã�", vbInformation, gstrSysName
        Else
            MsgBox "û�з���������" & IIf(mintDeptView = 0, "����", "����") & ",����ʹ��סԺҽ������վ��", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    
    '����Ǳ�ע�˵�,ִ���꼴�˳�
    If Control.ID > conMenu_��ע1 And Control.ID < conMenu_��ע���� Then
        Call SetPatiIcon(Control.Parameter)
        Exit Sub
    End If
    
    
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
            mlngSource = 999
            mbytSize = 0
            Call zlDatabase.SetPara("����", mbytSize, glngSys, pסԺҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_FontSize_L '������
        If mbytSize <> 1 Then
            mlngSource = 0
            mbytSize = 1
            Call zlDatabase.SetPara("����", mbytSize, glngSys, pסԺҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_Jump '��ת
        If Me.tbcSub.Selected.Index + 1 <= Me.tbcSub.ItemCount - 1 Then
            Me.tbcSub.Item(Me.tbcSub.Selected.Index + 1).Selected = True
        Else
            Me.tbcSub.Item(0).Selected = True
        End If
    Case conMenu_Manage_Bespeak 'ԤԼ�Һ�
        Control.Enabled = False
        Call mclsReg.zlExecuteCommandBars(Me, Control, "", mlng����ID)
        Control.Enabled = True
    Case conMenu_Edit_AppRequestManage, conMenu_Edit_AppRequest
        Control.Enabled = False
        Call mclsReg.zlExecuteCommandBars(Me, Control, "", mlng����ID)
        Control.Enabled = True
    Case conMenu_View_Option '"�Һ�ѡ������"
        Control.Enabled = False
        Call mclsReg.zlExecuteCommandBars(Me, Control, "")
        Control.Enabled = True
    Case conMenu_Tool_KssAudit '������ҩ���
        Call frmExamineKSS.ShowMe(Me, mclsMipModule)

    Case conMenu_Tool_OPSAudit '������˹���
        Call frmExamineOPS.ShowMe(Me, mclsMipModule)

    Case conMenu_Tool_OPSEmpower '������Ȩ����
        On Error Resume Next
        If gobjCISBase Is Nothing Then
            Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
            If gobjCISBase Is Nothing Then
                MsgBox "���ƻ�������(ZLCISBase)û����ȷ��װ���ù����޷�ִ�С�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        err.Clear: On Error GoTo 0
        Call gobjCISBase.CallOPSEmpower(Me, gcnOracle, glngSys, gstrDBUser)
    Case conMenu_Tool_TransAudit '��Ѫ��˹���
        On Error Resume Next
        Call frmExamineTransfuse.ShowMe(Me, 2, mclsMipModule)
    Case conMenu_Tool_CISMed  '�ٴ��Թ�ҩ
        Call Set�ٴ��Թ�ҩ(Me)
    Case conMenu_Tool_Archive '���Ӳ�������
        mblnUnRefresh = True
        Call frmArchiveView.ShowArchive(Me, mPatiInfo.����ID, mPatiInfo.��ҳID)
        mblnUnRefresh = False
    Case conMenu_Tool_ExaReport
        '���ó¸����Ǳ��ṩ�Ľӿ�
    Case conMenu_Tool_Reference_1 '������ϲο�
        mblnUnRefresh = True
        Call gobjKernel.ShowDiagHelp(vbModeless, Me)
        mblnUnRefresh = False
    Case conMenu_Tool_Reference_2 '���ƴ�ʩ�ο�
        mblnUnRefresh = True
        Call gobjKernel.ShowClincHelp(vbModeless, Me)
        mblnUnRefresh = False
    Case conMenu_Tool_MedRatio
        On Error Resume Next
        Call frmMedRatio.ShowMe(Me, mstrPrivs)
    Case conMenu_Edit_TraReactionRecord '��Ѫ��Ӧ
        Call FuncTraReactionRecord(Me, 1, pסԺҽ���´�)
    Case conMenu_Tool_MedRec '��ҳ����
        mblnUnRefresh = True
        Call ExecuteEditMediRec
        mblnUnRefresh = False
    Case conMenu_File_MedRecSetup '��ҳ��ӡ����
        Call PrintInMedRec(mclsInOutMedRec, 0, mPatiInfo.����ID, mPatiInfo.��ҳID, mobjReport, mPatiInfo.����ID, Me)
    Case conMenu_File_MedRecPreview '��ҳԤ��
        Call PrintInMedRec(mclsInOutMedRec, 1, mPatiInfo.����ID, mPatiInfo.��ҳID, mobjReport, mPatiInfo.����ID, Me)
    Case conMenu_File_MedRecPrint '��ҳ��ӡ
        Call PrintInMedRec(mclsInOutMedRec, 2, mPatiInfo.����ID, mPatiInfo.��ҳID, mobjReport, mPatiInfo.����ID, Me)
    Case conMenu_Tool_MeetIdea '��д/�鿴�������
        Call ExecuteMeetIdea(IIf(Control.Caption = "��д�������(&W)", 0, 1))
    Case conMenu_Tool_MeetOpen '���ܻ���
        Call Execute���ܻ���(IIf(Control.Caption = "���ܻ���(&O)", False, True))
    Case conMenu_Tool_MeetFinish '��ɻ���
        Call ExecuteMeetFinish
    Case conMenu_Tool_MeetCancel 'ȡ�����
        Call ExecuteMeetCancel
    Case conMenu_Tool_MedRecAuditSubmit '�ύ���
        '��Ժ���ˣ���δ�ύ��ܾ����״̬
        Call ExecuteMedRecAuditSubmit
    Case conMenu_Tool_MedRecAuditCancel 'ȡ���ύ
        '��Ժ���ˣ��Ѿ��ύ״̬
        Call ExecuteMedRecAuditCancel
    Case conMenu_Tool_MedRecAuditResponse '��鷴��
        '�����Ե��ã����ٿ��Բ鿴(��ǰ����ʷ)
        Call lbl���_Click
    Case conMenu_Tool_MedRecAuditWriteResponse '��д������
        If mclsChildQuestion Is Nothing Then
            Set mclsChildQuestion = New zlRichEPR.clsChildQuestion
        End If
        If Not mclsChildQuestion Is Nothing Then
            Call mclsChildQuestion.zlOpenQuestion(Me, mlng����ID, mlng��ҳID)
        End If
    Case conMenu_View_Find '����
        If Me.ActiveControl Is PatiIdentify Then
            PatiIdentify.SetFocus '��ʱ��Ҫ��λһ��
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
    Case conMenu_View_FindNext '������һ��
        If PatiIdentify.Text = "" Then
            PatiIdentify.SetFocus
        Else
            Call ExecuteFindPati(True)
        End If
    Case conMenu_View_Expend_CurCollapse '�۵���ǰ��
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                rptPati.SelectedRows(0).Expanded = False
            ElseIf Not rptPati.SelectedRows(0).ParentRow Is Nothing Then
                If rptPati.SelectedRows(0).ParentRow.GroupRow Then
                    rptPati.SelectedRows(0).ParentRow.Expanded = False
                End If
            End If
        End If
        '���۵���λ��������,�����Զ�������¼�
        Call rptPati_SelectionChanged
    Case conMenu_View_Expend_CurExpend 'չ����ǰ��
        If rptPati.SelectedRows.Count > 0 Then
            rptPati.SelectedRows(0).Expanded = True
        End If
    Case conMenu_View_Expend_AllCollapse '�۵�������
        For Each objRow In rptPati.Rows
            If objRow.GroupRow Then objRow.Expanded = False
        Next
        '���۵���λ��������,�����Զ�������¼�
        Call rptPati_SelectionChanged
    Case conMenu_View_Expend_AllExpend 'չ��������
        For Each objRow In rptPati.Rows
            If objRow.GroupRow Then objRow.Expanded = True
        Next
    Case conMenu_View_Notify '��������
        If rptNotify.Visible Then Call LoadNotify
    Case conMenu_View_Dept * 10# + 1, conMenu_View_Dept * 10# + 2 '������/������ʾ
        If mintDeptView <> Control.ID - conMenu_View_Dept * 10# - 1 Then
            mintDeptView = Control.ID - conMenu_View_Dept * 10# - 1
            Call zlDatabase.SetPara("������ʾ��ʽ", mintDeptView, glngSys, pסԺҽ��վ, InStr(";" & mstrPrivs & ";", ";��������;") > 0)
            
            Call ReLoadDept
        End If
    Case conMenu_View_Refresh 'ˢ��
        Call LoadPatients
        Call LoadNotify 'ˢ��ҽ������
         
    Case conMenu_File_Parameter '��������
        mblnUnRefresh = True
        frmInStationSetup.mstrPrivs = mstrPrivs
        frmInStationSetup.Show 1, Me
        If gblnOK Then
            Call GetLocalSetting
            Call LoadPatients
        End If
        mblnUnRefresh = False
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Help_Help '����
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit '�˳�
        Unload Me
    Case conMenu_Tool_HealthCard  '���񽡿���
        If Not mobjSquareCard Is Nothing Then
            Call mobjSquareCard.zlHealthArchivesShow(Me, p����ҽ��վ, mlng����ID, "")
        End If
    Case conMenu_Tool_Positive '���Խ���鿴
        Call mclsDis.ShowRegistByPati(Me, 1, mlng����ID, mlng��ҳID)
    Case conMenu_Tool_Critical
        Call ExecuteCritical
    Case Else
        mblnUnRefresh = True
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            'ִ�з�������ǰģ��ı���
            With mPatiInfo
                If Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_2" Or Split(Control.Parameter, ",")(1) = "ZL1_INSIDE_1261_3" Then
                    If cboDept.ListIndex = -1 Then
                        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
                    Else
                        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                            "��������=" & Split(cboDept.List(cboDept.ListIndex), "-")(1) & "|=" & CLng(cboDept.ItemData(cboDept.ListIndex)))
                    End If
                Else
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                        "����ID=" & .����ID, "��ҳID=" & .��ҳID, "סԺ��=" & .סԺ��, "���˿���=" & .����ID)
                End If
            End With
        ElseIf Between(Control.ID, conMenu_File_MedRecPrint * 100# + 1, conMenu_File_MedRecPrint * 100# + 6) Or Between(Control.ID, conMenu_File_MedRecPreview * 100# + 1, conMenu_File_MedRecPreview * 100# + 4) Then
            Call PrintInMedRec(mclsInOutMedRec, IIf(Between(Control.ID, conMenu_File_MedRecPrint * 100# + 1, conMenu_File_MedRecPrint * 100# + 6), 2, 1), mPatiInfo.����ID, mPatiInfo.��ҳID, mobjReport, mPatiInfo.����ID, Me, Val(Mid(Control.ID & "", Len(Control.ID & ""))))
        Else
            Select Case Me.tbcSub.Selected.Tag
            Case "·��"
               If rptPati.SelectedRows.Count >= 1 Then '��ѡ���в���ִ��rptPati.SelectedRows(0).GroupRow���ж�,����ᱨ��
                    If rptPati.SelectedRows(0).GroupRow = False Then
                        If rptPati.SelectedRows(0).Record(COL_Ӥ������ID).Value <> 0 Then
                            If rptPati.SelectedRows(0).Record(COL_Ӥ������ID).Value = cboDept.ItemData(cboDept.ListIndex) Or rptPati.SelectedRows(0).Record(COL_Ӥ������ID).Value = cboDept.ItemData(cboDept.ListIndex) Then
                                MsgBox "�ò����Ѿ�ת���������ˣ�ֻ��Ӥ�����ڱ����ң����������·����", vbInformation, Me.Caption
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                Call mclsPath.zlExecuteCommandBars(Control)
            Case "ҽ��"
                Call mclsAdvices.zlExecuteCommandBars(Control)
            Case "����"
                Call mclsEPRs.zlExecuteCommandBars(Control)
            Case "����"
                Call mclsTends.zlExecuteCommandBars(Control)
            Case "�°滤��"
                Call mclsTendsNew.zlExecuteCommandBars(Control)
            Case "������"
                Call mclsTendEPRs.zlExecuteCommandBars(Control)
            Case "�²���"
                Call mclsEMR.zlExecuteCommandBars(Control)
            Case "��������"
                Call mclsDisease.zlExecuteCommandBars(Control)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.ExeButtomClick(glngSys, pסԺҽ��վ, mcolSubForm("_" & tbcSub.Selected.Tag), tbcSub.Selected.Tag, Control.Caption, mlng����ID, mlng��ҳID, "")
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
            If .Count = 0 Then '��̬�Ӳ˵�,��1λ
                .Add xtpControlButton, conMenu_View_FindType * 100# + 1, "סԺ��(&1)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 2, "��  ��(&2)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 3, "���￨(&3)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 4, "��  ��(&4)"
            End If
        End With
    Case conMenu_File_MedRecPrint
        With CommandBar.Controls
            If .Count = 0 Then '��̬�Ӳ˵�,��1λ
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 1, "����(&1)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 2, "����(&2)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 3, "��ҳ1(&3)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 4, "��ҳ2(&4)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 5, "����+��ҳ1(&5)"
                .Add xtpControlButton, conMenu_File_MedRecPrint * 100# + 6, "����+��ҳ2(&6)"
            End If
        End With
    Case conMenu_File_MedRecPreview
        With CommandBar.Controls
            If .Count = 0 Then '��̬�Ӳ˵�,��1λ
                .Add xtpControlButton, conMenu_File_MedRecPreview * 100# + 1, "����(&1)"
                .Add xtpControlButton, conMenu_File_MedRecPreview * 100# + 2, "����(&2)"
                .Add xtpControlButton, conMenu_File_MedRecPreview * 100# + 3, "��ҳ1(&3)"
                .Add xtpControlButton, conMenu_File_MedRecPreview * 100# + 4, "��ҳ2(&4)"
            End If
        End With
    Case Else
       Select Case tbcSub.Selected.Tag
       Case "·��"
            Call mclsPath.zlPopupCommandBars(CommandBar)
       Case "ҽ��"
           Call mclsAdvices.zlPopupCommandBars(CommandBar)
       Case "����"
    
       Case "����"
       
       End Select
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    Dim i As Long
        
    If Not mblnIsInit Then
        mblnIsInit = True
        If Not mobjSquareCard Is Nothing Then
            Call PatiIdentify.zlInit(Me, glngSys, pסԺҽ��վ, gcnOracle, gstrDBUser, mobjSquareCard, mstrCardKind, "zl9CISJob")
            PatiIdentify.objIDKind.AllowAutoICCard = True
            PatiIdentify.objIDKind.AllowAutoIDCard = True
            If Not PatiIdentify.objIDKind.Cards Is Nothing Then
            For i = 0 To PatiIdentify.objIDKind.Cards.Count - 1
                If i = mintFindType Then
                    PatiIdentify.objIDKind.IDKind = i + 1
                    mstrFindType = PatiIdentify.objIDKind.Cards(i + 1).����
                    Exit For
                End If
            Next
            End If
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
    Case conMenu_View_Expend_CurExpend 'չ����ǰ��
        blnEnabled = False
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                blnEnabled = Not rptPati.SelectedRows(0).Expanded
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend_CurCollapse '�۵���ǰ��
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
    Case conMenu_View_Expend '�۵�/չ����
        Control.Enabled = rptPati.GroupsOrder.Count > 0 And rptPati.Rows.Count > 0
    Case conMenu_View_Dept * 10# + 1, conMenu_View_Dept * 10# + 2 '������/������ʾ
        Control.Checked = mintDeptView = Control.ID - conMenu_View_Dept * 10# - 1
        Control.Enabled = mblnDeptViewEnabled
    Case conMenu_Tool_KssAudit  '������ҩ���
        If GetInsidePrivs(p������ҩ���) = "" Then
            Control.Visible = False
        End If
    Case conMenu_Tool_OPSAudit  '������˹���
        If GetInsidePrivs(p������˹���) = "" Or Not gbln�����ּ����� Then
            Control.Visible = False
        End If
    Case conMenu_Tool_OPSEmpower  '������Ȩ����
        If GetInsidePrivs(p������Ȩ����) = "" Then
            Control.Visible = False
        End If
    Case conMenu_Tool_TransAudit '��Ѫ�ּ�����
        If GetInsidePrivs(p��Ѫ��˹���) = "" Or Not gbln��Ѫ�ּ����� Then
            Control.Visible = False
        End If
    Case conMenu_Tool_CISMed  '�ٴ��Թ�ҩ
        If InStr(GetInsidePrivs(pסԺҽ��վ), ";�ٴ��Թ�ҩ;") = 0 Then
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
    Case conMenu_Tool_Meet '���ﲡ��
        If InStr(mstrPrivs, "���ﲡ��") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = False
            If rptPati.SelectedRows.Count > 0 Then
                If Not rptPati.SelectedRows(0).GroupRow And rptPati.SelectedRows(0).Childs.Count = 0 Then
                    blnEnabled = Val(Mid(rptPati.SelectedRows(0).Record(col_����).Value, 1, 1)) = pt����
                End If
            End If
            Control.Enabled = blnEnabled
            If Me.Visible Then Control.Visible = tbcPati.Selected.Tag = "����"
        End If
    Case conMenu_Tool_MeetOpen '���ܻ���
        blnEnabled = False
        Control.Caption = IIf(mbln���ܻ���, "ȡ�����ܻ���(&X)", "���ܻ���(&O)")
        If rptPati.SelectedRows.Count > 0 Then
            If Not rptPati.SelectedRows(0).GroupRow And rptPati.SelectedRows(0).Childs.Count = 0 Then
                blnEnabled = True
            End If
        End If
        Control.Enabled = blnEnabled
        If Me.Visible Then Control.Visible = rptPati.SelectedRows(0).Record(col_ִ��״̬).Value = 0
    Case conMenu_Tool_MeetFinish, conMenu_Tool_MeetCancel '��ɻ���,ȡ�����
        blnEnabled = False
        If rptPati.SelectedRows.Count > 0 Then
            If Not rptPati.SelectedRows(0).GroupRow And rptPati.SelectedRows(0).Childs.Count = 0 Then
                blnEnabled = rptPati.SelectedRows(0).Record(col_ִ��״̬).Value = IIf(Control.ID = conMenu_Tool_MeetFinish, 0, 1)
            End If
        End If
        If Control.ID = conMenu_Tool_MeetFinish Then
            blnEnabled = blnEnabled And mbln���ܻ���
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Tool_MeetIdea
        blnEnabled = False
        If rptPati.SelectedRows.Count > 0 Then
            If Not rptPati.SelectedRows(0).GroupRow Then
                If Val(rptPati.SelectedRows(0).Record(COL_�������).Value) > 0 Then
                    blnEnabled = True
                    If Val(rptPati.SelectedRows(0).Record(col_ִ��״̬).Value) = 1 Then
                        Control.Caption = "�鿴�������(&V)"
                    Else
                        Control.Caption = "��д�������(&W)"
                    End If
                End If
            End If
        End If
        If Control.Caption = "��д�������(&W)" Then
            blnEnabled = blnEnabled And mbln���ܻ���
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Tool_MedRecAuditSubmit '�ύ���
        '��Ժ���ˣ���δ�ύ��ܾ����״̬
        If InStr(mstrPrivs, "��������ύ") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = False
            If Control.Caption = "�����ύ(&S)" Then Control.Caption = "�ύ���(&S)"
            If rptPati.SelectedRows.Count > 0 Then
                With rptPati.SelectedRows(0)
                    If Not .GroupRow Then
                        If (Int(Val(.Record(col_����).Value)) = pt��Ժ Or Int(Val(.Record(col_����).Value)) = pt����) Or (tbcPati.Selected.Tag = "��Ժ" And Val(.Record(col_����Id).Value) <> 0) Then
                            '��������Ժ��鷴��״̬����Ժ��δ�ύ��顣���ܾ������������ύ���
                            If .Record(col_ͼ��).Value <> 1 Or .Record(col_���).Value = 2 Then blnEnabled = True
                            If .Record(col_���).Value = 2 And Control.Caption = "�ύ���(&S)" Then Control.Caption = "�����ύ(&S)"
                        End If
                    End If
                End With
            End If
            Control.Enabled = blnEnabled
            If Me.Visible Then Control.Visible = tbcPati.Selected.Tag = "��Ժ"
        End If
    Case conMenu_Tool_MedRecAuditCancel 'ȡ���ύ
        '��Ժ���ˣ��Ѿ��ύ״̬
        If InStr(mstrPrivs, "��������ύ") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = False
            If rptPati.SelectedRows.Count > 0 Then
                With rptPati.SelectedRows(0)
                    If Not .GroupRow Then
                        If (Int(Val(.Record(col_����).Value)) = pt��Ժ Or Int(Val(.Record(col_����).Value)) = pt����) Or (tbcPati.Selected.Tag = "��Ժ" And Val(.Record(col_����Id).Value) <> 0) Then
                            If .Record(col_���).Value = 1 Then blnEnabled = True
                        End If
                    End If
                End With
            End If
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Tool_MedRecAuditResponse '��鷴��
        '�����Ե��ã����ٿ��Բ鿴(��ǰ����ʷ)
        Control.Enabled = rptPati.Rows.Count > 0
    Case conMenu_Tool_MedRecAuditWriteResponse '��д������
        Control.Enabled = rptPati.Rows.Count > 0 And InStr(GetInsidePrivs(p���Ӳ������), ";��鲡��;") <> 0
    Case conMenu_Tool_MedRatio
        If InStr(mstrPrivs, "ҩռ�Ȳ�ѯ") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_Tool_MedRec '��ҳ����
        If InStr(mstrPrivs, "��ҳ����") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = cboPages.ListIndex <> -1 And mPatiInfo.��ҳID > 0
        End If
    Case conMenu_File_MedRec '��ҳ��ӡ
        If InStr(mstrPrivs, "��ӡ��ҳ") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = cboPages.ListIndex <> -1
        End If
    Case conMenu_File_MedRecPreview, conMenu_File_MedRecPrint
        If InStr(mstrPrivs, "��ӡ��ҳ") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = cboPages.ListIndex <> -1
        End If
    Case conMenu_Edit_TraReactionRecord '��Ѫ��Ӧ
        Control.Visible = InStr(1, GetInsidePrivs(9005, , 2200), "��Ѫ��Ӧ�Ǽ�") <> 0
        Control.Enabled = Control.Visible And gblnѪ��ϵͳ
    Case conMenu_View_Notify '��������
        Control.Enabled = rptNotify.Visible
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
        '������ҩ����
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
        '��ҳ����
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
        Case "·��"
            Call mclsPath.zlUpdateCommandBars(Control)
        Case "ҽ��"
            Call mclsAdvices.zlUpdateCommandBars(Control)
        Case "����"
            Call mclsEPRs.zlUpdateCommandBars(Control)
        Case "����"
            Call mclsTends.zlUpdateCommandBars(Control)
        Case "�°滤��"
            Call mclsTendsNew.zlUpdateCommandBars(Control)
        Case "������"
            Call mclsTendEPRs.zlUpdateCommandBars(Control)
        Case "�²���"
            Call mclsEMR.zlUpdateCommandBars(Control)
        Case "��������"
            Call mclsDisease.zlUpdateCommandBars(Control)
        End Select
        
    End Select
End Sub

Private Sub SubWinDefCommandBar(ByVal objItem As TabControlItem)
'���ܣ�ˢ���Ӵ���˵���������
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
        
    Me.Caption = "סԺҽ������վ - " & objItem.Caption & "(��ǰ�û���" & UserInfo.���� & ")"
    If Not mbln�������� Then
    If InStr(mstrNotify, "1") > 0 Or Not Me.Visible Then
        dkpMain.Panes(2).Closed = False
        dkpMain.Panes(2).Hidden = Val(dkpMain.Panes(2).Tag) = 1
        dkpMain.Panes(2).Title = "��Ϣ����"
    Else
        dkpMain.Panes(2).Tag = IIf(dkpMain.Panes(2).Hidden, 1, 0)
        dkpMain.Panes(2).Close
    End If
    End If
    
    'ɾ�����ڵĹ������������˵���
    For lngCount = cbsMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.Count To 2 Step -1
        cbsMain(lngCount).Delete
    Next
    
    '���������¼���
    Call MainDefCommandBar
    
    If Not mclsReg Is Nothing And (InStr(GetInsidePrivs(9000), ";ԤԼ;") > 0 Or InStr(GetInsidePrivs(9000), ";ԤԼ�Ǽ�;") > 0) Then
        Call mclsReg.zlDefCommandBars(Me, Me.cbsMain, True)
    End If
    
    '�Ӵ������¼���
    Select Case objItem.Tag
    Case "·��"
        Call mclsPath.zlDefCommandBars(Me, Me.cbsMain, 0)
    Case "ҽ��"
        Call mclsAdvices.zlDefCommandBars(Me, Me.cbsMain, 0)
    Case "����"
        Call mclsEPRs.zlDefCommandBars(Me.cbsMain)
    Case "����"
        Call mclsTends.zlDefCommandBars(Me.cbsMain)
    Case "�°滤��"
        Call mclsTendsNew.zlDefCommandBars(Me.cbsMain)
    Case "������"
        Call mclsTendEPRs.zlDefCommandBars(Me.cbsMain)
    Case "�²���"
        Call mclsEMR.zlDefCommandBars(Me.cbsMain)
    Case "��������"
        Call mclsDisease.zlDefCommandBars(Me.cbsMain)
    Case "������������"
        
    Case Else
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strName = gobjPlugIn.GetButtomName(glngSys, pסԺҽ��վ, mcolSubForm("_" & objItem.Tag), objItem.Tag)
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

Private Sub SubWinRefreshData(ByVal objItem As TabControlItem)
'���ܣ�ˢ���Ӵ������ݼ�״̬
    Dim strInPatiNO As String, i As Integer
    Dim blnEdit As Boolean, lng·��״̬ As Long, bln���� As Boolean
    Dim lngType As Long, lng����ID As Long, lng����ID As Long
    Dim lng�������ID As Long, lngִ�п���ID As Long
    Dim lngState As TYPE_PATI_State, blnDis As Boolean

    If mlng����ID = 0 Or cboDept.ListIndex = -1 Then
        
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
        
        For i = 0 To tbcSub.ItemCount - 1 'Ĭ������������������ֲ���ʾ
            If tbcSub.Item(i).Tag = "������������" Then
                tbcSub.Item(i).Visible = False
                Exit For
            End If
        Next
        
        
        'Ҫ���Ӵ��尴�����ݴ������
        Select Case objItem.Tag
        Case "סԺһ��"
            Call mfrmInView.zlRefresh(0, 0, 0, 0)
        Case "·��"
            Call mclsPath.zlRefresh(0, 0, 0, 0, 0, False)
        Case "ҽ��"
            Call mclsAdvices.zlRefresh(0, 0, 0, 0, CDate(0), 0)
        Case "����"
            Call mclsEPRs.zlRefresh(0, 0, 0, False, False)
        Case "����"
            Call mclsTends.zlRefresh(0, 0, 0, False, True)
        Case "�°滤��"
            Call mclsTendsNew.zlRefresh(0, 0, 0, False, True, 0, 0, 1)
        Case "������"
            Call mclsTendEPRs.zlRefresh(0, 0, 0, False, False, False, True)
        Case "�໤"
            Call mclsWardMonitor.HideWindow
        Case "�²���"
            Call mclsEMR.zlRefresh(0, 0, 0, 0, 2)
        Case "��������"
            Call mclsDisease.zlRefresh(0, 0, 2, 0, False, False)
        Case "������������"
        
        Case Else
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                Call gobjPlugIn.RefreshForm(glngSys, pסԺҽ��վ, mcolSubForm("_" & objItem.Tag), objItem.Tag, 0, "", 0, False)
                Call zlPlugInErrH(err, "RefreshForm")
                err.Clear: On Error GoTo 0
            End If
        End Select
    Else
        With mPatiInfo
            lngType = Val(Mid(rptPati.SelectedRows(0).Record(col_����).Value, 1, 1))
            lng�������ID = cboDept.ItemData(cboDept.ListIndex)
            lng·��״̬ = Val(rptPati.SelectedRows(0).Record(col_·��״̬).Value)
            
            If lngType = pt���ת�� Then
                '��ȡ����ԭ���Ĳ����Ϳ���
                Call GetPatiLastChange(mlng����ID, mlng��ҳID, lng����ID, lng����ID)
                '������Կ��Ҳ鿴��ȡ��ǰ����;
                If mintDeptView = 0 Then
                    If cboDept.ListIndex <> -1 Then lng����ID = cboDept.ItemData(cboDept.ListIndex)
                End If
                lngState = ps���ת��
            Else
                lng����ID = .����ID
                lng����ID = .����ID
                
                If lngType = pt���� Then
                    lngִ�п���ID = rptPati.SelectedRows(0).Record(col_ִ�п���ID).Value
                    lngState = IIf(rptPati.SelectedRows(0).Record(col_ִ��״̬).Value = 0, ps����, ps����)
                Else
                    lngState = IIf(.��Ժ���� = CDate(0), IIf(.״̬ = 3, psԤ��, ps��Ժ), ps��Ժ)
                End If
            End If
            
            
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

            For i = 0 To tbcSub.ItemCount - 1 'Ĭ������������������ֲ���ʾ
                If tbcSub.Item(i).Tag = "������������" Then
                    tbcSub.Item(i).Visible = True
                    Exit For
                End If
            Next
            Select Case objItem.Tag
            Case "סԺһ��"
                Call mfrmInView.zlRefresh(.����ID, .��ҳID, lng����ID, .Ӥ��)
            Case "·��"
                Call mclsPath.zlRefresh(.����ID, .��ҳID, lng����ID, lng����ID, .״̬, .����ת��, True, lngState, , mclsMipModule)
            Case "ҽ��"
                If .״̬ = 1 Then '��Ժ����ס
                   If Val(zlDatabase.GetPara("���������ס�����´�ҽ��", glngSys, pסԺҽ���´�, 1)) = 0 Then
                        lngState = ps��ת�� 'lngState=ps��ת��ʱ�¿�ҽ���ȹ��ܲ�����
                   End If
                End If
                Call mclsAdvices.zlRefresh(.����ID, .��ҳID, lng����ID, lng����ID, lngState, .����ת��, , , lngִ�п���ID, lng·��״̬, _
                    cboDept.ItemData(cboDept.ListIndex), mclsMipModule, .Ӥ��, 0, mlng����ҽ��ID)
            Case "����"
                blnEdit = True
                With rptPati.SelectedRows(0)
                    If Int(lngType) = pt��Ժ Or Int(lngType) = pt���� Then
                        If Not (.Record(col_���).Value = 0 Or .Record(col_���).Value = 2 Or .Record(col_���).Value = 999) Then
                            '��������Ժ��鷴��״̬����Ժ��δ�ύ���
                            If .Record(col_ͼ��).Value = 1 Then blnEdit = False
                        End If
                    End If
                End With
                blnEdit = blnEdit And (lng�������ID = IIf(mintDeptView = 0, lng����ID, lng����ID) Or lngType = pt���� Or lngType = pt���ת��)
                '����ס�Ĳ��˲�����༭����
                blnEdit = blnEdit And tbcPati.Selected.Tag <> "����ס" And .����ID = mlng����ID
                'ҽ�����ٴ�·��������ɾ����Ӧ�Ĳ����ļ�������ǿ��ˢ��
                Call mclsEPRs.zlRefresh(.����ID, .��ҳID, IIf(mintDeptView = 0, lng�������ID, lng����ID), blnEdit, .����ת��, 0, True, lng����ID, lngState)
            Case "����"
                Call mclsTends.zlRefresh(.����ID, .��ҳID, lng����ID, False, True, lng����ID, lngState)
            Case "�°滤��"
                Call mclsTendsNew.zlRefresh(.����ID, .��ҳID, lng����ID, False, True, lng����ID, lngState, 1)
            Case "������"
                Call mclsTendEPRs.zlRefresh(.����ID, .��ҳID, lng����ID, False, False, .����ת��, True)
            Case "�໤"
                strInPatiNO = Trim(rptPati.SelectedRows(0).Record(col_סԺ��).Value)
                If strInPatiNO = "" Then
                    Call mclsWardMonitor.HideWindow
                Else
                    Call mclsWardMonitor.ShowInfor(strInPatiNO)
                End If
            Case "�²���"
                Call mclsEMR.zlRefresh(.����ID, .��ҳID, IIf(mintDeptView = 0 Or lngType = pt����, cboDept.ItemData(cboDept.ListIndex), lng����ID), lngState, 2)
            Case "��������"
                If objItem.Visible Then
                    '����ס�Ĳ��˲�����༭����
                    Call mclsDisease.zlRefresh(.����ID, .��ҳID, 2, IIf(mintDeptView = 0, cboDept.ItemData(cboDept.ListIndex), lng����ID), .����ת��, .����ID = mlng����ID And tbcPati.Selected.Tag <> "����ס", lngState)
                End If
            Case "������������"
                If Not gobjNurseIntegrate Is Nothing Then
                    If objItem.Visible Then
                        Call mcolSubForm("_������������").zlRefresh(.����ID, .��ҳID)
                    End If
                End If
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.RefreshForm(glngSys, pסԺҽ��վ, mcolSubForm("_" & objItem.Tag), objItem.Tag, .����ID, "", .��ҳID, .����ת��, _
                        lngִ�п���ID, cboDept.ItemData(cboDept.ListIndex), lng����ID, lng����ID, , lngState, , lng·��״̬)
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
            .Add xtpControlSplitButtonPopup, conMenu_File_MedRecPreview, "Ԥ����ҳ(&V)", -1, False
            .Add xtpControlSplitButtonPopup, conMenu_File_MedRecPrint, "��ӡ��ҳ(&P)", -1, False
        End With
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True '����
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
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Dept, "������ʾ(&D)") '����
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Dept * 10# + 1, "��������ʾ(&D)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Dept * 10# + 2, "��������ʾ(&U)", -1, False)
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)") '����
        objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FontSize, "�����С(&N)") '����
        objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_FontSize_S, "С����(&S)", -1, False '����
            .Add xtpControlButton, conMenu_View_FontSize_L, "������(&L)", -1, False '����
        End With

        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "չ��/�۵���(&X)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "�۵�������(&L)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "չ��������(&X)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "�۵���ǰ��(&C)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "չ����ǰ��(&E)", -1, False)
        End With
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FindType, "���ҷ�ʽ(&Y)"): objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "������һ��(&N)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True '����
        Set objControl = .Add(xtpControlButton, conMenu_View_Jump, "������ת(&J)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", -1, False)
    objMenu.ID = conMenu_ToolPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_KssAudit, "������ҩ���(&K)")
        objControl.IconId = 3551
        Set objControl = .Add(xtpControlButton, conMenu_Tool_OPSEmpower, "������Ȩ����(&N)")
        objControl.IconId = 3553
        Set objControl = .Add(xtpControlButton, conMenu_Tool_OPSAudit, "������˹���(&L)")
        objControl.IconId = 3551
        Set objControl = .Add(xtpControlButton, conMenu_Tool_TransAudit, "��Ѫ��˹���(&M)")
        objControl.IconId = 3551
        Set objControl = .Add(xtpControlButton, conMenu_Tool_CISMed, "�ٴ��Թ�ҩ(&J)")
        objControl.IconId = 3901
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "���Ӳ�������(&I)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_ExaReport, "��������ܼ챨��")
            objControl.IconId = conMenu_File_Preview
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "���ϲο�(&R)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_Reference_1, "������ϲο�(&D)", -1, False
            .Add xtpControlButton, conMenu_Tool_Reference_2, "���ƴ�ʩ�ο�(&C)", -1, False
        End With
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRatio, "ҩռ�Ȳ�ѯ")
            objControl.IconId = 813: objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRec, "��ҳ����(&M)"): objControl.BeginGroup = True
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_MedRecAudit, "�������(&Q)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRecAuditWriteResponse, "��д������", -1, False)
                objControl.IconId = conMenu_Tool_Sign
            Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRecAuditSubmit, "�ύ���(&S)", -1, False)
                objControl.IconId = conMenu_Manage_Complete
            Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRecAuditCancel, "ȡ���ύ(&C)", -1, False)
                objControl.IconId = conMenu_Edit_Untread
            Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRecAuditResponse, "��鷴��(&R)", -1, False)
                objControl.IconId = 3814
                objControl.BeginGroup = True
                objControl.ToolTipText = "�����鿴������鷴��"
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
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Meet, "���˻���(&E)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_MeetOpen, "���ܻ���(&O)", -1, False
            .Add xtpControlButton, conMenu_Tool_MeetIdea, "��д�������(&W)", -1, False
            .Add xtpControlButton, conMenu_Tool_MeetFinish, "��ɻ���(&F)", -1, False
            .Add xtpControlButton, conMenu_Tool_MeetCancel, "ȡ�����(&C)", -1, False
        End With
        If Not mobjSquareCard Is Nothing Then
            On Error Resume Next
            If mobjSquareCard.zlHealthArchiveIsSHow(Me, pסԺҽ��վ, strFunName, "") Then
                If err.Number = 0 Then
                    Set objControl = .Add(xtpControlButton, conMenu_Tool_HealthCard, strFunName)
                    objControl.BeginGroup = True
                    objControl.IconId = 3208
                End If
            End If
            On Error GoTo 0
        End If
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


    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��") '����
            
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Tool_MedRec, "��ҳ", objControl.Index + 1)
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlSplitButtonPopup, conMenu_File_MedRecPreview, "Ԥ����ҳ")
                objControl.IconId = conMenu_File_Preview
            Set objControl = .Add(xtpControlSplitButtonPopup, conMenu_File_MedRecPrint, "��ӡ��ҳ")
                objControl.IconId = conMenu_File_Print
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "��������")
            objControl.ToolTipText = "���Ӳ�������"

        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRecAuditSubmit, "�ύ")
            objControl.IconId = conMenu_Manage_Complete
            objControl.ToolTipText = "�ύ�������"
        
        Set objPopup = .Add(xtpControlPopup, conMenu_Tool_Meet, "����")
        objPopup.ID = conMenu_Tool_Meet
        objPopup.IconId = conMenu_Tool_Meet
        objPopup.Style = xtpButtonIconAndCaption
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_MeetOpen, "���ܻ���", -1, False
            .Add xtpControlButton, conMenu_Tool_MeetIdea, "��д�������(&W)", -1, False
            .Add xtpControlButton, conMenu_Tool_MeetFinish, "��ɻ���(&F)", -1, False
            .Add xtpControlButton, conMenu_Tool_MeetCancel, "ȡ�����(&C)", -1, False
        End With

        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����") '����
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�") '����
    End With
    
    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
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
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL1_INSIDE_1261_1", "ZL1_INSIDE_1261_4", "ZL1_INSIDE_1261_5", "ZL1_INSIDE_1261_6", "ZL1_INSIDE_1261_7", "ZL1_INSIDE_1261_8", _
                "ZL1_INSIDE_1261_9", "ZL1_INSIDE_1261_10")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnInView And UnloadMode = vbFormControlMenu Then
        Cancel = 1 'סԺһ���Ŵ�״̬,ȡ������ر�
        mfrmInView.zlExecuteCommandBars
    End If
End Sub

Private Sub mFrmConsultation_Unload(Cancel As Integer)
    '��������ˢ��
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
         '�Ŵ�
            SetParent mfrmInView.hwnd, Me.hwnd
            With mfrmInView
                .Tag = .Left & "," & .Top & "," & .Width & "," & .Height
                .Left = 0: .Width = Me.ScaleWidth
                .Top = 0: .Height = Me.ScaleHeight - Me.stbThis.Height
            End With
            mblnInView = True
        Else
        '��С
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

Private Sub mfrmInView_ViewPACSImage(ByVal ҽ��ID As Long)
'���ܣ�PACS��Ƭ����
    If CreateObjectPacs(gobjPublicPacs) Then
        Call gobjPublicPacs.ShowImage(ҽ��ID, Me, mPatiInfo.����ת��)
    End If
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    Dim lngPatiID As Long
    
    If objHisPati Is Nothing Then
        lngPatiID = 0
    Else
        lngPatiID = objHisPati.����ID
    End If
    
    If chkFilter.Value = 1 And chkFilter.Visible = True Then
        Call LoadPatients
    Else
        Call ExecuteFindPati(False, lngPatiID)
    End If
End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnIsInit = True Then mintFindType = Index - 1: mstrFindType = objCard.����
    If tbcPati.ItemCount <> 0 Then
        chkFilter.Visible = IIf(mstrFindType = "סԺ��" And tbcPati.Selected.Tag = "��Ժ", True, False)
        Call picPati_Resize
    End If
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub lbl���_Click()
    If cboDept.ListIndex = -1 Then Exit Sub
    
    '��ģ̬��ʾ��鷴������
    If mfrmResponse Is Nothing Then
        Set mfrmResponse = New frmAuditResponse
    End If
    mblnUnRefresh = True
    Call mfrmResponse.ShowMe(Me, cboDept.ItemData(cboDept.ListIndex), mintDeptView, mblnICU, 0, mstrPrivs)
    mblnUnRefresh = False
End Sub

Private Sub mclsAdvices_EditDiagnose(ParentForm As Object, ByVal ����ID As Long, ByVal ��ҳID As Long, ByVal ����ID As Long, ByVal str���� As String, Succeed As Boolean)
'���ܣ�¼�����
    If mclsInOutMedRec Is Nothing Then
        Set mclsInOutMedRec = New zlMedRecPage.clsInOutMedRec
        Call mclsInOutMedRec.InitMedRec(gcnOracle, glngSys, pסԺҽ��վ, mclsMipModule, gobjCommunity, gclsInsure)
    End If
    mblnUnRefresh = True
    
    If Not mclsInOutMedRec.ShowInMedRecEdit(ParentForm, ����ID, ��ҳID, ����ID, rptPati.SelectedRows(0).Record(col_·��״̬).Value, str����, mstrPrivs, , False) Then
        Succeed = False
    Else
        Succeed = mclsInOutMedRec.IsDiagInput
    End If
    mblnUnRefresh = False
End Sub

Private Sub mclsAdvices_RequestRefresh(ByVal RefreshNotify As Boolean)
'���ܣ�ҽ���Ӵ���Ҫ��ˢ��
    If Not RefreshNotify Then 'ע��Ҫ�ж�
        Call LoadPatients
    ElseIf rptNotify.Visible Then
        '��ˢ��ҽ����������
        Call LoadNotify
    End If
End Sub

Private Sub mclsAdvices_StatusTextUpdate(ByVal Text As String)
'���ܣ�ҽ���Ӵ���Ҫ�����״̬��
    Dim strTmp As String
    Dim intTmp As Long
    If Text = "" And rptPati.SelectedRows.Count > 0 Then
        With rptPati.SelectedRows(0)
            If Not .GroupRow Then
                If Val(.Record(col_����Id).Value) <> 0 Then intTmp = 1
            End If
            If intTmp = 1 Then
                stbThis.Panels(2).Text = IIf(stbThis.Panels(2).Tag = "", "", stbThis.Panels(2).Tag)
                lblFee(1).Caption = GetPati������Ϣ(mlng����ID, mlng��ҳID) & IIf(InStr(mstrPrivs, "ҩռ�Ȳ�ѯ") = 0, "", GetסԺ����ҩռ��(mlng����ID, mlng��ҳID))
                '��Ժ���˲���ʾ��Һ��
                If mPatiInfo.��Ժ���� = CDate(0) Then
                    lblFluid(0).Visible = True
                    lblFluid(1).Visible = True
                    strTmp = Get������Һ��(mlng����ID, mlng��ҳID)
                    lblFluid(1).Caption = "����" & Split(strTmp, ",")(0) & "ml,����" & Split(strTmp, ",")(1) & "ml"
                Else
                    lblFluid(0).Visible = False
                    lblFluid(1).Visible = False
                End If
                intTmp = Get����ҽ����ӡ(mlng����ID, mlng��ҳID)
                lblPrint(1).Caption = IIf(intTmp = 0, "δ��ӡ", IIf(intTmp = 1, "���ִ�ӡ", "ȫ����ӡ"))
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

Private Sub mclsAdvices_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
'���ܣ��鿴���Ӳ�������
    Call gobjRichEPR.ViewDocument(Me, ����ID, CanPrint)
End Sub

Private Sub mclsEPRs_Activate()
    If timNotify.Enabled And rptNotify.Visible Then
        '��ˢ��ҽ����������(�Զ�ˢ��ʱ)
        Call LoadNotify
    End If
End Sub

Private Sub mclsInOutMedRec_Closed(ByVal blnEditCancel As Boolean, ByVal str����ID As String, ByVal str���ID As String, ByVal strTag As String)
'����: �����ҳ�����ˢ������
'   EditCancel=ȡ���˳���ҳ
' strTag=������Ϣ�����ڴ洢���ﲡ����Ƭ�ļ���·�����Ժ���չʱ����|�ָ�
    Dim strSQL As String
    Dim lng·��ID As Long, lng·��״̬ As Long, lng����ID As Long, lng���ID As Long
    Dim rsPath As ADODB.Recordset, rsTmp As ADODB.Recordset, rsNext As ADODB.Recordset
    Dim bln��ҽ As Boolean
    Dim i As Long, blnNo As Boolean
    Dim str����IDs As String
    Dim str���IDs As String
    Dim objControl As CommandBarControl
    Dim blnNotView As Boolean
    
    On Error GoTo errH
    
    If Not blnEditCancel Then
        If mPatiInfo.��ҳID <> mlng��ҳID Then Exit Sub
        If InStr(";" & GetPrivFunc(glngSys, p����������д) & ";", ";������д;") > 0 Then
            '�鿴һ�������Ƿ���д����Ⱦ�����濨
            '�������Ƿ���Ҫ��д��Ⱦ�����濨������Ҫ���˳�
            '�����Ҫ��д�����һ�����Ƿ����ظ���д�ģ�û�о���д
            '�еĻ��͵���һ�����ظ���д����
            '����д�Ļ�ֱ���˳�����д�Ļ�����д
            mclsInOutMedRec.Hide
            Set rsTmp = mclsDisease.SatisfyEditDiseaseDoc(mlng����ID, mlng��ҳID, mPatiInfo.����ID, str����ID, str���ID)
            If Not rsTmp Is Nothing Then
                If rsTmp.RecordCount > 0 Then
                    If Not mclsDis.ShowDiseaseStation(Me, mlng����ID, mlng��ҳID, 2, mPatiInfo.����ID, str����ID, str���ID, blnNotView) Then
                        Call mclsDisease.EditDiseaseReport(Me, rsTmp, mlng����ID, mlng��ҳID, 2, mPatiInfo.����ID, blnNo)
                        If blnNo Then
                            Call mclsDis.EditNotFillReason(Me, mlng����ID, mlng��ҳID, 2)
                        End If
                    ElseIf blnNotView Then
                        Call mclsDisease.EditDiseaseReport(Me, rsTmp, mlng����ID, mlng��ҳID, 2, mPatiInfo.����ID, blnNo)
                        If blnNo Then
                            Call mclsDis.EditNotFillReason(Me, mlng����ID, mlng��ҳID, 2)
                        End If
                    End If
                End If
            End If
        End If
        
        Call LoadPatients
        '����·��
        lng·��״̬ = rptPati.SelectedRows(0).Record(col_·��״̬).Value
        '1.·��״̬-δ����
        '2.����-�����ٴ�·���Ŀ���
        '3.���������Ϣ�����仯
        '4.����״̬-������Ժ
        '5.��ǰ�û�ӵ�е���·����Ȩ��
        '��������5������ʱ�ŵ���·�����봰��
        If lng·��״̬ = -1 And mclsInOutMedRec.IsDiagChange _
            And mPatiInfo.״̬ = 0 And InStr(GetInsidePrivs(p�ٴ�·��Ӧ��), ";����·��;") <> 0 Then
            
            If HavePath(mPatiInfo.����ID) Then
                Set rsTmp = Get����ID(mlng����ID, mlng��ҳID, mPatiInfo.����ID, bln��ҽ)
                If bln��ҽ Then
                    rsTmp.Filter = "������� = 12 or ������� = 2 " 'ȡ��Ժ���
                    If rsTmp.RecordCount = 0 Then Exit Sub
                    For i = 1 To rsTmp.RecordCount
                        lng����ID = Val("" & rsTmp!����id)
                        lng���ID = Val("" & rsTmp!���id)
                        Set rsPath = GetPathTable(lng����ID, lng���ID, mPatiInfo.����ID)
                        If rsPath.RecordCount > 0 Then Exit For
                        rsTmp.MoveNext
                    Next
                Else
                    If rsTmp.RecordCount > 0 Then
                        lng����ID = Val("" & rsTmp!����id)
                        lng���ID = Val("" & rsTmp!���id)
                    End If
                    Set rsPath = GetPathTable(lng����ID, lng���ID, mPatiInfo.����ID)
                End If
                
                If rsPath.RecordCount = 0 Then
                    Set rsNext = Get����ID(mlng����ID, mlng��ҳID, mPatiInfo.����ID, , 1)
                    If rsNext.RecordCount = 0 Then Exit Sub
                End If
                
                Call mclsInOutMedRec.Hide '������ҳ
                Call mclsPath.zlRefresh(mlng����ID, mlng��ҳID, mPatiInfo.����ID, mPatiInfo.����ID, mPatiInfo.״̬, False, True)
                Call mclsPath.zlImportPath
                Call LoadPatients '��������ˢ��
                If rptPati.Visible Then rptPati.SetFocus
            End If
        '�кϲ�·���Ļ��͵���ϲ�·����mPatiInfo.״̬ =0:����סԺ��lng·��״̬ = 1:·������ִ����
        '�ҳ��Ѿ������·����
        ElseIf lng·��״̬ = 1 And mclsInOutMedRec.IsDiagChange And mPatiInfo.״̬ = 0 And InStr(GetInsidePrivs(p�ٴ�·��Ӧ��), ";����·��;") <> 0 Then
            If HavePath(mPatiInfo.����ID) Then
                '�Ѿ������·��
                strSQL = "Select a.ID,a.·��ID,A.�ϲ�·������,c.·��ID as ԭ·��ID,a.�汾��,a.״̬,a.��ǰ�׶�ID,a.��ǰ����,b.���� as δ����ԭ��,c.��ID,c.��֧ID,d.��֧ID as ǰһ�׶η�֧ID,e.����·������,a.�ϲ�·������,e.���� as ·������,a.������,a.����ʱ��,a.����ʱ��" & _
                        " From �����ٴ�·�� A,���쳣��ԭ�� B,�ٴ�·���׶� C,�ٴ�·���׶� D,�ٴ�·��Ŀ¼ E" & _
                        " Where a.����ID = [1] And a.��ҳID = [2] And a.·��ID=e.id And a.δ����ԭ�� = b.����(+) And a.��ǰ�׶�ID = c.ID(+) And a.ǰһ�׶�ID=d.id(+)" & _
                        " Order By a.����ʱ�� Desc"

                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, mPatiInfo.����ID)
                '��û�е���·���Ļ���ֱ���˳�
                If rsTmp.RecordCount = 0 Then
                    Exit Sub
                Else
                    lng·��ID = NVL(rsTmp!ID)
                    '�ϲ�·���Ѿ�5���˵Ļ�ֱ���˳�
                    If Val(rsTmp!�ϲ�·������ & "") >= 5 Then Exit Sub
                End If

                strSQL = "Select 1 From ����·������ A, ����·��ִ�� B" & vbNewLine & _
                         "Where a.·����¼id = b.·����¼id And a.·����¼id = [1] And A.�׶�id = [2] And b.���� = [3] And a.���� = b.����"

                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng·��ID, rsTmp!��ǰ�׶�ID & "", rsTmp!��ǰ���� & "")
                '��û�������Ļ�ֱ���˳�
                If rsTmp.RecordCount = 0 Then
                    Exit Sub
                End If

                Set rsTmp = Get����ID(mlng����ID, mlng��ҳID, mPatiInfo.����ID, , 2)
                If rsTmp.RecordCount = 0 Then Exit Sub

                '����Ǻϲ�·�����������зǵ��벡�ֵ�������ϻ򲢷�֢
                Do While Not rsTmp.EOF
                    If Val(rsTmp!����id & "") <> 0 Then
                        str����IDs = str����IDs & "," & rsTmp!����id
                    End If
                    If Val(rsTmp!���id & "") <> 0 Then
                        str���IDs = str���IDs & "," & rsTmp!���id
                    End If
                    rsTmp.MoveNext
                Loop
                rsTmp.MoveFirst
                str����IDs = Mid(str����IDs, 2)
                str���IDs = Mid(str���IDs, 2)

                '�����Distinct����Ϊ�����id�ͼ���id���˰󶨶�Ӧ�����ԣ�����������ظ�ֵ���ſ��Ѿ������˵ĺϲ�·��
                strSQL = "Select Distinct a.Id, a.����, a.����, a.����, a.˵��, Nvl(a.���ò���,'ͨ��') ���ò���, a.�����Ա�, a.��������, a.���°汾, c.��׼סԺ��,Nvl(a.��������,'��') as ��������,Nvl(a.ȷ������,0) as ȷ������,b.����ID,b.���ID" & vbNewLine & _
                        "From �ٴ�·��Ŀ¼ A, �ٴ�·������ B,�ٴ�·���汾 C" & vbNewLine & _
                        "Where a.Id = b.·��id And (instr(',' || [2] || ',',',' || b.����ID || ',')>0 and [2] is not null Or instr(',' || [4] || ',',',' || b.���ID || ',')>0 and [4] is not null)  And a.���°汾 is not null And a.id = b.·��ID And a.���°汾 = c.�汾��" & vbNewLine & _
                        "And a.Id = c.·��id And a.����=1 And b.����=0 And (a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From �ٴ�·������ D Where a.Id = d.·��id And d.����id = [1]))" & _
                        " And Not Exists(Select 1 From ���˺ϲ�·�� D Where a.id=d.·��ID  and d.��Ҫ·����¼ID=[3])"

                Set rsPath = zlDatabase.OpenSQLRecord(strSQL, "��ȡ·��Ŀ¼", mPatiInfo.����ID, str����IDs, lng·��ID, str���IDs)
                If rsPath.RecordCount = 0 Then Exit Sub

                Call mclsInOutMedRec.Hide     '������ҳ

                Set objControl = cbsMain.FindControl(, conMenu_Edit_ImportMerge, True, True)
                If Not objControl Is Nothing Then
                     Call mclsPath.zlExecuteCommandBars(objControl)
                End If

                Call mclsPath.zlRefresh(mlng����ID, mlng��ҳID, mPatiInfo.����ID, mPatiInfo.����ID, mPatiInfo.״̬, False, True)
                Call LoadPatients '��������ˢ��
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
    Dim lng·��״̬ As Long, lng����ID As Long, lng���ID As Long
    Dim rsTmp As ADODB.Recordset
    Dim rsPath As ADODB.Recordset, rsNext As ADODB.Recordset
    Dim bln��ҽ As Boolean
    Dim i As Long
    
    lng·��״̬ = rptPati.SelectedRows(0).Record(col_·��״̬).Value
    '1.·��״̬-δ����
    '2.����-�����ٴ�·���Ŀ���
    '3.���������Ϣ�����仯
    '4.����״̬-������Ժ
    '5.��ǰ�û�ӵ�е���·����Ȩ��
    '��������5������ʱ�ŵ���·�����봰��
    If lng·��״̬ = -1 And mPatiInfo.״̬ = 0 And InStr(GetInsidePrivs(p�ٴ�·��Ӧ��), ";����·��;") <> 0 Then
        If HavePath(mPatiInfo.����ID) Then
            Set rsTmp = Get����ID(mlng����ID, mlng��ҳID, mPatiInfo.����ID, bln��ҽ)
            If bln��ҽ Then
                rsTmp.Filter = "������� = 2 OR ������� = 12 "   'ȡ��Ժ���
                If rsTmp.RecordCount = 0 Then Exit Sub
                For i = 1 To rsTmp.RecordCount
                    lng����ID = Val("" & rsTmp!����id)
                    lng���ID = Val("" & rsTmp!���id)
                    Set rsPath = GetPathTable(lng����ID, lng���ID, mPatiInfo.����ID)
                    If rsPath.RecordCount > 0 Then Exit For
                    rsTmp.MoveNext
                Next
            Else
                If rsTmp.RecordCount > 0 Then
                    lng����ID = Val("" & rsTmp!����id)
                    lng���ID = Val("" & rsTmp!���id)
                End If
                Set rsPath = GetPathTable(lng����ID, lng���ID, mPatiInfo.����ID)
            End If

            If rsPath.RecordCount = 0 Then
                Set rsNext = Get����ID(mlng����ID, mlng��ҳID, mPatiInfo.����ID, , 1)
                If rsNext.RecordCount = 0 Then Exit Sub
            End If
            Call mclsDis.HideFrm(0) '���ش�Ⱦ�����Խ������
            Call mclsPath.zlRefresh(mlng����ID, mlng��ҳID, mPatiInfo.����ID, mPatiInfo.����ID, mPatiInfo.״̬, False, True)
            Call mclsPath.zlImportPath
            Call LoadPatients '��������ˢ��
            If rptPati.Visible Then rptPati.SetFocus
        End If
    End If
End Sub

Private Sub mclsMipModule_OpenLink(ByVal strMsgKey As String, ByVal strLinkPara As String)
'���ܣ����ð����Ϣ��λ����
    If InStr(",ZLHIS_PATIENT_002,ZLHIS_PATIENT_012,ZLHIS_PATIENT_009,ZLHIS_PATIENT_006,ZLHIS_PATIENT_010,", "," & strMsgKey & ",") > 0 Then
        If tbcPati.Item(1).Selected = False Then
            tbcPati.Item(1).Selected = True 'ѡ��л�ʱ��ˢ�²����б�
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
    Dim blnRecToLis As Boolean '�Ƿ���ص������б���
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
        Case "סԺ��"
            If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        Case "����"
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case "���￨"
            If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
                KeyAscii = 0
            Else
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End If
        Case "����"
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
'���ܣ��Ķ���Ϣ��ɾ����Ϣ����˫����Ϣ����ѡ����Ϣ���ٰ��س�����
'������Ķ�Σ��ֵ��Ϣ�򴥷���Ϣ
    Dim objControl As CommandBarControl
    Dim lngIndex As Long, lng����ID As Long, lng��ҳID As Long
    Dim lngҽ��ID As Long, str���� As String, strסԺ�� As String
    Dim strҵ�� As String, lng��ϢID As Long
    Dim blnFinded As Boolean, blnOk As Boolean
    Dim strNO As String, str���� As String
    Dim i As Long
    Dim str��Դ As String
    Dim strTmp As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    If KeyCode = vbKeyReturn Then
        If rptNotify.SelectedRows.Count > 0 Then
            With rptNotify.SelectedRows(0).Record
                strNO = .Item(C_��Ϣ).Value
                strҵ�� = .Item(C_ҵ��).Value
                lng����ID = Val(.Item(C_����Id).Value)
                lng��ҳID = Val(.Item(C_��ҳId).Value)
                lng��ϢID = Val(.Item(C_Id).Value)
                str���� = .Item(c_����).Value
                strסԺ�� = .Item(c_סԺ��).Value
                str���� = .Item(c_����).Value
                lngIndex = .Index
            End With
            
            If rptPati.SelectedRows.Count > 0 Then
                If Not rptPati.SelectedRows(0).GroupRow Then
                    '�ǲ����Ѿ���λ
                    blnFinded = InStr("_" & rptPati.SelectedRows(0).Record.Tag & "_", "_" & rptNotify.SelectedRows(0).Record.Tag & "_") > 0
                End If
            End If
            '���û�ҵ����ˣ��ҽ���ûѡ����Ժ�����б�ʱ�л��б��ٲ���һ��
            If tbcPati.Item(tbcPatiEnu.E��Ժ).Selected = False And Not blnFinded Then
                tbcPati.Item(tbcPatiEnu.E��Ժ).Selected = True
                blnFinded = LocatePati(rptNotify.SelectedRows(0).Record.Tag)
            End If
            
            If blnFinded And tbcSub.Tag = "ҽ��" And strҵ�� <> "" Then  '�ҵ����˺��پ����Ƿ�λҽ��
                lngҽ��ID = Val(strҵ��)
                If lngҽ��ID <> 0 Then
                    Call mclsAdvices.LocatedAdviceRow(lngҽ��ID)
                End If
            End If
          
            strTmp = ""
            If strNO = "ZLHIS_LIS_003" Then '����
                strTmp = "ZLHIS_CIS_014"
            ElseIf strNO = "ZLHIS_PACS_005" Then '���
                strTmp = "ZLHIS_CIS_025"
            End If
            If strTmp <> "" Then
                If Not (mclsMipModule Is Nothing) Then
                    If mclsMipModule.IsConnect Then
                        strSQL = "select ��Ժ����ID,��ǰ����ID from ������ҳ where ����ID=[1] and ��ҳID=[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
                        Call ZLHIS_CIS_MsgReadAfter(mclsMipModule, strTmp, lng����ID, str����, strסԺ��, , 2, _
                                lng��ҳID, Val(rsTmp!��ǰ����ID & ""), Val(rsTmp!��Ժ����ID & ""), str����, lngҽ��ID)
                    End If
                End If
            End If
            
            If strNO = "ZLHIS_RECIPEAUDIT_002" Then
                If strҵ�� = "������ҩ��" Then
                    '�������ⲿ������Ϣ   '���ýӿ��жϵ�ǰ���˵�ҽ���ǲ��Ƕ���ͨ����
                    blnOk = CheckZLPass(Me, lng����ID, lng��ҳID)
                    If blnOk Then
                        'δͨ��ҽ���༭���嵯���� �޸�
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
                            If objControl.Enabled Then objControl.Execute
                        End If
                    End If
                Else
                    'סԺ�Ĵ������ֻ�в��ϸ��Ȼ���������
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
                
                    Set objControl = cbsMain.FindControl(, conMenu_Edit_Blankoff, True, True)
                    If Not objControl Is Nothing Then
                        If objControl.Enabled Then objControl.Execute
                    End If
                End If
            End If
            
            If blnFinded And strNO = "ZLHIS_CIS_032" Then
                Call mclsDis.ShowDisRegist(Me, 1, Val(strҵ��), lng����ID, lng��ҳID)
            End If
            
            '�����ʿ���Ϣ
            If strNO = "ZLHIS_EMR_025" Then
                If tbcSub.Tag <> "�²���" Then
                    For i = 0 To tbcSub.ItemCount - 1
                        If tbcSub.Item(i).Visible Then
                            If tbcSub.Item(i).Tag = "�²���" Then
                                tbcSub.Item(i).Selected = True
                                cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
                                Exit For
                            End If
                        End If
                    Next
                End If
                If blnFinded Then
                    Set objControl = cbsMain.FindControl(, 3309, True, True)
                    If Not objControl Is Nothing Then
                        If objControl.Enabled Then
                            objControl.Parameter = strҵ��
                            objControl.Execute
                            Call ReadMsg(lng����ID, lng��ҳID, strNO, strҵ��, lng��ϢID)
                            Call ReadMsg����(lng����ID, lng��ҳID, strNO)
                        End If
                    End If
                End If
                Exit Sub
            End If
            
            If blnFinded And strNO = "ZLHIS_CIS_033" Then
            '��Ⱦ�����淴�޸���Ϣ�Ķ�
                blnOk = ReadMsgCIS033(lng����ID, lng��ҳID, strҵ��, lng��ϢID)
                If blnOk Then Call rptNotify.Records.RemoveAt(lngIndex)
            End If
            
            If blnFinded And strNO = "ZLHIS_BLOOD_006" Then
                If gobjPublicBlood Is Nothing And gblnѪ��ϵͳ Then InitObjBlood
                blnOk = gobjPublicBlood.zlIsBloodMessageDone(2, lng����ID, lng��ҳID, 2, cboDept.ItemData(cboDept.ListIndex))
                If blnOk Then
                    Call rptNotify.Records.RemoveAt(lngIndex)
                Else
                    If FuncTraReaction(Val(strҵ��), mlngModul, False, IIf(InStr(1, strҵ��, ":") > 0, Val(Split(strҵ��, ":")(1)), 0)) Then
                        If gobjPublicBlood.zlIsBloodMessageDone(2, lng����ID, lng��ҳID, 2, cboDept.ItemData(cboDept.ListIndex)) Then
                            Call rptNotify.Records.RemoveAt(lngIndex)
            End If
                    End If
                End If
            End If
            If strNO <> "ZLHIS_CIS_033" And strNO <> "ZLHIS_BLOOD_006" Then
                blnOk = ReadMsg(lng����ID, lng��ҳID, strNO, strҵ��, lng��ϢID)
                If blnOk Then Call rptNotify.Records.RemoveAt(lngIndex)
            End If
            Call rptNotify.Populate
        End If
    End If
End Sub

Private Function ReadMsgCIS033(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str��ʶ As String, ByVal lng��ϢID As Long) As Boolean
'���ܣ���Ⱦ�����淴�޸���Ϣ�Ķ�
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim lng�ļ�ID As Long
    Dim objControl As CommandBarControl
    Dim i As Long
    
    On Error GoTo errH
    'conMenu_Edit_Modify 3003 �޸İ�ť��
    lng�ļ�ID = Val(Split(str��ʶ, ",")(0))
    
    strSQL = "Select 1 From �����걨��¼ where �ļ�ID=[1] and ����״̬=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ļ�ID, 4)
    If rsTmp.RecordCount = 0 Then
    '����Ϣ���Ϊ�Ѷ�
        strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng��ҳID & ",'ZLHIS_CIS_033',2,'" & UserInfo.���� & "'," & cboDept.ItemData(cboDept.ListIndex) & ",null," & lng��ϢID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        ReadMsgCIS033 = True
        Exit Function
    End If
    
    If "�л����񹲺͹���Ⱦ�����濨" = Sys.RowValue("���Ӳ�����¼", lng�ļ�ID, "��������") Then
        '�������޸ı���
        '�Ƚ���Ƭ�л���ҽ����Ƭ������Ҳ˵�
        If tbcSub.Tag <> "��������" Then
            For i = 0 To tbcSub.ItemCount - 1
                If tbcSub.Item(i).Visible Then
                    If tbcSub.Item(i).Tag = "��������" Then
                        tbcSub.Item(i).Selected = True
                        cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
                        Exit For
                    End If
                End If
            Next
        End If
        
        cbsMain.RecalcLayout: Me.Refresh '����δ���ü�ˢ��
        If tbcSub.Selected.Tag = "��������" And tbcSub.Selected.Visible = True Then
            Set objControl = cbsMain.FindControl(, conMenu_Edit_Modify, True, True)
            If Not objControl Is Nothing Then
                If objControl.Enabled Then objControl.Execute
            End If
        End If
    Else
        '�������޸ı���
        Call mclsDis.ModifyDiseaseDoc(Me, lng�ļ�ID, mlng����ID, mlng��ҳID, 2, mPatiInfo.����ID)
    End If
    
    strSQL = "Select 1 From �����걨��¼ where �ļ�ID=[1] and ����״̬=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ļ�ID, 4)
    If rsTmp.RecordCount = 0 Then
    '����Ϣ���Ϊ�Ѷ�
        strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng��ҳID & ",'ZLHIS_CIS_033',2,'" & UserInfo.���� & "'," & cboDept.ItemData(cboDept.ListIndex) & ",null," & lng��ϢID & ")"
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

Private Sub mclspath_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
'���ܣ��ٴ�·���в鿴���Ӳ�������
    Call gobjRichEPR.ViewDocument(Me, ����ID, CanPrint)
End Sub

Private Sub mclspath_RequestRefresh(ByVal lngPathState As Long)
'���ܣ��ٴ�·����ˢ�²�����Ϣ�б��е�״̬,-1��ʾδ����״̬
    With rptPati.SelectedRows(0)
        .Record(col_·��״̬).Value = lngPathState
        .Record(col_·��״̬).Caption = " "
        .Record(col_·��״̬).Icon = -1 + Choose(lngPathState + 2, imgPati.ListImages("δ����").Index, imgPati.ListImages("������").Index, _
                imgPati.ListImages("ִ����").Index, imgPati.ListImages("��������").Index, imgPati.ListImages("�������").Index)
        
    End With
    If rptPati.Columns(col_·��״̬).Visible = False Then
        rptPati.Columns(col_·��״̬).Visible = True
    End If
    rptPati.Populate
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

Private Sub cboPages_Click()
'���ܣ�ѡ��ĳ��סԺ��¼ʱ����ȡ��صĲ�����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lng��ʶ As Long
    
    If cboPages.ListIndex = -1 Then Exit Sub
    If cboPages.ListIndex = mintPrePage Then Exit Sub
    mintPrePage = cboPages.ListIndex
    
    On Error GoTo errH
    
    mrsPati.Filter = "���=" & cboPages.ItemData(cboPages.ListIndex)
    
    mPatiInfo.����ID = mrsPati!����ID
    mPatiInfo.��ҳID = mrsPati!��ҳID
    
    strSQL = "Select NVL(B.����,A.����) ���� ,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����, A.����) ����, b.סԺ��, b.��Ժ����, b.ҽ�Ƹ��ʽ, d.��Ϣֵ As ҽ����, b.����, b.��ǰ����, c.���� As ����ȼ�, Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��) As ��Ժ����, b.��Ժ����, b.��Ŀ����," & vbNewLine & _
            "       b.��������, b.״̬, b.����ת��, b.��Ժ����id, b.��ǰ����id, a.סԺ����, e.�����" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B, �շ���ĿĿ¼ C, ������ҳ�ӱ� D, ��λ״����¼ E" & vbNewLine & _
            "Where a.����id = b.����id And a.����id = [1] And b.��ҳid = [2] And b.����ȼ�id = c.Id(+) And b.����id = d.����id(+) And" & vbNewLine & _
            "      b.��ҳid = d.��ҳid(+) And d.��Ϣ��(+) = 'ҽ����' And b.��Ժ����id = e.����id(+) And b.����id = e.����id(+) And b.��Ժ���� = e.����(+)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mPatiInfo.����ID, mPatiInfo.��ҳID)
    
    With rsTmp
        '���ղ���������ɫ��ʾ
        lbl����(1).Caption = "" & !סԺ��
        lbl����(1).ForeColor = zlDatabase.GetPatiColor(NVL(!��������))
        lblPatiName(1).Caption = "" & !����
        lblPatiName(1).ToolTipText = lblPatiName(1).Caption
        lblPatiName(1).ForeColor = lbl����(1).ForeColor
        
        lblҽ����(1).Caption = NVL(!ҽ����)
        lbl����(1).Caption = NVL(!����ȼ�)
        lbl����(1).Caption = NVL(!ҽ�Ƹ��ʽ)
        
        'Σ�ز��˲�����ɫ��ʾ
        lbl����(1).Caption = NVL(!��ǰ����)
        If NVL(!��ǰ����) = "Σ" Or NVL(!��ǰ����) = "��" Or NVL(!��ǰ����) = "��" Then
            lbl����(1).ForeColor = &HC0&
        Else
            lbl����(1).ForeColor = lblҽ����(1).ForeColor
        End If
        
        lbl��Ժ(1).Caption = Format(!��Ժ����, "yyyy-MM-dd HH:mm")
        If Not IsNull(!��Ժ����) Then
            lbl��Ժ(1).Caption = lbl��Ժ(1).Caption & "��" & Format(!��Ժ����, "yyyy-MM-dd HH:mm")
        
        End If
        
        lbl����(1).Caption = NVL(!��������)
        lbl����(1).Caption = IIf(IsNull(!�����), "", "(" & !����� & ")") & !��Ժ����
        
        '���
        lblDiag(1).Caption = GetPatiDiagnose(mPatiInfo.����ID, mPatiInfo.��ҳID, 2)
        
        '������Ϣ
        mPatiInfo.״̬ = NVL(!״̬, 0)
        mPatiInfo.סԺ�� = NVL(!סԺ��)
        mPatiInfo.���� = NVL(!��Ժ����)
                
        mPatiInfo.����ID = NVL(!��ǰ����ID, 0)
        mPatiInfo.����ID = NVL(!��Ժ����ID, 0)
                
        mPatiInfo.��Ժ���� = !��Ժ����
        If Not IsNull(!��Ժ����) Then
            mPatiInfo.��Ժ���� = !��Ժ����
        Else
            mPatiInfo.��Ժ���� = CDate(0)
        End If
        If Not IsNull(!��Ŀ����) Then
            mPatiInfo.��Ŀ���� = !��Ŀ����
        Else
            mPatiInfo.��Ŀ���� = CDate(0)
        End If
        mPatiInfo.סԺ���� = NVL(!סԺ����, 0)
        mPatiInfo.����ת�� = NVL(!����ת��, 0) <> 0
            
        Call SetPatiInfoCtlPos
    End With
    
    '��ʾ����ͼ��
    Call ShowPatiͼ��(mPatiInfo.����ID, mPatiInfo.��ҳID)
     
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
        
    '��ǰ����Ϊ��ǰҪ��λ�Ĳ���
    blnSeek = False
    If rptPati.SelectedRows.Count > 0 Then
        If Not rptPati.SelectedRows(0).GroupRow Then
            If rptPati.SelectedRows(0).Record(col_����Id).Value = PatiID _
                And rptPati.SelectedRows(0).Record(col_��ҳID).Value = PageID Then blnSeek = True
        End If
    End If
    '�Զ�Ѱ�Ҳ��л���ʾ��ǰҪ��λ�Ĳ���
    If Not blnSeek Then
        For Each objRow In rptPati.Rows
            If Not objRow.GroupRow Then
                If objRow.Record(col_����Id).Value = PatiID And objRow.Record(col_��ҳID).Value = PageID Then
                    blnEnabled = timNotify.Enabled
                    timNotify.Enabled = False '������������ˢ����������
                    Set rptPati.FocusedRow = objRow 'ѡ��,��ʾ,[����Change�¼�]
                    timNotify.Enabled = blnEnabled
                    blnSeek = True: Exit For
                End If
            End If
        Next
    End If
    If Not blnSeek Then
        MsgBox "��ǰ�����嵥��û���ҵ��ò��ˡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��λ����Ӧ������ҳ��
    strTab = Decode(ObjectType, 1, "ҽ��", 2, "����", 3, "����", 4, "����", 5, "", 6, "ҽ��", 7, "����", 8, "����", 9, "·��")
    If strTab <> "" And tbcSub.Selected.Tag <> strTab Then
        For i = 0 To tbcSub.ItemCount - 1
            If tbcSub(i).Tag = strTab Then
                tbcSub(i).Selected = True
                Me.Refresh: Exit For
            End If
        Next
        If tbcSub.Selected.Tag <> strTab Then
            MsgBox "���ܶ�λ��" & strTab & "���ݣ���������û����Ӧ��Ȩ�ޡ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    If ObjectType = 3 Or ObjectType = 4 Then '��λ������Ļ������ݽ���
        If Me.tbcSub.Item(mlngOldIndex).Visible Then
            Call mclsTends.zlLocateData(IIf(ObjectType = 3, 1, 0))
        ElseIf Me.tbcSub.Item(mlngNewIndex).Visible Then
            Call mclsTendsNew.zlLocateData(IIf(ObjectType = 3, 1, 0))
        End If
    End If
    On Error GoTo errH
    '�򿪶�Ӧ�Ķ���
    Select Case ObjectType
    Case 1 'סԺҽ��
    Case 2, 3, 7, 8 'סԺ����,������,����֤��,֪���ļ�
        If ObjectID = "0" Or ObjectID = "" Then Exit Sub
        If IsNumeric(ObjectID) Then
            Call gobjRichEPR.EditDocument(pסԺҽ��վ, Me, cboDept.ItemData(cboDept.ListIndex), ObjectID)
        Else '�°没��
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
                                MsgBox "ԭʼ�����Ѳ����ڣ��޷��鿴��", vbInformation, gstrSysName
                                Exit Sub
                        End If
            
            strPrivs = ";" & zl9ComLib.GetPrivFunc(glngSys, p���Ӳ�������) & ";"
            If NVL(rsEmr!completor) = "" Then
                If InStr(strPrivs, ";�ĵ���д;") > 0 Then '����дȨ��
                    Call gobjEmr.OpenFormForModifyDoc(Me.hwnd, rsEmr!masterid, rsEmr!actlogid, NVL(rsEmr!basiclogid), rsEmr!actionid, rsEmr!taskid, rsEmr!antetypeid, rsEmr!doctype, rsEmr!docid, Val(rsEmr!Occasion), Val(rsEmr!besealed), Val(rsEmr!docsecret), NVL(rsEmr!subdocid), "02", strPrivs)
                Else '��Ȩ��ֻ�ܲ鿴
                    Set objEmr = DynamicCreate("zlRichEMR.clsDockContent", "��ʾ����", True)
                    If Not objEmr Is Nothing Then
                        Call objEmr.Init(gobjEmr, gcnOracle, glngSys, 0)
                        Call objEmr.zlShowDoc(strDocID, strSubdocID)
                        Call objEmr.zlViewDoc(Me, "���Ĳ���", strSubdocID)
                    End If
                End If
            Else
                If InStr(strPrivs, ";�ĵ���;") > 0 Then '����дȨ��
                    Call gobjEmr.OpenFormForAuditDoc(Me.hwnd, rsEmr!masterid, rsEmr!actlogid, NVL(rsEmr!basiclogid), rsEmr!actionid, rsEmr!taskid, rsEmr!antetypeid, rsEmr!doctype, rsEmr!docid, Val(rsEmr!Occasion), Val(rsEmr!besealed), Val(rsEmr!docsecret), NVL(rsEmr!subdocid), "02", strPrivs)
                Else '��Ȩ��ֻ�ܲ鿴
                    Set objEmr = DynamicCreate("zlRichEMR.clsDockContent", "��ʾ����", True)
                    If Not objEmr Is Nothing Then
                        Call objEmr.Init(gobjEmr, gcnOracle, glngSys, 0)
                        Call objEmr.zlShowDoc(strDocID, strSubdocID)
                        Call objEmr.zlViewDoc(Me, "���Ĳ���", strSubdocID)
                    End If
                End If
            End If
        End If
    Case 4 '�����¼
    Case 5 '��ҳ��¼
        If InStr(mstrPrivs, "��ҳ����") > 0 Then
            Call ExecuteEditMediRec(True)
        Else
            MsgBox "��û�С�סԺҽ������վ������ҳ����Ȩ�ޣ����ܲ鿴�༭��ҳ!", vbInformation, gstrSysName
        End If
    Case 6 'ҽ������
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
    'Panne�е�Report�ؼ���Ҫǿ�д�����˳��
    '������ʱ���ܲ���vbKeyTab
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
    Dim lngҽ��ID As Long
    
    If rptNotify.SelectedRows.Count = 0 Then Exit Sub  '���������
    
    With rptNotify.SelectedRows(0)
        
        lngIndex = rptNotify.FocusedRow.Record.Index
        If rptNotify.Rows(lngIndex).Record(C_��Ϣ).Value = "ZLHIS_LIS_003" Or rptNotify.Rows(lngIndex).Record(C_��Ϣ).Value = "ZLHIS_PACS_005" Then '��������Ϣ��ҵ���д���� ҽ��id��������Դ
            lngҽ��ID = Val(Split(rptNotify.Rows(lngIndex).Record(C_ҵ��).Value, ",")(0))
        ElseIf rptNotify.Rows(lngIndex).Record(C_��Ϣ).Value = "ZLHIS_BLOOD_006" Then
            lngҽ��ID = Val(Split(rptNotify.Rows(lngIndex).Record(C_ҵ��).Value, ":")(0))
        Else
            lngҽ��ID = Val(rptNotify.Rows(lngIndex).Record(C_ҵ��).Value)
        End If
        
        '���ظ������ͬһ����Ŀ,�ҵ�ǰ����Ϊ��ǰ�������Ŀ,�򲻹�
        If .Record.Tag = mstrPreNotify Then
            If rptPati.SelectedRows.Count > 0 Then
                If Not rptPati.SelectedRows(0).GroupRow Then
                    If Val(rptPati.SelectedRows(0).Record(col_����Id).Value) <> 0 Then strCurPati = rptPati.SelectedRows(0).Record.Tag
                End If
            End If
        End If
        
        If .Record.Tag <> strCurPati Then
            mstrPreNotify = .Record.Tag
            If rptNotify.Rows(lngIndex).Record(C_��Ϣ).Value = "ZLHIS_CIS_020" Then
                If tbcPati.Selected.Index <> 4 Then
                    tbcPati.Item(4).Selected = True
                End If
            End If
            '�Զ�Ѱ�Ҳ��л���ʾ��ǰ���ѵĲ���
            If Not LocatePati(.Record.Tag) Then
                Call LoadPatients
                Call LocatePati(.Record.Tag)
            End If
        End If
        
        If lngҽ��ID <> 0 And tbcSub.Tag = "ҽ��" Then
            Call mclsAdvices.LocatedAdviceRow(lngҽ��ID)
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
        
    '����ʱ��ǿ���Ȱ����״̬����
    '������������Ч����������һ������
    If rptPati.SortOrder.Count = 1 Then
        If rptPati.SortOrder(0).Index <> col_��� Then
            Set objCol = rptPati.SortOrder(0)
            rptPati.SortOrder.DeleteAll
            rptPati.SortOrder.Add rptPati.Columns(col_���)
            rptPati.SortOrder.Add objCol
        End If
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "������ɫ" Then
        Call zlDatabase.ShowPatiColorTip(Me)
    End If
End Sub

Private Sub tbcPati_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim i As Long
    
    For i = 0 To picPara.Count - 1
        picPara(i).Visible = False
    Next
    If Item.Tag = "��Ժ" Then
        picPara(0).Visible = True
    ElseIf Item.Tag = "��Ժ" Then
        picPara(1).Visible = True
    ElseIf Item.Tag = "ת��" Then
        picPara(3).Visible = True
    ElseIf Item.Tag = "����" Then
        picPara(2).Visible = True
    End If
    
    chkFilter.Visible = IIf(mstrFindType = "סԺ��" And tbcPati.Selected.Tag = "��Ժ", True, False)
    Call picPati_Resize
    
    Call picPatiIn_Resize
    If Me.Visible Then
        Call LoadPatients
        Call LoadNotify 'ˢ��ҽ������
    End If

End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'���ܣ�ˢ���Ӵ�����漰����
'˵����������Ϊ�л����濨Ƭ����
    Dim Index As Long, objItem As TabControlItem
    
    If mblnTabTmp Then Exit Sub
    If Item.Tag = "" Then Exit Sub '��ʼ��ʱ,��û��ֵ
    If Item.Handle = picTmp.hwnd Then
        Index = Item.Index
        mblnTabTmp = True
        Screen.MousePointer = 11
        On Error GoTo errH
        Select Case Item.Tag
            Case "סԺһ��"
                Set objItem = tbcSub.InsertItem(Index, "סԺһ��", mcolSubForm("_סԺһ��").hwnd, 0)
                objItem.Tag = "סԺһ��"
            Case "·��"
                Set objItem = tbcSub.InsertItem(Index, "�ٴ�·��", mcolSubForm("_·��").hwnd, 0)
                objItem.Tag = "·��"
            Case "ҽ��"
                Set objItem = tbcSub.InsertItem(Index, "ҽ����Ϣ", mcolSubForm("_ҽ��").hwnd, 0)
                objItem.Tag = "ҽ��"
            Case "����"
                Set objItem = tbcSub.InsertItem(Index, "������Ϣ", mcolSubForm("_����").hwnd, 0)
                objItem.Tag = "����"
            Case "�²���"
                Set objItem = tbcSub.InsertItem(Index, "���Ӳ���", mcolSubForm("_�²���").hwnd, 0)
                objItem.Tag = "�²���"
            Case "����"
                Set objItem = tbcSub.InsertItem(Index, "������Ϣ", mcolSubForm("_����").hwnd, 0)
                objItem.Tag = "����"
            Case "�°滤��"
                Set objItem = tbcSub.InsertItem(Index, "������Ϣ", mcolSubForm("_�°滤��").hwnd, 0)
                objItem.Tag = "�°滤��"
            Case "������"
                Set objItem = tbcSub.InsertItem(Index, "������", mcolSubForm("_������").hwnd, 0)
                objItem.Tag = "������"
            Case "�໤"
                Set objItem = tbcSub.InsertItem(Index, "����໤", mcolSubForm("_�໤").hwnd, 0)
                objItem.Tag = "�໤"
            Case "��������"
                Set objItem = tbcSub.InsertItem(Index, "��������", mcolSubForm("_��������").hwnd, 0)
                objItem.Tag = "��������"
        End Select
        Call tbcSub.RemoveItem(Index + 1)
        objItem.Selected = True
        Screen.MousePointer = 0
        mblnTabTmp = False
    End If
    
    If Item.Tag = "������������" Then
        Me.dkpMain.Options.UseSplitterTracker = True '������ҳ�ؼ�����ʵʱ�϶�
    Else
        Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    End If
     
    'ˢ���Ӵ����Ӧ��CommandBar
    Call SubWinDefCommandBar(Item)
    '91136:�������ŵ�SubWinDefCommandBar֮����ǰ��ǰ�浼�����º���ʱ���ϰ�˵��޷�ˢ�¡�
    If mblnIsNot Then mblnIsNot = False: Exit Sub
    'ˢ���Ӵ�������
    Call SubWinRefreshData(Item)
    
    If Visible Then mfrmActive.SetFocus
    tbcSub.Tag = Item.Tag   '��¼��һ��ѡ��Ŀ�Ƭ
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
'���ܣ�ˢ�½�������
'˵�����Ӹ��¼���ʼ�᲻�ظ�������ص����ݶ�ȡ
    Dim lng����ID As Long, i As Long, lngidx As Long
    Dim blnIn���� As Boolean, rsTmp As Recordset, str����IDs As String
    
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
    lng����ID = cboDept.ItemData(cboDept.ListIndex)
    
    mstrPreNotify = ""
    rptNotify.Records.DeleteAll
    rptNotify.Populate
    rptNotify.TabStop = False
    
    
    '�Ƿ�Ǳ��Ƶ�ICU��
    str����IDs = "," & GetUser����IDs(True) & ","
    mblnICU = Sys.DeptHaveProperty(lng����ID, "ICU")
    If mblnICU = True Then
        If mintDeptView = 0 Then
            mblnICU = InStr(str����IDs, "," & lng����ID & ",") = 0
        Else
            '������ʾ���жϲ�����Ӧ�����п��Ҷ��ǲ���Ա�Ŀ���ʱ���ų���
            blnIn���� = True
            Set rsTmp = Sys.RowValue("�������Ҷ�Ӧ", lng����ID, , "����ID")
            Do While Not rsTmp.EOF
                If InStr(str����IDs, "," & rsTmp!����ID & ",") = 0 Then
                    blnIn���� = False
                End If
                rsTmp.MoveNext
            Loop
            mblnICU = Not blnIn����
        End If
    End If
    
    Call Sys.DeptHaveProperty(lng����ID, IIf(mintDeptView = 0, "�ٴ�", "����"), mblnOutDept)
        
    '�ر�ҵ����
    Set mclsInOutMedRec = Nothing
    If Not mfrmResponse Is Nothing Then
        Unload mfrmResponse
    End If
    
    '���¶�ȡ����
    Call LoadPatients
    
    '��ʼ���������
    Call Initͼ����Ϣ(lng����ID)
    
    '��ʾ�ٴ�·����Ƭ
    lngidx = -1
    For i = 0 To tbcSub.ItemCount - 1
        If tbcSub(i).Tag = "·��" Then
            lngidx = i
            Exit For
        End If
    Next
    If lngidx >= 0 Then
        If HavePath(lng����ID) = False Then
            tbcSub(lngidx).Visible = False
            rptPati.Columns(col_·��״̬).Visible = False
            rptPati.Columns(col_·��״̬).Width = 0
            rptPati.Populate
            If tbcSub.Tag = "·��" Or tbcSub.Tag = "" Then tbcSub.Item(lngidx + 1).Selected = True
        Else
            If tbcSub(lngidx).Visible = False Then
                tbcSub(lngidx).Visible = True
                rptPati.Columns(col_·��״̬).Visible = True
                rptPati.Columns(col_·��״̬).Width = 18
                rptPati.Populate
                If tbcSub.Tag = "·��" Or tbcSub.Tag = "" Then tbcSub.Item(lngidx).Selected = True
            End If
        End If
    End If
    If Me.Visible Then Call LoadNotify
    If Visible And rptPati.Visible Then rptPati.SetFocus
 
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngidx As Long, i As Long

    With rptPati
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)������,���Կ���Record(�к�)����������
        Set objCol = .Columns.Add(col_����, "����", 0, False)
            objCol.Sortable = False: objCol.Visible = False
        Set objCol = .Columns.Add(col_���, "", 16, False)
            objCol.TreeColumn = True: objCol.Visible = False
            objCol.Sortable = False: objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_ͼ��, "", 18, False)
            objCol.Sortable = False: objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentCenter
        
        lngidx = -1
        For i = 0 To tbcSub.ItemCount - 1
            If tbcSub(i).Tag = "·��" Then
                lngidx = i
                Exit For
            End If
        Next
        If lngidx >= 0 Then
            Set objCol = .Columns.Add(col_·��״̬, "·��״̬", 18, True)
        Else
            Set objCol = .Columns.Add(col_·��״̬, "·��״̬", 0, False): objCol.Visible = False
        End If
        
        Set objCol = .Columns.Add(col_����Id, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_��ҳID, "��ҳID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_����, "����", 55, True)
        Set objCol = .Columns.Add(col_סԺ��, "סԺ��", 62, True)
        Set objCol = .Columns.Add(col_����, "����", 50, True)
        Set objCol = .Columns.Add(col_�໤, "�໤", 30, True)
        If mclsWardMonitor.Enabled = False Or InStr(GetInsidePrivs(pסԺҽ��վ), "����໤") = 0 Then
            objCol.Visible = False
        End If
        Set objCol = .Columns.Add(col_�Ա�, "�Ա�", 30, True)
        Set objCol = .Columns.Add(col_����, "����", 30, True)
        Set objCol = .Columns.Add(col_�ѱ�, "�ѱ�", 55, True)
        Set objCol = .Columns.Add(col_����, "����", 70, True)
        Set objCol = .Columns.Add(col_����, "����", 70, True): objCol.Visible = False
        Set objCol = .Columns.Add(col_סԺҽʦ, "סԺҽʦ", 55, True)
        Set objCol = .Columns.Add(col_��Ժ����, "��Ժ����", 106, True)
        Set objCol = .Columns.Add(col_��Ժ����, "��Ժ����", 106, True)
        Set objCol = .Columns.Add(col_��������, "��������", 106, True)
        
        Set objCol = .Columns.Add(col_ҽ��ID, "ҽ��ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_���ͺ�, "���ͺ�", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_ִ��״̬, "ִ��״̬", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_ִ�п���ID, "ִ�п���ID", 0, False): objCol.Visible = False
        If ISPassShowCard Then
            Set objCol = .Columns.Add(col_���￨, "���￨", 0, False)
        Else
            Set objCol = .Columns.Add(col_���￨, "���￨", 70, True)
        End If
        Set objCol = .Columns.Add(col_סԺ����, "סԺ����", 56, True)
        Set objCol = .Columns.Add(col_������, "������", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_Ӥ������ID, "Ӥ������ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_Ӥ������ID, "Ӥ������ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_��ҽ���, "��ҽ���", 106, True)
        Set objCol = .Columns.Add(col_��ҽ���, "��ҽ���", 106, True)
        Set objCol = .Columns.Add(COL_�������, "�������", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_��Ⱦ��, "��Ⱦ��", 106, True)
        Set objCol = .Columns.Add(col_���λ�ʿ, "���λ�ʿ", 55, True)
        Set objCol = .Columns.Add(col_���ۺ�, "���ۺ�", 62, True)
        Set objCol = .Columns.Add(col_���֤��, "���֤��", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_�Ƿ���, "��", 30, False): objCol.Visible = False
        
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = objCol.Index = col_����
        Next
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
        
        .GroupsOrder.Add .Columns(col_����)
        .GroupsOrder(0).SortAscending = True '����֮��,��������в���ʾ,�����е������ǲ����
        
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.Add .Columns(col_�Ƿ���)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(col_���)
        .SortOrder(1).SortAscending = True
        .SortOrder.Add .Columns(col_����)
        .SortOrder(2).SortAscending = True
    End With
    
    With rptNotify
        Set objCol = .Columns.Add(c_ͼ��, "", 18, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(C_����Id, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_��ҳId, "��ҳID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_����, "����", 60, True)
        Set objCol = .Columns.Add(c_סԺ��, "סԺ��", 62, True)
        Set objCol = .Columns.Add(c_����, "����", 40, True)
        Set objCol = .Columns.Add(C_״̬, "״̬", 150, True)
         
        Set objCol = .Columns.Add(C_��Ϣ, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_���, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_����, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_ҵ��, "", 0, False): objCol.Visible = False
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
    
    '------���ͼ��󵯳������б�
    With rptTBPati
        Set objCol = .Columns.Add(CI_ͼ��1, "", 18, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(CI_ͼ��2, "", 18, True): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(CI_ͼ��3, "", 18, True): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(CI_����, "����", 40, True)
        Set objCol = .Columns.Add(CI_����ID, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(CI_��ҳID, "��ҳID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(CI_����, "����", 60, True)
        Set objCol = .Columns.Add(CI_סԺ��, "סԺ��", 60, True)
        Set objCol = .Columns.Add(CI_��Ժ����, "��Ժ����", 70, True)
        Set objCol = .Columns.Add(CI_��Ժ����, "��Ժ����", 70, True)
        Set objCol = .Columns.Add(CI_��������, "��������", 100, True)
        
                
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�в���..."
        End With
        .PreviewMode = False
        .AllowColumnSort = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList zlCommFun.GetPaitSignImageList(0)
        
        .SortOrder.Add .Columns.Find(CI_����)
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
    
    blnSetup = InStr(";" & mstrPrivs & ";", ";��������;") > 0
    Call zlDatabase.SetPara("���˲��ҷ�ʽ", mintFindType, glngSys, pסԺҽ��վ, blnSetup)
    strTmp = ""
    For i = 0 To chk��������.UBound
        strTmp = strTmp & IIf(chk��������(i).Value = 1, "1", "0")
    Next
    Call zlDatabase.SetPara("��ǰ��������", strTmp, glngSys, pסԺҽ��վ, blnSetup)
    Call zlDatabase.SetPara("��С����ʾ", Val(chkByTeam.Value), glngSys, pסԺҽ��վ, blnSetup)
    
    strTmp = chkHZ(0).Value & chkHZ(1).Value
    If strTmp = "00" Then
        strTmp = "11"
    End If
    strTmp = chkOut.Value & strTmp
    Call zlDatabase.SetPara("���ﲡ�˹���", strTmp, glngSys, pסԺҽ��վ, blnSetup)
    
    '���˷�Χ
    curDate = zlDatabase.Currentdate
    Call zlDatabase.SetPara("���ת������", Val(txtChange.Text), glngSys, mlngModul, blnSetup)
    
     If Not mclsInOutMedRec Is Nothing Then
        If Not mclsInOutMedRec.FormUnLoad Then
            Cancel = True
            Exit Sub
        Else
            Cancel = False
            Set mclsInOutMedRec = Nothing
        End If
    End If
    
    Call SetAllPatiͼ��(1)
        
    If Me.Visible Then
        If Not tbcSub.Selected Is Nothing Then
            Call zlDatabase.SetPara("ҽ������", tbcSub.Selected.Tag, glngSys, pסԺҽ��վ, blnSetup)
        End If
        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 And Not mbln�������� Then
            Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, dkpMain.SaveStateToString)
        End If
        If Not tbcPati.Selected Is Nothing Then
            Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(tbcPati), "tbcPati", tbcPati.Selected.Index)
        End If
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(chkOutByTeam), "chkOutByTeam", chkOutByTeam.Value)
        '���������̶�����һ���ؼ�����ʽ���棬����վ���������һ���Ǵ�ӡ����̶���ͼ����ʽ,������ָ�Ϊ������ť����ʽ
        cbsMain(2).Controls(1).Style = cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style
        Call SaveWinState(Me, App.ProductName)
    End If
    
    'ǿ��Unload,��Ȼ���ἤ���Ӵ�����¼�
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
    mstrList���� = ""
    Set mrsPatiNotes = Nothing
    Set mrsPati���� = Nothing
    mblnΣ��ֵ = False
    Set mclsDisease = Nothing
    Set mobjReport = Nothing
    If Not mFrmConsultation Is Nothing Then
        Unload mFrmConsultation
    End If
    Set mFrmConsultation = Nothing
    mlng����ҽ��ID = 0
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
    tbcPati.Height = picPati.ScaleHeight - tbcPati.Top - IIf(fra���.Visible, fra���.Height, 0)
    
    fra���.Left = 0
    fra���.Top = tbcPati.Top + tbcPati.Height
    fra���.Width = picPati.ScaleWidth
    
    picPatiIn.Width = picPati.ScaleWidth
End Sub

Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    'Panne�е�Report�ؼ���Ҫǿ�д�����˳��
    '������ʱ���ܲ���vbKeyTab
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
            ElseIf objHitTest.Row.Childs.Count = 0 Or Val(objHitTest.Row.Record(col_����Id).Value) <> 0 Then
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
    Dim str���֤�� As String
    Dim strTmp As String
    Dim str����IDs As String
    
    mblnIsNot = False
    
    If rptPati.SelectedRows.Count = 0 Then Exit Sub          '���������
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
            mlng����ID = Val(.Record(col_����Id).Value)
            mlng��ҳID = Val(.Record(col_��ҳID).Value)
            str���֤�� = .Record(col_���֤��).Value
            
            If InStr(strCurPati, "_") > 0 Then
                mPatiInfo.Ӥ�� = Val(Split(strCurPati, "_")(1))
            Else
                mPatiInfo.Ӥ�� = -1
            End If

            mbln���ܻ��� = False
            mlng����ҽ��ID = 0
            If tbcPati.Selected.Tag = "����" Then
                If Val(.Record(col_ҽ��ID).Value) <> 0 Then
                    mbln���ܻ��� = Get���ܻ���
                End If
                mlng����ҽ��ID = Val(.Record(col_ҽ��ID).Value)
            End If
            
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
                strSQL = "Select rownum as ���,����id,��ҳID,NVL(��������,0) ��������,סԺ��,To_Char(��Ժ����,'YYYY-MM-DD HH24:MI') as ��Ժ���� From ������ҳ Where ��ҳID<>0 And ����ID=[1] Order by ��ҳID Desc,��� Desc"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
                Call Cbo.SetListWidth(cboPages.hwnd, cboPages.Width * 1.5)
            Else
                strSQL = "Select rownum as ���,a.����id,a.��ҳID,NVL(a.��������,0) ��������,a.סԺ��,To_Char(a.��Ժ����,'YYYY-MM-DD HH24:MI') as ��Ժ���� From ������ҳ a" & _
                " where a.��ҳID<>0 And A.����ID In (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X)" & _
                " order by a.����id desc,a.��ҳid desc,��� Desc"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����IDs)
                Call Cbo.SetListWidth(cboPages.hwnd, cboPages.Width * 4)
            End If
            Set mrsPati = zlDatabase.CopyNewRec(rsTmp)
            cboPages.Clear
            Do While Not rsTmp.EOF
            
                If str����IDs = "" Then
                    strTmp = "�� " & rsTmp!��ҳID & " ��" & Decode(rsTmp!��������, 1, "(��������)", 2, "(סԺ����)", "")
                Else
                    strTmp = "�� " & rsTmp!��ҳID & IIf(IsNull(rsTmp!סԺ��), "", "_" & rsTmp!סԺ��) & " ��" & Decode(rsTmp!��������, 1, "(��������)", 2, "(סԺ����)", "") & ":" & rsTmp!��Ժ����
                End If
                
                cboPages.AddItem strTmp
                cboPages.ItemData(cboPages.NewIndex) = rsTmp!���
                If rsTmp!��ҳID = mlng��ҳID And rsTmp!����ID = mlng����ID Then
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
            If GetInsidePrivs(p�����¼����, True) <> "" Then
                strSQL = "Select 1 From ���˻����¼ A Where a.����id = [1] And a.��ҳid = [2]"
                If mPatiInfo.����ת�� Then
                    strSQL = Replace(strSQL, "���˻����¼", "H���˻����¼")
                End If
                
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
                If rsTmp.RecordCount > 0 Then
                    Me.tbcSub.Item(mlngOldIndex).Visible = True
                    Me.tbcSub.Item(mlngNewIndex).Visible = False
                    Me.tbcSub.Item(mlngNewIndex + 1).Visible = False '�°��ͬʱ���ػ�����
                    If tbcSub.Item(mlngOldIndex).Selected Or tbcSub.Item(mlngNewIndex).Selected Or tbcSub.Item(mlngNewIndex + 1).Selected Then
                        mblnIsNot = Not tbcSub.Item(mlngOldIndex).Selected
                        Me.tbcSub.Item(mlngOldIndex).Selected = True
                    End If
                Else
                    Me.tbcSub.Item(mlngNewIndex).Visible = True
                    Me.tbcSub.Item(mlngOldIndex).Visible = False
                    Me.tbcSub.Item(mlngNewIndex + 1).Visible = True '�°��ͬʱ��ʾ������
                    If tbcSub.Item(mlngOldIndex).Selected Or tbcSub.Item(mlngNewIndex).Selected Then
                        mblnIsNot = Not tbcSub.Item(mlngNewIndex).Selected
                        Me.tbcSub.Item(mlngNewIndex).Selected = True
                    End If
                End If
            End If
            
            
            Call LoadPatiAllergy(mlng����ID, cbo����)
                        
            '��Ժ���˶�ȡ�Ƿ����ύ���
            If (Int(Val(.Record(col_����).Value)) = pt��Ժ Or Int(Val(.Record(col_����).Value)) = pt����) Or (tbcPati.Selected.Tag = "��Ժ" And Val(.Record(col_����Id).Value) <> 0) Then
                If .Record(col_ͼ��).Value = -1 Then
                    '1-�ȴ����;2-�ܾ����;3-�������;4-��鷴��;5-���鵵
                    If .Record(col_���).Value = 0 Or .Record(col_���).Value = 999 Then
                        .Record(col_ͼ��).Value = 0
                    ElseIf .Record(col_���).Value = 1 Or .Record(col_���).Value = 2 Then
                        .Record(col_ͼ��).Value = 1
                    Else
                        .Record(col_ͼ��).Value = IIf(PatiMedRecHaveSubmit(mlng����ID, mlng��ҳID), 1, 0)
                    End If
                End If
            End If
            
            LockWindowUpdate 0
            
            stbThis.Panels(2).Text = IIf(stbThis.Panels(2).Tag = "", "", stbThis.Panels(2).Tag)
            lblFee(1).Caption = GetPati������Ϣ(mlng����ID, mlng��ҳID) & IIf(InStr(mstrPrivs, "ҩռ�Ȳ�ѯ") = 0, "", GetסԺ����ҩռ��(mlng����ID, mlng��ҳID))
            '��Ժ���˲���ʾ��Һ��
            If mPatiInfo.��Ժ���� = CDate(0) Then
                lblFluid(0).Visible = True
                lblFluid(1).Visible = True
                strSQL = Get������Һ��(mlng����ID, mlng��ҳID)
                lblFluid(1).Caption = "����" & Split(strSQL, ",")(0) & "ml,����" & Split(strSQL, ",")(1) & "ml"
            Else
                lblFluid(0).Visible = False
                lblFluid(1).Visible = False
            End If
            intTmp = Get����ҽ����ӡ(mlng����ID, mlng��ҳID)
            lblPrint(1).Caption = IIf(intTmp = 0, "δ��ӡ", IIf(intTmp = 1, "���ִ�ӡ", "ȫ����ӡ"))
            On Error Resume Next
            If Visible And rptPati.Visible Then rptPati.SetFocus
            If err.Number <> 0 Then err.Clear
            On Error GoTo errH
        Else
            Call ClearPatiInfo
            '��������ˢ���Ӵ���
            Call SubWinRefreshData(tbcSub.Selected)
            
            stbThis.Panels(2).Text = stbThis.Panels(2).Tag
            lblFee(1).Caption = ""
            lblFluid(1).Caption = ""
            lblPrint(1).Caption = ""
        End If
    End With
    '�����ؼ��ò������㣬����ͨ�����������ؼ��õ�����ķ�ʽ������BUG��74488��
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

Private Function Get���ܻ���() As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    strSQL = "SELECT A.����ʱ��,A.������ FROM ����ҽ������ A where ҽ��ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rptPati.SelectedRows(0).Record(col_ҽ��ID).Value))
    If Not rsTmp.EOF Then
        Get���ܻ��� = IIf(rsTmp!������ & "" = "", False, True)
    Else
        Get���ܻ��� = False
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
'���ܣ���ע����ȡ��Ժ���˵�ʱ�䷶Χ
    Dim curDate As Date, intDay As Integer
    Dim intType As Integer
    
    '������ʾ��Χ
    mintChange = Val(zlDatabase.GetPara("���ת������", glngSys, pסԺҽ��վ, 7))
    '�������30���ȡȱʡֵ
    If mintChange > 30 Then mintChange = 7
    
    '��Ժ����ʱ�䷶Χ���̶�Ϊ��ȥ3��
    curDate = zlDatabase.Currentdate
    mdtOutEnd = Format(curDate, "yyyy-MM-dd 23:59:59")
    mdtOutBegin = Format(mdtOutEnd - 3, "yyyy-MM-dd 00:00:00")
    
    '���ﲡ��ʱ�䷶Χ���̶�Ϊ��ȥ3��
    mdtMeetEnd = Format(curDate, "yyyy-MM-dd 23:59:59")
    mdtMeetBegin = Format(mdtMeetEnd - 3, "yyyy-MM-dd 00:00:00")
    
    '�Զ�ˢ�²������ļ��
    mintNotify = Val(zlDatabase.GetPara("�Զ�ˢ�²������ļ��", glngSys, pסԺҽ��վ))
    mintNotifyDay = Val(zlDatabase.GetPara("�Զ�ˢ�²�����������", glngSys, pסԺҽ��վ, 1))
    mstrNotify = zlDatabase.GetPara("�Զ�ˢ������", glngSys, pסԺҽ��վ, "0000")
    mbln��Ϣ���� = Val(zlDatabase.GetPara("����������ʾ", glngSys, pסԺҽ��վ)) = 1
    mblnΣ��ֵ���� = Val(zlDatabase.GetPara("סԺΣ��ֵ��������", glngSys, pסԺҽ��վ, 1)) = 1
    
    '������ʾ��ʽ
    mintDeptView = Val(zlDatabase.GetPara("������ʾ��ʽ", glngSys, pסԺҽ��վ, , , , intType))
    mblnDeptViewEnabled = Not ((intType = 3 Or intType = 15) And InStr(mstrPrivs, "��������") = 0)
    
    '������鷴������
    mlngMedRedDay = Val(zlDatabase.GetPara("������鷴������", glngSys, pסԺҽ��վ))
    '�����С
    mbytSize = Val(zlDatabase.GetPara("����", glngSys, pסԺҽ��վ, "0"))
    
    '������ҳ��׼
    mintMecStandard = Val(zlDatabase.GetPara("������ҳ��׼", glngSys, pסԺҽ��վ, "0"))
    
    mlngSource = IIf(mbytSize = 1, 0, 999)

End Sub

Private Function InitDepts() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
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
        cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        If rsTmp!ID = lngPreDept Then '����ԭ�ж�λ
            Call Cbo.SetIndex(cboDept.hwnd, cboDept.NewIndex)
        ElseIf InStr(mstrPrivs, "ȫԺ����") > 0 Then
            If UserInfo.����ID = rsTmp!ID And (lngPreDept = 0 Or cboDept.ListIndex = -1) Then 'ֱ����������
                Call Cbo.SetIndex(cboDept.hwnd, cboDept.NewIndex)
            End If
            If InStr("," & strDeptIDs & ",", "," & rsTmp!ID & ",") > 0 And cboDept.ListIndex = -1 Then
                Call Cbo.SetIndex(cboDept.hwnd, cboDept.NewIndex)
            End If
        Else
            '����ȱʡ���������Ŀ����ж��
            If rsTmp!ȱʡ = 1 And cboDept.ListIndex = -1 Then
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

Private Function Get����ͼ�����(ByVal lng״̬ As Long) As Long
    Dim i As Long
    
    i = imgPati.ListImages("Pati").Index
    Select Case lng״̬
        Case 1
            i = imgPati.ListImages("�ȴ����").Index
        Case 2
            i = imgPati.ListImages("�ܾ����").Index
        Case 13
            i = imgPati.ListImages("���ڳ��").Index
        Case 3
            i = imgPati.ListImages("�������").Index
        Case 14
            i = imgPati.ListImages("��鷴��").Index
        Case 4
            i = imgPati.ListImages("��鷴��").Index
        Case 16
            i = imgPati.ListImages("�������").Index
        Case 6
            i = imgPati.ListImages("�������").Index
    End Select
    Get����ͼ����� = i - 1 '����Ǵ�0��ʼ��
End Function

Private Function LoadPatients() As Boolean
'���ܣ���ȡ�����б�
    Dim rsPati As New ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim objParent As ReportRecord
    Dim objPt As ReportRecord '���鸸���
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim objRow As ReportRow, intBedLen As Integer
    Dim strPatiRow As String, lngPatiRow As Long, blnTeam As Boolean, blnOutTeam As Boolean
    
    Dim strSQL As String, strMonitor As String
    Dim i As Long, j As Long
    Dim lngCount(0 To 7) As Long, strState As String    '������ʾ���˷���ͳ����Ŀ
    Dim str�������� As String, strFilter As String
    Dim strTmpDate As String                            'ת�����˲�ѯʱ�䷶Χ����
    Dim blnIsFind As Boolean                            '�ж��Ƿ��ǲ���סԺ�ż�סԺ���Ƿ�Ϊ��
    Dim strTmpOut As String                             '��ѯ��Ժ����
    Dim strICUSQL As String
    Dim strICUOutSQL As String                          'ת�������õľ���ҽʦ
    Dim lng������� As Long
    
    Dim rs��Ⱦ��״̬ As ADODB.Recordset
    Dim blnDo��Ⱦ��״̬ As Boolean, blnTeamGroup As Boolean, blnVisible��� As Boolean
    Dim rsBaby As ADODB.Recordset
    Dim strSQLBaby As String
    Dim objBabyParent As ReportRecord
    Dim lngCol As Long
    Dim strPre�������� As String
    Dim str�ֶ���� As String, strPreС���� As String
    Dim strTab���� As String, strС���� As String, str���� As String
    Dim str����״̬ As String
    
    mblnUnRefresh = True
    Screen.MousePointer = 11
    On Error GoTo errH
    
    '��ҳ����������գ�F5ˢ�£�Ӧ�ûָ���һ����ֵ
    If cboDept.ListIndex = -1 Then Call Cbo.SetIndex(cboDept.hwnd, Val(cboDept.Tag))
    '�ж��Ƿ��ǲ���סԺ�ż�סԺ���Ƿ�Ϊ��
    If mstrFindType = "סԺ��" And Trim(PatiIdentify.Text) <> "" Then blnIsFind = True
    
    If blnIsFind Then
        '����ȡ������ʾ��Ժ���˲��������Բ���ʱ��ʾ���ҵ���Ա��ʱ�䷶Χ����Ա
        If chkFilter.Value = 1 And chkFilter.Visible = True Then
            strTmpOut = " And (B.��Ժ���� Between to_date('" & Format(mdtOutBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And " & _
                        " to_date('" & Format(mdtOutEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS') And b.סԺ��=[6] And B.��Ժ���� is Not Null) "
        Else
            strTmpOut = " And (B.��Ժ���� Between to_date('" & Format(mdtOutBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And " & _
                        " to_date('" & Format(mdtOutEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')  Or (b.סԺ��=[6]  And B.��Ժ���� is Not Null )) "
        End If
    Else
        '�û����ǲ��ң���ʾ������Ӧʱ���ڵĲ���
        strTmpOut = " And B.��Ժ���� Between to_date('" & Format(mdtOutBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And to_date('" & Format(mdtOutEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS') "
    End If
    Set mrsPatiNotes = Nothing
    mstrAllPatis = ""
    blnTeam = chkByTeam.Value = 1
    blnOutTeam = chkOutByTeam.Value = 1
    '��λ���ȹ̶�Ϊ10
    intBedLen = 10
    
    'Start ��������
    str�������� = ""
    strFilter = ""
    strTmpDate = ""
    For i = 0 To chk��������.UBound
        If chk��������(i).Value = 1 Then
            str�������� = str�������� & "," & chk��������(i).Caption
        End If
    Next
    str�������� = Mid(str��������, 2)
    If Not (UBound(Split(str��������, ",")) = chk��������.UBound Or str�������� = "") Then
        strFilter = " And Instr(','||[4]||',',','||B.��ǰ����||',')>0"
    End If
    
    '����״̬
    str����״̬ = chkHZ(0).Value & chkHZ(1).Value
    
    If str����״̬ = "01" Then
        str����״̬ = " and d.ִ��״̬=1"
    ElseIf str����״̬ = "10" Then
        str����״̬ = " and nvl(d.ִ��״̬,0)<>1"
    Else
        str����״̬ = ""
    End If
    
    If mintChange = 0 Then
        strTmpDate = ""
    Else
        strTmpDate = " And C.��ֹʱ�� Between Sysdate-[3] And Sysdate "
    End If
    
    If mblnICU Then
        If InStr(mstrPrivs, "ȫԺ����") = 0 And InStr(mstrPrivs, "���Ʋ���") = 0 Then
            strICUSQL = " And B.סԺҽʦ=[2] "
            strICUOutSQL = " And C.����ҽʦ=[2] "

        ElseIf InStr(mstrPrivs, "ȫԺ����") > 0 Then
            strICUSQL = ""
            strICUOutSQL = strICUSQL
        Else
            strICUSQL = " And exists(select 1 from ���˱䶯��¼ x where " & _
                " b.����id=x.����id and b.��ҳid=x.��ҳid and x.��ֹԭ�� in(2,3,15) and instr(','||[7]||',' , ','||x.����id||',')>0) "
            strICUOutSQL = strICUSQL
        End If
        
    End If
    
    If cboDept.ListIndex <> -1 Then
        If tbcPati.Selected.Tag = "��Ժ" Or tbcPati.Selected.Tag = "��Ժ" Then
            str�ֶ���� = ",first_value(Decode(Sign(h.�������-10),-1,h.�������,'')) " & _
                " Over(partition By h.����id,H.��ҳID Order By sign(h.�������-10),decode(h.��¼��Դ,4,0,h.��¼��Դ) desc,Decode(h.�������," & Decode(tbcPati.Selected.Tag, "��Ժ", "1,1,2,2,3,3,0", "��Ժ", "3,3,0") & ") DESC,h.��ϴ���) As ��ҽ���"
            If Sys.DeptHaveProperty(cboDept.ItemData(cboDept.ListIndex), "��ҽ��") Then
                str�ֶ���� = str�ֶ���� & ",first_value(Decode(Sign(h.�������-10),1,h.�������,'')) " & _
                " Over(partition By h.����id,H.��ҳID Order By sign(h.�������-10) desc,decode(h.��¼��Դ,4,0,h.��¼��Դ) desc,Decode(h.�������," & Decode(tbcPati.Selected.Tag, "��Ժ", "11,1,12,2,13,3,0", "��Ժ", "13,3,0") & ") DESC,h.��ϴ���) As ��ҽ���"
            Else
                str�ֶ���� = str�ֶ���� & ",null as ��ҽ���"
            End If
        Else
            str�ֶ���� = ",Null As ��ҽ���,Null As ��ҽ���"
        End If
        If mintDeptView = 0 Then
            '��Ժ����
            If tbcPati.Selected.Tag = "��Ժ" Or tbcPati.Selected.Tag = "����ס" Then
                If blnTeam And tbcPati.Selected.Tag = "��Ժ" Then    '��С��ģʽ
                    strSQL = _
                        "Select Distinct Decode(B.״̬,1,0,3,3,Decode(G.ID,Null,2,1)) as ����,Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2,G.Id ��ID,G.���� as С����," & _
                        " Decode(B.״̬,1,'����ס����',3,'Ԥ��Ժ����',Decode(G.����,Null,'��Ժ����',G.����)) as ����,A.����ID,B.��ҳID,A.���￨��,B.���ۺ�,A.�����,B.סԺ��," & _
                        " NVL(B.����,A.����) ���� ,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����, A.����) ����,NULL as ����,B.סԺҽʦ,LPAD(B.��Ժ����," & intBedLen & ",' ') as ����," & _
                        " B.�ѱ�,Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��) As ��Ժ����,B.��Ժ����,B.��������,B.״̬,B.����,B.����״̬,-Null as ҽ��ID,-Null as ���ͺ�,-Null as ִ��״̬,-Null as ִ�п���ID,Null as ��������," & _
                        " Nvl(b.·��״̬,-1) ·��״̬,A.���֤��,trunc(sysdate)-trunc(Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��)) as סԺ����,q.��� As ���������,b.������,b.Ӥ������ID,B.Ӥ������ID" & str�ֶ���� & _
                        " From ������Ϣ A,������ҳ B,�ٴ�ҽ��С�� G,������������¼ Q,��Ժ���� R,������ϼ�¼ H" & _
                        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And b.����id=q.����ID(+) And b.��ҳID=q.��ҳID(+) And H.����id(+)=b.����id And h.��ҳid(+)=b.��ҳid And (q.���=1 Or q.��� is Null)" & _
                        " And (R.����ID=[1] Or b.Ӥ������ID=[1]) And a.����ID=R.����ID And  A.��ǰ����ID=R.����ID And B.ҽ��С��ID=G.ID(+)" & _
                        IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "And (B.ҽ��С��ID is Null or b.ҽ��С��id in(select id from �ٴ�ҽ��С�� where ����id =[1]))", _
                            " And (B.ҽ��С��ID is Null And B.סԺҽʦ=[2] or B.ҽ��С��ID in (Select С��id From ҽ��С����Ա Where ��Աid = [5]))") & _
                        strICUSQL & " And  B.���ʱ�� is NULL" & strFilter & _
                        IIf(tbcPati.Selected.Tag = "����ס", " And Nvl(B.״̬,0)=1", " And Nvl(B.״̬,0)<>1")
                    blnTeamGroup = True
                Else
                    strSQL = _
                        "Select Distinct Decode(B.״̬,1,0,3,3,Decode(B.סԺҽʦ,[2],1,2)) as ����,Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2,Null as ��ID," & _
                        " Decode(B.״̬,1,'����ס����',3,'Ԥ��Ժ����',Decode(B.סԺҽʦ,[2],'" & UserInfo.���� & "����Ժ����','��Ժ����')) as ����," & _
                        " A.����ID,B.��ҳID,A.���￨��,B.���ۺ�,A.�����,A.���֤��,B.סԺ��,NVL(B.����,A.����) ���� ,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����, A.����) ����,NULL as ����,B.סԺҽʦ," & _
                        " LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,B.�ѱ�,Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��) As ��Ժ����,B.��Ժ����,B.��������,B.״̬,B.����,B.����״̬," & _
                        " -Null as ҽ��ID,-Null as ���ͺ�,-Null as ִ��״̬,-Null as ִ�п���ID,Null as ��������,Nvl(b.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��)) as סԺ����,q.��� As ���������,b.������,b.Ӥ������ID,B.Ӥ������ID" & _
                        str�ֶ���� & _
                        " From ������Ϣ A,������ҳ B,������������¼ Q,��Ժ���� R,������ϼ�¼ H" & _
                        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID  And b.����id=q.����ID(+) And b.��ҳID=q.��ҳID(+) And H.����id(+)=b.����id And h.��ҳid(+)=b.��ҳid And (q.���=1 Or q.��� is Null)" & _
                        " And (R.����ID=[1] Or b.Ӥ������ID=[1]) And a.����ID=R.����ID And A.��ǰ����ID=R.����ID " & _
                        IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And B.סԺҽʦ=[2]") & _
                        strICUSQL & _
                        " And B.���ʱ�� is NULL" & strFilter & _
                        IIf(tbcPati.Selected.Tag = "����ס", " And B.״̬=1", " And B.״̬<>1")
                End If
                strSQLBaby = "Select q.����id, q.��ҳid, q.���, q.Ӥ������, q.Ӥ���Ա�,Round(Decode(q.����ʱ��, Null, Sysdate, q.����ʱ��) - q.����ʱ��)||'��' As ����" & _
                    " From ������Ϣ A,������ҳ B,������������¼ Q,��Ժ���� R" & _
                    " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID  And b.����id=q.����ID And b.��ҳID=q.��ҳID" & _
                    " And (R.����ID=[1] Or b.Ӥ������ID=[1]) And a.����ID=R.����ID And A.��ǰ����ID=R.����ID " & _
                    IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And B.סԺҽʦ=[2]") & _
                    strICUSQL & " And  B.���ʱ�� is NULL" & strFilter & _
                    IIf(tbcPati.Selected.Tag = "����ס", " And B.״̬=1", " And B.״̬<>1")
            ElseIf tbcPati.Selected.Tag = "��Ժ" Then
                '��Ժ����:��Ժ���˿������ж��סԺ
                If blnOutTeam Then
                    strSQL = _
                        "Select Distinct Decode(B.��Ժ��ʽ,'����',Decode(B.סԺҽʦ,[2],5.1,5.2), Decode(B.סԺҽʦ,[2],4.1,4.2)) as ����," & _
                        " Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2,G.Id as ��ID,G.���� as С����," & _
                        " Decode(B.��Ժ��ʽ,'����',Decode(B.סԺҽʦ,[2],'" & UserInfo.���� & "����������','������������'),Decode(B.סԺҽʦ,[2],'" & UserInfo.���� & "�ĳ�Ժ����',Decode(g.����, Null, '������Ժ����', g.����))) as ����," & _
                        " A.����ID,B.��ҳID,A.���￨��,B.���ۺ�,A.�����,B.סԺ��,A.���֤��,NVL(B.����,A.����) ���� ,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����, A.����) ����,NULL as ����,B.סԺҽʦ," & _
                        " LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,B.�ѱ�,Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��) As ��Ժ����,B.��Ժ����,B.��������,B.״̬,B.����,B.����״̬," & _
                        " -Null as ҽ��ID,-Null as ���ͺ�,-Null as ִ��״̬,-Null as ִ�п���ID,Null as ��������,Nvl(b.·��״̬,-1) ·��״̬,trunc(B.��Ժ����)-trunc(Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��)) as סԺ����," & _
                        " q.��� As ���������,b.������,b.Ӥ������ID,B.Ӥ������ID,B.���λ�ʿ" & str�ֶ���� & _
                        " From ������Ϣ A,������ҳ B,�ٴ�ҽ��С�� G,������������¼ Q,������ϼ�¼ H" & _
                        " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.��Ժ����ID+0=[1] And b.����id=q.����ID(+) And b.��ҳID=q.��ҳID(+) And H.����id(+)=b.����id And h.��ҳid(+)=b.��ҳid And (q.���=1 Or q.��� is Null) " & strTmpOut & _
                        IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "And (B.ҽ��С��ID is Null or b.ҽ��С��id in(select id from �ٴ�ҽ��С�� where ����id =[1]))", _
                            " And (B.ҽ��С��ID is Null And B.סԺҽʦ=[2] or B.ҽ��С��ID in(Select С��id From ҽ��С����Ա Where ��Աid = [5]))") & _
                        strICUSQL & " And B.���ʱ�� is NULL And b.ҽ��С��id = g.Id(+)"
                    blnTeamGroup = True
                Else
                    strSQL = _
                        "Select Distinct Decode(B.��Ժ��ʽ,'����',Decode(B.סԺҽʦ,[2],5.1,5.2),Decode(B.סԺҽʦ,[2],4.1,4.2)) as ����," & _
                        " Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2,Null as ��ID," & _
                        " Decode(B.��Ժ��ʽ,'����',Decode(B.סԺҽʦ,[2],'" & UserInfo.���� & "����������','������������'),Decode(B.סԺҽʦ,[2],'" & UserInfo.���� & "�ĳ�Ժ����','������Ժ����')) as ����," & _
                        " A.����ID,B.��ҳID,A.���￨��,B.���ۺ�,A.�����,B.סԺ��,A.���֤��,NVL(B.����,A.����) ���� ,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����, A.����) ����,NULL as ����,B.סԺҽʦ," & _
                        " LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,B.�ѱ�,Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��) As ��Ժ����,B.��Ժ����,B.��������,B.״̬,B.����,B.����״̬," & _
                        " -Null as ҽ��ID,-Null as ���ͺ�,-Null as ִ��״̬,-Null as ִ�п���ID,Null as ��������,Nvl(b.·��״̬,-1) ·��״̬,trunc(B.��Ժ����)-trunc(Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��)) as סԺ����," & _
                        " q.��� As ���������,b.������,b.Ӥ������ID,B.Ӥ������ID,B.���λ�ʿ" & str�ֶ���� & _
                        " From ������Ϣ A,������ҳ B,������������¼ Q,������ϼ�¼ H" & _
                        " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.��Ժ����ID+0=[1] And b.����id=q.����ID(+) And b.��ҳID=q.��ҳID(+) And H.����id(+)=b.����id And h.��ҳid(+)=b.��ҳid And (q.���=1 Or q.��� is Null) " & strTmpOut & _
                        IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And B.סԺҽʦ=[2]") & _
                        strICUSQL & " And B.���ʱ�� is NULL"
                End If
                strSQLBaby = "Select q.����id, q.��ҳid, q.���, q.Ӥ������, q.Ӥ���Ա�,Round(Decode(q.����ʱ��, Null, Sysdate, q.����ʱ��) - q.����ʱ��)||'��' As ����" & _
                    " From ������ҳ B,������������¼ Q" & _
                    " Where Nvl(B.��ҳID,0)<>0 And B.��Ժ����ID+0=[1] And b.����id=q.����ID  And b.��ҳID=q.��ҳID " & strTmpOut & _
                    IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And B.סԺҽʦ=[2]") & _
                    strICUSQL & " And B.���ʱ�� is NULL"
            ElseIf tbcPati.Selected.Tag = "����" Then
                '���ﲡ��:��Ժ
                If InStr(mstrPrivs, "���ﲡ��") > 0 And Not mblnICU Then
                    strSQL = _
                        "Select 6||Decode(D.ִ��״̬,1,decode(d.�����,[2],1,2),0) as ����,Decode(D.ִ��״̬,1,1,0) as ����2,Null as ��ID,decode(d.�����,[2],d.�����,null,null,'����ҽ��')||Decode(d.ִ��״̬,1,'����ɻ���','δ��ɻ���') as ����," & _
                        " A.����ID,B.��ҳID,A.���￨��,B.���ۺ�,A.�����,B.סԺ��,NVL(B.����,A.����) ���� ,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����, A.����) ����,G.���� as ����,P.���� as ����,B.סԺҽʦ," & _
                        " LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,A.���֤��,B.�ѱ�,Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��) As ��Ժ����,B.��Ժ����,B.��������,B.״̬,B.����,B.����״̬,q.��� As ���������,b.������,b.Ӥ������ID," & _
                        " B.Ӥ������ID" & str�ֶ���� & ",D.ҽ��ID,D.���ͺ�,Decode(D.ִ��״̬,1,1,0) As ִ��״̬,D.ִ�в���ID as ִ�п���ID,e.�������," & _
                        " G.����||'����,'||Decode(D.ִ��״̬,1,'�����'||E.ҽ������,'����'||To_Char(E.��ʼִ��ʱ��,'MM.DD HH24:MI')||'����'||E.ҽ������||Decode(E.ҽ������,NULL,NULL,'('||E.ҽ������||')')) As ��������," & _
                        " Nvl(b.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��)) as סԺ����,Decode(E.������־,1,'��','��') as ������־" & _
                        " From ������Ϣ A,������ҳ B,����ҽ������ D,����ҽ����¼ E,������ĿĿ¼ F,���ű� G,������������¼ Q,���ű� P" & _
                        " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And E.���˿���ID=G.ID" & _
                        IIf(chkOut.Value = 1, "", " And B.��Ժ���� is NULL") & " And Nvl(B.״̬,0)<>3  And b.����id=q.����ID(+) AND B.��ǰ����ID=P.ID(+) And b.��ҳID=q.��ҳID(+) And (q.���=1 Or q.��� is Null)" & _
                        " And B.����ID=E.����ID And B.��ҳID=E.��ҳID And D.ҽ��ID=E.ID And E.������ĿID=F.ID" & _
                        " And E.�������='Z' And F.��������='7' And E.ִ�п���id+0=[1] And E.��ʼִ��ʱ�� Between to_date('" & Format(mdtMeetBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
                        " And to_date('" & Format(mdtMeetEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')" & _
                        " And B.���ʱ�� is NULL" & str����״̬
                End If
            ElseIf tbcPati.Selected.Tag = "ת��" Then
                'ת������:ҽ���ʹ�����ʾ����ת��ǰ�ģ������ѳ�Ժ��
                strSQL = _
                    "Select /*+ RULE */Distinct 7 as ����,Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2,Null as ��ID,'ת������' as ����," & _
                    " A.����ID,B.��ҳID,A.���￨��,B.���ۺ�,A.�����,B.סԺ��,NVL(B.����,A.����) ���� ,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����, A.����) ����,NULL as ����,C.����ҽʦ as סԺҽʦ," & _
                    " LPAD(C.����," & intBedLen & ",' ') as ����,A.���֤��,B.�ѱ�,Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��) As ��Ժ����,B.��Ժ����,B.��������,B.״̬,B.����,B.����״̬," & _
                    " -Null as ҽ��ID,-Null as ���ͺ�,-Null as ִ��״̬,-Null as ִ�п���ID,Null as ��������,Nvl(b.·��״̬,-1) ·��״̬,trunc(Nvl(b.��Ժ����, Sysdate))-trunc(Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��)) as סԺ����,q.��� As ���������,b.������," & _
                    " b.Ӥ������ID,B.Ӥ������ID" & str�ֶ���� & _
                    " From ������Ϣ A,������ҳ B,���˱䶯��¼ C,������������¼ Q" & _
                    " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And b.����id=q.����ID(+) And b.��ҳID=q.��ҳID(+) And (q.���=1 Or q.��� is Null)" & _
                    " And Nvl(B.״̬,0)<>2 And B.��Ժ����ID<>[1] And Nvl(C.���Ӵ�λ,0)=0" & _
                    " And B.����ID=C.����ID And B.��ҳID=C.��ҳID And C.����ID+0=[1]" & _
                    " And C.��ֹԭ�� =3 " & strTmpDate & _
                    IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And C.����ҽʦ=[2]") & _
                    strICUOutSQL & _
                    " And B.���ʱ�� is NULL "
                
                strSQLBaby = "Select q.����id, q.��ҳid, q.���, q.Ӥ������, q.Ӥ���Ա�,Round(Decode(q.����ʱ��, Null, Sysdate, q.����ʱ��) - q.����ʱ��)||'��' As ����" & _
                    " From ������ҳ B,���˱䶯��¼ C,������������¼ Q" & _
                    " Where Nvl(B.��ҳID,0)<>0 And b.����id=q.����ID And b.��ҳID=q.��ҳID " & _
                    " And Nvl(B.״̬,0)<>2 And B.��Ժ����ID<>[1] And Nvl(C.���Ӵ�λ,0)=0" & _
                    " And B.����ID=C.����ID And B.��ҳID=C.��ҳID And C.����ID+0=[1]" & _
                    " And C.��ֹԭ�� =3 " & strTmpDate & _
                    IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And C.����ҽʦ=[2]") & _
                    strICUOutSQL & _
                    " And B.���ʱ�� is NULL "
            End If
        Else
            '�������鿴
            '��Ժ����
            If tbcPati.Selected.Tag = "��Ժ" Or tbcPati.Selected.Tag = "����ס" Then
                If blnTeam And tbcPati.Selected.Tag = "��Ժ" Then
                    strSQL = _
                        "Select Distinct Decode(B.״̬,1,0,3,3,Decode(G.����,Null,2,1)) as ����,Decode(Nvl(B.����״̬,0),2,'����ת��',0,999,B.����״̬) as ����2,G.Id ��ID," & _
                        " Decode(B.״̬,1,'����ס����',3,'Ԥ��Ժ����',Decode(G.����,Null,'��Ժ����',G.����)) as ����,G.���� as С����," & _
                        " A.����ID,B.��ҳID,A.���￨��,B.���ۺ�,A.�����,B.סԺ��,A.���֤��,NVL(B.����,A.����) ���� ,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����, A.����) ����,C.���� as ����,B.סԺҽʦ," & _
                        " LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,B.�ѱ�,Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��) As ��Ժ����,B.��Ժ����,B.��������,B.״̬,B.����,B.����״̬," & _
                        " -Null as ҽ��ID,-Null as ���ͺ�,-Null as ִ��״̬,-Null as ִ�п���ID,Null as ��������,Nvl(b.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��)) as סԺ����,q.��� As ���������,b.������," & _
                        " b.Ӥ������ID,B.Ӥ������ID" & str�ֶ���� & _
                        " From ������Ϣ A,������ҳ B,���ű� C,�ٴ�ҽ��С�� G,������������¼ Q,��Ժ���� R,������ϼ�¼ H" & _
                        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And b.����id=q.����ID(+) And b.��ҳID=q.��ҳID(+) And H.����id(+)=b.����id And h.��ҳid(+)=b.��ҳid And (q.���=1 Or q.��� is Null)" & _
                        " And (R.����ID=[1] Or b.Ӥ������ID=[1]) And B.��Ժ����ID=C.ID And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null) And a.����ID=R.����ID And A.��ǰ����ID=R.����ID And B.ҽ��С��ID=G.ID(+)" & _
                        IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "And (B.ҽ��С��ID is Null or b.ҽ��С��id in(Select g.Id From �ٴ�ҽ��С�� G, �������Ҷ�Ӧ I Where g.����id = i.����id And i.����id = [1]))", _
                            " And (B.ҽ��С��ID is Null And B.סԺҽʦ=[2] or B.ҽ��С��ID in(Select С��id From ҽ��С����Ա Where ��Աid = [5]))") & _
                        strICUSQL & _
                        " And B.���ʱ�� is NULL" & strFilter & _
                        IIf(tbcPati.Selected.Tag = "����ס", " And B.״̬=1", " And B.״̬<>1")
                    blnTeamGroup = True
                Else
                    strSQL = _
                        "Select Distinct Decode(B.״̬,1,0,3,3,Decode(B.סԺҽʦ,[2],1,2)) as ����,Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2,Null as ��ID," & _
                        " Decode(B.״̬,1,'����ס����',3,'Ԥ��Ժ����',Decode(B.סԺҽʦ,[2],'" & UserInfo.���� & "����Ժ����','��Ժ����')) as ����," & _
                        " A.����ID,B.��ҳID,A.���￨��,B.���ۺ�,A.�����,B.סԺ��,A.���֤��,NVL(B.����,A.����) ���� ,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����, A.����) ����,C.���� as ����,B.סԺҽʦ," & _
                        " LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,B.�ѱ�,Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��) As ��Ժ����,B.��Ժ����,B.��������,B.״̬,B.����,B.����״̬," & _
                        " -Null as ҽ��ID,-Null as ���ͺ�,-Null as ִ��״̬,-Null as ִ�п���ID,Null as ��������,Nvl(b.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��) ) as סԺ����,q.��� As ���������,b.������,b.Ӥ������ID,B.Ӥ������ID" & _
                        str�ֶ���� & _
                        " From ������Ϣ A,������ҳ B,���ű� C,������������¼ Q,��Ժ���� R,������ϼ�¼ H" & _
                        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And b.����id=q.����ID(+) And b.��ҳID=q.��ҳID(+) And H.����id(+)=b.����id And h.��ҳid(+)=b.��ҳid And (q.���=1 Or q.��� is Null)" & _
                        " And (R.����ID=[1] Or b.Ӥ������ID=[1]) And B.��Ժ����ID=C.ID And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null) And a.����ID=R.����ID And A.��ǰ����ID=R.����ID " & _
                        IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And B.סԺҽʦ=[2]") & _
                        strICUSQL & _
                        " And B.���ʱ�� is NULL" & strFilter & _
                        IIf(tbcPati.Selected.Tag = "����ס", " And B.״̬=1", " And B.״̬<>1")
                End If
                strSQLBaby = "Select q.����id, q.��ҳid, q.���, q.Ӥ������, q.Ӥ���Ա�,Round(Decode(q.����ʱ��, Null, Sysdate, q.����ʱ��) - q.����ʱ��)||'��' As ����" & _
                    " From ������Ϣ A,������ҳ B,���ű� C,������������¼ Q,��Ժ���� R" & _
                    " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And b.����id=q.����ID And b.��ҳID=q.��ҳID " & _
                    " And (R.����ID=[1] Or b.Ӥ������ID=[1]) And B.��Ժ����ID=C.ID And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null) And a.����ID=R.����ID And A.��ǰ����ID=R.����ID " & _
                    IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And B.סԺҽʦ=[2]") & _
                    strICUSQL & " And B.���ʱ�� is NULL" & strFilter & _
                    IIf(tbcPati.Selected.Tag = "����ס", " And B.״̬=1", " And B.״̬<>1")
            ElseIf tbcPati.Selected.Tag = "��Ժ" Then
                '��Ժ����:��Ժ���˿������ж��סԺ
                If blnOutTeam Then
                    strSQL = _
                        "Select Distinct Decode(B.��Ժ��ʽ,'����',Decode(B.סԺҽʦ,[2],5.1,5.2),Decode(B.סԺҽʦ,[2],4.1,4.2)) as ����," & _
                        " Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2,G.Id as ��ID,G.���� as С����," & _
                        " Decode(B.��Ժ��ʽ,'����',Decode(B.סԺҽʦ,[2],'" & UserInfo.���� & "����������','������������'),Decode(B.סԺҽʦ,[2],'" & UserInfo.���� & "�ĳ�Ժ����',Decode(g.����, Null, '������Ժ����', g.����))) as ����," & _
                        " A.����ID,B.��ҳID,A.���￨��,B.���ۺ�,A.�����,B.סԺ��,A.���֤��,NVL(B.����,A.����) ���� ,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����, A.����) ����,C.���� as ����,B.סԺҽʦ," & _
                        " LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,B.�ѱ�,Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��) As ��Ժ����,B.��Ժ����,B.��������,B.״̬,B.����,B.����״̬," & _
                        " -Null as ҽ��ID,-Null as ���ͺ�,-Null as ִ��״̬,-Null as ִ�п���ID,Null as ��������,Nvl(b.·��״̬,-1) ·��״̬,trunc(B.��Ժ����)-trunc(Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��)) as סԺ����,q.��� As ���������,b.������,b.Ӥ������ID,B.Ӥ������ID,B.���λ�ʿ" & _
                        str�ֶ���� & _
                        " From ������Ϣ A,������ҳ B,���ű� C,�ٴ�ҽ��С�� G,������������¼ Q,������ϼ�¼ H" & _
                        " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And b.����id=q.����ID(+) And b.��ҳID=q.��ҳID(+) And H.����id(+)=b.����id And h.��ҳid(+)=b.��ҳid And (q.���=1 Or q.��� is Null)" & _
                        " And B.��ǰ����ID+0=[1] And B.��Ժ����ID=C.ID And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null) " & strTmpOut & _
                        IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "And (B.ҽ��С��ID is Null or b.ҽ��С��id in(select id from �ٴ�ҽ��С�� where ����id =[1]))", _
                            " And (B.ҽ��С��ID is Null And B.סԺҽʦ=[2] or B.ҽ��С��ID in(Select С��id From ҽ��С����Ա Where ��Աid = [5]))") & _
                        strICUSQL & " And B.���ʱ�� is NULL And b.ҽ��С��id = g.Id(+)"
                    blnTeamGroup = True
                Else
                    strSQL = _
                        "Select Distinct Decode(B.��Ժ��ʽ,'����',Decode(B.סԺҽʦ,[2],5.1,5.2),Decode(B.סԺҽʦ,[2],4.1,4.2)) as ����," & _
                        " Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2,Null as ��ID," & _
                        " Decode(B.��Ժ��ʽ,'����',Decode(B.סԺҽʦ,[2],'" & UserInfo.���� & "����������','������������'),Decode(B.סԺҽʦ,[2],'" & UserInfo.���� & "�ĳ�Ժ����','������Ժ����')) as ����," & _
                        " A.����ID,B.��ҳID,A.���￨��,B.���ۺ�,A.�����,B.סԺ��,A.���֤��,NVL(B.����,A.����) ���� ,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����, A.����) ����,C.���� as ����,B.סԺҽʦ," & _
                        " LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,B.�ѱ�,Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��) As ��Ժ����,B.��Ժ����,B.��������,B.״̬,B.����,B.����״̬," & _
                        " -Null as ҽ��ID,-Null as ���ͺ�,-Null as ִ��״̬,-Null as ִ�п���ID,Null as ��������,Nvl(b.·��״̬,-1) ·��״̬,trunc(B.��Ժ����)-trunc(Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��)) as סԺ����,q.��� As ���������,b.������,b.Ӥ������ID,B.Ӥ������ID,B.���λ�ʿ" & _
                        str�ֶ���� & _
                        " From ������Ϣ A,������ҳ B,���ű� C,������������¼ Q,������ϼ�¼ H" & _
                        " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And b.����id=q.����ID(+) And b.��ҳID=q.��ҳID(+) And H.����id(+)=b.����id And h.��ҳid(+)=b.��ҳid And (q.���=1 Or q.��� is Null)" & _
                        " And B.��ǰ����ID+0=[1] And B.��Ժ����ID=C.ID And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null) " & strTmpOut & _
                        IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And B.סԺҽʦ=[2]") & _
                        strICUSQL & " And B.���ʱ�� is NULL"
                End If
                strSQLBaby = "Select q.����id, q.��ҳid, q.���, q.Ӥ������, q.Ӥ���Ա�,Round(Decode(q.����ʱ��, Null, Sysdate, q.����ʱ��) - q.����ʱ��)||'��' As ����" & _
                    " From ������ҳ B,���ű� C,������������¼ Q" & _
                    " Where Nvl(B.��ҳID,0)<>0 And b.����id=q.����ID And b.��ҳID=q.��ҳID " & _
                    " And B.��ǰ����ID+0=[1] And B.��Ժ����ID=C.ID And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null) " & strTmpOut & _
                    IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And B.סԺҽʦ=[2]") & _
                    strICUSQL & " And B.���ʱ�� is NULL"
            ElseIf tbcPati.Selected.Tag = "����" Then
                '���ﲡ��:��Ժ(����վ������)
                If InStr(mstrPrivs, "���ﲡ��") > 0 And Not mblnICU Then
                    strSQL = _
                        "Select 6||Decode(D.ִ��״̬,1,decode(d.�����,[2],1,2),0) as ����,Decode(D.ִ��״̬,1,1,0) as ����2,Null as ��ID,decode(d.�����,[2],d.�����,null,null,'����ҽ��')||Decode(d.ִ��״̬,1,'����ɻ���','δ��ɻ���') as ����," & _
                        " A.����ID,B.��ҳID,A.���￨��,B.���ۺ�,A.�����,B.סԺ��,A.���֤��,NVL(B.����,A.����) ���� ,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����, A.����) ����,G.���� as ����,P.���� as ����,B.סԺҽʦ," & _
                        " LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,B.�ѱ�,Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��) As ��Ժ����,B.��Ժ����,B.��������,B.״̬,B.����,B.����״̬,q.��� As ���������,b.������,b.Ӥ������ID ,B.Ӥ������ID" & str�ֶ���� & "," & _
                        " D.ҽ��ID,D.���ͺ�,Decode(D.ִ��״̬,1,1,0) As ִ��״̬,D.ִ�в���ID as ִ�п���ID,e.�������," & _
                        " G.����||'����,'||Decode(D.ִ��״̬,1,'�����'||E.ҽ������,'����'||To_Char(E.��ʼִ��ʱ��,'MM.DD HH24:MI')||'����'||E.ҽ������||Decode(E.ҽ������,NULL,NULL,'('||E.ҽ������||')')) As ��������,Nvl(b.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��)) as סԺ����,Decode(E.������־,1,'��','��') as ������־" & _
                        " From ������Ϣ A,������ҳ B,����ҽ������ D,����ҽ����¼ E,������ĿĿ¼ F,���ű� G,������������¼ Q,���ű� P" & _
                        " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And E.���˿���ID=G.ID  AND B.��ǰ����ID=P.ID(+) And b.����id=q.����ID(+) And b.��ҳID=q.��ҳID(+) And (q.���=1 Or q.��� is Null)" & _
                        IIf(chkOut.Value = 1, "", " And B.��Ժ���� is NULL") & " And Nvl(B.״̬,0)<>3 " & _
                        " And B.����ID=E.����ID And B.��ҳID=E.��ҳID And D.ҽ��ID=E.ID And E.������ĿID=F.ID" & _
                        " And E.�������='Z' And F.��������='7' And E.��ʼִ��ʱ�� Between to_date('" & Format(mdtMeetBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
                        " And to_date('" & Format(mdtMeetEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')" & _
                        " And E.ִ�п���id+0 IN (Select ����ID From �������Ҷ�Ӧ Where ����ID=[1])" & _
                        " And B.���ʱ�� is NULL" & str����״̬
                End If
            ElseIf tbcPati.Selected.Tag = "ת��" Then
                'ת������:ҽ���ʹ�����ʾ����ת��ǰ�ģ������ѳ�Ժ��(����վ������)
                strSQL = _
                    "Select /*+ RULE */Distinct 7 as ����,Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2,Null as ��ID,'ת������' as ����," & _
                    " A.����ID,B.��ҳID,A.���￨��,B.���ۺ�,A.�����,B.סԺ��,A.���֤��,NVL(B.����,A.����) ���� ,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����, A.����) ����,D.���� as ����,C.����ҽʦ as סԺҽʦ," & _
                    " LPAD(C.����," & intBedLen & ",' ') as ����,B.�ѱ�,Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��) As ��Ժ����,B.��Ժ����,B.��������,B.״̬,B.����,B.����״̬," & _
                    " -Null as ҽ��ID,-Null as ���ͺ�,-Null as ִ��״̬,-Null as ִ�п���ID,Null as ��������,Nvl(b.·��״̬,-1) ·��״̬,trunc(Nvl(b.��Ժ����, Sysdate))-trunc(Decode(B.���ʱ��,NULL,B.��Ժ����,B.���ʱ��)) as סԺ����,q.��� As ���������,b.������,b.Ӥ������ID,B.Ӥ������ID" & _
                    str�ֶ���� & _
                    " From ������Ϣ A,������ҳ B,���˱䶯��¼ C,���ű� D,������������¼ Q" & _
                    " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.��Ժ����ID=D.ID And b.����id=q.����ID(+) And b.��ҳID=q.��ҳID(+) And (q.���=1 Or q.��� is Null) " & _
                    " And Nvl(B.״̬,0)<>2 And B.��ǰ����ID<>[1] And Nvl(C.���Ӵ�λ,0)=0" & _
                    " And B.����ID=C.����ID And B.��ҳID=C.��ҳID And C.����ID+0=[1]" & _
                    " And C.��ֹԭ�� =3 " & strTmpDate & _
                    IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And C.����ҽʦ=[2]") & _
                    strICUOutSQL & _
                    " And B.���ʱ�� is NULL "
                    
                strSQLBaby = "Select q.����id, q.��ҳid, q.���, q.Ӥ������, q.Ӥ���Ա�,Round(Decode(q.����ʱ��, Null, Sysdate, q.����ʱ��) - q.����ʱ��)||'��' As ����" & _
                    " From ������ҳ B,���˱䶯��¼ C,���ű� D,������������¼ Q" & _
                    " Where Nvl(B.��ҳID,0)<>0 And B.��Ժ����ID=D.ID And b.����id=q.����ID And b.��ҳID=q.��ҳID" & _
                    " And Nvl(B.״̬,0)<>2 And B.��ǰ����ID<>[1] And Nvl(C.���Ӵ�λ,0)=0" & _
                    " And B.����ID=C.����ID And B.��ҳID=C.��ҳID And C.����ID+0=[1]" & _
                    " And C.��ֹԭ�� =3 " & strTmpDate & _
                    IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And C.����ҽʦ=[2]") & _
                    strICUOutSQL & _
                    " And B.���ʱ�� is NULL "
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
            strSQL = strSQL & " Order by С����,����,����2,����,��ҳID Desc"
        Else
            strSQL = strSQL & " Order by ����,����2,����,��ҳID Desc"
        End If
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDept.ItemData(cboDept.ListIndex), UserInfo.����, _
            mintChange, str��������, UserInfo.ID, Val(Trim(PatiIdentify.Text)), mstrUserDeps)
        If strSQLBaby <> "" Then
            Set rsBaby = zlDatabase.OpenSQLRecord(strSQLBaby, Me.Caption, cboDept.ItemData(cboDept.ListIndex), UserInfo.����, _
                mintChange, str��������, UserInfo.ID, Val(Trim(PatiIdentify.Text)), mstrUserDeps)
        End If
        
        strSQL = ""
        If tbcPati.Selected.Tag = "��Ժ" Or tbcPati.Selected.Tag = "��Ժ" Then
            If mintDeptView = 0 Then
                If tbcPati.Selected.Tag = "��Ժ" Then
                    strTab���� = "Select b.����ID,B.��ҳID From ������ҳ B,��Ժ���� R " & _
                        " Where (R.����ID=[1] Or b.Ӥ������ID=[1]) And b.����ID=R.����ID And b.��Ժ����ID+0=R.����ID and b.��Ժ���� is null " & _
                        IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And B.סԺҽʦ=[2]") & _
                        strICUSQL & " And B.���ʱ�� is NULL" & strFilter & " And B.״̬<>1"
                Else
                    strTab���� = "Select b.����id, b.��ҳid" & vbNewLine & _
                                "From ������ҳ B" & vbNewLine & _
                                "Where (b.��Ժ����ID = [1] Or b.Ӥ������ID = [1]) And b.��Ժ���� Is Not Null And b.���ʱ�� Is Null " & strTmpOut
                End If
            Else
                If tbcPati.Selected.Tag = "��Ժ" Then
                    strTab���� = "Select b.����ID,B.��ҳID From ������ҳ B,��Ժ���� R " & _
                        " Where (R.����ID=[1] Or b.Ӥ������ID=[1]) And b.����ID=R.����ID And b.��Ժ����ID+0=R.����ID and b.��Ժ���� is null " & _
                        IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And B.סԺҽʦ=[2]") & _
                        strICUSQL & " And B.���ʱ�� is NULL" & strFilter & " And B.״̬<>1"
                Else
                    strTab���� = "Select b.����id, b.��ҳid" & vbNewLine & _
                                "From ������ҳ B" & vbNewLine & _
                                "Where (b.��ǰ����id = [1] Or b.Ӥ������id = [1]) And b.��Ժ���� Is Not Null And b.���ʱ�� Is Null " & strTmpOut
                End If
            End If
            strSQL = "select  m.����id,m.��ҳid,max(m.��¼) as ��¼,max(m.��д) as ��д,max(m.״̬) as ״̬ from " & vbNewLine & _
                "( " & _
                "select a.����id,a.��ҳid,1 as ��¼,0 as ��д,0 as ״̬ from ( " & strTab���� & ") a " & vbNewLine & _
                "where exists(select 1 from �������Լ�¼ b where a.����id=b.����id and a.��ҳid=b.��ҳid) " & vbNewLine & _
                "union all " & vbNewLine & _
                "Select  a.����id,a.��ҳid,0 as ��¼,1 as ��д,0 as ״̬ From ( " & strTab���� & ") a " & vbNewLine & _
                "Where exists( select 1 from  ���Ӳ�����¼ C where c.�������� = 5 And a.����id = c.����id And a.��ҳid = c.��ҳid  and c.�������� like '%��Ⱦ��%') " & vbNewLine & _
                "union all " & vbNewLine & _
                "Select  a.����id,a.��ҳid,0 as ��¼,1 as ��д,e.����״̬ as ״̬ From ( " & strTab���� & ") a ,�����걨��¼ E " & vbNewLine & _
                "Where a.����ID=E.����ID and A.��ҳID=e.��ҳID ) M " & vbNewLine & _
                "group by m.����id,m.��ҳid "
        End If
        
        If strSQL <> "" Then
            Set rs��Ⱦ��״̬ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDept.ItemData(cboDept.ListIndex), UserInfo.����, _
                mintChange, str��������, UserInfo.ID, Val(Trim(PatiIdentify.Text)), mstrUserDeps)
            If rs��Ⱦ��״̬.RecordCount > 0 Then blnDo��Ⱦ��״̬ = True
        End If
    End If
    
    If Not rsPati.EOF Then
        '��¼����ѡ�еĲ���
        If rptPati.SelectedRows.Count > 0 Then
            If Not rptPati.SelectedRows(0).GroupRow Then
                If rptPati.SelectedRows(0).Record.Tag <> "" Then
                    lngPatiRow = rptPati.SelectedRows(0).Index '���ڿ������¶�λ
                    strPatiRow = rptPati.SelectedRows(0).Record.Tag
                End If
            End If
        End If
        If mclsWardMonitor.Enabled And InStr(GetInsidePrivs(pסԺҽ��վ), "����໤") > 0 Then
            strMonitor = mclsWardMonitor.GetListPati
        End If
    End If
    
    rptPati.Records.DeleteAll
    
    If tbcPati.Selected.Tag = "��Ժ" Then
        rptPati.Columns(col_���λ�ʿ).Visible = True
    Else
        rptPati.Columns(col_���λ�ʿ).Visible = False
    End If
    
    'ˢ�º�����Զ�չ��
    For i = 1 To rsPati.RecordCount
        str���� = rsPati!���� & ""
        If blnTeamGroup Then
            strС���� = rsPati!С���� & ""
        Else
            strС���� = ""
        End If
        'С�������Ӹ���
        If str���� <> strС���� And strС���� <> "" Then
            blnVisible��� = True
            str���� = strС����
            If strС���� <> strPreС���� Then
                Set objPt = Nothing
                strPreС���� = strС����
            End If
            If objPt Is Nothing Then
                Set objPt = Me.rptPati.Records.Add()
            ElseIf objPt.Tag <> CStr(rsPati!��ID & "��") Then
                Set objPt = Me.rptPati.Records.Add()
            End If
            If objPt.Tag <> CStr(rsPati!��ID & "��") Then
                objPt.Tag = CStr(rsPati!��ID & "��")
                objPt.Expanded = True
                For j = 0 To rptPati.Columns.Count - 1
                    If j = col_���� Then
                        If IsNull(rsPati!��ID) Then
                            Set objItem = objPt.AddItem(Val(rsPati!����))
                        Else
                            Set objItem = objPt.AddItem(-1 * Val(rsPati!��ID))
                        End If
                        objItem.Caption = str����
                    ElseIf j = col_���� Then
                        Set objItem = objPt.AddItem(rsPati!���� & "")
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
        
        '�����ύ��������Ӹ���
        If NVL(rsPati!����״̬, 0) <> 0 And Val(Mid(rsPati!����, 1, 1)) <> pt���� Then
            blnVisible��� = True
            '���������ͷ����仯ʱҪ���⿪һ����֧������ᵼ�·��鲻��
            If strPre�������� <> rsPati!���� & "" Then
                Set objParent = Nothing
                strPre�������� = rsPati!���� & ""
            End If
            
            If objParent Is Nothing Then
                If objPt Is Nothing Then
                    Set objParent = Me.rptPati.Records.Add()
                Else
                    Set objParent = objPt.Childs.Add()
                End If
            ElseIf objParent.Tag <> CStr(rsPati!����״̬) Then
                Set objParent = Me.rptPati.Records.Add()
            End If
            If objParent.Tag <> CStr(rsPati!����״̬) Then
                objParent.Tag = CStr(rsPati!����״̬)
                objParent.Expanded = True
                For j = 0 To rptPati.Columns.Count - 1
                    If j = col_���� Then
                        If IsNull(rsPati!��ID) Then
                            Set objItem = objParent.AddItem(Val(rsPati!����))
                        Else
                            Set objItem = objParent.AddItem(-1 * Val(rsPati!��ID))
                        End If
                        objItem.Caption = str����
                    ElseIf j = col_��� Then
                        Set objItem = objParent.AddItem(Val(rsPati!����״̬))
                        objItem.Caption = " "
                    ElseIf j = col_���� Then
                        Set objItem = objParent.AddItem(CStr(Decode(rsPati!����״̬, 1, "�ȴ����", 2, "�ܾ����", 13, "���ڳ��", 3, "�������", 14, "��鷴��", 4, "��鷴��", 16, "���������", 6, "���������")))
                        objItem.ForeColor = rptPati.PaintManager.GroupForeColor
                    Else
                        Set objItem = objParent.AddItem("")
                        If j = col_ͼ�� Then objItem.Icon = Get����ͼ�����(rsPati!����״̬) 'rsPati!����״̬ + imgPati.ListImages("�ȴ����").Index - 2
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
        
        objRecord.Tag = CStr(rsPati!����ID & "," & rsPati!��ҳID) '���ڲ��˶�λ
        
        If IsNull(rsPati!��ID) Then
            Set objItem = objRecord.AddItem(Val(rsPati!����)) '������Value��������
        Else
            Set objItem = objRecord.AddItem(-1 * Val(rsPati!��ID))  '������Value��������
        End If
        objItem.Caption = str����

        Set objItem = objRecord.AddItem(Val(Decode(NVL(rsPati!����״̬, 0), 0, 999, rsPati!����״̬)))
        objItem.Caption = " "
        If NVL(rsPati!����״̬, 0) = 2 Then
            objRecord.PreviewText = "  ����:" & GetRefuseReason(rsPati!����ID, rsPati!��ҳID)
        End If

        'ͼ��:ע�����������Ǵ�0��ʼ��š�
        '     ͼ��Value���ڴ���Ƿ����ύ��飬����Ŷ�ȡ
        Set objItem = objRecord.AddItem(-1)
        objItem.Caption = " "
        If NVL(rsPati!����״̬, 0) <> 0 Then
            objItem.Icon = Get����ͼ�����(rsPati!����״̬)
        ElseIf "" & rsPati!������ <> "" Then
            objItem.Icon = imgPati.ListImages("������").Index - 1
        End If

        'lng·��״̬=-1:δ����,0-�����ϵ���������1-ִ���У�2-����������3-�������
        Set objItem = objRecord.AddItem(Val("" & rsPati!·��״̬))
        objItem.Caption = " "
        objItem.Icon = -1 + Choose(rsPati!·��״̬ + 2, imgPati.ListImages("δ����").Index, imgPati.ListImages("������").Index, _
            imgPati.ListImages("ִ����").Index, imgPati.ListImages("��������").Index, imgPati.ListImages("�������").Index)
        
        objRecord.AddItem Val(rsPati!����ID)
        objRecord.AddItem Val(rsPati!��ҳID)
        objRecord.AddItem CStr(NVL(rsPati!����))

        If mblnOutDept Then
            Set objItem = objRecord.AddItem("" & rsPati!�����)
            objItem.Caption = NVL(rsPati!�����, " ")
        Else
            Set objItem = objRecord.AddItem("" & rsPati!סԺ��)
            objItem.Caption = NVL(rsPati!סԺ��, " ")
        End If
        
        Set objItem = objRecord.AddItem(zlStr.Lpad(NVL(rsPati!����), 10)) 'Value��������
        objItem.Caption = CStr(Trim(NVL(rsPati!����, " "))) 'Ϊ��ʱ�ᱻValue���
        '������ͼ��
        If rsPati!��������� & "" <> "" Then
            objItem.Icon = imgPati.ListImages("Child").Index - 1
            If rsPati!Ӥ������ID & "" <> "" Then objItem.Icon = imgPati.ListImages("Out").Index - 1
        End If
        Set objItem = objRecord.AddItem(" ")    '����໤
        If strMonitor <> "" And Not IsNull(rsPati!סԺ��) Then
            If InStr("," & strMonitor & ",", "," & rsPati!סԺ�� & ",") > 0 Then
                objItem.Caption = "��"
            End If
        End If
        
        objRecord.AddItem CStr(NVL(rsPati!�Ա�))
        objRecord.AddItem CStr(NVL(rsPati!����))
        objRecord.AddItem CStr(NVL(rsPati!�ѱ�))
        objRecord.AddItem CStr(NVL(rsPati!����))
        If tbcPati.Selected.Tag = "����" Then
            objRecord.AddItem CStr(NVL(rsPati!����))
        Else
            objRecord.AddItem ""
        End If
        objRecord.AddItem CStr(NVL(rsPati!סԺҽʦ))
        objRecord.AddItem Format(rsPati!��Ժ����, "yyyy-MM-dd HH:mm")
        objRecord.AddItem Format(NVL(rsPati!��Ժ����), "yyyy-MM-dd HH:mm")
        objRecord.AddItem NVL(rsPati!��������)
        
        '���ڻ��ﲡ��
        objRecord.AddItem Val(NVL(rsPati!ҽ��ID, 0))
        objRecord.AddItem Val(NVL(rsPati!���ͺ�, 0))
        objRecord.AddItem Val(NVL(rsPati!ִ��״̬, 0))
        objRecord.AddItem Val(NVL(rsPati!ִ�п���ID, 0))
        objRecord.AddItem CStr(NVL(rsPati!���￨��))
        objRecord.AddItem Val(Trim(IIf(CStr("" & rsPati!סԺ����) = "0", "1", CStr("" & rsPati!סԺ����))))
        objRecord.AddItem "" & rsPati!������
        objRecord.AddItem Val("" & rsPati!Ӥ������ID)
        objRecord.AddItem Val("" & rsPati!Ӥ������ID)
        
        '������
        objRecord.AddItem CStr(NVL(rsPati!��ҽ���))
        objRecord.AddItem CStr(NVL(rsPati!��ҽ���))
        
        If tbcPati.Selected.Tag = "����" Then
            lng������� = Val("" & rsPati!�������)
        Else
            lng������� = 0
        End If
        objRecord.AddItem lng�������
        
        '��Ӵ�Ⱦ��״̬
        strSQL = ""
        If blnDo��Ⱦ��״̬ Then
            rs��Ⱦ��״̬.Filter = "����ID=" & Val(rsPati!����ID) & " and ��ҳID=" & Val(rsPati!��ҳID)
            If Not rs��Ⱦ��״̬.EOF Then strSQL = Get��Ⱦ��״̬(Val(rs��Ⱦ��״̬!��¼ & ""), Val(rs��Ⱦ��״̬!��д & ""), Val(rs��Ⱦ��״̬!״̬ & ""))
        End If
        objRecord.AddItem strSQL
        If tbcPati.Selected.Tag = "��Ժ" Then
            objRecord.AddItem rsPati!���λ�ʿ & ""
        Else
            '���������
            objRecord.AddItem ""
        End If
        
        objRecord.AddItem "" & rsPati!���ۺ�
        objRecord.AddItem "" & rsPati!���֤��
        If tbcPati.Selected.Tag = "����" Then
            objRecord.AddItem "" & rsPati!������־
        Else
            objRecord.AddItem ""
        End If
        
        '��ʾ������ɫ
        objRecord.Item(col_����).ForeColor = zlDatabase.GetPatiColor(NVL(rsPati!��������))
        For j = 0 To rptPati.Columns.Count - 1
            If j <> col_���� And j <> col_��� And j <> col_ͼ�� Then
                objRecord.Item(j).ForeColor = objRecord.Item(col_����).ForeColor
            End If
        Next
        
        '����ɵĻ��ﲡ���û�ɫ��ʾ
        If Val(Mid(rsPati!����, 1, 1)) = pt���� Then
            objRecord.PreviewText = "  " & rsPati!��������
            If NVL(rsPati!ִ��״̬, 0) = 1 Then
                For j = 0 To rptPati.Columns.Count - 1
                    objRecord.Item(j).ForeColor = &H808080
                Next
                objRecord.Item(col_ͼ��).Icon = 2
            Else
                objRecord.Item(col_ͼ��).Icon = 1
            End If
        End If
        'ͳ�Ʋ�����Ŀ
        lngCount(Val(Mid(rsPati!����, 1, 1))) = lngCount(Val(Mid(rsPati!����, 1, 1))) + 1
        
        '�����Ƿ���Ӥ�����Ӥ����
        If Not rsBaby Is Nothing Then
            Set objBabyParent = objRecord
            rsBaby.Filter = "����ID=" & objBabyParent(col_����Id).Value & " and ��ҳID=" & objBabyParent(col_��ҳID).Value
            If Not rsBaby.EOF Then
                blnVisible��� = True
                rsBaby.Sort = "���"
                objBabyParent.Expanded = False
                For lngCol = 1 To rsBaby.RecordCount
                    Set objRecord = objBabyParent.Childs.Add()
                    objRecord.Tag = objBabyParent.Tag & "_" & rsBaby!���
                    For j = 0 To rptPati.Columns.Count - 1
                        Set objItem = objRecord.AddItem(objBabyParent(j).Value)
                            objItem.Caption = " "
                            objItem.ForeColor = objBabyParent(j).ForeColor
                        Select Case j
                        Case col_����
                            objItem.Caption = "   " & rsBaby!Ӥ������
                        Case col_�Ա�
                            objItem.Caption = "" & rsBaby!Ӥ���Ա�
                        Case col_סԺ��
                            objItem.Caption = objBabyParent(j).Value & "-" & rsBaby!���
                        Case col_����
                            objItem.Caption = " "
                            If "" & rsBaby!Ӥ���Ա� = "��" Then
                                objItem.Icon = imgPati.ListImages("Child").Index - 1
                            Else
                                objItem.Icon = imgPati.ListImages("Fbaby").Index - 1
                            End If
                            If lngCol = 1 And objBabyParent(j).Icon = imgPati.ListImages("Child").Index - 1 Then
                                objBabyParent(j).Icon = objItem.Icon
                            End If
                        Case col_����
                            objItem.Caption = "" & rsBaby!����
                        End Select
                    Next
                    rsBaby.MoveNext
                Next
                objBabyParent.Tag = objBabyParent.Tag & "_0"
            End If
        End If
        
        '���Ų��˹ؼ���Ϣ
        mstrAllPatis = mstrAllPatis & "," & rsPati!����ID & ":" & rsPati!��ҳID
        rsPati.MoveNext
    Next
     
    Call ShowAllPatiͼ��(mstrAllPatis)

    If tbcPati.Selected.Tag = "��Ժ" Then
        rptPati.Columns.Find(col_��ҽ���).Visible = True
        rptPati.Columns.Find(col_��ҽ���).Caption = "��ҽ���"
        rptPati.Columns.Find(col_��ҽ���).Visible = Sys.DeptHaveProperty(cboDept.ItemData(cboDept.ListIndex), "��ҽ��")
        rptPati.Columns.Find(COL_��Ⱦ��).Visible = True
    ElseIf tbcPati.Selected.Tag = "��Ժ" Then
        rptPati.Columns.Find(col_��ҽ���).Visible = True
        rptPati.Columns.Find(col_��ҽ���).Caption = "��Ժ���"
        rptPati.Columns.Find(col_��ҽ���).Visible = False
        rptPati.Columns.Find(COL_��Ⱦ��).Visible = True
    Else
        rptPati.Columns.Find(col_��ҽ���).Visible = False
        rptPati.Columns.Find(col_��ҽ���).Visible = False
        rptPati.Columns.Find(COL_��Ⱦ��).Visible = False
    End If
    
    
    rptPati.Columns.Find(col_�Ƿ���).Visible = IIf(tbcPati.Selected.Tag = "����", True, False)
    rptPati.Columns.Find(col_����).Visible = IIf(tbcPati.Selected.Tag = "����", True, False)
    
    
    rptPati.Columns(col_���).Visible = blnVisible���
    If tbcPati.Selected.Tag = "����" Then
        rptPati.Columns.Find(col_����).Visible = True
    Else
        rptPati.Columns.Find(col_����).Visible = mintDeptView = 1
    End If
    If mblnOutDept Then
        rptPati.Columns.Find(col_סԺ��).Caption = "�����"
    Else
        rptPati.Columns.Find(col_סԺ��).Caption = "סԺ��"
    End If
    rptPati.Populate
    '���ݽ���ҽԺ������벡����Ŀͳ����Ϣ
    strState = " �� " & rsPati.RecordCount & " ������"
    For i = LBound(lngCount) To UBound(lngCount)
        If lngCount(i) > 0 Then
            Select Case i
            Case 0
                strState = strState & "������ס:"
            Case 1
                If blnTeam Then
                    strState = strState & "��ҽ��С��:"
                Else
                    strState = strState & IIf(tbcPati.Selected.Tag = "��Ժ", "��" & UserInfo.���� & "����Ժ:", "��" & UserInfo.���� & "�ĳ�Ժ:")
                End If
            Case 2
                strState = strState & "��������Ժ:"
            Case 3
                strState = strState & "������Ԥ��Ժ:"
            Case 4
                strState = strState & "�����Ƴ�Ժ:"
            Case 5
                strState = strState & "����������:"
            Case 6
                strState = strState & "������:"
            Case 7
                strState = strState & "��ת��:"
            End Select
            strState = strState & lngCount(i) & "��"
        End If
    Next
    stbThis.Panels(2).Text = strState
    stbThis.Panels(2).Tag = strState
    lblFee(1).Caption = ""
    
    '��λ������:��Populate֮��
    mstrPrePati = ""
    If rptPati.Rows.Count = 0 Or rsPati.RecordCount > 1 And lngPatiRow = 0 Then
        Call ClearPatiInfo
        '��������ˢ���Ӵ���
        Call SubWinRefreshData(tbcSub.Selected)
        
        mstrPreNotify = ""
        rptNotify.Records.DeleteAll
        rptNotify.Populate
        rptNotify.TabStop = False
        
        strState = " û�в���"
    Else
        'ȡָ��������
        If strPatiRow <> "" Then
            '�ȿ��ٶ�λ
            If lngPatiRow <= rptPati.Rows.Count - 1 Then
                If Not rptPati.Rows(lngPatiRow).GroupRow Then
                    If rptPati.Rows(lngPatiRow).Record.Tag = strPatiRow Then
                        Set objRow = rptPati.Rows(lngPatiRow)
                    End If
                End If
            End If
            '�ٽ��в���
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
        'ȡ��һ���Ƿ�����
        If objRow Is Nothing Then
            For i = 0 To rptPati.Rows.Count - 1
                If Not rptPati.Rows(i).GroupRow And rptPati.Rows(i).Childs.Count = 0 Then Set objRow = rptPati.Rows(i): Exit For
            Next
        End If
        Set rptPati.FocusedRow = objRow '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
    
    End If
    
    Screen.MousePointer = 0
    LoadPatients = True
    
    'ͬ��ˢ����鷴����Ϣ
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
'���ܣ��������������ص���ʾ��Ϣ
    
    Dim i As Long
    
    mlng����ID = 0
    mlng��ҳID = 0
    
    mPatiInfo.״̬ = 0
    mPatiInfo.סԺ�� = ""
    mPatiInfo.���� = ""
    mPatiInfo.Ӥ�� = -1
    mPatiInfo.��ҳID = 0
    mPatiInfo.����ID = 0
    mPatiInfo.����ID = 0
    mPatiInfo.��Ժ���� = CDate(0)
    mPatiInfo.��Ժ���� = CDate(0)
    mPatiInfo.��Ŀ���� = CDate(0)
    mPatiInfo.סԺ���� = 0
    mPatiInfo.����ת�� = False
    Set mPatiInfo.rsͼ�� = Nothing
        
    cboPages.Clear
    cbo����.Clear
    lbl����(1).Caption = ""
    lblPatiName(1).Caption = ""
    lblPatiName(1).ToolTipText = ""
    lbl����(1).Caption = ""
    lblҽ����(1).Caption = ""
    lbl����(1).Caption = ""
    lbl����(1).Caption = ""
    lbl����(1).Caption = ""
    lbl��Ժ(1).Caption = ""
    lblDiag(1).Caption = ""
    lbl����(1).Caption = ""
    lblFee(1).Caption = ""
    lblFluid(1).Caption = ""
    lblPrint(1).Caption = ""
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean, Optional ByVal lngPatiID As Long)
'���ܣ�����(��һ��)����
'������blnNext=�Ƿ������һ��
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long
    
    '��ʼ������
    If rptPati.SelectedRows.Count > 0 Then
        If Not rptPati.SelectedRows(0).GroupRow Then
            If Val(rptPati.SelectedRows(0).Record(col_����Id).Value) <> 0 Then blnHave = True
        End If
    End If
    If Not blnNext Or blnReStart Or Not blnHave Then
        i = 0 'ReportControl����������0��ʼ
    Else
        i = rptPati.SelectedRows(0).Index + 1
    End If
    
    '���Ҳ���
    For i = i To rptPati.Rows.Count - 1
        With rptPati.Rows(i)
            If Not .GroupRow Then
                If Val(.Record(col_����Id).Value) = lngPatiID And lngPatiID <> 0 Then Exit For
                If mstrFindType = "סԺ��" Then 'סԺ��
                    If .Record(col_סԺ��).Value = PatiIdentify.Text Then Exit For
                ElseIf mstrFindType = "���ۺ�" Then
                    If .Record(col_���ۺ�).Value = PatiIdentify.Text Then Exit For
                ElseIf mstrFindType = "����" Then '����
                    If UCase(Trim(.Record(col_����).Value)) = UCase(PatiIdentify.Text) Then Exit For
                ElseIf mstrFindType = "���￨" Then '���￨
                    If UCase(.Record(col_���￨).Value) = UCase(PatiIdentify.Text) Then Exit For
                ElseIf mstrFindType = "����" Then '����
                    If .Record(col_����).Value Like "*" & PatiIdentify.Text & "*" Then Exit For
                ElseIf mstrFindType = "�������" Then '�������
                    If tbcPati.Selected.Tag = "��Ժ" Or tbcPati.Selected.Tag = "��Ժ" Then
                        If .Record(col_��ҽ���).Value Like "*" & PatiIdentify.Text & "*" Then Exit For
                    End If
                End If
            End If
        End With
    Next

    If i <= rptPati.Rows.Count - 1 Then
        blnReStart = False
        '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
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
        If mstrFindType = "סԺ��" And Not mblnIsFindAgain Then 'סԺ��
            mblnIsFindAgain = True
            Call LoadPatients
            Call ExecuteFindPati
            mblnIsFindAgain = False
        Else
            blnReStart = True
            MsgBox IIf(blnNext, "������", "") & "�Ҳ������������Ĳ��ˡ�", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Function ExecuteMeetIdea(ByVal intType As Integer) As Boolean
'���ܣ��鿴����д�������
'������intType 0-��д��1���鿴
    Dim blnOk As Boolean
    Dim objFrm As Object
    With rptPati.SelectedRows(0)
        blnOk = mobjKernel.ShowConsultationApply(Me, Val(.Record(col_ҽ��ID).Value), IIf(intType = 0, 3, 4), objFrm)
        Set mFrmConsultation = objFrm
    End With
    If blnOk Then
        Call LoadPatients '��������ˢ��
        If rptPati.Visible Then rptPati.SetFocus
    End If
End Function

Private Function ExecuteMeetFinish() As Boolean
'���ܣ���ɶԵ�ǰ�����˻���
    Dim strSQL As String
    Dim strTmp As String
    Dim str������� As String
    
    Dim lng����ID As Long
    Dim lngҽ��ID As Long
    Dim lng������� As Long
    Dim lngTmp As Long
    Dim i As Long
    Dim lngType As Long '����ʽ 1���ڲ�����2���������崦��
    Dim lng�ж� As Long '����ж� 1���Ȳ����������ٴ���2��ֱ�Ӵ���
    
    Dim blnOk As Boolean
    Dim bln��� As Boolean
    Dim blnTrans As Boolean
    Dim blnDo As Boolean
    
    Dim rsTmp As ADODB.Recordset
    Dim rsҪ�� As ADODB.Recordset
    Dim arrSQL As Variant
    
    If mlng����ID = 0 Then Exit Function
    
    lngҽ��ID = Val(rptPati.SelectedRows(0).Record(col_ҽ��ID).Value)
    
    
    
    If Val(rptPati.SelectedRows(0).Record(COL_�������).Value) <> 0 Then
        '�ж��Ƿ��Ѿ�д�˻��������ҽ������
        strSQL = "select a.���� from ����ҽ������ a where a.ҽ��ID=[1] and a.��Ŀ='�������'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    
        If Not rsTmp.EOF Then
            str������� = rsTmp!���� & ""
            str������� = Trim(str�������)
        End If
    
        'ȡ������ƻ���
        If Val(zlDatabase.GetPara(237, glngSys)) = 1 Then
            If Is�������(lngҽ��ID) Then
                lng�ж� = 1
            Else
                lng�ж� = 2
            End If
        Else
            lng�ж� = 1
        End If
        
        If str������� = "" And lng�ж� = 1 Then
            If MsgBox("δ��д����������Ƿ���д���������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                lngType = 1
            Else
                lngType = 2
            End If
        Else
            lngType = 1
        End If
    
        
        If lngType = 2 Then
            blnOk = mobjKernel.ShowConsultationApply(Me, lngҽ��ID, 3)
            If blnOk Then
                '�ж��ڲ��Ƿ�������
                strSQL = "select 1 from ����ҽ������ a where a.ҽ��ID=[1] and a.ִ��״̬ = 1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
                
                If rsTmp.EOF Then
                
                    If cboDept.ListIndex <> -1 Then lng����ID = cboDept.ItemData(cboDept.ListIndex)
                    
                    With rptPati.SelectedRows(0)
                        strSQL = "ZL_����ҽ��ִ��_Finish(" & .Record(col_ҽ��ID).Value & "," & .Record(col_���ͺ�).Value & ",NULL,0,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lng����ID & ")"
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
            arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Delete(" & lngҽ��ID & ",'�������')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Delete(" & lngҽ��ID & ",'�������ʱ��')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Delete(" & lngҽ��ID & ",'������ɿ���')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Delete(" & lngҽ��ID & ",'����ҽ��')"
            
            strSQL = "select max(a.����) as ��� from ����ҽ������ a where a.ҽ��ID=[1] and nvl(a.��Ŀ,'��') not in ('�������','�������ʱ��','������ɿ���','����ҽ��')"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
            If Not rsTmp.EOF Then
                lng������� = Val(rsTmp!��� & "") + 1
            Else
                lng������� = 1
            End If
            
            '1.���븽��
            strSQL = "select a.id,a.������ from ����������Ŀ a where nvl(a.������,'��') in ('�������','�������ʱ��','������ɿ���','����ҽ��')"
            Set rsҪ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            For i = 1 To rsҪ��.RecordCount
                If rsҪ��!������ & "" = "�������" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'�������',0," & lng������� & "," & rsҪ��!ID & ",'" & str������� & "')"
                    lng������� = lng������� + 1
                ElseIf rsҪ��!������ & "" = "�������ʱ��" Then
                    strTmp = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'�������ʱ��',0," & lng������� & "," & rsҪ��!ID & ",'" & strTmp & "')"
                    lng������� = lng������� + 1
                ElseIf rsҪ��!������ & "" = "������ɿ���" Then
                    lngTmp = Val(rptPati.SelectedRows(0).Record(col_ִ�п���ID).Value)
                    strSQL = "select ���� from ���ű� where id=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngTmp)
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'������ɿ���',0," & lng������� & "," & rsҪ��!ID & ",'" & rsTmp!���� & "')"
                    lng������� = lng������� + 1
                ElseIf rsҪ��!������ & "" = "����ҽ��" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'����ҽ��',0," & lng������� & "," & rsҪ��!ID & ",'" & UserInfo.���� & "')"
                    lng������� = lng������� + 1
                End If
                rsҪ��.MoveNext
            Next
            
            If cboDept.ListIndex <> -1 Then lng����ID = cboDept.ItemData(cboDept.ListIndex)
            With rptPati.SelectedRows(0)
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_����ҽ��ִ��_Finish(" & .Record(col_ҽ��ID).Value & "," & .Record(col_���ͺ�).Value & ",NULL,0,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lng����ID & ")"
            End With
            gcnOracle.BeginTrans: blnTrans = True
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
            Next
            gcnOracle.CommitTrans: blnTrans = False
            blnOk = True
        End If
    Else
        If MsgBox("ȷʵҪ��ɶԸ�""" & lbl����(1).Caption & """�Ļ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        If cboDept.ListIndex <> -1 Then lng����ID = cboDept.ItemData(cboDept.ListIndex)
        
        With rptPati.SelectedRows(0)
            strSQL = "ZL_����ҽ��ִ��_Finish(" & .Record(col_ҽ��ID).Value & "," & .Record(col_���ͺ�).Value & ",NULL,0,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lng����ID & ")"
        End With
        
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
        blnOk = True
    End If
    
    If blnOk Then
        Call LoadPatients '��������ˢ��
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
'���ܣ�ȡ����ɶԵ�ǰ�����˻���
    Dim strSQL As String, lng����ID As Long
    
    If mlng����ID = 0 Then Exit Function
    If MsgBox("ȷʵҪȡ����ɶԸ�""" & lbl����(1).Caption & """�Ļ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    If cboDept.ListIndex <> -1 Then lng����ID = cboDept.ItemData(cboDept.ListIndex)
    
    With rptPati.SelectedRows(0)
        strSQL = "ZL_����ҽ��ִ��_Cancel(" & .Record(col_ҽ��ID).Value & "," & .Record(col_���ͺ�).Value & ",Null,0," & lng����ID & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
    End With
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    On Error GoTo 0
    
    Call LoadPatients '��������ˢ��
    If rptPati.Visible Then rptPati.SetFocus
    ExecuteMeetCancel = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Execute���ܻ���(blnCancel As Boolean) As Boolean
'���ܣ�ȡ����ɶԵ�ǰ�����˻���
    Dim strSQL As String, lng����ID As Long
    
    If mlng����ID = 0 Then Exit Function
    If MsgBox("ȷ��Ҫ" & IIf(blnCancel, "ȡ��", "") & "���ܶԸ�""" & lblPatiName(1).Caption & """�Ļ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    If cboDept.ListIndex <> -1 Then lng����ID = cboDept.ItemData(cboDept.ListIndex)
    
    With rptPati.SelectedRows(0)
        strSQL = "Zl_����ҽ������_���ﴦ��(" & .Record(col_ҽ��ID).Value & "," & .Record(col_���ͺ�).Value & "," & IIf(blnCancel = True, "1", "0") & ",'" & UserInfo.���� & "')"
    End With
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    On Error GoTo 0
    
    Call LoadPatients '��������ˢ��
    If rptPati.Visible Then rptPati.SetFocus
    Execute���ܻ��� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function ExecuteMedRecAuditSubmit() As Boolean
'���ܣ��ύ���˲������
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lng����ID As Long, lng��ҳID As Long, i As Long, lng·��״̬ As Long
    Dim strMsg As String
    
    If mlng����ID = 0 Then Exit Function
    On Error GoTo errH
    With rptPati.SelectedRows(0)
        lng����ID = .Record(col_����Id).Value
        lng��ҳID = .Record(col_��ҳID).Value
    End With

    lng·��״̬ = rptPati.SelectedRows(0).Record(col_·��״̬).Value
    If lng·��״̬ = 1 Then
        strMsg = "�ò��˻�����δ��ɵ��ٴ�·���������ύ������"
        MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    strSQL = "Select �������� From ���Ӳ�����¼ Where ����id = [1] And ��ҳid = [2] And ��ӡʱ�� Is Null Order By ����ʱ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鲡����ӡ", lng����ID, lng��ҳID)
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            If strMsg = "" Then
                strMsg = rsTmp!��������
            Else
                strMsg = strMsg & IIf((i Mod 2) = 0, "," & vbTab, vbCrLf) & rsTmp!��������
            End If
            If Len(strMsg) > 1000 Then
                strMsg = strMsg & "......"
                Exit For
            End If
            rsTmp.MoveNext
        Next
        strMsg = "���²���δ��ӡ��" & vbCrLf & strMsg & vbCrLf & "��ȷ��Ҫ������"
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    
    strSQL = "Zl_�����ύ��¼_Insert(" & lng����ID & "," & lng��ҳID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    On Error GoTo 0
    
    Call LoadPatients '��������ˢ��
    If rptPati.Visible Then rptPati.SetFocus
    ExecuteMedRecAuditSubmit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExecuteMedRecAuditCancel() As Boolean
'���ܣ�ȡ���ύ���˲������
    Dim strSQL As String
    
    If mlng����ID = 0 Then Exit Function
    If MsgBox("ȷʵҪ��""" & lbl����(1).Caption & """�Ĳ���ȡ���ύ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    With rptPati.SelectedRows(0)
        strSQL = "Zl_�����ύ��¼_Delete(" & .Record(col_����Id).Value & "," & .Record(col_��ҳID).Value & ")"
    End With
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    On Error GoTo 0
    
    Call LoadPatients '��������ˢ��
    If rptPati.Visible Then rptPati.SetFocus
    ExecuteMedRecAuditCancel = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ExecuteEditMediRec(Optional ByVal blnEditable As Boolean)
'���ܣ����в�����ҳ����
'������blnEditable=�Ƿ�����༭(��Ȩ�޼�ǩ������������)
    Dim blnReadOnly As Boolean
    
    If mlng����ID = 0 Then Exit Sub
    
    If mPatiInfo.����ת�� Then
        MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If rptPati.SelectedRows(0).GroupRow = False Then
        If rptPati.SelectedRows(0).Record(COL_Ӥ������ID).Value <> 0 Then
            If rptPati.SelectedRows(0).Record(COL_Ӥ������ID).Value = cboDept.ItemData(cboDept.ListIndex) Or rptPati.SelectedRows(0).Record(COL_Ӥ������ID).Value = cboDept.ItemData(cboDept.ListIndex) Then
                MsgBox "�ò����Ѿ�ת���������ˣ�ֻ��Ӥ�����ڱ����ң������������ҳ��", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
    End If
    
    '������Ŀ֮�󲻿�������
    If Not (CheckMecRed(mlng����ID, mlng��ҳID, Me.Caption) Or blnEditable) Then
        blnReadOnly = True
    End If
    
    If mclsInOutMedRec Is Nothing Then
        Set mclsInOutMedRec = New zlMedRecPage.clsInOutMedRec
        Call mclsInOutMedRec.InitMedRec(gcnOracle, glngSys, pסԺҽ��վ, mclsMipModule, gobjCommunity, gclsInsure)
    End If
    '��ģ̬��ʾ��ҳ����
    If Not mclsInOutMedRec.IsOpen Then
        If mclsInOutMedRec.ShowInMedRecEdit(Me, mlng����ID, mPatiInfo.��ҳID, mPatiInfo.����ID, rptPati.SelectedRows(0).Record(col_·��״̬).Value, , mstrPrivs, IIf(blnReadOnly, 1, 0), False) Then
            mstrPrePati = ""
        End If
    End If
End Sub

Private Sub ExecuteCritical()
'���ܣ�Σ��ֵ��ش���
    Dim lngΣ��ֵID As Long '���δ����Σ��ֵ��¼ID
    
    Call mobjKernel.ShowDealCritical(Me, mlng����ID, mlng��ҳID, "", lngΣ��ֵID)
    
    Call SetCriticalAdvice(lngΣ��ֵID)
End Sub

Private Sub timNotify_Timer()
    Static strPreTime1 As String
    Dim curTime As Date
    
    If mblnUnRefresh Then Exit Sub
    
    If mbln��Ϣ���� Then
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
        If mclsMipModule.IsConnect Then 'ʹ������Ϣƽ̨�����Զ�ˢ��
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
'���ܣ���ȡ������鷴��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngCount As Long
    Dim curDate As Date
    
    If cboDept.ListIndex = -1 Then
        fra���.Visible = False: LoadResponse = True: Exit Function
    End If
    
    On Error GoTo errH
    curDate = zlDatabase.Currentdate
    Screen.MousePointer = 11
    
    '��ȡ��ǰ����/��������Ժ����Ժ���ˣ���"����������¼"Ϊ׼ȫ��ɨ��
    strSQL = "Select Count(*) as ���� From ������ҳ B,����������¼ A" & _
        " Where A.����ID=B.����ID and A.��ҳID=B.��ҳID And A.��¼״̬=1 And A.�������� IN(1,2,5,6,7,8,9)" & _
        IIf(mintDeptView = 0, " And B.��Ժ����ID + 0=[1]", " And B.��ǰ����ID + 0=[1]") & _
        IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And B.סԺҽʦ=[2]") & _
        IIf(mblnICU And InStr(mstrPrivs, "ȫԺ����") = 0, " And B.סԺҽʦ=[2]", "") & _
        " And a.����ʱ�� Between [3] And [4]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LoadResponse", cboDept.ItemData(cboDept.ListIndex), UserInfo.����, CDate(Format(curDate - mlngMedRedDay, "yyyy-MM-dd")), CDate(Format(curDate, "yyyy-MM-dd HH:mm:ss")))
    If Not rsTmp.EOF Then lngCount = NVL(rsTmp!����, 0)
    
    lbl���.Caption = mlngMedRedDay & "���ڹ��� " & lngCount & " ��δ����Ĳ�����鷴��..."
    fra���.Visible = lngCount > 0
    If Decode(lngCount, 0, 0, 1) <> Decode(Val(lbl���.Tag), 0, 0, 1) Then
        Call picPati_Resize
    End If
    lbl���.Tag = lngCount
    
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
    
    If Mid(mstrNotify, m��������, 1) = "1" Then strTmp = strTmp & ",ZLHIS_EMR_021"
    If Mid(mstrNotify, mҽ������, 1) = "1" Then strTmp = strTmp & ",ZLHIS_OPER_001,ZLHIS_CIS_005,ZLHIS_CIS_015,ZLHIS_CIS_020"
    If Mid(mstrNotify, mΣ��ֵ, 1) = "1" Then strTmp = strTmp & ",ZLHIS_LIS_003,ZLHIS_PACS_005"
    If Mid(mstrNotify, m���泷��, 1) = "1" Then strTmp = strTmp & ",ZLHIS_LIS_002,ZLHIS_PACS_003"
    If Mid(mstrNotify, mҽ�����, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CIS_026,ZLHIS_CIS_027,ZLHIS_CIS_028,ZLHIS_CIS_029,ZLHIS_CIS_030"
    If Mid(mstrNotify, m�������, 1) = "1" Then strTmp = strTmp & ",ZLHIS_RECIPEAUDIT_002"
    If Mid(mstrNotify, m��Ⱦ��, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CIS_032,ZLHIS_CIS_033"
    If Mid(mstrNotify, m�����ʿ�, 1) = "1" Then strTmp = strTmp & ",ZLHIS_EMR_025"
    If Mid(mstrNotify, m��Ѫ���, 1) = "1" And gblnѪ��ϵͳ Then strTmp = strTmp & ",ZLHIS_BLOOD_001"  '����Ѫ����д���Ϣ�Ͳ���
    If Mid(mstrNotify, mУ������, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CIS_035"
    If Mid(mstrNotify, m��Ѫ���, 1) = "1" And gblnѪ��ϵͳ Then strTmp = strTmp & ",ZLHIS_BLOOD_004"  '����Ѫ����д���Ϣ�Ͳ���
    If Mid(mstrNotify, m��Ѫ��Ӧ, 1) = "1" And gblnѪ��ϵͳ Then strTmp = strTmp & ",ZLHIS_BLOOD_006"
    
    strTmp = Mid(strTmp, 2)
    If strTmp = "" Then LoadNotify = True: Exit Function
    
    strSQL = "Select b.id,b.����id, b.����id as ��ҳid,a.סԺ��,a.����,a.��ǰ���� As ����, Nvl(b.�������id, a.��ǰ����id) As �������id," & _
        " Nvl(b.���ﲡ��id, a.��ǰ����id) As ���ﲡ��id, b.������Դ, b.��Ϣ����, b.���ͱ���, b.ҵ���ʶ, b.���ȳ̶�, b.�Ǽ�ʱ��,a.����" & _
        " From ������Ϣ A, ҵ����Ϣ�嵥 B, ҵ����Ϣ���Ѳ��� C, ҵ����Ϣ������Ա D,������ҳ E" & _
        " Where a.����id = b.����id And b.Id = c.��Ϣid And b.Id = d.��Ϣid(+) And b.����id=e.����id and b.����id=e.��ҳid and e.��ҳid is not null And b.�Ǽ�ʱ�� >=Trunc(Sysdate-" & (mintNotifyDay - 1) & ") and substr(b.���ѳ���,[4],1)='1'" & _
        " And Nvl(b.�Ƿ�����, 0) = 0  And instr(','||[5]||',',','||b.���ͱ���||',')>0 and (c.����id = [1] Or d.������Ա = [3])" & _
        " Order By b.���ȳ̶�, b.�Ǽ�ʱ�� Desc"
        
    If strSQL = "" Then Exit Function
    Screen.MousePointer = 11
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, cboDept.ItemData(cboDept.ListIndex), , UserInfo.����, 2, strTmp)

    For i = 1 To rsTmp.RecordCount
        Select Case rsTmp!���ͱ���
        Case "ZLHIS_CIS_032", "ZLHIS_CIS_033", "ZLHIS_EMR_025"
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
            If InStr("<TB>" & strTag & "<TB>", "<TB>" & rsTmp!����ID & "," & rsTmp!��ҳID & "," & rsTmp!���ͱ��� & "<TB>") = 0 Then
                strTag = strTag & "<TB>" & rsTmp!����ID & "," & rsTmp!��ҳID & "," & rsTmp!���ͱ���
                blnDo = True
            End If
        End Select
        
        If blnDo Then
            Call AddReportRow(rsTmp!����ID & "," & rsTmp!��ҳID, rsTmp!����ID, rsTmp!��ҳID, NVL(rsTmp!����), NVL(rsTmp!סԺ��), NVL(rsTmp!����), NVL(rsTmp!��Ϣ����), _
                rsTmp!���ͱ��� & "", rsTmp!���ȳ̶� & "", Format(rsTmp!�Ǽ�ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), rsTmp!ҵ���ʶ & "", rsTmp!������Դ & "", rsTmp!ID)
                        blnDo = False
        End If
        rsTmp.MoveNext
    Next
    
    '�ϰ没���������
    If Mid(mstrNotify, m��������, 1) = "1" Then
        strSQL = " Select A.����ID,A.��ҳID,A.��������,A.��������,A.ǩ������,A.���ʱ��,B.����" & _
                " From ���Ӳ�����¼ A,�����ļ��б� B" & _
                " Where A.������Դ = 2 And A.�������� In (2,5,6) And Nvl(A.����״̬,0)<=0 And A.�鵵�� Is Null" & _
                " And A.�ļ�ID=B.ID(+) And A.���ʱ��>=Trunc(Sysdate-[2])" & _
                IIf(mintDeptView = 0, " And A.����ID=[1]", " And A.����ID IN(Select ����ID From �������Ҷ�Ӧ Where ����ID=[1])")
        strSQL = "Select A.����ID,A.��ҳID,B.סԺ��,C.����,NVL(D.����,B.����) ����,Min(A.���ʱ��) as ʱ��, -1 As ҽ��״̬,'' as ״̬" & _
            " From (" & strSQL & ") A,������Ϣ B,���˱䶯��¼ C ,������ҳ D" & _
            " Where A.����ID=B.����ID And (A.��������<>2 Or Nvl(A.����,0)>=0)" & _
            " And A.����ID=C.����ID And A.��ҳID=C.��ҳID And A.����ID=D.����ID And A.��ҳID=D.��ҳID" & _
            " And C.��ʼʱ�� Is Not Null And Nvl(C.���Ӵ�λ,0)=0 And (C.��ֹʱ�� Is Null Or C.��ֹԭ��=1)" & _
            " And A.ǩ������<Decode([3],C.����ҽʦ,4,C.����ҽʦ,2,C.����ҽʦ,1,0)" & _
            " Group by A.����ID,A.��ҳID,B.סԺ��,C.����,NVL(D.����,B.����) ,NVL(D.�Ա�,B.�Ա�),NVL(D.����, B.����)  Order by ʱ��"
        Set rsOld = zlDatabase.OpenSQLRecord(strSQL, Me.Name, cboDept.ItemData(cboDept.ListIndex), mintNotifyDay - 1, UserInfo.����)
    
        For i = 1 To rsOld.RecordCount
            Call AddReportRow(rsOld!����ID & "," & rsOld!��ҳID, rsOld!����ID, rsOld!��ҳID, NVL(rsOld!����), NVL(rsOld!סԺ��), NVL(rsOld!����), "����Ҫ��˵Ĳ�����", _
               "", 1, Format(rsOld!ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), "", "", 0)
            rsOld.MoveNext
        Next
    End If
    
    rptNotify.Populate 'ȱʡ��ѡ���κ���
    rptNotify.TabStop = rptNotify.Rows.Count > 0
    Screen.MousePointer = 0
    LoadNotify = True
    If mbln��Ϣ���� Then
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
'���ܣ�������ӡ�¼���д����ҳ��ӡ����
    Dim strSQL As String
    
    strSQL = _
            "Zl_���Ӳ�����ӡ_Insert(Null,9," & mlng����ID & "," & mPatiInfo.��ҳID & ",'" & UserInfo.���� & "')"
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetFontSize(ByVal blnSetMainFont As Boolean)
'���ܣ����н��������ͳһ����
'������blnSetMainFont  �Ƿ��������������� �����������ӽ����л���
    If blnSetMainFont Then
        Call zlControl.SetPubFontSize(Me, mbytSize)
        Call SetPatiIconScale
        Call SetpicPatiPosition
        Call SetPatiInfoCtlPos
    End If
    Select Case tbcSub.Selected.Tag
        Case "סԺһ��"
            Call mfrmInView.SetFontSize(mbytSize)
        Case "·��"
            Call mclsPath.SetFontSize(mbytSize)
        Case "ҽ��"
            Call mclsAdvices.SetFontSize(mbytSize)
        Case "����"
            Call mclsEPRs.SetFontSize(mbytSize)
        Case "����"
            Call mclsTends.SetFontSize(mbytSize)
        Case "�°滤��"
            Call mclsTendsNew.SetFontSize(mbytSize)
        Case "������"
            Call mclsTendEPRs.SetFontSize(mbytSize)
        Case "��������"
            Call mclsDisease.SetFontSize(mbytSize)
        Case "�²���"
            On Error Resume Next
            Call mclsEMR.SetFontSize(mbytSize)
            err.Clear: On Error GoTo 0
    End Select
     
End Sub

Private Sub SetpicPatiPosition()
'���ܣ������б�͹�����������ؿؼ���λ�����С����
    Dim i As Long
    Dim lngDistance As Long
        
    'checkBoxѡ�������������ʱ����ı��ȣ������Ҫ��ȥ100
    lngDistance = IIf(mbytSize = 0, 10, -50)
    Call zlControl.SetPubCtrlPos(False, 0, lbl��������, 10, chk��������(0), lngDistance, chk��������(1), lngDistance, chk��������(2), lngDistance, chkByTeam)
    Call zlControl.SetPubCtrlPos(False, 0, lbl��Ժʱ��, 50, cboSelectTime(1))
    Call zlControl.SetPubCtrlPos(False, 0, lbl��ʼʱ��, 50, cboSelectTime(2), 50, chkOut, 50, chkHZ(0), 50, chkHZ(1))
    Call zlControl.SetPubCtrlPos(False, 0, lblת��, 50, cmdRef)
        
    For i = 0 To picPara.Count - 1
        If i = 0 Then picPara(i).Height = IIf(mbytSize = 0, 320, 280)
        If i = 3 Then picPara(i).Height = IIf(mbytSize = 0, 320, 420)
    Next
    
    txtChange.Left = lblת��.Left + Me.TextWidth("��ʾ��� ")
    fraChange.Left = txtChange.Left
    fraChange.Top = txtChange.Top + txtChange.Height
    chkOutByTeam.Top = chkByTeam.Top
    chkOutByTeam.Left = chkByTeam.Left
    chkFilter.Height = PatiIdentify.Height
End Sub

Private Sub SetPatiInfoCtlPos()
'���ܣ����˵���ϸ��Ϣ����Ŀؼ�λ�õ�����������picInfo�еĿؼ�
    Dim lngDistance1 As Long, lngDistance2 As Long
    Dim lngTmp As Long
    
    lngDistance2 = 180: lngDistance1 = 10
    lngTmp = IIf(mbytSize = 0, 1080, 1300)
    
    lblPatiName(0).Top = IIf(mbytSize = 0, 190, 210)
    lbl����(0).Top = lblPatiName(0).Top
    cboPages.Width = 1600 ' IIf(mbytSize = 0, 1500, 1600)
    
    '1.סԺ����
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
    
    '2.������Ϣ
    lbl����(0).Left = lblPatiName(0).Left
    lbl����.Left = lbl����(0).Left
    lblFee(0).Left = lbl����(0).Left
    
    Call zlControl.SetPubCtrlPos(False, 0, lbl����(0), lngDistance1, lbl����(1), lngDistance2, lbl����(0), lngDistance1, lbl����(1), lngDistance2, lbl����(0), lngDistance1, lbl����(1), _
            lngDistance2, lbl����(0), lngDistance1, lbl����(1), lngDistance2, lbl����(0), lngDistance1, lbl����(1), lngDistance2, lblҽ����(0), lngDistance1, lblҽ����(1))
    
    lbl����.Top = lbl����(0).Height + lbl����(0).Top + 90
    Call zlControl.SetPubCtrlPos(False, 0, lbl����, lngDistance1, cbo����, lngDistance2, lbl����(0), lngDistance1, lbl����(1), lngDistance2, lbl��Ժ(0), lngDistance2, lblDiag(0), lngDistance1, lblDiag(1))
    lbl����(0).Left = lbl����.Left
    
    lblFee(0).Top = lbl����.Height + lbl����.Top + 90
    Call zlControl.SetPubCtrlPos(False, 0, lblFee(0), lngDistance1, lblFee(1), lngDistance2, lblFluid(0), lngDistance1, lblFluid(0))
    
    If lbl����(0).Left <= cbo����.Left + cbo����.Width + lngDistance2 Then
        lbl����(0).Left = cbo����.Left + cbo����.Width + lngDistance2
    End If
    lbl����(0).Left = lbl����(0).Left
    Call zlControl.SetPubCtrlPos(False, 0, lbl����(0), lngDistance1, lbl����(1), lngDistance2, lbl����(0), lngDistance1, lbl����(1), lngDistance2, lbl����(0), lngDistance1, lbl����(1), lngDistance2, lblҽ����(0), lngDistance1, lblҽ����(1))
    Call zlControl.SetPubCtrlPos(False, 0, lbl����(0), lngDistance1, lbl����(1), lngDistance2, lbl��Ժ(0), lngDistance2, lblDiag(0), lngDistance1, lblDiag(1))
    
    If lbl����(0).Left <= lbl��Ժ(0).Left Then
        lbl����(0).Left = lbl��Ժ(0).Left
    Else
        lbl��Ժ(0).Left = lbl����(0).Left
    End If
    
    Call zlControl.SetPubCtrlPos(False, 0, lbl����(0), lngDistance1, lbl����(1), lngDistance2, lbl����(0), lngDistance1, lbl����(1), lngDistance2, lblҽ����(0), lngDistance1, lblҽ����(1))
    Call zlControl.SetPubCtrlPos(False, 0, lbl��Ժ(0), lngDistance1, lbl��Ժ(1), lngDistance2, lblDiag(0), lngDistance1, lblDiag(1))
    
    If lblҽ����(0).Left >= lblDiag(0).Left Then
        If lbl����(0).Left >= lblDiag(0).Left Then
            lblDiag(0).Left = lbl����(0).Left
        Else
            lblDiag(0).Left = lblҽ����(0).Left
        End If
    Else
        lblҽ����(0).Left = lblDiag(0).Left
    End If

    
    Call zlControl.SetPubCtrlPos(False, 0, lbl����(0), lngDistance1, lbl����(1), lngDistance2, lblҽ����(0), lngDistance1, lblҽ����(1))
    Call zlControl.SetPubCtrlPos(False, 0, lblDiag(0), lngDistance1, lblDiag(1))
    
    lblFluid(0).Left = lbl����(0).Left
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
'���ܣ���ȡ���Ҳ����б����ݼ�¼��
'������strIn ��������
    Dim strSQL As String
    Dim blnYN As Boolean
    Dim strDeptIDs As String
    
    If strIn <> "" Then blnYN = True
    If mintDeptView = 0 Then
        '�����Ҷ�ȡ��ʾ
        '�����ż���۲��ҵĲ��˻�û���ϴ�������ֻ�Դ����в��˵Ŀ��ҵ�����
        If InStr(mstrPrivs, "ȫԺ����") > 0 Then
            strSQL = _
                " Select Distinct A.ID,A.����,A.����" & _
                " From ���ű� A,��������˵�� B" & _
                " Where B.����ID=A.ID And B.��������='�ٴ�'" & _
                " And ((B.������� IN(2,3) " & _
                IIf(mintDeptViewBed = 1, " And Exists (Select 1 From ��λ״����¼ C,  �������Ҷ�Ӧ D Where D.����ID = c.����id and A.ID = D.����ID) ", "") & _
                ")Or (B.�������=1 And Exists(Select 1 From ��λ״����¼ C Where B.����ID = C.����ID)))" & _
                " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                IIf(blnYN, " And (A.���� Like [2] Or A.���� Like [3] Or A.���� Like [3])", "") & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " Order by A.����"
        Else
            '����Ȩ�޵Ŀ��ң��������ڿ���+�������������Ŀ���
            strSQL = _
                " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
                " From ���ű� A,��������˵�� B,������Ա C" & _
                " Where B.����ID=A.ID And A.ID=C.����ID And C.��ԱID=[1]" & _
                " And (B.������� IN(2,3) Or (B.�������=1 And Exists(Select 1 From ��λ״����¼ C Where B.����ID = C.����ID)))" & _
                " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                IIf(blnYN, " And (A.���� Like [2] Or A.���� Like [3] Or A.���� Like [3])", "") & _
                " And B.��������='�ٴ�'"
            strSQL = strSQL & " Union " & _
                " Select C.ID,C.����,C.����,Nvl(A.ȱʡ,0) As ȱʡ" & _
                " From ������Ա A,�������Ҷ�Ӧ B,���ű� C" & _
                " Where A.����ID=B.����ID And B.����ID=C.ID And A.��ԱID=[1]" & _
                " And Exists(Select 1 From ��������˵�� Where ��������='����' And ����ID=B.����ID)" & _
                " And Not Exists(Select 1 From ��������˵�� Where ��������='�ٴ�' And ����ID=B.����ID)" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                IIf(blnYN, " And (C.���� Like [2] Or C.���� Like [3] Or C.���� Like [3])", "") & _
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)"
            If InStr(mstrPrivs, "ICU����") > 0 Then
                strSQL = strSQL & " Union " & _
                    " Select A.ID,A.����,A.����,0 As ȱʡ" & _
                    " From ���ű� A" & _
                    " Where Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='ICU')" & _
                    " And Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='�ٴ�')" & _
                    " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    IIf(blnYN, " And (A.���� Like [2] Or A.���� Like [3] Or A.���� Like [3])", "") & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
            End If
            strSQL = "Select ID,����,����,Max(ȱʡ) As ȱʡ From (" & strSQL & ") Group By ID,����,���� Order by ����"
        End If
    Else
        '��������ȡ��ʾ
        If InStr(mstrPrivs, "ȫԺ����") > 0 Then
            strSQL = _
                " Select Distinct A.ID,A.����,A.����" & _
                " From ���ű� A,��������˵�� B " & _
                " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
                IIf(mintDeptViewBed = 1, " And Exists (Select 1 From ��λ״����¼ C Where A.ID = c.����id) ", "") & _
                " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                IIf(blnYN, " And (A.���� Like [2] Or A.���� Like [3] Or A.���� Like [3])", "") & _
                " Order by A.����"
        Else
            '����Ȩ������ֱ�����ڲ���+���ڿ�����������
            strSQL = _
                " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
                " From ���ű� A,��������˵�� B,������Ա C" & _
                " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
                " And B.������� in(1,2,3) And B.��������='����'" & _
                " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                IIf(blnYN, " And (A.���� Like [2] Or A.���� Like [3] Or A.���� Like [3])", "") & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
            strSQL = strSQL & " Union " & _
                " Select C.ID,C.����,C.����,Nvl(A.ȱʡ,0) as ȱʡ" & _
                " From ������Ա A,�������Ҷ�Ӧ B,���ű� C" & _
                " Where A.����ID=B.����ID And B.����ID=C.ID And A.��ԱID=[1]" & _
                " And Exists(Select 1 From ��������˵�� Where ��������='�ٴ�' And ����ID=B.����ID)" & _
                " And Not Exists(Select 1 From ��������˵�� Where ��������='����' And ����ID=B.����ID)" & _
                " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                IIf(blnYN, " And (C.���� Like [2] Or C.���� Like [3] Or C.���� Like [3])", "") & _
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)"
            If InStr(mstrPrivs, "ICU����") > 0 Then
                strSQL = strSQL & " Union " & _
                    " Select A.ID,A.����,A.����,0 As ȱʡ" & _
                    " From ���ű� A" & _
                    " Where Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='ICU')" & _
                    " And Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='����')" & _
                    " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    IIf(blnYN, " And (A.���� Like [2] Or A.���� Like [3] Or A.���� Like [3])", "") & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
            End If
            strSQL = "Select ID,����,����,Max(ȱʡ) as ȱʡ From (" & strSQL & ") Group by ID,����,���� Order by ����"
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
'���ܣ������յ�����Ϣ���������б���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim blnTmp As Boolean
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim i As Long
    
    On Error GoTo errH
    
    If Mid(rsMsg!���ѳ���, 2, 1) <> "1" Then Exit Sub
    
    strSQL = "select ����id as id,�������� as ���� from ��������˵�� where ����id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(rsMsg!����IDs & ""))
    rsTmp.Filter = "����=" & IIf(mintDeptView = 0, "'�ٴ�'", "'����'")
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If Val(rsTmp!ID & "") = cboDept.ItemData(cboDept.ListIndex) Then
                blnTmp = True: Exit For
            End If
            rsTmp.MoveNext
        Next
    End If
    
    If blnTmp Or InStr("," & rsMsg!������Ա & ",", "," & UserInfo.���� & ",") > 0 Then
        
        '�ж��б��Ƿ��Ѿ���������Ϣ��
        For i = 0 To rptNotify.Rows.Count - 1
            If Not rptNotify.Rows(i).GroupRow Then
                If rptNotify.Rows(i).Record(C_��Ϣ).Value = rsMsg!���ͱ��� And rptNotify.Rows(i).Record.Tag = CStr(rsMsg!����ID & "," & rsMsg!����id) Then
                    Exit Sub
                End If
            End If
        Next
        
        strSQL = "Select a.סԺ��, a.����, a.�Ա�, a.����, a.��ǰ���� As ����, a.���� From ������Ϣ A Where a.����id =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMsg!����ID))
        
        Call AddReportRow(rsMsg!����ID & "," & rsMsg!����id, rsMsg!����ID, rsMsg!����id, rsTmp!����, NVL(rsTmp!סԺ��), NVL(rsTmp!����), NVL(rsMsg!��Ϣ����), _
             rsMsg!���ͱ��� & "", rsMsg!���ȳ̶� & "", Format(rsMsg!�Ǽ�ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), rsMsg!ҵ���ʶ & "", rsMsg!������Դ & "", 0)
        
        rptNotify.Populate
         
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AddReportRow(ParamArray arrInput() As Variant)
'���ܣ�����Ϣ�����б�������һ��
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strNO As String
    Dim strҵ�� As String
    Dim str������Դ As String
    Dim int���ȼ� As Integer
    Dim Index As Integer
    
    On Error GoTo errH
    Set objRecord = Me.rptNotify.Records.Add()
    objRecord.Tag = arrInput(Index): Index = Index + 1         'Tagֵ
    Set objItem = objRecord.AddItem(""): objItem.Icon = 3
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '����id
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '����id
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1 '����
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))      'סԺ��
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))      '����
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))      '״̬������
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    strNO = arrInput(Index)                                     '��Ϣ���
    objRecord.AddItem strNO: Index = Index + 1
    
    int���ȼ� = Val(arrInput(Index))                            '���
    objRecord.AddItem int���ȼ�: Index = Index + 1
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  '����
    
    strҵ�� = arrInput(Index): Index = Index + 1              'ҵ���ʶ
    str������Դ = arrInput(Index)                             '������Դ
    objRecord.AddItem strҵ��
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
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ReadMsg(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strNO As String, ByVal strҵ�� As String, ByVal lng��ϢID As Long) As Boolean
'���ܣ��Ķ���Ϣ
'˵������Ϣ�Ķ���ʽĿǰ��3�֣�����Ϣ�������Ķ�����ϢID�Ķ�����ҵ���ʶ�Ķ�
    Dim strSQL As String
    Dim strҽ��ID As String
    Dim blnDo As Boolean
    Dim lngΣ��ֵID As Long  '���δ����Σ��ֵ��¼ID
    Dim blnHisΣ��ֵ As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQLReadMsg As String
    Dim objControl As Object
    Dim i As Long
    On Error GoTo errH
    blnDo = True
    
    strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng��ҳID & ",'" & strNO & "',2,'" & UserInfo.���� & "'," & cboDept.ItemData(cboDept.ListIndex)
    Select Case strNO
    Case "ZLHIS_LIS_003", "ZLHIS_PACS_005"
        strSQL = strSQL & ",null,null,'" & strҵ�� & "'"
    Case "ZLHIS_CIS_032", "ZLHIS_CIS_033"
        strSQL = strSQL & ",null," & lng��ϢID
    End Select
    strSQL = strSQL & ")"
    
    strSQLReadMsg = strSQL
     
    If strNO = "ZLHIS_CIS_035" Or strNO = "ZLHIS_BLOOD_004" Then
        If strNO = "ZLHIS_CIS_035" Then
            'У��������Ϣ�Ĵ�����ҽ���༭���棬ҽ��������Զ�������Ϣ
            strSQL = "select a.ID,a.���ID,a.������� from ����ҽ����¼ a where A.ҽ��״̬=2 and a.����id=[1] and a.��ҳid=[2] order by a.���"
        Else
            strSQL = "select 1 from ����ҽ����¼ a where a.����id=[1] and a.��ҳid=[2] and a.ҽ��״̬=1 and a.�������='K' and a.��鷽��='1' and a.���״̬=1 and rownum<2"
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If rsTmp.EOF Then '����������Ϣ����Ϊ����
             Call zlDatabase.ExecuteProcedure(strSQLReadMsg, Me.Caption)
             ReadMsg = True
             Exit Function
        End If
        
        If strNO = "ZLHIS_CIS_035" Then
            '��λһ����Чҽ����
            For i = 1 To rsTmp.RecordCount
                If InStr(",5,6,", rsTmp!������� & "") > 0 Then
                    Call mclsAdvices.LocatedAdviceRow(Val(rsTmp!ID & ""))
                ElseIf "7" = rsTmp!������� & "" Then
                    Call mclsAdvices.LocatedAdviceRow(Val(rsTmp!���ID & ""))
                Else
                    Call mclsAdvices.LocatedAdviceRow(Val(rsTmp!ID & ""))
                End If
                Exit For
                rsTmp.MoveNext
            Next
        End If
  
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
        
        'ҽ���ж�λʧ���޸Ĳ˵����ܲ�����
        Set objControl = cbsMain.FindControl(, conMenu_Edit_Modify, True, True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then
                objControl.Execute
            End If
        End If
        Exit Function
    End If
    
    
    If strNO = "ZLHIS_LIS_003" Or strNO = "ZLHIS_PACS_005" Then
        If mblnΣ��ֵ Then
            'Σ��ֵ��Ϣ��ش���
            Call mobjKernel.ShowDealCritical(Me, lng����ID, lng��ҳID, "", lngΣ��ֵID)
            
            If lngΣ��ֵID <> 0 Then
                strSQL = "select a.�걾id,a.�������,a.ȷ���� from ����Σ��ֵ��¼ a where a.id=[1] and a.ȷ���� is not null"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngΣ��ֵID)
                If Not rsTmp.EOF Then
                    '����Ϣ����Ϊ����
                    Call zlDatabase.ExecuteProcedure(strSQLReadMsg, Me.Caption)
                    '�����LISΣ��ֵ����LIS�ӿ�
                    If strNO = "ZLHIS_LIS_003" Then
                        Call InitObjLis(pסԺҽ��վ)
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
                Call InitObjLis(pסԺҽ��վ)
                If Not gobjLIS Is Nothing Then
                    blnDo = gobjLIS.GetReadNotify(Me, strҽ��ID, UserInfo.����)
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

Private Function Is�������(ByVal lngҽ��ID As Long) As Boolean
'���ܣ���д�������ʱ�жϵ�ǰ�Ŀ����ǲ��Ǵ������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    strSQL = "select 1 from ����ҽ����¼ a,���ű� b where a.ִ�п���id=b.id and a.id=[1] and" & vbNewLine & _
        "exists (select 1 from ����ҽ������ c where c.ҽ��id =[1] and c.��Ŀ='����������'and b.����=c.����)"
        
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    
    If Not rsTmp.EOF Then
        Is������� = True
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

Private Sub ReadMsg����(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strNO As String)
'���ܣ���Ϣ�������е�һ����Ϣ
    Dim objRow As ReportRow
    Dim strTmp As String
    Dim strIndexs As String
    Dim arrTmp As Variant
    Dim i As Integer
    
    For Each objRow In rptNotify.Rows
        If objRow.GroupRow Then objRow.Expanded = True
        If Not objRow.GroupRow And objRow.Childs.Count = 0 Then
            If Val(objRow.Record(C_����Id).Value) = lng����ID And Val(objRow.Record(C_��ҳId).Value) = lng��ҳID And objRow.Record(C_��Ϣ).Value = strNO Then
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
'�޸�ͼ��
    If Button = 2 Then
        Call Initͼ��˵�(1, Index + 1)
    End If
End Sub

Private Sub fraPageId_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'�հ״�����ͼ�����ò˵�
    If Button = 2 Then
        Call Initͼ��˵�(0)
    End If
End Sub

Private Sub lblCountThis_Click(Index As Integer)
'����ѡ����
    Dim strTmp As String
    Dim strIcon As String
    
    strTmp = lblCountThis(Index).Tag
    If strTmp = "" Then Exit Sub
    mrsPati����.Filter = "˵��='" & Split(strTmp, "<Tab>")(0) & "' and ͼ������=" & Split(strTmp, "<Tab>")(1)
    If Not mrsPati����.EOF Then
        strTmp = mrsPati����!����
        If strTmp <> "" Then
            Call LoadIconSelect(strTmp)
        End If
    End If
End Sub

Private Sub imgIconPati_Click(Index As Integer)
'����ѡ����
    Call lblCountThis_Click(Index)
End Sub

Private Sub ShowPatiͼ��(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
'���ܣ���ʾ��ǰ���˵�ͼ��
'˵����ÿ������Ҫ�����5��ͼ�꣬��ѯ������ֻ�����ÿɼ��Լ���
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, j As Long
    Dim strTmp As String
    Dim lngIconIdx As Long
    
    On Error GoTo errH
    
    '����ͼ��
    For i = 0 To imgCurPati.Count - 1
        imgCurPati(i).Visible = False
        imgCurPati(i).Tag = ""
    Next
    Set mPatiInfo.rsͼ�� = Nothing
    
    '�ȶ�����
    If Not mrsPatiNotes Is Nothing Then
        mrsPatiNotes.Filter = "����id=" & lng����ID & " and ��ҳid=" & lng��ҳID
        If Not mrsPatiNotes.EOF Then
            Set rsTmp = zlDatabase.CopyNewRec(mrsPatiNotes)
        End If
    End If
    
    If rsTmp Is Nothing Then
        strSQL = "Select a.����id, a.��ҳid, a.���˳��,a.�������, nvl(a.���ⲡ��ID,0) as ���ⲡ��ID, a.������, a.����, Replace(b.˵��, '|', '') as ˵��, b.ͼ������, b.��Ч����, Floor(Sysdate - a.����) As ʵ������" & _
            " From ������Ǽ�¼ A, ����������� B,������ҳ c Where a.������� = b.������� And a.������ = b.������  And nvl(a.���ⲡ��ID,0) = nvl(b.����id,0) And (b.��Ч���� = 0 Or (b.��Ч���� > Floor(Sysdate - a.����))) " & _
            " and a.����id=c.����id and a.��ҳid=c.��ҳid and a.����id=c.��ǰ����id And a.����id = [1] And a.��ҳid = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
    End If
    
    If Not rsTmp.EOF Then
        Set mPatiInfo.rsͼ�� = zlDatabase.CopyNewRec(rsTmp)
        For i = 1 To rsTmp.RecordCount
            If InStr(",1,2,3,", rsTmp!���˳��) > 0 Then
                j = Val(rsTmp!���˳�� & "") - 1
                imgCurPati(j).Visible = True
                lngIconIdx = Val(rsTmp!ͼ������ & "") + 1
                If lngIconIdx > 0 And lngIconIdx <= zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages.Count Then
                    Set imgCurPati(j).Picture = zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages(lngIconIdx).Picture
                End If
                imgCurPati(j).ToolTipText = rsTmp!˵�� & ""
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

Private Sub SetAllPatiͼ��(Optional ByVal intFun As Integer)
'���ܣ�����ͼ��ĳ�ʼ����ж��
'������0-���أ�1-ж��
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
            mstrList���� = ""
        End If
    Next
    
End Sub

Private Sub ShowAllPatiͼ��(ByVal strPatis As String)
'���ܣ���ʾ��ǰ���в��˵Ļ���ͼ��,���ص�ǰ�����б��˵�ͼ��
'������strPatis ��ʽ��"����ID:��ҳID,����ID:��ҳID,..."
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strTmp As String
    Dim strParTable  As String
    Dim strTable As String
    Dim varArr As Variant
    Dim strPar As String
    Dim lngCnt As Long
    Dim str���� As String
    Dim lngIconIdx As Long
    
    
    On Error GoTo errH
    
    If strPatis = "" Then
        '����ͼ��
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
    
    strSQL = "Select a.����id, a.��ҳid, a.�������,a.���˳��, nvl(a.���ⲡ��ID,0) as ���ⲡ��ID,a.������, a.����, Replace(b.˵��, '|', '') as ˵��, b.ͼ������, b.��Ч����, Floor(Sysdate - a.����) As ʵ������" & _
        " From ������Ǽ�¼ A, ����������� B,������ҳ c Where a.������� = b.������� And a.������ = b.������  And nvl(a.���ⲡ��ID,0) = nvl(b.����id,0) And (b.��Ч���� = 0 Or (b.��Ч���� > Floor(Sysdate - a.����))) " & _
        " and a.����id=c.����id and a.��ҳid=c.��ҳid and a.����id=c.��ǰ����id and (a.����id,a.��ҳid) In (" & strTable & ")"
                
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
        
    '��ʼ������ͼ���¼��
    Set mrsPati���� = New ADODB.Recordset

    mrsPati����.Fields.Append "˵��", adVarChar, 40000
    mrsPati����.Fields.Append "ͼ������", adBigInt
    mrsPati����.Fields.Append "����", adBigInt
    mrsPati����.Fields.Append "����", adVarChar, 40000
    mrsPati����.CursorLocation = adUseClient
    mrsPati����.LockType = adLockOptimistic
    mrsPati����.CursorType = adOpenStatic
    mrsPati����.Open
    
    For i = 1 To mrsPatiNotes.RecordCount
         
        mrsPati����.Filter = "˵��='" & mrsPatiNotes!˵�� & "' and ͼ������=" & mrsPatiNotes!ͼ������
        
        If Not mrsPati����.EOF Then
            If InStr(str����, "," & mrsPatiNotes!����ID & ":" & mrsPatiNotes!��ҳID & ",") = 0 Then
                mrsPati����!���� = mrsPati����!���� & "," & mrsPatiNotes!����ID & ":" & mrsPatiNotes!��ҳID
                mrsPati����!���� = Val(mrsPati����!���� & "") + 1
                str���� = str���� & mrsPatiNotes!����ID & ":" & mrsPatiNotes!��ҳID & ","
            End If
        Else
            mrsPati����.AddNew
            mrsPati����!˵�� = mrsPatiNotes!˵��
            mrsPati����!ͼ������ = mrsPatiNotes!ͼ������
            mrsPati����!���� = 1
            mrsPati����!���� = mrsPatiNotes!����ID & ":" & mrsPatiNotes!��ҳID
            str���� = "," & mrsPati����!���� & ","
        End If
        mrsPati����.Update
        
        mrsPatiNotes.MoveNext
    Next
    mrsPati����.Filter = 0
    mrsPati����.Sort = "���� desc,ͼ������"
    
    '����ͼ�����
    For i = 0 To mrsPati����.RecordCount - 1
        lngIconIdx = Val(mrsPati����!ͼ������ & "") + 1
        If lngIconIdx > 0 And lngIconIdx <= zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages.Count Then
            Set imgIconPati(i).Picture = zlCommFun.GetPaitSignImageList(IIf(mlngSource = 999, 0, 1)).ListImages(lngIconIdx).Picture
        End If
        lblCountThis(i).Caption = "(" & mrsPati����!���� & ")"
        imgIconPati(i).Visible = True
        lblCountThis(i).Visible = True
        imgIconPati(i).ToolTipText = mrsPati����!˵��
        lblCountThis(i).ToolTipText = mrsPati����!˵��
        imgIconPati(i).Tag = lngIconIdx
        lblCountThis(i).Tag = mrsPati����!˵�� & "<Tab>" & mrsPati����!ͼ������
        lngCnt = lngCnt + 1
        mrsPati����.MoveNext
    Next
    
    For i = lngCnt To conIconAll - 1
        imgIconPati(i).Visible = False
        lblCountThis(i).Visible = False
    Next
    
    '����λ��
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

Private Sub Initͼ����Ϣ(ByVal lng����ID As Long)
'���ܣ���ʼ�����˸��Ի�ͼ����
'��ȡ��ǰ�����趨�ı�ע����
    Dim strSQL As String
    Dim i As Long
    
    On Error GoTo errH
    mstrList���� = ""
    strSQL = "Select Nvl(a.����id,0) as ����id, a.�������, a.������, Replace(a.˵��, '|', '') ˵��, a.ͼ������, a.��Ч����,a.�Ƿ�����" & vbNewLine & _
        "From ����������� a Where a.����id Is Null Or a.����id =[1]" & vbNewLine & _
        "Order By Nvl(a.����id, 0), a.�������, a.������"

    Set mrsNotes = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    For i = 1 To mrsNotes.RecordCount
        If Val(mrsNotes!������ & "") = 0 Then
            mstrList���� = mstrList���� & "<TabB>" & mrsNotes!����ID & "<TabA>" & mrsNotes!������� & "<TabA>" & mrsNotes!˵��
        End If
        mrsNotes.MoveNext
    Next
    mstrList���� = Mid(mstrList����, 7)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Initͼ��˵�(ByVal intType As Integer, Optional ByVal lng˳��� As Long)
'���ܣ�����ͼ�����ò˵�
'������intType 0-��հ״���1-�޸�ĳһ��
'      lng˳��� �޸ĵĵڼ���
'��ʾ�����б�ע���Ⲣ�ṩѡ��ע��������
    Dim int����1 As Integer
    Dim int����2 As Integer
    Dim int����3 As Integer
    
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrPopup As CommandBarPopup
    Dim var���� As Variant
    Dim varTmp As Variant
    Dim str���б��¼ As String
    Dim strBarName As String
    Dim i As Long
    Dim rsPatiͼ�� As ADODB.Recordset
    Dim intFun As Integer '1-ȫ��ʾ��2-����
    Dim int���� As Integer
    Dim blnDo As Boolean
    Dim blnDel As Boolean
    Dim strIdxChk As String
    Dim strDelInfo As String
    Dim lngNew��� As Long
    
    If mPatiInfo.����ID = 0 Then Exit Sub
    If mintDeptView = 0 Then
        Call Initͼ����Ϣ(mPatiInfo.����ID)
    End If
    If mrsNotes Is Nothing Then Exit Sub
    If mrsNotes.RecordCount = 0 Then Exit Sub
    
    'ֻ���������б���ͬ��һ��סԺ
    If mPatiInfo.����ID & "," & mPatiInfo.��ҳID <> mlng����ID & "," & mlng��ҳID Then Exit Sub
    
    If Not mPatiInfo.rsͼ�� Is Nothing Then
        Set rsPatiͼ�� = mPatiInfo.rsͼ��
    End If
    
    If intType = 0 Then
        If rsPatiͼ�� Is Nothing Then
            'ȫ��ʾ
            intFun = 1
            lngNew��� = 1
        Else
            If rsPatiͼ��.RecordCount < 3 Then
                intFun = 2
                lngNew��� = 1
                rsPatiͼ��.Sort = "���˳��"
                For i = 1 To rsPatiͼ��.RecordCount
                    If Val(rsPatiͼ��!���˳�� & "") = lngNew��� Then
                        lngNew��� = lngNew��� + 1
                    End If
                    rsPatiͼ��.MoveNext
                Next
            Else
                Exit Sub
            End If
        End If
    ElseIf intType = 1 Then
        intFun = 2
    End If
    
    Set cbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
    cbrPopupBar.Title = "��ע�趨"
    If mlngSource = 999 Then
        Call cbrPopupBar.SetIconSize(16, 16)
    Else
        Call cbrPopupBar.SetIconSize(24, 24)
    End If
    
    var���� = Split(mstrList����, "<TabB>")
    For i = 0 To UBound(var����)
        varTmp = Split(var����(i), "<TabA>")
        strBarName = varTmp(2)
        If intFun = 1 Or intFun = 2 Then
            blnDo = True
            If intFun = 2 Then
                rsPatiͼ��.Filter = "���ⲡ��ID=" & varTmp(0) & "  and �������=" & varTmp(1)
                blnDo = rsPatiͼ��.EOF
                If intType = 1 And Not blnDo Then
                    If lng˳��� = Val(rsPatiͼ��!���˳�� & "") Then
                        strIdxChk = varTmp(0) & "," & varTmp(1) & "," & rsPatiͼ��!ͼ������ & "," & rsPatiͼ��!˵��
                        
                        strDelInfo = varTmp(1) & ",0,0," & varTmp(0)
                        blnDel = True
                        blnDo = True
                    End If
                End If
            End If
            
            If blnDo Then
                mrsNotes.Filter = "����ID=" & varTmp(0) & "  and �������=" & varTmp(1) & " And ������>0"
                If mrsNotes.RecordCount <> 0 Then
                    int���� = int���� + 1
                    Set cbrPopup = cbrPopupBar.Controls.Add(xtpControlButtonPopup, conMenu_��ע1, strBarName)
                    If mlngSource = 999 Then
                        Call cbrPopup.CommandBar.SetIconSize(16, 16)
                    Else
                        Call cbrPopup.CommandBar.SetIconSize(24, 24)
                    End If
                    Do While Not mrsNotes.EOF
                        Set cbrPopupItem = cbrPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_��ע1 + mrsNotes.RecordCount, mrsNotes!˵��)
                        cbrPopupItem.IconId = conMenu_ͼ�� + mrsNotes!ͼ������
                        
                        
                        If intType = 1 Then
                            cbrPopupItem.Parameter = "1," & mrsNotes!������� & "," & mrsNotes!������ & "," & lng˳��� & "," & mrsNotes!����ID
                            cbrPopupItem.Checked = (strIdxChk = mrsNotes!����ID & "," & mrsNotes!������� & "," & mrsNotes!ͼ������ & "," & mrsNotes!˵��)
                        Else
                            cbrPopupItem.Parameter = "0," & mrsNotes!������� & "," & mrsNotes!������ & "," & lngNew��� & "," & mrsNotes!����ID
                        End If
                        
                        
                        mrsNotes.MoveNext
                    Loop
                    If blnDel Then
                        Set cbrPopupItem = cbrPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_��ע1 + mrsNotes.RecordCount + 1, "�����ע")
                            cbrPopupItem.BeginGroup = True
                            cbrPopupItem.IconId = 3014
                        cbrPopupItem.Parameter = "2," & strDelInfo    '����Ϊ-1,����Ϊimage��������1��ʼ,��ImageManager�Ǵ�0��ʼ
                    End If
                    blnDel = False
                End If
            End If
        End If
    Next
    mrsNotes.Filter = 0
    
    If int���� <> 0 Then
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
'���ܣ�����ͼ�꣬��������С����Ķ�Ӧ
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
    '��λ����
    Call LocatePati(rptTBPati.FocusedRow.Record.Tag)
End Sub

Private Sub LoadIconSelect(ByVal strPatis As String)
'���ܣ����ص��ͼ����ѡ�����б�
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
    
    strSQL = "select a.����id,a.��ҳid,a.����,lpad(a.��Ժ����,10,' ') AS ����,a.סԺ��," & vbNewLine & _
        " Decode(a.���ʱ��,NULL,a.��Ժ����,a.���ʱ��) AS ��Ժ���� ,a.��Ժ����,a.�������� from ������ҳ a" & vbNewLine & _
        " where (a.����id,a.��ҳid) In (" & strTable & ")" & _
        " order by ����"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatis)
    
    rptTBPati.Records.DeleteAll
 
    With rsTmp
        Do While Not .EOF
            Set objRecord = Me.rptTBPati.Records.Add()
            objRecord.Tag = CStr(!����ID & "," & !��ҳID)
            
            intIcon1 = -1
            intIcon2 = -1
            intIcon3 = -1
            
            If Not mrsPatiNotes Is Nothing Then
                mrsPatiNotes.Filter = "����ID=" & !����ID & " and ��ҳID=" & !��ҳID
                If Not mrsPatiNotes.EOF Then
                    For i = 1 To mrsPatiNotes.RecordCount
                        If mrsPatiNotes!���˳�� = 1 Then
                            intIcon1 = mrsPatiNotes!ͼ������
                        ElseIf mrsPatiNotes!���˳�� = 2 Then
                            intIcon2 = mrsPatiNotes!ͼ������
                        ElseIf mrsPatiNotes!���˳�� = 3 Then
                            intIcon3 = mrsPatiNotes!ͼ������
                        End If
                        mrsPatiNotes.MoveNext
                    Next
                End If
            End If
            
            'ͼ��1
            Set objItem = objRecord.AddItem("")
            If intIcon1 > -1 Then
                objItem.Icon = intIcon1
            End If
  
            'ͼ��2
            Set objItem = objRecord.AddItem("")
            If intIcon2 > -1 Then
                objItem.Icon = intIcon2
            End If
            
            'ͼ��3
            Set objItem = objRecord.AddItem("")
            If intIcon3 > -1 Then
                objItem.Icon = intIcon3
            End If
            
            Set objItem = objRecord.AddItem(zlStr.Lpad(NVL(!����), 10))
            objItem.Caption = Trim(NVL(!����, " "))
            objRecord.AddItem Val(!����ID)
            objRecord.AddItem Val(!��ҳID)
            objRecord.AddItem CStr(NVL(!����))
            Set objItem = objRecord.AddItem(CStr(NVL(!סԺ��)))
            objItem.Caption = NVL(!סԺ��, " ")
            
            Set objItem = objRecord.AddItem(Format(!��Ժ����, "yyyy-MM-dd"))
            objItem.Caption = Format(!��Ժ����, "yyyy-MM-dd")
            Set objItem = objRecord.AddItem(Format(!��Ժ����, "yyyy-MM-dd"))
            objItem.Caption = Format(!��Ժ����, "yyyy-MM-dd")
            
            Set objItem = objRecord.AddItem(NVL(!��������))
            objItem.Caption = NVL(!��������)
            
            .MoveNext
        Loop
    End With

    'picTBPati ��������
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    rptTBPati.Populate 'ȱʡ��ѡ���κ���
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
'���ܣ����²���ͼ��
'������ strIconInfo ���̲�����Ϣ
    '��������
    Dim strSQL As String
    Dim strSQLOther As String
    Dim varTmp As Variant
    Dim intType As Integer '0-����1-�ģ�2-ɾ
    Dim lng˳��� As Long
    Dim lng������� As Long
    Dim lng���ⲡ�� As Long
    Dim lng������ As Long
    Dim blnTrans As Boolean
    
    On Error GoTo errH
       
    varTmp = Split(strIconInfo, ",")
    intType = varTmp(0)
    lng������� = varTmp(1)
    lng������ = varTmp(2)
    lng˳��� = varTmp(3)
    lng���ⲡ�� = varTmp(4)
    
    If intType = 1 Then
        mPatiInfo.rsͼ��.Filter = "���˳��=" & lng˳���
        If mPatiInfo.rsͼ��!���ⲡ��ID & "," & mPatiInfo.rsͼ��!������� <> lng���ⲡ�� & "," & lng������� Then
            strSQLOther = "ZL_������Ǽ�¼_UPDATE(" & mPatiInfo.����ID & "," & mPatiInfo.����ID & "," & mPatiInfo.��ҳID & "," & mPatiInfo.rsͼ��!������� & ",0,0," & mPatiInfo.rsͼ��!���ⲡ��ID & ")"
        End If
        strSQL = "ZL_������Ǽ�¼_UPDATE(" & mPatiInfo.����ID & "," & mPatiInfo.����ID & "," & mPatiInfo.��ҳID & "," & lng������� & "," & lng������ & "," & lng˳��� & "," & IIf(0 = lng���ⲡ��, "null", lng���ⲡ��) & ")"
        
        gcnOracle.BeginTrans: blnTrans = True
        If strSQLOther <> "" Then Call zlDatabase.ExecuteProcedure(strSQLOther, Me.Caption)
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        gcnOracle.CommitTrans: blnTrans = False
    Else
        strSQL = "ZL_������Ǽ�¼_UPDATE(" & mPatiInfo.����ID & "," & mPatiInfo.����ID & "," & mPatiInfo.��ҳID & "," & lng������� & "," & lng������ & "," & lng˳��� & "," & IIf(0 = lng���ⲡ��, "null", lng���ⲡ��) & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    
    '���¿�Ƭ
    Set mrsPatiNotes = Nothing
    Call ShowAllPatiͼ��(mstrAllPatis)
    Call ShowPatiͼ��(mPatiInfo.����ID, mPatiInfo.��ҳID)
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

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

Private Sub ReadMsgAuto()
'���ܣ�Σ��ֵ��Ϣ�����Զ�����
    Dim i As Long
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    Dim strNO As String
    Dim strҵ�� As String
    Dim lng��ϢID As Long
    Dim blnRs As Boolean
    
    On Error GoTo errH
    
    For i = i To rptNotify.Rows.Count - 1
        With rptNotify.Rows(i)
            If Not .GroupRow Then
                strNO = .Record(C_��Ϣ).Value
                If strNO = "ZLHIS_LIS_003" Or strNO = "ZLHIS_PACS_005" Then
                    lng����ID = Val(.Record(C_����Id).Value)
                    lng��ҳID = Val(.Record(C_��ҳId).Value)
                    strҵ�� = .Record(C_ҵ��).Value
                    lng��ϢID = Val(.Record(C_Id).Value)
                    blnRs = ReadMsg(lng����ID, lng��ҳID, strNO, strҵ��, lng��ϢID)
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
