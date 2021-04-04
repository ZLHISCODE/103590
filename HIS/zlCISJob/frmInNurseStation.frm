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
   Caption         =   "סԺ��ʿ����վ"
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
            Caption         =   "סԺ����"
            Height          =   180
            Left            =   60
            TabIndex        =   70
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblPatiName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
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
         Begin VB.ComboBox cbo���� 
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
            Caption         =   "����:"
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
            Caption         =   "���:"
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
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҩ��:"
            Height          =   180
            Left            =   75
            TabIndex        =   63
            Top             =   450
            Width           =   810
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
            TabIndex        =   62
            Top             =   165
            Width           =   105
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ:"
            Height          =   180
            Index           =   0
            Left            =   3465
            TabIndex        =   61
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            Height          =   180
            Index           =   0
            Left            =   4140
            TabIndex        =   60
            Top             =   180
            Width           =   450
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
            TabIndex        =   59
            Top             =   165
            Width           =   90
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ��:"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   58
            Top             =   150
            Width           =   630
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            Height          =   180
            Index           =   0
            Left            =   7060
            TabIndex        =   57
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lblҽ���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ����:"
            Height          =   180
            Index           =   0
            Left            =   8595
            TabIndex        =   56
            Top             =   165
            Width           =   630
         End
         Begin VB.Label lbl��Ժ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ:"
            Height          =   180
            Index           =   0
            Left            =   5115
            TabIndex        =   55
            Top             =   450
            Width           =   450
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            Height          =   180
            Index           =   0
            Left            =   4065
            TabIndex        =   54
            Top             =   450
            Width           =   450
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            Height          =   180
            Index           =   0
            Left            =   2055
            TabIndex        =   53
            Top             =   165
            Width           =   450
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
            TabIndex        =   52
            Top             =   165
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
            TabIndex        =   51
            Top             =   450
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
            TabIndex        =   50
            Top             =   450
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
            TabIndex        =   49
            Top             =   165
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
            Left            =   7500
            TabIndex        =   48
            Top             =   165
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
            TabIndex        =   47
            Top             =   165
            Width           =   105
         End
         Begin VB.Label lblFluid 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Һ��:"
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
            Caption         =   "ҽ��:"
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
         Caption         =   "ȫ����"
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
         Caption         =   "���˸���"
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
         Caption         =   "���ѷ�Χ:"
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
            Text            =   "�������"
            TextSave        =   "�������"
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
      Begin VB.Frame fra��� 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   60
         TabIndex        =   9
         Top             =   4755
         Visible         =   0   'False
         Width           =   3360
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
               Caption         =   "ˢ��"
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
            Begin VB.Label lblת�� 
               AutoSize        =   -1  'True
               Caption         =   "��ʾ���    ���ת������"
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
            Begin VB.CommandButton cmd�������� 
               Height          =   240
               Left            =   2925
               Picture         =   "frmInNurseStation.frx":14F8
               Style           =   1  'Graphical
               TabIndex        =   27
               TabStop         =   0   'False
               ToolTipText     =   "ѡ����Ŀ(F4)"
               Top             =   30
               Width           =   270
            End
            Begin VB.TextBox txt�������� 
               Height          =   300
               Left            =   1060
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   0
               Width           =   2160
            End
            Begin VB.Label lbl�������� 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "����ȼ�(&N)"
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
            Begin VB.CheckBox chk�������� 
               Caption         =   "��"
               Height          =   195
               Index           =   2
               Left            =   2355
               TabIndex        =   20
               ToolTipText     =   "Ctrl+��ѡ������ѡ��"
               Top             =   30
               Value           =   1  'Checked
               Width           =   480
            End
            Begin VB.CheckBox chk�������� 
               Caption         =   "Σ"
               Height          =   195
               Index           =   1
               Left            =   1785
               TabIndex        =   21
               ToolTipText     =   "Ctrl+��ѡ������ѡ��"
               Top             =   30
               Value           =   1  'Checked
               Width           =   465
            End
            Begin VB.CheckBox chk�������� 
               Caption         =   "һ��"
               Height          =   195
               Index           =   0
               Left            =   1035
               TabIndex        =   22
               ToolTipText     =   "Ctrl+��ѡ������ѡ��"
               Top             =   30
               Value           =   1  'Checked
               Width           =   660
            End
            Begin VB.Label lbl�������� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ǰ����(&S)"
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
            Begin VB.Label lbl��Ժʱ�� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժʱ��"
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
            Begin VB.CheckBox chk���� 
               Caption         =   "δ����"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   1
               Left            =   1845
               TabIndex        =   34
               Top             =   50
               Value           =   1  'Checked
               Width           =   915
            End
            Begin VB.CheckBox chk���� 
               Caption         =   "�ѽ���"
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
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            IDKindStr       =   $"frmInNurseStation.frx":15EE
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
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label lblDept 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ����(&U)"
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
               Key             =   "�ȴ����"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":277E
               Key             =   "�ܾ����"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":2D18
               Key             =   "�������"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":32B2
               Key             =   "���ڳ��"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":3CC4
               Key             =   "��鷴��"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":46D6
               Key             =   "��鷴��"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":4C70
               Key             =   "�������"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":5682
               Key             =   "�������"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":6094
               Key             =   "δ����"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":662E
               Key             =   "ִ����"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":6BC8
               Key             =   "������"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":75DA
               Key             =   "��������"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":7B74
               Key             =   "�������"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":810E
               Key             =   "Child"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":86A8
               Key             =   "������"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":EF0A
               Key             =   "Out"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":F4A4
               Key             =   "����"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInNurseStation.frx":FA3E
               Key             =   "Fbaby"
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox pic�������� 
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
         Begin VB.ListBox lst�������� 
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
            ToolTipText     =   "ȷ��"
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
            ToolTipText     =   "ȡ��"
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
    col_����ȼ� = 12
    col_�ѱ� = 13
    col_���� = 14
    col_סԺҽʦ = 15
    col_��Ժ���� = 16
    col_��Ժ���� = 17
    col_�������� = 18
    col_���￨ = 19
    col_����ID = 20
    col_סԺ���� = 21
    col_������ = 22
    COL_Ӥ������ID = 23
    COL_Ӥ������ID = 24
    col_��ҽ��� = 25
    col_��ҽ��� = 26
    col_���λ�ʿ = 27
    col_��λ���� = 28
    col_���ۺ� = 29
    col_˳��� = 30
End Enum
Private Enum NOTIFYREPORT_COLUMN
    c_ͼ�� = 0
    C_����ID = 1
    C_��ҳID = 2
    c_���� = 3
    c_סԺ�� = 4
    c_���� = 5
    C_״̬ = 6
    
    '������
    C_��Ϣ = 7
    C_��� = 8
    C_���� = 9
    C_ҵ�� = 10
    C_���ﲡ�� = 11
End Enum
Private Enum PATI_TYPE
    pt��Ժ����ס = 0
    ptת�ƴ���ס = 1
    ptת��������ס = 2
    pt��Ժ = 3
    ptԤ�� = 4
    pt��Ժ = 5
    pt���� = 6
    pt���ת�� = 7
End Enum

Private Type PatiInfo
    ״̬ As Integer '������ҳ.״̬  0-����סԺ��1-��δ��ס��2-����ת�ƻ�����ת������3-��Ԥ��Ժ
    ���� As Integer '0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
    Ӥ�� As Integer
    סԺ�� As String
    ���� As String
    ��ҳID As Long
    �����ҳID As Long   '������Ϣ.��ҳID
    ����ID As Long
    ����ID As Long  'ת�Ʋ��ˣ��ǵ�ǰ����ID
    ���� As Boolean
    ��Ժ���� As Date
    ��Ժ���� As Date
    סԺ���� As Long
    ����ת�� As Boolean
    ���� As Integer
    ���� As Boolean
End Type
Public Enum EFun
    E��ס = 0
    Eת�� = 1
    E���� = 2
    E���� = 3
    E��Ժ = 4
    EתΪסԺ = 5
    E���Ĵ�λ�ȼ� = 6
    E����������Ϣ = 7
    E�������Ǽ� = 8
    E������� = 9
    Eҽ������ѡ�� = 10
    E���� = 11
    E�޸ĳ�Ժʱ�� = 12
    E��λ�Ի� = 13
    Eתҽ��С�� = 14
    Eת���� = 15
    Eת������ס = 16
    E���˱�ע�༭ = 17
End Enum

Public Enum tbcPatiEnu
    E����ס = 0
    E��Ժ = 1
    E��Ժ = 2
    Eת�� = 3
End Enum

'�Ӵ��������
Private mclsEMR As Object  '�°没��zlRichEMR.clsDockEMR
Private WithEvents mclsAdvices As zlPublicAdvice.clsDockInAdvices
Attribute mclsAdvices.VB_VarHelpID = -1
Private WithEvents mclsEPRs As zlRichEPR.cDockInEPRs
Attribute mclsEPRs.VB_VarHelpID = -1
Private WithEvents mclsTends As zlRichEPR.cDockInTends
Attribute mclsTends.VB_VarHelpID = -1
Private WithEvents mclsFeeQuery As zl9InExse.clsFeeQuery
Attribute mclsFeeQuery.VB_VarHelpID = -1
Private WithEvents mfrmResponse As frmAuditResponse '��鷴������
Attribute mfrmResponse.VB_VarHelpID = -1
Private mclsInPatient As zl9InPatient.clsInPatient
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private WithEvents mclsPath As zlPublicPath.clsDockPath
Attribute mclsPath.VB_VarHelpID = -1
Private mclsWardMonitor As clsWardMonitor     '�໤�ǽӿ�
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mobjKernel As zlPublicAdvice.clsPublicAdvice         '�ٴ����Ĳ���
Private mobjSquareCard As Object      '���������
Private mstrCardKind As String        '��������󷵻صĿ��õ�ҽ�ƿ�
Private mblnIsInit As Boolean

Private mcolSubForm As Collection
Private mfrmActive As Form

'54621:������,2013-02-28,��ʿվ�����ҳ������
Private mclsInOutMedRec As zlMedRecPage.clsInOutMedRec
'�������ñ���
Private mintChange As Integer
Private mdtOutBegin As Date, mdtOutEnd As Date
Private mintNotify As Integer 'ҽ�������Զ�ˢ�¼��(����)
Private mintNotifyDay As Integer '���Ѷ������ڵ�ҽ��
Private mstrNotifyAdvice As String '���ѵ�ҽ������
Private mlngMedRedDay As Long     '������鷴������

'�����������
Private mstrPrivs As String
Private mlngModul As Long
Private mPatiInfo As PatiInfo '��ʷסԺ��¼�е�,��һ��Ϊ��ǰ��
Private mlng����ID As Long, mlng��ҳID As Long, mlng����ID As Long '�����嵥�е�
Private mblnOutDept As Boolean '�Ƿ������������Ŀ��ң��������۲�����ʾ����ţ�

Private mintFindType As Integer '0-����,1-סԺ��,2-���￨,3-����
Private mstrFindType As String '�洢��ǰ�������͵�����
Private mblnFindTypeEnabled As Boolean
Private mintPreDept As Integer
Private mstrPrePati As String
Private mstrPreNotify As String
Private mintPrePage As Integer
Private mstrUnits As String '����Ա�����Ĳ�����
Private mblnNoCheck As Boolean
Private mblnUnRefresh As Boolean
Private mblnEditState As Boolean '�Ӵ����Ƿ��ڱ༭״̬
Private mblnReturn As Boolean        'cboUnit�س�����
Private mintOutPreTime As Integer

Private mblnMonitor As Boolean '�໤�ǳ����Ƿ����
Private mstrMonitor As String '�໤�ǳ���·��
Private mblnIsFindAgain As Boolean
Private mbytSize As Byte '�����С 0-С���壨9������) 1-�����壨12 �����壩
Private mblnNoRefNotify As Boolean '��ˢ��ҽ������
Private mintMecStandard As Integer  '������ҳ��ʽ 0-��������׼��1-�Ĵ�ʡ��׼��2-����ʡ��׼
Private mblnTabTmp As Boolean
Private mclsMsg As clsCISMsg
Private mrsMsg As ADODB.Recordset
Private mbln��Ϣ���� As Boolean
Private mblnCardOrder As Boolean '������true����λ��false������+��λ

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

Private Sub chk����_Click(Index As Integer)
    If chk����(0).Value = 0 And chk����(1).Value = 0 Then
        chk����((Index + 1) Mod 2).Value = 1
    End If
    If Me.Visible Then Call LoadPatients
End Sub

Private Sub InitSelectTime()
    Dim datCurr As Date
    
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    mdtOutEnd = datCurr
    mdtOutBegin = mdtOutEnd - 1
    
    cboSelectTime.Clear '��Ժ
    With cboSelectTime
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
    If cboSelectTime.ListCount > 0 Then cboSelectTime.ListIndex = 0
End Sub

Private Sub cboSelectTime_Click()
'���ܣ���ʱ�䷶Χ��ָ���ǣ�����ʱ��ѡ����
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    intDateCount = cboSelectTime.ItemData(cboSelectTime.ListIndex)
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If cboSelectTime.ListIndex = mintOutPreTime And intDateCount <> -1 Then Exit Sub
    If intDateCount = -1 Then
        If Not frmSelectTime.ShowMe(Me, mdtOutBegin, mdtOutEnd, cboSelectTime) Then
            'ȡ��ʱ�ָ�ԭ����ѡ��
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
        cboSelectTime.ToolTipText = "��Χ��" & Format(mdtOutBegin, "yyyy-MM-dd") & " �� " & Format(mdtOutEnd, "yyyy-MM-dd")
    End If
    '�����������֤ÿ���ط���ȡ�ĳ�Ժ���˶�����ͬһʱ�䷶Χ�ڣ�72783��
    Call zlDatabase.SetPara("��Ժ���˽������", DateDiff("d", datCurr, mdtOutEnd), glngSys, pסԺ��ʿվ)
    Call zlDatabase.SetPara("��Ժ���˿�ʼ���", DateDiff("d", mdtOutBegin, datCurr), glngSys, pסԺ��ʿվ)
    mintOutPreTime = cboSelectTime.ListIndex
    Call LoadPatients
End Sub

Private Sub cmdRef_Click()
'����ת��������ˢ��
    Call txtChange_KeyPress(vbKeyReturn)
End Sub

Private Sub cmdFilterCancel_Click()
    txt��������.SetFocus
    pic��������.Visible = False
End Sub

Private Sub cmdFilterCancel_LostFocus()
    If Not Me.ActiveControl Is cmdFilterOK _
        And Not Me.ActiveControl Is cmdFilterCancel _
        And Not Me.ActiveControl Is lst�������� _
        And Not Me.ActiveControl Is pic�������� Then pic��������.Visible = False
End Sub

Private Sub cmdFilterOK_Click()
    Dim i As Integer
    
    If lst��������.SelCount = 0 Then
        MsgBox "������ѡ��һ�ֻ���ȼ���", vbInformation, gstrSysName
        lst��������.SetFocus
    End If
    
    If lst��������.Selected(0) Then
        txt��������.Text = "ȫ��"
        txt��������.Tag = ""
    Else
        txt��������.Text = ""
        txt��������.Tag = ""
        For i = 1 To lst��������.ListCount - 1
            If lst��������.Selected(i) Then
                txt��������.Text = txt��������.Text & "," & lst��������.List(i)
                txt��������.Tag = txt��������.Tag & "," & lst��������.ItemData(i)
            End If
        Next
        txt��������.Text = Mid(txt��������.Text, 2)
        txt��������.Tag = Mid(txt��������.Tag, 2)
    End If
    
    txt��������.SetFocus
    pic��������.Visible = False
    
    '���¶�ȡ����
    Call LoadPatients
End Sub

Private Sub cmdFilterOK_LostFocus()
    If Not Me.ActiveControl Is cmdFilterOK _
        And Not Me.ActiveControl Is cmdFilterCancel _
        And Not Me.ActiveControl Is lst�������� _
        And Not Me.ActiveControl Is pic�������� Then pic��������.Visible = False
End Sub

Private Sub cmd��������_Click()
    Dim i As Integer
    
    For i = 0 To lst��������.ListCount - 1
        If txt��������.Tag = "" Then
            lst��������.Selected(i) = True
        ElseIf InStr("," & txt��������.Tag & ",", "," & lst��������.ItemData(i) & ",") > 0 Then
            lst��������.Selected(i) = True
        Else
            lst��������.Selected(i) = False
        End If
    Next
    lst��������.ListIndex = 0
    pic��������.Top = cmd��������.Top + cmd��������.Height + 30 + picPatiFilter.Top
    pic��������.Left = txt��������.Left
    pic��������.Width = txt��������.Width
    pic��������.Visible = True
    pic��������.ZOrder
    lst��������.SetFocus
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
    Dim blnCol As Boolean, strTmp As String, i As Long, bln·��״̬ As Boolean
    Dim intType As Integer
    Dim arrTmp As Variant
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    mblnNoCheck = False
    mblnNoRefNotify = False
    
    '��ȡ��������
    '-----------------------------------------------------
    mintPreDept = -1
    mstrPrePati = ""
    mstrPreNotify = ""
    mintPrePage = -1
    
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, pסԺ��ʿվ, GetInsidePrivs(pסԺ��ʿվ))
    Call AddMipModule(mclsMipModule)
    Set mobjKernel = New zlPublicAdvice.clsPublicAdvice
    
    Call GetLocalSetting '���ز���
    
    '����ָ�:Ĭ�ϲ������Ͷ�ȡ
    '-----------------------------------------------------
    mintFindType = Val(zlDatabase.GetPara("���˲��ҷ�ʽ", glngSys, pסԺ��ʿվ, , , , intType))
    mblnFindTypeEnabled = Not ((intType = 3 Or intType = 15) And InStr(mstrPrivs, "��������") = 0)
    
    mstrMonitor = ""
    mblnMonitor = Dir(App.Path & "\..\gdhs\AC2005.exe") <> ""
    If mblnMonitor Then mstrMonitor = App.Path & "\..\gdhs\AC2005.exe"
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)
    mstrCardKind = "��|����|0|0|0|0|0|0;ס|סԺ��|0|0|0|0|0|0;��|���￨|0|0|8|0|0|0;��|����|0|0|0|0|0|0;��|���ۺ�|0|0|0|0|0|0"
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
    Set objPane = Me.dkpMain.CreatePane(1, IIf(mbytSize = 0, 280, 320), 300, DockLeftOf, Nothing)
    objPane.Title = "סԺ�����б�"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    Set objPane = Me.dkpMain.CreatePane(2, 280, 100, DockBottomOf, objPane)
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
        mcolSubForm.Add mclsEMR.zlGetForm, "_�²���"
    End If
    mcolSubForm.Add mclsPath.zlGetForm, "_·��"
    mcolSubForm.Add mclsAdvices.zlGetForm, "_ҽ��"
    mcolSubForm.Add mclsFeeQuery.zlGetForm, "_����"
    mcolSubForm.Add mclsEPRs.zlGetForm, "_����"
    mcolSubForm.Add mclsTends.zlGetForm, "_����"
    If mclsWardMonitor.Enabled Then
        mcolSubForm.Add mclsWardMonitor.zlGetForm, "_�໤"
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
        If GetInsidePrivs(p�ٴ�·��Ӧ��, True) <> "" Then
            .InsertItem(intIdx, "�ٴ�·��", picTmp.hwnd, 0).Tag = "·��": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(pסԺҽ���´�, True) <> "" Or GetInsidePrivs(pסԺҽ������, True) <> "" Then
            .InsertItem(intIdx, "ҽ����Ϣ", picTmp.hwnd, 0).Tag = "ҽ��": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p���ò�ѯ, True) <> "" Then
            .InsertItem(intIdx, "������Ϣ", picTmp.hwnd, 0).Tag = "����": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(pסԺ��������, True) <> "" Then
            .InsertItem(intIdx, "������Ϣ", picTmp.hwnd, 0).Tag = "����": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p�°�סԺ����, True) <> "" And Not mclsEMR Is Nothing Then
            .InsertItem(intIdx, "���Ӳ���", picTmp.hwnd, 0).Tag = "�²���": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p�����¼����, True) <> "" Then
            .InsertItem(intIdx, "������Ϣ", picTmp.hwnd, 0).Tag = "����": intIdx = intIdx + 1
        End If
        If mclsWardMonitor.Enabled Then
            If InStr(GetInsidePrivs(pסԺ��ʿվ), "����໤") > 0 Then
                .InsertItem(intIdx, "����໤", picTmp.hwnd, 0).Tag = "�໤": intIdx = intIdx + 1
            End If
        End If
        
        '��Ҳ����еĿ�Ƭ
        Call CreatePlugInOK(pסԺ��ʿվ)
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strTmp = gobjPlugIn.GetFormCaption(glngSys, pסԺ��ʿվ)
            Call zlPlugInErrH(err, "GetFormCaption")
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ",")
                For i = 0 To UBound(arrTmp)
                    strTmp = arrTmp(i)
                    
                    mcolSubForm.Add gobjPlugIn.GetForm(glngSys, pסԺ��ʿվ, strTmp), "_" & strTmp
                    .InsertItem(intIdx, strTmp, mcolSubForm("_" & strTmp).hwnd, 0).Tag = strTmp: intIdx = intIdx + 1
                    Call zlPlugInErrH(err, "GetForm")
                Next
            End If
            err.Clear: On Error GoTo 0
        End If
        
        If .ItemCount = 0 Then
            MsgBox "��û��ʹ��סԺ��ʿ����վ��Ȩ�ޡ�", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        '�ָ��ϴ�ѡ��Ŀ�Ƭ
        strTab = zlDatabase.GetPara("ҽ������", glngSys, pסԺ��ʿվ)
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
        
        .Item(3).Selected = True
        .Item(1).Selected = True
        '��λ����ѡ�
        tbcPati.Item(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(tbcPati), "tbcPati", 1)).Selected = True
    End With
    
    '������������
    Call InitReportColumn
    picPatiFilter.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    chk��������(0).BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    chk��������(1).BackColor = chk��������(0).BackColor
    chk��������(2).BackColor = chk��������(0).BackColor
    picPati.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    picInfo.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    Call Cbo.SetListWidth(cbo����.hwnd, cbo����.Width * 2)
    Call Cbo.SetListWidth(cboPages.hwnd, cboPages.Width * 1.5)
    
    '��ʼ�����˹�������
    strTmp = zlDatabase.GetPara("��ǰ��������", glngSys, pסԺ��ʿվ, "111", _
        Array(lbl��������, chk��������(0), chk��������(1), chk��������(2)), InStr(mstrPrivs, "��������") > 0)
    For i = 0 To chk��������.UBound
        chk��������(i).Value = IIf(Mid(strTmp, i + 1, 1) = "1", 1, 0)
    Next
    If Not InitNurselevel Then Unload Me: Exit Sub
    If Not InitUnits Then Unload Me: Exit Sub
    If cboUnit.ListIndex = -1 Then
        If InStr(mstrPrivs, "ȫԺ����") > 0 Then
            MsgBox "û�з���סԺ������Ϣ,���ȵ����Ź��������ã�", vbInformation, gstrSysName
        Else
            MsgBox "û�з�������������,����ʹ��סԺ��ʿ����վ��", vbInformation, gstrSysName
        End If
        Unload Me: Exit Sub
    End If
    Call cboUnit_Click
    Call HaveRIS(True)
    
    'ת����������
    txtChange.Text = mintChange
    mintOutPreTime = -1
    Call InitSelectTime
    
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    End If
    If Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "NotifyRang", "0")) = 0 Then
        optNotify(1).Value = True
    Else
        optNotify(0).Value = True
    End If
    blnCol = rptPati.Columns(col_���).Visible
    bln·��״̬ = rptPati.Columns(col_·��״̬).Visible
    
    Call RestoreWinState(Me, App.ProductName, , True)
    Me.WindowState = vbMaximized
    Call SetFixedCommandBar(cbsMain(2).Controls)
    rptPati.Columns(col_���).Visible = blnCol
    
    rptPati.Columns(col_·��״̬).Visible = bln·��״̬
    If bln·��״̬ And rptPati.Columns(col_·��״̬).Width = 0 Then rptPati.Columns(col_·��״̬).Width = 18
    
    If tbcSub.Selected.Tag = "ҽ��" Then '�Է������������
        If dkpMain.Panes(2).Closed Then dkpMain.Panes(2).Closed = False
    End If
    Call LoadNotify
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long, strPrivs As String, strTmp As String
    Dim byt��ס��ʽ As Byte, blnRefresh As Boolean
    
    If Control.ID <> 0 Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    strPrivs = GetInsidePrivs(Enum_Inside_Program.p�������)
    Select Case Control.ID
    '����˵����������ת
    Case conMenu_Manage_Change_In
        If CheckBabyInOut Then Exit Sub
        If rptPati.SelectedRows(0).Record(col_����).Value = ptת��������ס Then
            blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.Eת������ס, Me, strPrivs, mlng����ID, mlng����ID, mlng��ҳID, "", 0)
        Else
            If rptPati.SelectedRows(0).Record(col_����).Value = ptת�ƴ���ס Then
                byt��ס��ʽ = 1 '0-��Ժ��ס��1-ת����ס
            Else
                byt��ס��ʽ = 0
            End If
            blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E��ס, Me, strPrivs, mlng����ID, mlng����ID, mlng��ҳID, "", _
                Val(rptPati.SelectedRows(0).Record(col_����ID).Value), byt��ס��ʽ)
        End If
    Case conMenu_Manage_Change_Turn
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.Eת��, Me, strPrivs, mlng����ID, mlng����ID, mlng��ҳID)
        
    Case conMenu_Manage_Change_TurnUnit
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.Eת����, Me, strPrivs, mlng����ID, mlng����ID, mlng��ҳID)
        
    Case conMenu_Manage_Change_TurnTeam
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.Eתҽ��С��, Me, strPrivs, mlng����ID, mlng����ID, mlng��ҳID)
        
    Case conMenu_Manage_Change_Bed
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E����, Me, strPrivs, mlng����ID, mlng����ID, mlng��ҳID, 0, "", "")
        
    Case conMenu_Manage_Change_TransposeBed
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E��λ�Ի�, Me, strPrivs, mlng����ID, mlng����ID, mlng��ҳID, mPatiInfo.����, "")
        
        
    Case conMenu_Manage_Change_House    'Ŀǰ�������û����ӿ�
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E����, Me, strPrivs, mlng����ID, mlng����ID, mlng��ҳID, 1, "", "")
        
    Case conMenu_Manage_Change_Out
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E��Ժ, Me, strPrivs, mlng����ID, mlng��ҳID)
        
    Case conMenu_Manage_Change_InPati
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.EתΪסԺ, Me, strPrivs, mlng����ID, mlng��ҳID, _
                Val(rptPati.SelectedRows(0).Record(col_סԺ��).Value), CStr(rptPati.SelectedRows(0).Record(col_����).Value))
        
        
    Case conMenu_Manage_Change_BedGrid
        If CheckBabyInOut Then Exit Sub
        Call mclsInPatient.zl_ExecPatiChange(EFun.E���Ĵ�λ�ȼ�, Me, strPrivs, mlng����ID, mlng��ҳID, _
            Trim(CStr(rptPati.SelectedRows(0).Record(col_����).Value)))
    Case conMenu_Manage_Change_PatiInfo
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E����������Ϣ, Me, strPrivs, mlng����ID, mlng����ID, mlng��ҳID)
       
    Case conMenu_Manage_Change_PaitNote
        If CheckBabyInOut Then Exit Sub
        Call mclsInPatient.zl_ExecPatiChange(EFun.E���˱�ע�༭, Me, strPrivs, mlng����ID, mlng��ҳID)
        
    Case conMenu_Manage_Change_Baby
        If CheckBabyInOut Then Exit Sub
        Call mclsInPatient.zl_ExecPatiChange(EFun.E�������Ǽ�, Me, strPrivs, mlng����ID, mlng��ҳID)
    Case conMenu_Manage_Change_ReCalcFee
        If CheckBabyInOut Then Exit Sub
        Call mclsInPatient.zl_ExecPatiChange(EFun.E�������, Me, strPrivs, mlng����ID, mlng��ҳID, _
        CStr(rptPati.SelectedRows(0).Record(col_����).Value))
    Case conMenu_Manage_Change_InsureSel
        If CheckBabyInOut Then Exit Sub
        Call mclsInPatient.zl_ExecPatiChange(EFun.Eҽ������ѡ��, Me, strPrivs, mlng����ID, mlng��ҳID, Val(mPatiInfo.����))
        
    Case conMenu_Manage_Change_Undo * 10 + 1    '����
        If CheckBabyInOut Then Exit Sub
        blnRefresh = mclsInPatient.zl_ExecPatiChange(EFun.E����, Me, strPrivs, mlng����ID, mlng����ID, mlng��ҳID, Val(mPatiInfo.����), Control.Caption)
        
    Case conMenu_Manage_Monitor '�໤��
        Call ExecuteMonitor
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
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_FontSize_L '������
        If mbytSize <> 1 Then
            mbytSize = 1
            Call SetFontSize(True)
            Me.cbsMain.RecalcLayout
        End If
    Case conMenu_View_Jump '��ת
        If Me.tbcSub.Selected.Index + 1 <= Me.tbcSub.ItemCount - 1 Then
            Me.tbcSub.Item(Me.tbcSub.Selected.Index + 1).Selected = True
        Else
            Me.tbcSub.Item(0).Selected = True
        End If
    Case conMenu_Tool_Archive '���Ӳ�������
        mblnUnRefresh = True
        Call frmArchiveView.ShowArchive(Me, mlng����ID, mPatiInfo.��ҳID)
        mblnUnRefresh = False
    Case conMenu_Tool_Reference_1 '������ϲο�
        mblnUnRefresh = True
        Call gobjKernel.ShowDiagHelp(vbModeless, Me)
        mblnUnRefresh = False
    Case conMenu_Tool_Reference_2 '���ƴ�ʩ�ο�
        mblnUnRefresh = True
        Call gobjKernel.ShowClincHelp(vbModeless, Me)
        mblnUnRefresh = False
    Case conMenu_Manage_FeeItemSet  '������Ŀ��������
        Call Set������Ŀ��������
    '54621:������,2013-02-28,��ʿվ�����ҳ������
    Case conMenu_Tool_MedRec '��ҳ����
        mblnUnRefresh = True
        Call ExecuteEditMediRec
        mblnUnRefresh = False
    Case conMenu_File_MedRecSetup '��ҳ��ӡ����
        Call PrintInMedRec(mclsInOutMedRec, 0, mlng����ID, mPatiInfo.��ҳID, mobjReport, mPatiInfo.����ID, Me)
    Case conMenu_File_MedRecPreview '��ҳԤ��
        Call PrintInMedRec(mclsInOutMedRec, 1, mlng����ID, mPatiInfo.��ҳID, mobjReport, mPatiInfo.����ID, Me)
    Case conMenu_File_MedRecPrint '��ҳ��ӡ
        Call PrintInMedRec(mclsInOutMedRec, 2, mlng����ID, mPatiInfo.��ҳID, mobjReport, mPatiInfo.����ID, Me)
    Case conMenu_Manage_Print_Label '�����ӡ
        If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_4", Me) Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_4", Me, "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID, 2)
        End If
    Case conMenu_Tool_MedRecAuditResponse '��鷴��
        '�����Ե��ã����ٿ��Բ鿴(��ǰ����ʷ)
        Call lbl���_Click

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
    Case conMenu_View_Notify And tbcSub.Selected.Tag <> "����" 'ҽ������(��һ���嵥��ID��ͬ��)
        If rptNotify.Visible Then Call LoadNotify
    Case conMenu_View_Refresh 'ˢ��
        blnRefresh = True
         
    Case conMenu_File_Parameter '��������
        mblnUnRefresh = True
        frmInStationSetup.mbln��ʿվ = True
        frmInStationSetup.mstrPrivs = mstrPrivs
        frmInStationSetup.Show 1, Me
        If gblnOK Then
            Call GetLocalSetting
            blnRefresh = True
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
    Case Else
        mblnUnRefresh = True
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            'ִ�з�������ǰģ��ı���
            With mPatiInfo
                strTmp = Split(Control.Parameter, ",")(1)
                If strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1132" Then 'סԺ�����ձ�
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strTmp, Me, _
                             "����=" & mlng����ID, "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID)
                ElseIf strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1139_2" Or strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1139_1" Then    '������ҳ�ʹ߿��
                    Call mclsFeeQuery.zlExecuteCommandBars(Control)
                Else
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strTmp, Me, _
                        "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID, "סԺ��=" & .סԺ��, "���˲���=" & .����ID, _
                        "���˿���=" & .����ID, "����=" & .����)
                End If
            End With
        ElseIf Between(Control.ID, conMenu_File_MedRecPrint * 100# + 1, conMenu_File_MedRecPrint * 100# + 6) Or Between(Control.ID, conMenu_File_MedRecPreview * 100# + 1, conMenu_File_MedRecPreview * 100# + 4) Then
            Call PrintInMedRec(mclsInOutMedRec, IIf(Between(Control.ID, conMenu_File_MedRecPrint * 100# + 1, conMenu_File_MedRecPrint * 100# + 6), 2, 1), mlng����ID, mPatiInfo.��ҳID, mobjReport, mPatiInfo.����ID, Me, Val(Mid(Control.ID & "", Len(Control.ID & ""))))
        Else
            Select Case Me.tbcSub.Selected.Tag
            Case "·��"
                If CheckBabyInOut Then Exit Sub
                Call mclsPath.zlExecuteCommandBars(Control)
            Case "ҽ��"
                Call mclsAdvices.zlExecuteCommandBars(Control)
            Case "����"
                Call mclsFeeQuery.zlExecuteCommandBars(Control)
            Case "����"
                Call mclsEPRs.zlExecuteCommandBars(Control)
            Case "����"
                Call mclsTends.zlExecuteCommandBars(Control)
            Case "�²���"
                Call mclsEMR.zlExecuteCommandBars(Control)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.ExeButtomClick(glngSys, pסԺ��ʿվ, mcolSubForm("_" & tbcSub.Selected.Tag), tbcSub.Selected.Tag, Control.Caption, mlng����ID, mlng��ҳID, "")
                    Call zlPlugInErrH(err, "ExeButtomClick")
                    err.Clear: On Error GoTo 0
                End If
            End Select
        End If
        mblnUnRefresh = False
    End Select
    
    If blnRefresh Then Call LoadPatients
    
    If Control.ID = conMenu_View_Refresh Then Call LoadNotify 'ˢ��ҽ������

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
            If .Count = 0 Then '��̬�Ӳ˵�,��1λ
                .Add xtpControlButton, conMenu_View_FindType * 100# + 1, "��  ��(&1)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 2, "סԺ��(&2)"
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
    Case conMenu_Manage_Change_Undo
        With CommandBar.Controls
            .DeleteAll
            If mlng����ID = 0 Then Exit Sub
            
            Set rsPatiLog = GetPatiLog(mlng����ID, mlng��ҳID)
            If rsPatiLog.RecordCount > 0 Then '��̬�Ӳ˵�,��1λ
                
                strPrivs = GetInsidePrivs(Enum_Inside_Program.p�������)
                rsPatiLog.MoveFirst
                For i = 1 To rsPatiLog.RecordCount
                    If Not IsNull(rsPatiLog!��ֹʱ��) And rsPatiLog!��ֹԭ�� = 1 Then
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Undo * 10 + i, "��Ժ")
                        j = j + 1
                        If InStr(";" & strPrivs & ";", ";������Ժ;") = 0 Or j > 1 Then objControl.Enabled = False
                    Else
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Undo * 10 + i, rsPatiLog!����)
                        If rsPatiLog.RecordCount > 1 And rsPatiLog!��ʼԭ�� = 1 Then objControl.Visible = False
                        j = j + 1
                        If j > 1 Then
                            objControl.Enabled = False
                        Else
                            If (objControl.Caption Like "*��ס" Or objControl.Caption = "ת������ס") Then
                                If InStr(strPrivs, "�������") = 0 Then objControl.Enabled = False
                            End If
                            If objControl.Caption = "תΪסԺ����" Then
                                If InStr(strPrivs, "סԺ����תסԺ") = 0 Then objControl.Enabled = False
                            ElseIf objControl.Caption = "Ԥ��Ժ" Then
                                If InStr(strPrivs, "����Ԥ��Ժ") = 0 Then objControl.Enabled = False
                                
                            ElseIf objControl.Caption = "����" Then
                                If InStr(strPrivs, "����") = 0 Then objControl.Enabled = False
                            End If
                        End If
                    End If
                    objControl.Category = "����"
                    If i <> 1 Then objControl.Enabled = False
                    rsPatiLog.MoveNext
                Next
            End If
        End With
    Case Else
       Select Case tbcSub.Selected.Tag
       Case "·��"
            Call mclsPath.zlPopupCommandBars(CommandBar)
       Case "ҽ��"
           Call mclsAdvices.zlPopupCommandBars(CommandBar)
       Case "����"
           Call mclsFeeQuery.zlPopupCommandBars(CommandBar)
       Case "����"
       
       Case "����"
       
       End Select
    End Select
    
End Sub


Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'���ܣ�����Ȩ�����ò�����صĲ˵��͹������Ŀɼ�״̬
    Dim blnVisible As Boolean, strPrivs As String

    'Ȩ��ֻ���ж�һ��,�Ѿ��жϹ�����������ж�
    If Control.Parameter = "���ж�" Then Exit Sub

    blnVisible = True
    strPrivs = GetInsidePrivs(Enum_Inside_Program.p�������)
    
    Select Case Control.ID
        Case conMenu_Manage_Change_In
            blnVisible = strPrivs <> ""
        Case conMenu_Manage_Change_Out
            blnVisible = InStr(strPrivs, "���˳�Ժ") > 0
        Case conMenu_Manage_Change_Turn
            blnVisible = InStr(strPrivs, "����ת��") > 0
        Case conMenu_Manage_Change_Bed, conMenu_Manage_Change_TransposeBed, conMenu_Manage_Change_House
            blnVisible = InStr(strPrivs, "����") > 0
        Case conMenu_Manage_Change_TurnUnit
            blnVisible = InStr(strPrivs, "ת����") > 0
        Case conMenu_Manage_Change_PatiInfo
            blnVisible = InStr(strPrivs, "����������Ϣ") > 0
        Case conMenu_Manage_Change_Baby
            blnVisible = InStr(strPrivs, "�������Ǽ�") > 0
        Case conMenu_Manage_Change_ReCalcFee
            blnVisible = InStr(strPrivs, "�������") > 0
        Case conMenu_Manage_Change_BedGrid
            blnVisible = InStr(strPrivs, "������λ�ȼ�") > 0
        Case conMenu_Manage_Change_InPati
            blnVisible = InStr(strPrivs, "סԺ����תסԺ") > 0
    End Select

    Control.Visible = blnVisible
    Control.Parameter = "���ж�"
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
                MsgBox "ҽ�ƿ�������zl9CardSquare����ʼ��ʧ��!", vbInformation, gstrSysName
            End If
            Call PatiIdentify.zlInit(Me, glngSys, pסԺ��ʿվ, gcnOracle, gstrDBUser, mobjSquareCard, mstrCardKind, "zl9CISJob")
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
    
    If Control.Category = "����" Then Exit Sub '��cbsMain_InitCommandsPopup������,�˳������Ӵ���������ɼ���
    
    If mblnEditState Then
        Select Case Control.ID
        Case conMenu_Help_Help, conMenu_File_Exit
            '������
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
    '��������£�Control.Enabled=falseʱ����Ҫ��ҵ�񲿼����жϣ����Բ����˳�
        
    If rptPati.SelectedRows.Count > 0 Then blnSelect = Not rptPati.SelectedRows(0).GroupRow
    If blnSelect Then
        lngType = Val(rptPati.SelectedRows(0).Record(col_����).Value)
        blnWaitIn = lngType = ptת�ƴ���ס Or lngType = pt��Ժ����ס Or lngType = ptת��������ס
        blnOut = lngType = pt��Ժ
        blnPreOut = lngType = ptԤ��
        blnOutTo = lngType = pt���ת��
    End If
    
    If Control.Category = "����" Then
        Call SetControlVisible(Control)
        If Not Control.Visible Then Exit Sub
        
        strPrivs = GetInsidePrivs(Enum_Inside_Program.p�������)
        If InStr(strPrivs, "���в���") = 0 Then
            If InStr("," & mstrUnits & ",", "," & mlng����ID & ",") = 0 Then Control.Enabled = False: Exit Sub
        End If
        If blnSelect = False Then Control.Enabled = False: Exit Sub
    End If
            
    Select Case Control.ID
    Case conMenu_Manage_Change_Undo
        Control.Enabled = blnSelect And Not blnWaitIn And Not blnOutTo And mlng��ҳID = mPatiInfo.�����ҳID
    Case conMenu_Manage_FeeItemSet  '������Ŀ��������,û��Ȩ��ʱ�ɲ鿴
        
    Case conMenu_Manage_Change_In   '��ס
        Control.Enabled = blnWaitIn
        
    Case conMenu_Manage_Change_InPati   'תΪסԺ
        Control.Enabled = Not blnWaitIn And Not blnOut And Not blnPreOut
        If Control.Enabled Then
            Control.Enabled = mPatiInfo.���� = 2
        End If
        
    'ת�ƣ�����������������������Ϣ���������,ת������תС��,��λ�Ի�
    Case conMenu_Manage_Change_Turn, conMenu_Manage_Change_Bed, conMenu_Manage_Change_House, _
         conMenu_Manage_Change_PatiInfo, conMenu_Manage_Change_ReCalcFee, conMenu_Manage_Change_TurnUnit, _
         conMenu_Manage_Change_TurnTeam, conMenu_Manage_Change_TransposeBed
         
        Control.Enabled = Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
        If Control.Enabled Then
            Control.Enabled = mPatiInfo.״̬ <> 2
            
            If Control.ID = conMenu_Manage_Change_TransposeBed Then '��λ�Ի�
                Control.Enabled = Trim(CStr(rptPati.SelectedRows(0).Record(col_����).Value)) <> ""
            End If
        End If
    Case conMenu_Manage_Change_InsureSel
        Control.Enabled = Not blnWaitIn And Not blnPreOut
        If Control.Enabled Then
            Control.Enabled = mPatiInfo.���� <> 0
        End If
    Case conMenu_Manage_Change_BedGrid
        Control.Enabled = Not blnWaitIn And Not blnOut And Not blnPreOut And Not blnOutTo
        If Control.Enabled Then
            Control.Enabled = Trim(CStr(rptPati.SelectedRows(0).Record(col_����).Value)) <> "" And mPatiInfo.״̬ <> 2
        End If
    Case conMenu_Manage_Change_Out
        Control.Enabled = Not blnWaitIn And Not blnOut
        If Control.Enabled Then
            Control.Enabled = (rptPati.SelectedRows(0).Record(col_����).Value = pt��Ժ Or blnPreOut) And mPatiInfo.״̬ <> 2
        End If
    Case conMenu_Manage_Change_Baby
        Control.Enabled = Not blnWaitIn And Not blnOut And Not blnPreOut
        If Control.Enabled Then
            Control.Enabled = mPatiInfo.���� And rptPati.SelectedRows(0).Record(col_�Ա�).Value = "Ů"
        End If
    Case conMenu_Manage_Monitor '�໤��
        Control.Visible = mblnMonitor
        
    Case conMenu_Manage_Change_PaitNote
        Control.Enabled = Not blnOutTo
        
        
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
    Case conMenu_Tool_Archive '���Ӳ�������
        If GetInsidePrivs(p���Ӳ�������) = "" Then
            Control.Visible = False
        Else
            Control.Enabled = mlng����ID <> 0
        End If
    Case conMenu_Tool_Reference_1 '������ϲο�
        If GetInsidePrivs(p������ϲο�) = "" Then Control.Visible = False
    Case conMenu_Tool_Reference_2 'ҩƷ�����Ʋο�
        If GetInsidePrivs(pҩƷ���Ʋο�) = "" Then Control.Visible = False
    Case conMenu_Tool_MedRecAuditResponse '��鷴��
        '�����Ե��ã����ٿ��Բ鿴(��ǰ����ʷ)
        Control.Enabled = rptPati.Rows.Count > 0
    Case conMenu_View_Notify And tbcSub.Selected.Tag <> "����" 'ҽ������
        Control.Enabled = rptNotify.Visible
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
    Case conMenu_Manage_Print_Label '�����ӡ
        If InStr(mstrPrivs, ";�����ӡ;") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = Not blnOut And blnSelect
        End If
         
    '54621:������,2013-02-28,��ʿվ�����ҳ������
    Case conMenu_Tool_MedRec '��ҳ����
        blnWriteMedRec = Val(zlDatabase.GetPara("ҽ���ͻ�ʿ�ֱ���д������ҳ", glngSys, pסԺҽ��վ, "0")) = 1
        Control.Enabled = cboPages.ListIndex <> -1 And mPatiInfo.��ҳID > 0 And blnWriteMedRec = True
        Control.Visible = blnWriteMedRec
    Case conMenu_File_Parameter '��������
        'If InStr(mstrPrivs, "��������") = 0 Then Control.Visible = False
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
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If Split(Control.Parameter, ",")(1) = "ZL" & glngSys \ 100 & "_INSIDE_1139_1" Then
                Control.Visible = tbcSub.Selected.Tag = "����"  '�߿��
                Exit Sub
            End If
            If Split(Control.Parameter, ",")(1) = "ZL" & glngSys \ 100 & "_INSIDE_1139_2" Then
                Control.Visible = tbcSub.Selected.Tag = "����"  '������ҳ
                Exit Sub
            End If
        End If
        '��ҳ����
        If Between(Control.ID, conMenu_File_MedRecPrint * 100# + 3, conMenu_File_MedRecPrint * 100# + 6) Or Between(Control.ID, conMenu_File_MedRecPreview * 100# + 3, conMenu_File_MedRecPreview * 100# + 4) Then
            If mintMecStandard = 0 Or mintMecStandard = 3  Or mintMecStandard = 1 Then
                Control.Visible = False
            Else
                Control.Visible = True
            End If
            Exit Sub
        End If
    
        Select Case tbcSub.Selected.Tag
        Case "·��"
            Call mclsPath.zlUpdateCommandBars(Control)
        Case "ҽ��"
            Call mclsAdvices.zlUpdateCommandBars(Control)
         Case "����"
            Call mclsFeeQuery.zlUpdateCommandBars(Control)
        Case "����"
            Call mclsEPRs.zlUpdateCommandBars(Control)
        Case "����"
            Call mclsTends.zlUpdateCommandBars(Control)
        Case "�²���"
            Call mclsEMR.zlUpdateCommandBars(Control)
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
        
    Me.Caption = "סԺ��ʿ����վ - " & objItem.Caption & "(��ǰ�û���" & UserInfo.���� & ")"
    
    If mstrNotifyAdvice = "000000000000" Then
        dkpMain.Panes(2).Tag = IIf(dkpMain.Panes(2).Hidden, 1, 0)
        dkpMain.Panes(2).Close
    Else
        dkpMain.Panes(2).Closed = False
        dkpMain.Panes(2).Hidden = Val(dkpMain.Panes(2).Tag) = 1
        dkpMain.Panes(2).Title = "��Ϣ����"
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
    
    '�Ӵ������¼���
    Select Case objItem.Tag
    Case "·��"
        Call mclsPath.zlDefCommandBars(Me, Me.cbsMain, 1)
    Case "ҽ��"
        Call mclsAdvices.zlDefCommandBars(Me, Me.cbsMain, 1)
    Case "����"
        Call mclsFeeQuery.zlDefCommandBars(Me, Me.cbsMain, 1)
    Case "����"
        Call mclsEPRs.zlDefCommandBars(Me.cbsMain)
    Case "����"
        Call mclsTends.zlDefCommandBars(Me.cbsMain)
    Case "�²���"
        Call mclsEMR.zlDefCommandBars(Me.cbsMain)
    Case Else
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strName = gobjPlugIn.GetButtomName(glngSys, pסԺ��ʿվ, mcolSubForm("_" & objItem.Tag), objItem.Tag)
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
    
    '�������RecalcLayout����������
    Call LockWindowUpdate(0)
    
    Set mfrmActive = mcolSubForm("_" & objItem.Tag)
End Sub

Private Sub SubWinRefreshData(ByVal objItem As TabControlItem)
'���ܣ�ˢ���Ӵ������ݼ�״̬
    Dim blnEdit As Boolean, strInPatiNO As String, lng·��״̬ As Long
    Dim lngType As PATI_TYPE, lng����ID As Long, lng����ID As Long
    Dim lngState As TYPE_PATI_State
    
    If mlng����ID = 0 Then
        'Ҫ���Ӵ��尴�����ݴ������
        Select Case objItem.Tag
        Case "·��"
            Call mclsPath.zlRefresh(0, 0, 0, 0, 0, False)
        Case "ҽ��"
            Call mclsAdvices.zlRefresh(0, 0, 0, 0, CDate(0), 0)
        Case "����"
            Call mclsFeeQuery.zlRefresh(0, 0, 0, 0, 0, False, False, False)
        Case "����"
            Call mclsEPRs.zlRefresh(0, 0, 0, False, False)
        Case "����"
            Call mclsTends.zlRefresh(0, 0, 0, False, False)
        Case "�໤"
            Call mclsWardMonitor.HideWindow
        Case "�²���"
            Call mclsEMR.zlRefresh(0, 0, 0, 0, 3)
        Case Else
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                Call gobjPlugIn.RefreshForm(glngSys, pסԺ��ʿվ, mcolSubForm("_" & objItem.Tag), objItem.Tag, 0, "", 0, False)
                Call zlPlugInErrH(err, "RefreshForm")
                err.Clear: On Error GoTo 0
            End If
        End Select
    Else
        With mPatiInfo
            lngType = Val(rptPati.SelectedRows(0).Record(col_����).Value)
            lng����ID = Val("" & rptPati.SelectedRows(0).Record(col_����ID).Value)    '���ת������Ϊԭ����ID
            
            If InStr("," & pt��Ժ����ס & "," & pt���ת�� & "," & ptת�ƴ���ס & "," & ptת��������ס & ",", "," & lngType & ",") > 0 Then
                '����ס���ˣ�ת�����ˣ�����ǰ����Ĳ���
                lng����ID = mlng����ID
            Else
                lng����ID = .����ID
            End If
            
            If lngType = pt���ת�� Then
                lngState = ps���ת��
            ElseIf lngType = ptת�ƴ���ס Or lngType = ptת��������ס Then
                lngState = ps��ת��
            Else
                lngState = IIf(.��Ժ���� = CDate(0), IIf(.״̬ = 3, psԤ��, ps��Ժ), ps��Ժ)
            End If
                        
        
            Select Case objItem.Tag
            Case "·��"
                Call mclsPath.zlRefresh(mlng����ID, .��ҳID, lng����ID, lng����ID, .״̬, .����ת��, True, , , mclsMipModule)
            Case "ҽ��"
                lng·��״̬ = Val(rptPati.SelectedRows(0).Record(col_·��״̬).Value)
                If .״̬ = 1 Then '��Ժ����ס
                   If Val(zlDatabase.GetPara("���������ס�����´�ҽ��", glngSys, pסԺҽ���´�, 1)) = 0 Then
                        lngState = ps��ת�� 'lngState=ps��ת��ʱ�¿�ҽ���ȹ��ܲ�����
                   End If
                End If
                Call mclsAdvices.zlRefresh(mlng����ID, .��ҳID, lng����ID, lng����ID, lngState, .����ת��, , , , lng·��״̬, mlng����ID, mclsMipModule, .Ӥ��)
                
            Case "����"
                Call mclsFeeQuery.zlRefresh(mlng����ID, mlng��ҳID, Val(.סԺ��), lng����ID, .����, .����ת��, .��Ժ���� <> CDate("0:00:00"), .����, False, _
                    lngType = pt���ת�� Or lngType = ptԤ�� Or lngType = pt��Ժ, lng����ID)
               
            Case "����"
                Call mclsEPRs.zlRefresh(mlng����ID, .��ҳID, mlng����ID, False, .����ת��, 0, tbcSub.Tag = "·��", lng����ID, lngState)
            Case "����"
                blnEdit = True
                With rptPati.SelectedRows(0)
                    If lngType = pt��Ժ Or lngType = pt���� Then
                        If Not (.Record(col_���).Value = 0 Or .Record(col_���).Value = 2 Or .Record(col_���).Value = 999) Then
                            '��������Ժ��鷴��״̬����Ժ��δ�ύ���
                            If .Record(col_ͼ��).Value = 1 Then blnEdit = False
                        End If
                    ElseIf lngType = ptת�ƴ���ס Or lngType = ptת��������ס Then
                        blnEdit = False
                    End If
                End With
                blnEdit = blnEdit And (mlng����ID = .����ID Or lngType = pt���ת��)
                Call mclsTends.zlRefresh(mlng����ID, .��ҳID, mlng����ID, blnEdit, False, lng����ID, lngState)
            Case "�໤"
                strInPatiNO = Trim(rptPati.SelectedRows(0).Record(col_סԺ��).Value)
                If strInPatiNO = "" Then
                    Call mclsWardMonitor.HideWindow
                Else
                    Call mclsWardMonitor.ShowInfor(strInPatiNO)
                End If
            Case "�²���"
                Call mclsEMR.zlRefresh(mlng����ID, .��ҳID, mlng����ID, lngState, 3)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.RefreshForm(glngSys, pסԺ��ʿվ, mcolSubForm("_" & objItem.Tag), objItem.Tag, mlng����ID, "", .��ҳID, .����ת��, , , _
                                    lng����ID, lng����ID, , lngState, , lng·��״̬)
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
            .Add xtpControlSplitButtonPopup, conMenu_File_MedRecPreview, "��ӡԤ��(&V)", -1, False
            .Add xtpControlSplitButtonPopup, conMenu_File_MedRecPrint, "��ӡ��ҳ(&P)", -1, False
        End With
        '49854:������,2013-10-31,���������ӡ
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Print_Label, "��ӡ���(&W)��")  '��ӡ���
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True '����
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "����(&P)", -1, False) '����
    objMenu.ID = conMenu_ManagePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_In, "��ס(&I)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Turn, "ת��(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_TurnUnit, "ת����(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_TurnTeam, "תС��(&T)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Bed, "����(&B)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_TransposeBed, "��λ�Ի�(&Q)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_House, "����(&H)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_BedGrid, "���Ĵ�λ�ȼ�(&G)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_PatiInfo, "����סԺ��Ϣ(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_PaitNote, "���˱�ע��Ϣ(&F)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Out, "��Ժ(&O)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_InPati, "תΪסԺ����(&Z)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_Baby, "�������Ǽ�(&N)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_ReCalcFee, "���ѱ��������(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Change_InsureSel, "ҽ������ѡ��(&M)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Manage_Change_Undo, "����(&U)"): objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Edit_Untread
                        
        '�໤��
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Monitor, "�໤��(&N)")
        objControl.BeginGroup = True
    End With
    For Each objControl In objMenu.CommandBar.Controls
        objControl.Category = "����"
    Next


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
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "���Ӳ�������(&I)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "���ϲο�(&R)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_Reference_1, "������ϲο�(&D)", -1, False
            .Add xtpControlButton, conMenu_Tool_Reference_2, "���ƴ�ʩ�ο�(&C)", -1, False
        End With
        
        '54621:������,2013-02-28,��ʿվ�����ҳ������
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRec, "��ҳ����(&M)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRecAuditResponse, "��鷴��(&S)")
            objControl.IconId = 3814
            objControl.BeginGroup = True
            objControl.ToolTipText = "�����鿴������鷴��"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_FeeItemSet, "������Ŀ��������(&C)")
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
    Set objBar = cbsMain.Add("������", xtpBarTop) '����
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��") '����
        
        '54621:������,2013-02-28,��ʿվ�����ҳ������
        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRec, "��ҳ"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlSplitButtonPopup, conMenu_File_MedRecPreview, "Ԥ����ҳ")
        objControl.IconId = conMenu_File_Preview
        Set objControl = .Add(xtpControlSplitButtonPopup, conMenu_File_MedRecPrint, "��ӡ��ҳ")
        objControl.IconId = conMenu_File_Print
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Monitor, "�໤��"): objControl.BeginGroup = True '����
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����") '����
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
'    With cbsMain.Options
'        .AddHiddenCommand conMenu_File_PrintSet '��ӡ����
'        .AddHiddenCommand conMenu_File_Excel '�����Excel
'        .AddHiddenCommand conMenu_View_Jump '��ת
'    End With
    
    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL1_INSIDE_1261_1", "ZL1_INSIDE_1261_5", "ZL1_INSIDE_1261_4", "ZL1_INSIDE_1261_6", "ZL1_INSIDE_1261_7", "ZL1_INSIDE_1261_8", _
            "ZL1_INSIDE_1261_9", "ZL1_INSIDE_1261_10")
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub lbl���_Click()
    If cboUnit.ListIndex = -1 Then Exit Sub
    
    '��ģ̬��ʾ��鷴������
    If mfrmResponse Is Nothing Then
        Set mfrmResponse = New frmAuditResponse
    End If
    mblnUnRefresh = True
    Call mfrmResponse.ShowMe(Me, cboUnit.ItemData(cboUnit.ListIndex), 1, False, 1, mstrPrivs)
    mblnUnRefresh = False
End Sub

Private Sub lst��������_ItemCheck(Item As Integer)
    Dim i As Integer
    
    If Item = 0 Then
        For i = 1 To lst��������.ListCount - 1
            lst��������.Selected(i) = lst��������.Selected(0)
        Next
    ElseIf Not lst��������.Selected(Item) Then
        lst��������.Selected(0) = False
    ElseIf lst��������.SelCount = lst��������.ListCount - 1 Then
        lst��������.Selected(0) = True
    End If
End Sub

Private Sub lst��������_LostFocus()
    If Not Me.ActiveControl Is cmdFilterOK _
        And Not Me.ActiveControl Is cmdFilterCancel _
        And Not Me.ActiveControl Is lst�������� _
        And Not Me.ActiveControl Is pic�������� Then pic��������.Visible = False
End Sub

Private Sub mclsAdvices_SetEditState(ByVal blnEditState As Boolean)
'���ܣ��Ӵ��崦�ڱ༭״̬ʱ����ֹʹ�ò˵���ת�ƽ�������й���
    picPati.Enabled = Not blnEditState
    picInfo.Enabled = Not blnEditState
    picNotify.Enabled = Not blnEditState
    
    mblnEditState = blnEditState
    mblnUnRefresh = blnEditState
End Sub

Private Sub mclsAdvices_ExecLogModi(ByVal ҽ��ID As Long, ByVal ���ͺ� As Long, ByVal ����ID As Long, ByVal ִ��ʱ�� As String, ��� As Boolean)
    On Error Resume Next
    mblnUnRefresh = True
    ��� = frmTechnicLog.ShowMe(Me, pסԺҽ������, ����ID, ҽ��ID, ���ͺ�, False, ִ��ʱ��)
    mblnUnRefresh = False
    On Error GoTo 0
End Sub

Private Sub mclsAdvices_ExecLogNew(ByVal ҽ��ID As Long, ByVal ���ͺ� As Long, ByVal ����ID As Long, ��� As Boolean)
    On Error Resume Next
    mblnUnRefresh = True
    ��� = frmTechnicLog.ShowMe(Me, pסԺҽ������, ����ID, ҽ��ID, ���ͺ�, False)
    mblnUnRefresh = False
    On Error GoTo 0
End Sub

Private Sub mclsAdvices_RequestRefresh(ByVal RefreshNotify As Boolean)
'���ܣ�ҽ���Ӵ���Ҫ��ˢ��
    If Not RefreshNotify Then
        Call LoadPatients
    ElseIf rptNotify.Visible Then
        '��ˢ��ҽ����������
        Call LoadNotify
    End If
End Sub

Private Sub mclsMipModule_OpenLink(ByVal strMsgKey As String, ByVal strLinkPara As String)
'���ܣ����ð����Ϣ��λ����
    Dim int�б� As Integer
    
    int�б� = -1
    If InStr(",ZLHIS_PATIENT_002,ZLHIS_PATIENT_012,ZLHIS_PATIENT_009,ZLHIS_PATIENT_006,ZLHIS_PATIENT_010,", "," & strMsgKey & ",") > 0 Then
        int�б� = E��Ժ
    ElseIf InStr(",ZLHIS_PATIENT_003,", "," & strMsgKey & ",") > 0 Then
        int�б� = E����ס
    End If
    
    If int�б� <> -1 Then
        If tbcPati.Item(int�б�).Selected = False Then
            tbcPati.Item(int�б�).Selected = True 'ѡ��л�ʱ��ˢ�²����б�
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
'���ܣ���Ϣ����
    Dim blnRecToLis As Boolean '�Ƿ���ص������б���
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

Private Sub mclsAdvices_StatusTextUpdate(ByVal Text As String)
'���ܣ�ҽ���Ӵ���Ҫ�����״̬��
    Dim strTmp As String
    Dim intTmp As String
    If Text = "" And rptPati.SelectedRows.Count > 0 Then
        With rptPati.SelectedRows(0)
            If Not .GroupRow Then
                If Val(.Record(col_����Id).Value) <> 0 Then intTmp = 1
            End If
            If intTmp = 1 Then
                stbThis.Panels(2).Text = IIf(stbThis.Panels(2).Tag = "", "", stbThis.Panels(2).Tag)
                lblFee(1).Caption = GetPati������Ϣ(mlng����ID, mlng��ҳID)
                
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

Private Sub cboPages_Click()
'���ܣ�ѡ��ĳ��סԺ��¼ʱ����ȡ��صĲ�����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If cboPages.ListIndex = -1 Then Exit Sub
    If cboPages.ListIndex = mintPrePage Then Exit Sub
    mintPrePage = cboPages.ListIndex
    
    On Error GoTo errH
    strSQL = "Select a.��ҳID As �����ҳID,NVL(b.����,a.����) ����, NVL(b.�Ա�,a.�Ա�) �Ա�, NVL(b.����,a.����) ����, b.סԺ��, b.��Ժ����, b.ҽ�Ƹ��ʽ, d.��Ϣֵ As ҽ����, b.����, b.��ǰ����, c.���� As ����ȼ�, Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) as ��Ժ����, b.��Ժ����, b.��Ŀ����," & vbNewLine & _
            "       b.��������, b.״̬, b.����ת��, b.��Ժ����id, b.��ǰ����id, a.סԺ����, e.�����" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B, �շ���ĿĿ¼ C, ������ҳ�ӱ� D, ��λ״����¼ E" & vbNewLine & _
            "Where a.����id = b.����id And a.����id = [1] And b.��ҳid = [2] And b.����ȼ�id = c.Id(+) And b.����id = d.����id(+) And" & vbNewLine & _
            "      b.��ҳid = d.��ҳid(+) And d.��Ϣ��(+) = 'ҽ����' And b.��Ժ����id = e.����id(+) And b.����id = e.����id(+) And b.��Ժ���� = e.����(+)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, cboPages.ItemData(cboPages.ListIndex))
    With rsTmp
        '���ղ���������ɫ��ʾ
        lbl����(1).Caption = "" & !סԺ��
        lbl����(1).ForeColor = zlDatabase.GetPatiColor(Nvl(!��������))
        lblPatiName(1).Caption = "" & !����
        lblPatiName(1).ToolTipText = lblPatiName(1).Caption
        lblPatiName(1).ForeColor = lbl����(1).ForeColor
        
        lblҽ����(1).Caption = Nvl(!ҽ����)
        lbl����(1).Caption = Nvl(!����ȼ�)
        lbl����(1).Caption = Nvl(!ҽ�Ƹ��ʽ)
        
        'Σ�ز��˲�����ɫ��ʾ
        lbl����(1).Caption = Nvl(!��ǰ����)
        If Nvl(!��ǰ����) = "Σ" Or Nvl(!��ǰ����) = "��" Or Nvl(!��ǰ����) = "��" Then
            lbl����(1).ForeColor = &HC0&
        Else
            lbl����(1).ForeColor = lblҽ����(1).ForeColor
        End If
        
        lbl��Ժ(1).Caption = Format(!��Ժ����, "yyyy-MM-dd HH:mm")
        If Not IsNull(!��Ժ����) Then
            lbl��Ժ(1).Caption = lbl��Ժ(1).Caption & "��" & Format(!��Ժ����, "yyyy-MM-dd HH:mm")

        End If
        
        lbl����(1).Caption = Nvl(!��������)
        lbl����(1).Caption = IIf(IsNull(!�����), "", "(" & !����� & ")") & !��Ժ����
        
        '���
        lblDiag(1).Caption = GetPatiDiagnose(mlng����ID, cboPages.ItemData(cboPages.ListIndex), 2)
        
        '������Ϣ
        mPatiInfo.״̬ = Nvl(!״̬, 0)
        mPatiInfo.סԺ�� = Nvl(!סԺ��)
        mPatiInfo.���� = Nvl(!��Ժ����)
        mPatiInfo.��ҳID = cboPages.ItemData(cboPages.ListIndex)
        mPatiInfo.�����ҳID = Nvl(!�����ҳID, 0)
        mPatiInfo.����ID = Nvl(!��ǰ����ID, 0)
        mPatiInfo.����ID = Nvl(!��Ժ����ID, 0)
        mPatiInfo.��Ժ���� = !��Ժ����
        If Not IsNull(!��Ժ����) Then
            mPatiInfo.��Ժ���� = !��Ժ����
        Else
            mPatiInfo.��Ժ���� = CDate(0)
        End If
        mPatiInfo.סԺ���� = Nvl(!סԺ����, 0)
        mPatiInfo.����ת�� = Nvl(!����ת��, 0) <> 0
        
        Call SetPatiInfoCtlPos

    End With
    
        
    '������Ϣȡ��ǰסԺ������
    strSQL = "Select B.סԺ��,B.��������,B.����,B.��Ժ����ID,B.��ǰ����ID,Decode(Nvl(X.�������, 0), 0, '��', '') As ����" & _
        " From ������ҳ B,������� X" & _
        " Where B.����ID=[1] And B.��ҳID=[2] And B.����ID = X.����ID(+) And X.����(+) = 1 And X.����(+) = 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    With rsTmp
        mPatiInfo.���� = Val("" & !����)
        mPatiInfo.���� = Not IsNull(!����)
        mPatiInfo.���� = Nvl(!��������, 0)
        mPatiInfo.���� = Sys.DeptHaveProperty(Val(!��Ժ����ID & ""), "����")
    End With
        
    'ˢ���Ӵ�������
    Call SubWinRefreshData(tbcSub.Selected)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mclsAdvices_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
'���ܣ��鿴���Ӳ�������
    Call gobjRichEPR.ViewDocument(Me, ����ID, CanPrint)
End Sub

Private Sub mclspath_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
'���ܣ��ٴ�·���в鿴���Ӳ�������
    Call gobjRichEPR.ViewDocument(Me, ����ID, CanPrint)
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
    strTab = Decode(ObjectType, 1, "ҽ��", 2, "����", 3, "����", 4, "����", 5, "", 6, "ҽ��", 7, "����", 8, "����")
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
        Call mclsTends.zlLocateData(IIf(ObjectType = 3, 1, 0))
    End If
    
    '�򿪶�Ӧ�Ķ���
    Select Case ObjectType
    Case 1 'סԺҽ��
    Case 2, 3, 7, 8 'סԺ����,������,����֤��,֪���ļ�
        If ObjectID = "0" Or ObjectID = "" Then Exit Sub
        If IsNumeric(ObjectID) Then
            Call gobjRichEPR.EditDocument(pסԺ��ʿվ, Me, cboUnit.ItemData(cboUnit.ListIndex), ObjectID)
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
                        "       RAWTOHEX(a.Id) Docid, 2 Occasion, a.Sealed Besealed, nvl(e.Code,5) Docsecret, b.Subdoc_Id Subdocid,b.completor" & vbNewLine & _
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
            If Nvl(rsEmr!completor) = "" Then
                If InStr(strPrivs, ";�ĵ���д;") > 0 Then '����дȨ��
                    Call gobjEmr.OpenFormForModifyDoc(Me.hwnd, rsEmr!masterid, rsEmr!actlogid, Nvl(rsEmr!basiclogid), rsEmr!actionid, rsEmr!taskid, rsEmr!antetypeid, rsEmr!doctype, rsEmr!docid, CInt(rsEmr!Occasion), CInt(rsEmr!besealed), CInt(rsEmr!docsecret), Nvl(rsEmr!subdocid), 2, strPrivs)
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
                    Call gobjEmr.OpenFormForAuditDoc(Me.hwnd, rsEmr!masterid, rsEmr!actlogid, Nvl(rsEmr!basiclogid), rsEmr!actionid, rsEmr!taskid, rsEmr!antetypeid, rsEmr!doctype, rsEmr!docid, CInt(rsEmr!Occasion), CInt(rsEmr!besealed), CInt(rsEmr!docsecret), Nvl(rsEmr!subdocid), 2, strPrivs)
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
        Call PrintInMedRec(mclsInOutMedRec, 1, mlng����ID, mlng��ҳID, mobjReport, mPatiInfo.����ID, Me)
    Case 6 'ҽ������
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
        lngPatiID = objHisPati.����ID
    End If
    
    Call ExecuteFindPati(False, lngPatiID)
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

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnIsInit = True Then mintFindType = Index - 1: mstrFindType = objCard.����
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
    txt��������.Width = cboUnit.Width
    cmd��������.Left = txt��������.Left + txt��������.Width - cmd��������.Width - 30
    picPara(0).Width = lbl��������.Width + txt��������.Width + 200
End Sub

Private Sub pic��������_GotFocus()
    lst��������.SetFocus
End Sub

Private Sub pic��������_Resize()
    On Error Resume Next
    
    lst��������.Left = -15
    lst��������.Top = -15
    lst��������.Width = pic��������.Width
    
    cmdFilterCancel.Left = pic��������.ScaleWidth - cmdFilterCancel.Width - 100
    cmdFilterOK.Left = cmdFilterCancel.Left - cmdFilterOK.Width - 60
    
    cmdFilterOK.Top = lst��������.Height + (pic��������.ScaleHeight - lst��������.Height - cmdFilterOK.Height) / 2
    cmdFilterCancel.Top = cmdFilterOK.Top
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
        picPara(1).Visible = True
    ElseIf Item.Tag = "��Ժ" Then
        picPara(2).Visible = True
        picPara(3).Visible = True
    ElseIf Item.Tag = "ת��" Then
        picPara(4).Visible = True
    End If
    Call picPatiIn_Resize
    
    If Me.Visible Then
        Call LoadPatients
        If mblnNoRefNotify = False Then
            Call LoadNotify 'ˢ��ҽ������
        End If
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
            Case "·��"
                Set objItem = tbcSub.InsertItem(Index, "�ٴ�·��", mcolSubForm("_·��").hwnd, 0)
                objItem.Tag = "·��"
            Case "ҽ��"
                Set objItem = tbcSub.InsertItem(Index, "ҽ����Ϣ", mcolSubForm("_ҽ��").hwnd, 0)
                objItem.Tag = "ҽ��"
            Case "����"
                Set objItem = tbcSub.InsertItem(Index, "������Ϣ", mcolSubForm("_����").hwnd, 0)
                objItem.Tag = "����"
            Case "����"
                Set objItem = tbcSub.InsertItem(Index, "������Ϣ", mcolSubForm("_����").hwnd, 0)
                objItem.Tag = "����"
            Case "�²���"
                Set objItem = tbcSub.InsertItem(Index, "���Ӳ���", mcolSubForm("_�²���").hwnd, 0)
                objItem.Tag = "�²���"
            Case "����"
                Set objItem = tbcSub.InsertItem(Index, "������Ϣ", mcolSubForm("_����").hwnd, 0)
                objItem.Tag = "����"
            Case "�໤"
                Set objItem = tbcSub.InsertItem(Index, "����໤", mcolSubForm("_�໤").hwnd, 0)
                objItem.Tag = "�໤"
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
    tbcSub.Tag = Item.Tag   '��¼��һ��ѡ��Ŀ�Ƭ
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
'���ܣ�ˢ�½�������
'˵�����Ӹ��¼���ʼ�᲻�ظ�������ص����ݶ�ȡ,����ҽ������
    Dim i As Long, lngidx As Long
    
    mblnReturn = True
    
    If cboUnit.ListIndex = mintPreDept Then Exit Sub
    
    mintPreDept = cboUnit.ListIndex
        
    mlng����ID = Val(cboUnit.ItemData(cboUnit.ListIndex))
   
    
    '�ر�ҵ����
    If Not mfrmResponse Is Nothing Then
        Unload mfrmResponse
    End If
    
    '54621:������,2013-02-28,��ʿվ�����ҳ������
    If Not mclsInOutMedRec Is Nothing Then
        Call mclsInOutMedRec.FormUnLoad
    End If
    
    Call Sys.DeptHaveProperty(mlng����ID, "����", mblnOutDept)
    '���¶�ȡ����
    Call LoadPatients
    
    '��ʾ�ٴ�·����Ƭ
    lngidx = -1
    For i = 0 To tbcSub.ItemCount - 1
        If tbcSub(i).Tag = "·��" Then
            lngidx = i
            Exit For
        End If
    Next
    If lngidx >= 0 Then
        If HavePath(mlng����ID) = False Then
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
    'If Visible And rptPati.Visible Then rptPati.SetFocus
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
        If mclsWardMonitor.Enabled = False Or InStr(GetInsidePrivs(pסԺ��ʿվ), "����໤") = 0 Then
            objCol.Visible = False
        End If
        Set objCol = .Columns.Add(col_�Ա�, "�Ա�", 30, True)
        Set objCol = .Columns.Add(col_����, "����", 30, True)
        Set objCol = .Columns.Add(col_����ȼ�, "����ȼ�", 56, True)
        Set objCol = .Columns.Add(col_�ѱ�, "�ѱ�", 55, True)
        Set objCol = .Columns.Add(col_����, "����", 70, True)
        Set objCol = .Columns.Add(col_סԺҽʦ, "סԺҽʦ", 55, True)
        Set objCol = .Columns.Add(col_��Ժ����, "��Ժ����", 106, True)
        Set objCol = .Columns.Add(col_��Ժ����, "��Ժ����", 106, True)
        Set objCol = .Columns.Add(col_��������, "��������", 106, True)
        If ISPassShowCard Then
            Set objCol = .Columns.Add(col_���￨, "���￨", 0, False)
        Else
            Set objCol = .Columns.Add(col_���￨, "���￨", 70, True)
        End If
        Set objCol = .Columns.Add(col_����ID, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_סԺ����, "סԺ����", 56, True)
        Set objCol = .Columns.Add(col_������, "������", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_Ӥ������ID, "Ӥ������ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_Ӥ������ID, "Ӥ������ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_��ҽ���, "��ҽ���", 106, True)
        Set objCol = .Columns.Add(col_��ҽ���, "��ҽ���", 106, True)
        Set objCol = .Columns.Add(col_���λ�ʿ, "���λ�ʿ", 55, True)
        Set objCol = .Columns.Add(col_��λ����, "��λ����", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_���ۺ�, "���ۺ�", 62, True)
        Set objCol = .Columns.Add(col_˳���, "˳���", 50, True)
        
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
        .SortOrder.Add .Columns(col_˳���)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(col_��λ����)
        .SortOrder(1).SortAscending = True
        .SortOrder.Add .Columns(col_����)
        .SortOrder(2).SortAscending = True
        .SortOrder.Add .Columns(col_���)
        .SortOrder(3).SortAscending = True

    End With
    
    With rptNotify
        Set objCol = .Columns.Add(c_ͼ��, "", 18, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(C_����ID, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_��ҳID, "��ҳID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_����, "����", 60, True)
        Set objCol = .Columns.Add(c_סԺ��, "סԺ��", 62, True)
        Set objCol = .Columns.Add(c_����, "����", 40, True)
        Set objCol = .Columns.Add(C_״̬, "״̬", 150, True)
        
        Set objCol = .Columns.Add(C_��Ϣ, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_���, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_����, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_ҵ��, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_���ﲡ��, "", 0, False): objCol.Visible = False
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
        
    mlng����ID = 0
    mlng��ҳID = 0
    mlng����ID = 0
    
    blnSetup = InStr(";" & mstrPrivs & ";", ";��������;") > 0
    Call zlDatabase.SetPara("���˲��ҷ�ʽ", mintFindType, glngSys, pסԺ��ʿվ, blnSetup)
    Call zlDatabase.SetPara("����", mbytSize, glngSys, pסԺ��ʿվ, blnSetup)

    strTmp = ""
    For i = 0 To chk��������.UBound
        strTmp = strTmp & IIf(chk��������(i).Value = 1, "1", "0")
    Next
    Call zlDatabase.SetPara("��ǰ��������", strTmp, glngSys, pסԺ��ʿվ, blnSetup)
    Call zlDatabase.SetPara("����ȼ�����", txt��������.Tag, glngSys, pסԺ��ʿվ, blnSetup)
    '���˷�Χ
    curDate = zlDatabase.Currentdate
    Call zlDatabase.SetPara("���ת������", Val(txtChange.Text), glngSys, mlngModul, blnSetup)
    
    If Not tbcSub.Selected Is Nothing Then
        Call zlDatabase.SetPara("ҽ������", tbcSub.Selected.Tag, glngSys, pסԺ��ʿվ, blnSetup)
    End If
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    End If
    If Me.Visible Then
        '���������̶�����һ���ؼ�����ʽ���棬����վ���������һ���Ǵ�ӡ����̶���ͼ����ʽ,������ָ�Ϊ������ť����ʽ
        cbsMain(2).Controls(1).Style = cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style
        Call SaveWinState(Me, App.ProductName)
        If Not tbcPati.Selected Is Nothing Then
            Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(tbcPati), "tbcPati", tbcPati.Selected.Index)
        End If
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "NotifyRang", IIf(optNotify(0).Value = True, 1, 0))
    End If
    
    'ǿ��Unload,��Ȼ���ἤ���Ӵ�����¼�
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
    '54621:������,2013-02-28,��ʿվ�����ҳ������
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
    tbcPati.Height = picPati.ScaleHeight - tbcPati.Top - IIf(fra���.Visible, fra���.Height, 0)
    
    fra���.Left = 0
    fra���.Top = tbcPati.Top + tbcPati.Height
    fra���.Width = picPati.ScaleWidth
    
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
        '�����ڴ������ʱ��picPara�ĸ߶Ȳ��������ָ��
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

Private Sub rptNotify_KeyUp(KeyCode As Integer, Shift As Integer)
'���ܣ��Զ�����ҽ��У�ԡ�ȷ��ֹͣ��ִ�н���
    Dim objControl As CommandBarControl
    Dim lngIndex As Long, lng����ID As Long, lng��ҳID As Long
    Dim lngҽ��ID As Long, strSQLRead As String, strSQL As String
    Dim strҵ�� As String, str���� As String, strסԺ�� As String, str���� As String
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
                strNO = .Item(C_��Ϣ).Value
                strҵ�� = .Item(C_ҵ��).Value
                lng����ID = Val(.Item(C_����ID).Value)
                lng��ҳID = Val(.Item(C_��ҳID).Value)
                str���� = .Item(c_����).Value
                strסԺ�� = .Item(c_סԺ��).Value
                str���� = .Item(c_����).Value
                lngIndex = .Index
            End With
            If strNO = "ZLHIS_PACS_006" Or strNO = "ZLHIS_PACS_007" Then
                strSQLRead = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng��ҳID & ",'" & strNO & "',3,'" & UserInfo.���� & "'," & mlng����ID & ",null,null,'" & strҵ�� & "')"
            ElseIf strNO = "ZLHIS_BLOOD_007" And gblnѪ��ϵͳ Then     'δ����ǰ��������Ϊ�Ѷ�
                If gobjPublicBlood Is Nothing And gblnѪ��ϵͳ Then InitObjBlood
                If gobjPublicBlood.zlIsBloodMessageDone(1, lng����ID, lng��ҳID, 3, mlng����ID) Then
                    Call rptNotify.Records.RemoveAt(lngIndex)
                    Call rptNotify.Populate
                End If
                Exit Sub
            Else
                strSQLRead = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng��ҳID & ",'" & strNO & "',3,'" & UserInfo.���� & "'," & mlng����ID & ")"
            End If
            
            If rptPati.SelectedRows.Count > 0 Then
                If Not rptPati.SelectedRows(0).GroupRow Then
                    '�ǲ����Ѿ���λ
                    blnFinded = InStr("_" & rptPati.SelectedRows(0).Record.Tag & "_", "_" & rptNotify.SelectedRows(0).Record.Tag & "_") > 0
                End If
            End If
            '���û�ҵ����ˣ��ҽ���ûѡ����Ժ�����б�ʱ�л��б��ٲ���һ��
            If tbcPati.Item(tbcPatiEnu.E��Ժ).Selected = False And Not blnFinded Then
                mblnNoRefNotify = True
                tbcPati.Item(tbcPatiEnu.E��Ժ).Selected = True
                mblnNoRefNotify = False
                blnFinded = LocatePati(rptNotify.SelectedRows(0).Record.Tag)
            End If
            
            If blnFinded And strҵ�� <> "" And tbcSub.Tag = "ҽ��" Then   '�ҵ����˺��پ����Ƿ�λҽ��
                If strNO = "ZLHIS_LIS_003" Or strNO = "ZLHIS_PACS_005" Then '��������Ϣ��ҵ���д���� ҽ��id��������Դ
                    lngҽ��ID = Val(Split(strҵ��, ",")(0))
                ElseIf strNO = "ZLHIS_BLOOD_007" Then
                    lngҽ��ID = Val(Split(strҵ��, ":")(0))
                Else
                    lngҽ��ID = Val(strҵ��)
                End If
                If lngҽ��ID <> 0 Then
                    Call mclsAdvices.LocatedAdviceRow(lngҽ��ID)
                End If
            End If
            
            If strNO = "ZLHIS_CIS_001" Or strNO = "ZLHIS_CIS_002" Then
                '˫���¿�����ͣҽ����Ϣ���ѡ�ֻ����У��ʱ���п��ܲ���������ˣ����ͺ�ȷ��ֹֻͣ�����������ˡ�
                '1.������¿���ֱ�ӷ���ģʽʱ����ѯ ����ҽ����¼
                If blnFinded Then
                    'ҽ���ж�Ȩ��
                    strTmp = GetInsidePrivs(pסԺҽ������)
                    If strNO = "ZLHIS_CIS_001" Then
                        If Not HaveOperateAdvice(lng����ID, lng��ҳID, 0) Then
                            Call zlDatabase.ExecuteProcedure(strSQLRead, Me.Caption)
                            Call rptNotify.Records.RemoveAt(lngIndex)
                            Call rptNotify.Populate
                        Else
                            If Val(zlDatabase.GetPara("����ǰ�Զ�У��", glngSys, pסԺҽ������, 0)) = 1 Then
                                If InStr(strTmp, ";����ҩ������;") > 0 Or InStr(strTmp, ";����ҩ�Ƴ���;") > 0 _
                                    Or InStr(strTmp, ";������������;") > 0 Or InStr(strTmp, ";������������;") > 0 Then
                                    
                                    blnTmp = mobjKernel.AdviceSend(Me, mlng����ID, lng����ID, lng��ҳID, gstrPrivs, mclsMipModule)
                                    
                                    If Not HaveOperateAdvice(lng����ID, lng��ҳID, 0) Then
                                        Call zlDatabase.ExecuteProcedure(strSQLRead, Me.Caption)
                                        Call rptNotify.Records.RemoveAt(lngIndex)
                                        Call rptNotify.Populate
                                    End If
                                    
                                    If blnTmp Then
                                        If tbcSub.Selected.Tag = "ҽ��" Then Call SubWinRefreshData(tbcSub.Selected)
                                    End If
                                End If
                            Else
                                If InStr(strTmp, ";ҽ��У�Դ���;") > 0 Then
                                    blnOnePati = Val(zlDatabase.GetPara("����ҽ��У��", glngSys, pסԺҽ������)) = 0
                                    blnTmp = mobjKernel.AdviceOperate(Me, gstrPrivs, 3, lng����ID, lng��ҳID, mlng����ID, lngҽ��ID, mclsMipModule, strPatis, blnOnePati)
                                    If strPatis <> "" And blnTmp Then Call ReadMsg�¿�(strPatis)
                                    If blnTmp Then
                                        If tbcSub.Selected.Tag = "ҽ��" Then Call SubWinRefreshData(tbcSub.Selected)
                                    Else
                                        If Not HaveOperateAdvice(lng����ID, lng��ҳID, 0) Then
                                            Call zlDatabase.ExecuteProcedure(strSQLRead, Me.Caption)
                                            Call rptNotify.Records.RemoveAt(lngIndex)
                                            Call rptNotify.Populate
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    ElseIf strNO = "ZLHIS_CIS_002" Then
                        If InStr(strTmp, ";ҽ��ȷ��ֹͣ;") > 0 Then
                            If Not HaveOperateAdvice(lng����ID, lng��ҳID, 1) Then
                                Call zlDatabase.ExecuteProcedure(strSQLRead, Me.Caption)
                                Call rptNotify.Records.RemoveAt(lngIndex)
                                Call rptNotify.Populate
                            Else
                                blnTmp = mobjKernel.AdviceOperate(Me, gstrPrivs, 2, lng����ID, lng��ҳID, mlng����ID, lngҽ��ID, mclsMipModule, strPatis, True)
                                If strPatis <> "" And blnTmp Then
                                    Call rptNotify.Records.RemoveAt(lngIndex)
                                    Call rptNotify.Populate
                                End If
                                If blnTmp Then
                                    If tbcSub.Selected.Tag = "ҽ��" Then Call SubWinRefreshData(tbcSub.Selected)
                                Else
                                    If Not HaveOperateAdvice(lng����ID, lng��ҳID, 1) Then
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
                '��ʿվ���ܴ���Σ��ֵ��Ϣ
                If strNO = "ZLHIS_LIS_003" Or strNO = "ZLHIS_PACS_005" Then Exit Sub
                If strSQLRead = "" Then Exit Sub

                Call zlDatabase.ExecuteProcedure(strSQLRead, Me.Caption)
                Call rptNotify.Records.RemoveAt(lngIndex)
                Call rptNotify.Populate
            End If
        End If
    End If
End Sub

Private Sub ReadMsg�¿�(ByVal strPatis As String)
'���ܣ���Ϣ����ҽ���¿���Ϣ��
    Dim objRow As ReportRow
    Dim strTmp As String
    Dim strIndexs As String
    Dim arrTmp As Variant
    Dim i As Integer
    
    For Each objRow In rptNotify.Rows
        If objRow.GroupRow Then objRow.Expanded = True
        If Not objRow.GroupRow And objRow.Childs.Count = 0 Then
            If InStr(";" & strPatis & ";", ";" & objRow.Record.Tag & ";") > 0 And objRow.Record(C_��Ϣ).Value = "ZLHIS_CIS_001" Then
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
    Dim lngҽ��ID As Long
    Dim strNO As String
    Dim lng���ﲡ��ID As Long
    Dim lng����ID As Long
    Dim lng��ҳID As Long

    
    If rptNotify.SelectedRows.Count = 0 Then Exit Sub  '���������
    
    With rptNotify.SelectedRows(0)
        
        lngIndex = rptNotify.FocusedRow.Record.Index
        strNO = CStr(rptNotify.Rows(lngIndex).Record(C_��Ϣ).Value)
        lng����ID = Val(rptNotify.Rows(lngIndex).Record(C_����ID).Value)
        lng��ҳID = Val(rptNotify.Rows(lngIndex).Record(C_��ҳID).Value)
        lng���ﲡ��ID = Val(rptNotify.Rows(lngIndex).Record(C_���ﲡ��).Value)
                        
        If rptNotify.Rows(lngIndex).Record(C_��Ϣ).Value = "ZLHIS_LIS_003" Or rptNotify.Rows(lngIndex).Record(C_��Ϣ).Value = "ZLHIS_PACS_005" Then '��������Ϣ��ҵ���д���� ҽ��id��������Դ
            lngҽ��ID = Val(Split(rptNotify.Rows(lngIndex).Record(C_ҵ��).Value, ",")(0))
        ElseIf rptNotify.Rows(lngIndex).Record(C_��Ϣ).Value = "ZLHIS_BLOOD_007" Then    '�����Ϣ��ҵ���д���� ҽ��ID:�շ�ID
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
            '�Զ�Ѱ�Ҳ��л���ʾ��ǰ���ѵĲ���
            If Not LocatePati(.Record.Tag) And tbcSub.Tag = "ҽ��" Then
                Call LoadPatients
                If Not LocatePati(.Record.Tag) Then
                    Call ReadAndSendMsg(strNO, lng����ID, lng��ҳID, lng���ﲡ��ID)
                    Call LoadNotify
                    Exit Sub
                End If
            End If
        End If
        
        If lngҽ��ID <> 0 And tbcSub.Tag = "ҽ��" Then
            Call mclsAdvices.LocatedAdviceRow(lngҽ��ID)
        End If
        
    End With
    rptNotify.SetFocus
End Sub

Private Function LocatePati(ByVal strTag As String) As Boolean
'���ܣ�ͨ��reportControl��Record.Tagֵ��λ����
'����   strTag   reportControl��Record.Tag�������ݸ�ʽΪ"����ID,��ҳID"

    Dim blnEnabled As Boolean
    Dim objRow As ReportRow
    
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
End Function

Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    'Panne�е�Report�ؼ���Ҫǿ�д�����˳��
    '������ʱ���ܲ���vbKeyTab
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
            ElseIf objHitTest.Row.Childs.Count = 0 Or Val(objHitTest.Row.Record(col_����Id).Value) <> 0 Then
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
    
    If rptPati.SelectedRows.Count = 0 Then Exit Sub         '���������
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
            mlng����ID = Val(.Record(col_����Id).Value)
            mlng��ҳID = Val(.Record(col_��ҳID).Value)
            If InStr(strCurPati, "_") > 0 Then
                mPatiInfo.Ӥ�� = Val(Split(strCurPati, "_")(1))
            Else
                mPatiInfo.Ӥ�� = -1
            End If
            
            LockWindowUpdate Me.hwnd
            
            On Error GoTo errH
            strSQL = "Select ��ҳID,NVL(��������,0) �������� From ������ҳ Where ��ҳID<>0 And ����ID=[1] Order by ��ҳID Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
            cboPages.Clear
            Do While Not rsTmp.EOF
                cboPages.AddItem "�� " & rsTmp!��ҳID & " ��" & Decode(rsTmp!��������, 1, "(��������)", 2, "(סԺ����)", "")
                cboPages.ItemData(cboPages.NewIndex) = rsTmp!��ҳID
                If rsTmp!��ҳID = mlng��ҳID Then
                    Call Cbo.SetIndex(cboPages.hwnd, cboPages.NewIndex)
                End If
                rsTmp.MoveNext
            Loop
            If cboPages.ListIndex = -1 Then
                Call Cbo.SetIndex(cboPages.hwnd, 0)
            End If
            
            mintPrePage = -1
            Call cboPages_Click
            
            Call LoadPatiAllergy(mlng����ID, cbo����)
            
            '��Ժ���˶�ȡ�Ƿ����ύ���
            If .Record(col_����).Value = pt��Ժ Or .Record(col_����).Value = pt���� Then
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
            lblFee(1).Caption = GetPati������Ϣ(mlng����ID, mlng��ҳID)
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

            lblPrint(0).Visible = True
            lblPrint(1).Visible = True
            intTmp = Get����ҽ����ӡ(mlng����ID, mlng��ҳID)
            lblPrint(1).Caption = IIf(intTmp = 0, "δ��ӡ", IIf(intTmp = 1, "���ִ�ӡ", "ȫ����ӡ"))
            If Visible And rptPati.Visible Then rptPati.SetFocus
        Else
            Call ClearPatiInfo
            '��������ˢ���Ӵ���
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
'���ܣ���ע����ȡ��Ժ���˵�ʱ�䷶Χ
    Dim curDate As Date, intDay As Integer
    
    '������ʾ��Χ
    mintChange = Val(zlDatabase.GetPara("���ת������", glngSys, pסԺ��ʿվ, 7))
    '�������30���ȡȱʡֵ
    If mintChange > 30 Then mintChange = 7
    
    '��Ժ����ʱ�䷶Χ���̶�Ϊ��ȥ3��
    curDate = zlDatabase.Currentdate
    mdtOutEnd = Format(curDate, "yyyy-MM-dd 23:59:59")
    mdtOutBegin = Format(mdtOutEnd - 3, "yyyy-MM-dd 00:00:00")
        
    'ҽ������ˢ������
    mstrNotifyAdvice = zlDatabase.GetPara("�Զ�ˢ��ҽ������", glngSys, pסԺ��ʿվ, "000000000000")
    mintNotifyDay = Val(zlDatabase.GetPara("�Զ�ˢ��ҽ������", glngSys, pסԺ��ʿվ, 1))
    mintNotify = Val(zlDatabase.GetPara("�Զ�ˢ��ҽ�����", glngSys, pסԺ��ʿվ))
    mbln��Ϣ���� = Val(zlDatabase.GetPara("����������ʾ", glngSys, pסԺ��ʿվ)) = 1
    '������鷴������
    mlngMedRedDay = Val(zlDatabase.GetPara("������鷴������", glngSys, pסԺ��ʿվ))
    '��������
    mbytSize = zlDatabase.GetPara("����", glngSys, pסԺ��ʿվ, "0")
    
    '������ҳ��׼
    mintMecStandard = Val(zlDatabase.GetPara("������ҳ��׼", glngSys, pסԺҽ��վ, "0"))
    
    mblnCardOrder = (Val(zlDatabase.GetPara("��λ��Ƭ����ʽ", glngSys, P�°滤ʿվ, 0)) = 0)
    
End Sub

Private Function InitNurselevel() As Boolean
'���ܣ���ʼ��סԺ����ȼ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSel As String
    Dim blnSelAll As Boolean
    
    txt��������.Text = ""
    txt��������.Tag = ""
    
    lst��������.AddItem "ȫ��"
    strSel = zlDatabase.GetPara("����ȼ�����", glngSys, pסԺ��ʿվ, "", Array(lbl��������, txt��������, cmd��������), InStr(mstrPrivs, "��������") > 0)
    blnSelAll = True
    
    strSQL = _
        " Select ID,����,���� From �շ���ĿĿ¼ Where ���='H' And ��Ŀ����>=1" & _
        " And (����ʱ�� is NULL Or Trunc(����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " And (վ��='" & gstrNodeNo & "' Or վ�� is Null)" & _
        " Order by ����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "InitNurselevel")
    Do While Not rsTmp.EOF
        lst��������.AddItem rsTmp!����
        lst��������.ItemData(lst��������.NewIndex) = rsTmp!ID
        If strSel = "" Or InStr("," & strSel & ",", "," & rsTmp!ID & ",") > 0 Then
            txt��������.Text = txt��������.Text & "," & rsTmp!����
            txt��������.Tag = txt��������.Tag & "," & rsTmp!ID
        Else
            blnSelAll = False
        End If
        rsTmp.MoveNext
    Loop
    
    If blnSelAll Then
        txt��������.Text = "ȫ��"
        txt��������.Tag = ""
    Else
        txt��������.Text = Mid(txt��������.Text, 2)
        txt��������.Tag = Mid(txt��������.Tag, 2)
    End If
    
    '����������С
    lst��������.Height = lst��������.ListCount * 210 + 30
    pic��������.Height = lst��������.Height + cmdFilterOK.Height + 120
    
    InitNurselevel = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    mstrUnits = GetUser����IDs
    
    cboUnit.Clear
    Set rsTmp = GetDataToUnits
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!���� & "-" & rsTmp!����
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If InStr(mstrPrivs, "ȫԺ����") > 0 Then
                If rsTmp!ID = UserInfo.����ID Then 'ֱ����������
                    Call Cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
                If InStr("," & mstrUnits & ",", "," & rsTmp!ID & ",") > 0 And cboUnit.ListIndex = -1 Then
                    Call Cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
            Else '����ȱʡ���������Ŀ����ж��
                If rsTmp!ȱʡ = 1 And cboUnit.ListIndex = -1 Then
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
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    Dim objRow As ReportRow
    Dim strPatiRow As String, lngPatiRow As Long, intBedLen As Integer
    
    Dim strFilter As String, strMonitor As String
    Dim strSQL As String, i As Long, j As Long
    Dim lngCount(0 To 7) As Long, strState As String    '������ʾ���˷���ͳ����Ŀ
    Dim strTmpDate As String                            'ת�����˲�ѯʱ�䷶Χ����
    Dim blnIsFind As Boolean                            '�ж��Ƿ��ǲ���סԺ�ż�סԺ���Ƿ�Ϊ��
    Dim strTmpOut As String                             '��ѯ��Ժ����
    Dim int���� As Integer                              '��Ժ���˵Ľ���״̬
    
    Dim rsBaby As ADODB.Recordset
    Dim strSQLBaby As String
    Dim objBabyParent As ReportRecord
    Dim lngCol As Long
    Dim strPre�������� As String, str�ֶ���� As String
    Dim bln��λ���� As Boolean
    
    '��ҳ����������գ�F5ˢ�£�Ӧ�ûָ���һ����ֵ
    If cboUnit.ListIndex = -1 Then Call Cbo.SetIndex(cboUnit.hwnd, mintPreDept)
     
    '��λ���ȹ̶�Ϊ10
    intBedLen = 10
    mblnUnRefresh = True
    '���˹�������
    strFilter = ""
    strTmpDate = ""
    If txt��������.Tag <> "" Then
        strFilter = strFilter & " And Instr(','||[3]||',',','||B.����ȼ�ID||',')>0"
    End If
    strState = ""
    For i = 0 To chk��������.UBound
        If chk��������(i).Value = 1 Then
            strState = strState & "," & chk��������(i).Caption
        End If
    Next
    strState = Mid(strState, 2)
    If Not (UBound(Split(strState, ",")) = chk��������.UBound Or strState = "") Then
        strFilter = strFilter & " And Instr(','||[4]||',',','||B.��ǰ����||',')>0"
    End If
    
    If mintChange = 0 Then
        strTmpDate = ""
    Else
        strTmpDate = " And C.��ֹʱ�� Between Sysdate-[2] And Sysdate "
    End If
    
    '��Ժ����ס��ת�ƴ���ס����(���˿��������Ĳ������ɽ���),ת��������ס
    'c.����id + 0,˵����ͨ��H����������ӹ��˺󣬼�¼�������٣�������B�������
    '��������Ϊ�����͡���value
    If tbcPati.Selected.Tag = "����ס" Then
        strSQL = _
            "Select /*+ RULE */Distinct" & vbNewLine & _
            " Decode(B.״̬,1,0,Decode(c.��ʼԭ��,3,1,2)) As ����, Decode(Nvl(b.����״̬, 0), 0, 999, b.����״̬) As ����2," & _
            " Decode(B.״̬,1,'��Ժ����ס����',Decode(c.��ʼԭ��,3,'ת�ƴ���ס����','ת��������ס����')) As ����," & _
            " a.����id, b.��ҳid, A.�����,B.סԺ��, NVL(b.����,a.����) ����, NVL(b.�Ա�,a.�Ա�) �Ա�, NVL(b.����,a.����) ����," & vbNewLine & _
            " d.���� As ����, c.����id, c.����ҽʦ As סԺҽʦ, b.����״̬, LPAD(c.����," & intBedLen & ",' ') as ����," & _
            " e.���� As ����ȼ�, b.�ѱ�, Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) as ��Ժ����, b.��Ժ����, b.��������, b.״̬, b.����, a.���￨��,b.���ۺ�," & vbNewLine & _
            " Nvl(b.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��)) as סԺ����,q.��� as ���������,b.������,B.Ӥ������ID,B.Ӥ������ID,'' AS ��ҽ���,'' AS ��ҽ��� " & vbNewLine & _
            "From ������Ϣ A, ������ҳ B, ���˱䶯��¼ C, ���ű� D, �շ���ĿĿ¼ E,������������¼ Q" & vbNewLine & _
            "Where a.��Ժ = 1 And a.����id = b.����id  And a.��ҳid = b.��ҳid And Nvl(b.��ҳid, 0) <> 0 And b.����id = c.����id And b.��ҳid = c.��ҳid And (C.����ID=[1] or C.����ID is null) And c.����id = d.Id  And b.����id=q.����ID(+) And b.��ҳID=q.��ҳID(+) And (q.���=1 Or q.��� is Null) " & vbNewLine & _
            "      And (d.վ��='" & gstrNodeNo & "' Or d.վ�� is Null)" & vbNewLine & _
            "      And b.����ȼ�id = e.Id(+) And Nvl(c.���Ӵ�λ, 0) = 0 And c.��ֹʱ�� Is Null" & vbNewLine & _
            "      And (c.��ʼԭ�� in(1,3) And Exists(Select 1 From �������Ҷ�Ӧ H Where c.����id = h.����id And h.����id = [1]) or c.��ʼԭ��=15 And c.����id = [1])" & vbNewLine & _
            "      And ((c.��ʼԭ�� = 1 And b.״̬ = 1) Or (c.��ʼԭ�� in (3,15) And c.��ʼʱ�� Is Null And b.״̬ = 2)) "
        
        strSQLBaby = "Select q.����id,q.��ҳid,q.���,q.Ӥ������,q.Ӥ���Ա�,Round(Decode(q.����ʱ��, Null, Sysdate, q.����ʱ��) - q.����ʱ��)||'��' As ����" & _
            " From ������Ϣ A, ������ҳ B, ���˱䶯��¼ C, ���ű� D, �շ���ĿĿ¼ E,������������¼ Q" & vbNewLine & _
            " Where a.��Ժ = 1 And a.����id = b.����id And Nvl(b.��ҳid, 0) <> 0 And b.����id = c.����id And b.��ҳid = c.��ҳid And (C.����ID=[1] or C.����ID is null) And c.����id = d.Id  And b.����id=q.����ID And b.��ҳID=q.��ҳID " & vbNewLine & _
            "      And (d.վ��='" & gstrNodeNo & "' Or d.վ�� is Null)" & vbNewLine & _
            "      And b.����ȼ�id = e.Id(+) And Nvl(c.���Ӵ�λ, 0) = 0 And c.��ֹʱ�� Is Null" & vbNewLine & _
            "      And (c.��ʼԭ�� in(1,3) And Exists(Select 1 From �������Ҷ�Ӧ H Where c.����id = h.����id And h.����id = [1]) or c.��ʼԭ��=15 And c.����id = [1])" & vbNewLine & _
            "      And ((c.��ʼԭ�� = 1 And b.״̬ = 1) Or (c.��ʼԭ�� in (3,15) And c.��ʼʱ�� Is Null And b.״̬ = 2)) "
    End If
    '��Ժ����
    If tbcPati.Selected.Tag = "��Ժ" Then
        str�ֶ���� = ",first_value(Decode(Sign(h.�������-10),-1,h.�������,'')) " & _
            " Over(partition By h.����id,H.��ҳID Order By sign(h.�������-10),decode(h.��¼��Դ,4,0,h.��¼��Դ) desc,Decode(h.�������,1,1,2,2,3,3,0) DESC,h.��ϴ���) As ��ҽ���"
        If Sys.DeptHaveProperty(cboUnit.ItemData(cboUnit.ListIndex), "��ҽ��") Then
            str�ֶ���� = str�ֶ���� & ",first_value(Decode(Sign(h.�������-10),1,h.�������,'')) " & _
            " Over(partition By h.����id,H.��ҳID Order By sign(h.�������-10) desc,decode(h.��¼��Դ,4,0,h.��¼��Դ) desc,Decode(h.�������,11,1,12,2,13,3,0) DESC,h.��ϴ���) As ��ҽ���"
        Else
            str�ֶ���� = str�ֶ���� & ",null as ��ҽ���"
        End If
        strSQL = _
            "Select /*+ RULE */ Distinct Decode(B.״̬,3,4,3) as ����," & _
            " Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2," & _
            " Decode(B.״̬,3,'Ԥ��Ժ����','��Ժ����') as ����," & _
            " A.����ID,B.��ҳID,A.�����,B.סԺ��,NVL(b.����,a.����) ����, NVL(b.�Ա�,a.�Ա�) �Ա�, NVL(b.����,a.����) ����,C.���� as ����,B.��Ժ����ID ����ID,B.סԺҽʦ,B.����״̬," & _
            " LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,E.���� as ����ȼ�,B.�ѱ�,Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) as ��Ժ����,B.��Ժ����,B.��������," & IIf(mblnCardOrder, "f.˳���,", "d.���� as ��λ����,f.˳���,") & _
            " B.״̬,B.����,A.���￨��,b.���ۺ�,Nvl(b.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��)) as סԺ����,q.��� as ���������,b.������,B.Ӥ������ID,B.Ӥ������ID " & str�ֶ���� & _
            " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ E,������������¼ Q,��Ժ���� R,������ϼ�¼ H" & IIf(mblnCardOrder, ",��λ״����¼ F", ",��λ���Ʒ��� D,��λ״����¼ F") & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And Nvl(B.״̬,0)<>1 And H.����id(+)=b.����id And h.��ҳid(+)=b.��ҳid " & _
            " And b.����id=q.����ID(+) And b.��ҳID=q.��ҳID(+) And (q.���=1 Or q.��� is Null)" & IIf(mblnCardOrder, "and b.����id=f.����id(+)", " and b.����id=f.����id(+) and f.��λ����=D.����(+)") & _
            " And B.��Ժ����ID=C.ID And B.����ȼ�ID=E.ID(+) And (R.����ID=[1] Or b.Ӥ������ID=[1]) And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
            " And a.����ID=R.����ID And A.��ǰ����ID+0=R.����ID And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL" & strFilter
        
        strSQLBaby = "Select q.����id,q.��ҳid,q.���,q.Ӥ������,q.Ӥ���Ա�,Round(Decode(q.����ʱ��, Null, Sysdate, q.����ʱ��) - q.����ʱ��)||'��' As ����" & _
            " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ E,������������¼ Q,��Ժ���� R" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And Nvl(B.״̬,0)<>1" & _
            " And b.����id=q.����ID And b.��ҳID=q.��ҳID " & _
            " And B.��Ժ����ID=C.ID And B.����ȼ�ID=E.ID(+) And (R.����ID=[1] Or b.Ӥ������ID=[1]) And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
            " And a.����ID=R.����ID And A.��ǰ����ID+0=R.����ID And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL" & strFilter
        bln��λ���� = Not mblnCardOrder
    End If
    '��Ժ����:��Ժ���˿������ж��סԺ
    If tbcPati.Selected.Tag = "��Ժ" Then
        '����δ�����
        If chk����(0).Value = 1 And chk����(1).Value = 1 Then
            int���� = 0               '����ʾ
        ElseIf chk����(0).Value = 0 And chk����(1).Value = 1 Then
            int���� = 1               'ֻ��ʾδ�����
        ElseIf chk����(0).Value = 1 And chk����(1).Value = 0 Then
            int���� = 2              'ֻ��ʾ�ѽ����
        End If
        
        '�ж��Ƿ��ǲ���סԺ�ż�סԺ���Ƿ�Ϊ��
        If mstrFindType = "סԺ��" And Trim(PatiIdentify.Text) <> "" Then blnIsFind = True
        
        '����ȡ������ʾ��Ժ���˲��������Բ���ʱ��ʾ���ҵ���Ա��ʱ�䷶Χ����Ա
        If blnIsFind Then
            '����ȡ������ʾ��Ժ���˲��������Բ���ʱ��ʾ���ҵ���Ա��ʱ�䷶Χ����Ա
'            strTmpOut = " And (B.��Ժ���� Between to_date('" & Format(mdtOutBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And " & _
'                        " to_date('" & Format(mdtOutEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')  Or (b.סԺ��=[5] And B.��Ժ���� is Not Null)) "

            strTmpOut = " And (B.��Ժ���� Between to_date('" & Format(mdtOutBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And " & _
                        " to_date('" & Format(mdtOutEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')  " & _
                        IIf(int���� = 0, "", " And " & IIf(int���� = 1, "", "Not") & " Exists(Select 1 From ������� Where a.����id = ����id And ���� = 1 And Nvl(�������, 0) <> 0)") & _
                        " Or (b.סԺ��=[5] And B.��Ժ���� is Not Null)) "
        Else
            '�û����ǲ��ң���ʾ������Ӧʱ���ڵĲ���
            strTmpOut = " And B.��Ժ���� Between to_date('" & Format(mdtOutBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And to_date('" & Format(mdtOutEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS') "
            strTmpOut = strTmpOut & IIf(int���� = 0, "", " And " & IIf(int���� = 1, "", "Not") & " Exists(Select 1 From ������� Where a.����id = ����id And ���� = 1 And Nvl(�������, 0) <> 0)")
        End If
    
        strSQL = _
            "Select /*+ RULE */ Decode(B.��Ժ��ʽ,'����',6,5) as ����," & _
            " Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2," & _
            " Decode(B.��Ժ��ʽ,'����','��������','��Ժ����') as ����," & _
            " A.����ID,B.��ҳID,A.�����,B.סԺ��,NVL(b.����,a.����) ����, NVL(b.�Ա�,a.�Ա�) �Ա�, NVL(b.����,a.����) ����,C.���� as ����,B.��Ժ����ID ����ID,B.סԺҽʦ,B.����״̬," & _
            " LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,E.���� as ����ȼ�,B.�ѱ�,Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) as ��Ժ����,B.��Ժ����,B.��������," & _
            " B.״̬,B.����,A.���￨��,b.���ۺ�,Nvl(b.·��״̬,-1) ·��״̬,trunc(b.��Ժ����)-trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��)) as סԺ����,q.��� as ���������,b.������,B.Ӥ������ID,B.Ӥ������ID,'' AS ��ҽ���,'' AS ��ҽ��� ,B.���λ�ʿ " & _
            " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ E,������������¼ Q" & _
            " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.״̬=0" & _
            " And b.����id=q.����ID(+) And b.��ҳID=q.��ҳID(+) And (q.���=1 Or q.��� is Null)" & _
            " And B.��Ժ����ID=C.ID And B.����ȼ�ID=E.ID(+) And B.��ǰ����ID+0=[1] And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
            " And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL" & strTmpOut
        
        strSQLBaby = "Select q.����id,q.��ҳid,q.���,q.Ӥ������,q.Ӥ���Ա�,Round(Decode(q.����ʱ��, Null, Sysdate, q.����ʱ��) - q.����ʱ��)||'��' As ����" & _
            " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ E,������������¼ Q" & _
            " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.״̬=0" & _
            " And b.����id=q.����ID And b.��ҳID=q.��ҳID " & _
            " And B.��Ժ����ID=C.ID And B.����ȼ�ID=E.ID(+) And B.��ǰ����ID+0=[1] And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
            " And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL" & strTmpOut
    End If
    'ת������:��Ժ,ҽ���ʹ�����ʾ����ת��ǰ��
    If tbcPati.Selected.Tag = "ת��" Then
        strSQL = _
            "Select /*+ RULE */ Distinct 7 as ����,Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2,'ת������' as ����," & _
            " A.����ID,B.��ҳID,A.�����,B.סԺ��,NVL(b.����,a.����) ����, NVL(b.�Ա�,a.�Ա�) �Ա�, NVL(b.����,a.����) ����,D.���� as ����,C.����ID,C.����ҽʦ as סԺҽʦ,B.����״̬," & _
            " LPAD(c.����," & intBedLen & ",' ') as ����,E.���� as ����ȼ�,B.�ѱ�,Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) as ��Ժ����,B.��Ժ����,B.��������," & _
            " B.״̬,B.����,A.���￨��,b.���ۺ�,Nvl(b.·��״̬,-1) ·��״̬,trunc(Nvl(b.��Ժ����, Sysdate))-trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��)) as סԺ����,q.��� as ���������,b.������,B.Ӥ������ID,B.Ӥ������ID,'' AS ��ҽ���,'' AS ��ҽ��� " & _
            " From ������Ϣ A,������ҳ B,���˱䶯��¼ C,���ű� D,�շ���ĿĿ¼ E,������������¼ Q" & _
            " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.����ȼ�ID=E.ID(+)" & _
            " And b.����id=q.����ID(+) And b.��ҳID=q.��ҳID(+) And (q.���=1 Or q.��� is Null)" & _
            " And B.����ID=C.����ID And B.��ҳID=C.��ҳID" & _
            " And B.��ǰ����ID<>[1] And C.����ID+0=[1] And C.����ID=D.ID" & _
            " And Nvl(C.���Ӵ�λ,0)=0 And C.��ֹԭ�� In(3,15) " & strTmpDate & _
            " And Nvl(B.״̬,0)<>2 And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL "
            
        strSQLBaby = "Select q.����id,q.��ҳid,q.���,q.Ӥ������,q.Ӥ���Ա�,Round(Decode(q.����ʱ��, Null, Sysdate, q.����ʱ��) - q.����ʱ��)||'��' As ����" & _
            " From ������Ϣ A,������ҳ B,���˱䶯��¼ C,���ű� D,�շ���ĿĿ¼ E,������������¼ Q" & _
            " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.����ȼ�ID=E.ID(+)" & _
            " And b.����id=q.����ID And b.��ҳID=q.��ҳID " & _
            " And B.����ID=C.����ID And B.��ҳID=C.��ҳID" & _
            " And B.��ǰ����ID<>[1] And C.����ID+0=[1] And C.����ID=D.ID" & _
            " And Nvl(C.���Ӵ�λ,0)=0 And C.��ֹԭ�� In(3,15) " & strTmpDate & _
            " And Nvl(B.״̬,0)<>2 And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL "
    End If
    
    If strSQL = "" Then
        rptPati.Records.DeleteAll
        rptPati.Populate
        mblnUnRefresh = False
        Screen.MousePointer = 0
        LoadPatients = True
        Exit Function
    End If
    
    strSQL = strSQL & " Order by" & IIf(tbcPati.Selected.Tag = "��Ժ", " ˳���,", "") & IIf(bln��λ����, " ��λ����,", "") & " ����,����,����2,��ҳID Desc"
 
    Screen.MousePointer = 11
    On Error GoTo errH
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex), _
        mintChange, txt��������.Tag, strState, Val(Trim(PatiIdentify.Text)))
    If strSQLBaby <> "" Then
        Set rsBaby = zlDatabase.OpenSQLRecord(strSQLBaby, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex), _
            mintChange, txt��������.Tag, strState, Val(Trim(PatiIdentify.Text)))
    End If
      
    '������Ԥ����
    If tbcPati.Selected.Tag = "��Ժ" Then
        rptPati.Columns(col_��ҽ���).Visible = True
        rptPati.Columns(col_��ҽ���).Visible = Sys.DeptHaveProperty(cboUnit.ItemData(cboUnit.ListIndex), "��ҽ��")
    Else
        rptPati.Columns(col_��ҽ���).Visible = False
        rptPati.Columns(col_��ҽ���).Visible = False
    End If
        If tbcPati.Selected.Tag = "��Ժ" Then
        rptPati.Columns(col_���λ�ʿ).Visible = True
    Else
        rptPati.Columns(col_���λ�ʿ).Visible = False
    End If
    
    '��¼����ѡ�еĲ���
    If rptPati.SelectedRows.Count > 0 Then
        If Not rptPati.SelectedRows(0).GroupRow Then
            If rptPati.SelectedRows(0).Record.Tag <> "" Then
                lngPatiRow = rptPati.SelectedRows(0).Index '���ڿ������¶�λ
                strPatiRow = rptPati.SelectedRows(0).Record.Tag
            End If
        End If
    End If
    rptPati.Records.DeleteAll
    rptPati.Columns(col_���).Visible = False
    
    If mclsWardMonitor.Enabled And InStr(GetInsidePrivs(pסԺ��ʿվ), "����໤") > 0 Then
        strMonitor = mclsWardMonitor.GetListPati
    End If
    
    
    'ˢ�º�����Զ�չ��
    For i = 1 To rsPati.RecordCount
        '�����ύ��������Ӹ���
        If Nvl(rsPati!����״̬, 0) <> 0 Then
            rptPati.Columns(col_���).Visible = True
            
            '���������ͷ����仯ʱҪ���⿪һ����֧������ᵼ�·��鲻��
            If strPre�������� <> rsPati!���� & "" Then
                Set objParent = Nothing
                strPre�������� = rsPati!���� & ""
            End If
            
            If objParent Is Nothing Then
                Set objParent = Me.rptPati.Records.Add()
            ElseIf objParent.Tag <> CStr(rsPati!����״̬) Then
                Set objParent = Me.rptPati.Records.Add()
            End If
            If objParent.Tag <> CStr(rsPati!����״̬) Then
                objParent.Tag = CStr(rsPati!����״̬)
                objParent.Expanded = True
                For j = 0 To rptPati.Columns.Count - 1
                    If j = col_���� Then
                        Set objItem = objParent.AddItem(Val(rsPati!����))
                        objItem.Caption = rsPati!����
                    ElseIf j = col_��� Then
                        Set objItem = objParent.AddItem(Val(rsPati!����״̬))
                        objItem.Caption = " "
                    ElseIf j = col_���� Then
                        Set objItem = objParent.AddItem(CStr(Decode(rsPati!����״̬, 1, "�ȴ����", 2, "�ܾ����", 3, "�������", 4, "��鷴��")))
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
            
        '��Ӿ���Ĳ���������(���л������)
        If Not objParent Is Nothing Then
            Set objRecord = objParent.Childs.Add()
        Else
            Set objRecord = Me.rptPati.Records.Add()
        End If
        
        objRecord.Tag = CStr(rsPati!����ID & "," & rsPati!��ҳID) '���ڲ��˶�λ
        
        Set objItem = objRecord.AddItem(Val(rsPati!����)) '������Value��������
        objItem.Caption = rsPati!����
        
        Set objItem = objRecord.AddItem(Val(Decode(Nvl(rsPati!����״̬, 0), 0, 999, rsPati!����״̬)))
        objItem.Caption = " "
        If Nvl(rsPati!����״̬, 0) = 2 Then
            objRecord.PreviewText = "  ����:" & GetRefuseReason(rsPati!����ID, rsPati!��ҳID)
        End If
        
        'ͼ��:ע�����������Ǵ�0��ʼ��š�
        '     ͼ��Value���ڴ���Ƿ����ύ��飬����Ŷ�ȡ
        Set objItem = objRecord.AddItem(-1)
        objItem.Caption = " "
        If Nvl(rsPati!����״̬, 0) <> 0 Then
            objItem.Icon = Get����ͼ�����(rsPati!����״̬)
        ElseIf "" & rsPati!������ <> "" Then
            objItem.Icon = imgPati.ListImages("������").Index - 1
        End If
        
        '      lng·��״̬=-1:δ����,0-�����ϵ���������1-ִ���У�2-����������3-�������
        Set objItem = objRecord.AddItem(Val("" & rsPati!·��״̬))
        objItem.Caption = " "
        objItem.Icon = -1 + Choose(rsPati!·��״̬ + 2, imgPati.ListImages("δ����").Index, imgPati.ListImages("������").Index, _
            imgPati.ListImages("ִ����").Index, imgPati.ListImages("��������").Index, imgPati.ListImages("�������").Index)
            
        
        objRecord.AddItem Val(rsPati!����ID)
        objRecord.AddItem Val(rsPati!��ҳID)
        objRecord.AddItem CStr(Nvl(rsPati!����))
        
        If mblnOutDept Then
            Set objItem = objRecord.AddItem("" & rsPati!�����)
            objItem.Caption = Nvl(rsPati!�����, " ")
        Else
            Set objItem = objRecord.AddItem("" & rsPati!סԺ��)
            objItem.Caption = Nvl(rsPati!סԺ��, " ")
        End If
        
        
        Set objItem = objRecord.AddItem(zlStr.Lpad(Nvl(rsPati!����), 10)) 'Value��������
        objItem.Caption = CStr(Trim(Nvl(rsPati!����, " "))) 'Ϊ��ʱ�ᱻValue���
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
        
        objRecord.AddItem CStr(Nvl(rsPati!�Ա�))
        objRecord.AddItem CStr(Nvl(rsPati!����))
        objRecord.AddItem CStr(Nvl(rsPati!����ȼ�))
        objRecord.AddItem CStr(Nvl(rsPati!�ѱ�))
        objRecord.AddItem CStr(Nvl(rsPati!����))
        objRecord.AddItem CStr(Nvl(rsPati!סԺҽʦ))
        objRecord.AddItem Format(rsPati!��Ժ����, "yyyy-MM-dd HH:mm")
        objRecord.AddItem Format(Nvl(rsPati!��Ժ����), "yyyy-MM-dd HH:mm")
        objRecord.AddItem CStr(Nvl(rsPati!��������))
        objRecord.AddItem CStr(Nvl(rsPati!���￨��))
        objRecord.AddItem Val("" & rsPati!����ID)
        objRecord.AddItem Val(Trim(IIf(CStr("" & rsPati!סԺ����) = "0", "1", CStr("" & rsPati!סԺ����))))     '����ס����Ϊ��
        objRecord.AddItem "" & rsPati!������
        objRecord.AddItem Val("" & rsPati!Ӥ������ID)
        objRecord.AddItem Val("" & rsPati!Ӥ������ID)
        '������
        objRecord.AddItem CStr(Nvl(rsPati!��ҽ���))
        objRecord.AddItem CStr(Nvl(rsPati!��ҽ���))
        
        If tbcPati.Selected.Tag = "��Ժ" Then
            objRecord.AddItem rsPati!���λ�ʿ & ""
        Else
            '���������
            objRecord.AddItem ""
        End If
        
        If bln��λ���� Then
            objRecord.AddItem rsPati!��λ���� & ""
        Else
            objRecord.AddItem ""
        End If
        '���ۺ�
        objRecord.AddItem "" & rsPati!���ۺ�
        
        If tbcPati.Selected.Tag = "��Ժ" Then
            Set objItem = objRecord.AddItem("" & Nvl(rsPati!˳���)) 'Value��������
            objItem.Caption = Nvl(rsPati!˳���, " ") 'Ϊ��ʱ�ᱻValue���
        Else
            objRecord.AddItem ""
        End If
        '��ʾ������ɫ
        objRecord.Item(col_����).ForeColor = zlDatabase.GetPatiColor(Nvl(rsPati!��������))
        For j = 0 To rptPati.Columns.Count - 1
            If j <> col_���� And j <> col_��� And j <> col_ͼ�� And j <> col_˳��� Then
                objRecord.Item(j).ForeColor = objRecord.Item(col_����).ForeColor
            End If
        Next
        
        'ͳ�Ʋ�����Ŀ
        lngCount(Val(rsPati!����)) = lngCount(Val(rsPati!����)) + 1
        '�����Ƿ���Ӥ�����Ӥ����
        If Not rsBaby Is Nothing Then
            Set objBabyParent = objRecord
            rsBaby.Filter = "����ID=" & objBabyParent(col_����Id).Value & " and ��ҳID=" & objBabyParent(col_��ҳID).Value
            If Not rsBaby.EOF Then
                rptPati.Columns(col_���).Visible = True
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
        rsPati.MoveNext
    Next
    If mblnOutDept Then
        rptPati.Columns.Find(col_סԺ��).Caption = "�����"
    Else
        rptPati.Columns.Find(col_סԺ��).Caption = "סԺ��"
    End If
    rptPati.Populate
    '���ݽ���ҽԺ������벡����Ŀͳ����Ϣ
    strState = "�� " & rsPati.RecordCount & " ������"
    For i = LBound(lngCount) To UBound(lngCount)
        If lngCount(i) > 0 Then
            Select Case i
            Case 0
                strState = strState & "����Ժ����ס:"
            Case 1
                strState = strState & "��ת�ƴ���ס:"
            Case 2
                strState = strState & "��ת��������ס:"
            Case 3
                strState = strState & "����Ժ:"
            Case 4
                strState = strState & "��Ԥ��Ժ:"
            Case 5
                strState = strState & "����Ժ:"
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
    
    '��λ������:��Populate֮��
    mstrPrePati = ""
    If rptPati.Rows.Count = 0 Or rsPati.RecordCount > 1 And lngPatiRow = 0 Then
        Call ClearPatiInfo
        '��������ˢ���Ӵ���
        Call SubWinRefreshData(tbcSub.Selected)
        
        If tbcPati.Selected.Tag = "��Ժ" And mblnNoRefNotify = False Then
            mstrPreNotify = ""
            rptNotify.Records.DeleteAll
            rptNotify.Populate
            rptNotify.TabStop = False
        End If
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
    mlng����ID = 0
    mlng��ҳID = 0
    
    mPatiInfo.״̬ = 0
    mPatiInfo.Ӥ�� = -1
    mPatiInfo.סԺ�� = ""
    mPatiInfo.���� = ""
    mPatiInfo.��ҳID = 0
    mPatiInfo.�����ҳID = 0
    mPatiInfo.����ID = 0
    mPatiInfo.����ID = 0
    mPatiInfo.��Ժ���� = CDate(0)
    mPatiInfo.��Ժ���� = CDate(0)
    mPatiInfo.סԺ���� = 0
    mPatiInfo.����ת�� = False
    mPatiInfo.���� = False
    mPatiInfo.���� = False
    mPatiInfo.���� = 0
    mPatiInfo.���� = 0
        
    cboPages.Clear
    cbo����.Clear
    
    lbl����(1).Caption = ""
    lbl����(1).Caption = ""
    lblPatiName(1).Caption = ""
    lblPatiName(1).ToolTipText = ""
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
    Dim strTmp As String
            
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
                If mstrFindType = "����" Then
                    If UCase(Trim(.Record(col_����).Value)) = UCase(PatiIdentify.Text) Then Exit For
                ElseIf mstrFindType = "סԺ��" Then
                    If .Record(col_סԺ��).Value = PatiIdentify.Text Then Exit For
                ElseIf mstrFindType = "���ۺ�" Then
                    If .Record(col_���ۺ�).Value = PatiIdentify.Text Then Exit For
                ElseIf mstrFindType = "���￨" Then
                    If UCase(.Record(col_���￨).Value) = UCase(PatiIdentify.Text) Then Exit For
                ElseIf mstrFindType = "����" Then
                    If .Record(col_����).Value Like "*" & PatiIdentify.Text & "*" Then Exit For
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
            If tbcPati.Selected.Tag = "��Ժ" And mstrFindType = "סԺ��" Then
                strTmp = GetNoPatiWhy(PatiIdentify.Text)
            Else
                strTmp = IIf(blnNext, "������", "") & "�Ҳ������������Ĳ��ˡ�"
            End If
            MsgBox strTmp, vbInformation, gstrSysName
        End If
    End If
End Sub

Function ExecuteMonitor() As Boolean
'���ܣ����ü໤��
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
'���ܣ���ȡ������鷴��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngCount As Long
    Dim curDate As Date
    
    If cboUnit.ListIndex = -1 Then
        fra���.Visible = False: LoadResponse = True: Exit Function
    End If
    
    On Error GoTo errH
    curDate = zlDatabase.Currentdate
    Screen.MousePointer = 11
    
    '��ȡ��ǰ��������Ժ����Ժ���ˣ���"����������¼"Ϊ׼ȫ��ɨ��
    strSQL = "Select Count(*) as ���� From ������ҳ B,����������¼ A" & _
        " Where A.����ID=B.����ID and A.��ҳID=B.��ҳID And A.��¼״̬=1" & _
        " And A.�������� IN(3,4) And B.��ǰ����ID + 0 =[1]" & _
        " And a.����ʱ�� Between [2] And [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LoadResponse", cboUnit.ItemData(cboUnit.ListIndex), CDate(Format(curDate - mlngMedRedDay, "yyyy-MM-dd")), CDate(Format(curDate, "yyyy-MM-dd HH:mm:ss")))
    If Not rsTmp.EOF Then lngCount = Nvl(rsTmp!����, 0)
    
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
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Long
    Dim blnTmp As Boolean
    
    mstrPreNotify = ""
    rptNotify.Records.DeleteAll
    
    If cboUnit.ListIndex = -1 Then LoadNotify = True: Exit Function
        
    Screen.MousePointer = 11
    On Error GoTo errH
    blnTmp = mobjKernel.GetAdviceRemind(rsTmp, cboUnit.ItemData(cboUnit.ListIndex), IIf(optNotify(0).Value = True, UserInfo.����, ""))
    Screen.MousePointer = 0
    If blnTmp = False Then Exit Function
    If rsTmp Is Nothing Then Exit Function
    If rsTmp.State = adStateClosed Then Exit Function
    strTmp = ","
    For i = 1 To rsTmp.RecordCount
        Select Case rsTmp!���ͱ��� & ""
        Case "ZLHIS_PACS_006", "ZLHIS_PACS_007"
            'ZLHIS_PACS_006 ZLHIS_PACS_007 ��ϢΪһ����Ϣ����ҽ����ĿΪ��λ����ʾһ��
            If InStr(strTmp, "," & rsTmp!���ͱ��� & "," & rsTmp!ҵ���ʶ & ",") = 0 Then
                strTmp = strTmp & rsTmp!���ͱ��� & "," & rsTmp!ҵ���ʶ & ","
                Call AddReportRow(rsTmp!����ID & "," & rsTmp!��ҳID, rsTmp!����ID, rsTmp!��ҳID, Nvl(rsTmp!����), Nvl(rsTmp!סԺ��), Nvl(rsTmp!����), Nvl(rsTmp!��Ϣ����), _
                    rsTmp!���ͱ��� & "", rsTmp!���ȳ̶� & "", Format(rsTmp!�Ǽ�ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), rsTmp!ҵ���ʶ & "", rsTmp!������Դ & "", Nvl(rsTmp!����, 0), Nvl(rsTmp!���ﲡ��id, 0))
            End If
        Case Else
            If InStr(strTmp, "," & rsTmp!����ID & "," & rsTmp!��ҳID & "," & rsTmp!���ͱ��� & ",") = 0 Then
                strTmp = strTmp & rsTmp!����ID & "," & rsTmp!��ҳID & "," & rsTmp!���ͱ��� & ","
                Call AddReportRow(rsTmp!����ID & "," & rsTmp!��ҳID, rsTmp!����ID, rsTmp!��ҳID, Nvl(rsTmp!����), Nvl(rsTmp!סԺ��), Nvl(rsTmp!����), Nvl(rsTmp!��Ϣ����), _
                    rsTmp!���ͱ��� & "", rsTmp!���ȳ̶� & "", Format(rsTmp!�Ǽ�ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), rsTmp!ҵ���ʶ & "", rsTmp!������Դ & "", Nvl(rsTmp!����, 0), Nvl(rsTmp!���ﲡ��id, 0))
            End If
        End Select
        rsTmp.MoveNext
    Next
    rptNotify.Populate 'ȱʡ��ѡ���κ���
    rptNotify.TabStop = rptNotify.Rows.Count > 0
    Screen.MousePointer = 0
    LoadNotify = True
    If mbln��Ϣ���� Then
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
'���ܣ����в�����ҳ����
'������blnEditable=�Ƿ�����༭(��Ȩ�޼�ǩ������������)
    Dim blnReadOnly As Boolean
    
    If mlng����ID = 0 Then Exit Sub
    
    If mPatiInfo.����ת�� Then
        MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '������Ŀ֮�󲻿�������
    If Not (CheckMecRed(mlng����ID, mlng��ҳID, Me.Caption) Or blnEditable) Then
        blnReadOnly = True
    End If
    If mclsInOutMedRec Is Nothing Then
        Set mclsInOutMedRec = New zlMedRecPage.clsInOutMedRec
        Call mclsInOutMedRec.InitMedRec(gcnOracle, glngSys, pסԺ��ʿվ, mclsMipModule, gobjCommunity, gclsInsure)
    End If
    '��ģ̬��ʾ��ҳ����
    If Not mclsInOutMedRec.IsOpen Then
        If mclsInOutMedRec.ShowInMedRecEdit(Me, mlng����ID, mPatiInfo.��ҳID, mPatiInfo.����ID, rptPati.SelectedRows(0).Record(col_·��״̬).Value, , mstrPrivs, IIf(blnReadOnly, 1, 0), False) Then
            mstrPrePati = "": Call rptPati_SelectionChanged
        End If
    End If
End Sub

Private Sub timNotify_Timer()
    Static strPreTime1 As String
    Static strPreTime2 As String
    Dim curTime As Date
    
    curTime = Now
    If gbln����Ӱ����ϢϵͳԤԼ Then
        If strPreTime2 = "" Then
            strPreTime2 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
        End If
        If DateDiff("s", CDate(strPreTime2), curTime) > 300 Then
            strPreTime2 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
            If mobjKernel.GetMsgRISReady(mlng����ID) Then
                Call LoadNotify
            End If
        End If
    End If
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

    'ˢ�²����������
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

Private Sub Set������Ŀ��������()
     On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "���ƻ�������(ZLCISBase)û����ȷ��װ���ù����޷�ִ�С�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    Call gobjCISBase.CallSetClinicCharge(mlng����ID, 1, Me, gcnOracle, glngSys, gstrDBUser, EסԺ����, InStr(GetInsidePrivs(pסԺ��ʿվ), ";������Ŀ��������;") = 0)
End Sub

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
        Call SetpicPatiPosition
        Call SetPatiInfoCtlPos
    End If
    Select Case tbcSub.Selected.Tag
        Case "·��"
            Call mclsPath.SetFontSize(mbytSize)
        Case "ҽ��"
            Call mclsAdvices.SetFontSize(mbytSize)
        Case "����"
            Call mclsFeeQuery.SetFontSize(mbytSize)
        Case "����"
            Call mclsEPRs.SetFontSize(mbytSize)
        Case "����"
            Call mclsTends.SetFontSize(mbytSize)
                Case "�²���"
            On Error Resume Next
            Call mclsEMR.SetFontSize(mbytSize)
            err.Clear: On Error GoTo 0
        End Select
End Sub
Private Sub SetpicPatiPosition()
'���ܣ������б��Լ�ɸѡ��������ؿؼ���λ�����С����

    Dim i As Long
    Dim lngDistance As Long
    
    Call picPatiIn_Resize
    Call zlControl.SetPubCtrlPos(False, 0, lblDept, 10, cboUnit)
    cboUnit.Width = picPatiFilter.ScaleWidth - cboUnit.Left
    
    Call zlControl.SetPubCtrlPos(False, 0, lbl��������, 10, txt��������)
    txt��������.Width = cboUnit.Width
    cmd��������.Width = 270
    cmd��������.Height = txt��������.Height - IIf(mbytSize = 0, 45, 60)
    cmd��������.Top = txt��������.Top + 30
    cmd��������.Left = txt��������.Left + txt��������.Width - cmd��������.Width - 30
    
    'checkBoxѡ�������������ʱ����ı��ȣ������Ҫ��ȥ100
    lngDistance = IIf(mbytSize = 0, 10, -50)
    Call zlControl.SetPubCtrlPos(False, 0, lbl��������, 10, chk��������(0), lngDistance, chk��������(1), lngDistance, chk��������(2))
    Call zlControl.SetPubCtrlPos(False, 0, lbl��Ժʱ��, 50, cboSelectTime)
    chk����(0).Left = cboSelectTime.Left + 0.5 * cboSelectTime.Width
    Call zlControl.SetPubCtrlPos(False, 0, chk����(0), 20, chk����(1))
    Call zlControl.SetPubCtrlPos(False, 0, lblת��, 50, cmdRef)

    txtChange.Left = lblת��.Left + Me.TextWidth("��ʾ��� ")
    fraChange.Left = txtChange.Left
    fraChange.Top = txtChange.Top + txtChange.Height
        
End Sub

Private Sub SetPatiInfoCtlPos()
'���ܣ��Բ��˵���ϸ��Ϣ����Ŀؼ�λ�õ�����������picInfo�еĿؼ�
    Dim lngDistance1 As Long, lngDistance2 As Long
    
    Dim lngTmp As Long
    
    lngDistance2 = 180: lngDistance1 = 10
    lngTmp = IIf(mbytSize = 0, 1080, 1270)
    
    lblPatiName(0).Top = IIf(mbytSize = 0, 190, 210)
    lbl����(0).Top = lblPatiName(0).Top
    
    '1.סԺ����
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

Private Function GetDataToUnits(Optional ByVal strIn As String = "") As ADODB.Recordset
'���ܣ���ȡ�����б����ݼ�¼��
'������strIn ��������
    Dim strSQL As String
    Dim blnYN As Boolean
    
    If strIn <> "" Then blnYN = True
    If InStr(mstrPrivs, "ȫԺ����") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
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
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            IIf(blnYN, " And (A.���� Like [2] Or A.���� Like [3] Or A.���� Like [3])", "") & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.����,C.����,Nvl(B.ȱʡ,0) as ȱʡ" & _
            " From �������Ҷ�Ӧ A,������Ա B,���ű� C" & _
            " Where A.����ID=C.ID And B.����ID=A.����ID And B.��ԱID=[1]" & _
            " And Exists(Select 1 From ��������˵�� Where ��������='�ٴ�' And ����ID=A.����ID)" & _
            " And Not Exists(Select 1 From ��������˵�� Where ��������='����' And ����ID=A.����ID)" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            IIf(blnYN, " And (C.���� Like [2] Or C.���� Like [3] Or C.���� Like [3])", "") & _
            " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = "Select ID,����,����,Max(ȱʡ) as ȱʡ From (" & strSQL & ") Group by ID,����,���� Order by ����"
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
'���ܣ����Ӥ����ĸ���Ƿ���룬�е�ǰ��Ӥ������,�������Ժҳ����ж�
    If rptPati.SelectedRows(0).GroupRow = False And tbcPati.Selected.Tag = "��Ժ" Then
        If rptPati.SelectedRows(0).Record(COL_Ӥ������ID).Value <> 0 Then
            If rptPati.SelectedRows(0).Record(COL_Ӥ������ID).Value = cboUnit.ItemData(cboUnit.ListIndex) Or rptPati.SelectedRows(0).Record(COL_Ӥ������ID).Value = cboUnit.ItemData(cboUnit.ListIndex) Then
                MsgBox "�ò����Ѿ�ת���������ˣ�ֻ��Ӥ�����ڱ����ң�������������ˡ�", vbInformation, Me.Caption
                CheckBabyInOut = True
            End If
        End If
    End If
End Function

Private Sub AddMsgToLis(ByVal rsMsg As ADODB.Recordset)
'���ܣ������յ�����Ϣ���������б���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim i As Long
    Dim blnAdd As Boolean '�Ƿ�Ҫ���һ��
    
    On Error GoTo errH
    
    If Mid(rsMsg!���ѳ���, 3, 1) <> "1" Then Exit Sub
    
    If InStr("," & rsMsg!����IDs & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") > 0 Or _
        InStr("," & rsMsg!������Ա & ",", "," & UserInfo.���� & ",") > 0 Then
        
        '�ж��б��Ƿ��Ѿ���������Ϣ�ˣ����� AddReportRow ���жϣ��������ܻ����һ��SQL��ѯ
        For i = 0 To rptNotify.Rows.Count - 1
            If Not rptNotify.Rows(i).GroupRow Then
                If rptNotify.Rows(i).Record(C_��Ϣ).Value = rsMsg!���ͱ��� And rptNotify.Rows(i).Record.Tag = CStr(rsMsg!����ID & "," & rsMsg!����id) Then
                    blnAdd = True
                End If
            End If
        Next
        
        '����ǽ������¿�����ͣҽ��Ҫ�滻��ǰ����Ϣ
        If blnAdd Then
            If InStr(",ZLHIS_CIS_001,ZLHIS_CIS_002,", "," & rsMsg!���ͱ��� & ",") > 0 And Val(rsMsg!���ȳ̶� & "") = 2 Then
                blnAdd = False
                rptNotify.Records.RemoveAt i
            End If
        End If
        
        If blnAdd Then Exit Sub
        
        strSQL = "Select a.סԺ��, a.����, a.�Ա�, a.����, a.��ǰ���� As ����, a.���� From ������Ϣ A Where a.����id =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMsg!����ID))
        
        Call AddReportRow(rsMsg!����ID & "," & rsMsg!����id, rsMsg!����ID, rsMsg!����id, Nvl(rsTmp!����), Nvl(rsTmp!סԺ��), Nvl(rsTmp!����), Nvl(rsMsg!��Ϣ����), _
             rsMsg!���ͱ��� & "", rsMsg!���ȳ̶� & "", Format(rsMsg!�Ǽ�ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), rsMsg!ҵ���ʶ & "", rsMsg!������Դ & "", Nvl(rsTmp!����, 0))
        
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
    objRecord.Tag = arrInput(Index): Index = Index + 1         'Tagֵ
    Set objItem = objRecord.AddItem(""): objItem.Icon = 1
    Set objItemIcon = objItem
    
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '����id
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '����id
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1 '����
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index))) 'סԺ��
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index))) '����
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))     '״̬������
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    strNO = arrInput(Index)                            '��Ϣ���
    objRecord.AddItem strNO: Index = Index + 1
    
    int���ȼ� = Val(arrInput(Index))                     '���
    objRecord.AddItem int���ȼ�: Index = Index + 1
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  '����
    
    strҵ�� = arrInput(Index): Index = Index + 1              'ҵ���ʶ
    str������Դ = arrInput(Index): Index = Index + 1          '������Դ
    int���� = arrInput(Index)
    
    If InStr(",ZLHIS_PACS_005,ZLHIS_LIS_003,", "," & strNO & ",") > 0 Then 'Σ��ֵ��Ϣ���⴦���Ķ�ʱ������Ϣ
        objRecord.AddItem strҵ�� & "," & Val(str������Դ)
    Else
        objRecord.AddItem strҵ��
    End If
    
    Index = Index + 1
    objRecord.AddItem Val(arrInput(Index))    '����ID
    
    If (strNO = "ZLHIS_CIS_001" Or strNO = "ZLHIS_CIS_002") And int���ȼ� = 2 Then objItemIcon.Icon = 18
    
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

Private Sub ReadAndSendMsg(ByVal strNO As String, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng���ﲡ��ID As Long)
    '���ܣ��¿���Ϣʱ������Ϣ�Ĳ����Ѿ����ٵ�ǰ���������Ƚ���Ϣ��Ϊ�Ѷ��������·�����Ϣ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim arrSQL() As String
    Dim lng��ǰ����ID As Long
    Dim lng��ǰ����ID As Long
    Dim blnTrans As Boolean
    
    On Error GoTo errH
    
    If strNO <> "ZLHIS_CIS_001" Then Exit Sub
    
    strSQL = "select nvl(A.��ǰ����ID,0) as ��ǰ����ID, nvl(A.��ǰ����ID,0) as ��ǰ����ID from ������Ϣ A where A.����ID = [1] and ��ҳID = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)

    If rsTmp.EOF Then Exit Sub
    
    lng��ǰ����ID = Val(rsTmp!��ǰ����id)
    lng��ǰ����ID = Val(rsTmp!��ǰ����ID)
    
    If lng���ﲡ��ID <> lng��ǰ����ID And lng��ǰ����ID <> 0 Then
        
        If Not HaveOperateAdvice(lng����ID, lng��ҳID, 0) Then
            '������ϢΪ�Ѷ�
            strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng��ҳID & ",'" & strNO & "','0010','" & _
            UserInfo.���� & "'," & cboUnit.ItemData(cboUnit.ListIndex) & ")"
            gcnOracle.BeginTrans: blnTrans = True
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            gcnOracle.CommitTrans: blnTrans = False
        Else
            strSQL = "select A.��Ϣ����, A.���ѳ���,A.���ͱ���,A.ҵ���ʶ,A.���ȳ̶� From ҵ����Ϣ�嵥 A Where a.����id=[1] And a.����id=[2] And a.���ͱ��� =[3] and a.���ﲡ��ID =[4]  And a.�Ƿ�����=0 And Rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID, strNO, lng���ﲡ��ID)
            If rsTmp.RecordCount > 0 Then
                For i = 0 To rsTmp.RecordCount - 1
                    ReDim Preserve arrSQL(i)
                    arrSQL(UBound(arrSQL)) = "Zl_ҵ����Ϣ�嵥_Insert(" & lng����ID & "," & lng��ҳID & "," & lng��ǰ����ID & "," & lng��ǰ����ID & ",2,'" & rsTmp!��Ϣ���� & "','" & rsTmp!���ѳ��� & "','" & rsTmp!���ͱ��� & "','" & rsTmp!ҵ���ʶ & "'," & rsTmp!���ȳ̶� & ",0,null," & lng��ǰ����ID & ",null)"
                Next
            End If
            
            '������ϢΪ�Ѷ�
            strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng��ҳID & ",'" & strNO & "','0010','" & _
            UserInfo.���� & "'," & cboUnit.ItemData(cboUnit.ListIndex) & ")"
            
            gcnOracle.BeginTrans: blnTrans = True
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            '���·�����Ϣ
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

Private Sub mclsAdvices_DoByAdvice(ByVal lngҽ��ID As Long, ByVal lng���ID As Long, ByVal lngWayID As Long, ByVal strTag As String)
'���ܣ���ҽ������  lngWayID��conMenu_Edit_AdvicePrice
    Dim lngTmp As Long, lng����ID As Long
    lngTmp = IIf(lng���ID = 0, lngҽ��ID, lng���ID)
    lng����ID = Val("" & rptPati.SelectedRows(0).Record(col_����ID).Value)
    Call mclsFeeQuery.zlPatiBilling(Me, mlng����ID, mlng����ID, mlng��ҳID, lng����ID, False, lngTmp)
End Sub

Private Function GetNoPatiWhy(ByVal strסԺ�� As String) As String
'���ܣ�������Ժ����δ��ʾ��ԭ�򣬸�סԺ�Ź��˳�Ժ����
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Long
    Dim strMsg As String
    Dim str������� As String
    
    On Error GoTo errH
    
    strSQL = "Select b.����id,b.��ҳid,b.����,b.��Ժ����,b.��ǰ����id,b.���ʱ��,Nvl(b.����״̬,0) as ����״̬,c.����||'-'||c.���� as ����" & vbNewLine & _
        "From ������Ϣ A, ������ҳ B,���ű� C Where a.����id = b.����id and b.��ǰ����id=c.id And b.סԺ�� =[1] order by b.��ҳID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strסԺ��, cboUnit.ItemData(cboUnit.ListIndex))
    
    If rsTmp.EOF Then
        strMsg = "���ݵ�ǰ�����סԺ��δ�ҵ��κβ��ˣ�ԭ��������һ�������סԺ�š�"
    Else
        strMsg = "���ݵ�ǰ�����סԺ���ҵ����²���(ID:" & rsTmp!����ID & ")��δ��ʾԭ��" & vbCrLf
        For i = 1 To rsTmp.RecordCount
            strTmp = "����""" & rsTmp!���� & """��" & rsTmp!��ҳID & "��סԺ""" & rsTmp!���� & """"
            If IsNull(rsTmp!��Ժ����) Then
                strTmp = strTmp & "δ��Ժ����ǰ������סԺ״̬��"
            ElseIf Val(rsTmp!��ǰ����ID & "") <> cboUnit.ItemData(cboUnit.ListIndex) Then
                strTmp = strTmp & "�ѳ�Ժ�������ڵ�ǰ������"
            ElseIf Not IsNull(rsTmp!���ʱ��) Or Val(rsTmp!����״̬ & "") = 5 Then
                strTmp = strTmp & "�ѳ�Ժ�������ѷ��鵵��"
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
