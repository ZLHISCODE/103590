VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "ZLIDKIND.OCX"
Begin VB.Form frmLabSampleRegister 
   Caption         =   "����걾�Ǽ�"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12285
   Icon            =   "frmLabSampleRegister.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   12285
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   855
      Index           =   0
      Left            =   2970
      TabIndex        =   38
      Top             =   5250
      Width           =   1245
      _Version        =   589884
      _ExtentX        =   2196
      _ExtentY        =   1508
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   855
      Index           =   1
      Left            =   1770
      TabIndex        =   39
      Top             =   5220
      Width           =   1245
      _Version        =   589884
      _ExtentX        =   2196
      _ExtentY        =   1508
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   855
      Index           =   2
      Left            =   2970
      TabIndex        =   40
      Top             =   6240
      Width           =   1245
      _Version        =   589884
      _ExtentX        =   2196
      _ExtentY        =   1508
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   855
      Index           =   3
      Left            =   1860
      TabIndex        =   41
      Top             =   6300
      Width           =   1245
      _Version        =   589884
      _ExtentX        =   2196
      _ExtentY        =   1508
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.Timer Timer 
      Interval        =   60000
      Left            =   11715
      Top             =   105
   End
   Begin VB.PictureBox picBarCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   9780
      ScaleHeight     =   585
      ScaleWidth      =   1125
      TabIndex        =   43
      Top             =   480
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox PicCuvetteCount 
      BorderStyle     =   0  'None
      Height          =   3105
      Left            =   4710
      ScaleHeight     =   3105
      ScaleWidth      =   2805
      TabIndex        =   35
      Top             =   4890
      Width           =   2805
      Begin XtremeReportControl.ReportControl rptCuvetteCount 
         Height          =   1125
         Left            =   300
         TabIndex        =   36
         Top             =   1230
         Width           =   2115
         _Version        =   589884
         _ExtentX        =   3731
         _ExtentY        =   1984
         _StockProps     =   0
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         EditOnClick     =   0   'False
      End
      Begin XtremeSuiteControls.ShortcutCaption srtCuvetteCount 
         Height          =   285
         Left            =   150
         TabIndex        =   37
         Top             =   240
         Width           =   2235
         _Version        =   589884
         _ExtentX        =   3942
         _ExtentY        =   503
         _StockProps     =   6
         Caption         =   "��ǰ�ѵǼ��Թ�"
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
   Begin VB.PictureBox PicCount 
      BorderStyle     =   0  'None
      Height          =   2835
      Left            =   7755
      ScaleHeight     =   2835
      ScaleWidth      =   3885
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4845
      Width           =   3885
      Begin XtremeReportControl.ReportControl rptCount 
         Height          =   1125
         Left            =   90
         TabIndex        =   33
         Top             =   450
         Width           =   2115
         _Version        =   589884
         _ExtentX        =   3731
         _ExtentY        =   1984
         _StockProps     =   0
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin XtremeSuiteControls.ShortcutCaption SrtCount 
         Height          =   285
         Left            =   150
         TabIndex        =   34
         Top             =   60
         Width           =   2235
         _Version        =   589884
         _ExtentX        =   3942
         _ExtentY        =   503
         _StockProps     =   6
         Caption         =   "��ǰ�ѵǼǵ�ҽ��"
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
      Height          =   4770
      Left            =   60
      ScaleHeight     =   4770
      ScaleWidth      =   7995
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   30
      Width           =   7995
      Begin VB.CheckBox chkDept 
         BackColor       =   &H00FDD6C6&
         Caption         =   "��ִ�п��ҹ���"
         Height          =   255
         Left            =   3720
         TabIndex        =   47
         ToolTipText     =   "���Ǽǵı걾�Ǽ���걾�������ʾ"
         Top             =   4455
         Width           =   1635
      End
      Begin VB.CheckBox chkUrgent 
         BackColor       =   &H00FDD6C6&
         Caption         =   "�����걾��ʾ"
         Height          =   255
         Left            =   6075
         TabIndex        =   46
         ToolTipText     =   "���Ǽǵı걾�Ǽ���걾�������ʾ"
         Top             =   4215
         Width           =   1515
      End
      Begin VB.CheckBox ChkBarCodeRegister 
         BackColor       =   &H00FDD6C6&
         Caption         =   "ɨ�������ֱ�ӵǼ�"
         Height          =   255
         Left            =   3720
         TabIndex        =   45
         Top             =   4215
         Width           =   1965
      End
      Begin VB.CheckBox chkComRequest 
         BackColor       =   &H00FDD6C6&
         Caption         =   "����걾�ͼ�󷽿ɽ��м���걾ǩ��"
         Height          =   225
         Left            =   60
         TabIndex        =   44
         Top             =   4215
         Width           =   3390
      End
      Begin zlIDKind.IDKind IDKind 
         Height          =   405
         Left            =   120
         TabIndex        =   28
         Top             =   300
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   714
         IDKindStr       =   "��|����|0;ҽ|ҽ����|1;��|���֤��|2;IC|IC����|3;��|�����|4;��|���￨|5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FDD6C6&
         Caption         =   "������Ϣ"
         Height          =   2385
         Left            =   60
         TabIndex        =   16
         Top             =   1800
         Width           =   7875
         Begin XtremeReportControl.ReportControl rptCuvette 
            Height          =   1635
            Left            =   120
            TabIndex        =   22
            Top             =   630
            Width           =   7605
            _Version        =   589884
            _ExtentX        =   13414
            _ExtentY        =   2884
            _StockProps     =   0
            AllowColumnRemove=   0   'False
            MultipleSelection=   0   'False
            SkipGroupsFocus =   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
         Begin VB.TextBox txt�ͼ��� 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5700
            MaxLength       =   20
            TabIndex        =   27
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "�Ǽ�(&G)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6765
            TabIndex        =   21
            Top             =   225
            Width           =   1065
         End
         Begin VB.TextBox txt����ʱ�� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1095
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   20
            Top             =   240
            Width           =   2010
         End
         Begin VB.TextBox txt������ 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3870
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   18
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ͼ���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   8
            Left            =   4950
            TabIndex        =   26
            Top             =   300
            Width           =   720
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   19
            Top             =   300
            Width           =   960
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   3150
            TabIndex        =   17
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.TextBox txtGoto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   780
         TabIndex        =   14
         ToolTipText     =   "����Ϊ���롢��������ͷΪ����ID��������סԺ�š���*������š���.���Һŵ��š���/���շѵ��ݺ�"
         Top             =   315
         Width           =   7140
      End
      Begin VB.Frame FraPatientInfo 
         BackColor       =   &H00FDD6C6&
         Caption         =   "������Ϣ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   60
         TabIndex        =   3
         Top             =   720
         Width           =   7875
         Begin VB.TextBox txt���� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5010
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   24
            Top             =   210
            Width           =   1155
         End
         Begin VB.TextBox txt�Ա� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   23
            Top             =   210
            Width           =   1095
         End
         Begin VB.TextBox txt���� 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   870
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   7
            Top             =   210
            Width           =   1635
         End
         Begin VB.TextBox txtBed 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6825
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   210
            Width           =   975
         End
         Begin VB.TextBox txtPatientDept 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   600
            Width           =   2445
         End
         Begin VB.TextBox txtID 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   870
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   600
            Width           =   1635
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   4515
            TabIndex        =   13
            Top             =   255
            Width           =   480
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ա�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   2730
            TabIndex        =   12
            Top             =   255
            Width           =   480
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ڿ���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   2730
            TabIndex        =   11
            Top             =   645
            Width           =   960
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   6330
            TabIndex        =   10
            Top             =   255
            Width           =   480
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ʶ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   645
            Width           =   720
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��  ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   255
            Width           =   720
         End
      End
      Begin VB.CheckBox chkRemberPer 
         BackColor       =   &H00FDD6C6&
         Caption         =   "�Ǽ��ͼ��˺����"
         Height          =   255
         Left            =   90
         TabIndex        =   48
         ToolTipText     =   "���Ǽǵı걾�Ǽ���걾�������ʾ"
         Top             =   4470
         Width           =   1935
      End
      Begin VB.Label lbl��ʾ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   1170
         TabIndex        =   29
         Top             =   30
         Width           =   60
      End
      Begin VB.Label lblGoto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���˲���"
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
         TabIndex        =   15
         Top             =   30
         Width           =   900
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   2835
      Left            =   8250
      ScaleHeight     =   2835
      ScaleWidth      =   3885
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1470
      Width           =   3885
      Begin XtremeReportControl.ReportControl rptPlist 
         Height          =   1125
         Left            =   780
         TabIndex        =   30
         Top             =   630
         Width           =   2115
         _Version        =   589884
         _ExtentX        =   3731
         _ExtentY        =   1984
         _StockProps     =   0
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin XtremeSuiteControls.ShortcutCaption srtPatient 
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   60
         Width           =   1455
         _Version        =   589884
         _ExtentX        =   2566
         _ExtentY        =   503
         _StockProps     =   6
         Caption         =   "������Ϣ"
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   8025
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLabSampleRegister.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16589
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin MSComctlLib.ImageList ImgList 
      Left            =   8340
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleRegister.frx":0E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleRegister.frx":0E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleRegister.frx":1424
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleRegister.frx":19BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   900
      Left            =   0
      TabIndex        =   25
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
         Name            =   "����"
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
   Begin XtremeSuiteControls.TabControl TabCtr 
      Height          =   1245
      Left            =   0
      TabIndex        =   42
      Top             =   5610
      Width           =   1965
      _Version        =   589884
      _ExtentX        =   3466
      _ExtentY        =   2196
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   8370
      Top             =   1050
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmLabSampleRegister.frx":1F58
      Left            =   8400
      Top             =   750
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLabSampleRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mDkp                                   '����ID
    �걾�Ǽ� = 0
    ҽ���б�
    �����б�
    �ǼǼ���
    �Ǽ��Թ�
End Enum
Private Enum mPcol                                  '�����б�
    ����ID = 1
    ����
    ��Դ
    ��������
    �Ա�
    ����
    ��ʶ��
    ����
    ���˿���
    δ�Ǽ�
    �ѵǼ�
    ����
    ��ִ��
    �ز�
    Ӥ����
End Enum

Private Enum mAcol                                  'ҽ���б�
    ID
    ѡ��
    ͼ��
    ��ִ��
    �ɼ���ʽ
    ҽ������
    ����
    ִ�п���
    ����ҽ��
    ����ʱ��
    ������
    ����ʱ��
    �걾
    ����ʱ��
    �Թ���ɫ
    �ϲ�ҽ��
    �Թܱ���
    ������
    �ͼ���
    ��Ѫ��
    �Թ�����
    ����
    ������Դ
    �������
    Ӥ��
    ����
    ���ID
    ҽ��id
    ����ID
    ����
    �Ա�
    ����
    ��ʶ��
    ����
    ���˿���
    ������
    ����ʱ��
    ������ĿID
    ִ��״̬
    ����
    �ͼ�ʱ��
End Enum
Private Enum mCuvette                               '�Թ�
    ѡ��
    ����
    ����
    ��Ӽ�
    ��Ѫ��
    ���
    ��ɫ
End Enum
Private Enum mCuvetteCount                          '�Թܼ���
    ����
    ����
    �ϼ�
End Enum
Private Enum mFilter                                '��������
    ��ʶ�� = 0
    ���￨
    ����
    ���ݺ�
    �걾
    �ɼ���ʽ
    ����
    סԺ
    ���
    ���˿���
    ���ʱ��
    ��ʼʱ��
    ����ʱ��
End Enum
Private mlngKey As Long                             '����ID
Private mlngDeptID As Long                          '����ID
Private mstrPrivs As String
Private mlngBatch As Long                           '����
Private mblnUse As Boolean                          '��ǰ�����Ƿ�ʹ��
Private mlngSelectBatch As Long                     '��ǰѡ�������
Private Enum IDKinds
    C0���� = 0
    C1ҽ���� = 1
    C2���֤�� = 2
    C3IC���� = 3
    C4����� = 4
    C5���￨ = 5
End Enum
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object, mbln���֤ As Boolean
Private Const conMenu_IDkind_Change  As Integer = 12345
Private mobjSquareCard As Object                                        'ȡ������
Private mobjLisInsideComm As Object                                     'LIS�ڲ��ӿ�
Private mblnShowPwd As Boolean                                          '�Ƿ���ʾ����
Private mintBabyNo  As Integer                                          'Ӥ����
'��������
Private Enum BarCodeType
    Code39 = 1
    Code128 = 2
End Enum
Private mstrFirstBarCode   As String        '��һ��ɨ����
Private mintCodeType As Integer             '39���128��
Private mrsSendPerson As Recordset          '�ͼ���Ա��¼��
Private mstrSendPerson As String            '�ͼ���
    
Private Sub CreateCbs(Optional ByVal blnSecond As Boolean)
    '���ܴ���������
    
    '�����˵�
    Dim Control As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim ControlFile As CommandBarPopup
    Dim ControlSelect As CommandBarPopup
    Dim ControlComboBox As CommandBarComboBox
    Dim intShowType As Integer
    
    'ȥ����չ��ť
    cbrthis.VisualTheme = xtpThemeOffice2003
    Set cbrthis.Icons = zlCommFun.GetPubIcons
    With cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbrthis.EnableCustomization False
    cbrthis.ActiveMenuBar.Controls.DeleteAll
    '-----------------------------------------------------
    
    '==�ļ��˵�
    Set ControlFile = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    With ControlFile.CommandBar.Controls
        .Add xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��"
        .Add xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)"
        .Add xtpControlButton, conMenu_File_Print, "��ӡ"
        Set Control = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): Control.BeginGroup = True
    End With
    
    '==�༭
    Set ControlFile = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    With ControlFile.CommandBar.Controls
        Set Control = .Add(xtpControlButton, conMenu_Manage_Request, "�Ǽ�(&G)")
        Set Control = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��(&C)")
        Set Control = .Add(xtpControlButton, conMenu_Edit_Insert, "��ʼ����(&I)")
        Set Control = .Add(xtpControlButton, conMenu_Edit_Import, "�����˶�(&P)")
        Set Control = .Add(xtpControlButton, conMenu_Edit_Delete, "����(&R)")
    End With
    
    '==�鿴�˵�
    Set ControlFile = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    With ControlFile.CommandBar.Controls
        Set ControlSelect = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        ControlSelect.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        ControlSelect.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        ControlSelect.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set Control = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): Control.Checked = True
        Set Control = .Add(xtpControlButton, conMenu_View_Filter, "����(&F)"): Control.BeginGroup = True
        Set Control = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
    End With
   
    '==�����˵�
    Set ControlFile = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    With ControlFile.CommandBar.Controls
        .Add xtpControlButton, conMenu_Help_Help, "��������(&H)"
        .Add xtpControlButton, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName
        .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&S)", -1, False
        Set Control = .Add(xtpControlButton, conMenu_Help_About, "����(&A)")
        Control.BeginGroup = True
    End With
    
    '==�б����
    If chkDept.Value = 0 Then
        Set ControlFile = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "�ɼ�����")
    Else
        Set ControlFile = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "ִ�п���")
    End If
    With ControlFile.CommandBar.Controls
        intShowType = zlDatabase.GetPara("�Ƿ񰴲�����ʾ", 100, 1212, 0)
        Set Control = .Add(xtpControlButton, conMenu_File_MedRecPreview, "�����Ҳ鿴"): Control.Checked = (intShowType = 0)
        Set Control = .Add(xtpControlButton, conMenu_File_MedRecPrint, "�������鿴"): Control.Checked = (intShowType = 1)
    End With
    ControlFile.Flags = xtpFlagRightAlign
    Set ControlComboBox = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlComboBox, conMenu_View_Busy, "����")
    ControlComboBox.ShortcutText = "����"
    ControlComboBox.Width = 130
    ControlComboBox.Flags = xtpFlagRightAlign
    ControlComboBox.Style = xtpButtonIconAndCaption
    ControlComboBox.DropDownListStyle = True
    
    '����
    Set Control = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "��������")
    Control.Flags = xtpFlagRightAlign
    Set Control = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlComboBox, conMenu_File_RoomSet, "����")
    Control.ShortcutText = "����"
    Control.Width = 130
    Control.Flags = xtpFlagRightAlign
    Control.Style = xtpButtonIconAndCaption
    
    If Not blnSecond Then
        '����������
        Dim Toolbar As CommandBar
        Dim ControlPopup As CommandBarPopup
        
        Set Toolbar = cbrthis.Add("������", xtpBarTop)
        Toolbar.ShowTextBelowIcons = False
        Toolbar.EnableDocking xtpFlagStretched
        With Toolbar.Controls
            .Add xtpControlButton, conMenu_File_Preview, "Ԥ��"
            .Add xtpControlButton, conMenu_File_Print, "��ӡ"
            Set Control = .Add(xtpControlButton, conMenu_Manage_Request, "�Ǽ�"): Control.BeginGroup = True
            Set Control = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��")
            Set Control = .Add(xtpControlButton, conMenu_Edit_ReprintReceipt, "�ش�����")
            Control.Enabled = False
            Set Control = .Add(xtpControlButton, conMenu_Edit_Insert, "��ʼ����")
            Set Control = .Add(xtpControlButton, conMenu_Edit_Import, "�����˶�")
            Set Control = .Add(xtpControlButton, conMenu_Edit_Delete, "����"): Control.BeginGroup = True
            
            Set Control = .Add(xtpControlButton, conMenu_View_Filter, "����"): Control.BeginGroup = True
            Set Control = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
            Set Control = .Add(xtpControlButton, conMenu_Help_Help, "����"): Control.BeginGroup = True
            Set Control = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): Control.BeginGroup = True
        End With
        
        For Each Control In Toolbar.Controls
            Control.Style = xtpButtonIconAndCaption
        Next
    End If
    
    '�����
    With cbrthis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add FCONTROL, Asc("F"), conMenu_View_Filter
        .Add 0, VK_F2, conMenu_Manage_Request
        .Add 0, VK_F4, conMenu_Edit_Untread
        .Add 0, VK_F10, conMenu_IDkind_Change
        .Add 0, VK_F6, conMenu_Manage_Plan
    End With
    '���ò����ò˵�
    With cbrthis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    Call zlDatabase.ShowReportMenu(Me.cbrthis, glngSys, glngModul, mstrPrivs)
End Sub

Private Function FindDept(ByVal cboCtrol As CommandBarComboBox, ByVal strTemp As String) As String
          '�����롢���ơ������ѯ����
          Dim i As Integer
          Dim intShowType As Integer
          Dim rsTmp As Recordset
          Dim strSQL As String
          
1     On Error GoTo FindDept_Error

2         For i = 1 To cboCtrol.ListCount
3             If cboCtrol.List(i) = strTemp Then
4                 FindDept = strTemp
5                 Exit Function
6             End If
7         Next
          
8         intShowType = IIf(Me.cbrthis.FindControl(, conMenu_File_MedRecPreview, , True).Checked, 0, 1)
          
9         If intShowType = 0 Then
10            strSQL = _
                      " Select Distinct A.ID,A.����,A.����" & _
                      " From ���ű� A,��������˵�� B,������Ա C " & _
                      " Where B.����ID = A.ID And A.ID=C.����ID " & _
                      " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
                      " And B.������� IN(1,2,3,4) And B.�������� IN('����','����','�ٴ�')"
11        Else
12            strSQL = _
                      " Select Distinct A.ID, A.����, A.����" & vbNewLine & _
                      " From ���ű� A, �������Ҷ�Ӧ B" & vbNewLine & _
                      " Where B.����id = A.ID And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)"
13        End If
          
14        strSQL = strSQL & " And (a.����=[1] Or a.����=[1] Or Upper(a.����)=Upper([1]))"
          
15        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strTemp)
          
16        If rsTmp.EOF Then
17            FindDept = ""
18            cboCtrol.ListIndex = 0
19        Else
20            For i = 1 To cboCtrol.ListCount
21                If cboCtrol.List(i) = rsTmp("����") & "-" & rsTmp("����") Then
22                    cboCtrol.ListIndex = i
23                    cboCtrol.Text = rsTmp("����") & "-" & rsTmp("����")
24                    FindDept = rsTmp("����") & "-" & rsTmp("����")
25                    Exit Function
26                End If
27            Next
28        End If


29        Exit Function
FindDept_Error:
30        MsgBox "zl9LisWork, frmLabSampleRegister, ִ��(FindDept)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl
31        Err.Clear
End Function

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strFilter As String                             '�����ִ�
    Dim cboCtrol As CommandBarComboBox                  '����
    Dim Controlcbo As CommandBarComboBox                '������
    Dim cbrControl As CommandBarControl                 '�ı���ǩ
    Dim strText As String
    
    Select Case Control.ID
        Case conMenu_File_PrintSet                                                  '��ӡ����
            ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1212_1", Me
        
        Case conMenu_File_Preview                                                   'Ԥ��
'            Call zlRptPrint(2)
            RegisterLisPrint (1)
            
        Case conMenu_File_Print                                                     '��ӡ
'            Call zlRptPrint(1)
            RegisterLisPrint (2)
       
        Case conMenu_File_Excel                                                     '�����Excel
            Call zlRptPrint(4)
        
        Case conMenu_File_Exit                                                      '�˳�
            Unload Me
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Manage_Request                                                 '�Ǽ�
            Call cmdOK_Click
            
        Case conMenu_Edit_Untread                                                   'ȡ��
            Call cmdOK_Click
        Case conMenu_Edit_ReprintReceipt                                            '�ش�����
            RePrintBarCode False
        Case conMenu_Edit_Insert                                                    '����
            BeginRegister   '��ʼ����
                    
        Case conMenu_Edit_Import                                                    '�����˶�
            frmLabSampleCheck.ShowMe Me
                    
        Case conMenu_Edit_Delete                                                    '����
            frmLabSampleRegisterRefuse.ShowMe Me, rptAlist(Me.TabCtr.Selected.Index).Records
            RefreshPatientData
            
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_View_ToolBar_Button                                            '��׼��ť
            Me.cbrthis(2).Visible = Not Me.cbrthis(2).Visible
            Me.cbrthis.RecalcLayout
        
        Case conMenu_View_ToolBar_Text                                              '�ı���ǩ
            For Each cbrControl In Me.cbrthis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbrthis.RecalcLayout
        Case conMenu_View_ToolBar_Size                                              '��ͼ��
            Me.cbrthis.Options.LargeIcons = Not Me.cbrthis.Options.LargeIcons
            Me.cbrthis.RecalcLayout
            
        Case conMenu_View_StatusBar                                                 '״̬��
            Control.Checked = Not Control.Checked
            Me.stbThis.Visible = Control.Checked
            Me.cbrthis.RecalcLayout
            
        Case conMenu_View_Filter                                                    '����
            frmLabSampleRegisterFilter.ShowMe Me, strFilter
            Me.rptPlist.Tag = strFilter
            If strFilter <> "" Then RefreshPatientData
            
        Case conMenu_View_Refresh                                                   'ˢ��
            RefreshPatientData
        
        Case conMenu_IDkind_Change
            Call IdKindChange
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Help_Help                                                      '��������
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web                                                       'Web�ϵ�����
            Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Home                                                  '��ҳ
            Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail                                                  '���ͷ���
            Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About                                                     '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_View_Busy                                                      '����ѡ��
            Set cboCtrol = Control
            If Trim(cboCtrol.Text) = "" Then Exit Sub
            strText = FindDept(cboCtrol, cboCtrol.Text)
            mlngDeptID = cboCtrol.ItemData(cboCtrol.ListIndex)
            RefreshPatientData
            If strText = "" Then cboCtrol.SetFocus
        Case conMenu_File_RoomSet                                                   '����ѡ��
            Set cboCtrol = Control
            mlngSelectBatch = cboCtrol.ItemData(cboCtrol.ListIndex)
            RefreshPatientData
        Case conMenu_File_MedRecPreview                                             '�����Ҳ鿴
            Control.Checked = Not Control.Checked
            Me.cbrthis.FindControl(, conMenu_File_MedRecPrint, , True).Checked = Not Control.Checked
            Call GetDept
        Case conMenu_File_MedRecPrint                                               '�������鿴
            Control.Checked = Not Control.Checked
            Me.cbrthis.FindControl(, conMenu_File_MedRecPreview, , True).Checked = Not Control.Checked
            Call GetDept
        Case Else

            If Control.ID < conMenu_ReportPopup * 100# + 1 Or Control.ID > conMenu_ReportPopup * 100# + 99 Then Exit Sub

            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
            
    End Select
End Sub

Private Sub cbrthis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.Visible = False Then Exit Sub
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
        
    Err = 0: On Error Resume Next
    Select Case Control.ID
        
        Case conMenu_File_Preview, conMenu_File_Print
            Control.Enabled = (mlngSelectBatch <> 0)
        
        Case conMenu_View_ToolBar_Button:                                                                   '��ť
            Control.Checked = Me.cbrthis(2).Visible
        
        Case conMenu_View_ToolBar_Text:                                                                     '��ť����
            Control.Checked = Not (Me.cbrthis(2).Controls(1).Style = xtpButtonIcon)
        
        Case conMenu_View_ToolBar_Size:                                                                     '��ͼ��
            Control.Checked = Me.cbrthis.Options.LargeIcons
        
        Case conMenu_View_StatusBar:                                                                        '״̬��
            Control.Checked = Me.stbThis.Visible
                    
        Case conMenu_Manage_Request                                                                         '�Ǽ�
            Control.Enabled = (Me.TabCtr.Selected.Index = 0 Or Me.TabCtr.Selected.Index = 3)
            
        Case conMenu_Edit_Untread                                                                           'ȡ��
            Control.Enabled = (Me.TabCtr.Selected.Index = 1)
            
        Case conMenu_Edit_Delete                                                                            '����
            Control.Enabled = (Me.TabCtr.Selected.Index = 0 Or Me.TabCtr.Selected.Index = 1)
        
        Case conMenu_File_MedRecPreview
            Control.Checked = Control.Checked
            
        Case conMenu_File_MedRecPrint
            Control.Checked = Control.Checked
        Case conMenu_Edit_ReprintReceipt
            RePrintBarCode True
        Case conMenu_Edit_Import                                                                            '�����˶�
            Control.Visible = InStr(";" & mstrPrivs & ";", ";�����˶�;") > 0
    End Select
End Sub

Private Sub Timer_Timer()
    Me.txt����ʱ��.Text = zlDatabase.Currentdate
End Sub

Private Sub chkDept_Click()
    Call CreateCbs(True)              '����������
    Call GetDept                      '�������
End Sub

Private Sub cmdOK_Click()
    '�ǼǺ�ȡ���Ǽ�
    If Me.TabCtr.Selected.Index <= 1 Or Me.TabCtr.Selected.Index = 4 Then
        SaveRegister Me.TabCtr.Selected.Index
        If Me.TabCtr.Selected.Index = 4 Then Me.cmdOK.Enabled = False
        Me.txtGoto.SetFocus
        txtGoto.Text = ""
    End If
    RefreshPatientData 1, mintBabyNo
End Sub

Private Sub Form_Load()

    On Error GoTo errH

    mstrPrivs = gstrPrivs       '��ʹ��Ȩ��
    mintCodeType = zlDatabase.GetPara("ʹ������", "100", "1211", 2)
    ChkBarCodeRegister.Value = zlDatabase.GetPara("ɨ�������ֱ�ӵǼ�", "100", "1212", 0)
    chkUrgent.Value = zlDatabase.GetPara("�����걾��ʾ", "100", "1212", 0)
    chkDept.Value = zlDatabase.GetPara("��ִ�п��ҹ���", "100", "1212", 0)
    Call CreateCbs              '����������
    Call CreateDkp              '��������
    Call CreateTab              '����Tab�б�
    Call CreateListHead         '������ͷ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hWnd)
    mbln���֤ = False
    chkComRequest = zlDatabase.GetPara("����걾�ͼ�󷽿ɽ��м���걾ǩ��", 100, 1212, 0)
    chkRemberPer.Value = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�Ǽ��ͼ��˺����", 1)
    '���ݶ���
    Call GetDept                                                '�������
    Call GetPerson                                              '�����ͼ���
    If mobjSquareCard Is Nothing Then
        Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
            MsgBox "IDKind��ʼ��ʧ��!", vbInformation, gstrSysName
        Else
            IDKind.IDKindStr = mobjSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
        End If
    End If
        
    Call RestoreWinState(Me, App.ProductName)                   '����ָ�
    RefreshPatientData
    
    Me.txt������.Text = UserInfo.����
    Me.txt����ʱ��.Text = zlDatabase.Currentdate
    
    If mobjLisInsideComm Is Nothing Then
        Dim strErr As String
        Set mobjLisInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        '��ʼ��LIS�ӿڲ���
        If Not mobjLisInsideComm Is Nothing Then
            If mobjLisInsideComm.InitComponentsHIS(glngSys, glngModul, gcnOracle, strErr) = False Then
                If strErr <> "" Then
                    MsgBox "��ʼ��LIS�ӿ�ʧ�ܣ�" & vbCrLf & strErr
                End If
                Set mobjLisInsideComm = Nothing
            End If
        End If
    End If
   
    Exit Sub
errH:
    MsgBox "��ʼ������ʱ�����������鲿�������ԣ�", vbInformation, "��ʼ��"
End Sub

Private Sub GetPerson()
    Dim strSQL As String
    strSQL = "Select Distinct d.Id, d.���, d.����, d.����" & vbNewLine & _
            "From ���ű� A, ��������˵�� B, ������Ա C, ��Ա�� D" & vbNewLine & _
            "Where a.Id = b.����id And a.Id = c.����id And c.��Աid = d.Id And a.����ʱ�� Is Not Null And" & vbNewLine & _
            "      a.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd') And (b.�������� = '�ٴ�' Or b.�������� = '����' Or b.�������� = '����' and 1=0)" & vbNewLine & _
            "Order By d.���"
    Set mrsSendPerson = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�ͼ���")
End Sub

Private Sub CreateDkp()
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, Pane4 As Pane, Pane5 As Pane
    
    dkpMan.SetCommandBars Me.cbrthis
    dkpMan.Options.DefaultPaneOptions = PaneNoCloseable
    dkpMan.Options.HideClient = True
    
    Set Pane1 = dkpMan.CreatePane(mDkp.�걾�Ǽ�, 400, 700, DockLeftOf, Nothing)
    Pane1.Title = "�걾�Ǽ�"
    Pane1.Handle = Me.picBarCodeWork.hWnd
    Pane1.Options = PaneNoCaption
    
    Set Pane2 = dkpMan.CreatePane(mDkp.ҽ���б�, 400, 300, DockBottomOf, Pane1)
    Pane2.Title = "ҽ����Ϣ"
    Pane2.Handle = Me.TabCtr.hWnd
    Pane2.Options = PaneNoCaption
    
    Set Pane3 = dkpMan.CreatePane(mDkp.�����б�, 600, 300, DockRightOf, Nothing)
    Pane3.Title = "���˲ɼ��嵥"
    Pane3.Handle = Me.picTab.hWnd
    Pane3.Options = PaneNoCaption
    
    Set Pane4 = dkpMan.CreatePane(mDkp.�ǼǼ���, 600, 150, DockBottomOf, Pane3)
    Pane4.Title = "�ǼǼ���"
    Pane4.Handle = Me.PicCount.hWnd
    Pane4.Options = PaneNoCaption
    
    Set Pane5 = dkpMan.CreatePane(mDkp.�Ǽ��Թ�, 600, 150, DockBottomOf, Pane4)
    Pane5.Title = "�Ǽ��Թ�"
    Pane5.Handle = Me.PicCuvetteCount.hWnd
    Pane5.Options = PaneNoCaption
    
    Pane1.Select
    
End Sub
Private Sub CreateTab()
    Dim Item As TabControlItem
    
    With Me.TabCtr
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.COLOR = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .InsertItem 0, "δ�Ǽ�", Me.rptAlist(0).hWnd, 0
        .InsertItem 1, "�ѵǼ�", Me.rptAlist(1).hWnd, 0
        .InsertItem 2, "ִ����", Me.rptAlist(2).hWnd, 0
        .InsertItem 3, "����", Me.rptAlist(3).hWnd, 0
'
        .PaintManager.LayOut = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .Item(0).Selected = True
    End With
End Sub

Private Sub Form_Resize()
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, Pane4 As Pane
    If Me.Visible = False Then Exit Sub
    Set Pane1 = Me.dkpMan.FindPane(mDkp.�걾�Ǽ�)
    Pane1.MaxTrackSize.SetSize 7995 / Screen.TwipsPerPixelX, 4770 / Screen.TwipsPerPixelY
    Pane1.MinTrackSize.SetSize 7995 / Screen.TwipsPerPixelX, 4770 / Screen.TwipsPerPixelY
    
    Me.cbrthis.RecalcLayout
    
    Set Pane2 = Me.dkpMan.FindPane(mDkp.�����б�)
    Pane2.MaxTrackSize.SetSize 7995 / Screen.TwipsPerPixelX, 4770 / Screen.TwipsPerPixelY
    Pane2.MinTrackSize.SetSize 7995 / Screen.TwipsPerPixelX, 4770 / Screen.TwipsPerPixelY
    
    Set Pane3 = Me.dkpMan.FindPane(mDkp.�ǼǼ���)
    Pane3.MaxTrackSize.SetSize 7995 / Screen.TwipsPerPixelX, 4770 / Screen.TwipsPerPixelY
    Pane3.MinTrackSize.SetSize 7995 / Screen.TwipsPerPixelX, 4770 / Screen.TwipsPerPixelY
'
    Set Pane4 = Me.dkpMan.FindPane(mDkp.�Ǽ��Թ�)
    Pane4.MaxTrackSize.SetSize 7995 / Screen.TwipsPerPixelX, 4770 / Screen.TwipsPerPixelY
    Pane4.MinTrackSize.SetSize 7995 / Screen.TwipsPerPixelX, 4770 / Screen.TwipsPerPixelY
    
    
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    
    Pane1.MinTrackSize.SetSize 100, 100
    Pane2.MinTrackSize.SetSize 100, 100
    Pane3.MinTrackSize.SetSize 100, 100
    Pane4.MinTrackSize.SetSize 100, 100

End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngSelectBatch = 0
    mlngBatch = 0
    Set mobjSquareCard = Nothing
    Call zlDatabase.SetPara("�Ƿ񰴲�����ʾ", IIf(Me.cbrthis.FindControl(, conMenu_File_MedRecPreview, , True).Checked, 0, 1), 100, 1212)
    zlDatabase.SetPara "����걾�ͼ�󷽿ɽ��м���걾ǩ��", chkComRequest, 100, 1212
    Call zlDatabase.SetPara("ɨ�������ֱ�ӵǼ�", IIf(ChkBarCodeRegister.Value = 1, 1, 0), 100, 1212)
    Call zlDatabase.SetPara("�����걾��ʾ", IIf(chkUrgent.Value = 1, 1, 0), 100, 1212)
    Call zlDatabase.SetPara("��ִ�п��ҹ���", IIf(chkDept.Value = 1, 1, 0), 100, 1212)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�Ǽ��ͼ��˺����", chkRemberPer.Value)
End Sub

Private Sub IDKind_Click()
    Dim lng�����ID As Long, strOutCardNO As String, strExpand As String, strOutPatiInforXML As String
    If IDKind.IDKind = IDKinds.C3IC���� Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtGoto.Text = mobjICCard.Read_Card()
            If txtGoto.Text <> "" Then Call txtGoto_KeyPress(vbKeyReturn)
        End If
    End If
    lng�����ID = Val(IDKind.GetKindItem("�����ID"))
    If lng�����ID = 0 Then Exit Sub
    
    If mobjSquareCard.zlReadCard(Me, glngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtGoto.Text = strOutCardNO
    If txtGoto.Text <> "" Then Call txtGoto_KeyPress(vbKeyReturn)
End Sub

Private Sub IDKind_ItemClick(Index As Integer)
    mblnShowPwd = Trim(IDKind.GetKindItem(7)) <> ""
    Me.txtGoto = ""
    If mblnShowPwd = True Then
        Me.txtGoto.PasswordChar = "*"
    Else
        Me.txtGoto.PasswordChar = ""
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    mbln���֤ = False
    If Not txtGoto.Locked And txtGoto.Text = "" And Me.ActiveControl Is txtGoto Then
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKinds.C2���֤��
        txtGoto.Text = strID
        mbln���֤ = True
        Call txtGoto_KeyPress(vbKeyReturn)
        mbln���֤ = False
        IDKind.IDKind = lngPreIDKind
    End If
End Sub

Private Sub PicCount_Resize()
    On Error Resume Next
    
    With SrtCount
        .Top = 0
        .Left = 0
        .Width = Me.picTab.ScaleWidth
    End With
    
    With rptCount
        .Top = Me.srtPatient.Top + Me.srtPatient.Height + 10
        .Left = 0
        .Width = Me.PicCount.ScaleWidth
        .Height = Me.PicCount.ScaleHeight - .Top
    End With
    

End Sub

Private Sub PicCuvetteCount_Resize()
    On Error Resume Next
    
    With srtCuvetteCount
        .Top = 0
        .Left = 0
        .Width = Me.PicCuvetteCount.ScaleWidth
    End With
    
    With rptCuvetteCount
        .Top = Me.srtCuvetteCount.Top + Me.srtCuvetteCount.Height + 10
        .Left = 0
        .Width = Me.PicCuvetteCount.ScaleWidth
        .Height = Me.PicCuvetteCount.ScaleHeight - .Top
    End With
End Sub

Private Sub picTab_Resize()
    On Error Resume Next
    
    With srtPatient
        .Top = 0
        .Left = 0
        .Width = Me.picTab.ScaleWidth
    End With
    
    With rptPlist
        .Top = Me.srtPatient.Top + Me.srtPatient.Height + 10
        .Left = 0
        .Width = Me.picTab.ScaleWidth
        .Height = Me.picTab.ScaleHeight - .Top
    End With
End Sub
Private Sub CreateListHead()
    '�����б�ͷ
    Dim Column As ReportColumn
    Dim intLoop As Integer
    
    '==ҽ���б�ͷ
    
    rptPlist.AllowColumnRemove = False
    rptPlist.ShowItemsInGroups = False
    
    With rptPlist.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "�϶��б��⵽����,�����з���..."
        .NoItemsText = "û�п���ʾ����Ŀ..."
        .VerticalGridStyle = xtpGridSolid
        
    End With
    rptPlist.SetImageList ImgList
    
    With Me.rptPlist.Columns
        Set Column = .Add(mPcol.����ID, "����ID", 0, False)
        
        Set Column = .Add(mPcol.��Դ, "��Դ", 45, True)
        Set Column = .Add(mPcol.��������, "��������", 75, True)
        Set Column = .Add(mPcol.�Ա�, "�Ա�", 60, True)
        Set Column = .Add(mPcol.����, "����", 60, True)
        Set Column = .Add(mPcol.���˿���, "���˿���", 75, True)
        Set Column = .Add(mPcol.��ʶ��, "��ʶ��", 60, True)
        Set Column = .Add(mPcol.����, "����", 60, True)
        
        Set Column = .Add(mPcol.δ�Ǽ�, "δ�Ǽ�", 45, True)
        Set Column = .Add(mPcol.�ѵǼ�, "�ѵǼ�", 45, True)
        Set Column = .Add(mPcol.����, "����", 30, True)
        Set Column = .Add(mPcol.��ִ��, "��ִ��", 45, True)
        Set Column = .Add(mPcol.�ز�, "�ز�", 30, True)
        Set Column = .Add(mPcol.Ӥ����, "Ӥ����", 0, False)
    End With
    
    For intLoop = 0 To 3
        '==�����б�ͷ
        rptAlist(intLoop).AllowColumnRemove = False
        rptAlist(intLoop).ShowItemsInGroups = False
        With rptAlist(intLoop).PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
            .HideSelection = True
        End With
        rptAlist(intLoop).SetImageList ImgList
        With Me.rptAlist(intLoop).Columns
            Set Column = .Add(mAcol.ID, "ID", 0, False): Column.Visible = False
            Set Column = .Add(mAcol.ѡ��, "Check", 18, False): Column.Icon = 0
            Set Column = .Add(mAcol.ͼ��, "", 18, False): Column.Icon = 3
            Set Column = .Add(mAcol.��ִ��, "��ִ��", 45, False): Column.Visible = False: Column.Alignment = xtpAlignmentCenter
            Set Column = .Add(mAcol.����, "����", 30, False): Column.Alignment = xtpAlignmentCenter
            Set Column = .Add(mAcol.�ɼ���ʽ, "�ɼ���ʽ", 75, True)
            Set Column = .Add(mAcol.�걾, "�걾", 55, True)
            Set Column = .Add(mAcol.ҽ������, "ҽ������", 75, True)
            Set Column = .Add(mAcol.����, "����", 75, True)
            Set Column = .Add(mAcol.ִ�п���, "ִ�п���", 75, True)
            Set Column = .Add(mAcol.����ҽ��, "����ҽ��", 75, True)
            Set Column = .Add(mAcol.����ʱ��, "����ʱ��", 75, True)
            Set Column = .Add(mAcol.������, "������", 65, True)
            Set Column = .Add(mAcol.�ͼ���, "�ͼ���", 65, True)
            Set Column = .Add(mAcol.������, "������", 65, True)
            Set Column = .Add(mAcol.����ʱ��, "����ʱ��", 75, True)
            Set Column = .Add(mAcol.����ʱ��, "����ʱ��", 75, True)
            Set Column = .Add(mAcol.����ʱ��, "����ʱ��", 75, True)
            Set Column = .Add(mAcol.�Թ���ɫ, "��ɫ����", 18, True): Column.Visible = False
            Set Column = .Add(mAcol.�Թܱ���, "�Թܱ���", 18, True): Column.Visible = False
            Set Column = .Add(mAcol.������, "������", 60, True)
            Set Column = .Add(mAcol.��Ѫ��, "��Ѫ��", 60, True): Column.Visible = False
            Set Column = .Add(mAcol.�Թ�����, "�Թ�����", 60, True): Column.Visible = False
            Set Column = .Add(mAcol.����, "����", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.������Դ, "������Դ", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.�������, "�������", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.Ӥ��, "Ӥ��", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.����, "����", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.���ID, "���ID", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.����ID, "����ID", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.����, "����", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.�Ա�, "�Ա�", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.����, "����", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.��ʶ��, "��ʶ��", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.����, "����", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.���˿���, "���˿���", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.������ĿID, "������ĿId", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.ִ��״̬, "ִ��״̬", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.ҽ��id, "ҽ��ID", 50, False): Column.Visible = False
            Set Column = .Add(mAcol.�ͼ�ʱ��, "�ͼ�ʱ��", 75, True)
            
        End With
    Next
    
    '==�����б�ͷ
    rptCount.AllowColumnRemove = False
    rptCount.ShowItemsInGroups = False
    With rptCount.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "�϶��б��⵽����,�����з���..."
        .NoItemsText = "û�п���ʾ����Ŀ..."
        .VerticalGridStyle = xtpGridSolid
        .HideSelection = True
    End With
    rptCount.SetImageList ImgList
    With rptCount.Columns
        Set Column = .Add(mAcol.ID, "ID", 0, False): Column.Visible = False
        Set Column = .Add(mAcol.ѡ��, "Check", 18, False): Column.Icon = 0
        Set Column = .Add(mAcol.ͼ��, "", 18, False): Column.Icon = 3
        Set Column = .Add(mAcol.��ִ��, "��ִ��", 45, False): Column.Visible = False: Column.Alignment = xtpAlignmentCenter
        Set Column = .Add(mAcol.����, "����", 30, False): Column.Alignment = xtpAlignmentCenter
        Set Column = .Add(mAcol.������Դ, "��Դ", 35, True)
        Set Column = .Add(mAcol.����, "����", 70, True)
        Set Column = .Add(mAcol.�Ա�, "�Ա�", 35, True)
        Set Column = .Add(mAcol.��ʶ��, "��ʶ��", 70, True)
        Set Column = .Add(mAcol.�ɼ���ʽ, "�ɼ���ʽ", 75, True)
        Set Column = .Add(mAcol.�걾, "�걾", 55, True)
        Set Column = .Add(mAcol.ҽ������, "ҽ������", 75, True)
        Set Column = .Add(mAcol.����, "����", 75, True)
        Set Column = .Add(mAcol.ִ�п���, "ִ�п���", 75, True)
        Set Column = .Add(mAcol.����ҽ��, "����ҽ��", 75, True)
        Set Column = .Add(mAcol.����ʱ��, "����ʱ��", 80, True)
        Set Column = .Add(mAcol.������, "������", 65, True)
        Set Column = .Add(mAcol.�ͼ���, "�ͼ���", 65, True)
        Set Column = .Add(mAcol.������, "������", 65, True)
        Set Column = .Add(mAcol.����ʱ��, "����ʱ��", 80, True)
        Set Column = .Add(mAcol.����ʱ��, "����ʱ��", 80, True)
        Set Column = .Add(mAcol.����ʱ��, "����ʱ��", 80, True)
        Set Column = .Add(mAcol.�Թ���ɫ, "��ɫ����", 18, True): Column.Visible = False
        Set Column = .Add(mAcol.�Թܱ���, "�Թܱ���", 18, True): Column.Visible = False
        Set Column = .Add(mAcol.������, "������", 60, True)
        Set Column = .Add(mAcol.��Ѫ��, "��Ѫ��", 60, True): Column.Visible = False
        Set Column = .Add(mAcol.�Թ�����, "�Թ�����", 60, True): Column.Visible = False
        Set Column = .Add(mAcol.����, "����", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.�������, "�������", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.Ӥ��, "Ӥ��", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.����, "����", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.���ID, "���ID", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.����ID, "����ID", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.����, "����", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.����, "����", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.���˿���, "���˿���", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.������ĿID, "������ĿId", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.ִ��״̬, "ִ��״̬", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.ҽ��id, "ҽ��ID", 50, False): Column.Visible = False
        Set Column = .Add(mAcol.�ͼ�ʱ��, "�ͼ�ʱ��", 80, True)
    End With
    
    '==�Թ�
    rptCuvette.AllowColumnRemove = False
    rptCuvette.ShowItemsInGroups = False
    
    With rptCuvette.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "�϶��б��⵽����,�����з���..."
        .NoItemsText = "û�п���ʾ����Ŀ..."
        .VerticalGridStyle = xtpGridSolid
        
    End With
    rptCuvette.SetImageList ImgList
    With Me.rptCuvette.Columns
        Set Column = .Add(mCuvette.ѡ��, "Check", 18, False): Column.Icon = 0
        Set Column = .Add(mCuvette.����, "����", 75, True)
        Set Column = .Add(mCuvette.����, "����", 120, True)
        Set Column = .Add(mCuvette.��Ӽ�, "��Ӽ�", 110, True)
        Set Column = .Add(mCuvette.��Ѫ��, "��Ѫ��", 80, True)
        Set Column = .Add(mCuvette.���, "���", 80, True)
        Set Column = .Add(mCuvette.��ɫ, "", 18, True): Column.Icon = 3
    End With
    
    '==�Թܼ���
    rptCuvetteCount.AllowColumnRemove = False
    rptCuvetteCount.ShowItemsInGroups = False
    
    With rptCuvetteCount.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "�϶��б��⵽����,�����з���..."
        .NoItemsText = "û�п���ʾ����Ŀ..."
        .VerticalGridStyle = xtpGridSolid
        
    End With
    rptCuvetteCount.SetImageList ImgList
    With Me.rptCuvetteCount.Columns
        Set Column = .Add(mCuvetteCount.����, "����", 75, True)
        Set Column = .Add(mCuvetteCount.����, "����", 120, True)
        Set Column = .Add(mCuvetteCount.�ϼ�, "�ϼ�", 45, True)
    End With
End Sub

Private Sub RefreshPatientData(Optional lngPatientType As Long = 0, Optional ByVal intBabyNo As Integer = 0)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String                                '����SQL�����
    Dim strSQL1 As String                               '����H��Ĳ���
    Dim Record As ReportRecord
    Dim Item As ReportColumn
    Dim intLoop As Integer
    Dim strTmp As String                                '��ʱ�ִ�����
    Dim varFilter As Variant                            '�����ִ�
    Dim strDateBegin As Date                            '��ʼʱ��
    Dim strDateEnd As Date                              '����ʱ��
    Dim blnDateMoved As Boolean                         '�Ƿ�ת��
    Dim lngPatientID As Long                            '����ID
    Dim strState As String                              '״̬
    Dim strDeptIDs As String                            '������ID
    Dim int��ҳID As Integer                            '��ҳID
    Dim str�Һŵ� As String                             '�Һŵ�
    Dim intPatientType As Integer                       '������Դ
    
    On Error GoTo errH
    
    '��������
    If TabCtr.Selected.Index > 0 Then
        GetBatch
    End If
    
    '��ע����ж�ȡ��������
    strTmp = zlDatabase.GetPara("�걾�Ǽǹ���", 100, 1212, "")
    
    '�ӹ��˴������������ʱ����
    If Me.rptPlist.Tag <> "" Then
        varFilter = Split(Me.rptPlist.Tag, ";")
    Else
        If strTmp <> "" Then
            varFilter = Split(strTmp, ";")
        End If
    End If
    
    zlCommFun.ShowFlash "���ڸ�������,���Ժ�...", Me
    
    strSQL = "Select ����id,������Դ,��������,���˿���,�Ա�,����,���￨��,��ʶ��,��ǰ����," & vbNewLine & _
                "       Sum(decode(״̬,'δ�Ǽ�',1,0)) As δ�Ǽ�,Sum(decode(״̬,'�ѵǼ�',1,0)) As �ѵǼ�," & vbNewLine & _
                "       Sum(decode(״̬,'����',1,0)) As ����,Sum(decode(״̬,'��ִ��',1,0)) As ��ִ��," & vbNewLine & _
                "       Sum(decode(����,'����',1,0)) As ����,Sum(�ز�) As �ز�,Nvl(sum(Ӥ��)/count(Ӥ��),0) as Ӥ���� From ("
    
    If chkDept.Value = 1 Then
        '��ִ�п��ҹ���
        strSQL = strSQL & "Select Distinct a.���id, a.����id, Decode(a.������Դ, 1, '����', 2, 'סԺ', 3, 'Ժ��', 4, '���') As ������Դ," & vbNewLine & _
                        "                Decode(Nvl(a.Ӥ��, 0), 0, c.����, t.Ӥ������) As ��������, e.���� As ���˿���, Decode(Nvl(a.Ӥ��, 0), 0, c.�Ա�, t.Ӥ���Ա�) As �Ա�," & vbNewLine & _
                        "                Decode(Nvl(a.Ӥ��, 0), 0, c.����, Nvl(Round(Nvl(t.����ʱ��, Sysdate) - t.����ʱ��), 0) || '��') As ����," & vbNewLine & _
                        "                c.���￨��, b.��������, Decode(b.ִ��״̬, 1, '��ִ��', 2, '����', 3, '��ִ��', Decode(b.������, Null, 'δ�Ǽ�', '�ѵǼ�')) As ״̬," & vbNewLine & _
                        "                a.Ӥ��, Decode(a.������Դ, 1, c.�����, 2, c.סԺ��) As ��ʶ��," & vbNewLine & _
                        "                Decode(c.��ǰ����, Null, Decode(l.��Ժ����, Null, l.��Ժ����, l.��Ժ����), c.��ǰ����) As ��ǰ����," & vbNewLine & _
                        "                Decode(a.������־, 1, '����', Decode(g.����, 1, '����')) As ����, Decode(b.ִ��״̬, 0, '', 2, '����') As ����," & vbNewLine & _
                        "                Nvl(b.�زɱ걾, 0) As �ز�, k.�Թܱ���" & vbNewLine & _
                        "From ����ҽ����¼ H, ����ҽ����¼ A, ����ҽ������ B, ������Ϣ C, ���ű� E, ������ĿĿ¼ F, ���˹Һż�¼ G, ������ĿĿ¼ K, ������ҳ L, ������������¼ T" & vbNewLine & _
                        "Where h.Id = a.���id And a.Id = b.ҽ��id And a.����id = c.����id And a.���˿���id = e.Id And a.������Ŀid = k.Id And h.������Ŀid = f.Id And" & vbNewLine & _
                        "      a.����id = t.����id(+) And a.��ҳid = t.��ҳid(+) And a.Ӥ�� = t.���(+) And a.�Һŵ� = g.No(+) And" & vbNewLine & _
                        "      (g.����id Is Null Or (g.��¼״̬ = 1 And g.��¼���� = 1)) And a.������� = 'C' And h.������� = 'E' And f.�������� = '6' And" & vbNewLine & _
                        "      a.����id = l.����id(+) And b.ִ�в���id  In (Select * From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) And" & vbNewLine & _
                        "      k.�Թܱ��� Is Not Null And b.ִ��״̬ In (0, 1, 2, 3) And b.�������� = 1 "
    Else
        '���ɼ����ҹ���
        strSQL = strSQL & "Select distinct h.���id,a.����id,decode(a.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���') as ������Դ, " & vbCrLf & _
                        " decode(nvl(a.Ӥ��,0),0,c.����,t.Ӥ������)  as ��������,e.���� as ���˿���,decode(nvl(a.Ӥ��,0),0,c.�Ա�,t.Ӥ���Ա�) as �Ա�," & vbCrLf & _
                        " decode(nvl(a.Ӥ��,0),0,c.����,Nvl(Round(Nvl(t.����ʱ��, Sysdate) - t.����ʱ��), 0) ||'��') as ����,c.���￨��,b.��������, " & vbCrLf & _
                        " decode(b.ִ��״̬,1, '��ִ��',2,'����',3,'��ִ��', decode(b.������,null,'δ�Ǽ�','�ѵǼ�')) as ״̬,a.Ӥ��, " & vbCrLf & _
                        " decode(A.������Դ, 1, C.�����, 2, C.סԺ��) As ��ʶ��, " & vbCrLf & _
                        " decode(c.��ǰ����,null,decode(l.��Ժ����,null,l.��Ժ����,l.��Ժ����),c.��ǰ����) as ��ǰ���� , " & vbCrLf & _
                        " decode(a.������־,1,'����',decode(g.����,1,'����')) as ���� , " & vbCrLf & _
                        " decode(b.ִ��״̬,0,'',2,'����') as ����,nvl(b.�زɱ걾,0) as �ز�, k.�Թܱ��� " & vbCrLf & _
                        " From ����ҽ����¼ A, ����ҽ������ B, ������Ϣ C, ���ű� E, ������ĿĿ¼ F,���˹Һż�¼ G,����ҽ����¼ H, " & vbCrLf & _
                        "      ������ĿĿ¼ K ,������ҳ L,����ҽ������ M,������������¼ T" & vbCrLf & _
                        " Where A.ID = H.���ID And H.id = B.ҽ��id And A.����id = C.����id And A.���˿���id = E.ID And A.������Ŀid = f.ID " & vbCrLf & _
                        "      And h.������ĿID = k.id " & vbCrLf & _
                        "and A.����ID=T.����ID(+) and A.��ҳID=T.��ҳID(+) and A.Ӥ��=T.���(+)" & vbCrLf & _
                        " And A.�Һŵ� = G.No(+) and (g.����ID is null or (g.��¼״̬=1 and g.��¼���� =1) ) and a.������� = 'E' and F.�������� = '6' and a.����id = l.����ID(+) " & vbCrLf & _
                        " and m.ִ�в���id  in (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) " & vbCrLf & _
                        " And A.ID = M.ҽ��ID And k.�Թܱ��� is not null and B.ִ��״̬ in (0,1,2,3) and m.�������� = 1 "
    End If
    
    '���Ҳ���ID
    gstrSql = "select ����ID from �������Ҷ�Ӧ where ����ID = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDeptID)
    If rsTmp.EOF = True Then
        strDeptIDs = mlngDeptID
    Else
        Do Until rsTmp.EOF
            strDeptIDs = strDeptIDs & "," & rsTmp("����ID")
            rsTmp.MoveNext
        Loop
        If strDeptIDs <> "" Then strDeptIDs = Mid(strDeptIDs, 2) & "," & mlngDeptID
    End If
    
    '����
    If mlngSelectBatch <> 0 And TabCtr.Selected.Index > 0 Then
        strSQL = strSQL & " and b.�������� = [11] "
    End If
    
    '���û��Ȩ�ޣ��Ͳ��ܿ���δ�����걾
    If InStr(mstrPrivs, "ǿ�ƵǼ�δ�����걾") = 0 Then
        strSQL = strSQL & " and b.������ is not null "
    End If
    
    If Me.rptPlist.Tag <> "" Then
        strSQL = strSQL & " And A.������Դ in (" & IIf(Val(varFilter(mFilter.����)) = 1, "1,3,", "") & _
                 Val(varFilter(mFilter.סԺ)) & "," & Val(varFilter(mFilter.���)) & ") "
    
        If varFilter(mFilter.��ʶ��) <> "" Then
            strSQL = strSQL & " And decode(a.������Դ,2,c.סԺ��,c.�����) = [2] "
        End If
        
        If varFilter(mFilter.���￨) <> "" Then
            strSQL = strSQL & " And C.���￨�� = [3] "
        End If
        
        If varFilter(mFilter.����) <> "" Then
            strSQL = strSQL & " And C.���� like [4] "
        End If
        
        If varFilter(mFilter.���ݺ�) <> "" Then
            strSQL = strSQL & " and B.NO = [5]"
        End If
        
        If varFilter(mFilter.�걾) <> "���б걾" Then
            strSQL = strSQL & " and A.�걾��λ = [6] "
        End If
        
        If varFilter(mFilter.�ɼ���ʽ) <> 0 Then
            strSQL = strSQL & " and f.ID +0 = [7] "
        End If
        
        If varFilter(mFilter.���˿���) <> 0 Then
            strSQL = strSQL & " And a.���˿���ID = [8] "
        End If
        
        If lngPatientType = 1 Then
            strSQL = strSQL & " And c.����id=[12]"
            'ʹ�ò���IDʱ��ʹ��
            strSQL = strSQL & " and b.����ʱ��+0 Between [9] and [10]"
        Else
            strSQL = strSQL & " and b.����ʱ�� Between [9] and [10]"
        End If
        
        If varFilter(mFilter.��ʼʱ��) = "" Then
            strDateBegin = zlDatabase.Currentdate - Val(varFilter(mFilter.���ʱ��))
            strDateEnd = zlDatabase.Currentdate
        Else
            strDateBegin = varFilter(mFilter.��ʼʱ��)
            strDateEnd = varFilter(mFilter.����ʱ��)
        End If
    Else
        If strTmp <> "" Then
            strSQL = strSQL & " And A.������Դ in (" & IIf(Val(varFilter(mFilter.����)) = 1, "1,3,", "") & _
                 Val(varFilter(mFilter.סԺ)) & "," & Val(varFilter(mFilter.���)) & ") "
    
            If varFilter(mFilter.�걾) <> "���б걾" Then
                strSQL = strSQL & " and A.�걾��λ = [6] "
            End If
            
            If varFilter(mFilter.�ɼ���ʽ) <> 0 Then
                strSQL = strSQL & " and f.ID +0= [7] "
            End If
            
            If varFilter(mFilter.���˿���) <> 0 Then
                strSQL = strSQL & " And a.���˿���ID = [8] "
            End If
            
            If lngPatientType = 1 Then
                strSQL = strSQL & " And c.����id=[12]"
                'ʹ�ò���IDʱ��ʹ��
                strSQL = strSQL & " and b.����ʱ��+0 Between [9] and [10]"
            Else
                strSQL = strSQL & " and b.����ʱ�� Between [9] and [10]"
            End If
            
            If Val(varFilter(mFilter.���ʱ��)) >= 0 Then
                strDateBegin = zlDatabase.Currentdate - Val(varFilter(mFilter.���ʱ��))
                strDateEnd = zlDatabase.Currentdate
            Else
                strDateBegin = varFilter(mFilter.��ʼʱ��)
                strDateEnd = varFilter(mFilter.����ʱ��)
            End If
        Else
            If lngPatientType = 1 Then
                strSQL = strSQL & " And c.����id=[12]"
                'ʹ�ò���IDʱ��ʹ��
                strSQL = strSQL & " and b.����ʱ��+0 Between [9] and [10]"
            Else
                strSQL = strSQL & " and b.����ʱ�� Between [9] and [10]"
            End If
            strDateBegin = zlDatabase.Currentdate - 3
            strDateEnd = zlDatabase.Currentdate
        End If
    End If
    
    strSQL = strSQL & ") Group By  ����id,������Դ,��������,���˿���,�Ա�,����,���￨��,��ʶ��,��ǰ���� "
    
    blnDateMoved = MovedByDate(CDate(strDateBegin)) '��ʱ�俴�Ƿ������ת��
    
    If blnDateMoved Then
        strSQL1 = strSQL
        strSQL1 = Replace(strSQL1, "����ҽ����¼", "H����ҽ����¼")
        strSQL1 = Replace(strSQL1, "����ҽ������", "H����ҽ������")
        strSQL = strSQL & " Union ALL " & strSQL1
    End If
    
    strSQL = strSQL & " Order by ���˿��� "
    
    If strTmp = "" And Me.rptPlist.Tag = "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, strDeptIDs, "", "", "", "", "", "", "", _
                    CDate(Format(strDateBegin, "yyyy-MM-dd 00:00:00")), CDate(Format(strDateEnd, "yyyy-MM-dd 23:59:59")), mlngSelectBatch, mlngKey)
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, strDeptIDs, Val(varFilter(mFilter.��ʶ��)), CStr(varFilter(mFilter.���￨)) _
                    , CStr(varFilter(mFilter.����)) & "%", CStr(varFilter(mFilter.���ݺ�)), CStr(varFilter(mFilter.�걾)), CLng(varFilter(mFilter.�ɼ���ʽ)) _
                    , mlngDeptID, CDate(Format(strDateBegin, "yyyy-MM-dd 00:00:00")), _
                    CDate(Format(strDateEnd, "yyyy-MM-dd 23:59:59")), mlngSelectBatch, mlngKey)
    End If
    If lngPatientType <> 1 Then
        '�����¼
        Me.rptPlist.Records.DeleteAll
        Me.rptAlist(TabCtr.Selected.Index).Records.DeleteAll
        Me.rptCuvette.Records.DeleteAll
        
        Do Until rsTmp.EOF
            Set Record = Me.rptPlist.Records.Add
                                    
            For intLoop = 0 To Me.rptPlist.Columns.Count + 1
                Record.AddItem ""
            Next
            
            Record(mPcol.����ID).Value = Nvl(rsTmp("����ID"))
            Record(mPcol.����).Value = Nvl(rsTmp("����"))
            If Nvl(rsTmp("����")) = "����" Then Record(mPcol.����).Icon = 2
            Record(mPcol.��Դ).Value = Nvl(rsTmp("������Դ"))
            Record(mPcol.��������).Value = Nvl(rsTmp("��������"))
            Record(mPcol.���˿���).Value = Nvl(rsTmp("���˿���"))
            Record(mPcol.�Ա�).Value = Nvl(rsTmp("�Ա�"))
            Record(mPcol.����).Value = Nvl(rsTmp("����"))
            Record(mPcol.��ʶ��).Value = Nvl(rsTmp("��ʶ��"))
            Record(mPcol.����).Value = Nvl(rsTmp("��ǰ����"))
            
            Record(mPcol.δ�Ǽ�).Value = Nvl(rsTmp("δ�Ǽ�"))
            Record(mPcol.�ѵǼ�).Value = Nvl(rsTmp("�ѵǼ�"))
            Record(mPcol.����).Value = Nvl(rsTmp("����"))
            Record(mPcol.��ִ��).Value = Nvl(rsTmp("��ִ��"))
            Record(mPcol.�ز�).Value = Nvl(rsTmp("�ز�"))
            Record(mPcol.Ӥ����).Value = Nvl(rsTmp("Ӥ����"))
            
            If Nvl(rsTmp("����"), 0) > 0 Then
                For intLoop = 0 To Me.rptPlist.Columns.Count + 1
                    Record(intLoop).ForeColor = vbRed
                Next
            End If
            
            If Nvl(rsTmp("�ز�"), 0) > 0 Then
                For intLoop = 0 To Me.rptPlist.Columns.Count + 1
                    Record(intLoop).Bold = True
                    Record(intLoop).ForeColor = vbBlue
                Next
            End If
            rsTmp.MoveNext
        Loop
    Else
        rsTmp.filter = "Ӥ����=" & "'" & intBabyNo & "'"
        Do Until rsTmp.EOF
            For intLoop = 0 To Me.rptPlist.Rows.Count - 1
                If Me.rptPlist.Rows(intLoop).Record(mPcol.����ID).Value = mlngKey And Me.rptPlist.Rows(intLoop).Record(mPcol.Ӥ����).Value = intBabyNo And Me.rptPlist.Rows(intLoop).Record(mPcol.��Դ).Value = Nvl(rsTmp("������Դ")) Then
                    Me.rptPlist.Rows(intLoop).Record(mPcol.δ�Ǽ�).Value = Nvl(rsTmp("δ�Ǽ�"))
                    Me.rptPlist.Rows(intLoop).Record(mPcol.�ѵǼ�).Value = Nvl(rsTmp("�ѵǼ�"))
                End If
            Next
            rsTmp.MoveNext
        Loop
    End If
    
    '����
    If Me.Visible = True Then
        
    Me.rptPlist.Populate
        
    End If
    If Me.Visible = True Then
        Me.rptAlist(TabCtr.Selected.Index).Populate
        Me.rptCuvette.Populate
    End If
    Me.stbThis.Panels(2).Text = "��ǰ���У�" & Me.rptPlist.Rows.Count & "�����ˣ�"
    
    '��λ���ϴ�ѡ�еĲ���
    With Me.rptPlist

        For intLoop = 0 To .Rows.Count - 1
            If .Rows(intLoop).Record(mPcol.����ID).Value = mlngKey And .Rows(intLoop).Record(mPcol.Ӥ����).Value = intBabyNo Then
                Set .FocusedRow = .Rows(intLoop)
                mlngKey = .Rows(intLoop).Record(mPcol.����ID).Value
                If .Rows(intLoop).Record(mPcol.��Դ).Value = "����" Then
                    intPatientType = 1
                ElseIf .Rows(intLoop).Record(mPcol.��Դ).Value = "סԺ" Then
                    intPatientType = 2
                ElseIf .Rows(intLoop).Record(mPcol.��Դ).Value = "Ժ��" Then
                    intPatientType = 3
                ElseIf .Rows(intLoop).Record(mPcol.��Դ).Value = "���" Then
                    intPatientType = 4
                Else
                    intPatientType = 1
                End If
                .Populate
                Me.rptPlist.Tag = ""
                Exit For
            End If
        Next
        
        If .FocusedRow Is Nothing And .Rows.Count > 0 Then
            Set .FocusedRow = .Rows(0)
            mlngKey = .Rows(0).Record(mPcol.����ID).Value
            If .Rows(0).Record(mPcol.��Դ).Value = "����" Then
                intPatientType = 1
            ElseIf .Rows(0).Record(mPcol.��Դ).Value = "סԺ" Then
                intPatientType = 2
            ElseIf .Rows(0).Record(mPcol.��Դ).Value = "Ժ��" Then
                intPatientType = 3
            ElseIf .Rows(0).Record(mPcol.��Դ).Value = "���" Then
                intPatientType = 4
            Else
                intPatientType = 1
            End If
            .Populate
        End If
        
        If Not .FocusedRow Is Nothing Then
            RefreshAdviceData mlngKey, Me.TabCtr.Selected.Index, intPatientType, False, intBabyNo
        End If
        
    End With
    
    '����������ִֻ��һ��
    Me.rptPlist.Tag = ""
    
    If Me.rptPlist.Rows.Count = 0 Then
        txt���� = ""
        txt����.Tag = ""
        txt�Ա�.Text = ""
        txt���� = ""
        txtBed = ""
        txtID = ""
        txtPatientDept = ""
    End If
    
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
    Dim intShowType As Integer
    
    mlngDeptID = zlDatabase.GetPara("����", 100, 1212, 0)
    intShowType = IIf(Me.cbrthis.FindControl(, conMenu_File_MedRecPreview, , True).Checked, 0, 1)
    
    Set Controlcbo = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_View_Busy, True, True)
    
    On Error GoTo errH
    
    If intShowType = 0 Then
        strSQL = _
                " Select Distinct A.ID,A.����,A.����" & _
                " From ���ű� A,��������˵�� B,������Ա C " & _
                " Where B.����ID = A.ID And A.ID=C.����ID " & _
                " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
                " And B.������� IN(1,2,3,4) And B.�������� IN('����','����','�ٴ�')"
    Else
        strSQL = _
                " Select Distinct A.ID, A.����, A.����" & vbNewLine & _
                " From ���ű� A, �������Ҷ�Ӧ B" & vbNewLine & _
                " Where B.����id = A.ID And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)"
    End If
'    If InStr(1, mstrPrivs, "���п���") <= 0 Then
'        strSQL = strSQL & " And C.��ԱID = [1] "
'    End If
    
    strSQL = strSQL & " Order by A.����"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    Controlcbo.Clear
    Controlcbo.ListIndex = -1
    Do Until rsTmp.EOF
        Controlcbo.AddItem rsTmp("����") & "-" & rsTmp("����")
        Controlcbo.ItemData(Controlcbo.ListCount) = rsTmp("ID")
        If rsTmp("id") = IIf(mlngDeptID = 0, UserInfo.����ID, mlngDeptID) Then
            Controlcbo.ListIndex = Controlcbo.ListCount
            mlngDeptID = IIf(mlngDeptID = 0, UserInfo.����ID, mlngDeptID)
        End If
        rsTmp.MoveNext
    Loop
    If Controlcbo.Text = "" Then
        Controlcbo.ListIndex = 1
        mlngDeptID = Controlcbo.ItemData(Controlcbo.ListIndex)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function RefreshAdviceData(lngPatientID As Long, intState As Integer, Optional intPatientType As Integer = 0, Optional blOnlyWhere As Boolean = False, Optional intBabyNo As Integer = 0) As Boolean
    '���ܣ�                         ˢ�²ɼ�ҽ����¼
    '������                         lngpatientId = ����ID ,
    '                               intPatientType = ������Դ
    '                               intState = ��ǰ״̬ 0=δ�Ǽ� 1=�ѵǼ� 2=��ִ��
    '                               blOnlyWhere = true ֻʹ�ò���ID���в���
    '                               intBabyNo  Ӥ����
    
    Dim blnDateMoved As Boolean                     '�Ƿ�ת��
    Dim strSQL As String                            'SQL���
    Dim strSQL1 As String                           '���ڲ�ѯH��
    Dim strTmp As String                            '��ʱ�ִ�����
    Dim varFilter As Variant                        '�����ִ�
    Dim strDateBegin As String                      '��ʼ����ʱ��
    Dim strDateEnd As String                        '��������ʱ��
    Dim rsTmp As New ADODB.Recordset                '���ݼ�
    Dim intLoop As Integer                          'ѭ������
    Dim Record As ReportRecord                      '�б����ݼ�
    Dim strOldAdvice As String                      '��¼�ϴ�ҽ��
    Dim strCuvetteNumber As String                  '���ڼ�¼�Թܱ���
    Dim intShowButtom As Integer                    '�����ڲ���ʱ��ʾ��Щ��ť,���� not null = 1 ������ not null = 2
    Dim strDeptIDs As String                        '������ID
    Dim int��ҳID As Integer                        '��ҳID
    Dim str�Һŵ� As String                         '�Һŵ�
    Dim str���շ�ҽ��ID  As String, str����ҽ��ID As String
    Dim strSQLbak As String
    Dim strGetSql As String, intMainID As Integer
    Dim rsTest As ADODB.Recordset
    Dim strPatiDept As String
    Dim strOldNO As String
    Dim strOldCodeBar As String

    On Error GoTo errH
    
    blnDateMoved = MovedByDate(Date) '��ʱ�俴�Ƿ������ת��
    
    '��ע����ж�ȡ��������
    strTmp = zlDatabase.GetPara("�걾�Ǽǹ���", 100, 1212, "")
    
    '�ӹ��˴������������ʱ����
    If Me.rptPlist.Tag <> "" Then
        varFilter = Split(Me.rptPlist.Tag, ";")
    Else
        If strTmp <> "" Then
            varFilter = Split(strTmp, ";")
        End If
    End If
    
    If chkDept.Value = 1 Then
        strSQL = "Select Distinct /*+ rule */ b.Id As ҽ��id, a.��ҳid, b.���id, g.��ɫ As �Թ���ɫ, d.���� As �ɼ���ʽ, b.ҽ������, c.��������, c.����ʱ��," & vbNewLine & _
                "                c.�걾�ͳ�ʱ�� As �ͼ�ʱ��, h.���� As ִ�п���, b.����ҽ��, b.����ʱ��, c.������, c.����ʱ��, g.���� As �Թܱ���, b.�걾��λ As �걾, i.���� As ��������," & vbNewLine & _
                "                i.�Ա�," & vbNewLine & _
                "                Decode(Nvl(a.Ӥ��, 0), 0, i.����, Decode(Sysdate - k.����ʱ��, '', '', Round(Sysdate - k.����ʱ��) || '��')) As ����," & vbNewLine & _
                "                i.��ǰ���� As ����, Decode(b.������Դ, 1, i.�����, 2, i.סԺ��) As ��ʶ��, k.Ӥ������, k.Ӥ���Ա�, l.���� As �������ڿ���," & vbNewLine & _
                "                Decode(c.ִ��״̬, 2, '����') As ����, i.����id, c.������, c.�ͼ���, g.��Ѫ��, g.���� As �Թ�����," & vbNewLine & _
                "                Decode(b.������־, 1, '����', '') As ����, Decode(b.������Դ, 1, '����', 2, 'סԺ', 3, 'Ժ��', 4, '���') As ������Դ, b.Ӥ��," & vbNewLine & _
                "                n.���� As ����, j.���� As ���˿���, c.����ʱ��, c.������, b.������Ŀid, c.ִ��״̬, o.��¼����, o.��¼״̬" & vbNewLine & _
                "From ����ҽ����¼ A, ����ҽ����¼ B, ����ҽ������ C, ������ĿĿ¼ D, ������ĿĿ¼ E, ��Ѫ������ G, ���ű� H, ������Ϣ I, ���ű� L, ���ű� J, ������������¼ K," & vbNewLine & _
                "     (Select ������Ŀid, ���� From ������Ŀ���� Where ���� = 9 And ���� = 1) N, ������ü�¼ O" & vbNewLine & _
                "Where a.Id = b.���id And b.Id = c.ҽ��id And a.������Ŀid = d.Id And b.������Ŀid = e.Id And e.��� = 'C' And e.�Թܱ��� = g.���� And" & vbNewLine & _
                "      b.ִ�п���id = h.Id And d.��� = 'E' And d.�������� = '6' And a.����id = [1] And c.����ʱ�� + 0 Between [3] And [4] And" & vbNewLine & _
                "      b.����id = i.����id And i.��ǰ����id = l.Id(+) And a.����id = k.����id(+) And a.��ҳid = k.��ҳid(+) And a.Ӥ�� = k.���(+) And" & vbNewLine & _
                "      e.Id = n.������Ŀid(+) And b.���˿���id = j.Id(+) And c.ҽ��id = o.ҽ�����(+) And c.��¼���� = Mod(o.��¼����(+), 10) And" & vbNewLine & _
                "      Nvl(o.��¼״̬, 0) In (0, 1) And b.������Դ = [11] "
    Else
        strSQL = " Select distinct /*+ rule */ B.ID as ҽ��ID,a.��ҳid, B.���id, G.��ɫ As �Թ���ɫ, D.���� As �ɼ���ʽ, B.ҽ������, C.��������,C.����ʱ��,c.�걾�ͳ�ʱ�� as �ͼ�ʱ��, " & vbCrLf & _
                 " H.���� As ִ�п���, B.����ҽ��,B.����ʱ��, C.������, C.����ʱ��, G.���� as �Թܱ���,b.�걾��λ as �걾, " & vbCrLf & _
                 " I.���� as ��������,I.�Ա�,decode(nvl(a.Ӥ��,0),0,i.����,decode(sysdate-k.����ʱ��,'','',round(sysdate-k.����ʱ��)||'��')) as ����," & vbCrLf & _
                 "i.��ǰ���� as ����,decode(b.������Դ,1,I.�����,2,i.סԺ��) as ��ʶ��,K.Ӥ������,K.Ӥ���Ա�, " & vbCrLf & _
                 " L.���� as �������ڿ���,Decode(C.ִ��״̬,2,'����') as ����,I.����ID,c.������,c.�ͼ���,G.��Ѫ��,G.���� as �Թ�����, " & vbCrLf & _
                 " DECODE(B.������־,1,'����','') as ����,decode(b.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���') as ������Դ, " & vbCrLf & _
                 " b.Ӥ��,N.���� as ����,J.���� as ���˿���,C.����ʱ��,C.������,b.������ĿID,C.ִ��״̬,O.��¼����,O.��¼״̬ " & vbCrLf & _
                 " From ����ҽ����¼ A, ����ҽ����¼ B, ����ҽ������ C, ������ĿĿ¼ D, ������ĿĿ¼ E, " & vbCrLf & _
                 " ��Ѫ������ G,���ű� H, ������Ϣ I,���ű� L,����ҽ������ M,���ű� J,������������¼ K , " & vbCrLf & _
                 " (select ������ĿID,���� from ������Ŀ���� where ���� = 9 and ���� = 1 ) N,סԺ���ü�¼ O " & vbCrLf & _
                 " Where A.ID = B.���id And B.ID = C.ҽ��id And A.������Ŀid = D.ID And B.������Ŀid = E.ID " & vbCrLf & _
                 " And E.��� = 'C' And E.�Թܱ��� = G.���� And B.ִ�п���id = H.ID " & vbCrLf & _
                 " And D.��� = 'E' And D.�������� = '6' And A.����id = [1] And c.����ʱ��+0 Between [3] and [4] " & vbCrLf & _
                 " And B.����ID = I.����ID and I.��ǰ����ID = L.ID(+) " & vbCrLf & _
                 " and A.����ID=K.����ID(+) and A.��ҳID=K.��ҳID(+) and A.Ӥ��=k.���(+)" & vbCrLf & _
                 " and a.id = m.ҽ��id And E.id = N.������ĿID(+) And b.���˿���id = J.id(+)  " & vbCrLf & _
                 " and c.ҽ��id = O.ҽ�����(+) and c.��¼���� =Mod(O.��¼����(+),10) and nvl(O.��¼״̬,0) in (0,1) And b.������Դ = [11] "
    End If
    
    '���Ҳ���Id
    gstrSql = "select ����ID from �������Ҷ�Ӧ where ����ID = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDeptID)
    If rsTmp.EOF = True Then
        strDeptIDs = mlngDeptID
    Else
        Do Until rsTmp.EOF
            strDeptIDs = strDeptIDs & "," & rsTmp("����ID")
            rsTmp.MoveNext
        Loop
        If strDeptIDs <> "" Then strDeptIDs = Mid(strDeptIDs, 2) & "," & mlngDeptID
    End If
    
    '����
    If mlngSelectBatch <> 0 And TabCtr.Selected.Index > 0 Then
        strSQL = strSQL & " and c.�������� = [7] "
    End If
    
    'ִ�п���
    If chkDept.Value = 1 Then
        strSQL = strSQL & " and b.ִ�п���ID + 0 in (Select * From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist))) "
    Else
        strSQL = strSQL & " and a.ִ�п���ID + 0 in (Select * From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist))) "
    End If
    
    '���˿���
    If rptPlist.FocusedRow Is Nothing Then
    Else
        strPatiDept = rptPlist.FocusedRow.Record(mPcol.���˿���).Value
        strSQL = strSQL & " and J.���� = [12] "
    End If
    
    '���û��Ȩ�ޣ��Ͳ��ܿ���δ�����걾
    If InStr(mstrPrivs, "ǿ�ƵǼ�δ�����걾") = 0 Then
        strSQL = strSQL & " and c.������ is not null "
    End If
    
    '�������ֲ�ͬ��״̬
    If intState = 0 Then
        strSQL = strSQL & " And C.ִ��״̬ in(0) and C.������ is null " & vbCrLf
    ElseIf intState = 1 Or intState = 4 Then
        strSQL = strSQL & " And C.ִ��״̬ in(0) and C.������ is not null " & vbCrLf
    ElseIf intState = 2 Then
        strSQL = strSQL & " And C.ִ��״̬ in (1,3) " & vbCrLf
    ElseIf intState = 3 Then
        strSQL = strSQL & " And C.ִ��״̬ in (2) " & vbCrLf
    End If
    
    '����
    If Me.rptPlist.Tag <> "" Or strTmp <> "" Then
        If varFilter(mFilter.�걾) <> "���б걾" Then
            strSQL = strSQL & " and A.�걾��λ = [5] "
        End If
        
        If varFilter(mFilter.�ɼ���ʽ) <> 0 Then
            strSQL = strSQL & " and d.ID + 0 = [6] "
        End If
        
        If Me.rptPlist.Tag <> "" Then
            strDateBegin = varFilter(mFilter.��ʼʱ��)
            strDateEnd = varFilter(mFilter.����ʱ��)
        Else
            strDateBegin = zlDatabase.Currentdate - Val(varFilter(mFilter.���ʱ��))
            strDateEnd = zlDatabase.Currentdate
        End If
    Else
        strDateBegin = zlDatabase.Currentdate - 3
        strDateEnd = zlDatabase.Currentdate
    End If
    
    If blnDateMoved Then
        strSQL1 = strSQL
        strSQL1 = Replace(strSQL1, "����ҽ����¼", "H����ҽ����¼")
        strSQL1 = Replace(strSQL1, "����ҽ������", "H����ҽ������")
        strSQL = strSQL & " Union ALL " & strSQL1
    End If
    
    If blOnlyWhere = True Then
        If chkDept.Value = 1 Then
            strSQL = "Select Distinct /*+ rule */ b.Id As ҽ��id, a.��ҳid, b.���id, g.��ɫ As �Թ���ɫ, d.���� As �ɼ���ʽ, b.ҽ������, c.��������, c.����ʱ��," & vbNewLine & _
                    "                c.�걾�ͳ�ʱ�� As �ͼ�ʱ��, h.���� As ִ�п���, b.����ҽ��, b.����ʱ��, c.������, c.����ʱ��, g.���� As �Թܱ���, b.�걾��λ As �걾, i.���� As ��������," & vbNewLine & _
                    "                i.�Ա�," & vbNewLine & _
                    "                Decode(Nvl(a.Ӥ��, 0), 0, i.����, Decode(Sysdate - k.����ʱ��, '', '', Round(Sysdate - k.����ʱ��) || '��')) As ����," & vbNewLine & _
                    "                i.��ǰ���� As ����, Decode(b.������Դ, 1, i.�����, 2, i.סԺ��) As ��ʶ��, k.Ӥ������, k.Ӥ���Ա�, l.���� As �������ڿ���," & vbNewLine & _
                    "                Decode(c.ִ��״̬, 2, '����') As ����, i.����id, c.������, c.�ͼ���, g.��Ѫ��, g.���� As �Թ�����," & vbNewLine & _
                    "                Decode(b.������־, 1, '����', '') As ����, Decode(b.������Դ, 1, '����', 2, 'סԺ', 3, 'Ժ��', 4, '���') As ������Դ, b.Ӥ��," & vbNewLine & _
                    "                n.���� As ����, j.���� As ���˿���, c.����ʱ��, c.������, b.������Ŀid, c.ִ��״̬, o.��¼����, o.��¼״̬, a.�Һŵ�" & vbNewLine & _
                    "From ����ҽ����¼ A, ����ҽ����¼ B, ����ҽ������ C, ������ĿĿ¼ D, ������ĿĿ¼ E, ��Ѫ������ G, ���ű� H, ������Ϣ I, ���ű� L, ���ű� J, ������������¼ K," & vbNewLine & _
                    "     (Select ������Ŀid, ���� From ������Ŀ���� Where ���� = 9 And ���� = 1) N, ������ü�¼ O" & vbNewLine & _
                    "Where a.Id = b.���id And b.Id = c.ҽ��id And a.������Ŀid = d.Id And b.������Ŀid = e.Id And e.��� = 'C' And e.�Թܱ��� = g.���� And" & vbNewLine & _
                    "      b.ִ�п���id = h.Id And d.��� = 'E' And d.�������� = '6' And a.����id = [1] And" & vbNewLine & _
                    "      b.����id = i.����id And i.��ǰ����id = l.Id(+) And a.����id = k.����id(+) And a.��ҳid = k.��ҳid(+) And a.Ӥ�� = k.���(+) And" & vbNewLine & _
                    "      e.Id = n.������Ŀid(+) And i.��ǰ����id = j.Id(+) And c.ҽ��id = o.ҽ�����(+) And c.��¼���� = Mod(o.��¼����(+), 10) And" & vbNewLine & _
                    "      Nvl(o.��¼״̬, 0) In (0, 1) "
        Else
            strSQL = " Select distinct /*+ rule */ B.ID as ҽ��ID,a.��ҳid, B.���id, G.��ɫ As �Թ���ɫ, D.���� As �ɼ���ʽ, B.ҽ������, C.��������,C.����ʱ��,c.�걾�ͳ�ʱ�� as �ͼ�ʱ��, " & vbCrLf & _
                 " H.���� As ִ�п���, B.����ҽ��,B.����ʱ��, C.������, C.����ʱ��, G.���� as �Թܱ���,b.�걾��λ as �걾, " & vbCrLf & _
                 " I.���� as ��������,I.�Ա�,decode(nvl(a.Ӥ��,0),0,i.����,decode(sysdate-k.����ʱ��,'','',round(sysdate-k.����ʱ��)||'��')) as ����," & vbCrLf & _
                 " i.��ǰ���� as ����,decode(b.������Դ,1,I.�����,2,i.סԺ��) as ��ʶ��, K.Ӥ������,K.Ӥ���Ա�, " & vbCrLf & _
                 " L.���� as �������ڿ���,Decode(C.ִ��״̬,2,'����') as ����,I.����Id,c.������,c.�ͼ���,G.��Ѫ��,G.���� as �Թ�����, " & vbCrLf & _
                 " DECODE(B.������־,1,'����','') as ����,decode(b.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���') as ������Դ,b.Ӥ��,N.���� as ����,J.���� as ���˿���,C.����ʱ��,C.������, " & vbCrLf & _
                 " b.������ĿID,C.ִ��״̬,O.��¼����,O.��¼״̬,a.�Һŵ� " & vbCrLf & _
                 " From ����ҽ����¼ A, ����ҽ����¼ B, ����ҽ������ C, ������ĿĿ¼ D, ������ĿĿ¼ E, " & vbCrLf & _
                 " ��Ѫ������ G,���ű� H, ������Ϣ I,���ű� L,����ҽ������ M,���ű� J,������������¼ K , " & vbCrLf & _
                 " (select ������ĿID,���� from ������Ŀ���� where ���� = 9 and ���� = 1 ) N,סԺ���ü�¼ O " & vbCrLf & _
                 " Where A.ID = B.���id And B.ID = C.ҽ��id And A.������Ŀid = D.ID And B.������Ŀid = E.ID " & vbCrLf & _
                 " And E.��� = 'C' And E.�Թܱ��� = G.���� And m.ִ�в���id = H.ID " & vbCrLf & _
                 " And D.��� = 'E' And D.�������� = '6' And A.����id = [1] And " & vbCrLf & _
                 " B.����ID = I.����ID and I.��ǰ����ID = L.ID(+) " & _
                 " and A.����ID=K.����ID(+) and A.��ҳID=K.��ҳID(+) and A.Ӥ��=k.���(+)" & vbCrLf & _
                 " And a.id  = m.ҽ��id And E.id = N.������ĿID(+) And I.��ǰ����id = J.id(+)  " & vbCrLf & _
                 " and c.ҽ��id = O.ҽ�����(+) and c.��¼���� =mod(O.��¼����(+),10) and nvl(O.��¼״̬,0) in (0,1) "
        End If
        
        '�������ֲ�ͬ��״̬
        If intState = 0 Or intState = 4 Then
            strSQL = strSQL & " And C.ִ��״̬ in(0) and C.������ is null " & vbCrLf
        ElseIf intState = 1 Then
            strSQL = strSQL & " And C.ִ��״̬ in(1,0,3) and C.������ is not null " & vbCrLf
        ElseIf intState = 2 Then
            strSQL = strSQL & " And C.ִ��״̬ in(1,3) and C.������ is not null " & vbCrLf
        ElseIf intState = 3 Then
            strSQL = strSQL & " And C.ִ��״̬ in(2) " & vbCrLf
        End If
        If IDKind.IDKind = IDKinds.C0���� And BlnIsNumber(txtGoto) Then
            strSQL = strSQL & " And c.�������� = [10]   "
        Else
            '�����жϲ�����סԺ��������
            gstrSql = "Select ��ҳid, ��Ժ���� From ������ҳ Where ����id = [1] Order By ��ҳid Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngPatientID)
            If rsTmp.EOF = False Then
                If Nvl(rsTmp("��Ժ����")) = "" Then
                    int��ҳID = Nvl(rsTmp("��ҳId"), 0)
                    strSQL = strSQL & " And a.��ҳid = [8] "
                Else
                    gstrSql = "Select NO " & vbNewLine & _
                            " From ���˹Һż�¼ A, ����ҽ����¼��b " & vbNewLine & _
                            " Where A.����id = B.����id And B.������Դ = 1 And a.��¼״̬=1 and a.��¼���� =1 and A.����id = [1] " & vbNewLine & _
                            " Order By A.ID Desc "
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngPatientID)
                    If rsTmp.EOF = False Then
                        str�Һŵ� = Nvl(rsTmp("NO"))
                        'strSQL = strSQL & " And A.�Һŵ� = [9] "
                    Else
                        strSQL = strSQL & " And nvl(c.��������,0) <> 0   "
                        'strSQL = strSQL & " And a.������Դ<> 2 "
                    End If
                End If
            Else
                gstrSql = "Select NO " & vbNewLine & _
                        " From ���˹Һż�¼ A, ����ҽ����¼��b " & vbNewLine & _
                        " Where A.����id = B.����id And B.������Դ = 1 And a.��¼״̬ =1 and a.��¼���� =1  and  A.����id = [1] " & vbNewLine & _
                        " Order By A.ID Desc "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngPatientID)
                If rsTmp.EOF = False Then
                    str�Һŵ� = Nvl(rsTmp("NO"))
                    'strSQL = strSQL & " And A.�Һŵ� = [9] "
                Else
                    strSQL = strSQL & " And a.������Դ<> 2 "
                End If
            End If
            If IDKind.IDKind = IDKinds.C0���� And BlnIsNumber(txtGoto) Then
                strSQL = strSQL & " And c.�������� = [10]   "
            End If
        End If
    End If
    
    If intPatientType <> 2 Then
        strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
    End If
    
'    strSQLbak = strSQL
'    strSQLbak = Replace$(strSQLbak, "סԺ���ü�¼", "������ü�¼")
'    strSQL = strSQL & " union all " & strSQLbak
    
    strSQL = strSQL & IIf(intBabyNo <> 0, " and nvl(a.Ӥ��,0)= " & intBabyNo, IIf(blOnlyWhere, "", "and nvl(a.Ӥ��,0) = 0 ")) & " Order By ��������, �Թܱ���, ���id, ҽ��id, �걾, ����ʱ�� "
    
    If strTmp <> "" Or rptPlist.Tag <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientID, strDeptIDs, CDate(Format(strDateBegin, "yyyy-mm-dd 00:00:00")), _
                CDate(Format(strDateEnd, "yyyy-mm-dd 23:59:59")), varFilter(mFilter.�걾), CLng(Val(varFilter(mFilter.�ɼ���ʽ))), mlngSelectBatch, _
                int��ҳID, str�Һŵ�, txtGoto, intPatientType, strPatiDept)
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientID, strDeptIDs, CDate(Format(strDateBegin, "yyyy-mm-dd 00:00:00")), _
                CDate(Format(strDateEnd, "yyyy-mm-dd 23:59:59")), "", 0, mlngSelectBatch, int��ҳID, str�Һŵ�, txtGoto, intPatientType, strPatiDept)
    End If
    
    Me.rptAlist(TabCtr.Selected.Index).Records.DeleteAll

    Me.rptCuvette.Records.DeleteAll
    Me.rptAlist(TabCtr.Selected.Index).Columns(mAcol.��ִ��).Visible = False
    If rsTmp.RecordCount < 1 Then
        Me.txt�ͼ��� = ""
    End If
    Do Until rsTmp.EOF
        If intState = 0 Then
            If intPatientType = 2 Then
                If intMainID = 0 Then
                    intMainID = Val(rsTmp("��ҳID") & "")
                    If intMainID <> 0 Then
                        strGetSql = "Select ��ҳid, ��Ժ���� From ������ҳ Where ����id = [1] and ��ҳid=[2] "
                        Set rsTest = zlDatabase.OpenSQLRecord(strGetSql, Me.Caption, lngPatientID, intMainID)
                        If rsTest.RecordCount > 0 Then
                            If Nvl(rsTest("��Ժ����")) <> "" Then
                                If MsgBox("�ò����ѳ�Ժ���Ƿ����ִ�У�", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                                    Me.rptAlist(TabCtr.Selected.Index).Records.DeleteAll
                                    Me.rptCuvette.Records.DeleteAll
                                    Me.rptAlist(TabCtr.Selected.Index).Populate
                                    Me.rptCuvette.Populate
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        'û�ж�Ӧ��ɫ����Ĳɼ���д��
        If IsNull(rsTmp("�Թ���ɫ")) = False Then
            If strOldAdvice <> rsTmp("���ID") & "" Or strOldNO <> rsTmp("�Թܱ���") & "" Or strOldCodeBar <> rsTmp("��������") & "" Then
                Set Record = Me.rptAlist(TabCtr.Selected.Index).Records.Add
                For intLoop = 0 To Me.rptAlist(TabCtr.Selected.Index).Columns.Count + 1
                    Record.AddItem ""
                Next
                
                If blOnlyWhere = True And (Nvl(rsTmp("ִ��״̬")) = 1 Or Nvl(rsTmp("ִ��״̬")) = 3) Then
                    Me.rptAlist(TabCtr.Selected.Index).Columns(mAcol.��ִ��).Visible = True
                End If
                
                If Nvl(rsTmp("ִ��״̬")) = 1 Or Nvl(rsTmp("ִ��״̬")) = 3 Then
                    Record(mAcol.��ִ��).Value = "��"
                Else
                    Record(mAcol.��ִ��).Value = ""
                End If
                
                Record(mAcol.ѡ��).HasCheckbox = True
                
                If str�Һŵ� = "" Then
                    Record(mAcol.ѡ��).Checked = IIf(Nvl(rsTmp("ִ��״̬")) = 0, True, False)
                Else
                    Record(mAcol.ѡ��).Checked = IIf(Nvl(rsTmp("ִ��״̬")) = 0 And str�Һŵ� = Nvl(rsTmp("�Һŵ�")), True, False)
                End If
                
                Record(mAcol.����).Value = IIf(rsTmp("��¼״̬") = 1, "��", "��")
                
                
                Record(mAcol.ID).Value = Nvl(rsTmp("ҽ��ID")) & "," & Nvl(rsTmp("���ID"))
                
                
                Record(mAcol.ͼ��).BackColor = Val(Nvl(rsTmp("�Թ���ɫ")))
                Record(mAcol.�ɼ���ʽ).Value = Nvl(rsTmp("�ɼ���ʽ"))
                Record(mAcol.ҽ������).Value = Nvl(rsTmp("ҽ������"))
                Record(mAcol.����).Value = Nvl(rsTmp("��������"))
                Record(mAcol.ִ�п���).Value = Nvl(rsTmp("ִ�п���"))
                Record(mAcol.����ҽ��).Value = Nvl(rsTmp("����ҽ��"))
                Record(mAcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                Record(mAcol.������).Value = Nvl(rsTmp("������"))
                Record(mAcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                Record(mAcol.�Թ���ɫ).Value = Nvl(rsTmp("�Թ���ɫ"))
                Record(mAcol.�Թܱ���).Value = Nvl(rsTmp("�Թܱ���"))
                Record(mAcol.�걾).Value = Nvl(rsTmp("�걾")) & IIf(Nvl(rsTmp("Ӥ��")) = 0, "", "(Ӥ��)")
                Record(mAcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                Record(mAcol.������).Value = Nvl(rsTmp("������"))
                Record(mAcol.�ͼ���).Value = Nvl(rsTmp("�ͼ���"))
                Record(mAcol.��Ѫ��).Value = Nvl(rsTmp("��Ѫ��"))
                Record(mAcol.�Թ�����).Value = Nvl(rsTmp("�Թ�����"))
                Record(mAcol.����).Value = Nvl(rsTmp("����"))
                Record(mAcol.������Դ).Value = Nvl(rsTmp("������Դ"))
                Record(mAcol.Ӥ��).Value = Nvl(rsTmp("Ӥ��"))
                Record(mAcol.����).Value = Nvl(rsTmp("����"))
                Record(mAcol.���ID).Value = Nvl(rsTmp("���ID"))
                
                Record(mAcol.����ID).Value = Nvl(rsTmp("����ID"))
                Record(mAcol.����).Value = Nvl(rsTmp("��������"))
                Record(mAcol.�Ա�).Value = Nvl(rsTmp("�Ա�"))
                Record(mAcol.����).Value = Nvl(rsTmp("����"))
                Record(mAcol.��ʶ��).Value = Nvl(rsTmp("��ʶ��"))
                Record(mAcol.����).Value = Nvl(rsTmp("����"))
                Record(mAcol.���˿���).Value = Nvl(rsTmp("���˿���"))
                Record(mAcol.������).Value = Nvl(rsTmp("������"))
                Record(mAcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                Record(mAcol.������ĿID).Value = Nvl(rsTmp("������ĿID"))
                Record(mAcol.ִ��״̬).Value = Nvl(rsTmp("ִ��״̬"))
                Record(mAcol.ҽ��id).Value = Nvl(rsTmp("ҽ��id"))
                Record(mAcol.�ͼ�ʱ��).Value = Nvl(rsTmp("�ͼ�ʱ��"))
                
                
                For intLoop = 0 To Me.rptAlist(TabCtr.Selected.Index).Columns.Count + 1
                    Record(intLoop).ForeColor = Val(Nvl(rsTmp("�Թ���ɫ")))
                Next
                
                If blOnlyWhere = True Then
                    If Record(mAcol.����).Value <> "" And intShowButtom <> 2 Then
                        intShowButtom = 1
                    End If
                    If Record(mAcol.������).Value <> "" Then
                        intShowButtom = 2
                    End If
                End If
            Else
                Record(mAcol.ҽ������).Value = Record(mAcol.ҽ������).Value & " " & Nvl(rsTmp("ҽ������"))
                Record(mAcol.�ϲ�ҽ��).Value = Record(mAcol.�ϲ�ҽ��).Value & ";" & _
                                               Nvl(rsTmp("ҽ��ID")) & "," & Nvl(rsTmp("���ID"))
                Record(mAcol.����).Value = Record(mAcol.����).Value & " " & Nvl(rsTmp("����"))
            End If
            strOldAdvice = rsTmp("���ID") & ""
            strOldNO = rsTmp("�Թܱ���") & ""
            strOldCodeBar = rsTmp("��������") & ""
            If InStr(1, strCuvetteNumber & ",", "," & Nvl(rsTmp("�Թܱ���")) & ",") <= 0 Then
                strCuvetteNumber = strCuvetteNumber & "," & Nvl(rsTmp("�Թܱ���"))
            End If
        End If
        If chkRemberPer.Value = 1 Then
            If Nvl(rsTmp("�ͼ���") & "") <> "" Then
                '�����ͼ��ˡ���ʾ�ͼ���
                txt�ͼ��� = Nvl(rsTmp("�ͼ���") & "")
            Else
                txt�ͼ��� = mstrSendPerson
            End If
        Else
            txt�ͼ��� = Nvl(rsTmp("�ͼ���"))
        End If
        rsTmp.MoveNext
    Loop
    If Me.Visible = True Then
        Me.rptAlist(TabCtr.Selected.Index).Populate
    End If
    
    '�м�¼ʱ��ʾд��ɹ�
    If rptAlist(TabCtr.Selected.Index).Records.Count > 0 Then
        RefreshAdviceData = True
    End If
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        mlngKey = Nvl(rsTmp("����ID"))
    End If
    '��ʹ�ò��˲���ʱ��д������Ϣ
    If blOnlyWhere = True Then
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            txt���� = IIf(IsNull(rsTmp("Ӥ������")), Nvl(rsTmp("��������")), Nvl(rsTmp("Ӥ������")))
            txt����.Tag = IIf(IsNull(rsTmp("Ӥ������")), Nvl(rsTmp("��������")), Nvl(rsTmp("Ӥ������")))
            On Error Resume Next
            txt�Ա� = IIf(IsNull(rsTmp("Ӥ������")), Nvl(rsTmp("�Ա�")), Nvl(rsTmp("Ӥ���Ա�")))
            txt���� = Nvl(rsTmp("����"))
            On Error GoTo 0
            txtBed = Nvl(rsTmp("����"))
            txtID = Nvl(rsTmp("��ʶ��"))
            txtPatientDept = Nvl(rsTmp("�������ڿ���"))
            mlngKey = Nvl(rsTmp("����ID"))
        End If
    Else
        '���ö���
        Select Case Me.TabCtr.Selected.Index
            Case 0
                
            Case 1
                
            Case 2
                
        End Select
    End If
    
    
    If strCuvetteNumber <> "" Then
        With Me.rptCuvette
            strSQL = "select ����,����,��Ӽ�,��Ѫ��,���,��ɫ from ��Ѫ������ where ���� in  " & _
                      "(Select * From Table(Cast(f_str2list([1]) As Zltools.t_strlist)))"
                       
                        
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strCuvetteNumber, 2))
            Do Until rsTmp.EOF
                Set Record = .Records.Add
                For intLoop = 0 To .Columns.Count + 1
                    Record.AddItem ""
                Next
                Record(mCuvette.ѡ��).HasCheckbox = True
                Record(mCuvette.ѡ��).Checked = True
                Record(mCuvette.����).Value = Nvl(rsTmp("����"))
                Record(mCuvette.����).Value = Nvl(rsTmp("����"))
                Record(mCuvette.��Ӽ�).Value = Nvl(rsTmp("��Ӽ�"))
                Record(mCuvette.��Ѫ��).Value = Nvl(rsTmp("��Ѫ��"))
                Record(mCuvette.���).Value = Nvl(rsTmp("���"))
                Record(mCuvette.��ɫ).BackColor = Nvl(rsTmp("��ɫ"))
                
                For intLoop = 0 To .Columns.Count + 1
                    Record(intLoop).ForeColor = Nvl(rsTmp("��ɫ"))
                Next
                rsTmp.MoveNext
            Loop
            
            If InStr(1, Mid(strCuvetteNumber, 2), ",") <= 0 Then
                .Records(0).Item(mCuvette.ѡ��).Checked = True
            End If
        End With
    End If
    
    Me.rptCuvette.Populate
    
    '��ʾ��ǰ��ѯ�ķ���
    If Me.rptAlist(TabCtr.Selected.Index).Records.Count > 0 Then
        With Me.rptAlist(TabCtr.Selected.Index)
            For intLoop = 0 To .Records.Count - 1
                If .Records(intLoop).Item(mAcol.����).Value = "��" Then
                    str���շ�ҽ��ID = str���շ�ҽ��ID & "," & .Records(intLoop).Item(mAcol.ID).Value
                End If
                str����ҽ��ID = str����ҽ��ID & "," & .Records(intLoop).Item(mAcol.ID).Value
            Next
            
            If Mid(str����ҽ��ID, 2) <> "" Then
                gstrSql = "Select /*+ rule */ Sum(ʵ�ս��) As ���ս��" & vbNewLine & _
                            "From סԺ���ü�¼" & vbNewLine & _
                            "Where ҽ����� In (Select * From Table(Cast(f_Num2list([1]) As ZLTOOLS.t_Numlist))) "
                            
                If intPatientType <> 2 Then
                    gstrSql = Replace(gstrSql, "סԺ���ü�¼", "������ü�¼")
                End If
                                            
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, str����ҽ��ID)
                lbl��ʾ���� = "(�ϼ�:Ӧ��<" & rsTmp("���ս��") & "Ԫ> "
            Else
                lbl��ʾ���� = ""
            End If
            If Mid(str���շ�ҽ��ID, 2) <> "" Then
                gstrSql = "Select /*+ rule */ Sum(ʵ�ս��) As ���ս��" & vbNewLine & _
                            "From סԺ���ü�¼" & vbNewLine & _
                            "Where ҽ����� In (Select * From Table(Cast(f_Num2list([1]) As ZLTOOLS.t_Numlist))) "
                            
                If intPatientType <> 2 Then
                    gstrSql = Replace(gstrSql, "סԺ���ü�¼", "������ü�¼")
                End If
                                            
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, str���շ�ҽ��ID)
                
                lbl��ʾ���� = lbl��ʾ���� & "ʵ��<" & rsTmp("���ս��") & "Ԫ>)"
            Else
                If lbl��ʾ���� <> "" Then
                    lbl��ʾ���� = lbl��ʾ���� & "ʵ��<0Ԫ>)"
                End If
            End If
        End With
    Else
        lbl��ʾ���� = ""
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub rptAlist_ItemCheck(Index As Integer, ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim i  As Integer
    If Item.Record(mAcol.ִ��״̬).Value = 1 Or Item.Record(mAcol.ִ��״̬).Value = 3 Then
        Item.Record(mAcol.ѡ��).Checked = False
        Me.rptAlist(TabCtr.Selected.Index).Populate
        MsgBox "��ִ�еı걾����ѡ��!", vbInformation, Me.Caption
    End If
    If Index = 1 Or Index = 0 Then
        'ѡ�񣬻�ȡ��ѡ��ͬһ������Ķ�ͬʱ������
        If Item.Record(mAcol.����).Value <> "" Then
            For i = 0 To rptAlist(Index).Rows.Count - 1
                If rptAlist(Index).Records(i).Item(i).Record(mAcol.����).Value = Item.Record(mAcol.����).Value Then
                    rptAlist(Index).Records(i).Item(i).Record(mAcol.ѡ��).Checked = Item.Record(mAcol.ѡ��).Checked
                End If
            Next
            rptAlist(Index).Redraw
        End If
    End If
End Sub

Private Sub rptAlist_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim hitColumn As ReportColumn
    Dim Record As ReportRecord
    Dim blSelect As Boolean

    With Me.rptAlist(TabCtr.Selected.Index)
        Set hitColumn = .HitTest(X, Y).Column
        
        If Not hitColumn Is Nothing Then
            If hitColumn.Caption = "Check" And .HitTest(X, Y).ht = xtpHitTestHeader Then
                hitColumn.AutoSize = True
                If .Records.Count > 0 Then blSelect = Not .Records(0).Item(mAcol.ѡ��).Checked
                For Each Record In .Records
                    If Record.Item(mAcol.ִ��״̬).Value = 0 Then
                        Record.Item(mAcol.ѡ��).Checked = blSelect
                    Else
                        Record.Item(mAcol.ѡ��).Checked = False
                    End If
                Next
            End If
        End If
        .Populate
    End With
End Sub

Private Sub rptAlist_SelectionChanged(Index As Integer)
    With Me.rptAlist(TabCtr.Selected.Index)
        .PaintManager.HighlightBackColor = .FocusedRow.Record(mAcol.�Թ���ɫ).Value
        .Populate
        If chkRemberPer.Value = 1 Then
            If .FocusedRow.Record(mAcol.�ͼ���).Value <> "" Then
                '�Ѿ����ͼ��˺���ʾ�ͼ���
                txt�ͼ���.Text = .FocusedRow.Record(mAcol.�ͼ���).Value
            Else
                txt�ͼ���.Text = mstrSendPerson
            End If
        Else
            txt�ͼ���.Text = .FocusedRow.Record(mAcol.�ͼ���).Value
        End If
    End With
    RePrintBarCode True
End Sub

Private Sub rptCount_SelectionChanged()
    With Me.rptCount
        txt�ͼ���.Text = .FocusedRow.Record(mAcol.�ͼ���).Value
    End With
End Sub

Private Sub rptCuvette_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
     'ѡ�й���
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
                If .Records.Count > 0 Then blSelect = Not .Records(0).Item(mCuvette.ѡ��).Checked
                For Each Record In .Records
                    Record.Item(mCuvette.ѡ��).Checked = blSelect
                Next
            End If
        End If
        .Populate
        'ѡ�й���
        Call SelectCuvette
    End With
End Sub

Private Sub rptCuvette_SelectionChanged()
    With Me.rptCuvette
        .PaintManager.HighlightBackColor = .FocusedRow.Record(mCuvette.��ɫ).ForeColor
        .Populate
    End With
End Sub

Private Sub rptPlist_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   '�ǼǺ�ȡ���Ǽ�
    If Me.TabCtr.Selected.Index <= 1 Then
        SaveRegister Me.TabCtr.Selected.Index
    End If
End Sub

Private Sub rptPlist_SelectionChanged()
    Dim intPatientType As Integer           '������Դ
    
    With Me.rptPlist.FocusedRow
        mlngKey = .Record(mPcol.����ID).Value
        If .Record(mPcol.��Դ).Value = "����" Then
            intPatientType = 1
        ElseIf .Record(mPcol.��Դ).Value = "סԺ" Then
            intPatientType = 2
        ElseIf .Record(mPcol.��Դ).Value = "Ժ��" Then
            intPatientType = 3
        ElseIf .Record(mPcol.��Դ).Value = "���" Then
            intPatientType = 4
        Else
            intPatientType = 1
        End If
        mintBabyNo = Val(.Record(mPcol.Ӥ����).Value)
    End With
'    lbl��ʾ����.Caption = ""
    'ʹ�ò���IDˢ��ҽ��
    RefreshAdviceData mlngKey, TabCtr.Selected.Index, intPatientType, False, mintBabyNo
    'ˢ����ʾ��Ϣ
    ShowPatientInfo
     Me.cbrthis.FindControl(, conMenu_Edit_ReprintReceipt).Enabled = False
End Sub
Private Sub SelectCuvette()
    '����               ѡ��ѡ�е��Թ�
    
    Dim RecordC As ReportRecord
    Dim RecordA As ReportRecord
    
    For Each RecordC In Me.rptCuvette.Records
        For Each RecordA In Me.rptAlist(TabCtr.Selected.Index).Records
            If RecordA(mAcol.�Թܱ���).Value = RecordC(mCuvette.����).Value And RecordA(mAcol.ִ��״̬).Value = 0 Then
                RecordA(mAcol.ѡ��).Checked = RecordC(mCuvette.ѡ��).Checked
            End If
        Next
    Next

    Me.rptAlist(TabCtr.Selected.Index).Populate
End Sub
Private Function SelectBarCode(strBarCode As String) As Boolean
    Dim RowA As ReportRow
    
    For Each RowA In Me.rptAlist(TabCtr.Selected.Index).Rows
        If RowA.Record(mAcol.����).Value = strBarCode Then
            RowA.Record(mAcol.ѡ��).Checked = True
            SelectBarCode = True
        Else
            RowA.Record(mAcol.ѡ��).Checked = False
        End If
    Next
    Me.rptAlist(TabCtr.Selected.Index).Populate
End Function

Private Sub ShowPatientInfo()
    
    'û�н�����ʱ�˳�
    If Me.rptPlist.FocusedRow Is Nothing Then Exit Sub
    On Error Resume Next
    With Me.rptPlist.FocusedRow
    
        
        txt���� = .Record(mPcol.��������).Value
        txt����.Tag = .Record(mPcol.��������).Value
        txt�Ա� = .Record(mPcol.�Ա�).Value
        txt���� = .Record(mPcol.����).Value
        
        txtBed = .Record(mPcol.����).Value
        txtID = .Record(mPcol.��ʶ��).Value
        txtPatientDept = .Record(mPcol.���˿���).Value
    End With
End Sub

Private Function SaveRegister(intState As Integer) As Boolean
    '����:              �Ǽǻ�ȡ���Ǽ�
    '����:              intState = 0 δ�Ǽ� = 1 �ѵǼ� = 2 ��ִ��
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strAdvice As String
    Dim intLoop As Integer
    Dim cbrControl As CommandBarControl
    Dim Record As ReportRecord
    Dim intTimeLimit As Integer         '�ͼ�ʱ�޵�λ����
    Dim blnTimeLimit As Boolean         '�Ƿ񳬹��ͼ�ʱ�� true = ����
    Dim strҽ��ID As String
    Dim strMsg As String
    Dim strUrgent As String
    Dim strTemp As String
    Dim strValue As String
    Dim intLen As Integer
    
    On Error GoTo errH
        
    Set cbrControl = Me.cbrthis.FindControl(, conMenu_Edit_Insert, True, True)
    
    If cbrControl.Caption = "��ʼ����" Then
        'û�п�ʼ����ʱǿ�п�ʼ
        BeginRegister
    End If
    
    If Me.rptAlist(TabCtr.Selected.Index).Rows.Count <= 0 Then
        MsgBox "û���ҵ����ԵǼǵ�ҽ����¼!", vbQuestion, gstrSysName
        Exit Function
    End If
    
    If intState = 0 Then
        With Me.rptAlist(TabCtr.Selected.Index)
            For intLoop = 0 To .Rows.Count - 1
                If .Rows(intLoop).Record(mAcol.ѡ��).Checked = True And .Rows(intLoop).Record(mAcol.ִ��״̬).Value = 0 Then
                    strҽ��ID = strҽ��ID & "," & .Rows(intLoop).Record(mAcol.ID).Value & "," & .Rows(intLoop).Record(mAcol.�ϲ�ҽ��).Value
                    
                    '�����걾��Ϣ
                    If .Rows(intLoop).Record(mAcol.����).Value = "����" And chkUrgent.Value = 1 Then
                        strUrgent = strUrgent & "," & .Rows(intLoop).Record(mAcol.ҽ������).Value
                    End If
                End If
            Next
        End With
        strҽ��ID = Mid(strҽ��ID, 2)
        If Chk���۷���(Me, strҽ��ID, 0) = False Then
            Exit Function
        End If
    End If
    With Me.rptAlist(TabCtr.Selected.Index)
        
        For intLoop = 0 To .Rows.Count - 1
            If .Rows(intLoop).Record(mAcol.ѡ��).Checked = True And .Rows(intLoop).Record(mAcol.ִ��״̬).Value = 0 Then
                '�����Ƿ񳬹��ɼ�ʱ��
                gstrSql = "select �ͼ�ʱ�� from ������Ŀѡ�� where ������Ŀid = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(.Rows(intLoop).Record(mAcol.������ĿID).Value))
                If rsTmp.EOF = True Then
                    intTimeLimit = 0
                Else
                    intTimeLimit = Val(Nvl(rsTmp("�ͼ�ʱ��")))
                End If
                
                If IsDate(.Rows(intLoop).Record(mAcol.����ʱ��).Value) = False And intTimeLimit > 0 Then
                    blnTimeLimit = True
                Else
                    If IsDate(.Rows(intLoop).Record(mAcol.����ʱ��).Value) = True Then
                        If DateDiff("n", .Rows(intLoop).Record(mAcol.����ʱ��).Value, zlDatabase.Currentdate) > intTimeLimit _
                            And intTimeLimit > 0 Then
                            '�����ͼ�ʱ��
                            blnTimeLimit = True
                        End If
                    Else
                        If intTimeLimit > 0 Then
                            blnTimeLimit = True
                        End If
                    End If
                End If
                If chkComRequest.Value = 1 And .Rows(intLoop).Record(mAcol.�ͼ�ʱ��).Value = "" Then
                    strMsg = strMsg & "��ǰ��" & .Rows(intLoop).Record(mAcol.ҽ������).Value & "��δ�ͼ�,������Ǽǣ�" & vbNewLine
                Else
                    If blnTimeLimit = True And (intState = 0 Or intState = 4) Then
                        '��ʱ�����鿴�Ƿ���Ȩ�ޣ���Ȩ��ʱֻ��ʾ
                        If InStr(mstrPrivs, "ǿ��ͨ���ͼ�ʱ��") > 0 Then
                            '��ʾ
                            MsgBox ("��ǰ�걾�Ĳ���ʱ��Ϊ��" & .Rows(intLoop).Record(mAcol.����ʱ��).Value & "��" & vbCrLf & _
                                    "�ѳ�������ʱ��" & intTimeLimit & "����,�ͼ��ӳ٣�")
                            strAdvice = strAdvice & "|" & .Rows(intLoop).Record(mAcol.ID).Value & _
                                Replace(.Rows(intLoop).Record(mAcol.�ϲ�ҽ��).Value, ";", "|")
                        Else
                            '�ܾ��Ǽ�
                            MsgBox ("��ǰ�걾�Ĳ���ʱ��Ϊ��" & .Rows(intLoop).Record(mAcol.����ʱ��).Value & "��" & vbCrLf & _
                                    "�ѳ�������ʱ��" & intTimeLimit & "����,������Ǽǣ�")
                        End If
                        
                    ElseIf .Rows(intLoop).Record(mAcol.����ʱ��).Value = "" And intState = 0 Then
                        '����ǿ�ƵǼ�δ�����걾
                        If InStr(mstrPrivs, "ǿ�ƵǼ�δ�����걾") > 0 Then
                            strAdvice = strAdvice & "|" & .Rows(intLoop).Record(mAcol.ID).Value & _
                                Replace(.Rows(intLoop).Record(mAcol.�ϲ�ҽ��).Value, ";", "|")
                        Else
                            '�ܾ��Ǽ�
                            MsgBox "��ǰ��" & .Rows(intLoop).Record(mAcol.ҽ������).Value & "��δ����,������Ǽǣ�", vbInformation
                        End If
                    Else
                        strAdvice = strAdvice & "|" & .Rows(intLoop).Record(mAcol.ID).Value & _
                                Replace(.Rows(intLoop).Record(mAcol.�ϲ�ҽ��).Value, ";", "|")
                    End If
                End If
            End If
        Next
        If strMsg <> "" Then MsgBox strMsg, vbInformation, "�걾ǩ��"
    End With
    If strAdvice <> "" Then
    
        If intState = 0 Then
            '���tat��ʱ
            If getTATTime(strAdvice) = False Then
                Exit Function
            End If
            If strAdvice = "" Then
                Exit Function
            End If
 
        End If
        
        '��ʾ�����걾
        If strUrgent <> "" And chkUrgent.Value = 1 Then
            If UBound(Split(strUrgent, ",")) > 2 Then
                MsgBox "��" & Split(strUrgent, ",")(1) & "," & Split(strUrgent, ",")(2) & ",......���ǼǱ걾Ϊ�����걾��", vbInformation, "�걾�Ǽ�"
            Else
                MsgBox "��" & Mid(strUrgent, 2) & "���ǼǱ걾Ϊ�����걾��", vbInformation, "�걾�Ǽ�"
            End If
        End If
        
        '�ǼǺ�ȡ���Ǽ�
        If Len(Mid(strAdvice, 2)) > 2000 Then
            strTemp = Mid(strAdvice, 2)
            Do While Len(strTemp) > 2000
                strValue = Mid(strTemp, 1, 2000)
                intLen = InStrRev(strValue, "|")
                strTemp = Mid(strValue, intLen + 1) & Mid(strTemp, 2001)
                strValue = Mid(strValue, 1, intLen - 1)
                
                strSQL = "Zl_����ҽ������_SampleInput('" & strValue
                If intState = 0 Or intState = 3 Or intState = 4 Then
                    strSQL = strSQL & "','" & UserInfo.���� & "'," & mlngBatch & ",'" & UserInfo.��� & "','" & UserInfo.���� & "','" & Trim(txt�ͼ���.Text) & "')"
                ElseIf intState = 1 Then
                    strSQL = strSQL & "',NULL," & mlngBatch & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
                End If
                zlDatabase.ExecuteProcedure strSQL, gstrSysName
            Loop
            
            If strTemp <> "" Then
                strSQL = "Zl_����ҽ������_SampleInput('" & strTemp
                If intState = 0 Or intState = 3 Or intState = 4 Then
                    strSQL = strSQL & "','" & UserInfo.���� & "'," & mlngBatch & ",'" & UserInfo.��� & "','" & UserInfo.���� & "','" & Trim(txt�ͼ���.Text) & "')"
                ElseIf intState = 1 Then
                    strSQL = strSQL & "',NULL," & mlngBatch & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
                End If
                zlDatabase.ExecuteProcedure strSQL, gstrSysName
            End If
        Else
            strSQL = "Zl_����ҽ������_SampleInput('" & Mid(strAdvice, 2)
            If intState = 0 Or intState = 3 Or intState = 4 Then
                strSQL = strSQL & "','" & UserInfo.���� & "'," & mlngBatch & ",'" & UserInfo.��� & "','" & UserInfo.���� & "','" & Trim(txt�ͼ���.Text) & "')"
            ElseIf intState = 1 Then
                strSQL = strSQL & "',NULL," & mlngBatch & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
            End If
            zlDatabase.ExecuteProcedure strSQL, gstrSysName
        End If
        
        SaveRegister = True
        mblnUse = True
        
        If intState = 0 Or intState = 3 Or intState = 4 Then
            Call WriterCheckSampleToLIS(Mid(strAdvice, 2), UserInfo.����, mlngBatch, Trim(Me.txt�ͼ���.Text))
        ElseIf intState = 1 Then
            Call WriterCheckSampleToLIS(Mid(strAdvice, 2), "", 0)
        End If
        
        'û��ҽ����¼ʱ�˳�
        If Me.rptAlist(TabCtr.Selected.Index).Rows.Count = 0 Then Exit Function
        
        If intState = 0 Or intState = 4 Then
            '����
            Call InsrOrDelAdvice(1, Replace(Mid(strAdvice, 2), "|", ","))
        Else
            'ȡ��
            Call InsrOrDelAdvice(0, Replace(Mid(strAdvice, 2), "|", ","))
        End If
    End If
    
    Me.rptAlist(TabCtr.Selected.Index).Records.DeleteAll
    Me.rptCuvette.Records.DeleteAll
    Me.rptAlist(TabCtr.Selected.Index).Populate
    Me.rptCuvette.Populate
    txt����.Text = ""
    txt�Ա�.Text = ""
    txt����.Text = ""
    txtBed.Text = ""
    txtID.Text = ""
    txtPatientDept.Text = ""
    
    SaveRegister = True
    
    Exit Function
errH:
    
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub TabCtr_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim intPatientType As Integer           '������Դ
    
    If Me.Visible = False Then Exit Sub
    Dim Controlcbo As CommandBarComboBox                '���οؼ�
    Set Controlcbo = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_File_RoomSet, True, True)
    Select Case Item.Index
        Case 0
            Me.cmdOK.Enabled = True
            Me.cmdOK.Caption = "�Ǽ�(&G)"
            Controlcbo.Clear
            Controlcbo.AddItem "��������"
            Controlcbo.ItemData(Controlcbo.ListCount) = 0
            Controlcbo.ListIndex = 1
            mlngSelectBatch = 0
        Case 1
            Me.cmdOK.Enabled = True
            Me.cmdOK.Caption = "ȡ��(&G)"
        Case 2
            Me.cmdOK.Enabled = False
            Me.cmdOK.Caption = "ȡ��(&G)"
        Case 3
            Me.cmdOK.Enabled = False
            Me.cmdOK.Caption = "ȡ��(&G)"
        Case 4
            Me.cmdOK.Enabled = False
            Me.cmdOK.Caption = "�Ǽ�(&G)"
    End Select
    RefreshPatientData 1, mintBabyNo
    Me.cbrthis.FindControl(, conMenu_Edit_ReprintReceipt).Enabled = False
    Me.stbThis.Panels(2).Text = "��ǰ���У�" & Me.rptPlist.Rows.Count & "�����ˣ�"
End Sub

Private Sub txtGoto_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtGoto.Text = "" And Me.ActiveControl Is txtGoto)
End Sub

Private Sub txtGoto_GotFocus()
    txtGoto.SelStart = 0
    txtGoto.SelLength = Len(txtGoto.Text)
    If Not mobjIDCard Is Nothing And txtGoto.Text = "" And Not txtGoto.Locked Then mobjIDCard.SetEnabled (True)
End Sub

Private Sub txtGoto_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub txtGoto_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim blFind As Boolean                           '�Ƿ���ҳɹ�
    Dim Row As ReportRow                            '�б��ж���
    Dim strFind As String                           '�ٽ������ִ�
    Dim blnCard As Boolean
    Dim str���շ�ҽ��ID As String                   '���շѵ�ҽ��ID��","�ָ�
    Dim str����ҽ��ID As String                     '���е�ҽ��ID��","�ָ�
    Dim intLoop As Integer
    Dim lng�����ID As Long
    Dim lng����ID As Long
    
    On Error GoTo errH
    
    If CheckIsInclude(UCase(Chr(KeyAscii)), "'����;��:��?��|,����""") = True Then KeyAscii = 0
    
    If IDKind.IDKind = IDKinds.C0���� Then
'        blnCard = zlCommFun.InputIsCard(txtGoto, KeyAscii, False)
    End If
    
    blnCard = False
    
    If IDKind.IDKind = IDKinds.C5���￨ Then
'        Call zlCommFun.InputIsCard(txtGoto, KeyAscii, True)
        gbytCardNOLen = Val(IDKind.GetKindItem("���ų���", IDKind.IDKind))
        blnCard = KeyAscii <> 8 And Len(txtGoto.Text) = gbytCardNOLen - 1 And txtGoto.SelLength <> Len(txtGoto.Text)
        If blnCard = True Then
            If KeyAscii <> 13 Then
                Me.txtGoto = Me.txtGoto & Chr(KeyAscii)
            End If
            KeyAscii = 0
        End If
    End If
    
    If KeyAscii = 13 Or (IDKind.IDKind = IDKinds.C5���￨ And blnCard = True) Then
    
        '������ٶ���
        
        If mstrFirstBarCode <> txtGoto.Text Then
            txt���� = ""
            txt����.Tag = ""
            txt�Ա� = ""
            txt���� = ""
            txtBed = ""
            txtID = ""
            lbl��ʾ����.Caption = ""
            txtPatientDept = ""
            Me.rptAlist(TabCtr.Selected.Index).Records.DeleteAll
            Me.rptCuvette.Records.DeleteAll
            Me.rptAlist(TabCtr.Selected.Index).Populate
            Me.rptCuvette.Populate
        End If
        
        Select Case Mid(txtGoto, 1, 1)
            Case "-"                                '����ID
                blFind = RefreshAdviceData(Mid(txtGoto, 2), Me.TabCtr.Selected.Index, 0, True)
                strFind = Val(Mid(txtGoto, 2))
            Case "+"                                'סԺ��
                strSQL = "select ����ID from ������Ϣ where סԺ�� = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, Val(Mid(txtGoto, 2)))
                If rsTmp.EOF = False Then blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 2, True)
                strFind = Mid(txtGoto, 2)
            Case "*"                                '�����
                strSQL = "select ����ID from ������Ϣ where ����� = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, Val(Mid(txtGoto, 2)))
                If rsTmp.EOF = False Then blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 1, True)
                strFind = Mid(txtGoto, 2)
            Case "."                                '�Һŵ���
                strSQL = "select ����ID from ����ҽ����¼��where �Һŵ� = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, Mid(txtGoto, 2))
                If rsTmp.EOF = False Then blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 1, True)
                strFind = Mid(txtGoto, 2)
            Case "/"                                '�շѵ��ݺ�
                strSQL = "select distinct ����ID from ������ü�¼ where No = [1] and ����id is not null and �����־ = 1 order by ����ID desc"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, zlCommFun.GetFullNO(Mid(txtGoto, 2)))
                If rsTmp.EOF = False Then blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 1, True)
                strFind = Mid(txtGoto, 2)
            Case Else                               '���￨������
                strFind = txtGoto
                If IDKind.IDKind = IDKinds.C0���� And BlnIsNumber(txtGoto) Then
                    strSQL = "select a.����id,a.������Դ from ����ҽ����¼ a , ����ҽ������ b " & _
                         " Where a.ID = b.ҽ��id And b.�������� = [1]  order by a.����ʱ�� desc     "
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, txtGoto)
                    If rsTmp.EOF = False Then
                        If mstrFirstBarCode = txtGoto.Text Then
                            mstrFirstBarCode = ""
                            '���õǼ�
                            Call LocationObj(txtGoto)
                            Call cmdOK_Click
                            
                            Exit Sub
                        Else
                            mstrFirstBarCode = txtGoto.Text
                            
                        End If
                        blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, Nvl(rsTmp(1), 0), True)
                    End If
                Else
                    If blnCard Or IDKind.IDKind = IDKinds.C5���￨ Then
                        strSQL = "select ����ID from ������Ϣ where ���￨�� = [1] "
                        strFind = UCase(txtGoto)
                    ElseIf IDKind.IDKind = IDKinds.C0���� Then
                        strSQL = "select ����ID from ������Ϣ where ���� = [1] "
                    ElseIf IDKind.IDKind = IDKinds.C1ҽ���� Then
                        strSQL = "select ����ID from ������Ϣ where ����ID = [1] "
                        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("ȫ��"), txtGoto, False, lng����ID) = False Then lng����ID = 0
                        strFind = lng����ID
                    ElseIf IDKind.IDKind = IDKinds.C2���֤�� Then
                        strSQL = "select ����ID from ������Ϣ where ����ID = [1] "
                        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("ȫ��"), txtGoto, False, lng����ID) = False Then lng����ID = 0
                        strFind = lng����ID
                    ElseIf IDKind.IDKind = IDKinds.C3IC���� Then
                        strSQL = "select ����ID from ������Ϣ where ����ID = [1] "
                        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("ȫ��"), txtGoto, False, lng����ID) = False Then lng����ID = 0
                        strFind = lng����ID
                    ElseIf IDKind.IDKind = IDKinds.C4����� Then
                        strSQL = "select ����ID from ������Ϣ where ����ID = [1] "
                        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("ȫ��"), txtGoto, False, lng����ID) = False Then lng����ID = 0
                        strFind = lng����ID
                    Else
                        If Val(IDKind.GetKindItem("�����ID")) <> 0 Then
                            lng�����ID = Val(IDKind.GetKindItem("�����ID"))
                            If mobjSquareCard.zlGetPatiID(lng�����ID, txtGoto, False, lng����ID) = False Then lng����ID = 0
                            If lng����ID = 0 Then lng����ID = 0
                        Else
                            If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("ȫ��"), txtGoto, False, lng����ID) = False Then lng����ID = 0
                        End If
                        strSQL = "select ����ID from ������Ϣ where ����ID = [1] "
                        strFind = lng����ID
                    End If
                  
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, strFind)
                    If rsTmp.EOF = False Then
                        blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 1, True)
                    End If
                End If
        End Select
        '���������������һ������
        If BlnIsNumber(txtGoto) = False Then
            mstrFirstBarCode = ""
        End If
        
        'û���ҵ�����ʱ����ʾ��Ϣ
        If blFind = False Then
            If IDKind.IDKind = IDKinds.C0���� And BlnIsNumber(txtGoto) Then
                '������ʱ�ж�һ�������״̬
                gstrSql = " Select b.ִ��״̬, b.������, b.����ʱ��, b.������, b.����ʱ��, b.�걾�ͳ�ʱ�� From ����ҽ����¼ a, ����ҽ������ b " & _
                         " Where a.id = b.ҽ��id and a.���id is not null and  b.�������� = [1] order by  a.����ʱ�� desc  "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.txtGoto)
                If rsTmp.EOF = True Then
                    MsgBox "û���ҵ�����<" & Me.txtGoto & ">!" & vbCrLf & _
                        "������ܱ�ȡ���󶨻���ҽ��������!", vbInformation, Me.Caption
                Else
                    If rsTmp("ִ��״̬") = 1 Or rsTmp("ִ��״̬") = 3 Then
                        MsgBox "��������Ϊ<" & Me.txtGoto & ">�ѱ�����!"
                    Else
                        If rsTmp("������") <> "" Then
                            MsgBox "��������<" & Me.txtGoto & ">�ѵǼ�  " & vbCrLf & _
                                  "�Ǽ�ʱ��<" & rsTmp("����ʱ��") & ">" & vbCrLf & _
                                  "�Ǽ���<" & rsTmp("������") & ">"
                        End If
                    End If
                End If
                
            End If
        Else
            If Me.TabCtr.Selected.Index = 4 Then
                Me.cmdOK.Enabled = True
            End If

        End If
        
        If Me.rptAlist(TabCtr.Selected.Index).Rows.Count > 0 And Me.cmdOK.Enabled = True Then
            If ChkBarCodeRegister.Value = 1 Then
                mstrFirstBarCode = ""
                '���õǼ�
                Call LocationObj(txtGoto)
                Call cmdOK_Click
            Else
                cmdOK.SetFocus
            End If
        Else
            Me.txtGoto.Text = ""
            Me.txtGoto.SetFocus
        End If
        If mstrFirstBarCode <> "" Then
            Call LocationObj(txtGoto)
        End If
        On Error Resume Next
        '�������ڶ�λ,�����ͺ���
        
'        If blFind = True Then
'            If Mid(txtGoto, 1, 1) <> "-" Then
'                rsTmp.MoveFirst
'                For Each Row In Me.rptPlist(0).Rows
'                    If Row.Record(mPcol.����ID).Value = Nvl(rsTmp(0)) Then
'                        Me.rptPlist(0).FocusedRow = Row
'                        Me.rptPlist(0).Populate
'                    End If
'                Next
'            Else
'                For Each Row In Me.rptPlist(0).Rows
'                    If Row.Record(mPcol.����ID).Value = Mid(txtGoto, 2) Then
'                        Me.rptPlist(0).FocusedRow = Row
'                        Me.rptPlist(0).Populate
'                    End If
'                Next
'            End If
'            Me.txtGoto.Text = ""
'            Me.txtGoto.SetFocus
'        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If Me.rptPlist.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '�������ݱ��
    If zlReportToVSFlexGrid(Me.vgdList, Me.rptPlist) = False Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vgdList
    objPrint.Title.Text = "���˽����嵥"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & zlDatabase.Currentdate)
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub BeginRegister()
    '����           ��ʼ�ǼǱ걾���������µ�����
   
    Dim rsTmp As New ADODB.Recordset
    Dim cbrControl As CommandBarControl
    
    Set cbrControl = Me.cbrthis.FindControl(, conMenu_Edit_Insert, True, True)
    
    If cbrControl.Caption = "��ʼ����" Then
        If Me.rptCount.Records.Count > 0 Then
            If MsgBox("�Ƿ�ֹͣ��ǰ������ʼ�µļ�����", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                Exit Sub
            End If
        End If
        cbrControl.Caption = "��������"
        
        'û��ʹ������ʱ������
        If mblnUse = True Or mlngBatch = 0 Then
            '�õ�һ���µ�����
            gstrSql = "select ����ҽ������_��������.nextval from dual "
            zlDatabase.OpenRecordset rsTmp, gstrSql, Me.Caption
            mlngBatch = rsTmp(0)
            mblnUse = False
        End If
        Me.txt�ͼ���.Enabled = True
    Else
        cbrControl.Caption = "��ʼ����"
        Me.txt�ͼ���.Enabled = False
        If Me.rptCount.Records.Count > 0 Then
            If MsgBox("�Ƿ��ӡ��ǰ������ɵ��嵥?", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
                RegisterLisPrint (1)
            End If
        End If
        
    End If
    If chkRemberPer.Value = 1 Then
        If Nvl(mstrSendPerson) <> "" Then
            txt�ͼ��� = mstrSendPerson
        Else
            txt�ͼ��� = ""
        End If
    Else
        Me.txt�ͼ���.Text = ""
    End If
    Me.cbrthis.RecalcLayout
    Me.rptCount.Records.DeleteAll
    Me.rptCount.Populate

End Sub
Private Sub GetBatch()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String                                '����SQL�����
    Dim strSQL1 As String                               '����H��Ĳ���
    Dim Record As ReportRecord
    Dim Item As ReportColumn
    Dim intLoop As Integer
    Dim strTmp As String                                '��ʱ�ִ�����
    Dim varFilter As Variant                            '�����ִ�
    Dim strDateBegin As Date                            '��ʼʱ��
    Dim strDateEnd As Date                              '����ʱ��
    Dim blnDateMoved As Boolean                         '�Ƿ�ת��
    Dim lngPatientID As Long                            '����ID
    Dim strState As String                              '״̬
    Dim Controlcbo As CommandBarComboBox                '���οؼ�
    
    On Error GoTo errH
    
    '��ע����ж�ȡ��������
    strTmp = zlDatabase.GetPara("�걾�Ǽǹ���", 100, 1212, "")
    Set Controlcbo = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_File_RoomSet, True, True)
    
    '�ӹ��˴������������ʱ����
    If Me.rptPlist.Tag <> "" Then
        varFilter = Split(Me.rptPlist.Tag, ";")
    Else
        If strTmp <> "" Then
            varFilter = Split(strTmp, ";")
        End If
    End If
    
    strSQL = "Select Distinct b.��������" & vbNewLine & _
            "From ����ҽ����¼ A, ����ҽ������ B, ������Ϣ C, ������ĿĿ¼ F" & vbNewLine & _
            "Where a.Id = b.ҽ��id And a.����id = c.����id And a.������Ŀid = f.Id And a.������� = 'E' And f.�������� = '6' And" & vbNewLine & _
            "      b.������ Is Not Null And b.ִ�в���id = [1]"
    
    If Me.rptPlist.Tag <> "" Then
        If Val(varFilter(mFilter.סԺ)) = 0 Then
            strSQL = strSQL & " And A.������Դ in (" & IIf(Val(varFilter(mFilter.����)) = 1, "1,3,", "") & _
                     Val(varFilter(mFilter.סԺ)) & "," & Val(varFilter(mFilter.���)) & ") "
        Else
            strSQL = strSQL & " and c.��Ժʱ�� Is Null And A.������Դ in (" & IIf(Val(varFilter(mFilter.����)) = 1, "1,3,", "") & _
                     Val(varFilter(mFilter.סԺ)) & "," & Val(varFilter(mFilter.���)) & ") "
        End If
        
        If varFilter(mFilter.���￨) <> "" Then
            strSQL = strSQL & " And C.���￨�� = [3] "
        End If
        
        If varFilter(mFilter.����) <> "" Then
            strSQL = strSQL & " And C.���� like [4] "
        End If
        
        If varFilter(mFilter.���ݺ�) <> "" Then
            strSQL = strSQL & " and B.NO = [5]"
        End If
        
        If varFilter(mFilter.�걾) <> "���б걾" Then
            strSQL = strSQL & " and A.�걾��λ = [6] "
        End If
        
        If varFilter(mFilter.�ɼ���ʽ) <> 0 Then
            strSQL = strSQL & " and f.ID = [7] "
        End If
        
        If varFilter(mFilter.���˿���) <> 0 Then
            strSQL = strSQL & " And a.���˿���ID = [8] "
        End If
        
        strSQL = strSQL & " and b.����ʱ�� Between [9] and [10]"
        
        If varFilter(mFilter.��ʼʱ��) = "" Then
            strDateBegin = zlDatabase.Currentdate - Val(varFilter(mFilter.���ʱ��))
            strDateEnd = zlDatabase.Currentdate
        Else
            strDateBegin = varFilter(mFilter.��ʼʱ��)
            strDateEnd = varFilter(mFilter.����ʱ��)
        End If
    Else
        If strTmp <> "" Then
            strSQL = strSQL & " And A.������Դ in (" & IIf(Val(varFilter(mFilter.����)) = 1, "1,3,", "") & _
                 Val(varFilter(mFilter.סԺ)) & "," & Val(varFilter(mFilter.���)) & ") "
            
            If varFilter(mFilter.�걾) <> "���б걾" Then
                strSQL = strSQL & " and A.�걾��λ = [6] "
            End If
            
            If varFilter(mFilter.�ɼ���ʽ) <> 0 Then
                strSQL = strSQL & " and f.ID = [7] "
            End If
            
            If varFilter(mFilter.���˿���) <> 0 Then
                strSQL = strSQL & " And a.���˿���ID = [8] "
            End If
            
            strSQL = strSQL & " and b.����ʱ�� Between [9] and [10]"
            
            If varFilter(mFilter.��ʼʱ��) = "" Then
                strDateBegin = zlDatabase.Currentdate - Val(varFilter(mFilter.���ʱ��))
                strDateEnd = zlDatabase.Currentdate
            Else
                strDateBegin = varFilter(mFilter.��ʼʱ��)
                strDateEnd = varFilter(mFilter.����ʱ��)
            End If
        Else
            strSQL = strSQL & " and b.����ʱ�� Between [9] and [10]"
            strDateBegin = zlDatabase.Currentdate - 3
            strDateEnd = zlDatabase.Currentdate
        End If
    End If
    
    blnDateMoved = MovedByDate(CDate(strDateBegin)) '��ʱ�俴�Ƿ������ת��
    
    If blnDateMoved Then
        strSQL1 = strSQL
        strSQL1 = Replace(strSQL1, "����ҽ����¼", "H����ҽ����¼")
        strSQL1 = Replace(strSQL1, "����ҽ������", "H����ҽ������")
        strSQL = strSQL & " Union ALL " & strSQL1
    End If
    
    strSQL = strSQL & " Order by b.�������� "
    
    If strTmp = "" And Me.rptPlist.Tag = "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngDeptID, "", "", "", "", "", "", "", _
                    CDate(Format(strDateBegin, "yyyy-MM-dd 00:00:00")), CDate(Format(strDateEnd, "yyyy-MM-dd 23:59:59")))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngDeptID, Val(varFilter(mFilter.��ʶ��)), CStr(varFilter(mFilter.���￨)) _
                    , CStr(varFilter(mFilter.����)) & "%", CStr(varFilter(mFilter.���ݺ�)), CStr(varFilter(mFilter.�걾)), CLng(varFilter(mFilter.�ɼ���ʽ)) _
                    , varFilter(mFilter.���˿���), CDate(Format(strDateBegin, "yyyy-MM-dd 00:00:00")), _
                    CDate(Format(strDateEnd, "yyyy-MM-dd 23:59:59")))
    End If
    
    intLoop = 1
    Controlcbo.Clear
    Controlcbo.AddItem "��������"
    Controlcbo.ItemData(Controlcbo.ListCount) = 0
    
    Do While Not rsTmp.EOF
        Controlcbo.AddItem "��" & intLoop & "����"
        Controlcbo.ItemData(Controlcbo.ListCount) = Val(Nvl(rsTmp("��������")))
        If Val(Nvl(rsTmp("��������"))) = mlngSelectBatch Then
            Controlcbo.ListIndex = Controlcbo.ListCount
        End If
        intLoop = intLoop + 1
        rsTmp.MoveNext
    Loop
    If Controlcbo.ListIndex < 1 Then
        Controlcbo.ListIndex = 1
        mlngSelectBatch = 0
    End If
    Exit Sub
errH:
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub RegisterLisPrint(Mode As Integer)
    '����       ��ӡ�Ǽ��嵥
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1212_1", Me, "��������=" & IIf(mlngSelectBatch = 0, mlngBatch, mlngSelectBatch), Mode)
End Sub

Private Sub IdKindChange()
    If Me.ActiveControl Is txtGoto Then
       IDKind.IDKind = IIf(IDKind.IDKind = IDKinds.C5���￨, 0, IDKind.IDKind + 1)
    End If
End Sub

Private Sub InsrOrDelAdvice(intType As Integer, strAdvice As String)
    '����       ���Ӻ�ɾ��������ҽ��
    '����       intType = 1 ����ҽ�� 0 = ɾ����ǰҽ��
    '           ҽ��ID�ִ�
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim intLoop As Integer
    Dim intItem As Integer
    Dim aItem() As String
    Dim RecordA As ReportRecord
    Dim strOldAdvice As String
    Dim Record As ReportRecord
    Dim strCuvetteNumber As String
    Dim strBarCode As String
    Dim intRow As Integer
    Dim strSQLbak As String
    Dim intPatientType As Integer                                             '������Դ
    Dim strTemp As String
    Dim strValue As String
    Dim intLen As Integer
    Dim strSqlAll As String
    
    On Error GoTo errH
    
    If intType = 0 Then
        'ɾ��
        For intLoop = Me.rptCount.Records.Count - 1 To 0 Step -1
            With Me.rptCount.Records(intLoop)
                If InStr("," & strAdvice & ",", "," & .Item(mAcol.ҽ��id).Value & ",") > 0 Then
                    Call Me.rptCount.Records.RemoveAt(intLoop)
                End If
            End With
        Next
        Me.rptCount.Populate
    End If
    
    If intType = 1 Then
        '����
        If Len(strAdvice) > 4000 Then
            strTemp = strAdvice
            Do While Len(strTemp) > 4000
                strValue = Mid(strTemp, 1, 4000)
                intLen = InStrRev(strValue, ",")
                strTemp = Mid(strValue, intLen + 1) & Mid(strTemp, 4001)
                strValue = Mid(strValue, 1, intLen - 1)
                
                strSQL = " Select /*+ rule */ B.ID as ҽ��ID, B.���id, G.��ɫ As �Թ���ɫ, D.���� As �ɼ���ʽ, B.ҽ������, C.��������,C.����ʱ��,c.�걾�ͳ�ʱ�� as �ͼ�ʱ��, " & vbCrLf & _
                         " H.���� As ִ�п���, B.����ҽ��,B.����ʱ��, C.������, C.����ʱ��, G.���� as �Թܱ���,b.�걾��λ as �걾, " & vbCrLf & vbCrLf & _
                         " I.���� as ��������,I.�Ա�,i.����,i.��ǰ���� as ����,decode(b.������Դ,1,I.�����,2,i.סԺ��) as ��ʶ��, " & vbCrLf & _
                         " L.���� as �������ڿ���,Decode(C.ִ��״̬,2,'����') as ����,I.����Id,c.������,c.�ͼ���,G.��Ѫ��,G.���� as �Թ�����, " & vbCrLf & _
                         " DECODE(B.������־,1,'����','') as ����,decode(b.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���') as ������Դ,b.Ӥ��,N.���� as ����,J.���� as ���˿���,C.����ʱ��,C.������, " & vbCrLf & _
                         " b.������ĿID,C.ִ��״̬,O.��¼����,O.��¼״̬ " & vbCrLf & _
                         " From ����ҽ����¼ A, ����ҽ����¼ B, ����ҽ������ C, ������ĿĿ¼ D, ������ĿĿ¼ E, " & vbCrLf & _
                         " ��Ѫ������ G,���ű� H, ������Ϣ I,���ű� L,����ҽ������ M,���ű� J, " & vbCrLf & _
                         " (select ������ĿID,���� from ������Ŀ���� where ���� = 9 and ���� = 1 ) N,סԺ���ü�¼ O " & vbCrLf & _
                         " Where A.ID = B.���id And B.ID = C.ҽ��id And A.������Ŀid = D.ID And B.������Ŀid = E.ID " & vbCrLf & _
                         " And E.��� = 'C' And E.�Թܱ��� = G.���� And m.ִ�в���id = H.ID " & vbCrLf & _
                         " And D.��� = 'E' And D.�������� = '6' And  " & vbCrLf & _
                         " B.����ID = I.����ID and I.��ǰ����ID = L.ID(+) " & _
                         " And a.id  = m.ҽ��id And E.id = N.������ĿID(+) And I.��ǰ����id = J.id(+) And ��Ժʱ�� is null " & vbCrLf & _
                         " and c.ҽ��id = O.ҽ�����(+) and c.��¼���� = mod(O.��¼����(+),10) and nvl(O.��¼״̬,0) in (0,1) " & vbCrLf & _
                         " and b.Id in (Select * From Table(Cast(f_Num2list('" & strValue & "') As zlTools.t_Numlist))) "
                strSqlAll = strSqlAll & strSQL & " union "
            Loop
            
            strSQL = " select /*+ rule */ ������Դ from ����ҽ����¼ where id in (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strValue)
            If rsTmp.EOF = True Then Exit Sub
            intPatientType = Nvl(rsTmp("������Դ"), 0)
            
            strSQL = " Select /*+ rule */ B.ID as ҽ��ID, B.���id, G.��ɫ As �Թ���ɫ, D.���� As �ɼ���ʽ, B.ҽ������, C.��������,C.����ʱ��,c.�걾�ͳ�ʱ�� as �ͼ�ʱ��, " & vbCrLf & _
                     " H.���� As ִ�п���, B.����ҽ��,B.����ʱ��, C.������, C.����ʱ��, G.���� as �Թܱ���,b.�걾��λ as �걾, " & vbCrLf & vbCrLf & _
                     " I.���� as ��������,I.�Ա�,i.����,i.��ǰ���� as ����,decode(b.������Դ,1,I.�����,2,i.סԺ��) as ��ʶ��, " & vbCrLf & _
                     " L.���� as �������ڿ���,Decode(C.ִ��״̬,2,'����') as ����,I.����Id,c.������,c.�ͼ���,G.��Ѫ��,G.���� as �Թ�����, " & vbCrLf & _
                     " DECODE(B.������־,1,'����','') as ����,decode(b.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���') as ������Դ,b.Ӥ��,N.���� as ����,J.���� as ���˿���,C.����ʱ��,C.������, " & vbCrLf & _
                     " b.������ĿID,C.ִ��״̬,O.��¼����,O.��¼״̬ " & vbCrLf & _
                     " From ����ҽ����¼ A, ����ҽ����¼ B, ����ҽ������ C, ������ĿĿ¼ D, ������ĿĿ¼ E, " & vbCrLf & _
                     " ��Ѫ������ G,���ű� H, ������Ϣ I,���ű� L,����ҽ������ M,���ű� J, " & vbCrLf & _
                     " (select ������ĿID,���� from ������Ŀ���� where ���� = 9 and ���� = 1 ) N,סԺ���ü�¼ O " & vbCrLf & _
                     " Where A.ID = B.���id And B.ID = C.ҽ��id And A.������Ŀid = D.ID And B.������Ŀid = E.ID " & vbCrLf & _
                     " And E.��� = 'C' And E.�Թܱ��� = G.���� And m.ִ�в���id = H.ID " & vbCrLf & _
                     " And D.��� = 'E' And D.�������� = '6' And  " & vbCrLf & _
                     " B.����ID = I.����ID and I.��ǰ����ID = L.ID(+) " & _
                     " And a.id  = m.ҽ��id And E.id = N.������ĿID(+) And I.��ǰ����id = J.id(+) And ��Ժʱ�� is null " & vbCrLf & _
                     " and c.ҽ��id = O.ҽ�����(+) and c.��¼���� = mod(O.��¼����(+),10) and nvl(O.��¼״̬,0) in (0,1) " & vbCrLf & _
                     " and b.Id in (Select * From Table(Cast(f_Num2list('" & strTemp & "') As zlTools.t_Numlist))) "
            strSqlAll = strSqlAll & strSQL
            
            If intPatientType <> 2 Then
                strSqlAll = Replace(strSqlAll, "סԺ���ü�¼", "������ü�¼")
            End If
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSqlAll, Me.Caption)
            
            Do Until rsTmp.EOF
                'û�ж�Ӧ��ɫ����Ĳɼ���д��
                If IsNull(rsTmp("�Թ���ɫ")) = False Then
                    
                    If strOldAdvice <> rsTmp("���ID") Then
                    
                        Set Record = Me.rptCount.Records.Add
                        For intLoop = 0 To rptCount.Columns.Count + 1
                            Record.AddItem ""
                        Next
                        
                        If (Nvl(rsTmp("ִ��״̬")) = 1 Or Nvl(rsTmp("ִ��״̬")) = 3) Then
                            Me.rptCount.Columns(mAcol.��ִ��).Visible = True
                        End If
                        
                        If Nvl(rsTmp("ִ��״̬")) = 1 Or Nvl(rsTmp("ִ��״̬")) = 3 Then
                            Record(mAcol.��ִ��).Value = "��"
                        Else
                            Record(mAcol.��ִ��).Value = ""
                        End If
                        Record(mAcol.����).Value = IIf(rsTmp("��¼״̬") = 1, "��", "��")
                        Record(mAcol.ID).Value = Nvl(rsTmp("ҽ��ID")) & "," & Nvl(rsTmp("���ID"))
                        Record(mAcol.ѡ��).HasCheckbox = True
                        Record(mAcol.ѡ��).Checked = IIf(Nvl(rsTmp("ִ��״̬")) = 0, True, False)
                        Record(mAcol.ͼ��).BackColor = Val(Nvl(rsTmp("�Թ���ɫ")))
                        Record(mAcol.�ɼ���ʽ).Value = Nvl(rsTmp("�ɼ���ʽ"))
                        Record(mAcol.ҽ������).Value = Nvl(rsTmp("ҽ������"))
                        Record(mAcol.����).Value = Nvl(rsTmp("��������"))
                        Record(mAcol.ִ�п���).Value = Nvl(rsTmp("ִ�п���"))
                        Record(mAcol.����ҽ��).Value = Nvl(rsTmp("����ҽ��"))
                        Record(mAcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                        Record(mAcol.������).Value = Nvl(rsTmp("������"))
                        Record(mAcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                        Record(mAcol.�Թ���ɫ).Value = Nvl(rsTmp("�Թ���ɫ"))
                        Record(mAcol.�Թܱ���).Value = Nvl(rsTmp("�Թܱ���"))
                        Record(mAcol.�걾).Value = Nvl(rsTmp("�걾")) & IIf(Nvl(rsTmp("Ӥ��")) = 0, "", "(Ӥ��)")
                        Record(mAcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                        Record(mAcol.������).Value = Nvl(rsTmp("������"))
                        Record(mAcol.�ͼ���).Value = Nvl(rsTmp("�ͼ���"))
                        Record(mAcol.��Ѫ��).Value = Nvl(rsTmp("��Ѫ��"))
                        Record(mAcol.�Թ�����).Value = Nvl(rsTmp("�Թ�����"))
                        Record(mAcol.����).Value = Nvl(rsTmp("����"))
                        Record(mAcol.������Դ).Value = Nvl(rsTmp("������Դ"))
                        Record(mAcol.Ӥ��).Value = Nvl(rsTmp("Ӥ��"))
                        Record(mAcol.����).Value = Nvl(rsTmp("����"))
                        Record(mAcol.���ID).Value = Nvl(rsTmp("���ID"))
                        
                        Record(mAcol.����ID).Value = Nvl(rsTmp("����ID"))
                        Record(mAcol.����).Value = Nvl(rsTmp("��������"))
                        Record(mAcol.�Ա�).Value = Nvl(rsTmp("�Ա�"))
                        Record(mAcol.����).Value = Nvl(rsTmp("����"))
                        Record(mAcol.��ʶ��).Value = Nvl(rsTmp("��ʶ��"))
                        Record(mAcol.����).Value = Nvl(rsTmp("����"))
                        Record(mAcol.���˿���).Value = Nvl(rsTmp("���˿���"))
                        Record(mAcol.������).Value = Nvl(rsTmp("������"))
                        Record(mAcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                        Record(mAcol.������ĿID).Value = Nvl(rsTmp("������ĿID"))
                        Record(mAcol.ִ��״̬).Value = Nvl(rsTmp("ִ��״̬"))
                        Record(mAcol.ҽ��id).Value = Nvl(rsTmp("ҽ��id"))
                        Record(mAcol.�ͼ�ʱ��).Value = Nvl(rsTmp("�ͼ�ʱ��"))
                        
                        For intLoop = 0 To Me.rptAlist(TabCtr.Selected.Index).Columns.Count + 1
                            Record(intLoop).ForeColor = Val(Nvl(rsTmp("�Թ���ɫ")))
                        Next
                    Else
                        If InStr(Record(mAcol.ҽ������).Value, Nvl(rsTmp("ҽ������"))) <= 0 Then
                            Record(mAcol.ҽ������).Value = Record(mAcol.ҽ������).Value & " " & Nvl(rsTmp("ҽ������"))
                        End If
                        Record(mAcol.�ϲ�ҽ��).Value = Record(mAcol.�ϲ�ҽ��).Value & ";" & _
                                                       Nvl(rsTmp("ҽ��ID")) & "," & Nvl(rsTmp("���ID"))
                        Record(mAcol.����).Value = Record(mAcol.����).Value & " " & Nvl(rsTmp("����"))
                    End If
                    strOldAdvice = rsTmp("���ID")
                    If InStr(1, strCuvetteNumber & ",", "," & Nvl(rsTmp("�Թܱ���")) & ",") <= 0 Then
                        strCuvetteNumber = strCuvetteNumber & "," & Nvl(rsTmp("�Թܱ���"))
                    End If
                End If
                If chkRemberPer.Value = 1 Then
                    If Nvl(rsTmp("�ͼ���") & "") <> "" Then
                        txt�ͼ��� = Nvl(rsTmp("�ͼ���") & "")
                    Else
                        txt�ͼ��� = mstrSendPerson
                    End If
                Else
                    txt�ͼ��� = Nvl(rsTmp("�ͼ���"))
                End If
                rsTmp.MoveNext
            Loop
        Else
            strSQL = " select /*+ rule */ ������Դ from ����ҽ����¼ where id in (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strAdvice)
            If rsTmp.EOF = True Then Exit Sub
            intPatientType = Nvl(rsTmp("������Դ"), 0)
            
            strSQL = " Select /*+ rule */ B.ID as ҽ��ID, B.���id, G.��ɫ As �Թ���ɫ, D.���� As �ɼ���ʽ, B.ҽ������, C.��������,C.����ʱ��,c.�걾�ͳ�ʱ�� as �ͼ�ʱ��, " & vbCrLf & _
                     " H.���� As ִ�п���, B.����ҽ��,B.����ʱ��, C.������, C.����ʱ��, G.���� as �Թܱ���,b.�걾��λ as �걾, " & vbCrLf & vbCrLf & _
                     " I.���� as ��������,I.�Ա�,i.����,i.��ǰ���� as ����,decode(b.������Դ,1,I.�����,2,i.סԺ��) as ��ʶ��, " & vbCrLf & _
                     " L.���� as �������ڿ���,Decode(C.ִ��״̬,2,'����') as ����,I.����Id,c.������,c.�ͼ���,G.��Ѫ��,G.���� as �Թ�����, " & vbCrLf & _
                     " DECODE(B.������־,1,'����','') as ����,decode(b.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���') as ������Դ,b.Ӥ��,N.���� as ����,J.���� as ���˿���,C.����ʱ��,C.������, " & vbCrLf & _
                     " b.������ĿID,C.ִ��״̬,O.��¼����,O.��¼״̬ " & vbCrLf & _
                     " From ����ҽ����¼ A, ����ҽ����¼ B, ����ҽ������ C, ������ĿĿ¼ D, ������ĿĿ¼ E, " & vbCrLf & _
                     " ��Ѫ������ G,���ű� H, ������Ϣ I,���ű� L,����ҽ������ M,���ű� J, " & vbCrLf & _
                     " (select ������ĿID,���� from ������Ŀ���� where ���� = 9 and ���� = 1 ) N,סԺ���ü�¼ O " & vbCrLf & _
                     " Where A.ID = B.���id And B.ID = C.ҽ��id And A.������Ŀid = D.ID And B.������Ŀid = E.ID " & vbCrLf & _
                     " And E.��� = 'C' And E.�Թܱ��� = G.���� And m.ִ�в���id = H.ID " & vbCrLf & _
                     " And D.��� = 'E' And D.�������� = '6' And  " & vbCrLf & _
                     " B.����ID = I.����ID and I.��ǰ����ID = L.ID(+) " & _
                     " And a.id  = m.ҽ��id And E.id = N.������ĿID(+) And I.��ǰ����id = J.id(+) And ��Ժʱ�� is null " & vbCrLf & _
                     " and c.ҽ��id = O.ҽ�����(+) and c.��¼���� = mod(O.��¼����(+),10) and nvl(O.��¼״̬,0) in (0,1) " & vbCrLf & _
                     " and b.Id in (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))  "
                     
            If intPatientType <> 2 Then
                strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
            End If
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strAdvice)
            
            Do Until rsTmp.EOF
                'û�ж�Ӧ��ɫ����Ĳɼ���д��
                If IsNull(rsTmp("�Թ���ɫ")) = False Then
                    
                    If strOldAdvice <> rsTmp("���ID") Then
                    
                        Set Record = Me.rptCount.Records.Add
                        For intLoop = 0 To rptCount.Columns.Count + 1
                            Record.AddItem ""
                        Next
                        
                        If (Nvl(rsTmp("ִ��״̬")) = 1 Or Nvl(rsTmp("ִ��״̬")) = 3) Then
                            Me.rptCount.Columns(mAcol.��ִ��).Visible = True
                        End If
                        
                        If Nvl(rsTmp("ִ��״̬")) = 1 Or Nvl(rsTmp("ִ��״̬")) = 3 Then
                            Record(mAcol.��ִ��).Value = "��"
                        Else
                            Record(mAcol.��ִ��).Value = ""
                        End If
                        Record(mAcol.����).Value = IIf(rsTmp("��¼״̬") = 1, "��", "��")
                        Record(mAcol.ID).Value = Nvl(rsTmp("ҽ��ID")) & "," & Nvl(rsTmp("���ID"))
                        Record(mAcol.ѡ��).HasCheckbox = True
                        Record(mAcol.ѡ��).Checked = IIf(Nvl(rsTmp("ִ��״̬")) = 0, True, False)
                        Record(mAcol.ͼ��).BackColor = Val(Nvl(rsTmp("�Թ���ɫ")))
                        Record(mAcol.�ɼ���ʽ).Value = Nvl(rsTmp("�ɼ���ʽ"))
                        Record(mAcol.ҽ������).Value = Nvl(rsTmp("ҽ������"))
                        Record(mAcol.����).Value = Nvl(rsTmp("��������"))
                        Record(mAcol.ִ�п���).Value = Nvl(rsTmp("ִ�п���"))
                        Record(mAcol.����ҽ��).Value = Nvl(rsTmp("����ҽ��"))
                        Record(mAcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                        Record(mAcol.������).Value = Nvl(rsTmp("������"))
                        Record(mAcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                        Record(mAcol.�Թ���ɫ).Value = Nvl(rsTmp("�Թ���ɫ"))
                        Record(mAcol.�Թܱ���).Value = Nvl(rsTmp("�Թܱ���"))
                        Record(mAcol.�걾).Value = Nvl(rsTmp("�걾")) & IIf(Nvl(rsTmp("Ӥ��")) = 0, "", "(Ӥ��)")
                        Record(mAcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                        Record(mAcol.������).Value = Nvl(rsTmp("������"))
                        Record(mAcol.�ͼ���).Value = Nvl(rsTmp("�ͼ���"))
                        Record(mAcol.��Ѫ��).Value = Nvl(rsTmp("��Ѫ��"))
                        Record(mAcol.�Թ�����).Value = Nvl(rsTmp("�Թ�����"))
                        Record(mAcol.����).Value = Nvl(rsTmp("����"))
                        Record(mAcol.������Դ).Value = Nvl(rsTmp("������Դ"))
                        Record(mAcol.Ӥ��).Value = Nvl(rsTmp("Ӥ��"))
                        Record(mAcol.����).Value = Nvl(rsTmp("����"))
                        Record(mAcol.���ID).Value = Nvl(rsTmp("���ID"))
                        
                        Record(mAcol.����ID).Value = Nvl(rsTmp("����ID"))
                        Record(mAcol.����).Value = Nvl(rsTmp("��������"))
                        Record(mAcol.�Ա�).Value = Nvl(rsTmp("�Ա�"))
                        Record(mAcol.����).Value = Nvl(rsTmp("����"))
                        Record(mAcol.��ʶ��).Value = Nvl(rsTmp("��ʶ��"))
                        Record(mAcol.����).Value = Nvl(rsTmp("����"))
                        Record(mAcol.���˿���).Value = Nvl(rsTmp("���˿���"))
                        Record(mAcol.������).Value = Nvl(rsTmp("������"))
                        Record(mAcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                        Record(mAcol.������ĿID).Value = Nvl(rsTmp("������ĿID"))
                        Record(mAcol.ִ��״̬).Value = Nvl(rsTmp("ִ��״̬"))
                        Record(mAcol.ҽ��id).Value = Nvl(rsTmp("ҽ��id"))
                        Record(mAcol.�ͼ�ʱ��).Value = Nvl(rsTmp("�ͼ�ʱ��"))
                        
                        For intLoop = 0 To Me.rptAlist(TabCtr.Selected.Index).Columns.Count + 1
                            Record(intLoop).ForeColor = Val(Nvl(rsTmp("�Թ���ɫ")))
                        Next
                    Else
                        If InStr(Record(mAcol.ҽ������).Value, Nvl(rsTmp("ҽ������"))) <= 0 Then
                            Record(mAcol.ҽ������).Value = Record(mAcol.ҽ������).Value & " " & Nvl(rsTmp("ҽ������"))
                        End If
                        Record(mAcol.�ϲ�ҽ��).Value = Record(mAcol.�ϲ�ҽ��).Value & ";" & _
                                                       Nvl(rsTmp("ҽ��ID")) & "," & Nvl(rsTmp("���ID"))
                        Record(mAcol.����).Value = Record(mAcol.����).Value & " " & Nvl(rsTmp("����"))
                    End If
                    strOldAdvice = rsTmp("���ID")
                    If InStr(1, strCuvetteNumber & ",", "," & Nvl(rsTmp("�Թܱ���")) & ",") <= 0 Then
                        strCuvetteNumber = strCuvetteNumber & "," & Nvl(rsTmp("�Թܱ���"))
                    End If
                End If
                txt�ͼ��� = Nvl(rsTmp("�ͼ���"))
                rsTmp.MoveNext
            Loop
        End If
        Me.rptCount.Populate
    End If
    '---------------------------------���¼��㵱ǰ�Ǽǵ�����---------------------------------------
    strCuvetteNumber = ""
    strBarCode = ""
    Me.rptCuvetteCount.Records.DeleteAll
    Me.rptCuvetteCount.Populate
    For intLoop = 0 To Me.rptCount.Rows.Count - 1
        With Me.rptCount.Rows(intLoop)
        
            If chkNO(.Record(mAcol.�Թܱ���).Value) = False Then
                Set Record = Me.rptCuvetteCount.Records.Add
                For intRow = 0 To rptCuvetteCount.Columns.Count + 1
                    Record.AddItem ""
                Next
                Record(mCuvetteCount.����).Value = .Record(mAcol.�Թܱ���).Value
                Record(mCuvetteCount.����).Value = .Record(mAcol.�Թ�����).Value
                For intRow = 0 To Me.rptCuvetteCount.Columns.Count - 1
                    Record(intRow).ForeColor = .Record(mAcol.�Թ���ɫ).Value
                Next
                Me.rptCuvetteCount.Populate
            End If
        End With
    Next
    
    Me.rptCuvetteCount.Populate
    For intLoop = 0 To Me.rptCount.Rows.Count - 1
        With Me.rptCount.Rows(intLoop)
            If ChkBarCode(.Record(mAcol.����).Value, intLoop) = False Then
                For intRow = 0 To Me.rptCuvetteCount.Rows.Count - 1
                    If Me.rptCuvetteCount.Rows(intRow).Record(mCuvetteCount.����).Value = .Record(mAcol.�Թܱ���).Value Then
                        Me.rptCuvetteCount.Rows(intRow).Record(mCuvetteCount.�ϼ�).Value = _
                            Val(Me.rptCuvetteCount.Rows(intRow).Record(mCuvetteCount.�ϼ�).Value) + 1
                        Exit For
                    End If
                Next
            End If

        End With
    Next
    
    Me.rptCuvetteCount.Populate
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function chkNO(strNO As String) As Boolean
    '�����Թܱ����Ƿ��ظ�
    Dim intLoop As Integer
    For intLoop = 0 To Me.rptCuvetteCount.Rows.Count - 1
        With Me.rptCuvetteCount.Rows(intLoop).Record
            If .Item(mCuvetteCount.����).Value = strNO Then
                chkNO = True
                Exit For
            End If
        End With
    Next
End Function

Private Function ChkBarCode(strBarCode As String, intIndex As Integer) As Boolean
    '���������Ƿ��ظ�
    Dim intLoop As Integer
    For intLoop = 0 To intIndex - 1
        With Me.rptCount.Rows(intLoop).Record
            If .Item(mAcol.����).Value = strBarCode Then
                ChkBarCode = True
                Exit For
            End If
        End With
    Next

End Function
Public Function CheckIsInclude(strSource As String, strTarge As String) As Boolean
    '���strSource�е�ÿһ���ַ��Ƿ���strTarge��
    Dim i As Long
    CheckIsInclude = False
    
    Select Case strTarge
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "ʱ��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+-_)(*&^%$#@!`~"
    Case "����ʱ��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+_)(*&^%$#@!`~"
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "������"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "��С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "�ɴ�ӡ�ַ�"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/."":;|\=+-_)(*&^%$#@!`~0123456789"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    CheckIsInclude = True
End Function
Private Sub WriterCheckSampleToLIS(strAdvices As String, strName As String, strBatchNO As Long, Optional ByVal strSentName As String)
    '����   ��ǩ����Ϣд��LIS
    Dim strErr As String
    If Not mobjLisInsideComm Is Nothing Then
        If mobjLisInsideComm.SampleCheckinInfoWrite(strAdvices, strName, strBatchNO, strErr, strSentName) = False Then
            MsgBox "д��ǩ����Ϣ��LIS���뵥����!" & vbCrLf & strErr
        End If
    End If
End Sub


Private Sub RePrintBarCode(ByVal blnNotRePrint As Boolean)
    Dim strSQL  As String
    Dim intFrom  As Integer
    With rptAlist(TabCtr.Selected.Index)
        If blnNotRePrint Then
            If .Records.Count = 0 Then
                Me.cbrthis.FindControl(, conMenu_Edit_ReprintReceipt).Enabled = False
                Exit Sub
            End If
            If Trim(.FocusedRow.Record(mAcol.����).Value) <> "" Then
                Me.cbrthis.FindControl(, conMenu_Edit_ReprintReceipt).Enabled = True
            Else
                Me.cbrthis.FindControl(, conMenu_Edit_ReprintReceipt).Enabled = False
            End If
        Else
            If .Records.Count = 0 Then
                MsgBox "δѡ��ҽ����", vbInformation, "��ʾ"
                Exit Sub
            End If
            If mintCodeType = BarCodeType.Code39 Then
                Bar39 Me.picBarCode, 3, Trim(.FocusedRow.Record(mAcol.����).Value), False, True
            Else
                Bar128 Me.picBarCode, 3, Trim(.FocusedRow.Record(mAcol.����).Value), True
            End If
            
            SavePicture Me.picBarCode.Image, App.path & "\BarCode.bmp"
            Select Case Trim(.FocusedRow.Record(mAcol.������Դ).Value)
                Case "����"
                    intFrom = 1
                Case "סԺ"
                    intFrom = 2
                Case "Ժ��"
                    intFrom = 3
                Case "���"
                    intFrom = 4
            End Select
            
            
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1212_2", Me, "��������=" & Trim(.FocusedRow.Record(mAcol.����).Value), _
            "��Ŀ = " & Trim(.FocusedRow.Record(mAcol.ҽ������).Value), _
            "�������� = " & IIf(Trim(.FocusedRow.Record(mAcol.����).Value) <> "", Trim(.FocusedRow.Record(mAcol.����).Value), "��"), _
            "�Ա� = " & IIf(Trim(.FocusedRow.Record(mAcol.�Ա�).Value) <> "", Trim(.FocusedRow.Record(mAcol.�Ա�).Value), "��"), _
            "���� = " & IIf(Trim(.FocusedRow.Record(mAcol.����).Value) <> "", Trim(.FocusedRow.Record(mAcol.����).Value), "��"), _
            "���� = " & IIf(Trim(.FocusedRow.Record(mAcol.����).Value) <> "", Trim(.FocusedRow.Record(mAcol.����).Value), "��"), _
            "��ʶ�� = " & IIf(Trim(.FocusedRow.Record(mAcol.��ʶ��).Value) <> "", Trim(.FocusedRow.Record(mAcol.��ʶ��).Value), "��"), _
            "���ڿ��� = " & IIf(Trim(.FocusedRow.Record(mAcol.���˿���).Value) <> "", Trim(.FocusedRow.Record(mAcol.���˿���).Value), "��"), _
            "�ɼ���ʽ = " & IIf(Trim(.FocusedRow.Record(mAcol.�ɼ���ʽ).Value) <> "", Trim(.FocusedRow.Record(mAcol.�ɼ���ʽ).Value), "��"), _
            "�걾 = " & IIf(Trim(.FocusedRow.Record(mAcol.�걾).Value) <> "", Trim(.FocusedRow.Record(mAcol.�걾).Value), "��"), _
            "ִ�п��� = " & IIf(Trim(.FocusedRow.Record(mAcol.ִ�п���).Value) <> "", Trim(.FocusedRow.Record(mAcol.ִ�п���).Value), "��"), _
            "����ҽ�� = " & IIf(Trim(.FocusedRow.Record(mAcol.����ҽ��).Value) <> "", Trim(.FocusedRow.Record(mAcol.����ҽ��).Value), "��"), _
            "����ʱ�� = " & IIf(Trim(.FocusedRow.Record(mAcol.����ʱ��).Value) <> "", Trim(.FocusedRow.Record(mAcol.����ʱ��).Value), "��"), _
            "������ = " & IIf(Trim(.FocusedRow.Record(mAcol.������).Value) <> "", Trim(.FocusedRow.Record(mAcol.������).Value), "��"), _
            "����ʱ�� = " & IIf(Trim(.FocusedRow.Record(mAcol.����ʱ��).Value) <> "", Trim(.FocusedRow.Record(mAcol.����ʱ��).Value), "��"), _
            "���� = " & IIf(Trim(.FocusedRow.Record(mAcol.�Թܱ���).Value) <> "", Trim(.FocusedRow.Record(mAcol.�Թܱ���).Value), "��"), _
            "��Ѫ�� = " & IIf(Trim(.FocusedRow.Record(mAcol.��Ѫ��).Value) <> "", Trim(.FocusedRow.Record(mAcol.��Ѫ��).Value), "��"), _
            "�Թ����� = " & IIf(Trim(.FocusedRow.Record(mAcol.�Թ�����).Value) <> "", Trim(.FocusedRow.Record(mAcol.�Թ�����).Value), "��"), _
            "���� = " & IIf(Trim(.FocusedRow.Record(mAcol.����).Value) <> "", Trim(.FocusedRow.Record(mAcol.����).Value), 0), _
            "������Դ = " & intFrom, _
            "����ͼ��1=" & App.path & "\BarCode.Bmp", 2)
            Kill App.path & "\BarCode.Bmp"
            
        End If
    End With
End Sub

Private Function getTATTime(strIDs As String) As Boolean
    '���TAT��ʱ,���ؿ����ͼ��ҽ��ID
    Dim strSex As String    '�Ա�
    Dim strDept As String   '�������
    Dim strItem As String   '������Ŀ   ��ĿID1,��Ŀ����1,����ʱ��1,����1;��ĿID2,��Ŀ����12,����ʱ��2,����2........
    Dim Record As ReportRecord
    Dim intMsg As Integer
    Dim strShowBef As String
    Dim strMsgShow As String
    Dim strMsgShowStop As String
    Dim strItemCode As String
    Dim strMsgNoTime As String 'û����һ��ʱ��ڵ����Ŀ
    
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
            '��ʼ��LIS�ӿڲ���
            If mobjLisInsideComm.InitComponentsHIS(glngSys, glngModul, gcnOracle, strErr) = False Then
                If strErr <> "" Then
                    MsgBox "��ʼ��LIS�ӿ�ʧ�ܣ�" & vbCrLf & strErr
                End If
                Set mobjLisInsideComm = Nothing
            End If
        End If
    End If
    
'    If Me.rptPlist.FocusedRow Is Nothing Then
'        getTATTime = False
'        Exit Function
'    End If
    
    '��ȡ�����Ա���������
'    With Me.rptPlist.FocusedRow
'        strSex = .Record(mPcol.�Ա�).Value
'        strDept = .Record(mPcol.���˿���).Value
'    End With
    
    '��ȡ��ĿID,��Ŀ����,����ʱ��,����
    strItem = ""
    For Each Record In Me.rptAlist(Me.TabCtr.Selected.Index).Records
        If Record(mAcol.ѡ��).Checked = True Then
'            If Record(mAcol.�ͼ�ʱ��).Value <> "" Then
                var_Item = Split(Mid(strIDs, 2), "|")
                For i = LBound(var_Item) To UBound(var_Item)
                    strItem = strItem & ";" & Record(mAcol.������ĿID).Value & "," & Record(mAcol.ҽ������).Value & _
                                            "," & Record(mAcol.�ͼ�ʱ��).Value & "," & IIf(Record(mAcol.����).Value = "����", 1, 0) & _
                                             "," & var_Item(i) & "," & Record(mAcol.����).Value
                Next
'            Else
'                strMsgNoTime = strMsgNoTime & Record(mAcol.ҽ������).Value & vbCrLf
'            End If
            '��ȡ�����Ա���������
            strSex = Record(mAcol.�Ա�).Value
            strDept = Record(mAcol.�������).Value
        End If
    Next
    
    If strMsgNoTime <> "" Then MsgBox strMsgNoTime & "δ�ͼ�,����ǩ��   ", vbInformation, Me.Caption
    If strItem <> "" Then
        strItem = Mid(strItem, 2)
    Else
        strIDs = ""
        getTATTime = False
        Exit Function
    End If
    
    '���TAT�Ƿ�ʱ
    On Error GoTo errold
    strTATItems = mobjLisInsideComm.GetTatTimeShow(2, strItem, strDept, "", "", strSex, intMsg, strShowBef, , UserInfo.����)
    If strTATItems <> "" Then
        var_Tmp = Split(strTATItems, ";")
        Do While UBound(Split(var_Tmp(0), ",")) < 9
            '����9��Ԫ�ص������ƴ��һ��0
            strTATItems = ""
            For i = LBound(var_Tmp) To UBound(var_Tmp)
                strTATItems = strTATItems & ";" & var_Tmp(i) & ",0"
            Next
            If strTATItems <> "" Then strTATItems = Mid(strTATItems, 2)
            var_Tmp = Split(strTATItems, ";")
        Loop
        
        '��ȡ���г�ʱ�ҽ�ֹ��Ŀ������
        For i = LBound(var_Tmp) To UBound(var_Tmp)
            If Split(var_Tmp(i), ",")(7) <> 0 And Split(var_Tmp(i), ",")(9) = 2 Then
                strItemCode = strItemCode & "," & Split(var_Tmp(i), ",")(6)
            End If
        Next
        
        strIDs = ""
        
        For i = LBound(var_Tmp) To UBound(var_Tmp)
            If Split(var_Tmp(i), ",")(7) <> 0 And Split(var_Tmp(i), ",")(9) = 1 And InStr(strItemCode, "," & Split(var_Tmp(i), ",")(6)) <= 0 And Split(var_Tmp(i), ",")(2) <> "" Then
                '�ѳ�ʱֻ��ʾ
                strMsgShow = strMsgShow & Replace(Replace(Split(var_Tmp(i), ",")(8), "[��Ŀ]", Split(var_Tmp(i), ",")(1)), "[��ʱ]", Split(var_Tmp(i), ",")(7) & "����") & vbCrLf
            ElseIf Split(var_Tmp(i), ",")(7) <> 0 And Split(var_Tmp(i), ",")(9) = 1 And InStr(strItemCode, "," & Split(var_Tmp(i), ",")(6)) > 0 And Split(var_Tmp(i), ",")(2) <> "" Then
                '����ͬ������Ŀ��
                strMsgShow = strMsgShow & Replace(Replace(Split(var_Tmp(i), ",")(8), "[��Ŀ]", Split(var_Tmp(i), ",")(1)), "[��ʱ]", "") & "����ͬ�����ֹ��Ŀ,���ܼ���" & vbCrLf
            ElseIf Split(var_Tmp(i), ",")(8) <> "0" And Split(var_Tmp(i), ",")(2) = "" Then
                'û��ǰһ��ʱ��ڵ��
                strMsgShowStop = strMsgShowStop & Split(var_Tmp(i), ",")(1) & "δ�ͼ�,����ǩ��" & vbCrLf
            ElseIf Split(var_Tmp(i), ",")(7) <> 0 And Split(var_Tmp(i), ",")(9) = 2 And Split(var_Tmp(i), ",")(2) <> "" Then
                '��ʱ����ֹ��
                strMsgShowStop = strMsgShowStop & Replace(Replace(Split(var_Tmp(i), ",")(8), "[��Ŀ]", Split(var_Tmp(i), ",")(1)), "[��ʱ]", Split(var_Tmp(i), ",")(7) & "����") & vbCrLf
            Else
                '��ͬ��Ŀͬ�����ʱ��,����һ����Ŀ��ʱ,�����и��������Ŀ�������ͼ�
                If InStr(strItemCode, "," & Split(var_Tmp(i), ",")(6)) <= 0 Then
                    strIDs = strIDs & "|" & Split(var_Tmp(i), ",")(4) & "," & Split(var_Tmp(i), ",")(5)
                End If
            End If
        Next
        
        '������Ϊ��ʾʱ,�������ʱ,���ͼ����й�ѡ����Ŀ,���˷�,��ֻ�ͼ�Ϊ��ʱ�ı걾
        If strMsgShow <> "" Then
            If MsgBox(strMsgShow & "�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                If strIDs <> "" Then
                    getTATTime = True
                Else
                    getTATTime = False
                End If
                Exit Function
            Else
                '�����,�����������й�ѡ����Ŀ
                strIDs = ""
                For i = LBound(var_Tmp) To UBound(var_Tmp)
                    If (Split(var_Tmp(i), ",")(7) = 0 Or Split(var_Tmp(i), ",")(9) = 1) And InStr(strItemCode, "," & Split(var_Tmp(i), ",")(6)) <= 0 And Split(var_Tmp(i), ",")(2) <> "" Then
                        strIDs = strIDs & "|" & Split(var_Tmp(i), ",")(4) & "," & Split(var_Tmp(i), ",")(5)
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

Private Sub txt�ͼ���_DblClick()
    Dim strVal As String
    Dim rsTmp As Recordset
    If Not mrsSendPerson Is Nothing Then
        mrsSendPerson.filter = ""
        If mrsSendPerson.RecordCount > 0 Then
            Set rsTmp = mrsSendPerson
            strVal = frmSelectPub.ShowMe(Me, rsTmp, "")
            If strVal <> "" Then
                txt�ͼ���.Text = Split(strVal, ",")(2)
                mstrSendPerson = txt�ͼ���.Text
                cmdOK.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txt�ͼ���_KeyPress(KeyAscii As Integer)
    Dim strVal As String
    Dim rsTmp As Recordset
    If KeyAscii = 13 Then
        If Not mrsSendPerson Is Nothing Then
            mrsSendPerson.filter = ""
            If mrsSendPerson.RecordCount > 0 Then
                Set rsTmp = mrsSendPerson
                strVal = frmSelectPub.ShowMe(Me, rsTmp, Trim(txt�ͼ���.Text))
                If strVal <> "" Then
                    txt�ͼ���.Text = Split(strVal, ",")(2)
                    mstrSendPerson = txt�ͼ���.Text
                    cmdOK.SetFocus
                End If
            End If
        End If
    End If
End Sub

