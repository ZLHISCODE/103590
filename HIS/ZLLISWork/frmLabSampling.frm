VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "*\A..\ZLIDKIND\zlIDKind.vbp"
Begin VB.Form frmLabSampling 
   AutoRedraw      =   -1  'True
   Caption         =   "����ɼ�����վ"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15360
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLabSampling.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   15360
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   615
      Index           =   0
      Left            =   13620
      TabIndex        =   39
      Top             =   3840
      Width           =   765
      _Version        =   589884
      _ExtentX        =   1349
      _ExtentY        =   1085
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   615
      Index           =   1
      Left            =   13980
      TabIndex        =   40
      Top             =   4320
      Width           =   765
      _Version        =   589884
      _ExtentX        =   1349
      _ExtentY        =   1085
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   615
      Index           =   2
      Left            =   13980
      TabIndex        =   41
      Top             =   4680
      Width           =   765
      _Version        =   589884
      _ExtentX        =   1349
      _ExtentY        =   1085
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   615
      Index           =   3
      Left            =   13980
      TabIndex        =   42
      Top             =   4800
      Width           =   765
      _Version        =   589884
      _ExtentX        =   1349
      _ExtentY        =   1085
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   615
      Index           =   5
      Left            =   13380
      TabIndex        =   43
      Top             =   4230
      Width           =   765
      _Version        =   589884
      _ExtentX        =   1349
      _ExtentY        =   1085
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAlist 
      Height          =   615
      Index           =   4
      Left            =   12540
      TabIndex        =   44
      Top             =   3720
      Width           =   765
      _Version        =   589884
      _ExtentX        =   1349
      _ExtentY        =   1085
      _StockProps     =   0
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.PictureBox picAdvice 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2004
      Left            =   8760
      ScaleHeight     =   2010
      ScaleWidth      =   3375
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   3372
      Begin XtremeSuiteControls.TabControl TabCtr 
         Height          =   1245
         Left            =   720
         TabIndex        =   50
         Top             =   420
         Width           =   1965
         _Version        =   589884
         _ExtentX        =   3466
         _ExtentY        =   2196
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picBarCodePrint 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   8220
      ScaleHeight     =   615
      ScaleWidth      =   645
      TabIndex        =   8
      Top             =   2130
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   8250
      ScaleHeight     =   2985
      ScaleWidth      =   7935
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   450
      Width           =   7935
      Begin XtremeReportControl.ReportControl rptPlist 
         Height          =   1125
         Left            =   1260
         TabIndex        =   9
         Top             =   1380
         Width           =   2115
         _Version        =   589884
         _ExtentX        =   3731
         _ExtentY        =   1984
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picFilter 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   6555
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   300
         Width           =   6555
         Begin VB.OptionButton optFilter 
            Caption         =   "����"
            Height          =   225
            Index           =   6
            Left            =   5580
            TabIndex        =   36
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "��ִ��"
            Height          =   225
            Index           =   5
            Left            =   4650
            TabIndex        =   35
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "���ͼ�"
            Height          =   225
            Index           =   4
            Left            =   3720
            TabIndex        =   34
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "�Ѳ���"
            Height          =   225
            Index           =   3
            Left            =   2730
            TabIndex        =   33
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "�Ѱ�"
            Height          =   225
            Index           =   2
            Left            =   1770
            TabIndex        =   32
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "δ��"
            Height          =   225
            Index           =   1
            Left            =   840
            TabIndex        =   31
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "ȫ��"
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   30
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   885
         End
      End
      Begin XtremeSuiteControls.ShortcutCaption srtPatient 
         Height          =   285
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2235
         _Version        =   589884
         _ExtentX        =   3942
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
      Height          =   6570
      Left            =   30
      ScaleHeight     =   6570
      ScaleWidth      =   8145
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   8145
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
         Height          =   1785
         Left            =   60
         TabIndex        =   51
         Top             =   750
         Width           =   8025
         Begin VB.TextBox txt����1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5550
            MaxLength       =   5
            TabIndex        =   63
            Top             =   210
            Width           =   555
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5820
            TabIndex        =   62
            Top             =   1350
            Width           =   285
         End
         Begin VB.ComboBox cboҽ�� 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3570
            TabIndex        =   61
            Top             =   960
            Width           =   1275
         End
         Begin VB.ComboBox cbo�������� 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmLabSampling.frx":058A
            Left            =   900
            List            =   "frmLabSampling.frx":058C
            TabIndex        =   60
            Top             =   960
            Width           =   1635
         End
         Begin VB.TextBox txtҽ������ 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   900
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   59
            Top             =   1350
            Width           =   4935
         End
         Begin VB.TextBox txtID 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   600
            Width           =   1635
         End
         Begin VB.TextBox txtPatientDept 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3570
            TabIndex        =   57
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox txtBed 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5310
            TabIndex        =   56
            Top             =   960
            Width           =   795
         End
         Begin VB.ComboBox cboAge 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmLabSampling.frx":058E
            Left            =   4770
            List            =   "frmLabSampling.frx":05A1
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   210
            Width           =   750
         End
         Begin VB.TextBox txt���� 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4320
            MaxLength       =   5
            TabIndex        =   54
            Top             =   210
            Width           =   435
         End
         Begin VB.ComboBox cbo�Ա� 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            IMEMode         =   3  'DISABLE
            ItemData        =   "frmLabSampling.frx":05BD
            Left            =   3210
            List            =   "frmLabSampling.frx":05BF
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   210
            Width           =   675
         End
         Begin VB.TextBox txt���� 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   900
            MaxLength       =   20
            TabIndex        =   52
            ToolTipText     =   "����Ϊ���￨�š���������ͷΪ����ID��������סԺ�š���*������š���.���Һŵ��š���/���շѵ��ݺ�"
            Top             =   210
            Width           =   1635
         End
         Begin VB.Image imgPatient 
            Height          =   975
            Left            =   6390
            Picture         =   "frmLabSampling.frx":05C1
            Stretch         =   -1  'True
            Top             =   210
            Width           =   1125
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   510
            Index           =   10
            Left            =   7320
            TabIndex        =   74
            Top             =   1200
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҽ��"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   9
            Left            =   2790
            TabIndex        =   73
            Top             =   990
            Width           =   720
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   150
            TabIndex        =   72
            Top             =   990
            Width           =   720
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������Ŀ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   150
            TabIndex        =   71
            Top             =   1350
            Width           =   720
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��       ��"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   180
            TabIndex        =   70
            Top             =   255
            Width           =   675
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   510
            Index           =   6
            Left            =   6330
            TabIndex        =   69
            Top             =   1200
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��  ʶ ��"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   180
            TabIndex        =   68
            Top             =   645
            Width           =   675
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   4920
            TabIndex        =   67
            Top             =   1005
            Width           =   360
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ڿ���"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   2790
            TabIndex        =   66
            Top             =   645
            Width           =   720
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ա�"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   2790
            TabIndex        =   65
            Top             =   255
            Width           =   360
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   3915
            TabIndex        =   64
            Top             =   255
            Width           =   360
         End
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   405
         Left            =   150
         TabIndex        =   47
         Top             =   300
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
         IDKindStr       =   $"frmLabSampling.frx":148B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "����"
         IDKind          =   -1
         DefaultCardType =   "0"
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.CheckBox chkFindMove 
         BackColor       =   &H00FDD6C6&
         Caption         =   "���ҵ����˺󽹵��ƶ�����������"
         Height          =   225
         Left            =   5100
         TabIndex        =   7
         Top             =   30
         Width           =   3135
      End
      Begin VB.Frame fraBarCode 
         BackColor       =   &H00FDD6C6&
         Caption         =   "����󶨺�����"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3960
         Left            =   60
         TabIndex        =   3
         Top             =   2550
         Width           =   8025
         Begin VB.Frame Frame2 
            BackColor       =   &H00FDD6C6&
            BorderStyle     =   0  'None
            Height          =   3090
            Left            =   60
            TabIndex        =   18
            Top             =   840
            Width           =   7875
            Begin XtremeReportControl.ReportControl rptCuvette 
               Height          =   1935
               Left            =   90
               TabIndex        =   19
               Top             =   90
               Width           =   6375
               _Version        =   589884
               _ExtentX        =   11245
               _ExtentY        =   3413
               _StockProps     =   0
               AllowColumnRemove=   0   'False
               MultipleSelection=   0   'False
               SkipGroupsFocus =   0   'False
               EditOnClick     =   0   'False
               AutoColumnSizing=   0   'False
            End
            Begin VB.CheckBox chkMaterial 
               BackColor       =   &H00FDD6C6&
               Caption         =   "�Զ���������"
               Height          =   225
               Left            =   90
               TabIndex        =   49
               Top             =   2820
               Width           =   2835
            End
            Begin VB.CheckBox chkApplyDept 
               BackColor       =   &H00FDD6C6&
               Caption         =   "��������ʱ�����������"
               Height          =   225
               Left            =   4620
               TabIndex        =   48
               Top             =   2580
               Width           =   2835
            End
            Begin VB.CheckBox chkBindPage 
               BackColor       =   &H00FDD6C6&
               Caption         =   "���ɻ���������ת���Ѱ�ҳ"
               Height          =   195
               Left            =   90
               TabIndex        =   46
               Top             =   2595
               Width           =   3015
            End
            Begin VB.CheckBox chkSendPrint 
               BackColor       =   &H00FDD6C6&
               Caption         =   "ȡ���ͼ쵥��ӡ"
               Height          =   195
               Left            =   90
               TabIndex        =   38
               Top             =   2370
               Width           =   1755
            End
            Begin VB.CheckBox chkDeptShow 
               BackColor       =   &H00FDD6C6&
               Caption         =   "ֻ��ʾ��ǰ�ɼ������Թ�"
               Height          =   225
               Left            =   1890
               TabIndex        =   37
               Top             =   2130
               Width           =   2505
            End
            Begin VB.CommandButton cmdBindBarCode 
               Caption         =   "������(&B)"
               Height          =   345
               Left            =   6600
               TabIndex        =   28
               Top             =   75
               Width           =   1185
            End
            Begin VB.CommandButton cmdNewBarcode 
               Caption         =   "��������(&N)"
               Height          =   345
               Left            =   6600
               TabIndex        =   27
               Top             =   465
               Width           =   1185
            End
            Begin VB.CheckBox chkBackBill 
               BackColor       =   &H00FDD6C6&
               Caption         =   "����ɴ�ӡ��ִ��"
               Height          =   225
               Left            =   4620
               TabIndex        =   26
               Top             =   2130
               Width           =   1785
            End
            Begin VB.CheckBox chkComPlete 
               BackColor       =   &H00FDD6C6&
               Caption         =   "���ɻ��������־Ϊ�Ѳɼ�"
               Height          =   225
               Left            =   4620
               TabIndex        =   25
               Top             =   2370
               Width           =   2835
            End
            Begin VB.CommandButton cmdBarcodePrint 
               Caption         =   "�����ӡ(&B)"
               Height          =   345
               Left            =   6600
               Picture         =   "frmLabSampling.frx":1530
               TabIndex        =   24
               Top             =   1260
               Width           =   1185
            End
            Begin VB.CommandButton cmdComplete 
               Caption         =   "��ɲɼ�(&P)"
               Height          =   345
               Left            =   6600
               Picture         =   "frmLabSampling.frx":167A
               TabIndex        =   23
               Top             =   870
               Width           =   1185
            End
            Begin VB.CommandButton cmdBakBillPrint 
               Caption         =   "��ִ����ӡ"
               Height          =   345
               Left            =   6600
               Picture         =   "frmLabSampling.frx":17C4
               TabIndex        =   22
               Top             =   1665
               Width           =   1185
            End
            Begin VB.CheckBox ChkBarCodePrint 
               BackColor       =   &H00FDD6C6&
               Caption         =   "���ɻ��������ӡ����"
               Height          =   225
               Left            =   1890
               TabIndex        =   21
               Top             =   2370
               Width           =   2505
            End
            Begin VB.CheckBox chkPrintBarCode 
               BackColor       =   &H00FDD6C6&
               Caption         =   "����ɴ�ӡ����"
               Height          =   225
               Left            =   90
               TabIndex        =   20
               Top             =   2130
               Width           =   1575
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FDD6C6&
            BorderStyle     =   0  'None
            Height          =   660
            Left            =   60
            TabIndex        =   11
            Top             =   210
            Width           =   7875
            Begin VB.TextBox TxtBarCode 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   990
               TabIndex        =   15
               Top             =   75
               Width           =   2145
            End
            Begin VB.TextBox TxtBarCodeCheck 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   4200
               TabIndex        =   14
               Top             =   75
               Width           =   2145
            End
            Begin VB.Frame fraSpace 
               BackColor       =   &H00FDD6C6&
               Height          =   45
               Left            =   60
               TabIndex        =   13
               Top             =   540
               Width           =   7785
            End
            Begin VB.CheckBox ChkContinuous 
               BackColor       =   &H00FDD6C6&
               Caption         =   "��������"
               Height          =   225
               Left            =   6480
               TabIndex        =   12
               Top             =   165
               Width           =   1095
            End
            Begin VB.Label LabCap 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   90
               TabIndex        =   17
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ȷ��"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   3300
               TabIndex        =   16
               Top             =   165
               Width           =   840
            End
         End
      End
      Begin VB.TextBox txtGoto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   795
         TabIndex        =   0
         ToolTipText     =   "��������ͷΪ����ID��������סԺ�š���*������š���.���Һŵ��š���/���շѵ��ݺ�"
         Top             =   285
         Width           =   7275
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
         TabIndex        =   2
         Top             =   30
         Width           =   930
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   9270
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampling.frx":190E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampling.frx":197A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampling.frx":1F14
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampling.frx":24AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampling.frx":2A48
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   7230
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLabSampling.frx":2FE2
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22013
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
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   900
      Left            =   0
      TabIndex        =   6
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
   Begin VB.Image imgLoad 
      Height          =   975
      Left            =   9390
      Picture         =   "frmLabSampling.frx":3876
      Stretch         =   -1  'True
      Top             =   3390
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image imgDefual 
      Height          =   975
      Left            =   8100
      Picture         =   "frmLabSampling.frx":4740
      Stretch         =   -1  'True
      Top             =   3270
      Visible         =   0   'False
      Width           =   1125
   End
   Begin XtremeSuiteControls.PopupControl PopupControl 
      Left            =   10020
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
      Width           =   300
      ShowDelay       =   6000
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   8370
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmLabSampling.frx":560A
      Left            =   8880
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLabSampling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mPcol                                  '�����б�
    ����ID = 1
    ����
    �ٴ�·������
    ��Դ
    ��������
    �Ա�
    ����
    ��ʶ��
    ����
    ���˿���
    ״̬
    ����
    ���￨
    ����
    �Һŵ�
    ��ִ��
    δ��
    �Ѱ�
    �Ѳ���
    �زɱ걾
    ���ͼ�
    �ϼ�
    ����ʱ��
    ��ҳID
End Enum

Private Enum mAcol                                  'ҽ���б�
    ���
    ID
    ѡ��
    ����
    ͼ��
    �ز�
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
    ��Ѫ��
    �Թ�����
    ����
    ������Դ
    �������
    Ӥ��
    ����
    ���ID
    ִ��״̬
    NO
    ���ʱ��
    ����ʱ��
    ��¼״̬
    ������
    �����ӡ
    ��������
    �ͳ�ʱ��
    Ӥ������
    �ɼ�����ID
    �ɼ�ִ�п���
    ������ĿID
    ������Ŀ���
    ����ִ�п���ID
    �Ʒ�״̬
    ��¼����
    Ӥ���Ա�
    �������ڿ���
End Enum
Private Enum mDkp                                   '����ID
    ������� = 0
    ҽ���б�
    �����б�
End Enum

Private Enum mFilter
    ��ʶ�� = 0
    ���￨
    ����
    ���ݺ�
    �걾
    �ɼ���ʽ
    ����
    סԺ
    ���
    ���ʱ��
    ���ͻ����ʱ��          '=0 ����ʱ�� = 1 ���ʱ��
    ��ʼʱ��
    ����ʱ��
    ��������
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

Private mlngDeptID As Long                                              '����ID
Private mlngKey As Long                                                 'ID
Private mstrPrivs As String                                             'Ȩ��
Private mblnSaveAdvice As Boolean                                       '�Ƿ���Ҫ����ҽ���������޸���Ժ���˱걾��Ϣ
Private PatientType As Integer, mlng����ID As Long, mstrNO As String    '�����շѵ��ݺ�
Private mblnBarCode As Boolean                                          '����
Private mlngReqDept As Long, mstrReqDoctor As String                    'Ĭ�ϵĵǼǿ��Һ�ҽ��
Private mstrKeys As String                                              '��ǰ���յ�����ҽ��ID
Private mintEditMode As Integer                                         '0�����ա�1���Ǽǡ�2�����º��ա�3����������
Private rsRelativeAdvice As ADODB.Recordset                             '�Ǽǵ����ҽ��
Private mlngCapID As Long                                               '�ɼ���ĿID
Private mstrExtData  As String                                           '�Ǽǵ�������Ŀ��Ϣ
Private ItemDeptID As Long, mlngDefaultDevice As Long
Private mbln΢������Ŀ As Boolean
Private mblFind As Boolean                                              '�Ƿ���ҵ��Ĳ���
Private mstrOldTime As String                                           '��¼�ɵ�ʱ�����ڶ�ʱ����
Private iInputType As Integer
Private mintTop As Integer                                              '��������
Private mintHeight As Integer                                           '�����
Private mobjSquareCard As Object                                        'ȡ������
Private mblnNowConsumption As Boolean                                   '�Ƿ���������

Private mblnShowPwd As Boolean                                          '�Ƿ���ʾ����
    
Private mstrIndex As String                                             '���˲��ҷ�ʽ
'����������ǰ����״̬�����һֱ�Ը�״̬���Բ�����ǰ����
'0�����￨
'1������ID
'2��סԺ��
'3�������
'4���Һŵ�
'5���շѵ��ݺ�
'6������
'-------------------------------------------- 2007-08-17 ����һ��֧ͨ��
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private Enum IDKinds
    C0���� = 0
    C1ҽ���� = 1
    C2���֤�� = 2
    C3IC���� = 3
    C4����� = 4
    C5���￨ = 5
End Enum
Private mbln���֤ As Boolean
Private Const conMenu_IDkind_Change  As Integer = 12345
Private mstrBarCodes As String                                          'ѡ��ǰ�����봮ʹ�ö��ŷָ��������


Private Function ReadPatPricture(ByVal lng����ID As Long, ByRef imgPatient As Image, Optional ByRef strFile As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ƭ
    '������lng����ID=��ȡָ�����˵���Ƭ
    '           imgPatient=��Ƭ����λ��
    '           strFile=��Ƭ�ı���·��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    imgPatient.Picture = Nothing
    strFile = ""
    strFile = Sys.ReadLob(glngSys, 27, lng����ID, strFile)
    If strFile <> "" Then
        imgPatient.Picture = LoadPicture(strFile)
        ReadPatPricture = True
        Kill strFile
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub CreateCbs()
    '���ܴ���������
    
    '�����˵�
    Dim Control As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim ControlFile As CommandBarPopup
    Dim ControlSelect As CommandBarPopup
    Dim intBarCode As Integer                       'ʹ������ 1=39Code , 2=128Code(Ĭ��)
    Dim intExecDept As Integer                      '������ִ�п��Ҵ�ӡ
    Dim intHideBarCode As Integer                   '����Ԥ������
    Dim lngTop As Long, lngLeft As Long, lngRight As Long, lngBottom As Long
    
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
    '-----------------------------------------------------
   intBarCode = zlDatabase.GetPara("ʹ������", "100", "1211", 2)
   intExecDept = zlDatabase.GetPara("������ִ�п��Ҵ�ӡ", "100", "1211", 1)
   intHideBarCode = zlDatabase.GetPara("����Ԥ������", "100", "1211", 0)
    '==�ļ��˵�
    Set ControlFile = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    With ControlFile.CommandBar.Controls
        Set Control = .Add(xtpControlPopup, conMenu_File_PrintSet, "��ӡ����(&S)��")
            Control.CommandBar.Controls.Add xtpControlButton, conMenu_File_MedRecSetup, "�����ӡ����", -1, False
            Control.CommandBar.Controls.Add xtpControlButton, conMenu_File_MedRecPreview, "��ִ����ӡ����", -1, False
            
        .Add xtpControlButton, conMenu_File_Preview, "�嵥Ԥ��(&V)"
        .Add xtpControlButton, conMenu_File_Print, "�嵥��ӡ(&P)"
        Set Control = .Add(xtpControlPopup, conMenu_File_MedRecPrint, "��ӡ")
            Control.CommandBar.Controls.Add xtpControlButton, conMenu_File_RowPrint, "��ӡ����(&C)", -1, False
            Control.CommandBar.Controls.Add xtpControlButton, conMenu_File_BatPrint, "��ӡ��ִ��(&B)", -1, False
        Set Control = .Add(xtpControlPopup, conMenu_Edit_Send, "��������")
            Set cbrControl = Control.CommandBar.Controls.Add(xtpControlButton, conMenu_Tool_SignNew, "ʹ��39Code", -1, False)
            If intBarCode = 1 Then cbrControl.Checked = True
            Set cbrControl = Control.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendOther, "ʹ��128Code", -1, False)
            If intBarCode = 2 Then cbrControl.Checked = True
            Set cbrControl = Control.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Transfer_Force, "������ִ�п��Ҵ�ӡ", -1, False)
            If intExecDept = 0 Then cbrControl.Checked = True
        .Add xtpControlButton, conMenu_File_Excel, "�����&Excel��"
        Set Control = .Add(xtpControlButton, conMenu_File_Parameter, "�豸����", -1, False)

        Control.BeginGroup = True
        Set Control = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
        Control.BeginGroup = True
    End With
    
    '==�༭
    Set ControlFile = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    With ControlFile.CommandBar.Controls
        Set Control = .Add(xtpControlButton, conMenu_Manage_RequestView, "�����(&B)")
        Set Control = .Add(xtpControlButton, conMenu_Manage_RequestPrint, "��������(&N)")
        Set Control = .Add(xtpControlButton, conMenu_Manage_RequestBatPrint, "��ɲɼ�(&P)"): Control.BeginGroup = True
        Set Control = .Add(xtpControlButton, conMenu_Tool_MedRec, "���������ӡ(&A)"): Control.BeginGroup = True
        Set Control = .Add(xtpControlButton, conMenu_Manage_Plan, "ֱ�ӵǼ�(&G)"): Control.BeginGroup = True
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
        Set Control = .Add(xtpControlButton, conMenu_View_PriceTable, "����Ԥ������(&H)"): Control.BeginGroup = True
        Control.Checked = (intHideBarCode = 1): Call HideBarCode
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
    
    If Not mobjZLIHISPlugIn Is Nothing Then
        Dim astrPlug() As String
        Dim intLoop As Integer
        '==����˵�
        astrPlug = Split(mobjZLIHISPlugIn.GetFuncNames(glngSys, glngModul), ",")
        Set ControlFile = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PlugIn, "��չ����(&P)", -1, False)
        With ControlFile.CommandBar.Controls
            For intLoop = 0 To UBound(astrPlug)
                .Add xtpControlButton, conMenu_PlugIn_Menu + intLoop + 1, astrPlug(intLoop)
            Next
            Control.BeginGroup = True
        End With
    End If
    
    '==�б����
    Set Control = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "��������")
    Control.Flags = xtpFlagRightAlign
    Set Control = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlComboBox, conMenu_View_Busy, "����")
    Control.ShortcutText = "����"
    Control.Width = 130
    Control.Flags = xtpFlagRightAlign
    Control.Style = xtpButtonIconAndCaption
    
    '����������
    Dim Toolbar As CommandBar
    Dim ControlPopup As CommandBarPopup
    
    Set Toolbar = cbrthis.Add("������", xtpBarTop)
    Toolbar.ShowTextBelowIcons = False
    Toolbar.EnableDocking xtpFlagStretched
    With Toolbar.Controls
        .Add xtpControlButton, conMenu_File_Preview, "Ԥ��"
        .Add xtpControlButton, conMenu_File_Print, "��ӡ"
        Set Control = .Add(xtpControlButton, conMenu_Tool_MedRec, "���������ӡ"): Control.BeginGroup = True
        
        Set Control = .Add(xtpControlButton, conMenu_View_Show, "�ͼ�˶�")
        If InStr(GetPrivFunc(2500, 2001), "�ͼ�˶�") = 0 Then
            Control.Visible = False
        End If
        
        Set Control = .Add(xtpControlButton, conMenu_View_Filter, "����"): Control.BeginGroup = True
        .Add xtpControlButton, conMenu_View_Refresh, "ˢ��"
        Set Control = .Add(xtpControlButton, conMenu_Help_Help, "����"): Control.BeginGroup = True
        .Add xtpControlButton, conMenu_File_Exit, "�˳�"
    End With
    
    For Each Control In Toolbar.Controls
        Control.Style = xtpButtonIconAndCaption
    Next
    
    '�����
    With cbrthis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add FCONTROL, Asc("F"), conMenu_View_Filter
        .Add FCONTROL, Asc("A"), conMenu_Tool_MedRec
'        .Add FCONTROL, Asc("C"), conMenu_File_RowPrint
        .Add FCONTROL, Asc("B"), conMenu_File_BatPrint
        .Add 0, VK_F2, conMenu_Manage_RequestView
        .Add 0, VK_F3, conMenu_Manage_RequestPrint
        .Add 0, VK_F4, conMenu_Manage_RequestBatPrint
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F6, conMenu_Manage_Plan
        .Add 0, VK_F10, conMenu_IDkind_Change
        
    End With
    '���ò����ò˵�
    With cbrthis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    Call zlDatabase.ShowReportMenu(Me.cbrthis, glngSys, glngModul, mstrPrivs, "ZL1_INSIDE_1211_1", "ZL1_INSIDE_1211_2", "ZL1_INSIDE_1211_3")
    Me.cbrthis.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    mintTop = lngTop
End Sub
Private Sub CreateDkp()
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane

    dkpMan.SetCommandBars Me.cbrthis
    dkpMan.Options.DefaultPaneOptions = PaneNoCloseable
    dkpMan.Options.HideClient = True
    
    Set Pane1 = dkpMan.CreatePane(mDkp.�������, 400, 700, DockLeftOf, Nothing)
    Pane1.Title = "�������"
'    Pane1.Handle = Me.picBarCodeWork.hWnd
    Pane1.Options = PaneNoCaption
    
    Set Pane2 = dkpMan.CreatePane(mDkp.ҽ���б�, 400, 300, DockBottomOf, Pane1)
    Pane2.Title = "ҽ����Ϣ"
'    Pane2.Handle = TabCtr.hWnd
    Pane2.Options = PaneNoCaption
    
    Set Pane3 = dkpMan.CreatePane(mDkp.�����б�, 600, 300, DockRightOf, Nothing)
    Pane3.Title = "���˲ɼ��嵥"
'    Pane3.Handle = Me.picTab.hWnd
    Pane3.Options = PaneNoCaption
    
    Pane1.Select
End Sub

Private Sub cboAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo��������_Click()
    If cbo��������.ListIndex > -1 Then InitDoctors cbo��������.ItemData(cbo��������.ListIndex)
End Sub

Private Sub cbo��������_GotFocus()
    Call zlControl.TxtSelAll(cbo��������)
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    '</CSCustomCode> 1
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo��������_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo��������.ListIndex <> -1 Then mlngReqDept = Me.cbo��������.ItemData(Me.cbo��������.ListIndex): Exit Sub '��ѡ��
    If cbo��������.Text = "" Then '������
        Exit Sub
    End If
    
    strInput = UCase(NeedName(cbo��������.Text))
    'ȫԺ�ٴ�����
    strSQL = _
        " Select Distinct A.ID,A.����,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where B.����ID = A.ID " & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
        " And (B.�������� IN('�ٴ�','���'))" & _
        " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
        " Order by A.����"
    
    On Error GoTo errH
    vRect = GetControlRect(cboҽ��.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��������", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo��������.Height, blnCancel, False, True, strInput & "%", strInput & "%")
    If Not rsTmp Is Nothing Then
        If Not zlControl.CboLocate(cbo��������, rsTmp!����) Then
            cbo��������.Text = ""
        End If
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    If Me.cbo��������.ListIndex > -1 Then mlngReqDept = Me.cbo��������.ItemData(Me.cbo��������.ListIndex)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    End If
End Sub

Private Sub cboҽ��_GotFocus()
    Call zlControl.TxtSelAll(cboҽ��)
End Sub

Private Sub cboҽ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboҽ��_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cboҽ��.ListIndex <> -1 Then mstrReqDoctor = Me.cboҽ��.Text: Exit Sub '��ѡ��
    If cboҽ��.Text = "" Then '������
        Exit Sub
    End If
    
    strInput = UCase(NeedName(cboҽ��.Text))
    'ȫԺҽ��
    strSQL = "Select Distinct ����ID From ��������˵�� Where ������� IN(1,2,3)"
    strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
        " And B.����ID IN(" & strSQL & ")" & _
        " And (Upper(A.���) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
        " And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
        " Order by A.����"
    
    On Error GoTo errH
    vRect = GetControlRect(cboҽ��.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҽ��", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cboҽ��.Height, blnCancel, False, True, strInput & "%", strInput & "%")
    If Not rsTmp Is Nothing Then
        cboҽ��.Text = rsTmp!����
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ��ҽ����", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    If Len(Trim(Me.cboҽ��.Text)) > 0 Then mstrReqDoctor = Me.cboҽ��.Text
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strFilter As String                             '�����ִ�
    Dim cboCtrol As CommandBarComboBox                  '����
    Dim Controlcbo As CommandBarComboBox                '������
    Dim cbrControl As CommandBarControl                 '�ı���ǩ
    Dim intExecDept As Integer                          '��ִͬ�п��ҵ��Ƿ�һ���ӡ
    Dim intDay As Integer
    Dim strTmp As String
    
    Select Case Control.ID
    
        Case conMenu_File_MedRecSetup                                               '�����ӡ����
            ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1211_1", Me
        
        Case conMenu_File_MedRecPreview                                             '��ִ����ӡ����
            ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1211_2", Me
            
        Case conMenu_File_Preview                                                   '�嵥Ԥ��
            Call zlRptPrint(2)
        
        Case conMenu_File_Print                                                     '�嵥��ӡ
            Call zlRptPrint(1)
            
        Case conMenu_File_RowPrint                                                  '�����ӡ
            Call CmdBarCodePrint_Click
        
        Case conMenu_File_BatPrint                                                  '��ִ����ӡ
            Call cmdBakBillPrint_Click
                    
        Case conMenu_Tool_SignNew                                                   'ʹ��39��
            Control.Checked = True
            Set cbrControl = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Edit_SendOther, True, True)
            cbrControl.Checked = False
            
        Case conMenu_Edit_SendOther                                                 'ʹ��128��
            Control.Checked = True
            Set cbrControl = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Tool_SignNew, True, True)
            cbrControl.Checked = False
        
        Case conMenu_Manage_Transfer_Force                                          '�����ֿ��Ҵ�ӡ
            Control.Checked = Not Control.Checked
        
        Case conMenu_File_Excel                                                     '�����Excel
            Call zlRptPrint(3)
        
        Case conMenu_File_Parameter                                                 '�豸����
            'frmLabSampleSetup.Show vbModal, Me
            Call zlCommFun.DeviceSetup(Me, 100, 1101)
        Case conMenu_File_Exit                                                      '�˳�
            Unload Me
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Manage_RequestView                                             '�����
            Call cmdBindBarCode_Click
        
        Case conMenu_Manage_RequestPrint                                            '��������
            Call cmdNewBarcode_Click
        
        Case conMenu_Manage_RequestBatPrint                                         '��ɲɼ�
            Call cmdComplete_Click
        
        Case conMenu_Tool_MedRec                                                    '����������ӡ
            Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Manage_Transfer_Force, True, True)
            intExecDept = IIf(Control.Checked, 0, 1)
            Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Manage_Transfer_Force, True, True)
            intExecDept = IIf(Control.Checked, 0, 1)
            Set cbrControl = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Edit_SendOther, True, True)
            frmLabBarCodeBatPrint.ShowMe Me, mstrPrivs, IIf(cbrControl.Checked, 2, 1), intExecDept, mblnNowConsumption
        Case conMenu_Manage_Plan                                                    'ֱ�ӵǼ�
            If InStr(mstrPrivs, "ֱ�ӵǼ�") > 0 Then
                frmLabSamplingRegister.ShowMe Me
            End If
        Case conMenu_View_Show                                                      '�ͼ�˶�
            strTmp = zlDatabase.GetPara("�ɼ�����վ����", 100, 1211, "")
            If Me.rptPlist.Tag <> "" Then
                intDay = Val(Split(Me.rptPlist.Tag, ";")(mFilter.���ʱ��))
            Else
                If strTmp <> "" Then
                    intDay = Val(Split(strTmp, ";")(mFilter.���ʱ��))
                End If
            End If
            
            If Not mobjLisInsideComm Is Nothing Then
                Call mobjLisInsideComm.ShowFrmSampleSendCheck(Me, 1, intDay)
            End If
            
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
            frmLabSamplingFilter.ShowMe Me, strFilter
            Me.rptPlist.Tag = strFilter
            If strFilter <> "" Then RefreshPatientData
            
        Case conMenu_View_PriceTable                                                '����Ԥ������
            Control.Checked = Not Control.Checked
            HideBarCode
                    
        Case conMenu_View_Refresh                                                   'ˢ��
            Me.rptPlist.Tag = ""
            RefreshPatientData
        
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
            mlngDeptID = cboCtrol.ItemData(cboCtrol.ListIndex)
            RefreshPatientData
        Case conMenu_IDkind_Change
            Call IdKindChange
        Case Else

            If Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99 Then

                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
            Else
                Call mobjZLIHISPlugIn.ExecuteFunc(glngSys, glngModul, Control.Caption, mlngKey, txtID, mstrBarCodes)
            End If
        
    End Select
End Sub

Private Sub IdKindChange()
    If Me.ActiveControl Is txtGoto Then
       IDKind.IDKind = IIf(IDKind.IDKind = IDKinds.C5���￨, 0, IDKind.IDKind + 1)
    End If
End Sub

Private Sub cbrthis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.Visible = False Then Exit Sub
    If Me.stbThis.Visible = True Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    If Me.Visible = False Then Exit Sub
        
    Err = 0: On Error Resume Next
    Select Case Control.ID
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel, conMenu_File_ExportToXML         '�嵥��ӡ���
            Control.Enabled = (Me.rptPlist.Records.Count <> 0)
        
        Case conMenu_View_ToolBar_Button:                                                                   '��ť
            Control.Checked = Me.cbrthis(2).Visible
        
        Case conMenu_View_ToolBar_Text:                                                                     '��ť����
            Control.Checked = Not (Me.cbrthis(2).Controls(1).Style = xtpButtonIcon)
        
        Case conMenu_View_ToolBar_Size:                                                                     '��ͼ��
            Control.Checked = Me.cbrthis.Options.LargeIcons
        
        Case conMenu_View_StatusBar                                                                         '״̬��
            Control.Checked = Me.stbThis.Visible
            
        Case conMenu_File_RowPrint, conMenu_File_BatPrint                                                   '����ͻ�ִ����ӡ
            Control.Enabled = (Me.rptAlist(Me.TabCtr.Selected.Index).Rows.Count > 0 And Me.TabCtr.Selected.Index > 0 And Me.TabCtr.Selected.Index <> 4)
        
        Case conMenu_Manage_RequestView                                                                     '�����
            Control.Enabled = (Me.rptAlist(Me.TabCtr.Selected.Index).Rows.Count > 0 And Me.TabCtr.Selected.Index <= 1)
            Select Case Me.TabCtr.Selected.Index
                Case 0
                    Control.Caption = "������(&B)"
                Case 1, 2
                    Control.Caption = "�����(&B)"
            End Select
        Case conMenu_Manage_Transfer_Force
            Control.Enabled = InStr(mstrPrivs, "��������") > 0
            
        Case conMenu_Manage_RequestPrint                                                                    '��������
            Control.Enabled = (Me.rptAlist(Me.TabCtr.Selected.Index).Rows.Count > 0 And Me.TabCtr.Selected.Index <= 1)
            Select Case Me.TabCtr.Selected.Index
                Case 0
                    Control.Caption = "��������(&N)"
                Case 1, 2
                    Control.Caption = "ȡ������(&N)"
            End Select
            
        Case conMenu_Manage_RequestBatPrint                                                                 '��ɲɼ�
            Control.Enabled = (Me.rptAlist(Me.TabCtr.Selected.Index).Rows.Count > 0 And Me.TabCtr.Selected.Index >= 1)
            Select Case Me.TabCtr.Selected.Index
                Case 0, 1
                    Control.Caption = "��ɲɼ�(&P)"
                Case 2
                    Control.Caption = "ȡ�����(&P)"
                Case 3
                    Control.Caption = "ȡ�����(&P)"
                    Control.Enabled = False
            End Select
        Case conMenu_Tool_SignNew, conMenu_Edit_SendOther                                                   'ʹ��39���128��
            Control.Checked = Control.Checked
            Control.Enabled = InStr(mstrPrivs, "���������ӡ��ʽ") > 0
            
        Case conMenu_Manage_Transfer_Force                                                                  '�����ֿ��Ҵ�ӡ
            Control.Checked = Control.Checked
            
        Case conMenu_View_PriceTable                                                                        '����Ԥ������
            Control.Checked = Control.Checked
    End Select
    
    '���Ѿ��ձ걾
    On Error Resume Next
    If mstrOldTime = "" Then
        mstrOldTime = Now
    End If
    If DateDiff("n", mstrOldTime, Now) >= 1 Then
        showJuShouPait (1)
    End If
End Sub

Private Sub showJuShouPait(ByVal intType As Integer)
    '���Ѿ��ձ걾
    'intType ��ʾ/�ر���ʾ��  1=��ʾ,0=�ر�
    On Error Resume Next
    Dim PopupItem As PopupControlItem
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim strJuShou As String
    Dim var_Tmp As Variant
    Dim lngItemTop As Long
            
    For intLoop = 0 To Me.rptPlist.Rows.Count - 1
        If Val(Me.rptPlist.Rows(intLoop).Record(mPcol.����).Value) > 0 Then
            intCount = intCount + Val(Me.rptPlist.Rows(intLoop).Record(mPcol.����).Value)
            strJuShou = strJuShou & "|" & Me.rptPlist.Rows(intLoop).Record(mPcol.��������).Value & "," & Val(Me.rptPlist.Rows(intLoop).Record(mPcol.����).Value)
        End If
    Next
    If strJuShou <> "" Then strJuShou = Mid(strJuShou, 2)
    var_Tmp = Split(strJuShou, "|")
    If intCount > 0 Then
        If mstrOldTime = "" Then
            mstrOldTime = Now
        End If
        If intType = 1 Then
            With Me.PopupControl
                .RemoveAllItems
                '���ظ�����Ŀ
                For intLoop = 0 To UBound(var_Tmp)
                    lngItemTop = (.Height / intCount) + (intLoop * 17)
                    Set PopupItem = .AddItem(45, lngItemTop, 300, 200, Split(var_Tmp(intLoop), ",")(0) & " ��" & Split(var_Tmp(intLoop), ",")(1) & "���걾������,����鿴")
                    PopupItem.TextColor = vbRed
                Next
                .VisualTheme = xtpPopupThemeOffice2003
                .SetSize 300, 120
                .Font.Bold = True
                .Animation = xtpPopupAnimationSlide
                .ShowDelay = 9999999
                .Show
            End With
            mstrOldTime = ""
            If intCount = 0 Then
                Me.Caption = "����ɼ�����վ"
            Else
                Me.Caption = "����ɼ�����վ(����" & intCount & "�����ձ걾������գ�)"
            End If
        End If
    Else
        With Me.PopupControl
            .Close
        End With
        Me.Caption = "����ɼ�����վ"
    End If
End Sub

Private Sub chkfilter_Click(Index As Integer)
    Call RefreshPatientData
End Sub

Private Sub cmdBakBillPrint_Click()
    'ֻ��ӡ��ִ��
    If Me.cmdBakBillPrint.Caption = "��ִ����ӡ" Then
        WriterBarCode 4, False, False, True
    Else
        WriterBarCode 3, True, True, False
        Call showJuShouPait(1)
    End If
End Sub

Private Sub CmdBarCodePrint_Click()
    'ֻ��ӡ����
    WriterBarCode 4, False, True, False
    
End Sub

Private Sub cmdBindBarCode_Click()
        
    If Me.cbo�Ա�.Tag = "����" Then
        If Not ValidAdvice Then Exit Sub
        
        mlngKey = SaveAdviceData
        If mlngKey = 0 Then
            MsgBox "����ҽ��ʧ��!", vbInformation, gstrSysName
            Exit Sub
        End If
                
        If mlng����ID <> 0 Then
            Me.txtGoto.Text = "-" & mlng����ID
            Call txtGoto_KeyPress(13)
        End If
    End If
    
    
    
    
    If Me.cmdBindBarCode.Caption = "������(&B)" Then
        '����ȷ��
        If CheckMoeny = False Then
    '        MsgBox "ȷ�Ϸ��ò��ɹ���", gstrSysName
            Exit Sub
        End If
        WriterBarCode 0, IIf(chkComPlete.Value = 1, True, False), IIf(ChkBarCodePrint.Value = 1, True, False)
    Else
        WriterBarCode 2, False
    End If
    ' ˢ�²�����Ϣ
    If Not rptPlist.FocusedRow Is Nothing And rptAlist(TabCtr.Selected.Index).Rows.Count = 0 And optFilter(TabCtr.Selected.Index + 1).Value = True Then
        rptPlist.Records(rptPlist.FocusedRow.Record.Index).DeleteAll
        rptPlist.Rows(rptPlist.FocusedRow.Index).Record.Visible = False
        rptPlist.Populate
    End If
End Sub

Private Sub cmdComplete_Click()
    Dim strItem As String
    Dim intLoop As Integer
    If Me.cmdComplete.Caption = "��ɲɼ�(&P)" Then
        WriterBarCode 3, True, IIf(chkPrintBarCode.Value = 1, True, False), IIf(chkBackBill.Value = 1, True, False)
    Else
        '��ʾ
        With Me.rptAlist(Me.TabCtr.Selected.Index)
            For intLoop = 0 To .Records.Count - 1
                If .Records(intLoop).Item(mAcol.ѡ��).Checked = True Then
                    strItem = strItem & vbCrLf & .Records(intLoop).Item(mAcol.ҽ������).Value
                End If
            Next
        End With
        If strItem <> "" Then
            If MsgBox("�Ƿ�ȷ��Ҫȡ������ҽ���������?" & strItem, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            WriterBarCode 2, False
        Else
            MsgBox "û���ҵ����Բ�����ҽ��!", vbInformation, Me.Caption
        End If
    End If
    ' ˢ�²�����Ϣ
    If Not rptPlist.FocusedRow Is Nothing And rptAlist(TabCtr.Selected.Index).Rows.Count = 0 And optFilter(TabCtr.Selected.Index + 1).Value = True Then
        rptPlist.Records(rptPlist.FocusedRow.Record.Index).DeleteAll
        rptPlist.Rows(rptPlist.FocusedRow.Index).Record.Visible = False
        rptPlist.Populate
    End If
End Sub

Private Sub cmdNewBarcode_Click()
    Dim strItem As String
    Dim intLoop As Integer
    Dim strIDs As String
    Dim strName As String
    Dim blnPrint As Boolean
    
    Dim rsSampleCode As Recordset
    Dim strSampleCode As String
    Dim strSQL As String
    
    If Me.cbo�Ա�.Tag = "����" Then
        If Not ValidAdvice Then Exit Sub
        
        mlngKey = SaveAdviceData
        If mlngKey = 0 Then
            MsgBox "����ҽ��ʧ��!", vbInformation, gstrSysName
            Exit Sub
        End If
                
        If mlng����ID <> 0 Then
            Me.txtGoto.Text = "-" & mlng����ID
            Call txtGoto_KeyPress(13)
        End If
    End If
    
    If Me.cmdNewBarcode.Caption = "��������(&N)" Then
        '����ȷ��
        If CheckMoeny = False Then
    '        MsgBox "ȷ�Ϸ��ò��ɹ���", gstrSysName
            Exit Sub
        End If
        WriterBarCode 1, IIf(chkComPlete.Value = 1, True, False), _
                         IIf(ChkBarCodePrint.Value = 1, True, False), _
                         IIf(chkBackBill.Value = 1, True, False)
    ElseIf Me.cmdNewBarcode.Caption = "�ͼ�걾(&C)" Or Me.cmdNewBarcode.Caption = "ȡ���ͼ�(&C)" Then

        With Me.rptAlist(Me.TabCtr.Selected.Index)
            For intLoop = 0 To .Records.Count - 1
                If .Records(intLoop).Item(mAcol.ѡ��).Checked = True Then
                    strIDs = strIDs & .Records(intLoop).Item(mAcol.ID).Value & "," & .Records(intLoop).Item(mAcol.�ϲ�ҽ��).Value
                End If
            Next
            strIDs = Replace(Replace(strIDs, ";", ","), "|", ",")
            
            If Me.cmdNewBarcode.Caption = "�ͼ�걾(&C)" Then
                '���tat��ʱ
                If getTATTime(strIDs) = False Then
                    Exit Sub
                End If
                If strIDs = "" Then
                    Exit Sub
                End If
                  
            End If
            
            If strIDs = "" Then
                MsgBox "û���ҵ����Բ�����ҽ����¼!", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '�����ͳ�ʱ��
            If strIDs <> "" Then
                
                If Me.cmdNewBarcode.Caption = "�ͼ�걾(&C)" And chkSendPrint.Value <> 1 Then
                    If frmLabSamplingSendInfo.ShowMe(Me, strName, blnPrint) = False Then
                        Exit Sub
                    End If
                End If
                If strName = "" Then
                    strName = UserInfo.����
                End If
                

                '���ɷ�������
                strSQL = "select ����ҽ������_�걾��������.NEXTVAL  from dual"
                Set rsSampleCode = zlDatabase.OpenSQLRecord(strSQL, "�걾��������", "")
                
                gstrSql = "Zl_LisԤ������_�걾�ͳ�('" & strIDs & "'" & IIf(Me.cmdNewBarcode.Caption = "ȡ���ͼ�(&C)", ",1", ",0") & _
                          ",'" & strName & "','" & rsSampleCode(0) & "')"
                zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                'д���ͼ�ʱ�䵽�������뵥��
                Call WriterSampleSendDateToLIS(strIDs, IIf(Me.cmdNewBarcode.Caption = "ȡ���ͼ�(&C)", "1", "0"), strName)
            End If
            
            If Me.cmdNewBarcode.Caption = "�ͼ�걾(&C)" And chkSendPrint.Value <> 1 Then
            
'                If MsgBox("�Ƿ��ӡ�ͳ��嵥?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                 If blnPrint = True Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1211_3", Me, "ҽ���ִ�=" & strIDs, 2)
                 End If
            End If
            RefreshPatientData
    
            Me.TxtBarCode.Text = ""
            Me.TxtBarCode.Tag = ""
            Me.TxtBarCodeCheck.Text = ""
            Me.txtGoto.SetFocus
        End With
    ElseIf Me.cmdNewBarcode.Caption = "ȡ������(&N)" Then
        With Me.rptAlist(Me.TabCtr.Selected.Index)
            For intLoop = 0 To .Records.Count - 1
                If .Records(intLoop).Item(mAcol.ѡ��).Checked = True Then
                    strItem = strItem & vbCrLf & .Records(intLoop).Item(mAcol.ҽ������).Value
                End If
            Next
        End With
        If MsgBox("�Ƿ�ȷ��Ҫȡ������ҽ�����ݵ�����?" & strItem, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            WriterBarCode 2, False
        End If
    ElseIf Me.cmdNewBarcode.Caption = "�ò�����(&N)" Then
        WriterBarCode 3, True, True, False, 1
    End If
    
    If Me.cbo�Ա�.Tag = "����" Then
        Me.txt����.SetFocus
        Me.cbo�Ա�.Tag = ""
    End If
    ' ˢ�²�����Ϣ
    If Not rptPlist.FocusedRow Is Nothing And rptAlist(TabCtr.Selected.Index).Rows.Count = 0 And optFilter(TabCtr.Selected.Index + 1).Value = True Then
        rptPlist.Records(rptPlist.FocusedRow.Record.Index).DeleteAll
        rptPlist.Rows(rptPlist.FocusedRow.Index).Record.Visible = False
        rptPlist.Populate
    End If
End Sub

Private Sub cmdSelect_Click()
    Dim strExtData As String
    Dim rsTmp As New ADODB.Recordset
    
    strExtData = frmLabSamplingSelect.ShowMe(Me, mlngDeptID)
    If strExtData <> "" Then
        '��ȡ�ɼ���ʽ
        Set rsTmp = SelectCap(Split(Split(strExtData, ";")(0), ",")(0))
        If rsTmp Is Nothing Then
            MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, gstrSysName
            Exit Sub
        End If
        mlngCapID = rsTmp("ID")
        Call AdviceSet�������(3, strExtData)
        txtҽ������.Text = Get�����������(2, "")
        txtҽ������.Text = txtҽ������.Text & "(" & Split(strExtData, ";")(1) & ")"
    End If
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case mDkp.�������
            Item.Handle = Me.picBarCodeWork.hWnd
        Case mDkp.ҽ���б�
            Item.Handle = picAdvice.hWnd
        Case mDkp.�����б�
            Item.Handle = Me.picTab.hWnd
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    IDKind.ActiveFastKey
End Sub

Private Sub Form_Load()
    Dim intItem As Integer
    Dim bln�������� As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strIndex As String
    
     On Error Resume Next
    '���Ӳ���˵�
    If mobjZLIHISPlugIn Is Nothing Then
        Set mobjZLIHISPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not mobjZLIHISPlugIn Is Nothing Then
            Call mobjZLIHISPlugIn.Initialize(gcnOracle, glngSys, glngModul)
        End If
    End If
     On Error GoTo errH
    
    '---------------------------------------
    '�����ʼ��
    mstrPrivs = gstrPrivs                                       '��ʹ��Ȩ��
    CreateCbs                                                   '����������
    CreateDkp                                                   '��������
    CreateListHead                                              '������ͷ
    CreateTab                                                   '�����б�
    Call RestoreWinState(Me, App.ProductName)                   '����ָ�
    '---------------------------------------
    '---------------------------------------
    bln�������� = InStr(";" & mstrPrivs & ";", ";��������;")
    ChkBarCodePrint = zlDatabase.GetPara("����������ӡ", 100, 1211, 1, Array(ChkBarCodePrint), bln��������)
    chkComPlete = zlDatabase.GetPara("���ɺ���Ϊ�����", 100, 1211, 0, Array(chkComPlete), bln��������)
    chkBackBill = zlDatabase.GetPara("����ɺ��ӡ��ִ��", 100, 1211, 0, Array(chkBackBill), bln��������)
    chkDeptShow = zlDatabase.GetPara("ֻ��ʾ��ǰ�ɼ������Թ�", 100, 1211, 0, Array(chkDeptShow), bln��������)
    ChkContinuous = zlDatabase.GetPara("��������", 100, 1211, 0, Array(ChkContinuous), bln��������)
    chkFindMove = zlDatabase.GetPara("���Ҳ��˺����ƶ�", 100, 1211, 0, Array(chkFindMove), bln��������)
    chkPrintBarCode = zlDatabase.GetPara("����ɺ��ӡ����", 100, 1211, 0, Array(chkPrintBarCode), bln��������)
    chkSendPrint = zlDatabase.GetPara("ȡ���ͼ쵥��ӡ", 100, 1211, 0, Array(chkSendPrint), bln��������)
    chkBindPage = zlDatabase.GetPara("��ת���Ѱ�ҳ", 100, 1211, 0, Array(chkBindPage), bln��������)
    chkApplyDept = zlDatabase.GetPara("��������ʱ�����������", 100, 1211, 0, Array(chkApplyDept), bln��������)
    chkMaterial = zlDatabase.GetPara("�Զ���������", 100, 1211, 0, Array(chkMaterial), bln��������)
    mblnNowConsumption = zlDatabase.GetPara("��Ŀִ��ǰ�������շѻ��ȼ������", 100, , False)
    intItem = zlDatabase.GetPara("������Ϣ����", 100, 1211, 0)
    Me.optFilter(intItem).Value = True
    
    '---------------------------------------
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    Set mobjICCard.gcnOracle = gcnOracle
    
    
    mbln���֤ = False
    
    '���ݶ���
    Call GetDept                                                '�������
    Call InitDepts                                              '�����������
    
    If mobjSquareCard Is Nothing Then
        Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
            MsgBox "IDKind��ʼ��ʧ��!", vbInformation, gstrSysName
        Else
            IDKind.IDKindStr = mobjSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
        End If
    End If
    
    RefreshPatientData                                          '��������
    
    
        If mobjLisInsideComm Is Nothing Then
            Dim strErr As String
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

    
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjSquareCard, , txtGoto)
    '62760;���没�˲��ҷ�ʽ
    strIndex = zlDatabase.GetPara("���˲��ҷ�ʽ", 100, 1211, 0)
    IDKind.IDKind = CInt(Val(strIndex))
    Exit Sub
errH:
    MsgBox "��ʼ������ʱ�����������鲿�������ԣ�", vbInformation, "��ʼ��"
End Sub

Private Sub Form_Resize()
    Dim Pane1 As Pane


    If Me.Visible = False Then Exit Sub



    Set Pane1 = Me.dkpMan.FindPane(mDkp.�������)
    Dim Control As CommandBarControl
    Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_View_PriceTable, True, True)

    On Error Resume Next

    Pane1.MaxTrackSize.SetSize 8145 / Screen.TwipsPerPixelX, IIf(Control.Checked, 5500, 6300) / Screen.TwipsPerPixelY + 15
    Pane1.MinTrackSize.SetSize 8145 / Screen.TwipsPerPixelX, IIf(Control.Checked, 5500, 6300) / Screen.TwipsPerPixelY

    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters

    Pane1.MinTrackSize.SetSize 100, 100
    '�����������󻯰�ťʱ,���ı����ý���
    Select Case Me.WindowState
        Case vbMaximized    '��󻯰�ť
            Me.txtGoto.SetFocus
    End Select
End Sub
Private Sub CreateListHead()
    '�����б�ͷ
    Dim Column As ReportColumn
    Dim intLoop As Integer
    
    '==�����б�ͷ
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
        Set Column = .Add(mPcol.�ٴ�·������, "�ٴ�·������", 18, False): Column.Icon = 4
        Set Column = .Add(mPcol.��Դ, "��Դ", 30, True)
        Set Column = .Add(mPcol.��������, "��������", 75, True)
        Set Column = .Add(mPcol.�Ա�, "�Ա�", 30, True)
        Set Column = .Add(mPcol.����, "����", 40, True)
        Set Column = .Add(mPcol.���˿���, "���˿���", 75, True)
        Set Column = .Add(mPcol.��ʶ��, "��ʶ��", 60, True)
        Set Column = .Add(mPcol.����, "����", 60, True)
        Set Column = .Add(mPcol.�Һŵ�, "�Һŵ�", 60, True): Column.Visible = False
        Set Column = .Add(mPcol.���￨, "���￨", 60, True): Column.Visible = False
        Set Column = .Add(mPcol.����, "����", 30, True)
        Set Column = .Add(mPcol.δ��, "δ��", 45, True)
        Set Column = .Add(mPcol.�Ѱ�, "�Ѱ�", 45, True)
        Set Column = .Add(mPcol.�Ѳ���, "�Ѳ���", 45, True)
        Set Column = .Add(mPcol.���ͼ�, "���ͼ�", 45, True)
        Set Column = .Add(mPcol.����, "����", 30, True)
        Set Column = .Add(mPcol.�زɱ걾, "�ز�", 30, True)
        Set Column = .Add(mPcol.��ִ��, "��ִ��", 45, True)
        Set Column = .Add(mPcol.�ϼ�, "�ϼ�", 45, True)
        Set Column = .Add(mPcol.����ʱ��, "����ʱ��", 75, True)
        Set Column = .Add(mPcol.״̬, "״̬", 30, False): Column.Visible = False
        Set Column = .Add(mPcol.����, "����", 30, False): Column.Visible = False
    End With
    
    '==ҽ���б�ͷ
    For intLoop = 0 To 5
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
            Set Column = .Add(mAcol.���, "���", 120, False): Column.Visible = False
            Set Column = .Add(mAcol.ID, "ID", 0, False): Column.Visible = False
            Set Column = .Add(mAcol.ѡ��, "Check", 18, False): Column.Icon = 0
            Set Column = .Add(mAcol.����, "", 18, False): Column.Icon = 2
            Set Column = .Add(mAcol.ͼ��, "", 18, False): Column.Icon = 3
            Set Column = .Add(mAcol.��¼״̬, "�շ�", 30, False): Column.Alignment = xtpAlignmentCenter
            Set Column = .Add(mAcol.�ز�, "�ز�", 30, False): Column.Alignment = xtpAlignmentCenter
            Set Column = .Add(mAcol.�ɼ���ʽ, "�ɼ���ʽ", 75, True)
            Set Column = .Add(mAcol.�걾, "�걾", 55, True)
            Set Column = .Add(mAcol.ҽ������, "ҽ������", 75, True)
            Set Column = .Add(mAcol.����, "����", 75, True)
            Set Column = .Add(mAcol.Ӥ������, "Ӥ��", 75, True)
            Set Column = .Add(mAcol.ִ�п���, "ִ�п���", 75, True)
            Set Column = .Add(mAcol.����ҽ��, "����ҽ��", 75, True)
            Set Column = .Add(mAcol.����ʱ��, "����ʱ��", 75, True)
            Set Column = .Add(mAcol.������, "������", 65, True)
            Set Column = .Add(mAcol.����ʱ��, "����ʱ��", 75, True)
            Set Column = .Add(mAcol.����ʱ��, "����ʱ��", 75, True)
            Set Column = .Add(mAcol.�Թ���ɫ, "��ɫ����", 18, True): Column.Visible = False
            Set Column = .Add(mAcol.�Թܱ���, "�Թܱ���", 18, True): Column.Visible = False
            Set Column = .Add(mAcol.������, "������", 60, True)
            Set Column = .Add(mAcol.NO, "���ݺ�", 60, True)
            Set Column = .Add(mAcol.���ʱ��, "���ʱ��", 75, True)
            Set Column = .Add(mAcol.����ʱ��, "����ʱ��", 75, True)
            Set Column = .Add(mAcol.��Ѫ��, "��Ѫ��", 60, True): Column.Visible = False
            Set Column = .Add(mAcol.�Թ�����, "�Թ�����", 60, True): Column.Visible = False
            Set Column = .Add(mAcol.����, "����", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.������Դ, "������Դ", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.�������, "�������", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.Ӥ��, "Ӥ��", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.����, "����", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.���ID, "���ID", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.ִ��״̬, "ִ��״̬", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.������, "������", 50, True): Column.Visible = False
            Set Column = .Add(mAcol.�����ӡ, "��ӡ", 65, True)
            Set Column = .Add(mAcol.��������, "��������", 400, True)
            Set Column = .Add(mAcol.�ͳ�ʱ��, "�ͳ�ʱ��", 120, True)
            Set Column = .Add(mAcol.�ɼ�����ID, "�ɼ�����ID", 120, False): Column.Visible = False
            Set Column = .Add(mAcol.�ɼ�ִ�п���, "�ɼ�ִ�п���", 120, False): Column.Visible = False
            Set Column = .Add(mAcol.������ĿID, "������ĿID", 120, False): Column.Visible = False
            Set Column = .Add(mAcol.������Ŀ���, "������Ŀ���", 120, False): Column.Visible = False
            Set Column = .Add(mAcol.����ִ�п���ID, "����ִ�п���ID", 120, False): Column.Visible = False
            Set Column = .Add(mAcol.�Ʒ�״̬, "�Ʒ�״̬", 120, False): Column.Visible = False
            Set Column = .Add(mAcol.��¼����, "��¼����", 120, False): Column.Visible = False
            Set Column = .Add(mAcol.�������ڿ���, "�������ڿ���", 120, False): Column.Visible = False
        End With
    Next
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
        Set Column = .Add(mCuvette.����, "����", 55, True)
        Set Column = .Add(mCuvette.����, "����", 80, True)
        Set Column = .Add(mCuvette.��Ӽ�, "��Ӽ�", 90, True)
        Set Column = .Add(mCuvette.��Ѫ��, "��Ѫ��", 60, True)
        Set Column = .Add(mCuvette.���, "���", 60, True)
        Set Column = .Add(mCuvette.��ɫ, "", 18, True): Column.Icon = 3
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Controlcbo As CommandBarComboBox
    Dim Control As CommandBarControl
    Dim strTmp As String
    Dim aItem() As String
    Dim intLoop As Integer
    
    Call SaveWinState(Me, App.ProductName)
'    Me.Visible = False


    zlDatabase.SetPara "����������ӡ", ChkBarCodePrint, 100, 1211
    zlDatabase.SetPara "���ɺ���Ϊ�����", chkComPlete, 100, 1211
    zlDatabase.SetPara "����ɺ��ӡ��ִ��", chkBackBill, 100, 1211
    zlDatabase.SetPara "��������", ChkContinuous, 100, 1211
    zlDatabase.SetPara "���Ҳ��˺����ƶ�", chkFindMove, 100, 1211
    zlDatabase.SetPara "����ɺ��ӡ����", chkPrintBarCode, 100, 1211
    zlDatabase.SetPara "ֻ��ʾ��ǰ�ɼ������Թ�", chkDeptShow, 100, 1211
    zlDatabase.SetPara "ȡ���ͼ쵥��ӡ", chkSendPrint, 100, 1211
    zlDatabase.SetPara "��ת���Ѱ�ҳ", chkBindPage, 100, 1211
    zlDatabase.SetPara "��������ʱ�����������", chkApplyDept, 100, 1211
    zlDatabase.SetPara "�Զ���������", chkMaterial, 100, 1211
    '�������ID�ѱ��´�ʹ��
    'Controlcbo = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_View_Busy, True, True)
    Set Controlcbo = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_View_Busy, True, True)
    zlDatabase.SetPara "����", Controlcbo.ItemData(Controlcbo.ListIndex), 100, 1211
    Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Tool_SignNew, True, True)
    zlDatabase.SetPara "ʹ������", IIf(Control.Checked, 1, 2), 100, 1211
    Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Manage_Transfer_Force, True, True)
    zlDatabase.SetPara "������ִ�п��Ҵ�ӡ", IIf(Control.Checked, 0, 1), 100, 1211
    Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_View_PriceTable, True, True)
    zlDatabase.SetPara "����Ԥ������", IIf(Control.Checked, 1, 0), 100, 1211

    For intLoop = 0 To Me.optFilter.Count - 1
        If Me.optFilter(intLoop).Value = True Then
            zlDatabase.SetPara "������Ϣ����", intLoop, 100, 1211
            Exit For
        End If
    Next
    '62760;���没�˲��ҷ�ʽ
    zlDatabase.SetPara "���˲��ҷ�ʽ", mstrIndex, 100, 1211
    
    '�Ѽ��ʱ��ָ�Ϊ���3���
    strTmp = zlDatabase.GetPara("�ɼ�����վ����", 100, 1211, "")
    If strTmp <> "" Then
        aItem = Split(strTmp, ";")
        strTmp = ""
        For intLoop = 0 To UBound(aItem)
            If intLoop = mFilter.���ʱ�� Then
                strTmp = strTmp & ";" & "3"
            ElseIf intLoop = mFilter.��ʼʱ�� Then
                strTmp = strTmp & ";"
            ElseIf intLoop = mFilter.����ʱ�� Then
                strTmp = strTmp & ";"
            Else
                strTmp = strTmp & ";" & aItem(intLoop)
            End If
        Next
        strTmp = Mid(strTmp, 2)
        zlDatabase.SetPara "�ɼ�����վ����", strTmp, 100, 1211
    End If
    Set mobjSquareCard = Nothing
    Set mobjLisInsideComm = Nothing
    imgPatient.Picture = Nothing
'    If Not mobjIDCard Is Nothing Then
'        Call mobjIDCard.SetEnabled(False)
'    End If
'    If Not mobjICCard Is Nothing Then
'        Call mobjICCard.SetEnabled(False)
'    End If
    
    

'    Set mobjIDCard = Nothing
'    Set mobjICCard.gcnOracle = Nothing
'    Set mobjICCard = Nothing
'    Me.TabCtr.RemoveAll
'    Me.cbrthis.DeleteAll
'    Me.dkpMan.CloseAll
'    Me.dkpMan.DestroyAll

End Sub






Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
'    objPatiInfor.����ID = 216178
    If objPatiInfor.����ID <> 0 Then
        txtGoto.Text = "-" & objPatiInfor.����ID
    ElseIf objPatiInfor.����ID = 0 Then
        txtGoto.Text = objPatiInfor.����
    End If
    Call txtGoto_KeyPress(vbKeyReturn)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
' 2007-08-17 ����һ��֧ͨ��
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

Private Sub optFilter_Click(Index As Integer)
    If Index <> 0 Then
        TabCtr.Item(Index - 1).Selected = True
    End If
    RefreshPatientData
End Sub



Private Sub CreateTab()
    Dim Item As TabControlItem
    
    With Me.TabCtr
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.COLOR = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .InsertItem 0, "δ��", Me.rptAlist(0).hWnd, 0
        .InsertItem 1, "�Ѱ�", Me.rptAlist(1).hWnd, 0
        .InsertItem 2, "�Ѳ���", Me.rptAlist(2).hWnd, 0
        .InsertItem 3, "���ͼ�", Me.rptAlist(3).hWnd, 0
        .InsertItem 4, "��ִ��", Me.rptAlist(4).hWnd, 0
        .InsertItem 5, "����", Me.rptAlist(5).hWnd, 0
        .PaintManager.LayOut = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .Item(0).Selected = True
    End With
End Sub

Private Sub picAdvice_Resize()
    With TabCtr
        .Top = 0
        .Left = 0
        .Width = Me.picAdvice.ScaleWidth
        .Height = Me.picAdvice.ScaleHeight
    End With
End Sub

Private Sub picFilter_Resize()
    On Error Resume Next
    If Me.Visible = False Then Exit Sub
'    With srtFilter
'        .Top = 0
'        .Left = 0
'        .Width = Me.picFilter.Width
'        .Height = Me.picFilter.Height
'    End With
End Sub

Private Sub picTab_Resize()
    On Error Resume Next
    If Me.Visible = False Then Exit Sub
    Me.srtPatient.Top = 0
    Me.srtPatient.Left = 0
    Me.srtPatient.Width = Me.picTab.ScaleWidth
    
    Me.picFilter.Top = Me.srtPatient.Top + Me.srtPatient.Height + 5
    Me.picFilter.Left = 0
    Me.picFilter.Width = Me.ScaleWidth
    
    Me.rptPlist.Top = Me.picFilter.Top + Me.picFilter.Height + 10
    Me.rptPlist.Left = 0
    Me.rptPlist.Width = Me.picTab.ScaleWidth
    Me.rptPlist.Height = Me.picTab.ScaleHeight - picFilter.Top - picFilter.Height
End Sub


Private Sub PopupControl_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    '������½ǵ�����ʾ���е���Ŀʱ,ѡ�в����б��еĲ���
    Dim strPaitName As String
    Dim rptRow As ReportRow
    
    Me.TabCtr.Item(5).Selected = True
    strPaitName = Mid(Item.Caption, 1, InStr(Item.Caption, " ") - 1)
    With Me.rptPlist
        For Each rptRow In .Rows
            If rptRow.Record(mPcol.��������).Value = strPaitName Then
                .FocusedRow = rptRow
                Exit For
            End If
        Next
    End With
End Sub

Private Sub rptAlist_ItemCheck(Index As Integer, ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim RecordC As ReportRecord
    Dim RecordA As ReportRecord
    Dim strResivePeo
    If Me.TabCtr.Selected.Index = 0 Then Exit Sub
    For Each RecordC In Me.rptAlist(Index).Records
        If RecordC(mAcol.����).Value = Item.Record(mAcol.����).Value Then
            RecordC(mAcol.ѡ��).Checked = Item.Checked
        End If
        If RecordC(mAcol.ѡ��).Checked = True Then
            If RecordC(mAcol.������).Value <> "" Then
                strResivePeo = RecordC(mAcol.������).Value
            End If
        End If
    Next
    If strResivePeo <> "" Then
        cmdComplete.Enabled = False
    Else
        cmdComplete.Enabled = True
    End If
    Me.rptAlist(Index).Populate
End Sub

Private Sub rptAlist_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim hitColumn As ReportColumn
    Dim Record As ReportRecord
    Dim blSelect As Boolean

    With Me.rptAlist(Index)
        Set hitColumn = .HitTest(X, Y).Column
        
        If Not hitColumn Is Nothing Then
            If hitColumn.Caption = "Check" And .HitTest(X, Y).ht = xtpHitTestHeader Then
                hitColumn.AutoSize = True
                If .Records.Count > 0 Then blSelect = Not .Records(0).Item(mAcol.ѡ��).Checked
                For Each Record In .Records
                    Record.Item(mAcol.ѡ��).Checked = blSelect
                Next
            End If
        End If
        .Populate
    End With
End Sub


Private Sub rptAlist_SelectionChanged(Index As Integer)
    With Me.rptAlist(Me.TabCtr.Selected.Index)
        If Not .FocusedRow Is Nothing And .FocusedRow.GroupRow = False Then
            .PaintManager.HighlightBackColor = Val(.FocusedRow.Record(mAcol.�Թ���ɫ).Value)
            .Populate
            '�����������ʱ������
            On Error Resume Next
            Me.cbo��������.Text = .FocusedRow.Record(mAcol.�������).Value
            Me.cboҽ��.Text = .FocusedRow.Record(mAcol.����ҽ��).Value
            Me.txtҽ������.Text = .FocusedRow.Record(mAcol.ҽ������).Value
            txtPatientDept.Text = .FocusedRow.Record(mAcol.�������ڿ���).Value
            On Error GoTo 0
        Else
            .PaintManager.HighlightBackColor = vbWhite
        End If
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

Private Sub rptPlist_GotFocus()
    Me.dkpMan.RecalcLayout
End Sub

Private Sub RefreshPatientData()
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
    Dim str�������� As String, strSample As String, strCapture As String
    Dim NowDate As Date                                 '��ǰʱ��
    Dim intPatientType As Integer                       '������Դ
    Dim blnPathPatient As Boolean                       '�ٴ�·������
    Dim intPatientPage As Integer
    Dim strBloodSQL As String                           '��Ѫ��ѯ���
    On Error GoTo errH
    
    '��ע����ж�ȡ��������
    strTmp = zlDatabase.GetPara("�ɼ�����վ����", 100, 1211, "")
    
    NowDate = zlDatabase.Currentdate
    
    '�ӹ��˴������������ʱ����
    If Me.rptPlist.Tag <> "" Then
        varFilter = Split(Me.rptPlist.Tag, ";")
    Else
        If strTmp <> "" Then
            varFilter = Split(strTmp, ";")
        End If
    End If
    
    lblGoto.Caption = "���˲���"
    
    zlCommFun.ShowFlash "���ڸ�������,���Ժ�...", Me
    
    strSQL = "Select ����ID,������Դ,��������,���˿���,�Ա�,����,���￨��,��ʶ��,��ǰ����,��ҳID,Sum(decode(״̬,'�����',1,'ִ����',1,0)) As ��ִ��," & vbNewLine & _
                "       Sum(decode(״̬,'����',1,0)) As ����,Sum(decode(״̬,'δ��',1,0)) As δ��," & vbNewLine & _
                "       Sum(decode(״̬,'�Ѱ�',1,0)) As �Ѱ�,Sum(decode(״̬,'�Ѳ���',1,0)) As �Ѳ���, " & vbNewLine & _
                "       Sum(decode(����,'����',1,0)) as ����,Sum(�زɱ걾) as �زɱ걾,sum(decode(״̬,'���ͼ�',1,0)) as ���ͼ�,max(����ʱ��) as ����ʱ��,�ٴ�·������" & vbNewLine & _
                "From ("
    strSQL = strSQL & "Select    distinct a.����id,decode(a.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���') as ������Դ, " & vbCrLf & _
             " c.���� as ��������,e.���� as ���˿���,c.�Ա�,c.����,c.���￨��,b.��������, " & vbCrLf & _
             " decode(b.ִ��״̬,1,'�����',2,'����',3,'ִ����', " & vbCrLf & _
             "         Decode(b.��������, Null, 'δ��', Decode(b.������, Null, '�Ѱ�', decode(b.�걾�ͳ�ʱ��,null,'�Ѳ���','���ͼ�'))) )as ״̬, " & vbCrLf & _
             " decode(A.������Դ, 1, C.�����, 2, C.סԺ��,4,c.�����) As ��ʶ��, " & vbCrLf & _
             " decode(c.��ǰ����,null,decode(l.��Ժ����,null,l.��Ժ����,l.��Ժ����),c.��ǰ����) as ��ǰ���� , " & vbCrLf & _
             " decode(a.������־,1,'����',decode(g.����,1,'����')) as ���� ,decode(a.������Դ , 2,a.��ҳID,0) ��ҳID, " & vbCrLf & _
             " decode(b.ִ��״̬,0,'',2,'����') as ����,b.ִ��״̬,nvl(b.�زɱ걾,0) as �زɱ걾,b.����ʱ��,nvl(s.·��״̬,0) as �ٴ�·������,a.ҽ������ " & vbCrLf & _
             " From ����ҽ����¼ A, ����ҽ������ B, ������Ϣ C, ���ű� E, ������ĿĿ¼ F,���˹Һż�¼ G,����ҽ����¼ H, " & vbCrLf & _
             "      ������ĿĿ¼ K ,������ҳ L,����ҽ������ M,����걾��¼ J,������ҳ S " & vbCrLf & _
             " Where A.ID = H.���ID And H.id = B.ҽ��id And A.����id = C.����id And A.���˿���id = E.ID And A.������Ŀid+0 = f.ID  " & vbCrLf & _
             "      And h.������ĿID = k.id and a.id = j.ҽ��id(+)  " & vbCrLf & _
             " And A.�Һŵ� = G.No(+) and a.����id = g.����id(+)  and a.����id = g.����id(+)  and (g.����ID is null or (g.��¼״̬ =1 and g.��¼���� =1) ) And  f.��� = 'E' And f.�������� = '6' and a.����id = l.����ID(+) and a.��ҳID = l.��ҳID(+) and m.ִ�в���id + 0 = [1] " & vbCrLf & _
             " And A.ID = M.ҽ��ID And k.�Թܱ��� is not null and a.����ID = S.����ID(+) and a.��ҳID = s.��ҳID(+) " & IIf(Me.rptPlist.Tag = "", "and a.��ʼִ��ʱ�� < to_date('" & Format(NowDate, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')", "")
                 
    
    If Me.rptPlist.Tag <> "" Then
        strSQL = strSQL & " And A.������Դ in (" & IIf(Val(varFilter(mFilter.����)) = 1, "1,3", "0") & "," & _
                 Val(varFilter(mFilter.סԺ)) & "," & Val(varFilter(mFilter.���)) & ") "

        If varFilter(mFilter.��ʶ��) <> "" Then
            strSQL = strSQL & " And decode(a.������Դ,2,c.סԺ��,c.�����) = [2] "
        End If
        
        If varFilter(mFilter.���￨) <> "" Then
            strSQL = strSQL & " And c.���￨�� = [3] "
            
        End If
        
        If varFilter(mFilter.����) <> "" Then
            strSQL = strSQL & " And C.���� like [4] "
            
        End If
        
        If varFilter(mFilter.���ݺ�) <> "" Then
            strSQL = strSQL & " and B.NO = [5]"
            
        End If
        
        If UBound(varFilter) >= mFilter.�걾 Then
            If Trim(varFilter(mFilter.�걾)) <> "" Then
                strSQL = strSQL & " And instr([6],','||H.�걾��λ||',') > 0 "
                
            End If
        End If
        
        If UBound(varFilter) >= mFilter.�ɼ���ʽ Then
            If Trim(varFilter(mFilter.�ɼ���ʽ)) <> "" Then
                strSQL = strSQL & " And instr([7],','|| f.ID ||',') > 0 "
                
            End If
        End If
        
        If varFilter(mFilter.���ͻ����ʱ��) = 0 Then
            strSQL = strSQL & " and m.����ʱ�� Between [8] and [9]"
            
        Else
            strSQL = strSQL & " and m.����ʱ�� Between [8] and [9]"
        End If
        
        If varFilter(mFilter.��ʼʱ��) = "" Then
            strDateBegin = NowDate - Val(varFilter(mFilter.���ʱ��))
            strDateEnd = NowDate
        Else
            strDateBegin = varFilter(mFilter.��ʼʱ��)
            strDateEnd = varFilter(mFilter.����ʱ��)
        End If
        If UBound(varFilter) >= mFilter.�������� Then
            If Trim(varFilter(mFilter.��������)) <> "" Then
                lblGoto.Caption = "���˲���(" & Mid(Trim(varFilter(mFilter.��������)), 2) & ")"
                strSQL = strSQL & " And instr([10],','|| k.�������� ||',') > 0 "
                
            End If
        End If
    Else
        If strTmp <> "" Then
            strSQL = strSQL & " And instr('" & IIf(Val(varFilter(mFilter.����)) = 1, "1,3", "0") & "," & _
                 Val(varFilter(mFilter.סԺ)) & "," & Val(varFilter(mFilter.���)) & "',A.������Դ)>0 "

    
            If UBound(varFilter) >= mFilter.�걾 Then
                If Trim(varFilter(mFilter.�걾)) <> "" Then
                    strSQL = strSQL & " And instr([6],','||H.�걾��λ||',') > 0 "
                    
                End If
            End If
            
            If UBound(varFilter) >= mFilter.�ɼ���ʽ Then
                If Trim(varFilter(mFilter.�ɼ���ʽ)) <> "" Then
                    strSQL = strSQL & " And instr([7],','|| f.ID ||',') > 0 "
                    
                End If
            End If
            
            If varFilter(mFilter.���ͻ����ʱ��) = 0 Then
                strSQL = strSQL & " and m.����ʱ�� Between [8] and [9]"
                
            Else
                strSQL = strSQL & " and m.����ʱ�� Between [8] and [9]"
                
            End If
            
            If Val(varFilter(mFilter.���ʱ��)) >= 0 Then
                strDateBegin = NowDate - Val(varFilter(mFilter.���ʱ��))
                strDateEnd = NowDate
            Else
                strDateBegin = varFilter(mFilter.��ʼʱ��)
                strDateEnd = varFilter(mFilter.����ʱ��)
            End If
            If UBound(varFilter) >= mFilter.�������� Then
                If Trim(varFilter(mFilter.��������)) <> "" Then
                    lblGoto.Caption = "���˲���(" & Mid(Trim(varFilter(mFilter.��������)), 2) & ")"
                    strSQL = strSQL & " And instr([10],','|| k.�������� ||',') > 0 "
                    
                End If
            End If
        Else
            strSQL = strSQL & " and m.����ʱ�� Between [8] and [9]"
            
            strDateBegin = NowDate - 3
            strDateEnd = NowDate
        End If
    End If
    
    strBloodSQL = GetBooldPatientDataSql
    strSQL = strSQL & " union all " & strBloodSQL
    strSQL = strSQL & ") a group by ����Id,������Դ,��������,���˿���,�Ա�,����,���￨��,��ʶ��,��ǰ����,��ҳID,�ٴ�·������ "
    
    blnDateMoved = MovedByDate(CDate(strDateBegin)) '��ʱ�俴�Ƿ������ת��
    
    If blnDateMoved Then
        strSQL1 = strSQL
        strSQL1 = Replace(strSQL1, "����ҽ����¼", "H����ҽ����¼")
        strSQL1 = Replace(strSQL1, "����ҽ������", "H����ҽ������")
        strSQL = strSQL & " Union ALL " & strSQL1
    End If
    
    strSQL = strSQL & " Order by ���˿��� "
    blnPathPatient = False
    If strTmp = "" And Me.rptPlist.Tag = "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngDeptID, "", "", "", "", "", "", _
                    CDate(Format(strDateBegin, "yyyy-MM-dd 00:00:00")), CDate(Format(strDateEnd, "yyyy-MM-dd 23:59:59")), "")
    Else
        If UBound(varFilter) >= mFilter.�������� Then
            str�������� = varFilter(mFilter.��������) & ","
        End If
        If UBound(varFilter) >= mFilter.�걾 Then
            strSample = varFilter(mFilter.�걾) & ","
        End If
        If UBound(varFilter) >= mFilter.�ɼ���ʽ Then
            strCapture = varFilter(mFilter.�ɼ���ʽ) & ","
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngDeptID, Val(varFilter(mFilter.��ʶ��)), CStr(varFilter(mFilter.���￨)) _
                    , CStr(varFilter(mFilter.����)) & "%", CStr(varFilter(mFilter.���ݺ�)), strSample, strCapture _
                    , CDate(Format(strDateBegin, "yyyy-MM-dd 00:00:00")), _
                    CDate(Format(strDateEnd, "yyyy-MM-dd 23:59:59")), str��������)
    End If
    
    '�����¼
    Me.rptPlist.Records.DeleteAll
    Me.rptAlist(Me.TabCtr.Selected.Index).Records.DeleteAll
    Me.rptCuvette.Records.DeleteAll
    
    Do Until rsTmp.EOF
    
        Set Record = rptPlist.Records.Add
        For intLoop = 0 To Me.rptPlist.Columns.Count + 1
            Record.AddItem ""
        Next
        
        If Val(Nvl(rsTmp("�ٴ�·������"), 0)) = 1 Then
            blnPathPatient = True
            Record(mPcol.�ٴ�·������).Icon = 4
        Else
            Record(mPcol.�ٴ�·������).Icon = -1
        End If
        Record(mPcol.����ID).Value = Nvl(rsTmp("����ID"))
        Record(mPcol.��Դ).Value = Nvl(rsTmp("������Դ"))
        Record(mPcol.��������).Value = Nvl(rsTmp("��������"))
        Record(mPcol.���˿���).Value = Nvl(rsTmp("���˿���"))
        Record(mPcol.�Ա�).Value = Nvl(rsTmp("�Ա�"))
        Record(mPcol.����).Value = Nvl(rsTmp("����"))
        Record(mPcol.��ʶ��).Value = Nvl(rsTmp("��ʶ��"))
        Record(mPcol.����).Value = Nvl(rsTmp("��ǰ����"))
        Record(mPcol.����).Value = Nvl(rsTmp("����"))
        Record(mPcol.���￨).Value = Nvl(rsTmp("���￨��"))
        
        Record(mPcol.����).Value = Nvl(rsTmp("����"))
        Record(mPcol.δ��).Value = Nvl(rsTmp("δ��"))
        Record(mPcol.�Ѱ�).Value = Nvl(rsTmp("�Ѱ�"))
        Record(mPcol.�Ѳ���).Value = Nvl(rsTmp("�Ѳ���"))
        Record(mPcol.���ͼ�).Value = Nvl(rsTmp("���ͼ�"))
        Record(mPcol.����).Value = Nvl(rsTmp("����"))
        Record(mPcol.�زɱ걾).Value = Nvl(rsTmp("�زɱ걾"))
        Record(mPcol.��ִ��).Value = Nvl(rsTmp("��ִ��"))
        Record(mPcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
        Record(mPcol.��ҳID).Value = Nvl(rsTmp("��ҳID"), 0)
        
        Record(mPcol.�ϼ�).Value = Val(rsTmp("δ��")) + Val(rsTmp("�Ѱ�")) + Val(rsTmp("�Ѳ���")) + Val(rsTmp("����")) + Val(rsTmp("��ִ��"))
        
        lngPatientID = Nvl(rsTmp("����ID"))
        
        If Nvl(rsTmp("����")) > 0 Then
            For intLoop = 0 To Me.rptPlist.Columns.Count + 1
                Record(intLoop).ForeColor = vbRed
            Next
        End If
        
        rsTmp.MoveNext
    Loop
    
    '����
    Me.rptPlist.Populate
    Me.rptPlist.Columns(1).Visible = blnPathPatient

    Me.rptAlist(Me.TabCtr.Selected.Index).Populate
    Me.rptCuvette.Populate
    
    If strTmp <> "" Then
    
        Me.stbThis.Panels(2).Text = "��ǰ��Χ<" & IIf(varFilter(mFilter.���ͻ����ʱ��) = 0, "����ʱ�� ", "����ʱ�� ") & _
                                Format(strDateBegin, "yyyy-mm-dd") & "---" & Format(strDateEnd, "yyyy-mm-dd") & "> �¹���:" & _
                                Me.rptPlist.Rows.Count & "������."
    Else
        Me.stbThis.Panels(2).Text = "��ǰ��Χ<" & "����ʱ�� " & _
                                Format(strDateBegin, "yyyy-mm-dd") & "---" & Format(strDateEnd, "yyyy-mm-dd") & "> �¹���:" & _
                                Me.rptPlist.Rows.Count & "������."
    End If
                                
    '��λ���ϴ�ѡ�еĲ���
    If Me.Visible = True Then
        With Me.rptPlist
            
            For intLoop = 0 To .Rows.Count - 1
                If .Rows(intLoop).Record(mPcol.����ID).Value = mlngKey Then
                    Set .FocusedRow = .Rows(intLoop)
                    mlngKey = .Rows(intLoop).Record(mPcol.����ID).Value
                    intPatientPage = .Rows(intLoop).Record(mPcol.��ҳID).Value
                    .Populate
    '                Me.rptPlist.Tag = ""
                    Exit For
                End If
            Next
            
            If .FocusedRow Is Nothing And .Rows.Count > 0 Then
                Set .FocusedRow = .Rows(0)
                intPatientType = IIf(.Rows(0).Record(mPcol.��Դ).Value = "סԺ", 2, 1)
                mlngKey = .Rows(0).Record(mPcol.����ID).Value
                intPatientPage = .Rows(0).Record(mPcol.��ҳID).Value
                .Populate
            End If
            
            If Not .FocusedRow Is Nothing Then
                RefreshAdviceData mlngKey, Me.TabCtr.Selected.Index, intPatientType, False, intPatientPage
            End If
            
        End With
    End If
    '����������ִֻ��һ��
'    Me.rptPlist.Tag = ""
    
    If Me.rptPlist.Rows.Count = 0 Then
        txt���� = ""
        txt����.Tag = ""
        cbo�Ա�.ListIndex = -1
        txt���� = ""
        txt����1 = ""
        txtBed = ""
        txtID = ""
        txtPatientDept = ""
        cbo��������.ListIndex = -1
        cboҽ��.ListIndex = -1
        txtҽ������.Text = ""
        txtҽ������.Tag = ""
        Me.lblCap(6).Visible = False
    End If
    
    '���˲�����Ϣ�б�
    Call FilterPatient
    
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
    
    mlngDeptID = zlDatabase.GetPara("����", 100, 1211, 0)
    
    Set Controlcbo = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_View_Busy, True, True)
    
    On Error GoTo errH
    
    strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID " & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And B.������� IN(1,2,3,4) And B.�������� IN('����','����','����')"
            
    If InStr(1, mstrPrivs, "���п���") <= 0 Then
        strSQL = strSQL & " And C.��ԱID = [1] "
    End If
    
    strSQL = strSQL & " Order by A.����"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    Controlcbo.Clear
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
    
    '�Ա�
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("�Ա�")
    cbo�Ա�.Clear
    If Not rsTmp Is Nothing Then
        For intLoop = 1 To rsTmp.RecordCount
            cbo�Ա�.AddItem rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cbo�Ա�.ItemData(cbo�Ա�.NewIndex) = 1
                cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function RefreshAdviceData(lngPatientID As Long, intState As Integer, intPatientType As Integer, Optional blOnlyWhere As Boolean = False, Optional intPatientPage As Integer) As Boolean
    '���ܣ�                         ˢ�²ɼ�ҽ����¼
    '������                         lngpatientId = ����ID ,
    '                               intPatientType = ������Դ
    '                               intState = ��ǰ״̬ 0=δ�� 1=�Ѱ� 2=�Ѳ���
    '                               blOnlyWhere = true ֻʹ�ò���ID���в���
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
    Dim intShowButtom As Integer                    '�����ڲ���ʱ��ʾ��Щ��ť,���� not null = 1 ������ not null = 2 ��ִ�� = 3
    Dim strNO As String                             'NO��
    Dim blnShowExec As Boolean                      '�Ƿ���ʾδ�շ����ﲡ��
    Dim strAge As String                            '�������
    Dim aAge() As String
    Dim rsBaby As New ADODB.Recordset               'Ӥ��������ѯ
    Dim NowDate As Date                             '��ǰʱ��
    Dim strSQLbak As String
    Dim strҽ������ As String                       'ҽ������
    Dim strAdvice  As String                        '��¼��ҽ����Ϣ�����Ѿ���ѡ��ҽ��
    Dim strReceivePeople As String                 '��¼������
    Dim strSample As String, strCapture As String
    Dim blnFL As Boolean                            '�Ƿ������ʾ,������Ѫҽ��ʱ,������ʾ
    Dim strBooldSql As String                       '��ѯ��Ѫsql
    blnDateMoved = MovedByDate(Date) '��ʱ�俴�Ƿ������ת��
                 
    '��ע����ж�ȡ��������
    strTmp = zlDatabase.GetPara("�ɼ�����վ����", 100, 1211, "")
    
    NowDate = zlDatabase.Currentdate
    
    '�ӹ��˴������������ʱ����
    If Me.rptPlist.Tag <> "" Then
        varFilter = Split(Me.rptPlist.Tag, ";")
    Else
        If strTmp <> "" Then
            varFilter = Split(strTmp, ";")
        End If
    End If
    
    On Error GoTo errH
                 
    strSQL = " select /*+ rule */ distinct a.���,a.ҽ��id,a.���id,a.�Թ���ɫ,a.�ɼ���ʽ,a.ҽ������,a.��������,a.����ʱ��,a.ִ�п���,a.����ҽ��,a.����ʱ��,a.������,a.����ʱ��,a.�Թܱ���," & vbCrLf & _
             "        a.�걾,a.��������,a.�Ա�,a.����,a.����,a.��ʶ��,Decode(A.������Դ, 2, nvl(A.�������ڿ���,A.�������), A.�������) As �������ڿ���,a.����,a.����ID,a.������,a.��Ѫ��,a.�Թ�����,a.����,a.������Դ,a.Ӥ��,a.����,a.ִ��״̬,a.NO," & vbCrLf & _
             "        a.���ʱ��,a.�������,a.����ʱ��,a.�����Ŀ,b.���ʷ���,b.��¼״̬,a.������,b.��¼����,a.�ɼ���ĿID,a.���￨��,a.�Һŵ�,a.�����ӡ,a.ִ��˵��,a.�ز�, " & vbCrLf & _
             "        a.�걾�ͳ�ʱ��,a.��ҳID,a.ִ�п���ID,a.�ɼ�ִ�п���,a.������ĿID,a.����ִ�п���ID,b.�����־,�Ʒ�״̬,����ģʽ" & vbCrLf & _
             " from "
    strSQL = strSQL & " ( Select decode(d.���,'K','��Ѫ','����') ���, B.ID as ҽ��ID, B.���id, G.��ɫ As �Թ���ɫ, decode(d.���,'K',b.ҽ������ ,d.����) As �ɼ���ʽ, decode(d.���,'K',d.����,b.ҽ������) as ҽ������, C.��������,C.����ʱ��, " & vbCrLf & _
             "   H.���� As ִ�п���, B.����ҽ��,B.����ʱ��, C.������, C.����ʱ��, G.���� as �Թܱ���,b.�걾��λ as �걾, " & vbCrLf & vbCrLf & _
             "   I.���� as ��������,I.�Ա�,i.����,i.��ǰ���� as ����,decode(b.������Դ,1,I.�����,2,i.סԺ��,4,i.�����) as ��ʶ��, " & vbCrLf & _
             "   L.���� as �������ڿ���,Decode(C.ִ��״̬,2,'����') as ����,I.����ID,c.������,G.��Ѫ��,G.���� as �Թ�����, " & vbCrLf & _
             "   DECODE(B.������־,1,'����','') as ����,b.������Դ,nvl(b.Ӥ��,0) as Ӥ��,N.���� as ����,decode(d.���, 'K', M.ִ��״̬,C.ִ��״̬) ִ��״̬,C.NO,j.���ʱ��,o.���� as �������,m.����ʱ��, " & vbCrLf & _
             "   E.�����Ŀ,C.������,c.��¼����,decode(d.���,'K',e.id ,d.id)  as �ɼ���ĿID,i.���￨��,b.�Һŵ�,C.�����ӡ,C.ִ��˵��,nvl(c.�زɱ걾,0) as �ز�,c.�걾�ͳ�ʱ��, " & vbCrLf & _
             "   a.��ҳID,Decode(d.���, 'K', b.ִ�п���ID, a.ִ�п���ID) ִ�п���ID,P.���� as �ɼ�ִ�п���,b.������ĿID,Decode(d.���, 'K', a.ִ�п���ID, b.ִ�п���ID) as ����ִ�п���ID,c.�Ʒ�״̬,i.����ģʽ " & vbCrLf & _
             "   From ����ҽ����¼ A, ����ҽ����¼ B, ����ҽ������ C, ������ĿĿ¼ D, ������ĿĿ¼ E, " & vbCrLf & _
             "   ��Ѫ������ G,���ű� H, ������Ϣ I,���ű� L,����ҽ������ M,����걾��¼ J ,���ű� O ,���ű� P, " & vbCrLf & _
             "   (select ������ĿID,���� from ������Ŀ���� where ���� = 9 and ���� = 1 ) N " & vbCrLf & _
             "  Where A.ID = B.���id And B.ID = C.ҽ��id And A.������Ŀid = D.ID And B.������Ŀid = E.ID  " & vbCrLf & _
             "    And (e.��� = 'E' Or e.��� = 'C') And E.�Թܱ��� = G.���� And B.ִ�п���id = H.ID(+) and a.ִ�п���ID = P.id(+)  " & vbCrLf & _
             "    And  d.��� = 'E'  And d.�������� = '6'  And A.����id = [1] " & IIf(InStr(txtGoto.Text, ".") = 1, "", "And c.����ʱ��+0 Between [3] and [4] ") & IIf(Me.rptPlist.Tag = "", "and a.��ʼִ��ʱ�� < to_date('" & Format(NowDate, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')", "") & vbCrLf & _
             "    and  m.ִ�в���id + 0 = [2] And B.����ID = I.����ID and I.��ǰ����ID = L.ID(+) " & vbCrLf & _
             "    and a.id = m.ҽ��id And E.id = N.������ĿID(+) and a.id = j.ҽ��id(+) and b.��������id = o.id  " & vbCrLf & _
             "    ) a , (Select ҽ�����,��¼����,��¼״̬,���ʷ���,�����־ From סԺ���ü�¼ Where  ����ID=[1]) b " & vbCrLf & _
             "where a.ҽ��id = b.ҽ�����(+) and a.��¼���� = mod(b.��¼����(+),10)  "
    
               '  IIf(Me.rptPlist.Tag = "", "and a.��ʼִ��ʱ�� between  to_date( '" & Mid(NowDate, 1, InStr(NowDate, " ")) & " 00:00:00' ,'yyyy-mm-dd hh24:mi:ss' ) and to_date('" & Mid(NowDate, 1, InStr(NowDate, " ")) & " 23:59:59','yyyy-mm-dd hh24:mi:ss')", "") & vbCrLf & _

    '�������ֲ�ͬ��״̬,
    If intState = 0 Then
        strSQL = strSQL & " And a.�������� is null And a.ִ��״̬ in (0) " & vbCrLf
    ElseIf intState = 1 Then
        strSQL = strSQL & " And a.�������� is not null And a.����ʱ�� is  null And a.ִ��״̬ in (0) " & vbCrLf
    ElseIf intState = 2 Then
        strSQL = strSQL & " And a.�������� is not null and  a.����ʱ�� is not null And a.ִ��״̬ in (0) and a.�걾�ͳ�ʱ�� is null " & vbCrLf
    ElseIf intState = 3 Then
        strSQL = strSQL & " And a.�������� is not null and  a.����ʱ�� is not null And a.ִ��״̬ in (0) and a.�걾�ͳ�ʱ�� is not null  " & vbCrLf
    ElseIf intState = 4 Then
        strSQL = strSQL & " And a.ִ��״̬ in (1,3) " & vbCrLf
    ElseIf intState = 5 Then
        strSQL = strSQL & " And a.ִ��״̬ in (2) " & vbCrLf
    End If
    
    '����
    If Me.rptPlist.Tag <> "" Or strTmp <> "" Then
       
        If UBound(varFilter) >= mFilter.�걾 Then
            If Trim(varFilter(mFilter.�걾)) <> "" Then
                strSQL = strSQL & " And instr([5],','||a.�걾||',') > 0 "
            End If
        End If
        
        If UBound(varFilter) >= mFilter.�ɼ���ʽ Then
            If Trim(varFilter(mFilter.�ɼ���ʽ)) <> "" Then
                strSQL = strSQL & " And instr([6],','||a.�ɼ���ĿID||',') > 0 "
            End If
        End If
        
        If Me.rptPlist.Tag <> "" Then
            strDateBegin = varFilter(mFilter.��ʼʱ��)
            strDateEnd = varFilter(mFilter.����ʱ��)
        Else
            strDateBegin = NowDate - Val(varFilter(mFilter.���ʱ��))
            strDateEnd = NowDate
        End If
    Else
        strDateBegin = NowDate - 3
        strDateEnd = NowDate
    End If
    
    If intPatientPage <> 0 Then
        strSQL = strSQL & " and a.��ҳid = [9] "
    End If
    
    
    If blnDateMoved Then
        strSQL1 = strSQL
        strSQL1 = Replace(strSQL1, "����ҽ����¼", "H����ҽ����¼")
        strSQL1 = Replace(strSQL1, "����ҽ������", "H����ҽ������")
        strSQL = strSQL & " Union ALL " & strSQL1
    End If
                 
    If blOnlyWhere = True Then
        strSQL = " select /*+ rule */ distinct a.���,a.ҽ��id,a.���id,a.�Թ���ɫ,a.�ɼ���ʽ,a.ҽ������,a.��������,a.����ʱ��,a.ִ�п���,a.����ҽ��,a.����ʱ��,a.������,a.����ʱ��,a.�Թܱ���," & vbCrLf & _
             "        a.�걾,a.��������,a.�Ա�,a.����,a.����,a.��ʶ��,Decode(A.������Դ, 2, nvl(A.�������ڿ���,A.�������), A.�������) As �������ڿ���,a.����,a.����ID,a.������,a.��Ѫ��,a.�Թ�����,a.����,a.������Դ,a.Ӥ��,a.����,a.ִ��״̬,a.NO," & vbCrLf & _
             "        a.���ʱ��,a.�������,a.����ʱ��,a.�����Ŀ,b.���ʷ���,b.��¼״̬,a.������,b.��¼����,a.�ɼ���ĿID,a.���￨��,a.�Һŵ�,a.�����ӡ,a.ִ��˵��,a.�ز�, " & vbCrLf & _
             "        a.�걾�ͳ�ʱ��,a.��ҳID,a.ִ�п���ID,a.�ɼ�ִ�п���,a.������ĿID,a.����ִ�п���ID,b.�����־,a.�Ʒ�״̬ , ����ģʽ" & vbCrLf & _
             " from "
        strSQL = strSQL & " ( Select decode(d.���,'K','��Ѫ','����') ���, B.ID as ҽ��ID, B.���id, G.��ɫ As �Թ���ɫ,decode(d.���,'K',b.ҽ������ ,d.����) As �ɼ���ʽ, decode(d.���,'K',d.����,b.ҽ������) as ҽ������, C.��������,C.����ʱ��, " & vbCrLf & _
             "   H.���� As ִ�п���, B.����ҽ��,B.����ʱ��, C.������, C.����ʱ��, G.���� as �Թܱ���,b.�걾��λ as �걾, " & vbCrLf & vbCrLf & _
             "   I.���� as ��������,I.�Ա�,i.����,i.��ǰ���� as ����,decode(b.������Դ,1,I.�����,2,i.סԺ��,4,i.�����) as ��ʶ��, " & vbCrLf & _
             "   L.���� as �������ڿ���,Decode(C.ִ��״̬,2,'����') as ����,I.����ID,c.������,G.��Ѫ��,G.���� as �Թ�����, " & vbCrLf & _
             "   DECODE(B.������־,1,'����','') as ����,b.������Դ,nvl(b.Ӥ��,0) as Ӥ��,N.���� as ����,C.ִ��״̬ ִ��״̬,C.NO,j.���ʱ��,o.���� as �������,m.����ʱ��, " & vbCrLf & _
             "   E.�����Ŀ,C.������,c.��¼����,decode(d.���,'K',e.id ,d.id) as �ɼ���ĿID,i.���￨��,a.�Һŵ�,c.�����ӡ,C.ִ��˵��,nvl(c.�زɱ걾,0) as �ز�,c.�걾�ͳ�ʱ��, " & vbCrLf & _
             "   A.��ҳID,Decode(d.���, 'K', b.ִ�п���ID, a.ִ�п���ID) ִ�п���ID,P.���� as �ɼ�ִ�п���,b.������ĿID, b.ִ�п���ID as ����ִ�п���ID,c.�Ʒ�״̬ ,i.����ģʽ " & vbCrLf & _
             "   From ����ҽ����¼ A, ����ҽ����¼ B, ����ҽ������ C, ������ĿĿ¼ D, ������ĿĿ¼ E, " & vbCrLf & _
             "   ��Ѫ������ G,���ű� H, ������Ϣ I,���ű� L,����ҽ������ M,����걾��¼ J ,���ű� O ,���ű� P, " & vbCrLf & _
             "   (select ������ĿID,���� from ������Ŀ���� where ���� = 9 and ���� = 1 ) N " & vbCrLf & _
             "  Where A.ID = B.���id And B.ID = C.ҽ��id And A.������Ŀid = D.ID And B.������Ŀid = E.ID " & vbCrLf & _
             "    And (e.��� = 'E' Or e.��� = 'C') And E.�Թܱ��� = G.���� And B.ִ�п���id = H.ID(+) and a.ִ�п���ID = P.id(+) " & vbCrLf & _
             "    And  d.��� = 'E'  And d.�������� = '6' And A.����id = [1] " & IIf(InStr(txtGoto.Text, ".") = 1, "", "And c.����ʱ��+0 Between [3] and [4] ") & IIf(Me.rptPlist.Tag = "", "and a.��ʼִ��ʱ�� < to_date('" & Format(NowDate, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')", "") & vbCrLf & _
             "    and  m.ִ�в���id + 0 = [2] And B.����ID = I.����ID and I.��ǰ����ID = L.ID(+) " & vbCrLf & _
             "    and a.id = m.ҽ��id And E.id = N.������ĿID(+) and a.id = j.ҽ��id(+) and b.��������id = o.id  " & vbCrLf & _
             "    ) a , (Select ҽ�����,��¼����,��¼״̬,���ʷ���,�����־ From סԺ���ü�¼ Where    ����ID=[1]) b " & vbCrLf & _
             "where a.ҽ��id = b.ҽ�����(+) and a.��¼���� = b.��¼����(+) "
            
            '�������ֲ�ͬ��״̬
            If intState = 0 Then
                strSQL = strSQL & " And a.�������� is null And a.ִ��״̬ in (0,2) " & vbCrLf
            ElseIf intState = 1 Then
                strSQL = strSQL & " And a.�������� is not null And a.����ʱ�� is  null And a.ִ��״̬ in (0,2) " & vbCrLf
            ElseIf intState = 2 Then
                strSQL = strSQL & " And a.�������� is not null and  a.����ʱ�� is not null And a.ִ��״̬ in (0,2) and �걾�ͳ�ʱ�� is null " & vbCrLf
            ElseIf intState = 3 Then
                strSQL = strSQL & " And a.�������� is not null and  a.����ʱ�� is not null And a.ִ��״̬ in (0,2) and �걾�ͳ�ʱ�� is not null  " & vbCrLf
            ElseIf intState = 4 Then
                strSQL = strSQL & " And a.ִ��״̬ in (1,3) " & vbCrLf
            ElseIf intState = 5 Then
                strSQL = strSQL & " And a.ִ��״̬ in (2) " & vbCrLf
            End If
        
        '���ݺ�
'        If Mid(Me.txtGoto.Text, 1, 1) = "/" Then
'            strNO = Mid(Me.txtGoto, 2)
'            If IsNumeric(strNO) = True Then
'                strsql = strsql & " And a.NO = [7] "
'            End If
'        End If
        '���ݲ������ж��Ƿ񰴲ɼ���������ʾ
        If chkDeptShow.Value <> 1 Then
            strSQL = Replace(strSQL, " and  m.ִ�в���id + 0 = [2] ", "")
        End If
        
        If Mid(Me.txtGoto.Text, 1, 1) = "*" Or Mid(Me.txtGoto.Text, 1, 1) = "." Then
            strSQL = strSQL & " And a.������Դ in (1,3,4) "
        End If
        
        If Mid(Me.txtGoto.Text, 1, 1) = "+" Then
            strSQL = strSQL & " And a.������Դ in ( 2,4) "
        End If
        
        
        '����
        If BlnIsNumber(txtGoto) Then
            strSQL = strSQL & " And (a.�������� = [8] or a.���￨�� = [8]) "
        End If
        
    End If
    
    If intPatientType <> 2 Then
        strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
    End If
    strBooldSql = GetBloodAdviceSql(intState, intPatientType, blOnlyWhere, intPatientPage)
    strSQL = strSQL & " Union ALL " & strBooldSql & " order by ���,�Թܱ���,���ID,ִ�п���,�걾,��������,ҽ������,����ʱ��,�����Ŀ desc "
             
    If strTmp <> "" Or rptPlist.Tag <> "" Then
        If UBound(varFilter) >= mFilter.�걾 Then
            strSample = varFilter(mFilter.�걾) & ","
        End If
        If UBound(varFilter) >= mFilter.�ɼ���ʽ Then
            strCapture = varFilter(mFilter.�ɼ���ʽ) & ","
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientID, mlngDeptID, CDate(Format(strDateBegin, "yyyy-mm-dd 00:00:00")), _
                CDate(Format(strDateEnd, "yyyy-mm-dd 23:59:59")), strSample, strCapture, zlCommFun.GetFullNO(strNO), txtGoto, _
                intPatientPage)
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientID, mlngDeptID, CDate(Format(strDateBegin, "yyyy-mm-dd 00:00:00")), _
                CDate(Format(strDateEnd, "yyyy-mm-dd 23:59:59")), "", 0, zlCommFun.GetFullNO(strNO), txtGoto, intPatientPage)
    End If
    
    Me.rptAlist(Me.TabCtr.Selected.Index).Records.DeleteAll
    Me.rptCuvette.Records.DeleteAll
    
    Do Until rsTmp.EOF
        blnShowExec = True
        '����Ȩ�����ж��Ƿ���ʾδ�շѵ������¼
        If InStr(mstrPrivs, "��ʾ���ۼ�¼") <= 0 And Nvl(rsTmp("��¼״̬"), "NULL") <> "NULL" Then
            If Nvl(rsTmp("�����־"), 1) = 1 Then
                'ֻ�������ﲡ��
                If Nvl(rsTmp("��¼״̬"), 0) <> 1 Then blnShowExec = False
            End If
        Else
            '�˷ѵ���ĿҲ����ʾ
            If Nvl(rsTmp("��¼״̬"), 0) > 1 Then blnShowExec = False
        End If
        
        'û�ж�Ӧ��ɫ����Ĳɼ���д��
        If IsNull(rsTmp("�Թ���ɫ")) = False And blnShowExec = True Then
            
            If strOldAdvice <> rsTmp("���ID") Then
                
                Set Record = Me.rptAlist(Me.TabCtr.Selected.Index).Records.Add
                For intLoop = 0 To Me.rptAlist(Me.TabCtr.Selected.Index).Columns.Count + 1
                    Record.AddItem ""
                Next
                
                Record(mAcol.ID).Value = Nvl(rsTmp("ҽ��ID")) & "," & Nvl(rsTmp("���ID"))
                Record(mAcol.ѡ��).HasCheckbox = True
                
                '�����뵥��������ѡ��Ҫ�󶨵�����
                If InStr(strAdvice, ";" & Nvl(rsTmp("ҽ������")) & Nvl(rsTmp("�걾")) & Nvl(rsTmp("Ӥ��"))) <= 0 Then
                    If blOnlyWhere = True Then
                        Select Case Mid(txtGoto, 1, 1)
                            Case "+", "*"                           'סԺ��,�����
    '                            If Nvl(rsTmp("��ʶ��")) = Mid(txtGoto, 2) Then
                                    Record(mAcol.ѡ��).Checked = True
    '                            End If
                            Case "."                                '�Һŵ���
    '                            If Nvl(rsTmp("�Һŵ�")) = Mid(txtGoto, 2) Then
                                    Record(mAcol.ѡ��).Checked = True
    '                            End If
                            Case "/"                                '�շѵ��ݺ�
                                If Nvl(rsTmp("NO")) = zlCommFun.GetFullNO(Mid(txtGoto, 2)) Then
                                    Record(mAcol.ѡ��).Checked = True
                                End If
                            Case Else
                                Record(mAcol.ѡ��).Checked = True
                        End Select
                        strAdvice = strAdvice & ";" & Nvl(rsTmp("ҽ������")) & Nvl(rsTmp("�걾")) & Nvl(rsTmp("Ӥ��"))
                    Else
                        Record(mAcol.ѡ��).Checked = True
                        strAdvice = strAdvice & ";" & Nvl(rsTmp("ҽ������")) & Nvl(rsTmp("�걾")) & Nvl(rsTmp("Ӥ��"))
                    End If
                End If
                If Nvl(rsTmp("Ӥ��"), 0) > 0 Then
                
                    If rsTmp("������Դ") = 2 Then
                        gstrSql = "select Ӥ������,Ӥ���Ա� from ������������¼ where ����ID = [1] and ��ҳID = [2] and ��� = [3] "
                        Set rsBaby = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Nvl(rsTmp("����ID"), 0)), CLng(Nvl(rsTmp("��ҳID"), 0)), CInt(rsTmp("Ӥ��")))
                        If rsBaby.EOF = False Then
                            Record(mAcol.Ӥ������).Value = "Ӥ��(" & rsBaby("Ӥ������") & ")"
                            Record(mAcol.Ӥ���Ա�).Value = "Ӥ��(" & rsBaby("Ӥ���Ա�") & ")"
                        Else
                            Record(mAcol.Ӥ������).Value = "Ӥ��(" & rsTmp("Ӥ��") & ")"
                            Record(mAcol.Ӥ���Ա�).Value = "δ֪"
                        End If
                    Else
                        Record(mAcol.Ӥ������).Value = "Ӥ��(" & rsTmp("Ӥ��") & ")"
                        Record(mAcol.Ӥ���Ա�).Value = "δ֪"
                    End If
                End If
                
'                Record(mAcol.ѡ��).Checked = True
                'û��������ʱҲ����Ϊ���̡�
                If IsNull(rsTmp("��¼״̬")) = True Then
                    Record(mAcol.��¼״̬).Value = "��"
                Else
                    Record(mAcol.��¼״̬).Value = IIf(rsTmp("��¼״̬") = 1, "��", "��")
                End If
                Record(mAcol.�ز�).Value = IIf(rsTmp("�ز�") = 1, "��", "")
                Record(mAcol.����).Icon = IIf(rsTmp("����") = "����", 2, 10)
                Record(mAcol.ͼ��).BackColor = Val(Nvl(rsTmp("�Թ���ɫ")))
                Record(mAcol.�ɼ���ʽ).Value = Nvl(rsTmp("�ɼ���ʽ"))
                Record(mAcol.ҽ������).Value = Nvl(rsTmp("ҽ������"))
                Record(mAcol.����).Value = Nvl(rsTmp("��������"))
                Record(mAcol.ִ�п���).Value = Nvl(rsTmp("ִ�п���"))
                Record(mAcol.����ҽ��).Value = Nvl(rsTmp("����ҽ��"))
                Record(mAcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                Record(mAcol.������).Value = Nvl(rsTmp("������"))
                Record(mAcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                Record(mAcol.�Թ���ɫ).Value = Val(Nvl(rsTmp("�Թ���ɫ")))
                Record(mAcol.�Թܱ���).Value = Nvl(rsTmp("�Թܱ���"))
                Record(mAcol.�걾).Value = Nvl(rsTmp("�걾")) & IIf(Nvl(rsTmp("Ӥ��")) = 0, "", "(Ӥ��" & rsTmp("Ӥ��") & " )")
                Record(mAcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                Record(mAcol.������).Value = Nvl(rsTmp("������"))
                Record(mAcol.��Ѫ��).Value = Nvl(rsTmp("��Ѫ��"))
                Record(mAcol.�Թ�����).Value = Nvl(rsTmp("�Թ�����"))
                Record(mAcol.����).Value = Nvl(rsTmp("����"))
                Record(mAcol.������Դ).Value = Nvl(rsTmp("������Դ"))
                Record(mAcol.Ӥ��).Value = Nvl(rsTmp("Ӥ��"))
                Record(mAcol.����).Value = Nvl(rsTmp("����"))
                Record(mAcol.���ID).Value = Nvl(rsTmp("���ID"))
                Record(mAcol.ִ��״̬).Value = Nvl(rsTmp("ִ��״̬"))
                Record(mAcol.NO).Value = Nvl(rsTmp("NO"))
                Record(mAcol.���ʱ��).Value = Nvl(rsTmp("���ʱ��"))
                Record(mAcol.�������).Value = Nvl(rsTmp("�������"))
                Record(mAcol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                Record(mAcol.������).Value = Nvl(rsTmp("������"))
                Record(mAcol.�ͳ�ʱ��).Value = Nvl(rsTmp("�걾�ͳ�ʱ��"))
                Record(mAcol.�����ӡ).Value = IIf(Val(Nvl(rsTmp("�����ӡ"))) = 0, "δ��ӡ", "�Ѵ�ӡ")
                Record(mAcol.��������).Value = Nvl(rsTmp("ִ��˵��"))
                Record(mAcol.�ɼ�����ID).Value = Nvl(rsTmp("ִ�п���ID"))
                Record(mAcol.�ɼ�ִ�п���).Value = Nvl(rsTmp("�ɼ�ִ�п���"))
                Record(mAcol.������ĿID).Value = Val(Nvl(rsTmp("������ĿID")))
                Record(mAcol.����ִ�п���ID).Value = Val(Nvl(rsTmp("����ִ�п���ID")))
                Record(mAcol.�Ʒ�״̬).Value = Val(Nvl(rsTmp("�Ʒ�״̬")))
                Record(mAcol.��¼����).Value = Val(Nvl(rsTmp("��¼����")))
                Record(mAcol.���).Value = Nvl(rsTmp("���"))
                Record(mAcol.�������ڿ���).Value = Nvl(rsTmp("�������ڿ���"))
                If Nvl(rsTmp("���")) = "��Ѫ" Then blnFL = True    '��������Ѫҽ��ʱ,��Ҫ������ʾ,������ʾ��ʦ������Ѫҽ��,��Ҫ��������ҽ������
                For intLoop = 0 To Me.rptAlist(Me.TabCtr.Selected.Index).Columns.Count + 1
                    Record(intLoop).ForeColor = Val(Nvl(rsTmp("�Թ���ɫ")))
                Next
                Record(mAcol.����).Value = IIf(Trim(Nvl(rsTmp("����"))) = "", Nvl(rsTmp("ҽ������")), Nvl(rsTmp("����")))
                If blOnlyWhere = True Then
                    If Record(mAcol.����).Value <> "" And intShowButtom <> 2 Then
                        intShowButtom = 1
                    End If
                    If Record(mAcol.������).Value <> "" Then
                        If Nvl(rsTmp("�걾�ͳ�ʱ��")) = "" Then
                            intShowButtom = 2
                        Else
                            intShowButtom = 3
                        End If
                    End If
                    
                    If Record(mAcol.ִ��״̬).Value = 1 Then
                        intShowButtom = 4
                    End If
                End If
                If strReceivePeople = "" Then
                    If Record(mAcol.������).Value <> "" Then
                        strReceivePeople = Record(mAcol.������).Value
                    End If
                End If
                If Record(mAcol.�ز�).Value = 1 Then
                    For intLoop = 0 To Me.rptAlist(Me.TabCtr.Selected.Index).Columns.Count + 1
                        Record(intLoop).Bold = True
                    Next
                End If
            Else
                strҽ������ = Nvl(rsTmp("ҽ������"))
                If InStr(";" & Record(mAcol.ҽ������).Value & ";", ";" & strҽ������ & ";") <= 0 Then
                    Record(mAcol.ҽ������).Value = Record(mAcol.ҽ������).Value & ";" & Nvl(rsTmp("ҽ������"))
                End If
                
                Record(mAcol.�ϲ�ҽ��).Value = Record(mAcol.�ϲ�ҽ��).Value & "," & _
                                               Nvl(rsTmp("ҽ��ID")) & "," & Nvl(rsTmp("���ID"))
                
                strҽ������ = IIf(Trim(Nvl(rsTmp("����"))) = "", Nvl(rsTmp("ҽ������")), Nvl(rsTmp("����")))
                If InStr(";" & Record(mAcol.����).Value & ";", ";" & strҽ������ & ";") <= 0 Then
                    Record(mAcol.����).Value = Record(mAcol.����).Value & ";" & strҽ������
                End If
                Record(mAcol.������Ŀ���).Value = Record(mAcol.������Ŀ���).Value & ";" & Val(Nvl(rsTmp("������ĿID")))
                
            End If
            strOldAdvice = rsTmp("���ID")
            If InStr(1, strCuvetteNumber & ",", "," & Nvl(rsTmp("�Թܱ���")) & ",") <= 0 Then
                strCuvetteNumber = strCuvetteNumber & "," & Nvl(rsTmp("�Թܱ���"))
            End If
            
            If InStr(1, mstrBarCodes & ",", "," & Nvl(rsTmp("��������")) & ",") <= 0 Then
                mstrBarCodes = mstrBarCodes & "," & Nvl(rsTmp("��������"))
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    If mstrBarCodes <> "" Then
        mstrBarCodes = Mid$(mstrBarCodes, 2)
    End If
    
    If Me.Visible = True Then
        Me.rptAlist(Me.TabCtr.Selected.Index).Populate
    End If
    
    '�м�¼ʱ��ʾд��ɹ�
    If rptAlist(Me.TabCtr.Selected.Index).Records.Count > 0 Then
        RefreshAdviceData = True
    End If
    
    '��ʹ�ò��˲���ʱ��д������Ϣ
    Me.txt����.Text = ""
'    Me.cboAge.Text = ""
    Me.txt����1.Text = ""
    cbo��������.Text = ""
    cboҽ��.Text = ""
    txtҽ������.Text = ""
    txtҽ������.Tag = ""
    
    If blOnlyWhere = True Then
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            txt���� = Nvl(rsTmp("��������"))
            txt����.Tag = Nvl(rsTmp("��������"))
            On Error Resume Next
            cbo�Ա� = Nvl(rsTmp("�Ա�"))
            cbo�Ա�.Tag = ""
            strAge = Nvl(rsTmp("����"))
            
            strAge = Replace(strAge, "Сʱ", "ʱ")
            strAge = Replace(strAge, "����", "��")
        
            If Trim(Replace(Replace(Replace(Replace(Replace(strAge, "��", ""), "��", ""), "��", ""), "ʱ", ""), "��", "")) <> "" Then
                If InStr(strAge, "����") > 0 Or InStr(strAge, "Ӥ��") > 0 Then
                    Me.txt����.Text = ""
                    Me.cboAge.Text = Trim(strAge)
                Else
                    strAge = Replace(Replace(Replace(Replace(Replace(strAge, "��", "��;"), "��", "��;"), "��", "��;"), "ʱ", "ʱ;"), "��", "��;")
                    aAge = Split(strAge, ";")
                    If UBound(aAge) = 1 Then
                        Me.txt����.Text = Val(aAge(0))
                        Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "��", "����"), "ʱ", "Сʱ")
                    Else
                        Me.txt����.Text = Val(aAge(0))
                        Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "��", "����"), "ʱ", "Сʱ")
                        Me.txt����1.Text = Val(aAge(1)) & Replace(Replace(Right(aAge(1), 1), "��", "����"), "ʱ", "Сʱ")
                    End If
                End If
            Else
                Me.txt����.Text = ""
                Me.cboAge.ListIndex = 0
            End If
'            txt���� = Val(Nvl(rsTmp("����")))
'            If IsNumeric(Nvl(rsTmp("����"))) = False And Len(Nvl(rsTmp("����"))) > 0 Then
'                Me.cboAge = Mid(Nvl(rsTmp("����")), Len(rsTmp("����")))
'            End If

            On Error GoTo 0
            txtBed = Nvl(rsTmp("����"))
            txtID = Nvl(rsTmp("��ʶ��"))
            If Nvl(rsTmp("������Դ")) = 2 Then
                lblCap(0).Caption = "ס  Ժ ��"
            Else
                If Nvl(rsTmp("��ʶ��")) = "" Then
                    lblCap(0).Caption = "��  �� ��"
                    txtID = Nvl(rsTmp("NO"))
                Else
                    lblCap(0).Caption = "��  �� ��"
                End If
            End If
            txtPatientDept = Nvl(rsTmp("�������ڿ���"))
'            cbo��������.ListIndex = -1
'            cboҽ��.ListIndex = -1
'            txtҽ������.Text = ""
'            txtҽ������.Tag = ""
            
            cbo��������.Text = Nvl(rsTmp("�������"))
            cboҽ��.Text = Nvl(rsTmp("����ҽ��"))
            txtҽ������.Text = Nvl(rsTmp("ҽ������"))
            
            If Nvl(rsTmp("����")) <> "" Then
                lblCap(6).Visible = True
            Else
                lblCap(6).Visible = False
            End If
            If Val(Nvl(rsTmp("����ģʽ"))) = 1 Then
                lblCap(10).Visible = True
            Else
                lblCap(10).Visible = False
            End If
            mlngKey = Nvl(rsTmp("����ID"))
            
            Select Case intState
                Case 1
                    Me.cmdBindBarCode.Enabled = True
                    Me.cmdNewBarcode.Enabled = True
                    Me.cmdComplete.Enabled = True
                    Me.cmdBarcodePrint.Enabled = True
                    Me.cmdBakBillPrint.Enabled = True
                    Me.cmdBindBarCode.Caption = "�����(&B)"
                    Me.cmdNewBarcode.Caption = "ȡ������(&N)"
                    Me.cmdComplete.Caption = "��ɲɼ�(&P)"
                Case 2
                    Me.cmdBindBarCode.Enabled = False
                    Me.cmdNewBarcode.Enabled = True
                    If strReceivePeople = "" Then
                        Me.cmdComplete.Enabled = True
                    Else
                        Me.cmdComplete.Enabled = False
                    End If
                    Me.cmdBarcodePrint.Enabled = True
                    Me.cmdBakBillPrint.Enabled = True
                    Me.cmdBindBarCode.Caption = "�����(&B)"
                    Me.cmdNewBarcode.Caption = "�ͼ�걾(&C)"
                    Me.cmdComplete.Caption = "ȡ�����(&P)"
                Case 3
                    Me.cmdBindBarCode.Enabled = False
                    Me.cmdNewBarcode.Enabled = True
                    Me.cmdComplete.Enabled = False
                    Me.cmdBarcodePrint.Enabled = True
                    Me.cmdBakBillPrint.Enabled = True
                    Me.cmdBindBarCode.Caption = "�����(&B)"
                    Me.cmdNewBarcode.Caption = "ȡ���ͼ�(&C)"
                    Me.cmdComplete.Caption = "ȡ�����(&P)"
                Case 4
                    Me.cmdBindBarCode.Enabled = False
                    Me.cmdNewBarcode.Enabled = False
                    Me.cmdComplete.Enabled = False
                    Me.cmdBarcodePrint.Enabled = False
                    Me.cmdBakBillPrint.Enabled = False
                    Me.cmdBindBarCode.Caption = "�����(&B)"
                    Me.cmdNewBarcode.Caption = "ȡ������(&N)"
                    Me.cmdComplete.Caption = "ȡ�����(&P)"
                Case Else
                    Me.cmdBindBarCode.Enabled = True
                    Me.cmdNewBarcode.Enabled = True
                    Me.cmdComplete.Enabled = False
                    Me.cmdBarcodePrint.Enabled = False
                    Me.cmdBakBillPrint.Enabled = False
                    Me.cmdBindBarCode.Caption = "������(&B)"
                    Me.cmdNewBarcode.Caption = "��������(&N)"
                    Me.cmdComplete.Caption = "��ɲɼ�(&P)"
                            
            End Select
        End If
    Else
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            txt���� = Nvl(rsTmp("��������"))
            txt����.Tag = Nvl(rsTmp("��������"))
            On Error Resume Next
            cbo�Ա� = Nvl(rsTmp("�Ա�"))
            cbo�Ա�.Tag = ""
            
            strAge = Nvl(rsTmp("����"))
            
            
            strAge = Replace(strAge, "Сʱ", "ʱ")
            strAge = Replace(strAge, "����", "��")
            
            
            If Trim(Replace(Replace(Replace(Replace(Replace(strAge, "��", ""), "��", ""), "��", ""), "ʱ", ""), "��", "")) <> "" Then
                If InStr(strAge, "����") > 0 Or InStr(strAge, "Ӥ��") > 0 Then
                    Me.txt����.Text = ""
                    Me.cboAge.Text = Trim(strAge)
                Else
                    strAge = Replace(Replace(Replace(Replace(Replace(strAge, "��", "��;"), "��", "��;"), "��", "��;"), "ʱ", "ʱ;"), "��", "��;")
                    aAge = Split(strAge, ";")
                    If UBound(aAge) = 1 Then
                        Me.txt����.Text = Val(aAge(0))
                        Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "��", "����"), "ʱ", "Сʱ")
                    Else
                        Me.txt����.Text = Val(aAge(0))
                        Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "��", "����"), "ʱ", "Сʱ")
                        Me.txt����1.Text = Val(aAge(1)) & Replace(Replace(Right(aAge(1), 1), "��", "����"), "ʱ", "Сʱ")
                    End If
                End If
            Else
                Me.txt����.Text = ""
                Me.cboAge.ListIndex = 0
            End If
'            txt���� = Val(Nvl(rsTmp("����")))
'            If IsNumeric(Nvl(rsTmp("����"))) = False And Len(Nvl(rsTmp("����"))) > 0 Then
'                Me.cboAge = Mid(Nvl(rsTmp("����")), Len(rsTmp("����")))
'            End If
            On Error GoTo 0
            txtBed = Nvl(rsTmp("����"))
            txtID = Nvl(rsTmp("��ʶ��"))
            If Nvl(rsTmp("������Դ")) = 2 Then
                lblCap(0).Caption = "ס  Ժ ��"
            Else
                If Nvl(rsTmp("��ʶ��")) = "" Then
                    lblCap(0).Caption = "��  �� ��"
                    txtID = Nvl(rsTmp("NO"))
                Else
                    lblCap(0).Caption = "��  �� ��"
                End If
            End If
            
            
            txtPatientDept = Nvl(rsTmp("�������ڿ���"))
'            cbo��������.ListIndex = -1
'            cboҽ��.ListIndex = -1
'            txtҽ������.Text = ""
'            txtҽ������.Tag = ""
            
            cbo��������.Text = Nvl(rsTmp("�������"))
            cboҽ��.Text = Nvl(rsTmp("����ҽ��"))
            txtҽ������.Text = Nvl(rsTmp("ҽ������"))
            
            If Nvl(rsTmp("����")) <> "" Then
                lblCap(6).Visible = True
            Else
                lblCap(6).Visible = False
            End If
            If Val(Nvl(rsTmp("����ģʽ"))) = 1 Then
                lblCap(10).Visible = True
            Else
                lblCap(10).Visible = False
            End If
            mlngKey = Nvl(rsTmp("����ID"))
        End If
        '���ö���
        Select Case Me.TabCtr.Selected.Index
            Case 0
            Me.cmdBindBarCode.Enabled = True
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = False
            Me.cmdBarcodePrint.Enabled = False
            Me.cmdBakBillPrint.Enabled = False
            Me.cmdBindBarCode.Caption = "������(&B)"
            Me.cmdNewBarcode.Caption = "��������(&N)"
            Me.cmdComplete.Caption = "��ɲɼ�(&P)"
            Me.cmdBakBillPrint.Caption = "��ִ����ӡ"
        Case 1
            Me.cmdBindBarCode.Enabled = True
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = True
            Me.cmdBarcodePrint.Enabled = True
            Me.cmdBakBillPrint.Enabled = True
            Me.cmdBindBarCode.Caption = "�����(&B)"
            Me.cmdNewBarcode.Caption = "ȡ������(&N)"
            Me.cmdComplete.Caption = "��ɲɼ�(&P)"
            Me.cmdBakBillPrint.Caption = "��ִ����ӡ"
        Case 2
            Me.cmdBindBarCode.Enabled = False
            Me.cmdNewBarcode.Enabled = True
            If strReceivePeople = "" Then
                Me.cmdComplete.Enabled = True
            Else
                Me.cmdComplete.Enabled = False
            End If
            Me.cmdBarcodePrint.Enabled = True
            Me.cmdBakBillPrint.Enabled = True
            Me.cmdBindBarCode.Caption = "�����(&B)"
            Me.cmdNewBarcode.Caption = "�ͼ�걾(&C)"
            Me.cmdComplete.Caption = "ȡ�����(&P)"
            Me.cmdBakBillPrint.Caption = "��ִ����ӡ"
        Case 3
            Me.cmdBindBarCode.Enabled = False
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = False
            Me.cmdBarcodePrint.Enabled = True
            Me.cmdBakBillPrint.Enabled = True
            Me.cmdBindBarCode.Caption = "�����(&B)"
            Me.cmdNewBarcode.Caption = "ȡ���ͼ�(&C)"
            Me.cmdComplete.Caption = "ȡ�����(&P)"
            Me.cmdBakBillPrint.Caption = "��ִ����ӡ"
        Case 4
            Me.cmdBindBarCode.Enabled = False
            Me.cmdNewBarcode.Enabled = False
            Me.cmdComplete.Enabled = False
            Me.cmdBarcodePrint.Enabled = False
            Me.cmdBakBillPrint.Enabled = False
            Me.cmdBindBarCode.Caption = "�����(&B)"
            Me.cmdNewBarcode.Caption = "ȡ������(&N)"
            Me.cmdComplete.Caption = "ȡ�����(&P)"
            Me.cmdBakBillPrint.Caption = "��ִ����ӡ"
        Case 5
            Me.cmdBindBarCode.Enabled = False
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = True
            Me.cmdBarcodePrint.Enabled = False
            Me.cmdBakBillPrint.Enabled = True
            Me.cmdBindBarCode.Caption = "�����(&B)"
            Me.cmdNewBarcode.Caption = "�ò�����(&N)"
            Me.cmdComplete.Caption = "ȡ�����(&P)"
            Me.cmdBakBillPrint.Caption = "�ز�����(&R)"
        End Select
    End If
    
    
    If strCuvetteNumber <> "" Then
        With Me.rptCuvette
            strSQL = "select ����,����,��Ӽ�,��Ѫ��,���,��ɫ from ��Ѫ������ where ���� in " & _
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
            
            If InStr(1, Mid(strCuvetteNumber, 2), ",") <= 0 And .Records.Count > 0 Then
                .Records(0).Item(mCuvette.ѡ��).Checked = True
            End If
        End With
    End If
    If strTmp <> "" Then
    
        Me.stbThis.Panels(2).Text = "��ǰ��Χ<" & IIf(varFilter(mFilter.���ͻ����ʱ��) = 0, "����ʱ�� ", "����ʱ�� ") & _
                                Format(strDateBegin, "yyyy-mm-dd") & "---" & Format(strDateEnd, "yyyy-mm-dd") & "> �¹���:" & _
                                Me.rptPlist.Rows.Count & "������."
    Else
        Me.stbThis.Panels(2).Text = "��ǰ��Χ<" & "����ʱ�� " & _
                                Format(strDateBegin, "yyyy-mm-dd") & "---" & Format(strDateEnd, "yyyy-mm-dd") & "> �¹���:" & _
                                Me.rptPlist.Rows.Count & "������."
    End If
    Me.rptCuvette.Populate
    
    Me.rptAlist(Me.TabCtr.Selected.Index).GroupsOrder.DeleteAll
    If blnFL = True Then
        With Me.rptAlist(Me.TabCtr.Selected.Index)
            Call .GroupsOrder.Add(.Columns.Column(mAcol.���))
            .Populate
        End With
    End If
    If mlngKey <> 0 Then
        Call ReadPatPricture(mlngKey, imgLoad)
        If imgLoad.Picture = 0 Then
            imgPatient.Picture = imgDefual.Picture
        Else
            imgPatient.Picture = imgLoad.Picture
        End If
    End If
'    SelectCuvette
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetBooldPatientDataSql() As String
    Dim strBloodSQL
    Dim strTmp As String
    Dim strBloodType As String
    Dim varFilter As Variant
    Dim NowDate As Date

    On Error GoTo errH

    '��ע����ж�ȡ��������
    strTmp = zlDatabase.GetPara("�ɼ�����վ����", 100, 1211, "")
    strBloodType = zlDatabase.GetPara(273, 100)
    NowDate = zlDatabase.Currentdate
    '�ӹ��˴������������ʱ����
    If Me.rptPlist.Tag <> "" Then
        varFilter = Split(Me.rptPlist.Tag, ";")
    Else
        If strTmp <> "" Then
            varFilter = Split(strTmp, ";")
        End If
    End If

    strBloodSQL = strBloodSQL & "Select    distinct a.����id,decode(a.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���') as ������Դ, " & vbCrLf & _
                " c.���� as ��������,e.���� as ���˿���,c.�Ա�,c.����,c.���￨��,b.��������, " & vbCrLf & _
                " decode(b.ִ��״̬,1,'�����',2,'����',3,'ִ����', " & vbCrLf & _
                "         Decode(b.��������, Null, 'δ��', Decode(b.������, Null, '�Ѱ�', decode(b.�걾�ͳ�ʱ��,null,'�Ѳ���','���ͼ�'))) )as ״̬, " & vbCrLf & _
                " decode(A.������Դ, 1, C.�����, 2, C.סԺ��,4,c.�����) As ��ʶ��, " & vbCrLf & _
                " decode(c.��ǰ����,null,decode(l.��Ժ����,null,l.��Ժ����,l.��Ժ����),c.��ǰ����) as ��ǰ���� , " & vbCrLf & _
                " decode(a.������־,1,'����',decode(g.����,1,'����')) as ���� ,decode(a.������Դ , 2,a.��ҳID,0) ��ҳID, " & vbCrLf & _
                " decode(b.ִ��״̬,0,'',2,'����') as ����,b.ִ��״̬,nvl(b.�زɱ걾,0) as �زɱ걾,b.����ʱ��,nvl(s.·��״̬,0) as �ٴ�·������,a.ҽ������ " & vbCrLf & _
                " From ����ҽ����¼ A, ����ҽ������ B, ������Ϣ C, ���ű� E, ������ĿĿ¼ F,���˹Һż�¼ G,����ҽ����¼ H, " & vbCrLf & _
                "      ������ĿĿ¼ K ,������ҳ L,����ҽ������ M,����걾��¼ J,������ҳ S " & vbCrLf & _
                " Where A.ID = H.���ID And H.id = B.ҽ��id And A.����id = C.����id And A.���˿���id = E.ID And A.������Ŀid+0 = f.ID  " & vbCrLf & _
                "      And h.������ĿID = k.id and a.id = j.ҽ��id(+)  " & vbCrLf & _
                " And A.�Һŵ� = G.No(+) and a.����id = g.����id(+)  and a.����id = g.����id(+)  and (g.����ID is null or (g.��¼״̬ =1 and g.��¼���� =1) ) And f.��� = 'K'  and Decode(f.���, 'K', '9') = k.�������� and a.����id = l.����ID(+) and a.��ҳID = l.��ҳID(+) and b.ִ�в���id + 0 = [1] " & vbCrLf & _
                " And A.ID = M.ҽ��ID And k.�Թܱ��� is not null and a.����ID = S.����ID(+) and a.��ҳID = s.��ҳID(+) " & IIf(Me.rptPlist.Tag = "", "and a.��ʼִ��ʱ�� < to_date('" & Format(NowDate, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')", "")


    If Me.rptPlist.Tag <> "" Then
        strBloodSQL = strBloodSQL & " And A.������Դ in (" & IIf(Val(varFilter(mFilter.����)) = 1, "1,3", "0") & "," & _
                      Val(varFilter(mFilter.סԺ)) & "," & Val(varFilter(mFilter.���)) & ") "
        If varFilter(mFilter.��ʶ��) <> "" Then
            strBloodSQL = strBloodSQL & " And decode(a.������Դ,2,c.סԺ��,c.�����) = [2] "
        End If

        If varFilter(mFilter.���￨) <> "" Then
            strBloodSQL = strBloodSQL & " And c.���￨�� = [3] "

        End If

        If varFilter(mFilter.����) <> "" Then
            strBloodSQL = strBloodSQL & " And C.���� like [4] "

        End If

        If varFilter(mFilter.���ݺ�) <> "" Then
            strBloodSQL = strBloodSQL & " and B.NO = [5]"

        End If

        If UBound(varFilter) >= mFilter.�걾 Then
            If Trim(varFilter(mFilter.�걾)) <> "" Then
                strBloodSQL = strBloodSQL & " And decode(f.���,'K',1,instr([6],','||H.�걾��λ||',')) > 0 "

            End If
        End If

        If UBound(varFilter) >= mFilter.�ɼ���ʽ Then
            If Trim(varFilter(mFilter.�ɼ���ʽ)) <> "" Then
                strBloodSQL = strBloodSQL & " And instr([7],','|| decode(f.���,'K',K.id,f.ID) ||',') > 0 "

            End If
        End If

        If varFilter(mFilter.���ͻ����ʱ��) = 0 Then
            strBloodSQL = strBloodSQL & " and b.����ʱ�� Between [8] and [9]"

        Else
            strBloodSQL = strBloodSQL & " and b.����ʱ�� Between [8] and [9]"
        End If

        If UBound(varFilter) >= mFilter.�������� Then
            If Trim(varFilter(mFilter.��������)) <> "" Then
                strBloodSQL = strBloodSQL & " And instr([10],','|| decode(f.���,'K','" & strBloodType & "',k.��������) ||',') > 0 "

            End If
        End If
    Else
        If strTmp <> "" Then
            strBloodSQL = strBloodSQL & " And instr('" & IIf(Val(varFilter(mFilter.����)) = 1, "1,3", "0") & "," & _
                          Val(varFilter(mFilter.סԺ)) & "," & Val(varFilter(mFilter.���)) & "',A.������Դ)>0 "

            If UBound(varFilter) >= mFilter.�걾 Then
                If Trim(varFilter(mFilter.�걾)) <> "" Then
                    strBloodSQL = strBloodSQL & " And decode(f.���,'K',1,instr([6],','||H.�걾��λ||',')) > 0 "

                End If
            End If

            If UBound(varFilter) >= mFilter.�ɼ���ʽ Then
                If Trim(varFilter(mFilter.�ɼ���ʽ)) <> "" Then
                    strBloodSQL = strBloodSQL & " And instr([7],','|| decode(f.���,'K',K.id,f.ID) ||',') > 0 "

                End If
            End If

            If varFilter(mFilter.���ͻ����ʱ��) = 0 Then
                strBloodSQL = strBloodSQL & " and b.����ʱ�� Between [8] and [9]"
            Else
                strBloodSQL = strBloodSQL & " and b.����ʱ�� Between [8] and [9]"

            End If

            If UBound(varFilter) >= mFilter.�������� Then
                If Trim(varFilter(mFilter.��������)) <> "" Then
                    strBloodSQL = strBloodSQL & " And instr([10],','|| decode(f.���,'K','" & strBloodType & "',k.��������) ||',') > 0 "

                End If
            End If
        Else
            strBloodSQL = strBloodSQL & " and m.����ʱ�� Between [8] and [9]"
        End If
    End If
    GetBooldPatientDataSql = strBloodSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function GetBloodAdviceSql(intState As Integer, intPatientType As Integer, Optional blOnlyWhere As Boolean = False, Optional intPatientPage As Integer) As String
    Dim blnDateMoved As Boolean                     '�Ƿ�ת��
    Dim strSQL As String                            'SQL���
    Dim strTmp As String                            '��ʱ�ִ�����
    Dim varFilter As Variant                        '�����ִ�
    Dim strDateBegin As String                      '��ʼ����ʱ��
    Dim strDateEnd As String                        '��������ʱ��
    Dim rsTmp As New ADODB.Recordset                '���ݼ�
    Dim intLoop As Integer                          'ѭ������
    Dim Record As ReportRecord                      '�б����ݼ�
    Dim strOldAdvice As String                      '��¼�ϴ�ҽ��
    Dim strCuvetteNumber As String                  '���ڼ�¼�Թܱ���
    Dim NowDate As Date                             '��ǰʱ��
    Dim blnFL As Boolean                            '�Ƿ������ʾ,������Ѫҽ��ʱ,������ʾ
    Dim strSQL1 As String


    blnDateMoved = MovedByDate(Date)    '��ʱ�俴�Ƿ������ת��

    '��ע����ж�ȡ��������
    strTmp = zlDatabase.GetPara("�ɼ�����վ����", 100, 1211, "")

    NowDate = zlDatabase.Currentdate

    '�ӹ��˴������������ʱ����
    If Me.rptPlist.Tag <> "" Then
        varFilter = Split(Me.rptPlist.Tag, ";")
    Else
        If strTmp <> "" Then
            varFilter = Split(strTmp, ";")
        End If
    End If

    On Error GoTo errH
    strSQL = " select /*+ rule */ distinct a.���,a.ҽ��id,a.���id,a.�Թ���ɫ,a.�ɼ���ʽ,a.ҽ������,a.��������,a.����ʱ��,a.ִ�п���,a.����ҽ��,a.����ʱ��,a.������,a.����ʱ��,a.�Թܱ���," & vbCrLf & _
           "        a.�걾,a.��������,a.�Ա�,a.����,a.����,a.��ʶ��,Decode(A.������Դ, 2, nvl(A.�������ڿ���,A.�������), A.�������) As �������ڿ���,a.����,a.����ID,a.������,a.��Ѫ��,a.�Թ�����,a.����,a.������Դ,a.Ӥ��,a.����,a.ִ��״̬,a.NO," & vbCrLf & _
           "        a.���ʱ��,a.�������,a.����ʱ��,a.�����Ŀ,b.���ʷ���,b.��¼״̬,a.������,b.��¼����,a.�ɼ���ĿID,a.���￨��,a.�Һŵ�,a.�����ӡ,a.ִ��˵��,a.�ز�, " & vbCrLf & _
           "        a.�걾�ͳ�ʱ��,a.��ҳID,a.ִ�п���ID,a.�ɼ�ִ�п���,a.������ĿID,a.����ִ�п���ID,b.�����־,a.�Ʒ�״̬ , ����ģʽ" & vbCrLf & _
           " from "

    strSQL = strSQL & "(  Select decode(d.���,'K','��Ѫ','����') ���, B.ID as ҽ��ID, B.���id, G.��ɫ As �Թ���ɫ, decode(d.���,'K',b.ҽ������ ,d.����) As �ɼ���ʽ, decode(d.���,'K',d.����,b.ҽ������) as ҽ������, C.��������,C.����ʱ��, " & vbCrLf & _
           "   H.���� As ִ�п���, B.����ҽ��,B.����ʱ��, C.������, C.����ʱ��, G.���� as �Թܱ���,b.�걾��λ as �걾, " & vbCrLf & vbCrLf & _
           "   I.���� as ��������,I.�Ա�,i.����,i.��ǰ���� as ����,decode(b.������Դ,1,I.�����,2,i.סԺ��,4,i.�����) as ��ʶ��, " & vbCrLf & _
           "   L.���� as �������ڿ���,Decode(C.ִ��״̬,2,'����') as ����,I.����ID,c.������,G.��Ѫ��,G.���� as �Թ�����, " & vbCrLf & _
           "   DECODE(B.������־,1,'����','') as ����,b.������Դ,nvl(b.Ӥ��,0) as Ӥ��,N.���� as ����,decode(d.���, 'K', M.ִ��״̬,C.ִ��״̬) ִ��״̬,C.NO,j.���ʱ��,o.���� as �������,m.����ʱ��, " & vbCrLf & _
           "   E.�����Ŀ,C.������,c.��¼����,decode(d.���,'K',e.id ,d.id)  as �ɼ���ĿID,i.���￨��,b.�Һŵ�,C.�����ӡ,C.ִ��˵��,nvl(c.�زɱ걾,0) as �ز�,c.�걾�ͳ�ʱ��, " & vbCrLf & _
           "   a.��ҳID,Decode(d.���, 'K', b.ִ�п���ID, a.ִ�п���ID) ִ�п���ID,P.���� as �ɼ�ִ�п���,b.������ĿID,Decode(d.���, 'K', a.ִ�п���ID, b.ִ�п���ID) as ����ִ�п���ID,c.�Ʒ�״̬,i.����ģʽ " & vbCrLf & _
           "   From ����ҽ����¼ A, ����ҽ����¼ B, ����ҽ������ C, ������ĿĿ¼ D, ������ĿĿ¼ E, " & vbCrLf & _
           "   ��Ѫ������ G,���ű� H, ������Ϣ I,���ű� L,����ҽ������ M,����걾��¼ J ,���ű� O ,���ű� P, " & vbCrLf & _
           "   (select ������ĿID,���� from ������Ŀ���� where ���� = 9 and ���� = 1 ) N " & vbCrLf & _
           "  Where A.ID = B.���id And B.ID = C.ҽ��id And A.������Ŀid = D.ID And B.������Ŀid = E.ID  " & vbCrLf & _
           "    And (e.��� = 'E' Or e.��� = 'C') And E.�Թܱ��� = G.���� And B.ִ�п���id = H.ID(+) and a.ִ�п���ID = P.id(+)  " & vbCrLf & _
           "    and   d.��� = 'K' And  e.��������= '9' And A.����id = [1] " & IIf(InStr(txtGoto.Text, ".") = 1, "", "And c.����ʱ��+0 Between [3] and [4] ") & IIf(Me.rptPlist.Tag = "", "and a.��ʼִ��ʱ�� < to_date('" & Format(NowDate, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')", "") & vbCrLf & _
           "    and c.ִ�в���id + 0 = [2] And B.����ID = I.����ID and I.��ǰ����ID = L.ID(+) " & vbCrLf & _
           "    and a.id = m.ҽ��id And E.id = N.������ĿID(+) and a.id = j.ҽ��id(+) and b.��������id = o.id  " & vbCrLf & _
           "    ) a , (Select ҽ�����,��¼����,��¼״̬,���ʷ���,�����־ From סԺ���ü�¼ Where  ����ID=[1]) b " & vbCrLf & _
             "where a.ҽ��id = b.ҽ�����(+) and a.��¼���� = mod(b.��¼����(+),10)  "

    '  IIf(Me.rptPlist.Tag = "", "and a.��ʼִ��ʱ�� between  to_date( '" & Mid(NowDate, 1, InStr(NowDate, " ")) & " 00:00:00' ,'yyyy-mm-dd hh24:mi:ss' ) and to_date('" & Mid(NowDate, 1, InStr(NowDate, " ")) & " 23:59:59','yyyy-mm-dd hh24:mi:ss')", "") & vbCrLf & _

       '�������ֲ�ͬ��״̬,
    If intState = 0 Then
        strSQL = strSQL & " And a.�������� is null And a.ִ��״̬ in (0) " & vbCrLf
    ElseIf intState = 1 Then
        strSQL = strSQL & " And a.�������� is not null And a.����ʱ�� is  null And a.ִ��״̬ in (0) " & vbCrLf
    ElseIf intState = 2 Then
        strSQL = strSQL & " And a.�������� is not null and  a.����ʱ�� is not null And a.ִ��״̬ in (0) and a.�걾�ͳ�ʱ�� is null " & vbCrLf
    ElseIf intState = 3 Then
        strSQL = strSQL & " And a.�������� is not null and  a.����ʱ�� is not null And a.ִ��״̬ in (0) and a.�걾�ͳ�ʱ�� is not null  " & vbCrLf
    ElseIf intState = 4 Then
        strSQL = strSQL & " And a.ִ��״̬ in (1,3) " & vbCrLf
    ElseIf intState = 5 Then
        strSQL = strSQL & " And a.ִ��״̬ in (2) " & vbCrLf
    End If

    '����
    If Me.rptPlist.Tag <> "" Or strTmp <> "" Then

        If UBound(varFilter) >= mFilter.�걾 Then
            If Trim(varFilter(mFilter.�걾)) <> "" Then
                strSQL = strSQL & " And decode(a.���,'��Ѫ',1,instr([5],','||a.�걾||',')) > 0 "
            End If
        End If

        If UBound(varFilter) >= mFilter.�ɼ���ʽ Then
            If Trim(varFilter(mFilter.�ɼ���ʽ)) <> "" Then
                strSQL = strSQL & " And instr([6],','||a.�ɼ���ĿID||',') > 0 "
            End If
        End If

        If Me.rptPlist.Tag <> "" Then
            strDateBegin = varFilter(mFilter.��ʼʱ��)
            strDateEnd = varFilter(mFilter.����ʱ��)
        Else
            strDateBegin = NowDate - Val(varFilter(mFilter.���ʱ��))
            strDateEnd = NowDate
        End If
    Else
        strDateBegin = NowDate - 3
        strDateEnd = NowDate
    End If

    If intPatientPage <> 0 Then
        strSQL = strSQL & " and a.��ҳid = [9] "
    End If


    If blnDateMoved Then
        strSQL1 = strSQL
        strSQL1 = Replace(strSQL1, "����ҽ����¼", "H����ҽ����¼")
        strSQL1 = Replace(strSQL1, "����ҽ������", "H����ҽ������")
        strSQL = strSQL & " Union ALL " & strSQL1
    End If

    If blOnlyWhere = True Then
        strSQL = " select /*+ rule */ distinct a.���,a.ҽ��id,a.���id,a.�Թ���ɫ,a.�ɼ���ʽ,a.ҽ������,a.��������,a.����ʱ��,a.ִ�п���,a.����ҽ��,a.����ʱ��,a.������,a.����ʱ��,a.�Թܱ���," & vbCrLf & _
               "        a.�걾,a.��������,a.�Ա�,a.����,a.����,a.��ʶ��,Decode(A.������Դ, 2, nvl(A.�������ڿ���,A.�������), A.�������) As �������ڿ���,a.����,a.����ID,a.������,a.��Ѫ��,a.�Թ�����,a.����,a.������Դ,a.Ӥ��,a.����,a.ִ��״̬,a.NO," & vbCrLf & _
               "        a.���ʱ��,a.�������,a.����ʱ��,a.�����Ŀ,b.���ʷ���,b.��¼״̬,a.������,b.��¼����,a.�ɼ���ĿID,a.���￨��,a.�Һŵ�,a.�����ӡ,a.ִ��˵��,a.�ز�, " & vbCrLf & _
               "        a.�걾�ͳ�ʱ��,a.��ҳID,a.ִ�п���ID,a.�ɼ�ִ�п���,a.������ĿID,a.����ִ�п���ID,b.�����־,a.�Ʒ�״̬ , ����ģʽ" & vbCrLf & _
               " from "

        strSQL = strSQL & "(   Select decode(d.���,'K','��Ѫ','����') ���, B.ID as ҽ��ID, B.���id, G.��ɫ As �Թ���ɫ,decode(d.���,'K',b.ҽ������ ,d.����) As �ɼ���ʽ, decode(d.���,'K',d.����,b.ҽ������) as ҽ������, C.��������,C.����ʱ��, " & vbCrLf & _
               "   H.���� As ִ�п���, B.����ҽ��,B.����ʱ��, C.������, C.����ʱ��, G.���� as �Թܱ���,b.�걾��λ as �걾, " & vbCrLf & vbCrLf & _
               "   I.���� as ��������,I.�Ա�,i.����,i.��ǰ���� as ����,decode(b.������Դ,1,I.�����,2,i.סԺ��,4,i.�����) as ��ʶ��, " & vbCrLf & _
               "   L.���� as �������ڿ���,Decode(C.ִ��״̬,2,'����') as ����,I.����ID,c.������,G.��Ѫ��,G.���� as �Թ�����, " & vbCrLf & _
               "   DECODE(B.������־,1,'����','') as ����,b.������Դ,nvl(b.Ӥ��,0) as Ӥ��,N.���� as ����,decode(d.���, 'K', M.ִ��״̬,C.ִ��״̬) ִ��״̬,C.NO,j.���ʱ��,o.���� as �������,m.����ʱ��, " & vbCrLf & _
               "   E.�����Ŀ,C.������,c.��¼����,decode(d.���,'K',e.id ,d.id) as �ɼ���ĿID,i.���￨��,a.�Һŵ�,c.�����ӡ,C.ִ��˵��,nvl(c.�زɱ걾,0) as �ز�,c.�걾�ͳ�ʱ��, " & vbCrLf & _
               "   A.��ҳID,Decode(d.���, 'K', b.ִ�п���ID, a.ִ�п���ID) ִ�п���ID,P.���� as �ɼ�ִ�п���,b.������ĿID,Decode(d.���, 'K', a.ִ�п���ID, b.ִ�п���ID) as ����ִ�п���ID,c.�Ʒ�״̬ ,i.����ģʽ " & vbCrLf & _
               "   From ����ҽ����¼ A, ����ҽ����¼ B, ����ҽ������ C, ������ĿĿ¼ D, ������ĿĿ¼ E, " & vbCrLf & _
               "   ��Ѫ������ G,���ű� H, ������Ϣ I,���ű� L,����ҽ������ M,����걾��¼ J ,���ű� O ,���ű� P, " & vbCrLf & _
               "   (select ������ĿID,���� from ������Ŀ���� where ���� = 9 and ���� = 1 ) N " & vbCrLf & _
               "  Where A.ID = B.���id And B.ID = C.ҽ��id And A.������Ŀid = D.ID And B.������Ŀid = E.ID " & vbCrLf & _
               "    And (e.��� = 'E' Or e.��� = 'C') And E.�Թܱ��� = G.���� And B.ִ�п���id = H.ID(+) and a.ִ�п���ID = P.id(+) " & vbCrLf & _
               "    and d.��� = 'K'  And  e.�������� = '9' And A.����id = [1] " & IIf(InStr(txtGoto.Text, ".") = 1, "", "And c.����ʱ��+0 Between [3] and [4] ") & IIf(Me.rptPlist.Tag = "", "and a.��ʼִ��ʱ�� < to_date('" & Format(NowDate, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')", "") & vbCrLf & _
               "    and c.ִ�в���id + 0 = [2] And B.����ID = I.����ID and I.��ǰ����ID = L.ID(+) " & vbCrLf & _
               "    and a.id = m.ҽ��id And E.id = N.������ĿID(+) and a.id = j.ҽ��id(+) and b.��������id = o.id  " & vbCrLf & _
               "    ) a , (Select ҽ�����,��¼����,��¼״̬,���ʷ���,�����־ From סԺ���ü�¼ Where    ����ID=[1]) b " & vbCrLf & _
                 "where a.ҽ��id = b.ҽ�����(+) and a.��¼���� = b.��¼����(+) "

        '�������ֲ�ͬ��״̬
        If intState = 0 Then
            strSQL = strSQL & " And a.�������� is null And a.ִ��״̬ in (0,2) " & vbCrLf
        ElseIf intState = 1 Then
            strSQL = strSQL & " And a.�������� is not null And a.����ʱ�� is  null And a.ִ��״̬ in (0,2) " & vbCrLf
        ElseIf intState = 2 Then
            strSQL = strSQL & " And a.�������� is not null and  a.����ʱ�� is not null And a.ִ��״̬ in (0,2) and �걾�ͳ�ʱ�� is null " & vbCrLf
        ElseIf intState = 3 Then
            strSQL = strSQL & " And a.�������� is not null and  a.����ʱ�� is not null And a.ִ��״̬ in (0,2) and �걾�ͳ�ʱ�� is not null  " & vbCrLf
        ElseIf intState = 4 Then
            strSQL = strSQL & " And a.ִ��״̬ in (1,3) " & vbCrLf
        ElseIf intState = 5 Then
            strSQL = strSQL & " And a.ִ��״̬ in (2) " & vbCrLf
        End If

        '���ݺ�
        '        If Mid(Me.txtGoto.Text, 1, 1) = "/" Then
        '            strNO = Mid(Me.txtGoto, 2)
        '            If IsNumeric(strNO) = True Then
        '                strsql = strsql & " And a.NO = [7] "
        '            End If
        '        End If
        '���ݲ������ж��Ƿ񰴲ɼ���������ʾ
        If chkDeptShow.Value <> 1 Then
            strSQL = Replace(strSQL, " and c.ִ�в���id + 0 = [2] ", "")
        End If

        If Mid(Me.txtGoto.Text, 1, 1) = "*" Or Mid(Me.txtGoto.Text, 1, 1) = "." Then
            strSQL = strSQL & " And a.������Դ in (1,3,4) "
        End If

        If Mid(Me.txtGoto.Text, 1, 1) = "+" Then
            strSQL = strSQL & " And a.������Դ in ( 2,4) "
        End If


        '����
        If BlnIsNumber(txtGoto) Then
            strSQL = strSQL & " And (a.�������� = [8] or a.���￨�� = [8]) "
        End If

    End If

    If intPatientType <> 2 Then
        strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
    End If
    GetBloodAdviceSql = strSQL

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub rptPlist_SelectionChanged()
    Dim intPatientType As Integer                   '������Դ
    Dim intPatientPage As Integer                   '��ҳID
    
    If Me.rptPlist.FocusedRow Is Nothing Then Exit Sub
    
    With Me.rptPlist.FocusedRow
        mlngKey = .Record(mPcol.����ID).Value
        intPatientType = IIf(.Record(mPcol.��Դ).Value = "סԺ", 2, 1)
        intPatientPage = .Record(mPcol.��ҳID).Value
    End With
    'ʹ�ò���IDˢ��ҽ��
    RefreshAdviceData mlngKey, Me.TabCtr.Selected.Index, intPatientType, False, intPatientPage
    'ˢ����ʾ��Ϣ
    ShowPatientInfo
    '��λ���㵽��������
'    If Me.Visible = True Then Me.TxtBarCode.SetFocus
    mblFind = False
End Sub

Private Sub ShowPatientInfo()
    Dim strAge As String
    Dim aAge() As String
    
    'û�н�����ʱ�˳�
    If Me.rptPlist.FocusedRow Is Nothing Then Exit Sub
    On Error Resume Next
    With Me.rptPlist.FocusedRow
    
        Call AdjustEditState(True)
        
        txt���� = .Record(mPcol.��������).Value
        txt����.Tag = .Record(mPcol.��������).Value
        cbo�Ա� = .Record(mPcol.�Ա�).Value
        cbo�Ա�.Tag = ""
        strAge = .Record(mPcol.����).Value
        
        strAge = Replace(strAge, "Сʱ", "ʱ")
        strAge = Replace(strAge, "����", "��")
        
        If Trim(Replace(Replace(Replace(Replace(Replace(strAge, "��", ""), "��", ""), "��", ""), "ʱ", ""), "��", "")) <> "" Then
            If InStr(strAge, "����") > 0 Or InStr(strAge, "Ӥ��") > 0 Then
                Me.txt����.Text = ""
                Me.cboAge.Text = Trim(strAge)
            Else
                strAge = Replace(Replace(Replace(Replace(Replace(strAge, "��", "��;"), "��", "��;"), "��", "��;"), "ʱ", "ʱ;"), "��", "��;")
                aAge = Split(strAge, ";")
                If UBound(aAge) = 1 Then
                    Me.txt����.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "��", "����"), "ʱ", "Сʱ")
                Else
                    Me.txt����.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "��", "����"), "ʱ", "Сʱ")
                    Me.txt����1.Text = Val(aAge(1)) & Replace(Replace(Right(aAge(1), 1), "��", "����"), "ʱ", "Сʱ")
                End If
            End If
        Else
            Me.txt����.Text = ""
            Me.cboAge.ListIndex = 0
        End If
'        txt���� = Val(.Record(mPcol.����).Value)
'        If IsNumeric(.Record(mPcol.����).Value) = False And Len(.Record(mPcol.����).Value) > 0 Then
'                Me.cboAge = Mid(Nvl(.Record(mPcol.����).Value), Len(.Record(mPcol.����).Value))
'            End If
        txtBed = .Record(mPcol.����).Value
        txtID = .Record(mPcol.��ʶ��).Value
        txtPatientDept = .Record(mPcol.���˿���).Value
        cbo��������.ListIndex = -1
        cboҽ��.ListIndex = -1
'        Me.txtҽ������.Text = ""
'        Me.txtҽ������.Tag = ""
        
        Call AdjustEditState(False)
        
        If .Record(mPcol.����).Value <> 0 Then
            lblCap(6).Visible = True
        Else
            lblCap(6).Visible = False
        End If
    End With
    
End Sub

Private Sub SelectCuvette()
    '����               ѡ��ѡ�е��Թ�
    
    Dim RecordC As ReportRecord
    Dim RecordA As ReportRecord
    
    For Each RecordC In Me.rptCuvette.Records
        For Each RecordA In Me.rptAlist(Me.TabCtr.Selected.Index).Records
            If RecordA(mAcol.�Թܱ���).Value = RecordC(mCuvette.����).Value Then
                RecordA(mAcol.ѡ��).Checked = RecordC(mCuvette.ѡ��).Checked
            End If
        Next
    Next

    Me.rptAlist(Me.TabCtr.Selected.Index).Populate
End Sub

Private Sub TabCtr_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    '���ö���
    Select Case Item.Index
        Case 0
            Me.cmdBindBarCode.Enabled = True
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = False
            Me.cmdBarcodePrint.Enabled = False
            Me.cmdBakBillPrint.Enabled = False
            Me.cmdBindBarCode.Caption = "������(&B)"
            Me.cmdNewBarcode.Caption = "��������(&N)"
            Me.cmdComplete.Caption = "��ɲɼ�(&P)"
            Me.cmdBakBillPrint.Caption = "��ִ����ӡ"
        Case 1
            Me.cmdBindBarCode.Enabled = True
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = True
            Me.cmdBarcodePrint.Enabled = True
            Me.cmdBakBillPrint.Enabled = True
            Me.cmdBindBarCode.Caption = "�����(&B)"
            Me.cmdNewBarcode.Caption = "ȡ������(&N)"
            Me.cmdComplete.Caption = "��ɲɼ�(&P)"
            Me.cmdBakBillPrint.Caption = "��ִ����ӡ"
        Case 2
            Me.cmdBindBarCode.Enabled = False
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = True
            Me.cmdBarcodePrint.Enabled = True
            Me.cmdBakBillPrint.Enabled = True
            Me.cmdBindBarCode.Caption = "�����(&B)"
            Me.cmdNewBarcode.Caption = "�ͼ�걾(&C)"
            Me.cmdComplete.Caption = "ȡ�����(&P)"
            Me.cmdBakBillPrint.Caption = "��ִ����ӡ"
        Case 3
            Me.cmdBindBarCode.Enabled = False
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = False
            Me.cmdBarcodePrint.Enabled = True
            Me.cmdBakBillPrint.Enabled = True
            Me.cmdBindBarCode.Caption = "�����(&B)"
            Me.cmdNewBarcode.Caption = "ȡ���ͼ�(&C)"
            Me.cmdComplete.Caption = "ȡ�����(&P)"
            Me.cmdBakBillPrint.Caption = "��ִ����ӡ"
        Case 4
            Me.cmdBindBarCode.Enabled = False
            Me.cmdNewBarcode.Enabled = False
            Me.cmdComplete.Enabled = False
            Me.cmdBarcodePrint.Enabled = False
            Me.cmdBakBillPrint.Enabled = False
            Me.cmdBindBarCode.Caption = "�����(&B)"
            Me.cmdNewBarcode.Caption = "ȡ������(&N)"
            Me.cmdComplete.Caption = "ȡ�����(&P)"
            Me.cmdBakBillPrint.Caption = "��ִ����ӡ"
        Case 5
            Me.cmdBindBarCode.Enabled = False
            Me.cmdNewBarcode.Enabled = True
            Me.cmdComplete.Enabled = True
            Me.cmdBarcodePrint.Enabled = False
            Me.cmdBakBillPrint.Enabled = True
            Me.cmdBindBarCode.Caption = "�����(&B)"
            Me.cmdNewBarcode.Caption = "�ò�����(&N)"
            Me.cmdComplete.Caption = "ȡ�����(&P)"
            Me.cmdBakBillPrint.Caption = "�ز�����(&R)"
    End Select
    
    If Me.Visible = True Then
        Call RefreshPatientData
    End If
    
    If InStr(1, mstrPrivs, "�걾�ɼ�") <= 0 Then
        Me.cmdBindBarCode.Enabled = False
        Me.cmdNewBarcode.Enabled = False
        Me.cmdComplete.Enabled = False
        Me.cmdBarcodePrint.Enabled = False
        Me.cmdBakBillPrint.Enabled = False
    End If
    
    With Me.rptPlist
        If .Rows.Count > 0 Then
'            RefreshAdviceData mlngKey, Me.TabCtr.Selected.Index
            If .FocusedRow Is Nothing Then
                .FocusedRow = Me.rptPlist.Rows(0)
                .Populate
            Else
                RefreshAdviceData .FocusedRow.Record(mPcol.����ID).Value, Me.TabCtr.Selected.Index, IIf(.FocusedRow.Record(mPcol.��Դ).Value = "סԺ", 2, 1), False, .FocusedRow.Record(mPcol.��ҳID).Value
            End If
        Else
            Me.rptAlist(Me.TabCtr.Selected.Index).Records.DeleteAll
            Me.rptCuvette.Records.DeleteAll
            Me.rptAlist(Me.TabCtr.Selected.Index).Populate
            Me.rptCuvette.Populate
            txt���� = ""
            txt����.Tag = ""
            cbo�Ա�.ListIndex = -1
            txt���� = ""
            txt����1 = ""
            txtBed = ""
            txtID = ""
            txtPatientDept = ""
            cbo��������.ListIndex = -1
            cboҽ��.ListIndex = -1
            txtҽ������.Text = ""
            txtҽ������.Tag = ""
            Me.lblCap(6).Visible = False
        End If
    End With
    
    Me.stbThis.Panels(2).Text = Mid(Me.stbThis.Panels(2).Text, 1, InStr(1, Me.stbThis.Panels(2).Text, "��:") + 1) & _
                                Me.rptPlist.Rows.Count & "������."
End Sub

Private Sub txtBarCode_GotFocus()
    Me.TxtBarCode.SelStart = 0
    Me.TxtBarCode.SelLength = Len(Me.TxtBarCode)
End Sub

Private Sub TxtBarCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtBarCodeCheck.SetFocus
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub TxtBarCode_LostFocus()
    Dim Record As ReportRecord
    Dim strItems As String
    If Len(Trim(Me.TxtBarCode.Text)) < 2 Then Exit Sub
    
    If Me.TxtBarCode.Text = Me.TxtBarCode.Tag Then Exit Sub
    
    strItems = ""
    
    For Each Record In Me.rptAlist(Me.TabCtr.Selected.Index).Records
        If Val(Mid(Me.TxtBarCode.Text, 1, Len(Record(mAcol.�Թܱ���).Value))) = Record(mAcol.�Թܱ���).Value Then
            If InStr("," & strItems & ",", "," & Record(mAcol.������ĿID).Value & ",") <= 0 Then
                Record(mAcol.ѡ��).Checked = True
                Me.TxtBarCode.Tag = Me.TxtBarCode.Text
            Else
                Record(mAcol.ѡ��).Checked = False
            End If
            strItems = strItems & "," & Record(mAcol.������ĿID).Value
        Else
            Record(mAcol.ѡ��).Checked = False
        End If
    Next
    Me.rptAlist(Me.TabCtr.Selected.Index).Populate
End Sub

Private Sub TxtBarCodeCheck_GotFocus()
    TxtBarCodeCheck.SelStart = 0
    TxtBarCodeCheck.SelLength = Len(TxtBarCodeCheck.Text)
End Sub

Private Sub TxtBarCodeCheck_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.TxtBarCodeCheck.Text <> "" Then
            '������
            cmdBindBarCode_Click
        Else
            If Me.TxtBarCode.Text = "" And Me.TxtBarCodeCheck.Text = "" Then
                '��������
                cmdNewBarcode_Click
            End If
        End If
    End If
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    
End Sub

Private Sub WriterBarCode(Mode As Integer, Optional WriterUser As Boolean = False, _
                          Optional PrintBarCode As Boolean = False, Optional PrintBackBill As Boolean = False, _
                          Optional ByVal intContinue As Integer)
    '����                       д�������
    '����                       Mode =0 ������ =1 �������� =2 ������� = 3 ��ɲɼ� = 4 ��ӡ������ִ��
    '                           WriterUser �Ƿ�д�������(д������˱�ʾ�ɼ����)
    '                           PrintBarCode �Ƿ��ӡ����
    '                           PrintBackBill �Ƿ��ӡ��ִ��
    '                           intContinue   1=�ò�����
    
    Dim blSelect As Boolean                         '�鿴�Ƿ���ѡ�е���
    Dim Record As ReportRecord                      '���м�¼������
    Dim strSQL As String                            'SQL���
    Dim rsTmp As New ADODB.Recordset                '���ݼ�¼��
    Dim intLoop As Integer                          'ѭ������
    Dim strTmp As String                            '��ʱ�ִ�����
    Dim varAdvice As Variant                        'һ��ҽ�����ж����Ŀʱʹ��
    Dim strCuvetteNumber As String                  '�Թܱ���
    Dim strAdvice As String                         'ҽ��ID�����ID
    Dim varItem As Variant                          '���ڷֽ�ҽ��ID�����ID
    Dim strUnion As String                          '�ϲ�ҽ��ID ���ID �����ִ�����"|"�ָ�
    Dim strBarCodePrint As String                   '�������Ŀ�ִ����ڴ�ӡ
    Dim varBarcodePrint As Variant                  '�����ӡ
    Dim strBarCode As String                        '�����ִ�(�ֽ��)
    Dim strItem As String                           '����ֽ����Ŀ
    Dim strBackBill As String                       '��ִ��ҽ���ִ�
    Dim Control As CommandBarControl                '�������ؼ������ж�ʹ����������
    Dim intBaby As Integer                          '�Ƿ���Ӥ��,>0 ��ʾӤ������
    Dim strSample As String                         '�걾
    Dim strAdviceContent As String                  'ҽ������
    Dim lngConnectID As Long                        '���ID
    Dim varFilter As Variant                        '������ͬ����Ŀ
    Dim blnMsgBox As Boolean                        '�Ƿ�����ʾ��ʾ
    Dim strDept As String                           'ִ�п���
    Dim intExecDept As Integer                      '������ִ�п��Ҵ�ӡ
    Dim str���� As String                           '����
    Dim strInfo As String                           '��ʾ��Ϣ
    Dim strҽ��ID As String                         '���ҽ��ID��","�ָ�
    Dim str���� As String                           '��¼��ǰ���ɵ�����
    Dim rsNumber As ADODB.Recordset                 '�����Թܱ������ݼ�
    Dim astrSQL() As String                         'SQL�ִ�
    Dim blnRollBak As Boolean                       '�Ƿ��ǻ�������
    Dim strҽ��ID�� As String                       'ҽ��ID��
    Dim blnPrint As Boolean
    
    ReDim astrSQL(0)
    
    For Each Record In Me.rptAlist(Me.TabCtr.Selected.Index).Records
        If Record(mAcol.ѡ��).Checked = True Then
            If chkDept(Record(mAcol.�ɼ�����ID).Value) = False Then
                '��¼��ʾ��Ϣ����������ʾ
                Record(mAcol.ѡ��).Checked = False
                strInfo = strInfo & "��Ŀ:<" & Record(mAcol.ҽ������).Value & ">�Ĳɼ�ִ�п���<" & Record(mAcol.�ɼ�ִ�п���).Value & ">��������Բ����Ŀ��ҷ�Χ��,���ܰ�����!"
            Else
                strҽ��ID = strҽ��ID & "," & Record(mAcol.ID).Value & "," & Record(mAcol.�ϲ�ҽ��).Value
            End If
        End If
    Next
    
    Me.rptAlist(Me.TabCtr.Selected.Index).Populate
    
    strҽ��ID = Mid(strҽ��ID, 2)
    strҽ��ID = Replace(Replace(strҽ��ID, ";", ","), "|", ",")
    
    
    '���黮�۵�����(����ʱ�ż��)
    If Mode = 0 Or Mode = 1 Then
        If Chk���۷���(Me, strҽ��ID, 0, "E") = False Then
            If strInfo <> "" Then
                MsgBox strInfo
            End If
            Exit Sub
        End If
    End If
    
    For Each Record In Me.rptAlist(Me.TabCtr.Selected.Index).Records
        If Record(mAcol.ѡ��).Checked = True Then
            blSelect = True
            Exit For
        End If
    Next
    

        
    'û�м�¼ʱ�˳�
    If blSelect = False Then
        If strInfo <> "" Then
            MsgBox strInfo
        Else
            MsgBox "û���ҵ����Բ�����ҽ�����ݣ�", vbInformation, gstrSysName
        End If
        If Me.ChkContinuous.Value = 1 Then
            Me.TxtBarCode.SetFocus
        Else
            Me.txtGoto.SetFocus
        End If
        Exit Sub
    End If
        
    '��ʱ�鿴�Ƿ�������
    If Mode = 0 Then
        If Trim(Me.TxtBarCode.Text) = "" Or Trim(Me.TxtBarCodeCheck.Text) = "" Then
            MsgBox "��ɨ�����������!", vbInformation, gstrSysName
            Me.TxtBarCode.SetFocus
            Exit Sub
        End If
        
        If Me.TxtBarCode <> Me.TxtBarCodeCheck Then
            MsgBox "����ɨ�����벻һ��!������ɨ��!", vbInformation, gstrSysName
            Me.TxtBarCode.SetFocus
            Exit Sub
        End If
        
        strSQL = "select �������� from ����ҽ������ where �������� = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, Me.TxtBarCode)
        If rsTmp.EOF = False Then
            If MsgBox("�����Ѵ����Ƿ�ȷ�����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Me.TxtBarCode.SetFocus
                Exit Sub
            End If
        End If
        
    End If
    
    InitRecordSet rsNumber
    
    For Each Record In Me.rptAlist(Me.TabCtr.Selected.Index).Records
        If Record(mAcol.ѡ��).Checked = True And chkDept(Record(mAcol.�ɼ�����ID).Value) = True Then
            
            
            '�Ƿ�����ִ�п��Ҵ�ӡ����
            Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Manage_Transfer_Force, True, True)
            intExecDept = IIf(Control.Checked, 0, 1)
            
            
            Select Case Mode
                Case 0                          '��
                    MakeBarCode rsNumber, Record, Mode, intExecDept, Me.TxtBarCode.Text
                Case 1                          '����
                    MakeBarCode rsNumber, Record, Mode, intExecDept
                Case 2                          'ȡ��
                    MakeBarCode rsNumber, Record, Mode
                Case 3, 4                       '��ɡ���ӡ
                    MakeBarCode rsNumber, Record, Mode
                
            End Select
            
        End If
    Next
    
    On Error GoTo errH
    
    If rsNumber.RecordCount = 0 Then Exit Sub
    rsNumber.MoveFirst
    Select Case Mode
        Case 0, 1                                   '�󶨻���������
            Do Until rsNumber.EOF
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "Zl_����Ԥ������_��������('" & rsNumber("ҽ��ID��") & "','" & rsNumber("��������") & "')"
                If WriterUser = True Then
                    'ִ�����
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    If rsNumber("���") & "" = "��Ѫ" Then
                        astrSQL(UBound(astrSQL)) = "Zl_����Ԥ������_�ɼ����('" & rsNumber("ҽ��ID��") & "','" & UserInfo.��� & "','" & UserInfo.���� & "',0,1)"
                    Else
                        astrSQL(UBound(astrSQL)) = "Zl_����Ԥ������_�ɼ����('" & rsNumber("ҽ��ID��") & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
                    End If
                End If
                rsNumber.MoveNext
            Loop
        Case 2                                     'ȡ����ɻ������
            Do Until rsNumber.EOF
                If TabCtr.Selected.Index = 2 Then
                    'ȡ���ɼ�
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    If rsNumber("���") & "" = "��Ѫ" Then
                        astrSQL(UBound(astrSQL)) = "Zl_����Ԥ������_�ɼ����('" & rsNumber("ҽ��ID��") & "','" & UserInfo.��� & "','" & UserInfo.���� & "',1,1)"
                    Else
                        astrSQL(UBound(astrSQL)) = "Zl_����Ԥ������_�ɼ����('" & rsNumber("ҽ��ID��") & "','" & UserInfo.��� & "','" & UserInfo.���� & "',1)"
                    End If
                    If chkComPlete.Value = 1 Then
                        '���ݲ������Ƿ�ȡ����
                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                        astrSQL(UBound(astrSQL)) = "Zl_����Ԥ������_��������('" & rsNumber("ҽ��ID��") & "')"
                    End If
                Else
                    'ȡ����
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "Zl_����Ԥ������_��������('" & rsNumber("ҽ��ID��") & "')"
                    If TabCtr.Selected.Index = 5 Then
                        intContinue = 2
                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                        If rsNumber("���") & "" = "��Ѫ" Then
                            astrSQL(UBound(astrSQL)) = "Zl_����Ԥ������_�ɼ����('" & rsNumber("ҽ��ID��") & "','" & UserInfo.��� & "','" & UserInfo.���� & "',1,1)"
                        Else
                            astrSQL(UBound(astrSQL)) = "Zl_����Ԥ������_�ɼ����('" & rsNumber("ҽ��ID��") & "','" & UserInfo.��� & "','" & UserInfo.���� & "',1)"
                        End If
                    End If
                End If
                rsNumber.MoveNext
            Loop
        Case 3                                      '��ɲɼ���ȡ����ɡ��زɱ걾
            Do Until rsNumber.EOF
                If TabCtr.Selected.Index = 4 Then
                    '�ز�ʱ��ȡ���걾�ͼ�
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "Zl_LisԤ������_�걾�ͳ�('" & rsNumber("ҽ��ID��") & "',1)"
                    '���²ɼ�
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    If rsNumber("���") & "" = "��Ѫ" Then
                        astrSQL(UBound(astrSQL)) = "Zl_����Ԥ������_�ɼ����('" & rsNumber("ҽ��ID��") & "','" & UserInfo.��� & "','" & UserInfo.���� & "',0,1)"
                    Else
                        astrSQL(UBound(astrSQL)) = "Zl_����Ԥ������_�ɼ����('" & rsNumber("ҽ��ID��") & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
                    End If
                Else
                    '�ò�����ʱ�ͼ�걾
                    If intContinue = 1 Then
                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                        astrSQL(UBound(astrSQL)) = "Zl_LisԤ������_�걾�ͳ�('" & rsNumber("ҽ��ID��") & "',0)"
                    ElseIf TabCtr.Selected.Index = 5 And rsNumber("��������") <> "" Then
                        intContinue = 3
                    End If
                    
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    If rsNumber("���") & "" = "��Ѫ" Then
                        astrSQL(UBound(astrSQL)) = "Zl_����Ԥ������_�ɼ����('" & rsNumber("ҽ��ID��") & "','" & UserInfo.��� & "','" & UserInfo.���� & "',0,1)"
                    Else
                        astrSQL(UBound(astrSQL)) = "Zl_����Ԥ������_�ɼ����('" & rsNumber("ҽ��ID��") & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
                    End If
                End If
                rsNumber.MoveNext
            Loop
    End Select
    
    gcnOracle.BeginTrans
    blnRollBak = True
    
    For intLoop = 1 To UBound(astrSQL)
        If astrSQL(intLoop) <> "" Then
            zlDatabase.ExecuteProcedure astrSQL(intLoop), Me.Caption
        End If
    Next
    gcnOracle.CommitTrans
    
    If ((Mode = 0 Or Mode = 1) And WriterUser = True) Or Mode = 3 Or Mode = 2 Then
        Call WriterBarCodeToLIS(rsNumber, IIf(Mode = 2, 2, 3), intContinue)
    ElseIf (Mode = 0 Or Mode = 1) Then
        Call WriterBarCodeToLIS(rsNumber, 3, 4)
    End If
    
    '��ӡ����
    If PrintBarCode = True And intContinue <> 1 Then
        
        blnPrint = CheckPlugIn(glngSys, glngModul, rsNumber)
        If blnPrint = True Then
            rsNumber.MoveFirst
            Do Until rsNumber.EOF
                '�������뵽PIC
                Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Tool_SignNew, True, True)
                If Control.Checked = True Then
                    Bar39 Me.picBarCodePrint, 3, Nvl(rsNumber("��������")), False, True
                Else
                    Bar128 Me.picBarCodePrint, 3, Nvl(rsNumber("��������")), True
                End If
                SavePicture Me.picBarCodePrint.Image, App.path & "\BarCode.Bmp"
                '��ʼ��ӡ
                Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1211_1", Me, "��������=" & Nvl(rsNumber("��������")), _
                "��Ŀ=" & Replace(Nvl(rsNumber("ҽ������")), ",", " "), _
                "�������� = " & IIf(Nvl(rsNumber("Ӥ��"), 0) = 0, IIf(txt���� <> "", txt����, "��"), rsNumber("Ӥ������")), _
                "�Ա� = " & IIf(Nvl(rsNumber("Ӥ��"), 0) = 0, IIf(cbo�Ա� <> "", cbo�Ա�, "��"), rsNumber("Ӥ���Ա�")), _
                "���� = " & IIf(Nvl(rsNumber("Ӥ��"), 0) = 0, IIf(txt���� & cboAge <> "", txt���� & cboAge & txt����1, "��"), "Ӥ"), _
                "���� = " & IIf(txtBed <> "", txtBed, "��"), _
                "��ʶ�� = " & IIf(txtID <> "", txtID, "��"), _
                "���ڿ��� = " & IIf(Nvl(rsNumber("���ڿ���")) <> "", Nvl(rsNumber("���ڿ���")), "��"), _
                "�ɼ���ʽ = " & IIf(Nvl(rsNumber("�ɼ���ʽ")) <> "", Nvl(rsNumber("�ɼ���ʽ")), "��"), _
                "�걾 = " & IIf(Nvl(rsNumber("�걾")) <> "", Nvl(rsNumber("�걾")), "��"), _
                "ִ�п��� = " & IIf(Nvl(rsNumber("ִ�п���")) <> "", Nvl(rsNumber("ִ�п���")), "��"), _
                "����ҽ�� = " & IIf(Nvl(rsNumber("����ҽ��")) <> "", Nvl(rsNumber("����ҽ��")), "��"), _
                "����ʱ�� = " & IIf(Nvl(rsNumber("����ʱ��")) <> "", Nvl(rsNumber("����ʱ��")), "��"), _
                "������ = " & IIf(Nvl(rsNumber("������")) <> "", Nvl(rsNumber("������")), "��"), _
                "����ʱ�� = " & IIf(Nvl(rsNumber("����ʱ��")) <> "", Nvl(rsNumber("����ʱ��")), "��"), _
                "���� = " & IIf(Nvl(rsNumber("����")) <> "", Nvl(rsNumber("����")), "��"), _
                "��Ѫ�� = " & IIf(Nvl(rsNumber("��Ѫ��")) <> "", Nvl(rsNumber("��Ѫ��")), "��"), _
                "�Թ����� = " & IIf(Nvl(rsNumber("�Թ�����")) <> "", Nvl(rsNumber("�Թ�����")), "��"), _
                "���� = " & IIf(Nvl(rsNumber("������־")) <> "", Nvl(rsNumber("������־")), "��"), _
                "������Դ = " & IIf(Nvl(rsNumber("������Դ")) <> "", Nvl(rsNumber("������Դ")), "��"), _
                "����ͼ��1=" & App.path & "\BarCode.Bmp", 2)
                'ɾ������ͼ��
                Kill App.path & "\BarCode.Bmp"
                strSQL = "Zl_LisԤ������_�����ӡ('" & Replace(rsNumber("ҽ��ID��"), ",,", ",") & "')"
                zlDatabase.ExecuteProcedure strSQL, gstrSysName
                rsNumber.MoveNext
            Loop
        End If
    End If
    
    '��ӡ��ִ��
    If PrintBackBill = True Then
        rsNumber.MoveFirst
        Do Until rsNumber.EOF
            strҽ��ID�� = strҽ��ID�� & "," & rsNumber("ҽ��ID��")
            rsNumber.MoveNext
        Loop
        Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1211_2", Me, "����ID=" & mlngKey, "ҽ��ID��=" & Mid(strҽ��ID��, 2), 2)
    End If
    
    'ˢ������
    If mblFind = False Then
        '�ǲ��ҵĲ���
        With Me.rptPlist
            If .Rows.Count > 0 Then
    '            RefreshAdviceData mlngKey, Me.TabCtr.Selected.Index
                If .FocusedRow Is Nothing Then
                    .FocusedRow = Me.rptPlist.Rows(0)
                    .Populate
                Else
                    RefreshAdviceData .FocusedRow.Record(mPcol.����ID).Value, Me.TabCtr.Selected.Index, IIf(.FocusedRow.Record(mPcol.��Դ).Value = "סԺ", 2, 1), False, .FocusedRow.Record(mPcol.��ҳID).Value
                End If
            Else
                Me.rptAlist(Me.TabCtr.Selected.Index).Records.DeleteAll
                Me.rptCuvette.Records.DeleteAll
                Me.rptAlist(Me.TabCtr.Selected.Index).Populate
                Me.rptCuvette.Populate
                txt���� = ""
                txt����.Tag = ""
                cbo�Ա�.ListIndex = -1
                txt���� = ""
                txt����1 = ""
                txtBed = ""
                txtID = ""
                txtPatientDept = ""
                cbo��������.ListIndex = -1
                cboҽ��.ListIndex = -1
                txtҽ������.Text = ""
                txtҽ������.Tag = ""
                Me.lblCap(6).Visible = False
            End If
        End With
    
        Me.TxtBarCode.Text = ""
        Me.TxtBarCode.Tag = ""
        Me.TxtBarCodeCheck.Text = ""
        
        If Me.ChkContinuous.Value = 1 Then
            Me.TxtBarCode.SetFocus
        Else
            Me.txtGoto.SetFocus
        End If
    Else
        '���Ҳ���
        With Me.rptPlist
            If .Rows.Count > 0 Then
    '            RefreshAdviceData mlngKey, Me.TabCtr.Selected.Index
                If .FocusedRow Is Nothing Then
                    .FocusedRow = Me.rptPlist.Rows(0)
                    .Populate
                Else
                    RefreshAdviceData .FocusedRow.Record(mPcol.����ID).Value, Me.TabCtr.Selected.Index, IIf(.FocusedRow.Record(mPcol.��Դ).Value = "סԺ", 2, 1), False, .FocusedRow.Record(mPcol.��ҳID).Value
                End If
            Else
                Me.rptAlist(Me.TabCtr.Selected.Index).Records.DeleteAll
                Me.rptCuvette.Records.DeleteAll
                Me.rptAlist(Me.TabCtr.Selected.Index).Populate
                Me.rptCuvette.Populate
                txt���� = ""
                txt����.Tag = ""
                cbo�Ա�.ListIndex = -1
                txt���� = ""
                txt����1 = ""
                txtBed = ""
                txtID = ""
                txtPatientDept = ""
                cbo��������.ListIndex = -1
                cboҽ��.ListIndex = -1
                txtҽ������.Text = ""
                txtҽ������.Tag = ""
                Me.lblCap(6).Visible = False
            End If
        End With
        
        Me.TxtBarCode.Text = ""
        Me.TxtBarCode.Tag = ""
        Me.TxtBarCodeCheck.Text = ""
        
        If Mode = 1 Then
            If chkComPlete.Value = 1 Then
                Me.cmdBindBarCode.Enabled = False
                Me.cmdNewBarcode.Enabled = False
                Me.cmdComplete.Enabled = True
                Me.cmdBarcodePrint.Enabled = True
                Me.cmdBakBillPrint.Enabled = True
                Me.cmdBindBarCode.Caption = "�����(&B)"
                Me.cmdNewBarcode.Caption = "ȡ������(&N)"
                Me.cmdComplete.Caption = "ȡ�����(&P)"
            Else
                Me.cmdBindBarCode.Enabled = True
                Me.cmdNewBarcode.Enabled = True
                Me.cmdComplete.Enabled = True
                Me.cmdBarcodePrint.Enabled = True
                Me.cmdBakBillPrint.Enabled = True
                Me.cmdBindBarCode.Caption = "�����(&B)"
                Me.cmdNewBarcode.Caption = "ȡ������(&N)"
                Me.cmdComplete.Caption = "��ɲɼ�(&P)"
            End If
        End If
        
        If strInfo <> "" Then
            MsgBox strInfo
        End If
        
        If Me.ChkContinuous.Value = 1 Then
            Me.TxtBarCode.SetFocus
        Else
            Me.txtGoto.SetFocus
        End If
    End If
    
    '���ɻ���������ת���Ѱ�ҳ
    If chkBindPage.Value = 1 Then
        If Mode = 0 Or Mode = 1 Then TabCtr.Item(1).Selected = True
    End If
    
    Exit Sub
errH:
    If blnRollBak = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckPlugIn(ByVal lngSys As Long, ByVal lngModual As Long, ByVal rsMoneyNow As ADODB.Recordset) As Boolean
'    rsNumber.Fields.Append "���", adVarChar, 20
'    rsNumber.Fields.Append "����", adVarChar, 18
'    rsNumber.Fields.Append "���ID", adBigInt
'    rsNumber.Fields.Append "��������", adVarChar, 18
'    rsNumber.Fields.Append "ִ�п���ID", adVarChar, 18
'    rsNumber.Fields.Append "������ĿID", adVarChar, 18
'    rsNumber.Fields.Append "Ӥ��", adBigInt
'    rsNumber.Fields.Append "������־", adBigInt
'    rsNumber.Fields.Append "�걾", adVarChar, 30
'    rsNumber.Fields.Append "ҽ������", adVarChar, 500
'    rsNumber.Fields.Append "�ɼ���ʽ", adVarChar, 100
'    rsNumber.Fields.Append "����ҽ��", adVarChar, 50
'    rsNumber.Fields.Append "����ʱ��", adDate
'    rsNumber.Fields.Append "������", adVarChar, 50
'    rsNumber.Fields.Append "����ʱ��", adDate
'    rsNumber.Fields.Append "��Ѫ��", adVarChar, 20
'    rsNumber.Fields.Append "�Թ�����", adVarChar, 50
'    rsNumber.Fields.Append "������Դ", adInteger
'    rsNumber.Fields.Append "ҽ��ID��", adVarChar, 500
'    rsNumber.Fields.Append "ִ�п���", adVarChar, 50
'    rsNumber.Fields.Append "Ӥ������", adVarChar, 50
'    rsNumber.Fields.Append "Ӥ���Ա�", adVarChar, 50
'    rsNumber.Fields.Append "�������", adVarChar, 50
    
    Dim blnTmp As Boolean
        On Error Resume Next
        CheckPlugIn = True
        If Not mobjZLIHISPlugIn Is Nothing Then
            blnTmp = mobjZLIHISPlugIn.LisPrintCodeBefore(lngSys, lngModual, rsMoneyNow)
            Call zlPlugInErrH(Err, "LisPrintCodeBefore")
            If Err.Number <> 0 Then
                '�ӿڳ�����,������ӡ
                blnTmp = True
            End If
        Else
            blnTmp = True
        End If
        CheckPlugIn = blnTmp
    Err.Clear: On Error GoTo 0

End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Private Sub cmdNewBarcode_GotFocus()
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (Me.ActiveControl Is cmdNewBarcode And txtGoto.Tag <> "")
End Sub

Private Sub cmdNewBarcode_LostFocus()
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
End Sub

Private Sub txtGoto_Change()
'    If Me.ActiveControl Is txtGoto Then
'        If IDKind.IDKind = IDKinds.C0���� Then
'            If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtGoto.Text = "" And Me.ActiveControl Is txtGoto)
'            If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtGoto.Text = "" And Me.ActiveControl Is txtGoto)
'        ElseIf IDKind.IDKind = IDKinds.C3IC���� Then
'            If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtGoto.Text = "" And Me.ActiveControl Is txtGoto)
'        End If
'    End If
End Sub

Private Sub txtGoto_GotFocus()
    If Me.Visible = False Then Exit Sub
    txtGoto.SelStart = 0
    txtGoto.SelLength = Len(txtGoto.Text)
    If txtGoto.Text = "" And Not txtGoto.Locked Then
'        If IDKind.IDKind = IDKinds.C0���� Then
'            If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (True)
'            If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (True)
'        ElseIf IDKind.IDKind = IDKinds.C3IC���� Then
'            If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (True)
'        ElseIf IDKind.IDKind = IDKinds.C2���֤�� Then
'            If Not mobjIDCard Is Nothing And txtGoto.Text = "" And Not txtGoto.Locked Then mobjIDCard.SetEnabled (True)
'        End If
    End If
End Sub
Private Sub txtGoto_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim blFind As Boolean                           '�Ƿ���ҳɹ�
    Dim Row As ReportRow                            '�б��ж���
    Dim strFind As String                           '�ٽ������ִ�
    Dim blnBarCode As Boolean                       '�Ƿ�����
    Dim blnNo As Boolean                            '�Ƿ񵥾ݺ�
    Dim blnCard As Boolean
    Dim lng�����ID As Long
    Dim lng����ID As Long
    Dim strTmp As String
    Dim NowDate As Date
    Dim strDateBegin As Date
    Dim strDateEnd As Date
    
    On Error GoTo errH
    
    
    If Trim(Me.txtGoto.Text) = "" Then Me.txtGoto.SetFocus: Exit Sub
'    blnCard = zlCommFun.InputIsCard(txtGoto, KeyAscii, mblnShowPwd)
    mstrIndex = IDKind.IDKind
    If IDKind.IDKind = IDKind.GetKindIndex("����") Then
'        blnCard = zlCommFun.InputIsCard(txtGoto, KeyAscii, False)
    End If
    If IDKind.IDKind = IDKind.GetKindIndex("���￨") Then
        If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        If CheckIsInclude(UCase(Chr(KeyAscii)), "'����;��:��?��|,����""") = True Then KeyAscii = 0
        'gbytCardNOLen = Val(IDKind.GetKindItem("���ų���", IDKind.IDKind))
        gbytCardNOLen = IDKind.GetCardNoLen
'        Call zlCommFun.InputIsCard(txtGoto, KeyAscii, True)
        txtGoto.Text = ReplaseSpecial(txtGoto.Text)
        blnCard = KeyAscii <> 8 And Len(txtGoto.Text) = gbytCardNOLen - 1 And txtGoto.SelLength <> Len(txtGoto.Text)
        If blnCard = True And KeyAscii <> 0 Then
            If KeyAscii <> 13 Then
                Me.txtGoto = Me.txtGoto & Chr(KeyAscii)
            End If
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Or (IDKind.IDKind = IDKind.GetKindIndex("���￨") And blnCard = True) Then
        'ˢ�º��ٲ���
'        RefreshPatientData
        '������ٶ���
        txt���� = ""
        txt����.Tag = ""
        cbo�Ա�.ListIndex = -1
        txt���� = ""
        txt����1 = ""
        txtBed = ""
        txtID = ""
        lblCap(0).Caption = "��  ʶ ��"
        txtPatientDept = ""
        cbo��������.ListIndex = -1
        cboҽ��.ListIndex = -1
        txtҽ������.Text = ""
        txtҽ������.Tag = ""
        Me.rptAlist(Me.TabCtr.Selected.Index).Records.DeleteAll
        Me.rptCuvette.Records.DeleteAll
        Me.rptAlist(Me.TabCtr.Selected.Index).Populate
        Me.rptCuvette.Populate
        
        If mbln���֤ Or IDKind.IDKind = IDKind.GetKindIndex("���֤��") Then
'            strsql = "select ����ID from ������Ϣ where ���֤�� = [1] "
'            Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, txtGoto)
'            If Not rsTmp.EOF Then
'                txtGoto = "-" & rsTmp.Fields("����ID")
'            End If
            If mobjSquareCard.zlGetPatiID(IDKind.GetCurCard.����, txtGoto, False, lng����ID) = False Then lng����ID = 0
            If lng����ID <> 0 Then
                txtGoto = "-" & lng����ID
            End If
        ElseIf IDKind.IDKind = IDKinds.C1ҽ���� Then
            
        End If
    
    
        Select Case Mid(txtGoto, 1, 1)
            Case "-"                                '����ID
                blFind = RefreshAdviceData(Val(Mid(txtGoto, 2)), Me.TabCtr.Selected.Index, 1, True)
                strFind = Val(Mid(txtGoto, 2))
            Case "+"                                'סԺ��
                strSQL = "select ����ID from ������Ϣ where סԺ�� = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, Val(Mid(txtGoto, 2)))
                If rsTmp.EOF = False Then blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 2, True)
                strFind = Val(Mid(txtGoto, 2))
            Case "*"                                '�����
                strSQL = "select ����ID from ������Ϣ where ����� = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, Val(Mid(txtGoto, 2)))
                If rsTmp.EOF = False Then blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 1, True)
                strFind = Val(Mid(txtGoto, 2))
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
                blnNo = True
            
            Case Else                               '���￨������
                strFind = txtGoto
                If IDKind.IDKind = IDKind.GetKindIndex("����") And BlnIsNumber(txtGoto) Then
                    strSQL = "select a.����id,a.������Դ from ����ҽ����¼ a , ����ҽ������ b " & _
                         " Where a.ID = b.ҽ��id And b.�������� = [1] order by a.����ʱ�� desc    "
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, txtGoto)
                    If rsTmp.EOF = False Then
                        blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, Nvl(rsTmp(1), 0), True)
                        blnBarCode = True
                    Else
                        strSQL = "select ����ID from ������Ϣ where ���￨�� = [1] "
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, txtGoto)
                        If rsTmp.EOF = False Then
                            blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 1, True)
                        End If
                    End If
                    
                Else
'                    MsgBox IDKind.GetKindIndex("���￨")
                    If blnCard Or IDKind.IDKind = IDKind.GetKindIndex("���￨") Then
                        strSQL = "select ����ID from ������Ϣ where ���￨�� = [1] "
                        strFind = UCase(txtGoto)
                    ElseIf IDKind.IDKind = IDKind.GetKindIndex("ҽ����") Then
                        strSQL = "select ����ID from ������Ϣ where ����ID = [1] "
                        If mobjSquareCard.zlGetPatiID(IDKind.GetCurCard.����, txtGoto, False, lng����ID) = False Then lng����ID = 0
                        strFind = lng����ID
                    ElseIf IDKind.IDKind = IDKind.GetKindIndex("���֤��") Then
                        strSQL = "select ����ID from ������Ϣ where ����ID = [1] "
                        If mobjSquareCard.zlGetPatiID(IDKind.GetCurCard.����, txtGoto, False, lng����ID) = False Then lng����ID = 0
                        strFind = lng����ID
                    ElseIf IDKind.IDKind = IDKind.GetKindIndex("IC����") Then
                        strSQL = "select ����ID from ������Ϣ where ����ID = [1] "
                        If mobjSquareCard.zlGetPatiID(IDKind.GetCurCard.����, txtGoto, False, lng����ID) = False Then lng����ID = 0
                        strFind = lng����ID
                    ElseIf IDKind.IDKind = IDKind.GetKindIndex("�����") Then
                        strSQL = "select ����ID from ������Ϣ where ����ID = [1] "
                        If mobjSquareCard.zlGetPatiID(IDKind.GetCurCard.����, txtGoto, False, lng����ID) = False Then lng����ID = 0
                        strFind = lng����ID
                        'strFind = Val(txtGoto)
                    ElseIf IDKind.IDKind = IDKind.GetKindIndex("����") Then 'ͨ����������ʱ��Ӧע�⵱ǰʱ�䷶Χ
'                        strSQL = "select ����ID from ������Ϣ where ���� = [1] "
'                        strSQL = strSQL & " or ���￨�� = [1] "
                        
                        If rptPlist.Tag = "" Then
                            strTmp = zlDatabase.GetPara("�ɼ�����վ����", 100, 1211, "") '��ע����ж�ȡ��������
                            NowDate = zlDatabase.Currentdate
                            strDateBegin = CDate(Format(NowDate - Val(Split(strTmp, ";")(9)), "yyyy-mm-dd 00:00:00"))
                            strDateEnd = CDate(Format(NowDate, "yyyy-mm-dd 23:59:59"))
                        Else
                            strDateBegin = CDate(Format(Split(rptPlist.Tag, ";")(11), "yyyy-mm-dd 00:00:00"))
                            strDateEnd = CDate(Format(Split(rptPlist.Tag, ";")(12), "yyyy-mm-dd 23:59:59"))
                        End If
                        
                        strSQL = "Select Distinct a.����id" & vbNewLine & _
                            "From ������Ϣ A, ����ҽ����¼ B, ����ҽ������ C" & vbNewLine & _
                            "Where a.����id = b.����id And b.Id = c.ҽ��id And c.����ʱ��+0 Between To_Date('" & strDateBegin & "', 'yyyy-mm-dd hh24:mi:ss') And" & vbNewLine & _
                            "      To_Date('" & strDateEnd & "', 'yyyy-mm-dd hh24:mi:ss') And (a.���� = [1] Or a.���￨�� = [1]) "
                    Else
                        If IDKind.GetCurCard.�ӿ���� <> 0 Then
                            lng�����ID = IDKind.GetCurCard.�ӿ����
                            If mobjSquareCard.zlGetPatiID(lng�����ID, txtGoto, False, lng����ID) = False Then lng����ID = 0
                            If lng����ID = 0 Then lng����ID = 0
                        Else
                            If mobjSquareCard.zlGetPatiID(IDKind.GetCurCard.����, txtGoto, False, lng����ID) = False Then lng����ID = 0
                        End If
                        strSQL = "select ����ID from ������Ϣ where ����ID = [1] "
                        strFind = lng����ID
                    End If
                  
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, strFind)
                    If rsTmp.RecordCount = 0 Then
                        If IDKind.IDKind = IDKind.GetKindIndex("����") Then
                            strFind = Trim(strFind)
                            strFind = Replace(strFind, Chr(&HD), "")
                            If IDKind.GetKindIndex("���￨") <> -1 Then lng�����ID = IDKind.GetIDKindCard("���￨").�ӿ����
                            strSQL = "select ����ID from ����ҽ�ƿ���Ϣ where ���� = [1] and �����id =[2] "
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, strFind, lng�����ID)
                        End If
                    End If
                    If rsTmp.EOF = False Then
                        strFind = rsTmp(0)
                        blFind = RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 1, True)
                    End If
                    
                End If
                    
                
        End Select
        
        On Error Resume Next
        '�������ڶ�λ,�����ͺ���
        mblFind = blFind                '�����ж��Ƿ���ҵ��Ĳ���
        If blFind = True Then
            '���ҳɹ�
            If Me.chkFindMove.Value = 1 Then
                Me.TxtBarCode.SetFocus
            Else
                If cmdNewBarcode.Enabled = True And cmdNewBarcode.Caption = "��������(&N)" Then
                    cmdNewBarcode.SetFocus
                ElseIf cmdComplete.Enabled = True And cmdComplete.Caption = "��ɲɼ�(&P)" Then
                    Me.cmdComplete.SetFocus
                Else
                    cmdNewBarcode.SetFocus
                End If
            End If
            FindPatient txtGoto.Text
        Else
            'û���ҵ�����ʱ�����б��в���һ��
            If FindPatient(txtGoto.Text) = False Then
                'û�в��ҵ�����
                Me.txtGoto.SelStart = 0
                Me.txtGoto.SelLength = Len(Me.txtGoto.Text)
                Me.txtGoto.SetFocus
            Else
                '����͵���ʱֻ�����Ӧ�ļ�¼
                If blnNo = True Or blnBarCode = True Then
                     If rsTmp.EOF = False Then
                        If blnNo = True Then
                            Call RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, 1, True)
                        Else
                            Call RefreshAdviceData(rsTmp(0), Me.TabCtr.Selected.Index, Nvl(rsTmp(1), 1), True)
                        End If
                     End If
                End If
                
                If Me.chkFindMove.Value = 1 Then
                    Me.TxtBarCode.SetFocus
                Else
                    If cmdNewBarcode.Enabled = True And cmdNewBarcode.Caption = "��������(&N)" Then
                        cmdNewBarcode.SetFocus
                    Else
                        cmdComplete.SetFocus
                    End If
                End If
            End If
        End If
        Me.txtGoto.Text = ""
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function FindPatient(strFind As String) As Boolean
    '���Ҳ���
    '����                   '�ɹ�����True ʧ�ܷ���False
    Dim Row  As ReportRow
    Dim intLoop As Integer
    
    For intLoop = 0 To 2
        For Each Row In Me.rptPlist.Rows '����š�סԺ�š�����id�����������￨�š�����
            If (Trim(strFind) = "*" & Trim(Row.Record(mPcol.��ʶ��).Value) And InStr("����,���", Trim(Row.Record(mPcol.��Դ).Value)) > 0) Or _
                (Trim(IDKind.IDKind) = Trim(IDKind.GetKindIndex("�����")) And Trim(strFind) = Trim(Row.Record(mPcol.��ʶ��).Value) And InStr("����,���", Trim(Row.Record(mPcol.��Դ).Value)) > 0) Or _
                (Trim(strFind) = "+" & Trim(Row.Record(mPcol.��ʶ��).Value) And "סԺ" = Trim(Row.Record(mPcol.��Դ).Value)) Or _
                Trim(strFind) = "-" & Trim(Row.Record(mPcol.����ID).Value) Or Trim(Row.Record(mPcol.��������).Value) Like Trim(strFind) & "*" Or _
                Trim(strFind) = Trim(Row.Record(mPcol.���￨).Value) Or InStr(1, "," & Trim(Row.Record(mPcol.����).Value) & ",", "," & Trim(strFind) & ",") > 0 Then
                
                Set Me.rptPlist.FocusedRow = Row
                Me.rptPlist.Populate
                FindPatient = True
                
                Exit Function
            End If
        Next
    Next
End Function

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
    objPrint.Title.Text = "�����ļ��嵥"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub txtGoto_LostFocus()
    If IDKind.IDKind = IDKinds.C0���� Then
        If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    ElseIf IDKind.IDKind = IDKinds.C3IC���� Then
        If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
    ElseIf IDKind.IDKind = IDKinds.C2���֤�� Then
        If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    End If
End Sub


Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(Me.txtҽ������)) > 0 Then
            If Me.cmdNewBarcode.Enabled = True Then
                Me.cmdNewBarcode.SetFocus
            End If
        Else
            zlCommFun.PressKey vbKeyTab
        End If
        Exit Sub
    Else
        KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
    End If
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        KeyCode = Asc(UCase(Chr(KeyCode)))
    Else
        Call AdjustEditState(True)
'        Me.cbo�Ա�.SetFocus
        zlCommFun.PressKey vbKeyTab
    End If
End Sub






Private Function InitDepts() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strOldText As String
    
    On Error GoTo errH
    strOldText = Me.cbo��������.Text
    Me.cbo��������.Clear
    
    strSQL = _
        " Select Distinct A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where B.����ID = A.ID " & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
        " And (B.�������� IN('�ٴ�','���'))" & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    For i = 1 To rsTmp.RecordCount
        cbo��������.AddItem rsTmp!����
        cbo��������.ItemData(cbo��������.NewIndex) = rsTmp!ID
        
        rsTmp.MoveNext
    Next
    
    On Error Resume Next
    Me.cbo��������.Text = strOldText
    If cbo��������.ListCount > 0 And Me.cbo��������.ListIndex = -1 Then cbo��������.ListIndex = 0
    
    InitDepts = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitDoctors(ByVal lng����ID As Long)
'���ܣ���ȡ��ǰ���������а�����������Ա
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    Me.cboҽ��.Clear
    
    '����ҽ����ʿ
    strSQL = _
        "Select Distinct A.ID,B.����ID,A.���,A.����,Upper(A.����) as ����," & _
        " C.��Ա����,Nvl(A.Ƹ�μ���ְ��,0) as ְ��" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID" & _
        " And C.��Ա���� IN('ҽ��') And B.����ID=[1] " & _
        " And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) "
        
    strSQL = strSQL & " Order by ����,��Ա���� Desc"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboҽ��.AddItem rsTmp!����
            cboҽ��.ItemData(cboҽ��.ListCount - 1) = rsTmp!����ID
            
            If rsTmp!ID = UserInfo.ID And cboҽ��.ListIndex = -1 Then cboҽ��.ListIndex = cboҽ��.NewIndex
            rsTmp.MoveNext
        Next
        
        If cboҽ��.ListCount = 1 And cboҽ��.ListIndex = -1 Then cboҽ��.ListIndex = 0
    End If
End Sub


Private Function Get�����������(ByVal int���� As Integer, ByVal txtMainAdvice As String) As String
'���ܣ��������ɼ���������ݵ�ҽ������
'������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
    Dim lngBegin As Long, i As Long
    Dim str���� As String, strTmp As String
    Dim strDate As String
    
    If rsRelativeAdvice Is Nothing Or int���� = 1 Then Get����������� = txtMainAdvice: Exit Function
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("����"))) > 0 Then
            strTmp = strTmp & "," & rsRelativeAdvice("����")
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    
    If strTmp <> "" Then
        Get����������� = IIf(Len(Trim(txtMainAdvice)) = 0, "", txtMainAdvice & " �� ") & Mid(strTmp, 2)
    Else
        Get����������� = txtMainAdvice
    End If
End Function

Private Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
'������intNum=��Ŀ���,Ϊ0ʱ�̶��������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date
    
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(strNO, "0000000")
        Exit Function
    End If
    GetFullNO = strNO
    
    strSQL = "Select ��Ź���,Sysdate as ���� From ������Ʊ� Where ��Ŀ���=" & intNum
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic")
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!��Ź���, 0)
        curDate = rsTmp!����
    End If

    If intType = 1 Then
        '���ձ��
        strSQL = Format(CDate("1992-" & Format(rsTmp!����, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
    Else
        '������
        GetFullNO = PreFixNO & Format(strNO, "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetDictData(strDict As String) As ADODB.Recordset
'���ܣ���ָ�����ֵ��ж�ȡ����
'������strDict=�ֵ��Ӧ�ı���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strDict & " Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then Set GetDictData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtҽ������_GotFocus()
    Call zlControl.TxtSelAll(txtҽ������)
End Sub

Private Sub txtҽ������_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If txtҽ������.Text = txtҽ������.Tag Then
'            zlcommfun.PressKey vbKeyTab
            Me.cmdBindBarCode.Enabled = True
            Me.cmdNewBarcode.Enabled = True
            If chkFindMove.Value = 0 Then
                cmdNewBarcode.SetFocus
            Else
                Me.TxtBarCode.SetFocus
            End If
            Exit Sub
        End If
        
        With txtҽ������
            Set rsTmp = SelectDiagItem()
        End With
        
        If rsTmp Is Nothing Then 'ȡ����������
            '�ָ�ԭֵ
            txtҽ������.Text = txtҽ������.Tag
            zlControl.TxtSelAll txtҽ������
            txtҽ������.SetFocus: Exit Sub
        End If
        '����Ŀ��¼��
        '����ѡ����Ŀ����ȱʡҽ����Ϣ
        If AdviceInput(rsTmp) Then
            DoEvents
            '��ʾ��ȱʡ���õ�ֵ
            txtҽ������.Tag = txtҽ������.Text
        Else
            DoEvents
            '�ָ�ԭֵ
            txtҽ������.Text = txtҽ������.Tag
            zlControl.TxtSelAll txtҽ������

            txtҽ������.SetFocus: Exit Sub
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub
Private Function SelectDiagItem() As ADODB.Recordset
'ѡ�������Ŀ
    Dim strSQL As String
    Dim objPoint As POINTAPI
    
    strSQL = "Select Distinct A.ID,A.����,A.����,nvl(A.���㵥λ,'��') As ���㵥λ,nvl(A.�걾��λ,' ') As �걾��λ," + _
        "Decode(A.���,'H',Decode(A.��������,'1','����ȼ�','������')," + _
        "'E',Decode(A.��������,'1','��������','2','��ҩ;��','3','��ҩ�巨',4,'��ҩ�÷�','����')," + _
        "'Z',Decode(A.��������,'1','����','2','סԺ','3','ת��','4','����','5','��Ժ','6','תԺ','����'),A.��������) As ��Ŀ����,A.��� As ���ID,A.ID As ������ĿID,nvl(ִ��Ƶ��,0) As ִ��Ƶ��ID,nvl(���㷽ʽ,0) As ���㷽ʽID,nvl(ִ�а���,0) As ִ�а���ID,nvl(�Ƽ�����,0) As �Ƽ�����ID,nvl(ִ�п���,0) As ִ�п���ID "
    strSQL = strSQL + "From ������ĿĿ¼ A,������Ŀ���� C,����ִ�п��� D Where A.ID=C.������ĿID And A.ID=D.������ĿID And A.���='C' "       'And D.ִ�п���ID=" & mlngDeptID
    strSQL = strSQL + " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
        "And A.������� IN(" & PatientType & ",3) And Nvl(A.����Ӧ��,0)=1 And Nvl(A.�����Ա�,0) IN (" + _
        IIf(Me.cbo�Ա�.Text Like "*��*", "1,0)", "2,0)") + _
        " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
        " And (A.���� Like '" + txtҽ������ + "%' Or Upper(A.����) Like '" + txtҽ������ + "%' Or Upper(C.����) Like '" + UCase(txtҽ������) + "%')"
            
    Call ClientToScreen(txtҽ������.hWnd, objPoint)
    Set SelectDiagItem = zlDatabase.ShowSelect(Me, strSQL, 0, "ѡ��������Ŀ", True, Me.txtҽ������.Text, "", True, True, True, objPoint.X * 15, objPoint.Y * 15, Me.txtҽ������.Height, False, True)
End Function

Private Function AdviceInput(Optional rsInput As ADODB.Recordset = Nothing) As Boolean
'���ܣ����������������Ŀ(���������)����ȱʡ��ҽ������
'������rsInput=�����ѡ�񷵻صļ�¼��
'���أ�����¼���Ƿ���Ч
    Dim rsTmp As ADODB.Recordset
    Dim strHelpText As String
    Dim strSQL As String
    Dim strExtData As String
    Dim blnOk As Boolean
    Dim t_Pati As TYPE_PatiInfoEx
    
    
    On Error GoTo errH
    '��Ŀ�����������뼰����Ϸ��Լ��
    '---------------------------------------------------------------------------------------------------------------
    If Not rsInput Is Nothing Then txtҽ������.Text = rsInput!����    '��ʱ��ʾ

    '��Ҫ����������ݵ�һЩ��Ŀ
    '---------------------------------------------------------------------------------------------------------------
    '������Ŀѡ�����걾
    strHelpText = "������Ŀ"
    
    If Not rsInput Is Nothing Then
        strExtData = rsInput!������ĿID & ";" & rsInput!�걾��λ    '��������Ŀ
    Else
        strExtData = mstrExtData   '��������Ŀ
    End If
    
    On Error Resume Next
    '�ӿڸ��죺int����û�д������ڴ�Ϊ0�� bytUseType ��ǰû�������ڴ�Ϊ0
    blnOk = frmAdviceEditEx.ShowMe(Me, Me.txtҽ������.hWnd, t_Pati, 0, 4, 0, 1, PatientType, , , , 0, strExtData, , , , , True)
    On Error GoTo errH

    If Not blnOk Then Exit Function
    If strExtData = "" Or Mid(strExtData, 1, 1) = ";" Then Exit Function
    
    '��ȡ�ɼ���ʽ
    Set rsTmp = SelectCap(Split(Split(strExtData, ";")(0), ",")(0))
    If rsTmp Is Nothing Then
        MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, gstrSysName
        Exit Function
    End If
    mlngCapID = rsTmp("ID")
    
    strSQL = "Select C.��Ŀ��� From ������ĿĿ¼ A,���鱨����Ŀ B,������Ŀ C " & _
        "Where A.ID=B.������ĿID And B.������ĿID=C.������ĿID And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Split(Split(strExtData, ";")(0), ",")(0))
    If rsTmp.EOF Then
        mbln΢������Ŀ = False
    Else
        mbln΢������Ŀ = IIf(Nvl(rsTmp("��Ŀ���"), 0) = 2, True, False)
    End If
    
    mstrExtData = strExtData
    
    
    Call AdviceSet�������(3, mstrExtData)
    txtҽ������.Text = Get�����������(2, "")
    txtҽ������.Text = txtҽ������.Text & "(" & Split(mstrExtData, ";")(1) & ")"
    
    '����ҽ��
    On Error Resume Next
    If Me.cboҽ��.Text = "" Then Me.cboҽ��.ListIndex = 0
    
    AdviceInput = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function SelectCap(Optional ByVal lngItemID As Long = 0) As ADODB.Recordset
'��ȡ�ɼ���ʽ
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim tmpRect As RECT
    
    On Error GoTo DBError
        
    strSQL = "Select Distinct A.ID,A.����,A.���� " + _
        "From ������ĿĿ¼ A,�����÷����� D Where A.ID=D.�÷�ID" + _
        " And A.���='E' And A.��������='6'" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
        " And A.������� IN(" & PatientType & ",3) And Nvl(A.�����Ա�,0) IN (" + _
        IIf(Me.cbo�Ա�.Text Like "*��*", "1,0)", "2,0)") + _
        " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
        " And D.��ĿID=" & lngItemID
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.EOF Then
        strSQL = "Select Distinct A.ID,A.����,A.���� " + _
            "From ������ĿĿ¼ A Where " + _
            " A.���='E' And A.��������='6'" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
            " And A.������� IN(" & PatientType & ",3) And Nvl(A.�����Ա�,0) IN (" + _
            IIf(Me.cbo�Ա�.Text Like "*��*", "1,0)", "2,0)") + _
            " And Nvl(A.ִ��Ƶ��,0) IN(0,1)"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then Set SelectCap = rsTmp
    
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceSet�������(ByVal int���� As Integer, ByVal strDataIDs As String)
'���ܣ�1.��������ָ����������Ŀ�Ĳ�λ��,�����������������Ŀ���޸Ĳ�λ
'      2.��������ָ��������Ŀ�ĸ���������������Ŀ��,����������������Ŀ��������Ŀ�ĸ���������������Ŀ
'������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
'      strDataIDs=���:������鲿λ��Ϣ,����:��������������������Ŀ��Ϣ,���п���û�и�������������
    Dim strSQL As String, i As Long
    Dim arrIDs As Variant
    
    On Error GoTo errH
            
    '���������Ŀ
    strDataIDs = Mid(strDataIDs, 1, InStr(strDataIDs, ";") - 1)
    
    If strDataIDs <> "" Then
        If Not rsRelativeAdvice Is Nothing Then
            rsRelativeAdvice.Close
        Else
            Set rsRelativeAdvice = New ADODB.Recordset
        End If
        strSQL = "Select ID,����,����,nvl(�걾��λ,' ') As �걾��λ," + _
        "���,nvl(�Ƽ�����,0) As �Ƽ�����,nvl(ִ�п���,0) As ִ�п���,�������� From ������ĿĿ¼ Where ID IN(" & strDataIDs & ")"
        Set rsRelativeAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Else
        If Not rsRelativeAdvice Is Nothing Then rsRelativeAdvice.Close: Set rsRelativeAdvice = Nothing
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'���ܣ����ش�д�ĵ��ݺ���ǰ׺
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function
'���ҽ�����ݵĺϷ���
Private Function ValidAdvice() As Boolean
    ValidAdvice = True
    
    On Error Resume Next
    If txt����.Text = "" Then
        ValidAdvice = False
        MsgBox "�����벡�˵�������", vbInformation, gstrSysName: DoEvents
'        mintFocusItem = FocusItem.����
        txt����.SetFocus: Exit Function
    End If
    
    If Len(Trim(Me.txtҽ������)) = 0 Then
        ValidAdvice = False
        MsgBox "��������������Ŀ��", vbInformation, gstrSysName: DoEvents
'        mintFocusItem = FocusItem.ҽ������
        Me.txtҽ������.SetFocus: Exit Function
    End If
    If Me.cbo��������.ListIndex = -1 Then
        ValidAdvice = False
        MsgBox "��ָ���������ң�", vbInformation, gstrSysName: DoEvents
'        mintFocusItem = FocusItem.��������
        Me.cbo��������.SetFocus: Exit Function
    End If
    If Len(Trim(Me.cboҽ��.Text)) = 0 Then
        ValidAdvice = False
        MsgBox "��ָ������ҽ����", vbInformation, gstrSysName: DoEvents
'        mintFocusItem = FocusItem.ҽ��
        Me.cboҽ��.SetFocus: Exit Function
    End If
End Function


Private Function SaveAdviceData() As Long
    Dim strSQL As String, strDate As String, strNO As String
    Dim lngAdviceID As Long, lngTmpID As Long, lngSendNO As Long
    Dim iMaxSeq As Integer, iSendSeq As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim lng��������ID As Long, lng����ID As Long, strDoctor As String, i As Integer
    Dim strִ�п���ID As String, strִ�п���ID1 As String, lngDept As Long
    Dim rsCard As ADODB.Recordset
    Dim tmpstr��� As String, tmplngClinicID As Long, tmpint�Ƽ����� As Integer, tmpintִ������ As Integer
    Dim rsDept As ADODB.Recordset
    Dim intPatientSource As Integer                     '������Դ
    Dim astrSQL() As String
    Dim blnRollBack As Boolean
    Dim intLoop As Integer
    Dim strCostType As String, lngJ As Long
    On Error GoTo ErrHand
    ReDim astrSQL(0)

    On Error GoTo ErrHand
    
    
    '���没����Ϣ
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    If PatientType = 1 Then '���ﲡ��
        If mlng����ID > 0 Then '���еĲ���
'            strsql = _
'                "zl_�ҺŲ��˲���_INSERT(3," & mlng����ID & ",Null," & _
'                "'',''," & _
'                "'" & txt����.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & txt����.Text & Me.cboAge.Text & Me.txt����1.Text & "'," & _
'                "'�Է�','�Է�'," & _
'                "'','',''," & _
'                "'','','',0,'','','','',''," & strDate & ",NULL)"
        Else '�²���
            '��ӻ�ȡĬ�Ϸѱ�
            strSQL = "select ����,ȱʡ��־ from �ѱ� order by ���� "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork")
            Do While Not rsTmp.EOF
                lngJ = lngJ + 1
                If lngJ = 1 Then
                    strCostType = rsTmp("����")
                End If
                If rsTmp("ȱʡ��־") = 1 Then
                    strCostType = rsTmp("����")
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            If strCostType = "" Then strCostType = "�Է�"
        
            mlng����ID = zlDatabase.GetNextNo(1)
            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
            astrSQL(UBound(astrSQL)) = _
                "zl_�ҺŲ��˲���_INSERT(1," & mlng����ID & ",Null," & _
                "'',''," & _
                "'" & txt����.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & txt����.Text & Me.cboAge.Text & Me.txt����1.Text & "'," & _
                "'" & strCostType & "','" & strCostType & "'," & _
                "'','',''," & _
                "'','','',0,'','','','',''," & strDate & ",NULL)"
        End If
    End If
    '����ҽ��������
    lngAdviceID = zlDatabase.GetNextId("����ҽ����¼")
    iMaxSeq = 0
    
    lng��������ID = Me.cbo��������.ItemData(Me.cbo��������.ListIndex)
    strDoctor = NeedName(Me.cboҽ��.Text)
    
    If rsRelativeAdvice.RecordCount = 0 Then
        strִ�п���ID = mlngDeptID
    Else
        'PatientType
        If mlng����ID > 0 Then
            strSQL = "select  ִ�п���ID from  ����ִ�п��� where ������Դ = [1] and ������ĿID = [2] "
        Else
            strSQL = "select ִ�п���id from ����ִ�п��� where ������Ŀid = [2]"
        End If
        rsRelativeAdvice.MoveFirst
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, PatientType, CLng(rsRelativeAdvice("Id")))
        strִ�п���ID = Val(Nvl(rsTmp("ִ�п���ID")))
    End If
    
    iSendSeq = 1
    '������Ŀ���ɼ���ʽ��Ϊ��ҽ��
    tmplngClinicID = mlngCapID
    'ȡ�ɼ���ʽ��ִ�в���
    strִ�п���ID1 = "NULL"
    
    lngSendNO = zlDatabase.GetNextNo(10)
    strNO = zlDatabase.GetNextNo(IIf(PatientType = 2, 14, 13))
    
    '�������ҽ��
    If Not rsRelativeAdvice Is Nothing Then
        i = 2
        rsRelativeAdvice.MoveFirst
        Do While Not rsRelativeAdvice.EOF
            lngTmpID = zlDatabase.GetNextId("����ҽ����¼")
            With rsRelativeAdvice
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "ZL_����ҽ����¼_Insert(" & lngTmpID & "," & lngAdviceID & "," & _
                    (iMaxSeq + i) & ",3," & mlng����ID & ",NULL," & _
                    "0,1," & _
                    "1,'" & .Fields("���") & "'," & _
                    .Fields("ID") & ",NULL,NULL,NULL,NULL," & _
                    "'" & Replace(.Fields("����"), "'", "''") & "',''," & _
                    "'" & .Fields("�걾��λ") & "','һ����',NULL,NULL,'',NULL," & _
                    .Fields("�Ƽ�����") & "," & _
                    strִ�п���ID & "," & _
                    .Fields("ִ�п���") & ",0," & strDate & ",NULL," & _
                    IIf(Me.txtPatientDept.Tag = 0, lng��������ID, Me.txtPatientDept.Tag) & "," & lng��������ID & ",'" & strDoctor & "'," & _
                    "Sysdate,'',Null)"
                iSendSeq = iSendSeq + 1
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "ZL_����ҽ������_Insert(" & _
                    lngTmpID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
                    iSendSeq & ",NULL,NULL,NULL," & _
                    "Sysdate+1/(24*3600)," & _
                    "0," & strִ�п���ID & ",0,0)"
                i = i + 1
                .MoveNext
            End With
        Loop
    End If
    '��������Ĳɼ���ʽ�ŵ����
    iMaxSeq = iMaxSeq + 1
    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
    astrSQL(UBound(astrSQL)) = "ZL_����ҽ����¼_Insert(" & lngAdviceID & ",NULL," & _
        iMaxSeq & ",3," & mlng����ID & ",NULL," & _
        "0,1," & _
        "1,'E'," & mlngCapID & ",NULL,NULL,NULL,NULL," & _
        "'" & Replace(Me.txtҽ������, "'", "''") & "',''," & _
        "'','һ����',NULL,NULL,'',NULL,2," & _
        strִ�п���ID1 & ",3,0," & strDate & ",NULL," & _
        IIf(Me.txtPatientDept.Tag = 0, lng��������ID, Me.txtPatientDept.Tag) & "," & lng��������ID & ",'" & strDoctor & "'," & _
        "Sysdate,'',Null)"
    iSendSeq = iSendSeq + 1
    '������ҽ��
    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
    astrSQL(UBound(astrSQL)) = "ZL_����ҽ������_Insert(" & _
        lngAdviceID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
        iSendSeq & ",NULL,NULL,NULL," & _
        "Sysdate+1/(24*3600)," & _
        "0," & strִ�п���ID & ",0,1)"
    
    gcnOracle.BeginTrans
    blnRollBack = True
    For intLoop = 1 To UBound(astrSQL)
        If astrSQL(intLoop) <> "" Then
            zlDatabase.ExecuteProcedure astrSQL(intLoop), "д��ҽ��"
        End If
    Next
    gcnOracle.CommitTrans
    SaveAdviceData = mlng����ID
    Exit Function
ErrHand:
    mlng����ID = 0
    If blnRollBack = True Then
        gcnOracle.RollbackTrans
    End If
'    Err.Raise Err.Number, "�걾����"
    Exit Function
End Function

Private Sub AdjustEditState(blEnable As Boolean)
    '����:              �����༭״̬
    'Me.txt����.Enabled = blEnable
    cbo�Ա�.Enabled = blEnable
    txt����.Enabled = blEnable
    txt����1.Enabled = blEnable
    cboAge.Enabled = blEnable
    cbo��������.Enabled = blEnable
    cboҽ��.Enabled = blEnable
    txtҽ������.Enabled = blEnable
    cmdSelect.Enabled = blEnable
End Sub

Private Sub HideBarCode()
    '����Ԥ������
    Dim Control As CommandBarControl
    Set Control = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_View_PriceTable, True, True)
    
    If Control.Checked = True Then
        Me.Frame1.Visible = False
        Me.Frame2.Top = Me.Frame1.Top - 20
        Me.fraBarCode.Height = Me.Frame2.Top + Me.Frame2.Height + 120
    Else
        Me.Frame1.Visible = True
        Me.Frame2.Top = Me.Frame1.Top + Me.Frame1.Height + 20
        Me.fraBarCode.Height = Me.Frame2.Top + Me.Frame2.Height + 120
    End If
    
    Call Form_Resize
End Sub

Private Function chkDept(lngDept As Long) As Boolean
    '����ִ�п����Ƿ��ڵ�ǰ�ɲ����Ŀ�����
    
    Dim cboCtrol As CommandBarComboBox              '����
    Dim intLoop As Integer
    
    Set cboCtrol = Me.cbrthis.FindControl(, conMenu_View_Busy, True, True)
    
    For intLoop = 1 To cboCtrol.ListCount
        If cboCtrol.ItemData(intLoop) = lngDept Then
            chkDept = True
            Exit Function
        End If
    Next

End Function
Private Sub InitRecordSet(rsNumber As ADODB.Recordset)
    '��ʼ����¼��
    
    '��¼�Թܱ���
    Set rsNumber = New ADODB.Recordset
    rsNumber.Fields.Append "���", adVarChar, 20
    rsNumber.Fields.Append "����", adVarChar, 18
    rsNumber.Fields.Append "���ID", adBigInt
    rsNumber.Fields.Append "��������", adVarChar, 18
    rsNumber.Fields.Append "ִ�п���ID", adVarChar, 18
    rsNumber.Fields.Append "������ĿID", adVarChar, 18
    rsNumber.Fields.Append "Ӥ��", adBigInt
    rsNumber.Fields.Append "������־", adBigInt
    rsNumber.Fields.Append "�걾", adVarChar, 30
    rsNumber.Fields.Append "ҽ������", adVarChar, 500
    rsNumber.Fields.Append "�ɼ���ʽ", adVarChar, 100
    rsNumber.Fields.Append "����ҽ��", adVarChar, 100
    rsNumber.Fields.Append "����ʱ��", adDate
    rsNumber.Fields.Append "������", adVarChar, 50
    rsNumber.Fields.Append "����ʱ��", adDate
    rsNumber.Fields.Append "��Ѫ��", adVarChar, 20
    rsNumber.Fields.Append "�Թ�����", adVarChar, 50
    rsNumber.Fields.Append "������Դ", adInteger
    rsNumber.Fields.Append "ҽ��ID��", adVarChar, 500
    rsNumber.Fields.Append "ִ�п���", adVarChar, 200
    rsNumber.Fields.Append "Ӥ������", adVarChar, 50
    rsNumber.Fields.Append "Ӥ���Ա�", adVarChar, 50
    rsNumber.Fields.Append "�������", adVarChar, 200
    rsNumber.Fields.Append "���ڿ���", adVarChar, 200
    
    rsNumber.CursorLocation = adUseClient
    rsNumber.LockType = adLockOptimistic
    rsNumber.CursorType = adOpenStatic
    rsNumber.Open
    
End Sub

Public Function MakeBarCode(rsNumber As ADODB.Recordset, RowRecord As ReportRecord, intMode As Integer, Optional intExecDept As Integer, Optional strBarCode As String) As Boolean
'����                   �������벢��¼������汣�浽���ݻ��ӡ
'����                   ���ڼ�¼�ļ�¼��
'                       RowRecord������
'                       'ִ�п����Ƿ�Ҫ����
'                       Mode =0 ������ =1 �������� =2 ������� = 3 ��ɲɼ� = 4 ��ӡ������ִ��
'                       strBarCode <> ""ʱ��ʾʹ�ð�����
    Dim strFilter As String
    Dim blnNew As Boolean
    Dim strҽ������ As String
    Dim strErr As String

    blnNew = False
    Select Case intMode
    Case 0                              '��
        If rsNumber.RecordCount = 0 Then blnNew = True
    Case 1                              '����
        strFilter = "������ĿID=" & RowRecord.Item(mAcol.������ĿID).Value
        rsNumber.filter = strFilter
        If rsNumber.EOF = False Then
            '��������Ŀ��ͬʱ����һ������
            blnNew = True
        Else
            On Error GoTo GoOn  '����LIS����10.35.130û��IsToleranceItem��ڵ����
            If mobjLisInsideComm.IsToleranceItem(Val(RowRecord.Item(mAcol.������ĿID).Value), strErr) Then
                '��������������������
                If strErr <> "" Then
                    MsgBox strErr, vbInformation, Me.Caption
                End If
                blnNew = True
            End If
GoOn:
            Err.Clear
            On Error GoTo 0
            strFilter = "����='" & RowRecord.Item(mAcol.�Թܱ���).Value & _
                        "' And Ӥ��=" & RowRecord.Item(mAcol.Ӥ��).Value & _
                      " And ������־=" & IIf(RowRecord.Item(mAcol.����).Value = "����", 1, 0) & _
                      " And �걾='" & RowRecord.Item(mAcol.�걾).Value & "'" & _
                        IIf(Me.chkApplyDept.Value = 1, " and �������='" & RowRecord.Item(mAcol.�������).Value & "'", "")
            If intExecDept = 1 Then strFilter = strFilter & " And ִ�п���id=" & RowRecord.Item(mAcol.����ִ�п���ID).Value
            rsNumber.filter = strFilter
            If rsNumber.EOF = True Then
                '����������
                blnNew = True
            End If
        End If
    Case 2                              'ȡ������
        If rsNumber.RecordCount = 0 Then blnNew = True

    Case 3, 4                           '���������ӡ
        strFilter = "��������='" & RowRecord.Item(mAcol.����).Value & "'"
        rsNumber.filter = strFilter
        If rsNumber.EOF = True Then
            blnNew = True
        End If
    End Select
    If blnNew = True Then
        rsNumber.AddNew
        '�󶨺���������
        rsNumber!��� = RowRecord.Item(mAcol.���).Value
        If strBarCode <> "" Then
            rsNumber!�������� = strBarCode
        Else
            If intMode = 3 Or intMode = 4 Then
                rsNumber!�������� = RowRecord.Item(mAcol.����).Value
            Else
                rsNumber!�������� = zlDatabase.GetNextNo(125, Split(RowRecord.Item(mAcol.ID).Value, ",")(0))
            End If
        End If
        rsNumber!������� = RowRecord.Item(mAcol.�������).Value
        rsNumber!�ɼ���ʽ = RowRecord.Item(mAcol.�ɼ���ʽ).Value
        rsNumber!�걾 = RowRecord.Item(mAcol.�걾).Value
        rsNumber!ִ�п���ID = RowRecord.Item(mAcol.����ִ�п���ID).Value
        rsNumber!����ҽ�� = RowRecord.Item(mAcol.����ҽ��).Value
        rsNumber!����ʱ�� = RowRecord.Item(mAcol.����ʱ��).Value
        rsNumber!������ = RowRecord.Item(mAcol.������).Value
        If RowRecord.Item(mAcol.����ʱ��).Value <> "" Then
            rsNumber!����ʱ�� = RowRecord.Item(mAcol.����ʱ��).Value
        End If
        rsNumber!���� = RowRecord.Item(mAcol.�Թܱ���).Value
        rsNumber!��Ѫ�� = RowRecord.Item(mAcol.��Ѫ��).Value
        rsNumber!�Թ����� = RowRecord.Item(mAcol.�Թ�����).Value
        rsNumber!������־ = IIf(RowRecord.Item(mAcol.����).Value = "����", 1, 0)
        rsNumber!������Դ = RowRecord.Item(mAcol.������Դ).Value
        rsNumber!Ӥ�� = RowRecord.Item(mAcol.Ӥ��).Value
        rsNumber!ִ�п��� = RowRecord.Item(mAcol.ִ�п���).Value
        rsNumber!ҽ������ = RowRecord.Item(mAcol.����).Value
        rsNumber!������ĿID = RowRecord.Item(mAcol.������ĿID).Value
        rsNumber!Ӥ������ = RowRecord.Item(mAcol.Ӥ������).Value
        rsNumber!Ӥ���Ա� = RowRecord.Item(mAcol.Ӥ���Ա�).Value
        rsNumber!���ڿ��� = RowRecord.Item(mAcol.�������ڿ���).Value
        rsNumber!ҽ��ID�� = Replace(Replace(RowRecord.Item(mAcol.ID).Value & "," & RowRecord.Item(mAcol.�ϲ�ҽ��).Value, ";", ","), ",,", ",")
        If Left(rsNumber!ҽ��ID��, 1) = "," Then rsNumber!ҽ��ID�� = Mid(rsNumber!ҽ��ID��, 2)
        If Right(rsNumber!ҽ��ID��, 1) = "," Then rsNumber!ҽ��ID�� = Mid(rsNumber!ҽ��ID��, 1, Len(rsNumber!ҽ��ID��) - 1)
        rsNumber.Update
    Else
        If rsNumber.RecordCount > 0 Then
            rsNumber.MoveLast
            strҽ������ = IIf(Trim(RowRecord.Item(mAcol.����).Value) = "", RowRecord.Item(mAcol.ҽ������).Value, RowRecord.Item(mAcol.����).Value)
            If InStr(";" & rsNumber!ҽ������ & ";", ";" & strҽ������ & ";") <= 0 Then
                rsNumber!ҽ������ = rsNumber!ҽ������ & ";" & strҽ������
            End If

            rsNumber!ҽ��ID�� = rsNumber!ҽ��ID�� & "," & Replace(Replace(RowRecord.Item(mAcol.ID).Value & RowRecord.Item(mAcol.�ϲ�ҽ��).Value, ";", ","), ",,", ",")
            If Left(rsNumber!ҽ��ID��, 1) = "," Then rsNumber!ҽ��ID�� = Mid(rsNumber!ҽ��ID��, 2)
            If Right(rsNumber!ҽ��ID��, 1) = "," Then rsNumber!ҽ��ID�� = Mid(rsNumber!ҽ��ID��, 1, Len(rsNumber!ҽ��ID��) - 1)
            rsNumber.Update
        End If
    End If
    rsNumber.filter = ""
End Function

Private Sub mobjICCard_ShowICCardInfo(ByVal strNO As String)
    Dim lngPreIDKind As Long
    
    If Not txtGoto.Locked And txtGoto.Text = "" And Me.ActiveControl Is txtGoto And strNO <> "" Then
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKinds.C3IC����
        txtGoto.Text = strNO
        txtGoto.Tag = strNO
        Call txtGoto_KeyPress(vbKeyReturn)
        If txtGoto.Text = "" Then Call mobjICCard.SetEnabled(False)
       
        IDKind.IDKind = lngPreIDKind
    ElseIf Me.ActiveControl Is cmdNewBarcode And txtGoto.Tag = strNO Then
        Call cmdNewBarcode_Click
        txtGoto.Tag = ""
        Call mobjICCard.SetEnabled(False)
    End If
End Sub

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
Private Sub FilterPatient()
    '���ܹ��˲����б�
    Dim intLoop As Integer
    
    For intLoop = 0 To Me.rptPlist.Rows.Count - 1
        'δ��
        If Me.optFilter(1).Value = True Then
            If Val(Me.rptPlist.Rows(intLoop).Record.Item(mPcol.δ��).Value) > 0 Then
                Me.rptPlist.Rows(intLoop).Record.Visible = True
            Else
                Me.rptPlist.Rows(intLoop).Record.Visible = False
            End If
        End If
        '�Ѱ�
        If Me.optFilter(2).Value = True Then
            If Val(Me.rptPlist.Rows(intLoop).Record.Item(mPcol.�Ѱ�).Value) > 0 Then
                Me.rptPlist.Rows(intLoop).Record.Visible = True
            Else
                Me.rptPlist.Rows(intLoop).Record.Visible = False
            End If
        End If
        '�Ѳ���
        If Me.optFilter(3).Value = True Then
            If Val(Me.rptPlist.Rows(intLoop).Record.Item(mPcol.�Ѳ���).Value) > 0 Then
                Me.rptPlist.Rows(intLoop).Record.Visible = True
            Else
                Me.rptPlist.Rows(intLoop).Record.Visible = False
            End If
        End If
        '���ͼ�
        If Me.optFilter(4).Value = True Then
            If Val(Me.rptPlist.Rows(intLoop).Record.Item(mPcol.���ͼ�).Value) > 0 Then
                Me.rptPlist.Rows(intLoop).Record.Visible = True
            Else
                Me.rptPlist.Rows(intLoop).Record.Visible = False
            End If
        End If
        '��ִ��
        If Me.optFilter(5).Value = True Then
            If Val(Me.rptPlist.Rows(intLoop).Record.Item(mPcol.��ִ��).Value) > 0 Then
                Me.rptPlist.Rows(intLoop).Record.Visible = True
            Else
                Me.rptPlist.Rows(intLoop).Record.Visible = False
            End If
        End If
        '����
        If Me.optFilter(6).Value = True Then
            If Val(Me.rptPlist.Rows(intLoop).Record.Item(mPcol.����).Value) > 0 Then
                Me.rptPlist.Rows(intLoop).Record.Visible = True
            Else
                Me.rptPlist.Rows(intLoop).Record.Visible = False
            End If
        End If
    Next
    Me.rptPlist.Populate
End Sub

Private Function CheckMoeny() As Boolean
    '����           ����Ƿ��շ�����շѴ��շ�ȷ�ϴ���
    Dim strAdvice As String
    Dim lngLoop As Long
    Dim intProperties As Integer
    Dim intPatientType As Integer
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    '���ݲ������ж��Ƿ���
    If mblnNowConsumption = False Then: CheckMoeny = True: Exit Function
    If mobjSquareCard Is Nothing Then Exit Function
    
    With Me.rptAlist(Me.TabCtr.Selected.Index)
        For lngLoop = 0 To .Records.Count - 1
            If .Records(lngLoop).Item(mAcol.ѡ��).Checked = True Then
                If .Records(lngLoop).Visible = True And .Records(lngLoop).Item(mAcol.�Ʒ�״̬).Value <> 3 Then
                    strAdvice = strAdvice & "," & .Records(lngLoop).Item(mAcol.ID).Value & "," & .Records(lngLoop).Item(mAcol.���ID).Value & "," & _
                        .Records(lngLoop).Item(mAcol.�ϲ�ҽ��).Value
                    intProperties = Val(.Records(lngLoop).Item(mAcol.��¼����).Value)
                    intPatientType = Val(.Records(lngLoop).Item(mAcol.������Դ).Value)
                End If
            End If
        Next
        'ֻ�����ﲡ�˲Ŵ���
        If intPatientType = 1 Then
            If strAdvice <> "" Then strAdvice = Mid(strAdvice, 2)
            If mobjSquareCard.zlSquareAffirm(Me, glngModul, "", mlngKey, 0, False, , , strAdvice) = False Then
                Exit Function
            End If
        Else
            '����סԺ���͵������շѵ����
            strSQL = "Select Count(ID) Count" & vbNewLine & _
                    "From ������ü�¼" & vbNewLine & _
                    "Where ����id = [1] And ҽ����� In (Select * From Table(Cast(F_Num2list([2]) As Zltools.T_Numlist))) And ��¼״̬ = 0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����ȷ��", mlngKey, strAdvice)
            If rsTmp.RecordCount > 0 Then
                If Val(rsTmp("Count") & "") > 0 Then
                    If mobjSquareCard.zlSquareAffirm(Me, glngModul, "", mlngKey, 0, False, , , strAdvice) = False Then
                        Exit Function
                    End If
                End If
            End If

        End If
    End With
    CheckMoeny = True
End Function

Private Function ReplaseSpecial(strTmp As String) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����               �滻�����ַ�
    '����
    '                   ���滻���ַ�
    '����               ���滻�������ַ�����ִ�
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intLoop As Integer
    Dim strSpecial As String
    Dim astrTmp() As String
    strSpecial = "'^��^��^;^��^:^��^?^��^|^,^��^.^��^"""
    astrTmp = Split(strSpecial, "^")
    For intLoop = 0 To UBound(astrTmp)
        strTmp = Replace$(strTmp, astrTmp(intLoop), "")
    Next
    ReplaseSpecial = strTmp
    
End Function

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
    
    If Me.rptPlist.FocusedRow Is Nothing Then
        getTATTime = False
        Exit Function
    End If
    
    '��ȡ�����Ա���������
    With Me.rptPlist.FocusedRow
        strSex = .Record(mPcol.�Ա�).Value
        strDept = .Record(mPcol.���˿���).Value
    End With
    
    '��ȡ��ĿID,��Ŀ����,����ʱ��,����
    strItem = ""
    For Each Record In Me.rptAlist(Me.TabCtr.Selected.Index).Records
        If Record(mAcol.ѡ��).Checked = True Then
'            If Record(mAcol.����ʱ��).Value <> "" Then
                var_Item = Split(Record(mAcol.ҽ������).Value, ";")
                For i = LBound(var_Item) To UBound(var_Item)
                    strItem = strItem & ";" & Record(mAcol.������ĿID).Value & "," & var_Item(i) & _
                                            "," & Record(mAcol.����ʱ��).Value & "," & IIf(Record(mAcol.����).Value = "����", 1, 0) & _
                                             "," & Record(mAcol.ID).Value & "," & Record(mAcol.����).Value
                Next
'            Else
'                strMsgNoTime = strMsgNoTime & Record(mAcol.ҽ������).Value & vbCrLf
'            End If
        End If
    Next
    
    If strMsgNoTime <> "" Then MsgBox strMsgNoTime & "δ����,�����ͼ�", vbInformation, Me.Caption
    If strItem <> "" Then
        strItem = Mid(strItem, 2)
    Else
        getTATTime = False
        strIDs = ""
        Exit Function
    End If
    
    '���TAT�Ƿ�ʱ
    On Error GoTo errold
    strTATItems = mobjLisInsideComm.GetTatTimeShow(1, strItem, strDept, "", "", strSex, intMsg, strShowBef, , UserInfo.����)
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
        
        '��ȡ������Ŀ������
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
                strMsgShowStop = strMsgShowStop & Split(var_Tmp(i), ",")(1) & "δ����,�����ͼ�" & vbCrLf
            ElseIf Split(var_Tmp(i), ",")(7) <> 0 And Split(var_Tmp(i), ",")(9) = 2 And Split(var_Tmp(i), ",")(2) <> "" Then
                '��ʱ����ֹ��
                strMsgShowStop = strMsgShowStop & Replace(Replace(Split(var_Tmp(i), ",")(8), "[��Ŀ]", Split(var_Tmp(i), ",")(1)), "[��ʱ]", Split(var_Tmp(i), ",")(7) & "����") & vbCrLf
            Else
                '��ͬ��Ŀͬ�����ʱ��,����һ����Ŀ��ʱ,�����и��������Ŀ�������ͼ�
                If InStr(strItemCode, "," & Split(var_Tmp(i), ",")(6)) <= 0 Then
                    strIDs = strIDs & "," & Split(var_Tmp(i), ",")(4) & "," & Split(var_Tmp(i), ",")(5)
                End If
            End If
        Next
        
        If strIDs <> "" Then
            strIDs = Mid(strIDs, 2)
        End If
        
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
                        strIDs = strIDs & "," & Split(var_Tmp(i), ",")(4) & "," & Split(var_Tmp(i), ",")(5)
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



