VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlIDKind.ocx"
Begin VB.Form frmArchiveViewAndPrint 
   BackColor       =   &H80000005&
   Caption         =   "������ѯ��ӡ"
   ClientHeight    =   11460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16485
   Icon            =   "frmArchiveViewAndPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11460
   ScaleWidth      =   16485
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ProgressBar psb 
      Height          =   225
      Left            =   10680
      TabIndex        =   70
      Top             =   11160
      Visible         =   0   'False
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   480
      ScaleHeight     =   8055
      ScaleWidth      =   3975
      TabIndex        =   9
      Top             =   1080
      Width           =   3975
      Begin VB.CheckBox chkFilter 
         Height          =   255
         Left            =   3360
         Picture         =   "frmArchiveViewAndPrint.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         ToolTipText     =   "���ղ��������Բ��˽��й�����ʾ"
         Top             =   1133
         Width           =   270
      End
      Begin VB.Frame fraType 
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   4815
         Begin VB.OptionButton optType 
            Caption         =   "סԺ"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   22
            Top             =   120
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optType 
            Caption         =   "����"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   21
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lblType 
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Left            =   60
            TabIndex        =   23
            Top             =   150
            Width           =   450
         End
      End
      Begin VB.ComboBox cboDept 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   960
         TabIndex        =   19
         Text            =   "cboDept"
         Top             =   600
         Width           =   2655
      End
      Begin VB.PictureBox picPatiIn 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   0
         ScaleHeight     =   3495
         ScaleWidth      =   4005
         TabIndex        =   10
         Top             =   3840
         Width           =   4005
         Begin XtremeReportControl.ReportControl rptPati 
            Height          =   1620
            Left            =   480
            TabIndex        =   11
            Top             =   1800
            Width           =   2280
            _Version        =   589884
            _ExtentX        =   4022
            _ExtentY        =   2857
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
            Index           =   1
            Left            =   0
            ScaleHeight     =   315
            ScaleWidth      =   3855
            TabIndex        =   24
            Top             =   0
            Visible         =   0   'False
            Width           =   3855
            Begin VB.ComboBox cboSelectTime 
               Height          =   300
               Index           =   1
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   0
               Width           =   1230
            End
            Begin VB.Label lblSeeTim 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ʱ��"
               Height          =   180
               Left            =   120
               TabIndex        =   26
               Top             =   45
               Width           =   720
            End
         End
         Begin VB.PictureBox picPara 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   320
            Index           =   0
            Left            =   120
            ScaleHeight     =   315
            ScaleWidth      =   3855
            TabIndex        =   12
            Top             =   480
            Visible         =   0   'False
            Width           =   3855
            Begin VB.ComboBox cboSelectTime 
               Height          =   300
               Index           =   0
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   0
               Width           =   1230
            End
            Begin VB.Label lbl��Ժʱ�� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժʱ��"
               Height          =   180
               Left            =   0
               TabIndex        =   14
               Top             =   60
               Width           =   720
            End
         End
      End
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   300
         Left            =   960
         TabIndex        =   15
         Top             =   1110
         Width           =   2295
         _ExtentX        =   4683
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmArchiveViewAndPrint.frx":D0A4
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
         InputAppearance =   0
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
      Begin XtremeSuiteControls.TabControl tbcPati 
         Height          =   1335
         Left            =   360
         TabIndex        =   18
         Top             =   2400
         Width           =   3015
         _Version        =   589884
         _ExtentX        =   5318
         _ExtentY        =   2355
         _StockProps     =   64
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&D)��"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   660
         Width           =   810
      End
      Begin VB.Label lblFind 
         BackStyle       =   0  'Transparent
         Caption         =   "����(F3)"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1140
         Width           =   735
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   4800
      ScaleHeight     =   7935
      ScaleWidth      =   4935
      TabIndex        =   6
      Top             =   2640
      Width           =   4935
      Begin VB.ComboBox cboVisit 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   0
         Width           =   3765
      End
      Begin MSComctlLib.TreeView tvwArchive 
         Height          =   1785
         Left            =   960
         TabIndex        =   8
         Top             =   3600
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   3149
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ils16"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.TabControl tbcHistory 
         Height          =   2895
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   3255
         _Version        =   589884
         _ExtentX        =   5741
         _ExtentY        =   5106
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   11790
      ScaleHeight     =   285
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   2670
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picRpt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   11625
      ScaleHeight     =   780
      ScaleWidth      =   915
      TabIndex        =   2
      Top             =   3255
      Width           =   915
      Begin SHDocVwCtl.WebBrowser webRpt 
         Height          =   450
         Left            =   135
         TabIndex        =   3
         Top             =   150
         Width           =   450
         ExtentX         =   794
         ExtentY         =   794
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.Frame fraPati 
      BackColor       =   &H80000005&
      Caption         =   "������Ϣ"
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   11565
      Begin VB.Frame fraOutPati 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   1560
         TabIndex        =   48
         Top             =   360
         Width           =   9375
         Begin VB.Label lbl�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   8880
            TabIndex        =   67
            Top             =   0
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   11
            Left            =   2580
            TabIndex        =   66
            Top             =   0
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "�Ա�:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   12
            Left            =   4860
            TabIndex        =   65
            Top             =   0
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "�����:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   13
            Left            =   180
            TabIndex        =   64
            Top             =   0
            Width           =   630
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   14
            Left            =   7140
            TabIndex        =   63
            Top             =   0
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "���֤��:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   15
            Left            =   2220
            TabIndex        =   62
            Top             =   360
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��������:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   17
            Left            =   4500
            TabIndex        =   61
            Top             =   720
            Width           =   810
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����׿��"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   11
            Left            =   3060
            TabIndex        =   60
            Top             =   0
            Width           =   720
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "28��"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   12
            Left            =   7620
            TabIndex        =   59
            Top             =   0
            Width           =   360
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Ů"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   13
            Left            =   5340
            TabIndex        =   58
            Top             =   0
            Width           =   180
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "20150101"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   14
            Left            =   825
            TabIndex        =   57
            Top             =   0
            Width           =   720
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "500101198810121245"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   15
            Left            =   3060
            TabIndex        =   56
            Top             =   360
            Width           =   1620
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "2015-10-11"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   17
            Left            =   5340
            TabIndex        =   55
            Top             =   720
            Width           =   900
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��ַ:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   18
            Left            =   4860
            TabIndex        =   54
            Top             =   360
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��������:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   19
            Left            =   0
            TabIndex        =   53
            Top             =   360
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����ҽʦ:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   0
            TabIndex        =   52
            Top             =   720
            Width           =   810
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��������������"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   16
            Left            =   5340
            TabIndex        =   51
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����ӱ"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   825
            TabIndex        =   50
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "2015-10-11"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   18
            Left            =   825
            TabIndex        =   49
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame fraInPati 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   1560
         TabIndex        =   27
         Top             =   360
         Width           =   9375
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   2580
            TabIndex        =   47
            Top             =   0
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "�Ա�:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   2
            Left            =   4860
            TabIndex        =   46
            Top             =   0
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "סԺ��:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   4
            Left            =   180
            TabIndex        =   45
            Top             =   0
            Width           =   630
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   7140
            TabIndex        =   44
            Top             =   0
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "���֤��:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   3
            Left            =   2220
            TabIndex        =   43
            Top             =   360
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   5
            Left            =   2580
            TabIndex        =   42
            Top             =   720
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��Ժ:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   6
            Left            =   4860
            TabIndex        =   41
            Top             =   720
            Width           =   450
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����׿��"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   3060
            TabIndex        =   40
            Top             =   0
            Width           =   720
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "28��"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   2
            Left            =   7620
            TabIndex        =   39
            Top             =   0
            Width           =   360
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Ů"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   5340
            TabIndex        =   38
            Top             =   0
            Width           =   180
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "20150101"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   4
            Left            =   825
            TabIndex        =   37
            Top             =   0
            Width           =   720
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "#"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   3
            Left            =   3060
            TabIndex        =   36
            Top             =   360
            Width           =   90
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "#"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   5
            Left            =   3060
            TabIndex        =   35
            Top             =   720
            Width           =   90
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "2015-10-11"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   6
            Left            =   5340
            TabIndex        =   34
            Top             =   720
            Width           =   900
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��ַ:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   8
            Left            =   4860
            TabIndex        =   33
            Top             =   360
            Width           =   450
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��������:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   10
            Left            =   0
            TabIndex        =   32
            Top             =   360
            Width           =   810
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "סԺҽʦ:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   9
            Left            =   0
            TabIndex        =   31
            Top             =   720
            Width           =   810
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��������������"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   8
            Left            =   5340
            TabIndex        =   30
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����ӱ"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   9
            Left            =   825
            TabIndex        =   29
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "2015-10-11"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   10
            Left            =   825
            TabIndex        =   28
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Image imgPatient 
         Height          =   1185
         Left            =   120
         Picture         =   "frmArchiveViewAndPrint.frx":D187
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   11100
      Width           =   16485
      _ExtentX        =   29078
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16060
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "�༭"
            TextSave        =   "�༭"
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
   Begin XtremeSuiteControls.TabControl tbcArchive 
      Height          =   2595
      Left            =   10920
      TabIndex        =   5
      Top             =   3720
      Width           =   4125
      _Version        =   589884
      _ExtentX        =   7276
      _ExtentY        =   4577
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2040
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":E051
            Key             =   "סԺ"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":148B3
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":14E4D
            Key             =   "home"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":1B6AF
            Key             =   "object_report"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":1BC49
            Key             =   "object_case"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":1C1E3
            Key             =   "object_tend"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":1C77D
            Key             =   "object_first"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":1CD17
            Key             =   "object_advice"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":1D2B1
            Key             =   "object_file"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":1D84B
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":240AD
            Key             =   "Path"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgPati 
      Left            =   3000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":24647
            Key             =   "Boy"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":2AEA9
            Key             =   "Girl"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":3170B
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":31CA5
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":3223F
            Key             =   "print"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3720
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":38AA1
            Key             =   "��ҳ����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":3BADB
            Key             =   "Boy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":4233D
            Key             =   "��鱨��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":47DFF
            Key             =   "���鱨��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":4D8C1
            Key             =   "Girl"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":54123
            Key             =   "Patient"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":54FFD
            Key             =   "unCheckAll"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":55597
            Key             =   "CheckAll"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":55B31
            Key             =   "סԺ����"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":58B6B
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":5BBA5
            Key             =   "����֤��"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":5EBDF
            Key             =   "��ҳ��ҳһ"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":5FAB9
            Key             =   "�ٴ�·��"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":62AF3
            Key             =   "��ҳ��ҳ��"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":65B2D
            Key             =   "������"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":68B67
            Key             =   "סԺҽ��"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":6BBA1
            Key             =   "�����¼"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":6EBDB
            Key             =   "֪���ļ�"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":71C15
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":7224F
            Key             =   "CheckFill"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":72889
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":72EC3
            Key             =   "��ҳ����"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":75EFD
            Key             =   "down"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":7C75F
            Key             =   "up"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveViewAndPrint.frx":82FC1
            Key             =   "סԺ֤"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmArchiveViewAndPrint.frx":86453
      Left            =   600
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmArchiveViewAndPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'API����
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Enum INPATIREPORT_COLUMN
    col_ѡ�� = 0
    col_ͼ�� = 1
    col_��ӡͼ�� = 2
    col_�Ƿ��Ŀ = 3
    col_��Ŀ���� = 4
    col_סԺ�� = 5
    COL_NO = 6             '����
    COL_����� = 7          '����
    col_���� = 8
    col_���� = 9
    col_�Ա� = 10
    col_���� = 11
    col_���֤�� = 12
    col_�������� = 13
    COL_ִ��ʱ�� = 14        '����
    col_��Ժ���� = 15
    col_��Ժ���� = 16
    col_סԺҽʦ = 17       '�������ҽ��
    col_��ͥ��ַ = 18
    col_���￨�� = 19
    col_���ۺ� = 20
    '������
    col_�������� = 21
    col_����Id = col_�������� + 1        '����
    col_��ҳID = col_�������� + 2         '���� ����Ϊ�Һ�ID
    col_����ID = col_�������� + 3       '����
    COL_����ת�� = col_�������� + 4     '����ת��
    col_��ӡ��¼ = col_�������� + 5
End Enum

Private Enum PATI_INFO
    lbl_���� = 0
    lbl_�Ա� = 1
    lbl_���� = 2
    lbl_���֤�� = 3
    lbl_סԺ�� = 4
    lbl_���� = 5
    lbl_��Ժ���� = 6
    lbl_��ͥ��ַ = 8
    lbl_סԺҽʦ = 9
    lbl_�������� = 10
    
    lblOUT_���� = 11
    lblOUT_���� = 12
    lblOUT_�Ա� = 13
    lblOUT_����� = 14
    lblOUT_���֤�� = 15
    lblOUT_��ͥ��ַ = 16
    lblOUT_�������� = 17
    lblOUT_�������� = 18
    lblOUT_����ҽʦ = 7
End Enum

Private Type PatiInfo
    ״̬ As Integer '������ҳ.״̬
    Ӥ�� As Integer
    סԺ�� As String
    ���� As String
    ��ҳID As Long
    ����ID As Long
    ����ID As Long
    ����ID As Long
    ��Ժ���� As Date
    ��Ժ���� As Date
    ��Ŀ���� As Date
    סԺ���� As Long
    ����ת�� As Boolean
End Type

Private Enum E_TYPE
    E_סԺ = 0
    E_���� = 1
End Enum
'����
Private Const M_CON_CATE As String = "R11,R12,R1,R2,R3,R4,R5,R6,R7,R8,R9,R10"
'���￨������󷵻صĿ��õ�ҽ�ƿ�
Private Const mstrCardKindOut         As String = "��|���￨|0|0|8|0|0|0;��|�����|0|0|0|0|0|0;��|�Һŵ�|0|0|0|0|0|0;��|����|0|0|0|0|0|0;��|�������֤|0|0|0|0|0|0;�ɣ�|�ɣÿ�|1|0|0|0|0|0"
'סԺ��������󷵻صĿ��õ�ҽ�ƿ�
Private Const mstrCardKindIN          As String = "��|���￨|0|0|8|0|0|0;ס|סԺ��|0|0|0|0|0|0;��|����|0|0|0|0|0|0;��|����|0|0|0|0|0|0;��|�������֤|0|0|0|0|0|0;��|���ۺ�|0|0|0|0|0|0"
'ֱ�Ӳ��ҿ�������󷵻صĿ��õ�ҽ�ƿ�
Private Const mstrCardKindFind        As String = "��|���￨|0|0|8|0|0|0;��|�����|0|0|0|0|0|0;ס|סԺ��|0|0|0|0|0|0;��|���ݺ�|0|0|0|0|0|0;��|����|0|0|0|0|0|0;��|�������֤|0|0|0|0|0|0;�ɣ�|�ɣÿ�|1|0|0|0|0|0;ҽ|ҽ����|0|0|0|0|0|0"

'�¼�
Private WithEvents mclsTendsNew     As zl9TendFile.clsTendFile    '�°滤ʿ����վ
Attribute mclsTendsNew.VB_VarHelpID = -1
Private WithEvents mobjReport       As zl9Report.clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private WithEvents mclsDockAduits   As zlRichEPR.clsDockAduits
Attribute mclsDockAduits.VB_VarHelpID = -1
'
Private mclsOutAdvices          As zlPublicAdvice.clsDockOutAdvices
Private mclsInAdvices           As zlPublicAdvice.clsDockInAdvices
Private mclsPath                As zlPublicPath.clsDockPath
Private mclsArchive             As zlMedRecPage.clsArchive '���Ӳ������Ĵ�����
Private mobjRichEMR             As Object
Private mobjPublicPACS          As Object
Private mobjSquareCard          As Object      '���������
Private mobjReportForm          As Object      '����Ԥ������
Private mobjPatient             As Object
Private mobjInfection           As Object      '��Ⱦ�����濨,ģ��Ԥ������
'
Private mstr�Һŵ�              As String
Private mstrPrivs               As String

Private mstr���鱨���ӡ        As String        '0-�ϰ�LIS�������;1-�°�LIS����ʽ
Private mstr�����Ӧ����        As String
Private mstr����Ӧ����        As String
Private mstrFindType           As String       '�����洢��ǰ�������͵�����
Private mstrPrintDocIDs        As String       '�����������ĵ�ֻ��ӡһ��
Private mstrTempDel            As String        'ɾ����ʱ�ļ�
Private mstrPrintMedRec        As String        '��¼�Ѿ���ӡ����

'PDF��ӡ
'Public gstrInputSeverName As String
'Public gstrInputUser As String
'Public gstrInputPwd As String
'
''
Private mlng����ID      As Long
Private mlng����ID      As Long '���˵�ǰ�������ľ���ID������Ϊ�Һ�ID,סԺ����ҳID
Private mlng����ID      As Long
Private mlng����ID      As Long
Private mlngPreDept     As Long

Private mintPatiCount   As Integer   '��ѡ������Ŀ
Private mintPreDept     As Integer
Private mintDeptView    As Integer '0-��������ʾ��1-��������ʾ
Private mintDeptViewBed As Integer '0��1-ֻ��ʾ�д�λ�Ĳ������߿���
Private mintMecStandard As Integer  '������ҳ��ʽ 0-��������׼��1-�Ĵ�ʡ��׼��2-����ʡ��׼,3-����ʡ��׼
Private mintFindType    As Integer '0-סԺ��,1-����,2-���￨,3-����
Private mintOutPreTime  As Integer

'
Private mbytType        As Byte             '0-סԺ;1-����
Private mbytPDFStatu    As Byte             '0-δ��ʼ��;1-��ʼ���ɹ�
Private mbytPrintType   As Byte             '1-��ӡ��ҳ
'
Private mblnLIS         As Boolean         '�Ƿ����°�LIS
Private mblnOutDept     As Boolean '�Ƿ������������Ŀ��ң��������۲�����ʾ����ţ�
Private mblnMoved       As Boolean '��ǰ���������Ƿ�ת��
Private mblnNewTends    As Boolean 'T-�°滤���¼;F-�ϰ滤���¼
Private mblnICU         As Boolean '�Ƿ�Ǳ��Ƶ�ICU��
Private mblnUndo        As Boolean
Private mblnTabTmp      As Boolean
Private mblnTvwTmp      As Boolean
Private mblnSeePic      As Boolean           'T-��ʾ��Ƭ;F-���ع�Ƭ
Private mblnPrint        As Boolean           'T-�����¼������ӡ
'
Private mcolSubForm     As Collection
Private mcolReport      As Collection
Private mcolPrint       As Collection       '�����ӡ��
'


Private mdatOutBegin As Date, mdatOutEnd As Date    '��Ժָ��ʱ��
Private mDatBegin As Date, mDatEnd As Date          '����ָ��ʱ��

Private mrsPati         As ADODB.Recordset '������Ϣ���ϣ�����ͬһ���֤�ŵ����в���
Private mrsData         As ADODB.Recordset
Private mrsMedRec       As ADODB.Recordset
Private mblnReturn      As Boolean

Public Sub ShowArchive(ByVal frmParent As Object, ByVal strPrivs As String)
'���ܣ������ӿڷ��������� ShowMe����
'����:frmParent-������
'     strPrivs-ģ��Ȩ��
    mstrPrivs = strPrivs

    Me.Show 0, frmParent  ' ����Ϊ��ģʽ;��ģʽ������� �ٴ�·�������Ʊ��桢���µ��Ӵ������ʱ������
End Sub

Private Sub InitBasicData()
'���ܣ���ʼ��һЩ�������ݣ��������б���ص�
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim strSQL As String
    Dim objTab As TabControlItem
    Dim strTmp As String
    Dim str����IDs As String
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim strErr As String
    Dim blnTmp As Boolean
    Dim str���֤�� As String
    Dim strOrder As String
    
    Screen.MousePointer = 11
    LockWindowUpdate Me.hwnd
    mstr�Һŵ� = "": mlngPreDept = -1

    Call tbcHistory.RemoveAll
    Call cboVisit.Clear
    
    If mlng����ID = 0 Then
        Set objTab = tbcHistory.InsertItem(tbcHistory.ItemCount, "", tvwArchive.hwnd, 0)
        If mbytType = E_���� Then
            Call ShowOutPatiInfo
        Else
            Call ShowInPatiInfo
        End If
        Call ShowArchiveTree
    Else
        On Error GoTo errH
        strSQL = "select a.���֤�� from ������Ϣ a where a.����id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        strTmp = rsTmp!���֤�� & ""
        If strTmp <> "" Then
            '��֤���֤�ŵĺϷ���
            If mobjPatient Is Nothing Then
                On Error Resume Next
                Set mobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
                err.Clear: On Error GoTo 0
            End If
            If mobjPatient Is Nothing Then
                MsgBox "����������Ϣ����������zlPublicPatient.clsPublicPatient��ʧ�ܣ�", vbInformation, Me.Caption
            Else
                Call mobjPatient.zlInitCommon(gcnOracle, glngSys, UserInfo.�û���)
                If mobjPatient.CheckPatiIdcard(strTmp) Then
                    str���֤�� = strTmp
                End If
            End If
        End If
        
        On Error GoTo errH
        If chkFilter.Value = vbChecked Then
            strOrder = "��ʼʱ�� Desc"
        Else
            If mbytType = E_סԺ Then
                strOrder = "���� Desc,��ʼʱ�� Desc"
            Else
                strOrder = "���� ASC,��ʼʱ�� Desc"
            End If
        End If
        If str���֤�� <> "" Then
            strSQL = "select a.����id from ������Ϣ a where a.����id<>[1] and a.���֤��=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, str���֤��)
            Do While Not rsTmp.EOF
                str����IDs = str����IDs & "," & rsTmp!����ID
                rsTmp.MoveNext
            Loop
        End If
        If str����IDs = "" Then
            strSQL = " Select ����id,ID as ����ID,NO,����ʱ�� as ��ʼʱ��,Null as ����ʱ��,ִ�в���ID as ����ID,0 as ����ת��,-1 as ��������,null as �����,1 as ���� From ���˹Һż�¼ Where ����ID=[1] And ��¼����=1 And ��¼״̬=1" & _
                " Union ALL" & _
                " Select ����id,ID as ����ID,NO,����ʱ�� as ��ʼʱ��,Null as ����ʱ��,ִ�в���ID as ����ID,1 as ����ת��,-1 as ��������,null as �����,1 as ���� From H���˹Һż�¼ Where ����ID=[1] And ��¼����=1 And ��¼״̬=1" & _
                " Union ALL" & _
                " Select ����id,��ҳID as ����ID,Null,��Ժ���� as ��ʼʱ��,��Ժ����,��Ժ����ID,����ת��,NVL(��������,0) as ��������,null as �����,2 as ���� From ������ҳ Where ����ID=[1] And Nvl(��ҳID,0)<>0"
            strSQL = "Select Rownum As ���,a.����ID,A.����ID,A.NO,A.��ʼʱ��,A.����ʱ��,B.���� as ����,A.����ת�� ,A.��������,a.����� From (" & strSQL & ") A,���ű� B Where A.����ID=B.ID Order by " & strOrder
            Set mrsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        Else
            str����IDs = mlng����ID & str����IDs
            strTmp = " ����ID In (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X) "
            strSQL = " Select ����id,ID as ����ID,NO,����ʱ�� as ��ʼʱ��,Null as ����ʱ��,ִ�в���ID as ����ID,0 as ����ת��,-1 as ��������,null as �����,1 as ���� From ���˹Һż�¼ Where " & strTmp & " And ��¼����=1 And ��¼״̬=1 and NO is not null" & _
                " Union ALL" & _
                " Select ����id,ID as ����ID,NO,����ʱ�� as ��ʼʱ��,Null as ����ʱ��,ִ�в���ID as ����ID,1 as ����ת��,-1 as ��������,null as �����,1 as ���� From H���˹Һż�¼ Where " & strTmp & " And ��¼����=1 And ��¼״̬=1 and NO is not null" & _
                " Union ALL" & _
                " Select ����id,��ҳID as ����ID,Null,��Ժ���� as ��ʼʱ��,��Ժ����,��Ժ����ID,����ת��,NVL(��������,0) as ��������,סԺ�� as �����,2 as ���� From ������ҳ Where " & strTmp & " And Nvl(��ҳID,0)<>0"
            strSQL = "Select Rownum As ���,a.����ID,A.����ID,A.NO,A.��ʼʱ��,A.����ʱ��,B.���� as ����,A.����ת�� ,A.��������,a.����� From (" & strSQL & ") A,���ű� B Where A.����ID=B.ID  Order by " & strOrder
            Set mrsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����IDs)
        End If
        Do While Not mrsData.EOF
            strTmp = IIf(IsNull(mrsData!NO), "��" & mrsData!����id & "��" & IIf(mrsData!�������� = 1, "��������", IIf(mrsData!�������� = 2, "סԺ����", "סԺ")), "�������") & ":" & mrsData!���� & "," & Format(mrsData!��ʼʱ��, "yyyy-MM-dd HH:mm") & _
            IIf(Not IsNull(mrsData!����ʱ��), "��" & Format(mrsData!����ʱ��, "yyyy-MM-dd HH:mm"), "")
        
            If mrsData.AbsolutePosition = 1 Then
                Set objTab = tbcHistory.InsertItem(tbcHistory.ItemCount, strTmp, tvwArchive.hwnd, IIf(IsNull(mrsData!NO), 0, 1))
            End If
             
            cboVisit.AddItem strTmp
            cboVisit.ItemData(cboVisit.NewIndex) = Val(mrsData!���)
            mrsData.MoveNext
        Loop
        If cboVisit.ListCount > 0 Then
            Call Cbo.SetIndex(cboVisit.hwnd, 0)
            Call cboVisit_Click
        End If
    End If
    LockWindowUpdate 0
    Screen.MousePointer = 0
    Exit Sub
errH:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox
    Dim objMenu As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim arrTmp  As Variant
    Dim strFunc As String
    Dim strTmp  As String
    Dim i As Long
    
    '������----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False) '����
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Dept, "������ʾ(&D)") '����
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Dept * 10# + 1, "��������ʾ(&D)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Dept * 10# + 2, "��������ʾ(&U)", -1, False)
        End With
    End With
        
    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_FilePopup, "�ļ�")
        objPopup.IconId = conMenu_File_Open
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet * 10# + 1, "��ӡ����")
            Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet * 10# + 2, "��������")
                objControl.IconId = conMenu_File_Parameter
            Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet * 10# + 3, "PDFλ��")
                objControl.ToolTipText = "����PDF���λ��"
                objControl.IconId = conMenu_File_PDF
        End With
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_File_PDF, "PDF")
        Set objControl = .Add(xtpControlButton, conMenu_Img_Look, "��Ƭ")
        objControl.IconId = conMenu_Edit_MarkMap: objControl.BeginGroup = True
        '��չ����
        If CreatePlugInOK(P�������Ĵ�ӡ) Then
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Tool_PlugIn, "��չ����", objControl.Index + 1)
            objPopup.ID = conMenu_Tool_PlugIn
            objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                On Error Resume Next
                strFunc = gobjPlugIn.GetFuncNames(glngSys, P�������Ĵ�ӡ)
                Call zlPlugInErrH(err, "GetFuncNames")
                err.Clear: On Error GoTo 0
                If strFunc <> "" Then
                    arrTmp = Split(strFunc, ",")
                    strTmp = Replace(strFunc, "Auto:", "")
                    arrTmp = Split(strTmp, ",")
                    For i = 0 To UBound(arrTmp)
                        Set objControl = .Add(xtpControlButton, conMenu_Tool_PlugIn_Item + i + 1, CStr(arrTmp(i)))
                        If i <= 9 Then objControl.Caption = objControl.Caption & "(&" & IIf(i = 9, 0, i + 1) & ")"
                        objControl.IconId = conMenu_Tool_PlugIn_Item
                        objControl.Parameter = arrTmp(i)
                    Next
                Else
                    objPopup.Visible = False
                End If
            End With
        End If
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): objControl.BeginGroup = True
    End With
    
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add FALT, vbKeyS, conMenu_Edit_Save
    End With

End Sub

Private Sub InitReportColumn()
'����:bytFunc=0 סԺ;bytFunc=1 ����
    Dim objCol As ReportColumn, lngidx As Long, i As Long

    With rptPati
        .Columns.DeleteAll
        
        Set objCol = .Columns.Add(col_ѡ��, "", 20, False)
            objCol.Icon = imgPati.ListImages("UnCheck").Index - 1
            objCol.EditOptions.AllowEdit = True
            objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_ͼ��, "", 20, False)  'ͼ��
            objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_��ӡͼ��, "", 20, False)  'col_��ӡͼ��
            objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_�Ƿ��Ŀ, "�Ƿ��Ŀ", 60, True)
        Set objCol = .Columns.Add(col_��Ŀ����, "��Ŀ����", 80, True)
        Set objCol = .Columns.Add(col_סԺ��, "סԺ��", 80, True)
        Set objCol = .Columns.Add(COL_NO, "NO", 80, True)
        Set objCol = .Columns.Add(COL_�����, "�����", 80, True)
        Set objCol = .Columns.Add(col_����, "����", 50, True)
        Set objCol = .Columns.Add(col_����, "����", 80, True)
        Set objCol = .Columns.Add(col_�Ա�, "�Ա�", 45, True)
        Set objCol = .Columns.Add(col_����, "����", 45, True)
        Set objCol = .Columns.Add(col_���֤��, "���֤��", 150, True)
        Set objCol = .Columns.Add(col_��������, "��������", 80, True)
        Set objCol = .Columns.Add(COL_ִ��ʱ��, "ִ��ʱ��", 80, True)
        Set objCol = .Columns.Add(col_��Ժ����, "��Ժ����", 80, True)
        Set objCol = .Columns.Add(col_��Ժ����, "��Ժ����", 80, True)
        Set objCol = .Columns.Add(col_סԺҽʦ, "סԺҽʦ", 80, True)
        Set objCol = .Columns.Add(col_��ͥ��ַ, "��ַ", 150, True)
        If ISPassShowCard Then
            Set objCol = .Columns.Add(col_���￨��, "���￨��", 0, False)
        Else
            Set objCol = .Columns.Add(col_���￨��, "���￨��", 70, True)
        End If
        Set objCol = .Columns.Add(col_���ۺ�, "���ۺ�", 62, True)
        '������
        Set objCol = .Columns.Add(col_��������, "��������", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_����Id, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_��ҳID, "��ҳID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_����ID, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_����ת��, "����ת��", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_��ӡ��¼, "��ӡ��¼", 0, False): objCol.Visible = False
        Call ShowReportColumn
        
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
        mblnUndo = True
        .MultipleSelection = False '������SelectionChanged�¼�
         mblnUndo = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgPati
        
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
        If mbytType = 0 Then
            .SortOrder.Add .Columns(col_��Ժ����)
            .SortOrder(0).SortAscending = True
        Else
            .SortOrder.Add .Columns(COL_ִ��ʱ��)
            .SortOrder(0).SortAscending = True
        End If
    End With
End Sub

Private Sub ShowReportColumn()
    With rptPati.Columns
    If tbcPati.Selected.Tag = "��Ժ" Then
        .Find(col_�Ƿ��Ŀ).Visible = True
        .Find(col_��Ŀ����).Visible = True
        .Find(col_��ӡͼ��).Visible = True
        .Find(col_סԺ��).Visible = True
        .Find(col_��Ժ����).Visible = True
        .Find(col_��Ժ����).Visible = True
        .Find(col_����).Visible = True
        .Find(col_���ۺ�).Visible = True
        
        .Find(COL_NO).Visible = False
        .Find(COL_�����).Visible = False
        .Find(COL_ִ��ʱ��).Visible = False
        .Find(col_סԺҽʦ).Caption = "סԺҽʦ"
    ElseIf tbcPati.Selected.Tag = "��Ժ" Then
        .Find(col_�Ƿ��Ŀ).Visible = False
        .Find(col_��Ŀ����).Visible = False
        .Find(col_��ӡͼ��).Visible = False
        .Find(col_סԺ��).Visible = True
        .Find(col_��Ժ����).Visible = True
        .Find(col_��Ժ����).Visible = True
        .Find(col_����).Visible = True
        .Find(col_���ۺ�).Visible = True
        
        .Find(COL_NO).Visible = False
        .Find(COL_�����).Visible = False
        .Find(COL_ִ��ʱ��).Visible = False
        .Find(col_סԺҽʦ).Caption = "סԺҽʦ"
    ElseIf tbcPati.Selected.Tag = "����" Then
        .Find(col_�Ƿ��Ŀ).Visible = False
        .Find(col_��Ŀ����).Visible = False
        .Find(col_��ӡͼ��).Visible = False
        .Find(col_סԺ��).Visible = False
        .Find(col_��Ժ����).Visible = False
        .Find(col_��Ժ����).Visible = False
        .Find(col_����).Visible = False
        .Find(col_���ۺ�).Visible = False
         
        .Find(COL_NO).Visible = True
        .Find(COL_�����).Visible = True
        .Find(COL_ִ��ʱ��).Visible = True
        .Find(col_סԺҽʦ).Caption = "����ҽ��"
    ElseIf tbcPati.Selected.Tag = "����" Then
        .Find(col_�Ƿ��Ŀ).Visible = False
        .Find(col_��Ŀ����).Visible = False
        .Find(col_��ӡͼ��).Visible = False
        .Find(col_סԺ��).Visible = False
        .Find(col_��Ժ����).Visible = False
        .Find(col_��Ժ����).Visible = False
        .Find(col_����).Visible = False
        .Find(col_���ۺ�).Visible = False
        
        .Find(COL_NO).Visible = True
        .Find(COL_�����).Visible = True
        .Find(COL_ִ��ʱ��).Visible = True
        .Find(col_סԺҽʦ).Caption = "����ҽ��"
    End If
    End With
End Sub

Private Sub cboDept_Click()
'���ܣ�ˢ�½�������
'˵�����Ӹ��¼���ʼ�᲻�ظ�������ص����ݶ�ȡ
    Dim lng����ID As Long, i As Long, lngidx As Long
    Dim blnIn���� As Boolean, rsTmp As Recordset, str����IDs As String
    
    If cboDept.ListIndex = -1 Then
        Call ClearPatiInfo
        Exit Sub
    End If
    cboDept.Tag = cboDept.ListIndex
    mintPreDept = cboDept.ListIndex

    '���¶�ȡ����
    Call LoadPatients
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

Private Sub cboDept_Validate(Cancel As Boolean)
    If mblnReturn Then
        mblnReturn = False
    Else
        Call Cbo.SetIndex(cboDept.hwnd, Val(cboDept.Tag))
    End If
End Sub

Private Sub cboSelectTime_Click(Index As Integer)
'����:Index 0-��Ժ ;1-����
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    If Index = 0 Then
        intDateCount = cboSelectTime(Index).ItemData(cboSelectTime(Index).ListIndex)
        datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
        If cboSelectTime(Index).ListIndex = mintOutPreTime And intDateCount <> -1 Then Exit Sub
        If intDateCount = -1 Then
            If Not frmSelectTime.ShowMe(Me, mdatOutBegin, mdatOutEnd, cboSelectTime(Index)) Then
                'ȡ��ʱ�ָ�ԭ����ѡ��
                Call Cbo.SetIndex(cboSelectTime(Index).hwnd, mintOutPreTime)
                Exit Sub
            End If
        Else
            mdatOutEnd = datCurr
            mdatOutBegin = mdatOutEnd - intDateCount
        End If
        If mdatOutBegin = CDate(0) Or mdatOutEnd = CDate(0) Then
            cboSelectTime(Index).ToolTipText = ""
        Else
            cboSelectTime(Index).ToolTipText = "��Χ��" & Format(mdatOutBegin, "yyyy-MM-dd") & " �� " & Format(mdatOutEnd, "yyyy-MM-dd")
        End If
        mintOutPreTime = cboSelectTime(Index).ListIndex
    Else
        intDateCount = cboSelectTime(Index).ItemData(cboSelectTime(Index).ListIndex)
        datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        If Me.Visible Then
            If intDateCount = -1 Then
                If Not frmSelectTime.ShowMe(Me, mDatBegin, mDatEnd, cboSelectTime(Index)) Then
                    'ȡ��ʱ�ָ�ԭ����ѡ��
                    Call Cbo.SetIndex(cboSelectTime(Index).hwnd, mintOutPreTime)
                    Exit Sub
                End If
            ElseIf intDateCount = 0 Then
                '����  86114
                mDatBegin = Format(datCurr, "yyyy-MM-dd 00:00:00")
                mDatEnd = Format(datCurr, "yyyy-MM-dd 23:59:59")
            Else
                mDatEnd = Format(datCurr, "yyyy-MM-dd 23:59:59")
                mDatBegin = Format(mDatEnd - intDateCount, "yyyy-MM-dd 00:00:00")
            End If
        End If
        'ѡ����ʱ��֮������Һŵ�����
        cboSelectTime(Index).ToolTipText = Format(mDatBegin, "yyyy-MM-dd") & " - " & Format(mDatEnd, "yyyy-MM-dd")
        lblSeeTim.ToolTipText = cboSelectTime(Index).ToolTipText
        mintOutPreTime = cboSelectTime(Index).ListIndex
    End If
    If Me.Visible = True Then Call LoadPatients
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID

    Case conMenu_File_PrintSet * 10# + 1 '��ӡ����
        frmPrintSet.Show 1
    Case conMenu_File_PrintSet * 10# + 2 '��������
         Call SetPrintPara
    Case conMenu_File_PrintSet * 10# + 3
        Call SetPDFPath
    Case conMenu_File_Preview
        Call FuncPrintOrView(1) 'Ԥ��
    Case conMenu_File_Print
        If Control.Parameter = "DO" Then Exit Sub
        Control.Parameter = "DO"
        Control.Enabled = False
        Call FuncPrintOrView(2) '��ӡ
        Control.Parameter = ""
    Case conMenu_File_PDF
        If Control.Parameter = "DO" Then Exit Sub
        Control.Parameter = "DO"
        Control.Enabled = False
        Call FuncPrintOrView(3) 'PDF
        Control.Parameter = ""
    Case conMenu_File_Exit '�˳�
        Unload Me
    Case conMenu_View_Dept * 10# + 1, conMenu_View_Dept * 10# + 2 '������/������ʾ
        If mintDeptView <> Control.ID - conMenu_View_Dept * 10# - 1 Then
            mintDeptView = Control.ID - conMenu_View_Dept * 10# - 1
            Call LoadDept
        End If
    Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            Call gobjPlugIn.ExecuteFunc(glngSys, P�������Ĵ�ӡ, Control.Parameter, mlng����ID, mlng����ID, 0)
            Call zlPlugInErrH(err, "ExecuteFunc")
            err.Clear: On Error GoTo 0
        End If
    Case conMenu_Img_Look
        Call FuncLookPicture
    Case Else
    
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With Me.fraPati
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
    End With
    With Me.tbcArchive
        .Left = lngLeft: .Width = lngRight - lngLeft
        .Top = lngTop + fraPati.Height: .Height = lngBottom - lngTop - fraPati.Height
    End With
    Me.fraInPati.Width = fraPati.Width - imgPatient.Width - 500
    Me.fraOutPati.Width = fraPati.Width - imgPatient.Width - 500
    
    psb.Top = stbThis.Top + Screen.TwipsPerPixelY * 4
    psb.Left = stbThis.Panels(3).Left + Screen.TwipsPerPixelX * 2
    psb.Width = stbThis.Panels(3).Width - Screen.TwipsPerPixelX * 7
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    Case conMenu_FilePopup
        If InStr(";" & mstrPrivs & ";", ";��ӡ;") = 0 And InStr(";" & mstrPrivs & ";", ";��������;") = 0 And InStr(";" & mstrPrivs & ";", ";PDF;") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_File_PrintSet * 10# + 1, conMenu_File_Print
        If InStr(";" & mstrPrivs & ";", ";��ӡ;") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_File_PrintSet * 10# + 2
        If InStr(";" & mstrPrivs & ";", ";��������;") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_File_PrintSet * 10# + 3, conMenu_File_PDF   'PDFλ��  'PDF���
        If InStr(";" & mstrPrivs & ";", ";PDF;") = 0 Then
            Control.Visible = False
        End If
    Case conMenu_Img_Look
        Control.Visible = mblnSeePic
    Case conMenu_View_Dept * 10# + 1, conMenu_View_Dept * 10# + 2 '������/������ʾ
        Control.Checked = mintDeptView = Control.ID - conMenu_View_Dept * 10# - 1
    End Select
    If Control.Visible Then
        Select Case Control.ID
        Case conMenu_File_Print, conMenu_File_PDF, conMenu_File_Preview
            If cboVisit.ListIndex = -1 Or Control.Parameter = "DO" Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        End Select
    End If
     
    DeleteLISTempFile
End Sub

Private Sub chkFilter_Click()
    Dim i As Long
    
    PatiIdentify.Text = ""
    If PatiIdentify.Visible And PatiIdentify.Enabled Then PatiIdentify.SetFocus
    If chkFilter.Value = vbChecked Then
        cboDept.Enabled = False
        optType(0).Enabled = False
        optType(1).Enabled = False
        chkFilter.Tag = mbytType
        rptPati.Tag = ""
        rptPati.Records.DeleteAll
        rptPati.Populate
        cboSelectTime(0).Enabled = False
        cboSelectTime(1).Enabled = False
        
        On Error Resume Next
        If mobjSquareCard Is Nothing Then Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        err.Clear: On Error GoTo 0
        If mobjSquareCard.zlInitComponents(Me, P�������Ĵ�ӡ, glngSys, UserInfo.�û���, gcnOracle, False) = False Then
            Set mobjSquareCard = Nothing
            MsgBox "ҽ�ƿ�������zl9CardSquare����ʼ��ʧ��!", vbInformation, gstrSysName
        End If
        Call PatiIdentify.zlInit(Me, glngSys, P�������Ĵ�ӡ, gcnOracle, gstrDBUser, mobjSquareCard, mstrCardKindFind, "zl9CISJob")
        PatiIdentify.objIDKind.AllowAutoICCard = True
        PatiIdentify.objIDKind.AllowAutoIDCard = True
        chkFilter.ToolTipText = "��ˢ��������[-����ID]��[+סԺ��]��[*�����]�ȷ�ʽ��ȡ���˵���Ϣ��"
        For i = 0 To tbcPati.ItemCount - 1
            If Not tbcPati.Item(i).Selected And tbcPati.Item(i).Visible Then
                tbcPati.Item(i).Visible = False
            End If
        Next
        Call ClearPatiInfo
        Call InitBasicData
    Else
        cboDept.Enabled = True
        optType(0).Enabled = True
        optType(1).Enabled = True
        cboSelectTime(0).Enabled = True
        cboSelectTime(1).Enabled = True
        chkFilter.ToolTipText = "���ղ��������Բ��˽��й�����ʾ"
        rptPati.Tag = ""
        rptPati.Records.DeleteAll
        Call optType_Click(Val(chkFilter.Tag))
    End If
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picPati.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picLeft.hwnd
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
    PatiIdentify.ActiveFastKey '����
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
    Dim objTab As TabControlItem
    Dim frmTendBody As Object
    Dim intIdx As Integer
    Dim intType As Integer
    Dim objPane As XtremeDockingPane.Pane
    
    '����ָ���Ĭ���������Ͷ�ȡ
    '-----------------------------------------------------
    mintPatiCount = 0
    DeleteLISTempFile
    mintFindType = Val(zlDatabase.GetPara("���˲��ҷ�ʽ", glngSys, P�������Ĵ�ӡ, , , , intType))
    mintDeptViewBed = Val(zlDatabase.GetPara("����ʾ�޴�λ�Ĳ�������", glngSys, P�������Ĵ�ӡ, , , , intType))
    mintMecStandard = Val(zlDatabase.GetPara("������ҳ��׼", glngSys, pסԺҽ��վ, "0"))
    chkFilter.Value = Val(zlDatabase.GetPara("������ʾģʽ", glngSys, P�������Ĵ�ӡ, , , , intType))
    mblnLIS = Sys.IsSysSetUp(2500)
    mstrPrintDocIDs = ""
    Call InitCommandBar
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, 320, 400, DockLeftOf, Nothing)
    objPane.Title = "�����б�"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    Set objPane = Me.dkpMain.CreatePane(2, 360, 400, DockRightOf, objPane)
    objPane.Title = "�ļ��б�"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    '��������Ӳ�������RichEditor��ͻ,�������PDF�����ļ�,������Ӳ���ÿ�����ʱ��ѡ��·��������
    If zlCommFun.PDFInitialize() Then mbytPDFStatu = 1
    If Not gobjEmr Is Nothing Then
        If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
            Set gobjEmr = Nothing
        Else
            Set mobjRichEMR = DynamicCreate("zlRichEMR.clsDockContent", "�°没��", False)
            If Not mobjRichEMR Is Nothing Then Call mobjRichEMR.Init(gobjEmr, gcnOracle, glngSys, 0)
        End If
    End If
    If mclsArchive Is Nothing Then
        Set mclsArchive = New zlMedRecPage.clsArchive
        Call mclsArchive.InitArchiveMedRec(gcnOracle, glngSys)
    End If
    Set mclsOutAdvices = New clsDockOutAdvices
    Set mclsInAdvices = New clsDockInAdvices
    Set mclsDockAduits = New clsDockAduits
    Set mclsPath = New clsDockPath
    Set mclsTendsNew = New zl9TendFile.clsTendFile
    Call mclsTendsNew.InitTendFile(gcnOracle, glngSys)
    Set frmTendBody = mclsDockAduits.zlGetFormTendBody
    Call zlControl.FormSetCaption(frmTendBody, False, False)
    Call CreateObjectPacs(mobjPublicPACS)
    If mobjReport Is Nothing Then Set mobjReport = New zl9Report.clsReport
    '�Ӵ���
    '-----------------------------------------------------
    Set mcolSubForm = New Collection
    mcolSubForm.Add mclsArchive.zlGetForm(0), "_������ҳ"
    mcolSubForm.Add mclsArchive.zlGetForm(1), "_סԺ��ҳ"
    mcolSubForm.Add mclsDockAduits.zlGetFormEPR, "_������Ϣ"
    mcolSubForm.Add mclsOutAdvices.zlGetForm, "_����ҽ��"
    mcolSubForm.Add mclsInAdvices.zlGetForm, "_סԺҽ��"
    mcolSubForm.Add frmTendBody, "_���¼�¼��"
    mcolSubForm.Add mclsDockAduits.zlGetFormTendFile, "_�����¼��"
    mcolSubForm.Add mclsPath.zlGetForm, "_�ٴ�·��"
    mcolSubForm.Add mclsTendsNew.zlGetfrmInTendFile, "_�°滤��"
    If Not mobjRichEMR Is Nothing Then mcolSubForm.Add mobjRichEMR.zlGetForm, "_���Ӳ���"
    If Not mobjPublicPACS Is Nothing Then mcolSubForm.Add mobjPublicPACS.zlDocGetForm, "_��鱨��"
    With tbcArchive
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .Color = xtpTabColorOffice2003
            .Layout = xtpTabLayoutAutoSize
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        '��ʽ����Form_Load��ȡ���һ��ͼƬ��ʽ���л���ʱ�����������¼���
        Set objTab = .InsertItem(intIdx, "������ҳ", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "סԺ��ҳ", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "������Ϣ", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "����ҽ��", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "סԺҽ��", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "���¼�¼��", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "�����¼��", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "�ٴ�·��", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "�°滤��", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        If Not mobjRichEMR Is Nothing Then
            Set objTab = .InsertItem(intIdx, "���Ӳ���", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        End If
        Set objTab = .InsertItem(intIdx, "��鱨��", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "����", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "��������", picRpt.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        
    End With
    
    Call ClearPatiInfo
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
        .InsertItem(0, "��Ժ", picPatiIn.hwnd, 0).Tag = "��Ժ"
        .InsertItem(1, "��Ժ", picPatiIn.hwnd, 0).Tag = "��Ժ"
        .InsertItem(2, "����", picPatiIn.hwnd, 0).Tag = "����"
        .InsertItem(3, "����", picPatiIn.hwnd, 0).Tag = "����"
        .Item(1).Selected = True
        .Item(0).Selected = True
        '��λ����ѡ�
        tbcPati.Item(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(tbcPati), "tbcPati", 1)).Selected = True
    End With
    
    '������ʷ
    '-----------------------------------------------------
    With tbcHistory
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .Color = xtpTabColorOffice2003
            .DisableLunaColors = False
            .BoldSelected = True
            .HotTracking = True
            .ShowIcons = True
        End With
        .SetImageList ils16
    End With
    'RIS�ӿڴ���
    Call InitReportColumn
    Call HaveRIS(True)
    If mblnLIS Then Call InitObjLis(P�������Ĵ�ӡ)
    Call FuncLoadReport
    Call InitSelectTime
    'ȱʡ��λסԺ
    Call optType_Click(0)
    Call RestoreWinState(Me, App.ProductName, , True)
    Call ShowReportColumn
    If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Terminate()
     
    Call DeleteLISTempFile
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    'ǿ��Unload,��Ȼ���ἤ���Ӵ�����¼�
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next
    If picRpt.Tag <> "" Then mstrTempDel = picRpt.Tag
    Set mclsArchive = Nothing
    Set mclsDockAduits = Nothing
    Set mclsOutAdvices = Nothing
    Set mclsInAdvices = Nothing
    Set mclsPath = Nothing
    Set mclsTendsNew = Nothing
    Set mobjRichEMR = Nothing
    Set mrsData = Nothing
    Set mobjPublicPACS = Nothing
    Set mobjPatient = Nothing
    Set mobjSquareCard = Nothing
    
    If Not mobjReportForm Is Nothing Then Unload mobjReportForm
    Set mobjReportForm = Nothing
    Set mobjReport = Nothing
    Set mobjInfection = Nothing
End Sub

Private Sub cboVisit_Click()

    If cboVisit.Text = "" Then Exit Sub
    
    If mlngPreDept = cboVisit.ItemData(cboVisit.ListIndex) Then Exit Sub
    
    mlngPreDept = cboVisit.ItemData(cboVisit.ListIndex)
    
    mrsData.Filter = "���=" & mlngPreDept
    
    mlng����ID = mrsData!����id
    mlng����ID = mrsData!����ID
    mblnMoved = False
    If Not mrsData.EOF Then
        mstr�Һŵ� = NVL(mrsData!NO, "")
        mblnMoved = Val(NVL(mrsData!����ת��, "")) = 1
    End If
    '��ʾ������Ϣ
    If mstr�Һŵ� <> "" Then
        Call ShowOutPatiInfo
    Else
        Call ShowInPatiInfo
    End If
    
    '��ʾ����Ŀ¼
    Me.tbcHistory(0).Caption = cboVisit.Text
    Call ShowArchiveTree
    If tvwArchive.Visible And tvwArchive.Enabled Then tvwArchive.SetFocus
End Sub

Private Sub FuncLookPicture()
'��Ƭ����
    Dim lngҽ��ID As Long
    
    If Not tvwArchive.SelectedItem Is Nothing Then
        lngҽ��ID = Val(Split(tvwArchive.SelectedItem.Tag, ";")(1) & "")
        If lngҽ��ID <> 0 Then
            If CreateObjectPacs(gobjPublicPacs) Then
                Call gobjPublicPacs.ShowImage(lngҽ��ID, Me, mblnMoved)
            End If
        End If
    End If
End Sub

Private Sub lblDept_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    Dim vPoint As POINTAPI
    If mbytType = E_���� Then Exit Sub
    If chkFilter.Value = vbChecked Then Exit Sub
    Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_View_Dept, , True)
    If Not objPopup Is Nothing Then
        vPoint.X = lblDept.Left / Screen.TwipsPerPixelX
        vPoint.Y = (lblDept.Top + lblDept.Height + 30) / Screen.TwipsPerPixelY
        ClientToScreen picPati.hwnd, vPoint
        objPopup.CommandBar.ShowPopup , vPoint.X * Screen.TwipsPerPixelX, vPoint.Y * Screen.TwipsPerPixelY
    End If
End Sub

Private Sub mclsDockAduits_AfterEprPrint(ByVal lngRecordId As Long)
    mstrPrintDocIDs = mstrPrintDocIDs & lngRecordId & ","
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
'���ܣ�������ӡ�¼���д����ҳ��ӡ����
    Dim strSQL As String
    
    If mblnPrint And mbytPrintType = 1 Then
        If InStr("," & mstrPrintMedRec & ",", ",0_9,") = 0 Then
            strSQL = "Zl_���Ӳ�����ӡ_Insert(Null,9," & mlng����ID & "," & mlng����ID & ",'" & UserInfo.���� & "')"
            On Error GoTo errH
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            '���Ӳ�����ӡΨһ����(�ļ�ID, ����, ��ӡʱ��) ������ҳ:�ļ�IDΪNULL,����:9 ��ӡʱ��Oracle�����Զ���ȡ
            '������ӡ��ҳ������ʱ,ֻ��¼һ��
            mstrPrintMedRec = mstrPrintMedRec & "," & "0_9"
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optType_Click(Index As Integer)
    Dim i As Integer
    Dim blnVisible As Boolean
    Dim strCardKind  As String
    mbytType = Index
    fraInPati.Visible = False: fraOutPati.Visible = False
    
    If Index = 0 Then 'סԺ
        fraInPati.Visible = True
        On Error Resume Next
        If mobjSquareCard Is Nothing Then Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        err.Clear: On Error GoTo 0
        strCardKind = mstrCardKindIN
        If mobjSquareCard.zlInitComponents(Me, P�������Ĵ�ӡ, glngSys, UserInfo.�û���, gcnOracle, False) = False Then
            Set mobjSquareCard = Nothing
            MsgBox "ҽ�ƿ�������zl9CardSquare����ʼ��ʧ��!", vbInformation, gstrSysName
        End If
        Call PatiIdentify.zlInit(Me, glngSys, P�������Ĵ�ӡ, gcnOracle, gstrDBUser, mobjSquareCard, strCardKind, "zl9CISJob")
        PatiIdentify.objIDKind.AllowAutoICCard = True
        PatiIdentify.objIDKind.AllowAutoIDCard = True
        For i = 0 To tbcPati.ItemCount - 1
            If InStr(",��Ժ,��Ժ,", tbcPati.Item(i).Tag) > 0 Then
                blnVisible = True
            Else
                blnVisible = False
            End If
            tbcPati.Item(i).Visible = blnVisible
        Next
        mblnUndo = True
        tbcPati.Item(0).Selected = True 'ȱʡѡ����Ժ
        mblnUndo = False
    ElseIf Index = 1 Then '����
        fraOutPati.Visible = True
        '�����������
        On Error Resume Next
        If mobjSquareCard Is Nothing Then Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        err.Clear: On Error GoTo 0
        strCardKind = mstrCardKindOut
        If Not mobjSquareCard Is Nothing Then
            If mobjSquareCard.zlInitComponents(Me, P�������Ĵ�ӡ, glngSys, gstrDBUser, gcnOracle, False) = False Then
                Set mobjSquareCard = Nothing
                MsgBox "ҽ�ƿ�������zl9CardSquare����ʼ��ʧ��!", vbInformation, gstrSysName
            Else
                strCardKind = mobjSquareCard.zlGetIDKindStr(strCardKind)
            End If
        End If
        Call PatiIdentify.zlInit(Me, glngSys, P�������Ĵ�ӡ, gcnOracle, gstrDBUser, mobjSquareCard, strCardKind, "zl9CISJob")
        PatiIdentify.objIDKind.AllowAutoICCard = True
        PatiIdentify.objIDKind.AllowAutoIDCard = True
    
        For i = 0 To tbcPati.ItemCount - 1
            If InStr(",����,����,", tbcPati.Item(i).Tag) > 0 Then
                blnVisible = True
            Else
                blnVisible = False
            End If
            tbcPati.Item(i).Visible = blnVisible
        Next
        mblnUndo = True
        tbcPati.Item(2).Selected = True 'ȱʡѡ������
        mblnUndo = False
    End If
    Call LoadDept
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

Private Sub PatiIdentify_FindPatiBefore(ByVal objCard As zlIDKind.Card, blnCard As Boolean, strShowText As String, objCardData As zlIDKind.PatiInfor, blnFindPatied As Boolean, blnCancel As Boolean)
    Dim strPati As String, vRect As RECT, strName As String
    Dim rsTmp As ADODB.Recordset
    Dim lngPatiID As Long
    
    If chkFilter.Value = vbUnchecked Then Exit Sub
    
    strName = Trim(PatiIdentify.Text)
'    "��ˢ��������[-����ID]��[+סԺ��]��[*�����]�ȷ�ʽ��ȡ���˵���Ϣ��"
    On Error GoTo ErrHand
    blnCancel = False
            
    If objCard.���� Like "*��*��*" And blnCard = False And strName <> "" And InStr("-*+/", Left(Trim(PatiIdentify.Text), 1)) = 0 Then

        strPati = "Select 1 As ����id, a.����id As ID, a.����id, a.����, a.�Ա�, a.����, a.סԺ��, a.�����, a.סԺ����, Trunc(a.��Ժʱ��, 'dd') As ��Ժ����, a.��������," & vbNewLine & _
                "       a.���֤��, a.��ͥ��ַ, a.������λ, a.��������" & vbNewLine & _
                "From ������Ϣ A" & vbNewLine & _
                "Where ���� = [1]"
        strPati = strPati & " Order by ����ID,��Ժ���� Desc"
        
        vRect = zlControl.GetControlRect(PatiIdentify.hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, _
            vRect.Left, vRect.Top, PatiIdentify.Height, blnCancel, False, True, strName)
        If Not rsTmp Is Nothing Then
            If NVL(rsTmp!ID) = 0 Then
                blnCancel = True: Exit Sub
            Else '�Բ���ID��ȡ
                lngPatiID = NVL(rsTmp!ID)
            End If
        Else 'ȡ��ѡ��
            If blnCancel = False Then
                MsgBox "û���ҵ����������Ĳ��ˡ�" & strName & "����", vbInformation, gstrSysName
            End If
            blnCancel = True: Exit Sub
        End If
        Call ExecuteFindPati(False, lngPatiID)
        blnCancel = True: Exit Sub
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    mintFindType = Index - 1: mstrFindType = objCard.����
End Sub

Private Sub PatiIdentify_KeyPress(KeyAscii As Integer)
    Select Case mstrFindType
        Case "סԺ��", "�����"
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

Private Sub picLeft_Resize()
    On Error Resume Next
    With cboVisit
        .Top = 60
        .Left = 60
        .Width = picLeft.ScaleWidth - 120
        .Height = 300
    End With
    With tbcHistory
        .Left = 60
        .Top = cboVisit.Top + cboVisit.Height
        .Width = picLeft.ScaleWidth - 120
        .Height = picLeft.ScaleHeight - (cboVisit.Height + 60)
    End With
End Sub

Private Sub picPati_GotFocus()
    On Error Resume Next
'    If rptPati.Visible Then rptPati.SetFocus
End Sub

Private Sub picPati_Resize()
    Dim lngSplit As Long
    
    On Error Resume Next
    fraType.Move 0, 0
    lngSplit = 75
     
    lblDept.Top = fraType.Height + lngSplit + (cboDept.Height - lblDept.Height) / 2
    lblDept.Left = lngSplit
    cboDept.Top = fraType.Top + fraType.Height + lngSplit
    cboDept.Left = lblDept.Left + lblDept.Width + lngSplit / 2
    cboDept.Width = picPati.ScaleWidth - cboDept.Left - lblDept.Left
    
    PatiIdentify.Left = cboDept.Left
    PatiIdentify.Top = cboDept.Top + cboDept.Height + lngSplit
    PatiIdentify.Width = cboDept.Width - chkFilter.Width - 60
    
    chkFilter.Left = PatiIdentify.Left + PatiIdentify.Width + 60
    chkFilter.Top = PatiIdentify.Top + (PatiIdentify.Height - chkFilter.Height) / 2
    lblFind.Left = lblDept.Left
    lblFind.Top = lngSplit + PatiIdentify.Top + (PatiIdentify.Height - lblFind.Height) / 2
    lblFind.Width = lblDept.Width

    tbcPati.Left = 0
    tbcPati.Top = PatiIdentify.Top + PatiIdentify.Height + lngSplit
    tbcPati.Width = picPati.ScaleWidth
    tbcPati.Height = picPati.ScaleHeight - tbcPati.Top
    
    picPatiIn.Width = picPati.ScaleWidth
End Sub

Private Sub picPatiIn_Resize()
    Dim lngTop As Long
    Dim i As Long
    On Error Resume Next
    For i = 0 To picPara.Count - 1
        If picPara(i).Visible Then
            picPara(i).Width = picPatiIn.ScaleWidth
            picPara(i).Top = lngTop
            lngTop = picPara(i).Top + picPara(i).Height
        End If
    Next
    rptPati.Move 60, lngTop, picPatiIn.ScaleWidth - 60, picPatiIn.ScaleHeight - lngTop
End Sub

Private Sub rptPati_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    
    If Item.Checked = True Then
        mintPatiCount = mintPatiCount + 1
    Else
        mintPatiCount = mintPatiCount - 1
    End If
    If mintPatiCount = rptPati.Records.Count Then
        rptPati.Columns(col_ѡ��).Icon = imgPati.ListImages("Check").Index - 1
    Else
        rptPati.Columns(col_ѡ��).Icon = imgPati.ListImages("UnCheck").Index - 1
    End If
    stbThis.Panels(2).Text = IIf(mintPatiCount = 0, "", "��ѡ��" & mintPatiCount & "�����ˣ�")
End Sub

Private Sub rptPati_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim hitColumn As ReportColumn
    Dim lngHit As Long
    
    If Button = 1 Then
        Set hitColumn = rptPati.HitTest(X, Y).Column
        If Not hitColumn Is Nothing Then
            If hitColumn.ItemIndex = col_ѡ�� Then
                lngHit = rptPati.HitTest(X, Y).ht
                If xtpHitTestHeader = lngHit Then
                    If rptPati.Records.Count = 0 Then Exit Sub  '������ʱ��ֹ�л�
                    If hitColumn.Icon = imgPati.ListImages("Check").Index - 1 Then
                        hitColumn.Icon = imgPati.ListImages("UnCheck").Index - 1
                        SelectItems 2
                    Else
                        hitColumn.Icon = imgPati.ListImages("Check").Index - 1
                        SelectItems 1
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptPati_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim hitColumn As ReportColumn
    Dim Item As ReportRecordItem
    Dim strTipInfo As String
    Dim vPos As POINTAPI
    Dim lngHwnd As Long
    
    On Error Resume Next
    Set hitColumn = rptPati.HitTest(X, Y).Column
    If Not hitColumn Is Nothing Then
        If hitColumn.Index = col_��ӡͼ�� Then
            Set Item = rptPati.HitTest(X, Y).Item
            If Not Item Is Nothing Then
                If Item.Record(col_��ӡͼ��).Icon <> -1 Then
                    strTipInfo = Item.Record(col_��ӡ��¼).Value
                    If strTipInfo = "" Then '���û�л�ȡ������������ȡ����¼���б���
                        strTipInfo = GetPrintLog(Item.Record(col_����Id).Value, Item.Record(col_��ҳID).Value) '��ȡ��ӡ��¼
                        Item(col_��ӡ��¼).Value = strTipInfo
                    End If
                    GetCursorPos vPos
                    lngHwnd = WindowFromPoint(vPos.X, vPos.Y)
                    Call zlCommFun.ShowTipInfo(lngHwnd, strTipInfo, True)
                End If
            End If
        Else
            Call zlCommFun.ShowTipInfo(lngHwnd, "")
        End If
    End If
End Sub

Private Sub rptPati_SelectionChanged()
'����:
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    If mblnUndo Then Exit Sub       '�л�����סԺ�ᴥ��
    If Not Me.Visible Then Exit Sub
    If rptPati.SelectedRows.Count = 0 Then Exit Sub          '���������
    
    On Error GoTo errH
    
    With rptPati.SelectedRows(0)
        If .GroupRow Then
            Call ClearPatiInfo
        Else
            '������Ƭ
            If rptPati.Tag = Val(.Record(col_����Id).Value & "") & "_" & Val(.Record(col_��ҳID).Value & "") Then Exit Sub
            If Not ReadPatPricture(Val(.Record(col_����Id).Value & ""), imgPatient) Then
               Set imgPatient.Picture = imgList.ListImages("Patient").Picture
            End If
            mlng����ID = Val(.Record(col_����Id).Value & "")
            mlng����ID = Val(.Record(col_��ҳID).Value & "")
            rptPati.Tag = Val(.Record(col_����Id).Value & "") & "_" & Val(.Record(col_��ҳID).Value & "")
            Call InitBasicData
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tbcArchive_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'���ܣ�ˢ���Ӵ�����漰����
'˵����������Ϊ�л����濨Ƭ����
    Dim Index As Long, objItem As TabControlItem
    
    If mblnTabTmp Then Exit Sub
    If Item.Tag = "" Then Exit Sub '��ʼ��ʱ,��û��ֵ
    If Item.Handle = picTmp.hwnd Then
        Screen.MousePointer = 11
        Index = Item.Index
        mblnTabTmp = True
        On Error GoTo errH
        Select Case Item.Tag
            Case "������ҳ"
                Set objItem = tbcArchive.InsertItem(Index, "������ҳ", mcolSubForm("_������ҳ").hwnd, 0)
                objItem.Tag = "������ҳ"
            Case "סԺ��ҳ"
                Set objItem = tbcArchive.InsertItem(Index, "סԺ��ҳ", mcolSubForm("_סԺ��ҳ").hwnd, 0)
                objItem.Tag = "סԺ��ҳ"
            Case "������Ϣ"
                Set objItem = tbcArchive.InsertItem(Index, "������Ϣ", mcolSubForm("_������Ϣ").hwnd, 0)
                objItem.Tag = "������Ϣ"
            Case "����ҽ��"
                Set objItem = tbcArchive.InsertItem(Index, "����ҽ��", mcolSubForm("_����ҽ��").hwnd, 0)
                objItem.Tag = "����ҽ��"
            Case "סԺҽ��"
                Set objItem = tbcArchive.InsertItem(Index, "סԺҽ��", mcolSubForm("_סԺҽ��").hwnd, 0)
                objItem.Tag = "סԺҽ��"
            Case "���¼�¼��"
                Set objItem = tbcArchive.InsertItem(Index, "���¼�¼��", mcolSubForm("_���¼�¼��").hwnd, 0)
                objItem.Tag = "���¼�¼��"
            Case "�����¼��"
                Set objItem = tbcArchive.InsertItem(Index, "�����¼��", mcolSubForm("_�����¼��").hwnd, 0)
                objItem.Tag = "�����¼��"
            Case "�ٴ�·��"
                Set objItem = tbcArchive.InsertItem(Index, "�ٴ�·��", mcolSubForm("_�ٴ�·��").hwnd, 0)
                objItem.Tag = "�ٴ�·��"
            Case "�°滤��"
                Set objItem = tbcArchive.InsertItem(Index, "�°滤��", mcolSubForm("_�°滤��").hwnd, 0)
                objItem.Tag = "�°滤��"
            Case "���Ӳ���"
                Set objItem = tbcArchive.InsertItem(Index, "���Ӳ���", mcolSubForm("_���Ӳ���").hwnd, 0)
                objItem.Tag = "���Ӳ���"
            Case "��鱨��"
                Set objItem = tbcArchive.InsertItem(Index, "��鱨��", mcolSubForm("_��鱨��").hwnd, 0)
                objItem.Tag = "��鱨��"
            Case "����"
                If mobjReportForm Is Nothing Then
                    Set objItem = tbcArchive.InsertItem(Index, "����", picTmp.hwnd, 0)
                Else
                    Set objItem = tbcArchive.InsertItem(Index, "����", mobjReportForm.hwnd, 0)
                End If
                objItem.Tag = "����"
        End Select
        Call tbcArchive.RemoveItem(Index + 1)
        objItem.Selected = True
        mblnTabTmp = False
        Screen.MousePointer = 0
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowArchiveTab(ByVal strShow As String, ByVal strCaption As String)
'���ܣ��л���ʾ��ͬ�ĵ���ҳ�棬������ս���
    Dim i As Long
    
    For i = 0 To tbcArchive.ItemCount - 1
        If tbcArchive(i).Tag = strShow Then
            'Ĭ�ϵĿ�Ƭ����ǰ����Ҫչʾ��һ��ʱ�����ܴ��廹δ����ȥ������ͨ�������ж�һ���ֶ���һ�Ρ�������ֶ��ظ�ִ��
            If tbcArchive.Item(i).Handle = picTmp.hwnd Then Call tbcArchive_SelectedChanged(tbcArchive.Item(i))
            tbcArchive(i).Caption = strCaption
            If Not tbcArchive(i).Visible Then
                tbcArchive(i).Visible = True
                tbcArchive(i).Selected = True
                Exit For
            End If
        End If
    Next
    
    For i = 0 To tbcArchive.ItemCount - 1
        If tbcArchive(i).Tag <> strShow Then
            If tbcArchive(i).Visible Then tbcArchive(i).Visible = False
        End If
    Next
End Sub

Private Sub tbcPati_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim i As Long
    If chkFilter.Value = Checked Then Exit Sub
    For i = 0 To picPara.Count - 1
        picPara(i).Visible = False
    Next
    
    If Item.Tag = "��Ժ" Then
        picPara(0).Visible = True
    ElseIf Item.Tag = "����" Then
        picPara(1).Visible = True
    End If
    
    Call picPatiIn_Resize
    
    If Me.Visible And Not mblnUndo Then
        LoadPatients
    End If
End Sub

Private Sub FuncPrintOrView(ByVal bytFunc As Byte)
'����:Ԥ����ӡ
'����:bytFunc=2����ӡ����=1��Ԥ����0=����;3-PDF
'˵��:�������˴�ӡ�����ļ�;�������˴�ӡ����ѡ���ļ�;
'     ������ӡ�����ļ�;������ӡ����ļ�
    Dim objFSO As New Scripting.FileSystemObject    'FSO����
    Dim rsTemp As ADODB.Recordset
    
    Dim strKey As String
    Dim strPath As String
    Dim strSQL As String
    Dim strRegRange As String
    Dim strFile   As String
    Dim str�Һŵ�   As String
    Dim strPatiName As String
    Dim strPatiNo  As String
    Dim strErr      As String
    
    Dim i As Long, j As Long
    Dim lngSel As Long
    Dim lngNo  As Long
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    Dim lng����ID As Long
    
    Dim blnMoveData As Boolean
    Dim blnPath     As Boolean
    Dim strDeviceName   As String
    
    
    If mlng����ID = 0 Then Exit Sub
    On Error GoTo errH
    If bytFunc = 3 Then
        strPath = GetRegister(˽��ģ��, "��ӡ����", "PDFλ��", App.Path)
        If strPath = "" Then SetPDFPath
        If mbytPDFStatu = 0 Then
            If Not zlCommFun.PDFInitialize(strErr) Then
                MsgBox "PDF�豸��ʼ��ʧ�ܣ�" & strErr, vbExclamation, gstrSysName
                Exit Sub
            Else
                mbytPDFStatu = 1
            End If
        End If
        '����Ƿ����TinyPDF(32λϵͳ) Foxit Reader PDF Printer (64λϵͳ)��ӡ��
        strDeviceName = zlCommFun.PDFPrinterDeviceName()
    Else
        If Not LoadPrint Then Exit Sub
    End If
    With tvwArchive
        If bytFunc = 2 Or bytFunc = 3 Then
            Me.MousePointer = 11
            '��ӡѡ����Ŀ
            If mintPatiCount = 0 Then
                lngSel = 0
                mstrPrintDocIDs = ""
                mstrPrintMedRec = ""
                If mstr�Һŵ� = "" Then
                    strPatiName = lblShow(lbl_����).Caption
                    strPatiNo = lblShow(lbl_סԺ��).Caption
                    If bytFunc = 2 Then
                        lngNo = 1
                        strSQL = "Select Nvl(Max(��ӡ����),0)+1 As ��ӡ���� From ������ӡ��¼ Where ����id=[1] And ��ҳid=[2]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
                        If rsTemp.BOF = False Then
                            lngNo = rsTemp!��ӡ����
                        End If
                        mblnPrint = True
                    End If
                Else
                    lngNo = -1
                    strPatiName = lblShow(lbl_����).Caption
                    strPatiNo = ""
                    mblnPrint = False
                End If
                mrsMedRec.Filter = ""
                Do While Not mrsMedRec.EOF
                    mrsMedRec!ѡ�� = 0
                    mrsMedRec.MoveNext
                Loop
                For i = 1 To .Nodes.Count
                    strKey = .Nodes.Item(i).Key
                    If .Nodes.Item(i).Checked And Not InStr(",R0,R1,R2,R3,R4,R5,R6,R7,R9,R10,R11,R12,", strKey) > 0 Then
                        mrsMedRec.Filter = "ID='" & strKey & "'"
                        mrsMedRec!ѡ�� = 1
                        lngSel = lngSel + 1
                    End If
                Next
                mrsMedRec.Filter = "ѡ��=1"
                If mrsMedRec.RecordCount > 0 Then psb.Visible = True: psb.Max = mrsMedRec.RecordCount
                For i = 1 To mrsMedRec.RecordCount
                    psb.Value = i
                    Call PrintOrView(bytFunc, mlng����ID, mlng����ID, mlng����ID, mrsMedRec!ID & "", mrsMedRec!���� & "", strPath, strPatiName, strPatiNo, lngNo, mblnMoved, strDeviceName)
                    mrsMedRec.MoveNext
                Next
                If psb.Visible Then psb.Visible = False
                If lngSel = 0 Then
                    strKey = .SelectedItem.Key
                    Call PrintOrView(bytFunc, mlng����ID, mlng����ID, mlng����ID, strKey, .SelectedItem.Tag, strPath, strPatiName, strPatiNo, lngNo, mblnMoved, strDeviceName)
                End If
            Else
                '������ӡ
                'ͳ�Ʋ���¼��ӡ���
                strRegRange = ""
                For i = 1 To .Nodes.Count
                    If .Nodes.Item(i).Checked And InStr("R1,R2,R3,R4,R5,R6,R7,R8,R9,R10,R11,R12,", .Nodes.Item(i).Key) > 0 Then
                        If .Nodes.Item(i).Key = "R8" Then
                            strRegRange = strRegRange & " OR ID ='R8'"
                        Else
                            strRegRange = strRegRange & " OR �ϼ�ID ='" & .Nodes.Item(i).Key & "'"
                        End If
                    End If
                Next
                If strRegRange <> "" Then strRegRange = Mid(strRegRange, 5)
                If strRegRange = "" Then
                    MsgBox "�����ļ��б���ѡ����Ҫ������ļ����͡�", vbInformation, Me.Caption
                    GoTo errMsg
                End If
                For i = 0 To rptPati.Records.Count - 1
                    If rptPati.Records(i).Item(col_ѡ��).Checked = True Then
                        With rptPati.Records(i)
                            lng����ID = Val(.Item(col_����Id).Value & "")
                            lng��ҳID = Val(.Item(col_��ҳID).Value & "")
                            lng����ID = Val(.Item(col_����ID).Value & "")
                            
                            strPatiName = .Item(col_����).Value & ""
                            mstrPrintDocIDs = ""
                            mstrPrintMedRec = ""
                            If mbytType = E_סԺ Then
                                str�Һŵ� = ""
                                blnMoveData = Val(.Item(COL_����ת��).Value & "") = 1
                                strPatiNo = .Item(col_סԺ��).Value & ""
                                blnPath = False
                                If GetInsidePrivs(p�ٴ�·��Ӧ��) <> "" Then
                                    blnPath = HavePath(lng����ID)
                                End If
                                If bytFunc = 2 Then
                                    lngNo = 1
                                    strSQL = "Select Nvl(Max(��ӡ����),0)+1 As ��ӡ���� From ������ӡ��¼ Where ����id=[1] And ��ҳid=[2]"
                                    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
                                    If rsTemp.BOF = False Then
                                        lngNo = rsTemp!��ӡ����
                                    End If
                                    mblnPrint = True
                                End If
                            Else
                                lngNo = -1
                                str�Һŵ� = .Item(COL_NO).Value
                                blnMoveData = False
                                blnPath = False
                                mblnPrint = False
                            End If

                            Set rsTemp = GetCISStruct(lng����ID, lng��ҳID, str�Һŵ�, blnPath, blnMoveData)
                            If Not rsTemp Is Nothing Then
                                rsTemp.Filter = strRegRange
                                If rsTemp.RecordCount > 0 Then psb.Visible = True: psb.Max = rsTemp.RecordCount
                                For j = 1 To rsTemp.RecordCount
                                    psb.Visible = j
                                    Call PrintOrView(bytFunc, lng����ID, lng��ҳID, lng����ID, rsTemp!ID & "", rsTemp!���� & "", strPath, strPatiName, strPatiNo, lngNo, blnMoveData, strDeviceName)
                                    rsTemp.MoveNext
                                Next
                                If psb.Visible Then psb.Visible = False
                            End If
                        End With
                    End If
                Next
            End If
            mstrPrintMedRec = ""
            Me.MousePointer = 0
        ElseIf bytFunc = 1 Then
            If InStr(",R0,R1,R2,R3,R4,R5,R6,R7,R9,R10,R11,R12,", "," & .SelectedItem.Key & ",") > 0 Then Exit Sub
            Call PrintOrView(1, mlng����ID, mlng����ID, mlng����ID, .SelectedItem.Key, .SelectedItem.Tag, , , , , mblnMoved)
        ElseIf bytFunc = 0 Then
            Call PrintInMedRec(Nothing, 0, mlng����ID, mlng����ID, mobjReport, mlng����ID, Me)
        End If
    End With
    Exit Sub
errMsg:
    If bytFunc > 1 Then Me.MousePointer = 0
    Exit Sub
errH:
    If bytFunc > 1 Then Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PrintOrView(ByVal bytFunc As Byte, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, _
    ByVal strKey As String, Optional ByVal strPara As String, Optional ByVal strPath As String, Optional ByVal strPatiName As String, _
    Optional ByVal strHosNo As String, Optional ByVal lngNo As Long, Optional ByVal blnMoveData As Boolean, Optional ByVal strDeviceName As String)
'����:Ԥ����ӡ
'����:bytFunc=2����ӡ����=1��Ԥ����0=����;3-PDF
    Dim intPage As Integer
    Dim intMode As Integer
    Dim intSel  As Integer
    Dim objFile As New Scripting.FileSystemObject
    Dim varParam As Variant
    Dim blnMod  As Boolean
    Dim blnFoxitPDF As Boolean
    
    Dim strSQL          As String
    Dim strReportNO     As String
    Dim strInfo          As String
    Dim strMsg          As String
    Dim strFileName     As String
    Dim strType         As String
    Dim strPrint        As String
    Dim rsTemp          As ADODB.Recordset
    Dim objPrint        As Object
    Dim strPDFPath      As String

    On Error GoTo errH

    If InStr(",R0,R1,R2,R3,R4,R5,R6,R7,R9,R10,R11,R12,", "," & strKey & ",") > 0 Then Exit Sub
    '��ȡȱʡ��ӡ��
    varParam = Split(strPara, ";")
    If bytFunc = 1 Or bytFunc = 2 Then
        If InStr(strPara, ";EMR;") > 0 Then
            strPrint = mcolPrint(varParam(4))
        ElseIf strKey Like "R7P*" Or strKey Like "R7L*" Then
            strPrint = mcolPrint("R7")
        Else
            On Error Resume Next
            strPrint = mcolPrint(Split(strKey, "K")(0))
            If err.Number <> 0 Then strPrint = mcolPrint("R11")
            err.Clear: On Error GoTo errH
        End If
    End If
    If strKey Like "R11K*" Then
        intPage = Val(Replace(strKey, "R11K", ""))
        If bytFunc = 3 Then
            If intPage = 1 Then
                strType = "��ҳ����"
            ElseIf intPage = 2 Then
                strType = "��ҳ����"
            ElseIf intPage = 3 Then
                strType = "��ҳ��ҳһ"
            ElseIf intPage = 4 Then
                strType = "��ҳ��ҳ��"
            End If
            strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng��ҳID & "_" & strType & ".PDF"
            Call zlCommFun.PDFFile(strFileName)
            Call PrintInMedRec(Nothing, 4, lng����ID, lng��ҳID, mobjReport, lng����ID, Me, intPage, strFileName)
        Else
            If bytFunc = 2 Then mbytPrintType = 1 '�����ҳ��ӡ
            Call PrintInMedRec(Nothing, bytFunc, lng����ID, lng��ҳID, mobjReport, lng����ID, Me, intPage, IIf(bytFunc = 2, strPrint, ""))
            If bytFunc = 2 Then mbytPrintType = 0 'ȡ�����
        End If
    ElseIf strKey Like "R12K*" Then
        If strKey = "R12K1" Then
            If bytFunc = 3 Then
                strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng��ҳID & "_��ʱҽ��.PDF"
                Call zlCommFun.PDFFile(strFileName)
            End If
             '�ٴ�ӡ����
            Call gobjKernel.zlPrintAdvice(Me, lng����ID, lng��ҳID, 0, 1, IIf(bytFunc = 3, strFileName, strPrint), IIf(bytFunc = 3, 4, bytFunc))
        ElseIf strKey = "R12K2" Then
            If bytFunc = 3 Then
                strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng��ҳID & "_����ҽ��.PDF"
                Call zlCommFun.PDFFile(strFileName)
            End If
            '�ȴ�ӡ����
            Call gobjKernel.zlPrintAdvice(Me, lng����ID, lng��ҳID, 0, 0, IIf(bytFunc = 3, strFileName, strPrint), IIf(bytFunc = 3, 4, bytFunc))
        End If
    ElseIf strKey Like "R1K*" Or strKey Like "R2K*" Or strKey Like "R5K*" Or strKey Like "R6K*" Or strKey Like "R4K*" Then  '1-���ﲡ��,2-סԺ����;4-������;5-����֤��;6-֪���ļ�
        If bytFunc = 1 Then
            If strKey Like "R5K*" And Val(varParam(5)) = 2 Then 'Ԥ��������Ⱦ���濨
                Call FuncViewDisReportCard(Val(varParam(0)))
            Else
                Call mclsDockAduits.zlPrintDocument(3, 1, Val(varParam(0)), strPrint, blnMoveData)
            End If
        Else
            'סԺ�����ͻ����� ���ڹ����ĵ�ʱֻ��ӡ���һ��
            If (strKey Like "R2K*" Or strKey Like "R4K*") And InStr("," & mstrPrintDocIDs, "," & varParam(0) & ",") > 0 Then Exit Sub      '
            If bytFunc = 3 Then
                If strKey Like "R1K*" Then
                    strType = "���ﲡ��"
                ElseIf strKey Like "R2K*" Then
                    strType = "סԺ����"
                ElseIf strKey Like "R4K*" Then
                    strType = "������"
                ElseIf strKey Like "R5K*" Then
                    strType = "����֤��"
                ElseIf strKey Like "R6K*" Then
                    strType = "֪���ļ�"
                End If
                strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng��ҳID & "_" & strType & "_" & varParam(0) & ".PDF"
                Call zlCommFun.PDFFile(strFileName): blnFoxitPDF = True
                strDeviceName = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
            End If
            Call mclsDockAduits.zlPrintDocument(3, 2, Val(varParam(0)), IIf(bytFunc = 3, strDeviceName, strPrint), blnMoveData)
            If bytFunc = 2 Then
                Call RecordEprPrintInfo(1, Val(varParam(0)), lngNo)
            End If
        End If
    ElseIf strKey Like "R3K*" Then   '3-�����¼
        If mblnNewTends = False Then
            If UBound(varParam) >= 1 Then
                If Val(varParam(1)) = -1 Then '���µ�
                    If bytFunc = 3 Then
                        strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng��ҳID & "_�����¼_" & Val(varParam(3)) & ".PDF"
                        Call zlCommFun.PDFFile(strFileName): blnFoxitPDF = True
                        strDeviceName = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
                    End If
                    Call mclsDockAduits.zlRefreshTendBody(lng����ID, lng��ҳID, Val(Split(varParam(0), "_")(0)), Val(varParam(4)), blnMoveData)
                    Call mclsDockAduits.zlPrintDocument(1, IIf(bytFunc = 3, 2, bytFunc), , IIf(bytFunc = 3, strDeviceName, strPrint), blnMoveData)
                    
                    If bytFunc = 2 Then
                        Call RecordEprPrintInfo(2, "���µ�", lngNo, lng����ID, lng��ҳID)
                    End If
                Else '�����¼
                    If bytFunc = 3 Then
                        strFileName = strPath & "\" & strFileName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng��ҳID & "_�����¼_" & Val(varParam(3)) & ".PDF"
                        Call zlCommFun.PDFFile(strFileName): blnFoxitPDF = True
                        strDeviceName = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
                    End If
                    Call mclsDockAduits.zlRefresh(3, Val(varParam(3)), lng����ID, lng��ҳID, Val(varParam(0)), CStr(varParam(2)), , Val(varParam(4)), blnMoveData)
                    Call mclsDockAduits.zlPrintDocument(2, IIf(bytFunc = 3, 2, bytFunc), , IIf(bytFunc = 3, strDeviceName, strPrint), blnMoveData)
                    
                    If bytFunc = 2 Then
                        Call RecordEprPrintInfo(3, Val(varParam(3)), lngNo, lng����ID, lng��ҳID)
                    End If
                End If
            End If
        Else
            If UBound(varParam) >= 1 Then
                Select Case Val(varParam(1))
                    Case -1 '���µ�
                        intSel = 1
                    Case 1  '����ͼ
                        intSel = 3
                    Case Else '��¼��
                        intSel = 2
                End Select
                strInfo = "��ʼ���" & strPatiName & Decode(intSel, 1, "���µ�", 2, "�����¼", "����ͼ") & "_" & Val(varParam(3))
                If bytFunc = 3 Then
                    strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng��ҳID & "_" & Decode(intSel, 1, "���µ�", 2, "�����¼", "����ͼ") & "_" & Val(varParam(3)) & ".PDF"
                    Call zlCommFun.PDFFile(strFileName): blnFoxitPDF = True
                End If
                Call mclsTendsNew.zlPrintDocument(lng����ID, lng��ҳID, Val(varParam(4)), Val(varParam(0)), Val(varParam(3)), intSel, IIf(bytFunc = 3, strDeviceName, strPrint), IIf(bytFunc = 1, False, True))
                If bytFunc = 2 Then
                    Call RecordEprPrintInfo(2, Decode(intSel, 1, "���µ�", 2, "�����¼", "����ͼ"), lngNo, lng����ID, lng��ҳID)
                End If
            End If
        End If
    ElseIf strKey Like "R7K*" Then   '7-���Ʊ���
        If bytFunc = 3 Then
            If Val(varParam(0)) = 0 Then
                strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng��ҳID & "_" & varParam(4) & "_" & varParam(2) & ".PDF"
            Else
                strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng��ҳID & "_" & varParam(4) & "_" & Val(varParam(0)) & ".PDF"
            End If
            Call zlCommFun.PDFFile(strFileName)
            strDeviceName = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
        End If
        If varParam(2) = "E" Then
            If mstr�����Ӧ���� <> "" Then
                strReportNO = Split(mstr�����Ӧ����, ",")(2)
                If bytFunc = 3 Then
                    Call mobjReport.ReportOpen(gcnOracle, 0, strReportNO, Me, "����id=" & lng����ID, "��ҳid=" & lng��ҳID, "ҽ��ID=" & varParam(1), "PDF=" & strFileName, 4)
                Else
                    Call mobjReport.SetReportPrintSet(gcnOracle, glngSys, strReportNO, "printer", strPrint)  '����ָ����ӡ��
                    Call mobjReport.ReportOpen(gcnOracle, 0, strReportNO, Me, "����id=" & lng����ID, "��ҳid=" & lng��ҳID, "ҽ��ID=" & varParam(1), bytFunc)
                End If
            ElseIf mblnLIS And Val(mstr���鱨���ӡ) = 1 And Not gobjLIS Is Nothing Then
                blnMod = gobjLIS.PrintLisReport(Me, Val(varParam(1)), IIf(bytFunc = 3, 4, bytFunc), , IIf(bytFunc = 3, strFileName, ""), strPrint, strMsg)     '1-Ԥ��;2-��ӡ;4-PDF
            Else
                Call mclsDockAduits.zlPrintDocument(4, IIf(bytFunc = 3, 2, bytFunc), Val(varParam(0)), IIf(bytFunc = 3, strDeviceName, strPrint), blnMoveData)
                If bytFunc = 3 Then blnFoxitPDF = True
            End If
        ElseIf varParam(2) = "D" Then
            If mstr����Ӧ���� <> "" Then
                strReportNO = Split(mstr����Ӧ����, ",")(2)
                If bytFunc = 3 Then
                    Call mobjReport.ReportOpen(gcnOracle, 0, strReportNO, Me, "����id=" & lng����ID, "��ҳid=" & lng��ҳID, "ҽ��ID=" & varParam(1), "PDF=" & strFileName, 4)
                Else
                    Call mobjReport.SetReportPrintSet(gcnOracle, glngSys, strReportNO, "printer", strPrint)  '����ָ����ӡ��
                    Call mobjReport.ReportOpen(gcnOracle, 0, strReportNO, Me, "����id=" & lng����ID, "��ҳid=" & lng��ҳID, "ҽ��ID=" & varParam(1), bytFunc)
                End If
            Else
                If Val(varParam(3)) <> 0 Then
                    'RIS
                    If Not gobjRis Is Nothing Then
                        If bytFunc = 1 Then
                            Call gobjRis.ShowViewReport(Me.hwnd, Val(varParam(1)), True, Val(varParam(3)))
                        ElseIf bytFunc = 2 Then
                            Call gobjRis.HISPrintReportByAppno(Me.hwnd, Val(varParam(1)), Val(varParam(3)), strPrint)
                        ElseIf bytFunc = 3 Then
                            strPDFPath = gobjRis.HISGetPrintFile(Me.hwnd, 1, Val(varParam(1)), Val(varParam(3)))
                            If strPDFPath <> "" Then
                                If objFile.FileExists(strPDFPath) And strPath <> "" Then
                                    objFile.MoveFile strPDFPath, strPath & "\"
                                End If
                            End If
                        End If
                    End If
                ElseIf Val(varParam(0)) <> 0 Then
                    Call mclsDockAduits.zlPrintDocument(4, IIf(bytFunc = 3, 2, bytFunc), Val(varParam(0)), IIf(bytFunc = 3, strDeviceName, strPrint), blnMoveData)
                    If bytFunc = 3 Then blnFoxitPDF = True
                End If
            End If
        End If
    ElseIf strKey Like "R7P*" Then  '��鱨��
        If Not mobjPublicPACS Is Nothing Then
            If bytFunc = 3 Then strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng��ҳID & "_" & varParam(1) & ".PDF"
            Call mobjPublicPACS.PrintReport(varParam(0), IIf(bytFunc = 3, strFileName, strPrint), IIf(bytFunc = 1, True, False))  'TrueԤ��
            If bytFunc = 3 Then blnFoxitPDF = True
        End If
    ElseIf strKey Like "R7L*" Then  '��������
        If bytFunc = 3 Then
            strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng��ҳID & "_" & Split(varParam(2), "<sTab>")(1) & "_" & varParam(0) & ".PDF"
            Call zlCommFun.PDFFile(strFileName)
            strFileName = GetLisRptFile(strPara, strFileName)
        End If
    ElseIf strKey Like "R8*" Then  '8-�ٴ�·��
        If bytFunc = 3 Then
            strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng��ҳID & "_�ٴ�·��.PDF"
            Call mclsPath.zlFuncPathTableOutPut(4, True, strFileName, lng����ID, lng��ҳID, strDeviceName)
            blnFoxitPDF = True
        Else
            Call mclsPath.zlRefreshReadOnly(lng����ID, lng��ҳID)
            Call mclsPath.zlFuncPathTableOutPut(IIf(bytFunc = 1, 2, 1), True, "", 0, 0, strPrint) '2-Ԥ��;1-��ӡ
            Call RecordEprPrintInfo(2, "�ٴ�·��", lngNo, lng����ID, lng��ҳID)
        End If
    ElseIf strKey Like "R9K*" Then   '9-סԺ֤
        strReportNO = "ZLCISBILL" & Format(varParam(0), "00000") & "-1"
        If bytFunc = 3 Then
            strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng��ҳID & "_סԺ֤.PDF"
            Call zlCommFun.PDFFile(strFileName)
            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "NO=" & varParam(1), "����=" & varParam(2), "ҽ��ID=0", "PDF=" & strFileName, 4)
        Else
            Call mobjReport.SetReportPrintSet(gcnOracle, glngSys, strReportNO, "printer", strPrint)  '����ָ����ӡ��
            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReportNO, Me, "NO=" & varParam(1), "����=" & varParam(2), "ҽ��ID=0", bytFunc)
        End If
    ElseIf strKey Like "R10K*" Then   '10-��������
        If bytFunc = 3 Then
            strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng��ҳID & "_��������_" & varParam(0) & ".PDF"
            Call zlCommFun.PDFFile(strFileName)
            Call mobjReport.ReportOpen(gcnOracle, 0, varParam(2), Me, "����id=" & lng����ID, "��ҳid=" & lng��ҳID, "PDF=" & strFileName, 4)
        Else
            Call mobjReport.SetReportPrintSet(gcnOracle, glngSys, varParam(2), "printer", strPrint)  '����ָ����ӡ��
            Call mobjReport.ReportOpen(gcnOracle, 0, varParam(2), Me, "����id=" & lng����ID, "��ҳid=" & lng��ҳID, bytFunc)
        End If
    ElseIf InStr(strKey, "R") = 0 And Len(strPara) >= 32 Then
        'EMR����Ԥ��
        If Not mobjRichEMR Is Nothing Then
            If varParam(1) <> "" Then
                Call mobjRichEMR.zlShowDoc(varParam(0), varParam(1))
            Else
                Call mobjRichEMR.zlShowDoc(varParam(0), "")
            End If
            If bytFunc = 1 Then
                Call mobjRichEMR.zlPrintDoc(True)
            ElseIf bytFunc = 2 Or bytFunc = 3 Then
                '���ڹ����ĵ�ʱֻ��ӡ���һ��
                If InStr("," & mstrPrintDocIDs, "," & varParam(0) & ",") > 0 Then Exit Sub
                If bytFunc = 3 Then
                    strFileName = strPath & "\" & strPatiName & IIf(strHosNo = "", "", "_" & strHosNo) & "_" & lng��ҳID & "_" & varParam(2) & varParam(0) & ".PDF"
                    Call zlCommFun.PDFFile(strFileName)
                    strDeviceName = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
                End If
                err.Number = 0: On Error Resume Next
                Call mobjRichEMR.zlPrintDoc(False, IIf(bytFunc = 3, strDeviceName, strPrint))
                If err.Number = 450 Then
                    err.Number = 0
                    Call mobjRichEMR.zlPrintDoc(False)
                End If
                err.Clear: On Error GoTo 0
                mstrPrintDocIDs = mstrPrintDocIDs & varParam(0) & ","  '
                If bytFunc = 3 Then blnFoxitPDF = True
            End If
        End If
    End If
    If blnFoxitPDF Then
        Call zlCommFun.PDFFileSuccess
    End If
    Exit Sub
errH:
    mstrPrintDocIDs = ""
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tvwArchive_NodeCheck(ByVal node As MSComctlLib.node)
    CheckNode node
End Sub

Private Sub tvwArchive_NodeClick(ByVal node As MSComctlLib.node)
    Dim arrPar As Variant
    Dim intSel As Integer
    Dim strFile As String
    Dim i As Long
    If mblnTvwTmp Then Exit Sub
    If tvwArchive.Tag = node.Key Then Exit Sub
    mblnTvwTmp = True
    LockWindowUpdate Me.hwnd
    
    arrPar = Split(node.Tag, ";")
     
    If node.Key Like "R1K*" Or node.Key Like "R2K*" Or node.Key Like "R4K*" Or node.Key Like "R5K*" Or node.Key Like "R6K*" Or node.Key Like "R7K*" Then
        Call ShowArchiveTab("������Ϣ", node.Text)
    End If
    mblnSeePic = False
    If node.Key = "R11" Or node.Key = "R0" Then
        Call ShowArchiveTab(IIf(mstr�Һŵ� <> "", "������ҳ", "סԺ��ҳ"), tbcHistory.Selected.Caption)
        Call mclsArchive.zlRefresh(IIf(mstr�Һŵ� <> "", 0, 1), mlng����ID, mlng����ID, mblnMoved)
    ElseIf node.Key = "R12" Then  'ҽ����¼
        If mstr�Һŵ� <> "" Then
            Call ShowArchiveTab("����ҽ��", tbcHistory.Selected.Caption)
            Call mclsOutAdvices.zlRefresh(mlng����ID, mstr�Һŵ�, False, mblnMoved)
        Else
            Call ShowArchiveTab("סԺҽ��", tbcHistory.Selected.Caption)
            Call mclsInAdvices.zlRefresh(mlng����ID, mlng����ID, mlng����ID, mlng����ID, 0, mblnMoved)
        End If
    ElseIf node.Key Like "R1K*" Then '���ﲡ��
        Call mclsDockAduits.zlRefresh(1, Val(arrPar(0)), , , , , , , mblnMoved)
    ElseIf node.Key Like "R2K*" Then 'סԺ����
        Call mclsDockAduits.zlRefresh(2, Val(arrPar(0)), , , , , , , mblnMoved)
    ElseIf node.Key Like "R3K*" Then '�����¼
        If UBound(arrPar) >= 1 Then
            If mblnNewTends = False Then
                If Val(arrPar(1)) = -1 Then
                    Call ShowArchiveTab("���¼�¼��", node.Text)
                    Call mclsDockAduits.zlRefreshTendBody(mlng����ID, mlng����ID, Val(arrPar(0)), 0, mblnMoved)
                Else
                    Call ShowArchiveTab("�����¼��", node.Text)
                    Call mclsDockAduits.zlRefresh(3, Val(arrPar(3)), mlng����ID, mlng����ID, Val(arrPar(0)), CStr(arrPar(2)), , , mblnMoved)
                End If
            Else
                Select Case Val(arrPar(1))
                    Case -1
                        intSel = 0
                    Case 1
                        intSel = 2
                    Case Else
                        intSel = 1
                End Select
                Call ShowArchiveTab("�°滤��", node.Text)
                Call mclsTendsNew.zlRefreshTendFile(mlng����ID, mlng����ID, Val(arrPar(4)), Val(arrPar(0)), False, IIf(glngModul = pסԺҽ��վ, True, False), intSel, Val(arrPar(3)), 1)
            End If
        End If
    ElseIf node.Key Like "R4K*" Then '������
        Call mclsDockAduits.zlRefresh(4, Val(arrPar(0)), , , , , , , mblnMoved)
    ElseIf node.Key Like "R5K*" Then '����֤��
        Call mclsDockAduits.zlRefresh(5, Val(arrPar(0)), , , , , , , mblnMoved)
    ElseIf node.Key Like "R6K*" Then '֪���ļ�
        Call mclsDockAduits.zlRefresh(6, Val(arrPar(0)), , , , , , , mblnMoved)
    ElseIf node.Key Like "R7K*" Then '���Ʊ���
        Call mclsDockAduits.zlRefresh(7, Val(arrPar(0)), , , , , , , mblnMoved)
        mblnSeePic = arrPar(2) = "D"
    ElseIf node.Key = "R8" Then
        If mstr�Һŵ� = "" Then
            Call ShowArchiveTab("�ٴ�·��", node.Text)
            Call mclsPath.zlRefreshReadOnly(mlng����ID, mlng����ID)
        End If
    ElseIf node.Key Like "R7P*" Then  '��鱨��
        mblnSeePic = True
        Call ShowArchiveTab("��鱨��", node.Text)
        If Not mobjPublicPACS Is Nothing Then Call mobjPublicPACS.zlDocRefresh(Split(node.Tag, ";")(0))
    ElseIf node.Key Like "R7L*" Then  '��������
        strFile = GetLisRptFile(node.Tag)
        If strFile <> "" Then
            If picRpt.Tag <> "" And picRpt.Tag <> mstrTempDel And picRpt.Tag <> strFile Then mstrTempDel = picRpt.Tag
            webRpt.Navigate strFile
            picRpt.Tag = strFile
        End If
        Call ShowArchiveTab("��������", node.Text)
    ElseIf node.Key Like "R9K*" Or node.Key Like "R10K*" Or node.Key Like "R11K*" Or node.Key Like "R12K*" Then   'סԺ֤;��������;��ҳ��ҽ��
        Call FuncShowReport(node)
    ElseIf InStr(node.Key, "R") = 0 And Len(node.Tag) >= 32 Then
        'EMR����Ԥ��
        If Not mobjRichEMR Is Nothing Then
            Call ShowArchiveTab("���Ӳ���", node.Text)
            If arrPar(1) <> "" Then
                Call mobjRichEMR.zlShowDoc(arrPar(0), arrPar(1))
            Else
                Call mobjRichEMR.zlShowDoc(arrPar(0), "")
            End If
        End If
    Else
        LockWindowUpdate 0
        mblnTvwTmp = False
        Exit Sub
    End If
    tvwArchive.Tag = node.Key
    mblnTvwTmp = False
    If tvwArchive.Visible And tvwArchive.Enabled Then tvwArchive.SetFocus
    LockWindowUpdate 0
End Sub

Private Function ShowArchiveTree() As Boolean
'���ܣ���ʾ���˵�������Ŀ¼
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, objNode As node, strSQL1 As String
    Dim blnPath As Boolean
    Dim strSel As String
    Dim strKey As String
    Dim i As Long
    
    Screen.MousePointer = 11
    
    If Not tvwArchive.SelectedItem Is Nothing Then
        strKey = tvwArchive.SelectedItem.Key
        If strKey Like "R11K*" Or strKey = "R11" Or strKey Like "R12K*" Or strKey = "R12" Then
            strSel = Split(strKey, "K")(0)
        End If
    End If
    
    '���˿��Ҵ��ڿ��õ��ٴ�·��ʱ����ʾ�ٴ�·����¼
    If mstr�Һŵ� = "" Then
        If GetInsidePrivs(p�ٴ�·��Ӧ��) <> "" Then
            blnPath = HavePath(mlng����ID)
        End If
    End If
    Set rsTmp = GetCISStruct(mlng����ID, mlng����ID, mstr�Һŵ�, blnPath, mblnMoved)
    Set mrsMedRec = rsTmp

    tvwArchive.Tag = ""
    tvwArchive.Nodes.Clear

    Do While Not rsTmp.EOF
        If NVL(rsTmp!�ϼ�ID) = "" Then
            Set objNode = tvwArchive.Nodes.Add(, , CStr(rsTmp!ID), rsTmp!����, NVL(rsTmp!ͼ��))
        Else
            Set objNode = tvwArchive.Nodes.Add(CStr(rsTmp!�ϼ�ID), tvwChild, CStr(rsTmp!ID), rsTmp!����, NVL(rsTmp!ͼ��))
        End If

        objNode.Tag = NVL(rsTmp!����)
        objNode.Expanded = True

        If tvwArchive.Nodes.Count = 1 Then
            objNode.Selected = True
        ElseIf objNode.Key = strSel Then
            objNode.Selected = True
        End If

        rsTmp.MoveNext
    Loop
    If Not tvwArchive.SelectedItem Is Nothing Then
        tvwArchive.SelectedItem.EnsureVisible
        Call tvwArchive_NodeClick(tvwArchive.SelectedItem)
    End If
    
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadDept()
'���ܣ����ﰴ���ҷ�ʽ��ȡ���ò���;סԺ ������/������ʽ��ȡ���ò���
'ע��:�˷�����Ҫ
    If mbytType = E_סԺ Then
        lblDept.Caption = IIf(mintDeptView = 0, "����(&D)��", "����(&D)��")
    Else
        lblDept.Caption = "����(&D)"
    End If
    mintPreDept = -1
    Call InitDepts
    Call cboDept_Click
    
    If cboDept.ListIndex = -1 Then
        If mbytType = E_סԺ Then
            If InStr(mstrPrivs, "ȫԺ����") > 0 Then
                MsgBox "û�з���סԺ" & IIf(mintDeptView = 0, "����", "����") & "��Ϣ,���ȵ����Ź��������ã�", vbInformation, gstrSysName
            Else
                MsgBox "û�з���������" & IIf(mintDeptView = 0, "����", "����") & ",����ʹ��סԺҽ������վ��", vbInformation, gstrSysName
            End If
        Else
            If InStr(mstrPrivs, "ȫԺ����") > 0 Then
                MsgBox "û�з������������Ϣ,���ȵ����Ź��������ã�", vbInformation, gstrSysName
            Else
                MsgBox "û�з�������������,����ʹ�ò�����ѯ��ӡ��", vbInformation, gstrSysName
            End If
        End If
    End If
End Sub

Private Function LoadPatients(Optional ByVal lngPatiID As Long) As Boolean
'����:��ȡ������Ϣ
    Dim strSQL As String
    Dim rsPati As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim lngColor As Long
    Dim strFilter As String
    Dim strCaption As String
    Dim strTemp As String
    Dim intBedLen As Integer
    
    On Error GoTo errH
    '��λ���ȹ̶�Ϊ10
    intBedLen = 10
    
    If lngPatiID <> 0 Then
        strSQL = "Select ����id,����,����,��Ժ,ִ��״̬ " & vbNewLine & _
                "From (Select a.��ҳid As ����id, a.��Ժ���� As ʱ��, 1 As ����, ��Ժ����id As ����, Decode(a.��Ժ����, Null, 1, 0) As ��Ժ, NULL as ִ��״̬ " & vbNewLine & _
                "       From ������ҳ A" & vbNewLine & _
                "       Where a.����id = [1] And Nvl(��ҳid, 0) <> 0" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select ID As ����id, b.ִ��ʱ�� As ʱ��, 0 As ����, b.ִ�в���id As ����, 0 As ��Ժ,ִ��״̬ " & vbNewLine & _
                "       From ���˹Һż�¼ B" & vbNewLine & _
                "       Where b.����id = [1] And ��¼���� = 1 And ��¼״̬ = 1 And ִ��״̬ In (1,2)) A" & vbNewLine & _
                "Order By a.ʱ�� Desc"

        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatiID)
        If Not rsPati.EOF Then
            If Val(rsPati!���� & "") = 1 Then
                mbytType = E_סԺ
                If (rsPati!��Ժ & "") = 1 Then
                    strCaption = "��Ժ"
                    strSQL = "Select Distinct b.����id, b.��ҳid, b.��Ŀ����, b.סԺ��,b.���ۺ�,b.����, b.�Ա�, a.����, b.��ͥ��ַ, a.���֤��, a.��������, b.��Ժ����, b.��Ժ����, b.סԺҽʦ,B.����ת��, b.��������, b.��������," & vbNewLine & _
                            "       Decode(d.����id || '_' || d.��ҳid, '_', 0, 1) As �Ƿ��ӡ, b.��Ժ����id as ����ID, B.סԺҽʦ,a.���￨��,LPAD(B.��Ժ����," & intBedLen & ",' ') as ���� " & vbNewLine & _
                            "From ������Ϣ A, ������ҳ B, ������ӡ��¼ D" & vbNewLine & _
                            "Where a.����id = b.����id And b.����id = d.����id(+) And b.��ҳid = d.��ҳid(+) And b.����id = [1] And b.��ҳid = [2]"
                    
                Else
                    strCaption = "��Ժ"
                    strSQL = "Select Distinct b.����id, b.��ҳid, b.��Ŀ����,b.סԺ��,b.���ۺ�,b.����, b.�Ա�, a.����, b.��ͥ��ַ, a.���֤��, a.��������, b.��Ժ����, b.��Ժ����, b.סԺҽʦ,B.����ת��, b.��������, b.��������," & vbNewLine & _
                            "  Decode(d.����id || '_' || d.��ҳid, '_', 0, 1) As �Ƿ��ӡ,b.��Ժ����id as ����ID, b.סԺҽʦ,a.���￨��,LPAD(B.��Ժ����," & intBedLen & ",' ') as ���� " & vbNewLine & _
                            "From ������Ϣ A, ������ҳ B, ������ӡ��¼ D" & vbNewLine & _
                            "Where a.����id = b.����id And b.����id = d.����id(+) And b.��ҳid = d.��ҳid(+) And b.����id = [1] And b.��ҳid = [2]"
                    
                End If
                
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatiID, Val(rsPati!����id & ""))
                If rsPati.RecordCount > 0 Then
                    If InStr(mstrPrivs, "���Ʋ���") = 0 And InStr(mstrPrivs, "ȫԺ����") = 0 Then
                        rsPati.Filter = "סԺҽʦ ='" & UserInfo.���� & "'"
                        If rsPati.RecordCount = 0 Then
                            MsgBox "�û���" & UserInfo.���� & "��Ȩ�޲���,����������ò��ˡ�" & rsPati!���� & "����", vbInformation, gstrSysName
                            Exit Function
                        End If
                    ElseIf InStr(mstrPrivs, "ȫԺ����") = 0 Then
                        strSQL = "Select 1 From ������Ա Where ����id = [1] And ��Աid = [2]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, NVL(rsPati!��Ժ����ID), UserInfo.ID)
                        If rsTemp.RecordCount = 0 Then
                            MsgBox "�û���" & UserInfo.���� & "��Ȩ�޲���,����������ò��ˡ�" & rsPati!���� & "����", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
            Else
                mbytType = E_����
                If Val(rsPati!ִ��״̬ & "") = 1 Then '����
                    strCaption = "����"
                ElseIf Val(rsPati!ִ��״̬ & "") = 2 Then '���ھ���
                    strCaption = "����"
                End If
                strSQL = "Select Distinct b.Id, b.No, b.����id, b.�����, b.����, b.�Ա�, b.����, b.ִ��ʱ��, b.ִ����, b.ִ�в���id, a.��ͥ��ַ, a.��������, a.���֤��, a.��������,a.���￨��" & vbNewLine & _
                        "From ������Ϣ A, ���˹Һż�¼ B" & vbNewLine & _
                        "Where b.����id = a.����id And b.Id = [1]"
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsPati!����id & ""))
                If rsPati.RecordCount > 0 Then
                    If InStr(mstrPrivs, "���Ʋ���") = 0 And InStr(mstrPrivs, "ȫԺ����") = 0 Then
                        rsPati.Filter = "ִ���� ='" & UserInfo.���� & "'"
                        If rsPati.RecordCount = 0 Then
                            MsgBox "�û���" & UserInfo.���� & "��Ȩ�޲���,����������ò��ˡ�" & rsPati!���� & "����", vbInformation, gstrSysName
                            Exit Function
                        End If
                    ElseIf InStr(mstrPrivs, "ȫԺ����") = 0 Then
                        strSQL = "Select 1 From ������Ա Where ����id = [1] And ��Աid = [2]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, NVL(rsPati!ִ�в���ID), UserInfo.ID)
                        If rsTemp.RecordCount = 0 Then
                            MsgBox "�û���" & UserInfo.���� & "��Ȩ�޲���,����������ò��ˡ�" & rsPati!���� & "����", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
            End If
 
            For i = 0 To tbcPati.ItemCount - 1
                If strCaption = tbcPati.Item(i).Tag Then
                    tbcPati.Item(i).Visible = True
                    mblnUndo = True
                    tbcPati.Item(i).Selected = True 'ȱʡѡ������
                    mblnUndo = False
                Else
                    tbcPati.Item(i).Visible = False
                End If
            Next
        Else
            Exit Function
        End If
    Else
        If mbytType = E_סԺ Then
            '������ҳ.״̬��0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ
            If mintDeptView = 0 Then
                '��Ժ����
                If tbcPati.Selected.Tag = "��Ժ" Then
                    strSQL = "Select Distinct b.����id, b.��ҳid, b.��Ŀ����, b.סԺ��,b.���ۺ�, Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�, NVL(b.����,a.����) As ����, b.��ͥ��ַ, a.���֤��, a.��������, b.��Ժ����," & vbNewLine & _
                                "       b.��Ժ����, b.סԺҽʦ,b.��������, B.����ת��, b.��������, Decode(d.����id, Null, 0, 1) As �Ƿ��ӡ, R.����ID,a.���￨��,LPAD(B.��Ժ����," & intBedLen & ",' ') as ����  " & vbNewLine & _
                                "From ������Ϣ A,������ҳ B,������ӡ��¼ D, ��Ժ���� R" & vbNewLine & _
                                "Where r.����id = d.����id(+) And r.��ҳid = d.��ҳid(+) And r.����id = a.����id And r.����id = b.����id " & vbNewLine & _
                                " And r.��ҳid = b.��ҳid And (R.����ID=[1] Or b.Ӥ������ID=[1]) And b.״̬<>1" & _
                                    IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And b.סԺҽʦ=[2]")
                ElseIf tbcPati.Selected.Tag = "��Ժ" Then
                    strFilter = " And B.��Ժ���� Between to_date('" & Format(mdatOutBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And to_date('" & Format(mdatOutEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS') "
                    
                    strSQL = "Select Distinct b.����id, b.��ҳid, b.��Ŀ����, b.סԺ��,b.���ۺ�, Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�, NVL(b.����,a.����) As ����, b.��ͥ��ַ, a.���֤��, a.��������, b.��Ժ����," & vbNewLine & _
                            "       b.��Ժ����, b.סԺҽʦ, b.��������, B.����ת��,b.��������, Decode(d.����id, Null, 0, 1) As �Ƿ��ӡ, B.��Ժ����ID As ����ID,a.���￨��,LPAD(B.��Ժ����," & intBedLen & ",' ') as ����  " & vbNewLine & _
                            "From ������Ϣ A, ������ҳ B, ������ӡ��¼ D" & vbNewLine & _
                            "Where a.����id = b.����id And Nvl(b.��ҳid, 0) <> 0 And a.����id = d.����id(+) And a.��ҳid = d.��ҳid(+) And b.��Ժ����id + 0 = [1]" & vbNewLine & _
                            IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And b.סԺҽʦ=[2]") & vbNewLine & _
                            " And b.���ʱ�� Is Null " & strFilter
                    
                End If
            Else
                '�������鿴
                '��Ժ����
                If tbcPati.Selected.Tag = "��Ժ" Then
                    strSQL = "Select Distinct b.����id, b.��ҳid, b.��Ŀ����, b.סԺ��,b.���ۺ�, Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�, NVL(b.����,a.����) As ����, b.��ͥ��ַ, a.���֤��, a.��������, b.��Ժ����," & vbNewLine & _
                                "       b.��Ժ����, b.סԺҽʦ,b.��������,b.����ת��,b.��������, Decode(d.����id, Null, 0, 1) As �Ƿ��ӡ, R.����ID,a.���￨��,LPAD(B.��Ժ����," & intBedLen & ",' ') as ����  " & vbNewLine & _
                                "From ������Ϣ A,������ҳ B,���ű� C,������ӡ��¼ D, ��Ժ���� R" & vbNewLine & _
                                "Where r.����id = d.����id(+) And r.��ҳid = d.��ҳid(+) And r.����id = a.����id And r.����id = b.����id " & vbNewLine & _
                                " And r.��ҳid = b.��ҳid And (R.����ID=[1] Or b.Ӥ������ID=[1]) And b.״̬<>1" & _
                                " And B.��Ժ����ID=C.ID And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null) " & vbNewLine & _
                                IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And b.סԺҽʦ=[2]") & vbNewLine & _
                                " And b.���ʱ�� Is Null " & strFilter
                     
                Else
                    strFilter = " And B.��Ժ���� Between to_date('" & Format(mdatOutBegin, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And to_date('" & Format(mdatOutEnd, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS') "
                    strSQL = "Select Distinct b.����id, b.��ҳid, b.��Ŀ����, b.סԺ��,b.���ۺ�,Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�, NVL(b.����,a.����) As ����, b.��ͥ��ַ, a.���֤��, a.��������, b.��Ժ����," & vbNewLine & _
                            "       b.��Ժ����, b.סԺҽʦ, b.��������, b.��������,b.����ת��, Decode(d.����id, Null, 0, 1) As �Ƿ��ӡ, B.��Ժ����ID As ����ID,a.���￨��,LPAD(B.��Ժ����," & intBedLen & ",' ') as ���� " & vbNewLine & _
                            "From ������Ϣ A, ������ҳ B,���ű� C, ������ӡ��¼ D" & vbNewLine & _
                            "Where a.����id = b.����id And Nvl(b.��ҳid, 0) <> 0 And a.����id = d.����id(+) And a.��ҳid = d.��ҳid(+) And B.��ǰ����ID+0=[1] " & vbNewLine & _
                            " And B.��Ժ����ID=C.ID And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null) " & vbNewLine & _
                             IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And b.סԺҽʦ=[2]") & vbNewLine & _
                            " And b.���ʱ�� Is Null " & strFilter
        
                End If
            End If
        Else
            '���ﲡ��
            If tbcPati.Selected.Tag = "����" Then
                strSQL = "Select B.ID, b.No, b.����id, b.�����, b.����, b.�Ա�, b.����, b.ִ��ʱ��, b.ִ����, b.ִ�в���ID, a.��ͥ��ַ, a.��������, a.���֤��, a.��������,a.���￨�� " & vbNewLine & _
                    "From ������Ϣ A, ���˹Һż�¼ B" & vbNewLine & _
                    "Where b.����id = a.����id And b.��¼���� = 1 And b.��¼״̬ = 1 And b.ִ��״̬ = 2 And b.ִ�в���ID = [1]" & vbNewLine & _
                    IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And b.ִ����=[2]")
            ElseIf tbcPati.Selected.Tag = "����" Then
                strSQL = "Select B.ID, b.No, b.����id, b.�����, b.����, b.�Ա�, b.����, b.ִ��ʱ��, b.ִ����, b.ִ�в���ID, a.��ͥ��ַ, a.��������, a.���֤��, a.��������,a.���￨�� " & vbNewLine & _
                    "From ������Ϣ A, ���˹Һż�¼ B" & vbNewLine & _
                    "Where b.����id = a.����id And b.��¼���� = 1 And b.��¼״̬ = 1 And b.ִ��״̬ = 1 And b.ִ�в���ID = [1]" & vbNewLine & _
                    IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And b.ִ����=[2]")
                strSQL = strSQL & " And B.ִ��ʱ�� Between To_Date('" & Format(mDatBegin, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(mDatEnd, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
            End If
        End If
        If strSQL = "" Then
            rptPati.Tag = ""
            rptPati.Records.DeleteAll
            rptPati.Populate
            Screen.MousePointer = 0
            LoadPatients = True
            Exit Function
        End If
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDept.ItemData(cboDept.ListIndex), UserInfo.����)
    End If

    '���ز����б�
    Call ClearPatiInfo
    rptPati.Tag = ""
    rptPati.Records.DeleteAll
    mintPatiCount = 0
    stbThis.Panels(2).Text = ""
     
    For i = 1 To rsPati.RecordCount
        Set objRecord = rptPati.Records.Add()
        Set objItem = objRecord.AddItem("")    'ͼ��
        objItem.HasCheckbox = True
        Set objItem = objRecord.AddItem("")  'ͼ1��
        If InStr(rsPati!�Ա� & "", "��") > 0 Then
            objItem.Icon = imgPati.ListImages("Boy").Index - 1
        ElseIf InStr(rsPati!�Ա� & "", "Ů") > 0 Then
            objItem.Icon = imgPati.ListImages("Girl").Index - 1
        End If
        

        If tbcPati.Selected.Tag = "��Ժ" Then
            Set objItem = objRecord.AddItem("")  'ͼ��
            If Val(rsPati!�Ƿ��ӡ & "") = 1 Then
                objItem.Icon = imgPati.ListImages("print").Index - 1
            End If
            objRecord.AddItem IIf(NVL(rsPati!��Ŀ����) <> "", "�ѱ�Ŀ", "δ��Ŀ")
            objRecord.AddItem Format(rsPati!��Ŀ���� & "", "YYYY-MM-dd")
        Else
            objRecord.AddItem ""   'ͼ��
            objRecord.AddItem ""
            objRecord.AddItem ""
        End If
        
        If mbytType = E_סԺ Then
            objRecord.AddItem rsPati!סԺ�� & ""
            objRecord.AddItem ""
            objRecord.AddItem ""
            Set objItem = objRecord.AddItem(zlStr.Lpad(NVL(rsPati!����), 10)) 'Value��������
            objItem.Caption = CStr(Trim(NVL(rsPati!����, " "))) 'Ϊ��ʱ�ᱻValue���
        Else
            objRecord.AddItem ""
            objRecord.AddItem rsPati!NO & ""
            objRecord.AddItem rsPati!����� & ""
            objRecord.AddItem ""
        End If
        
        objRecord.AddItem rsPati!���� & ""
        objRecord.AddItem rsPati!�Ա� & ""
        objRecord.AddItem rsPati!���� & ""
        objRecord.AddItem rsPati!���֤�� & ""
        objRecord.AddItem Format(rsPati!�������� & "", "YYYY-MM-DD")
        If mbytType = E_סԺ Then
            objRecord.AddItem ""
        Else
            objRecord.AddItem Format(rsPati!ִ��ʱ�� & "", "YYYY-MM-DD")
        End If
       
        If mbytType = E_סԺ Then
            objRecord.AddItem Format(rsPati!��Ժ���� & "", "YYYY-MM-DD")
            objRecord.AddItem Format(rsPati!��Ժ���� & "", "YYYY-MM-DD")
            objRecord.AddItem rsPati!סԺҽʦ & ""
        Else
            objRecord.AddItem ""
            objRecord.AddItem ""
            objRecord.AddItem rsPati!ִ���� & ""
        End If
        objRecord.AddItem rsPati!��ͥ��ַ & ""
        objRecord.AddItem rsPati!���￨�� & ""
        If mbytType = E_סԺ Then
            objRecord.AddItem rsPati!���ۺ� & ""
        Else
            objRecord.AddItem ""
        End If
        '������
        objRecord.AddItem rsPati!�������� & ""
        objRecord.AddItem CLng(rsPati!����ID)

        If mbytType = E_סԺ Then
            objRecord.AddItem NVL(rsPati!��ҳID)
            objRecord.AddItem rsPati!����ID & ""
            objRecord.AddItem rsPati!����ת�� & ""
        Else
            objRecord.AddItem NVL(rsPati!ID)
            objRecord.AddItem rsPati!ִ�в���ID & ""
            objRecord.AddItem ""
        End If
         '��ʾ������ɫ
        lngColor = zlDatabase.GetPatiColor(NVL(rsPati!��������))
        objRecord.Item(col_����).ForeColor = lngColor

        rsPati.MoveNext
    Next
    Call ShowReportColumn
    rptPati.Populate
    If chkFilter.Value = vbChecked Then
        If rptPati.Records.Count > 0 Then Set rptPati.FocusedRow = rptPati.Rows(0)
    Else
        Call InitBasicData '����ļ��б�
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowOutPatiInfo() As Boolean
'���ܣ�ѡ�����ﲡ��ĳ����ʷ�����¼ʱ����ȡ��صĲ�����Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    
    On Error GoTo errH
    fraOutPati.Visible = True: fraInPati.Visible = False
    If mlng����ID <> 0 Then
        strSQL = "Select B.Id,B.NO,B.�����,B.����,B.�Ա�,B.����,A.ҽ�Ƹ��ʽ,A.���֤��,A.��ͥ��ַ," & _
            " A.�ѱ�,A.����,A.��������,A.ҽ����,B.����,B.����ʱ��,B.ִ����,B.ִ��״̬,B.ִ��ʱ��," & _
            " B.ִ�в���ID as ����ID,B.����,B.����,D.������,C.���� as ����" & _
            " From ������Ϣ A,���˹Һż�¼ B,���ű� C,����������Ϣ D" & _
            " Where A.����ID=B.����ID And B.ID=[1] And B.ִ�в���ID=C.ID" & _
            " And B.����ID=D.����ID(+) And B.����=D.����(+) And B.��¼����=1 And B.��¼״̬=1"
        If mblnMoved Then
            strSQL = Replace(strSQL, "���˹Һż�¼", "H���˹Һż�¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        With rsTmp
            '���ղ���������ɫ��ʾ
            lblShow(lblOUT_����).Caption = NVL(!����)
            
            If Not IsNull(!����) Then
                lblShow(lblOUT_����).ForeColor = vbRed
            Else
                lblShow(lblOUT_����).ForeColor = lblShow(lblOUT_�����).ForeColor
            End If
            lblShow(lblOUT_�����).Caption = NVL(!�����)
            lblShow(lblOUT_����).Caption = NVL(!����)
            lblShow(lblOUT_�Ա�).Caption = NVL(!�Ա�)
            lblShow(lblOUT_���֤��).Caption = NVL(!���֤��)
            lblShow(lblOUT_��������).Caption = NVL(!ִ��ʱ��)
            lblShow(lblOUT_��ͥ��ַ).Caption = NVL(!��ͥ��ַ)
            lblShow(lblOUT_��������).Caption = NVL(!��������)
            lblShow(lblOUT_����ҽʦ).Caption = NVL(!ִ����)
            lbl��.Visible = NVL(!����, 0) <> 0
            mlng����ID = NVL(!����ID, 0)
            mlng����ID = 0
        End With
    Else
        Call ClearPatiInfo
    End If
    ShowOutPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ClearPatiInfo()
'����:���������Ϣ��ʾ��
    Dim i As Long
    mlng����ID = 0
    mlng����ID = 0
    mlng����ID = 0
    mlng����ID = 0
    Set imgPatient.Picture = imgList.ListImages("Patient").Picture
    For i = lblShow.LBound To lblShow.UBound
        lblShow(i).Caption = ""
    Next

End Sub

Private Function ShowInPatiInfo() As Boolean
'���ܣ�ѡ��ĳ��סԺ��¼ʱ����ȡ��صĲ�����Ϣ
'���أ�blnMoved=����סԺ�����Ƿ�ת����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    
    On Error GoTo errH
    fraInPati.Visible = True: fraOutPati.Visible = False
    If mlng����ID <> 0 Then
        strSQL = "Select NVL(B.����,A.����) ����, NVL(B.�Ա�,A.�Ա�) �Ա�, NVL(B.����,A.����) ����," & _
            " NVL(B.��ͥ��ַ,A.��ͥ��ַ) As ��ͥ��ַ, B.סԺ��,B.��Ժ����,B.ҽ�Ƹ��ʽ," & _
            " D.��Ϣֵ as ҽ����,B.����,B.��ǰ����,C.���� as ����ȼ�,B.��Ժ����," & _
            " B.��Ժ����,B.��������,B.״̬,B.��Ժ����ID,B.��ǰ����ID,A.סԺ����,A.���֤�� " & _
            " From ������Ϣ A,������ҳ B,�շ���ĿĿ¼ C,������ҳ�ӱ� D" & _
            " Where A.����ID=B.����ID And A.����ID=[1] And B.��ҳID=[2] And B.����ȼ�ID=C.ID(+)" & _
            " And B.����ID=D.����ID(+) And B.��ҳID=D.��ҳID(+) And D.��Ϣ��(+)='ҽ����'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID)
    
        With rsTmp
            '���ղ�����ɫ������ʾ
            lblShow(lbl_����).Caption = NVL(!����)
            lblShow(lbl_����).ForeColor = zlDatabase.GetPatiColor(NVL(!��������))

            lblShow(lbl_סԺ��).Caption = NVL(!סԺ��)
            lblShow(lbl_�Ա�).Caption = NVL(!�Ա�)
            lblShow(lbl_����).Caption = NVL(!����)
            lblShow(lbl_���֤��).Caption = NVL(!���֤��)
            lblShow(lbl_��ͥ��ַ).Caption = NVL(!���֤��)
            'Σ�ز��˲�����ɫ��ʾ
            lblShow(lbl_����).Caption = NVL(!��ǰ����)
            If NVL(!��ǰ����) = "Σ" Or NVL(!��ǰ����) = "��" Or NVL(!��ǰ����) = "��" Then
                lblShow(lbl_����).ForeColor = vbRed
            Else
                lblShow(lbl_����).ForeColor = lblList(lbl_סԺ��).ForeColor
            End If
            lblShow(lbl_��Ժ����).Caption = Format(!��Ժ����, "yyyy-MM-dd HH:mm")
            If Not IsNull(!��Ժ����) Then
                lblShow(lbl_��Ժ����).Caption = lblShow(lbl_��Ժ����).Caption & "��" & Format(!��Ժ����, "yyyy-MM-dd HH:mm")
            End If
            mlng����ID = NVL(!��Ժ����ID, 0)
            mlng����ID = NVL(!��ǰ����ID, 0)
        End With
    Else
        Call ClearPatiInfo
    End If
    ShowInPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetEmrCISStruct(ByVal lngPatiID As Long, ByVal lngPageID As Long) As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strExtendTag As String, strReturn As String, strSQL As String, strSQLNew As String
    
    On Error GoTo errH
    If gobjEmr Is Nothing Then Set GetEmrCISStruct = Nothing: Exit Function
    strExtendTag = GetEMRIn_Tag(lngPatiID, lngPageID)
    If strExtendTag = "" Then Set GetEmrCISStruct = Nothing: Exit Function
    
    '�ϼ�ID��ID�����ƣ�������ͼ��
    strSQL = "Select Decode(e.Kind, '01', 'R1', '02', 'R2', '03', 'R4', '04', 'R5', '05', 'R6', 'R2') �ϼ�id," & vbNewLine & _
            "       Nvl(d.Subdoc_Id, Rawtohex(d.Real_Doc_Id)) As ID, d.Subdoc_Id As ���ĵ�id," & vbNewLine & _
            "       e.Title ||" & vbNewLine & _
            "        Decode(d.Completor, Null, ''," & vbNewLine & _
            "               '�� ' || d.Completor || ' ��' || To_Char(d.Complete_Time, 'yyyy-mm-dd hh24:mi') || 'ǩ����') As ����," & vbNewLine & _
            "       Rawtohex(d.Real_Doc_Id) || Decode(d.Subdoc_Id, Null, ';', ';' || d.Subdoc_Id) || ';' ||Nvl(d.Subdoc_Title, E.Title) As ����, 'object_case' As ͼ��" & vbNewLine & _
            "From (Select Distinct d.Id, c.Antetype_Id, c.Subdoc_Id, c.Subdoc_Title, c.Real_Doc_Id, c.Complete_Time, c.Completor" & vbNewLine & _
            "       From Bz_Act_Log A, Bz_Act_Log D, Bz_Doc_Tasks C" & vbNewLine & _
            "       Where a.Extend_Tag = :etag And (a.Id = d.Id Or a.Id = d.Basiclog_Id) And d.Id = c.Actlog_Id And" & vbNewLine & _
            "             c.Real_Doc_Id Is Not Null) D, Antetype_List E" & vbNewLine & _
            "Where d.Antetype_Id = e.Id  And e.Title = Decode(e.Type, 3, d.Subdoc_Title, e.Title)" & vbNewLine & _
            "Order By Rawtohex(d.Real_Doc_Id), e.Code, d.Complete_Time"
            
    strSQLNew = "Select Decode(e.Kind, '01', 'R1', '02', 'R2', '03', 'R4', '04', 'R5', '05', 'R6', 'R2') �ϼ�id," & vbNewLine & _
                "       Nvl(d.Subdoc_Id, Rawtohex(d.Real_Doc_Id)) As ID, d.Subdoc_Id As ���ĵ�id," & vbNewLine & _
                "       e.Title ||" & vbNewLine & _
                "        Decode(d.Completor, Null, ''," & vbNewLine & _
                "               '�� ' || d.Completor || ' ��' || To_Char(d.Complete_Time, 'yyyy-mm-dd hh24:mi') || 'ǩ����') As ����," & vbNewLine & _
                "       Rawtohex(d.Real_Doc_Id) || Decode(d.Subdoc_Id, Null, ';', ';' || d.Subdoc_Id) || ';' ||Nvl(d.Subdoc_Title, E.Title) As ����, 'object_case' As ͼ��" & vbNewLine & _
                "From (Select Distinct d.Id, c.Antetype_Id, c.Subdoc_Id, c.Subdoc_Title, c.Real_Doc_Id, c.Complete_Time, c.Completor, c.Order_No" & vbNewLine & _
                "       From Bz_Act_Log A, Bz_Act_Log D, Bz_Doc_Tasks C" & vbNewLine & _
                "       Where a.Extend_Tag = :etag And (a.Id = d.Id Or a.Id = d.Basiclog_Id) And d.Id = c.Actlog_Id And" & vbNewLine & _
                "             c.Real_Doc_Id Is Not Null And Nvl(c.Intead, 0) = 0) D, Antetype_List E" & vbNewLine & _
                "Where d.Antetype_Id = e.Id " & vbNewLine & _
                "Order By Rawtohex(d.Real_Doc_Id), e.Code, d.Order_No"
    
    err.Clear
    On Error Resume Next
    strReturn = gobjEmr.OpenSQLRecordset(strSQLNew, strExtendTag & "^16^etag", rsTemp)
    If err.Number <> 0 Or strReturn <> "" Then
        err.Clear
        strReturn = gobjEmr.OpenSQLRecordset(strSQL, strExtendTag & "^16^etag", rsTemp)
    End If
    
    If strReturn <> "" Then
        MsgBox strReturn, vbCritical, gstrSysName
        Set GetEmrCISStruct = Nothing: Exit Function
    End If
    
    Set GetEmrCISStruct = rsTemp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetEMRIn_Tag(ByVal lngPatiID As Long, ByVal lngPageID As Long) As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    If mbytType = E_���� Then
        GetEMRIn_Tag = "MZ_" & mlng����ID
    Else
        strSQL = "Select Nvl(a.Id, b.Id) ID" & vbNewLine & _
                    "From (Select Max(ID) ID From ���˱䶯��¼ Where ����id = [1] And ��ҳid = [2] And ��ʼԭ�� = 2 And Nvl(���Ӵ�λ, 0) = 0) A," & vbNewLine & _
                    "     (Select Max(ID) ID From ���˱䶯��¼ Where ����id = [1] And ��ҳid = [2] And ��ʼԭ�� = 1 And Nvl(���Ӵ�λ, 0) = 0) B"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������ԺID", lngPatiID, lngPageID)
        
        If rsTmp Is Nothing Then Exit Function
        If NVL(rsTmp!ID) = "" Then Exit Function
        GetEMRIn_Tag = "BD_" & rsTmp!ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetLisRptFile(ByVal strTag As String, Optional ByVal strFile As String) As String
'���ܣ���LIS�����ļ��鿴����ȡ��ʱ�ļ�·��
'
    Dim lngRetu As Long, strInfo As String
    Dim objFile As New FileSystemObject
    Dim strTmp As String
    Dim lng����ID As String
    Dim str������ As String
    Dim lng���� As String
    Dim varTmp As Variant
    Dim strSuffix As String '�ļ���׺��
    
    Screen.MousePointer = 11
    
    varTmp = Split(strTag, ";")
    lng����ID = varTmp(0)
    strTmp = Replace(strTag, varTmp(0) & ";" & varTmp(1) & ";", "")
    varTmp = Split(strTmp, "<sTab>")
    lng���� = varTmp(0)
    If lng���� = 0 Then
        strSuffix = "pdf"
    ElseIf lng���� = 1 Then
        strSuffix = "html"
    Else
        strSuffix = "xps"
    End If
    str������ = varTmp(1) & "_" & lng����ID
    If strFile = "" Then
        strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\tmpReport_" & lng����ID & "." & strSuffix
    End If
    If Not objFile.FileExists(strFile) Then
        strFile = Sys.ReadLob(glngSys, 22, lng����ID, strFile)
        If Not objFile.FileExists(strFile) Then
            MsgBox "�ļ����ݶ�ȡʧ�ܣ�", vbInformation, gstrSysName:
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    GetLisRptFile = strFile
    Screen.MousePointer = 0
End Function

Private Sub picRpt_Resize()
    On Error Resume Next
    webRpt.Move 0, 0, picRpt.Width, picRpt.Height
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
    Set rsTmp = GetDataToDepts()
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

Private Function GetDataToDepts(Optional ByVal strIn As String = "") As ADODB.Recordset
'���ܣ���ȡ���Ҳ����б����ݼ�¼��
'������strIn ��������
'      bytFunc=0 סԺ;1-����
    Dim strSQL As String
    Dim blnYN As Boolean
    Dim strDeptIDs As String
    
    If strIn <> "" Then blnYN = True

    If mbytType = E_סԺ Then
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
    ElseIf mbytType = E_���� Then
        If InStr(mstrPrivs, "ȫԺ����") > 0 Then
            strSQL = _
                    " Select Distinct A.ID,A.����,A.����" & _
                    " From ���ű� A,��������˵�� B" & _
                    " Where B.����ID=A.ID And B.��������='�ٴ�' And B.������� In(1,3) " & _
                    " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    IIf(blnYN, " And (A.���� Like [2] Or A.���� Like [3] Or A.���� Like [3])", "") & _
                    " Order by A.����"
        Else
            strSQL = "Select Distinct B.ID,B.����,B.����,A.ȱʡ" & _
                " From ������Ա A,���ű� B,��������˵�� C" & _
                " Where A.����ID=B.ID And B.ID=C.����ID And C.������� In(1,3) And C.��������='�ٴ�'" & _
                " And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is Null)" & _
                " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) And A.��ԱID=[1]" & _
                IIf(blnYN, " And (B.���� Like [2] Or B.���� Like [3] Or B.���� Like [3])", "") & _
                " Order by B.����"
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

Private Sub InitSelectTime()
    Dim datCurr As Date
     
   datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
   mdatOutEnd = datCurr
   mdatOutBegin = mdatOutEnd - 1
   mDatEnd = datCurr
   mDatBegin = mDatEnd - 1
   
   cboSelectTime(0).Clear '��Ժ
   With cboSelectTime(0)
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
   If cboSelectTime(0).ListCount > 0 Then cboSelectTime(0).ListIndex = 0

   cboSelectTime(1).Clear
   With cboSelectTime(1)
       .AddItem "����"
       .ItemData(.NewIndex) = 0
       .AddItem "����(������)"
       .ItemData(.NewIndex) = 1
       .AddItem "һ����"
       .ItemData(.NewIndex) = 7
       .AddItem "[ָ��...]"
       .ItemData(.NewIndex) = -1
   End With
   If cboSelectTime(1).ListCount > 0 Then cboSelectTime(1).ListIndex = 0

End Sub

Private Sub CheckNode(ByVal node As MSComctlLib.node, Optional ByVal bytFunc As Byte = 0)
'����:���ڵ㹴ѡ|��ѡ���ӽڵ��Ӧ��ѡ|��ѡ��
'   �ӽڵ����ж�����ѡ���򸸽ڵ�Ҳ��ѡ��
'   �ӽڵ�ֻҪһ������ѡ�����ڵ�ҲĬ�ϲ���ѡ��
'����:Node-��ǰ���
    Dim objNode As MSComctlLib.node
    If bytFunc = 0 Then
        If node.Children > 0 Then
            If Not node.Child Is Nothing Then
                Set objNode = node.Child
                objNode.Checked = node.Checked
                If objNode.Children > 0 Then Call CheckNode(objNode)
                Do While Not objNode.Next Is Nothing
                    Set objNode = objNode.Next
                    objNode.Checked = node.Checked
                    If objNode.Children > 0 Then Call CheckNode(objNode)
                Loop
            End If
        End If

        If Not node.Parent Is Nothing Then Call CheckNode(node, 1)
    Else
        '�ӽڵ����һ����ѡ,���ڵ㲻��ѡ;��ǰ�ڵ�������ڸ��ڵ㣬������Ҷ�ӽڵ㶼ѡ��,��ô��Ӧ�ĸ��ڵ�Ҳѡ��
        If node.Checked = False Then
            If node.Parent.Checked = True Then
                node.Parent.Checked = False
                If Not node.Parent.Parent Is Nothing Then Call CheckNode(node.Parent, 1)
            End If
        Else
            If node.Parent.Checked = False Then
                Set objNode = node.FirstSibling
                Do While Not objNode Is Nothing
                    If objNode.Checked = False Then Exit Sub
                    Set objNode = objNode.Next
                Loop
                node.Parent.Checked = True
                If Not node.Parent.Parent Is Nothing Then Call CheckNode(node.Parent, 1)
            End If
        End If
    End If
End Sub

Private Sub FuncLoadReport()
    Dim objControl As CommandBarControl
    Dim objPop As Object
    Dim strHide As String
    Dim i As Long
    
    strHide = ",ZL1_INSIDE_1254_1,ZL1_INSIDE_1254_2,ZL1_INSIDE_1261_1,ZL1_INSIDE_1261_4,ZL1_INSIDE_1261_5,ZL1_INSIDE_1261_6,ZL1_INSIDE_1261_7,ZL1_INSIDE_1261_8,ZL1_INSIDE_1261_9,ZL1_INSIDE_1261_10,"
    mstr����Ӧ���� = zlDatabase.GetPara("����Ӧ����", glngSys, P�������Ĵ�ӡ)
    mstr�����Ӧ���� = zlDatabase.GetPara("�����Ӧ����", glngSys, P�������Ĵ�ӡ)
    If mstr����Ӧ���� <> "" Then strHide = strHide & "," & Split(mstr����Ӧ����, ",")(2) & ","
    If mstr�����Ӧ���� <> "" Then strHide = strHide & "," & Split(mstr�����Ӧ����, ",")(2) & ","
    mstr���鱨���ӡ = zlDatabase.GetPara("���鱨���ӡ", glngSys, P�������Ĵ�ӡ)
    '��ջ���
    Set mcolReport = New Collection
    For i = 1 To cbsMain.ActiveMenuBar.Controls.Count
        If cbsMain.ActiveMenuBar.Controls(i).ID = conMenu_ReportPopup _
            Or cbsMain.ActiveMenuBar.Controls(i).Caption Like "����*" Then
                cbsMain.ActiveMenuBar.Controls.Item(i).Delete
            Exit For
        End If
    Next
    
    Call zlDatabase.ShowReportMenu(cbsMain, glngSys, P�������Ĵ�ӡ, mstrPrivs, strHide)
    
    For i = 1 To cbsMain.ActiveMenuBar.Controls.Count
        If cbsMain.ActiveMenuBar.Controls(i).ID = conMenu_ReportPopup _
            Or cbsMain.ActiveMenuBar.Controls(i).Caption Like "����*" Then
            Set objControl = cbsMain.ActiveMenuBar.Controls.Item(i)
            Exit For
        End If
    Next
    
    If Not objControl Is Nothing Then
        With objControl.CommandBar.Controls
            For i = 1 To .Count
                Set objPop = .Item(i)
                mcolReport.Add Split(objPop.Caption, "(&")(0) & "," & objPop.Parameter, "_" & i     '��������,ϵͳ��,������
            Next
        End With
    End If
End Sub

Private Sub SetPrintPara()
    Dim i As Long
    Dim objFrm As New frmParaSet
    Dim lngRow As Long
    
    Call objFrm.ShowMe(Me, glngSys, P�������Ĵ�ӡ, mstrPrivs)
    '���¼���
    Call FuncLoadReport
End Sub

Private Function RecordEprPrintInfo(ByVal bytMode As Byte, ByVal strRecordKey As String, ByVal lngNo As Long, Optional ByVal lngPatientKey As Long, Optional ByVal lngPatientPageKey As Long) As Boolean
    
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    
    If Not mblnPrint Then Exit Function
    
    If lngNo = 0 Then
        lngNo = 1
        strSQL = "Select Nvl(Max(��ӡ����),0)+1 As ��ӡ���� From ������ӡ��¼ Where ����id=[1] And ��ҳid=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISAudit", lngPatientKey, lngPatientPageKey)
        If rsTmp.BOF = False Then
            lngNo = rsTmp("��ӡ����").Value
        End If
    End If
    
    Select Case bytMode
    Case 1
        strSQL = "Select ����id,��ҳid,�������� From�����Ӳ�����¼ a Where a.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlCISAudit", Val(strRecordKey))
        If rs.BOF = False Then
            strSQL = "Zl_������ӡ��¼_Insert(" & Val(rs("����id").Value) & "," & Val(rs("��ҳid").Value) & "," & lngNo & ",'" & rs("��������").Value & "','" & UserInfo.���� & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "mdlCISAudit")
        End If
    Case 2
        strSQL = "Zl_������ӡ��¼_Insert(" & lngPatientKey & "," & lngPatientPageKey & "," & lngNo & ",'" & strRecordKey & "','" & UserInfo.���� & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "mdlCISAudit")
    Case 3
        strSQL = "Select ���� From�������ļ��б� a Where a.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlCISAudit", Val(strRecordKey))
        If rs.BOF = False Then
            strSQL = "Zl_������ӡ��¼_Insert(" & lngPatientKey & "," & lngPatientPageKey & "," & lngNo & ",'" & rs("����").Value & "','" & UserInfo.���� & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "mdlCISAudit")
        End If
    End Select
    
    RecordEprPrintInfo = True
    
End Function

Private Sub SelectItems(ByVal bytFunc As Byte)
'����:
'   bytFunc=1 ȫѡ,=2ȡ��ȫѡ
    Dim i As Long
    
    With rptPati
        For i = 0 To .Records.Count - 1
            If bytFunc = 1 Then
                .Records(i).Item(col_ѡ��).Checked = True
            Else
                .Records.Record(i).Item(0).Checked = False
            End If
        Next
        mintPatiCount = IIf(bytFunc = 1, .Records.Count, 0)
    End With
    stbThis.Panels(2).Text = IIf(mintPatiCount = 0, "", "��ѡ��" & mintPatiCount & "�����ˣ�")
End Sub

Private Function GetPrintLog(ByVal lngPatient As Long, ByVal lngPageID As Long) As String
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ��ӡ���� As ��ӡ��, ��ӡ����, ��ӡ��, ��ӡʱ�� From ������ӡ��¼ Where ����id = [1] And ��ҳid = [2] Order By ��ӡʱ��, ��ӡ���"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatient, lngPageID)
    Do Until rs.EOF
        
        GetPrintLog = GetPrintLog & vbCrLf & Rpad(rs!��ӡ��, 10) & Rpad(Format(rs!��ӡʱ��, "yyyy-mm-dd hh:MM"), 20) & Rpad(rs!��ӡ����, 40)
        rs.MoveNext
    Loop
    GetPrintLog = Rpad("��ӡ��", 10) & Rpad("��ӡʱ��", 20) & Rpad("��ӡ����", 40) & GetPrintLog
    
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    Rpad = zl9ComLib.zlStr.Rpad(strCode, lngLen, strChar, True)
End Function

Private Function GetCISStruct(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str�Һŵ� As String, ByVal blnPath As Boolean, ByVal blnDataMove As Boolean) As ADODB.Recordset
'����:lng��ҳID סԺ����Ϊ��ҳid,���ﲡ��Ϊ�Һ�id
    Dim strSQL As String, strSQL1 As String
    Dim rsTmp As ADODB.Recordset
    Dim rsMedRec As ADODB.Recordset
    Dim strRptIDs As String
    
    Dim i As Long
    
    On Error GoTo errH
    '1-���ﲡ��;2-סԺ����;3-�����¼;4-������;5-����֤��;6-֪���ļ�;7-���Ʊ���,11-��ҳ��Ϣ,12-ҽ����¼,8-�ٴ�·��;9-סԺ֤;10-��������
    strSQL = strSQL & _
        " Select 'R0' As ID, '' As �ϼ�id, '�����ļ�' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'00' As ���� From Dual Union All" & _
        " Select 'R11' As ID, 'R0' As �ϼ�id, '��ҳ��Ϣ' As ����, '' As ����,0 As ĩ��,'home' As ͼ��,'01' As ���� From Dual Union All" & _
        " Select 'R12' As ID, 'R0' As �ϼ�id, 'ҽ����¼' As ����, '' As ����,0 As ĩ��,'object_advice' As ͼ��,'02' As ���� From Dual Union All" & _
        " Select 'R1' As ID, 'R0' As �ϼ�id, '���ﲡ��' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'03' As ���� From Dual Where [3]=0 Union All" & _
        " Select 'R2' As ID, 'R0' As �ϼ�id, 'סԺ����' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'04' As ���� From Dual Where [3]=1 Union All" & _
        " Select 'R3' As ID, 'R0' As �ϼ�id, '�����¼' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'05' As ���� From Dual Where [3]=1 Union All" & _
        " Select 'R4' As ID, 'R0' As �ϼ�id, '������' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'06' As ���� From Dual Where [3]=1 Union All" & _
        " Select 'R7' As ID, 'R0' As �ϼ�id, '���Ʊ���' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'07' As ���� From Dual Union All" & _
        " Select 'R5' As ID, 'R0' As �ϼ�id, '����֤��' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'08' As ���� From Dual Union All" & _
        " Select 'R6' As ID, 'R0' As �ϼ�id, '֪���ļ�' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'09' As ���� From Dual "
    strSQL = strSQL & _
            IIf(blnPath, " Union All Select 'R8' As ID, 'R0' As �ϼ�id, '�ٴ�·��' As ����, '' As ����,0 As ĩ��,'Path' As ͼ��,'10' As ���� From Dual", "") & _
            IIf(str�Һŵ� = "", " Union All Select 'R9' As ID, 'R0' As �ϼ�id, 'סԺ֤' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'11' As ���� From Dual", "") & _
            IIf(str�Һŵ� = "", " Union All Select 'R10' As ID, 'R0' As �ϼ�id, '��������' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,'12' As ���� From Dual", "")
    If lng����ID = 0 Then
        strSQL = " Select * From (" & strSQL & ") Order By Decode(�ϼ�id,Null,' ',�ϼ�id),����"
        Set GetCISStruct = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 0, 0, IIf(mbytType = E_סԺ, 1, 0))
        Exit Function
    End If
    '��ҳ�����:��ҳ����,��ҳ����,��ҳ��ҳһ,��ҳ��ҳ��
    If str�Һŵ� = "" Then
        strSQL = strSQL & _
                " Union All Select 'R11K1' As ID, 'R11' As �ϼ�id, '��ҳ����' As ����, '' As ����,1 As ĩ��,'home' As ͼ��,'1' As ���� From Dual Where [3]=1 " & _
                " Union All Select 'R11K2' As ID, 'R11' As �ϼ�id, '��ҳ����' As ����, '' As ����,1 As ĩ��,'home' As ͼ��,'2' As ���� From Dual Where [3]=1 "
        If mintMecStandard = 1 Or mintMecStandard = 2 Then
            strSQL = strSQL & _
                  " Union All Select 'R11K3' As ID, 'R11' As �ϼ�id, '��ҳ��ҳһ' As ����, '' As ����,1 As ĩ��,'home' As ͼ��,'3' As ���� From Dual Where [3]=1 " & _
                  " Union All Select 'R11K4' As ID, 'R11' As �ϼ�id, '��ҳ��ҳ��' As ����, '' As ����,1 As ĩ��,'home' As ͼ��,'4' As ���� From Dual Where [3]=1 "
        End If
    End If
    'ҽ������
    If str�Һŵ� = "" Then
        strSQL = strSQL & " Union All Select 'R12K1' As ID, 'R12' As �ϼ�id, '��ʱҽ��' As ����, '' As ����,1 As ĩ��,'object_advice' As ͼ��,'1' As ���� From Dual"
        strSQL = strSQL & " Union All Select 'R12K2' As ID, 'R12' As �ϼ�id, '����ҽ��' As ����, '' As ����,1 As ĩ��,'object_advice' As ͼ��,'2' As ���� From Dual"
        'סԺ֤
        strSQL = strSQL & " Union All" & _
            " Select * From (Select �ϼ�id||'K'||ID as ID, �ϼ�id, ����, ����, ĩ��, ͼ��, ����" & vbNewLine & _
                "From (Select a.Id, 'R9' As �ϼ�id, c.����||'������ʱ�䣺'|| To_Char(D.����ʱ��, 'yyyy-mm-dd hh24:mi')||' �����ˣ�'||D.������||'��' as ����," & vbNewLine & _
                "  c.���|| ';' ||d.No || ';' || d.��¼���� As ����, 1 As ĩ��, 'Folder' As ͼ��, To_Char(d.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����" & vbNewLine & _
                "       From ����ҽ����¼ A, ����ҽ������ D, ��������Ӧ�� B, �����ļ��б� C" & vbNewLine & _
                "       Where a.Id = d.ҽ��id And a.����id = [1] And a.������Ŀid = b.������Ŀid And b.�����ļ�id = c.Id And c.���� = 7 And" & vbNewLine & _
                "             b.Ӧ�ó��� = 1 And c.���� Like '%סԺ֤%'" & vbNewLine & _
                "       Order By D.����ʱ�� Desc)) Where Rownum<10 "
    End If
    
    '��������
    'ID=�ϼ�ID+K����ID,ҽ��ID,0
    '����=����ID;ҽ��ID
    strSQL = strSQL & " Union All" & _
        " Select A.�ϼ�id||'K'||Trim(To_Char(A.ID))||','||Trim(To_Char(Nvl(A.ҽ��id,0)))||',0' As ID,A.�ϼ�id," & _
        "       Decode(A.ҽ��id,Null,A.����||'('||To_Char(A.����ʱ��, 'YYYY-MM-DD')||')',A.����||'��'||B.ҽ������||'('||To_Char(A.����ʱ��, 'YYYY-MM-DD')||')') As ����," & _
        "       Trim(To_Char(A.ID))||';'||Decode(A.ҽ��id,Null,'0',Trim(To_Char(A.ҽ��id))) || ';'|| B.������� || ';'|| A.RISID||';'|| A.����||';'||A.�༭��ʽ As ����," & _
        "       1 As ĩ��,Decode(��������,1,'object_case',2,'object_case',4,'object_case',7,'object_report','object_file') As ͼ��,���� " & _
        " From (Select A.ID, 'R'||A.�������� As �ϼ�id, A.�������� As ����,C.ҽ��id,C.RISID,A.��������,A.�༭��ʽ,A.����ʱ��,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') As ����" & _
        "       From ���Ӳ�����¼ A,����ҽ������ C " & _
        "       Where A.����id = [1] And A.��ҳid = [2] And (A.������Դ=2 And [3]=1 Or Nvl(A.������Դ,0)<>2 And [3]=0)" & _
        "           And C.����id(+)=A.ID And A.�������� In (1, 2, 3, 4, 5, 6, 7)" & _
        "       ) A,����ҽ����¼ B Where A.ҽ��id=B.Id(+)"
    '������
    'ID=�ϼ�ID+K�ļ�ID,0,����ID
    '����=����ID;����;��ʼ����ֹ;�ļ�ID
    '��鱾�β�����ʹ�õ����ϰ廹���°�
    strSQL1 = "Select 1 From ���˻����¼ A Where a.����id = [1] And a.��ҳid = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL1, "����Ƿ�����ϰ�����", lng����ID, mlng����ID)
    If rsTmp.RecordCount > 0 Then
        mblnNewTends = False
        strSQL = strSQL & " Union All" & _
            " Select 'R3K'||Trim(To_Char(A.ID))||',0,'||Trim(To_Char(A.����Id)) As ID,'R3' As �ϼ�id," & _
            "       A.����||'('||B.����||'��'||To_Char(A.��ʼ, 'YYYY-MM-DD HH24:MI') || '��' ||To_Char(A.��ֹ, 'YYYY-MM-DD HH24:MI') || ')' As ����," & _
            "       Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(����,0)))||';'||To_Char(A.��ʼ, 'YYYY-MM-DD HH24:MI')||'��'||To_Char(A.��ֹ, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID)) ||';'||'0' As ����," & _
            "       1 As ĩ��,'object_tend' As ͼ��,To_Char(a.��ʼ,'YYYY-MM-DD HH24:MI:SS') As ����" & _
            " From (" & _
            "   Select F.ID, F.���, F.����, R.��ʼ, R.��ֹ, R.����id, ����" & _
            "   From (" & _
            "       Select ID, ���, ����, 3 As ������, ͨ��, 0 As ����id, ����" & _
            "          From �����ļ��б� Where ���� = 3 And ���� < 0" & _
            "       Union All" & _
            "       Select L.ID, L.���, L.����, F.���� As ������, L.ͨ��, A.����id, L.����" & _
            "          From ����ҳ���ʽ F, �����ļ��б� L, ����Ӧ�ÿ��� A" & _
            "          Where L.���� = 3 And L.���� = 0 And L.���� = F.���� And L.��� = F.��� And L.ID = A.�ļ�id(+)" & _
            "       ) F,(" & _
            "       Select R.����id, Nvl(Min(R.������), 3) As ������, Min(R.����ʱ��) As ��ʼ, Max(R.����ʱ��) As ��ֹ" & _
            "          From ���˻����¼ R" & _
            "          Where R.������Դ = 2 And R.����id = [1] And Nvl(R.��ҳid, 0) = [2] And Nvl(R.Ӥ��, 0) = 0" & _
            "          Group By R.����id" & _
            "       ) R" & _
            "       Where (F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = R.����id) And F.������ >= R.������" & _
            "   ) A, ���ű� B Where A.����id = B.ID "
    Else
        mblnNewTends = True
        strSQL = strSQL & " Union All" & _
                " Select 'R3K'||Trim(To_Char(A.ID))||',0,'||Trim(To_Char(A.����Id)) As ID,'R3' As �ϼ�id," & vbNewLine & _
                "     A.����||'('||B.����||'��'||To_Char(A.��ʼ, 'YYYY-MM-DD HH24:MI') || '��' ||To_Char(A.��ֹ, 'YYYY-MM-DD HH24:MI') || ')' As ����," & vbNewLine & _
                "      Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(����,0)))||';'||To_Char(A.��ʼ, 'YYYY-MM-DD HH24:MI')||'��'||To_Char(A.��ֹ, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID))||';'||Trim(To_Char(A.Ӥ��)) As ����," & vbNewLine & _
                "       1 As ĩ��,'object_tend' As ͼ��,To_Char(a.��ʼ,'YYYY-MM-DD HH24:MI:SS') As ����" & vbNewLine & _
                " From (" & vbNewLine & _
                "   Select R.ID, F.���, R.����,R.Ӥ��, R.��ʼ, NVL(R.��ֹ,nvl(R.ʱ��,R.��ʼ)) ��ֹ, R.����id, ����" & vbNewLine & _
                "   From (" & vbNewLine & _
                "       Select L.ID, L.���, L.����, F.���� As ������, L.ͨ��, L.����" & vbNewLine & _
                "          From ����ҳ���ʽ F, �����ļ��б� L" & vbNewLine & _
                "          Where L.���� = 3 And L.���� = F.���� And L.��� = F.��� And (L.ͨ��=1 OR L.ͨ��=2)" & vbNewLine & _
                "" & vbNewLine & _
                "       ) F,(" & vbNewLine & _
                "       Select R.ID,R.����id,R.�ļ����� ����,R.��ʽID,nvl(R.Ӥ��,0) Ӥ��,Min(R.��ʼʱ��) As ��ʼ, Max(R.����ʱ��) As ��ֹ,MAX(T.����ʱ��) ʱ��" & vbNewLine & _
                "          From ���˻����ļ� R,���˻������� T" & vbNewLine & _
                "          Where R.ID=T.�ļ�ID(+) And R.����id = [1] And Nvl(R.��ҳid, 0) = [2]" & vbNewLine & _
                "          Group By R.ID,R.�ļ�����,R.����id,R.��ʽID,R.Ӥ��" & vbNewLine & _
                "       ) R" & vbNewLine & _
                "       Where F.ID=R.��ʽID" & vbNewLine & _
                "   ) A, ���ű� B Where A.����id = B.ID And DECODE(A.����,-1,0,A.Ӥ��)=A.Ӥ��"
    End If
    
    strSQL = " Select * From (" & strSQL & ") Order By Decode(�ϼ�id,Null,' ',�ϼ�id),����"
    
    If blnDataMove And mlng����ID <> 0 Then
        strSQL = Replace(strSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
        strSQL = Replace(strSQL, "���˻����¼", "H���˻����¼")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "���˻����ļ�", "H���˻����ļ�")
        strSQL = Replace(strSQL, "���˻�������", "H���˻�������")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID, IIf(str�Һŵ� = "", 1, 0))
    Set rsMedRec = zlDatabase.CopyNewRec(rsTmp, False, "", Array("ѡ��", adInteger, 2, Empty))
    'EMR
    Set rsTmp = Nothing
    Set rsTmp = GetEmrCISStruct(lng����ID, lng��ҳID)
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF

	If InStr("," & strRptIDs & ",", "," & rsTmp!ID & ",") = 0 Then
            rsMedRec.AddNew
            rsMedRec!ID = rsTmp!ID
            rsMedRec!�ϼ�ID = rsTmp!�ϼ�ID
            rsMedRec!���� = rsTmp!����
            rsMedRec!���� = NVL(rsTmp!����) & ";EMR;" & rsTmp!�ϼ�ID
            rsMedRec!ͼ�� = NVL(rsTmp!ͼ��)
            rsMedRec!ĩ�� = 1
            rsMedRec.Update
	     strRptIDs = strRptIDs & "," & rsTmp!ID
            End If

            rsTmp.MoveNext
        Loop
    End If
    '�°�PACS
    Set rsTmp = Nothing
    If Not mobjPublicPACS Is Nothing Then
        Set rsTmp = mobjPublicPACS.zlDocGetList(lng����ID, lng��ҳID, str�Һŵ�)
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                rsMedRec.AddNew
                rsMedRec!ID = "R7P" & rsTmp!����ID
                rsMedRec!�ϼ�ID = "R7"
                rsMedRec!���� = rsTmp!�ĵ����� & ""
                rsMedRec!���� = rsTmp!����ID & ";" & rsTmp!ҽ��ID
                rsMedRec!ͼ�� = "object_report"
                rsMedRec!ĩ�� = 1
                rsMedRec.Update
                rsTmp.MoveNext
            Loop
        End If
    End If
    
    '����LIS����
    If str�Һŵ� = "" Then
        strSQL = "select b.id as ����ID,b.������ as �ĵ�����,c.ҽ��ID,b.���� from ����ҽ����¼ a, ҽ���������� b,����ҽ������ c where b.id=c.����id and a.id=c.ҽ��id and c.����id is not null and a.����id=[1] and a.��ҳid=[2]"
    Else
        strSQL = "select b.id as ����ID,b.������ as �ĵ�����,c.ҽ��ID,b.���� from ����ҽ����¼ a, ҽ���������� b,����ҽ������ c where b.id=c.����id and a.id=c.ҽ��id and c.����id is not null and a.�Һŵ�=[3]"
    End If
    If blnDataMove Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
  
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID, str�Һŵ�)
    If Not rsTmp Is Nothing Then
    strRptIDs = ""
        Do While Not rsTmp.EOF
	If InStr("," & strRptIDs & ",", "," & rsTmp!����ID & ",") = 0 Then
            rsMedRec.AddNew
            rsMedRec!ID = "R7L" & rsTmp!����ID
            rsMedRec!�ϼ�ID = "R7"
            rsMedRec!���� = rsTmp!�ĵ����� & ""
            rsMedRec!���� = rsTmp!����ID & ";" & rsTmp!ҽ��ID & ";" & rsTmp!���� & "<sTab>" & rsTmp!�ĵ�����
            rsMedRec!ͼ�� = "object_report"
            rsMedRec!ĩ�� = 1
            rsMedRec.Update

	     strRptIDs = strRptIDs & "," & rsTmp!����ID
            End If

            rsTmp.MoveNext
        Loop
    End If
    '׷����������
    If str�Һŵ� = "" And lng����ID <> 0 Then
        For i = 1 To mcolReport.Count
            rsMedRec.AddNew
            rsMedRec!ID = "R10K" & i
            rsMedRec!�ϼ�ID = "R10"
            rsMedRec!���� = Split(mcolReport(i), ",")(0)
            rsMedRec!���� = Split(mcolReport(i), ",")(0) & ";" & Split(mcolReport(i), ",")(1) & ";" & Split(mcolReport(i), ",")(2)  '��������,ϵͳ��,������
            rsMedRec!ͼ�� = "object_report"
            rsMedRec!ĩ�� = 1
            rsMedRec.Update
        Next
    End If
    rsMedRec.Filter = ""
    Set GetCISStruct = rsMedRec
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetPDFPath()
    Dim objFSO As New Scripting.FileSystemObject
    Dim strPath As String
    
    strPath = GetRegister(˽��ģ��, "��ӡ����", "PDFλ��", App.Path)
    strPath = OS.OpenDir(Me.hwnd, "��ѡ�񵼳�PDF�ļ�λ��", strPath)
    If strPath = "" Then Exit Sub
    If Not objFSO.FolderExists(strPath) Then
        Call objFSO.CreateFolder(strPath)
    End If
    err.Clear: On Error Resume Next
    Call SetRegister(˽��ģ��, "��ӡ����", "PDFλ��", strPath)
End Sub

Private Sub FuncShowReport(ByVal node As MSComctlLib.node)
    Dim objItem As TabControlItem
    Dim lngIndex As Long
    Dim i As Long
    Dim intPage As Integer
    Dim strCaption As String
    Dim strMsg     As String
    Dim arrPar As Variant
    
    If mlng����ID = 0 Then Exit Sub
    If Not LoadPrint Then Exit Sub  '���ش�ӡ��
    lngIndex = -1
    arrPar = Split(node.Tag, ";")
    strCaption = node.Text
    If Not mobjReportForm Is Nothing Then Unload mobjReportForm
    Set mobjReportForm = Nothing
    If node.Key Like "R11K*" Then
        intPage = Val(Replace(node.Key, "R11K", ""))
        Call PrintInMedRec(Nothing, 5, mlng����ID, mlng����ID, mobjReport, mlng����ID, Me, intPage, mcolPrint("R11"), mobjReportForm)
    ElseIf node.Key Like "R12K*" Then
        If node.Key = "R12K1" Then
            Call gobjKernel.zlPrintAdvice(Me, mlng����ID, mlng����ID, 0, 1, mcolPrint("R12"), 5, mobjReportForm, strMsg)
        ElseIf node.Key = "R12K2" Then
            Call gobjKernel.zlPrintAdvice(Me, mlng����ID, mlng����ID, 0, 0, mcolPrint("R12"), 5, mobjReportForm, strMsg)
        End If
    ElseIf node.Key Like "R9K*" Then
        strCaption = "סԺ֤"
        Call mobjReport.SetReportPrintSet(gcnOracle, glngSys, "ZLCISBILL" & Format(arrPar(0), "00000") & "-1", "printer", mcolPrint("R9"))  '����ָ����ӡ��
        Call mobjReport.LoadReport(gcnOracle, glngSys, "ZLCISBILL" & Format(arrPar(0), "00000") & "-1", Me, mobjReportForm, Nothing, "NO=" & arrPar(1), "����=" & arrPar(2), "ҽ��ID=0", 1)
    ElseIf node.Key Like "R10K*" Then
        strCaption = "��������"
        Call mobjReport.SetReportPrintSet(gcnOracle, glngSys, arrPar(2), "printer", mcolPrint("R10"))
        Call mobjReport.LoadReport(gcnOracle, 0, arrPar(2), Me, mobjReportForm, Nothing, "����id=" & mlng����ID, "��ҳid=" & mlng����ID, 1)
    End If
    
    If mobjReportForm Is Nothing Then MsgBox "����δ��ȡ�ɹ�������ϵ����Ա��" & IIf(strMsg <> "", vbCrLf & "��ʾ:" & strMsg, ""), vbInformation, Me.Caption
    For i = 0 To tbcArchive.ItemCount - 1
        If tbcArchive.Item(i).Tag = "����" Then
            If mobjReportForm Is Nothing Then
                Set objItem = tbcArchive.InsertItem(i, strCaption, picTmp.hwnd, 0)
            Else
                Set objItem = tbcArchive.InsertItem(i, strCaption, mobjReportForm.hwnd, 0)
            End If
            objItem.Tag = "����"
            Call tbcArchive.RemoveItem(i + 1)
            objItem.Selected = True
            Exit For
        End If
    Next
    Call ShowArchiveTab("����", strCaption)
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean, Optional ByVal lngPatiID As Long)
'���ܣ�����(��һ��)����
'������blnNext=�Ƿ������һ��
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long
    'ֱ�Ӳ��Ҳ���
    If chkFilter.Value = vbChecked Then
        If lngPatiID = 0 Then
            MsgBox "�Ҳ������������Ĳ��ˡ�", vbInformation, gstrSysName
        Else
            Call LoadPatients(lngPatiID)
        End If
        Exit Sub
    End If
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
                ElseIf mstrFindType = "�����" Then '�����
                    If .Record(COL_�����).Value = PatiIdentify.Text Then Exit For
                ElseIf mstrFindType = "�Һŵ�" Then '�Һŵ�
                    If UCase(.Record(COL_NO).Value) = UCase(PatiIdentify.Text) Then Exit For
                ElseIf mstrFindType = "����" Then '����
                    If UCase(Trim(.Record(col_����).Value)) = UCase(PatiIdentify.Text) Then Exit For
                ElseIf mstrFindType = "���￨" Then '���￨
                    If UCase(.Record(col_���￨��).Value) = UCase(PatiIdentify.Text) Then Exit For
                ElseIf mstrFindType = "���ۺ�" Then '���ۺ�
                    If UCase(.Record(col_���ۺ�).Value) = UCase(PatiIdentify.Text) Then Exit For
                ElseIf mstrFindType = "����" Then '����
                    If .Record(col_����).Value Like "*" & PatiIdentify.Text & "*" Then Exit For
                ElseIf mstrFindType = "�������֤" Then '�������֤
                    If .Record(col_���֤��).Value = UCase(PatiIdentify.Text) Then Exit For
                End If
            End If
        End With
    Next
    
    If i <= rptPati.Rows.Count - 1 Then
        blnReStart = False
        '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        Set rptPati.FocusedRow = rptPati.Rows(i)
        If rptPati.Visible Then rptPati.SetFocus
    Else
        blnReStart = True
        MsgBox IIf(blnNext, "������", "") & "�Ҳ������������Ĳ��ˡ�", vbInformation, gstrSysName
    End If
End Sub

Private Sub FuncViewDisReportCard(ByVal lngEPRid As Long)
    Dim objFrm As New frmChildScale
    '�����Ⱦ�����濨δ��ʼ���ɹ��������³�ʼ��
    If mobjInfection Is Nothing Then
        Set mobjInfection = DynamicCreate("zlDisReportCard.clsDisReportCard", "��Ⱦ�����濨", True)
        If Not mobjInfection Is Nothing Then
            mobjInfection.Init gcnOracle, glngSys
            
        End If
    End If
    
    If Not mobjInfection Is Nothing Then
        objFrm.zlInitData mobjInfection.zlGetForm
        mobjInfection.zlRefresh mlng����ID, mlng����ID, lngEPRid, mblnMoved
        objFrm.Show 1, Me
    End If
End Sub

Private Function LoadPrint() As Boolean
    Dim varTemp As Variant
    Dim strPrint As String
    Dim i As Long
    
    If Printers.Count = 0 Then
        MsgBox "ע�⣺" & Chr(13) _
            & "    δ��װ��ӡ������ͨ��ϵͳ���õĴ�ӡ��" & Chr(13) _
            & "������Ӱ�װ��ӡ����", vbCritical + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set mcolPrint = New Collection
    varTemp = Split(M_CON_CATE, ",")
    For i = LBound(varTemp) To UBound(varTemp)
        strPrint = GetRegister(˽��ģ��, "��ӡ����", "��ӡ��" & varTemp(i), Printer.DeviceName)
        mcolPrint.Add strPrint, varTemp(i)
    Next
    LoadPrint = True
End Function

Private Function DeleteLISTempFile() As Boolean
    Dim objFile As New FileSystemObject
    Dim i As Long
    If mstrTempDel = "" Then Exit Function
    If objFile.FileExists(mstrTempDel) Then
        Do While i < 1000
            On Error Resume Next
            objFile.DeleteFile mstrTempDel, True
            If err.Number = 0 Then
                mstrTempDel = ""
                Exit Do
            End If
            err.Clear: On Error GoTo 0
        Loop
    End If
End Function


