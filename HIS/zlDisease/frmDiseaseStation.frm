VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#7.0#0"; "zlIDKind.ocx"
Begin VB.Form frmDiseaseStation 
   Caption         =   "��Ⱦ��������վ"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18240
   Icon            =   "frmDiseaseStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   18240
   StartUpPosition =   3  '����ȱʡ
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
               Caption         =   "�����"
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
               Caption         =   "������"
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
               Caption         =   "���޴����"
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
               Caption         =   "״̬(S):"
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
               Caption         =   "���ϱ�"
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
               Caption         =   "���ϱ�"
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
               Caption         =   "״̬(S):"
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
               Caption         =   "����"
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
               Caption         =   "���"
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
               Caption         =   "ְҵ"
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
               Caption         =   "*����"
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
               Caption         =   "*����������д��"
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
               Caption         =   "�������ʱ��"
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
               Caption         =   "�������ʱ��"
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
               Caption         =   "�Ǵ�Ⱦ����¼"
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
               Caption         =   "����д���濨"
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
               Caption         =   "����������"
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
               Caption         =   "״̬(S):"
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
               Caption         =   "�������Ǽ�ʱ��"
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
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmDiseaseStation.frx":6852
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindAppearance=   0
         ShowSortName    =   -1  'True
         DefaultCardType =   "���￨"
         IDKindWidth     =   555
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Caption         =   "����:"
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
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
            Text            =   "27��"
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   29
            Text            =   "��"
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Text            =   "����"
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
            Caption         =   "����:"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��ʶ��:"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "ְ    ҵ:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   35
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "��    ��:"
            Height          =   255
            Index           =   6
            Left            =   3720
            TabIndex        =   34
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   " ��ͥ��ַ:"
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
            Name            =   "����"
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
         Text            =   "�ȴ�������������"
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
            Text            =   "���"
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "�������2:"
            Height          =   255
            Index           =   11
            Left            =   6960
            TabIndex        =   22
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "�������1:"
            Height          =   255
            Index           =   10
            Left            =   6960
            TabIndex        =   21
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "ȷ������:"
            Height          =   255
            Index           =   9
            Left            =   3720
            TabIndex        =   20
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "��������:"
            Height          =   255
            Index           =   8
            Left            =   3720
            TabIndex        =   19
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "���Ƽ���:"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   29263
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
            Object.Tag             =   "��ɾ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":117F6
            Key             =   ""
            Object.Tag             =   "����д"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":18058
            Key             =   ""
            Object.Tag             =   "������"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":1E8BA
            Key             =   ""
            Object.Tag             =   "���ϱ�"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":2511C
            Key             =   ""
            Object.Tag             =   "�����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":2B97E
            Key             =   ""
            Object.Tag             =   "���޴����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":321E0
            Key             =   ""
            Object.Tag             =   "���ϱ�"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":38A42
            Key             =   ""
            Object.Tag             =   "������"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseStation.frx":3F2A4
            Key             =   ""
            Object.Tag             =   "�Ǵ�Ⱦ��"
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
    ͼ�� = 0
    ID = 1
    ״̬ = 2
    ���� = 3
    ��Դ = 4
    ���� = 5
    ����� = 6
    ���� = 7
    �Ա� = 8
    ���� = 9
    �ʱ�� = 10
    ��� = 11
    ������� = 12
    ���Ƽ��� = 13
    �Ǽ��� = 14
    �Ǽ�ʱ�� = 15
    ���ע = 16
    ɾ���� = 17
    ɾ��ʱ�� = 18
    ������ = 19
    ����ʱ�� = 20
    ���͵�λ = 21
    ���ͱ�ע = 22
    ����ת�� = 23
    ����ID = 24
    ��ҳID = 25
    �ļ�ID = 26
    �༭��ʽ = 27
    ��Ϣ = 28
End Enum

Private Enum mRptCol
    ID = 0
    ���� = 1
    �������� = 2
    ����˵�� = 3
    �Ǽ��� = 4
    �Ǽ�ʱ�� = 5
    ������ = 6
    ����ʱ�� = 7
End Enum

Private Enum mSendRptCol
    ID = 0
    ���� = 1
    �������� = 2
    ����˵�� = 3
    �Ǽ��� = 4
    �Ǽ�ʱ�� = 5
End Enum

Private Enum mCtlID
    txt���� = 0
    txt�Ա� = 1
    txt���� = 2
    txt��ʶ�� = 3
    txt���� = 4
    txtְҵ = 5
    txt�绰 = 6
    txt��ַ = 7
    txt�������� = 8
    txtȷ������ = 9
    txt�������1 = 10
    txt�������2 = 11
    txt���Ƽ��� = 12

    chk����� = 0
    chk������ = 1
    chk���޴���� = 2
    chk����д = 0
    chk�Ǵ�Ⱦ�� = 1
    chk������ = 2
    chk���ϱ� = 0
    chk���ϱ� = 1
End Enum

Private Enum mTcbID
    tcbδ��д = 0
    tcb��� = 1
    tcb�ϱ� = 2
    tcb��ɾ�� = 3
    tcb���ع��� = 4
End Enum

Private Const conPane_Reports = 1
Private Const conPane_AppInfo = 2
Private Const conPane_Preview = 3
Private Const conPane_Feedback = 4

Private mTcbSelectID As mTcbID           'ѡ����TabControl��ҳ��
Private mstrState As String              'ɸѡʱ����Ĵ���״̬

Private mstrPrivs As String              '��ǰʹ����Ȩ�޴�
Private mstrFiles As String              '��������ı����ļ�

Private mintWaitIndex As Integer  'δ��дʱʱ��ѡ�����
Private mintDelIndex As Integer   '��ɾ��ʱʱ��ѡ�����
Private mintIndex As Integer      '��˺��ϱ�����ʱʱ��ѡ�����

'��˹������ϱ����� �鿴����ļ�¼�����뷶Χ
Private mintDates As Integer              'Ĭ�ϲ鿴�����¼��������Ϊ0ʱ��˵���������ִ�в������ã�Ҫ�����ڷ�Χ�鿴
Private mdtFrom As Date, mdtTo As Date    '����Χ�鿴��ʼ���ڣ���mintDates=0ʱ��Ч;����Χ�鿴��ֹ���ڣ���mintDates=0ʱ��Ч

'δ��д �鿴����ļ�¼�����뷶Χ
Private mintWaitDays As Integer
Private mdtWaitBegin As Date, mdtWaitEnd As Date

'��ɾ�� �鿴����ļ�¼�����뷶Χ
Private mintDelDays As Integer
Private mdtDelBegin As Date, mdtDelEnd As Date

Private mfrmPreview As frmDockEPRContent                 '��������Ԥ������
Private mfrmPreFeedBack As frmDockEPRContent             '�ѹ����ķ�������������Ԥ������
Private mobjInfection As Object                          '�л����񹲺͹���Ⱦ�����濨

Private mstrCurId As String               '��ǰ��¼ID EMR���ID���ַ���
Private mstrContent As String             '�²�����XML����
Private mIntState As Integer              '0-�����գ�-1-�Ѿ��գ�1-����ˣ�4-�����ޣ�3-���ϱ���2-���ϱ���5-���޴���6-��ɾ��
Private mblnCurMoved As Boolean           '��ǰ��¼ת��״̬ 0-δת�� 1-��ת��
Private mstrFindType As String            '���Ҳ���ʱ���ҵ�����
Private mdatTime As Date                  '�鿴ѡ�е��ϱ�����ķ���˵����¼�Ĵ���ʱ��

Private mblnReportCheck As Boolean        '�Ƿ�ֻ����ʾ���ص�ҳ�棬ҽ��վ���ýӿ���ʾһ�����Ƿ��д��벡�˵��ظ�����ʱ��
Private mrsOld As ADODB.Recordset         '����ʱ�ϰ���Ӳ�����ѯ��������
Private mblnReport As Boolean             '��ǰ��ʾ���Ƿ��Ǳ��濨��true-���濨��false-��������
Private mlngID As Long                    '��ǰ��ʾ�ı��濨���߷�������ID


'�鿴���Խ��������ʱ���øý���ʱ��һЩ����
Private mstrName As String
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mIntPatiFrom As Integer
Private mlng����ID As Long
Private mstr����ID As String
Private mstr���ID As String

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl, strInfo As String
    Dim strSqlFile As String, rsTmp As ADODB.Recordset
    
    If mblnCurMoved And (Control.ID = conMenu_File_Open Or Control.ID = conMenu_Edit_Reuse Or Control.ID = conMenu_Edit_Send Or Control.ID = conMenu_Edit_Untread) Then
        MsgBox "�ò��˵ı��������Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                        "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    Select Case Control.ID
        Case conMenu_File_PrintSet: Call zlPrintSet
        Case conMenu_File_Preview: Call zlEPRPrint(True)
        Case conMenu_File_Print: Call zlEPRPrint(False)
        Case conMenu_File_RowPrint: Call zlRptPrint(1)
        Case conMenu_File_Parameter
            Call frmDiseaseStationSet.ShowMe(Me, InStr(1, mstrPrivs, "��Χ����") > 0, mstrFiles)
            Call zlRefList
        Case conMenu_File_Exit
             Unload Me
        Case conMenu_Edit_Audit
                Dim intAduitState As Integer
                 With rptList
                    strInfo = .FocusedRow.Record.Item(mCol.����).Value
                    strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.����).Value
                    strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.����).Value
                    strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.�Ա�).Value
                    strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.����).Value
                    strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.����ID).Value
                    strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.��ҳID).Value
                    strInfo = strInfo & "|" & IIf(.FocusedRow.Record.Item(mCol.��Դ).Value = "סԺ", "2", "1")
                    strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.���).Value
                    strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.�ʱ��).Value
                End With
                If frmDiseaseAduit.ShowDiseaseAudit(Me, mstrCurId, strInfo, intAduitState) Then
                    '(intAduitState = 3)���ͨ��������״̬=3;Ҫ���ޣ�����״̬=4��������Ϣ������/סԺҽ��վ
                    If intAduitState = 4 Then
                        Call SendMsg    '������Ϣ
                    End If
                    Call zlRefList(mstrCurId)
                End If
        Case conMenu_Edit_Delete
            If mstrCurId <> "" Then
                If MsgBox("��ȷ��Ҫɾ���ñ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
            'strInfo=����|����|����|�Ա�|����|�����|���|�ʱ��|����ID|��ҳID
            With rptList
                strInfo = .FocusedRow.Record.Item(mCol.����).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.����).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.����).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.�Ա�).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.����).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.�����).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.���).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.�ʱ��).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.����ID).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.��ҳID).Value
                strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.�ļ�ID).Value
            End With
            If frmDiseaseReportSend.ShowMe(Me, mstrCurId, strInfo, txtInfo(txt�������1).Text, txtInfo(txt�������2).Text) Then Call zlRefList(mstrCurId)
        Case conMenu_Edit_Untread  '�ջ�
            Dim strMsg As String
            Dim strSQL As String
            Select Case mIntState
                Case 2
                    If CheckUntread() Then
                        strMsg = "���ȡ���ü�������ġ��걨�Ǽǡ���"
                    Else
                        MsgBox "���ϱ��ı����Ѿ������˺�������������ȡ����", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Case 3:  strMsg = "���ȡ���ü�������ġ����ͨ������"
                Case 4:  strMsg = "���ȡ���ü�������ġ�Ҫ���ޡ���"
                Case 6:  strMsg = "���ȡ���ü�������ġ�ɾ��������"
                Case Else: Exit Sub
            End Select
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            strSQL = "Zl_�����걨��¼_Untread('" & mstrCurId & "',1)"
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
            strSqlFile = "select t.����id,t.��ҳid,t.����id,t.Ӥ��,t.������Դ from ���Ӳ�����¼ t where t.id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSqlFile, "�ļ���Ϣ��ѯ", mstrCurId)
            If rsTmp.RecordCount <> 0 Then
                 mobjInfection.OpenDoc Me, 1, rsTmp!����ID, rsTmp!��ҳID, rsTmp!������Դ, Val(rsTmp!Ӥ�� & ""), rsTmp!����ID, mstrCurId
            End If
        Case conMenu_Help_Web_Home: Call gobjComlib.zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Forum '������̳
            Call gobjComlib.zlWebForum(Me.hwnd)
        Case conMenu_Help_Web_Mail: Call gobjComlib.zlMailTo(Me.hwnd)
        Case conMenu_Help_About:    Call gobjComlib.ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case Else
            'ִ�з�������ǰģ��ı���
            Dim lng����ID As Long
            If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                If rptList.SelectedRows.Count > 0 Then
                    If Not rptList.SelectedRows(0).GroupRow Then
                        lng����ID = Val(rptList.SelectedRows(0).Record(mCol.ID).Value)
                    End If
                End If
                If lng����ID <> 0 Then
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "����ID=" & lng����ID)
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
            If Control.Enabled Then Control.Enabled = (Me.rptList.FocusedRow.Record.Item(mCol.�༭��ʽ).Value <> 2)
        Case conMenu_File_Print
             Control.Enabled = (mlngID > 0)
        Case conMenu_File_RowPrint
             Control.Enabled = (Me.rptList.Records.Count <> "")
        Case conMenu_Edit_Audit
            Control.Visible = (InStr(1, mstrPrivs, "���") > 0)
            If Control.Visible Then Control.Visible = (mTcbSelectID = tcb���)
            If Control.Visible Then Control.Visible = chkAduitState(chk�����).Value = 1
            If Control.Visible Then Control.Visible = (tbcMain.Item(0).Selected)
            Control.Enabled = (mstrCurId <> "" And (mIntState = -1 Or mIntState = 0 Or mIntState = 1 Or mIntState = 5))
        Case conMenu_Edit_Send
            Control.Visible = (InStr(1, mstrPrivs, "����") > 0)
            If Control.Visible Then Control.Visible = (mTcbSelectID = tcb�ϱ�)
            If Control.Visible Then Control.Visible = chkSendState(chk���ϱ�).Value = 1
            If Control.Visible Then Control.Visible = (tbcMain.Item(0).Selected)
            Control.Enabled = (mstrCurId <> "" And mIntState = 3)
        Case conMenu_Edit_Untread
            Control.Visible = (InStr(1, mstrPrivs, "����") > 0)
            If Control.Visible Then Control.Visible = (mTcbSelectID <> tcb���ع��� And mTcbSelectID <> tcbδ��д)
            If Control.Visible Then Control.Visible = (tbcMain.Item(0).Selected)
            Control.Enabled = (mstrCurId <> "" And (mIntState = 2 Or mIntState = 3 Or mIntState = 4 Or mIntState = 6))
        Case conMenu_Edit_Delete
            Control.Visible = (InStr(1, mstrPrivs, "ɾ��") > 0)
            If Control.Visible Then Control.Visible = (mTcbSelectID = tcb���ع���)
            If Control.Visible Then Control.Visible = (tbcMain.Item(0).Selected)
            Control.Enabled = (mstrCurId <> "" And (mIntState = -1 Or mIntState = 1 Or mIntState = 2 Or mIntState = 3 Or mIntState = 5))
        Case conMenu_Edit_NewTable
            Control.Visible = mIntPatiFrom <> 0 And mblnReportCheck
        Case conMenu_Edit_Add
            Control.Visible = (InStr(1, mstrPrivs, "����") > 0)
            If Control.Visible Then Control.Visible = (mTcbSelectID = tcb�ϱ�)
            If Control.Visible Then Control.Visible = chkSendState(chk���ϱ�).Value = 1
            If Control.Visible Then Control.Visible = (tbcMain.Item(0).Selected)
            Control.Enabled = (mstrCurId <> "" And mIntState = 2)
            
        Case conMenu_Edit_EditInfo           '�޸ı��濨
            Control.Visible = (InStr(1, mstrPrivs, "���") > 0) And mblnReport And Me.rptList.FocusedRow.Record.Item(mCol.�༭��ʽ).Value = 2
            If Control.Visible Then Control.Visible = (mTcbSelectID = tcb���)
            If Control.Visible Then Control.Visible = (tbcMain.Item(0).Selected)
            Control.Enabled = (mstrCurId <> "" And (mIntState = -1 Or mIntState = 0 Or mIntState = 1 Or mIntState = 5)) And Me.rptList.FocusedRow.Record.Item(mCol.�༭��ʽ).Value = 2
            
        Case conMenu_Edit_Modify
            Control.Visible = (InStr(1, mstrPrivs, "����") > 0)
            If Control.Visible Then Control.Visible = (mTcbSelectID = tcb�ϱ�)
            If Control.Visible Then Control.Visible = chkSendState(chk���ϱ�).Value = 1
            If Control.Visible Then Control.Visible = (tbcMain.Item(0).Selected)
            Control.Enabled = (mstrCurId <> "" And mIntState = 2 And mdatTime <> 0)
        Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsMain(2).Visible
        Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
        Case conMenu_View_Refresh
            Control.Visible = (Not mblnReportCheck And mTcbSelectID <> tcb���ع���)
            Control.Enabled = (Trim(mstrFiles) <> "")
    End Select
End Sub

Private Sub chkAduitState_Click(Index As Integer)
'��˹�������������
    Dim i As Integer
    Dim strState As String, strTemp As String
    mstrState = ""
    For i = chkAduitState.LBound To chkAduitState.UBound
        If chkAduitState(i).Value = 1 Then
            Select Case i
                Case chk�����
                    strState = strState & ",-1,0,1"
                    strTemp = " or S.����״̬ is null "
                Case chk������
                    strState = strState & ", 4"
                Case chk���޴����
                    strState = strState & ", 5"
            End Select
        End If
    Next
    If strState <> "" Then
        If strTemp <> "" Then
            mstrState = " and (S.����״̬ in (" & Mid(strState, 2) & ")" & strTemp & ") "
        Else
            mstrState = " and S.����״̬ in (" & Mid(strState, 2) & ") "
        End If
    End If
    If Me.Visible And (Index <> -1) Then Call zlRefList
End Sub

Private Sub chkDisState_Click(Index As Integer)
'δ��д����������
    Dim i As Integer
    Dim strState As String
    For i = chkDisState.LBound To chkDisState.UBound
        If chkDisState(i).Value = 1 Then
            Select Case i
                Case chk����д
                    strState = strState & ", 2, 4"
                Case chk�Ǵ�Ⱦ��
                    strState = strState & ", 3"
                Case chk������
                    strState = strState & ", 1"
            End Select
        End If
    Next
    mstrState = IIf(strState <> "", " and A.��¼״̬ in (" & Mid(strState, 2) & ") ", "")
    If Me.Visible And (Index <> -1) Then Call zlRefList
End Sub

Private Sub chkSendState_Click(Index As Integer)
''�ϱ���������������
    Dim i As Integer
    Dim strState As String
    For i = chkSendState.LBound To chkSendState.UBound
        If chkSendState(i).Value = 1 Then
            Select Case i
                Case chk���ϱ�
                    strState = strState & ", 3"
                Case chk���ϱ�
                    strState = strState & ", 2"
            End Select
        End If
    Next
    mstrState = IIf(strState <> "", " and S.����״̬ in (" & Mid(strState, 2) & ") ", "")
    If Me.Visible And (Index <> -1) Then Call zlRefList
End Sub

Private Sub Form_Load()
    Dim strState As String
    Dim arrayState() As String
    Dim dtCurDate As Date
    Dim strBegin As String, strEnd As String
On Error Resume Next
    If mblnReportCheck Then
        Me.Caption = "��Ⱦ�����濨�ظ�����"
    Else
         Me.Caption = "��Ⱦ��������վ"
    End If

    Call PatiIdentify.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser)
    PatiIdentify.objIDKind.AllowAutoICCard = True
    PatiIdentify.objIDKind.AllowAutoIDCard = True
    PatiIdentify.ShowSortName = True

    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
    mstrPrivs = gstrPrivs

    mstrFiles = Trim(gobjComlib.zlDatabase.GetPara("������վ�ɹ����ļ�", glngSys, 1278))
    strState = CStr(gobjComlib.zlDatabase.GetPara("��Ⱦ��ϵͳ�鿴״̬��Χ", glngSys, 1278))
    
    If strState <> "" Then
        arrayState = Split(strState, ",")
        chkAduitState(chk�����).Value = Val(arrayState(0))
        chkAduitState(chk������).Value = Val(arrayState(1))
        chkAduitState(chk���޴����).Value = Val(arrayState(2))
        chkSendState(chk���ϱ�).Value = Val(arrayState(3))
        chkSendState(chk���ϱ�).Value = Val(arrayState(4))
        chkDisState(chk����д).Value = Val(arrayState(5))
        chkDisState(chk�Ǵ�Ⱦ��).Value = Val(arrayState(6))
        chkDisState(chk������).Value = Val(arrayState(7))
    End If

      '��ѯ����
    dtCurDate = gobjComlib.zlDatabase.Currentdate
    mintDates = Val(gobjComlib.zlDatabase.GetPara("������ϱ�����״̬�²鿴��������ı���", glngSys, 1278))

    If mintDates = -1 Then
        strBegin = gobjComlib.zlDatabase.GetPara("������ϱ�����״̬�²鿴ָ�������ı������ʼ����", glngSys, 1278)
        strEnd = gobjComlib.zlDatabase.GetPara("������ϱ�����״̬�²鿴ָ�������ı���Ľ�������", glngSys, 1278)
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

    mintWaitDays = Val(gobjComlib.zlDatabase.GetPara("δ��д״̬�²鿴��������ı���", glngSys, 1278))
    If mintWaitDays = -1 Then
        strBegin = gobjComlib.zlDatabase.GetPara("δ��д״̬�²鿴ָ�������ı������ʼ����", glngSys, 1278)
        strEnd = gobjComlib.zlDatabase.GetPara("δ��д״̬�²鿴ָ�������ı���Ľ�������", glngSys, 1278)
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

    mintDelDays = Val(gobjComlib.zlDatabase.GetPara("��ɾ��״̬�²鿴��������ı���", glngSys, 1278))
    If mintDelDays = -1 Then
        strBegin = gobjComlib.zlDatabase.GetPara("��ɾ��״̬�²鿴ָ�������ı������ʼ����", glngSys, 1278)
        strEnd = gobjComlib.zlDatabase.GetPara("��ɾ��״̬�²鿴ָ�������ı���Ľ�������", glngSys, 1278)
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

     '���ôʾ���ʾͣ������
    Set mfrmPreview = New frmDockEPRContent
    Set mfrmPreFeedBack = New frmDockEPRContent
    Set mobjInfection = DynamicCreate("zlDisReportCard.clsDisReportCard", "��Ⱦ�����濨", True)
    If Not mobjInfection Is Nothing Then
        mobjInfection.Init gcnOracle, glngSys
    End If

    Call InitCommandBar
    Call InitDkpMain
    Call InitReportControl
    Call InitTabContol
    Call InitCboSelectTime

    If Not mblnReportCheck Then
        mTcbSelectID = Val(gobjComlib.zlDatabase.GetPara("��ǰ�鿴����Ĺ���״̬", glngSys, 1278))
        Me.tbcReportList.Item(mTcbSelectID).Selected = True
    End If

'     ����װ��
    If mblnReportCheck Then
        Call SetDuplicateReportData(mrsOld)
    Else
        If mstrFiles = "" Then
            Me.stbThis.Panels(2).Text = "δ���ñ�����վ�ļ������淶Χ"
        Else
            Call zlRefList
        End If
        '����ָ�
        Call gobjComlib.RestoreWinState(Me, App.ProductName)
    End If
    Me.WindowState = vbMaximized
End Sub

Private Sub InitReportControl()
    Dim rptCol As ReportColumn

    With Me.rptList
        Set rptCol = .Columns.Add(mCol.ͼ��, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.״̬, "״̬", 90, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����, "����", 0, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.��Դ, "��Դ", 40, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.����, "����", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.�����, "��ʶ��", 75, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.����, "����", 60, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.�Ա�, "�Ա�", 40, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.����, "����", 40, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.�ʱ��, "�ʱ��", 100, True): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.���, "���", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.�������, "�������", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.���Ƽ���, "���Ƽ���", 100, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.�Ǽ���, "�Ǽ���", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.�Ǽ�ʱ��, "�Ǽ�ʱ��", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.���ע, "���ע", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.ɾ����, "ɾ����", 50, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.ɾ��ʱ��, "ɾ��ʱ��", 50, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.������, "������", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.����ʱ��, "����ʱ��", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.���͵�λ, "���͵�λ", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.���ͱ�ע, "���ͱ�ע", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.����ת��, "����ת��", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����ID, "����ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.��ҳID, "��ҳID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.�ļ�ID, "�ļ�ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.�༭��ʽ, "�༭��ʽ", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.��Ϣ, "��Ϣ", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
    
        .GroupsOrder.Add .Columns.Find(mCol.״̬)
        .GroupsOrder(0).SortAscending = True
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With

    With Me.rptAuditContent
        Set rptCol = .Columns.Add(mRptCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mRptCol.����, "����", 300, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mRptCol.��������, "��������", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mRptCol.����˵��, "����˵��", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mRptCol.�Ǽ���, "�Ǽ���", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mRptCol.�Ǽ�ʱ��, "�Ǽ�ʱ��", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mRptCol.������, "������", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mRptCol.����ʱ��, "����ʱ��", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        .GroupsOrder.Add .Columns(1)
        .GroupsOrder(0).SortAscending = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
    
    With Me.rptSendContent
        Set rptCol = .Columns.Add(mSendRptCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mSendRptCol.����, "����", 300, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mSendRptCol.��������, "��������", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mSendRptCol.����˵��, "����˵��", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mSendRptCol.�Ǽ���, "�Ǽ���", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mSendRptCol.�Ǽ�ʱ��, "�Ǽ�ʱ��", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        
        .GroupsOrder.Add .Columns(1)
        .GroupsOrder(0).SortAscending = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
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
        .InsertItem(0, "���濨", mfrmPreview.hwnd, 0).Tag = "���濨"
        .InsertItem(1, "������", picDis.hwnd, 0).Tag = "������"
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
            .InsertItem(mTcbID.tcbδ��д, "δ��д", picReportList.hwnd, 0).Tag = "δ��д"
            .InsertItem(mTcbID.tcb���, "��˹���", picReportList.hwnd, 0).Tag = "��˹���"
            .InsertItem(mTcbID.tcb�ϱ�, "�ϱ�����", picReportList.hwnd, 0).Tag = "�ϱ�����"
            .InsertItem(mTcbID.tcb��ɾ��, "��ɾ��", picReportList.hwnd, 0).Tag = "��ɾ��"
            .InsertItem(mTcbID.tcb���ع���, "���ع���", picReportList.hwnd, 0).Tag = "���ع���"
            .Item(1).Selected = True
         Else
            .InsertItem(0, mstrName & " �����б�", picReportList.hwnd, 0).Tag = "��Ⱦ������"
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

        .InsertItem(0, "��˷���˵��", PicAuditContent.hwnd, 0).Tag = "��˷���˵��"
        .InsertItem(1, "�ϱ�����˵��", PicSendContent.hwnd, 0).Tag = "�ϱ�����˵��"
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
End Sub

Private Sub SetFeedbackContent(ByVal strID As String, ByVal intState As Integer, Optional ByVal intType As Integer = 0)
'���ܣ���ʾѡ�б���ķ������
'����: strID ѡ�б����ID
'      intState ѡ�б���Ĵ���״̬
'      intType:0-������˷������ϱ��������л�����˷���ҳ�棻1-������˷������л�����˷���ҳ�棻2-�����ϱ��������л����ϱ�����ҳ��
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
            strSQL = "Select Rownum As ����, �ļ�id, �Ǽ���, �Ǽ�ʱ��, ��¼״̬, ��������, ������, ����ʱ��, �������˵��" & vbNewLine & _
                    "From (Select �ļ�id, �Ǽ���, �Ǽ�ʱ��, ��¼״̬, ��������, ������, ����ʱ��, �������˵��" & vbNewLine & _
                    "       From �������淴�� Where �ļ�id = [1] Order By �Ǽ�ʱ��)"
            Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strID))
            If rsTemp.RecordCount > 0 Then
                paneFeedBack.Closed = False
                blnAudit = True
                Me.rptAuditContent.Records.DeleteAll
                Do While Not rsTemp.EOF
                    Set rptRcd = Me.rptAuditContent.Records.Add()
                    rptRcd.AddItem CStr(NVL(rsTemp!�ļ�ID))
                    rptRcd.AddItem "��" & CStr(NVL(rsTemp!����)) & "�η���"
                    rptRcd.AddItem CStr(NVL(rsTemp!��������))
                    rptRcd.AddItem CStr(NVL(rsTemp!�������˵��))
                    rptRcd.AddItem CStr(NVL(rsTemp!�Ǽ���))
                    rptRcd.AddItem CStr(NVL(rsTemp!�Ǽ�ʱ��))
                    rptRcd.AddItem CStr(NVL(rsTemp!������))
                    rptRcd.AddItem CStr(NVL(rsTemp!����ʱ��))
                    rsTemp.MoveNext
                Loop
                rptAuditContent.Populate
            End If
        End If
        
        If (intType = 0 Or intType = 2) And intState = 2 Then
            If IsNumeric(strID) Then
                strSQL = "Select Rownum As ����, �걨id, ������Ϣ, �Ǽ���, �Ǽ�ʱ��, �������˵��" & vbNewLine & _
                            "From (Select �걨id, ������Ϣ, �Ǽ���, �Ǽ�ʱ��, �������˵�� From �����걨���� Where �걨id = [1] Order By �Ǽ�ʱ��)"
    
                Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strID))
                If rsTemp.RecordCount > 0 Then
                    blnSend = True
                    paneFeedBack.Closed = False
                    Me.rptSendContent.Records.DeleteAll
                    Do While Not rsTemp.EOF
                        Set rptRcd = Me.rptSendContent.Records.Add()
                        rptRcd.AddItem CStr(NVL(rsTemp!�걨ID))
                        rptRcd.AddItem "��" & CStr(NVL(rsTemp!����)) & "�η���"
                        rptRcd.AddItem CStr(NVL(rsTemp!������Ϣ))
                        rptRcd.AddItem CStr(NVL(rsTemp!�������˵��))
                        rptRcd.AddItem CStr(NVL(rsTemp!�Ǽ���))
                        rptRcd.AddItem CStr(NVL(rsTemp!�Ǽ�ʱ��))
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
    objPane.Title = "���˱����б�"
    objPane.MinTrackSize.Width = 350
    objPane.MaxTrackSize.Width = 360

    Set objPane = Me.dkpMain.CreatePane(conPane_Feedback, 300, 250, DockBottomOf, objPane)
    objPane.Options = PaneNoFloatable Or PaneNoHideable
    objPane.Title = "����˵��"

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
        .UseFadedIcons = True           'ͼ����ʾΪ��ɫЧ��
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsMain.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    Me.cbsMain.ActiveMenuBar.Title = "�˵�"

    Me.cbsMain.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        If Not mblnReportCheck Then
            Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
            Set cbrControl = .Add(xtpControlButton, conMenu_File_RowPrint, "�嵥��ӡ(&L)��"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&M)��"): cbrControl.BeginGroup = True
        End If
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        If Not mblnReportCheck Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "���(&A)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_EditInfo, "�޸ı��濨(&E)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "����(&S)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "����(&B)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Add, "������ע(&X)"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸ı�ע(&M)")
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewTable, "��д���濨(&N)")
        End If
        cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With

    '�����
    With Me.cbsMain.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, vbKeyF12, conMenu_File_Parameter
        .Add FCONTROL, Asc("A"), conMenu_Edit_Audit
        .Add FCONTROL, Asc("S"), conMenu_Edit_Send
        .Add FCONTROL, Asc("U"), conMenu_Edit_Untread
        .Add 0, VK_F5, conMenu_View_Refresh
    End With

    '����������
    Set cbrToolBar = Me.cbsMain.Add("������", xtpBarTop)
    cbrToolBar.ContextMenuPresent = False                   '�������ϵ������Ҽ�ʱ���������ò˵�
    cbrToolBar.ShowTextBelowIcons = False                   '�������еİ�ť������ʾ��ͼ���Ҳ�
    cbrToolBar.EnableDocking xtpFlagHideWrap                '��������Ȳ���ʱҲ������
    With cbrToolBar.Controls
        If Not mblnReportCheck Then
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
            Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_EditInfo, "�޸ı��濨"): cbrControl.Visible = False
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "���")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "����")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Add, "����˵��"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�˵��")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "����")
            Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewTable, "��д���濨")
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
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
        strChkState = chkAduitState(chk�����).Value & "," & chkAduitState(chk������).Value & "," & chkAduitState(chk���޴����).Value & "," & chkSendState(chk���ϱ�).Value & "," & chkSendState(chk���ϱ�).Value & "," & chkDisState(chk����д).Value & "," & chkDisState(chk�Ǵ�Ⱦ��).Value & "," & chkDisState(chk������).Value
        Call gobjComlib.zlDatabase.SetPara("��Ⱦ��ϵͳ�鿴״̬��Χ", strChkState, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("������ϱ�����״̬�²鿴��������ı���", mintDates, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("������ϱ�����״̬�²鿴ָ�������ı������ʼ����", mdtFrom, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("������ϱ�����״̬�²鿴ָ�������ı���Ľ�������", mdtTo, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("δ��д״̬�²鿴��������ı���", mintWaitDays, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("δ��д״̬�²鿴ָ�������ı������ʼ����", mdtWaitBegin, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("δ��д״̬�²鿴ָ�������ı���Ľ�������", mdtWaitEnd, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("��ɾ��״̬�²鿴��������ı���", mintDelDays, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("��ɾ��״̬�²鿴ָ�������ı������ʼ����", mdtDelBegin, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("��ɾ��״̬�²鿴ָ�������ı���Ľ�������", mdtDelEnd, glngSys, 1278)
        Call gobjComlib.zlDatabase.SetPara("��ǰ�鿴����Ĺ���״̬", mTcbSelectID, glngSys, 1278)
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
        If mTcbSelectID = tcb�ϱ� Or mTcbSelectID = tcb��� Then
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
'���ܣ�ѡ�еı��淢���仯
    Dim strInfo As String
    If Not Me.Visible Then Exit Sub

    '����������ϵĲ�����Ϣ
    Call ClearTxtInfo

    mstrContent = "": rptList.Tag = ""
    With Me.rptList
        If .FocusedRow Is Nothing Then
            mstrCurId = "": mIntState = 0: strInfo = "": mblnCurMoved = False
        ElseIf .FocusedRow.GroupRow = True Then
            mstrCurId = "":  mIntState = 0: strInfo = "": mblnCurMoved = False
        Else
            mstrCurId = .FocusedRow.Record.Item(mCol.ID).Value
            mIntState = .FocusedRow.Record.Item(mCol.ͼ��).Value
            mblnCurMoved = (.FocusedRow.Record.Item(mCol.����ת��).Value = 1)
        End If
    End With

'   �ڽ�������ʾ���˵Ļ�����Ϣ
    Call SetPatiInfo

'   ��ʾѡ�б���ķ������
    Call SetFeedbackContent(mstrCurId, mIntState)
    Call GetFeedbackIDs(Val(mstrCurId))
    tbcMain.Item(0).Selected = True
    
    If IsNumeric(mstrCurId) Then
        Call RefreshReport(Val(mstrCurId), mblnCurMoved, NVL(rptList.FocusedRow.Record.Item(mCol.�༭��ʽ).Value, 0))
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
'����:��������ϲ��˵Ļ�����Ϣ
    txtInfo(txt����).Text = ""
    txtInfo(txt�Ա�).Text = ""
    txtInfo(txt����).Text = ""
    txtInfo(txt��ʶ��).Text = ""
    txtInfo(txt����).Text = ""
    txtInfo(txtְҵ).Text = ""
    txtInfo(txt��ַ).Text = ""
    txtInfo(txt�绰).Text = ""
    txtInfo(txt��������).Text = ""
    txtInfo(txtȷ������).Text = ""
    txtInfo(txt�������1).Text = ""
    txtInfo(txt�������2).Text = ""
    fraInfo(0).Visible = False
    fraInfo(1).Visible = False
    imgPati.Visible = False
    txtState.Visible = False
End Sub

Public Sub zlRptPrint(ByVal bytMode As Byte)
'����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
'����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub
    '-------------------------------------------------
    '�������ݱ��
    If zlReportToVSFlexGrid(Me.vfgTemp, Me.rptList) = False Then Exit Sub
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow

    Set objPrint.Body = Me.vfgTemp
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

'-------------------------------------------------------
'���ܣ�  ����Ԥ������ӡ
'������  blnPreview  :�Ƿ���Ԥ��ģʽ
'-------------------------------------------------------
Private Sub zlEPRPrint(blnPreview As Boolean)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim clsPrint As zlRichEPR.clsDockAduits
    
    If mstrCurId = "" Then Exit Sub

    Err = 0: On Error GoTo ErrHand
    If mblnReport And IsNumeric(mstrCurId) Then
        strSQL = "Select l.������Դ, l.����id, l.��ҳid,l.�༭��ʽ, f.ҳ�� From ���Ӳ�����¼ l, �����ļ��б� f Where l.�ļ�id = f.Id And l.Id = [1]"
        If mblnCurMoved Then
            strSQL = Replace(strSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
        End If
        Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(mstrCurId))
        With rsTemp
            If .RecordCount <= 0 Then MsgBox "�ü�����������Ѿ����ٴ�ɾ����", vbExclamation, gstrSysName: Exit Sub
            If NVL(rsTemp!�༭��ʽ, 0) = 0 Or NVL(rsTemp!�༭��ʽ, 0) = 1 Then
                Set clsPrint = New zlRichEPR.clsDockAduits
                Call clsPrint.zlPrintDocument(3, IIf(blnPreview, 1, 2), CLng(mstrCurId))
                Set clsPrint = Nothing
            ElseIf rsTemp!�༭��ʽ = 2 Then
                mobjInfection.PrintDoc Me, !����ID, !��ҳID, CLng(mstrCurId), ""
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
'���ܣ� �ڽ�������ʾ���˵Ļ�����Ϣ
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
            strInfo = .FocusedRow.Record.Item(mCol.��Ϣ).Value
            aryInfo = Split(strInfo, "|")
            imgPati.Visible = True
            txtInfo(txt����).Text = .FocusedRow.Record.Item(mCol.����).Value
            txtInfo(txt�Ա�).Text = .FocusedRow.Record.Item(mCol.�Ա�).Value
            txtInfo(txt����).Text = .FocusedRow.Record.Item(mCol.����).Value & "��"
            txtInfo(txt����).Text = Replace(txtInfo(txt����).Text, "����", "��")
            txtInfo(txt��ʶ��).Text = .FocusedRow.Record.Item(mCol.�����).Value
            txtInfo(txt����).Text = .FocusedRow.Record.Item(mCol.����).Value
            If .FocusedRow.Record.Item(mCol.��Դ).Value = "����" Then
                lblInfo(txt��ʶ��) = "�����:"
            ElseIf .FocusedRow.Record.Item(mCol.��Դ).Value = "סԺ" Then
                lblInfo(txt��ʶ��) = "סԺ��:"
            Else
                lblInfo(txt��ʶ��) = "��ʶ��:"
            End If
            txtInfo(txtְҵ).Text = aryInfo(0)
            txtInfo(txt�绰).Text = aryInfo(1)
            txtInfo(txt��ַ).Text = aryInfo(2)

            Call ReadPatPricture(Val(.FocusedRow.Record.Item(mCol.����ID).Value), imgPati)
            fraInfo(0).Visible = True
            txtInfo(txt��������).Text = aryInfo(3)
            txtInfo(txtȷ������).Text = aryInfo(4)
            txtInfo(txt�������1).Text = aryInfo(5)
            txtInfo(txt�������2).Text = aryInfo(6)
            txtInfo(txt���Ƽ���).Text = .FocusedRow.Record.Item(mCol.���Ƽ���).Value
            fraInfo(1).Visible = True
        End If
    End With
End Sub

Public Sub ReadPatPricture(ByVal lng����ID As Long, ByRef imgPatient As Image)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ƭ
    '������lng����ID=��ȡָ�����˵���Ƭ
    '           imgPatient=��Ƭ����λ��
    '           strFile=��Ƭ�ı���·��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFile As String
    On Error GoTo ErrHand
    If lng����ID = 0 Then Exit Sub
    strFile = ""
    strFile = gobjComlib.sys.Readlob(glngSys, 27, lng����ID, strFile)
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
        lngPatiID = objHisPati.����ID
    End If

    Call ExecuteFindPati(False, lngPatiID)
End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
     mstrFindType = objCard.����
End Sub

Private Sub PatiIdentify_KeyPress(KeyAscii As Integer)
    Select Case mstrFindType
        Case "סԺ��"
            If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        Case "�����"
            If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        Case "���￨"
            If InStr(":��';��?��", Chr(KeyAscii)) > 0 Then
                KeyAscii = 0
            Else
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End If
    End Select
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean, Optional ByVal lngPatiID As Long)
'���ܣ�����(��һ��)����
'������blnNext=�Ƿ������һ��
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long

    '��ʼ������
    If rptList.SelectedRows.Count > 0 Then
        If Not rptList.SelectedRows(0).GroupRow Then
            If Val(rptList.SelectedRows(0).Record(mCol.����ID).Value) <> 0 Then blnHave = True
        End If
    End If

    If Not blnNext Or blnReStart Or Not blnHave Then
        i = 0       'ReportControl����������0��ʼ
    Else
        i = rptList.SelectedRows(0).Index + 1
    End If

    '���Ҳ���
    For i = i To rptList.Rows.Count - 1
        With rptList.Rows(i)
            If Not .GroupRow Then
                If Val(.Record(mCol.����ID).Value) = lngPatiID And lngPatiID <> 0 Then Exit For

                If mstrFindType = "סԺ��" Then 'סԺ��
                    If UCase(Trim(.Record(mCol.�����).Value)) = UCase(PatiIdentify.Text) And .Record(mCol.��Դ).Value <> "����" Then Exit For
                ElseIf mstrFindType = "�����" Then
                    If UCase(Trim(.Record(mCol.�����).Value)) = UCase(PatiIdentify.Text) And .Record(mCol.��Դ).Value <> "סԺ" Then Exit For
                ElseIf mstrFindType = "����" OR mstrFindType = "��������￨" Then '����
                    If .Record(mCol.����).Value Like "*" & PatiIdentify.Text & "*" Then Exit For
                End If
            End If
        End With
    Next

    If i <= rptList.Rows.Count - 1 Then
        blnReStart = False
        '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        Set rptList.FocusedRow = rptList.Rows(i)
        If rptList.Visible Then rptList.SetFocus
    Else
        MsgBox IIf(blnNext, "������", "") & "�Ҳ������������Ĳ��ˡ�", vbInformation, gstrSysName
    End If
End Sub

Private Function DeleteReport(ByVal strID As String) As Boolean
'���ܣ�ɾ������
'������strID ��Ҫɾ���ı����ID
    Dim strSQL As String
On Error GoTo ErrHand

    strSQL = "Zl_�����걨��¼_Delete(" & strID & ")"
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
'���ܣ�����ѯ�����ı������Ϣ���ص�ReportControl�ؼ���
    Dim strPatiInfo As String
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem

    If rsTemp Is Nothing Then Exit Sub
    If rsTemp.RecordCount > 0 Then
        Do While Not rsTemp.EOF
            strPatiInfo = CStr(NVL(rsTemp!ְҵ)) & "|" & CStr(NVL(rsTemp!��ͥ�绰)) & "|" & CStr(NVL(rsTemp!��ͥ��ַ)) & "|" & CStr(NVL(rsTemp!��������)) & "|" & CStr(NVL(rsTemp!ȷ������)) & "|" & CStr(NVL(rsTemp!�������1)) & "|" & CStr(NVL(rsTemp!�������2))
            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(CStr(NVL(rsTemp!״̬)))
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
            Select Case rsTemp!״̬
                Case -1: rptRcd.AddItem CStr("�����")
                Case 0: rptRcd.AddItem CStr("�����")
                Case 1: rptRcd.AddItem CStr("�����")
                Case 2: rptRcd.AddItem CStr("�ѱ���")
                Case 3: rptRcd.AddItem CStr("���ϱ�")
                Case 4: rptRcd.AddItem CStr("������")
                Case 5: rptRcd.AddItem CStr("���޴����")
                Case 6: rptRcd.AddItem CStr("����д")
                Case 7: rptRcd.AddItem CStr("�Ǵ�Ⱦ��")
                Case 8: rptRcd.AddItem CStr("��ɾ��")
                Case 9: rptRcd.AddItem CStr("������")
                Case Else: rptRcd.AddItem ""
            End Select
            rptRcd.AddItem CStr(NVL(rsTemp!����))
            rptRcd.AddItem IIf(Val(rsTemp!������Դ) = 1, "����", IIf(Val(rsTemp!������Դ) = 2, "סԺ", ""))
            rptRcd.AddItem CStr(NVL(rsTemp!����))
            rptRcd.AddItem CStr(NVL(rsTemp!�����))
            rptRcd.AddItem CStr(NVL(rsTemp!����))
            rptRcd.AddItem CStr(NVL(rsTemp!�Ա�))
            rptRcd.AddItem CStr(NVL(rsTemp!����))
            rptRcd.AddItem CStr(NVL(rsTemp!�ʱ��))
            rptRcd.AddItem CStr(NVL(rsTemp!���))
            rptRcd.AddItem CStr(NVL(rsTemp!��������))
            rptRcd.AddItem CStr(NVL(rsTemp!���Ƽ���))
            rptRcd.AddItem CStr(NVL(rsTemp!�Ǽ���))
            rptRcd.AddItem CStr(NVL(rsTemp!�Ǽ�ʱ��))
            rptRcd.AddItem CStr(NVL(rsTemp!���ע))
            rptRcd.AddItem CStr(NVL(rsTemp!������))
            rptRcd.AddItem CStr(NVL(rsTemp!����ʱ��))
            rptRcd.AddItem CStr(NVL(rsTemp!������))
            rptRcd.AddItem CStr(NVL(rsTemp!����ʱ��))
            rptRcd.AddItem CStr(NVL(rsTemp!���͵�λ))
            rptRcd.AddItem CStr(NVL(rsTemp!���ͱ�ע))
            rptRcd.AddItem CStr(NVL(rsTemp!����ת��, 0))
            rptRcd.AddItem CStr(NVL(rsTemp!����ID, 0))
            rptRcd.AddItem CStr(NVL(rsTemp!��ҳID, 0))
            rptRcd.AddItem CStr(NVL(rsTemp!�ļ�ID, 0))
            rptRcd.AddItem CStr(NVL(rsTemp!�༭��ʽ, 0))
            rptRcd.AddItem CStr(strPatiInfo)
            rsTemp.MoveNext
        Loop
    End If
End Sub

Private Function GetReportFiles() As String
'���ܣ���ȡ����ı����ļ�
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
'���ܣ�ˢ����ʾ�����ļ�
'������strCurId ��Ҫ��λ���ı����ļ�ID
    Dim rptRow As ReportRow
On Error GoTo ErrHand
    If mblnReportCheck Then Exit Function
    Me.rptList.Records.DeleteAll
    If Trim(mstrFiles) = "" Then Exit Function
    
    dkpMain.FindPane(conPane_Feedback).Close
    If mTcbSelectID = tcbδ��д Then
        Call zlRefWaitFullData
    ElseIf mTcbSelectID = tcb�ϱ� Or mTcbSelectID = tcb��� Then
        Call zlRefOldData(mintDates, mdtFrom, mdtTo)
    ElseIf mTcbSelectID = tcb��ɾ�� Then
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

    Me.stbThis.Panels(2).Text = "����" & Me.rptList.Records.Count & "�ݼ������档"
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
End Function

Private Sub zlRefWaitFullData()
'���ܣ���ѯδ��д�ı����ļ�
    Dim blnMoved As Boolean, strTemp As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strSqlDate As String
On Error GoTo ErrHand

    If mTcbSelectID <> tcbδ��д Then Exit Sub
    If mstrState = "" Then Exit Sub

    If mintWaitDays <> -1 Then
        strSqlDate = " And a.�Ǽ�ʱ�� >= trunc(Sysdate - [1])"
    Else
        blnMoved = MovedByDate(CDate(mdtWaitBegin))
        strSqlDate = " And a.�Ǽ�ʱ�� Between [2] And [3]"
    End If
    
    strSQL = "Select A.ID as ID, null As �ļ�id, a.����id, a.��ҳid, '��Ⱦ�����Խ��������' As ����, Nvl2(A.�Һŵ�, 1, 2) as ������Դ , e.���� As ����," & vbNewLine & _
           "Nvl(c.סԺ��, b.�����) As �����, Nvl(c.����, b.����) As ����, Nvl(c.�Ա�, b.�Ա�) As �Ա�, Nvl(c.����, b.����) As ����, Null As �ʱ��," & vbNewLine & _
           "Null As ���, Null As ��������, 3 As �༭��ʽ, Decode(A.��¼״̬,1,9,3,7,6) As ״̬, Null As ������,Null as ������ ,Null as ����ʱ��, " & vbNewLine & _
           "Null As ����ʱ��, Null As ���͵�λ, Null As ���ͱ�ע, Null As �Ǽ���, Null As �Ǽ�ʱ��, F.ְҵ," & vbNewLine & _
           "F.��ͥ��ַ, F.��ͥ�绰, Null As ��������, Null As ȷ������,  Null As �������1, Null As �������2," & vbNewLine & _
           "Null As ���ע, a.��Ⱦ������ as  ���Ƽ��� , 0 As ����ת�� " & vbNewLine & _
           "From �������Լ�¼ A, ���˹Һż�¼ B, ������ҳ C, ���ű� E,������Ϣ F " & vbNewLine & _
           "Where a.����id = c.����id(+) And a.��ҳid = c.��ҳid(+) And a.����id = b.����id(+) And a.�Һŵ� = b.No(+) " & vbNewLine & _
           mstrState & " And Nvl(b.ִ�в���id, c.��Ժ����id) = e.Id and A.����ID =F.����ID  And a.�ļ�id Is Null " & strSqlDate
    If blnMoved Then
        strTemp = Replace(strTemp, "�������Լ�¼", "H�������Լ�¼")
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
'���ܣ���ѯ�ϰ���Ӳ����ı����ļ�
'������intDates     ��ѯ��������ı����ļ�
'      dtFrom  ��ѯָ��ʱ��α����ļ�����ʼ����
'      dtTo    ��ѯָ��ʱ��α����ļ�����ֹ����
    Dim blnMoved As Boolean, strTemp As String, strFiles As String
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strDelSql As String
On Error GoTo ErrHand

    strFiles = GetReportFiles()

    If Trim(strFiles) = "" Then Exit Sub
    If mstrState = "" Then Exit Sub
    If mTcbSelectID = tcb�ϱ� Or mTcbSelectID = tcb��� Then
        strDelSql = " and S.������ Is Null And S.����ʱ�� Is Null "
    ElseIf mTcbSelectID = tcb��ɾ�� Then
        strDelSql = " and S.������ Is Not Null And S.����ʱ�� Is Not Null "
    End If

    If intDates <> -1 Then
        strSQL = " And l.���ʱ�� >= trunc(Sysdate - [1])  "
    Else
        blnMoved = MovedByDate(CDate(dtFrom))
        strSQL = " And l.���ʱ�� Between [2] And [3]"
    End If

    strSQL = "Select l.Id, l.�ļ�id, l.����id, l.��ҳid, l.�������� As ����,l.������Դ, d.���� As ����," & vbNewLine & _
            " Decode(l.������Դ, 1, p.�����, 2, p.סԺ��) As �����, Nvl(l.����, p.����) As ����, Nvl(l.�Ա�, p.�Ա�) As �Ա�, Nvl(l.����, p.����) As ����," & vbNewLine & _
            " To_Char(l.���ʱ��, 'yyyy-mm-dd hh24:mi') As �ʱ��, l.������ As ���, l.��������, l.�༭��ʽ," & vbNewLine & _
            " l.״̬ As ״̬, l.������," & vbNewLine & _
            " To_Char(l.����ʱ��, 'yyyy-mm-dd hh24:mi') as ����ʱ��, l.���͵�λ,l.���ͱ�ע,l.�Ǽ���," & vbNewLine & _
            " To_Char(l.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi') as �Ǽ�ʱ��, Nvl(l.ְҵ,P.ְҵ) as ְҵ , Nvl(l.��ͥ��ַ,P.��ͥ��ַ) as ��ͥ��ַ, Nvl(l.��ͥ�绰,P.��ͥ�绰) as ��ͥ�绰 ," & vbNewLine & _
            " To_Char(l.��������, 'yyyy-mm-dd') as ��������,To_Char(l.ȷ������, 'yyyy-mm-dd') as ȷ������ , l.�������1 ,l.�������2," & vbNewLine & _
            " l.���ע ,L.������,L.����ʱ�� , null as  ���Ƽ���,0 As ����ת��" & vbNewLine & _
            " From (Select l.Id, l.�ļ�id, l.����id, l.��ҳid, l.��������, l.������Դ, l.����id, l.���ʱ��, l.������, l.����ʱ��, l.�༭��ʽ, decode(S.������, null ,Nvl(s.����״̬, 0),8) As ״̬," & vbNewLine & _
            "       s.������, s.����ʱ��, s.���͵�λ, s.���ͱ�ע, s.�Ǽ���, s.�Ǽ�ʱ��, s.����, s.�Ա�, s.����, s.ְҵ, s.��ͥ��ַ, s.��ͥ�绰," & vbNewLine & _
            "        s.��������, s.ȷ������, s.�������1, s.�������2, s.���ע, s.��������,s.������,S.����ʱ�� " & vbNewLine & _
            "       From ���Ӳ�����¼ L, �����걨��¼ S " & vbNewLine & _
            "       Where l.Id = s.�ļ�id(+) And l.�������� = 5 And l.�ļ�id In (" & strFiles & ") " & strSQL & _
               mstrState & strDelSql & " ) L, ������Ϣ P, ���ű� D" & vbNewLine & _
            "Where l.����id = p.����id And l.����id = d.Id "

    If blnMoved Then
        strTemp = Replace(strSQL, "0 as ����ת��", "1 as ����ת��")
        strTemp = Replace(strTemp, "���Ӳ�����¼", "H���Ӳ�����¼")
        strTemp = Replace(strTemp, "�����걨��¼", "H�����걨��¼")
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
'˫���޸ķ���˵��
    With rptSendContent
        If .FocusedRow Is Nothing Then
            mdatTime = 0
            Exit Sub
        ElseIf .FocusedRow.GroupRow = True Then
            mdatTime = 0
            Exit Sub
        Else
            mdatTime = Format(.FocusedRow.Record.Item(mSendRptCol.�Ǽ�ʱ��).Value, "yyyy-mm-dd HH:MM:SS")
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
            mdatTime = Format(.FocusedRow.Record.Item(mSendRptCol.�Ǽ�ʱ��).Value, "yyyy-mm-dd HH:MM:SS")
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
        If tbcMain.Selected.Tag = "���濨" Then
            picDis.Visible = False
            If IsNumeric(mstrCurId) Then
                Call RefreshReport(Val(mstrCurId), mblnCurMoved, NVL(rptList.FocusedRow.Record.Item(mCol.�༭��ʽ).Value, 0))
            End If
        ElseIf tbcMain.Selected.Tag = "������" Then
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
    Me.rptList.Columns(mCol.ɾ����).Visible = False
    Me.rptList.Columns(mCol.ɾ��ʱ��).Visible = False

    picState(mTcbSelectID).Visible = True
    txtInfo(txt��������).Visible = True
    txtInfo(txtȷ������).Visible = True
    txtInfo(txt�������1).Visible = True
    txtInfo(txt�������2).Visible = True
    lblInfo(txt��������).Visible = True
    lblInfo(txtȷ������).Visible = True
    lblInfo(txt�������1).Visible = True
    lblInfo(txt�������2).Visible = True

    If Item.Tag = "��˹���" Then
        picState(5).Visible = True
        picState(5).Move 100, 50
        picState(mTcbSelectID).Move 100, 450
        Call chkAduitState_Click(-1)
    ElseIf Item.Tag = "�ϱ�����" Then
        picState(5).Visible = True
        picState(5).Move 100, 50
        picState(mTcbSelectID).Move 100, 450
        Call chkSendState_Click(-1)
    ElseIf Item.Tag = "δ��д" Then
        picState(mTcbSelectID).Move 100, 50
        txtInfo(txt��������).Visible = False
        txtInfo(txtȷ������).Visible = False
        txtInfo(txt�������1).Visible = False
        txtInfo(txt�������2).Visible = False
        lblInfo(txt��������).Visible = False
        lblInfo(txtȷ������).Visible = False
        lblInfo(txt�������1).Visible = False
        lblInfo(txt�������2).Visible = False
        Call chkDisState_Click(-1)
    ElseIf Item.Tag = "��ɾ��" Then
        Me.rptList.Columns(mCol.ɾ����).Visible = True
        Me.rptList.Columns(mCol.ɾ��ʱ��).Visible = True
        picState(mTcbSelectID).Move 100, 50
        mstrState = " and S.������ is not null "
    ElseIf Item.Tag = "���ع���" Then
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
    cboSelectTime(0).Clear         'δ��д
    With cboSelectTime(0)
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
        .AddItem "30����"
        .ItemData(.NewIndex) = 30
        .AddItem "[ָ��...]"
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

    cboSelectTime(1).Clear         '��ɾ��
    With cboSelectTime(1)
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
         .AddItem "30����"
        .ItemData(.NewIndex) = 30
        .AddItem "[ָ��...]"
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

    cboSelectTime(2).Clear         '������ϱ�����
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
         .AddItem "30����"
        .ItemData(.NewIndex) = 30
        .AddItem "[ָ��...]"
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
'������Index 0δ��д 1��ɾ�� 2������ϱ�����
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
            cboSelectTime(Index).ToolTipText = "��Χ��" & Format(mdtWaitBegin, "yyyy-MM-dd") & " �� " & Format(mdtWaitEnd, "yyyy-MM-dd")
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
                    'ȡ��ʱ�ָ�ԭ����ѡ��
                    Call gobjComlib.zlControl.CboSetIndex(cboSelectTime(Index).hwnd, intOldIndex)
                    Exit Sub
                End If
            End If
        End If
        If mdtDelBegin = CDate(0) Or mdtDelEnd = CDate(0) Then
            cboSelectTime(Index).ToolTipText = ""
        Else
            cboSelectTime(Index).ToolTipText = "��Χ��" & Format(mdtDelBegin, "yyyy-MM-dd") & " �� " & Format(mdtDelEnd, "yyyy-MM-dd")
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
                    'ȡ��ʱ�ָ�ԭ����ѡ��
                    Call gobjComlib.zlControl.CboSetIndex(cboSelectTime(Index).hwnd, intOldIndex)
                    Exit Sub
                End If
            End If
        End If

        If mdtFrom = CDate(0) Or mdtTo = CDate(0) Then
            cboSelectTime(Index).ToolTipText = ""
        Else
            cboSelectTime(Index).ToolTipText = "��Χ��" & Format(mdtFrom, "yyyy-MM-dd") & " �� " & Format(mdtTo, "yyyy-MM-dd")
        End If
        If Me.Visible = True Then Call zlRefList
    End If
End Sub

Private Sub cmdDuplicateCheck_Click()
'���ܣ���ʾ���ص����б���
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
'���ܣ��������
'������rsOld �ϰ���Ӳ�����ѯ���������ݼ�
'      strName Ҫ��Ѱ���˵�����
'      strProfession  Ҫ��Ѱ���˵�ְҵ
'      strDiagnose    Ҫ��Ѱ���˵����
    Dim strSQL As String
    Dim blnMoved As Boolean
    Dim strTemp As String

On Error GoTo ErrHand
    If strProfession <> "" Then
        strSQL = " And a.ְҵ = [2] "
    End If
    If strDiagnose <> "" Then
        strSQL = strSQL & " And (a.�������1 Like [3] Or a.�������2 Like [3]) "
    End If
    blnMoved = MovedByDate(DateAdd("m", -12, CDate(gobjComlib.zlDatabase.Currentdate)))

    strSQL = "Select a.�ļ�id as ID, l.�������� As ����, Null As �����, a.����״̬ As ״̬, a.������, a.����ʱ��, a.���͵�λ, a.���ͱ�ע," & vbNewLine & _
            "       a.�Ǽ���, a.�Ǽ�ʱ��, a.����, a.�Ա�, a.����, a.ְҵ, a.��ͥ��ַ, a.��ͥ�绰, a.��������, a.ȷ������, a.�������1, a.�������2, a.���ע, a.��������," & vbNewLine & _
            "       a.����ҽ��, l.������ As ���, l.����ʱ�� As �ʱ��, l.������Դ, b.���� As ����, NULL As ���Ƽ���, Null As �����, l.�༭��ʽ As �༭��ʽ, a.������, a.����ʱ��," & vbNewLine & _
            "       l.����id, l.��ҳid, l.�ļ�ID, 0 As ����ת��" & vbNewLine & _
            "From �����걨��¼ A, ���Ӳ�����¼ L, ���ű� B" & vbNewLine & _
            "Where a.������ Is Null And a.����ʱ�� Is Null And a.�ļ�id = l.Id And l.�������� = 5 and a.����״̬= 3 And b.Id = l.����id " & vbNewLine & _
            " and a.���� Like [1] " & strSQL

    If blnMoved Then
        strTemp = Replace(strSQL, "0 As ����ת��", "1 as ����ת��")
        strTemp = Replace(strTemp, "���Ӳ�����¼", "H���Ӳ�����¼")
        strTemp = Replace(strTemp, "�����걨��¼", "H�����걨��¼")
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
'���ܣ����������ʾ
'������rsOld �ϰ���Ӳ�����ѯ���������ݼ�
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
        Me.stbThis.Panels(2).Text = "�����ҵ�" & Me.rptList.Records.Count & "�ݷ��������������档"
    Else
        Me.stbThis.Panels(2).Text = "û�в��ҵ����������ļ������档"
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

Public Function ShowDiseaseStation(ByVal frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
                ByVal intPatiFrom As Integer, ByVal lng����ID As Long, ByVal str����ID As String, ByVal str���ID As String, ByRef blnNotView As Boolean) As Boolean
'���ܣ���ѯָ����Աһ�����Ƿ���д����Ⱦ�����濨
'������lng����ID    ����ID

'      lng��ҳID    סԺΪ ��ҳID������Ϊ �Һ�ID
'      intPatiFrom  ������Դ סԺΪ 2�� ����Ϊ 1
'      lng����ID    ���� ID
'      str����ID    ����ID
'      str���ID    ���ID
    Dim blnMoved As Boolean, strTemp As String
    Dim strSQL As String
    Dim blnHasData As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim str��ҳids As String
    Dim blnDiag As Boolean
    Dim vMsg As String
    On Error GoTo errHand

    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mIntPatiFrom = intPatiFrom
    mlng����ID = lng����ID
    mstr����ID = str����ID
    mstr���ID = str���ID

    If mstr���ID = "" And mstr����ID = "" Then Exit Function

    mblnReportCheck = True
    blnMoved = MovedByDate(DateAdd("m", -12, CDate(gobjComlib.zlDatabase.Currentdate)))
    

    strSQL = "Select distinct l.Id, l.�ļ�id, l.����id, l.��ҳid, l.�������� As ����, l.������Դ,  d.���� As ����," & vbNewLine & _
            " Decode(l.������Դ, 1, p.�����, 2, p.סԺ��) As �����, Nvl(l.����, p.����) As ����, Nvl(l.�Ա�, p.�Ա�) As �Ա�, Nvl(l.����, p.����) As ����," & vbNewLine & _
            " To_Char(l.���ʱ��, 'yyyy-mm-dd hh24:mi') As �ʱ��, l.������ As ���, l.��������, l.�༭��ʽ," & vbNewLine & _
            " l.״̬ As ״̬,l.������," & vbNewLine & _
            " To_Char(l.����ʱ��, 'yyyy-mm-dd hh24:mi') as ����ʱ��, l.���͵�λ,l.���ͱ�ע,l.�Ǽ���," & vbNewLine & _
            " To_Char(l.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi') as �Ǽ�ʱ��, Nvl(l.ְҵ,P.ְҵ) as ְҵ , Nvl(l.��ͥ��ַ,P.��ͥ��ַ) as ��ͥ��ַ, Nvl(l.��ͥ�绰,P.��ͥ�绰) as ��ͥ�绰 ," & vbNewLine & _
            " To_Char(l.��������, 'yyyy-mm-dd') as ��������,To_Char(l.ȷ������, 'yyyy-mm-dd') as ȷ������ , l.�������1 ,l.�������2," & vbNewLine & _
            " l.���ע ,L.������,L.����ʱ�� , Q.��Ⱦ������ as  ���Ƽ���,0 As ����ת��" & vbNewLine & _
            " From (Select l.Id, l.�ļ�id, l.����id, l.��ҳid, l.��������, l.������Դ, l.����id, l.���ʱ��, l.������, l.����ʱ��, l.�༭��ʽ, decode(S.������, null ,Nvl(s.����״̬, 0),7) As ״̬," & vbNewLine & _
            "      s.������, s.����ʱ��, s.���͵�λ, s.���ͱ�ע, s.�Ǽ���, s.�Ǽ�ʱ��, s.����, s.�Ա�, s.����, s.ְҵ, s.��ͥ��ַ, s.��ͥ�绰," & vbNewLine & _
            "        s.��������, s.ȷ������, s.�������1, s.�������2, s.���ע, s.��������,s.������,S.����ʱ�� " & vbNewLine & _
            "       From ���Ӳ�����¼ L, �����걨��¼ S " & vbNewLine & _
            "       Where l.Id = s.�ļ�id(+) And l.�������� = 5 And l.����ID = [1] And trunc(l.���ʱ��) >=trunc(ADD_MONTHS(sysdate,-12)) " & vbNewLine & _
            " and S.������ Is Null And S.����ʱ�� Is Null ) L, ������Ϣ P, ���ű� D,�������Լ�¼ Q " & vbNewLine & _
            "Where l.����id = p.����id And l.����id = d.Id and l.id =Q.�ļ�ID(+) and l.����id =Q.����ID(+)"
    If blnMoved Then
        strTemp = Replace(strSQL, "0 as ����ת��", "1 as ����ת��")
        strTemp = Replace(strTemp, "���Ӳ�����¼", "H���Ӳ�����¼")
        strTemp = Replace(strTemp, "�����걨��¼", "H�����걨��¼")
        strTemp = Replace(strTemp, "�������Լ�¼", "H�������Լ�¼")
        strSQL = strSQL & " Union All " & strTemp
    End If
    Set mrsOld = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ�ò���һ���ڵĴ�Ⱦ�����濨", lng����ID)
    blnDiag = True
    If mrsOld.RecordCount > 0 Then
        '��ѯ�ò���һ���ڲ�ͬ�����
        Do While Not mrsOld.EOF
            If mrsOld!��ҳid & "" <> "" Then str��ҳids = str��ҳids & "," & mrsOld!��ҳid & ""
            mrsOld.MoveNext
        Loop
		mrsOld.MoveFirst 
        If str��ҳids <> "" Then
            str��ҳids = Mid(str��ҳids, 2)
            If str����ID <> "" Then
                strSQL = " Union Select �ļ�ID,����id,���id From ��������ǰ�� Where ����ID IN (Select Column_Value From Table(f_Num2list([2])))"
            End If
            If str���ID <> "" Then
                strSQL = strSQL & " Union Select �ļ�ID,����id,���id From ��������ǰ�� Where ���ID IN (Select Column_Value From Table(f_Num2list([3])))"
            End If
            strSQL = "(" & Mid(strSQL, 8) & ")"
            strSQL = "Select /*+ Rule*/ b.����id,b.���id,0 as ����ת��" & vbNewLine & _
                    "From �����ļ��б� A ,(" & strSQL & ") B Where A.ID=B.�ļ�ID  And" & vbNewLine & _
                    "(a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From ����Ӧ�ÿ��� C Where c.�ļ�id = a.Id And c.����id = [4]))"
            strSQL = "(" & strSQL & ") Minus Select A.����id, A.���id,0 as ����ת�� From ������ϼ�¼ A" & vbNewLine & _
                    "Where a.����id = [1] and a.��ҳid in (Select Column_Value From Table(f_Num2list([5]))) And Trunc(a.��¼����) >= Trunc(Add_Months(Sysdate, -12)) And a.������� = 1"
        
            If blnMoved Then
                strTemp = Replace(strSQL, "0 as ����ת��", "1 as ����ת��")
                strTemp = Replace(strTemp, "������ϼ�¼", "H������ϼ�¼")
                strSQL = strSQL & " Union All " & strTemp
            End If
        
            Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ�ò���һ���ڵĴ�Ⱦ�����", lng����ID, str����ID, str���ID, lng����ID, str��ҳids)
            blnDiag = rsTmp.RecordCount = 0
        End If
    End If
    
    If mrsOld.RecordCount > 0 And blnDiag Then
        blnHasData = True
        ShowDiseaseStation = True
        vMsg = zlCommFun.ShowMsgBox("��Ⱦ�����濨��ʾ", "��Ҫ��д��Ⱦ�����濨�����Ǹò����ڹ�ȥ��һ�����Ѿ���д����Ⱦ�����濨���Ƿ�鿴��д���ı��濨��", "!�鿴(&R),����(&W),?����(&Q)", frmParent, vbQuestion)
        If vMsg = "����" Then
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
'���ܣ���д����
    Dim clsDisease As New cDockDisease
    WriteReport = clsDisease.EditDiseaseDoc(Me, mlng����ID, mlng��ҳID, mIntPatiFrom, mlng����ID, mstr����ID, mstr���ID)
    Set clsDisease = Nothing
End Function

Private Function GetFeedbackIDs(ByVal lngID As Long) As Boolean
'���ܣ���ѯѡ�еı�������Է�����ID
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strDis As String
    Dim i As Long
On Error GoTo ErrHand
    If lngID = 0 Then Exit Function
    
    strSQL = "select A.ID,A.�Ǽ�ʱ��,B.���� as ����,A.��Ⱦ������ from  �������Լ�¼ A, ���ű� B where A.�ļ�ID = [1] and a.�Ǽǿ���id = B.ID "
    Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ�ñ����Ӧ�����Է�����", lngID)
    
    If rsTemp.RecordCount > 0 Then
         With Me.tbcDis
            .removeAll
            For i = 0 To rsTemp.RecordCount - 1
                If Not InStr(strDis, rsTemp!��Ⱦ������ & "") > 0 Then
                    strDis = ";" & rsTemp!��Ⱦ������ & strDis
                End If
                 .InsertItem(i, rsTemp!���� & "(" & rsTemp!�Ǽ�ʱ�� & ")", mfrmPreFeedBack.hwnd, 0).Tag = rsTemp!ID
                rsTemp.MoveNext
            Next
            strDis = Mid(strDis, 2)
            .InsertItem(i, "i", mfrmPreFeedBack.hwnd, 0).Tag = i
            .Item(1).Selected = True
            .Item(0).Selected = True
            .Item(i).Visible = False
            txtInfo(txt���Ƽ���).Text = strDis
             If Not rptList.FocusedRow Is Nothing Then
                If Not rptList.FocusedRow.GroupRow Then
                    rptList.FocusedRow.Record.Item(mCol.���Ƽ���).Value = strDis
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
'���ܣ�������Ϣ ��Ⱦ�����淵��
    Dim strSQL As String, strXML As String
    Dim rsPati As ADODB.Recordset
    Dim lng����ID As Long, lng����ID As Long
    Dim strTmp As String
    Dim lng������Դ As Long
    On Error GoTo errH

    lng����ID = Val(rptList.FocusedRow.Record.Item(mCol.����ID).Value)
    lng����ID = Val(rptList.FocusedRow.Record.Item(mCol.��ҳID).Value)
    lng������Դ = IIf(rptList.FocusedRow.Record.Item(mCol.��Դ).Value = "סԺ", 2, 1)
    strTmp = rptList.FocusedRow.Record.Item(mCol.����).Value

    If lng������Դ = 1 Then
        strSQL = "select a.����,null as סԺ��,a.�����,null as ����ID,a.ִ�в���id as ����ID,null as ����,a.���� from ���˹Һż�¼ a where a.ID=[2]"
    ElseIf lng������Դ = 2 Then
        strSQL = "select a.����,a.סԺ��,null as �����,a.��ǰ����id as ����ID,a.��Ժ����id as ����ID,a.��Ժ���� as ����,null as ���� from ������ҳ a where a.����ID=[1] and a.��ҳid=[2]"
    End If
    Set rsPati = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng����ID)
    strXML = "<patient_info><patient_id>" & lng����ID & "</patient_id><patient_name>" & rsPati!���� & "</patient_name>"
    strXML = strXML & "<in_number>" & rsPati!סԺ�� & "</in_number>"
    strXML = strXML & "<out_number>" & rsPati!����� & "</out_number>"
    strXML = strXML & "</patient_info><patient_clinic><patient_source>" & lng������Դ & "</patient_source>"
    strXML = strXML & "<clinic_id>" & lng����ID & "</clinic_id>"
    strXML = strXML & "<clinic_area_id>" & rsPati!����ID & "</clinic_area_id>"
    strTmp = "" & gobjComlib.sys.RowValue("���ű�", Val("" & rsPati!����ID), "����")
    strXML = strXML & "<clinic_area_title>" & strTmp & "</clinic_area_title>"
    strXML = strXML & "<clinic_dept_id>" & rsPati!����ID & "</clinic_dept_id>"
    strTmp = "" & gobjComlib.sys.RowValue("���ű�", Val("" & rsPati!����ID), "����")
    strXML = strXML & "<clinic_dept_title>" & strTmp & "</clinic_dept_title>"
    strXML = strXML & "<clinic_room>" & rsPati!���� & "</clinic_room>"
    strXML = strXML & "<clinic_bed>" & rsPati!���� & "</clinic_bed>"
    strXML = strXML & "</patient_clinic><disease_report><file_id>" & mstrCurId & "</file_id><doc_id>" & mstrCurId & "</doc_id>"
    strXML = strXML & "<report_name>" & rptList.FocusedRow.Record.Item(mCol.����).Value & "</report_name>"
    strXML = strXML & "<create_time>  </create_time>"   '����ʱ�䣬Ϊ��
    strXML = strXML & "<create_doctor>" & UserInfo.���� & "</create_doctor>"
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
'���ܣ�����ϱ������Ƿ���Գ����ϱ�
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngID  As Long
On Error GoTo ErrHand
    If Val(mstrCurId) = 0 Then Exit Function
    If IsNumeric(mstrCurId) Then
        lngID = CLng(Val(mstrCurId))
        strSQL = "select count(1) as num from �����걨���� where �걨ID = [1] "
        Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ�ñ����Ӧ�����Է�����", lngID)
        
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
'���ܣ��޸Ļ��������ϱ���ע˵��
'������ intType��1-������2-�޸�
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
