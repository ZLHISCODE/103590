VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "*\A..\ZLIDKIND\zlIDKind.vbp"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\ZL9PACSCONTROL\zl9PacsControl.vbp"
Begin VB.Form frmPacsMainNew 
   Caption         =   "Ӱ����վ"
   ClientHeight    =   10575
   ClientLeft      =   8535
   ClientTop       =   870
   ClientWidth     =   15240
   Icon            =   "frmPacsMainNew.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicFollowHistory 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   4560
      ScaleHeight     =   1215
      ScaleWidth      =   3375
      TabIndex        =   43
      Top             =   1800
      Visible         =   0   'False
      Width           =   3375
      Begin RichTextLib.RichTextBox rftHistoryFollow 
         Height          =   615
         Left            =   240
         TabIndex        =   44
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1085
         _Version        =   393217
         BackColor       =   12648447
         BorderStyle     =   0
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmPacsMainNew.frx":1CFA
      End
   End
   Begin VB.Timer TimerHistory 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7920
      Top             =   840
   End
   Begin MSComctlLib.ImageList img24 
      Left            =   5400
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":1D8D
            Key             =   "PACS����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":2507
            Key             =   "��Ƭ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":2C81
            Key             =   "PACS��д"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":33FB
            Key             =   "PACS���"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":3B75
            Key             =   "PACS�鿴������Ϣ"
         EndProperty
      EndProperty
   End
   Begin VB.Timer timFun 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   7320
      Top             =   840
   End
   Begin VB.PictureBox PicFucs 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3120
      ScaleHeight     =   855
      ScaleWidth      =   2175
      TabIndex        =   37
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Image imgFun 
         Height          =   495
         Index           =   3
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   240
         Width           =   495
      End
      Begin VB.Image imgFun 
         Height          =   495
         Index           =   2
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   240
         Width           =   495
      End
      Begin VB.Image imgFun 
         Height          =   495
         Index           =   1
         Left            =   600
         Stretch         =   -1  'True
         Top             =   240
         Width           =   495
      End
      Begin VB.Image imgFun 
         Height          =   495
         Index           =   0
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Timer TimFlicker 
      Interval        =   500
      Left            =   7920
      Top             =   120
   End
   Begin VB.PictureBox picFollowUp 
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   17
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Label Label3 
         Caption         =   "���"
         Height          =   495
         Left            =   0
         TabIndex        =   19
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox picExtra 
      Height          =   1935
      Left            =   7080
      ScaleHeight     =   1875
      ScaleWidth      =   2715
      TabIndex        =   16
      Top             =   5400
      Width           =   2775
      Begin RichTextLib.RichTextBox rtxtAppend 
         Height          =   1575
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2778
         _Version        =   393217
         BackColor       =   16635590
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmPacsMainNew.frx":426F
      End
   End
   Begin VB.PictureBox picDataSearchContainer 
      BackColor       =   &H00E0E0E0&
      Height          =   2415
      Left            =   7080
      ScaleHeight     =   2355
      ScaleWidth      =   4635
      TabIndex        =   15
      Top             =   7680
      Width           =   4695
      Begin VB.PictureBox picDataSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   -3960
         ScaleHeight     =   4095
         ScaleMode       =   0  'User
         ScaleWidth      =   5200
         TabIndex        =   26
         Top             =   -2520
         Width           =   5200
      End
      Begin VB.CommandButton cmdMore 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPacsMainNew.frx":430C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "��ʾȫ����ѯ����"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1920
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPacsMainNew.frx":47C2
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "���ò�ѯ����"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdDo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��  ѯ"
         Height          =   735
         Left            =   1920
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPacsMainNew.frx":4CB4
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "��ѯ"
         Top             =   120
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Timer timerVideoEvent 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   9015
      Top             =   165
   End
   Begin VB.Timer timerCapture 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   8505
      Top             =   135
   End
   Begin VB.PictureBox picWindow 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   9240
      ScaleHeight     =   4575
      ScaleWidth      =   5535
      TabIndex        =   3
      Top             =   600
      Width           =   5535
      Begin zl9PacsControl.TranControl tcDisable 
         Height          =   975
         Left            =   4560
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1720
         AlphaValue      =   25
      End
      Begin VB.PictureBox picLoadState 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   960
         ScaleHeight     =   1095
         ScaleWidth      =   3855
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   3855
         Begin VB.PictureBox picSmile 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   360
            Left            =   240
            Picture         =   "frmPacsMainNew.frx":5386
            ScaleHeight     =   360
            ScaleWidth      =   360
            TabIndex        =   8
            Top             =   240
            Width           =   360
         End
         Begin VB.Label labLoadState 
            Caption         =   " ���ڼ��ع���ģ�飬�����ĵȴ�..."
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   600
            TabIndex        =   7
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.PictureBox picReportContainer 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00808080&
         Height          =   2055
         Left            =   3720
         ScaleHeight     =   2055
         ScaleWidth      =   1815
         TabIndex        =   5
         Top             =   2520
         Visible         =   0   'False
         Width           =   1815
      End
      Begin XtremeSuiteControls.TabControl TabWindow 
         Height          =   2415
         Left            =   600
         TabIndex        =   4
         Top             =   720
         Width           =   4125
         _Version        =   589884
         _ExtentX        =   7276
         _ExtentY        =   4260
         _StockProps     =   64
      End
   End
   Begin DicomObjects.DicomViewer dcmRelateViewer 
      Height          =   1095
      Left            =   12600
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      _Version        =   262147
      _ExtentX        =   4471
      _ExtentY        =   1931
      _StockProps     =   35
   End
   Begin VB.Timer TimerRefresh 
      Enabled         =   0   'False
      Left            =   7320
      Top             =   120
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   10215
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   4154
            MinWidth        =   4154
            Picture         =   "frmPacsMainNew.frx":5DFD
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10345
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   6675
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   36
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":6691
            Key             =   "����"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":773B
            Key             =   "����"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":87E5
            Key             =   "���"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":988F
            Key             =   "��д"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":A939
            Key             =   "���"
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":B9E3
            Key             =   "���"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":CA8D
            Key             =   "���"
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":DB37
            Key             =   "����"
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":EBE1
            Key             =   "�ܾ�"
            Object.Tag             =   "9"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   6060
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":FC8B
            Key             =   "��ѡ����"
            Object.Tag             =   "90000"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":10225
            Key             =   "��ѡѡ��"
            Object.Tag             =   "90001"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":107BF
            Key             =   "��ѡ����"
            Object.Tag             =   "90002"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainNew.frx":10ED1
            Key             =   "��ѡѡ��"
            Object.Tag             =   "90003"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8100
      Left            =   0
      ScaleHeight     =   8100
      ScaleWidth      =   6540
      TabIndex        =   1
      Top             =   1320
      Width           =   6540
      Begin XtremeSuiteControls.TabControl tabScheme 
         Height          =   735
         Left            =   4080
         TabIndex        =   45
         Tag             =   "0"
         Top             =   2160
         Width           =   1215
         _Version        =   589884
         _ExtentX        =   2143
         _ExtentY        =   1296
         _StockProps     =   64
      End
      Begin VB.CommandButton cmdLocate 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2640
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPacsMainNew.frx":115E3
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "��λ"
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdFind 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPacsMainNew.frx":11A15
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "����"
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.PictureBox pic�������ڵ� 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   4560
         ScaleHeight     =   1095
         ScaleWidth      =   1455
         TabIndex        =   35
         Top             =   5160
         Width           =   1455
         Begin VB.Label labNoScheme 
            AutoSize        =   -1  'True
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Left            =   360
            TabIndex        =   36
            Top             =   480
            Width           =   1080
         End
      End
      Begin VB.PictureBox PicLine 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   50
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   5775
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3120
         Width           =   5775
      End
      Begin VB.PictureBox picDetail 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   3735
         TabIndex        =   22
         Top             =   4200
         Width           =   3735
         Begin VB.Label labPatientAge 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1080
            TabIndex        =   42
            Top             =   240
            Width           =   120
         End
         Begin VB.Label LabFlag���� 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF0000&
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   3375
            TabIndex        =   38
            Top             =   30
            Width           =   270
         End
         Begin VB.Image imgStep 
            Height          =   375
            Left            =   120
            Top             =   120
            Width           =   495
         End
         Begin VB.Label labCollectionInfo 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "����"
               Size            =   7.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Left            =   840
            TabIndex        =   34
            Top             =   480
            Width           =   75
         End
         Begin VB.Label labPatientInfo 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   840
            TabIndex        =   33
            Top             =   120
            Width           =   120
         End
         Begin VB.Label LabFlag��Ⱦ��״̬ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   3135
            TabIndex        =   32
            Top             =   30
            Width           =   270
         End
         Begin VB.Label LabFlagΣ��״̬ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF00FF&
            Caption         =   "Σ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   2895
            TabIndex        =   31
            Top             =   30
            Width           =   270
         End
         Begin VB.Label LabFlag��ɫͨ�� 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000C000&
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   2655
            TabIndex        =   30
            Top             =   30
            Width           =   270
         End
         Begin VB.Label LabFlagӤ�� 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Ӥ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   2415
            TabIndex        =   29
            Top             =   30
            Width           =   270
         End
         Begin VB.Label LabFlag���� 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   2175
            TabIndex        =   28
            Top             =   30
            Width           =   270
         End
         Begin VB.Image imgState 
            Height          =   255
            Index           =   0
            Left            =   3000
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.PictureBox picEvent 
         Height          =   615
         Left            =   120
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   18
         Top             =   6840
         Visible         =   0   'False
         Width           =   1335
         Begin VB.Label lab1 
            Caption         =   "��������"
            Height          =   495
            Left            =   120
            TabIndex        =   20
            Top             =   0
            Width           =   975
         End
      End
      Begin XtremeSuiteControls.TabControl TabExtra 
         Height          =   855
         Left            =   1200
         TabIndex        =   14
         Top             =   5640
         Width           =   1455
         _Version        =   589884
         _ExtentX        =   2566
         _ExtentY        =   1508
         _StockProps     =   64
      End
      Begin VB.PictureBox picTemp 
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   1320
         ScaleHeight     =   585
         ScaleWidth      =   825
         TabIndex        =   13
         Top             =   3360
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.PictureBox picFilter 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   2895
         TabIndex        =   12
         Top             =   120
         Width           =   2895
         Begin XtremeCommandBars.CommandBars cbrFilter 
            Left            =   0
            Top             =   120
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
         End
      End
      Begin VB.PictureBox ptemp 
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   11
         ToolTipText     =   "��û��ʲô��"
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Bindings        =   "frmPacsMainNew.frx":11E47
         Height          =   1695
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   3255
         _cx             =   1996035693
         _cy             =   1996032942
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
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
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
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   300
         Left            =   0
         TabIndex        =   21
         Top             =   840
         Width           =   2385
         _ExtentX        =   4207
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
         IDKindStr       =   $"frmPacsMainNew.frx":11E6F
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
         CaptionAlignment=   0
         ShowPropertySet =   -1  'True
         DefaultCardType =   "���￨"
         IDkindBorderStyle=   1
         IDKindWidth     =   1800
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
         AllowAutoCommCard=   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
      End
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPacsMainNew.frx":11F22
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPacsMainNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

#Const DebugImmediately = False

Private Const C_LAYOUT_BASEHEIGHTOFTAB As Long = 5000 ' ������Ϣ5000
Private Const C_LAYOUT_BASEHEIGHTOFDETAILINFO As Long = 800 ' ��ϸ��Ϣ��׼�߶�5000

Private Const C_STEPIMG_�Ǽ� As String = "����" '
Private Const C_STEPIMG_���� As String = "����" '
Private Const C_STEPIMG_��� As String = "���" '
Private Const C_STEPIMG_�ܾ� As String = "�ܾ�" '
Private Const C_STEPIMG_���� As String = "����" '
Private Const C_STEPIMG_��� As String = "���" '
Private Const C_STEPIMG_��� As String = "���" '
Private Const C_STEPIMG_��� As String = "���" '
Private Const C_STEPIMG_��д As String = "��д" '

Private mobjCurStudyInfo As New clsStudyInfo  '���ڲ����ļ����Ϣ
Private mobjHistoryStudyInfo As New clsStudyInfo  '������ʷ���ļ����Ϣ
Private mstrFirstTab As String '�״���ʾ��ҳ��
Private mlngMove As Long
Private mintQueryState As Integer '��ѯ����״̬  0 δ��ʼ��  ��1 ����  ��2 û���κ���Ч����   3��û���Ѿ����õķ���
Private mblHistory As Boolean '�Ƿ����μ��
Private mblHaveHistory As Boolean '������ʷ���
Private mintAutoRefreshTimer As Integer '�Զ�ˢ�¼�ʱ����
Private mintAutoRefreshTimerCount As Integer '�Զ�ˢ�¼�ʱ����
Private mobjPublicPACS As Object             'PACSҵ���װ��������
Private mobjPacsInterface As Object

Private mlngPicHistoryX As Long
Private mlngPicHistoryY As Long
Private mlngpicHistoryOldW As Long
Private mlngpicHistoryOldH As Long

Private mblnpicHistoryMoving As Boolean
'--------------------------------------------------

Private Const M_BLN_ALL_FUNCTIONS_OPEN As Boolean = True
Private Const M_STR_MODULE_MENU_TAG As String = "Main"

'��û������ʱ��ʹ�ô���ʾ��Ϣ
Private Const M_STR_HINT_NoSelectData As String = "��ѡ����Ҫִ�еļ�����ݡ�"

'��˸��ʱ��Ϣ
Private Type TFlickerInfo
    LngSchemeNo As Long '��ǰ������
    strName As String '��˸�ֶ��� �磺 "������"
    strInfo As String '��ϸ��Ϣ ��"�ѵǼ�,����ʱ��,30|�ѱ���,����ʱ��,40|"
End Type

'ϵͳ�������Ͷ���
Private Type TSystemPar

    '���ز���
    blnLockAfterCall As Boolean                         '�Ƿ���к������ɼ�
    strFirstTab As String                               '�״���ʾ��ҳ��
    blnֱ�Ӽ�� As Boolean                              '�ǼǺ�ֱ�ӽ�����
    blnWriteCapDoctor As Boolean                        '�Ƿ��ڲɼ�ͼ����Զ��ѵ�ǰ�û���дΪ��鼼ʦ
    blnAutoOpenReport As Boolean                        '��ʼ����Զ��򿪱���
    blnChoosePrintFormat As Boolean                    '�Ƿ񱨵���ӡǰѡ���ʽ
    strLocalRoom As String                              '����ִ�м�����
    lngImageValid As Long                               'ͼ��У��
    
    '���̲���
    blnCompleteCommit As Boolean                        '��˺������ٴ�ȷ��
    blnFinallyCompleteCommit As Boolean                 '�����ֱ�����
    blnIgnoreResult As Boolean                          '���������� '=true ����
    
    blnReportWithImage As Boolean                       '��ͼ�����д���棬��ͼ�񲻿�д����
    blnNoSignFinish As Boolean                              '����δǩ�������ӡ���
    blnReportWithResult As Boolean                      '�������Խ������д����
    
    blnPrintCommit As Boolean                           '��ӡ��ֱ�����
    blnCanPrint As Boolean                              'ƽ����Ҫ��˲��ܴ�ӡ =true
    blnAuditAutoPrint As Boolean                           '�����ֱ�Ӵ�ӡ
    lngBeforeDays As Long                               'Ĭ�ϲ�ѯ������
    blnUseQueue As Boolean                              '�Ƿ������Ŷӽк�
    blnSynStudylist As Boolean                          '�Ŷӽк�ʱ������Ŷ��б������б����ݺ��Ƿ�ͬ����λ������б�
    blnAutoInQueue As Boolean                           '�����Ŷӽкź��Ƿ��Զ����
    blnQueueQuick As Boolean                            '�����Ŷӽкź��Ƿ��Զ�������ݽкŴ���
    
    blnRelatingPatient As Boolean                       '�Ƿ����ù�������
    lngConformDetermine As Long                         '�������
    strImageLevel As String                             'Ӱ�������ȼ���
    strReportLevel As String                            '���������ȼ���
    lngImageLevel As Long                               'Ӱ�������ж�
    lngReportLevel As Long                              '���������ж�
    
    lngHintType As Long                                 '��Ͻ����ʾ����
    
    blnIsPetitionScan As Boolean                        '�Ƿ��������뵥ɨ��
    blnChangeUser As Boolean                            '�Ƿ������û�����
    blnSwitchUser As Boolean                            '�Ƿ������û��л�
    
    lngVideoStationMoneyExeModle As Long                '�ɼ�����ִ��ģʽ 0-����ʱִ�У�1-���ʱִ�У�2-����ʱִ��
    lngPacsStationMoneyExeModle As Long                 'ҽ������ִ��ģʽ 0-����ʱִ�У�1-����ʱִ��
    lngPatholStationMoneyExeModle As Long               '�������ִ��ģʽ 0-����ʱִ�У�1-���ʱִ�У�2-����ʱִ��
    
    lngListColorMark As Long                            'Ϊ0ʱ����б�ǰ��ɫ��Ϊ1ʱ����б���ɫ
    blnNameColColorCfg As Boolean                       '�Ƿ���ݲ������������б���������ɫ
    blnOrdinaryNameColColorCfg As Boolean               'ȱʡ���͵Ĳ����Ƿ���ݲ�����������������ɫ
    
    blnAutoSendWorkList As Boolean                      '�Ƿ񱨵�ʱ�Զ�����WorkList
    blnNameFuzzySearch As Boolean                       '�Ƿ�����Ĭ��ģ����ѯ
    blnNameQueryTimeLimit As Boolean                    '����������ʱ�Ƿ����ʱ������
    blnAutoPrint As Boolean    '�������Զ���ӡ���뵥
    blnAutoPrintCheck As Boolean                   '�Զ������ظ���ӡ
    
    blnShowImgAfterReport As Boolean                    '����ʱ��Ƭ
    blnIsLocateReport As Boolean
    blnPEISNoCheckMoneyFinish  As Boolean    '����鱨����ɲ��жϷ���
    blnQuickTabDisplayScheme  As Boolean    '���ÿ��tab��ǩչʾ����
    blnQueryValidTime  As Boolean                  '�л����Ҳ��ı�Ĭ��ʱ�䷶Χ
    strPDFFTPdevice As String
    blnPDFTested As Boolean 'PDF�Ƿ��Ѿ���ʼ�����ԣ�ÿ���л����Һ���Ҫ���²��ԣ��״�ʹ��PDFǰִ�в��ԡ�
    
    blnPopChangGuiWindow  As Boolean
    blnPopKuaiShuWindow As Boolean
    blnPopBingDongWindow As Boolean
    blnPopXiBaoWindow As Boolean
    blnPopHuiZhenWindow As Boolean
    blnPopShiJianWindow As Boolean
    
    bln�����Ǽ� As Boolean
End Type


'��Ƶ�ɼ��¼���Ϣ
Private Type TVideoEventInf
    vetEventType As TVideoEventType
    lngAdviceId As Long
    lngSendNo As Long
    strOtherInf As String
    dcmImage As DicomImage
End Type

'��Ƶ�ɼ���Ϣ����
Private Type TCaptureMsgInf
    lngMsg As Long
    lngVirtualKey As Long
    lngScanKey As Long
    lngFlags As Long
End Type

Private mintInterface() As TInterface   '�Զ�ִ�еĲ��
Private mintInterfaceCount As Integer '��Ҫ�Զ�ִ�еĲ��������1 ��ʼ����

Private mintToolBarWriteReg As Integer        '������ע���״ֵ̬

Private mstrPrivs As String, mlngModule As Long              'ģ��ţ���ģ��Ȩ��
Private mstrPublicAdvicePrivs As String                     '9001ģ��Ȩ��

'�Ӵ������
Private WithEvents mobjEvent As clsEvent            '�¼��������
Attribute mobjEvent.VB_VarHelpID = -1
Private WithEvents mfrmRISRequest As frmRISRequest
Attribute mfrmRISRequest.VB_VarHelpID = -1

'��Ϣ��������
Private WithEvents mobjMsgCenter As clsPacsMsgProcess
Attribute mobjMsgCenter.VB_VarHelpID = -1

'����ģ�������ˢ��ģʽ�����������
'1.����ģ��ֻҪ���ڣ�ǿ�ƶ����е����ݽ���ˢ��
'2.����ģ������ʾʱ���Ŷ����е����ݽ���ˢ��
'3.����ģ����������ݱ仯ʱ����ʾ��ģ���ǵ�ǰģ�飬�Ŷ����е����ݽ���ˢ��

Private mfrmWork_PacsImg As frmWork_Image           'Ӱ���Ӵ���
Attribute mfrmWork_PacsImg.VB_VarHelpID = -1
Private mobjWork_Pathol As clsWorkModule_Pathol     '�������ģ��
Private mobjWork_His As clsWorkModule_His           'HIS���ģ��
Private mobjAppendBill As Object


Private WithEvents mobjPacsQueryWrap As clsPacsQueryWrap      '�Զ����ѯ���ܷ�װ��
Attribute mobjPacsQueryWrap.VB_VarHelpID = -1
Private mobjWork_ImageCap As Object  ' zl9PacsImageCap.clsPacsCapture  '��Ƶ�ɼ�ģ��
Attribute mobjWork_ImageCap.VB_VarHelpID = -1
Private WithEvents mobjWork_Report As clsWorkModule_Report     '����ģ��
Attribute mobjWork_Report.VB_VarHelpID = -1
Private WithEvents mobjPacsCore As zl9PacsCore.clsViewer            '��Ƭվ����
Attribute mobjPacsCore.VB_VarHelpID = -1
Private WithEvents mobjQueue As frmWork_Queue  'zlQueueManage.cLSQueueManage          '�Ŷӽк�
Attribute mobjQueue.VB_VarHelpID = -1

Private WithEvents mobjPetitionCap As frmPetitionCapture                 '���뵥
Attribute mobjPetitionCap.VB_VarHelpID = -1

Private WithEvents mfrmHistory As frmHistoryStudy                  '���μ��
Attribute mfrmHistory.VB_VarHelpID = -1

Private mfrmPatholSpecimen As frmPatholSpecimen              '�걾����

Private mclsCISKernel As clsCISKernel   'ֻʹ���˸���鿴���뵥����
'���ڱ���
Private mlngCur����ID As Long                               '��ǰ����ID
Private mstrCur���� As String                               '��ǰ���� ����-����
Private mstrCanUse���� As String                            '��ǰ���ÿ���  ID_����-����
Private mlngFilterTab As Long                               '����tabҳ
Private mblnInitOk As Boolean   '��ʼ�����,װ�ر��
Private mblnLoadSubFrom As Boolean                          '�Ƿ����ڼ����Ӵ���
Private mblnAllDepts As Boolean                             '�Ƿ�ѡ��ȫ������
Private mstrCanUse����IDs As String                         '��ǰ���õĿ���ID�����á������ָ�������ֱ����ΪSQL��ѯ����
Private mblnMenuDownState As Boolean                        '����˫����������������
Private mblnIsLoadPatholModule As Boolean                   '�Ƿ������˲���ģ��
Private mblnFormLoadState As Boolean
Private mblnIsScheduleDept As Boolean                       '��ǰѡ�п��ң��Ƿ�����ԤԼ
Private mblnIsScheduleOrder As Boolean                      '��ǰ����Ƿ�����ԤԼ������ԤԼ�豸�ж�

Private mblnIsPrintMode As Boolean                          '�Ƿ����嵥��ӡ

Private mstrDefaultPatientType As String                    'ȱʡ��������
Private mlngOldAdviceId As Long                             'ǰһ��ѡ��ļ���ҽ��ID

Private mstrRPTExecutor As String                           '����ѡ��ı�����
Private mrtReportType As ReportType
Private mblnLockState As Boolean                           '�Ƿ����û���������״̬

'���̿��Ʊ���
Private mSysPar As TSystemPar                               'ϵͳ����

Private mblnObserve As Boolean                              '�Ƿ��й�Ƭ����Ȩ��   true��  false��
Private mintImgCount As Integer                             '��ɨ�����뵥����

Private WithEvents mobjCaptureHot As zl9PacsControl.clsHookKey
Attribute mobjCaptureHot.VB_VarHelpID = -1
Private mVideoEventInf As TVideoEventInf
Private mstrCaptureHot As String                                    '�ɼ��ȼ�����
Private mstrCaptureAfterHot As String                               '��̨�ɼ��ȼ�����
Private mstrCaptureAfterTagHot As String                            '��Ǹ����ȼ�����
Private mCaptureMsg As TCaptureMsgInf
Private mobjSquareCard As Object

'�������ز���
Private mstrSelQueueRooms As String                         'ֻ����ִ�м��ڵĲ���
Private mstrAllQueueRooms As String

Private mblnMoved As Boolean                                '��ǰʱ������Ƿ�ת�ƹ�
Private mstrWorkModule As String

Private mblnAssignment As Boolean
Private mlngLocateFindType As Long
Private mstrAllExamineRoomCfg As String    '���п���ִ�м�ѡ�����
Private mstrCurExamineRoomCfg As String    '��ǰ����ִ�м�ѡ�����

'��ʷ��¼����ʾ
Private mblnIsHistory As Boolean

'˫�û���¼
Private mcnOracleHIS As New ADODB.Connection    '��¼HIS����̨��½ʱʹ�õ����ݿ����Ӵ�
Private mstrUserNameHIS As String               '��¼HIS����̨��½ʱʹ�õ��û���
Private mstrUserIDHIS As String                 '��¼HIS����̨��¼ʱʹ�õ��û�ID
Private mstrUserNameNew As String               '��¼˫�û���½�ĵڶ����û���
Private mstrUserIDNew As String                 '��¼˫�û���¼�ĵڶ����û�ID
Private mblnCnOracleIsHIS As Boolean            '��ǰ���ݿ������Ƿ�HIS����̨������
Private mintChangeUserState As Integer          '��¼�û������������1- ͳһ��2-����

'�ղع���
Private mlngShareFatherID As Long
Private mlngCollectionFatherID As Long
Private mblnIsLoading As Boolean

Private mblnIsCallModuleRefresh As Boolean          '�Ƿ����ģ��ˢ�²���
Private mblnAutoRefreshList As Boolean          '�Ƿ��Զ�ˢ�¼���б�
Private mobjPublicAdvice As Object
Private mobjMedicalRecord As Object
Private mblnIsValid As Boolean

Private mintState As Integer
Private mblnRefreshWord As Boolean              '�Ƿ�ǿ��ˢ�´ʾ����

Property Get StartDate() As Date
    StartDate = mobjPacsQueryWrap.StartDate
End Property

Property Get EndDate() As Date
    EndDate = mobjPacsQueryWrap.EndDate
End Property

Property Get StudyInfo() As clsStudyInfo
    Set StudyInfo = mobjCurStudyInfo
End Property

Property Get IsValid() As Boolean
    IsValid = mblnIsValid
End Property


Private Sub DynamicCreateModuleObj()
On Error Resume Next
    '���������㲿��
    Set mobjSquareCard = CreateObject("zlOneCardComLib.clsOneCardComLib")
    
    'mobjAppendBill���mobjAppendBill��Ϊ�գ���ʾʹ�û��ģʽ
    Set mobjAppendBill = CreateObject("ZlSoft.HIS.Charge.AppendCharge")
err.Clear
End Sub

Public Sub ShowStation(ByVal lngModule As Long, Owner As Object)
    
    Dim t1 As Long
    Dim i As Integer
    
    mblnIsValid = True
    mblnInitOk = False
    mblnLoadSubFrom = False
    mlngModule = lngModule
    mblnAutoRefreshList = False
    mstrPublicAdvicePrivs = "-1"
    mintState = 0
    mblnLockState = False
    Set mrsDeptParas = Nothing  'ʹ���Ҳ����������½��м���
    
    Call DynamicCreateModuleObj
    
    '��ʼ�������㲿��
    If Not mobjSquareCard Is Nothing Then
        mobjSquareCard.zlInitComponents Me, mlngModule, glngSys, gstrDBUser, gcnOracle
    End If
    
    '��仰����ʡ�ԣ����һ�������������⣬ֻҪ��ʽ��ȷ���ɣ������ᱻ�޸�
    PatiIdentify.zlInit Me, glngSys, mlngModule, gcnOracle, gstrDBUser, mobjSquareCard, InitCardType("����;")

    Call WriteLog("ShowStation -> Step 1������Ӱ�������ڳ�ʼ�����̡�", "frmPacsWork")

    If Not mblnFormLoadState Then Call InitForm
    
    Call WriteLog("ShowStation -> Step 2", "frmPacsWork")
    
    '����ʾ����ǰϵͳ����
    Me.Show , Owner
    If Me.WindowState = 1 Then Me.WindowState = 0
    
    DoEvents
    
    Call WriteLog("ShowStation -> Step 3����ʼ��������ģ�顣", "frmPacsWork")
    '��������Ĺ���ģ��

    Call Me.InitSubForm

    DoEvents
    
    Call WriteLog("ShowStation -> Step 4��������ʾ��ģ�顣", "frmPacsWork")
    
    If Not TabWindow.Selected Is Nothing Then
        Call ConfigSubForm(TabWindow.Selected)
    End If
    
    mblnInitOk = True
    
    Call WriteLog("ShowStation -> Step 5��ˢ�������б�", "frmPacsWork")
    'ˢ�¼������
    
    If mintQueryState = 1 Then
        Call ExecuteDefaultQueryScheme
    Else
        If mSysPar.strFirstTab <> "" Then '��Ϊ�ձ�ʾ��������ҳ��ʾ,��TabWindow����ˢ��
                
            For i = 0 To TabWindow.ItemCount - 1
                If InStr(TabWindow.Item(i).tag, mSysPar.strFirstTab) > 0 And TabWindow.Item(i).Visible Then
                    Exit For
                End If
            Next
            
            If i = TabWindow.ItemCount Then    'ûѭ�����˴�����1������TAB
                For i = 0 To TabWindow.ItemCount - 1
                    If TabWindow.Item(i).Visible Then
                        Exit For
                    End If
                Next i
            End If
            
            'ˢ��ҳ�棬����ʾ������ҳ
            If TabWindow.Item(i).Selected Then
                Call RefreshTabWindow
            Else
                TabWindow.Item(i).Selected = True
            End If
        Else
            TabWindow.Item(0).Selected = True
        End If
    End If

    '������Ҫ����һ������ �����޸Ĺ��˲˵����� �����ɵ���Ϊֻ�Բ˵��ؼ��޸ġ�
    
    Call ReSetFormFontSize
    
    
    
    DoEvents
    Call WriteLog("ShowStation -> Step 6������ģ��˵���", "frmPacsWork")
    '����ģ��˵�
    Call CreateWorkModuleMenu

    'δ����ϵͳ�������ܿ�����Ƶ���棬��Ҫ����һ����ƵԤ��
    If Not mobjWork_ImageCap Is Nothing Then
        Call WriteLog("ShowStation -> Step 6.1��������ƵԤ����", "frmPacsWork")
        Call mobjWork_ImageCap.zlRePreview
    End If
    
    Call WriteLog("ShowStation -> Step End.������Ӱ�������ڳ�ʼ�����̡�", "frmPacsWork")
    
    Debug.Print "ShowStation��ʱ" & GetTickCount - t1
    
    If mSysPar.lngImageValid > 0 Then
        If Len(Dir(Mid(App.Path, 1, InStrRev(App.Path, "\")) & "zlPacsImageValid.exe")) > 0 Then
            If InitRegister Then
                Shell Mid(App.Path, 1, InStrRev(App.Path, "\")) & "zlPacsImageValid.exe   " & gstrServerName & "||" & gstrUserName & "||" & gstrUserPswd & "||" & mlngCur����ID & "||" & mSysPar.lngImageValid & "||" & "" & "||2", 1
            End If
        End If
    End If
End Sub


Public Function MainWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo errhandle
    Dim strLog As String
    
    '��Ϣ����
    Select Case uMsg
        Case WM_XWREPORT_IMG
            strLog = Now & " umsg = " & uMsg & ";wparam = " & wParam & ";lparam = " & lParam & vbCrLf
    
            If gblnXWLog = True Then
                Call WriteCommLog("XWWindowProc", "XW�ӿ�", strLog)
            End If
            
            '�����������͵�ϵͳ������ı���ͼ��
            If lParam <> 0 Then
                If gblnXWLog Then
                    Call WriteCommLog("XWWindowProc", "XW�ӿ�", "���뱨��ͼ������̡�")
                End If
    
                Call XWSaveReportImages(Me, lParam)
            End If
        Case WM_LIST_SYNCROW
'            MsgBox "ˢ��������"
        Case WM_LIST_REFRESH
'            MsgBox "ˢ���б�����"
    End Select

Exit Function
errhandle:

End Function


Private Sub Menu_File_Excel_click()
'����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
'����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
'       lngSelectedRow,��¼���ô�ӡ����ǰ��ѡ���У����嵥�رպ�ָ�
On Error GoTo errhandle
    Dim bytMode As Byte
    Dim lngSelectedRow As Long
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = vsfList
    objPrint.title.Text = "��鲡���嵥"
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & zlDatabase.Currentdate())
    Call objPrint.BelowAppRows.Add(objAppRow)

    '�� �Ƿ��Ǵ�ӡ�嵥������ֵ
    mblnIsPrintMode = True
    '�õ���ӡ�嵥ǰ�ĵ�ǰѡ����
    lngSelectedRow = vsfList.RowSel
    
    bytMode = zlPrintAsk(objPrint)
    If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    
    '��ӡ��Ԥ�������� �ָ�ѡ����
    vsfList.Row = lngSelectedRow
    mblnIsPrintMode = False
    
    
    Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_RichEPR(ByVal cbrID As Long)
'�Զ��򿪱���༭����ͬʱ������PACS����༭���͵��Ӳ����༭��
On Error GoTo errhandle
    Dim cbrControl As CommandBarControl, i As Long
    
    '���û��ѡ�������ݣ���ֱ���˳�ִ��
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    '����ҳ�治�ɼ�ʱ��ִ���κβ���
    If TabWindow.Selected.tag <> "������д" Then
        For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
            If TabWindow(i).tag = "������д" And TabWindow(i).Visible = True Then TabWindow(i).Selected = True
        Next
        If TabWindow.Selected.tag <> "������д" Then Exit Sub
    Else
        If TabWindow.Selected.Visible = False Then Exit Sub
    End If
    
    '�ҵ�����ҳ�棬�ٴ��������ҳ��
        'ˢ��Ƕ��ҳ������
        If Not mobjWork_Report Is Nothing Then
            Call mobjWork_Report.zlUpdateAdviceInf(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, mobjCurStudyInfo.blnMoved, mobjCurStudyInfo.lngBaby)
            Call mobjWork_Report.zlUpdateOtherInf(picReportContainer, vsfList, mblnIsHistory, mobjCurStudyInfo.blnCanPrint, mobjCurStudyInfo.strDoDoctor, mobjCurStudyInfo.strStudyUID)
            
            Call mobjWork_Report.zlRefreshFace
        End If
    
    '�жϰ���������
    Set cbrControl = Me.cbrMain.FindControl(, conMenu_PacsReport_Open + 1000000)
    
    If cbrControl Is Nothing Then
        Set cbrControl = Me.cbrMain.FindControl(, cbrID + 1000000)
        If cbrControl Is Nothing Then Exit Sub
    End If
    
    Call cbrMain_Update(cbrControl)
    If cbrControl.Enabled = False Then Exit Sub
        
    '����˫����ť����ı���������Ҫ�������ó�False����Ϊ��������ʱ�򿪱��洰�塱ʱ��ʵ���ϴ˱���ΪTrue
    mblnMenuDownState = False
    Call cbrMain_Execute(cbrControl)
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_File_Parmeter_click()
On Error GoTo errhandle
    With frmTechnicSetup
        .mlngModul = mlngModule
        .mlng����ID = mlngCur����ID
        .mstrPrivs = mstrPrivs
        .Show 1, Me
        
        If .mblnOk Then
            InitLocalPars
            
            If Not mobjWork_Report Is Nothing Then
                '���¼��غͱ�����ص����ò���
                Call mobjWork_Report.InitReportParameter
            End If
            mSysPar.blnAutoPrint = Val(zlDatabase.GetPara("�������Զ���ӡ���뵥", glngSys, mlngModule, 0)) '�������Զ���ӡ���뵥
            mSysPar.blnAutoPrintCheck = Val(zlDatabase.GetPara("�Զ�����ظ������ӡ", glngSys, mlngModule, 0))
            mSysPar.blnShowImgAfterReport = (Val(zlDatabase.GetPara("����ʱ��Ƭ", glngSys, mlngModule, 0)) = 1)
            mSysPar.blnAutoOpenReport = (Val(zlDatabase.GetPara("��ʼ����Զ��򿪱���", glngSys, mlngModule, 0)) = 1)
            
            
        End If
    End With
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub


'��ʾ��ݷ�ʽ����
Private Sub Menu_File_ShortcutSet_click()
    Dim frmShortcut As New frmShortcutConfig
    
On Error GoTo errhandle
    Dim lngCount As Long
    
    Call frmShortcut.ShowShortcutConfig(App.ProductName, mlngModule, Me)
        
    If frmShortcut.blnIsOk Then Call ReCreatCbrMenu(cbrMain)
    
    Call Unload(frmShortcut)
    Set frmShortcut = Nothing
Exit Sub
errhandle:
    Call Unload(frmShortcut)
    Set frmShortcut = Nothing
End Sub


Private Sub Menu_Help_About_click()
On Error GoTo errhandle
    ShowAbout Me, App.title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Help_click()
'���ܣ����ð�������
On Error GoTo errhandle
    ShowHelp App.ProductName, Me.hwnd, Me.Name
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Forum_click()
On Error GoTo errhandle
    Call zlWebForum(Me.hwnd)
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Help_Web_Mail_click()
On Error GoTo errhandle
    zlMailTo hwnd
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Manage_ȡ������()
'ȡ��������������ǣ�ÿ��ȡ��������ͼ��ȫ���������б���ɢ��N����ʱ��¼
On Error GoTo errhandle
    Dim lngResult As Long
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If

    lngResult = -1
    
    '�����ģ���Ϊ1298��RIS����վ����������������ݿ��ѯ��ƥ���ͼ���¼
    If mlngModule = G_LNG_PACSSTATION_MODULE And mobjCurStudyInfo.intImageLocation = 1 Then
        lngResult = XWShowMatched(Me, mobjCurStudyInfo.lngAdviceId)
    Else
        frmSelectMuli.ShowImageReleation mlngModule, mobjCurStudyInfo.lngAdviceId, mstrPrivs, mobjCurStudyInfo.blnMoved, IIf(mlngModule = G_LNG_PACSSTATION_MODULE, False, True), mlngCur����ID, 1
        
        If frmSelectMuli.mblnOk = True Then lngResult = 0
    End If
    
    If lngResult <> 0 Then Exit Sub
    
    Call AfterReleationImage(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, 1, True)

Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_��ɲ�����()
'���ģʽ��ʹ��
    Dim objPatholPrice As New frmPatholPrice
    
    objPatholPrice.zlInitModule -1, mstrPrivs, mlngCur����ID, Me
    objPatholPrice.zlRefresh mlngCur����ID, mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.blnMoved
    
    objPatholPrice.Show 1, Me
End Sub

Private Sub Menu_Manage_������()
'���ģʽ�µĲ����Ѵ���
On Error GoTo errH
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngSystemFrom As Long
    Dim strPar As String
    
    strSQL = "select B.���ӱ�־ From ����ҽ����¼ A, ���˹Һż�¼ B Where A.�Һŵ�=B.No And A.ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ӱ�־", mobjCurStudyInfo.lngAdviceId)
    
    If rsData.RecordCount <= 0 Then
        '�����ϰ油�Ѵ���
        lngSystemFrom = 1
    Else
        If Val(nvl(rsData!���ӱ�־)) = 3 Then
            '�����°油��
            lngSystemFrom = 2
        Else
            '�����ϰ油�Ѵ���
            lngSystemFrom = 1
        End If
    End If
    
    strPar = GetJsonPar(mobjCurStudyInfo.lngAdviceId)
    
    Call mobjAppendBill.EditChargeBill(strPar)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetJsonPar(ByVal lngAdviceId As Long) As String
On Error GoTo errH
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strUserName As String
    Dim strUserPswd As String
    Dim lngUerResId As Long
    Dim strNodeNo As String
    Dim strNodeName As String
    Dim strSysFrom As String
    Dim strUerResId As String
    
    GetJsonPar = ""
    
    If gobjRegister Is Nothing Then Set gobjRegister = VBA.Interaction.GetObject("", "zlRegister.clsRegister")
    If gobjRegister Is Nothing Then
        Set gobjRegister = CreateObject("zlRegister.clsRegister")
    End If
    
    lngUerResId = UserInfo.ID
    strNodeName = ""
    strNodeNo = ""
    
    '��ѯ������Դϵͳ
    strSysFrom = "01"
    strSQL = "Select ���ӱ�־ From ���˹Һż�¼ Where ����ID=[1] and No=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ӱ�־", mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.strRegNo)
    If rsData.RecordCount > 0 Then
        If Val(nvl(rsData!���ӱ�־)) = 3 Then strSysFrom = "02"
    End If
    
            
    strUserName = gobjRegister.GetUserName
    strUserPswd = gstrInputPwd ' GetLoginPassword 'gobjRegister.GetPassword(App.hInstance)
    
    If strSysFrom = "02" Then
        strSQL = "Select ��ԴID From ��Ա�� Where ID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��Ա����ԴID", UserInfo.ID)
        If rsData.RecordCount > 0 Then
            strUerResId = nvl(rsData!��ԴID)
        End If
    
        strSQL = "Select a.����ID," & _
                    " '' As �����ʶ, " & _
                    " Decode(a.������Դ, 4, 2, 2, 1, 0) As ������Դ, " & _
                    " a.ID As ҽ�����, b.���ͺ�, " & _
                    " c.��Դid As ��ǰ���ұ�ʶ, " & _
                    " c.���� As ��ǰ���ұ���, c.���� As ��ǰ��������" & _
                    " From ����ҽ����¼ A, ����ҽ������ B, ���ű� C " & _
                    " Where a.Id = b.ҽ��id And b.ִ�в���id = c.Id And a.Id = [1]"

    Else
        strNodeNo = gstrNodeNo
        strNodeName = gstrNodeName
    
        strSQL = "Select a.����ID," & _
                    " To_Char(a.��ҳid) As �����ʶ, " & _
                    " Decode(a.������Դ, 4, 2, 2, 1, 0) As ������Դ, " & _
                    " b.ҽ��id As ҽ�����, b.���ͺ�, " & _
                    " To_Char(b.ִ�в���id) As ��ǰ���ұ�ʶ, " & _
                    " c.���� As ��ǰ���ұ���, c.���� As ��ǰ��������" & _
                    " From ����ҽ����¼ A, ����ҽ������ B, ���ű� C " & _
                    " Where a.Id = b.ҽ��id And b.ִ�в���id = c.Id And a.Id = [1]"
                
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯҽ��Json����", lngAdviceId)
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetJsonPar = "{" & _
            """��Դϵͳ"":""" & strSysFrom & """," & _
            """������Դ"":""" & nvl(rsData!������Դ) & """," & _
            """���˱�ʶ"":""" & nvl(rsData!����ID) & """," & _
            IIf(strSysFrom <> "02", """�����ʶ"":""" & nvl(rsData!�����ʶ) & """,", "") & _
            """ҽ�����"":""" & nvl(rsData!ҽ�����) & """," & _
            """ҽ�����ͺ�"":""" & nvl(rsData!���ͺ�) & """," & _
            """��ǰ���ұ�ʶ"":""" & nvl(rsData!��ǰ���ұ�ʶ) & """," & _
            """��ǰ���ұ���"":""" & nvl(rsData!��ǰ���ұ���) & """," & _
            """��ǰ��������"":""" & nvl(rsData!��ǰ��������) & """," & _
            """����Ա��ʶ"":""" & IIf(strSysFrom <> "02", lngUerResId, strUerResId) & """," & _
            """����Ա����"":""" & UserInfo.��� & """," & _
            """����Ա����"":""" & UserInfo.���� & """," & _
            """Ժ������"":""" & strNodeNo & """," & _
            """Ժ������"":""" & strNodeName & """," & _
            """�û���"":""" & strUserName & """," & _
            """�û�����"":""" & strUserPswd & """" & _
        "}"
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetLoginPassword()
    '��ȡ��ǰ�û���¼����
    Dim objLogin As Object
   
    On Error Resume Next
    
    GetLoginPassword = ""
    
    Set objLogin = CreateObject("zlLogin.clsLogin")
    If objLogin Is Nothing Then
        err.Clear
        Exit Function
    End If
   
    GetLoginPassword = objLogin.InputPwd
End Function


Private Function Menu_Manage_�ޱ������() As Boolean
'ֻ�н����еı�����Բ����ò˵�,��Ϊ��ʱ��û��ǩ��
On Error GoTo errhandle
    Dim lngID As Long
    Dim intDelete As Integer '�������Ƿ�ɾ������  1ɾ��   0��ɾ��
    
    intDelete = 0
    Menu_Manage_�ޱ������ = False
    
    If (mobjCurStudyInfo.strReportDoctor <> "" Or mobjCurStudyInfo.strReportOperation <> "") Then
        If MsgBoxD(Me, "��ɼ��ǰ�Ƿ�ɾ���Ѿ���д�ı���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            intDelete = 0
        Else
            intDelete = 1
        End If
    End If

    If mblnAllDepts Then
        If mobjCurStudyInfo.lngExeDepartmentId > 0 Then
            lngID = mobjCurStudyInfo.lngExeDepartmentId
        Else
            lngID = 0
        End If
    Else
        lngID = mlngCur����ID
    End If
    
    If mrtReportType = �����ĵ��༭�� Then
        gstrSQL = "Zl_Ӱ����_״̬����(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & ",'',6," & intDelete & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lngID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ")"
    ElseIf mrtReportType = PACS����༭�� Then
        gstrSQL = "Zl_Ӱ����_State(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & ",6," & intDelete & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lngID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ",1)"
    Else
        gstrSQL = "Zl_Ӱ����_State(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & ",6," & intDelete & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lngID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ",2)"
    End If

    Call zlDatabase.ExecuteProcedure(gstrSQL, "�ı������")
    Menu_Manage_�ޱ������ = True
    
    Exit Function
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function getRegID(ByVal strRegNo As String) As Long
'����:��ȡ�Һ�id
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errhandle
    
    getRegID = 0
    
    strSQL = "select id from ���˹Һż�¼ where no=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, GetWindowCaption, strRegNo)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    getRegID = nvl(rsTemp!ID, 0)
    
    Exit Function

errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function IsAlreadyInputQuality(ByVal lngAdviceId As Long) As Boolean
On Error GoTo errH
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    IsAlreadyInputQuality = False
    
    strSQL = "select �ۺ����� from ��������Ϣ where ҽ��ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, GetWindowCaption, lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    If nvl(rsData!�ۺ�����) <> "" Then IsAlreadyInputQuality = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Menu_Manage_����������(Optional lngAdviceId As Long = 0, Optional blnRefresh As Boolean = True, Optional strReportId As String = "")
'�������������̵��ã���ʱ������ҽ��ID������ҪȨ���ж�
On Error GoTo errhandle
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim intState As Integer
    Dim blnAllReportFinished As Boolean
    Dim objStudyInfo As New clsStudyInfo
    Dim intCol As Integer
    Dim lngRow As Long
    Dim lngAdviceIDSub As Long '�������е�ҽ��ID
    Dim lngID As Long
    
        '���ִ�й���=6 ˵���������Ѿ��������״̬����ʱ�˳������̲��Ҳ���Ҫ��ʾ��������XX���Զ���ɲ�����
    If lngAdviceId > 0 Then
        strSQL = "select ҽ��ID from ����ҽ������ where ҽ��ID=[1] and ִ�й���=6"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�Ƿ��Ѿ��������״̬", lngAdviceId)
        If rsData.RecordCount > 0 Then
             Exit Sub
        End If
    End If
    
    If InStr(mstrPrivs, ";������;") <= 0 Then
        MsgBoxD Me, "û��Ȩ�ޣ������������", vbInformation, Me.Caption
        Exit Sub
    End If

    '��δ����ҽ��ID,��ȡѡ����ҽ��ID
    lngAdviceIDSub = lngAdviceId
    If lngAdviceIDSub = 0 Then
        If vsfList.Rows > 1 Then
            intCol = vsfList.ColIndex("ҽ��ID")
            lngRow = vsfList.Row
            lngAdviceIDSub = Val(vsfList.TextMatrix(lngRow, intCol))
            
        End If
    End If
    
    If lngAdviceIDSub = 0 Then
        MsgBoxD Me, "��ȡ�������ʧ��", vbInformation, Me.Caption
        Exit Sub
    End If
            
    Set objStudyInfo = mobjPacsQueryWrap.GetBaseInfo(lngAdviceIDSub, GetMovedState(lngRow, vsfList))
    
    If objStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Not mSysPar.blnNoSignFinish Then
    '�����ѡ����δǩ������򲻱ؽ���������ж�
        If Is_ExistReportWriting(lngAdviceIDSub) Then
            MsgBoxD Me, "�����Ѿ��޸Ļ�δǩ��������������ɡ�", vbInformation, Me.Caption
            Exit Sub
        ElseIf objStudyInfo.intStep < 4 Then
            MsgBoxD Me, "���滹δǩ��������������ɡ�", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    
    '������֮ǰ�����ж��Ƿ�����������������������ɣ�
        '1��סԺ���ߣ��Ѿ���Ժ������δ��˵Ļ��۵���ʹ�á�ִ�к��Զ���˻��۵�������
        '2�����ﻼ�ߣ���δ���ѵĵ��ݡ�
    If objStudyInfo.lngPatientFrom = 2 Then
        'סԺ���ߣ��ж��Ƿ��Ѿ���Ժ������δ��˻��۵�
        If bln����δ��˳�Ժ(objStudyInfo.lngPatId, objStudyInfo.lngPageID, nvl(objStudyInfo.lngAdviceId), objStudyInfo.lngPatientFrom) Then
            'ִ�к��Զ���˻��۵���Ч�����Ҳ����ѳ�Ժ������δ��˵Ļ��۵�
            MsgBoxD Me, "�ò����ѳ�Ժ������δ��˵Ļ��۵���������ɣ�", vbExclamation, gstrSysName
            Exit Sub
        End If
    ElseIf objStudyInfo.lngPatientFrom = 4 And mSysPar.blnPEISNoCheckMoneyFinish Then
        '�����ɲ��жϷ��� 133458
    Else
        '������ﻼ��,�ж��Ƿ���δ�ɷ���
        If blnδ�ɷ���(objStudyInfo.lngAdviceId) = True Then
            If objStudyInfo.intGreenChannel = 1 Or objStudyInfo.intEmergentTag = 1 Then
                If MsgBoxD(Me, "�û��߻���δ�ɷѵ���Ŀ���Ƿ�Ҫ��ɣ�", vbYesNo, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            Else
                MsgBoxD Me, "�û��߻���δ�ɷѵ���Ŀ��������ɡ�", vbExclamation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    If lngAdviceId = 0 Then
    '����ǻ��б���δ���ʱ��ɼ��
        If mrtReportType = �����ĵ��༭�� Then
            intState = getStudyStateRich(objStudyInfo.lngAdviceId, strReportId, False, blnAllReportFinished)
        
            If intState = 4 And blnAllReportFinished = False Then
                If MsgBoxD(Me, "���б���û��д�꣬�����ʱ��ɼ�飬��Ҫ�С���¼���桱Ȩ�޵��˲��ܼ�����д����!" & vbCrLf & "ȷ��Ҫ���������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If
    
    '��մ�������
    Call Menu_Manage_SendAudit("")

    '����ǲ���ϵͳ��������ʱ������Ҫ�����������ƴ���
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        
        If (mSysPar.blnPopChangGuiWindow And objStudyInfo.intPathoType = 0) _
            Or (mSysPar.blnPopKuaiShuWindow And objStudyInfo.intPathoType = 5) _
            Or (mSysPar.blnPopBingDongWindow And objStudyInfo.intPathoType = 1) _
            Or (mSysPar.blnPopXiBaoWindow And objStudyInfo.intPathoType = 2) _
            Or (mSysPar.blnPopHuiZhenWindow And objStudyInfo.intPathoType = 3) _
            Or (mSysPar.blnPopShiJianWindow And objStudyInfo.intPathoType = 4) Then

            If Not IsAlreadyInputQuality(objStudyInfo.lngAdviceId) Then
                If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.zlMenu.zlExecuteMenu(conMenu_Pathol_Quality_Manage)
            End If
    
            If Not IsAlreadyInputQuality(objStudyInfo.lngAdviceId) Then
                Call MsgBoxD(Me, "δ¼��������������ִ����ɲ�����", vbInformation, GetWindowCaption)
                Exit Sub
            End If
        End If
            
    End If
    
    If mblnAllDepts Then
        If objStudyInfo.lngExeDepartmentId > 0 Then
            lngID = objStudyInfo.lngExeDepartmentId
        Else
            lngID = 0
        End If
    Else
        lngID = mlngCur����ID
    End If
    
    If mrtReportType = �����ĵ��༭�� Then
        strSQL = "Zl_Ӱ����_״̬����(" & objStudyInfo.lngAdviceId & "," & objStudyInfo.lngSendNo & ",'',6,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lngID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ")"
    ElseIf mrtReportType = PACS����༭�� Then
        strSQL = "Zl_Ӱ����_State(" & objStudyInfo.lngAdviceId & "," & objStudyInfo.lngSendNo & ",6,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lngID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ",1)"
    Else
        strSQL = "Zl_Ӱ����_State(" & objStudyInfo.lngAdviceId & "," & objStudyInfo.lngSendNo & ",6,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lngID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ",2)"
    End If
    
    Call zlDatabase.ExecuteProcedure(strSQL, "�ı������")

        
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        gstrSQL = "Zl_������_���(" & objStudyInfo.lngAdviceId & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���������")
    End If
    
    Call CheckExecuteInterface(EInterfaceExeTime.�����ɺ�)
        
    'ȡ���Ŷ���Ϣ
    If mSysPar.blnUseQueue = True And Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Call mobjQueue.zlCompletePacsQueue(objStudyInfo.lngAdviceId)
    End If

    Call UpdateQueryListData(Nothing, objStudyInfo.lngAdviceId)
        
    Call NotificationAllModuleRefresh
    Call RefreshTabWindow(True)
    
    '���ͼ�������Ϣ
    Call mobjMsgCenter.Send_Msg_StudyComplete(objStudyInfo.lngAdviceId, strReportId)
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_ȡ��������()
On Error GoTo errhandle
    Dim strSQL As String
    Dim intState As Integer

    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If

    If mobjCurStudyInfo.blnMoved Then
        MsgBoxD Me, "�ò��˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        If CheckIsArchived(mobjCurStudyInfo.lngAdviceId) Then
            MsgBoxD Me, "�ò��˵ĵ����Ѿ��鵵�������������", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If mrtReportType = �����ĵ��༭�� Then
        intState = getStudyStateRich(mobjCurStudyInfo.lngAdviceId, "", True)
        strSQL = "Zl_Ӱ����_״̬����(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & ",''," & intState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ")"
    ElseIf mrtReportType = PACS����༭�� Then
        intState = getStudyState(mobjCurStudyInfo.lngAdviceId, True)
        strSQL = "Zl_Ӱ����_State(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & "," & intState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ",1)"
    Else
        intState = getStudyState(mobjCurStudyInfo.lngAdviceId, True)
        strSQL = "Zl_Ӱ����_State(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & "," & intState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ",2)"
    End If
    
    zlDatabase.ExecuteProcedure strSQL, "ȡ��������"
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        strSQL = "Zl_������_ȡ�����(" & mobjCurStudyInfo.lngAdviceId & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "������ȡ�����")
    End If
    
    Call CheckExecuteInterface(EInterfaceExeTime.ȡ�����ʱ)
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    Call NotificationAllModuleRefresh
    Call RefreshTabWindow(True)
    
    '���ͼ�鳷�������Ϣ
    Call mobjMsgCenter.Send_Msg_CancelComplete(mobjCurStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function CheckIsArchived(lngAdviceId As Long) As Boolean
 '���ò��˵����Ƿ��Ѿ��鵵���ѹ鵵�ļ�飬��Ҫ��������ȡ�����  0--δ�鵵  1--�ѹ鵵
 On Error GoTo errhandle
 
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "select distinct c.����״̬ as ״̬ from ��������Ϣ a,����鵵��Ϣ b,��������Ϣ c where a.����ҽ��ID = b.����ҽ��ID and b.����id = c.id and a.ҽ��ID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ��ѹ鵵", lngAdviceId)
    
    If rsTemp.RecordCount < 1 Then
        CheckIsArchived = False
        Exit Function
    End If
    
    CheckIsArchived = IIf(nvl(rsTemp!״̬, 0) = 1, True, False)
Exit Function
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Menu_Manage_CriticalMark(ByVal lngID As Long)
'Σ��ֵ����
On Error GoTo errhandle
    Dim strSQL As String
    Dim intCritical As Integer
    Dim rsData As ADODB.Recordset
    Dim lngCriticalId As Long
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mobjPublicAdvice Is Nothing Then
        Set mobjPublicAdvice = DynamicCreate("zlPublicAdvice.clsPublicAdvice", "zlPublicAdvice")
        If mobjPublicAdvice Is Nothing Then Exit Sub
        
        Call mobjPublicAdvice.InitCommon(gcnOracle, glngSys, gstrNodeNo, gfrmMain, glngModul, gstrPrivs, mobjMsgCenter.Msg)
        Call mobjPublicAdvice.InitDisease(gcnOracle, glngSys, gfrmMain, glngModul, gstrPrivs)

    End If

    Select Case lngID
        Case conMenu_Manage_PacsCriticalReg     'Σ�����ߵǼ�
            If mobjCurStudyInfo.lngPatientFrom = 1 Then        '����
                Call mobjPublicAdvice.ShowAppCritical(Me, True, 0, 1, _
                            mobjCurStudyInfo.lngPatId, 0, mobjCurStudyInfo.strRegNo, mobjCurStudyInfo.lngBaby, lngCriticalId, _
                            mobjCurStudyInfo.lngAdviceId, , , , mlngCur����ID, gstrUserName, mobjMsgCenter.Msg)
            ElseIf mobjCurStudyInfo.lngPatientFrom = 2 Then    'סԺ
                Call mobjPublicAdvice.ShowAppCritical(Me, True, 0, 2, _
                            mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.lngPageID, mobjCurStudyInfo.strRegNo, mobjCurStudyInfo.lngBaby, lngCriticalId, _
                            mobjCurStudyInfo.lngAdviceId, , , , mlngCur����ID, gstrUserName, mobjMsgCenter.Msg)
            Else                                            '���������
                Call mobjPublicAdvice.ShowAppCritical(Me, True, 0, 3, _
                            mobjCurStudyInfo.lngPatId, 0, "", mobjCurStudyInfo.lngBaby, lngCriticalId, _
                            mobjCurStudyInfo.lngAdviceId, , , , mlngCur����ID, gstrUserName, mobjMsgCenter.Msg)
            End If
    
        Case conMenu_Manage_PacsCriticalManage  'Σ�����߹���
            If mobjPublicAdvice.ShowQueryCritical(Me, True, 2, 1, mlngCur����ID, 0, mobjMsgCenter.Msg) = False Then Exit Sub
    End Select

    '��ѯҽ��Σ�����...
    strSQL = "Select ID From ����Σ��ֵ��¼ Where ҽ��ID=[1] and nvl(״̬, 0)<>0"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯΣ��ҽ��״̬", mobjCurStudyInfo.lngAdviceId)
    If rsData.RecordCount > 0 Then
        intCritical = 1         'Σ��
    Else
        intCritical = 0         '��Σ��
    End If
    
    '����Ӱ��Σ��״̬
    If intCritical = 1 Then
        strSQL = "zl_Ӱ����_Σ������(" & mobjCurStudyInfo.lngAdviceId & ",1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)

        mobjCurStudyInfo.intDangerState = 1

        Menu_Manage_������� conMenu_Manage_Negative
        
        '����Σ��ֵ��Ϣ
        'Call mobjMsgCenter.Send_Msg_Critical(mobjCurStudyInfo.lngAdviceId)
    ElseIf intCritical = 0 Then
        strSQL = "Zl_Ӱ��Σ��ֵ��¼_ȡ��(" & mobjCurStudyInfo.lngAdviceId & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)

        mobjCurStudyInfo.intDangerState = 0
    End If
        
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)

    Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_�������(ByVal lngID As Long)
On Error GoTo errhandle
    Dim strSQL As String
    Dim iResult As Integer
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    Select Case lngID
        Case conMenu_Manage_Negative
            iResult = 1
        Case conMenu_Manage_Positive
            iResult = 0
    End Select
    
    If mrtReportType = �����ĵ��༭�� Then
        Call mobjWork_Report.Menu_Manage_�������(mobjCurStudyInfo.lngAdviceId, iResult)
        Exit Sub
    End If
    
    strSQL = "ZL_Ӱ����_���(" & mobjCurStudyInfo.lngAdviceId & "," & iResult & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "���������")

    mobjCurStudyInfo.intPositive = iResult
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_��ɫͨ��(ByVal lngID As Long)
On Error GoTo errhandle
    Dim strSQL As String
    Dim intResult As Integer
    Dim blnCanPrint As Boolean
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    Select Case lngID
        Case conMenu_Manage_GChannelOk
            intResult = "1"
        Case conMenu_Manage_GChannelCancel
            intResult = "0"
    End Select
    
    strSQL = "Zl_��ɫͨ��_Update(" & mobjCurStudyInfo.lngAdviceId & ",'" & intResult & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "��ɫͨ��")
    
    mobjCurStudyInfo.intGreenChannel = intResult
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    If mrtReportType = �����ĵ��༭�� Then

        blnCanPrint = mobjCurStudyInfo.intEmergentTag <> 0 Or mobjCurStudyInfo.intGreenChannel <> 0
        
        If Not mobjWork_Report Is Nothing Then
            Call mobjWork_Report.zlUpdateOtherInf(picReportContainer, vsfList, mblnIsHistory, blnCanPrint, mobjCurStudyInfo.strDoDoctor, mobjCurStudyInfo.strStudyUID)
            Call mobjWork_Report.zlRefreshFace(True, False, False)
        End If
    End If

Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_�������(ByVal lngID As Long)
On Error GoTo errhandle
    Dim strResult As String
    Dim strSQL As String
    Dim lngColIndex As Long

    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If

    Select Case lngID
        Case conMenu_Manage_FuHe
            strResult = "����"
        Case conMenu_Manage_JiBenFuHe
            strResult = "��������"
        Case conMenu_Manage_BuFuHe
            strResult = "������"
    End Select

    strSQL = "Zl_�������_Update(" & mobjCurStudyInfo.lngAdviceId & ",'" & strResult & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "�������")
        
    mobjCurStudyInfo.strAccord = strResult
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)

Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_CheckList()
    If mobjCurStudyInfo.lngAdviceId > 0 Then
        Set mclsCISKernel = New clsCISKernel
        Call mclsCISKernel.ShowPacsApplication(Me, mobjCurStudyInfo.lngAdviceId)
        Set mclsCISKernel = Nothing
    Else
        MsgBox "û��ѡ���ˡ�", vbInformation + vbOKOnly, gstrSysName
    End If
End Sub

'�ֲ�λִ��
Private Sub menu_Manage_ExecOnePart()
    Dim frmExecForm As frmExecOnePart
    
    Set frmExecForm = New frmExecOnePart
    
    '��ʾ�ֲ�λִ�к�ȡ������
    Call frmExecForm.ZlShowMe(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.strPatientName, mobjCurStudyInfo.strPatientAge, mobjCurStudyInfo.strPatientSex, mobjCurStudyInfo.strStuStateDesc, Me)
    
    'ˢ�·���ҳ��
    If TabWindow.Selected.tag = "�������" Or TabWindow.Selected.tag = "סԺҽ��" Or TabWindow.Selected.tag = "����ҽ��" Then
        Call RefreshTabWindow
    End If
End Sub

'��Ⱦ���Ǽ�
Private Sub Menu_Manage_DiseaseRegist()
    Dim strReportResult As String
    Dim strCurrDocId As String
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    
On Error GoTo errhandle
    If mobjPublicAdvice Is Nothing Then
        Set mobjPublicAdvice = DynamicCreate("zlPublicAdvice.clsPublicAdvice", "zlPublicAdvice")
        If mobjPublicAdvice Is Nothing Then Exit Sub
        
        Call mobjPublicAdvice.InitCommon(gcnOracle, glngSys, gstrNodeNo, gfrmMain, glngModul, gstrPrivs, mobjMsgCenter.Msg)
        Call mobjPublicAdvice.InitDisease(gcnOracle, glngSys, gfrmMain, glngModul, gstrPrivs)
    End If
    
    If mrtReportType = �����ĵ��༭�� Then
        strCurrDocId = mobjWork_Report.GetCurrDocId(mobjCurStudyInfo.lngAdviceId)
        
        If Trim(strCurrDocId) <> "" Then
            strSQL = "Select ������ From Ӱ�񱨸��¼ Where ID = [1]"
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������", strCurrDocId)
            
            If rsData.RecordCount > 0 Then strReportResult = nvl(rsData!������)
        End If
    Else
        strSQL = "Select  b.�����ı� As ���� From ���Ӳ������� a,���Ӳ������� b, ����ҽ������ c " & _
                 "Where c.ҽ��id = [1] And a.�����ı� = '������' And a.�������� = 3 And a.Id = b.��ID " & _
                 "And a.�ļ�id = c.����id And b.�������� = 2 And b.��ֹ�� = 0"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������", mobjCurStudyInfo.lngAdviceId)
        
        If rsData.RecordCount > 0 Then strReportResult = nvl(rsData!����)
    End If
    
    If mobjCurStudyInfo.lngPatientFrom = 1 Then        '����
        Call mobjPublicAdvice.ShowDisRegist(Me, 0, , mobjCurStudyInfo.lngPatId, , mobjCurStudyInfo.strRegNo, mobjCurStudyInfo.lngAdviceId, mlngCur����ID, , , , , strReportResult)
    ElseIf mobjCurStudyInfo.lngPatientFrom = 2 Then    'סԺ
        Call mobjPublicAdvice.ShowDisRegist(Me, 0, , mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.lngPageID, , mobjCurStudyInfo.lngAdviceId, mlngCur����ID, , , , , strReportResult)
    Else                                            '���������
        Call mobjPublicAdvice.ShowDisRegist(Me, 0, , mobjCurStudyInfo.lngPatId, , , mobjCurStudyInfo.lngAdviceId, mlngCur����ID, , , , , strReportResult)
    End If
    
    Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'��Ⱦ����ѯ
Private Sub Menu_Manage_DiseaseQuery()
On Error GoTo errhandle
    If mobjPublicAdvice Is Nothing Then
        Set mobjPublicAdvice = DynamicCreate("zlPublicAdvice.clsPublicAdvice", "zlPublicAdvice")
        If mobjPublicAdvice Is Nothing Then Exit Sub
        Call mobjPublicAdvice.InitDisease(gcnOracle, glngSys, gfrmMain, glngModul, gstrPrivs)
    End If
    
    Call mobjPublicAdvice.ShowDisQuery(mlngCur����ID)

    Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Manage_�޸�()
On Error GoTo errhandle
    Dim strOldName As String
    Dim strOldRoom As String
    Dim strQueueName As String
    Dim strCodeNo As String
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        With frmRISRequest
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = mobjCurStudyInfo.lngSendNo
            .mlngAdviceId = mobjCurStudyInfo.lngAdviceId
            .mstrPatientName = mobjCurStudyInfo.strPatientName
            .mintEditMode = IIf(mobjCurStudyInfo.intStep > 1, 3, 1)  '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
            .mlngCurDeptId = IIf(mblnAllDepts, mobjCurStudyInfo.lngExeDepartmentId, mlngCur����ID)
            .mstrCur���� = zlStr.NeedName(mstrCur����)
            
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            
            If .mlngResultState <> 0 Then
                strOldName = mobjCurStudyInfo.strPatientName
                strOldRoom = mobjCurStudyInfo.strExeRoom
                
                Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
                
                If mSysPar.blnUseQueue And Not mobjQueue Is Nothing Then
                    '����Ǳ������޸ģ��Ҹı���ִ�м䣬����Ҫ���½����Ŷ�
                    If .mintEditMode = 3 And .mlngResultState = 3 Then
                        If .mstrTechnicRoom <> strOldRoom Then
                            If .mstrTechnicRoom = "" Then
                                '���Ϊ�գ�����Ҫ����ü����Ŀ��Ӧ����Ŀ������߿��ҵĶ�����
                                Call mobjQueue.zlGetInQueueInf(mobjCurStudyInfo.lngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                            Else
                                '�����Ϊ�գ���д���Ӧ��ִ�м�����
                                strQueueName = .mstrCur���� & "-" & .mstrTechnicRoom
                                strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                            End If
                            
                            Call mobjQueue.zlUpdatePacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
                        Else
                            '������ʽ���޸ģ���ֻ���Ŷӽк��е������Ϣ���и���
                            If .mstrPatientName <> strOldName Then
                                Call mobjQueue.zlUpdatePacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId)
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Else
        With frmPatholRIS
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = mobjCurStudyInfo.lngSendNo
            .mlngAdviceId = mobjCurStudyInfo.lngAdviceId
            .mintEditMode = IIf(mobjCurStudyInfo.intStep > 1, 3, 1)  '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
            .mlngCurDeptId = IIf(mblnAllDepts, mobjCurStudyInfo.lngExeDepartmentId, mlngCur����ID)
            .mintImgCount = mintImgCount
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            
            If .RefreshPatiInfor(False) = True Then  'ˢ�²���
                .mblnOk = False
                .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            End If
            
            If .mblnOk Then Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId) '�ɹ�����
        End With
    End If
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Manage_ModifBaseInfo()
'������Ϣ����
On Error GoTo errhandle
    Dim zlPubPatient As Object
    
    Dim int���� As Integer
    Dim str����ID As String

    Set zlPubPatient = VBA.Interaction.GetObject("", "zlPublicPatient.clsPublicPatient")
    If zlPubPatient Is Nothing Then Set zlPubPatient = CreateObject("zlPublicPatient.clsPublicPatient")
    
    If Not zlPubPatient Is Nothing Then Call zlPubPatient.zlInitCommon(gcnOracle, glngSys)
    
    With mobjCurStudyInfo
        int���� = Decode(.lngPatientFrom, 1, 1, 2, 2, 3, 3, 4, 4)

        str����ID = Decode(.lngPatientFrom, 1, getRegID(.strRegNo), 2, .lngPageID, 3, .lngAdviceId, 4, .strRegNo)

        If zlPubPatient.ModiPatiBaseInfo(Me, mlngModule, .lngPatId, str����ID, int����) Then
            Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
        End If
        
    End With
    
    Set zlPubPatient = Nothing
Exit Sub
errhandle:
    Set zlPubPatient = Nothing
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Manage_���ƵǼ�()
    Dim strQueueName As String
    Dim strCodeNo As String
    
On Error GoTo errhandle
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        With frmRISRequest
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = 0
            .mlngAdviceId = 0
            .mintEditMode = 0 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
            .mlngCurDeptId = IIf(mblnAllDepts, mobjCurStudyInfo.lngExeDepartmentId, mlngCur����ID)
            .mstrCur���� = zlStr.NeedName(mstrCur����)
            .mlngResultState = 0
            
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1), mblnAllDepts, mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo
            
            If .mlngResultState <> 0 Then '�ɹ�����
                Call CheckExecuteInterface(EInterfaceExeTime.�ǼǺ�)
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
                
                If .mintEditMode = 2 Then
                    Call CheckExecuteInterface(EInterfaceExeTime.������)
                    Call mobjPacsQueryWrap.LocateRow(1)
                End If
                
                '���ͬʱ��ѡ����ʼ����Զ��򿪱��桱�͡��ǼǺ��Զ�������������ô���Զ��򿪱������
                If mSysPar.blnAutoOpenReport And mSysPar.blnֱ�Ӽ�� Then Call Menu_RichEPR(conMenu_Edit_Modify)
                
                If .mlngResultState = 2 Then
                    '��������Ŷӽкţ��򱨵�����Ҫ�����ŶӽкŶ���......
                    If mSysPar.blnUseQueue And Not mobjQueue Is Nothing Then
                        '������Ҫ����Ķ�������
                        If .mstrTechnicRoom = "" Then
                            '���δ�գ�����Ҫ����ü����Ŀ��Ӧ����Ŀ������߿��ҵĶ�����
                            Call mobjQueue.zlGetInQueueInf(mobjCurStudyInfo.lngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                        Else
                            '�����Ϊ�գ���д���Ӧ��ִ�м�����
                            strQueueName = .mstrCur���� & "-" & .mstrTechnicRoom
                            strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                        End If
                        
                        Call mobjQueue.zlInPacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
                    End If
                    
                End If
                
                '������������Ϣ
                Call mobjMsgCenter.Send_Msg_Request(.mlngAdviceId)
            End If
        End With
    Else
        With frmPatholRIS
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = 0
            .mlngAdviceId = 0
            .mintEditMode = 0 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
            .mlngCurDeptId = IIf(mblnAllDepts, mobjCurStudyInfo.lngExeDepartmentId, mlngCur����ID)
            .mblnOk = False
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            
            If .CopyCheck(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo) = True Then  'ˢ�²���
                .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            End If
            
            If .mblnOk Then '�ɹ�����
                Call CheckExecuteInterface(EInterfaceExeTime.�ǼǺ�)
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
            End If
        End With
    End If
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Manage_�Ǽ�()
On Error GoTo errhandle
    Dim strQueueName As String
    Dim strCodeNo As String
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Set mfrmRISRequest = New frmRISRequest
        With mfrmRISRequest
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = 0
            .mlngAdviceId = 0
            .mstrPatientName = ""
            .mintEditMode = 0 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
            .mlngCurDeptId = IIf(mblnAllDepts, mobjCurStudyInfo.lngExeDepartmentId, mlngCur����ID)
            .mstrCur���� = zlStr.NeedName(mstrCur����)
            .mlngResultState = 0
            
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1), mblnAllDepts
            
            If .mlngResultState <> 0 Then '�ɹ�����
                Call CheckExecuteInterface(EInterfaceExeTime.�ǼǺ�)
                
                If mSysPar.bln�����Ǽ� Then
                    Call RefreshList
                Else
                    Call UpdateQueryListData(Nothing, .mlngAdviceId)
                End If
                
                If .mintEditMode = 2 Then
                    Call CheckExecuteInterface(EInterfaceExeTime.������)
                    Call mobjPacsQueryWrap.LocateRow(1)
                End If
                
                '���ͬʱ��ѡ����ʼ����Զ��򿪱��桱�͡��ǼǺ��Զ�������������ô���Զ��򿪱������
                If mSysPar.blnAutoOpenReport And mSysPar.blnֱ�Ӽ�� Then Call Menu_RichEPR(conMenu_Edit_Modify)
                If .mlngResultState = 2 Then
                    Call CheckExecuteInterface(EInterfaceExeTime.������)
                    '��������Ŷӽкţ��򱨵�����Ҫ�����ŶӽкŶ���......
                    If mSysPar.blnUseQueue And Not mobjQueue Is Nothing Then
                        '������Ҫ����Ķ�������
                        If .mstrTechnicRoom = "" Then
                            '���δ�գ�����Ҫ����ü����Ŀ��Ӧ����Ŀ������߿��ҵĶ�����
                            Call mobjQueue.zlGetInQueueInf(mobjCurStudyInfo.lngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                        Else
                            '�����Ϊ�գ���д���Ӧ��ִ�м�����
                            strQueueName = .mstrCur���� & "-" & .mstrTechnicRoom
                            strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                        End If
                        Call mobjQueue.zlInPacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
                    End If
                    
                End If
                
                '������������Ϣ
                Call mobjMsgCenter.Send_Msg_Request(.mlngAdviceId)
            End If
            
        End With
    Else
        With frmPatholRIS
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = 0
            .mlngAdviceId = 0
            .mintEditMode = 0 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
            .mlngCurDeptId = IIf(mblnAllDepts, mobjCurStudyInfo.lngExeDepartmentId, mlngCur����ID)
            .mintImgCount = 0
            .mblnOk = False
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            
            If .mblnOk Then '�ɹ�����
                Call CheckExecuteInterface(EInterfaceExeTime.�ǼǺ�)
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
     
                If .mintEditMode = 2 Then
                    Call CheckExecuteInterface(EInterfaceExeTime.������)
                    Call mobjPacsQueryWrap.LocateRow(1)
                End If
                
                If mSysPar.blnֱ�Ӽ�� Then Call CheckExecuteInterface(EInterfaceExeTime.������)
                '���ͬʱ��ѡ����ʼ����Զ��򿪱��桱�͡��ǼǺ��Զ�������������ô���Զ��򿪱������
                If mSysPar.blnAutoOpenReport And mSysPar.blnֱ�Ӽ�� Then Call Menu_RichEPR(conMenu_Edit_Modify)
                
                '������������Ϣ
                Call mobjMsgCenter.Send_Msg_Request(.mlngAdviceId)
            End If
        End With
    End If
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Manage_ȡ���Ǽ�()
On Error GoTo errhandle
    Dim strSQL As String
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If MsgBoxD(Me, "ȷ��Ҫȡ����ǰ������" & Chr(10) & Chr(13) & "����ȡ�������Ӧ��ҽ�����ܾ�ִ�У�", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strSQL = "ZL_����ҽ��ִ��_�ܾ�ִ��(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & ",null,null," & mlngCur����ID & ")"
    
    Call zlDatabase.ExecuteProcedure(strSQL, "�����Ǽ�")
    Call CheckExecuteInterface(EInterfaceExeTime.ȡ���Ǽ�ʱ)
    
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    '����ҽ��������Ϣ
    Call mobjMsgCenter.Send_Msg_CancelAdvice(mobjCurStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_�ٻ�ȡ��()
'���ܣ��ٻر�ȡ���ĵǼ�
On Error GoTo errhandle
    Dim strSQL As String
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If MsgBoxD(Me, "ȷʵҪ�ٻر�ȡ���Ǽǵ���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strSQL = "ZL_����ҽ��ִ��_ȡ���ܾ�(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & ",null,null," & mlngCur����ID & ")"
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    '����״̬ͬ����Ϣ
    Call mobjMsgCenter.Send_Msg_StateSync(mobjCurStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_����()
On Error GoTo errhandle
    Dim blnFocusFind As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim strQueueName As String
    Dim strCodeNo As String
    Dim blnIsCurDayReservations As Boolean '�Ƿ�����ԤԼ����
    Dim strSQL As String
    Dim blnIsClearQueue As Boolean
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    blnIsCurDayReservations = False
    blnIsClearQueue = False
    If mblnIsScheduleOrder Then
        '�ж��Ƿ�ԤԼ����
        strSQL = "Select ID,ԤԼ��ʼʱ�� From Ӱ��ԤԼ��¼ Where ҽ��Id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����ԤԼ��Ϣ", mobjCurStudyInfo.lngAdviceId)
        If rsTemp.RecordCount > 0 Then
            blnIsCurDayReservations = True
            
            '�Ѿ�ԤԼ�����ж�ԤԼ���ں͵�ǰʱ���Ƿ�һ�£������һ���򵯳�������ʾ
            '���ԤԼ���ں͵�ǰ����һ�£���ֱ�ӽ��뱨��
            If Format(nvl(rsTemp!ԤԼ��ʼʱ��), "yyyy-mm-dd") <> Format(zlDatabase.Currentdate, "yyyy-mm-dd") Then
                If MsgBoxD(Me, "��ǰ����ԤԼ�ļ������Ϊ " & Format(nvl(rsTemp!ԤԼ��ʼʱ��), "yyyy-mm-dd") & "���뵱ǰʱ�䲻һ�£��Ƿ����������", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                    Exit Sub
                Else
                    blnIsClearQueue = True
                    blnIsCurDayReservations = False
                End If
            End If
        End If
    End If
    
    If mobjCurStudyInfo.lngPatientFrom = 4 Then    '�������첡�˲�ִ�����¹���
        Call zlDatabase.ExecuteProcedure("zl_PeisLockAdviceState(" & mobjCurStudyInfo.lngAdviceId & ")", Me.Caption)
    End If
    
    If Me.ActiveControl Is Nothing Then
        blnFocusFind = False
    Else
        blnFocusFind = (Me.ActiveControl.Name = "txtFilter")
    End If
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        With frmRISRequest
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = mobjCurStudyInfo.lngSendNo
            .mlngAdviceId = mobjCurStudyInfo.lngAdviceId
            .mstrPatientName = mobjCurStudyInfo.strPatientName
            .mintEditMode = 2 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
            .mlngCurDeptId = IIf(mblnAllDepts, mobjCurStudyInfo.lngExeDepartmentId, mlngCur����ID)
            .mstrCur���� = zlStr.NeedName(mstrCur����)
            .mlngResultState = 0
            
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            
            If .mlngResultState <> 0 Then  '�ɹ�����
                Call CheckExecuteInterface(EInterfaceExeTime.������)
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
                
                If .mblnIsRelationImage = True Then
                    '�������ǰ����ͼ��������Զ��������������ｫ��Ӱ��ͼ��ģ�����ˢ��
                    If Not mfrmWork_PacsImg Is Nothing Then
                        Call mfrmWork_PacsImg.zlUpdateAdviceInf(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, mobjCurStudyInfo.blnMoved)
                        Call mfrmWork_PacsImg.zlRefreshFace(True)
                    End If
                End If
                
                If mSysPar.blnAutoOpenReport Then Call Menu_RichEPR(conMenu_Edit_Modify)              '��ʼ����Զ��򿪱���
                
                If .mlngResultState = 2 Then
                    '��������Ŷӽкţ����ұ������Զ��Ŷӣ��򱨵�����Ҫ�����ŶӽкŶ���......
                    
                    If mSysPar.blnUseQueue And mSysPar.blnAutoInQueue And Not mobjQueue Is Nothing Then
                        strSQL = "Select Id from �ŶӽкŶ��� Where ҵ������=1 And ҵ��ID=[1]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯԤԼ����", .mlngAdviceId)
                        
                        If blnIsCurDayReservations And rsTemp.RecordCount > 0 Then
                            Call mobjQueue.ReservationQueue(.mlngAdviceId)
                        Else
                            If blnIsClearQueue Then
                                'ɾ��֮ǰԤԼʱ���Ŷӣ������������ɶ���
                                strSQL = "zl_�ŶӽкŶ���_�Զ������(1," & "'ҵ��ID=" & mobjCurStudyInfo.lngAdviceId & "',0)"
                                Call zlDatabase.ExecuteProcedure(strSQL, "ɾ����������")
                            End If
                            
                            '������Ҫ����Ķ�������
                            If .mstrTechnicRoom = "" Then
                                '���δ�գ�����Ҫ����ü����Ŀ��Ӧ����Ŀ������߿��ҵĶ�����
                                Call mobjQueue.zlGetInQueueInf(mobjCurStudyInfo.lngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                            Else
                                '�����Ϊ�գ���д���Ӧ��ִ�м�����
                                strQueueName = .mstrCur���� & "-" & .mstrTechnicRoom
                                strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                            End If
                            
                            Call mobjQueue.zlInPacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
                        End If
                    End If
                    
                End If
                
                '����״̬ͬ����Ϣ
                Call mobjMsgCenter.Send_Msg_StateSync(.mlngAdviceId)
                
                If mobjCurStudyInfo.lngPatientFrom <> 3 Then
                    Call mobjMsgCenter.Send_Msg_Arrange(.mlngAdviceId)
                End If
            End If

        End With
    Else
        With frmPatholRIS
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = mobjCurStudyInfo.lngSendNo
            .mlngAdviceId = mobjCurStudyInfo.lngAdviceId
            .mintEditMode = 2 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
            .mlngCurDeptId = IIf(mblnAllDepts, mobjCurStudyInfo.lngExeDepartmentId, mlngCur����ID)
            .mintImgCount = mintImgCount
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            If .RefreshPatiInfor(True) = True Then  'ˢ�²���
                .mblnOk = False
                .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            End If
            If .mblnOk Then  '�ɹ�����
                Call CheckExecuteInterface(EInterfaceExeTime.������)
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
                If mSysPar.blnAutoOpenReport Then Call Menu_RichEPR(conMenu_Edit_Modify)              '��ʼ����Զ��򿪱���
                
                '����״̬ͬ����Ϣ
                Call mobjMsgCenter.Send_Msg_StateSync(.mlngAdviceId)
            End If
            
        End With
    End If
    
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'�Ŷӽк����
Private Sub zlInPacsQueue()
    Dim strQueueName As String
    Dim strCodeNo As String
    If mobjQueue Is Nothing Then Exit Sub
    
    '������Ҫ����Ķ�������
    If Trim(mobjCurStudyInfo.strExeRoom) = "" Then
        '���δ�գ�����Ҫ����ü����Ŀ��Ӧ����Ŀ������߿��ҵĶ�����
        Call mobjQueue.zlGetInQueueInf(mobjCurStudyInfo.lngAdviceId, mlngCur����ID, strQueueName, strCodeNo)
    Else
        '�����Ϊ�գ���д���Ӧ��ִ�м�����
        strQueueName = zlStr.NeedName(mstrCur����) & "-" & mobjCurStudyInfo.strExeRoom
        strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(mobjCurStudyInfo.strExeRoom, mlngCur����ID)
    End If
    Call mobjQueue.zlInQueue(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.strPatientName, mlngCur����ID, strQueueName, mobjCurStudyInfo.strExeRoom, strCodeNo)
End Sub




Private Sub Menu_Manage_ȡ������()
On Error GoTo errhandle
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim lngResult As Long
    Dim strMsg As String

    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
  
    If mobjCurStudyInfo.intStep <= 1 Then Call Menu_Manage_ȡ���Ǽ�: Exit Sub  '����������
    '------------------------------------��ǩ������Ҫ�Ȼ���ǩ�����ٳ���
    strSQL = "Select Distinct B.���ʱ�� From ����ҽ������ A, ���Ӳ�����¼ B Where A.����ID=B.Id And A.ҽ��ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Ƿ�ǩ��", mobjCurStudyInfo.lngAdviceId)
    
    If Not rsTemp.EOF Then
        If nvl(rsTemp!���ʱ��, "") <> "" Then 'ǩ������
            MsgBoxD Me, "��ǰ���˵ļ�鱨���Ѿ�ǩ��,����ȡ�����,���Ȼ���ǩ��!", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    
    '��������ȡ�Ļ�����Ƭ�����ܽ���ȡ��
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        strSQL = "select count(1) as ���� from ��������Ϣ a, ����ȡ����Ϣ b where a.����ҽ��ID=b.����ҽ��ID and a.ҽ��ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, GetWindowCaption, mobjCurStudyInfo.lngAdviceId)
        If rsTemp.RecordCount > 0 Then
            If Val(nvl(rsTemp!����)) > 0 Then
                Call MsgBoxD(Me, "�ü����ִ��ȡ�Ĳ��������ܽ���ȡ����", vbInformation, GetWindowCaption)
                Exit Sub
            End If
        End If
    End If

    If mobjCurStudyInfo.strStudyUID <> "" And Not CheckPopedom(mstrPrivs, "���ͼ��") Then
        MsgBoxD Me, "��û��������ͼ��Ȩ��,�������ͼ��,���Բ���ȡ��������!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strMsg = "������Ϣ��������" & mobjCurStudyInfo.strPatientName & "   �Ա�" & mobjCurStudyInfo.strPatientSex & "   ���䣺" & mobjCurStudyInfo.strPatientAge & "   ���ţ�" & mobjCurStudyInfo.strStudyNum & "��" & vbCrLf & _
             "ȡ�����˱��μ�齫ɾ����Ӧ�ļ��ͼ��ͼ�鱨�棬�Ƿ������"

    If MsgBoxD(Me, strMsg, vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    'ȡ���Ŷ���Ϣ
    If mSysPar.blnUseQueue = True And Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Call mobjQueue.zlCancelPacsQueue(mobjCurStudyInfo.lngAdviceId)
    End If
    
    '�����RIS����վ������ͼ��������PACS�У�����Ҫ��ȡ��������Ȼ���ٵ���ZL_Ӱ����_CANCEL����ȡ������
    If mlngModule = G_LNG_PACSSTATION_MODULE And mobjCurStudyInfo.intImageLocation = 1 Then
        'ȡ��ͼ�����
        Call XWUnmatchImage(mobjCurStudyInfo.lngAdviceId, 0)
    End If
    
    'ȡ�����棬�޸����ݿ�״̬��ɾ����Ӱ�����¼��
    strSQL = "ZL_Ӱ����_CANCEL(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & ",0," & mlngCur����ID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        strSQL = "ZL_������_����(" & mobjCurStudyInfo.lngAdviceId & ")"
        zlDatabase.ExecuteProcedure strSQL, GetWindowCaption
    End If
        
        Call CheckExecuteInterface(EInterfaceExeTime.ȡ������ʱ)
    
    '���ͼ��������PACS����ɾ��Ӱ���ļ���Ŀ¼
    If mobjCurStudyInfo.intImageLocation = 0 Then
        RemoveCheckImages mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo
    End If
    
    If TabWindow.Selected.tag = "Ӱ��ɼ�" Then
        If Not mobjWork_ImageCap Is Nothing Then
            Call mobjWork_ImageCap.zlRefreshData(True)
        End If
    End If
    
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    '����״̬������Ϣ
    Call mobjMsgCenter.Send_Msg_StateCancel(mobjCurStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_����Ӱ��()
On Error GoTo errhandle
    Dim lngResult As Long
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If

    lngResult = -1
    '�����ģ���ΪRIS����վ����������������ݿ��ѯδƥ���ͼ���¼
    If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
        lngResult = XWShowUnMatched(Me, mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.strImgType)
        
        If lngResult = 0 Then
            'ͼ������ɹ���,ʹ��ֵΪ1
            mobjCurStudyInfo.intImageLocation = 1
            Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
        End If
    Else
        frmSelectMuli.ShowImageReleation mlngModule, mobjCurStudyInfo.lngAdviceId, mstrPrivs, mobjCurStudyInfo.blnMoved, IIf(mlngModule = G_LNG_PACSSTATION_MODULE, False, True), mlngCur����ID, 2, mobjCurStudyInfo.strImgType
        
        If Not frmSelectMuli.mblnOk Then Exit Sub
        lngResult = 0
    End If
    
    If lngResult <> 0 Then Exit Sub
    
    Call AfterReleationImage(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, 2, True)
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Menu_Dept_Select(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errhandle
    Dim i As Integer
    Dim objDepartmentMenu As CommandBarControl
    Dim objControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim CtlFont As StdFont
    Dim strFontType As String
    Dim strOldSchemeValue(4) As String
    Dim cbrMenuBar As CommandBarPopup

    If Not mblnInitOk Then Exit Sub
    
    Set CtlFont = New StdFont
    strFontType = IIf(IsUseClearType = True, "΢���ź�", "����")
    CtlFont.Name = strFontType
    CtlFont.Size = gbytFontSize
            
    mstrSelQueueRooms = ""
    
    If mlngCur����ID <> Control.DescriptionText Or (Control.DescriptionText <> 0 And mblnAllDepts = True) Then
        mstrRPTExecutor = UserInfo.����
        
        If Not mobjWork_Report Is Nothing And mrtReportType = �����ĵ��༭�� Then
            Call mobjWork_Report.SetDocCreator(mstrRPTExecutor)
        End If
        
        stbThis.Panels(4).Text = "����ҽ����" & mstrRPTExecutor & "   ���ҽ����" & Split(stbThis.Panels(4).Text, "���ҽ����")(1)
                
        Set mobjCurStudyInfo = GetNullAdviceInf
        
        '�����л�������û�����´����˵��͹���ģ�飬Ҳû�е���cbrMain.RecalcLayout�������Ҫʹ�øö������ÿ����л���Ŀ�����Ϣ
        Set objDepartmentMenu = cbrMain.FindControl(, conMenu_View_Filter * 10#)
        
        If Control.DescriptionText = 0 Then
            'ѡ�����п���
            mblnAllDepts = True
            mlngCur����ID = 0
        
            If Not objDepartmentMenu Is Nothing Then objDepartmentMenu.Caption = "��ǰ����:ȫ������"
            
            Call mobjPacsQueryWrap.DepChange(mstrCanUse����IDs)
            Set cbrFilter.options.Font = CtlFont
            
            If Not mobjQueue Is Nothing And mlngModule = G_LNG_PACSSTATION_MODULE Then
                mobjQueue.ChangeToAllDept mblnAllDepts
            End If
        Else
            'ѡ�񵥸�����
            mblnAllDepts = False
            
            mlngCur����ID = Control.DescriptionText
            mstrCur���� = Mid(Control.Caption, 1, InStrRev(Control.Caption, "(") - 1)
            
            mrtReportType = GetDeptPara(mlngCur����ID, "����༭��", 0)                 '����༭��
            Call mobjPacsQueryWrap.ReportTypeChange(mrtReportType)
            
            If Not objDepartmentMenu Is Nothing Then objDepartmentMenu.Caption = "��ǰ����:" & mstrCur����
            
            Call SetParaUseImgSignValid(mlngCur����ID)
            Call InitModuleParameter

            If Not mobjWork_ImageCap Is Nothing Then
                Call mobjWork_ImageCap.zlInitModule(gcnOracle, glngSys, mlngModule, mstrPrivs, mlngCur����ID, Me.hwnd, Me, True)
                '�����������ڸ����Ƿ�ʹ�ú�̨ͼ
                mobjWork_ImageCap.ModuleNo = mlngModule
            End If
            
            If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
            If Not mfrmWork_PacsImg Is Nothing Then Call mfrmWork_PacsImg.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
            If Not mobjWork_His Is Nothing Then
                If mblnAllDepts Then
                    Call mobjWork_His.zlInitModule(mlngModule, mstrPrivs, UserInfo.����ID, Me)
                Else
                    Call mobjWork_His.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
                End If
            End If
            
            '�����л�������������Ŷӽкţ�������Ŷӽк�ҳ��
            If mSysPar.blnUseQueue = True Then
                If mobjQueue Is Nothing Then
                    mstrWorkModule = mstrWorkModule & ";�Ŷӽк�ģ��;"
                    
                    Set mobjQueue = New frmWork_Queue
                    Call mobjQueue.zlInitPacsQueueCfg(mlngModule, mlngCur����ID, zlStr.NeedName(mstrCur����), mstrPrivs, mblnAllDepts, Me)
                    
                    TabWindow.InsertItem 13, "�Ŷӽк�", mobjQueue.hwnd, 10011
                    TabWindow.Item(TabWindow.ItemCount - 1).tag = "�Ŷӽк�"
                    
                    Call picWindow_Resize
                Else
                    Call mobjQueue.zlInitPacsQueueCfg(mlngModule, mlngCur����ID, zlStr.NeedName(mstrCur����), mstrPrivs, mblnAllDepts, Me)
                End If
                
                '��ݽкŽ���
                If mSysPar.blnQueueQuick Then
                    If Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
                        Call mobjQueue.OpenQueueQuick(GetSelQueueRooms(True), Me)
                    End If
                End If
            Else
                If mSysPar.blnUseQueue = False And Not mobjQueue Is Nothing Then
                    mstrWorkModule = Replace(mstrWorkModule, ";�Ŷӽк�ģ��;", "")
                    
                    For i = 0 To TabWindow.ItemCount - 1
                        If TabWindow.Item(i).tag = "�Ŷӽк�" Then
                            Call TabWindow.RemoveItem(i)
                            Exit For
                        End If
                    Next i
                    
                    mobjQueue.CloseQueueQuick
                    
                    Set mobjQueue = Nothing
                    
                    Call picWindow_Resize
                End If
            End If
            
            If mlngModule = G_LNG_PACSSTATION_MODULE Then
                If Not mfrmWork_PacsImg Is Nothing And InStr(mstrWorkModule, ";Ӱ��ͼ��ģ��;") > 0 Then
                    '����Ӱ���������Ӳ˵��͹�����
                    Call mfrmWork_PacsImg.zlMenu.zlCreateMenu(Me.cbrMain)
                    Call mfrmWork_PacsImg.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
                End If
            End If
            
            'Ϊ���ֱ���˵��ܹ�һֱ��ʾ��������Ҫ�Ա���˵����д���
            If Not mobjWork_Report Is Nothing And (InStr(mstrWorkModule, ";Ӱ�񱨸�ģ��;") > 0 Or InStr(mstrWorkModule, ";�������ģ��;") > 0) Then
                Call mobjWork_Report.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
                
                '���������Ӧ�˵��͹�����������༭��ʹ�ò�ͬ��ʽ��ʱ�򣬴����Ĳ˵���ͬ��
                Call mobjWork_Report.zlMenu.zlCreateMenu(Me.cbrMain)
                Call mobjWork_Report.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
                         
                If TabWindow.Selected.tag = "������д" Then
                    Call mobjWork_Report.SetReportWindow(True)
                Else
                    Call mobjWork_Report.SetReportWindow(False)
                End If
                
            End If
            
            '�л���Ϣ�Ľ��տ���
            Call mobjMsgCenter.ChangeMsgReceiveDept(mlngCur����ID)
            
            With mobjPacsQueryWrap.CurPacsQuery.GetSqlScheme
                strOldSchemeValue(0) = .Query
                strOldSchemeValue(1) = .FilterCfgCount
                strOldSchemeValue(2) = .Detail
                strOldSchemeValue(3) = .SerachCfgCount
                strOldSchemeValue(4) = .ShowCfgCount
            End With
            
            Call mobjPacsQueryWrap.DepChange(mlngCur����ID)
            '�ж��Ƿ���Ҫ�л�����
            Call mobjPacsQueryWrap.CurPacsQuery.LoadQueryScheme(glngSys, mlngModule, mlngCur����ID, UserInfo.ID)
            
'            With mobjPacsQueryWrap.CurPacsQuery.GetSqlScheme
'                If strOldSchemeValue(0) <> .Query Or strOldSchemeValue(1) <> .FilterCfgCount Or strOldSchemeValue(2) <> .Detail Or strOldSchemeValue(3) <> .SerachCfgCount Or strOldSchemeValue(4) <> .ShowCfgCount Then
'                    '��������ȫ��ͬ�����¼����µ�Ĭ�Ϸ���  LSQ TODO  �����Ż�����Ҫ���¼��ط����Ĵ���ʽ
                    Call ExecuteDefaultQueryScheme
                                        
                    Set cbrMenuBar = cbrMain.FindControl(, conMenu_Manage_Query)
                    Call mobjPacsQueryWrap.RefreshCustomQueryMenu(cbrMenuBar, mintQueryState, tabScheme, mSysPar.blnQuickTabDisplayScheme)
                    With cbrMenuBar.CommandBar
                        Set objControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_QueryCFG, "��ѯ����", "", 0, True)
                        Set objControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_QueryCfgUserScheme, "���÷�������", "", 0, False)
                        Set objControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_QueryTabDisplayScheme, "��ʾ���÷�����ǩ", "", 0, True)
                        objControl.Checked = mSysPar.blnQuickTabDisplayScheme
                        objControl.CloseSubMenuOnClick = False
                        
                        Set objControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_QueryValidTime, "��������", "", 0, False)
                        objControl.Checked = mSysPar.blnQueryValidTime
                        objControl.CloseSubMenuOnClick = False
                    End With
'                End If
'            End With
            
            Set cbrFilter.options.Font = CtlFont
        End If
        
        Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
        
        Call CreateWorkModuleMenu
        
        Call cbrMain.RecalcLayout
        
        'ˢ���Ŷӽк�ģ�����ݣ�����Ѿ�����
        Call RefreshPacsQueueData
        
        Call CreateAuditorMenu(cbrMain.FindControl(, conMenu_ManagePopup).CommandBar.FindControl(, conMenu_Manage_SendAudit))
        
        'ˢ���Ƿ�����ԤԼ
        Call IsSchedule(mlngCur����ID, mobjCurStudyInfo.lngAdviceId, 0, mblnIsScheduleDept, mblnIsScheduleOrder)
    End If
    
    If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
        glngXWDeptID = mlngCur����ID
    End If
    
    Call ReCreatCbrMenu(cbrMain)
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub AddPlugInToolBarMenu(cbrControls As CommandBarControls, ByVal lngModule As Long)

    Dim cbrControl As CommandBarControl
    Dim i As Long, j As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blFirst As Boolean

On Error GoTo ErrorHand
    
    blFirst = True
    strSQL = "Select a.id,a.���� as ��������,a.�Ƿ����� as ��������,a.ִ������,b.�������,b.���� as ��������,b.�Ƿ����� as ��������,b.�Ƿ�����Ҽ��˵�,b.�Ƿ���빤����,b.vbs�ű� from Ӱ�����ҽ� a, Ӱ�������� b " & _
             "Where a.�Ƿ�����=1 and  b.�Ƿ�����=1 and a.id = b.���id And (a.����ģ��=0 or a.����ģ��=[1]) Order By a.id,b.�������"
             
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��������������˵�", lngModule)
    
    If rsTemp.RecordCount > 0 Then

        While Not rsTemp.EOF
                
            j = j + 1
            
            If Val(nvl(rsTemp!�Ƿ���빤����)) = 1 Then
                If blFirst = True Then
                    Set cbrControl = CreateModuleMenu(cbrControls, xtpControlButton, conMenu_Manage_PacsPlugIn * 10000# + j, nvl(rsTemp!��������), "", 2325, True)
                    blFirst = False
                Else
                    Set cbrControl = CreateModuleMenu(cbrControls, xtpControlButton, conMenu_Manage_PacsPlugIn * 10000# + j, nvl(rsTemp!��������), "", 2325, False)
                End If
                
                cbrControl.Parameter = nvl(rsTemp!VBS�ű�)
                cbrControl.DescriptionText = Val(nvl(rsTemp!ִ������))
                cbrControl.Category = Val(nvl(rsTemp!��������)) & "," & Val(nvl(rsTemp!�Ƿ�����Ҽ��˵�)) & "," & Val(nvl(rsTemp!�Ƿ���빤����))
            End If
            
            Call rsTemp.MoveNext
        Wend
    End If
            
    Exit Sub
ErrorHand:
    Call err.Raise(0, , "����˵���ӵ��������쳣-" & err.Description)
End Sub

Private Sub RefreshCustomPlugInMenu(objQueryMenu As Object, ByVal lngModule As Long)
    Dim objCurQueryMenu As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim blFirstMenu As Boolean '�Ƿ��һ�����ܲ˵��������ж��Ƿ���Ҫ�ӷָ��ߣ�
    Dim i As Long, j As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngAppId As Long

On Error GoTo ErrorHnad
    
    blFirstMenu = True
    If objQueryMenu Is Nothing Then Exit Sub
    
    Set objCurQueryMenu = objQueryMenu
    
    For i = 1 To objCurQueryMenu.CommandBar.Controls.Count
        objCurQueryMenu.CommandBar.Controls(1).Delete
    Next
    
    strSQL = "Select a.id,a.���� as ��������,a.�Ƿ����� as ��������,a.ִ������,b.�������,b.���� as ��������,b.�Ƿ����� as ��������,b.�Ƿ�����Ҽ��˵�,b.�Ƿ���빤����,b.vbs�ű� from Ӱ�����ҽ� a, Ӱ�������� b " & _
             "Where a.id = b.���id and a.�Ƿ�����=1 and b.�Ƿ�����=1 And (a.����ģ��=0 or a.����ģ��=[1]) Order By a.id,b.�������"
             
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��������˵�", lngModule)
    
    With objCurQueryMenu.CommandBar
        If rsTemp.RecordCount > 0 Then
            i = 65
            While Not rsTemp.EOF
                j = j + 1
                
                If lngAppId <> nvl(rsTemp!ID) Then
                    Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_Manage_PacsPlugLevel2 * 10000# + nvl(rsTemp!ID), nvl(rsTemp!��������), "", , False)
                    lngAppId = nvl(rsTemp!ID)
                Else
                    Set cbrPopControl = cbrMain.FindControl(, conMenu_Manage_PacsPlugLevel2 * 10000# + nvl(rsTemp!ID), , True)
                End If

                If Not cbrPopControl Is Nothing Then
                    If blFirstMenu Then
                        Set cbrControl = CreateModuleMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsPlugIn * 10000# + j, nvl(rsTemp!��������), "", , True)
                    Else
                        Set cbrControl = CreateModuleMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsPlugIn * 10000# + j, nvl(rsTemp!��������), "", , False)
                    End If
                End If
                                
                cbrControl.Parameter = nvl(rsTemp!VBS�ű�)
                cbrControl.DescriptionText = Val(nvl(rsTemp!ִ������))
                cbrControl.Category = Val(nvl(rsTemp!��������)) & "," & Val(nvl(rsTemp!�Ƿ�����Ҽ��˵�)) & "," & Val(nvl(rsTemp!�Ƿ���빤����))
                
                blFirstMenu = False
                
                Call rsTemp.MoveNext
            Wend
        End If
            
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_PacsPlugCfg, "�������", "", 181, False)
    End With

    Exit Sub
ErrorHnad:
    Call err.Raise(0, , "���²���˵��쳣-" & err.Description)
End Sub

Private Sub Menu_View_Refresh_click()
On Error GoTo errhandle
    Call RefreshList
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Home_click()
On Error GoTo errhandle
    zlHomePage hwnd
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_StatusBar_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errhandle
    Me.stbThis.Visible = Not Me.stbThis.Visible
    Control.Checked = Not Control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Button_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errhandle
    Dim i As Integer
    
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    Control.Checked = Not Control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Size_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errhandle
    Me.cbrMain.options.LargeIcons = Not Me.cbrMain.options.LargeIcons
    Control.Checked = Not Control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errhandle
    Dim i As Integer, cbrControl As CommandBarControl
    Dim intStyle As Integer

    For i = 2 To cbrMain.Count
        If Me.cbrMain(i).Controls.Count >= 1 Then
            intStyle = Me.cbrMain(i).Controls(i).Style
            If intStyle = xtpButtonIconAndCaption Then
                intStyle = xtpButtonIcon
                Me.cbrMain(i).ShowTextBelowIcons = False
            Else
                intStyle = xtpButtonIconAndCaption
                Me.cbrMain(i).ShowTextBelowIcons = True
            End If
        End If
        
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = intStyle
        Next
    Next
    
    Control.Checked = Not Control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function GetDeptName(lngDeptId As Long, strDeptStrings As String) As String
'ͨ�����õĿ��Ҵ�����ȡָ������ID�Ŀ�������
On Error GoTo errhandle
    Dim strDepts() As String
    Dim i As Integer
    
    strDepts = Split(strDeptStrings, "|")
    For i = 0 To UBound(strDepts)
        If Split(strDepts(i), "_")(0) = lngDeptId Then
            GetDeptName = Split(strDepts(i), "_")(1)
            Exit For
        End If
    Next i
Exit Function
errhandle:
    If ErrCenter = 1 Then Resume
End Function

Private Function GetPatholNum(ByVal strSureNum As String) As String
'�ֽ�ȷ�Ϻ���
    Dim lngFindSplitChar As Long
    
    lngFindSplitChar = InStr(1, strSureNum, "-")
    
    If lngFindSplitChar > 0 Then
        GetPatholNum = UCase(Mid(strSureNum, 1, lngFindSplitChar - 1))
    Else
        GetPatholNum = UCase(strSureNum)
    End If
    
End Function

Private Sub cmdClear_Click()
    Call mobjPacsQueryWrap.CurPacsQuery.EmbedConditionRestore
End Sub

Private Sub cmdDo_Click()
    Call mobjPacsQueryWrap.ExecuteQuery(C_QUERY_���ݼ���)
    If mobjPacsQueryWrap.SqlScheme.AutoRefreshTimeLen > 0 Then TimerRefresh.Enabled = True
End Sub

Private Sub cmdMore_Click()
    Call mobjPacsQueryWrap.ExecuteQuery(C_QUERY_�������)
    If mobjPacsQueryWrap.SqlScheme.AutoRefreshTimeLen > 0 Then TimerRefresh.Enabled = True
End Sub

Private Sub Form_Activate()
On Error GoTo errhandle
    '�жϵ�ǰ����ģ���Ƿ�Ӱ��ɼ�ģ�飬����ǣ����жϲɼ�ģ���Ƿ��ʼ��������Ѿ���ʼ�������˳��ù��̣�����Ͷ�����г�ʼ��������ʾ
    '��Ϊ��ͬһ����̨�У����ͬʱ�򿪲�����Ƶ�ɼ�ģ�齫���л�������һϵͳ�˳�ʱ���ɼ�ģ��Ҳ�����ͷţ�����л��ص�ǰϵͳ����Ҫ�ж��Ƿ���³�ʼ���ɼ�ģ��
    Call Form_Resize
    If Not mobjWork_ImageCap Is Nothing Then
        If mobjWork_ImageCap.ModuleNo <> 0 And mobjWork_ImageCap.ModuleNo <> mlngModule Then mobjWork_ImageCap.ModuleNo = mlngModule
    End If
    If Not mblnInitOk Then Exit Sub
    If TabWindow.Selected Is Nothing Then Exit Sub
    If TabWindow.Selected.tag <> "Ӱ��ɼ�" Then Exit Sub
    If Not mobjWork_ImageCap Is Nothing Then
        With mobjWork_ImageCap
            Call .zlUpdateStudyInf(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, mobjCurStudyInfo.blnMoved, mobjCurStudyInfo.blnIsReported)
            Call .zlRefreshVideoWindow
            Call .zlRefreshData(False)
        End With
    End If
    
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub imgFun_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    timFun.Enabled = False
End Sub

Private Sub mfrmHistory_OnDoWork(ByVal LngAdvice As Long, ByVal strFuncName As String)
    Select Case strFuncName
        Case "��Ƭ"
            Call OpenViewer(1, mobjPacsCore, LngAdvice, False, Me)
        Case "�Ա�"
            Call OpenViewer(1, mobjPacsCore, LngAdvice, True, Me)
        Case "�鿴����"
            Call OpenReport
    End Select

End Sub

Private Sub mfrmHistory_OnListLostFocus()
    TimerHistory.Enabled = True
End Sub

Private Sub mfrmHistory_OnListMouseClick(ByVal LngAdvice As Long, ByVal X As Long, ByVal Y As Long, ByVal blnClear As Boolean)
On Error GoTo errH

    If blnClear Then
        Call mobjPacsQueryWrap.ClearHistoryFollow(False)
    Else
        Call mobjPacsQueryWrap.DohistoryFollow(LngAdvice, X, Y)
    End If
    
    Exit Sub
errH:
End Sub

Private Sub mfrmHistory_OnListMove()
    Call mobjPacsQueryWrap.ClearHistoryFollow(True)
End Sub

Private Sub mfrmHistory_OnLoadCfg(strListCfg As String)
    strListCfg = mobjPacsQueryWrap.HistoryCfg
End Sub

Private Sub mfrmHistory_OnRefresh(ByVal lngCount As Long)
    If lngCount = 0 Then
        TabExtra.Item(2).Caption = "���μ��"
    Else
        TabExtra.Item(2).Caption = "���μ��(" & lngCount & ")"
    End If
End Sub

Private Sub mfrmHistory_OnSaveCfg(ByVal strListCfg As String)
     mobjPacsQueryWrap.HistoryCfg = strListCfg
End Sub
    
Private Sub mfrmHistory_OnSelectStudy(ByVal LngAdvice As Long, ByVal strAdvice As String, ByVal blnEmbed As Boolean)
On Error GoTo errhandle
    Dim i As Integer
    Dim StudyInfoTMP As New clsStudyInfo  '���ڲ����ļ����Ϣ
    
    If blnEmbed Then mobjPacsQueryWrap.AdviceId = 0
    mobjPacsQueryWrap.AdviceId = 0
        
    If LngAdvice = 0 Then
        Exit Sub
    End If
    
    If InStr(strAdvice, ",") > 0 Then
        mblHaveHistory = True
    Else
        mblHaveHistory = False
    End If

    If LngAdvice <> mobjHistoryStudyInfo.lngAdviceId Then
        mblnIsHistory = True
        
        If blnEmbed Then
            Set StudyInfoTMP = mobjCurStudyInfo
            Set mobjCurStudyInfo = mobjPacsQueryWrap.GetBaseInfo(LngAdvice, 0)
            Set mobjHistoryStudyInfo = mobjCurStudyInfo
            
            Call RefreshModuleAdviceInf
        
            Call ShowTab '���ݲ����ṩ��ͬѡ�
        
            Call ShowBillList(cbrMain.FindControl(, conMenu_Manage_RequestPrint, , True))  '��ʾ�ɴ�ӡ�����Ƶ���:֮���Լ�ʱ����,��Ϊ��ʹ��F2�ȼ�
        
            If Not TabWindow.Selected Is Nothing Then
                Call ConfigSubForm(TabWindow.Selected)
            End If
        
            Call NotificationAllModuleRefresh

            If mstrFirstTab <> "" Then '��Ϊ�ձ�ʾ��������ҳ��ʾ,��TabWindow����ˢ��
    
                For i = 0 To TabWindow.ItemCount - 1
                    If InStr(TabWindow.Item(i).tag, mstrFirstTab) > 0 And TabWindow.Item(i).Visible Then
                        Exit For
                    End If
                Next
    
                If i = TabWindow.ItemCount Then    'ûѭ�����˴�����1������TAB
                    For i = 0 To TabWindow.ItemCount - 1
                        If TabWindow.Item(i).Visible Then
                            Exit For
                        End If
                    Next i
                End If
    
                'ˢ��ҳ�棬����ʾ������ҳ
                If TabWindow.Item(i).Selected Then
                    Call RefreshTabWindow
                Else
                    TabWindow.Item(i).Selected = True
                End If
                
                Set mobjCurStudyInfo = StudyInfoTMP
            Else
                Call RefreshTabWindow
            End If
        Else
            Set mobjHistoryStudyInfo = mobjPacsQueryWrap.GetBaseInfo(LngAdvice, 0)
        End If
        
        Call mobjPacsQueryWrap.FillAppend(mobjHistoryStudyInfo.lngAdviceId, mobjHistoryStudyInfo.lngPatId, mobjHistoryStudyInfo.blnMoved, rftHistoryFollow)
        If rftHistoryFollow.Text = "" Then
            PicFollowHistory.Visible = False
        End If
        
    End If
    Exit Sub
errhandle:
    MsgBoxD Me, "OnSelectStudy ������Ϣ��" & err.Description, vbInformation, Me.Caption
End Sub

Private Sub mfrmHistory_OnViewReport(ByVal LngAdvice As Long)
    Call OpenReport
End Sub

Private Sub mobjPacsQueryWrap_OnLocateBackColor()
    cmdLocate.BackColor = IIf(mobjPacsQueryWrap.DefaultLocate, &HFF00&, &H8000000F)
    cmdFind.BackColor = IIf(mobjPacsQueryWrap.DefaultLocate = False, &HFF00&, &H8000000F)
End Sub

Private Sub mobjPacsQueryWrap_OnQueueRoomChanged()
    Call RefreshPacsQueueData
End Sub

Private Sub mobjPacsQueryWrap_OnSwipeCard()
On Error GoTo errH
    Call VsfListDbClick(True)
errH:
End Sub

Private Sub mobjPacsQueryWrap_OnClearFace()
'�������ݣ�ִ�в�ѯ��û������ʱ����ս���ؼ���ʾ
On Error GoTo errhandle
    Dim i As Integer
    
    If vsfList.Rows < 2 Then
        '��û������ʱ��֪ͨˢ�¹���ģ������ص�����
        Set mobjCurStudyInfo = GetNullAdviceInf
        
        Call RefreshModuleAdviceInf
        Call NotificationAllModuleRefresh

        If TabWindow.Selected Is Nothing Then
            'ѡ���һ������ģ��
            For i = 0 To TabWindow.ItemCount - 1
                If TabWindow.Item(i).Visible Then
                    TabWindow(i).Selected = True

                    mblnAutoRefreshList = False
                    Exit For
                End If
            Next i
        End If

        Call RefreshTabWindow

        mblnAutoRefreshList = False
        mblnIsLoading = False
        
        '���½�TAb����  ������Ϣ  ��ʷ��� ״̬ͼ
        
        For i = imgState.Count - 1 To 0 Step -1
            imgState(i).Visible = False
        Next
    
        imgStep.Visible = False
        LabFlag����.Visible = False
        LabFlagӤ��.Visible = False
        LabFlag��ɫͨ��.Visible = False
        LabFlagΣ��״̬.Visible = False
        LabFlag��Ⱦ��״̬.Visible = False
        LabFlag����.Visible = False
        
        labCollectionInfo.Visible = False
        labPatientInfo.Visible = False
        labPatientAge.Visible = False
        
        Call mfrmHistory.ClearData
        TabExtra.Item(2).Caption = "���μ��"
        
        Call mobjPacsQueryWrap.FillAppend(0, 0, False, rtxtAppend)
        
        stbThis.Panels(2).Text = "�� " & 0 & " ����¼": stbThis.Panels(2).Alignment = sbrCenter
        stbThis.Panels(3).Text = ""
        
    End If
    Exit Sub
errhandle:
    MsgBoxD Me, "OnClearFace ������Ϣ��" & err.Description, vbInformation, Me.Caption
End Sub

Private Sub mobjWork_Report_AfterReportSave()
    If mobjPacsQueryWrap.SqlScheme.AutoRefreshTimeLen > 0 Then TimerRefresh.Enabled = True
End Sub

Private Sub mobjWork_Report_AfterSetRptQuality(ByVal lngAdviceId As Long, ByVal strValue As String)
    mobjCurStudyInfo.strImageQuality = strValue
    Call UpdateQueryListData(Nothing, lngAdviceId)
End Sub

Private Sub mobjWork_Report_BeforeReportSave()
    TimerRefresh.Enabled = False
End Sub

Private Sub mobjWork_Report_BeforeBatPrint(ByRef strIds As String)
    Dim i As Integer 

    strIds = ""
    For i = 1 To vsfList.Rows - 1
        If strIds <> "" Then strIds = strIds & ","
        strIds = strIds & vsfList.Cell(flexcpText, i, vsfList.ColIndex("ҽ��ID"))
    Next

End Sub

Private Sub PatiIdentify_KeyPress(KeyAscii As Integer)
    TimerRefresh.Enabled = False
End Sub

Private Sub picDataSearchContainer_Resize()
'���� ���ݼ���������ȴ���9000��ʱ�򣬰�ť����߼���ﵽ600���Ҳ����������м䰴��������
'��ʼ״̬��
On Error GoTo errhandle
    Dim intTMP As Single '���������ʵ����Ӱ�ť����ѯ����ľ���
    Dim lngWidth As Integer '��ѯ������
    Dim lngBaseWidth As Long '��ť�Ͳ�ѯ����ľ���
    Dim lngBaseWidthDataSearchContainer As Long '�����������
    Dim lngMove As Long

    If picDataSearchContainer.Width = Screen.Width Then Exit Sub
    lngBaseWidth = 200

    If gbytFontSize = 9 Then
        If picDataSearchContainer.Width <= 5500 Then
            lngWidth = 4000
        ElseIf picDataSearchContainer.Width >= 6500 Then
            lngWidth = 5000
        Else
            lngWidth = picDataSearchContainer.Width - 1500
        End If

    ElseIf gbytFontSize = 12 Then
        If picDataSearchContainer.Width <= 6000 Then
            lngWidth = 4500
        ElseIf picDataSearchContainer.Width >= 7000 Then
            lngWidth = 5500
        Else
            lngWidth = picDataSearchContainer.Width - 1500
        End If
    Else
        If picDataSearchContainer.Width <= 6500 Then
            lngWidth = 5500
        ElseIf picDataSearchContainer.Width >= 8000 Then
            lngWidth = 6500
        Else
            lngWidth = 5500 + 0.66 * (picDataSearchContainer.Width - 6500)
        End If
    End If

'    If gbytFontSize = 9 Then
        intTMP = 0
'    ElseIf gbytFontSize = 12 Then
'        intTMP = 150
'    Else
'        intTMP = 300
'    End If

    lngBaseWidthDataSearchContainer = lngBaseWidth + lngWidth + 2 * intTMP + cmdDo.Width

    If picDataSearchContainer.Width > lngBaseWidthDataSearchContainer Then
        lngMove = (picDataSearchContainer.Width - lngBaseWidthDataSearchContainer) / 2
        Call picDataSearch.Move(lngMove - 400, 0, lngWidth, picDataSearchContainer.Height)
        Call cmdDo.Move(lngMove + lngWidth + lngBaseWidth + intTMP - 400)
    Else
        Call picDataSearch.Move(-400, 0, lngWidth, picDataSearchContainer.Height)
        Call cmdDo.Move(lngWidth + lngBaseWidth + intTMP - 400)
    End If

    
    Call mobjPacsQueryWrap.CurPacsQuery.EmbedSize(picDataSearch)
    cmdMore.Visible = mobjPacsQueryWrap.CurPacsQuery.IsMoreEmbedInput And cmdDo.Visible
    
    If Not cmdMore.Visible Then
        Call cmdClear.Move(cmdDo.Left, cmdClear.Top, cmdDo.Width)
        cmdClear.Width = cmdDo.Width
    Else
        Call cmdClear.Move(cmdDo.Left, cmdClear.Top, 0.5 * cmdDo.Width)
    End If
    Call cmdMore.Move(cmdClear.Left + cmdClear.Width)
    
errhandle:
End Sub

Private Sub cmdFind_Click()
On Error GoTo errhandle
    mobjPacsQueryWrap.DefaultLocate = False
    
    cmdLocate.BackColor = IIf(mobjPacsQueryWrap.DefaultLocate, &HFF00&, &H8000000F)
    cmdFind.BackColor = IIf(mobjPacsQueryWrap.DefaultLocate = False, &HFF00&, &H8000000F)
    
    If Me.MousePointer = 0 Then
        Me.MousePointer = 13
        vsfList.Row = 0
        Call mobjPacsQueryWrap.Find(True, True)
        TimerRefresh.Enabled = False
        Me.MousePointer = 0
    Else
        Exit Sub
    End If
    Exit Sub
errhandle:
    MsgBoxD Me, "���Ҳ�����������Ϣ��" & err.Description, vbInformation, Me.Caption
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '���ع���ģ��ʱ���������˳�����
    If Not mblnInitOk Then
        Cancel = True
        Exit Sub
    End If
    
    If mblnMenuDownState Then
        If MsgBoxD(Me, "��ǰ������δ��ɣ�ǿ���˳�������ɳ����쳣���Ƿ������", vbYesNo, "����") = vbNo Then Cancel = True
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call AdjustFace(picList.Height, picList.Width)
End Sub

Private Sub imgFun_Click(Index As Integer)
'Ŀǰ�ṩ�ĸ��� "����" ȡ������"  �޸���Ϣ" ��д����"
On Error GoTo errH
    Dim i As Integer
    
    If mblnMenuDownState Then Exit Sub

    Select Case imgFun(Index).ToolTipText
        Case C_FUNC_STR_����
            Call Menu_Manage_����
        Case C_FUNC_STR_��д����
            Call Menu_RichEPR(conMenu_Edit_Modify)
        Case C_FUNC_STR_�鿴������Ϣ
            frmDegreeCard.ShowMe mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.lngPageID, Me
        Case C_FUNC_STR_��Ƭ
            If Not mfrmWork_PacsImg Is Nothing Then
                Call mfrmWork_PacsImg.zlMenu.zlExecuteMenu(conMenu_Img_Look)
            End If
        Case C_FUNC_STR_���
            Call Menu_Manage_����������
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mfrmRISRequest_HaveRegist()
    Dim strQueueName As String
    Dim strCodeNo As String
    With mfrmRISRequest
        If .mlngResultState <> 0 Then '�ɹ�����
            '��������Ŷӽкţ��򱨵�����Ҫ�����ŶӽкŶ���......
            If mSysPar.blnUseQueue And Not mobjQueue Is Nothing And .mlngResultState = 2 Then
                '������Ҫ����Ķ�������
                If .mstrTechnicRoom = "" Then
                    '���δ�գ�����Ҫ����ü����Ŀ��Ӧ����Ŀ������߿��ҵĶ�����
                    Call mobjQueue.zlGetInQueueInf(mobjCurStudyInfo.lngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                Else
                    '�����Ϊ�գ���д���Ӧ��ִ�м�����
                    strQueueName = .mstrCur���� & "-" & .mstrTechnicRoom
                    strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                End If
                
                Call mobjQueue.zlInPacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
            End If
            
            '������������Ϣ
            Call mobjMsgCenter.Send_Msg_Request(.mlngAdviceId)
        End If
    End With
End Sub

Private Sub mobjCaptureHot_OnKeyBoardLHook(ByVal lngMsg As Long, ByVal lngVkCode As Long, ByVal lngScanCode As Long, ByVal lngFlags As Long)
On Error GoTo errhandle
    Dim lngWindowPID As Long
    Dim lngVideoPID As Long
    Dim lngCurrentPID As Long

'    If lngMsg <> WM_KEYDOWN Then Exit Sub
    If Trim(mstrCaptureHot) = "" And Trim(mstrCaptureAfterHot) = "" And Trim(mstrCaptureAfterTagHot) = "" Then Exit Sub
    
    mCaptureMsg.lngMsg = lngMsg
    mCaptureMsg.lngVirtualKey = lngVkCode
    mCaptureMsg.lngScanKey = lngScanCode
    mCaptureMsg.lngFlags = lngFlags
    
    '����ֱ����Hook�ص�������ʹ��ActiveExe�������ط�������������δ֪�������
    timerCapture.Enabled = True
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mobjEvent_OnWork(objEvent As Object, ByVal lngWorkType As TWorkEventType, ByVal lngAdviceId As Long, ByVal other As Variant)
'��Ӧ����ģ��ִ�в����󴥷����¼�
On Error GoTo errhandle
    Dim strSQL As String
    Dim strRoom As String
    Dim i As Integer
    Dim j As Integer
    Dim strStudyUID As String
    Dim strGrades() As String
    
    Dim lngCurRow As Long
    Dim lngColIndex As Long
    
    Select Case lngWorkType
        Case TWorkEventType.wetDelImg
            Call CheckExecuteInterface(EInterfaceExeTime.ɾ��ͼ��ʱ)
        Case TWorkEventType.wetGetImg           '��ȡͼ��QR��***************************************
            Call UpdateQueryListData(Nothing, lngAdviceId)
            
        Case TWorkEventType.wetTechDo           '��ʦִ��***************************************
            If mobjCurStudyInfo.lngAdviceId = lngAdviceId Then
            
                mobjCurStudyInfo.blnIsTechincalSure = Val(other)
                If Val(other) = 1 Then mobjCurStudyInfo.strDoDoctor = UserInfo.����
                
                Call UpdateQueryListData(Nothing, lngAdviceId)
            End If
            
        Case TWorkEventType.wetChangeImgType    '�ı�Ӱ������***************************************
           Call UpdateQueryListData(Nothing, lngAdviceId)
        
        Case TWorkEventType.wetLockStudy, TWorkEventType.wetUnLockStudy        '�������,�������
            '�޸ı�ǩҳ����ʾ��ʽ�ͱ���
            For i = 0 To TabWindow.ItemCount - 1
                If TabWindow(i).Caption Like "*Ӱ��ɼ�*" Then
                    If lngWorkType = wetLockStudy Then
                        TabWindow(i).Image = 10013
                        TabWindow(i).Caption = "��" & other & "�� Ӱ��ɼ�"
                    Else
                        TabWindow(i).Image = conMenu_Cap_Dynamic
                        TabWindow(i).Caption = "Ӱ��ɼ�"
                    End If
                    Exit For
                End If
            Next i
            
            'ˢ��Ƕ�뱨���е�����ͼͼ�������Ƶ�ɼ���ͼ��
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngWorkType, lngAdviceId, other)
            
        Case TWorkEventType.wetCaptureFirstImg, TWorkEventType.wetDelAllImg, TWorkEventType.wetUpdateImg  '�ɼ���һ��ͼ��***************************************
            '���¼���б�
            
            strStudyUID = other
            
            If lngWorkType = wetCaptureFirstImg Then
                
                '���¼���б�
                Call UpdateStudyListState(lngAdviceId, strStudyUID, True, True)

                
                If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(mtProRep)
                
                '����״̬ͬ����Ϣ
                Call mobjMsgCenter.Send_Msg_StateSync(lngAdviceId)
            ElseIf lngWorkType = wetDelAllImg Then
                '���¼���б�
                Call UpdateStudyListState(lngAdviceId, strStudyUID, False, True)

                '����״̬ͬ����Ϣ
                Call mobjMsgCenter.Send_Msg_StateCancel(lngAdviceId)
                Call CheckExecuteInterface(EInterfaceExeTime.ɾ��ͼ��ʱ)
            End If

            If mobjCurStudyInfo.lngAdviceId <> lngAdviceId Then Exit Sub
            
            'ˢ��Ƕ�뱨���е�����ͼͼ�������Ƶ�ɼ���ͼ��
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngWorkType, lngAdviceId, other)
            
            'ˢ��Ƕ���ؼ챨��������½�����ͼͼ��
            If lngWorkType = wetUpdateImg Then If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(mtProRep)
        Case wetChangeUser
            '�����û�ʱ����Ҫ���жϱ����Ƿ���Ҫ����
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
        
            Call ChangeUser
            
            '�����û�����Ҫˢ�±���༭������Ϊ�û�������ԭ�б���ı༭�û����ߴ����û���Ҫ���и���
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
            
        Case wetPatholRequest       '��������
            Call UpdateQueryListData(Nothing, lngAdviceId)
            
        Case wetPatholQuality       '��������
        
            Call UpdateQueryListData(Nothing, lngAdviceId)
            
        Case wetPatholBatSlices     '��Ƭ��������
            Call UpdateQueryListData(Nothing, lngAdviceId)
            
        Case wetPatholBatSpeExm     '�ؼ���������
            Call UpdateQueryListData(Nothing, lngAdviceId)
            
        Case wetSpecimenAccept      '�걾����
            Call UpdateQueryListData(Nothing, lngAdviceId)
            lngCurRow = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("ҽ��ID"))
        
            If lngCurRow > 0 Then
                'ˢ����������ģ������
                If Not mobjWork_Pathol Is Nothing Then
                    Call mobjWork_Pathol.zlUpdateAdviceInf(lngAdviceId, 0, 2, False)
                    Call mobjWork_Pathol.NotificationRefresh(mtAll)
                End If
            End If
            
        Case wetSpecimenSave        '�걾����
            '�걾�����ˢ��ȡ��ģ������
            If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(TModuleType.mtMaterial)
            
        Case wetMaterialSure        'ȡ��ȷ��
            
            Call UpdateQueryListData(Nothing, lngAdviceId)
            
            'ˢ����Ƭģ������
            If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(TModuleType.mtSlices)
            
        Case wetMaterialSave        '�Ŀ鱣��
            'ˢ����Ƭģ������
            If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(TModuleType.mtSlices)
            
        Case wetSlicesSure          '��Ƭȷ��
            Call UpdateQueryListData(Nothing, lngAdviceId)
        Case wetSpeExamSure         '�ؼ�ȷ��
            Call UpdateQueryListData(Nothing, lngAdviceId)
        Case wetViewEprReport       'Ԥ�����Ӳ�������
            Dim strRepInf() As String
            
            strRepInf = Split(other & ",,", ",")
            
            If Val(strRepInf(0)) <= 0 Then Exit Sub
            
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.ViewEPRReport(Val(strRepInf(0)), IIf(Val(strRepInf(1)) = 1, True, False))
        
        Case wetViewPacsImage       'Ԥ��Pacsͼ��
            '����100��ͼ������У�Ĭ��ÿ��5�Ŵ�һ��
            Call OpenViewer(2, mobjPacsCore, lngAdviceId, False, Me)
            
        Case wetRejectReport        '���汻����
            lngCurRow = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("ҽ��ID"))

            If lngCurRow <= 0 Then Exit Sub
                        
            Call UpdateQueryListData(Nothing, lngAdviceId)
            '����״̬ͬ����Ϣ
            Call mobjMsgCenter.Send_Msg_StateSync(lngAdviceId)
        Case wetPrintFilm
            '����Ƭ��ӡ��Ϣ
            lngCurRow = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("ҽ��ID"))

            If lngCurRow <= 0 Then Exit Sub
            Call UpdateQueryListData(Nothing, lngAdviceId)

        Case wetImageQuality
            strGrades = Split(mSysPar.strImageLevel, ",")
            If Val(other) - 1 <= UBound(strGrades) Then
            
                mobjCurStudyInfo.strImageQuality = strGrades(Val(other) - 1)
                Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
                
            End If
        End Select
    
    Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mobjMsgCenter_OnRecevieMsg(ByVal strMsgItemIdentity As String, ByVal strXmlContext As String, rsData As ADODB.Recordset, objMsgPro As clsMipModule, objXML As clsXML)
'��Ϣ���մ���
On Error GoTo errH
    Dim lngRowIndex As Long
    Dim lngAdviceId As Long
    Dim lngStudyState As Long
    Dim strHint As String
    Dim strSQL As String
    Dim rsReport As ADODB.Recordset
    Dim rsDataMulite As ADODB.Recordset
    Dim rsDataMuliteClone As ADODB.Recordset
    Dim strCurNo As String
    Dim strNodeId As String
    Dim lngChargeState As Long
    
    
    lngAdviceId = 0
    
    '��ȡ��Ϣ�ж�Ӧ��ҽ��ID����
    If strMsgItemIdentity = G_STR_MSG_ZLHIS_PACS_003 Then
        rsData.Filter = "node_name='study_order_id'"
    Else
        rsData.Filter = "node_name='order_id'"
    End If
    
    If rsData.RecordCount > 0 Then
        lngAdviceId = Val(nvl(rsData!node_value))
    End If
    
    
    Select Case strMsgItemIdentity
        Case G_STR_MSG_ZLHIS_CIS_017    '�������
            '������Ϣ��ʾ@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='patient_name'"
            strHint = "���� " & nvl(rsData!node_value) & " ��Ҫ���м�飬�뼰ʱ����"
            
            Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
            
            '�����ݿ���ˢ������
            Call UpdateQueryListData(Nothing, lngAdviceId)
            
        Case G_STR_MSG_ZLHIS_CIS_024    'ҽ������
            '����������ʾ@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='patient_name'"
            strHint = "���� " & nvl(rsData!node_value) & " �ļ��ҽ���ѱ������� "
        
            Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
        
            '�����ݿ���ˢ������
            Call UpdateQueryListData(Nothing, lngAdviceId)
            
        Case G_STR_MSG_ZLHIS_CIS_025    'Σ��ֵ�Ķ�
            '����Ϣƽ̨���õ�����ʾ
            
        Case G_STR_MSG_ZLHIS_CHARGE_003 '���ﻼ�߻��۵���
            'ˢ���շ�״̬��ʾ
            '���ݵ��ݺŲ��Ҷ�Ӧ��ҽ��ID
            rsData.Filter = "node_name='bill_no'"
            If rsData.RecordCount <= 0 Then
                Exit Sub
            End If
            
             lngRowIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("ҽ��ID"))
            If lngRowIndex > 0 Then Call UpdateQueryListData(Nothing, lngAdviceId)
        
        Case G_STR_MSG_ZLHIS_PACS_001   '��鱨����ɣ������ɲ����鱨���������
            '�����б��е���ʾ״̬
            lngRowIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("ҽ��ID"))
            If lngRowIndex > 0 Then Call UpdateQueryListData(Nothing, lngAdviceId)
            
        Case G_STR_MSG_ZLHIS_PACS_002, G_STR_MSG_ZLHIS_PACS_003  '���״̬ͬ������״̬���˴���
            '������汻���أ���Ҫ��������@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='study_cur_state'"
            If nvl(rsData!node_value) = -1 Then
                
                '��Ҫ�жϵ�ǰ�û��Ƿ�Ϊ������
                strSQL = "select ������ from Ӱ�����¼ where ҽ��ID=[1]"
                Set rsReport = zlDatabase.OpenSQLRecord(strSQL, "��ѯ������", lngAdviceId)
                If rsReport.RecordCount > 0 Then
                    If nvl(rsReport!������) = UserInfo.���� Then
                        '������Ϣ
                        rsData.Filter = "node_name='patient_name'"
                        strHint = "����" & nvl(rsData!node_value) & "�ı����ѱ����أ���ע�⴦��"
                        
                        Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
                    End If
                End If
            End If
            
            'ˢ���б��Ӧ��ʾ
            lngRowIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("ҽ��ID"))
            
            If lngRowIndex > 0 Then Call UpdateQueryListData(Nothing, lngAdviceId)
        Case G_STR_MSG_ZLHIS_PACS_004   '��鱨�泷��
            '�����б��е���ʾ״̬
            lngRowIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("ҽ��ID"))
            
            If lngRowIndex > 0 Then Call UpdateQueryListData(Nothing, lngAdviceId)
            
        Case G_STR_MSG_ZLHIS_PACS_005   '���Σ��ֵ֪ͨ
            '�ڿ����ڵ���Σ������@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='patient_name'"
            strHint = "���� " & nvl(rsData!node_value) & "��"
            
            rsData.Filter = "node_name='check_item_title'"
            strHint = strHint & "�����Ŀ " & nvl(rsData!node_value) & " ����Σ�������"
            
            Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
        
            '�����б��е���ʾ״̬
            lngRowIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("ҽ��ID"))
            
            If lngRowIndex > 0 Then Call UpdateQueryListData(Nothing, lngAdviceId)
            
    End Select
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mobjPacsCore_AfterSaveOuterImage(strStudyUID As String)
    '�������ⲿͼ��ˢ��ͼ��������б�
On Error GoTo errhandle
    
    'û�м�¼���˳�
    If mobjCurStudyInfo.lngAdviceId = 0 Then Exit Sub
    
    '�ǵ�ǰ�ļ�飬��ˢ�¼��������б�
    If mobjCurStudyInfo.strStudyUID = strStudyUID Then
        Call mfrmWork_PacsImg.zlRefreshFace(True)
    End If
    
    Exit Sub
errhandle:
    '������
End Sub


Public Sub OnStateChange(ByVal lngEventType As TVideoEventType, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal strOther As String, ByVal dcmImage As DicomImage)
'��Ƶ�ɼ������ص��¼�
    mVideoEventInf.vetEventType = lngEventType
    mVideoEventInf.lngAdviceId = lngAdviceId
    mVideoEventInf.lngSendNo = lngSendNo
    mVideoEventInf.strOtherInf = strOther
    Set mVideoEventInf.dcmImage = dcmImage

    timerVideoEvent.Enabled = True
End Sub

Public Sub OnDockClose()
'�������ڹرջص��¼�
End Sub

Private Sub DoOnStateChange(ByVal lngEventType As TVideoEventType, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal strOther As String, ByVal dcmImage As DicomImage)
'��Ӧ����ģ��ִ�в����󴥷����¼�
On Error GoTo errhandle
    Dim strSQL As String
    Dim strRoom As String
    Dim strStudyUID As String
    Dim i As Long
    Dim lngID As Long '����ִ�з���-���п���
    
    Select Case lngEventType
        Case TVideoEventType.vetImgDeled '����ɾ��ͼ�� ���ڲ���Զ�ִ��
            Call CheckExecuteInterface(EInterfaceExeTime.ɾ��ͼ��ʱ)
            If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(mtProRep)
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngEventType, lngAdviceId, strOther)
        Case TVideoEventType.vetImgCaped
        Case TVideoEventType.vetUseAfterImage, TVideoEventType.vetNotUseAfterImage
            If lngEventType = TVideoEventType.vetUseAfterImage And mlngModule = G_LNG_VIDEOSTATION_MODULE Then
                If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UseAfterImgChanged(True)
            Else
                If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UseAfterImgChanged(False)
            End If
        Case TVideoEventType.vetLockStudy, TVideoEventType.vetUnLockStudy         '�������,�������
            '�޸ı�ǩҳ����ʾ��ʽ�ͱ���
            For i = 0 To TabWindow.ItemCount - 1
                If TabWindow(i).Caption Like "*Ӱ��ɼ�*" Then
                    If lngEventType = vetLockStudy Then
                        TabWindow(i).Image = 10013
                        TabWindow(i).Caption = "��" & strOther & "�� Ӱ��ɼ�"
                        mblnLockState = True
                    Else
                        TabWindow(i).Image = conMenu_Cap_Dynamic
                        TabWindow(i).Caption = "Ӱ��ɼ�"
                        mblnLockState = False
                    End If
                    Exit For
                End If
            Next i
            
     
            'ˢ��Ƕ�뱨���е�����ͼͼ�������Ƶ�ɼ���ͼ��
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngEventType, lngAdviceId, strOther)

            
        Case TVideoEventType.vetCaptureFirstImg, TVideoEventType.vetDelAllImg, TVideoEventType.vetUpdateImg  '�ɼ���һ��ͼ��***************************************
            '���¼���б�
            
            strStudyUID = strOther
            
            If lngEventType = TVideoEventType.vetCaptureFirstImg Then
                '����ʱִ�з��û�ΪӰ��ɼ�ϵͳʱִ�з���
                If (mlngModule = G_LNG_VIDEOSTATION_MODULE And mSysPar.lngVideoStationMoneyExeModle = 1) Or _
                   (mlngModule = G_LNG_PATHSTATION_MODULE And mSysPar.lngPatholStationMoneyExeModle = 1) Then
                    
                    If mblnAllDepts Then
                        If mobjCurStudyInfo.lngExeDepartmentId > 0 Then
                            lngID = mobjCurStudyInfo.lngExeDepartmentId
                        Else
                            lngID = 0
                        End If
                    Else
                        lngID = mlngCur����ID
                    End If
                    
                    strSQL = "Zl_Ӱ�����ִ��(" & lngAdviceId & "," & lngSendNo & ",3,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lngID & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, "ִ�м�����")
                End If
                
                If mblnLockState Then
                
                Else
                    Call UpdateStudyListState(lngAdviceId, strStudyUID, True, True)
                End If
                
                Call RefreshTab(True)
                
                Call CheckExecuteInterface(EInterfaceExeTime.��ͼ��)
                
                If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(mtProRep)
            ElseIf lngEventType = TVideoEventType.vetDelAllImg Then
                If mblnLockState Then
                
                Else
                    Call UpdateStudyListState(lngAdviceId, strStudyUID, False, True)
                End If
                
                Call RefreshTab(False)
                
                Call CheckExecuteInterface(EInterfaceExeTime.ɾ��ͼ��ʱ)
            End If

            If lngEventType = TVideoEventType.vetUpdateImg Then Call CheckExecuteInterface(EInterfaceExeTime.��ͼ��)
                        
            If mobjCurStudyInfo.lngAdviceId <> lngAdviceId Then Exit Sub
            
            'ˢ��Ƕ�뱨���е�����ͼͼ�������Ƶ�ɼ���ͼ��
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngEventType, lngAdviceId, strOther)
            
            'ˢ��Ƕ���ؼ챨��������½�����ͼͼ��
            If lngEventType = TVideoEventType.vetUpdateImg Then If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(mtProRep)
        
        Case TVideoEventType.vetAfterUpdateImg
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngEventType, lngAdviceId, strOther)
            Call CheckExecuteInterface(EInterfaceExeTime.��ͼ��)
            
        Case TVideoEventType.vetImportImage
            Call AfterReleationImage(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, 2, False)
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngEventType, lngAdviceId, strOther)
            
        Case TVideoEventType.vetExportImage
            Call AfterReleationImage(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, 1, False)
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngEventType, lngAdviceId, strOther)
            
        Case TVideoEventType.vetAddReportImg
            '���뱨��ͼ
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngEventType, lngAdviceId, strOther, dcmImage)
    End Select
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AfterReleationImage(ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal intStep As Integer, ByVal lngReleationType As Long, ByVal blnUseMenuReleation As Boolean)
On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    If lngReleationType = 1 Then
        If InStr("345", intStep) > 0 Then
            gstrSQL = "Select ���uid From Ӱ�����¼ Where  ҽ��ID=[1] And ���ͺ�=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngAdviceId, lngSendNo)
            
            If rsTemp.RecordCount > 0 Then
                If IsNull(rsTemp!���UID) Then
                    '����Ӱ����״̬�������ǰҽ���Ѿ�û��ͼ�񣬶��Ҽ�����Ϊ3�����޸�Ϊ2
                    If intStep = 3 Then
                        gstrSQL = "Zl_Ӱ����_State(" & lngAdviceId & "," & lngSendNo & ",2,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
                        zlDatabase.ExecuteProcedure gstrSQL, "ȡ������"
                    End If
                End If
            End If
        End If
    Else
        '����Ӱ����״̬�����ԭ����״̬���ѱ��������޸ĳ��Ѽ�飬
        If intStep = 2 Then
            '��������Ѿ���ͼ�����޸ĳ��Ѽ��
            strSQL = "Select ���UID From Ӱ�����¼ Where ҽ��ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ���ͼ��", lngAdviceId)
            
            If Not IsNull(rsTemp!���UID) Then
                strSQL = "Zl_Ӱ����_State(" & lngAdviceId & "," & lngSendNo & ",3,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
                zlDatabase.ExecuteProcedure strSQL, "����Ӱ��"
            End If
        End If
    End If
    
    Call UpdateQueryListData(Nothing, lngAdviceId)
    
    If Not mobjWork_ImageCap Is Nothing Then
        Call mobjWork_ImageCap.zlRefreshData(True)
    End If
    
    If Not mfrmWork_PacsImg Is Nothing Then
        Call mfrmWork_PacsImg.zlRefreshFace(True)
    End If
    
    If Not mobjWork_Report Is Nothing And blnUseMenuReleation Then
        Call mobjWork_Report.UpdateVideoCaptureState(TVideoEventType.vetAfterUpdateImg, lngAdviceId, "")
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mobjPacsQueryWrap_OnColStatistics(ByVal strStatisticsInfo As String)
    stbThis.Panels(2).Text = "�� " & vsfList.Rows - 1 & " ����¼": stbThis.Panels(2).Alignment = sbrCenter
    stbThis.Panels(3).Text = strStatisticsInfo
End Sub

Private Sub mobjPacsQueryWrap_OnDoStateImage(ByVal lngRow As Long)
'����״̬ͼ
On Error GoTo errH
    Dim i As Integer, j As Integer, k1 As Integer, k2 As Integer
    Dim objClsRelation As New clsScRowRelation
    Dim intImgCount As Integer
    Dim lngLeft As Long
    Dim strValue As String
    
    '�������״̬ͼ
    For i = imgState.Count - 1 To 0 Step -1
        imgState(i).Visible = False
    Next
    intImgCount = 0

    If mobjPacsQueryWrap Is Nothing Then Exit Sub
    If mobjPacsQueryWrap.SqlScheme Is Nothing Then Exit Sub
    If mobjPacsQueryWrap.SqlScheme.ShowCfgCount < 1 Then Exit Sub
    
    With mobjPacsQueryWrap.SqlScheme
        
        For i = 1 To .ShowCfgCount 'i ��������ʾ����
            If .ShowCfg(i).RowRelationCount > 0 Then
                
                For j = 1 To .ShowCfg(i).RowRelationCount 'j�����й���
                
                    Set objClsRelation = .ShowCfg(i).RowRelation(j)
                    If Len(objClsRelation.Icon) > 0 And objClsRelation.IsStateIcon Then '�����ж��Ƿ���������ʾͼ��
                        
                        strValue = vsfList.Cell(flexcpText, lngRow, vsfList.ColIndex(.ShowCfg(i).Name))
                        
                        If (LTrim(strValue) = objClsRelation.TiggerData And objClsRelation.TiggerData <> "[�ǿ�]" And objClsRelation.TiggerData <> "[��]") _
                    Or (Len(Trim(strValue)) = 0 And objClsRelation.TiggerData = "[��]") Or (Len(Trim(strValue)) > 0 And objClsRelation.TiggerData = "[�ǿ�]") Then
                    
                            '���״̬ͼ
                            If intImgCount = 0 Then
                                Set imgState(0).Picture = mobjPacsQueryWrap.GetIcon(objClsRelation.Icon)
                                Call imgState(0).Move(picDetail.Width - imgState(0).Width, C_LAYOUT_BASEHEIGHTOFDETAILINFO - GetMaxImgHeight - 30)
                                imgState(0).Visible = True
                                
                                intImgCount = 1
                            Else
                                If imgState.Count <= intImgCount Then Load imgState(intImgCount)

                                Set imgState(intImgCount).Picture = mobjPacsQueryWrap.GetIcon(objClsRelation.Icon)

'                                ��������λ��
                                lngLeft = 0
                                For k1 = intImgCount To 0 Step -1
                                    '���ȼ����Ѿ����ڵ�ͼ��Ŀ��֮��
                                    lngLeft = lngLeft + imgState(k1).Width
                                Next
                                
                                lngLeft = picDetail.Width - lngLeft

                                Call imgState(intImgCount).Move(lngLeft, C_LAYOUT_BASEHEIGHTOFDETAILINFO - GetMaxImgHeight - 30)
                                imgState(intImgCount).Visible = True

                                intImgCount = intImgCount + 1
                            End If
                            
                        End If
                    End If
                    
                Next  ' for j
            End If
        Next 'for i
    End With
    
    Exit Sub
errH:
    err.Raise -1, "frmPacsQuery", "[DoStateImage]" & vbCrLf & err.Description
    Resume
End Sub

Private Sub mobjPacsQueryWrap_OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'�Ҽ��˵�����˵����ʹ�ò�ѯ����¼��е��ö������б�ؼ�ֱ�ӵ���ԭ�� ��ѯ�������OnMouseUp�ᴦ�����ܸ��湦��
'pacsmain��ߴ������Ҽ��˵����ܣ�������߶�ʹ��vsflist_onMouseUp��λ�õķ��ա�
On Error GoTo errH
    Dim Control As CommandBarControl, Menucontrol As CommandBarControl
    Dim controlPlugIn As CommandBarControl
    Dim plugins As CommandBarControl
    Dim Popup As CommandBar
    Dim strTmp As String
    Dim i As Long
    
    If mobjPacsQueryWrap.ShowingRowCount < 1 Then Exit Sub

    If Button = 2 Then
        Set Popup = cbrMain.Add("�Ҽ��˵�", xtpBarPopup)


        For i = 1 To cbrMain.ActiveMenuBar.Controls.Count
            Set Menucontrol = cbrMain.ActiveMenuBar.Controls(i)

            If (Menucontrol.ID = conMenu_ManagePopup Or Menucontrol.ID = conMenu_Collection) And Menucontrol.type = xtpControlPopup Then
                For Each Control In Menucontrol.CommandBar.Controls
                    '�����Ҽ� "�ղص�" �˵�
                    If Control.ID <> conMenu_Collection_ViewShare And Control.ID <> conMenu_Collection_Manage _
                    And Mid(Control.ID, 1, Decode(InStr(Control.ID, "0") - 1, -1, 0, InStr(Control.ID, "0") - 1)) <> comMenu_Collection_Type _
                    And Mid(Control.ID, 1, Decode(InStr(Control.ID, "0") - 1, -1, 0, InStr(Control.ID, "0") - 1)) <> conMenu_Collection_ViewShare Then
                        '���ޱ������֮ǰ������ģ�鴴�����Ҽ��˵�
                        If Control.ID = conMenu_Manage_Complete Then
                            If Not mfrmWork_PacsImg Is Nothing Then Call mfrmWork_PacsImg.zlMenu.zlPopupMenu(Popup)
                            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.zlMenu.zlPopupMenu(Popup)
                        End If

                        Control.Copy Popup
                    End If
                Next
            ElseIf Menucontrol.ID = conMenu_Manage_PacsPlugIn Then
                For Each Control In Menucontrol.CommandBar.Controls '���������˵�
                    If Control.ID >= conMenu_Manage_PacsPlugLevel2 * 10000# And Control.ID <= conMenu_Manage_PacsPlugLevel2 * 10000# + 9999 Then

                        For Each controlPlugIn In Control.CommandBar.Controls

                            If UBound(Split(controlPlugIn.Category, ",")) = 2 Then '����ĩ���˵�
                                strTmp = Split(controlPlugIn.Category, ",")(1)
                            Else
                                strTmp = controlPlugIn.Category
                            End If
                            
                            If plugins Is Nothing Then
                                Set plugins = Popup.Controls.Add(xtpControlPopup, conMenu_Manage_PacsPlugIn, "�������")
                            End If

                            If Val(strTmp) = 1 Then controlPlugIn.Copy plugins.CommandBar

                        Next

                    End If
                Next
            End If
        Next i


        Popup.ShowPopup
    End If

    Exit Sub
errH:
    MsgBox err.Description & "mobjQueryShow_OnMouseUp", vbExclamation, gstrSysName
End Sub

Private Sub mobjPacsQueryWrap_OnChangeData(ByVal blnAutoRefresh As Boolean, ByVal blnIsSelChange As Boolean)
On Error GoTo errH
    Dim i As Integer
    Dim intCol As Integer
    Dim lngRow As Long
    Dim lngAdviveID As Long 'ҽ��ID
    Dim strInfo As String
    Dim intCount As Integer
        
    If Not mfrmHistory Is Nothing Then
        mfrmHistory.ListRow = 0
    End If

    PicFollowHistory.Visible = False
    intCol = vsfList.ColIndex("ҽ��ID")
    
    lngRow = vsfList.RowSel
    If lngRow = -1 Then Exit Sub

    lngAdviveID = Val(vsfList.TextMatrix(lngRow, intCol))

    mblnIsHistory = False
    
    Set mobjCurStudyInfo = mobjPacsQueryWrap.StudyInfo
    
    If blnIsSelChange Then Call LocateMainWorkModuleTab
    Call DoLabFlag
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����Ǵ�ӡ�嵥�Ĳ��� ��ֹͣ�иı��¼�����Ȼ����ɽ������ˢ��
    If mblnIsPrintMode Then Exit Sub
    
    mblnIsHistory = False
    
    If Not mobjWork_Report Is Nothing And Not TabWindow.Selected Is Nothing Then
        If TabWindow.Selected.tag = "������д" Then
            Call mobjWork_Report.AllowLocate(mblnAutoRefreshList)
        End If
    End If

    mblnRefreshWord = True
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then '�޼�¼ʱ����

        If Not TabWindow.Selected Is Nothing Then
            Call ConfigSubForm(TabWindow.Selected)
        End If

        Call RefreshModuleAdviceInf
        Call RefreshTabWindow
    Else

        mintImgCount = GetScanRequestCount(mobjCurStudyInfo.lngAdviceId)

        Call RefreshModuleAdviceInf

        Call ShowTab '���ݲ����ṩ��ͬѡ�

        Call ShowBillList(cbrMain.FindControl(, conMenu_Manage_RequestPrint, , True))  '��ʾ�ɴ�ӡ�����Ƶ���:֮���Լ�ʱ����,��Ϊ��ʹ��F2�ȼ�

        If Not TabWindow.Selected Is Nothing Then
            Call ConfigSubForm(TabWindow.Selected)
        End If

        '�ж��Ƿ��ֶ�ˢ�µļ���б�������ֶ�ˢ�£�����Ҫ֪ͨ��������ģ�����ˢ�£�...
        If mblnIsCallModuleRefresh Then
        

            Call NotificationAllModuleRefresh
        End If

        If mstrFirstTab <> "" Then '��Ϊ�ձ�ʾ��������ҳ��ʾ,��TabWindow����ˢ��

            For i = 0 To TabWindow.ItemCount - 1
                If InStr(TabWindow.Item(i).tag, mstrFirstTab) > 0 And TabWindow.Item(i).Visible Then
                    Exit For
                End If
            Next

            If i = TabWindow.ItemCount Then    'ûѭ�����˴�����1������TAB
                For i = 0 To TabWindow.ItemCount - 1
                    If TabWindow.Item(i).Visible Then
                        Exit For
                    End If
                Next i
            End If
            
            'ˢ��ҳ�棬����ʾ������ҳ
            If TabWindow.Item(i).Selected Then
                Call RefreshTabWindow
            Else
                TabWindow.Item(i).Selected = True
            End If
        Else
            Call RefreshTabWindow
        End If
    End If

    '���п��ҵĴ���
    If mblnAllDepts Then
        Call mfrmHistory.RefreshHistoryList(mobjCurStudyInfo.lngAdviceId, mlngModule, mobjCurStudyInfo.lngPatId, _
              mobjCurStudyInfo.lngPatientFrom = 2, mlngCur����ID, mstrCanUse����IDs, _
              mobjCurStudyInfo.lngLinkId, True, mSysPar.blnRelatingPatient, , mobjCurStudyInfo.lngBaby)
    Else
        Call mfrmHistory.RefreshHistoryList(mobjCurStudyInfo.lngAdviceId, mlngModule, mobjCurStudyInfo.lngPatId, _
              mobjCurStudyInfo.lngPatientFrom = 2, mlngCur����ID, mstrCanUse����, _
              mobjCurStudyInfo.lngLinkId, False, mSysPar.blnRelatingPatient, , mobjCurStudyInfo.lngBaby)
    End If
    
    ''������ϸ��Ϣ
    With mobjCurStudyInfo
        
        If .strImgType <> "" And .strStudyNum <> "" Then
            strInfo = "(" & .strImgType & ":" & .strStudyNum & ")"
        ElseIf .strImgType <> "" And .strStudyNum = "" Then
            strInfo = "(" & .strImgType & ")"
        ElseIf .strImgType = "" And .strStudyNum <> "" Then
            strInfo = "(" & .strStudyNum & ")"
        Else
            strInfo = ""
        End If
        
        labPatientInfo.Caption = .strPatientName & strInfo & "  " & .strPatientSex
        labPatientInfo.Visible = True
        labPatientAge.Caption = .strPatientAge
        labPatientAge.Visible = True
    End With
    
    If mobjCurStudyInfo.lngPatientFrom = 1 Then
        If Val(mobjCurStudyInfo.strMarkNum) > 0 Then labCollectionInfo = "��:" & mobjCurStudyInfo.strMarkNum & "  "
    ElseIf mobjCurStudyInfo.lngPatientFrom = 2 Then
        If Val(mobjCurStudyInfo.strMarkNum) > 0 Then labCollectionInfo = "ס:" & mobjCurStudyInfo.strMarkNum & "  "
    Else
        labCollectionInfo = ""
    End If
    
    labCollectionInfo = labCollectionInfo & mobjCurStudyInfo.strAdviceContext
    labCollectionInfo = labCollectionInfo & IIf(mobjCurStudyInfo.strCollectionInfo = "", "", "  (��" & mobjCurStudyInfo.strCollectionInfo & ")")
    
    If labCollectionInfo = "" Then
        Call labPatientInfo.Move(2 * C_LAYOUT_LISTLEFT + imgStep.Width + 60, C_LAYOUT_LISTLEFT + (540 - labPatientInfo.Height) / 2)
    Else
        labCollectionInfo.Visible = True
        labPatientAge.Visible = True
        Call labPatientInfo.Move(2 * C_LAYOUT_LISTLEFT + imgStep.Width + 60, C_LAYOUT_LISTLEFT)
    End If
    Call labCollectionInfo.Move(2 * C_LAYOUT_LISTLEFT + imgStep.Width + 60, labPatientInfo.Top + labPatientInfo.Height)
    Call labPatientAge.Move(labPatientInfo.Left + labPatientInfo.Width + TextWidth("  "), labPatientInfo.Top)
    
    If mobjCurStudyInfo.blnInfancy Then
        labPatientAge.ForeColor = vbRed
    Else
        labPatientAge.ForeColor = vbBlack
    End If
    
    Select Case mobjCurStudyInfo.strStuStateDesc
        Case "�ѵǼ�"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_�Ǽ�).Picture        '          "�Ǽ�"
        Case "�ѱ���"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_����).Picture        '          "����"
        Case "�Ѽ��"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_���).Picture        '          "���"
        Case "�ѱ���"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_���).Picture        '          "���"
        Case "�����"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_���).Picture        '          "���"
        Case "�����"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_���).Picture        '          "���"
        Case "��д��"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_��д).Picture        '          "��д"
        Case "�Ѳ���"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_����).Picture        '          "����"
        Case "�Ѿܾ�"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_�ܾ�).Picture        '          "�ܾ�"
        Case Else
            If App.LogMode = 0 Then
                MsgBox err.Description & "δ֪�ļ�����", vbInformation, gstrSysName
            End If
    End Select
    
    imgStep.Visible = True
    
    If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.SetblHaveReport
    
    On Error Resume Next


    'ˢ���Ƿ�����ԤԼ
    Call IsSchedule(mlngCur����ID, mobjCurStudyInfo.lngAdviceId, 0, mblnIsScheduleDept, mblnIsScheduleOrder)
    Exit Sub
errH:
    MsgBox err.Description & "mobjPacsQueryWrap_OnSelChange", vbExclamation, gstrSysName
End Sub

Private Sub mobjPetitionCap_RefreshState(ByVal blnState As Long)
    Dim lngAdviceId As Long
    Dim intCol As Integer
    
    intCol = vsfList.ColIndex("ҽ��ID")
    lngAdviceId = Val(vsfList.TextMatrix(vsfList.RowSel, intCol))

    If lngAdviceId > 0 Then Call UpdateQueryListData(Nothing, lngAdviceId)
End Sub

Private Sub mobjQueue_OnCallAboutLock(ByVal lngType As Long, strLockedName As String, ByVal blnLockPara As Boolean)
On Error GoTo errhandle
'104686��أ����к�������飬
'lngType����  1:�ж��Ƿ������˲��������Ƿ��Ѿ��б������ļ��,����ֱ�ӽ���        2:���²���
'strLockedName   ��="" ������û��Ӱ�죬����˵���Ѿ����ò������ҷ���֮ǰ�����ļ�黼������
'blnLockPara   ���ڸ���PacsMain�еĲ���
    Dim i As Integer
    Dim intPosition As Integer
    Dim strTmp As String
            
    If lngType = 1 Then
    '�ж��Ƿ������˲������ж��Ƿ������˼��
        If mSysPar.blnLockAfterCall Then
            strLockedName = ""
            '�ж��Ƿ��Ѿ��������
            For i = 0 To TabWindow.ItemCount - 1
                If TabWindow(i).Caption Like "*Ӱ��ɼ�*" And TabWindow(i).Image = 10013 Then
                    '�������
                    Call mobjWork_ImageCap.LockStudy(2, 0, 0, 0, 0)
'                    strTmp = TabWindow(i).Caption
'
'                    intPosition = InStr(strTmp, "��")
'                    If intPosition > 0 Then
'                        strLockedName = Mid(strTmp, 1, intPosition)
'                    Else
'                        strLockedName = "δ֪��ʽ�ļ��"
'                    End If

'                    MsgBox "���������ļ��" & strLockedName

                    Exit For
                End If
            Next i
        End If
    ElseIf lngType = 2 Then
    '���²���
        mSysPar.blnLockAfterCall = blnLockPara
    End If
    
    Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mobjQueue_OnCalled(ByVal lngAdviceId As Long, ByVal strRoom As String, ByVal TCallWay As zlQueueOper.TCallWay)
    Dim intRowIndex As Integer
    Dim lngSendNo As Long
    Dim lngStudyState As Long
    Dim blnMoved As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
On Error GoTo errhandle

    intRowIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("ҽ��ID"))
    Call QueueDataConsistency(lngAdviceId, strRoom, intRowIndex)
    
    If TCallWay = cwBroadcast Or TCallWay = cwWaitRoom Then Exit Sub
        
    If mSysPar.blnLockAfterCall Then
    
        '�����߼��ж��Ƿ����á�ͬ����λ������б�����δ���ã���Ҫ����ҵ��ID��ȡ��Ҫ�����ļ�飬���Ѿ����ã�ֻ��Ҫ������
        'intRowIndex=-1˵������б���û����ʾ�Ŷ��б������ݣ���Ҫ����������
        If mSysPar.blnSynStudylist Then
            If intRowIndex = -1 Then
            
                '���ݿ��л�÷��ͺţ����״̬��ת��״̬
                strSQL = "Select b.���ͺ�,b.ִ�й��� from  Ӱ�����¼ a,����ҽ������ b where a.ҽ��ID =[1] and a.ҽ��id = b.ҽ��id "
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�����Ҫ��������Ϣ", lngAdviceId)
                
                If rsTemp.RecordCount > 0 Then
                    lngSendNo = Val(nvl(rsTemp!���ͺ�))
                    lngStudyState = Val(nvl(rsTemp!ִ�й���))
                    blnMoved = 0
                Else
                    MsgBoxD Me, "����ȷ����Ҫ��������Ϣ���Զ�����ʧ�ܣ����ֶ�����", vbInformation, "���к��Զ�����"
                    Exit Sub
                End If
                
                '�������
                Call mobjWork_ImageCap.LockStudy(1, lngAdviceId, lngSendNo, lngStudyState, blnMoved)
            Else
                '�������
                Call mobjWork_ImageCap.LockStudy(3, 0, 0, 0, False)
            End If
            
        Else
            If intRowIndex = -1 Then
                '���ݿ��л�÷��ͺţ����״̬��ת��״̬
                strSQL = "Select b.���ͺ�,b.ִ�й��� from  Ӱ�����¼ a,����ҽ������ b where a.ҽ��ID =[1] and a.ҽ��id = b.ҽ��id "
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�����Ҫ��������Ϣ", lngAdviceId)
                
                If rsTemp.RecordCount > 0 Then
                    lngSendNo = Val(nvl(rsTemp!���ͺ�))
                    lngStudyState = Val(nvl(rsTemp!ִ�й���))
                    blnMoved = 0
                Else
                    MsgBoxD Me, "����ȷ����Ҫ��������Ϣ���Զ�����ʧ�ܣ����ֶ�����", vbInformation, "���к��Զ�����"
                    Exit Sub
                End If
                
            Else
                lngSendNo = mobjCurStudyInfo.lngSendNo
                lngStudyState = mobjCurStudyInfo.intStep
                blnMoved = 0
            End If
            
            '�������
            Call mobjWork_ImageCap.LockStudy(1, lngAdviceId, lngSendNo, lngStudyState, blnMoved)
        End If
        
    End If
        
    Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mobjQueue_OnQueueQuick(blnOpenQuick As Boolean)
    On Error GoTo errhandle
    
    mSysPar.blnQueueQuick = blnOpenQuick
    
    If mSysPar.blnUseQueue = True Then
        '��ݽкŽ���
        If mSysPar.blnQueueQuick Then
            If Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
                Call mobjQueue.OpenQueueQuick(GetSelQueueRooms(True), Me)
            End If
        End If
    End If
    Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mobjWork_Report_AfterOpenRich(ByVal lngOrderID As Long, ByVal strDocId As String)
'����д���ں���
    '�����ѡ�򿪱���ͬʱ��Ƭ��������򿪹�Ƭվ
    If mSysPar.blnShowImgAfterReport = True Then
        If Not mfrmWork_PacsImg Is Nothing Then
            Call mfrmWork_PacsImg.zlMenu.zlExecuteMenu(conMenu_Img_Look)
        End If
    End If
End Sub

Private Sub mobjWork_Report_AfterReleationImage(ByVal lngOrderID As Long, ByVal lngSendNo As Long, ByVal intStep As Integer, ByVal lngReleationType As Long)
On Error GoTo errhandle
    Call AfterReleationImage(lngOrderID, lngSendNo, intStep, lngReleationType, False)
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mobjWork_Report_DocPluginAction(ByVal actionType As Long, ByVal data As String, ByVal tag As String)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
On Error GoTo errhandle
    If actionType = 5 And Trim(data) <> "" And (Trim(tag) = "����ͼ��" Or Trim(tag) = "ȡ������") Then
        '����ҽ��ID��ȡ���ͺźͼ�����
        strSQL = "select b.���ͺ�,b.ִ�й��� from  Ӱ�����¼ a,����ҽ������ b where a.ҽ��ID =[1] and a.ҽ��id = b.ҽ��id"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", Val(data))
        
        If rsTemp.RecordCount > 0 Then
            Call AfterReleationImage(data, Val(nvl(rsTemp!���ͺ�)), Val(nvl(rsTemp!ִ�й���)), IIf(Trim(tag) = "����ͼ��", 2, 1), False)
        End If
    End If
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbrMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub conMenu_WorkModule_Click()
On Error GoTo errhandle
    Dim frmWorkModule As New frmWorkModuleCfg
    
    frmWorkModule.blnIsUseQueue = mSysPar.blnUseQueue
    Call frmWorkModule.ShowWorkModuleCfg(mlngModule, Me)
    
    '�������ù���ģ��ҳ��
    If frmWorkModule.blnIsOk Then
        
        mblnInitOk = False '��ֹ���Ӵ�����ع����ж��Ӵ������ˢ��
        
        Call InitSubForm
        
        mblnInitOk = True
        
        Call ShowTab
        
        Call picWindow_Resize
        
        '���û�м�����ݣ��������������ģ�飬ֻ��ʾģ�鱳��
        If tcDisable.Visible Then Call tcDisable.Translucence
        
        If Not TabWindow.Selected Is Nothing Then Call TabWindow_SelectedChanged(TabWindow.Selected)
        
    End If
    
    Call Unload(frmWorkModule)
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbrMain_Execute(ByVal objControl As XtremeCommandBars.ICommandBarControl)
On Error GoTo errhandle
    Dim Control As XtremeCommandBars.ICommandBarControl
    Dim i As Long
    Dim str��ʦһ As String, str��ʦ�� As String, strִ�м� As String
    Dim intRowIndex As Integer
    Dim strSys1 As String
    Dim strSys2 As String
    Dim bytSize As Byte
    
    If mintQueryState <> 1 And objControl.ID <> conMenu_Manage_Query And objControl.ID <> conMenu_Manage_QueryCFG Then
        MsgBoxD Me, "û���Ѿ����������������õĲ�ѯ���ã����Ƚ�������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mblnMenuDownState Then Exit Sub

    '������Ҫ����id���Ҷ�Ӧ�Ĳ˵���Ŀ����Ϊͨ���󶨿�ݼ�ִ��ʱ����������һ��ֻ��id��û�������κ���Ϣ��control�˵���
    Set Control = cbrMain.FindControl(, objControl.ID, True, True)
    If Control Is Nothing Then
        '����ò˵�Ϊ���Ӳ����༭�����Ҽ��˵�������Ҫ�޸��Ҽ��˵���id����Ϣ
        If Not mobjWork_Report Is Nothing Then
            Call mobjWork_Report.ReplacePopupMenu(objControl)
            
            Set Control = cbrMain.FindControl(, objControl.ID, True, True)
        End If
        
        If Control Is Nothing Then Exit Sub
    End If
    
    If Control.ID = 0 Then Exit Sub
    
    mblnMenuDownState = True
        
    cbrMain.RecalcLayout
    
    'ִ��Ӱ��ͼ���Ӧ�Ĺ���
    If Not mfrmWork_PacsImg Is Nothing Then
        If mfrmWork_PacsImg.zlMenu.zlIsModuleMenu(Control) Then
            Call mfrmWork_PacsImg.zlMenu.zlExecuteMenu(Control.ID)
            
            mblnMenuDownState = False
            Exit Sub
        End If
    End If
    
    If Not mobjWork_ImageCap Is Nothing Then
'            If mobjWork_ImageCap.zlMenu.zlIsModuleMenu(control) Then
'                'ִ��ActivexExe��Ƶ�ɼ���Ӧ�˵�����
'                Call mobjWork_ImageCap.zlMenu.zlExecuteMenu(control.ID)
'
'                mblnMenuDownState = False
'                Exit Sub
'            End If
    End If
    
    'ִ�в������Ӧ����
    If Not mobjWork_Pathol Is Nothing Then
        If mobjWork_Pathol.zlMenu.zlIsModuleMenu(Control) Then
            Call mobjWork_Pathol.zlMenu.zlExecuteMenu(Control.ID)
            
            mblnMenuDownState = False
            Exit Sub
        End If
    End If
    
    'ִ��HISģ���Ӧ����
    If Not mobjWork_His Is Nothing Then
        If mobjWork_His.zlMenu.zlIsModuleMenu(Control) Then
            If mintChangeUserState = 2 Then  '�������û������������
                MsgBoxD Me, "��ͳһ�û����ٲ�����", vbInformation, gstrSysName
            Else
                Call mobjWork_His.zlMenu.zlExecuteMenu(Control.ID)
            End If

            mblnMenuDownState = False
            Exit Sub
        End If
    End If
    
    If Not mobjWork_Report Is Nothing Then
        If mobjWork_Report.zlMenu.zlIsModuleMenu(Control) Then
            'ִ�б�����ع���ʱ���������л�������ģ�飬��������ִ��

            If TabWindow.Selected.tag <> "������д" Then
                For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
                    If TabWindow(i).tag = "������д" And TabWindow(i).Visible = True Then TabWindow(i).Selected = True
                Next
            End If
            
            If Control.Caption <> "������ӡ" Then
                If TabWindow.Selected.tag <> "������д" Then
                    mblnMenuDownState = False
                    Exit Sub
                End If
            End If
            
            Call mobjWork_Report.zlMenu.zlExecuteMenu(Control.ID)
            
            '�����ѡ�򿪱���ͬʱ��Ƭ��������򿪹�Ƭվ
            'ʹ�ñ����ĵ��༭��ʱ����AfterOpenRich�¼��д���
            If (Control.ID = conMenu_PacsReport_Open + 1000000 Or Control.ID = conMenu_Edit_Modify + 1000000 _
                Or Control.ID = conMenu_Edit_Audit + 1000000 Or Control.ID = conMenu_File_Open + 1000000) And _
                mrtReportType <> �����ĵ��༭�� And mSysPar.blnShowImgAfterReport = True Then
                If Not mfrmWork_PacsImg Is Nothing Then
                    Call mfrmWork_PacsImg.zlMenu.zlExecuteMenu(conMenu_Img_Look)
                End If
            End If
            
            mblnMenuDownState = False
            Exit Sub
        Else
            If Control.ID = conMenu_Manage_ReportFirst Or Control.ID = conMenu_Manage_ReportSecond Or _
                Control.ID = conMenu_Manage_ReportThird Or Control.ID = conMenu_Manage_ReportFourth Then
                
                Call mobjWork_Report.zlMenu.zlExecuteMenu(Control.ID + 1000000)
                mblnMenuDownState = False
                Exit Sub
            End If
        End If
    End If
    
    
    Select Case Control.ID

'--------------------------�ļ�------------------
        Case conMenu_File_PrintSet '��ӡ����
            
            Call zlPrintSet
            
        Case conMenu_File_Excel '�嵥��ӡ
            Call Menu_File_Excel_click
            
        Case conMenu_File_Parameter '��������
            Call Menu_File_Parmeter_click
            
        Case ConMenu_File_ShortcutSet '��ݼ�����
            Call Menu_File_ShortcutSet_click
            
        Case conMenu_Pathol_WorkModule  'վ��ģʽ����
            Call conMenu_WorkModule_Click
            
'        Case conMenu_Manage_SetXWParam  '��������PACS�Ĳ���
'            Call Menu_Manage_SetXWParam_click
            
        Case conMenu_File_SendImg '����ͼ��
            Call conMenu_File_SendImg_click
        
        Case conMenu_Check_ViewLink
            Call ViewLinkChecks
        
        Case conMenu_Cap_DevSet         '��Ƶ����
            If Not mobjWork_ImageCap Is Nothing Then
                Call mobjWork_ImageCap.zlShowVideoConfig
                mstrCaptureHot = GetSetting("ZLSOFT", "����ģ��", "�ɼ��ȼ�", "F8")
                mstrCaptureAfterHot = GetSetting("ZLSOFT", "����ģ��", "��̨�ɼ��ȼ�", "F7")
                mstrCaptureAfterTagHot = GetSetting("ZLSOFT", "����ģ��", "��Ǹ����ȼ�", "F6")
            End If
            
        Case conMenu_Manage_ChangeUser
            '�����û�ʱ����Ҫ���жϱ����Ƿ���Ҫ����
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
        
            Call ChangeUser
            
            '�����û�����Ҫˢ�±���༭������Ϊ�û�������ԭ�б���ı༭�û����ߴ����û���Ҫ���и���
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
            
        Case conMenu_Manage_SwitchUser
            '�л��û�ʱ����Ҫ���жϱ����Ƿ���Ҫ����
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
            
            Call SwitchUser
            
            '�л��û�����Ҫˢ�±���༭������Ϊ�û��л���ԭ�б���ı༭�û����ߴ����û���Ҫ���и���
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
            
        Case conMenu_Manage_Change_In   '�����б�
            If dkpMain.Panes(1).hidden = False Then
                dkpMain.Panes(1).Hide
            Else
                dkpMain.ShowPane (1)
            End If
            
        Case conMenu_File_Exit '�˳�
            mblnMenuDownState = False
            Unload Me

'---------------------------���-----------------
        Case conMenu_Manage_RequestPrint * 10# + 1 To conMenu_Manage_RequestPrint * 10# + 9 '��ӡ���Ƶ���
            Call FuncBillPrint(Control)
            
        Case comMenu_Petition_Capture                       'ɨ�����뵥
            Call Menu_Petition_ɨ�����뵥(1)
            
        Case comMenu_Petition_View
            Call Menu_Petition_ɨ�����뵥(0)                '�鿴���뵥
            
        Case conMenu_Manage_Regist                          '�Ǽ�
            Call Menu_Manage_�Ǽ�

            
        Case conMenu_Manage_CopyCheck                       '���ƵǼ�
            Call Menu_Manage_���ƵǼ�
            
        Case conMenu_Manage_Receive                         '����
            Call Menu_Manage_����
            
        Case conMenu_Manage_Redo                            'ȡ���Ǽ�
            Call Menu_Manage_ȡ���Ǽ�
            
        Case conMenu_Manage_ReGet                           '�ٻ�ȡ��
            Call Menu_Manage_�ٻ�ȡ��
            
        Case conMenu_Manage_ThingModi                       '�޸ĵǼ�
            Call Menu_Manage_�޸�
        
        Case conMenu_Manage_CheckList                       '�鿴�������뵥
            Call Menu_Manage_CheckList
            
        Case conMenu_Manage_ExecOnePart                     '�ֲ�λִ��
            Call menu_Manage_ExecOnePart
            
        Case conMenu_Manage_DiseaseQuery                    '��Ⱦ����ѯ
            Call Menu_Manage_DiseaseQuery
            
        Case conMenu_Manage_DiseaseRegist                   '��Ⱦ���Ǽ�
            Call Menu_Manage_DiseaseRegist
        
        Case conMenu_Manage_ModifBaseInfo               '������Ϣ����
            Call Menu_Manage_ModifBaseInfo
        
        Case conMenu_Manage_Logout                          'ȡ������
            Call Menu_Manage_ȡ������
            
        Case conMenu_Manage_InQueue                         '�Ŷӽк����
            Call zlInPacsQueue
            
        Case conMenu_Manage_Schedule                        '���ԤԼ
            Call Menu_Manage_���ԤԼ
            
        Case conMenu_Manage_ScheduleManage                  'ԤԼ����
            Call Menu_Manage_ԤԼ����
            
        Case conMenu_Manage_Transfer                        '����Ӱ��
            Call Menu_Manage_����Ӱ��
            
        Case conMenu_Manage_Cancel                          'ȡ������
            Call Menu_Manage_ȡ������
            
        Case conMenu_Manage_AttachMoney                     '������
            Call Menu_Manage_������
            
        Case conMenu_Manage_CompleteAttach                  '������ɲ���
            Call Menu_Manage_��ɲ�����
            
        Case conMenu_Manage_Review                          '���
            Call Menu_Manage_���
            
        Case conMenu_Tool_Analyse
            Call OpenViewer(1, mobjPacsCore, mobjCurStudyInfo.lngAdviceId, True, Me, "", mobjCurStudyInfo.blnMoved)
        
        Case conMenu_Manage_ReportRelease                   '���淢��
            Call Menu_Manage_���淢��
            
        Case conMenu_Manage_FilmRelease                     '��Ƭ����
            Call Menu_Manage_��Ƭ����
            
        Case conMenu_Manage_SendArrange                     '���Ͱ���
            Call frmSendArrange.ShowMe(Me, mlngCur����ID, mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, str��ʦһ, str��ʦ��, strִ�м�)
            If str��ʦһ <> "" Then
                Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
            End If

        Case conMenu_Manage_ReportExecutor                  '����ִ�У�����Ǳ�����
            Call Menu_Manage_ReportExecutor
            
        Case conMenu_Manage_SendAudit * 10# + 1 To conMenu_Manage_SendAudit * 10# + 99    '�������
            Call Menu_Manage_SendAudit(Control.Caption)
            
        Case conMenu_Manage_PacsCriticalReg, conMenu_Manage_PacsCriticalManage        'Σ��ֵ����
            Call Menu_Manage_CriticalMark(Control.ID)
            
        Case conMenu_Manage_Negative, conMenu_Manage_Positive                  '���������
            Call Menu_Manage_�������(Control.ID)
           
        Case conMenu_Manage_FuHe, conMenu_Manage_JiBenFuHe, conMenu_Manage_BuFuHe   '�������
            Call Menu_Manage_�������(Control.ID)
            
        Case conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel
            Call Menu_Manage_��ɫͨ��(Control.ID)
        
        Case conMenu_Manage_Complete                        '������
            Call Menu_Manage_����������
                
        Case conMenu_Manage_Undone                          'ȡ��������
            Call Menu_Manage_ȡ��������
            
        Case conMenu_Manage_RelatingPatiet                  '��������
            Call Menu_Manage_��������
            
        Case conMenu_Manage_ReLoadPDF
            If mrtReportType = PACS����༭�� Then
                Call mobjWork_Report.ReUpLoad(mobjCurStudyInfo.lngAdviceId)
            ElseIf mrtReportType = ���Ӳ����༭�� Then
                If Not mSysPar.blnPDFTested Then Call TestPDF
                Call CreateReportPDFAndUpLoad(mobjCurStudyInfo.lngAdviceId, Me, mSysPar.strPDFFTPdevice)
            End If
            
        Case conMenu_Manage_Burn                            'ͼ���¼
            Call Menu_Manage_ͼ���¼

        Case conMenu_Manage_LookMecRecord                   '��������
            Call Menu_Manage_��������
            
'----------------------------------------�ղ�---------------------------------------
        Case conMenu_Collection_Manage  '�ղع���
           Call Menu_Manage_�ղع���
        Case conMenu_Collection_To      '�ղص�
           Call Menu_Manage_�ղص�
        Case comMenu_Collection_Type * 10000# To comMenu_Collection_Type * 10000# + 9999  '��̬�ղ����˵�
           Call Menu_Manage_�ղ�������ʾ(Control, 0)
        Case conMenu_Collection_ViewShare * 10000# To conMenu_Collection_ViewShare * 10000# + 9999   '�鿴����
           Call Menu_Manage_�ղ�������ʾ(Control, 1)
        Case conMenu_Manage_QueryCFG
            
            strSys1 = "[ϵͳ.ϵͳ��],[ϵͳ.ģ���],[ϵͳ.����ID],[ϵͳ.�û�ID],[ϵͳ.�û��˺�],[ϵͳ.�û�����]"
            strSys1 = strSys1 & ",[ϵͳ.����������],[ϵͳ.������ʱ��],[ϵͳ.��������],[ϵͳ.����ʱ��]"
            strSys1 = strSys1 & ",[ϵͳ.��ʼ����],[ϵͳ.��������]"


            strSys2 = "[ϵͳ.����ID],[ϵͳ.ҽ��ID]"

            If gbytFontSize = 9 Then
                bytSize = 0
            ElseIf gbytFontSize = 12 Then
                bytSize = 1
            Else
                bytSize = 2
            End If

            Call mobjPacsQueryWrap.CurPacsQuery.ShowSchemeCfg(mlngModule, strSys1, strSys2, bytSize, Me)
            
        Case conMenu_Manage_QueryCfgUserScheme
            Call mobjPacsQueryWrap.CurPacsQuery.ShowUserScheme(mlngModule, mlngCur����ID, Me)
        Case conMenu_Manage_QueryTabDisplayScheme
            '�������ݿ�����ͻ������,��������ѡ��tab��Ŀ
            mSysPar.blnQuickTabDisplayScheme = Not mSysPar.blnQuickTabDisplayScheme
            
            zlDatabase.SetPara "��ʾ���÷�����ǩ", IIf(mSysPar.blnQuickTabDisplayScheme, 1, 0), glngSys, mlngModule

            tabScheme.Visible = mSysPar.blnQuickTabDisplayScheme
            tabScheme.tag = IIf(mSysPar.blnQuickTabDisplayScheme, "1", "0")
            
            
            Call AdjustFace(picList.Height, picList.Width)
        Case conMenu_Manage_QueryValidTime
            '�������ݿ�����ͻ������,��������ѡ��tab��Ŀ
            mSysPar.blnQueryValidTime = Not mSysPar.blnQueryValidTime
            zlDatabase.SetPara "�л���������ʱ�䷶Χ", IIf(mSysPar.blnQueryValidTime, 1, 0), glngSys, mlngModule

            
'----------------------------------------�������������---------------------
        Case conMenu_Manage_PacsPlugCfg
            Call ShowPacsInterfaceCfg
        Case conMenu_Manage_PacsPlugIn * 10000# To conMenu_Manage_PacsPlugIn * 10000# + 100
            Call ExecuteInterfaceFun(Control.Parameter, Control.DescriptionText, False)
'-------------------------------------------------------------------
        Case conMenu_View_Filter '����
            Call mobjPacsQueryWrap.ExecuteQuery(C_QUERY_�������)
'---------------------------�鿴----------------
        Case conMenu_View_ToolBar_Button '������
            Call Menu_View_ToolBar_Button_click(Control)
            
        Case conMenu_View_FontSize_S    'С����
            Call SetFontSize(0)
        Case conMenu_View_FontSize_M    '������
            Call SetFontSize(1)
        Case conMenu_View_FontSize_L    '������
            Call SetFontSize(2)
            
        Case conMenu_View_ToolBar_Text '��ť����
            Call Menu_View_ToolBar_Text_click(Control)
        Case conMenu_View_ToolBar_Size '��ͼ��
            Call Menu_View_ToolBar_Size_click(Control)
            
        Case conMenu_View_StatusBar '״̬��
            Call Menu_View_StatusBar_click(Control)
            
        Case conMenu_View_Refresh 'ˢ��
            mblnIsCallModuleRefresh = True
            
            Call RefreshList
            Call RefreshPacsQueueData
                        
                        mblnIsCallModuleRefresh = False
        Case comMenu_Cap_Process
            Call Menu_Manage_�����ɼ�
'---------------------------����----------------
        Case conMenu_Tool_Valid         'ͼ��У�Թ���
            
            If Len(Dir(Mid(App.Path, 1, InStrRev(App.Path, "\")) & "zlPacsImageValid.exe")) > 0 Then
                If InitRegister Then
                    Shell Mid(App.Path, 1, InStrRev(App.Path, "\")) & "zlPacsImageValid.exe   " & gstrServerName & "||" & gstrUserName & "||" & gstrUserPswd & "||" & mlngCur����ID & "||" & mSysPar.lngImageValid & "||" & "" & "||1", 1
                End If
            End If
'--------------------------����-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum
            'Case Menu_Help_Web_Forum_click
        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click
        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click
        Case conMenu_Help_About
            Call Menu_Help_About_click
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse����, "|")) + 1
            Call Menu_Dept_Select(Control)
        Case conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99
            If Control.Parameter <> "" Then 'ִ�з�������ǰģ��ı���
        
                If mobjCurStudyInfo.lngAdviceId <> 0 Then
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                        "ִ�п���=" & mobjCurStudyInfo.lngExeDepartmentId, "ҽ��ID=" & mobjCurStudyInfo.lngAdviceId, "���ͺ�=" & mobjCurStudyInfo.lngSendNo, _
                            "NO=" & mobjCurStudyInfo.strNO, "����ID=" & mobjCurStudyInfo.lngPatId, "�Һŵ�=" & mobjCurStudyInfo.strRegNo)
                Else
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "", 1)
                End If
                
            End If
        '----------------------------------------�Զ����ѯ---------------------------------------
        Case conMenu_Manage_CustomQuery * 100# + 1 To conMenu_Manage_CustomQuery * 100# + 99
            Call ChangeScheme(Control.Caption, Val(Control.Parameter), True)
            
        Case Else
            If mobjCurStudyInfo.lngAdviceId = 0 Then
                mblnMenuDownState = False
                Exit Sub
            End If
            
            Select Case TabWindow.Selected.tag
                    
                    
                Case "�Ŷӽк�"
                    If Not mobjQueue Is Nothing Then
                        If mintChangeUserState = 2 Then  '�������û������������
                            MsgBoxD Me, "��ͳһ�û����ٲ�����", vbInformation, gstrSysName
                        Else
                            mobjQueue.zlExecuteCommandbar Control
                        End If
                    End If
                Case "�������", "סԺҽ��", "����ҽ��", "סԺ����", "���ﲡ��", "������Ӳ���", "סԺ���Ӳ���"
                    If Not mobjWork_His Is Nothing Then
                        Call mobjWork_His.zlMenu.zlExecuteMenu(Control.ID)
                    End If
                Case "������д"
                    If Not mobjWork_Report Is Nothing Then
                        Call mobjWork_Report.zlMenu.zlExecuteMenu(Control.ID)
                    End If
            End Select
            
    End Select
    
    mblnMenuDownState = False
Exit Sub
errhandle:
    mblnMenuDownState = False
        mblnIsCallModuleRefresh = False
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ShowPacsInterfaceCfg()
On Error GoTo ErrorHnad
    Dim lngCount As Long
         
    If Not CheckPopedom(mstrPrivs, "������ù���") Then
        Call MsgBox("��û�иò�����Ȩ�ޣ�����ϵ����Ա��", vbInformation, "��ʾ")
        Exit Sub
    End If
    
    If Not ChechHaveTlbinf32 Then
        Call MsgBox("ϵͳ��ȱ��TLBINF32.DLL�ļ������²�����ù��ܲ�������ʹ�ã�����ϵ���������Ա���(�����������ϵͳĿ¼����Ӳ�ע��TLBINF32.DLL�ļ�)��", vbInformation, "��ʾ")
        Exit Sub
    End If
    Call frmPacsInterfaceCfg.ShowPacsInterfaceCfg(Me, mlngModule, mstrPrivs, mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.lngPatId)
    
    Call ReCreatCbrMenu(cbrMain)
    
    Exit Sub
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Function ExecuteInterfaceFun(ByVal strVBS As String, ByVal lngExecuteType As Long, ByVal blnAutoDo As Boolean) As Boolean
'blnAutoDo �Ƿ��Զ�ִ�У�Ӱ���������ʾ��Ϣ����ʽ��
'����vbs�ű�ʵ�ֹ���
    Dim i As Integer
    Dim lngStart As Long, lngEnd As Long
    Dim ary() As String
    Dim strTmpVBS As String, strParaName As String, strParaVal As String
    Dim objCall As Object
    
On Error GoTo ErrorHnad
    
    ary = Split(strVBS, vbCrLf)
    
    For i = 0 To UBound(ary)
        '����Ԥ����������ڲ���ֵ
        strTmpVBS = ary(i)
        
        Do While InStr(strTmpVBS, "[[") > 0
            lngStart = InStr(strTmpVBS, "[[")
            lngEnd = InStr(strTmpVBS, "]]") + 2
            
            strParaName = Mid(strTmpVBS, lngStart, lngEnd - lngStart)
            
            Select Case strParaName
                Case "[[�û���]]"
                    strParaVal = UserInfo.����
                                
                Case "[[�˺���]]"
                    strParaVal = UserInfo.�û���
                    
                Case "[[ϵͳ��]]"
                    strParaVal = glngSys
                    
                Case "[[ģ���]]"
                    strParaVal = mlngModule
                
                Case "[[����ID]]"
                    strParaVal = mlngCur����ID
                
                Case "[[����ID]]"
                    strParaVal = mobjCurStudyInfo.lngPatId
                    
                Case "[[ҽ��ID]]"
                    strParaVal = mobjCurStudyInfo.lngAdviceId
                    
                Case "[[���ͺ�]]"
                    strParaVal = mobjCurStudyInfo.lngSendNo
                    
                Case "[[����]]"
                    strParaVal = mobjCurStudyInfo.strStudyNum
                    
                Case "[[�����]]", "[[סԺ��]]"
                    strParaVal = mobjCurStudyInfo.strMarkNum
                    
                Case "[[���֤��]]"
                    strParaVal = mobjCurStudyInfo.strIIDNumber
                    
                Case "[[Ӱ�����]]"
                    strParaVal = mobjCurStudyInfo.strImgType
                                        
                Case "[[��ǰ���ھ��]]"
                     strParaVal = Me.hwnd
                                         
                Case Else
                    strParaVal = "------"
                    
            End Select
            
            If strParaVal <> "------" Then strVBS = Replace(strVBS, strParaName, strParaVal)
            
            strTmpVBS = Trim(Mid(strTmpVBS, lngEnd))
        Loop
    Next
    
    If ExecuteSub(strVBS, lngExecuteType) = True Then ExecuteInterfaceFun = True
    
    ExecuteInterfaceFun = True
    
    Exit Function
ErrorHnad:
    If blnAutoDo Then
        err.Raise 0, , err.Description
    Else
        MsgBox err.Description, vbExclamation, gstrSysName
    End If
    ExecuteInterfaceFun = False
End Function

Private Function ExecuteSub(ByVal strVBS As String, ByVal lngExecuteType As Long, Optional ByVal blnCheckVBS As Boolean = False) As Boolean
'����vbs�ű�ʵ�ֹ���
    Dim objCall As Object
    Dim strTempVBS As String
    
On Error GoTo ErrorHnad
    
    ExecuteSub = False
    
    '�����ű�ִ�ж���
    Set objCall = CreateObject("ScriptControl")
    objCall.TimeOut = 60000
    objCall.Language = "vbscript"
    
    Call objCall.AddCode(strVBS)
    
    If blnCheckVBS Then ExecuteSub = True: Exit Function
    
    Call objCall.Run(Trim("ExcuteSub"))
    
    Exit Function
ErrorHnad:
    err.Raise 0, , err.Description
End Function

Private Sub RefreshPacsQueueData(Optional blnSetFocus As Boolean = True)
'ˢ���Ŷ�ģ������
    If mSysPar.blnUseQueue And Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Call mobjQueue.zlRefreshQueueData(GetSelQueueRooms(), blnSetFocus)
    End If
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
    
    '���������С
    gbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, IIf(bytSize = 2, 15, bytSize)))
    
    Call ReSetFormFontSize
    Call ReSetModuleFontSize(gbytFontSize, IIf(bytSize = 2, 1, bytSize))
End Sub


Private Sub ReSetModuleFontSize(ByVal bytFontSize As Byte, ByVal bytSize As Byte)
'����:�������ø���ҵ��ģ�鴰��������С
    On Error Resume Next
        
        '�����ֺŴ�С����98496
    If Not mobjWork_Report Is Nothing Then
        Call mobjWork_Report.SetFontSize(gbytFontSize)
    End If

    '�ж� ��ǰѡ�е�
    Select Case mlngModule
        Case 1290
            If Not mfrmWork_PacsImg Is Nothing Then
                If TabWindow.Selected.tag = "Ӱ��ͼ��" Then
                    Call mfrmWork_PacsImg.ReSetFormFontSize(gbytFontSize)
                End If
            End If
            
            If Not mobjWork_His Is Nothing Then
                If Not mobjWork_His.GetExpenseObj Is Nothing Then Call mobjWork_His.GetExpenseObj.SetFontSize(bytSize)
                If Not mobjWork_His.GetAdviceObj Is Nothing Then Call mobjWork_His.GetAdviceObj.SetFontSize(bytSize)
                If Not mobjWork_His.GetEPRsObj Is Nothing Then Call mobjWork_His.GetEPRsObj.SetFontSize(bytSize)
            End If
            
        Case 1291
            If Not mobjWork_His Is Nothing Then
               If Not mobjWork_His.GetExpenseObj Is Nothing Then Call mobjWork_His.GetExpenseObj.SetFontSize(bytSize)
               If Not mobjWork_His.GetAdviceObj Is Nothing Then Call mobjWork_His.GetAdviceObj.SetFontSize(bytSize)
               If Not mobjWork_His.GetEPRsObj Is Nothing Then Call mobjWork_His.GetEPRsObj.SetFontSize(bytSize)
            End If
                        
            If Not mobjWork_ImageCap Is Nothing Then
                Call mobjWork_ImageCap.SetFontSize(gbytFontSize)
            End If
            
        Case 1294
        
            If Not mobjWork_Pathol Is Nothing Then
                Select Case TabWindow.Selected.tag
                    Case "�걾����"
                        Call mobjWork_Pathol.GetModule(mtSpecimen).ReSetFormFontSize(gbytFontSize)
                        
                    Case "����ȡ��"
                        Call mobjWork_Pathol.GetModule(mtMaterial).ReSetFormFontSize(gbytFontSize)
                        
                    Case "������Ƭ"
                        Call mobjWork_Pathol.GetModule(mtSlices).ReSetFormFontSize(gbytFontSize)
                        
                        
                    Case "�����ؼ�"
                        Call mobjWork_Pathol.GetModule(mtSpeExam).ReSetFormFontSize(gbytFontSize)
                        
                    Case "���̱���"
                        Call mobjWork_Pathol.GetModule(mtProRep).ReSetFormFontSize(gbytFontSize)
                        
                    Case "�������"
                        If Not mobjWork_His Is Nothing Then Call mobjWork_His.GetExpenseObj.SetFontSize(gbytFontSize, bytSize)
                        
                    Case "����ҽ��", "סԺҽ��"
                        If Not mobjWork_His Is Nothing Then Call mobjWork_His.GetAdviceObj.SetFontSize(bytSize)
                    
                End Select
            End If
    End Select
End Sub

Private Sub ReSetFormFontSize()
'����:�������ù���վ����������С
    On Error Resume Next
    
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim strFontType As String
    Dim i As Integer
    
    Me.FontSize = gbytFontSize
    Set CtlFont = New StdFont
    strFontType = IIf(IsUseClearType = True, "΢���ź�", "����")
    CtlFont.Name = strFontType
    
    If gblUsePacsQuery Then
        Call mobjPacsQueryWrap.CurPacsQuery.RefreshCfgFontSize(gbytFontSize)
        
        If Not mfrmHistory Is Nothing Then
            Call mfrmHistory.SetFontSize(gbytFontSize)
        End If
    End If
    
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("TabStrip") 'ҳ��ؼ�
            objCtrl.Font.Name = strFontType
            objCtrl.Font.Size = gbytFontSize
        Case UCase("Label")
            If objCtrl.Name = "LabFlag����" Or objCtrl.Name = "LabFlagӤ��" Or objCtrl.Name = "LabFlag��ɫͨ��" _
                Or objCtrl.Name = "LabFlagΣ��״̬" Or objCtrl.Name = "LabFlag��Ⱦ��״̬" _
                Or objCtrl.Name = "labNoScheme" Or objCtrl.Name = "LabFlag����" Then
            ElseIf objCtrl.Name = "labCollectionInfo" Then
                objCtrl.Font.Name = strFontType
                objCtrl.Font.Bold = False
                objCtrl.FontSize = gbytFontSize
            Else
                objCtrl.Font.Name = strFontType
                objCtrl.FontSize = gbytFontSize
                objCtrl.Height = TextHeight("��") + 60
            End If
        Case UCase("vsFlexGrid")
        
            Dim lngRow As Long
            
            objCtrl.Cell(flexcpFontSize, 0, 0, objCtrl.Rows - 1, objCtrl.Cols - 1) = gbytFontSize
            objCtrl.HeadFont.Size = gbytFontSize
            objCtrl.FontSize = gbytFontSize
            objCtrl.RowHeight(0) = TextHeight("��") + 150
            '��������к��޸ĵ�һ�еĿ��
            If objCtrl.Rows < 11 Then
                objCtrl.ColWidth(0) = TextWidth("XX")
            ElseIf 10 < objCtrl.Rows And objCtrl.Rows < 101 Then
                objCtrl.ColWidth(0) = TextWidth("XXX")
            ElseIf 100 < objCtrl.Rows And objCtrl.Rows < 1001 Then
                objCtrl.ColWidth(0) = TextWidth("XXXX")
            Else
                objCtrl.ColWidth(0) = TextWidth("XXXXX")
            End If
            
            If objCtrl.Rows - 1 = objCtrl.BottomRow Then
                lngRow = objCtrl.BottomRow
            Else
                If objCtrl.Rows - objCtrl.BottomRow > 30 Then
                    lngRow = objCtrl.BottomRow + 29
                Else
                    lngRow = objCtrl.Rows - 1
                End If
            End If
            
            For i = objCtrl.TopRow To lngRow
                objCtrl.RowHeight(i) = TextHeight("��") + 120
            Next
            
        Case UCase("ComboBox")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = gbytFontSize
        Case UCase("OptionButton")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = gbytFontSize
            objCtrl.Width = TextWidth("�޹�" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = gbytFontSize
            objCtrl.Width = TextWidth("�޹�" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Name = strFontType
            objCtrl.Font.Size = gbytFontSize
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("��") * 1.5
        Case UCase("textBox")
          objCtrl.FontName = strFontType
          objCtrl.FontSize = gbytFontSize
        Case UCase("ReportControl")
            CtlFont.Size = gbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
            
        Case UCase("DockingPane")
            CtlFont.Size = gbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            Call dkpMain.RedrawPanes

        Case UCase("CommandBars")
            CtlFont.Size = gbytFontSize
            Set objCtrl.options.Font = CtlFont

        Case UCase("TabControl")
            CtlFont.Size = gbytFontSize
            Set objCtrl.PaintManager.Font = CtlFont

        Case UCase("CommandButton")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = gbytFontSize

        Case UCase("PatiIdentify")
            objCtrl.CardNoShowFont.Size = gbytFontSize
            objCtrl.Font.Size = gbytFontSize
            objCtrl.IDKindFont.Size = gbytFontSize
            If gbytFontSize = 9 Then
                objCtrl.Height = 330
            ElseIf gbytFontSize = 12 Then
                objCtrl.Height = 360
            ElseIf gbytFontSize = 15 Then
                objCtrl.Height = 390
            End If
            objCtrl.Refrash
            
        Case UCase("richtextbox")
            If rtxtAppend.Text <> "" Then
                Call mobjPacsQueryWrap.SetRichtxtFontSize
            End If
        End Select
    Next
    
End Sub

Private Sub cbrMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
On Error GoTo errhandle
    Dim objControl As CommandBarControl, i As Integer
    Dim aryKindInfo() As String
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
        Case conMenu_View_Filter * 10#
            With CommandBar.Controls
                If .Count = 0 Then
                    If mlngModule = G_LNG_PACSSTATION_MODULE Then
                        'ֻ��ҽ����Ҫ��ӡ�ȫ�����ҡ��Ŀ���ѡ��˵�
                        Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100#, "ȫ������")
                    
                        objControl.Category = "Main"
                        objControl.DescriptionText = 0
                        If mblnAllDepts = True Then objControl.Checked = True
                    End If
                    
                    '�����ÿһ���������
                    For i = 0 To UBound(Split(mstrCanUse����, "|"))  'mstrCanUse����=id_����-����|id_����-����
                        Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100# + i + 1, Split(Split(mstrCanUse����, "|")(i), "_")(1) & "(&" & i & ")")
                        objControl.Category = "Main"
                        objControl.DescriptionText = Split(Split(mstrCanUse����, "|")(i), "_")(0)
                        
                        If mblnAllDepts = False And mlngCur����ID = objControl.DescriptionText Then
                            objControl.Checked = True
                        End If
                    Next
                End If
            End With
        Case Else
            Select Case Me.TabWindow.Selected.tag
                Case "סԺҽ��", "����ҽ��", "�������"
                    Call mobjWork_His.zlMenu.zlRefreshSubMenu(CommandBar)
            End Select
    End Select
errhandle:
End Sub


Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errhandle
    Dim blnNoRecord As Boolean
    Dim intState As Integer
    Dim strTmp As String
    Dim blnCancel As Boolean
    Dim tt As CommandBarControl
    Dim objControl As XtremeCommandBars.ICommandBarControl
    
    If Not mblnInitOk Then Exit Sub

    '����ò˵�Ϊ���Ӳ����༭�����Ҽ��˵�������Ҫ�޸Ĳ˵�id����Ϣ
    Set objControl = cbrMain.FindControl(, Control.ID, True, True)
    If objControl Is Nothing Then
        If Not mobjWork_Report Is Nothing Then
            Call mobjWork_Report.ReplacePopupMenu(Control)
        End If
    End If
    
    blnNoRecord = True
    If vsfList.Cols <= 1 Or vsfList.Rows <= 1 Or vsfList.RowSel < 1 Then
        blnNoRecord = True
    Else
        blnNoRecord = mobjCurStudyInfo.lngAdviceId = 0
    End If
    
    If Not blnNoRecord Then
        intState = mobjCurStudyInfo.intStep   'ִ�й���
        blnCancel = mobjCurStudyInfo.strStuStateDesc = "�Ѿܾ�"
    End If
    
    If TabWindow.ItemCount > 0 Then
        If TabWindow.Selected Is Nothing Then Exit Sub
        
        '����Ӱ��ͼ��˵�
        If Not mfrmWork_PacsImg Is Nothing Then
            If mfrmWork_PacsImg.zlMenu.zlIsModuleMenu(Control) Then
                Call mfrmWork_PacsImg.zlMenu.zlUpdateMenu(Control)
                Exit Sub
            End If
        End If
        
        '���²�����˵�
        If Not mobjWork_Pathol Is Nothing Then
            If mobjWork_Pathol.zlMenu.zlIsModuleMenu(Control) Then

                Select Case Control.ID
                    Case conMenu_PatholSpecimen
                        Control.Visible = IIf(TabWindow.Selected.tag = "�걾����", True, False)
                        
                        Exit Sub
                    Case conMenu_PatholMaterial
                        Control.Visible = IIf(TabWindow.Selected.tag = "����ȡ��", True, False)
                        
                        Exit Sub
                    Case conMenu_PatholSlices
                        Control.Visible = IIf(TabWindow.Selected.tag = "������Ƭ", True, False)
                        
                        Exit Sub
                    Case conMenu_PatholSpeExam
                        Control.Visible = IIf(TabWindow.Selected.tag = "�����ؼ�", True, False)
                        
                        Exit Sub
                    Case conMenu_PatholProRep
                        Control.Visible = IIf(TabWindow.Selected.tag = "���̱���", True, False)
                        
                        Exit Sub
                End Select
                
                Call mobjWork_Pathol.zlMenu.zlUpdateMenu(Control)
                
                Exit Sub
            End If
        End If
        
        '����HISģ��˵�
        If Not mobjWork_His Is Nothing Then
            
            If InStr("�������, סԺҽ��, ����ҽ��, סԺ����, ���ﲡ��, ������Ӳ���, סԺ���Ӳ���", TabWindow.Selected.tag) > 0 Then
                If mobjWork_His.zlMenu.zlIsModuleMenu(Control) Then
                    Call mobjWork_His.zlMenu.zlUpdateMenu(Control)
                    
                    '����ɳ�����,�Լ�ҽ���б���鿴��ӡ����Ƭ�˵����������
                    If mobjCurStudyInfo.intStep = 6 Then
                        Select Case Control.ID
                            Case conMenu_Edit_MarkMap, conMenu_Tool_PlugIn, conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99, conMenu_Edit_Compend, conMenu_Manage_ReportLisView, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3
                                Control.Enabled = True
                            Case conMenu_Edit_Copy, conMenu_File_ExportToXML, conMenu_Tool_Search, conMenu_File_Open, conMenu_EditPopup, conMenu_Edit_ChargeDelAudit
                                '�⼸���˵�������
                            Case Else
                                Control.Enabled = False
                        End Select
                    End If
                    
                    Exit Sub
                End If
            End If
        End If
        
        If Not mobjWork_ImageCap Is Nothing Then
'                If mobjWork_ImageCap.zlMenu.zlIsModuleMenu(control) Then
'                    '������Ƶ�ɼ��˵�...
'                    Call mobjWork_ImageCap.zlMenu.zlUpdateMenu(control)
'                    Exit Sub
'                End If
        End If

        
        '���±���ģ��˵�
        If Not mobjWork_Report Is Nothing Then
            If mobjWork_Report.zlMenu.zlIsModuleMenu(Control) Then
                Call mobjWork_Report.zlMenu.zlUpdateMenu(Control)
                
                '��ǰ�鿴�������μ�¼��˵���������  LSQ���μ��
'                If cboTimes.ListIndex <> -1 Then
'                    If mobjCurStudyInfo.lngAdviceID <> cboTimes.ItemData(cboTimes.ListIndex) Then
'                        If Control.ID = conMenu_Edit_Copy + 1000000 Or Control.ID = conMenu_File_ExportToXML + 1000000 Or Control.ID = conMenu_EditPopup + 1000000 _
'                            Or Control.ID = conMenu_Tool_Search + 1000000 Or Control.ID = conMenu_File_Preview + 1000000 Or Control.ID = conMenu_File_Print + 1000000 Or Control.ID = conMenu_File_NoAskPrint + 1000000 Then
'                            '�⼸���˵�������
'                        Else
'                            Control.Enabled = False
'                        End If
'                    End If
'                End If
            
                Exit Sub
            End If
        End If
    End If
    
    
    Select Case Control.ID
        Case conMenu_Manage_LocateValue
            Control.Enabled = Not blnNoRecord
        Case comMenu_Cap_Process
            Control.Enabled = True 'Not blnNoRecord
        Case conMenu_View_Filter * 10#
            Control.Caption = "��ǰ����:" & IIf(mblnAllDepts = True, "ȫ������", mstrCur����)
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse����, "|")) + 1
            If mblnAllDepts = True Then
                Control.Checked = (Control.DescriptionText = 0)
            Else
                Control.Checked = (Control.DescriptionText = mlngCur����ID)
            End If
        Case conMenu_View_ToolBar_Button '������
            If cbrMain.Count >= 2 Then
                Control.Checked = Me.cbrMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text 'ͼ������
            If cbrMain.Count >= 2 Then
                Control.Checked = Not (Me.cbrMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '��ͼ��
            Control.Checked = Me.cbrMain.options.LargeIcons
        Case conMenu_View_StatusBar '״̬��
            Control.Checked = Me.stbThis.Visible
        Case conMenu_View_Filter   '����
        
        Case conMenu_View_Refresh  'ˢ��
        
        Case conMenu_Check_ViewLink
            Control.Enabled = Not blnNoRecord
        
        Case conMenu_Manage_RequestPrint
            Control.Enabled = Control.CommandBar.Controls.Count > 0 And Not blnNoRecord
            
        Case conMenu_Manage_Regist   '���Ǽ�(&I)
        Case conMenu_Manage_CopyCheck '���ƵǼ�
            If Not blnNoRecord Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
        Case conMenu_Manage_Redo   'ȡ���Ǽ�(&R)
            If Not blnNoRecord Then
                Control.Enabled = intState <= 1 And intState <> -1 And Not blnCancel
            Else
                Control.Enabled = False
            End If
        Case conMenu_Manage_ReGet   '�ٻ�ȡ��
            If Not blnNoRecord Then
                Control.Enabled = blnCancel
            Else
                Control.Enabled = False
            End If
        Case conMenu_Manage_ThingModi   '�޸���Ϣ(&M)
            If Not blnNoRecord Then
                Control.Enabled = intState < 6 And Not blnCancel
            Else
                Control.Enabled = False
            End If
        Case conMenu_Manage_CheckList   '�鿴���뵥
            Control.Visible = True
            If mobjCurStudyInfo.lngAdviceId > 0 And mobjCurStudyInfo.lngPatientFrom <> 3 Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
            
        Case conMenu_Manage_ExecOnePart     '�ֲ�λִ��
            If Not blnNoRecord Then
                '2, "�ѱ���", 3, "�Ѽ��", 4, "�ѱ���", 5, "�����"
                Control.Enabled = (intState >= 2 And intState <= 5) And Not blnCancel
            Else
                Control.Enabled = False
            End If
            
        Case conMenu_Manage_Disease, conMenu_Manage_DiseaseQuery, conMenu_Manage_DiseaseRegist
            If mstrPublicAdvicePrivs = "-1" Then mstrPublicAdvicePrivs = ";" & GetPrivFunc(100, 9001) & ";"
            
            If Control.ID = conMenu_Manage_Disease Then
                Control.Visible = InStr(mstrPublicAdvicePrivs, "��Ⱦ�����Խ���Ǽ�") > 0 Or InStr(mstrPublicAdvicePrivs, "��Ⱦ�����Խ����ѯ") > 0
                Control.Enabled = mobjCurStudyInfo.lngAdviceId > 0
            ElseIf Control.ID = conMenu_Manage_DiseaseQuery Then
                Control.Visible = InStr(mstrPublicAdvicePrivs, "��Ⱦ�����Խ����ѯ") > 0
                Control.Enabled = mobjCurStudyInfo.lngAdviceId > 0
            Else
                Control.Visible = InStr(mstrPublicAdvicePrivs, "��Ⱦ�����Խ���Ǽ�") > 0
                Control.Enabled = mobjCurStudyInfo.lngAdviceId > 0 And intState >= 4
            End If
        Case conMenu_Manage_ModifBaseInfo '������Ϣ����
            If Not blnNoRecord Then
                Control.Enabled = intState < 6 And Not blnCancel And mobjCurStudyInfo.lngPatientFrom <= 2 And mobjCurStudyInfo.lngBaby <= 0
            Else
                Control.Enabled = False
            End If
        Case conMenu_Manage_Receive   '��鱨��(&L)
            Control.Enabled = Not Control.Enabled
            Control.Enabled = Not Control.Enabled
            If Not blnNoRecord Then
                Control.Enabled = intState <= 1 And intState <> -1 And Not blnCancel
            Else
                Control.Enabled = False
            End If
        
        Case conMenu_Manage_Logout   'ȡ������(&D)
            If blnNoRecord Then
                Control.Enabled = False
            ElseIf Control.Parent Is Nothing Then '��ʹ���ȼ�ʱ��������ж�parent����������쳣
                Exit Sub
            ElseIf Control.Parent.type = xtpControlPopup Then
                Control.ToolTipText = "ȡ������"
                Control.Caption = "ȡ������(&D)"
                Control.Enabled = (intState = 2 Or intState = 3)

            Else ' �������е���ȡ��������ȡ���Ǽ�,ͬһ�������ȡ���ǼǺ�ȡ����鹦��
                Control.Enabled = (intState = 2 Or intState = 3) Or (intState <= 1 And intState <> -1 And Not blnCancel) '���ܾ��Ĳ��ܱ��ٴξܾ�
                Control.ToolTipText = IIf(intState <= 1 And intState <> -1, "ȡ���Ǽ�", "ȡ������")
                Control.Caption = "ȡ��"
            End If

            If Control.ToolTipText = "ȡ���Ǽ�" Then
                Control.Enabled = Control.Enabled And CheckPopedom(mstrPrivs, "���Ǽ�")
            Else
                Control.Enabled = Control.Enabled And CheckPopedom(mstrPrivs, "ȡ������")
            End If
        Case conMenu_Manage_InQueue    '�Ŷӽк����
            Control.Visible = mSysPar.blnUseQueue And Not mSysPar.blnAutoInQueue
            Control.Enabled = (intState >= 2 And intState <= 5)
            
        Case conMenu_Manage_Schedule                        '���ԤԼ
            If mblnIsScheduleDept = False Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Enabled = (intState = 0 Or intState = 1)
                If Control.Enabled = True Then
                    'ֻ��ԤԼ��Ŀ�����ܴ򿪼��ԤԼ
                    Control.Enabled = mblnIsScheduleOrder
                End If
            End If
            
        Case conMenu_Manage_ScheduleManage                  'ԤԼ����
                Control.Visible = mblnIsScheduleDept
                Control.Enabled = mblnIsScheduleDept
            
        Case conMenu_Manage_Transfer   '����Ӱ��(&C)
            Control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '��2---5֮�����
            
        Case conMenu_Manage_Cancel   'ȡ������(&B)
            If (intState >= 2 And intState <= 5) Or intState = -1 Then
                Control.Enabled = mobjCurStudyInfo.strStudyUID <> ""
            Else
                Control.Enabled = False
            End If
            
        Case conMenu_Manage_AttachMoney, conMenu_Manage_CompleteAttach
            Control.Enabled = intState >= 1 And intState < 6
            
        Case conMenu_Manage_Review  '���
            If (Not blnNoRecord And intState > 1 And intState <= 6) Or intState = -1 Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
        Case conMenu_Tool_Analyse   '�߼�ͼ����
            If (Not blnNoRecord And intState > 1 And intState < 6) Or intState = -1 Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
        Case conMenu_Manage_LookMecRecord '��������
            If mobjCurStudyInfo.lngPageID > 0 Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
        Case conMenu_Manage_Release     '���淢��,��������ɺ󶼿���ִ��
        
            Control.Enabled = IIf(intState >= 2, True, False)
            If Not blnNoRecord Then
              '�޸ı��淢�Ű�ť�ı���
                 If Not blnNoRecord Then
                     If mobjCurStudyInfo.intReportGiveOut = 1 And mobjCurStudyInfo.intFilmGiveOut = 1 Then
                         Control.Caption = "�ջ�"
                         Control.ToolTipText = "�ջ��Ѿ����ŵı����Ƭ"
                         Control.Enabled = Control.Enabled And CheckPopedom(mstrPrivs, "ȡ������")
                     Else
                         Control.Caption = "����"
                         Control.ToolTipText = IIf(Control.ID = conMenu_Manage_Release, "�����Ƭ����", "����ͽ�Ƭͬʱ����")
                     End If
                 End If
            End If

            
            Control.Enabled = Not Control.Enabled
            Control.Enabled = Not Control.Enabled
            
        Case conMenu_Manage_FilmRelease
            Control.Enabled = IIf(intState >= 2, True, False)
            
            If Not blnNoRecord Then
                If mobjCurStudyInfo.intFilmGiveOut = 1 Then
                    Control.Caption = "��Ƭ�ջ�"
                    Control.ToolTipText = "�ջ��Ѿ����ŵĽ�Ƭ"
                    Control.Enabled = Control.Enabled And CheckPopedom(mstrPrivs, "ȡ������")
                Else
                    Control.Caption = "��Ƭ����"
                    Control.ToolTipText = "��Ƭ����"
                    Control.Enabled = Control.Enabled And mobjCurStudyInfo.strStudyUID <> ""
                End If
            End If

        Case conMenu_Manage_ReportRelease   'LSQ527
            Control.Enabled = IIf(intState >= 4, True, False)
            
            If Not blnNoRecord Then
                If mrtReportType = �����ĵ��༭�� Then
                    If mobjWork_Report.GetReportReleaseState(mobjCurStudyInfo.lngAdviceId) > 1 Then
                        Control.Caption = "�����ջ�"
                        Control.ToolTipText = "�ջ��Ѿ����ŵı���"
                        Control.Enabled = Control.Enabled And CheckPopedom(mstrPrivs, "ȡ������")
                    Else
                        Control.Caption = "���淢��"
                        Control.ToolTipText = "���淢��"
                        Control.Enabled = Control.Enabled And CheckPopedom(mstrPrivs, "���淢��")
                    End If
                Else
                    If mobjCurStudyInfo.intReportGiveOut = 1 Then
                        Control.Caption = "�����ջ�"
                        Control.ToolTipText = "�ջ��Ѿ����ŵı���"
                        Control.Enabled = Control.Enabled And CheckPopedom(mstrPrivs, "ȡ������")
                    Else
                        Control.Caption = "���淢��"
                        Control.ToolTipText = "���淢��"
                        Control.Enabled = Control.Enabled And CheckPopedom(mstrPrivs, "���淢��")
                    End If
                End If
            End If
'            Control.Enabled = Not Control.Enabled
'            Control.Enabled = Not Control.Enabled
        
        Case conMenu_Manage_SendArrange                     '���Ͱ���
            Control.Enabled = IIf(intState >= 2 And intState < 6, True, False)
            
        Case conMenu_Manage_SendAudit               '�������
            Control.Enabled = IIf(intState = 4, True, False)
            
        Case conMenu_Manage_ReportExecutor      '����ִ��
            Control.Enabled = IIf(intState >= 2 And intState <= 6, True, False)
            
        Case conMenu_Manage_PacsCritical
            Control.Enabled = intState >= 2 Or intState = -1   '��2---6֮�����
            
        Case conMenu_Manage_PacsCriticalReg
            Control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1  '��2---5֮�����
            
        Case conMenu_Manage_PacsCriticalManage
            Control.Enabled = intState >= 2 Or intState = -1   '��2---6֮�����

        Case conMenu_Manage_Result, conMenu_Manage_Negative, conMenu_Manage_Positive   '���������(&X)
            If mSysPar.blnIgnoreResult = True Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '��2---5֮�����
                If mobjCurStudyInfo.intDangerState = 1 And Control.ID = conMenu_Manage_Result Then Control.Enabled = False
            End If
            
        Case conMenu_Manage_FuHe, conMenu_Manage_JiBenFuHe, conMenu_Manage_BuFuHe, conMenu_Manage_FuHeLevel '�������
            If mSysPar.lngConformDetermine = 0 Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '��2---5֮�����
            End If
        
        Case conMenu_Manage_GChannel, conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel '��ɫͨ�����/ȡ��
            Control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '��2---5֮�����

        Case conMenu_Manage_Complete   '������(&E)
            Control.Enabled = ((intState = 4 Or intState = 5) Or ((intState = 2 Or intState = 3) And (mSysPar.blnNoSignFinish)))

        Case conMenu_Manage_Undone   'ȡ�����(&U)
            Control.Enabled = intState = 6

        Case conMenu_File_SendImg  '����ͼ��
        Case conMenu_Img_Contrast, conMenu_Img_Look     'Ӱ��Ա�,Ӱ���Ƭ
            If mblnObserve Then
                If blnNoRecord Then Control.Enabled = False: Exit Sub

                Control.Enabled = mobjCurStudyInfo.strStudyUID <> ""
            Else
                Control.Visible = False
            End If
        Case conMenu_Manage_RelatingPatiet  '��������
            If blnNoRecord Or (intState < 2 And intState <> -1) Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_Manage_ReLoadPDF
            Control.Enabled = Not blnNoRecord And (mobjCurStudyInfo.blnIsPrinted Or intState > 4)
            Control.Visible = mSysPar.blnPDFTested And mSysPar.strPDFFTPdevice <> "" And mSysPar.strPDFFTPdevice <> M_STR_PDF_NOPRINTER
            
        Case conMenu_Manage_Burn
        Case conMenu_File_SendImg
        Case conMenu_File_PrintSet     '��ӡ����(&S)
        Case conMenu_File_Excel         '�嵥��ӡ(&L)
            Control.Enabled = Not blnNoRecord
        Case conMenu_File_Parameter, conMenu_Cap_DevSet
        
        Case conMenu_Manage_ChangeUser  '�û�����
            If mSysPar.blnChangeUser Then
                Control.Visible = True
            Else
                Control.Visible = False
            End If
            
        Case conMenu_Manage_SwitchUser  '�л��û�
            If mSysPar.blnSwitchUser Then
                Control.Visible = True
            Else
                Control.Visible = False
            End If
        
'        Case conMenu_Manage_SetXWParam      '����PACS�������ã�����д˲˵�������ʾ
        Case conMenu_ReportPopup, conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99 '����
        Case conMenu_FilePopup, conMenu_ManagePopup, conMenu_ViewPopup, conMenu_HelpPopup
        Case conMenu_ToolPopup, conMenu_Tool_Valid
        Case conMenu_Help_Help, conMenu_Help_About  '����
        Case conMenu_Help_Web, conMenu_Help_Web_Forum, conMenu_Help_Web_Home, conMenu_Help_Web_Mail '����WEB
        Case conMenu_File_Exit
        Case ConMenu_File_ShortcutSet
        Case conMenu_Pathol_WorkModule
        Case conMenu_View_ToolBar
        Case conMenu_Manage_Query
        Case conMenu_Manage_QueryCFG
        Case conMenu_Manage_QueryCfgUserScheme
            Control.Enabled = IIf(mlngCur����ID = 0, False, True)
        Case conMenu_Manage_QueryTabDisplayScheme
            Control.Checked = mSysPar.blnQuickTabDisplayScheme
        Case conMenu_Manage_QueryValidTime
            Control.Checked = mSysPar.blnQueryValidTime
        Case conMenu_Manage_PacsPlugIn, conMenu_Manage_PacsPlugCfg
        Case conMenu_Manage_PacsPlugIn * 10000# To conMenu_Manage_PacsPlugIn * 10000# + 100
            '100908             Category������չΪ3��
            'strTmp:����Ƿ�����
            strTmp = IIf(UBound(Split(Control.Category, ",")) = 2, Split(Control.Category, ",")(0), Control.Category)
            Control.Enabled = Val(strTmp)
        Case conMenu_Manage_PacsPlugLevel2 * 10000# To conMenu_Manage_PacsPlugLevel2 * 10000# + 9999
        Case conMenu_Cap_DevSet     'Ӱ���豸����
        Case conMenu_Manage_Change_In   '�����б�
        Case conMenu_Img_3D_MMPR, conMenu_Img_3D_MPR, conMenu_Img_3D_PF, conMenu_Img_3D_SA, conMenu_Img_3D_VA, conMenu_Img_3D_VE '��ά�ؽ��ļ����Ӳ˵�����Ҫ����
        Case conMenu_View_FontSize_S    'С����
             Control.Checked = gbytFontSize = 9
        Case conMenu_View_FontSize_M    '������
             Control.Checked = gbytFontSize = 12
        Case conMenu_View_FontSize_L    '������
             Control.Checked = gbytFontSize = 15
        
   '-------------------------------------------------�ղع�����----------------------------------------------------------
 
        Case conMenu_Collection    '�ղ�(&C)
            Control.Enabled = True
        Case conMenu_Collection_Manage  '�ղع���˵�
            Control.Enabled = True
        Case conMenu_Collection_ViewShare      '�鿴����
            Control.Enabled = True
        Case comMenu_Collection_Type * 10000# To comMenu_Collection_Type * 10000# + 9999  '��̬�ղز˵�
            Control.Enabled = True
        Case conMenu_Collection_ViewShare * 10000# To conMenu_Collection_ViewShare * 10000# + 9999  '��̬����˵�
            Control.Enabled = True
         Case conMenu_Collection_To
            
            
    '-------------------------------------------ɨ�����뵥����-----------------------------------------------

        'ɨ�����뵥
        Case comMenu_Petition_Capture
            If blnCancel Then
                Control.Enabled = False
            Else
                Control.Enabled = IIf((intState >= 2 And intState <= 5) Or intState = -1, True, False)
            End If
            
        '�鿴���뵥
        Case comMenu_Petition_View
            
        Case conMenu_Manage_CustomQuery * 100# + 1 To conMenu_Manage_CustomQuery * 100# + 99
            Control.Enabled = True

            If Control.Parameter = mobjPacsQueryWrap.SchemeNo Then
                Control.iconid = 3558
            Else
                Control.iconid = 0
            End If
            
        Case conMenu_Manage_CustomQuery * 100 + 500
        
        Case Else
            If Control.Caption = "Toolbar Options" Or Control.Caption = "������ѡ��" Then
                Control.Enabled = True
                Exit Sub
            End If
            
            If blnNoRecord Then
                Control.Enabled = False
                Exit Sub
            End If
                    
            
            '����ɳ�����,�Լ�ҽ���б���鿴��ӡ����Ƭ�˵����������
            If mobjCurStudyInfo.intStep = 6 Then
                Control.Enabled = False
            End If
            
    End Select
errhandle:
End Sub

Private Sub InitModuleParameter()
'����:��ʼ��ģ�鼶����,���������ʱ����һ��
On Error GoTo errH
    Dim rsTemp As ADODB.Recordset
    
    mSysPar.lngListColorMark = nvl(GetDeptPara(mlngCur����ID, "��ɫ��ʾ����", 0))
    mSysPar.blnNameColColorCfg = GetDeptPara(mlngCur����ID, "������ɫ����", 0) = "1"         '������ɫ����
    mSysPar.blnOrdinaryNameColColorCfg = GetDeptPara(mlngCur����ID, "ȱʡ���Ͳ���������ɫ����", 0) = "1"       'ȱʡ���Ͳ���������ɫ����
    
    If mSysPar.blnNameColColorCfg Then
        gstrSQL = "select ���� from �������� where ȱʡ��־=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡȱʡ��������")
        
        If rsTemp.RecordCount > 0 Then mstrDefaultPatientType = nvl(rsTemp!����)
    End If
    mSysPar.blnAutoPrint = Val(zlDatabase.GetPara("�������Զ���ӡ���뵥", glngSys, mlngModule, 0)) '�������Զ���ӡ���뵥
    mSysPar.blnAutoPrintCheck = Val(zlDatabase.GetPara("�Զ�����ظ������ӡ", glngSys, mlngModule, 0))
    
    mSysPar.blnChangeUser = GetDeptPara(mlngCur����ID, "�������û�", 0) = "1"              '�������û�
    mSysPar.blnSwitchUser = GetDeptPara(mlngCur����ID, "�����л��û�", 0) = "1"              '�����л��û�
    
    mSysPar.blnIsPetitionScan = IIf(Val(GetDeptPara(mlngCur����ID, "�������뵥ɨ��", 1)) = 1, True, False)   '��ȡ�������뵥ɨ�����
    mSysPar.strImageLevel = nvl(GetDeptPara(mlngCur����ID, "Ӱ�������ȼ�", "��,��"))
    mSysPar.strReportLevel = nvl(GetDeptPara(mlngCur����ID, "���������ȼ�", "��,��"))
    mSysPar.blnֱ�Ӽ�� = (Val(GetDeptPara(mlngCur����ID, "�ǼǺ�ֱ�Ӽ��", 0)) = 1)         '�ǼǺ�ֱ�Ӽ��

'    mSysPar.lngCriticalValues = Val(GetDeptPara(mlngCur����ID, "Σ������ж�", 0))           'Σ������ж�
    mSysPar.blnIgnoreResult = GetDeptPara(mlngCur����ID, "���Խ��������", 0) = "1" '        '���Խ��������
    mSysPar.lngConformDetermine = Val(GetDeptPara(mlngCur����ID, "��������ж�", 0))         '��������ж�
    mSysPar.lngImageLevel = Val(GetDeptPara(mlngCur����ID, "Ӱ�������ж�", 0))               'Ӱ�������ж�
    mSysPar.lngReportLevel = Val(GetDeptPara(mlngCur����ID, "���������ж�", 0))
    
    mSysPar.lngHintType = Val(GetDeptPara(mlngCur����ID, "��Ͻ����ʾ����", 0))
    
    mSysPar.blnReportWithImage = GetDeptPara(mlngCur����ID, "��ͼ�����д����", 0) = "1" '   '��ͼ�����д����
    mSysPar.blnReportWithResult = GetDeptPara(mlngCur����ID, "��Ӱ�����Ϊ����", 0) = "1" '  '��Ӱ�����Ϊ����
    mSysPar.blnCompleteCommit = GetDeptPara(mlngCur����ID, "��˺�ֱ�����", 0) = "1" '      '��˺�ֱ�����
    mSysPar.blnFinallyCompleteCommit = GetDeptPara(mlngCur����ID, "�����ֱ�����", 0) = "1" '�����ֱ�����
    mSysPar.blnAuditAutoPrint = IIf(Val(GetDeptPara(mlngCur����ID, "�����ֱ�Ӵ�ӡ", 0)) = 1, True, False) '�����ֱ�Ӵ�ӡ
    mSysPar.blnNoSignFinish = GetDeptPara(mlngCur����ID, "����δǩ�������ӡ���", 0) = "1" '       '�ޱ���򱨸�δǩ���������
    mSysPar.strPDFFTPdevice = GetDeptPara(mlngCur����ID, "PDFת���洢�豸", "")
    mSysPar.blnPDFTested = False
    
    mSysPar.bln�����Ǽ� = Val(zlDatabase.GetPara("�����Ǽ�����", glngSys, mlngModule, 0)) '�����Ǽ�
    
    mSysPar.lngBeforeDays = Val(GetDeptPara(mlngCur����ID, "Ĭ�Ϲ�������", 2)) '                   'Ĭ�Ϲ�������
    If mSysPar.lngBeforeDays > 15 Or mSysPar.lngBeforeDays <= 0 Then
        mSysPar.lngBeforeDays = 2
    End If
    
    mSysPar.blnWriteCapDoctor = GetDeptPara(mlngCur����ID, "�ɼ�ͼ����Ϊ��鼼ʦ", 0) = "1"  '�ɼ�ͼ����Ϊ��鼼ʦ
    
    mSysPar.blnPrintCommit = GetDeptPara(mlngCur����ID, "��ӡ��ֱ�����", 0) = "1" '           '��ӡ��ֱ�����
    mSysPar.blnCanPrint = GetDeptPara(mlngCur����ID, "ƽ������˲��ܴ򱨸�") = "1"             'ƽ����Ҫ��˲��ܴ�ӡ =true
    mSysPar.blnAutoSendWorkList = GetDeptPara(mlngCur����ID, "����ʱ�Զ�����WorkList") = "1"   '����ʱ�Զ�����WorkList

    '����������
    mSysPar.blnNameFuzzySearch = GetDeptPara(mlngCur����ID, "����Ĭ��ģ����ѯ", "1") = "1"     '����Ĭ��ģ����ѯ
    mSysPar.blnNameQueryTimeLimit = GetDeptPara(mlngCur����ID, "������ѯʱ������", "1") = "1"  '����������ʱ�Ƿ����ʱ������
    
    '����ʱ��Ƭ
    mSysPar.blnShowImgAfterReport = (Val(zlDatabase.GetPara("����ʱ��Ƭ", glngSys, mlngModule, 0)) = 1)
    
    '�Ƿ�λ����
    mSysPar.blnIsLocateReport = Val(GetDeptPara(mlngCur����ID, "����л�ʱ��λ����༭", "1")) = 1
    
    '��첡�����ʱ���жϷ���
    mSysPar.blnPEISNoCheckMoneyFinish = (Val(zlDatabase.GetPara("��첡�����ʱ���жϷ���", glngSys, mlngModule, 0)) = 1)
    
    '��ʾ���÷�����ǩ
    mSysPar.blnQuickTabDisplayScheme = Val(zlDatabase.GetPara("��ʾ���÷�����ǩ", glngSys, mlngModule, 0)) = 1
    mSysPar.blnQueryValidTime = Val(zlDatabase.GetPara("�л���������ʱ�䷶Χ", glngSys, mlngModule, 0)) = 1
    
    If CheckPopedom(mstrPrivs, "�Ŷӽк�") And mlngModule <> G_LNG_PATHOLSYS_NUM And CheckPopedom(";" & GetPrivFunc(glngSys, 1160) & ";", "����") Then      '��Ȩ��ʹ�òŸ��ݲ�������
        mSysPar.blnUseQueue = GetDeptPara(mlngCur����ID, "�����Ŷӽк�", 0) = "1" '          'Ĭ�ϲ������Ŷӽк�
        
        If mSysPar.blnUseQueue Then
            mSysPar.blnSynStudylist = GetDeptPara(mlngCur����ID, "ͬ����λ����б�", 0)
            mSysPar.blnAutoInQueue = GetDeptPara(mlngCur����ID, "�������Զ��Ŷ�", 1)
        End If
    Else
        mSysPar.blnUseQueue = False
    End If
    
    mSysPar.blnRelatingPatient = GetDeptPara(mlngCur����ID, "������������", 0) = "1"       '�Ƿ�ʹ�ù�
    gblnXWLog = (Val(zlDatabase.GetPara("XW��¼�ӿ���־", glngSys, G_LNG_XWPACSVIEW_MODULE, "0")) = 1) '�Ƿ��¼�ӿ���־
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub TestPDF()
On Error GoTo errH
    Dim objPDF As clsPDF
    
    mSysPar.blnPDFTested = True
    
    If mrtReportType = ���Ӳ����༭�� Then
        If mSysPar.strPDFFTPdevice <> "" Then
            If Dir(App.Path & "\TmpImage\PDF\", vbDirectory) = "" Then
                    Call MkLocalDir(App.Path & "\TmpImage\PDF")
            End If
    
            Set objPDF = New clsPDF
            If objPDF.PDFInitialize() = False Then
                MsgBoxD Me, "����PDFת��������֤ʧ�ܣ�����ϵ������Աȷ�������ӡ�豸�Ƿ���ȷ��װ��", vbExclamation, gstrSysName
                mSysPar.strPDFFTPdevice = M_STR_PDF_NOPRINTER
            End If
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picList.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picWindow.hwnd
    End If
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
On Error GoTo errhandle
    '��ֹ����б� �϶�
    Cancel = IIf(((Action = 4 Or Action = 6 Or Action = 5) And Not Pane.hidden), True, False)
errhandle:
End Sub

Private Sub InitForm()
    Dim strKinds As String
    Dim blnDo As Boolean
    Dim lngKey As Long
    Dim bytFontSize As Byte
    Dim objPar As clsQueryPar
        
    Call WriteLog("InitForm -> Step 1����ʼִ��...", "frmPacsWork")
    
    '�õ����Ի�������
    blnDo = Val(zlDatabase.GetPara("ʹ�ø��Ի����")) <> 0
    
    mstrPrivs = gstrPrivs 'Ȩ��
    mlngModule = glngModul 'ģ���
    mlngCur����ID = 0
    mstrCur���� = ""
    mstrCanUse���� = ""
    mblnAllDepts = False
    
    '��ȡ�����С
    bytFontSize = Val(zlDatabase.GetPara("��ʾ�����С", glngSys, glngModul))
    gbytFontSize = IIf(bytFontSize = 0, 9, IIf(bytFontSize = 1, 12, 15))

    mblnInitOk = False  '��ʼ����,��ʼ�����֮ǰ���������ݵ���ȡ
    mblnMenuDownState = False
    mlngFilterTab = 0
    
    '�жϵ�ǰ�û��Ƿ���� ��Ƭվ�Ļ���Ȩ��
    mblnObserve = CheckPopedom(";" & GetPrivFunc(glngSys, 1289) & ";", "����")
    
    Call WriteLog("InitForm -> Step 2�����뱾��ע������...", "frmPacsWork")
    
'    '�жϵ�ǰ�û��Ƿ���С�Ӱ���豸Ŀ¼����Ȩ�ޣ��д�Ȩ�޲ſ�������������PACS����
'    mblnSetXWParam = IIf(InStr(GetPrivFunc(glngSys, G_LNG_XWPACSVIEW_MODULE), "PACS��������") > 0, True, False)
    
    Call InitLocalPars '����ע������
    
    Call WriteLog("InitForm -> Step 3�����벿�������Ϣ...", "frmPacsWork")
    If Not InitDepts Then Unload Me: Exit Sub '��ʼ��ҽ������
    
    mrtReportType = GetDeptPara(mlngCur����ID, "����༭��", 0)                 '����༭��
    
    ReDim gConnectedShardDir(0) As String   '��ʼ������Ŀ¼���Ӵ�
    
    '��ʼ�Ӵ���
    Set mobjEvent = New clsEvent
    Set gobjEvent = mobjEvent
    
    '���ݲ����ж��Ƿ�������Ϣ����
    Set mobjMsgCenter = New clsPacsMsgProcess
    Call mobjMsgCenter.OpenMsgCenter(mlngModule, mlngCur����ID, mstrPrivs)
    
    Set mobjPacsCore = New zl9PacsCore.clsViewer
    Set mfrmHistory = New frmHistoryStudy
    
    Call WriteLog("InitForm -> Step 4����ʼ���Զ����ѯ�������...", "frmPacsWork")
    If mobjPacsQueryWrap Is Nothing Then
        Set mobjPacsQueryWrap = New clsPacsQueryWrap
        Set objPar = New clsQueryPar
        Set objPar.cmdFind = cmdFind
        Set objPar.ImageList16 = img16
        Set objPar.ImageList24 = img24
        Set objPar.img1 = imgFun(0)
        Set objPar.img2 = imgFun(1)
        
        Set objPar.img3 = imgFun(2)
        Set objPar.img4 = imgFun(3)
        Set objPar.objFilterCmdBar = cbrFilter
        Set objPar.objPatiIdentify = PatiIdentify
        Set objPar.picContainer = picDataSearch
        Set objPar.cmdDo = cmdDo
        
        Set objPar.picFollow = PicFucs
        Set objPar.picList = picList
        Set objPar.rtpAppend = rtxtAppend
        Set objPar.TimerFunc = timFun
        Set objPar.vsfList = vsfList
        
        Set objPar.TabCtl = TabExtra
        Set objPar.rtfHisFollow = rftHistoryFollow
        Set objPar.PicHisFollow = PicFollowHistory
        Set objPar.TimerHisFunc = TimerHistory
        Set objPar.picTemp = picTemp
        
        Set objPar.labPatiInfo = labPatientInfo
        
        Call rftHistoryFollow.Move(50, 50, PicFollowHistory.Width - 100, PicFollowHistory.Height - 100)
                
        Call mobjPacsQueryWrap.Init(mlngCur����ID, UserInfo.ID, mlngModule, mrtReportType, mSysPar.blnCanPrint, mobjSquareCard, Me, objPar)
        
        mobjPacsQueryWrap.DefaultLocate = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\", "DEFLOCATE", True)
        
        cmdLocate.BackColor = IIf(mobjPacsQueryWrap.DefaultLocate, &HFF00&, &H8000000F)
        cmdFind.BackColor = IIf(mobjPacsQueryWrap.DefaultLocate = False, &HFF00&, &H8000000F)
    End If
    
    Call WriteLog("InitForm -> Step 5����ʼ�����ż�����...", "frmPacsWork")
    Call InitModuleParameter '��ʼ��ģ�鼶����
    
    Call WriteLog("InitForm -> Step 6����ʼ�����ڲ˵�...", "frmPacsWork")
    Call InitCommandBars

    Call WriteLog("InitForm -> Step 7����ʼ�����沼��...", "frmPacsWork")
    Call initTabExtra
    Call InitFaceScheme
    
     '���ע����й��������ֵΪ�� ���� �ѹ�ѡ���Ի����ã���ô��ע���д�빤������ʾģʽֵ
    If mintToolBarWriteReg = 9 Or (mintToolBarWriteReg = 0 And blnDo) Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\CommandBars", "cbrMainButtonText", 3
    End If
    
    '�ָ������״̬   ע���ָ�����״̬ ������� ��ע���д�빤������ʾģʽֵ �������棬�������ɹ�������ʾģʽ����
    Call RestoreWinState(Me, App.ProductName)
    
     '������--- �ı���ǩ ������ʹ��RestoreWinState �ָ����ˣ�����Ҫ����������δ��ѡ���Ի����ã��򹤾���Ĭ����ʾͼ����ı�
    If blnDo Then
        If Me.cbrMain(2).Controls(1).Style = xtpButtonIconAndCaption Then
            Me.cbrMain(2).ShowTextBelowIcons = True
        Else
            Me.cbrMain(2).ShowTextBelowIcons = False
        End If
    Else
        Me.cbrMain(2).ShowTextBelowIcons = True
    End If
    
    ClearCacheFolder App.Path & "\TmpImage\"    '����ʱĿ¼���ˣ�����ո�Ŀ¼
    
    
    '�ж���ʱĿ¼�Ƿ����
    If Dir(IIf(Len(App.Path) > 3, App.Path & "\", App.Path & "") & "TmpImage", vbDirectory) = "" Then
        Call MkDir(IIf(Len(App.Path) > 3, App.Path & "\", App.Path & "") & "TmpImage")
    End If
    
    
    '��ʼ��˫�û���½�Ĳ���
    mblnCnOracleIsHIS = True
    mintChangeUserState = 1
    mstrUserNameHIS = UserInfo.����
    mstrUserNameNew = UserInfo.����
    mstrUserIDHIS = UserInfo.�û���
    mstrUserIDNew = UserInfo.�û���
    
    Set mcnOracleHIS = gcnOracle
    
    Me.stbThis.Panels(4).Text = "����ҽ����" & mstrUserNameHIS & "   ���ҽ����" & mstrUserNameNew
    
    ReDim mobjPacsReportArry(0) As frmReport
    
    gblnUseXinWangView = False
    
    If mlngModule = G_LNG_PACSSTATION_MODULE Then
        gblnUseXinWangView = IsUseXwViewer
    
'        '�����RIS����վ���������������ݿ⣬��ȡ����
'        If gblnUseXinWangView Then
'            '���Ͻػ���Ϣ��hook
'            plngXWPreWndProc = XWHook(Me.hWnd)
'        End If
    End If

    
    mblnFormLoadState = True
    
    Call WriteLog("InitForm -> Step 10������ִ��...", "frmPacsWork")
End Sub

Private Function GetWindowCaption() As String
    GetWindowCaption = Mid(Me.Caption & " ", 1, InStr(Me.Caption & " ", " "))
End Function


Private Sub DisposeObj()
    TimerRefresh.Enabled = False
    
    If Not mfrmHistory Is Nothing Then
        Call mfrmHistory.Free
        Set mfrmHistory = Nothing
    End If
    
    If Not mobjPacsQueryWrap Is Nothing Then
        Call mobjPacsQueryWrap.Free
        Set mobjPacsQueryWrap = Nothing
    End If
    
    If Not mobjAppendBill Is Nothing Then
        Set mobjAppendBill = Nothing
    End If
    
    If Not mfrmWork_PacsImg Is Nothing Then
        Unload mfrmWork_PacsImg
        Set mfrmWork_PacsImg = Nothing
    End If
    
    If Not mobjQueue Is Nothing Then
        Unload mobjQueue
        Set mobjQueue = Nothing
    End If
    
    If Not mobjPacsCore Is Nothing Then
        mobjPacsCore.Closefrom
        Set mobjPacsCore = Nothing
    End If
    
    If Not mobjWork_Pathol Is Nothing Then
        Call mobjWork_Pathol.Free
        Set mobjWork_Pathol = Nothing
    End If
    
    If Not mobjWork_His Is Nothing Then
        Call mobjWork_His.Free
        Set mobjWork_His = Nothing
    End If
    
    If Not mobjWork_Report Is Nothing Then
        Call mobjWork_Report.Free
        Set mobjWork_Report = Nothing
    End If
    
    If mlngModule <> G_LNG_PACSSTATION_MODULE Then
        If Not mobjCaptureHot Is Nothing Then
            Call mobjCaptureHot.FreeHook
            Set mobjCaptureHot = Nothing
        End If
    End If
    
    'ʹ��Activex����Ƶ�ɼ���ʽ�˳�
    Set mobjWork_ImageCap = Nothing
    
    If Not gobjMsgCenter Is Nothing Then
        Set gobjMsgCenter = Nothing
    End If
        
    Set mobjEvent = Nothing
    Set mobjSquareCard = Nothing
    
    If Not mobjPublicAdvice Is Nothing Then Set mobjPublicAdvice = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    TimFlicker.Enabled = False
    
    Select Case mlngModule
        Case 1290
            Call UnAttachModuleMsgProc(Me.hwnd, mtImage)
        Case 1291
            Call UnAttachModuleMsgProc(Me.hwnd, mtVideo)
        Case 1294
            Call UnAttachModuleMsgProc(Me.hwnd, mtPathol)
    End Select
    
    
    If Not mobjPacsInterface Is Nothing Then Set mobjPacsInterface = Nothing
    
    If Not mobjWork_ImageCap Is Nothing Then
        Call mobjWork_ImageCap.zlNotifyQuit
    End If
    
    '�ر���Ϣ����
    If Not gobjMsgCenter Is Nothing Then
        Call gobjMsgCenter.CloseMsgCenter
    End If
    
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.���� & App.ProductName & "\", "��ʷ��鸡��������", PicFollowHistory.Width)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.���� & App.ProductName & "\", "��ʷ��鸡������߶�", PicFollowHistory.Height)
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.���� & App.ProductName & "\" & Me.Name & "\", "�б�����Ϣ�߶�����", mlngMove)
    
'    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsList), vsList.Name, mstrCol)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "StudyListWidth", picList.Width / Me.ScaleWidth)
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\", "DEFLOCATE", mobjPacsQueryWrap.DefaultLocate)
    '���������С
    zlDatabase.SetPara "��ʾ�����С", IIf(gbytFontSize = 9, 0, IIf(gbytFontSize = 12, 1, IIf(gbytFontSize = 15, 2, gbytFontSize))), glngSys, glngModul
    '�ָ���������
    Me.Caption = GetWindowCaption
    
    Call SaveWinState(Me, App.ProductName)
    
    Call DisposeObj
    
    '�ָ�����̨�����ݿ�����
    If mblnCnOracleIsHIS = False Then
        Set gcnOracle = mcnOracleHIS
        InitCommon gcnOracle
'        RegCheck
        SetDbUser mstrUserIDHIS
        Call GetUserInfo
        Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
    End If
    
    frmTwoUser.intDBState = 1
    
    mblnFormLoadState = False
    
    mblnIsValid = False
    Set mobjCurStudyInfo = Nothing
    Set mfrmPatholSpecimen = Nothing
    Set mobjHistoryStudyInfo = Nothing
    Set mclsCISKernel = Nothing
    Set mcnOracleHIS = Nothing
    Set mobjMedicalRecord = Nothing
    Set mfrmRISRequest = Nothing
    Set mobjMsgCenter = Nothing
    Set mobjPetitionCap = Nothing
    Set mobjPublicPACS = Nothing
End Sub

Private Function InitCardType(ByVal strCardNames As String) As String
'��ָ����ʽ��ʼ��������
    Dim i As Integer
    Dim aryKindInfo() As String
    Dim strKinds As String
    
    aryKindInfo = Split(strCardNames, ";")
    
    strKinds = ""
    For i = 0 To UBound(aryKindInfo) - 1
        If strKinds <> "" Then strKinds = strKinds & ";"
        strKinds = strKinds & aryKindInfo(i) & "|" & aryKindInfo(i) & "|-1"
    Next i
    
    InitCardType = strKinds & ";"
End Function

Private Sub InitLocalPars()
    Dim strTemp As String
    Dim strTempArry() As String
    Dim i As Integer
'��ʼ����ʱ���ز������Ը�������Ϊ��,������أ����ˣ��������õȵ���

    mstrCaptureHot = GetSetting("ZLSOFT", "����ģ��", "�ɼ��ȼ�", "F8")
    mstrCaptureAfterHot = GetSetting("ZLSOFT", "����ģ��", "��̨�ɼ��ȼ�", "F7")
    mstrCaptureAfterTagHot = GetSetting("ZLSOFT", "����ģ��", "��Ǹ����ȼ�", "F6")
    
    mlngMove = Val(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.���� & App.ProductName & "\" & Me.Name & "\", "�б�����Ϣ�߶�����", 0))
    
errContinue2:
    mSysPar.blnLockAfterCall = zlDatabase.GetPara("���к������ɼ�", glngSys, mlngModule, "0")
    mSysPar.strFirstTab = zlDatabase.GetPara("������ҳ", glngSys, mlngModule, "") 'Ϊ�ձ�ʾ��ʹ�ö��ƹ�����ҳ����
    mSysPar.blnAutoOpenReport = (Val(zlDatabase.GetPara("��ʼ����Զ��򿪱���", glngSys, mlngModule, 0)) = 1)
    mSysPar.blnChoosePrintFormat = (Val(zlDatabase.GetPara("������ӡǰѡ���ʽ", glngSys, mlngModule, 0)) = 1)
    mSysPar.strLocalRoom = zlDatabase.GetPara("����ִ�м�����", glngSys, mlngModule, "")
    mSysPar.blnQueueQuick = IIf(Val(zlDatabase.GetPara("�Զ�������ݺ��д���", glngSys, mlngModule, "1")) = 1, True, False)
    mSysPar.lngImageValid = Val(zlDatabase.GetPara("ͼ��У��", glngSys, mlngModule, 0))
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        
        mSysPar.blnPopChangGuiWindow = (Val(zlDatabase.GetPara("������������", glngSys, mlngModule, 0)) = 1)
        mSysPar.blnPopKuaiShuWindow = (Val(zlDatabase.GetPara("����ʯ����������", glngSys, mlngModule, 0)) = 1)
        mSysPar.blnPopBingDongWindow = (Val(zlDatabase.GetPara("������������", glngSys, mlngModule, 0)) = 1)
        mSysPar.blnPopXiBaoWindow = (Val(zlDatabase.GetPara("ϸ����������", glngSys, mlngModule, 0)) = 1)
        mSysPar.blnPopHuiZhenWindow = (Val(zlDatabase.GetPara("������������", glngSys, mlngModule, 0)) = 1)
        mSysPar.blnPopShiJianWindow = (Val(zlDatabase.GetPara("ʬ����������", glngSys, mlngModule, 0)) = 1)
    End If
    
    If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
        '����ǲɼ�ģ�飬����Ҫִ�иò���
        mSysPar.lngVideoStationMoneyExeModle = Val(zlDatabase.GetPara("�ɼ�����ִ��ģʽ", glngSys, mlngModule, 0))
    ElseIf mlngModule = G_LNG_PACSSTATION_MODULE Then
        mSysPar.lngPacsStationMoneyExeModle = Val(zlDatabase.GetPara("ҽ������ִ��ģʽ", glngSys, mlngModule, 0))
    Else
        mSysPar.lngPatholStationMoneyExeModle = Val(zlDatabase.GetPara("�������ִ��ģʽ", glngSys, mlngModule, 0))
    End If
    
    '�õ�ע����й��ڹ�������ʾ״̬��ֵ�����Ϊ�������9
    mintToolBarWriteReg = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\CommandBars", "cbrMainButtonText", 9))
    
End Sub

Private Function InitDepts() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
On Error GoTo errH
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str����IDs As String, str��Դ As String
    
    str��Դ = "1,2,3"
    If CheckPopedom(mstrPrivs, "���п���") Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where B.����ID = A.ID " & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " and (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null ) " & _
            " And instr([1],','||B.�������||',')> 0 And B.�������� IN('���')" & _
            " Order by A.����"
    Else
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID And C.��ԱID=" & UserInfo.ID & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " and (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null ) " & _
            " And instr([1],','||B.�������||',')>0  And B.�������� IN('���')" & _
            " Order by A.����"
    End If
   

    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, GetWindowCaption, CStr("," & str��Դ & ","))
    
    If rsTmp.EOF Then
        MsgBoxD Me, "û�з���ҽ��������Ϣ,���ȵ����Ź��������á�", vbInformation, gstrSysName
        Exit Function
    Else
        str����IDs = GetUser����IDs
        Do Until rsTmp.EOF
            mstrCanUse���� = mstrCanUse���� & "|" & rsTmp!ID & "_" & rsTmp!���� & "-" & rsTmp!����
            mstrCanUse����IDs = mstrCanUse����IDs & "," & rsTmp!ID
            
            If rsTmp!ID = UserInfo.����ID Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� '��ȡĬ�Ͽ���
            If InStr("," & str����IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur����ID = 0 Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� 'û��Ĭ�Ͽ���,ȡ���������ҵ�һ��
            rsTmp.MoveNext
        Loop
        
        mstrCanUse���� = Mid(mstrCanUse����, 2)
        mstrCanUse����IDs = Mid(mstrCanUse����IDs, 2)
        
        If CheckPopedom(mstrPrivs, "���п���") And mlngCur����ID = 0 Then
            mlngCur����ID = Split(Split(mstrCanUse����, "|")(0), "_")(0)
            mstrCur���� = Split(Split(mstrCanUse����, "|")(0), "_")(1)
        End If
        
        If mlngCur����ID = 0 And Not CheckPopedom(mstrPrivs, "���п���") Then  'û�����п��Ҳ���Ȩ��,���Ҳ����߿��Ҳ����ڼ�������
            MsgBoxD Me, "û�з�������������,����ʹ�ô˹���վ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        Call SetParaUseImgSignValid(mlngCur����ID)
        InitDepts = True
    End If
    
    
    If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
        glngXWDeptID = mlngCur����ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitFaceScheme()
    Dim lngListWidth As Double
    
    '��ʼ���沼��
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane
    With Me.dkpMain
        .SetCommandBars cbrMain
        .options.HideClient = True
        .options.UseSplitterTracker = False 'ʵʱ�϶�
        .options.ThemedFloatingFrames = True
        .options.AlphaDockingContext = True
    End With
    
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    
    lngListWidth = nvl(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "StudyListWidth", 0.35))
    If lngListWidth >= 1 Then lngListWidth = 0.35
    
    'ע����б���Ľ��沼��Pnae�������ԣ������Ĭ�ϵ�Pane����
    If dkpMain.PanesCount <> 3 Then
        dkpMain.DestroyAll
        
        Set Pane1 = dkpMain.CreatePane(1, lngListWidth * 100, 250, DockLeftOf, Nothing)
        Pane1.title = "����б�"
        Pane1.Handle = picList.hwnd
        Pane1.options = PaneNoCloseable Or PaneNoFloatable
        
        Set Pane2 = dkpMain.CreatePane(2, (1 - lngListWidth) * 100, 300, DockRightOf, Nothing)
        Pane2.title = "�Ӵ���"
        Pane2.Handle = picWindow.hwnd
        Pane2.options = PaneNoCaption Or PaneNoCloseable

    End If
End Sub

Private Sub InitCommandBars()
    '���ܴ���������
On Error GoTo errH
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim str3DFuncs() As String
    Dim blnShowCaption As Boolean
    
    Dim rsCollection As ADODB.Recordset
    Dim rsViewShare As ADODB.Recordset
    Dim rsShareCount As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    
    Dim i As Integer
    Dim i3DFunc As Integer
    Dim intTxtLen As Integer
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
    With Me.cbrMain.options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    If mstrPublicAdvicePrivs = "-1" Then mstrPublicAdvicePrivs = ";" & GetPrivFunc(100, 9001) & ";"
'�˵�����
'Begin------------------------�ļ��˵�--------------------------------------Ĭ�Ͽɼ�
    Me.cbrMain.ActiveMenuBar.title = "�˵�"
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_FilePopup, "�ļ�", "", 0, False)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_PrintSet, "��ӡ����", "", 181, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Excel, "�嵥��ӡ", "", 103, False)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Parameter, "��������", "", 181, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, ConMenu_File_ShortcutSet, "��ݼ�����", "", 181, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_WorkModule, "վ��ģʽ����", "", 9004, False)
    
'        If mblnSetXWParam = True And mlngModule = G_LNG_PACSSTATION_MODULE Then    '�С�Ӱ���豸Ŀ¼����Ȩ�ޣ���������������PACS�Ĳ���
'            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_SetXWParam, "PACS��������", "", 9004, False)
'        End If
        
        '������Ƶ�ɼ����ò˵�
        If mlngModule <> G_LNG_PACSSTATION_MODULE Then
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_DevSet, "��Ƶ����", "��Ƶ����", 815, False)
        End If
        
        If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
            '�����û������˵�
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ChangeUser, "�û�����", "�������ҽ���ͱ���ҽ��", 3012, True)
        End If
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_SwitchUser, "�л��û�", "�л��û�", 3012, True)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_SendImg, "����ͼ��", "", 3061, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Change_In, "�����б�", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Exit, "�˳�", "", 191, True)
    End With



'Begin----------------------���˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_ManagePopup, "���", "", 0, False)
    With cbrMenuBar.CommandBar
    
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_Request, "���뵥", "���뵥", 0, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButtonPopup, conMenu_Manage_RequestPrint, "��ӡ���뵥��", "", 0, False)
        
            '����������뵥ɨ����� ��ѡ������ء�ɨ�����뵥���˵���δ��ѡ�� ������
            If mSysPar.blnIsPetitionScan Then
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, comMenu_Petition_Capture, "ɨ�����뵥", "", 5020, , False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, comMenu_Petition_View, "�鿴���뵥", "�鿴��ɨ������뵥ͼ��", 3935, True)
            End If
            
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_CheckList, "�鿴��������", "�鿴�������뵥", 8044, False)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Check_ViewLink, "�鿴�������", "�鿴�������", 102, False)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Regist, "���Ǽ�", "", 2110, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_CopyCheck, "���ƵǼ�", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ReGet, "�ٻ�ȡ��", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Receive, "��鱨��", "", 744, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_InQueue, "���", "��ʼ�Ŷ�", 3534, True)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Schedule, "���ԤԼ", "", 0, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ScheduleManage, "ԤԼ����", "", 0, False)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Transfer, "����Ӱ��", "", 505, True)
                
        If mlngModule = G_LNG_PACSSTATION_MODULE Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_SendArrange, "���Ͱ���", "", 232, False)
        End If
        
        '�����
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_SendAudit, "�������", "���͵������", 0, False)
        Call CreateAuditorMenu(cbrControl)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_LookMecRecord, "��������", "", 102, False)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ReportExecutor, "����ִ��", "ָ����ǰ����ļ�¼��", 5008, True)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Complete, "������", "", 225, False, , False)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_Change_Undo, "��������", "��������", 0, True)
        If Not cbrControl Is Nothing Then
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Redo, "ȡ���Ǽ�", "", 742, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Logout, "ȡ������", "", 743, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Undone, "ȡ�����", "", 2615, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Cancel, "ȡ������", "", 506, False)
        End If
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_State, "�����", "�����", 0, True)

            If mlngModule = G_LNG_PACSSTATION_MODULE Then
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlPopup, conMenu_Manage_Release, "���Ŵ���", "�����Ƭ���Ŵ���", 3013, False)
                If Not cbrPopControl Is Nothing Then
                    Call CreateModuleMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "���淢��", "", 8215, False)
                    Call CreateModuleMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FilmRelease, "��Ƭ����", "", 8216, False)
                End If
            Else
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "���淢��", "", 8215, False)
            End If
    
             
            '�����
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlPopup, conMenu_Manage_Result, "������", "", 0, False)
            If Not cbrPopControl Is Nothing Then
                Call CreateModuleMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Negative, "�������", "", 3506, False)
                Call CreateModuleMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Positive, "�������", "", 3507, False)
            End If
            '�������
            If mlngModule <> G_LNG_PATHOLSYS_NUM Then
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlPopup, conMenu_Manage_FuHeLevel, "�������", "", 0, False)
                If Not cbrPopControl Is Nothing Then
                    Call CreateModuleMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FuHe, "����", "", 3587, False)
                    Call CreateModuleMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_JiBenFuHe, "��������", "", 3010, False)
                    Call CreateModuleMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_BuFuHe, "������", "", 3010, False)
                End If
            End If
                
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlPopup, conMenu_Manage_GChannel, "��ɫͨ��", "", 0, False, , False)
            If Not cbrPopControl Is Nothing Then
                Call CreateModuleMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_GChannelOk, "���", "", 0, False, , False)
                Call CreateModuleMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_GChannelCancel, "ȡ��", "", 0, False, , False)
            End If
        
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_More, "�������", "�������", 0, True)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ThingModi, "�޸���Ϣ", "", 0, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ModifBaseInfo, "������Ϣ����", "", 4113, False)
            
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ExecOnePart, "�ֲ�λִ��", "�ֲ�λִ�к�ȡ��ҽ��", 0, True)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Review, "������Ϣ", "", 232, False)
    
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_DiseaseRegist, "��Ⱦ���Ǽ�", "��Ⱦ���Ǽ�", 3564, True)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_DiseaseQuery, "��Ⱦ����ѯ", "��Ⱦ����ѯ", 102, False)
            
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsCriticalReg, "Σ�����ߵǼ�", "", 8344, True)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsCriticalManage, "Σ�����߹���", "", 8345, False)
        
        
            If Not (mobjAppendBill Is Nothing) And GetInsidePrivs(pҽ�����ѹ���, True) <> "" Then
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_AttachMoney, "���ӷ���", "", 3011, True)
                
                If glngModul = G_LNG_PATHOLSYS_NUM Then
                    Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_CompleteAttach, "��ɲ���", "", 3816, False)
                End If
            End If
        

            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_RelatingPatiet, "��������", "", 803, True)
            
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Burn, "ͼ���¼", "", 0, True)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReLoadPDF, "�����ϴ�PDF", "", 0, True)
            
        
        If mlngModule <> G_LNG_PACSSTATION_MODULE Then
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Tool_Analyse, "�߼�����", "�߼�ͼ����", 0, True)
        End If
    
    End With
    
    
    
'Begin-------------------------------------------------------�ղز˵�(Ĭ�Ͽɼ�)----------------------------------------------------------

    'gstrSQL = "select ID ,�ϼ�id,������,�ղ���� from Ӱ���ղ���� where ������='" & UserInfo.���� & "' Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id"
        gstrSQL = "select a.ID ,a.�ϼ�id,b.���� as ������,a.�ղ���� from Ӱ���ղ���� a,��Ա�� b where a.������ID=" & UserInfo.ID & " and a.������id=b.ID(+) Start With a.�ϼ�id Is Null Connect By Prior a.ID = a.�ϼ�id"
    Set rsCollection = zlDatabase.OpenSQLRecord(gstrSQL, GetWindowCaption)

    'gstrSQL = "select ID ,�ϼ�id,������,�ղ����,�Ƿ��� from Ӱ���ղ���� where ������<>'" & UserInfo.���� & "' Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id"
        gstrSQL = "select a.ID ,a.�ϼ�id,b.���� as ������,a.�ղ����,a.�Ƿ��� from Ӱ���ղ���� a,��Ա�� b where a.������ID<>" & UserInfo.ID & " and a.������id=b.ID(+) Start With a.�ϼ�id Is Null Connect By Prior a.ID = a.�ϼ�id"
    Set rsViewShare = zlDatabase.OpenSQLRecord(gstrSQL, GetWindowCaption)
        
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_Collection, "�ղ�", "", 0, False) ' Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Collection, "�ղ�", -1, False)
    With cbrMenuBar.CommandBar
        
        '��¡���� ɸѡ����������ݽ����ж�
        Set rsShareCount = zlDatabase.CopyNewRec(rsViewShare)
        rsShareCount.Filter = "�Ƿ���=1"
        
        If rsShareCount.RecordCount <> 0 Then
           '�ݹ鴴������˵�
           mlngShareFatherID = 0
           Set rsTemp = zlDatabase.CopyNewRec(rsViewShare)
           rsViewShare.Filter = "�ϼ�ID=" & nvl(rsViewShare!�ϼ�ID, 1) & " and ������<> '" & UserInfo.���� & "'"
           
           Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_Collection_ViewShare, "����鿴", "", 0, False)
           Call RecursionCreateShareMenu(rsViewShare, rsTemp, cbrControl)
        End If

        If rsCollection.RecordCount > 0 Then
            '�ݹ鴴���ղ����˵�
                 mlngCollectionFatherID = 0
                 Set rsTemp = zlDatabase.CopyNewRec(rsCollection)
                 rsCollection.Filter = "�ϼ�ID=" & nvl(rsCollection!�ϼ�ID, 1)
                 Call RecursionCreateCollectionMenu(rsCollection, rsTemp, cbrMenuBar)
        End If
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Collection_To, "�ղص�...", "", 0, True) '.Add(xtpControlButton, conMenu_Collection_To, "�ղص�...")
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Collection_Manage, "�ղع���", "", 0, False) '.Add(xtpControlButton, conMenu_Collection_Manage, "�ղع���", -1, False)
        
    End With
    
    '��ȡ��������ģ��ı���(��������ģ���)
'-----------------------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "����(&R)")
    cbrMenuBar.ID = conMenu_ReportPopup
    
    Call zlDatabase.ShowReportMenu(cbrMain, glngSys, mlngModule, mstrPrivs, _
                                    "ZL1_INSIDE_1294_01", _
                                    "ZL1_INSIDE_1294_02", _
                                    "ZL1_INSIDE_1294_03", _
                                    "ZL1_INSIDE_1294_04", _
                                    "ZL1_INSIDE_1294_05", _
                                    "ZL1_INSIDE_1294_06", _
                                    "ZL1_INSIDE_1294_07", _
                                    "ZL1_INSIDE_1294_08", _
                                    "ZL1_INSIDE_1294_09", _
                                    "ZL1_INSIDE_1294_10", _
                                    "ZL1_INSIDE_1294_11", _
                                    "ZL1_INSIDE_1294_12", _
                                    "ZL1_INSIDE_1294_13", _
                                    "ZL1_INSIDE_1294_15")
    If cbrMenuBar.CommandBar.Controls.Count > 0 Then
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
        
        For i = 1 To cbrMenuBar.CommandBar.Controls.Count
            cbrMenuBar.CommandBar.Controls(i).Category = M_STR_MODULE_MENU_TAG
        Next i
    Else
        cbrMenuBar.Delete
    End If
    
'Begin----------------------�Զ����ѯ�˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_Manage_Query, "��ѯ", "", 0, False)
    
    Call mobjPacsQueryWrap.RefreshCustomQueryMenu(cbrMenuBar, mintQueryState, tabScheme, mSysPar.blnQuickTabDisplayScheme)
    
    Call CheckHaveScheme(False, "")
    
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_QueryCFG, "��ѯ����", "", 0, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_QueryCfgUserScheme, "���÷�������", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_QueryTabDisplayScheme, "��ʾ���÷�����ǩ", "", 0, True)
        cbrControl.Checked = mSysPar.blnQuickTabDisplayScheme
        cbrControl.CloseSubMenuOnClick = False
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_QueryValidTime, "��������", "", 0, False)
        cbrControl.Checked = mSysPar.blnQuickTabDisplayScheme
        cbrControl.CloseSubMenuOnClick = False
    End With
'Begin----------------------���������ܲ���˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_Manage_PacsPlugIn, "���", "", 0, False)
    Call RefreshCustomPlugInMenu(cbrMenuBar, mlngModule)
    Call initInterface(mlngModule)

'Begin----------------------�鿴�˵�--------------------------------------
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_ViewPopup, "�鿴", "", 0, False)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_View_ToolBar, "������", "", 0, False)
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar '�����˵�
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť", "", 0, False): cbrPopControl.Checked = True
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ", "", 0, False): cbrPopControl.Checked = True
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��", "", 0, False): cbrPopControl.Checked = True
            End With
            
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_View_FontSize, "�����С", "", 0, False)
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar '�����˵�
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_FontSize_S, "С����", "", 0, False): cbrPopControl.Checked = True
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_FontSize_M, "������", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_FontSize_L, "������", "", 0, False)
            End With
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_StatusBar, "״̬��", "", 0, True): cbrControl.Checked = True
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_View_Filter * 10#, "������", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_Refresh, "ˢ��", "", 0, False)
    End With
    
'Begin----------------------���߲˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_ToolPopup, "����", "", 0, False)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Tool_Valid, "ͼ��У�Թ���", "", 0, False)
    End With

'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_HelpPopup, "����", "", 0, False)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Help_Help, "��������", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_Help_Web, "WEB�ϵ�����", "", 0, False)
            With cbrControl.CommandBar
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Help_Web_Forum, "������̳", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Help_Web_Home, "������ҳ", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���", "", 0, False)
            End With
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Help_About, "���ڡ�", "", 0, True)
    End With
    

'---------------------�������Ͻǵ�ǰ����----------------------------------
    Set cbrControl = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_View_Filter * 10#, "������", "", 0, False): cbrControl.flags = xtpFlagRightAlign
            
    '���ұ���ʾ�����ɼ���ť
    If mlngModule <> G_LNG_PACSSTATION_MODULE Then
        Set cbrControl = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlButton, comMenu_Cap_Process, "�����ɼ�", "���������ɼ�����", 0, False): cbrControl.flags = xtpFlagRightAlign
    End If
        
'---------------------����������------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True

    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Regist, "�Ǽ�", "���Ǽ�", 211, True)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Receive, "����", "��鱨��", 744, False)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Logout, "ȡ��", "ȡ������", 743, False)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Schedule, "ԤԼ", "���ԤԼ", 6823, True)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_InQueue, "���", "��ʼ�Ŷ�", 3534, True)
    
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_View_Filter, "����", "����", 0, True)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_View_Refresh, "ˢ��", "ˢ��", 0, False)
        
    Call AddPlugInToolBarMenu(cbrToolBar.Controls, mlngModule)  '100908
    
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Review, "��ע", "������Ϣ", 232, True)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, comMenu_Petition_View, "�鿴���뵥", "�鿴��ɨ������뵥ͼ��", 3935, False)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_CheckList, "�鿴��������", "�鿴�������뵥", 8044, False)
    
    If Not (mobjAppendBill Is Nothing) And GetInsidePrivs(pҽ�����ѹ���, True) <> "" Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_AttachMoney, "������", "������", 3011, False)
        If glngModul = G_LNG_PATHOLSYS_NUM Then
            Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_CompleteAttach, "��ɲ���", "��ɲ���", 3816, False)
        End If
    End If
    
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_Disease, "��Ⱦ��", "��Ⱦ��", 3842, False)
    If Not cbrControl Is Nothing Then
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_DiseaseRegist, "��Ⱦ���Ǽ�", "��Ⱦ���Ǽ�", 3564, False)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_DiseaseQuery, "��Ⱦ����ѯ", "��Ⱦ����ѯ", 102, False)
    End If
    
    If mlngModule <> G_LNG_PACSSTATION_MODULE Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Tool_Analyse, "�߼�", "�߼�ͼ����")
    End If
    
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_SwitchUser, "�л�", "�л��û�", 3012, False, conMenu_Tool_Analyse)
    
    If mlngModule = G_LNG_PACSSTATION_MODULE Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_Release, "���Ŵ���", "�����Ƭ���Ŵ���", 3013, False)
        If Not cbrControl Is Nothing Then
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "���淢��", "���淢��", 8215, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FilmRelease, "��Ƭ����", "��Ƭ����", 8216, False)
        End If
    Else
        Set cbrPopControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "���淢��", "���淢��", 8215, False)
    End If
    
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_ReportExecutor, "����ִ��", "ָ����ǰ����ļ�¼��", 5008, False)
    
    If mlngModule = G_LNG_PACSSTATION_MODULE Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_SendArrange, "���Ͱ���", "���Ͱ���", 232, False)
    End If
    
    'Σ�����
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_PacsCritical, "Σ��ֵ", "Σ�����", 8338, False)
    If Not cbrControl Is Nothing Then
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsCriticalReg, "Σ��ֵ�Ǽ�", "Σ��ֵ���ߵǼ�", 8345, False)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsCriticalManage, "Σ��ֵ����", "Σ��ֵ���߹���", 8338, True)
    End If
    
    '�����������
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_Result, "���", "�����������", 3506, False)
    If Not cbrControl Is Nothing Then
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Negative, "����", "����", 3506, False)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Positive, "����", "����", 3507, False)
    End If
    
    '����ǲ���ϵͳ����û�з��������ť
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_FuHeLevel, "�������", "�������", 8044, False)
        If Not cbrControl Is Nothing Then
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FuHe, "����", "����", 0, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_JiBenFuHe, "��������", "��������", 0, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_BuFuHe, "������", "������", 0, False)
        End If
    End If
        
    'ֻ��Ӱ��ɼ�ϵͳ�ž����û���������
    If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_ChangeUser, "����", "�������ҽ���ͱ���ҽ��", 3012, False)
    End If
    
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Complete, "���", "����������", 225, False, , False)
  
     '��ʼ���������� �ӵ�����Ϊ�˷�ֹ��һЩ���������ʱ�򣬻ᵼ������ָ��ɳ�ʼ��
    Call SetFontSize(IIf(gbytFontSize = 12, 1, IIf(gbytFontSize = 15, 2, 0)))
'    '��������ģ������Ĳ˵�
'    Call CreateWorkModuleMenu
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, _
    Optional ByVal lngIndex As Long = -1, Optional blnIsControlCreate As Boolean = True) As CommandBarControl
'������ģ���ڵĲ˵�
On Error GoTo err
    Dim blHavePrives As Boolean '�Ƿ�߱���Ӧ�˵�Ȩ��
    'blnIsControlCreate �Ƿ�û��Ȩ��Ҳ�������˵���
    
    '�����˵�ǰ����ID ��Ȩ���ж��Ƿ���ֹ��������
    'ע��  conMenu_Manage_GChannel  conMenu_Manage_Complete conMenu_Manage_Result conMenu_Edit_Audit
    'conMenu_PacsReport_RepFormat ���봴��
    blHavePrives = True
    
    Select Case lngID
        Case conMenu_File_SendImg '����ͼ��
            If Not CheckPopedom(mstrPrivs, "�ļ�����") Then blHavePrives = False
            
        Case conMenu_Manage_Regist, conMenu_Manage_CopyCheck, conMenu_Manage_Redo, conMenu_Manage_ThingModi, comMenu_Petition_View
        '���Ǽǣ����ƵǼǣ�ȡ���Ǽ�, �޸���Ϣ,�鿴���뵥
            If Not CheckPopedom(mstrPrivs, "���Ǽ�") Then blHavePrives = False
            
        Case conMenu_Manage_Receive '��鱨��
            If Not CheckPopedom(mstrPrivs, "��鱨��") Then blHavePrives = False
            
        Case conMenu_Manage_Logout 'ȡ������
            If Not CheckPopedom(mstrPrivs, "ȡ������") Then blHavePrives = False
            
        Case conMenu_Manage_Transfer, conMenu_Manage_Cancel '����Ӱ�� ȡ������
            If Not CheckPopedom(mstrPrivs, "ͼ�����") Then blHavePrives = False
            
        Case conMenu_Manage_Review '���
            If Not CheckPopedom(mstrPrivs, "���") Then blHavePrives = False
            
        Case conMenu_Manage_Disease
            If Not (CheckPopedom(mstrPublicAdvicePrivs, "��Ⱦ�����Խ���Ǽ�") Or CheckPopedom(mstrPublicAdvicePrivs, "��Ⱦ�����Խ����ѯ")) Then blHavePrives = False
            
        Case conMenu_Manage_DiseaseRegist
            If Not CheckPopedom(mstrPublicAdvicePrivs, "��Ⱦ�����Խ���Ǽ�") Then blHavePrives = False
            
        Case conMenu_Manage_DiseaseQuery
            If Not CheckPopedom(mstrPublicAdvicePrivs, "��Ⱦ�����Խ����ѯ") Then blHavePrives = False
            
        Case conMenu_Manage_PacsCritical, conMenu_Manage_PacsCriticalReg, conMenu_Manage_PacsCriticalManage
            If Not CheckPopedom(mstrPublicAdvicePrivs, "Σ��ֵ����") Then blHavePrives = False
            
        Case conMenu_Manage_Undone
            If Not CheckPopedom(mstrPrivs, "ȡ��������") Then blHavePrives = False
            
        Case conMenu_Manage_RelatingPatiet
            If Not (CheckPopedom(mstrPrivs, "��������") And mSysPar.blnRelatingPatient) Then blHavePrives = False
        
        Case conMenu_Manage_ReLoadPDF
            blHavePrives = True
            
        Case conMenu_Manage_Burn
            If Not CheckPopedom(mstrPrivs, "ͼ���¼") Then blHavePrives = False
            
        Case conMenu_Tool_Analyse '�߼�ͼ����
            If Not CheckPopedom(";" & GetPrivFunc(glngSys, 1289) & ";", "����") Then blHavePrives = False
        '------------------
        Case conMenu_Manage_GChannel, conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel
        '��ɫͨ�� ��� ȡ��
            If Not CheckPopedom(mstrPrivs, "��ɫͨ��") Then blHavePrives = False
            
        Case conMenu_Manage_Complete  '������
            If Not CheckPopedom(mstrPrivs, "������") Then blHavePrives = False
            
        Case conMenu_Manage_ModifBaseInfo '������Ϣ����
            If Not CheckPopedom(mstrPrivs, "ǿ���޸�סԺ������Ϣ") Then blHavePrives = False
            
        Case conMenu_Manage_ExecOnePart '�ֲ�λִ��
            If Not CheckPopedom(mstrPrivs, "ȡ������") Then blHavePrives = False

        Case conMenu_Manage_ConfigQuery, conMenu_Manage_QueryCFG
            If Not CheckPopedom(mstrPrivs, "��ѯ����") Then blHavePrives = False
        
        Case conMenu_Manage_ReportExecutor '����ִ��
            If Not CheckPopedom(mstrPrivs, "����ִ��") Then blHavePrives = False
        
        Case conMenu_Manage_Schedule, conMenu_Manage_ScheduleManage       '���ԤԼ,ԤԼ����
            If Not CheckPopedom(mstrPrivs, "���ԤԼ") Then blHavePrives = False
            
        Case Else
    End Select
    
    If blHavePrives Or Not blnIsControlCreate Then
    
        If lngIndex >= 0 Then
            Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
        Else
            Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
        End If
    
        CreateModuleMenu.ID = lngID '������ﲻָ��id�����ܽ���Щ�˵���ӵ��Ҽ��˵���
        
        If lngIconId <> 0 Then CreateModuleMenu.iconid = lngIconId
        If blnStartGroup Then CreateModuleMenu.BeginGroup = True
        If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
        
        If Not blHavePrives Then
            CreateModuleMenu.Visible = False
        End If
        
        CreateModuleMenu.Category = M_STR_MODULE_MENU_TAG
    End If
    Exit Function
err:
End Function


Private Sub CreateWorkModuleMenu()
'��������ģ��˵�
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
On Error GoTo err
    If Not mobjWork_Pathol Is Nothing And mblnIsLoadPatholModule Then
        Call mobjWork_Pathol.zlMenu.zlCreateMenu(Me.cbrMain)
        Call mobjWork_Pathol.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
    End If
    
    '����Ӱ��ͼ��ģ����ز˵���������
    If Not mfrmWork_PacsImg Is Nothing And InStr(mstrWorkModule, ";Ӱ��ͼ��ģ��;") > 0 Then
        Call mfrmWork_PacsImg.zlMenu.zlCreateMenu(Me.cbrMain)
        Call mfrmWork_PacsImg.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
    End If
    
    If Not mobjWork_ImageCap Is Nothing And InStr(mstrWorkModule, ";Ӱ��ɼ�ģ��;") > 0 Then
        'TODO:������Ƶ�ɼ�ģ��˵�
'            Call mobjWork_ImageCap.zlMenu.zlCreateMenu(Me.cbrMain)
'            Call mobjWork_ImageCap.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
    End If
    
    '���뽫����˵��Ĵ�������mobjWork_His�����˵�֮ǰ���������л�������ģ��ʱ����Ӧ��ģ��˵����ܹ�������
    If Not mobjWork_Report Is Nothing And _
        (InStr(mstrWorkModule, ";Ӱ�񱨸�ģ��;") > 0 Or InStr(mstrWorkModule, ";�������ģ��;") > 0) Then
        Call mobjWork_Report.zlMenu.zlCreateMenu(Me.cbrMain)
        Call mobjWork_Report.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
    End If
    
    If Not mobjWork_His Is Nothing Then
        '��Ϊ��PACSϵͳ�� ����ӡ�� �˵����ڱ༭�˵����£������������ļ��˵��£������ڵ��ò����Ĳ˵���������ʱ��
        '���ļ��˵����Ҳ�����ӡ�˵����������PACS�У��嵥��ӡ���ļ��˵��£����Ե��ò����Ĳ˵���������ʱ��
        '�嵥��ӡ��id�ĳɴ�ӡ��id��������󣬻ָ��嵥��ӡԭ����id
        If Not TabWindow.Selected Is Nothing Then
            If TabWindow.Selected.tag = "������Ӳ���" Or TabWindow.Selected.tag = "סԺ���Ӳ���" Then
                Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
                Set cbrControl = cbrMenuBar.CommandBar.Controls.Find(, conMenu_File_Excel)
                cbrControl.ID = conMenu_File_Print
            End If
        End If
        
        Call mobjWork_His.zlMenu.zlCreateMenu(Me.cbrMain)
        If Not TabWindow.Selected Is Nothing Then
            If TabWindow.Selected.tag = "������Ӳ���" Or TabWindow.Selected.tag = "סԺ���Ӳ���" Then
                cbrControl.ID = conMenu_File_Excel
            End If
        End If
    End If

    Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
    
    Call cbrMain.RecalcLayout
    Exit Sub
err:
    cbrControl.ID = conMenu_File_Excel
End Sub

Private Sub RecursionCreateShareMenu(rsFilterADO As ADODB.Recordset, rsFullADO As ADODB.Recordset, cbrParentControl As CommandBarControl, Optional blnIsShare As Boolean = False)
'�ݹ�ѭ����������˵�
    Dim rsFilterTemp As ADODB.Recordset
    Dim i As Long
    Dim cbrControl As CommandBarControl
    Static j As Long
    
    If rsFilterADO.RecordCount = 0 Then Exit Sub
    rsFilterADO.MoveFirst
    
    With cbrParentControl.CommandBar.Controls
        If mlngShareFatherID <> 0 Then
            Set cbrControl = .Add(xtpControlButton, CLng(conMenu_Collection_ViewShare) * 10000# + mlngShareFatherID, "�鿴��ǰ�ղ�", -1, False)
            cbrControl.Category = M_STR_MODULE_MENU_TAG
        End If
        
        For i = 1 To rsFilterADO.RecordCount
            rsFullADO.Filter = " �ϼ�ID=" & nvl(rsFilterADO!ID)

            If rsFullADO.RecordCount > 0 Then
                Set cbrControl = Nothing
  
                If nvl(rsFilterADO!�Ƿ���) = 1 Or blnIsShare = True Then
                    mlngShareFatherID = nvl(rsFilterADO!ID)
                    '���������˵� ����ϼ�ID=1 ����ʾ����������
                    Set cbrControl = .Add(xtpControlButtonPopup, CLng(conMenu_Collection_ViewShare) * 10000# + j, nvl(rsFilterADO!�ղ����) & Decode(cbrParentControl.ID, conMenu_Collection_ViewShare, "(" & nvl(rsFilterADO!������) & ")", ""), -1, False)
                    cbrControl.DescriptionText = nvl(rsFilterADO!������)
                    cbrControl.Category = M_STR_MODULE_MENU_TAG
                    
                    j = j + 1
                    If i = 1 Then cbrControl.BeginGroup = True
                End If
                
                Set rsFilterTemp = zlDatabase.CopyNewRec(rsFullADO)
                '�����Լ�
                Call RecursionCreateShareMenu(rsFilterTemp, rsFullADO, IIf(cbrControl Is Nothing, cbrParentControl, cbrControl), IIf(nvl(rsFilterADO!�Ƿ���) = 0, False, True))
            Else
            '�����Ӽ��˵�
                If nvl(rsFilterADO!�Ƿ���) = 1 Or blnIsShare = True Then
                    Set cbrControl = .Add(xtpControlButton, CLng(conMenu_Collection_ViewShare) * 10000# + j, nvl(rsFilterADO!�ղ����) & Decode(cbrParentControl.ID, conMenu_Collection_ViewShare, "(" & nvl(rsFilterADO!������) & ")", ""), -1, False)
                    cbrControl.DescriptionText = nvl(rsFilterADO!������)
                    cbrControl.Category = M_STR_MODULE_MENU_TAG
                    
                    j = j + 1
                    If i = 1 Then cbrControl.BeginGroup = True
                End If
                mlngShareFatherID = 0
            End If

            If Not rsFilterADO.EOF Then rsFilterADO.MoveNext
        Next
    End With
End Sub



Private Sub RecursionCreateCollectionMenu(rsFilterADO As ADODB.Recordset, rsFullADO As ADODB.Recordset, cbrMenuBar As CommandBarPopup)
'�ݹ�ѭ�������ղ����˵�
    Dim rsFilterTemp As ADODB.Recordset
    Dim cbrControl As CommandBarControl
    Dim i As Long
    Static j As Long

    If rsFilterADO.RecordCount = 0 Then Exit Sub
    rsFilterADO.MoveFirst

    With cbrMenuBar.CommandBar.Controls
        If mlngCollectionFatherID <> 0 Then
            Set cbrControl = .Add(xtpControlButton, CLng(comMenu_Collection_Type) * 10000# + mlngCollectionFatherID, "�鿴��ǰ�ղ�", -1, False)
            cbrControl.Category = M_STR_MODULE_MENU_TAG
        End If

        For i = 1 To rsFilterADO.RecordCount

            rsFullADO.Filter = " �ϼ�ID=" & nvl(rsFilterADO!ID)
            mlngCollectionFatherID = nvl(rsFilterADO!ID)
            If rsFullADO.RecordCount > 0 Then
            '���������˵�
                Set cbrControl = .Add(xtpControlButtonPopup, CLng(comMenu_Collection_Type) * 10000# + j, nvl(rsFilterADO!�ղ����), -1, False)
                cbrControl.Category = M_STR_MODULE_MENU_TAG
                
                j = j + 1
                
                Set rsFilterTemp = zlDatabase.CopyNewRec(rsFullADO)
                '�����Լ�
                Call RecursionCreateCollectionMenu(rsFilterTemp, rsFullADO, cbrControl)
                
            Else
            '�����Ӽ��˵�
                Set cbrControl = .Add(xtpControlButton, CLng(comMenu_Collection_Type) * 10000# + j, nvl(rsFilterADO!�ղ����), -1, False)
                cbrControl.Category = M_STR_MODULE_MENU_TAG
                
                j = j + 1
            End If
            If i = 1 Then cbrControl.BeginGroup = True

            If Not rsFilterADO.EOF Then rsFilterADO.MoveNext
        Next
    End With

End Sub


Private Sub ReadWorkModuleCfg()
    '���õ�ǰ��Ҫ�����Ĺ���ҳ��
    mstrWorkModule = zlDatabase.GetPara("վ��ģ��", glngSys, mlngModule, "")
    mstrWorkModule = IIf(mstrWorkModule <> "", ";" & mstrWorkModule & ";", "")
    
    '���ģ��Ϊ�գ������������Ŷӽкţ���ֻ��ʾ�ŶӽкŹ���ģ��
    If mstrWorkModule = "" Then 'And Not mblnUseQueue
        Select Case mlngModule
            Case G_LNG_PACSSTATION_MODULE
                mstrWorkModule = ";Ӱ��ͼ��ģ��;Ӱ�񱨸�ģ��;������¼ģ��;���ü�¼ģ��;ҽ����¼ģ��;"
            
            Case G_LNG_VIDEOSTATION_MODULE
                mstrWorkModule = ";Ӱ��ɼ�ģ��;Ӱ�񱨸�ģ��;������¼ģ��;���ü�¼ģ��;ҽ����¼ģ��;"
            
            Case G_LNG_PATHOLSYS_NUM
                mstrWorkModule = ";�걾����ģ��;Ӱ��ɼ�ģ��;����ȡ��ģ��;������Ƭģ��;�����ؼ�ģ��;���̱���ģ��;�������ģ��;������¼ģ��;���ü�¼ģ��;ҽ����¼ģ��;"
            Case Else
                Exit Sub
        End Select
    End If
    
'    '���Դ���
'    mstrWorkModule = ";Ӱ��ͼ��ģ��;Ӱ��ɼ�ģ��;�걾����ģ��;����ȡ��ģ��;������Ƭģ��;�����ؼ�ģ��;���̱���ģ��;Ӱ�񱨸�ģ��;���ü�¼ģ��;ҽ����¼ģ��;������¼ģ��;"
End Sub


Private Sub InitPatholModuleObj()
    '��ʼ���������ģ�����
    If mobjWork_Pathol Is Nothing Then
        Set mobjWork_Pathol = New clsWorkModule_Pathol
        Call mobjWork_Pathol.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
    End If
End Sub

Private Sub InitHisModuleObj()
    '��ʼ��HIS���ģ�����
    If mobjWork_His Is Nothing Then
        Set mobjWork_His = New clsWorkModule_His
        
        If mblnAllDepts Then
            Call mobjWork_His.zlInitModule(mlngModule, mstrPrivs, UserInfo.����ID, Me)
        Else
            Call mobjWork_His.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
        End If
    End If
End Sub

Private Sub InitActiveVideoModuleObj()
'��ʼ��ActivexExe��Ƶ�ɼ�ģ�����
    If mlngModule = G_LNG_PACSSTATION_MODULE Then Exit Sub
    If Not CheckPopedom(mstrPrivs, "��Ƶ�ɼ�") Then Exit Sub
    If InStr(mstrWorkModule, ";Ӱ��ɼ�ģ��;") < 0 Then Exit Sub
    
    If mobjWork_ImageCap Is Nothing Then
        Set mobjWork_ImageCap = CreateObject("zl9PacsImageCap.clsPacsCapture") ' New zl9PacsCapture.clsPacsCapture
        With mobjWork_ImageCap
            If .ModuleNo <> mlngModule And .ModuleNo <> 0 Then .ModuleNo = mlngModule
            .ParentWindowKey = Me.Name
            .AllowEventNotify = True
            .ImgLoadType = IIf(GetServiceStatus = SERVICE_RUNNING, FileLoadType.Service, FileLoadType.Normal)
            
            Call .RegEventObj(Me)
            
            Call .zlInitModule(gcnOracle, glngSys, mlngModule, mstrPrivs, mlngCur����ID, Me.hwnd, Me, True, gblnUseDebugLog)
        End With
    End If
End Sub

Private Sub ShowModuleLoadState(Optional ByVal strState As String = "")
'��ʾ����״̬
On Error GoTo errhandle
    picLoadState.Left = 0
    picLoadState.Top = 350
    picLoadState.Width = picWindow.Width - 0
    picLoadState.Height = picWindow.Height - 350
    
    
    If strState <> "" Then
        labLoadState.Caption = strState
        Call picLoadState_Resize
    End If
    
    picLoadState.Visible = True
    
errhandle:
End Sub

Private Sub HideModuleLoadState()
'��������״̬
    picLoadState.Visible = False
End Sub

Public Sub InitSubForm()
    Dim i As Integer
    Dim blnDoEvents As Boolean

    mblnIsLoadPatholModule = False   '���ñ��������ȻΪfalseʱ�����������ɾ������˵�
    blnDoEvents = True  '��ֵΪtrueʱ�������ι���ģ����ع����е��¼�����
    
    Call ShowModuleLoadState
    DoEvents
    
    With TabWindow
        .RemoveAll
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.ColorSet.ButtonNormal = &HE0E0E0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ButtonMargin.Top = 4
        .PaintManager.ButtonMargin.Bottom = 4
        .PaintManager.ShowIcons = True
        .RemoveAll
        
        '��ȡ����ģ������
        Call ReadWorkModuleCfg
    
        If InStr(mstrWorkModule, ";Ӱ��ͼ��ģ��;") > 0 Then
            '����Ӱ���¼ģ��
            If mfrmWork_PacsImg Is Nothing Then
                Set mfrmWork_PacsImg = New frmWork_Image
                
                Set mfrmWork_PacsImg.PacsCore = mobjPacsCore
                Call mfrmWork_PacsImg.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
            End If
    
            .InsertItem 0, "Ӱ���¼", picTemp.hwnd, conMenu_Img_Look
            .Item(TabWindow.ItemCount - 1).tag = "Ӱ��ͼ��"
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
            
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mfrmWork_PacsImg Is Nothing Then
                Call mfrmWork_PacsImg.zlMenu.zlClearMenu
                Call mfrmWork_PacsImg.zlMenu.zlClearToolBar
            End If
        End If
                        
        If mlngModule <> G_LNG_PACSSTATION_MODULE And CheckPopedom(mstrPrivs, "��Ƶ�ɼ�") _
            And InStr(mstrWorkModule, ";Ӱ��ɼ�ģ��;") > 0 Then
            
            If mobjCaptureHot Is Nothing Then
                Set mobjCaptureHot = New zl9PacsControl.clsHookKey
                Call mobjCaptureHot.EnableHook(WM_KEYDOWN, True)
            End If

            Call InitActiveVideoModuleObj
            
            .InsertItem 1, "Ӱ��ɼ�", mobjWork_ImageCap.ContainerHwnd, conMenu_Cap_Dynamic
            .Item(TabWindow.ItemCount - 1).tag = "Ӱ��ɼ�"

            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
        
        If CheckPopedom(mstrPrivs, "�걾����") And InStr(mstrWorkModule, ";�걾����ģ��;") > 0 Then
            Call InitPatholModuleObj
            
            .InsertItem 2, "�걾����", picTemp.hwnd, G_INT_ICONID_SPECIMEN
            .Item(TabWindow.ItemCount - 1).tag = "�걾����"
            
            mblnIsLoadPatholModule = True

            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
        
        If CheckPopedom(mstrPrivs, "����ȡ��") And InStr(mstrWorkModule, ";����ȡ��ģ��;") > 0 Then
            Call InitPatholModuleObj
            
            .InsertItem 3, "����ȡ��", picTemp.hwnd, G_INT_ICONID_MATERIAL
            .Item(TabWindow.ItemCount - 1).tag = "����ȡ��"
            
            mblnIsLoadPatholModule = True
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
        
        If CheckPopedom(mstrPrivs, "������Ƭ") And InStr(mstrWorkModule, ";������Ƭģ��;") > 0 Then
            Call InitPatholModuleObj
            
            .InsertItem 4, "������Ƭ", picTemp.hwnd, G_INT_ICONID_SLICES
            .Item(TabWindow.ItemCount - 1).tag = "������Ƭ"
            
            mblnIsLoadPatholModule = True
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
        
        If (CheckPopedom(mstrPrivs, "�����黯") Or CheckPopedom(mstrPrivs, "����Ⱦɫ") Or CheckPopedom(mstrPrivs, "���Ӳ���")) _
            And InStr(mstrWorkModule, ";�����ؼ�ģ��;") > 0 Then
            Call InitPatholModuleObj
            
            .InsertItem 5, "�����ؼ�", picTemp.hwnd, G_INT_ICONID_SPEEXAM
            .Item(TabWindow.ItemCount - 1).tag = "�����ؼ�"
            
            mblnIsLoadPatholModule = True
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
        
        If (CheckPopedom(mstrPrivs, "��������") Or CheckPopedom(mstrPrivs, "��Ⱦ����") _
            Or CheckPopedom(mstrPrivs, "���ӱ���") Or CheckPopedom(mstrPrivs, "���߱���") _
            Or CheckPopedom(mstrPrivs, "�����ؼ챨�����")) And InStr(mstrWorkModule, ";���̱���ģ��;") > 0 Then
            Call InitPatholModuleObj
            
            .InsertItem 6, "����/�ؼ챨��", picTemp.hwnd, G_INT_ICONID_PROREPORT
            .Item(TabWindow.ItemCount - 1).tag = "���̱���"
            
            mblnIsLoadPatholModule = True
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
        
        If GetInsidePrivs(p���Ʊ������, True) <> "" And _
            (InStr(mstrWorkModule, ";Ӱ�񱨸�ģ��;") > 0 Or InStr(mstrWorkModule, ";�������ģ��;") > 0) Then
            
            If mobjWork_Report Is Nothing Then
                Set mobjWork_Report = New clsWorkModule_Report
                
                Call mobjWork_Report.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
                
                Set mobjWork_Report.PacsCore = mobjPacsCore
            End If
            
            .InsertItem 7, "Ӱ�񱨸�", picReportContainer.hwnd, 10008 'conMenu_Edit_Compend
            .Item(TabWindow.ItemCount - 1).tag = "������д"
            
            mblnIsLoadPatholModule = True
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlMenu.zlClearMenu
                Call mobjWork_Report.zlMenu.zlClearToolBar
            End If
        End If
        
        
        If Not mblnIsLoadPatholModule And Not mobjWork_Pathol Is Nothing Then
            'û�м��ز���ģ�飬��mobjWork_Pathol��Ϊ��ʱ��ɾ������˵�
            Call mobjWork_Pathol.zlMenu.zlClearMenu
            Call mobjWork_Pathol.zlMenu.zlClearToolBar
        End If
        
        If mobjAppendBill Is Nothing Then   'ʹ�û��ģʽʱ������ʾǶ��Ĳ����ѹ���
            If GetInsidePrivs(pҽ�����ѹ���, True) <> "" And InStr(mstrWorkModule, ";���ü�¼ģ��;") > 0 Then
                Call InitHisModuleObj
                
                .InsertItem 8, "���ü�¼", picTemp.hwnd, 10007
                .Item(TabWindow.ItemCount - 1).tag = "�������"
                
                If Not blnDoEvents Then
                    DoEvents
                    blnDoEvents = True
                End If
            Else
                'ɾ����Ӧ�˵��͹�����
                If Not mobjWork_His Is Nothing Then
                    '�ݲ�����hisģ��Ĳ˵�ֻ���ڸ�ģ�鱻��ʾ������±�����...
                End If
            End If
        End If
        
        If GetInsidePrivs(pסԺҽ���´�, True) <> "" And InStr(mstrWorkModule, ";ҽ����¼ģ��;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 9, "ҽ����¼", picTemp.hwnd, 10010
            .Item(TabWindow.ItemCount - 1).tag = "סԺҽ��"
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mobjWork_His Is Nothing Then
                '�ݲ�����hisģ��Ĳ˵�ֻ���ڸ�ģ�鱻��ʾ������±�����...
            End If
        End If
        
        If GetInsidePrivs(p����ҽ���´�, True) <> "" And InStr(mstrWorkModule, ";ҽ����¼ģ��;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 10, "ҽ����¼", picTemp.hwnd, 10010  ' conMenu_Edit_NewItem
            .Item(TabWindow.ItemCount - 1).tag = "����ҽ��": .Item(TabWindow.ItemCount - 1).Visible = False
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mobjWork_His Is Nothing Then
                '�ݲ�����hisģ��Ĳ˵�ֻ���ڸ�ģ�鱻��ʾ������±�����...
            End If
        End If
        
        If GetInsidePrivs(pסԺ��������, True) <> "" And InStr(mstrWorkModule, ";������¼ģ��;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 11, "������¼", picTemp.hwnd, 10009 ' conMenu_Edit_Archive
            .Item(TabWindow.ItemCount - 1).tag = "סԺ����"
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mobjWork_His Is Nothing Then
                '�ݲ�����hisģ��Ĳ˵�ֻ���ڸ�ģ�鱻��ʾ������±�����...
            End If
        End If
        
        If GetInsidePrivs(p���ﲡ������, True) <> "" And InStr(mstrWorkModule, ";������¼ģ��;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 12, "������¼", picTemp.hwnd, 10009 ' conMenu_Edit_Archive
            .Item(TabWindow.ItemCount - 1).tag = "���ﲡ��": .Item(TabWindow.ItemCount - 1).Visible = False
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mobjWork_His Is Nothing Then
                '�ݲ�����hisģ��Ĳ˵�ֻ���ڸ�ģ�鱻��ʾ������±�����...
            End If
        End If
        
        If GetInsidePrivs(p������Ӳ���, True) <> "" And InStr(mstrWorkModule, ";���Ӳ���ģ��;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 13, "���Ӳ���", picTemp.hwnd, 10009
            .Item(TabWindow.ItemCount - 1).tag = "������Ӳ���": .Item(TabWindow.ItemCount - 1).Visible = False
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mobjWork_His Is Nothing Then
                '�ݲ�����hisģ��Ĳ˵�ֻ���ڸ�ģ�鱻��ʾ������±�����...
            End If
        End If
        
        If GetInsidePrivs(pסԺ���Ӳ���, True) <> "" And InStr(mstrWorkModule, ";���Ӳ���ģ��;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 14, "���Ӳ���", picTemp.hwnd, 10009
            .Item(TabWindow.ItemCount - 1).tag = "סԺ���Ӳ���": .Item(TabWindow.ItemCount - 1).Visible = False
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            'ɾ����Ӧ�˵��͹�����
            If Not mobjWork_His Is Nothing Then
                '�ݲ�����hisģ��Ĳ˵�ֻ���ڸ�ģ�鱻��ʾ������±�����...
            End If
        End If
        
        '����Ŷӽк�ҳ��
        If mSysPar.blnUseQueue = True Then
            mstrWorkModule = mstrWorkModule & ";�Ŷӽк�ģ��;"
            
            If mobjQueue Is Nothing Then
                Set mobjQueue = New frmWork_Queue
                Call mobjQueue.zlInitPacsQueueCfg(mlngModule, mlngCur����ID, zlStr.NeedName(mstrCur����), mstrPrivs, mblnAllDepts, Me)
            End If
            
            .InsertItem 15, "�Ŷӽк�", picTemp.hwnd, 10011
            .Item(TabWindow.ItemCount - 1).tag = "�Ŷӽк�"
            
            '��ݽкŽ���
            If mSysPar.blnQueueQuick Then
                If Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
                    Call mobjQueue.OpenQueueQuick(GetSelQueueRooms(True), Me)
                End If
            End If
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
    
'        If Not GetVideoForm Is Nothing Then Call GetVideoForm.ShowVideoWindow(picVideoContainer)
    End With
    
    DoEvents
    
    If GetWorkModuleCount = 1 Then
        TabWindow.PaintManager.ClientMargin.Top = -30
    Else
        TabWindow.PaintManager.ClientMargin.Top = 0
    End If
    
    Call HideModuleLoadState
End Sub

Private Function GetWorkModuleCount() As Long
'��ȡ�ɼ�tabwindow������
    Dim i As Long
    Dim lngCount As Long
    Dim aryWorkModule() As String
    
    
    aryWorkModule = Split(mstrWorkModule, ";")
    
    For i = LBound(aryWorkModule) To UBound(aryWorkModule)
        If aryWorkModule(i) <> "" Then lngCount = lngCount + 1
    Next i
    
    GetWorkModuleCount = lngCount
End Function


Private Function GetTabWindowIndex() As Long
'��ȡ��һ���ɼ�tabwindow������
    Dim i As Long
    
    GetTabWindowIndex = -1
    For i = 0 To TabWindow.ItemCount - 1
        If TabWindow.Item(i).Visible Then
            GetTabWindowIndex = i
            Exit Function
        End If
    Next i
End Function

Private Sub mobjWork_Report_AfterDeleted(ByVal lngOrderID As Long)
    Call CheckExecuteInterface(EInterfaceExeTime.ȡ������ʱ)
    Call AfterDeleted(lngOrderID)
End Sub

Private Sub mobjWork_Report_AfterDeletedRich(ByVal lngOrderID As Long, ByVal strDocId As String)
    Call AfterDeletedRich(lngOrderID, strDocId)
End Sub

Private Sub mobjWork_Report_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub

Private Sub mobjWork_Report_AfterPrintedRich(ByVal lngOrderID As Long, ByVal strDocId As String)
    Call AfterPrintedRich(lngOrderID, strDocId)
End Sub

Private Sub mobjWork_Report_AfterSaved(ByVal lngOrderID As Long, frmOwnerForm As Object, ByVal lngSaveType As Long, ByVal isRefreshFace As Boolean)
    Call AfterReportSaved(lngOrderID, frmOwnerForm, lngSaveType, isRefreshFace)
End Sub

Private Sub mobjWork_Report_AfterSavedRich(ByVal lngOrderID As Long, ByVal strDocId As String, frmOwnerForm As Object, ByVal lngSaveType As Long)
    Call AfterReportSavedRich(lngOrderID, strDocId, frmOwnerForm, lngSaveType)
End Sub

Private Sub mobjPacsCore_AfterSaveReportImage(strStudyUID As String)
    On Error GoTo errhandle
    
    If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.RefreshReportImage
    
    Exit Sub
    
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub mobjQueue_OnDiagnose(ByVal lngAdviceId As Long, ByVal strExeRoom As String, ByVal strTurnPage As String)
'�Ŷӽ����¼�
On Error GoTo errhandle
    Dim lngIndex As String
    Dim i As Long
    Dim rsTemp As ADODB.Recordset
    
    lngIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("ҽ��ID"))
    If lngIndex = -1 Then

        
        If mSysPar.blnSynStudylist Then
            If vsfList.FindRow(lngAdviceId, 1, vsfList.ColIndex("ҽ��ID")) > 0 Then Exit Sub
    
            Set rsTemp = mobjPacsQueryWrap.CurPacsQuery.ExecuteWithAttach("[ϵͳ.ҽ��ID]", lngAdviceId, 1)
            
            If rsTemp.RecordCount > 0 Then
                Call UpdateQueryListData(rsTemp, lngAdviceId, SyncDataType.NoSync)
                '����Ĵ������ڱ�֤ѡ�е�һ�в��Ҹ��¶�Ӧ������Ϣ��������
            End If
        End If
    End If
    
    If lngIndex > 0 Then
        Call mobjPacsQueryWrap.LocateRow(lngIndex)
        
        If Trim(strTurnPage) <> "" Then
            '��ת��ָ���Ĺ���ģ��

            For i = 0 To TabWindow.ItemCount - 1
                If InStr(TabWindow(i).tag, strTurnPage) > 0 And TabWindow(i).Visible Then
                    TabWindow(i).Selected = True
                    Exit For
                End If
            Next i
        End If
    End If
    
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub mobjQueue_OnCompleted(ByVal lngAdviceId As Long, ByVal strExeRoom As String)
'�Ŷ�����¼�
On Error GoTo errhandle
    Dim lngIndex As String
    lngIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("ҽ��ID"))
    
    If lngIndex > 0 Then
        Call mobjPacsQueryWrap.LocateRow(lngIndex)
    End If
    
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mobjQueue_OnSelChange(ByVal lngAdviceId As Long)
'��ѡ��ı��¼�
On Error GoTo errhandle
    Dim lngIndex As Long
    
    If mSysPar.blnSynStudylist Then
        With vsfList
            lngIndex = .FindRow(lngAdviceId, 1, .ColIndex("ҽ��ID"), False, False)
            .Row = lngIndex
            
            '����λ��δ�����ڿɼ����귶Χ�ڣ�������ɼ�����
            If (lngIndex < .TopRow Or lngIndex > .BottomRow) And lngIndex > 0 Then
                .TopRow = lngIndex
            End If
        
            lngIndex = .FindRow(lngAdviceId, 1, .ColIndex("ҽ��ID"))
            
            If lngIndex > 0 Then
                Call mobjPacsQueryWrap.LocateRow(lngIndex)
            End If
        End With
    End If
    
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub AfterDeletedRich(ByVal lngOrderID As Long, ByVal strDocId As String)
    Dim intState As Integer
    Dim lngSendNo As Long
    Dim blnAllReportFinished As Boolean
    
On Error GoTo errhandle
    intState = getStudyStateRich(lngOrderID, strDocId, False, , lngSendNo)
    If intState = 6 Then Exit Sub
    
    gstrSQL = "Zl_Ӱ����_״̬����(" & lngOrderID & "," & lngSendNo & ",''," & intState & ",0,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������״̬��������")
    
    If intState < 4 Then
        gstrSQL = "ZL_Ӱ�񱨸���_Clear(" & lngOrderID & ")"
        zlDatabase.ExecuteProcedure gstrSQL, "��ձ��"
        
        '��մ�������
        Call Menu_Manage_SendAudit("")
    End If
    
    Call UpdateQueryListData(Nothing, lngOrderID)
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AfterDeleted(ByVal lngOrderID As Long)
On Error GoTo errhandle
    gstrSQL = "ZL_Ӱ�񱨸���_Clear(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "��ձ��"
    
    gstrSQL = "Zl_Ӱ����ͼ��_����ͼ(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "��Ǳ���ͼ"
    
    Call UpdateQueryListData(Nothing, lngOrderID)
    
    '���汨���ˢ����Ƶ�ɼ����ڵı���ͼ���
    If Not mobjWork_ImageCap Is Nothing Then
        Call mobjWork_ImageCap.zlRefreshData(True)
    End If
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub AfterPrintedRich(ByVal lngOrderID As Long, ByVal strDocId As String)
On Error GoTo errhandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strResultInput As String
    Dim bln���������� As Boolean
    Dim blnCriticalValues As Boolean
    Dim blnImageQuality As Boolean
    Dim blnReportQuality As Boolean
    Dim blnConformDetermine As Boolean
    Dim blnAllReportFinished As Boolean
    Dim intState As Integer, lngSendID As Long
    
    strResultInput = ""
    
    intState = getStudyStateRich(lngOrderID, strDocId, False, blnAllReportFinished, lngSendID, bln����������, blnCriticalValues, blnImageQuality, blnReportQuality, blnConformDetermine)
    If intState = 6 Then Exit Sub
    
    strSQL = "Select B.Σ��״̬, A.�������, B.Ӱ������, A.��������, B.�������,B.ҽ��ID " & _
                 "From Ӱ�񱨸��¼ A, Ӱ�����¼ B " & _
                 "Where A.ID=[1] and A.ҽ��id = B.ҽ��id"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������", strDocId)
    
'    If (Not blnCriticalValues And mSysPar.lngCriticalValues <> 0) Then strResultInput = "Σ��״̬|"    �����ڽ��������¼��Σ��ֵ
    If (Not bln���������� And mSysPar.blnIgnoreResult = False) Then strResultInput = strResultInput & "�������|"
    If (Not blnImageQuality And mSysPar.strImageLevel <> "") And mSysPar.lngImageLevel <> 0 And CheckPopedom(mstrPrivs, "Ӱ���ʿ�") Then strResultInput = strResultInput & "Ӱ������|"
    If (Not blnReportQuality And mSysPar.strReportLevel <> "") And mSysPar.lngReportLevel <> 0 And CheckPopedom(mstrPrivs, "�����ʿ�") Then strResultInput = strResultInput & "��������|"
    If (Not blnConformDetermine And mSysPar.lngConformDetermine <> 0) Then strResultInput = strResultInput & "�������|"
    
    If strResultInput <> "" Then Call PromptResultRich(lngOrderID, strDocId, mlngModule, Me, mlngCur����ID, strResultInput)
    
    If mSysPar.blnPrintCommit = True Then
        If blnAllReportFinished Then Call Menu_Manage_����������(lngOrderID, False, strDocId)
    End If
    
    Call UpdateQueryListData(Nothing, lngOrderID)
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub AfterPrinted(lngOrderID As Long)
On Error GoTo errhandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strResultInput As String
    
    If Not mSysPar.blnPDFTested Then Call TestPDF
    
    If mSysPar.blnPDFTested And mSysPar.strPDFFTPdevice <> "" And mrtReportType = ���Ӳ����༭�� Then
        Call CreateReportPDFAndUpLoad(lngOrderID, Me, mSysPar.strPDFFTPdevice)
    End If
    strResultInput = ""
    
    gstrSQL = "ZL_Ӱ�񱨸��ӡ_Update(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "���´�ӡ���"
    
    strSQL = "Select B.Σ��״̬, A.�������, B.Ӱ������, B.��������, B.������� " & _
             "From ����ҽ������ A, Ӱ�����¼ B " & _
             "Where A.ҽ��id = B.ҽ��id and B.ҽ��ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������", lngOrderID)
    
'    If IsNull(rsTemp!Σ��״̬) And mSysPar.lngCriticalValues <> 0 Then strResultInput = "Σ��״̬|"    '���ڱ�����������¼��Σ��ֵ
    If IsNull(rsTemp!�������) And Not mSysPar.blnIgnoreResult Then strResultInput = strResultInput & "�������|"
    If IsNull(rsTemp!Ӱ������) And mSysPar.strImageLevel <> "" And mSysPar.lngImageLevel <> 0 And CheckPopedom(mstrPrivs, "Ӱ���ʿ�") Then strResultInput = strResultInput & "Ӱ������|"
    If IsNull(rsTemp!��������) And mSysPar.strReportLevel <> "" And mSysPar.lngReportLevel <> 0 And CheckPopedom(mstrPrivs, "�����ʿ�") Then strResultInput = strResultInput & "��������|"
    If IsNull(rsTemp!�������) And mSysPar.lngConformDetermine <> 0 Then strResultInput = strResultInput & "�������|"

    If strResultInput <> "" Then Call PromptResult(lngOrderID, mlngModule, Me, mlngCur����ID, strResultInput)
    
    If mSysPar.blnPrintCommit = True Then
        Call Menu_Manage_����������(lngOrderID, False)
    End If
    
    Call UpdateQueryListData(Nothing, lngOrderID)
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub AfterReportSavedRich(ByVal lngOrderID As Long, ByVal strDocId As String, frmOwnerForm As Form, ByVal lngSaveType As Long)
'���汨��֮��Ĵ���
'ִ�й��̣�2-�ѱ�����3-�Ѽ�飻4-�ѱ��棻5-����ˣ�6-�����
On Error GoTo errhandle
    Dim intState As Integer, lngSendID As Long
    Dim strǩ�� As String
    Dim str������ As String
    Dim str������ As String
    Dim bln���������� As Boolean
    Dim blnCriticalValues As Boolean
    Dim blnImageQuality As Boolean
    Dim blnReportQuality As Boolean
    Dim blnConformDetermine As Boolean
    Dim arrSQL() As Variant
    Dim blnInTrans As Boolean
    Dim i As Integer
    Dim blnAllReportFinished As Boolean
    Dim lngID As Long
        
    arrSQL = Array()

    'Call mobjWork_Report.zlRefreshFace(True)
    
    'intState =1--�ѵǼǣ�2--�ѱ�����3--�Ѽ�飻4--�ѱ��棻5--����ˣ�6--����ɣ������̲������������ֵ��
    
    '��ȡ���μ���ִ�й���
    intState = getStudyStateRich(lngOrderID, strDocId, False, blnAllReportFinished, lngSendID, bln����������, blnCriticalValues, blnImageQuality, blnReportQuality, blnConformDetermine)
    If intState = 6 Then Exit Sub
    
    If intState = 4 And lngSaveType = 2 Then
    '������˺�
        '��մ�������
        Call Menu_Manage_SendAudit("")
    End If
    
    If intState = 2 Or intState = 3 Or intState = 4 Then
        '���汣��ʱִ�з���
        If (mlngModule = G_LNG_VIDEOSTATION_MODULE And mSysPar.lngVideoStationMoneyExeModle = 2) Or _
           (mlngModule = G_LNG_PATHSTATION_MODULE And mSysPar.lngPatholStationMoneyExeModle = 2) Or _
           (mlngModule = G_LNG_PACSSTATION_MODULE And mSysPar.lngPacsStationMoneyExeModle = 1) Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            
            If mblnAllDepts Then
                If mobjCurStudyInfo.lngExeDepartmentId > 0 Then
                    lngID = mobjCurStudyInfo.lngExeDepartmentId
                Else
                    lngID = 0
                End If
            Else
                lngID = mlngCur����ID
            End If
            
            gstrSQL = "Zl_Ӱ�����ִ��(" & lngOrderID & "," & lngSendID & ",4,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lngID & ")"
            arrSQL(UBound(arrSQL)) = gstrSQL
        End If
    End If
    
    gstrSQL = "Zl_Ӱ����_״̬����(" & lngOrderID & "," & lngSendID & ",'" & strDocId & "'," & intState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
                    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    gcnOracle.BeginTrans        '----------������״̬��������
    
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "������״̬��������")
    Next i
    
    gcnOracle.CommitTrans
    blnInTrans = False
    
    If (intState = 4 Or intState = 5) And IIf(mSysPar.lngHintType = 0, lngSaveType = 1, IIf(mSysPar.lngHintType = 1, lngSaveType = 2, False)) Then
        Dim strResultInput As String
        
        strResultInput = ""
        If mSysPar.blnReportWithResult Then '��Ӱ�����Ϊ����  -����ʾ�Զ����
            Call mobjWork_Report.Menu_Manage_�������(mobjCurStudyInfo.lngAdviceId, "0")
        End If
            
'        If (Not blnCriticalValues And mSysPar.lngCriticalValues <> 0) Then strResultInput = "Σ��״̬|"    '���ڱ�����������¼��Σ��ֵ
        If (Not bln���������� And mSysPar.blnIgnoreResult = False) Then strResultInput = strResultInput & "�������|"
        If (Not blnImageQuality And mSysPar.strImageLevel <> "") And mSysPar.lngImageLevel <> 0 And CheckPopedom(mstrPrivs, "Ӱ���ʿ�") Then strResultInput = strResultInput & "Ӱ������|"
        If (Not blnReportQuality And mSysPar.strReportLevel <> "") And mSysPar.lngReportLevel <> 0 And CheckPopedom(mstrPrivs, "�����ʿ�") Then strResultInput = strResultInput & "��������|"
        If (Not blnConformDetermine And mSysPar.lngConformDetermine <> 0) Then strResultInput = strResultInput & "�������|"
 
        If strResultInput <> "" Then Call PromptResultRich(lngOrderID, strDocId, mlngModule, frmOwnerForm, mlngCur����ID, strResultInput)
    End If
    
    '�������˺�ֱ����ɡ��������ֱ����ɡ�
    If (blnAllReportFinished And mSysPar.blnCompleteCommit) Or (intState = 5 And mSysPar.blnFinallyCompleteCommit) Then
        Call Menu_Manage_����������(lngOrderID, False, strDocId)
    End If
    
    '����״̬����
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    '����״̬ͬ����Ϣ
    Call mobjMsgCenter.Send_Msg_StateSync(lngOrderID)
    Exit Sub
errhandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub AfterReportSaved(lngOrderID As Long, frmOwnerForm As Form, ByVal lngSaveType As Long, ByVal isRefreshFace As Boolean)
'ִ�й��̣�2-�ѱ�����3-�Ѽ�飻4-�ѱ��棻5-����ˣ�6-�����
'------------------------------------------------
'���ܣ����汨��֮��Ĵ���
'������ lngOrderID -- ҽ��ID
'       frmOwnerForm -- ������ID
'       lngSaveType -- ��������, 0-��ͨ���棬1-���ǩ����2-���ǩ����3-�����޶� , 4-����ǩ��, 5-������ˣ�6-���������ǩ��ֱ�����ǩ��,7-���˲��������ǩ��ֱ�����ǩ��
'       isRefreshFace -- �Ƿ�ˢ�±������
'���أ�
'------------------------------------------------
On Error GoTo errhandle
    Dim intState As Integer, lngSendID As Long
    Dim strǩ�� As String
    Dim str������ As String
    Dim str������ As String
    Dim bln���������� As Boolean
    Dim blnCriticalValues As Boolean
    Dim blnImageQuality As Boolean
    Dim blnReportQuality As Boolean
    Dim blnConformDetermine As Boolean
    Dim arrSQL() As Variant
    Dim blnInTrans As Boolean
    Dim i As Integer
    Dim blnAllReportFinished As Boolean
    Dim lngID As Long
    Dim blnDoPDF As Boolean '����PDF�������ϴ�
    
    arrSQL = Array()
    blnDoPDF = False
    'ˢ�±������
    If isRefreshFace Then
        Call mobjWork_Report.zlRefreshFace(True)
    End If
    
    'intState =1--�ѵǼǣ�2--�ѱ�����3--�Ѽ�飻4--�ѱ��棻5--����ˣ�6--����ɣ������̲������������ֵ��

    '��ȡ���μ���ִ�й���
    intState = getStudyState(lngOrderID, False, lngSendID, str������, strǩ��, str������, bln����������, blnCriticalValues, blnImageQuality, blnReportQuality, blnConformDetermine)
        
    '���ǩ���ı��������մ�������
    If mintState = 4 Then
        If intState < 4 Then
            Call Menu_Manage_SendAudit("")
        End If
    End If
    mintState = intState
    
    '����ʱ���Ƿ�����Ҫ�Զ�ִ�еĲ������
    If lngSaveType = 0 Then
    '���汣���
        Call CheckExecuteInterface(EInterfaceExeTime.���汣���)
    ElseIf intState = 4 And lngSaveType = 1 Then
    '����ǩ����
        Call CheckExecuteInterface(EInterfaceExeTime.����ǩ����)
    ElseIf intState = 5 And lngSaveType = 2 Then
    '������˺�
        '��մ�������
        Call Menu_Manage_SendAudit("")
        
        If Not mSysPar.blnPDFTested Then Call TestPDF
        If mSysPar.blnPDFTested And mSysPar.strPDFFTPdevice <> "" And mrtReportType = ���Ӳ����༭�� Then blnDoPDF = True
        
        Call CheckExecuteInterface(EInterfaceExeTime.������˺�)
    ElseIf lngSaveType = 4 Then
    'ȡ��ǩ��ʱ
        Call CheckExecuteInterface(EInterfaceExeTime.ȡ��ǩ��ʱ)
    ElseIf lngSaveType = 5 Then
    'ȡ�����ʱ
        Call CheckExecuteInterface(EInterfaceExeTime.ȡ�����ʱ)
    ElseIf lngSaveType = 6 Then
    'ֱ�����
        
        Call CheckExecuteInterface(EInterfaceExeTime.����ǩ����)
        Call CheckExecuteInterface(EInterfaceExeTime.������˺�)
    ElseIf lngSaveType = 7 Then
    'ֱ����˻���ʱ
        Call CheckExecuteInterface(EInterfaceExeTime.ȡ�����ʱ)
        Call CheckExecuteInterface(EInterfaceExeTime.ȡ��ǩ��ʱ)
    End If
        
    '2--�ѱ�����3--�Ѽ��
    If intState = 2 Or intState = 3 Then
        gstrSQL = "Zl_Ӱ����_State(" & lngOrderID & "," & lngSendID & "," & intState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
        
        gstrSQL = "ZL_Ӱ�񱨸汣��_Update(" & lngOrderID & ",'" & IIf(mstrRPTExecutor <> "", mstrRPTExecutor, str������) & "','')"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
        
        '���汣��ʱִ�з���
        If (mlngModule = G_LNG_VIDEOSTATION_MODULE And mSysPar.lngVideoStationMoneyExeModle = 2) Or _
           (mlngModule = G_LNG_PATHSTATION_MODULE And mSysPar.lngPatholStationMoneyExeModle = 2) Or _
           (mlngModule = G_LNG_PACSSTATION_MODULE And mSysPar.lngPacsStationMoneyExeModle = 1) Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            
            If mblnAllDepts Then
                If mobjCurStudyInfo.lngExeDepartmentId > 0 Then
                    lngID = mobjCurStudyInfo.lngExeDepartmentId
                Else
                    lngID = 0
                End If
            Else
                lngID = mlngCur����ID
            End If
            gstrSQL = "Zl_Ӱ�����ִ��(" & lngOrderID & "," & lngSendID & ",4,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lngID & ")"
            
            arrSQL(UBound(arrSQL)) = gstrSQL
        End If
    Else
        If intState = 4 Then        '4--�ѱ���
            '���ǩ�������һ��ǩ��Ϊҽʦ,ִ�й���Ϊ�ѱ���
            '�п��ܵ���� 1-ҽʦ��N��ǩ�� 2-���μ������һ����ǩ 3-�޶�ģʽ�±���(ǩ������=0)
            gstrSQL = "Zl_Ӱ����_State(" & lngOrderID & "," & lngSendID & "," & intState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
            
            'Ӧ����д�����˲�׼ȷ�����˵�ʱ�򣬻��˵����Ǳ����ˣ����ǲ��Ǳ��洴����
            'ҽ�����ǩ��,�����ǵ�N�Σ���ʱ����������Ҫ���棬��������Ҫ���;
            gstrSQL = "ZL_Ӱ�񱨸汣��_Update(" & lngOrderID & ",'" & IIf(mstrRPTExecutor <> "", mstrRPTExecutor, str������) & "','')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
        ElseIf intState = 5 Then         '5--�����
            '���ǩ�������μ����ϼ���ǩ����ǩ������>=2,ִ�й���Ϊ�����
            gstrSQL = "Zl_Ӱ����_State(" & lngOrderID & "," & lngSendID & "," & intState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
            
            gstrSQL = "ZL_Ӱ�񱨸汣��_Update(" & lngOrderID & ",'" & IIf(mstrRPTExecutor <> "", mstrRPTExecutor, str������) & "','" & IIf(strǩ�� <> "", strǩ��, str������) & "')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
        End If
    End If
    
    '���±���ͼ���
    gstrSQL = "Zl_Ӱ����ͼ��_����ͼ(" & lngOrderID & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    gcnOracle.BeginTrans        '----------������״̬��������
    
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "������״̬��������")
    Next i
    
    gcnOracle.CommitTrans
    blnInTrans = False
    
    If blnDoPDF Then
        Call CreateReportPDFAndUpLoad(lngOrderID, Me, mSysPar.strPDFFTPdevice)
    End If
    '��ʾ���뱨�渽�ӽ���������Ե�
    '4--�ѱ��棻5--�����;lngHintType ��Ͻ����ʾ���ͣ�lngSaveType 1-���ǩ����2-���ǩ����6-���������ǩ��ֱ�����ǩ��
    
    If (intState = 4 Or intState = 5) And IIf(mSysPar.lngHintType = 0, lngSaveType = 1, IIf(mSysPar.lngHintType = 1, (lngSaveType = 2 Or lngSaveType = 6), False)) Then
        Dim strResultInput As String
        
        strResultInput = ""
        If mSysPar.blnReportWithResult Then '��Ӱ�����Ϊ����  -����ʾ�Զ����
            gstrSQL = "ZL_Ӱ����_���(" & lngOrderID & ",0)"
            zlDatabase.ExecuteProcedure gstrSQL, "���������"
        End If
            
'        If (Not blnCriticalValues And mSysPar.lngCriticalValues <> 0) Then strResultInput = "Σ��״̬|"    '���ڱ�����������¼��Σ��ֵ
        If (Not bln���������� And mSysPar.blnIgnoreResult = False) Then strResultInput = strResultInput & "�������|"
        If (Not blnImageQuality And mSysPar.strImageLevel <> "") And mSysPar.lngImageLevel <> 0 And CheckPopedom(mstrPrivs, "Ӱ���ʿ�") Then strResultInput = strResultInput & "Ӱ������|"
        If (Not blnReportQuality And mSysPar.strReportLevel <> "") And mSysPar.lngReportLevel <> 0 And CheckPopedom(mstrPrivs, "�����ʿ�") Then strResultInput = strResultInput & "��������|"
        If (Not blnConformDetermine And mSysPar.lngConformDetermine <> 0) Then strResultInput = strResultInput & "�������|"
 
        If strResultInput <> "" Then Call PromptResult(lngOrderID, mlngModule, frmOwnerForm, mlngCur����ID, strResultInput)
    End If
    
    If intState = 5 And mSysPar.blnCompleteCommit Then   '�������˺�ֱ����ɡ�
        Call Menu_Manage_����������(lngOrderID, False)
    End If
    '����״̬����
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    '���汨���ˢ����Ƶ�ɼ����ڵı���ͼ���
    If Not mobjWork_ImageCap Is Nothing Then
        Call mobjWork_ImageCap.zlRefreshData(True)
        mobjWork_ImageCap.IsReported = mobjCurStudyInfo.blnIsReported   '�ѱ���
    End If
    
    '����״̬ͬ����Ϣ
    Call mobjMsgCenter.Send_Msg_StateSync(lngOrderID)
    
    Exit Sub
errhandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub UpdateStudyListState(lngAdviceId As Long, strStudyUID As String, blnAddImage As Boolean, blnStateChanged As Boolean)
On Error GoTo errH
    Dim strSQL As String   '�����ϰ�Ĵ���ͼ��ɼ�״̬��صĲ�����ˢ��ѡ���У����ݿ�ˢ�£�
    Dim intRowIndex As Integer
    Dim rsData As ADODB.Recordset
    Dim lngCol As Long

    With vsfList

        intRowIndex = .FindRow(lngAdviceId, , .ColIndex("ҽ��ID"))
        '�������ø���Ӱ���鼼ʦ
        If mSysPar.blnWriteCapDoctor = True And blnStateChanged = True Then
            If mblnCnOracleIsHIS Then
                strSQL = "Zl_Ӱ����_��鼼ʦ( " & lngAdviceId & ",'" & IIf(blnAddImage = True, mstrUserNameNew, "") & "')"
            Else
                strSQL = "Zl_Ӱ����_��鼼ʦ( " & lngAdviceId & ",'" & IIf(blnAddImage = True, mstrUserNameHIS, "") & "')"
            End If

            zlDatabase.ExecuteProcedure strSQL, GetWindowCaption
        End If
        
        If blnStateChanged Then
            Call UpdateQueryListData(Nothing, lngAdviceId)
        End If
        
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ShowBillList(objPopup As CommandBarPopup) As Boolean
'���ܣ���ʾ��ǰִ��ҽ�����Դ�ӡ�����Ƶ����ڲ˵���
    Dim rsTmp As New ADODB.Recordset
    Dim objControl As CommandBarControl
    Dim strSQL As String
        
    On Error GoTo errH
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Function
    End If
    
    If Not objPopup Is Nothing Then objPopup.CommandBar.Controls.DeleteAll
    
    strSQL = "Select Distinct C.���,C.����,C.˵��" & _
        " From ����ҽ����¼ A,��������Ӧ�� B,�����ļ��б� C" & _
        " Where A.ID=[1] And A.���ID IS NULL" & _
        " And A.������ĿID=B.������ĿID" & _
        " And B.Ӧ�ó���=[2] And B.�����ļ�ID=C.ID And C.����=7" & _
        " Order by C.���"
        
    If mobjCurStudyInfo.blnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngPatientFrom)
    
    If Not rsTmp.EOF Then
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Manage_RequestPrint * 10# + 1, rsTmp!���� & "(&0)")
            objControl.Parameter = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1" '��Ӧ���Զ��屨����
            objControl.Category = M_STR_MODULE_MENU_TAG
        End With
        cbrMain.KeyBindings.Add 0, vbKeyF10, conMenu_Manage_RequestPrint * 10# + 1
    End If
    
    ShowBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub FuncBillPrint(objControl As CommandBarControl)
'���ܣ���ӡ���Ƶ���
On Error GoTo errhandle
    If objControl.Parameter = "" Then '��֣�ֱ�Ӱ�F10ʱ����һ���յ�Control
        Set objControl = cbrMain.FindControl(, conMenu_Manage_RequestPrint * 10# + 1, , True)
        If objControl Is Nothing Then Exit Sub
    End If
    
    If objControl.Parameter = "" Then Exit Sub
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If ReportPrintSet(gcnOracle, glngSys, objControl.Parameter, Me) Then
        Call ReportOpen(gcnOracle, glngSys, objControl.Parameter, Me, "NO=" & mobjCurStudyInfo.strNO, _
                       "����=" & mobjCurStudyInfo.lngRecordKind, "ҽ��ID=" & mobjCurStudyInfo.lngAdviceId, 1)
    End If
    Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub NotificationAllModuleRefresh()
'֪ͨ����ģ��ˢ��
    If Not mobjWork_His Is Nothing Then Call mobjWork_His.NotificationRefresh(hmAll)
    If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(mtAll)
    If Not mfrmWork_PacsImg Is Nothing Then Call mfrmWork_PacsImg.NotificationRefresh
    If Not mobjWork_ImageCap Is Nothing Then Call mobjWork_ImageCap.zlNotifyRefresh
    If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.NotificationRefresh
End Sub

Private Sub NotificationImageCapRefresh()
'֪ͨ�ɼ�ģ��ˢ�£���Ҫ��ˢ�±���ͼ���
    If Not mobjWork_ImageCap Is Nothing Then
        Call mobjWork_ImageCap.zlNotifyRefresh
    End If
End Sub

Private Sub DisableWorkModule()
'���ù���ģ��
    tcDisable.Visible = True
    tcDisable.Translucence
End Sub


Private Sub EnableWorkModule()
'�򿪹���ģ��
    tcDisable.Visible = False
End Sub

Public Sub RefreshList()
'blClick �Ƿ���ˢ�´�����ˢ���б�
'ˢ�������б�
    Dim i As Integer
    
    If mblnIsLoading = True Then
        MsgBoxD Me, "���ݼ����У����Ժ�����...", vbInformation, Me.Caption
        Exit Sub
    End If
    
    mblnIsLoading = True

On Error GoTo errhandle

    mblnAutoRefreshList = True
        
    Call mobjPacsQueryWrap.ExecuteQuery(C_QUERY_ˢ��)
    
    If mobjPacsQueryWrap.SqlScheme.AutoRefreshTimeLen > 0 Then TimerRefresh.Enabled = True
    
    If mobjCurStudyInfo.lngAdviceId > 0 Then
'        RefreshDisPlay���������� 2��ʾ���²���
        Call mobjPacsQueryWrap.RefreshDisplay(vsfList.Row, mobjCurStudyInfo.lngAdviceId, 2)
    End If
    
    If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.SetblHaveReport
    
    'ֱ�ӿ�ʼ��λ
    If vsfList.Rows <= 1 Then
        '��û������ʱ��֪ͨˢ�¹���ģ������ص�����
        
        Set mobjCurStudyInfo = GetNullAdviceInf
        
        Call RefreshModuleAdviceInf
        Call NotificationAllModuleRefresh

        If TabWindow.Selected Is Nothing Then
            'ѡ���һ������ģ��
            For i = 0 To TabWindow.ItemCount - 1
                If TabWindow.Item(i).Visible Then
                    TabWindow(i).Selected = True

                    mblnAutoRefreshList = False
                    Exit For
                End If
            Next i
        End If

        Call RefreshTabWindow

        mblnAutoRefreshList = False
        mblnIsLoading = False

        Exit Sub
    End If

    
    mblnAutoRefreshList = False
    mblnIsLoading = False

    Exit Sub
errhandle:
    mblnIsLoading = False
End Sub

Private Sub picDetail_Resize()
On Error Resume Next
    Dim i As Integer, j As Integer, k As Integer
    Dim lngLeft As Long
    Dim intCnt As Integer

    intCnt = imgState.Count
    
    For i = 0 To intCnt - 1
        '��������λ��
        lngLeft = 0

        For k = i To 0 Step -1
            lngLeft = lngLeft + imgState(k).Width
        Next

        lngLeft = picDetail.Width - lngLeft
        Call imgState(i).Move(lngLeft, C_LAYOUT_BASEHEIGHTOFDETAILINFO - GetMaxImgHeight - 30)
    Next
End Sub

Private Sub PicFollowHistory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngPicHistoryX = X
    mlngPicHistoryY = Y
    mlngpicHistoryOldW = PicFollowHistory.Width
    mlngpicHistoryOldH = PicFollowHistory.Height
    TimerHistory.Enabled = False
    
    If PicFollowHistory.MousePointer = vbSizeNWSE Then
        '����ԤԼ��ǩ���
        mblnpicHistoryMoving = True
    End If
End Sub

Private Sub PicFollowHistory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errH
    If mblnpicHistoryMoving And Button = 1 Then
        Call MoveCtrHistroyFollow(X, Y)
    Else
        If X > PicFollowHistory.Width - 100 Then
            PicFollowHistory.MousePointer = vbSizeNWSE
        Else
            PicFollowHistory.MousePointer = vbDefault
        End If
    End If
    Exit Sub
errH:
End Sub

Private Sub PicFollowHistory_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnpicHistoryMoving = False
    TimerHistory.Enabled = True
End Sub

Private Sub PicLine_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandle
errhandle:
End Sub

Private Sub picLine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���·���ϸ��Ϣ�߶ȿ��Ըı�
On Error GoTo errhandle
    Dim i As Integer
    
    'Y��ֵ �����ƶ�Ϊ��ֵ �����ƶ�Ϊ��ֵ
    If Button = 1 Then
        '��ֵ�ﵽһ����Χ���˳�����

        If Y > 0 Then
        '�����϶����ж�
        ElseIf Y < 0 Then
        '�����϶����жϣ���Ҫ�����������б�ͷ��������500
            If (PicLine.Top + Y) - vsfList.Top < 800 Then
                Exit Sub
            End If
        Else
        'Y=0
        End If

        PicLine.Top = PicLine.Top + Y
        picDetail.Top = picDetail.Top + Y
        TabExtra.Top = TabExtra.Top + Y

        vsfList.Height = vsfList.Height + Y
        TabExtra.Height = TabExtra.Height - Y

        mlngMove = TabExtra.Height - C_LAYOUT_BASEHEIGHTOFTAB

        If Not mobjPacsQueryWrap Is Nothing Then
            For i = vsfList.TopRow To vsfList.BottomRow
                Call mobjPacsQueryWrap.RefreshRowRelation(i)
            Next
        End If
    End If

errhandle:
End Sub

Private Sub picLine_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandle
    Call AdjustFace(picList.Height, picList.Width)
errhandle:
End Sub

Private Sub picList_Resize()
On Error GoTo errhandle
    If picList.Width < 1000 Then picList.Width = 1000
    Call AdjustFace(picList.Height, picList.Width)
errhandle:
End Sub

Private Sub picLoadState_Resize()
On Error GoTo errhandle
    labLoadState.Left = Fix((picLoadState.Width - labLoadState.Width) / 2)
    labLoadState.Top = Fix((picLoadState.Height - labLoadState.Height) / 2)
    
    picSmile.Left = labLoadState.Left - picSmile.Width
    picSmile.Top = labLoadState.Top - 80
    
errhandle:
End Sub

Private Sub cmdLocate_Click()
On Error GoTo errhandle
    mobjPacsQueryWrap.DefaultLocate = True
    
    cmdLocate.BackColor = IIf(mobjPacsQueryWrap.DefaultLocate, &HFF00&, &H8000000F)
    cmdFind.BackColor = IIf(mobjPacsQueryWrap.DefaultLocate = False, &HFF00&, &H8000000F)
    
    If Me.MousePointer = 0 Then
        Me.MousePointer = 13
        Call mobjPacsQueryWrap.Find(False, True)
        TimerRefresh.Enabled = False
        Me.MousePointer = 0
    Else
        Exit Sub
    End If
    Exit Sub
errhandle:
    MsgBoxD Me, "��λ������������Ϣ��" & err.Description, vbInformation, Me.Caption
End Sub

Private Sub picReportContainer_Resize()
On Error GoTo errhandle
    
    If mobjWork_Report Is Nothing Then Exit Sub
    
    Call mobjWork_Report.UpdateSize
    
errhandle:
End Sub

Private Sub picWindow_Resize()
On Error GoTo errhandle
    With TabWindow
        If GetWorkModuleCount = 1 Then
            TabWindow.PaintManager.ClientMargin.Top = -30
        Else
            TabWindow.PaintManager.ClientMargin.Top = 0
        End If
        .Top = 0
        .Left = 0
        .Width = picWindow.ScaleWidth
        .Height = picWindow.ScaleHeight + IIf(GetWorkModuleCount = 1, ScaleY(30, vbTwips, vbPixels), 0)
    End With
    
    tcDisable.Left = 0
    tcDisable.Top = IIf(TabWindow.PaintManager.ClientMargin.Top < 0, 0, IIf(gbytFontSize = 9, 440, 470))
    tcDisable.Width = picWindow.ScaleWidth
    tcDisable.Height = picWindow.ScaleHeight - IIf(TabWindow.PaintManager.ClientMargin.Top < 0, 0, IIf(gbytFontSize = 9, 440, 470))
errhandle:
End Sub

Private Sub ConfigSubForm(ByVal Item As XtremeSuiteControls.ITabControlItem)
'�����Ӵ��ڽ���
On Error GoTo errhandle
    Dim lngIndex As Integer
    Dim objItem As XtremeSuiteControls.TabControlItem
    
    If mblnLoadSubFrom Then Exit Sub
    If Item.Handle <> picTemp.hwnd Then Exit Sub
    
    mblnLoadSubFrom = True
    lngIndex = Item.Index
    
    Set objItem = Nothing
    
    Select Case Item.tag
        Case "Ӱ��ͼ��"
            Set objItem = TabWindow.InsertItem(lngIndex, "Ӱ���¼", mfrmWork_PacsImg.hwnd, Item.Image)
                
        Case "�걾����"
            Set objItem = TabWindow.InsertItem(lngIndex, "�걾����", mobjWork_Pathol.GetModule(mtSpecimen).hwnd, Item.Image)

        Case "����ȡ��"
            Set objItem = TabWindow.InsertItem(lngIndex, "����ȡ��", mobjWork_Pathol.GetModule(mtMaterial).hwnd, Item.Image)
            
        Case "������Ƭ"
            Set objItem = TabWindow.InsertItem(lngIndex, "������Ƭ", mobjWork_Pathol.GetModule(mtSlices).hwnd, Item.Image)
            
        Case "�����ؼ�"
            Set objItem = TabWindow.InsertItem(lngIndex, "�����ؼ�", mobjWork_Pathol.GetModule(mtSpeExam).hwnd, Item.Image)
        
        Case "���̱���"
            Set objItem = TabWindow.InsertItem(lngIndex, "����/�ؼ챨��", mobjWork_Pathol.GetModule(mtProRep).hwnd, Item.Image)
            
        Case "�������"
            If mobjAppendBill Is Nothing Then
                Set objItem = TabWindow.InsertItem(lngIndex, "���ü�¼", mobjWork_His.GetModule(hmExpense).hwnd, Item.Image)
            End If
        Case "סԺҽ��"
            Set objItem = TabWindow.InsertItem(lngIndex, "ҽ����¼", mobjWork_His.GetModule(hmInAdvice).hwnd, Item.Image)
            
        Case "����ҽ��"
            Set objItem = TabWindow.InsertItem(lngIndex, "ҽ����¼", mobjWork_His.GetModule(hmOutAdvices).hwnd, Item.Image)
            
        Case "סԺ����"
            Set objItem = TabWindow.InsertItem(lngIndex, "������¼", mobjWork_His.GetModule(hmInEPRs).hwnd, Item.Image)
            
        Case "���ﲡ��"
            Set objItem = TabWindow.InsertItem(lngIndex, "������¼", mobjWork_His.GetModule(hmOutEPRs).hwnd, Item.Image)
           
        Case "������Ӳ���", "סԺ���Ӳ���"
            Set objItem = TabWindow.InsertItem(lngIndex, "���Ӳ���", mobjWork_His.GetModule(hmEMR).hwnd, Item.Image)
              
        Case "�Ŷӽк�"
            Set objItem = TabWindow.InsertItem(lngIndex, "�Ŷӽк�", mobjQueue.hwnd, Item.Image)
            
        Case "Ӱ��ɼ�", "������д"
            '���ﲻ���д���
    End Select
    
    Call RefreshModuleAdviceInf
    
    If Not objItem Is Nothing Then
        objItem.tag = Item.tag
        objItem.Selected = True
        
        Call TabWindow.RemoveItem(lngIndex + 1)
    End If
    
    mblnLoadSubFrom = False
Exit Sub
errhandle:
    If Not objItem Is Nothing Then
        If objItem.tag = "" Then
            Call TabWindow.RemoveItem(objItem.Index)
        End If
    End If
    
    mblnLoadSubFrom = False
End Sub

Private Sub rftHistoryFollow_LostFocus()
    TimerHistory.Enabled = True
End Sub

Private Sub TabExtra_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not mobjPacsQueryWrap Is Nothing Then
        cmdMore.Visible = mobjPacsQueryWrap.CurPacsQuery.IsMoreEmbedInput
        
        If Not cmdMore.Visible Then
            Call cmdClear.Move(cmdDo.Left, cmdClear.Top, cmdDo.Width)
            cmdClear.Width = cmdDo.Width
        Else
            Call cmdClear.Move(cmdDo.Left, cmdClear.Top, 0.5 * cmdDo.Width)
        End If
        Call cmdMore.Move(cmdClear.Left + cmdClear.Width)
    End If
End Sub

Private Sub tabScheme_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errH
        If Not mblnInitOk Then Exit Sub
    Call ChangeScheme(Item.Caption, Val(Item.tag), False)
    Exit Sub
errH:
    MsgBoxD Me, "�л����������쳣,��ϸ��Ϣ��" & err.Description, vbInformation, Me.Caption
End Sub

Private Sub TabWindow_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errhandle
    Dim intStyle As Integer
    Dim blnVisible As Boolean
    Dim blnLargeIcon As Boolean
    Dim cbrControl As CommandBarControl
    'LSQ Debug
    Call ConfigSubForm(Item)

    If Not mblnInitOk Then Exit Sub

    Call ReSetModuleFontSize(gbytFontSize, IIf(gbytFontSize = 9, 0, 1))

    If Not mobjWork_Report Is Nothing And Item.tag = "������д" Then
        Call mobjWork_Report.AllowLocate(True)
    End If

    mblnRefreshWord = True
    Call RefreshTabWindow

    'ˢ���Ŷӽк�ģ�����ݣ�����Ѿ����ò�����ѡ����Ŷӽк�ҳ��
    If Trim(Item.tag) = "�Ŷӽк�" Then
        Call RefreshPacsQueueData(False)
    End If

    Call LockWindowUpdate(Me.hwnd)

    '�еĲ˵���ֻ�ڹ���ģ����ʾ��ʱ�� ����ʾ
    Call CreateWorkModuleMenu

    If mobjCurStudyInfo.lngAdviceId <> 0 Then
        '��ʾ�ɴ�ӡ�����Ƶ���:֮���Լ�ʱ����,��Ϊ��ʹ��F2�ȼ�
        Call ShowBillList(cbrMain.FindControl(, conMenu_Manage_RequestPrint, , True))
    End If

    Call LockWindowUpdate(0)

    Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub timerCapture_Timer()
On Error GoTo errhandle
    Dim strKeyAlias As String
    
    If Not mblnInitOk Then Exit Sub
    timerCapture.Enabled = False
    
    strKeyAlias = GetKeyAlias(mCaptureMsg.lngMsg, 0)
    
    'ʹ���ȼ����вɼ�
'    If strKeyAlias = mstrCaptureHot Then
    If strKeyAlias = mstrCaptureHot Then
        If Not mobjWork_ImageCap Is Nothing Then
            Call mobjWork_ImageCap.zlCaptureImg
        End If

    'ʹ���ȼ����к�̨�ɼ�
    ElseIf strKeyAlias = mstrCaptureAfterHot Then
        If Not mobjWork_ImageCap Is Nothing Then
            Call mobjWork_ImageCap.zlCaptureAfterImg
        End If
    
    'ʹ���ȼ����б�Ǹ���
    ElseIf strKeyAlias = mstrCaptureAfterTagHot Then
        If Not mobjWork_ImageCap Is Nothing Then
            Call mobjWork_ImageCap.zlUpdateAfterCaptureInfo
        End If
    End If
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume

End Sub

Private Sub timerRefresh_Timer()
On Error GoTo errhandle
    'ˢ�²����б�
    Dim blNeedColStatistics As Boolean
    
    If Not mblnInitOk Then Exit Sub
    If Not Me.Visible Then Exit Sub
    If mobjPacsQueryWrap Is Nothing Then Exit Sub
    
    blNeedColStatistics = False

    If mintAutoRefreshTimerCount > 1 Then
        mintAutoRefreshTimerCount = mintAutoRefreshTimerCount - 1
        Exit Sub
    Else
        mintAutoRefreshTimerCount = mintAutoRefreshTimer
        TimerRefresh.Enabled = False
        
        Call RefreshList
        TimerRefresh.Enabled = True
    End If
    
errhandle:
End Sub


Private Sub ChangeUser()
    Dim strPrivs As String
    Dim strUserID As String
    
    frmTwoUser.intDBState = mintChangeUserState
    frmTwoUser.strUserNameHIS = mstrUserNameHIS
    frmTwoUser.strUserIDHIS = mstrUserIDHIS
    frmTwoUser.Show 1, Me
    
    If frmTwoUser.blnOk = True Then
        If frmTwoUser.intDBState = 1 Then   'ͳһ����ָ���HISԭ�������ݿ����Ӻ��û���
            mstrUserNameNew = mstrUserNameHIS
            mstrUserIDNew = mstrUserIDHIS
            mblnCnOracleIsHIS = True
            mintChangeUserState = 1
            Set gcnOracle = mcnOracleHIS
            
            
            InitCommon gcnOracle
            
            SetDbUser mstrUserIDHIS
'            RegCheck
            
            Call GetUserInfo
            
            strPrivs = ";" & GetPrivFunc(100, mlngModule) & ";"      'Ӱ��ɼ�����վ
            
            Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
            Call mobjWork_Report.zlInitModule(mlngModule, strPrivs, mlngCur����ID, Me)
        ElseIf frmTwoUser.intDBState = 2 Then   '�������򽻻����ݿ�����
            '�����ʹ�������ݿ����ӣ��ȼ��Ȩ��
            mstrUserNameNew = frmTwoUser.strUserNameNew
            mstrUserIDNew = frmTwoUser.strUserIDNew
            mintChangeUserState = 2
            If frmTwoUser.blnCnOracleIsNew = True Then
                Set gcnOracle = frmTwoUser.cnOracle
                mblnCnOracleIsHIS = False
                
                '��ʼ��zlComLib������ȷ��GetPrivFunc��ȡ������ȷ����Ϣ
                InitCommon gcnOracle
'                RegCheck
                SetDbUser mstrUserIDNew
                
                '�����û�Ȩ��
                strPrivs = GetPrivFunc(100, mlngModule)       'Ӱ��ɼ�����վ
                If strPrivs = "" Then
                    MsgBoxD Me, "�㲻�߱�ʹ�á�Ӱ��ɼ�����վ��ģ���Ȩ�ޣ�", vbInformation, gstrSysName
                    
                    '�л���ԭ�����û�
                    Set gcnOracle = mcnOracleHIS
                    
                    InitCommon gcnOracle
'                    RegCheck
                    SetDbUser mstrUserIDHIS
                
                    mstrUserNameNew = mstrUserNameHIS
                    mstrUserIDNew = mstrUserIDHIS
                    mblnCnOracleIsHIS = True
                    mintChangeUserState = 1
                End If
                
                strPrivs = GetPrivFunc(100, 1258)       '���Ʊ������
                If strPrivs = "" Then
                    MsgBoxD Me, "�㲻�߱�ʹ�á����Ʊ��桱ģ���Ȩ�ޣ�", vbInformation, gstrSysName
                    
                    '�л���ԭ�����û�
                    Set gcnOracle = mcnOracleHIS
                    
                    InitCommon gcnOracle
'                    RegCheck
                    SetDbUser mstrUserIDHIS
                    
                    mstrUserNameNew = mstrUserNameHIS
                    mstrUserIDNew = mstrUserIDHIS
                    mblnCnOracleIsHIS = True
                    mintChangeUserState = 1
                End If
            Else
                Set gcnOracle = mcnOracleHIS
                
                InitCommon gcnOracle
'                RegCheck
                SetDbUser mstrUserIDHIS
                
                mblnCnOracleIsHIS = True
            End If
            
            Call GetUserInfo
            Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
            
            strPrivs = ";" & GetPrivFunc(100, mlngModule) & ";"       'Ӱ��ɼ�����վ
            Call mobjWork_Report.zlInitModule(mlngModule, strPrivs, mlngCur����ID, Me)
        End If
        
    End If
    
    If mblnCnOracleIsHIS Then
        Me.stbThis.Panels(4).Text = "����ҽ����" & mstrUserNameHIS & "   ���ҽ����" & mstrUserNameNew
    Else
        Me.stbThis.Panels(4).Text = "����ҽ����" & mstrUserNameNew & "   ���ҽ����" & mstrUserNameHIS
    End If
End Sub

Private Sub SwitchUser()
'��ȡ���û�Ȩ��˵����ʹ�� GetPrivFuncByUser ���ұ�֤strDBUser������gstrDBUser��һ���������õ���¼�û�Ȩ�ޣ����� GetPrivFuncByUser ��Ҫ����SetDbUser ֮ǰ
'���� InitCommon ��ִ�� SetDbUser
'����114781�Ķ��㣺�޸��ж��Ƿ��л������û����߼����л��û�������mstrPrivs��ֵ����
    Dim strPrivs As String

    Call frmSwitchUser.SetModule(mlngModule)
    frmSwitchUser.Show 1, Me

    If frmSwitchUser.blnOk Then
'        �����ʹ�������ݿ����ӣ��ȼ��Ȩ��
        mstrUserNameNew = frmSwitchUser.strUserNameNew
        mstrUserIDNew = frmSwitchUser.strUserIDNew

        Set gcnOracle = frmSwitchUser.mcnOracle
        mblnCnOracleIsHIS = False

        If gstrDBUser <> mstrUserIDNew Then
            mstrPublicAdvicePrivs = "-1"

            mstrPrivs = ";" & GetPrivFuncByUser(100, mlngModule, mstrUserIDNew) & ";"
            
            InitCommon gcnOracle
            gstrDBUser = mstrUserIDNew
            
            Call GetUserInfo
            Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)

            Call mobjWork_Report.zlInitModule(mlngModule, mstrPrivs, mlngCur����ID, Me)
        
            Call ReCreatCbrMenu(cbrMain)
        End If

        Me.stbThis.Panels(4).Text = "����ҽ����" & mstrUserNameNew & "   ���ҽ����" & mstrUserNameNew
    End If

End Sub

Private Sub Menu_Manage_���()
On Error GoTo errhandle
    Dim strReview As String
    Dim strDeptName As String

    If mobjCurStudyInfo.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    strDeptName = Split(mstrCur����, "-")(1)
    If frmReview.ShowMe(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, Me, strDeptName, strReview) = True Then
            
        mobjCurStudyInfo.strFollowUpDescribe = strReview
        Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
        
    End If

Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_���淢��()
'���淢��
On Error GoTo errhandle
    Dim strSQL As String

    If mobjCurStudyInfo.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mrtReportType = �����ĵ��༭�� Then
        Call mobjWork_Report.Menu_Manage_���淢��(mobjCurStudyInfo.lngAdviceId, IIf(mobjWork_Report.GetReportReleaseState(mobjCurStudyInfo.lngAdviceId) > 1, 0, 1))
    Else
        strSQL = "Zl_Ӱ�񱨸淢��(" & mobjCurStudyInfo.lngAdviceId & ",'" & UserInfo.���� & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "���淢��")

        mobjCurStudyInfo.intReportGiveOut = IIf(mobjCurStudyInfo.intReportGiveOut = 1, 0, 1)
        Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
            
    End If
    
    Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_��Ƭ����()
'��Ƭ����
On Error GoTo errhandle
    Dim strSQL As String

    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    strSQL = "Zl_Ӱ��Ƭ����(" & mobjCurStudyInfo.lngAdviceId & ",'" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "��Ƭ����")
    
    mobjCurStudyInfo.intFilmGiveOut = IIf(mobjCurStudyInfo.intFilmGiveOut = 1, 0, 1)
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)

    Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_ReportExecutor()
    Dim strSQL As String
    
    Dim strRPTExecutor As String
On Error GoTo errhandle
    strRPTExecutor = frmSelectRPTExecutor.GetRPTExecutor(mlngCur����ID, Me, mstrRPTExecutor)
    
    If strRPTExecutor <> "" Then
        '���±�����
        strSQL = "ZL_Ӱ�񱨸汣��_���±�����(" & mobjCurStudyInfo.lngAdviceId & ",'" & strRPTExecutor & "')"
        Call zlDatabase.ExecuteProcedure(CStr(strSQL), "���±�����")
        
        'ˢ�¶�Ӧ���ı�����
        mstrRPTExecutor = strRPTExecutor
        
        mobjCurStudyInfo.strReportDoctor = strRPTExecutor
        Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)

        If Not mobjWork_Report Is Nothing And mrtReportType = �����ĵ��༭�� Then Call mobjWork_Report.SetDocCreator(mstrRPTExecutor)
        
        stbThis.Panels(4).Text = "����ҽ����" & strRPTExecutor & "   ���ҽ����" & Split(stbThis.Panels(4).Text, "���ҽ����")(1)
    End If
    
    Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_SendAudit(strName As String)
    Dim strSQL As String

    On Error GoTo errhandle
    
    If mobjCurStudyInfo.lngAdviceId > 0 Then
        strSQL = "Zl_Ӱ�����¼_�����������(" & mobjCurStudyInfo.lngAdviceId & ",'" & strName & "')"
        zlDatabase.ExecuteProcedure strSQL, "�����������"
        
        If Len(Trim(strName)) > 0 Then
            Call MsgBoxD(Me, "�ɹ����͵�����ˡ�" & strName & "����", vbInformation, "��ʾ")
        End If
    Else
        Call MsgBoxD(Me, "����ѡ��һ����顣", vbInformation, "��ʾ")
        Exit Sub
    End If
    
    'ͬ��ˢ�¼���б�
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub timerVideoEvent_Timer()
On Error GoTo errhandle
    timerVideoEvent.Enabled = False
    TimerRefresh.Enabled = False
    
    Call DoOnStateChange(mVideoEventInf.vetEventType, mVideoEventInf.lngAdviceId, mVideoEventInf.lngSendNo, mVideoEventInf.strOtherInf, mVideoEventInf.dcmImage)
    If mobjPacsQueryWrap.SqlScheme.AutoRefreshTimeLen > 0 Then TimerRefresh.Enabled = True
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume

End Sub

Private Function GetStudyNumberDisplayName() As String
'��ȡ��������ʾ����
    GetStudyNumberDisplayName = IIf(mlngModule = G_LNG_PATHOLSYS_NUM, "�����", "����")
End Function


Private Sub vsfList_OnDblClick()
On Error GoTo errhandle
    If mobjCurStudyInfo.lngAdviceId <> 0 Then
        '˫�����˼���б�ʱ��������˼��״̬Ϊ �Ѿܾ���Ŀǰ�����κδ���
        If mobjCurStudyInfo.strStuStateDesc = "�Ѿܾ�" Then Exit Sub
        
        Select Case mobjCurStudyInfo.intStep
            Case 1, 0
                Call Menu_Manage_����
            Case 2, 3               '˫������д����,�����ʱ�����趨�Ƿ�򿪹�Ƭվ
                Call Menu_RichEPR(conMenu_Edit_Modify)
            Case -1, 4, 5               '˫���޶�����,�����ʱ�����趨�Ƿ�򿪹�Ƭվ
                Call Menu_RichEPR(conMenu_Edit_Audit)
            Case 6                  '����
                Call Menu_RichEPR(conMenu_File_Open)
        End Select
    End If
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function GetNullAdviceInf() As clsStudyInfo
    Dim ObjClsStudyInfo As New clsStudyInfo

    With ObjClsStudyInfo
        .lngPatId = 0
        .strPatientName = ""
        .lngPatDept = 0
        .lngAdviceId = 0
        .lngUnit = 0
        .lngSendNo = 0
        .strStudyUID = ""
        .blnCanPrint = False
        .blnIsInsidePatient = False
        .blnMoved = False
        .intState = -1
        .intStep = -1
        .strRegNo = ""
'        .lngRegId = 0
        .lngExeDepartmentId = 0
'        .strExeRoom = ""
        .lngPatientFrom = 0
        .strDoDoctor = ""
        .strStudyNum = ""
'        .strBedNum = ""
        .strMarkNum = "0"
        .lngBaby = 0
        .strPatientSex = ""
        .strPatientAge = ""
        .strNO = ""
'        .lngRecordKind = 0
        .intFilmGiveOut = 0
        .intReportGiveOut = 0
        .strAdviceContext = ""
        .strAdviceDepartAndMethod = ""
        .strStuStateDesc = ""
        .blnIsTechincalSure = False
        .intDangerState = 0
        .intEmergentTag = 0
        .intGreenChannel = 0
        .blnInfancy = False
    End With
    
    Set GetNullAdviceInf = ObjClsStudyInfo
End Function

Private Function GetScanRequestCount(ByVal lngAdviceId As Long) As Long
'��ȡɨ�����뵥������
On Error GoTo errhandle
    Dim lngCount As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    GetScanRequestCount = 0
    
    '����������뵥ɨ����� ��ѡ������ִ�в�ѯ�õ����뵥ͼ��������δ��ѡ�� ��ִ��
    If mSysPar.blnIsPetitionScan Then
        '����ҽ��ID��ѯ Ӱ�����뵥ͼ����õ���ɨ������ ����ҽ����������� VSList
        strSQL = "select count(*) as ͼ���� from Ӱ�����뵥ͼ�� where ҽ��ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�õ�ͼ������", lngAdviceId)
        
        lngCount = Val(rsTemp!ͼ����)
    Else
        lngCount = 0
    End If
    
    GetScanRequestCount = lngCount
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowTab()
'���ݲ�����Դ���Ʋ�����ҽ��ѡ�
On Error GoTo errhandle
    Dim i As Integer
    Dim intDefaultIndex As Integer
    Dim blnShowReport As Boolean
    Dim strFirstTab As String
    
    If TabWindow.ItemCount <= 0 Then Exit Sub
    
    blnShowReport = False
     
    If Not mblnIsHistory Then '-------------------------------------------�б�ѡ�����
        '�ж� ��ͼ����д����
        blnShowReport = True
        
        If mSysPar.blnReportWithImage = True Then
            If mobjCurStudyInfo.strStudyUID = "" Then blnShowReport = False
        End If
    End If
    
    If mobjCurStudyInfo.lngPatientFrom <> 2 Then '���ݲ�����Դ���Ʋ�����ҽ��ѡ�
        For i = 0 To TabWindow.ItemCount - 1
            Select Case TabWindow(i).tag
                Case "���ﲡ��", "����ҽ��"
                    TabWindow(i).Visible = True
                    
                Case "סԺ����", "סԺҽ��"
                    TabWindow(i).Visible = False
                    
                Case "������Ӳ���"
                    TabWindow(i).Visible = True
                
                Case "סԺ���Ӳ���"
                    TabWindow(i).Visible = False
                    
                Case "Ӱ��ͼ��"
                    TabWindow(i).Visible = True
                Case "������д"
                    TabWindow(i).Visible = IIf(Not mblnIsHistory, (mobjCurStudyInfo.intStep > 1 Or mobjCurStudyInfo.intStep = -1) And blnShowReport Or GetWorkModuleCount = 1, True)
                Case "�Ŷӽк�"
                    TabWindow(i).Visible = mSysPar.blnUseQueue 'True '
            End Select
        Next
    Else
        For i = 0 To TabWindow.ItemCount - 1
            Select Case TabWindow(i).tag
                Case "���ﲡ��", "����ҽ��"
                    TabWindow(i).Visible = False

                Case "סԺ����", "סԺҽ��"
                    TabWindow(i).Visible = True
                
                Case "������Ӳ���"
                    TabWindow(i).Visible = False
                
                Case "סԺ���Ӳ���"
                    TabWindow(i).Visible = True

                Case "Ӱ��ͼ��"
                    TabWindow(i).Visible = True
                Case "������д"
                    TabWindow(i).Visible = IIf(Not mblnIsHistory, (mobjCurStudyInfo.intStep > 1 Or mobjCurStudyInfo.intStep = -1) And blnShowReport Or GetWorkModuleCount = 1, True)
                Case "�Ŷӽк�"
                    TabWindow(i).Visible = mSysPar.blnUseQueue 'True '
            End Select
        Next
    End If
    
    
    
    intDefaultIndex = GetTabWindowIndex
    
    
    '�����ǰ��ѡ���ҳ�治�ɼ�������ʾ�û�����Ҫ����ҳ��
    If TabWindow.Selected Is Nothing Then
        strFirstTab = mstrFirstTab
        For i = 0 To TabWindow.ItemCount - 1
            If InStr(TabWindow(i).tag, mSysPar.strFirstTab) > 0 And TabWindow(i).Visible Then
                TabWindow(i).Selected = True
                Exit For
            End If
        Next i
    End If
    
    If TabWindow.Selected Is Nothing Then TabWindow(intDefaultIndex).Selected = True

    If TabWindow.Selected.Visible = False Then
        strFirstTab = mstrFirstTab
        For i = 0 To TabWindow.ItemCount - 1
            If InStr(TabWindow(i).tag, mSysPar.strFirstTab) > 0 And TabWindow(i).Visible Then
                TabWindow(i).Selected = True
                Exit For
            End If
        Next i
    End If
    
    If TabWindow.Selected.Visible = False Then
        If intDefaultIndex < 0 Then
            TabWindow.Selected.Visible = True
        Else
            TabWindow(intDefaultIndex).Selected = True
            TabWindow(intDefaultIndex).Visible = True
        End If
    End If
    
    Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub RefreshModuleAdviceInf()
'ˢ��ģ��ҽ����Ϣ
On Error GoTo errhandle
    Dim intStep As Long
    Dim bln�����¿� As Boolean

    If mobjCurStudyInfo.intState = 2 Then intStep = -2
    
    'ˢ��Ӱ��ҽ��ģ���ҽ����Ϣ
    If Not mfrmWork_PacsImg Is Nothing Then
        If Not mobjCurStudyInfo Is Nothing Then Set mfrmWork_PacsImg.StudyInfo = mobjCurStudyInfo
        Call mfrmWork_PacsImg.zlUpdateAdviceInf(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, mobjCurStudyInfo.blnMoved)
        Call mfrmWork_PacsImg.zlUpdateOtherInf(mblHaveHistory, mobjCurStudyInfo.blnIsTechincalSure, mobjCurStudyInfo.strDoDoctor)
    End If
    
    'ˢ����Ƶ�ɼ�ģ���ҽ����Ϣ
    If Not mobjWork_ImageCap Is Nothing Then
        Call mobjWork_ImageCap.zlUpdateStudyInf(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, mobjCurStudyInfo.blnMoved, mobjCurStudyInfo.blnIsReported)
    End If

    'ˢ�²������ģ���ҽ����Ϣ
    If Not mobjWork_Pathol Is Nothing Then
        Call mobjWork_Pathol.zlUpdateAdviceInf(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, mobjCurStudyInfo.blnMoved)
    End If
    
    'ˢ��HIS���ģ���ҽ����Ϣ
    If Not mobjWork_His Is Nothing Then
        '�����¿�����:
        
        bln�����¿� = Not ((mobjCurStudyInfo.lngPatientFrom = 1 And mobjCurStudyInfo.lng����ִ��״̬ = 1) Or (mobjCurStudyInfo.lngPatientFrom = 2 And Not bln������Ժ(mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.lngPageID)))
        
        Call mobjWork_His.zlUpdateAdviceInf(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, mobjCurStudyInfo.blnMoved)
        Call mobjWork_His.zlUpdateOtherInf(mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.lngUnit, mobjCurStudyInfo.lngPatDept, mobjCurStudyInfo.lngPageID, _
            mobjCurStudyInfo.intState, mobjCurStudyInfo.strRegNo, mblnIsHistory, mobjCurStudyInfo.blnIsInsidePatient, bln�����¿�)
    End If
    
    'ˢ�±���ģ������ҽ����Ϣ
    If Not mobjWork_Report Is Nothing Then
        'δ����ǰ������༭���治��ʾ
        If mobjCurStudyInfo.intStep < 2 And mobjCurStudyInfo.intStep <> -1 Then
            Call mobjWork_Report.zlUpdateAdviceInf(0, 0, 0, 0, 0, 0)
            Call mobjWork_Report.zlRefreshFace(, , , , , mobjCurStudyInfo.lngAdviceId)
        Else
            Call mobjWork_Report.zlUpdateAdviceInf(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, mobjCurStudyInfo.blnMoved, mobjCurStudyInfo.lngBaby)
        End If
        
        Call mobjWork_Report.zlUpdateOtherInf(picReportContainer, vsfList, mblnIsHistory, mobjCurStudyInfo.blnCanPrint, mobjCurStudyInfo.strDoDoctor, mobjCurStudyInfo.strStudyUID)
    End If
    
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub RefreshTabWindow(Optional blnRefresh As Boolean)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ܣ�ˢ��TABҳ��
'������
'       blnRefresh ��ɺ�ȡ�������֪ͨPACS����༭��ˢ��
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo errhandle
    
    If TabWindow.Selected Is Nothing Then Exit Sub
    
    If TabWindow.Selected.tag = "" Then Exit Sub
    
    Select Case TabWindow.Selected.tag
        Case "Ӱ��ͼ��"
            If Not mobjCurStudyInfo Is Nothing Then Set mfrmWork_PacsImg.StudyInfo = mobjCurStudyInfo
            Call mfrmWork_PacsImg.zlRefreshFace
            
        Case "�걾����"
            Call mobjWork_Pathol.GetModule(mtSpecimen).zlRefreshFace
            
        Case "����ȡ��"
            Call mobjWork_Pathol.GetModule(mtMaterial).zlRefreshFace
            
        Case "������Ƭ"
            Call mobjWork_Pathol.GetModule(mtSlices).zlRefreshFace
            
        Case "�����ؼ�"
            Call mobjWork_Pathol.GetModule(mtSpeExam).zlRefreshFace
            
        Case "���̱���"
            Call mobjWork_Pathol.GetModule(mtProRep).zlRefreshFace
            
        Case "������д"
            If GetActiveWindow = Me.hwnd Then Call mobjWork_Report.zlShowReportVideo
            Call mobjWork_Report.zlUpdateAdviceInf(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngPatId, _
                mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, mobjCurStudyInfo.blnMoved, mobjCurStudyInfo.lngBaby)
            '��ɺ�ȡ�������֪ͨPACS����༭��ˢ��
            If blnRefresh Then mobjWork_Report.NotificationRefresh
            Call mobjWork_Report.zlRefreshFace(False, False, True, mobjWork_Report.IsDockActive, mblnRefreshWord)
            mblnRefreshWord = False
                
            
        Case "�������", "סԺҽ��", "����ҽ��", "סԺ����", "���ﲡ��", "������Ӳ���", "סԺ���Ӳ���"
            Call mobjWork_His.zlRefreshFace(, mobjCurStudyInfo.lngPatientFrom)
            
        Case "Ӱ��ɼ�"
            If Not mobjWork_ImageCap Is Nothing Then
                Call mobjWork_ImageCap.zlUpdateStudyInf(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, mobjCurStudyInfo.blnMoved, mobjCurStudyInfo.blnIsReported)
                Call mobjWork_ImageCap.zlRefreshData
                Call mobjWork_ImageCap.zlRefreshVideoWindow
            End If

    End Select
    
    If Not mobjWork_ImageCap Is Nothing And TabWindow.Selected.tag <> "Ӱ��ɼ�" Then
        '�����л����ǲɼ�ҳ�棬Ȼ���л����󣬲ɼ�����ͼ�������
        Call mobjWork_ImageCap.zlUpdateStudyInf(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, mobjCurStudyInfo.blnMoved, mobjCurStudyInfo.blnIsReported)
        'Call mobjWork_ImageCap.zlRefreshVideoWindow
        Call mobjWork_ImageCap.zlRefreshData
    End If
    
    If TabWindow.Selected.tag <> "Ӱ��ɼ�" And TabWindow.Selected.tag <> "�Ŷӽк�" Then
        If mobjCurStudyInfo.lngAdviceId <= 0 Then
            Call DisableWorkModule
        Else
            Call EnableWorkModule
        End If
    Else
        EnableWorkModule
    End If
    
    Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_��������()
'��������
On Error GoTo errhandle
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    Call frmReferencePatient.ZlShowMe(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.strPatientName, Me, True, mlngCur����ID)
    
    'ˢ�²����б�
     Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Menu_Manage_�����ɼ�()
On Error GoTo errhandle

    If Not GetIsValidOfStorageDevice(mlngCur����ID) Then
      MsgBoxD Me, "Ӱ��洢�豸δ�������ͣ�ã����飡", vbInformation, gstrSysName
      Exit Sub
    End If
    
    If Not mobjWork_ImageCap Is Nothing Then
        Call mobjWork_ImageCap.zlShowPopupVideo
        
        If mlngOldAdviceId <> mobjCurStudyInfo.lngAdviceId And TabWindow.Selected.Caption <> "Ӱ��ɼ�" Then
            Call mobjWork_ImageCap.zlRefreshData
            mlngOldAdviceId = mobjCurStudyInfo.lngAdviceId
        End If
    End If
    
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Manage_ͼ���¼()
'ͼ���¼
    Dim lngCurAdviceId As Long
    Dim objBurn As Object
    Dim frmBurn As frmImageBurn
    
    If mobjCurStudyInfo.intImageLocation = 1 Then
        Call subXWShowArchiveManager(3)
    Else
        On Error GoTo errExit
            Set objBurn = CreateObject("IMAPI2.MsftDiscMaster2")
            Set objBurn = Nothing
            GoTo continueBurn
errExit:
            Call MsgBoxD(Me, "���ܴ�����¼�������ڰ�װIMAPI2��¼��������½��롣", vbOKOnly, Me.Caption)
            Exit Sub
            
continueBurn:
            
            Set frmBurn = New frmImageBurn
        On Error GoTo errFree
            
            lngCurAdviceId = mobjCurStudyInfo.lngAdviceId
            
            Set frmBurn = New frmImageBurn
            Call frmBurn.ShowBurn(mlngModule, mlngCur����ID, lngCurAdviceId, mobjCurStudyInfo.blnMoved, Me)
errFree:
            Call Unload(frmBurn)
            Set frmBurn = Nothing
    End If
End Sub

Private Sub Menu_Manage_��������()
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If InStr(";" & GetPrivFunc(100, 1259) & ";", ";����;") = 0 Then
        MsgBoxD Me, "��û�в��ĵ��Ӳ�����Ȩ�ޣ�����ϵ����Ա��", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Set mobjMedicalRecord = Nothing
    If mobjMedicalRecord Is Nothing Then
        Set mobjMedicalRecord = DynamicCreate("zlPublicAdvice.clsPublicAdvice", "zlPublicAdvice")
        If mobjMedicalRecord Is Nothing Then Exit Sub
        
        Call mobjMedicalRecord.InitCommon(gcnOracle, glngSys, gstrNodeNo, gfrmMain, glngModul, gstrPrivs, mobjMsgCenter.Msg)
        
        If mobjCurStudyInfo.lngPageID <= 0 Then
            MsgBoxD Me, "�ò�����δ����������", vbInformation, Me.Caption
        Else
            Call mobjMedicalRecord.ShowArchive(Me, mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.lngPageID, True)
            
            Set mobjMedicalRecord = Nothing
        End If
    End If
    
End Sub

Private Sub Menu_Manage_�ղع���()
'�ղع���
On Error GoTo errFree
    Dim frmCollectionManage As New frmCollectionManage
    Dim lngCount As Long

    Call frmCollectionManage.ShowCollectionManageWind(Me)
    
    Call ReCreatCbrMenu(cbrMain)
    
errFree:
    Call Unload(frmCollectionManage)
    Set frmCollectionManage = Nothing
End Sub

Private Sub Menu_Manage_�ղص�()
'�ղص�
    Dim frmToCollection As New frmToCollection
    Dim rsTemp As ADODB.Recordset
    Dim lngAdviceId As Long
    Dim lngSendNo As Long
    Dim intMovedState As Integer

On Error GoTo errFree

    lngAdviceId = mobjCurStudyInfo.lngAdviceId
    lngSendNo = mobjCurStudyInfo.lngSendNo
    intMovedState = mobjCurStudyInfo.blnMoved
    
    If lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    gstrSQL = "select �״�ʱ�� from ����ҽ������ where ҽ��ID= " & lngAdviceId & ""
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    '�ж�ѡ�м�¼�Ƿ񱨵������û�б������ܽ����ղز���
    Do While Not rsTemp.EOF
        If nvl(rsTemp!�״�ʱ��) = "" Then
            Call MsgBoxD(Me, "�ü��δ�����������ղأ�", vbOKOnly, "Ӱ������վ")
            Exit Sub
        End If
        
        rsTemp.MoveNext
    Loop
    
    Call frmToCollection.ShowToCollectionWind(Me, lngAdviceId, lngSendNo)
    
    Set mobjCurStudyInfo = mobjPacsQueryWrap.GetBaseInfo(lngAdviceId, intMovedState + 1)
    
    If mobjCurStudyInfo.lngPatientFrom = 1 Then
        If Val(mobjCurStudyInfo.strMarkNum) > 0 Then labCollectionInfo = "��:" & mobjCurStudyInfo.strMarkNum & "  "
    ElseIf mobjCurStudyInfo.lngPatientFrom = 2 Then
        If Val(mobjCurStudyInfo.strMarkNum) > 0 Then labCollectionInfo = "ס:" & mobjCurStudyInfo.strMarkNum & "  "
    Else
        labCollectionInfo = ""
    End If
    
    labCollectionInfo = labCollectionInfo & mobjCurStudyInfo.strAdviceContext
    labCollectionInfo = labCollectionInfo & IIf(mobjCurStudyInfo.strCollectionInfo = "", "", "  (��" & mobjCurStudyInfo.strCollectionInfo & ")")
    
errFree:
    Call Unload(frmToCollection)
    Set frmToCollection = Nothing
End Sub


Private Sub Menu_Manage_�ղ�������ʾ(ByVal Control As XtremeCommandBars.ICommandBarControl, ByVal bytStyle As Byte)
'�ղ�������ʾ����
On Error GoTo errHand
    Dim strCollectionType As String
    Dim lngFatherID As Long
    Dim strLink As String
    
    '�����ղ�����ַ���
    If InStr(Control.Caption, "(") = 0 Then
        strCollectionType = Control.Caption
    Else
        strCollectionType = Mid(Control.Caption, 1, InStr(Control.Caption, "(") - 1)
    End If
    
    '������ID�ַ���
    If bytStyle = 0 Then
        lngFatherID = CLng(Control.ID) - CLng(comMenu_Collection_Type) * 10000#
    ElseIf bytStyle = 1 Then
        lngFatherID = CLng(Control.ID) - CLng(conMenu_Collection_ViewShare) * 10000#
    End If
    
    If Control.Caption = "�鿴��ǰ�ղ�" Then
        strLink = " select ҽ��ID from Ӱ���ղ���� A ,Ӱ���ղ����� B where A.Id=b.�ղ�Id and A.ID=" & lngFatherID & " union " & _
                        " select ҽ��ID from Ӱ���ղ���� A ,Ӱ���ղ����� B,Ӱ���ղ���� C where C.Id=b.�ղ�Id and A.Id=C.�ϼ�id  and A.ID=" & lngFatherID & ""
    Else
        strLink = "select ҽ��ID from Ӱ���ղ���� A ,Ӱ���ղ����� B where A.Id=b.�ղ�Id and  A.�ղ����='" & strCollectionType & "'"
    End If
    
    Call mobjPacsQueryWrap.ExecuteWithLink(strLink)
    TimerRefresh.Enabled = False
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Petition_ɨ�����뵥(ByVal intType As Integer)
'intType:0--�鿴���뵥��1--ɨ�����뵥

On Error GoTo errFree
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim strPatientDepartment As String
    Dim lngDepID As Long
    
    '���ж��Ƿ���ɨ������뵥�����û�У�ֱ��Ԥ������
    If intType = 0 Then
        strSQL = "select ���뵥ͼ�� from Ӱ�����뵥ͼ�� where ҽ��ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ����ɨ�����뵥", mobjCurStudyInfo.lngAdviceId)
        If rsTemp.RecordCount = 0 Then
            Call ViewEPRPetition(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.lngClinicId, mobjCurStudyInfo.lngPatientFrom)
            Exit Sub
        End If
    End If
    
    Set mobjPetitionCap = New frmPetitionCapture
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    lngDepID = IIf(mlngCur����ID = 0, mobjCurStudyInfo.lngExeDepartmentId, mlngCur����ID)
    With mobjCurStudyInfo
        strSQL = "Select ���� From ���ű� Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˿���", .lngPatDept)
        strPatientDepartment = ""
        If rsTemp.RecordCount > 0 Then strPatientDepartment = nvl(rsTemp!����)
    
        Call mobjPetitionCap.ShowPetitionCaptureWind(mstrPrivs, _
                                    lngDepID, _
                                    strPatientDepartment, _
                                    .strPatientName, _
                                    .strPatientAge, _
                                    .strPatientSex, _
                                    .strAdviceContext, _
                                    .strAdviceDepartAndMethod, _
                                    IIf(Not CheckPopedom(mstrPrivs, "���Ǽ�"), True, IIf(intType = 0, True, False)), _
                                    False, _
                                    .lngAdviceId, _
                                    IIf(.strStuStateDesc = "�Ѿܾ�", 1, IIf(.strStuStateDesc = "�����", 2, 0)))
    End With
errFree:
    Call Unload(mobjPetitionCap)
    Set mobjPetitionCap = Nothing
End Sub

Private Sub conMenu_File_SendImg_click()
'------------------------------------------------
'���ܣ�����ͼ��
'���أ�
'------------------------------------------------
    On Error GoTo err
    If mlngModule = G_LNG_PACSSTATION_MODULE Then
        If mobjCurStudyInfo.lngAdviceId <= 0 Or mobjCurStudyInfo.intImageLocation = 1 Then
            Call subXWShowArchiveManager(2)
        Else
            frmPacsSendImage.ShowMe Me
        End If
    Else
        frmPacsSendImage.ShowMe Me
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub mobjWork_Report_OnImageCountChanged(ByVal intType As Integer, ByVal isNeedRefreshTitle As Boolean)
    If Not mobjWork_ImageCap Is Nothing Then
        Call mobjWork_ImageCap.showAfterCapInfo(intType, isNeedRefreshTitle)
    End If
End Sub

Private Sub initInterface(ByVal lngModule As Long)
'��ʼ����Ҫ�Զ�ִ�еĲ��
On Error GoTo errH

    Dim i As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim intExeTime As Integer
    Dim intType As Integer
    Dim strVBS As String

    mintInterfaceCount = 0
    strSQL = "Select a.���� as ������, b.���� as ������ , b.�Զ�ִ��ʱ��,b.vbs�ű�  from Ӱ�����ҽ� a, Ӱ�������� b " & _
             "Where   b.�Ƿ�����=1 and  a.�Ƿ�����=1 and a.id = b.���id And (a.����ģ��=0 or a.����ģ��=[1]) Order By a.id,b.�������"
             
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ʼ�����", lngModule)
    
    If rsTemp.RecordCount > 0 Then
    ReDim mintInterface(rsTemp.RecordCount)

        While Not rsTemp.EOF
    
            intExeTime = Val(nvl(rsTemp!�Զ�ִ��ʱ��))
            
            If intExeTime > 0 Then
                strVBS = nvl(rsTemp!VBS�ű�)
                
                mintInterfaceCount = mintInterfaceCount + 1
                mintInterface(mintInterfaceCount).intID = mintInterfaceCount
                mintInterface(mintInterfaceCount).strVBS = strVBS
                mintInterface(mintInterfaceCount).intExeTime = intExeTime
                mintInterface(mintInterfaceCount).strName = nvl(rsTemp!������) & "-" & nvl(rsTemp!������)
            End If
            
            Call rsTemp.MoveNext
        Wend
    End If
        
    Exit Sub
errH:
    MsgBoxD Me, "��ʼ���Զ�ִ�в�����̷�������,��ϸ��Ϣ��" & err.Description, vbInformation, Me.Caption
End Sub

Private Sub CheckExecuteInterface(ByVal intTime As Integer)
'���ܣ�����ʱ���Ƿ�����Ҫ�Զ�ִ�еĲ������
'intTime:ִ��ʱ��
On Error GoTo errH

    Dim i As Integer
        
    If mintInterfaceCount <= 0 Then Exit Sub
    
    For i = 1 To mintInterfaceCount
        If mintInterface(i).intExeTime = intTime Then
            Call ExecuteInterfaceFun(mintInterface(i).strVBS, 0, True)
        End If
    Next

    Exit Sub
errH:
    MsgBoxD Me, "���[" & mintInterface(i).strName & "]ִ���쳣��������Ϣ��" & err.Description, vbInformation, Me.Caption
    err.Clear
End Sub

Private Function ChechHaveTlbinf32() As Boolean
On Error Resume Next
    Dim objtest As Object
    
    ChechHaveTlbinf32 = False
    Set objtest = CreateObject("TLI.TLIApplication")
    
    If Not objtest Is Nothing Then ChechHaveTlbinf32 = True
    
    Set objtest = Nothing
End Function

Public Sub DoFontSize(ByVal blIsDock As Boolean, ByVal intFontSize As Integer)
    Call mobjWork_Report.DoFontSize(blIsDock, intFontSize)
End Sub

Private Sub AdjustFace(ByVal lngH As Long, ByVal lngW As Long)
'�ֺ� Ŀǰ����վ֧��9,12,15����;lngH �߶ȣ�lngW ���   C_LAYOUT_LISTLEFT
''������ؼ����ϵ��� mobjFilterCmdBar��mobjFindPati+mobjFindCmd��mobjList��mobjIconPanel��mobjTab
On Error GoTo errH
    Dim lng���ٹ��� As Long
    Dim lng���Ҳ��� As Long
    Dim lngList As Long
    Dim lngInfo As Long
    Dim lngTab As Long
    Dim lngMoreW As Long
    
    If Val(tabScheme.tag) = 1 Then
        If gbytFontSize = 9 Then
            lngMoreW = 320
        ElseIf gbytFontSize = 12 Then
            lngMoreW = 380
        Else
            lngMoreW = 490
        End If
    Else
        lngMoreW = 0
    End If
    
    '�����Ǵ�Ź涨�ķָ�����Ч�ƶ���Χ
    If mlngMove > 6000 Then mlngMove = 6000
    If mlngMove < -4000 Then mlngMove = -4000

    If Not mobjPacsQueryWrap.blShowPatiIdentify Then
        lng���Ҳ��� = 0
    Else
        If gbytFontSize = 15 Then
            lng���Ҳ��� = 400
        Else
            lng���Ҳ��� = 350
        End If
    End If

    If mobjPacsQueryWrap.SqlScheme Is Nothing Then
        lng���ٹ��� = 0
    Else
        If Not mobjPacsQueryWrap.SqlScheme.FilterCfgCount > 0 Then
            lng���ٹ��� = 0
        Else
            If gbytFontSize = 15 Then
                lng���ٹ��� = 550
            ElseIf gbytFontSize = 12 Then
                lng���ٹ��� = 450
            Else
                lng���ٹ��� = 400
            End If
        End If
    End If
    
    lngInfo = C_LAYOUT_BASEHEIGHTOFDETAILINFO
    If gbytFontSize = 15 Then
        lngInfo = C_LAYOUT_BASEHEIGHTOFDETAILINFO + 200
    ElseIf gbytFontSize = 12 Then
        lngInfo = C_LAYOUT_BASEHEIGHTOFDETAILINFO + 70
    Else
        lngInfo = C_LAYOUT_BASEHEIGHTOFDETAILINFO
    End If
    
    lngTab = C_LAYOUT_BASEHEIGHTOFTAB + mlngMove
    lngList = lngH - lng���Ҳ��� - lng���ٹ��� - lngInfo - lngTab
    If lngList < 0 Then lngList = 0
    
    Call tabScheme.Move(0, 0, lngMoreW + C_LAYOUT_LISTLEFT, lngH)
    
    Call picFilter.Move(lngMoreW + C_LAYOUT_LISTLEFT, 0, lngW - lngMoreW, lng���ٹ���)
    Call PatiIdentify.Move(lngMoreW + C_LAYOUT_LISTLEFT, picFilter.Top + picFilter.Height, lngW - lngMoreW - 0.5 * C_LAYOUT_LISTLEFT - cmdLocate.Width - cmdFind.Width, lng���Ҳ���)
    
    Call cmdLocate.Move(lngMoreW + PatiIdentify.Width, PatiIdentify.Top, cmdLocate.Width, lng���Ҳ���)
    Call cmdFind.Move(cmdLocate.Left + cmdLocate.Width, PatiIdentify.Top, cmdFind.Width, lng���Ҳ���)
    
    If mobjPacsQueryWrap.blShowPatiIdentify Then
        Call vsfList.Move(lngMoreW + C_LAYOUT_LISTLEFT, PatiIdentify.Top + PatiIdentify.Height, lngW - lngMoreW - C_LAYOUT_LISTLEFT, lngList)
        cmdLocate.Visible = True
        cmdFind.Visible = True
    Else
        Call vsfList.Move(lngMoreW + C_LAYOUT_LISTLEFT, picFilter.Top + picFilter.Height, lngW - lngMoreW - C_LAYOUT_LISTLEFT, lngList)
        cmdLocate.Visible = False
        cmdFind.Visible = False
    End If
    
    Call PicLine.Move(C_LAYOUT_LISTLEFT, vsfList.Top + vsfList.Height, lngW - C_LAYOUT_LISTLEFT, 50)

    Call picDetail.Move(lngMoreW + C_LAYOUT_LISTLEFT, vsfList.Top + vsfList.Height + 50, lngW - lngMoreW - C_LAYOUT_LISTLEFT, lngInfo)
    
    Call imgStep.Move(C_LAYOUT_LISTLEFT, C_LAYOUT_LISTLEFT)
    
    If labCollectionInfo = "" Then
        Call labPatientInfo.Move(2 * C_LAYOUT_LISTLEFT + imgStep.Width + 60, C_LAYOUT_LISTLEFT + (540 - labPatientInfo.Height) / 2)
    Else
        Call labPatientInfo.Move(2 * C_LAYOUT_LISTLEFT + imgStep.Width + 60, C_LAYOUT_LISTLEFT)
    End If
    Call labCollectionInfo.Move(2 * C_LAYOUT_LISTLEFT + imgStep.Width + 60, labPatientInfo.Top + labPatientInfo.Height)
    Call labPatientAge.Move(labPatientInfo.Left + labPatientInfo.Width + TextWidth("  "), labPatientInfo.Top)
    
    Call TabExtra.Move(lngMoreW + C_LAYOUT_LISTLEFT, picDetail.Top + picDetail.Height, lngW - lngMoreW - C_LAYOUT_LISTLEFT, lngTab)
    picDataSearchContainer.Width = lngW - C_LAYOUT_LISTLEFT
    
    Call rtxtAppend.Move(0, 0, lngW - C_LAYOUT_LISTLEFT, TabExtra.Height)
    
    Call pic�������ڵ�.Move(0, 0, picList.Width, picList.Height)
    Call labNoScheme.Move((picList.Width - labNoScheme.Width) / 2, (picList.Height - labNoScheme.Height) / 2)
    
    Call DoLabFlag
errH:
End Sub

Private Sub initTabExtra()
'��ʼ���������½�Tab�ؼ�
' ��ؿؼ��� TabExtra  picDataSearch�����ݼ����� picExtra(������Ϣ)  picFollowUp(���)  picEvent(����)
''���ݼ��� ������Ϣ ���μ�� ������� �������� ���ƹ̶���Ҫ�޸�ע���ѯcls �����޸�
    Dim strSelect As String
    Dim i As Integer
    Dim CtlFont As StdFont
    Dim lngW As Long
    Dim lngH As Long
    
    
    With TabExtra
        .RemoveAll

        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.ColorSet.ButtonNormal = &HE0E0E0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ButtonMargin.Top = 4
        .PaintManager.ButtonMargin.Bottom = 4
        .PaintManager.ShowIcons = True
        .RemoveAll
        
        .InsertItem 1, "���ݼ���", picDataSearchContainer.hwnd, 0
        .Item(TabExtra.ItemCount - 1).tag = "���ݼ���"
        
        .InsertItem 2, "������Ϣ", picExtra.hwnd, 0
        .Item(TabExtra.ItemCount - 1).tag = "������Ϣ"
        
        .InsertItem 3, "���μ��", mfrmHistory.hwnd, 0
        .Item(TabExtra.ItemCount - 1).tag = "���μ��"
        
        
        
'        .InsertItem 4, "�������", picFollowUp.hWnd, 0
'        .Item(TabExtra.ItemCount - 1).tag = "�������"
'
'        .InsertItem 5, "��������", picEvent.hWnd, 0
'        .Item(TabExtra.ItemCount - 1).tag = "��������"
        
        
        strSelect = mobjPacsQueryWrap.GetTabSelectName(False)
        .Item(0).Selected = True
        
        .Width = Screen.Width
        
        For i = 0 To .ItemCount - 1
            If strSelect = .Item(i).tag And .Item(i).Visible Then
                .Item(i).Selected = True
                Exit For
            End If
        Next
        
        lngW = Val(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.���� & App.ProductName & "\", "��ʷ��鸡��������", 3375))
        lngH = Val(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.���� & App.ProductName & "\", "��ʷ��鸡������߶�", 1215))
    
        If lngW < 1500 Or lngW > 8000 Then
            lngW = 3375
        End If
        
        If lngH < 500 Or lngH > 2500 Then
            lngH = 1215
        End If
        Call PicFollowHistory.Move(0, 0, lngW, lngH)
        Call rftHistoryFollow.Move(50, 50, lngW - 100, lngH - 100)
        
        '���ݼ��� ������Ϣ ���μ�� ������� ��������
    End With
    
End Sub

Public Sub ExecuteDefaultQueryScheme()
'ִ���Զ����ѯĬ�Ϸ���
On Error GoTo errH
    Dim i As Long
    Dim lngShemeNo As Long
    Dim lngShemeNoFirst As Long
    Dim t1 As Long
    Dim blUseFirst As Boolean
    Dim intIndexFirst As Integer
    Dim dtStartDate As Date, dtEndDate As Date
    
    t1 = GetTickCount
    lngShemeNo = -1

    If mobjPacsQueryWrap.CurPacsQuery Is Nothing Then Exit Sub
    
    With mobjPacsQueryWrap.CurPacsQuery
    
        For i = 1 To .SchemeCount
            If .SchemeInfo(i).IsDefault Then
                lngShemeNo = .SchemeInfo(i).SchemeId
                blUseFirst = False
                Exit For
            ElseIf Not .SchemeInfo(i).IsDefault And .SchemeInfo(i).IsOften Then
                lngShemeNoFirst = .SchemeInfo(i).SchemeId
                intIndexFirst = i
                blUseFirst = True
                If lngShemeNo <> -1 Then Exit For
            End If
        Next
        
        If lngShemeNo = -1 Then lngShemeNo = lngShemeNoFirst
        
        If lngShemeNo <> -1 Then
            labPatientInfo.Caption = ""
            labCollectionInfo.Caption = ""
            labPatientAge.Caption = ""
            
            dtStartDate = mobjPacsQueryWrap.StartDate
            dtEndDate = mobjPacsQueryWrap.EndDate
            If dtStartDate <> Empty And dtEndDate <> Empty Then
                Call mobjPacsQueryWrap.ExecuteMenu(lngShemeNo, mSysPar.blnQueryValidTime, dtStartDate, dtEndDate)
            Else
                Call mobjPacsQueryWrap.ExecuteMenu(lngShemeNo)
            End If

            Call InitAutoRefresh 'ExecuteMenu�����ִ��
            gblnXWMoved = mobjPacsQueryWrap.CurPacsQuery.IsMoved 'ExecuteMenu�����ִ��
            
            If blUseFirst Then
                dkpMain.FindPane(1).title = .SchemeInfo(intIndexFirst).Name
            Else
                dkpMain.FindPane(1).title = .SchemeInfo(i).Name
            End If
            Call mobjPacsQueryWrap.RefreshTabLeft(tabScheme, dkpMain.FindPane(1).title)
            
            Call AdjustFace(picList.Height, picList.Width)
        End If
    
    End With
    
    cmdDo.Visible = True
    cmdClear.Visible = True
    cmdMore.Visible = mobjPacsQueryWrap.CurPacsQuery.IsMoreEmbedInput
    If Not cmdMore.Visible Then
        Call cmdClear.Move(cmdDo.Left, cmdClear.Top, cmdDo.Width)
        cmdClear.Width = cmdDo.Width
    Else
        Call cmdClear.Move(cmdDo.Left, cmdClear.Top, 0.5 * cmdDo.Width)
    End If
    Call cmdMore.Move(cmdClear.Left + cmdClear.Width)
    
    Exit Sub
errH:
    MsgBoxD Me, "ִ��Ĭ�Ϸ����쳣��������Ϣ��" & err.Description, vbInformation, Me.Caption
End Sub

Public Sub UpdateQueryListData(ByRef rsData As Recordset, ByVal lngAdviceId As Long, Optional ByVal intSyncDataType As Integer = SyncDataType.rsDataAndrsShow)
'���²�ѯ�б�ĳһ������
'ͬʱ����¸��л������ݣ�ע��Ҫ���жϸ������Ƿ��ǵ�ǰѡ����
'blIsAdd �Ƿ���������
'lngAdviceID�仯�е�ҽ��ID
'blRaiseEventSelChange �Ƿ񴥷��б�selchange�¼�
On Error GoTo errH
    If Not mobjPacsQueryWrap Is Nothing Then Call mobjPacsQueryWrap.UpdateRow(rsData, lngAdviceId, intSyncDataType)
    Exit Sub
errH:
    MsgBoxD Me, "�����������쳣��������Ϣ��" & err.Description, vbInformation, Me.Caption
End Sub

Private Sub DoLabFlag()
    Dim lng��ʶ�߳� As Long
    Dim test As Boolean
    Dim lngTop��� As Long
    
    lngTop��� = 30
    lng��ʶ�߳� = 270
    
    Call LabFlag����.Move(picDetail.Width - lng��ʶ�߳�, lngTop���, lng��ʶ�߳�, lng��ʶ�߳�)
    Call LabFlag��Ⱦ��״̬.Move(picDetail.Width - 2 * lng��ʶ�߳�, lngTop���, lng��ʶ�߳�, lng��ʶ�߳�)
    Call LabFlagΣ��״̬.Move(picDetail.Width - 3 * lng��ʶ�߳�, lngTop���, lng��ʶ�߳�, lng��ʶ�߳�)
    Call LabFlag��ɫͨ��.Move(picDetail.Width - 4 * lng��ʶ�߳�, lngTop���, lng��ʶ�߳�, lng��ʶ�߳�)
    Call LabFlagӤ��.Move(picDetail.Width - 5 * lng��ʶ�߳�, lngTop���, lng��ʶ�߳�, lng��ʶ�߳�)
    Call LabFlag����.Move(picDetail.Width - 6 * lng��ʶ�߳�, lngTop���, lng��ʶ�߳�, lng��ʶ�߳�)
    
    If mobjCurStudyInfo.lngAdviceId < 1 Then
        LabFlag����.Visible = False
        LabFlagӤ��.Visible = False
        LabFlag��ɫͨ��.Visible = False
        LabFlagΣ��״̬.Visible = False
        LabFlag��Ⱦ��״̬.Visible = False
        LabFlag����.Visible = False
    Else
        If mobjCurStudyInfo.intEmergentTag Then
            LabFlag����.Visible = True
        Else
            LabFlag����.Visible = False
            Call LabFlag��Ⱦ��״̬.Move(LabFlag��Ⱦ��״̬.Left + lng��ʶ�߳�)
            Call LabFlagΣ��״̬.Move(LabFlagΣ��״̬.Left + lng��ʶ�߳�)
            Call LabFlag��ɫͨ��.Move(LabFlag��ɫͨ��.Left + lng��ʶ�߳�)
            Call LabFlagӤ��.Move(LabFlagӤ��.Left + lng��ʶ�߳�)
            Call LabFlag����.Move(LabFlag����.Left + lng��ʶ�߳�)
        End If
    
        If mobjCurStudyInfo.blnIsInfectious Then
            LabFlag��Ⱦ��״̬.Visible = True
        Else
            LabFlag��Ⱦ��״̬.Visible = False
            Call LabFlagΣ��״̬.Move(LabFlagΣ��״̬.Left + lng��ʶ�߳�)
            Call LabFlag��ɫͨ��.Move(LabFlag��ɫͨ��.Left + lng��ʶ�߳�)
            Call LabFlagӤ��.Move(LabFlagӤ��.Left + lng��ʶ�߳�)
            Call LabFlag����.Move(LabFlag����.Left + lng��ʶ�߳�)
        End If
        
        If mobjCurStudyInfo.intDangerState = 1 Then
            LabFlagΣ��״̬.Visible = True
        Else
            LabFlagΣ��״̬.Visible = False
            Call LabFlag��ɫͨ��.Move(LabFlag��ɫͨ��.Left + lng��ʶ�߳�)
            Call LabFlagӤ��.Move(LabFlagӤ��.Left + lng��ʶ�߳�)
            Call LabFlag����.Move(LabFlag����.Left + lng��ʶ�߳�)
        End If
        
        If mobjCurStudyInfo.intGreenChannel = 1 Then
            LabFlag��ɫͨ��.Visible = True
        Else
            LabFlag��ɫͨ��.Visible = False
            Call LabFlagӤ��.Move(LabFlagӤ��.Left + lng��ʶ�߳�)
            Call LabFlag����.Move(LabFlag����.Left + lng��ʶ�߳�)
        End If

        If mobjCurStudyInfo.lngBaby > 0 Then
            LabFlagӤ��.Visible = True
        Else
            LabFlagӤ��.Visible = False
            Call LabFlag����.Move(LabFlag����.Left + lng��ʶ�߳�)
        End If
        
        Select Case mobjCurStudyInfo.lngMoneyState
            Case ChargeState.δ�շ�
                LabFlag����.Caption = "Ƿ"
'                LabFlag����.ForeColor = &H80FF&
            Case ChargeState.���շ�
                LabFlag����.Caption = "��"
'                LabFlag����.ForeColor = &H8000&
            Case ChargeState.�޷���
                LabFlag����.Caption = "��"
'                LabFlag����.ForeColor = &HC00000
            Case ChargeState.�Ѳ���
                LabFlag����.Caption = "��"
'                LabFlag����.ForeColor = &HFF&
            Case ChargeState.�Ѽ���
                LabFlag����.Caption = "��"
'                LabFlag����.ForeColor = &HFF00FF
            Case ChargeState.���˷�
                LabFlag����.Caption = "��"
'                LabFlag����.ForeColor = &H80000011
            Case ChargeState.������
                LabFlag����.Caption = "��"
'                LabFlag����.ForeColor = &H8080&
            Case ChargeState.�ѵ���
                LabFlag����.Caption = "��"
'                LabFlag����.ForeColor = &H94
        End Select
        LabFlag����.Visible = True

    End If
End Sub

Private Sub TimFlicker_Timer()
On Error GoTo errH
'   ��ʱ��˸�Ĵ���
    Dim i As Integer, j As Integer
    Dim lngCol As Long, lngColContrast As Long
    Dim strTmp As String
    Dim lngStateColor As Long, lngNextStateColor As Long, lngPreStateColor As Long
    Dim objRowRelation As Object
    
    Static intsta As Integer
    Static TPFlickerInfo As TFlickerInfo '��ʱ��˸����
    
    '������һ�μ���ʱ��ȡ��ʱ��˸�����Ϣ
    If TPFlickerInfo.LngSchemeNo <> mobjPacsQueryWrap.SchemeNo Then
        TPFlickerInfo.strName = ""
        TPFlickerInfo.strInfo = ""
    
        If mobjPacsQueryWrap.SqlScheme Is Nothing Then Exit Sub
        TPFlickerInfo.LngSchemeNo = mobjPacsQueryWrap.SchemeNo
        
        For i = 1 To mobjPacsQueryWrap.SqlScheme.ShowCfgCount
            For j = 1 To mobjPacsQueryWrap.SqlScheme.ShowCfg(i).RowRelationCount
                Set objRowRelation = mobjPacsQueryWrap.SqlScheme.ShowCfg(i).RowRelation(j)
                
                If objRowRelation.FlickerTimeOut > 0 Then
                    TPFlickerInfo.strName = mobjPacsQueryWrap.SqlScheme.ShowCfg(i).Name
                    TPFlickerInfo.strInfo = TPFlickerInfo.strInfo & objRowRelation.TiggerData & "," & objRowRelation.TimeOutReferCol & "," & objRowRelation.FlickerTimeOut & "|"

                End If
            Next
        Next
        
        intsta = 0
        Exit Sub
        
    End If
    
    intsta = intsta + 1
    If intsta = 4 Then intsta = 1

    lngCol = vsfList.ColIndex(TPFlickerInfo.strName)
    If vsfList.TopRow = vsfList.BottomRow Then Exit Sub
    For i = vsfList.TopRow To vsfList.BottomRow   '�����ɼ���  For 1
        For j = 0 To UBound(Split(TPFlickerInfo.strInfo, "|")) - 1 '�ж��Ƿ����㳬ʱ���� For 2
            strTmp = Split(TPFlickerInfo.strInfo, "|")(j)
            If Split(strTmp, ",")(0) = vsfList.TextMatrix(i, lngCol) Then
                lngColContrast = vsfList.ColIndex(Split(strTmp, ",")(1))
                
                If IsDate(vsfList.TextMatrix(i, lngColContrast)) Then
                
                    If DateDiff("N", vsfList.TextMatrix(i, lngColContrast), Now) >= Val(Split(strTmp, ",")(2)) Then    '���������õĳ�ʱʱ��
                    
                        '���Ȳ�����˸����
                        lngStateColor = C_COLOR_LISTCOL0
                        lngNextStateColor = vbYellow
                        lngPreStateColor = RGB(0, 0, 0)
    
                        If intsta = 1 Then
                            vsfList.Cell(flexcpBackColor, i, 0) = lngPreStateColor
                        ElseIf intsta = 2 Then
                            vsfList.Cell(flexcpBackColor, i, 0) = C_COLOR_LISTCOL0
                        Else
                            vsfList.Cell(flexcpBackColor, i, 0) = lngNextStateColor
                        End If
                    End If
                End If
                
                Exit For   '�����㳬ʱ���� �˳�For 2
            End If
        Next
    Next
    Exit Sub
errH:
'    err.Raise -1, "frmPacsQuery", "[TimFlicher_Timer]" & vbCrLf & err.Description
End Sub

Private Sub timFun_Timer()
    PicFucs.Visible = False
End Sub

Private Sub rftHistoryFollow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'�����μ��tab����ѡ�У���ֻҪ����ƶ������������ϣ��������屣�ֲ���ʧ
On Error GoTo errH
    If TabExtra.Item(2).Selected Then
        TimerHistory.Enabled = False
    End If
    
    If mblnpicHistoryMoving Then
        PicFollowHistory.MousePointer = vbSizeNWSE
    Else
        PicFollowHistory.MousePointer = vbDefault
    End If
    
    rftHistoryFollow.MousePointer = vbDefault
    
    Exit Sub
errH:
End Sub

Private Sub vsfList_DblClick()
On Error GoTo errH
    Call VsfListDbClick(False)
errH:
End Sub

Private Sub CheckHaveScheme(ByVal blLoadFail As Boolean, ByVal strHint As String)
'�����Ƿ��Ѿ��з���������������ʾ��Ϣ
'Ŀǰ�У��������޷����������÷������������
    
    If blLoadFail Then
        pic�������ڵ�.Visible = True
            labNoScheme.Visible = True
            Call pic�������ڵ�.Move(picList.Left, picList.Top, picList.Width, picList.Height)
            
            If Trim(strHint) <> "" Then
                labNoScheme.Caption = "��ѯ�������ش���" & vbLf & strHint
            Else
                labNoScheme.Caption = "��ѯ�������ش�������ϵ���������Ա"
            End If
    Else
        If mintQueryState = 1 Then
            pic�������ڵ�.Visible = False
            labNoScheme.Visible = False
        Else
            pic�������ڵ�.Visible = True
            labNoScheme.Visible = True
            Call pic�������ڵ�.Move(picList.Left, picList.Top, picList.Width, picList.Height)
            
            If mintQueryState = 2 Then
                labNoScheme.Caption = "û����Ч��ѯ��������������"
            ElseIf mintQueryState = 3 Then
                labNoScheme.Caption = "û�����÷���"
            Else
                labNoScheme.Caption = "��ѯ�������ش�������ϵ���������Ա"
            End If
        End If
    End If
    
    Call picList_Resize
End Sub

Private Sub vsfList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    timFun.Enabled = True
End Sub

Private Sub OpenReport()
' ��ʷ����鿴����
On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim intType As Integer
    Dim lngDocID As Long
    Dim objPublicPACS As Object
    
    If mobjHistoryStudyInfo.lngAdviceId <= 0 Then
        Exit Sub
    End If
    
    intType = -1
    intType = GetDeptPara(mobjHistoryStudyInfo.lngExeDepartmentId, "����༭��", 0)                '����༭��
    
    If intType = PACS����༭�� Or intType = ���Ӳ����༭�� Then
        strSQL = "select ����ID FROM ����ҽ������ where ҽ��ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ID", mobjHistoryStudyInfo.lngAdviceId)
        
        If rsTemp.RecordCount = 1 Then
            lngDocID = Val(rsTemp!����Id)
        End If
        
        If intType = ���Ӳ����༭�� Then
            If mobjPublicPACS Is Nothing Then
                Set mobjPublicPACS = CreateObject("zlPublicPACS.clsPublicPACS")
                Call mobjPublicPACS.initInterface(gcnOracle, UserInfo.�û���)
            End If
            
            Call mobjPublicPACS.ViewEPRReport(Me, lngDocID, mobjHistoryStudyInfo.lngAdviceId, True)
            Exit Sub
        End If
        Call frmReportHistory.ZlShowMe(Me, mobjHistoryStudyInfo.lngAdviceId, lngDocID, mobjHistoryStudyInfo.blnMoved)
        Call MoveWindow(frmReportHistory.hwnd, (Me.ScaleWidth - frmReportHistory.ScaleWidth) / (2 * Screen.TwipsPerPixelX), _
        (Me.ScaleHeight - frmReportHistory.ScaleHeight) / (2 * Screen.TwipsPerPixelY), _
        frmReportHistory.ScaleWidth / Screen.TwipsPerPixelX, frmReportHistory.ScaleHeight / Screen.TwipsPerPixelY, 1)
    
    Else
        If mobjPacsInterface Is Nothing Then
            If mobjPacsInterface Is Nothing Then Set mobjPacsInterface = DynamicCreate("ZLSoft.ZLPacs.Interface.PacsInterface", "PACS���ܱ���༭���ӿڲ���zlPacsInterfaceC")
        End If
        
        If Not mobjPacsInterface Is Nothing Then
            Call mobjPacsInterface.RefreshReportForm(mobjHistoryStudyInfo.lngAdviceId, mobjHistoryStudyInfo.lngPatId, mobjHistoryStudyInfo.lngExeDepartmentId, 6, False, False)
            Call mobjPacsInterface.ForceLockReport(True)
            Call mobjPacsInterface.OpenFormForEditReport(Me.hwnd, mobjHistoryStudyInfo.lngAdviceId, mobjHistoryStudyInfo.lngPatId, 6, False)
            Call mobjPacsInterface.ForceLockReport(False)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub RefreshTab(ByVal blExistImg As Boolean)
'Ŀǰר���ڴ���123803 ����blnReportWithImage  ���á���ͼ�����д���桱 �������    �ɼ�ͼ��� ���� ��ʾ�����ǩ�����ͼ��� ���� ���ر����ǩ��
'blExistImg   true:��ͼ��    false û��ͼ��
On Error GoTo errH
    Dim i As Integer
    
    If Not mSysPar.blnReportWithImage Then Exit Sub
    For i = 0 To TabWindow.ItemCount - 1
        If TabWindow(i).tag = "������д" Then
            TabWindow(i).Visible = blExistImg Or GetWorkModuleCount = 1
            Exit Sub
        End If
    Next
    Exit Sub
errH:
    MsgBoxD Me, "RefreshTabִ���쳣��������Ϣ��" & err.Description, vbInformation, Me.Caption
End Sub

Private Sub CreateAuditorMenu(objControl As CommandBarControl)
'��������˲˵�
On Error GoTo errH
    Dim cbrPopControl As CommandBarControl
    Dim rsTemp As Recordset
    Dim strSQL As String
    Dim i As Long
    
    If Not objControl Is Nothing Then
        objControl.CommandBar.Controls.DeleteAll
    End If
    
    If mblnAllDepts Then
        strSQL = "Select Distinct a.Id, a.����" & vbNewLine & _
            "From ��Ա�� a, ������Ա b, ��������˵�� c" & vbNewLine & _
            "Where a.Id = b.��Աid And b.����id = c.����id And c.�������� = '���'"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����˱����ʸ��ҽ��")
    Else
        strSQL = "select A.id,A.���� from ��Ա�� A,������Ա B where B.����ID=[1] AND A.ID=B.��ԱID"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����˱����ʸ��ҽ��", mlngCur����ID)
    End If
    
    If rsTemp.RecordCount < 1 Then Exit Sub
    For i = 1 To rsTemp.RecordCount
        If GetUserSignLevel(rsTemp!ID) >= cprSL_���� Then
            Set cbrPopControl = CreateModuleMenu(objControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_SendAudit * 10# + i, rsTemp!����, "", 0, False)
        End If
        rsTemp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_���ԤԼ()
'------------------------------------------------
'���ܣ��򿪼��ԤԼ����
'��������
'���أ���
'------------------------------------------------
    On Error GoTo err
    Dim i As Integer
    Dim strIds As String
    Dim lngID() As Long
    Dim blnCheckin As Boolean
    
    blnCheckin = True
    strIds = frmSchSchedule.ZlShowMe(mstrPrivs, mobjCurStudyInfo.lngAdviceId, IIf(mlngCur����ID = 0, mstrCanUse����IDs, mlngCur����ID), Me, blnCheckin)
    
    If strIds = "" Then Exit Sub

    If blnCheckin = True Then
        Call Menu_Manage_����
    End If
    
    '������ֵ
    If InStr(strIds, ",") > 0 Then
        lngID = Split(strIds, ",")

        For i = 0 To UBound(lngID)
            If lngID(i) > 0 Then Call UpdateQueryListData(Nothing, lngID(i))
        Next
    Else
        ReDim lngID(0)
        lngID(0) = Val(strIds)
        If lngID(0) > 0 Then Call UpdateQueryListData(Nothing, lngID(i))
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Manage_ԤԼ����()
'------------------------------------------------
'���ܣ���ԤԼ������
'��������
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    frmSchManage.ZlShowMe mstrPrivs, IIf(mlngCur����ID = 0, mstrCanUse����IDs, mlngCur����ID), mobjCurStudyInfo.lngAdviceId, Me
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function GetSelQueueRooms(Optional blnQuick As Boolean = False) As String
On Error GoTo errH
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strID As String
    Dim strTmp As String
    
    If blnQuick Then
        If mstrSelQueueRooms <> "" Then
            GetSelQueueRooms = mstrSelQueueRooms
            Exit Function
        Else
            mstrSelQueueRooms = ""
        End If
        
        If mblnAllDepts Then
            If CheckPopedom(mstrPrivs, "���п���") Then
                strSQL = "select ����,ִ�м� from ҽ��ִ�з��� a, ���ű� b where a.����Id=b.Id and instr([1],b.ID)>0 "
                
                strID = mstrCanUse����IDs
            Else
                '��ѯ��Ӧ��Ա���ڿ�������������ִ�м�
                strSQL = "select ����,ִ�м� from ҽ��ִ�з��� a, ������Ա b,���ű� c where a.����id=b.����id and a.����Id=c.Id and b.��Աid = [1]"
                
                strID = UserInfo.ID
            End If
                    
        Else
            strSQL = "Select ����,ִ�м� From ҽ��ִ�з��� a, ���ű� b Where a.����Id=b.Id and  ����ID=[1]"
            
            strID = mlngCur����ID
            
        End If
        
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strID)
        
        
        While rsData.EOF = False
        
            If mstrSelQueueRooms <> "" Then mstrSelQueueRooms = mstrSelQueueRooms & ","
            mstrSelQueueRooms = mstrSelQueueRooms & nvl(rsData!����) & "-" & nvl(rsData!ִ�м�)
            rsData.MoveNext
            
        Wend
        
        GetSelQueueRooms = mstrSelQueueRooms
    Else
        GetSelQueueRooms = mobjPacsQueryWrap.SelQueueRooms
    End If
    
    Exit Function
errH:
    MsgBoxD Me, "GetSelQueueRoomsִ���쳣��������Ϣ��" & err.Description, vbInformation, Me.Caption
End Function

Private Sub InitAutoRefresh()
'�����Զ�ˢ�£��������л���������ִ��Ĭ�Ϸ�����ִ��
On Error GoTo errH

    If mobjPacsQueryWrap Is Nothing Then Exit Sub
    If mobjPacsQueryWrap.SqlScheme Is Nothing Then Exit Sub
    
    If mobjPacsQueryWrap.SqlScheme.AutoRefreshTimeLen <= 0 Then
        TimerRefresh.Enabled = False
    Else
        mintAutoRefreshTimer = mobjPacsQueryWrap.SqlScheme.AutoRefreshTimeLen
        mintAutoRefreshTimerCount = mintAutoRefreshTimer
        TimerRefresh.Interval = 60000
        If App.LogMode = 0 Then
            TimerRefresh.Interval = 20000
        End If
        TimerRefresh.Enabled = True
    End If
    Exit Sub
errH:
    MsgBoxD Me, "��ȡ�Զ�ˢ�¼��ִ��ʧ�ܡ�������Ϣ��" & err.Description, vbInformation, Me.Caption
End Sub

Public Function GetBaseInfo(ByVal lngAdviceId As Long, Optional intMovedState As Integer = 0) As clsStudyInfo
    Set GetBaseInfo = mobjPacsQueryWrap.GetBaseInfo(lngAdviceId, intMovedState)
End Function

Private Sub QueueDataConsistency(ByVal lngAdviceId As Long, ByVal strRoom As String, ByVal intRowIndex As Integer)
'�Ŷ�����һ���Դ�����Ҫ��ִ�м�����
On Error GoTo errH
    Dim lngSendNo As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    '�Ŷ�����һ���Դ����жϼ�¼�����Ƿ���ڣ������ڣ�������¼�����ݣ�UpdateSourceData����
     '�ж��б��Ƿ��Ѿ���ʾ������ʾ��������б����ݣ�"ִ�м�"��
     '����Ҫ�������ݿ����ݣ����з��ͺ���Դ�����Ǽ�¼�����ݣ�Ҳ���ܴ����ݿ��е�����ѯ��
     
    '���ִ�м�����û�б仯����ֹ����
    If intRowIndex > -1 Then
        If mobjPacsQueryWrap.Text(intRowIndex, "ִ�м�") = strRoom Then
            Exit Sub
        End If
    End If

    strSQL = "select ���ͺ� from ����ҽ������ Where ҽ��ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��÷��ͺ�", lngAdviceId)
    If rsTemp.RecordCount = 1 Then
        lngSendNo = Val(nvl(rsTemp!���ͺ�))
    End If
        
    Call UpdateQueryListData(Nothing, lngAdviceId)
    
    '�������ݿ�����
    strSQL = "ZL_Ӱ�����¼_���Ͱ���(" & lngAdviceId & "," & lngSendNo & ",null,null,null,'" & strRoom & "',1)"
    Call zlDatabase.ExecuteProcedure(strSQL, "����ִ�м�")
    
    Exit Sub
errH:
    MsgBoxD Me, "QueueDataConsistency ������Ϣ��" & err.Description, vbInformation, Me.Caption
End Sub

Private Sub ReCreatCbrMenu(ObjCbrMain As CommandBars)
On Error GoTo errH
    Dim lngCount As Long
    
    Call LockWindowUpdate(Me.hwnd)
        
    For lngCount = ObjCbrMain.ActiveMenuBar.Controls.Count To 1 Step -1
        ObjCbrMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = ObjCbrMain.Count To 2 Step -1
        ObjCbrMain(lngCount).Delete
    Next
    
    Call InitCommandBars
    Call CreateWorkModuleMenu
    Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
    
    Call LockWindowUpdate(0)
    
    Exit Sub
errH:
    Call LockWindowUpdate(0)
    MsgBoxD Me, "ReCreatCbrMenuִ���쳣��������Ϣ��" & err.Description, vbInformation, Me.Caption
End Sub

Private Sub VsfListDbClick(ByVal blnIsLocate As Boolean)
On Error GoTo errhandle
    
    If Not blnIsLocate Then
        If vsfList.MouseRow = -1 Or vsfList.MouseRow = 0 Then Exit Sub
    End If
    
    If mobjCurStudyInfo.lngAdviceId <> 0 Then
        '˫�����˼���б�ʱ��������˼��״̬Ϊ �Ѿܾ���Ŀǰ�����κδ���
        If mobjCurStudyInfo.strStuStateDesc = "�Ѿܾ�" Then Exit Sub
        
        Select Case mobjCurStudyInfo.intStep
            Case 1, 0
                Call Menu_Manage_����
            Case 2, 3               '˫������д����,�����ʱ�����趨�Ƿ�򿪹�Ƭվ
                Call Menu_RichEPR(conMenu_Edit_Modify)
            Case -1, 4, 5               '˫���޶�����,�����ʱ�����趨�Ƿ�򿪹�Ƭվ
                Call Menu_RichEPR(conMenu_Edit_Audit)
            Case 6                  '����
                Call Menu_RichEPR(conMenu_File_Open)
        End Select
    End If

Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub MoveCtrHistroyFollow(ByVal lngX As Long, ByVal lngY As Long)
'PicFollowHistory rftHistoryFollow
On Error GoTo errH
    Dim lngNewW As Long
    Dim lngNewH As Long
    
    lngNewW = mlngpicHistoryOldW + (lngX - mlngPicHistoryX)
    If lngNewW > 8000 Then lngNewW = 8000
    If lngNewW < 1500 Then lngNewW = 1500
    
    lngNewH = mlngpicHistoryOldH + (lngY - mlngPicHistoryY)
    If lngNewH > 2500 Then lngNewH = 2500
    
    If lngNewH < 500 Then lngNewH = 500

    Call PicFollowHistory.Move(PicFollowHistory.Left, PicFollowHistory.Top, lngNewW, lngNewH)
    
    Call rftHistoryFollow.Move(50, 50, lngNewW - 100, lngNewH - 100)
    
    Exit Sub
errH:
End Sub

Private Function Is_ExistReportWriting(ByVal lngAdviceId As Long) As Boolean
'�Ƿ��б��洦���޶�״̬
On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "Select e.Id,  -null as ������, l.���汾 As �汾, '�����޶���' As ����, l.������ As ��Ա" & vbNewLine & _
            "From ���Ӳ�����¼ l," & vbNewLine & _
            "    (Select Max(c.��ʼ��) As ��ʼ��, Max(Id + 1) As Id,Max(������+1) as ������" & vbNewLine & _
            "     From ���Ӳ������� c ,����ҽ������ d" & vbNewLine & _
            "     Where c.�ļ�id = d.����id  And c.�������� = 8 and d.ҽ��id=[1]) e ,����ҽ������ f" & vbNewLine & _
            "Where L.ID =f.����id  And L.���汾 > e.��ʼ�� and f.ҽ��id=[1]"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ж��Ƿ�����޶��еı���", lngAdviceId)
    Is_ExistReportWriting = rsTemp.RecordCount > 0
            
    Exit Function
errH:
    err.Raise -1, , "�ж��Ƿ�����޶��еı����쳣" & vbCrLf & err.Description
End Function

Private Sub ChangeScheme(ByVal strName As String, ByVal lngID As Long, ByVal blnMenuClick As Boolean)
'blnMenuClick �Ƿ�˵������������������ָ��true: �˵����������  false�����tab�������
On Error GoTo errH
    Dim i As Integer
    Dim strResult As String
    
    If lngID <= 0 Then Exit Sub
    
    If Not mobjPacsQueryWrap Is Nothing Then
        For i = imgState.Count - 1 To 0 Step -1
            imgState(i).Visible = False
        Next
        
        labPatientInfo.Caption = ""
        labCollectionInfo.Caption = ""
        labPatientAge.Caption = ""
        
        strResult = mobjPacsQueryWrap.ExecuteMenu(lngID, mSysPar.blnQueryValidTime, mobjPacsQueryWrap.StartDate, mobjPacsQueryWrap.EndDate)
        Call InitAutoRefresh 'ExecuteMenu�����ִ��
        gblnXWMoved = mobjPacsQueryWrap.CurPacsQuery.IsMoved 'ExecuteMenu�����ִ��
        
        Call CheckHaveScheme(False, strResult)
        
        dkpMain.FindPane(1).title = strName
        
        If blnMenuClick Then Call mobjPacsQueryWrap.RefreshTabLeft(tabScheme, dkpMain.FindPane(1).title)
        
        Call AdjustFace(picList.Height, picList.Width)
        Call picDataSearchContainer_Resize
        Call ReSetFormFontSize
    End If
    Exit Sub
errH:
    err.Raise -1, , "ChangeScheme�쳣" & vbCrLf & err.Description
End Sub

Private Function GetMaxImgHeight() As Long
On Error GoTo errH
    Dim lngReturn As Long
    Dim i As Integer
    
    lngReturn = imgState(0).Height
    For i = 0 To imgState.Count - 1
        If lngReturn < imgState(i).Height Then lngReturn = imgState(i).Height
    Next
    
    GetMaxImgHeight = lngReturn
    Exit Function
errH:
    GetMaxImgHeight = imgState(0).Height
End Function

Private Sub SetParaUseImgSignValid(ByVal lngID As Long)
On Error GoTo errH
'���ݿ���ID����һ��ȫ�ֲ���
    gblUseImgSignValid = False
    If Len(GetSetting("ZLSOFT", "����ģ��\ZL9PACSWork", "����ͼ��ǩ����֤")) > 0 Then
        gblUseImgSignValid = GetSetting("ZLSOFT", "����ģ��\ZL9PACSWork", "����ͼ��ǩ����֤", "0") = "1" And GetSignVerifyType(lngID) = 1
    Else
        gblUseImgSignValid = GetDeptPara(lngID, "ͼ��ǩ����֤") = "1" And GetSignVerifyType(lngID) = 1
    End If
errH:
End Sub

Private Sub ViewLinkChecks()
'�鿴������飬���ݵ�ǰ����ID �鵽���еĹ���ҽ��ID,Ȼ��ʹ�� ExecuteWithLink(strLink)
'1 ȫ������ID�ļ����Ϣ
'2 ���ʹ���˹�����飬���й�������ҽ��ID
On Error GoTo errH
    Dim strSQL As String
    Dim rsData As Recordset
    Dim strAppend As String
    Dim strLink As String
    Dim i As Long
    
    If mobjCurStudyInfo Is Nothing Then Exit Sub
    
    If mSysPar.blnRelatingPatient And mobjCurStudyInfo.lngLinkId > 0 Then
        strLink = "Select A.ID as ҽ��ID From ����ҽ����¼ A Where A.����id = " & mobjCurStudyInfo.lngPatId & " UNION ALL Select ҽ��ID  from Ӱ�����¼ Where ����ID =" & mobjCurStudyInfo.lngLinkId & ""
    Else
        strLink = "Select A.ID as ҽ��ID From ����ҽ����¼ A Where A.����id = " & mobjCurStudyInfo.lngPatId & ""
    End If
        
    Call mobjPacsQueryWrap.ExecuteWithLink(strLink)
    TimerRefresh.Enabled = False
    
    For i = 1 To vsfList.Rows - 1
        vsfList.TextMatrix(i, 0) = i
    Next

    Exit Sub
errH:
    err.Raise -1, , "ViewLinkChecks�쳣" & vbCrLf & err.Description
End Sub

Private Sub LocateMainWorkModuleTab()
On Error GoTo errH
'�ָ���Ҫ����ҳ�棬�����������Ҫ����ҳ�棬�л����ʱ�����л�����Ӧҳ�� And Not TabWindow.Item(i).Selected
    Dim i As Integer
    
    If Len(mSysPar.strFirstTab) <= 0 Then Exit Sub
    
    For i = 0 To TabWindow.ItemCount - 1
        If InStr(TabWindow(i).tag, mSysPar.strFirstTab) > 0 And TabWindow.Item(i).Visible Then
            If Not TabWindow.Item(i).Selected Then
                TabWindow.Item(i).Selected = True
                Exit Sub
            End If
        End If
    Next
errH:
End Sub

Private Sub ViewEPRPetition(ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal lngClinicId As Long, ByVal intSourceType As Long)
On Error GoTo errH
    Dim rsTemp As ADODB.Recordset, strBillNo As String, strExseNo As String, intExseKind As Integer
    Dim strSQL As String
    
    
    strSQL = "select NO,��¼���� from ����ҽ������ where ҽ��ID=[1] and ���ͺ�=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡNO", lngAdviceId, lngSendNo)
    If rsTemp.EOF Then Exit Sub
    
    strExseNo = rsTemp!no: intExseKind = rsTemp!��¼����
    
    strSQL = "Select B.ͨ��,B.ID, B.���" & vbNewLine & _
            "From ��������Ӧ�� A, �����ļ��б� B" & vbNewLine & _
            "Where A.������Ŀid =[1] And A.Ӧ�ó��� =[2] And A.�����ļ�id = B.ID And B.���� = 7"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���ݱ��", lngClinicId, CLng(Decode(intSourceType, 1, 1, 2, 2, 1)))
    
    If rsTemp.EOF Then Exit Sub
    
    strBillNo = "ZLCISBILL" & Format(rsTemp!���, "00000") & "-1"
    
    ReportOpen gcnOracle, glngSys, strBillNo, Me, "NO=" & strExseNo, "����=" & intExseKind, "ҽ��ID=" & lngAdviceId, 1
    Exit Sub
errH:
    err.Raise -1, , "ViewEPRPetition�쳣" & vbCrLf & err.Description
End Sub

