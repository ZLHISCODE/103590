VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "*\A..\ZLIDKind\ZLIDKIND.vbp"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPacsMainV2 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Ӱ����վ"
   ClientHeight    =   10575
   ClientLeft      =   8535
   ClientTop       =   870
   ClientWidth     =   15240
   Icon            =   "frmPacsMainV2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10575
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timerHelper 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   6000
      Top             =   600
   End
   Begin VB.PictureBox picHelper 
      BorderStyle     =   0  'None
      Height          =   8055
      Left            =   6720
      ScaleHeight     =   8055
      ScaleWidth      =   3375
      TabIndex        =   35
      Top             =   1440
      Width           =   3375
      Begin zl9PACSWork.ucPacsHelper ucPacsHelper1 
         Height          =   8175
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   14420
      End
   End
   Begin VB.PictureBox picTabFace 
      BackColor       =   &H00DBE2E3&
      BorderStyle     =   0  'None
      Height          =   395
      Left            =   6240
      ScaleHeight     =   390
      ScaleWidth      =   1335
      TabIndex        =   33
      Top             =   0
      Width           =   1335
      Begin XtremeSuiteControls.TabControl TabWindow 
         Height          =   315
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   1245
         _Version        =   589884
         _ExtentX        =   2196
         _ExtentY        =   556
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.ImageList img24 
      Left            =   1440
      Top             =   600
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
            Picture         =   "frmPacsMainV2.frx":1CFA
            Key             =   "PACS����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":2474
            Key             =   "��Ƭ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":2BEE
            Key             =   "PACS��д"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":3368
            Key             =   "PACS���"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":3AE2
            Key             =   "PACS�鿴������Ϣ"
         EndProperty
      EndProperty
   End
   Begin VB.Timer timFun 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   5400
      Top             =   600
   End
   Begin VB.PictureBox PicFucs 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4080
      ScaleHeight     =   855
      ScaleWidth      =   2175
      TabIndex        =   26
      Top             =   1440
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
      Left            =   4200
      Top             =   600
   End
   Begin VB.PictureBox picExtra 
      BackColor       =   &H00E0E0E0&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2355
      ScaleWidth      =   2715
      TabIndex        =   9
      Top             =   7080
      Width           =   2775
      Begin RichTextLib.RichTextBox rtxtAppend 
         Height          =   1575
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2778
         _Version        =   393217
         BackColor       =   14737632
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmPacsMainV2.frx":41DC
      End
   End
   Begin VB.PictureBox picDataSearchContainer 
      BackColor       =   &H00E0E0E0&
      Height          =   2415
      Left            =   3000
      ScaleHeight     =   2355
      ScaleWidth      =   3435
      TabIndex        =   8
      Top             =   7080
      Width           =   3495
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
         TabIndex        =   15
         Top             =   -2520
         Width           =   5200
      End
      Begin VB.CommandButton cmdMore 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPacsMainV2.frx":4279
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Picture         =   "frmPacsMainV2.frx":472F
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "frmPacsMainV2.frx":4C21
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "��ѯ"
         Top             =   120
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Timer timerCapture 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   4800
      Top             =   600
   End
   Begin VB.PictureBox picWindow 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   10320
      ScaleHeight     =   4575
      ScaleWidth      =   5175
      TabIndex        =   2
      Top             =   3240
      Width           =   5175
   End
   Begin VB.Timer TimerRefresh 
      Enabled         =   0   'False
      Left            =   3600
      Top             =   600
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
            Picture         =   "frmPacsMainV2.frx":52F3
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
      Left            =   2880
      Top             =   600
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
            Picture         =   "frmPacsMainV2.frx":5B87
            Key             =   "����"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":6C31
            Key             =   "����"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":7CDB
            Key             =   "���"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":8D85
            Key             =   "��д"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":9E2F
            Key             =   "���"
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":AED9
            Key             =   "���"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":BF83
            Key             =   "���"
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":D02D
            Key             =   "����"
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":E0D7
            Key             =   "�ܾ�"
            Object.Tag             =   "9"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2160
      Top             =   600
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
            Picture         =   "frmPacsMainV2.frx":F181
            Key             =   "��ѡ����"
            Object.Tag             =   "90000"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":F71B
            Key             =   "��ѡѡ��"
            Object.Tag             =   "90001"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":FCB5
            Key             =   "��ѡ����"
            Object.Tag             =   "90002"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":103C7
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
      Height          =   5580
      Left            =   0
      ScaleHeight     =   5580
      ScaleWidth      =   6540
      TabIndex        =   1
      Top             =   1320
      Width           =   6540
      Begin XtremeSuiteControls.TabControl tabScheme 
         Height          =   735
         Left            =   4440
         TabIndex        =   32
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
         Picture         =   "frmPacsMainV2.frx":10AD9
         Style           =   1  'Graphical
         TabIndex        =   30
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
         Picture         =   "frmPacsMainV2.frx":10F0B
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "����"
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.PictureBox pic�������ڵ� 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   4440
         ScaleHeight     =   1095
         ScaleWidth      =   1455
         TabIndex        =   24
         Top             =   3600
         Width           =   1455
         Begin VB.Label labNoScheme 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            TabIndex        =   25
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3120
         Width           =   5775
      End
      Begin VB.PictureBox picDetail 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   3735
         TabIndex        =   11
         Top             =   3600
         Width           =   3735
         Begin VB.Label labPatientAge 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            TabIndex        =   31
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
            TabIndex        =   27
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
            BackStyle       =   0  'Transparent
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
            TabIndex        =   23
            Top             =   480
            Width           =   75
         End
         Begin VB.Label labPatientInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
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
      Begin XtremeSuiteControls.TabControl TabExtra 
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   4560
         Width           =   3735
         _Version        =   589884
         _ExtentX        =   6588
         _ExtentY        =   1085
         _StockProps     =   64
      End
      Begin VB.PictureBox picTemp 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   435
         TabIndex        =   6
         Top             =   3240
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox picFilter 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   2895
         TabIndex        =   5
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
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   4
         ToolTipText     =   "��û��ʲô��"
         Top             =   3240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Bindings        =   "frmPacsMainV2.frx":1133D
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   3735
         _cx             =   6588
         _cy             =   2990
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
         TabIndex        =   10
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
         IDKindStr       =   $"frmPacsMainV2.frx":11365
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
      Bindings        =   "frmPacsMainV2.frx":11418
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPacsMainV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

#Const DebugImmediately = False

Implements IEventNotify

Private Const C_MODULE_NAME = "frmPacsMainV2"
'Private Const C_HISTORY_VIEW_TAG = "-��"        '��ʷ���ݲ鿴���

Private Const C_LAYOUT_BASEHEIGHTOFTAB As Long = 5000 ' ������Ϣ5000
Private Const C_LAYOUT_BASEHEIGHTOFDETAILINFO As Long = 800 ' ��ϸ��Ϣ��׼�߶�5000

Private Const C_LNG_TAB_MENU_ID = 123456780

Private Const C_STEPIMG_�Ǽ� As String = "����" '
Private Const C_STEPIMG_���� As String = "����" '
Private Const C_STEPIMG_��� As String = "���" '
Private Const C_STEPIMG_�ܾ� As String = "�ܾ�" '
Private Const C_STEPIMG_���� As String = "����" '
Private Const C_STEPIMG_��� As String = "���" '
Private Const C_STEPIMG_��� As String = "���" '
Private Const C_STEPIMG_��� As String = "���" '
Private Const C_STEPIMG_��д As String = "��д" '


'��ʽ��ʽ
Public Enum TColorStyle
    sBlue = 0   '��ɫ
    sAshen = 1  '�Ұ�
    sGray = 2   '��ɫ
    sBlack = 3  '��ɫ
    sSys = 4    'ϵͳ
End Enum

Private Type TWorkModuleInfo
'����ģ����Ϣ
    ModuleName As String
    objModule As Object
    hwnd As Long
    FontSize As Double
    DeptId As Long      '��������
    tag As String           '������־
End Type

Private mAryWorkModule() As TWorkModuleInfo

Private mobjCurStudyInfo As New clsStudyInfo  '���ڲ����ļ����Ϣ
Private mstrFirstTab As String '�״���ʾ��ҳ��
Private mlngMove As Long
Private mintQueryState As Integer '��ѯ����״̬  0 δ��ʼ��  ��1 ����  ��2 û���κ���Ч����   3��û���Ѿ����õķ���
Private mblHistory As Boolean '�Ƿ����μ��
Private mblHaveHistory As Boolean '������ʷ���
Private mintAutoRefreshTimer As Integer '�Զ�ˢ�¼�ʱ����
Private mintAutoRefreshTimerCount As Integer '�Զ�ˢ�¼�ʱ����

'---------------------------------------------------
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
    blnֱ�Ӽ�� As Boolean                               '�ǼǺ�ֱ�ӽ�����
    blnWriteCapDoctor As Boolean                        '�Ƿ��ڲɼ�ͼ����Զ��ѵ�ǰ�û���дΪ��鼼ʦ
    blnAutoOpenReport As Boolean                        '��ʼ����Զ��򿪱���
    blnChoosePrintFormat As Boolean                     '�Ƿ񱨵���ӡǰѡ���ʽ
    strLocalRoom As String                              '����ִ�м�����
    lngImageValid As Long                               'ͼ��У��
    lngAutoImageDays As Long                            '�Զ�����ʷͼ���������Χ
    
    '���̲���
    blnCompleteCommit As Boolean                        '��˺������ٴ�ȷ��
    blnFinallyCompleteCommit As Boolean                 '�����ֱ�����
    blnIgnoreResult As Boolean                          '���������� '=true ����
    
    blnReportWithImage As Boolean                       '��ͼ�����д���棬��ͼ�񲻿�д����
    blnNoSignFinish As Boolean                              '����δǩ�������ӡ���
    blnReportWithResult As Boolean                      '�������Խ������д����
    
    blnPrintCommit As Boolean                           '��ӡ��ֱ�����
    blnCanPrint As Boolean                              'ƽ����Ҫ��˲��ܴ�ӡ =true
    blnAuditAutoPrint As Boolean                        '�����ֱ�Ӵ�ӡ
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
    blnAutoPrint As Boolean                             '�������Զ���ӡ���뵥
    blnAutoPrintCheck As Boolean                        '�Զ������ظ���ӡ
    blnDirectSendRepImg As Boolean                      'ֱ�ӽ���Ƭ��ͼ���͵�����
    
    blnShowImgAfterReport As Boolean                    '����ʱ��Ƭ
    blnIsLocateReport As Boolean
    blnPEISNoCheckMoneyFinish  As Boolean    '����鱨����ɲ��жϷ���
    blnQuickTabDisplayScheme  As Boolean    '���ÿ��tab��ǩչʾ����
    lngReportType As Long
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


Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long


Private mintInterface() As TInterface   '�Զ�ִ�еĲ��
Private mintInterfaceCount As Integer '��Ҫ�Զ�ִ�еĲ��������1 ��ʼ����

Private mintToolBarWriteReg As Integer        '������ע���״ֵ̬

Private mstrPrivs As String, mlngModule As Long              'ģ��ţ���ģ��Ȩ��

Private WithEvents mfrmRISRequest As frmRISRequest
Attribute mfrmRISRequest.VB_VarHelpID = -1

'��Ϣ��������
Private WithEvents mobjMsgCenter As clsPacsMsgProcess
Attribute mobjMsgCenter.VB_VarHelpID = -1

'����ģ�������ˢ��ģʽ�����������
'1.����ģ��ֻҪ���ڣ�ǿ�ƶ����е����ݽ���ˢ��
'2.����ģ������ʾʱ���Ŷ����е����ݽ���ˢ��
'3.����ģ����������ݱ仯ʱ����ʾ��ģ���ǵ�ǰģ�飬�Ŷ����е����ݽ���ˢ��

Private mobjWork_PacsImg As frmWork_ImageV2             'Ӱ���Ӵ���
Attribute mobjWork_PacsImg.VB_VarHelpID = -1
Private mobjWork_Pathol As clsWorkModule_PatholV2       '�������ģ��
Private mobjWork_His As clsWorkModule_HisV2             'HIS���ģ��
Private mobjWork_Report As clsWorkModule_ReportV2       '����ģ��

Private mobjWork_ImageCap As zl9PacsImageCap.clsPacsCaptureV2  '��Ƶ�ɼ�ģ��
Private mobjRichReportWrap As frmEPREditWrapV2

Private WithEvents mobjCapLinker As clsCapLinker
Attribute mobjCapLinker.VB_VarHelpID = -1


Private WithEvents mobjPacsCore As zl9PacsCore.clsViewer        '��Ƭվ����
Attribute mobjPacsCore.VB_VarHelpID = -1
Private WithEvents mobjQueue As frmWork_Queue                   '�Ŷӽк�
Attribute mobjQueue.VB_VarHelpID = -1



Private mobjSelModule As Object
Private mlngSelHwnd As Long
Private mstrSelTabName As String
Private mstrSelModuleTag As String
Private mobjAppendBill As Object


Private WithEvents mobjPacsQueryWrap As clsPacsQueryWrap      '�Զ����ѯ���ܷ�װ��
Attribute mobjPacsQueryWrap.VB_VarHelpID = -1

'���ڱ���
Private mlngCur����ID As Long                               '��ǰ����ID
Private mstrCur���� As String                               '��ǰ���� ����-����
Private mstrCanUse���� As String                            '��ǰ���ÿ���  ID_����-����

Private mblnInitOk As Boolean   '��ʼ�����,װ�ر��
Private mblnAllDepts As Boolean                             '�Ƿ�ѡ��ȫ������
Private mstrCanUse����IDs As String                         '��ǰ���õĿ���ID�����á������ָ�������ֱ����ΪSQL��ѯ����
Private mblnMenuDownState As Boolean                        '����˫����������������
Private mblnIsHasPatholModule As Boolean                   '�Ƿ������˲���ģ��

Private mblnFormLoadState As Boolean
Private mblnIsScheduleDept As Boolean                       '��ǰѡ�п��ң��Ƿ�����ԤԼ
Private mblnIsScheduleOrder As Boolean                      '��ǰ����Ƿ�����ԤԼ������ԤԼ�豸�ж�

Private mblnIsPrintMode As Boolean                          '�Ƿ����嵥��ӡ

Private mstrDefaultPatientType As String                    'ȱʡ��������

Private mstrRPTExecutor As String                           '����ѡ��ı�����
Private mblnLockState As Boolean                           '�Ƿ����û���������״̬

'���̿��Ʊ���
Private mSysPar As TSystemPar                               'ϵͳ����

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

'˫�û���¼
Private mcnOracleHIS As New ADODB.Connection    '��¼HIS����̨��½ʱʹ�õ����ݿ����Ӵ�
Private mstrHisUserName As String               '��¼HIS����̨��½ʱʹ�õ��û���
Private mstrHisUserID As String                 '��¼HIS����̨��¼ʱʹ�õ��û�ID
Private mstrOtherUserName As String             '��¼˫�û���½�ĵڶ����û���
Private mstrOtherUserID As String               '��¼˫�û���¼�ĵڶ����û�ID
Private mblnCnOracleIsHIS As Boolean            '��ǰ���ݿ������Ƿ�HIS����̨������
Private mintChangeUserState As Integer          '��¼�û������������1- ͳһ��2-����

'�ղع���
Private mlngShareFatherID As Long
Private mlngCollectionFatherID As Long
Private mblnIsLoading As Boolean
 
Private mblnIsForceRefresh As Boolean          '�Ƿ����ģ��ǿ��ˢ�²���

Private mobjPublicAdvice As Object
Private mobjMedicalRecord As Object
Private mblnIsValid As Boolean                  '��������Ƿ���Ч

Private mintState As Integer
Private mblnIsHistoryMode As Boolean            '�Ƿ���ʷ״̬
Private mblnIsHideStudyList As Boolean
Private mblnIsHideHelper As Boolean


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



'***********************************************IEventNotifyʵ��***********************************************


Public Function MainWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo errhandle
    Dim strLog As String
    Dim strResult As String
    
    '��Ϣ����
    Select Case uMsg
        Case WM_XWREPORT_IMG
            strLog = Now & " umsg = " & uMsg & ";wparam = " & wParam & ";lparam = " & lParam & vbCrLf
    
            If gblnXWLog Then Call WriteCommLog("XWWindowProc", "XW�ӿ�", strLog)
            
            '�����������͵�ϵͳ������ı���ͼ��
            If lParam <> 0 Then
                If gblnXWLog Then Call WriteCommLog("XWWindowProc", "XW�ӿ�", "���뱨��ͼ������̡�")
    
                strResult = XWSaveReportImagesV2(Me, lParam)
                
                If Len(strResult) <= 0 Then Exit Function
                
                Call IEventNotify_Broadcast(uMsg, , lParam, strResult)
                
            End If
        Case Else
            Call IEventNotify_SendRequest(uMsg, , lParam)
            
    End Select
    
Exit Function
errhandle:
    Notify.PrintErr err, infWaring, , C_MODULE_NAME, "MainWindowProc"
End Function


Property Get Notify() As IEventNotify
    Set Notify = Me
End Property

Public Function IEventNotify_Owner() As Object
    Set IEventNotify_Owner = Me
End Function


Public Function IEventNotify_Hwnd() As Long
    IEventNotify_Hwnd = hwnd
End Function


Public Function IEventNotify_MainPrivs() As String
'��ȡȨ�޴�
    IEventNotify_MainPrivs = mstrPrivs
End Function


Public Function IEventNotify_PrintErr(objErr As ErrObject, ByVal lngInfoType As Long, _
    Optional ByVal lngHwnd As Long = 0, Optional ByVal strUnitName As String = "", Optional ByVal strMethodName As String = "") As Long
    IEventNotify_PrintErr = IEventNotify_PrintInfo(objErr.Description, lngInfoType, lngHwnd, strUnitName, strMethodName)
End Function

Public Function IEventNotify_PrintInfo(ByVal strErr As String, ByVal lngInfoType As Long, _
    Optional ByVal lngHwnd As Long = 0, Optional ByVal strUnitName As String = "", Optional ByVal strMethodName As String = "") As Long
'��ӡ������Ϣ
'����
'
'���Դ���.
'[1290-ZLHIS] [0227 13:36:46] [frmPacsMain.PrintErr]
'
'On Error GoTo errHandle

    Dim strMsg As String
    
    IEventNotify_PrintInfo = 0
    
    strMsg = strErr & vbCrLf & "[" & mlngModule & "-" & UserInfo.�û��� & "] [" & Format(Now, "mmdd hh:mm:ss") & "] [" & strUnitName & "." & strMethodName & "]"
    
    Debug.Print strMsg
    OutputDebugString strMsg
    
    Select Case lngInfoType
        Case infNone
            '��ִ���κβ���
            
        Case infHint
            If lngHwnd = 0 Then
                MsgBoxD Me, strErr, vbOKOnly, "��ʾ"
            Else
                MsgboxH lngHwnd, strErr, vbOKOnly, "��ʾ"
            End If
            
        Case infWaring
            If lngHwnd = 0 Then
                MsgBoxD Me, strErr, vbOKOnly, "����"
            Else
                MsgboxH lngHwnd, strErr, vbOKOnly, "����"
            End If
            
        Case infNormalErr
            If lngHwnd = 0 Then
                MsgBoxD Me, strMsg, vbOKOnly, "����"
            Else
                MsgboxH lngHwnd, strMsg, vbOKOnly, "����"
            End If
            
        Case infDataErr
            IEventNotify_PrintInfo = IIf(ErrCenter() = 1, True, False)
            Call SaveErrLog
            
        Case Else
            If lngHwnd = 0 Then
                IEventNotify_PrintInfo = MsgBoxD(Me, strErr, lngInfoType, "��ʾ")
            Else
                IEventNotify_PrintInfo = MsgboxH(lngHwnd, strErr, lngInfoType, "��ʾ")
            End If
            
    End Select
    
'Exit Function
'errHandle:
'    Debug.Print "IEventNotify_PrintErr Err:" & err.Description
End Function


Private Sub IEventNotify_SendRequest(ByVal lngEventNo As Long, Optional ByVal strTag As String = "", _
    Optional data1, Optional data2, Optional data3, Optional strExPro As String = "")
'lngEventNo:�¼���
'strTag:�¼����
'data:��������
    Dim lngSendNo As Long
    Dim blnIsMoved As Boolean
    Dim strStudyUID As String
    
    Dim curData1
    Dim curData2
    Dim curData3
    Dim strCurExPro As String
    
On Error GoTo errhandle
    If IsError(data1) = False Then curData1 = data1
    If IsError(data2) = False Then curData2 = data2
    If IsError(data3) = False Then curData3 = data3
    
    strCurExPro = strExPro
    
    Select Case lngEventNo
        Case WM_LIST_SYNCROW    'ͬ������ʾ
            Call UpdateQueryListData(Nothing, curData1)
            
        Case WM_LIST_MOVEUP '����
            If vsfList.Row > 1 Then vsfList.Row = vsfList.Row - 1
            
        Case WM_LIST_MOVEDOWN '����
            If vsfList.Row + 1 < vsfList.Rows Then vsfList.Row = vsfList.Row + 1
            
        Case WM_LIST_GETLASTADVICE
            'data1���ݵ�ǰʹ�õ�ҽ��ID
            data1 = TraversalAdvice(data1, False, lngSendNo, blnIsMoved)
            data2 = lngSendNo
            data3 = blnIsMoved
        
        Case WM_LIST_GETNEXTADVICE
            'data1���ݵ�ǰʹ�õ�ҽ��ID
            data1 = TraversalAdvice(data1, True, lngSendNo, blnIsMoved)
            data2 = lngSendNo
            data3 = blnIsMoved
        
        Case WM_IMG_OPENVIEW
            '��Ƭ
            If mobjPacsCore Is Nothing Then
                HintMsg "ͼ��鿴������Ч�����ܽ��д˲�����", "cbrMain_Execute", vbOKOnly
                Exit Sub
            End If
            
            If mobjCurStudyInfo.lngAdviceId = Val(curData1) Then
                If mobjCurStudyInfo.strStudyUID <> "" Then
                    Call OpenViewer(1, mobjPacsCore, mobjCurStudyInfo.lngAdviceId, False, Me, "", mobjCurStudyInfo.blnMoved)
                Else
                    '�򿪹����ļ��ͼ��
                    Call OpenLatestImage(Me, mobjPacsCore, mobjCurStudyInfo, mSysPar.lngAutoImageDays)
                End If
            Else
                Call GetSendNo(Val(curData1), strStudyUID)
                
                If Len(strStudyUID) <= 0 Then
                    '�򿪹����ļ��ͼ��
                    Call OpenLatestImage(Me, mobjPacsCore, GetBaseInfo(curData1), mSysPar.lngAutoImageDays)
                Else
                    Call OpenViewer(1, mobjPacsCore, curData1, False, Me)
                End If
            End If
            
        Case WM_IMG_CONTRASTVIEW
            '�Ա�
            If mobjPacsCore Is Nothing Then
                HintMsg "ͼ��鿴������Ч�����ܽ��д˲�����", "cbrMain_Execute", vbOKOnly
                Exit Sub
            End If
            
            Call OpenViewer(1, mobjPacsCore, curData1, True, Me)
            
        Case WM_REPORT_VIEW
            '����Ԥ��
            Call ReoprtPrint(curData1, curData2, False, curData3)
            
        Case WM_REPORT_PRINT
            '�����ӡ
            Call ReoprtPrint(curData1, curData2, True, curData3)
            
    End Select
    
Exit Sub
errhandle:
    HintError err, "SendRequest", False
End Sub

Private Function GetSendNo(ByVal lngAdviceId As Long, _
    Optional ByRef strStudyUID As String, _
    Optional ByRef strRecDate As String, _
    Optional ByRef lngStep As Long) As Long
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    GetSendNo = 0
    strStudyUID = ""
    
    strSQL = "Select a.���ͺ�,a.ִ�й���, b.���UID, to_Char(B.��������,'YYYYMMDD')  as �������� " & _
            " From ����ҽ������ a, Ӱ�����¼ B " & _
            " where a.ҽ��ID=b.ҽ��ID(+) and a.���ͺ�=b.���ͺ�(+) and a.ҽ��Id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ͺ�", lngAdviceId)
    
    '�����漰��ת�����ݽ��й�Ƭ�Ȳ����������Ҫ��ѯת���������
    If rsData.RecordCount <= 0 Then
        '����ʷ���ٽ��в�ѯ
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
        
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ͺ�", lngAdviceId)
    End If
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    strStudyUID = NVL(rsData!���UID)
    strRecDate = NVL(rsData!��������)
    lngStep = Val(NVL(rsData!ִ�й���))
    
    GetSendNo = Val(NVL(rsData!���ͺ�))
End Function

Private Function GetLastSignInfo(ByVal lngAdviceId As Long, Optional ByRef strCreateUser As String) As TReportSignInfo
'�ù��̽�֧��δת�������ݲ�ѯ����resetstate�����б����ã���ȡ���ǩ����Ϣ���л��˺������ش���
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim reportSignInfo As TReportSignInfo
    Dim lngSendNo As Long
    Dim strStudyUID As String
    Dim dblEPRID As Double
    Dim lngLastVer As Long
    
    reportSignInfo.ID = 0
    
    strCreateUser = UserInfo.����
    
    strSQL = "Select a.ҽ��id,a.����id, b.������, b.������, b.ǩ������, b.���ʱ��, b.���汾 " _
             & "  From ����ҽ������ a, ���Ӳ�����¼ b Where a.ҽ��id = [1] And a.����id = b.Id"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ȡǩ����Ϣ", lngAdviceId)
    If rsData.RecordCount <= 0 Then Exit Function
     
    strCreateUser = NVL(rsData!������)
    dblEPRID = Val(NVL(rsData!����Id))
    lngLastVer = Val(NVL(rsData!���汾))
             
    '����˵��������Ҫ����ǩ���ˡ��������ݵ�����������˲�һ����ǩ���ˣ����Ҫ�������һ��ǩ����
    strSQL = "Select ID  From ���Ӳ������� Where �ļ�ID=[1] And ��������= 8 And ��ʼ�� = [2] "
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ȡǩ������", dblEPRID, lngLastVer)
    If rsData.RecordCount <= 0 Then
        '�ж�ǰ�ΰ汾�Ƿ�Ϊǩ��
        If lngLastVer > 1 Then
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ȡǩ������", dblEPRID, lngLastVer - 1)
        Else
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ȡǩ������", dblEPRID, lngLastVer)
        End If
        
        
        If rsData.RecordCount <= 0 Then Exit Function
    End If
    
    Call GetReportSignInfo(Val(NVL(rsData!ID)), reportSignInfo, False)
    
     
    GetLastSignInfo = reportSignInfo
End Function

Public Function IEventNotify_StudyInfo() As clsStudyInfo
'��ȡ�����Ϣ
    Set IEventNotify_StudyInfo = mobjCurStudyInfo
End Function

Private Sub IEventNotify_Broadcast(ByVal lngEventNo As Long, Optional ByVal strTag As String = "", _
    Optional data1, Optional data2, Optional data3, Optional strExPro As String = "")
     
    Dim curData1
    Dim curData2
    Dim curData3
    Dim strCurExPro As String
    
    Dim blnSyncRow As Boolean
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    Dim curSignInfo As TReportSignInfo
    Dim lngSendNo As Long
    Dim strStudyUID As String
    Dim lngStep As Long
    Dim blnOk As Boolean

'On Error GoTo errHandle:
    If IsError(data1) = False Then curData1 = data1
    If IsError(data2) = False Then curData2 = data2
    If IsError(data3) = False Then curData3 = data3
    
    strCurExPro = strExPro
    
    blnSyncRow = False
    
    If Val(strTag) = 0 Then
        Call ExecutePlugin(lngEventNo, 0, curData1, curData2, curData3)
    End If
    
    Select Case lngEventNo
        Case BM_REPORT_EVENT_PRINT  '�����ӡ�¼�*****************************************************************************
            'curData1��ʾҽ��ID
            'curData2��ʾ����ID ��δʹ��
            'curData3�༭������ '-1��ʾ�����༭��
            
            '���strTag=0, curData3���ʾ�����Ƿ�����ִ�д�ӡ����
            
            If Val(curData3) = -1 Then
                If mSysPar.blnIgnoreResult = False Then Call ReportResultHint(curData1)
                
                '�����ת�����ݽ��д�ӡ����Ӧ�ý��б�Ǹ��£�ִ�д˴洢���̲����ѯ������
                strSQL = "ZL_Ӱ�񱨸��ӡ_Update(" & curData1 & ")"
                zlDatabase.ExecuteProcedure strSQL, "���´�ӡ���"
    
                blnSyncRow = True
                
                If mSysPar.blnPrintCommit = True Then   '��ӡ��ֱ�����
                    If Menu_Manage_����������(Val(curData1), False) Then blnSyncRow = False '�����ڲ����Զ��Լ���б��н��и���
                End If
                
                
            Else
                '������ǲ����༭������Ҫ�ж��Ƿ��ӡǰ
                If Val(strTag) <> 1 Then    '��ӡǰ
                    If mSysPar.blnIgnoreResult = False And mSysPar.lngHintType = 2 Then Call ReportResultHint(curData1): blnSyncRow = True
                Else
                    '����Ǵ�ӡ�󣬾���Ҫ��֤�Ƿ�����ʾ
                    Call ReportResultHint(curData1)
                    
                    strSQL = "ZL_Ӱ�񱨸��ӡ_Update(" & curData1 & ")"
                    zlDatabase.ExecuteProcedure strSQL, "���´�ӡ���"
                    
                    blnSyncRow = True
                    
                    If mSysPar.blnPrintCommit = True Then   '��ӡ��ֱ�����
                        If Menu_Manage_����������(Val(curData1), False) Then blnSyncRow = False
                    End If
                    
                End If
            End If
            
'        Case BM_RIS_EVENT_COMPLETE  '�������¼�*****************************************************************************
'            If Val(strTag) = 1 Then
'                '���ͼ�������Ϣ
'                Call mobjMsgCenter.Send_Msg_StudyComplete(objStudyInfo.lngAdviceId, strReportId)
'
'                blnSyncRow = True
'            End If
            
'        Case BM_RIS_EVENT_CANCELCOMP    'ȡ���������¼�*****************************************************************************
'            If Val(strTag) = 1 Then
'                '���ͼ�鳷�������Ϣ
'                Call mobjMsgCenter.Send_Msg_CancelComplete(mobjCurStudyInfo.lngAdviceId)
'
'                blnSyncRow = True
'            End If
            
            
        Case BM_REPORT_EVENT_AUDIT   '���ǩ��֪ͨ*****************************************************************************
            'curData1 ��ʾҽ��ID
            'curData2 ��ʾ����ID ��δʹ��
            'curData3 ��ʾ�༭������
            
            If Val(curData3) = -1 Or Val(strTag) = 1 Then
                If mSysPar.blnIgnoreResult = False And mSysPar.lngHintType = 1 Then Call ReportResultHint(curData1)
                 
                '���������
                lngSendNo = GetSendNo(curData1, strStudyUID)
                Call ResetState(curData1, lngSendNo, strStudyUID)
                
                blnOk = False
                If mSysPar.blnAuditAutoPrint Then '�����ֱ�Ӵ�ӡ
                    '��˺��ӡ�������ǰ�����ܹ�������ˣ�˵��û�н���ת��
                    If ReoprtPrint(Val(curData1), False, True, , strExPro) Then
                        '���´�ӡ״̬
                        strSQL = "ZL_Ӱ�񱨸��ӡ_Update(" & curData1 & ")"
                        zlDatabase.ExecuteProcedure strSQL, "���´�ӡ���"
                        
                        blnOk = True
                    End If
                End If
                
                blnSyncRow = True
                If mSysPar.blnCompleteCommit Then   '�������˺�ֱ����ɡ�
                    If Menu_Manage_����������(Val(curData1), False) Then blnSyncRow = False
                Else
                    If mSysPar.blnPrintCommit = True And blnOk Then   '��ӡ��ֱ�����
                        If Menu_Manage_����������(Val(curData1), False) Then blnSyncRow = False
                    End If
                End If
                
                '����״̬ͬ����Ϣ
                Call mobjMsgCenter.Send_Msg_StateSync(curData1)
                
                
            End If
            
        Case BM_REPORT_EVENT_SIGN   '���ǩ��֪ͨ*****************************************************************************
            'curData1 ��ʾҽ��ID
            'curData2 ��ʾ����ID ��δʹ��
            'curData3 ��ʾ�༭������
            
            If Val(curData3) = -1 Or Val(strTag) = 1 Then
                If mSysPar.blnIgnoreResult = False And mSysPar.lngHintType = 0 Then Call ReportResultHint(curData1)
                
                '���汨����
                lngSendNo = GetSendNo(curData1, strStudyUID)
                Call ResetState(curData1, lngSendNo, strStudyUID)
                
                '����״̬ͬ����Ϣ
                Call mobjMsgCenter.Send_Msg_StateSync(curData1)
                
                blnSyncRow = True
            End If
            
        Case BM_REPORT_EVENT_SAVE, BM_REPORT_EVENT_POPUPEXIT   '���汣��֪ͨ*****************************************************************************
            'curData1 ��ʾҽ��ID
            'curData2 ��ʾ����ID ��δʹ��
            'curData3 ��ʾ�༭������
            
            If Val(curData3) = -1 Or Val(strTag) = 1 Then
                
                '���汨����
                lngSendNo = GetSendNo(curData1, strStudyUID)
                Call ResetState(curData1, lngSendNo, strStudyUID)
                
                '����Ӱ����ͼ��ı���ͼ��ǣ����ݱ����б���ı���ͼ����Ӱ����ͼ���¼
                strSQL = "Zl_Ӱ����ͼ��_����ͼ(" & curData1 & ")"
                Call zlDatabase.ExecuteProcedure(strSQL, "����Ӱ�񱨸�ͼ���")
    
                If (mlngModule = G_LNG_VIDEOSTATION_MODULE And mSysPar.lngVideoStationMoneyExeModle = 2) Or _
                   (mlngModule = G_LNG_PATHSTATION_MODULE And mSysPar.lngPatholStationMoneyExeModle = 2) Or _
                   (mlngModule = G_LNG_PACSSTATION_MODULE And mSysPar.lngPacsStationMoneyExeModle = 1) Then
                    'ִ�з���
                    Call ExecuteExpense(curData1, GetSendNo(curData1), 4)
                End If
                
                If ucPacsHelper1.AdviceId = CLng(curData1) Then
                     Call ucPacsHelper1.SyncReportImgState(GetReportImgs(curData1))
                End If
                
                '����״̬ͬ����Ϣ
                Call mobjMsgCenter.Send_Msg_StateSync(curData1)
                
                If lngEventNo = BM_REPORT_EVENT_POPUPEXIT Then
                    If CLng(curData1) = mobjCurStudyInfo.lngAdviceId And mstrSelTabName = C_TAB_NAME_��鱨�� Then
                        'ˢ��Ƕ��ʽ��������
                        If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.zlRefreshFace(mobjCurStudyInfo, mstrSelModuleTag, True)
                    End If
                End If

                blnSyncRow = True
            End If
            
        Case BM_REPORT_EVENT_DELETE '����ɾ��֪ͨ*****************************************************************************
            'curData1 ��ʾҽ��ID
            'curData2 ��ʾ����ID ��δʹ��
            'curData3 ��ʾ�༭������
            
            If Val(curData3) = -1 Or Val(strTag) = 1 Then
            
                strSQL = "ZL_Ӱ�񱨸���_Clear(" & curData1 & ")"
                zlDatabase.ExecuteProcedure strSQL, "��ձ�����"
                
                strSQL = "Zl_Ӱ����ͼ��_����ͼ(" & curData1 & ")"
                zlDatabase.ExecuteProcedure strSQL, "��ձ���ͼ"
                
                '��������ͼ״̬ͬ����ʾ
                Call ucPacsHelper1.ClearReportImgState
                
                blnSyncRow = True
            End If
            
        Case BM_REPORT_EVENT_REJECT '���沵��֪ͨ*****************************************************************************
            'curData1 ��ʾҽ��ID
            'curData2 ��ʾ����ID ��δʹ��
            'curData3 ��ʾ�༭������
            
            If Val(curData3) = -1 Or Val(strTag) = 1 Then
                blnSyncRow = Not Menu_Manage_SendAudit(curData1, "")   '�÷����л�ˢ���б���
                
                '����״̬ͬ����Ϣ
                Call mobjMsgCenter.Send_Msg_StateSync(curData1)
            End If
            
        Case BM_REPORT_EVENT_BACK '�������֪ͨ*****************************************************************************
            'curData1 ��ʾҽ��ID
            'curData2 ��ʾ����ID ��δʹ��
            'curData3 ��ʾ�༭������
            
            If Val(curData3) = -1 Or Val(strTag) = 1 Then
                lngSendNo = GetSendNo(curData1, strStudyUID)
                
                Call ResetState(curData1, lngSendNo, strStudyUID)
                blnSyncRow = Menu_Manage_SendAudit(curData1, "")    '�÷����л�ˢ���б���
            End If
            
        Case BM_REPORT_EVENT_OPEN  '������¼�
            'curData1 ��ʾҽ��ID
            'curData2 ��ʾ����ID ��δʹ��
            'curData3 ��ʾ�༭������
            
            If Val(curData3) = -1 Or Val(strTag) = 1 Then
                'ҽ��ģ��򿪱�����Զ���Ƭ
                If mSysPar.blnShowImgAfterReport = True And mlngModule = G_LNG_PACSSTATION_MODULE Then
                    If mobjPacsCore Is Nothing Then
                        HintMsg "ͼ��鿴������Ч�����ܽ��д˲�����", "cbrMain_Execute", vbOKOnly
                        Exit Sub
                    End If
                
                    Call OpenViewer(1, mobjPacsCore, curData1, False, Me)
                End If
            End If
            
        Case BM_REPORT_EVENT_QUALITY    '�����������
            If Val(strTag) = 1 Then
                blnSyncRow = True
            End If
            
        Case BM_IMAGE_EVENT_QUALITYTAG  'Ӱ���������
            If Val(strTag) = 1 Then
                blnSyncRow = True
            End If
            
        Case BM_IMAGE_EVENT_GETIMAGE '��ȡӰ��
            If Val(strTag) = 1 Then
                blnSyncRow = True
            End If
            
        Case BM_IMAGE_EVENT_TECHDO  '��ʦִ��
            If Val(strTag) = 1 Then
                blnSyncRow = True
            End If
            
        Case BM_IMAGE_EVENT_CHANGEDEVICE   '�ı�Ӱ���豸
            If Val(strTag) = 1 Then
                blnSyncRow = True
            End If
            
        Case BM_IMAGE_EVENT_XWFILMPRINT       '��Ƭ�����ӡ
            If Val(strTag) = 1 Then
                blnSyncRow = True
            End If
    
        Case BM_IMAGE_EVENT_DEL         'ɾ��ͼ���֪ͨ
            If Val(strTag) = 1 And Val(curData3) = -1 Then      '�ж��Ƿ�Ϊɾ�����һ��ͼ
                
''                Ӱ��ҽ��frmWork_ImageV2ɾ��ͼ��ʱ����Ҫȷ��ֻ��ɾ�����һ��ͼ��ʱ���ܴ�������Ϣ
'                If mobjCurStudyInfo.lngAdviceId = Val(curData1) Then
'                    If mobjCurStudyInfo.intStep = 3 Then
'                        strSQL = "Zl_Ӱ����_State(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & ",2,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & GetCurDeptId & ")"
'                        zlDatabase.ExecuteProcedure strSQL, "ɾ�����һ��ͼ��"
'
'                        mobjCurStudyInfo.intStep = 2
'                    End If
'                Else
                    lngSendNo = GetSendNo(curData1, strStudyUID, , lngStep)
            
                    '������״̬Ϊ�Ѽ�飬����ɾ������ͼ�����Ҫ��ͼ����л���
                    If lngStep = 3 And Len(strStudyUID) <= 0 Then
                        '����Ӱ����״̬�����ɾ�����һ��ͼ����ԭ������Ϊ3�����޸�Ϊ2
                        strSQL = "Zl_Ӱ����_State(" & Val(curData1) & "," & lngSendNo & ",2,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & GetCurDeptId & ")"
                        zlDatabase.ExecuteProcedure strSQL, "ɾ�����һ��ͼ��"
                    End If
'                End If
                
                '����״̬ͬ����Ϣ
                Call mobjMsgCenter.Send_Msg_StateSync(curData1)
                    
                blnSyncRow = True
            End If
            
            If Val(strTag) = 1 Then Call SyncHelperDataState(curData1, strCurExPro, 0)
            
        Case BM_IMAGE_EVENT_FIRST '�״�ͼ��ɼ�֪ͨ
            If Val(curData2) = -1 Then      '��ʾ�״�ͼ��ɼ�
                Call WriteEprExChangeData(curData1)
                
'                If mobjCurStudyInfo.lngAdviceId = Val(curData1) Then
'                    If mobjCurStudyInfo.intStep < 3 Then
'                        strSQL = "Zl_Ӱ����_State(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & ",3,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & GetCurDeptId & ")"
'                        zlDatabase.ExecuteProcedure strSQL, "����״βɼ�"
'
'                        mobjCurStudyInfo.intStep = 3
'                    End If
'                Else
                    '�����ݿ��ȡָ�����״̬
                    lngSendNo = GetSendNo(curData1, strStudyUID, , lngStep)
                    
                    If lngStep < 3 Then
                        strSQL = "Zl_Ӱ����_State(" & Val(curData1) & "," & lngSendNo & ",3,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & GetCurDeptId & ")"
                        zlDatabase.ExecuteProcedure strSQL, "����״βɼ�"
                    End If
'                End If
                
                '����״̬ͬ����Ϣ
                Call mobjMsgCenter.Send_Msg_StateSync(curData1)
                
                blnSyncRow = True
            ElseIf Val(curData2) = -2 Or Val(curData2) = -3 Then
                lngSendNo = GetSendNo(curData1, strStudyUID, , lngStep)
                
                If lngStep < 3 Then
                    strSQL = "Zl_Ӱ����_State(" & Val(curData1) & "," & lngSendNo & ",3,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & GetCurDeptId & ")"
                    zlDatabase.ExecuteProcedure strSQL, "����״βɼ�"
                    
                    '����״̬ͬ����Ϣ
                    Call mobjMsgCenter.Send_Msg_StateSync(curData1)
                    
                    blnSyncRow = True
                End If
            End If
            
            '�����pacs����༭��
            Call SyncHelperDataState(curData1, Val(strCurExPro), 0)
            
        Case WM_XWREPORT_IMG    '���뱨��ͼ
            Call SyncHelperDataState(curData1, 0, 0)
            
            'ͬ������Ƭ�����ı���ͼ��ӵ�����༭����
            If mSysPar.blnDirectSendRepImg Then Call AddViewImageToReport(curData1, curData2)
            
        Case BM_REPORT_EVENT_ADDIMG
            Call AddViewImageToReport(curData1, curData2)
            
        Case BM_REPORT_EVENT_CLOSEEPR   '�������ڹرպ���Ҫ����Ƕ�봰������
        
            If mstrSelTabName = C_TAB_NAME_��鱨�� And Not mobjWork_Report Is Nothing Then
                If Val(curData1) = mobjWork_Report.StudyInfo.lngAdviceId Then
                    Call mobjWork_Report.zlRefreshFace(mobjWork_Report.StudyInfo, GetWorkModuleTag(C_TAB_NAME_��鱨��), True)
                End If
            End If
            
        Case BM_REPORT_EVENT_REFWCHR    'ˢ�´ʾ�   �������ڵĴʾ��ַ�����󣬻ᷢ�ʹ���Ϣ����ͬ������
            Call ReinitWordChar(curData1)
            
        Case BM_REPORT_EVENT_REFFRAGMENT    'ˢ�´ʾ�Ƭ��  �������ڵĴʾ�Ƭ�α���󣬻ᷢ�ʹ���Ϣ����ͬ������
            Call ReinitWordFragment(curData1)
            
        Case BM_PATHOL_EVENT_BASE + wetPatholQuality    '��������
            blnSyncRow = True
            
        Case BM_PATHOL_EVENT_BASE + wetPatholBatSlices  '������Ƭ
            blnSyncRow = True
            
        Case BM_PATHOL_EVENT_BASE + wetPatholBatSpeExm  '�����ؼ�
            blnSyncRow = True
            
        Case BM_PATHOL_EVENT_BASE + wetMaterialSave '�Ŀ鱣��
            blnSyncRow = True
            
        Case BM_PATHOL_EVENT_BASE + wetSlicesSure   '��Ƭȷ��
            blnSyncRow = True
            
        Case BM_PATHOL_EVENT_BASE + wetSpeExamSure  '�ؼ�ȷ��
            blnSyncRow = True
            
        Case BM_PATHOL_EVENT_BASE + wetPatholRequest, BM_PATHOL_EVENT_BASE + wetSpecimenAccept, BM_PATHOL_EVENT_BASE + wetMaterialSure, BM_PATHOL_EVENT_BASE + wetMaterialSave
            blnSyncRow = True
            
            If Not mobjWork_Pathol Is Nothing Then
                If mobjWork_Pathol.AdviceId = Val(curData1) Then
                    'ˢ�²�������ģ�������
                    Call ForceRefreshPatholModule
                End If
            End If
        
    End Select
    
    If Val(strTag) = 1 Or Val(curData3) = -1 Then
        '���ܲ�����ɺ�ִ�в��
        Call ExecutePlugin(lngEventNo, 1, curData1, curData2, curData3)
    End If
    
    If blnSyncRow Then Call UpdateQueryListData(Nothing, curData1)
'Exit Sub
'errHandle:
'    err.Raise 513, , err.Description
End Sub


Private Function VideoIsAttachReportWindow(Optional ByVal lngVideoRootHwnd As Long = 0)
'�ж���Ƶ�Ƿ�Ƕ��ĵ���ʽ���洰��
    Dim objForm As Object
    Dim lngCurVideoRootHwnd As Long
    
    VideoIsAttachReportWindow = False
    
    If mobjWork_ImageCap Is Nothing Then Exit Function
    
    lngCurVideoRootHwnd = lngVideoRootHwnd
    If lngCurVideoRootHwnd = 0 Then
        lngCurVideoRootHwnd = GetAncestor(mobjWork_ImageCap.VideoHwnd, GA_ROOT)
    End If
    
    For Each objForm In Forms
        If TypeOf objForm Is frmReportV2 Then
            If objForm.IsLinkHelper = False And objForm.hwnd = lngCurVideoRootHwnd Then 'And objForm.AdviceId = mobjCurStudyInfo.lngAdviceId Then
                VideoIsAttachReportWindow = True
                Exit Function
            End If
        End If
    Next
    
End Function

Private Sub ForceRefreshPatholModule()
'ǿ��ˢ�²���ģ��
    Dim i As Long
    
    For i = 0 To UBound(mAryWorkModule)
        If InStr(mAryWorkModule(i).ModuleName, "����") > 0 Then
            If Not mAryWorkModule(i).objModule Is Nothing Then
                Call mAryWorkModule(i).objModule.zlRefreshFace(True)
            End If
        End If
    Next

End Sub

Private Sub ReinitWordChar(ByVal lngSourceHwnd As Long)
'���ó��ôʾ��ַ�
    Dim objForm As Object
        
    For Each objForm In Forms
        If TypeOf objForm Is frmReportV2 Then
            If objForm.hwnd <> lngSourceHwnd Then
                Call objForm.ReinitWordChar
            End If
        End If
    Next
End Sub

Private Sub ReinitWordFragment(ByVal lngSourceHwnd As Long)
'���ó��ôʾ�Ƭ��
    Dim objForm As Object
        
    For Each objForm In Forms
        If TypeOf objForm Is frmReportV2 Then
            If objForm.hwnd <> lngSourceHwnd Then
                Call objForm.ReinitWordFragment
            End If
        End If
    Next
End Sub


Private Sub AddViewImageToReport(ByVal lngAdviceId As String, ByVal strImgFile As String)
'��ӹ�Ƭͼ�񵽱���
    Dim objForm As Object
    Dim strFileName As String
    Dim objFSO As New FileSystemObject
    Dim objPopup As Object
    Dim objEmbed As Object
    
    strFileName = objFSO.GetFileName(strImgFile)
    
    For Each objForm In Forms
        If TypeOf objForm Is frmReportV2 Then
            If objForm.AdviceId = lngAdviceId Then
                If objForm.Caption = "����༭" Then 'Ƕ��ʽ����༭�������setTitle������ʾ ���� �Ա�Ȼ��߻�����Ϣ
                    Set objEmbed = objForm
                Else
                    Set objPopup = objForm
                End If
            End If
        End If
    Next
    
    If Not objPopup Is Nothing Then
        Call objPopup.AddRepImgFile(strImgFile, 0, strFileName)
        Exit Sub
    End If
    
    If Not objEmbed Is Nothing Then
        Call objEmbed.AddRepImgFile(strImgFile, 0, strFileName)
        Exit Sub
    End If
     
End Sub

Private Sub SyncHelperDataState(ByVal lngAdviceId As Long, ByVal lngSourceHwnd As Long, ByVal lngSyncType As Long)
'ͬ��helperģ��������ʾ
    Dim objForm As Object
    
    If lngSourceHwnd < 0 Then Exit Sub
    
    For Each objForm In Forms
        If TypeOf objForm Is frmReportV2 Then
            Call objForm.SyncHelper(lngAdviceId, lngSourceHwnd, lngSyncType)
        End If
    Next
    
    If ucPacsHelper1.hwnd <> lngSourceHwnd And lngAdviceId = mobjCurStudyInfo.lngAdviceId Then
        If lngSyncType = 0 And ucPacsHelper1.SelTabName <> "ͼ��" Then Exit Sub
        If lngSyncType = 1 And ucPacsHelper1.SelTabName <> "�ʾ�" Then Exit Sub
        If lngSyncType = 2 And ucPacsHelper1.SelTabName <> "��ʷ" Then Exit Sub
        If lngSyncType = 3 And ucPacsHelper1.SelTabName <> "����" Then Exit Sub
        
        Call ucPacsHelper1.zlRefresh(mobjCurStudyInfo, 0, True)
    End If
End Sub

Private Sub WriteEprExChangeData(ByVal lngAdviceId As Long)
'д��Ͳ����༭������ͼ�񽻻��ı��
    Dim strStudyUID As String
    Dim strRecDate As String
    Dim strIniContext As String
    Dim strFile As String
    
    Call GetSendNo(lngAdviceId, strStudyUID, strRecDate)
    
    strIniContext = "[DATA]" & vbCrLf & _
                                "STUDYUID=" & strStudyUID & vbCrLf & _
                                "IMGPATH=" & GetTempImgPath & strRecDate & "\" & strStudyUID & "\"
                                
    strFile = GetTempImgPath() & "DataExchange\"
    If DirExists(strFile) = False Then Call MkLocalDir(strFile)
    
    strFile = strFile & lngAdviceId & ".dat"
    
    Call WritTextFile(strFile, strIniContext)
    Call SetFileHide(strFile)
End Sub
'***********************************************IEventNotifyʵ��***********************************************

Private Function HintError(objErr As ErrObject, ByVal strMethodName As String, _
    Optional ByVal blnIsDataErr As Boolean = True) As Long
    If blnIsDataErr Then
        HintError = Notify.PrintErr(objErr, infDataErr, , C_MODULE_NAME, strMethodName)
    Else
        HintError = Notify.PrintErr(objErr, infNormalErr, , C_MODULE_NAME, strMethodName)
    End If
End Function

Private Function HintMsg(ByVal strMsg As String, ByVal strMethodName As String, _
    Optional ByVal lngMsgType As Long = infHint) As Long
        HintMsg = Notify.PrintInfo(strMsg, lngMsgType, , C_MODULE_NAME, strMethodName)
End Function


Private Function GetReportImgs(ByVal lngAdviceId As Long) As String
'��ȡ����ͼ��UID
'���汣������
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
On Error GoTo errhandle
    GetReportImgs = ""
    
    strSQL = "select ͼ��UID from Ӱ����ͼ�� a, Ӱ�������� b, Ӱ�����¼ c where a.����UID =b.����UID and b.���UID = c.���UID and c.ҽ��ID=[1] and nvl(a.����ͼ,0)<>0"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��鱨��ͼ", lngAdviceId)
    If rsData.RecordCount <= 0 Then Exit Function
    
    While Not rsData.EOF
        GetReportImgs = GetReportImgs & ";" & NVL(rsData!ͼ��UID) & ";"
        Call rsData.MoveNext
    Wend
    
Exit Function
errhandle:
    If HintError(err, "GetReportImgs") = 1 Then Resume
End Function

Private Sub ResetState(ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal strStudyUID As String)
'�ù���ֻ���û��ת�������ݽ��д���
'2-�ѱ�����3-�Ѽ�� 4-�ѱ��棬5-����ˣ�6-�����
    Dim curSignInfo As TReportSignInfo
    Dim strSQL As String
    Dim lngState As Long
    Dim strCreateUser As String
    
    curSignInfo = GetLastSignInfo(lngAdviceId, strCreateUser)
    
    If curSignInfo.ID = 0 Then
        '���洦��δǩ��״̬
        '�п����ǻ��˵�δǩ��״̬
        '��Ҫ�ж��Ƿ���ͼ�����û��ͼ�������ѱ���״̬�������ͼ�������Ѽ��״̬
        lngState = 2
        If Len(strStudyUID) > 0 Then lngState = 3
        
        '��û��ǩ��ʱ��д��ʵ�ʵı�����
        strSQL = "ZL_Ӱ�񱨸汣��_Update(" & lngAdviceId & ",'" & IIf(mstrRPTExecutor <> "", mstrRPTExecutor, strCreateUser) & "','')"
        Call zlDatabase.ExecuteProcedure(strSQL, "���±�����Ա")
    Else
        If curSignInfo.ǩ������ > 1 Then
            '���洦�����ǩ��״̬
            lngState = 5
            
            strSQL = "ZL_Ӱ�񱨸汣��_Update(" & lngAdviceId & ",'" & strCreateUser & "','" & curSignInfo.���� & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "���±�����Ա")
        Else
            '���洦�����ǩ��״̬
            lngState = 4
            
            '��ո�����
            strSQL = "ZL_Ӱ�񱨸汣��_Update(" & lngAdviceId & ",'" & strCreateUser & "','')"
            Call zlDatabase.ExecuteProcedure(strSQL, "���±�����Ա")
        End If
    End If
    
    strSQL = "Zl_Ӱ����_State(" & lngAdviceId & "," & lngSendNo & "," & lngState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "���±���״̬")
End Sub

Private Sub ExecuteExpense(ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal lngProcState As Long)
'ִ�з���
    Dim lngID As Long
    Dim strSQL As String
    
    If mblnAllDepts Then
        lngID = UserInfo.����ID
    Else
        lngID = mlngCur����ID
    End If
    
    strSQL = "Zl_Ӱ�����ִ��(" & lngAdviceId & "," & lngSendNo & "," & lngProcState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lngID & ")"
    zlDatabase.ExecuteProcedure strSQL, "ִ�з���"
End Sub

Private Sub ExecutePlugin(ByVal lngEventNo As Long, ByVal lngTimeTag As Long, _
    Optional data1, Optional data2, Optional data3)
'ִ�в��
'lngTimeTag:0-����ִ��ǰ��1-����ִ�к�
    Dim lngTimeType As Long
    
'On Error GoTo errHandle

    Select Case lngEventNo
        Case BM_RIS_EVENT_REGISTER
            lngTimeType = EInterfaceExeTimeV2.���Ǽ�
        
        Case BM_RIS_EVENT_RECEVIE
            lngTimeType = EInterfaceExeTimeV2.��鱨��
            
        Case BM_RIS_EVENT_COMPLETE
            lngTimeType = EInterfaceExeTimeV2.������
            
        Case BM_RIS_EVENT_CANCELREG
            lngTimeType = EInterfaceExeTimeV2.ȡ���Ǽ�
            
        Case BM_RIS_EVENT_CANCELREC
            lngTimeType = EInterfaceExeTimeV2.ȡ������
        
        Case BM_RIS_EVENT_CANCELCOMP
            lngTimeType = EInterfaceExeTimeV2.ȡ�����
            
        Case BM_IMAGE_EVENT_CAPTURE
            lngTimeTag = EInterfaceExeTimeV2.ͼ��ɼ�
            
        Case BM_IMAGE_EVENT_DEL
            lngTimeType = EInterfaceExeTimeV2.ɾ��ͼ��
            
        Case BM_REPORT_EVENT_AUDIT
            lngTimeType = EInterfaceExeTimeV2.�������
            
        Case BM_REPORT_EVENT_SIGN
            lngTimeType = EInterfaceExeTimeV2.����ǩ��
            
        Case BM_REPORT_EVENT_SAVE
            lngTimeType = EInterfaceExeTimeV2.���汣��
            
        Case BM_REPORT_EVENT_REJECT
            lngTimeType = EInterfaceExeTimeV2.���沵��
            
        Case BM_REPORT_EVENT_DELETE
            lngTimeType = EInterfaceExeTimeV2.ɾ������
            
        Case BM_REPORT_EVENT_BACK
            lngTimeType = EInterfaceExeTimeV2.�������
            
        Case BM_SYS__EVENT_MENU
            lngTimeType = EInterfaceExeTimeV2.�˵�ִ��
            
    End Select
    
    Call ExecutePluginInterface(lngTimeType, lngTimeTag, data1, data2, data3)
'Exit Sub
'errHandle:
'    err.Raise -1, , err.Description
End Sub

Private Function TraversalAdvice(ByVal lngAdviceId As Long, ByVal blnIsMoveDown As Boolean, _
    Optional ByRef lngSendNo As Long = 0, Optional ByRef blnIsMoved As Boolean = False) As Long
'����ҽ��
    Dim lngRowIndex As Long
    Dim lngNewRow As Long
    Dim lngResult As Long
    Dim lngIdCol As Long
    Dim objBaseInfo As clsStudyInfo
    
    TraversalAdvice = lngAdviceId
    
    lngIdCol = vsfList.ColIndex("ҽ��ID")
    lngRowIndex = vsfList.FindRow(lngAdviceId, , lngIdCol)
    
    If blnIsMoveDown Then
        lngNewRow = lngRowIndex + 1
        If lngNewRow >= vsfList.Rows Then Exit Function
    Else
        lngNewRow = lngRowIndex - 1
        If lngNewRow < 1 Then Exit Function
    End If
    
    lngResult = Val(vsfList.TextMatrix(lngNewRow, lngIdCol))
    Set objBaseInfo = GetBaseInfo(lngResult, vsfList.Cell(flexcpData, lngNewRow))

    TraversalAdvice = lngResult
    
    lngSendNo = objBaseInfo.lngSendNo
    blnIsMoved = objBaseInfo.blnMoved
End Function

Private Sub DynamicCreateModuleObj()
    Dim strDllName As String
On Error GoTo errhandle
    '���������㲿��
    
    strDllName = "zlOneCardComLib.clsOneCardComLib"
    Set mobjSquareCard = CreateObject(strDllName)
    
    'mobjAppendBill: ���mobjAppendBill��Ϊ�գ���ʾʹ�û��ģʽ
    strDllName = ""
    Set mobjAppendBill = CreateObject("ZlSoft.HIS.Charge.AppendCharge")
Exit Sub
errhandle:
    If strDllName <> "" Then
        HintError err, "DynamicCreateModuleObj<" & strDllName & ">", False
    End If
    
    Set mobjAppendBill = Nothing
End Sub

Private Sub InitBaseComponent()
On Error GoTo errhandle
    '��ʼ����Ƭ����
    If mobjPacsCore Is Nothing Then
        Set mobjPacsCore = New zl9PacsCore.clsViewer
        
        '��Ӱ��ҽ������վ��Ƭʱ������ʾ��汨��ͼ��ť
        If mlngModule <> G_LNG_PACSSTATION_MODULE Then
            mobjPacsCore.ReportImgButtonVisible = False
        End If
    End If
    
    '��ʼ�������㲿��
    If Not mobjSquareCard Is Nothing Then
        mobjSquareCard.zlInitComponents Me, mlngModule, glngSys, gstrDBUser, gcnOracle
    End If
    
    '��仰����ʡ�ԣ����һ�������������⣬ֻҪ��ʽ��ȷ���ɣ������ᱻ�޸�
    PatiIdentify.zlInit Me, glngSys, mlngModule, gcnOracle, gstrDBUser, mobjSquareCard, InitCardType("����;")
Exit Sub
errhandle:
    HintError err, "InitBaseComponent", False
End Sub

Private Sub StartMsgCenter(ByVal lngDeptId As Long)
'������Ϣ����
    If mobjMsgCenter Is Nothing Then
        Set mobjMsgCenter = New clsPacsMsgProcess
    End If
    
    Call mobjMsgCenter.OpenMsgCenter(mlngModule, lngDeptId, mstrPrivs)
End Sub


Private Sub RestoreFormState()
    Dim blnDo As Boolean
    Dim strLayout As String
On Error GoTo errhandle
    '�õ����Ի�������
    blnDo = Val(zlDatabase.GetPara("ʹ�ø��Ի����")) <> 0
    
     '���ע����й��������ֵΪ�� ���� �ѹ�ѡ���Ի����ã���ô��ע���д�빤������ʾģʽֵ
    If mintToolBarWriteReg = 9 Or (mintToolBarWriteReg = 0 And blnDo) Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\CommandBars", "cbrMainButtonText", 3
    End If
    
    '�ָ������״̬   ע���ָ�����״̬ ������� ��ע���д�빤������ʾģʽֵ �������棬�������ɹ�������ʾģʽ����
    Call RestoreWinState(Me, App.ProductName)
    
    strLayout = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\", "HELPER", "")
    Call ucPacsHelper1.SetLayout(strLayout)
    
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
Exit Sub
errhandle:
    HintError err, "RestoreFormState", False
End Sub

Public Sub ShowStation(ByVal lngModule As Long, Owner As Object)
    
'    Dim t1 As Long
    Dim i As Integer
    
    mlngSelHwnd = 0
    mstrSelTabName = ""
    mstrSelModuleTag = ""
    Set mobjSelModule = Nothing
    
    mblnIsValid = True
    mblnInitOk = False
    mlngModule = lngModule
    mintState = 0
    mblnLockState = False
    mblnIsHistoryMode = False
    
    'Ĭ�ϱ�����Ϊ��ǰ��¼�û�
    mstrRPTExecutor = UserInfo.����
    
    Set mrsDeptParas = Nothing  'ʹ���Ҳ����������½��м���
    
    mstrPrivs = gstrPrivs & ";" & GetPrivFunc(100, 9001) & ";"  '�����жϴ�Ⱦ����Σ��ֵ�Ȳ˵�Ȩ��
    
    Call WriteLog("ShowStation -> Step 0�����ô�����ʽ����ʼ������������")
    
    Call StyleChange(sAshen)
    
    If IsExistsBGServer() = "" Then
        'δ��⵽��̨�����������������ʾ
        Call HintMsg("δ��⵽��̨�������ͼ�񽫲������ú�̨���䡣", "ShowStation", infWaring)
    Else
        'TODO:�򿪷�����ʧ�ܴ���...
    End If
    
    Call DynamicCreateModuleObj
    
    Call InitBaseComponent

    Call WriteLog("ShowStation -> Step 1������Ӱ�������ڳ�ʼ�����̡�")

    If Not mblnFormLoadState Then
        If Not InitDepts Then '��ʼ��ҽ������
            Unload Me
            Exit Sub
        End If
        
        If mlngModule <> 1290 Then
            ucPacsHelper1.AllowEmbedVideo = IIf(Val(GetDeptPara(mlngCur����ID, "��ʾ��Ƶ�ɼ�", "0")) <> 0, True, False)
            
            Set mobjCapLinker = New clsCapLinker
            Set mobjCapLinker.MainHelper = ucPacsHelper1
            
            mobjCapLinker.Init Me, mlngCur����ID, mstrPrivs
        Else
            ucPacsHelper1.AllowEmbedVideo = False
            Call ucPacsHelper1.HideEmbedVideo
        End If
        
        ReDim mAryWorkModule(0)
    
        Call StartMsgCenter(mlngCur����ID)  '������Ϣ����
        
        Call InitPars                       '��ʼ������
        
        Call InitQueryWrapComponent
        
        Call initInterface(mlngModule)
        
        Call ReSetFormFontSize              '���ý��������С
        
        Call InitLayout                     '���ý��沼��
        
        Call InitPacsHelper                 '��ʼ��pacshelper����
        Call InitWorkModuleTab              '���ù���ģ��tab��ǩ
        Call InitCommandBars                '����ϵͳ�˵�
        
        Call initTabExtra                   '��ʼ���б�����Ϣ
         
'        Call RestoreFormState               '�ָ�����״̬
        
        mblnFormLoadState = True
    End If
    
    Call WriteLog("ShowStation -> Step 2����ʾϵͳ���档")
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then Set gobjEvent = Me
    
    If Not mobjPacsCore Is Nothing Then mobjPacsCore.DirectSendRepImg = mSysPar.blnDirectSendRepImg
    
    '����ʾ����ǰϵͳ����
    Me.Show , Owner
    
    Call RestoreFormState               '�ָ�����״̬
    
    If Me.WindowState = 1 Then Me.WindowState = 0
    
    Call WriteLog("ShowStation -> Step 3��ˢ�������б�")
    'ˢ�¼������
    
    If mintQueryState = 1 Then
        Call ExecuteDefaultQueryScheme
    End If

    mblnInitOk = True

    Call WriteLog("ShowStation -> Step 4����������ģ����ʾ��")
    If Not TabWindow.Selected Is Nothing Then
        Call TabWindow_SelectedChanged(TabWindow.Selected)
    End If
 
    'δ����ϵͳ�������ܿ�����Ƶ���棬��Ҫ����һ����ƵԤ��
    If Not mobjWork_ImageCap Is Nothing Then
        Call WriteLog("ShowStation -> Step 5��������ƵԤ����")
        Call mobjWork_ImageCap.zlRePreview
    End If
    
    Call WriteLog("ShowStation -> Step End.������Ӱ�������ڳ�ʼ�����̡�")
    
'    Debug.Print "ShowStation��ʱ" & GetTickCount - t1

    '����������ȼ�������Ҫ�����ȼ�����
    If mobjCaptureHot Is Nothing And (mstrCaptureHot <> "" Or mstrCaptureAfterHot <> "" Or mstrCaptureAfterTagHot <> "") Then
        Set mobjCaptureHot = New zl9PacsControl.clsHookKey
        Call mobjCaptureHot.EnableHook(WM_KEYDOWN, True)
    End If
    
    
    '������Ҫ����һ������ �����޸Ĺ��˲˵����� �����ɵ���Ϊֻ�Բ˵��ؼ��޸ġ�
    '��ʼ���������� �ӵ�����Ϊ�˷�ֹ��һЩ���������ʱ�򣬻ᵼ������ָ��ɳ�ʼ��
    Call SetFontSize(IIf(gbytFontSize = 12, 1, IIf(gbytFontSize = 15, 2, 0)))
    
    Call StartImageValid
End Sub

Private Sub StartImageValid()
On Error GoTo errhandle
    If mSysPar.lngImageValid > 0 Then
        If Len(Dir(GetAppRootPath & "zlPacsImageValid.exe")) > 0 Then
            If InitRegister Then
                Shell GetAppRootPath & "zlPacsImageValid.exe   " & gstrServerName & "||" & gstrUserName & "||" & gstrUserPswd & "||" & mlngCur����ID & "||" & mSysPar.lngImageValid & "||" & "" & "||2", 1
            End If
        End If
    End If
Exit Sub
errhandle:
    HintError err, "StartImageValid<����ͼ����֤>", False
End Sub



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
    
    mblnInitOk = False
    
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
    
    mblnInitOk = True
    
    Exit Sub
errhandle:
    mblnInitOk = True
    If HintError(err, "Menu_File_Excel_click") = 1 Then Resume
End Sub

Private Sub Menu_RichEPR(ByVal cbrID As Long)
'�Զ��򿪱���༭����ͬʱ������PACS����༭���͵��Ӳ����༭��
On Error GoTo errhandle
    Dim cbrControl As CommandBarControl, i As Long
    Dim strCurModuleTag As String
    
    '���û��ѡ�������ݣ���ֱ���˳�ִ��
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_RichEPR", vbInformation
        Exit Sub
    End If
    
    '����ҳ�治�ɼ�ʱ��ִ���κβ���
    If TabWindow.Selected.Caption <> C_TAB_NAME_��鱨�� Then
        For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
            If TabWindow(i).Caption = C_TAB_NAME_��鱨�� And TabWindow(i).Visible = True Then
                TabWindow(i).Selected = True
                Exit For
            End If
        Next
        
        If TabWindow.Selected.Caption <> C_TAB_NAME_��鱨�� Then Exit Sub
    Else
        If TabWindow.Selected.Visible = False Then Exit Sub
    End If

    strCurModuleTag = GetWorkModuleName(mstrSelTabName, mobjCurStudyInfo.lngExeDepartmentId, mobjCurStudyInfo.lngPatientFrom)
    If strCurModuleTag <> mstrSelModuleTag Then
       Call SelectModule(mstrSelTabName, strCurModuleTag)
       TabWindow.Selected.tag = strCurModuleTag
    End If
    
    '�ҵ�����ҳ�棬�ٴ��������ҳ��
    'ˢ��Ƕ��ҳ������
    If Not mobjWork_Report Is Nothing Then
        Call mobjWork_Report.zlRefreshFace(mobjCurStudyInfo, strCurModuleTag, True)
        
        If strCurModuleTag = C_WORKMODULE_NAME_�ϰ汨�� Then
            Call mobjWork_Report.zlMenu.zlExecuteMenu(strCurModuleTag, conMenu_PacsReport_Open + mobjWork_Report.BaseMenuId)
        Else
            If cbrID = conMenu_PacsReport_Write Then
                Call mobjWork_Report.zlMenu.zlExecuteMenu(strCurModuleTag, conMenu_Edit_Modify + mobjWork_Report.BaseMenuId)
            Else
                Call mobjWork_Report.zlMenu.zlExecuteMenu(strCurModuleTag, cbrID + mobjWork_Report.BaseMenuId)
            End If
        End If
    End If
    
Exit Sub
errhandle:
    Call HintError(err, "Menu_RichEPR", False)
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
            
            '�ж��Ƿ���Ҫ�����ȼ�
            If mobjCaptureHot Is Nothing And (mstrCaptureHot <> "" Or mstrCaptureAfterHot <> "" Or mstrCaptureAfterTagHot <> "") Then
                Set mobjCaptureHot = New zl9PacsControl.clsHookKey
                Call mobjCaptureHot.EnableHook(WM_KEYDOWN, True)
            End If
            
        End If
    End With
Exit Sub
errhandle:
    Call HintError(err, "Menu_File_Parmeter_click", False)
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
    
    Call HintError(err, "Menu_File_ShortcutSet_click", False)
End Sub


Private Sub Menu_Help_About_click()
On Error GoTo errhandle
    ShowAbout Me, App.title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
errhandle:
    Call HintError(err, "Menu_Help_About_click", False)
End Sub

Private Sub Menu_Help_Help_click()
'���ܣ����ð�������
On Error GoTo errhandle
    ShowHelp App.ProductName, Me.hwnd, "ZL9PACSWORK"
Exit Sub
errhandle:
    Call HintError(err, "Menu_Help_Help_click", False)
End Sub

Private Sub Menu_Help_Web_Forum_click()
On Error GoTo errhandle
    Call zlWebForum(Me.hwnd)
Exit Sub
errhandle:
    Call HintError(err, "Menu_Help_Web_Forum_click", False)
End Sub


Private Sub Menu_Help_Web_Mail_click()
On Error GoTo errhandle
    zlMailTo hwnd
Exit Sub
errhandle:
    Call HintError(err, "Menu_Help_Web_Mail_click", False)
End Sub

Private Sub Menu_Manage_ȡ������()
'ȡ��������������ǣ�ÿ��ȡ��������ͼ��ȫ���������б���ɢ��N����ʱ��¼
On Error GoTo errhandle
    Dim lngResult As Long
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_ȡ������", vbInformation
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
    
    Call ReleationImage(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, 1, True)

Exit Sub
errhandle:
    Call HintError(err, "Menu_Manage_ȡ������", False)
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
        If Val(NVL(rsData!���ӱ�־)) = 3 Then
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
    If HintError(err, "Menu_Manage_������") = 1 Then Resume
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
     
    If gobjRegister Is Nothing Then Set gobjRegister = CreateObject("zlRegister.clsRegister")
    
    lngUerResId = UserInfo.ID
    strNodeName = ""
    strNodeNo = ""
    
    '��ѯ������Դϵͳ
    strSysFrom = "01"
    strSQL = "Select ���ӱ�־ From ���˹Һż�¼ Where ����ID=[1] and No=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ӱ�־", mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.strRegNo)
    If rsData.RecordCount > 0 Then
        If Val(NVL(rsData!���ӱ�־)) = 3 Then strSysFrom = "02"
    End If
    
            
    strUserName = gobjRegister.GetUserName
    strUserPswd = gstrInputPwd ' GetLoginPassword 'gobjRegister.GetPassword(App.hInstance)
    
    If strSysFrom = "02" Then
        strSQL = "Select ��ԴID From ��Ա�� Where ID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��Ա����ԴID", UserInfo.ID)
        If rsData.RecordCount > 0 Then
            strUerResId = NVL(rsData!��ԴID)
        End If
    
        strSQL = "Select a.����ID," & _
                    " '' As �����ʶ, " & _
                    " Decode(a.������Դ, 4, 2, 2, 1, 0) As ������Դ, " & _
                    "Nvl(a.���ID, a.ID) As ҽ�����, b.���ͺ�, " & _
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
            """������Դ"":""" & NVL(rsData!������Դ) & """," & _
            """���˱�ʶ"":""" & NVL(rsData!����ID) & """," & _
            IIf(strSysFrom <> "02", """�����ʶ"":""" & NVL(rsData!�����ʶ) & """,", "") & _
            """ҽ�����"":""" & NVL(rsData!ҽ�����) & """," & _
            """ҽ�����ͺ�"":""" & NVL(rsData!���ͺ�) & """," & _
            """��ǰ���ұ�ʶ"":""" & NVL(rsData!��ǰ���ұ�ʶ) & """," & _
            """��ǰ���ұ���"":""" & NVL(rsData!��ǰ���ұ���) & """," & _
            """��ǰ��������"":""" & NVL(rsData!��ǰ��������) & """," & _
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
    If HintError(err, "GetJsonPar") = 1 Then Resume
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
    
    getRegID = NVL(rsTemp!ID, 0)
    
    Exit Function

errhandle:
    If HintError(err, "getRegID") = 1 Then Resume
End Function

Private Function IsAlreadyInputQuality(ByVal lngAdviceId As Long) As Boolean
On Error GoTo errH
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    IsAlreadyInputQuality = False
    
    strSQL = "select �ۺ����� from ��������Ϣ where ҽ��ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, GetWindowCaption, lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    If NVL(rsData!�ۺ�����) <> "" Then IsAlreadyInputQuality = True
    Exit Function
errH:
    If HintError(err, "IsAlreadyInputQuality") = 1 Then Resume
End Function

Private Function Menu_Manage_����������(Optional lngAdviceId As Long = 0, Optional blnRefresh As Boolean = True, Optional strReportId As String = "") As Boolean
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
    
    Menu_Manage_���������� = False
    
    '���ִ�й���=6 ˵���������Ѿ��������״̬����ʱ�˳������̲��Ҳ���Ҫ��ʾ��������XX���Զ���ɲ�����
    If lngAdviceId > 0 Then
        strSQL = "select ҽ��ID from ����ҽ������ where ҽ��ID=[1] and ִ�й���=6"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�Ƿ��Ѿ��������״̬", lngAdviceId)
        If rsData.RecordCount > 0 Then
            Menu_Manage_���������� = True
            Exit Function
        End If
    End If
    
    If InStr(mstrPrivs, ";������;") <= 0 Then
        HintMsg "û��Ȩ�ޣ������������", "Menu_Manage_����������", vbInformation
        Exit Function
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
        HintMsg "��ȡ�������ʧ��", "Menu_Manage_����������", vbInformation
        Exit Function
    End If
        
    Set objStudyInfo = mobjPacsQueryWrap.GetBaseInfo(lngAdviceIDSub, GetMovedState(lngRow, vsfList))
    
    If objStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_����������", vbInformation
        Exit Function
    End If
    
    If Not mSysPar.blnNoSignFinish Then
    '�����ѡ����δǩ������򲻱ؽ���������ж�
        If Is_ExistReportWriting(lngAdviceIDSub) Then
            HintMsg "�����Ѿ��޸Ļ�δǩ��������������ɡ�", "Menu_Manage_����������", vbInformation
            Exit Function
        ElseIf objStudyInfo.intStep < 4 Then
            HintMsg "���滹δǩ��������������ɡ�", "Menu_Manage_����������", vbInformation
            Exit Function
        End If
    End If
    
    '������֮ǰ�����ж��Ƿ�����������������������ɣ�
        '1��סԺ���ߣ��Ѿ���Ժ������δ��˵Ļ��۵���ʹ�á�ִ�к��Զ���˻��۵�������
        '2�����ﻼ�ߣ���δ���ѵĵ��ݡ�
    If objStudyInfo.lngPatientFrom = 2 Then
        'סԺ���ߣ��ж��Ƿ��Ѿ���Ժ������δ��˻��۵�
        If bln����δ��˳�Ժ(objStudyInfo.lngPatId, objStudyInfo.lngPageID, NVL(objStudyInfo.lngAdviceId), objStudyInfo.lngPatientFrom) Then
            'ִ�к��Զ���˻��۵���Ч�����Ҳ����ѳ�Ժ������δ��˵Ļ��۵�
            HintMsg "�ò����ѳ�Ժ������δ��˵Ļ��۵���������ɣ�", "Menu_Manage_����������", vbExclamation
            Exit Function
        End If
    ElseIf objStudyInfo.lngPatientFrom = 4 And mSysPar.blnPEISNoCheckMoneyFinish Then
        '�����ɲ��жϷ��� 133458
    Else
        '������ﻼ��,�ж��Ƿ���δ�ɷ���
        If blnδ�ɷ���(objStudyInfo.lngAdviceId) = True Then
            If objStudyInfo.intGreenChannel = 1 Or objStudyInfo.intEmergentTag = 1 Then
                If HintMsg("�û��߻���δ�ɷѵ���Ŀ���Ƿ�Ҫ��ɣ�", "Menu_Manage_����������", vbYesNo) = vbNo Then
                    Exit Function
                End If
            Else
                HintMsg "�û��߻���δ�ɷѵ���Ŀ��������ɡ�", "Menu_Manage_����������", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    
    Call Notify.Broadcast(BM_RIS_EVENT_COMPLETE, 0, mobjCurStudyInfo.lngAdviceId)

    '����ǲ���ϵͳ��������ʱ������Ҫ�����������ƴ���
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        If Not IsAlreadyInputQuality(objStudyInfo.lngAdviceId) Then
            If Not mobjWork_Pathol Is Nothing Then
                Call mobjWork_Pathol.zlMenu.zlExecuteMenu("", conMenu_Pathol_Quality_Manage)
            End If
        End If
        
        If Not IsAlreadyInputQuality(objStudyInfo.lngAdviceId) Then
            HintMsg "δ¼��������������ִ����ɲ�����", "Menu_Manage_����������", vbInformation
            Exit Function
        End If
    End If
    
    lngID = GetCurDeptId
    
    '��մ�������
    strSQL = "Zl_Ӱ�����¼_�����������(" & lngAdviceIDSub & ",'')"
    zlDatabase.ExecuteProcedure strSQL, "�����������"
    
    
    If objStudyInfo.lngReportType = 1 Then  'pacs����༭��
        strSQL = "ZL_Ӱ����_STATE(" & objStudyInfo.lngAdviceId & "," & objStudyInfo.lngSendNo & ",6,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lngID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ",1)"
    Else
        strSQL = "ZL_Ӱ����_STATE(" & objStudyInfo.lngAdviceId & "," & objStudyInfo.lngSendNo & ",6,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lngID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ",2)"
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, "�ı������")

        
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        gstrSQL = "Zl_������_���(" & objStudyInfo.lngAdviceId & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���������")
    End If
        
    'ȡ���Ŷ���Ϣ
    If mSysPar.blnUseQueue = True And Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Call mobjQueue.zlCompletePacsQueue(objStudyInfo.lngAdviceId)
    End If
    
    
    Menu_Manage_���������� = True
    
    Call UpdateQueryListData(Nothing, objStudyInfo.lngAdviceId)
    
    '������ɵ�ҽ���뵱ǰ�б�ѡ���ҽ����ͬʱ����ͬ��ˢ��ģ������
    If lngAdviceId <> 0 And lngAdviceId = mobjCurStudyInfo.lngAdviceId Then
        Call RefreshModuleData(mstrSelTabName, mstrSelModuleTag, mobjSelModule)
    End If
    
    '���ͼ�������Ϣ
    Call mobjMsgCenter.Send_Msg_StudyComplete(objStudyInfo.lngAdviceId, strReportId)
    
    Call Notify.Broadcast(BM_RIS_EVENT_COMPLETE, 1, mobjCurStudyInfo.lngAdviceId)
    
    
Exit Function
errhandle:
    If HintError(err, "Menu_Manage_����������") = 1 Then Resume
End Function

Private Sub Menu_Manage_ȡ��������()
On Error GoTo errhandle
    Dim strSQL As String
    Dim intState As Integer

    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_ȡ��������", vbInformation
        Exit Sub
    End If

    If mobjCurStudyInfo.blnMoved Then
        HintMsg "�ò��˵ı��μ�������Ѿ�ת���������ݿ⣬�����������", "Menu_Manage_ȡ��������", vbInformation
        Exit Sub
    End If
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        If CheckIsArchived(mobjCurStudyInfo.lngAdviceId) Then
            HintMsg "�ò��˵ĵ����Ѿ��鵵�������������", "Menu_Manage_ȡ��������", vbInformation
            Exit Sub
        End If
    End If
    
    Call Notify.Broadcast(BM_RIS_EVENT_CANCELCOMP, 0, mobjCurStudyInfo.lngAdviceId)
    
    If mobjCurStudyInfo.lngReportType = 1 Then  '1-pacs����༭����2-�����༭��
        intState = getStudyState(mobjCurStudyInfo.lngAdviceId, True)
        strSQL = "ZL_Ӱ����_STATE(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & "," & intState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ",1)"
    Else
        intState = getStudyState(mobjCurStudyInfo.lngAdviceId, True)
        strSQL = "ZL_Ӱ����_STATE(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & "," & intState & ",NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCur����ID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ",2)"
    End If
    
    zlDatabase.ExecuteProcedure strSQL, "ȡ��������"
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        strSQL = "Zl_������_ȡ�����(" & mobjCurStudyInfo.lngAdviceId & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "������ȡ�����")
    End If
    
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    Call RefreshModuleData(mstrSelModuleTag, mstrSelModuleTag, mobjSelModule)
    
    '���ͼ�鳷�������Ϣ
    Call mobjMsgCenter.Send_Msg_CancelComplete(mobjCurStudyInfo.lngAdviceId)
    
    Call Notify.Broadcast(BM_RIS_EVENT_CANCELCOMP, 1, mobjCurStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_ȡ��������") = 1 Then Resume
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
    
    CheckIsArchived = IIf(NVL(rsTemp!״̬, 0) = 1, True, False)
Exit Function
errhandle:
    If HintError(err, "CheckIsArchived") = 1 Then Resume
End Function

Private Sub Menu_Manage_CriticalMark(ByVal lngID As Long)
'Σ��ֵ����
On Error GoTo errhandle
    Dim strSQL As String
    Dim intCritical As Integer
    Dim rsData As ADODB.Recordset
    Dim lngCriticalId As Long
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_CriticalMark", vbInformation
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

    '�����ǰģ���Ǳ���ģ�飬��ͬ�����±���״̬����ʾ
    If mobjWork_Report Is Nothing Then Exit Sub
    If TypeOf mobjSelModule Is frmReportV2 Then Call mobjSelModule.ReadRepStateTag
    
    Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_CriticalMark") = 1 Then Resume
End Sub

Private Sub Menu_Manage_�������(ByVal lngID As Long)
On Error GoTo errhandle
    Dim strSQL As String
    Dim iResult As Integer
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_�������", vbInformation
        Exit Sub
    End If
    
    Select Case lngID
        Case conMenu_Manage_Negative
            iResult = 1
        Case conMenu_Manage_Positive
            iResult = 0
    End Select
    
    strSQL = "ZL_Ӱ����_���(" & mobjCurStudyInfo.lngAdviceId & "," & iResult & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "���������")

    mobjCurStudyInfo.intPositive = iResult
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    '�����ǰģ���Ǳ���ģ�飬��ͬ�����±���״̬����ʾ
    If mobjWork_Report Is Nothing Then Exit Sub
    If TypeOf mobjSelModule Is frmReportV2 Then Call mobjSelModule.ReadRepStateTag
    
    
Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_�������") = 1 Then Resume
End Sub

Private Sub Menu_Manage_��ɫͨ��(ByVal lngID As Long)
On Error GoTo errhandle
    Dim strSQL As String
    Dim intResult As Integer
    Dim blnCanPrint As Boolean
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_��ɫͨ��", vbInformation
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

Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_��ɫͨ��") = 1 Then Resume
End Sub

Private Sub Menu_Manage_�������(ByVal lngID As Long)
On Error GoTo errhandle
    Dim strResult As String
    Dim strSQL As String
    Dim lngColIndex As Long

    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_�������", vbInformation
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
    If HintError(err, "Menu_Manage_�������") = 1 Then Resume
End Sub

Private Sub Menu_Manage_CheckList()
On Error GoTo errhandle
    Dim objCisKernel As clsCISKernel
    
    If mobjCurStudyInfo.lngAdviceId > 0 Then
        Set objCisKernel = DynamicCreate("zlCISKernel.clsCISKernel", "CISKernel")
        
        If objCisKernel Is Nothing Then Exit Sub
        
        Call objCisKernel.ShowPacsApplication(Me, mobjCurStudyInfo.lngAdviceId)
        
        Set objCisKernel = Nothing
    Else
        HintMsg "û��ѡ���ˡ�", "Menu_Manage_CheckList", vbInformation + vbOKOnly
    End If
Exit Sub
errhandle:
    Call HintError(err, "Menu_Manage_CheckList", False)
End Sub

'�ֲ�λִ��
Private Sub menu_Manage_ExecOnePart()
    Dim frmExecForm As frmExecOnePart
    
    Set frmExecForm = New frmExecOnePart
    
    '��ʾ�ֲ�λִ�к�ȡ������
    Call frmExecForm.ZlShowMe(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.strPatientName, mobjCurStudyInfo.strPatientAge, mobjCurStudyInfo.strPatientSex, mobjCurStudyInfo.strStuStateDesc, Me)
    
    'ˢ�·���ҳ��
    If TabWindow.Selected.tag = "�������" Or TabWindow.Selected.tag = "סԺҽ��" Or TabWindow.Selected.tag = "����ҽ��" Then
        Call RefreshModuleData(mstrSelTabName, mstrSelModuleTag, mobjSelModule)
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
    
 
    strSQL = "Select  b.�����ı� As ���� From ���Ӳ������� a,���Ӳ������� b, ����ҽ������ c " & _
             "Where c.ҽ��id = [1] And a.�����ı� = '������' And a.�������� = 3 And a.Id = b.��ID " & _
             "And a.�ļ�id = c.����id And b.�������� = 2 And b.��ֹ�� = 0"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������", mobjCurStudyInfo.lngAdviceId)
    
    If rsData.RecordCount > 0 Then strReportResult = NVL(rsData!����)
    
    If mobjCurStudyInfo.lngPatientFrom = 1 Then        '����
        Call mobjPublicAdvice.ShowDisRegist(Me, 0, , mobjCurStudyInfo.lngPatId, , mobjCurStudyInfo.strRegNo, mobjCurStudyInfo.lngAdviceId, mlngCur����ID, , , , , strReportResult)
    ElseIf mobjCurStudyInfo.lngPatientFrom = 2 Then    'סԺ
        Call mobjPublicAdvice.ShowDisRegist(Me, 0, , mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.lngPageID, , mobjCurStudyInfo.lngAdviceId, mlngCur����ID, , , , , strReportResult)
    Else                                            '���������
        Call mobjPublicAdvice.ShowDisRegist(Me, 0, , mobjCurStudyInfo.lngPatId, , , mobjCurStudyInfo.lngAdviceId, mlngCur����ID, , , , , strReportResult)
    End If
    
    Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_DiseaseRegist") = 1 Then Resume
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
    Call HintError(err, "Menu_Manage_DiseaseQuery", False)
End Sub

Private Sub Menu_Manage_�޸�()
On Error GoTo errhandle
    Dim strOldName As String
    Dim strOldRoom As String
    Dim strQueueName As String
    Dim strCodeNo As String
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_�޸�", vbInformation
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
    Call HintError(err, "Menu_Manage_�޸�", False)
End Sub

Private Sub Menu_Manage_ModifBaseInfo()
'������Ϣ����
On Error GoTo errhandle
    Dim zlPubPatient As Object
    
    Dim int���� As Integer
    Dim str����ID As String

    Set zlPubPatient = CreateObject("zlPublicPatient.clsPublicPatient")
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
    
    Call HintError(err, "Menu_Manage_ModifBaseInfo", False)
End Sub

Private Sub Menu_Manage_���ƵǼ�()
    Dim strQueueName As String
    Dim strCodeNo As String
    Dim lngNewAdviceId As Long
    Dim lngResultState As Long
    
On Error GoTo errhandle
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_���ƵǼ�", vbInformation
        Exit Sub
    End If
    
    Call Notify.Broadcast(BM_RIS_EVENT_REGISTER, 0)
    
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
            
            lngResultState = .mlngResultState
            
            If .mlngResultState <> 0 Then '�ɹ�����
                lngNewAdviceId = .mlngAdviceId
                
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
                
                '���ͬʱ��ѡ����ʼ����Զ��򿪱��桱�͡��ǼǺ��Զ�������������ô���Զ��򿪱������
                If mSysPar.blnAutoOpenReport And mSysPar.blnֱ�Ӽ�� Then Call Menu_RichEPR(conMenu_PacsReport_Write)
                
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
            lngResultState = 0
            
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            If .CopyCheck(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo) = True Then  'ˢ�²���
                .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            End If
            
            If .mblnOk Then '�ɹ�����
                lngResultState = 1
                lngNewAdviceId = .mlngAdviceId
                
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
            End If
        End With
    End If
    
    Call Notify.Broadcast(BM_RIS_EVENT_REGISTER, 1, lngNewAdviceId, lngResultState)
Exit Sub
errhandle:
    Call HintError(err, "Menu_Manage_���ƵǼ�", False)
End Sub

Private Sub Menu_Manage_�Ǽ�()
On Error GoTo errhandle
    Dim strQueueName As String
    Dim strCodeNo As String
    Dim lngNewAdviceId As Long
    Dim lngResultState As Long
    
    Call Notify.Broadcast(BM_RIS_EVENT_REGISTER, 0)
    
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
            
            lngResultState = .mlngResultState
            
            If .mlngResultState <> 0 Then '�ɹ�����
                lngNewAdviceId = .mlngAdviceId
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
                
                '���ͬʱ��ѡ����ʼ����Զ��򿪱��桱�͡��ǼǺ��Զ�������������ô���Զ��򿪱������
                If mSysPar.blnAutoOpenReport And mSysPar.blnֱ�Ӽ�� Then Call Menu_RichEPR(conMenu_PacsReport_Write)
                
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
            .mintImgCount = 0
            .mblnOk = False
            lngResultState = 0
            
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            
            If .mblnOk Then '�ɹ�����
                lngResultState = 1
                lngNewAdviceId = .mlngAdviceId
            
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
                
                '���ͬʱ��ѡ����ʼ����Զ��򿪱��桱�͡��ǼǺ��Զ�������������ô���Զ��򿪱������
                If mSysPar.blnAutoOpenReport And mSysPar.blnֱ�Ӽ�� Then Call Menu_RichEPR(conMenu_PacsReport_Write)
                
                '������������Ϣ
                Call mobjMsgCenter.Send_Msg_Request(.mlngAdviceId)
            End If
        End With
    End If
    
    Call Notify.Broadcast(BM_RIS_EVENT_REGISTER, 1, lngNewAdviceId, lngResultState)
Exit Sub
errhandle:
    Call HintError(err, "Menu_Manage_�Ǽ�", False)
End Sub

Private Sub Menu_Manage_ȡ���Ǽ�()
On Error GoTo errhandle
    Dim strSQL As String
    Dim lngCancelAdviceId As Long
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_ȡ���Ǽ�", vbInformation
        Exit Sub
    End If
    
    If HintMsg("ȷ��Ҫȡ����ǰ������" & Chr(10) & Chr(13) & "����ȡ�������Ӧ��ҽ�����ܾ�ִ�У�", "Menu_Manage_ȡ���Ǽ�", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

    lngCancelAdviceId = mobjCurStudyInfo.lngAdviceId
    Call Notify.Broadcast(BM_RIS_EVENT_CANCELREG, 0, lngCancelAdviceId)
    
    strSQL = "ZL_����ҽ��ִ��_�ܾ�ִ��(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & ",null,null," & mlngCur����ID & ")"
    
    Call zlDatabase.ExecuteProcedure(strSQL, "�����Ǽ�")
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    '����ҽ��������Ϣ
    Call mobjMsgCenter.Send_Msg_CancelAdvice(mobjCurStudyInfo.lngAdviceId)
    
    Call Notify.Broadcast(BM_RIS_EVENT_CANCELREG, 1, lngCancelAdviceId)
Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_ȡ���Ǽ�") = 1 Then Resume
End Sub

Private Sub Menu_Manage_�ٻ�ȡ��()
'���ܣ��ٻر�ȡ���ĵǼ�
On Error GoTo errhandle
    Dim strSQL As String
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_�ٻ�ȡ��", vbInformation
        Exit Sub
    End If
    
    If HintMsg("ȷʵҪ�ٻر�ȡ���Ǽǵ���Ŀ��", "Menu_Manage_�ٻ�ȡ��", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    strSQL = "ZL_����ҽ��ִ��_ȡ���ܾ�(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & ",null,null," & mlngCur����ID & ")"
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    '����״̬ͬ����Ϣ
    Call mobjMsgCenter.Send_Msg_StateSync(mobjCurStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_�ٻ�ȡ��") = 1 Then Resume
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
    Dim lngResultState As Long
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_����", vbInformation
        Exit Sub
    End If
    
    Call Notify.Broadcast(BM_RIS_EVENT_RECEVIE, 0, mobjCurStudyInfo.lngAdviceId)
    
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
            If Format(NVL(rsTemp!ԤԼ��ʼʱ��), "yyyy-mm-dd") <> Format(zlDatabase.Currentdate, "yyyy-mm-dd") Then
                If HintMsg("��ǰ����ԤԼ�ļ������Ϊ " & Format(NVL(rsTemp!ԤԼ��ʼʱ��), "yyyy-mm-dd") & "���뵱ǰʱ�䲻һ�£��Ƿ����������", "Menu_Manage_����", vbInformation + vbYesNo) = vbNo Then
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
            
            lngResultState = .mlngResultState
            
            If .mlngResultState <> 0 Then  '�ɹ�����
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
                
                If .mblnIsRelationImage = True Then
                    '�������ǰ����ͼ��������Զ��������������ｫ��Ӱ��ͼ��ģ�����ˢ��
                    If Not mobjWork_PacsImg Is Nothing Then
                        Call mobjWork_PacsImg.zlRefreshFace(mobjCurStudyInfo, True)
                    End If
                End If
                
                If mSysPar.blnAutoOpenReport Then Call Menu_RichEPR(conMenu_PacsReport_Write)              '��ʼ����Զ��򿪱���
                
                If .mlngResultState = 2 Then
                    '��������Ŷӽкţ����ұ������Զ��Ŷӣ��򱨵�����Ҫ�����ŶӽкŶ���......
                    If mSysPar.blnUseQueue And mSysPar.blnAutoInQueue And Not mobjQueue Is Nothing Then
                        If blnIsCurDayReservations Then
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
            lngResultState = 0
            
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            If .RefreshPatiInfor(True) = True Then  'ˢ�²���
                .mblnOk = False
                .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            End If
            
            If .mblnOk Then  '�ɹ�����
                lngResultState = 1
                
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
                If mSysPar.blnAutoOpenReport Then Call Menu_RichEPR(conMenu_PacsReport_Write)              '��ʼ����Զ��򿪱���
                
                '����״̬ͬ����Ϣ
                Call mobjMsgCenter.Send_Msg_StateSync(.mlngAdviceId)
            End If
            
        End With
    End If
    
    Call Notify.Broadcast(BM_RIS_EVENT_RECEVIE, 1, mobjCurStudyInfo.lngAdviceId, lngResultState)
    
Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_����") = 1 Then Resume
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
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_ȡ������", vbInformation
        Exit Sub
    End If
    
  
    If mobjCurStudyInfo.intStep <= 1 Then Call Menu_Manage_ȡ���Ǽ�: Exit Sub  '����������
    '------------------------------------��ǩ������Ҫ�Ȼ���ǩ�����ٳ���
    
    Call Notify.Broadcast(BM_RIS_EVENT_CANCELREC, 0, mobjCurStudyInfo.lngAdviceId)
    
    strSQL = "Select Distinct B.���ʱ��, B.ǩ������ From ����ҽ������ A, ���Ӳ�����¼ B Where A.����ID=B.Id And A.ҽ��ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Ƿ�ǩ��", mobjCurStudyInfo.lngAdviceId)
    
    If Not rsTemp.EOF Then
        If NVL(rsTemp!���ʱ��, "") <> "" And Val(NVL(rsTemp!ǩ������)) > 0 Then 'ǩ������
            HintMsg "��ǰ���˵ļ�鱨���Ѿ�ǩ��,����ȡ�����,���Ȼ���ǩ��!", "Menu_Manage_ȡ������", vbInformation
            Exit Sub
        End If
    End If
    
    '��������ȡ�Ļ�����Ƭ�����ܽ���ȡ��
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        strSQL = "select count(1) as ���� from ��������Ϣ a, ����ȡ����Ϣ b where a.����ҽ��ID=b.����ҽ��ID and a.ҽ��ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, GetWindowCaption, mobjCurStudyInfo.lngAdviceId)
        If rsTemp.RecordCount > 0 Then
            If Val(NVL(rsTemp!����)) > 0 Then
                HintMsg "�ü����ִ��ȡ�Ĳ��������ܽ���ȡ����", "Menu_Manage_ȡ������", vbInformation
                Exit Sub
            End If
        End If
    End If

    If mobjCurStudyInfo.strStudyUID <> "" And Not CheckPopedom(mstrPrivs, "���ͼ��") Then
        HintMsg "��û��������ͼ��Ȩ��,�������ͼ��,���Բ���ȡ��������!", "Menu_Manage_ȡ������", vbInformation
        Exit Sub
    End If
    
    strMsg = "������Ϣ��������" & mobjCurStudyInfo.strPatientName & "   �Ա�" & mobjCurStudyInfo.strPatientSex & "   ���䣺" & mobjCurStudyInfo.strPatientAge & "   ���ţ�" & mobjCurStudyInfo.strStudyNum & "��" & vbCrLf & _
             "ȡ�����˱��μ�齫ɾ����Ӧ�ļ��ͼ��ͼ�鱨�棬�Ƿ������"

    If HintMsg(strMsg, "Menu_Manage_ȡ������", vbDefaultButton2 + vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
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
    
    '���ͼ��������PACS����ɾ��Ӱ���ļ���Ŀ¼
    If mobjCurStudyInfo.intImageLocation = 0 Then
        RemoveCheckImages mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo
    End If
    
    If TabWindow.Selected.tag = "Ӱ��ɼ�" Then
        'TODO:����������Զ�����������Ҫˢ��helper������ͼ��
'        If Not mobjWork_ImageCap Is Nothing Then
'            Call mobjWork_ImageCap.zlRefreshData(True)
'        End If
    End If
    
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    '����״̬������Ϣ
    Call mobjMsgCenter.Send_Msg_StateCancel(mobjCurStudyInfo.lngAdviceId)
    
    Call Notify.Broadcast(BM_RIS_EVENT_CANCELREC, 1, mobjCurStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_ȡ������") = 1 Then Resume
End Sub

Private Sub Menu_Manage_����Ӱ��()
On Error GoTo errhandle
    Dim lngResult As Long
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_����Ӱ��", vbInformation
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
    
    Call ReleationImage(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, 2, True)
Exit Sub
errhandle:
    Call HintError(err, "Menu_Manage_����Ӱ��", False)
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
    Dim strModuleTag As String

    If Not mblnInitOk Then Exit Sub
    
    mblnInitOk = False
    
    Set CtlFont = New StdFont
    
    strFontType = IIf(IsUseClearType = True, "΢���ź�", "����")
    
    CtlFont.Name = strFontType
    CtlFont.Size = gbytFontSize
            
    mstrSelQueueRooms = ""
    
    If mlngCur����ID <> Control.DescriptionText Or (Control.DescriptionText <> 0 And mblnAllDepts = True) Then
        mstrRPTExecutor = UserInfo.����
        
        stbThis.Panels(4).Text = "����ҽ����" & mstrRPTExecutor & "   ���ҽ����" & Split(stbThis.Panels(4).Text, "���ҽ����")(1)
                
        Set mobjCurStudyInfo = GetNullAdviceInf
        
        '�����л�������û�����´����˵��͹���ģ�飬Ҳû�е���cbrMain.RecalcLayout�������Ҫʹ�øö������ÿ����л���Ŀ�����Ϣ
        Set objDepartmentMenu = cbrMain.FindControl(, conMenu_View_Filter * 10#)
        
        If Control.DescriptionText = 0 Then
            'ѡ�����п���
            mblnAllDepts = True
            mlngCur����ID = 0
        
            If Not objDepartmentMenu Is Nothing Then objDepartmentMenu.Caption = "ȫ������"
            
            Call mobjPacsQueryWrap.DepChange(mstrCanUse����IDs, True)
            Set cbrFilter.options.Font = CtlFont
            
            If Not mobjQueue Is Nothing And mlngModule = G_LNG_PACSSTATION_MODULE Then
                mobjQueue.ChangeToAllDept mblnAllDepts
            End If
        Else
            'ѡ�񵥸�����
            mblnAllDepts = False
            
            mlngCur����ID = Control.DescriptionText
            mstrCur���� = Mid(Control.Caption, 1, InStrRev(Control.Caption, "(") - 1)
             
            If Not objDepartmentMenu Is Nothing Then objDepartmentMenu.Caption = mstrCur����
            
            Call SetParaUseImgSignValid(mlngCur����ID)
            Call InitDeptParameter(mlngCur����ID)
            
            Call ucPacsHelper1.Init(Me, mlngModule, mlngCur����ID, mstrPrivs, True)

            
            If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.zlInitModule(Me, mlngModule, mstrPrivs, mlngCur����ID)
            If Not mobjWork_PacsImg Is Nothing Then Call mobjWork_PacsImg.zlInitModule(Me, mlngModule, mstrPrivs, mlngCur����ID)
            If Not mobjWork_His Is Nothing Then Call mobjWork_His.zlModule.zlInitModule(Me, mlngModule, mstrPrivs, mlngCur����ID)
            If Not mobjWork_ImageCap Is Nothing Then Call mobjWork_ImageCap.zlInitModule(gcnOracle, mobjCapLinker, glngSys, mlngModule, mstrPrivs, mlngCur����ID, Me.hwnd, gblnUseDebugLog)
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.zlInitModule(Me, mlngModule, mstrPrivs, mlngCur����ID, mobjCapLinker, ucPacsHelper1)

     
            '�������༭�����Ͳ�ͬ������Ҫ��ʼ������ͬ���ҿ���ʹ�ò�ͬ�ı���༭��
            If TabWindow.Selected.Caption = C_TAB_NAME_��鱨�� Then
                Call SelectModule(C_TAB_NAME_��鱨��, strModuleTag, True)
                TabWindow.Selected.tag = strModuleTag
            Else
                'ֻ��Ҫ������ģ���tag����Ϊ�գ���Ϊ����ģ�鲻ͬ���ҿ���ʹ�ò�ͬ�ı༭����
                For i = 1 To TabWindow.ItemCount
                    If TabWindow.Item(i).Caption = C_TAB_NAME_��鱨�� Then
                        TabWindow.Item(i).tag = ""
                        Exit For
                    End If
                Next
            End If
            
            
            If Not mobjCapLinker Is Nothing Then Call mobjCapLinker.Init(Me, mlngCur����ID, mstrPrivs)


            '�����л�������������Ŷӽкţ�������Ŷӽк�ҳ��
            If mSysPar.blnUseQueue = True Then
                Call CreateTabItem(13, C_TAB_NAME_�Ŷӽк�, 10011, "")
                
                If mobjQueue Is Nothing Then
'                    mstrWorkModule = mstrWorkModule & ";�Ŷӽк�ģ��;"
'
'                    Set mobjQueue = New frmWork_Queue
'                    Call mobjQueue.zlInitPacsQueueCfg(mlngModule, mlngCur����ID, zlStr.NeedName(mstrCur����), mstrPrivs, mblnAllDepts, Me)
'
'                    TabWindow.InsertItem 13, "�Ŷӽк�", mobjQueue.hwnd, 10011
'                    TabWindow.Item(TabWindow.ItemCount - 1).tag = "�Ŷӽк�"
'
'                    Call picWindow_Resize
                    Call VerifyModuleObj(C_TAB_NAME_�Ŷӽк�)
                Else
                    Call mobjQueue.zlInitPacsQueueCfg(mlngModule, mlngCur����ID, zlStr.NeedName(mstrCur����), mstrPrivs, mblnAllDepts, Me)
                End If
                
                
                Call picTabFace_Resize
                
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
                            If TabWindow.Item(i).Selected Then
                                TabWindow.Item(0).Selected = True
                            End If
                            
                            Call TabWindow.RemoveItem(i)
                            Exit For
                        End If
                    Next i
                    
                    mobjQueue.CloseQueueQuick
                    
                    mobjQueue.Visible = False
                    
'                    Unload mobjQueue
'                    Set mobjQueue = Nothing
                    
                    Call picTabFace_Resize
'                    Call picWindow_Resize
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
            
            Call mobjPacsQueryWrap.DepChange(mlngCur����ID, False)
            
            '�ж��Ƿ���Ҫ�л�����
            Call mobjPacsQueryWrap.CurPacsQuery.LoadQueryScheme(glngSys, mlngModule, mlngCur����ID, UserInfo.ID)
            
            Call ExecuteDefaultQueryScheme
            
            Set cbrMenuBar = cbrMain.FindControl(, conMenu_Manage_Query)
            
            Call mobjPacsQueryWrap.RefreshCustomQueryMenu(cbrMenuBar, mintQueryState, tabScheme, mSysPar.blnQuickTabDisplayScheme)
            With cbrMenuBar.CommandBar
                Set objControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_QueryCFG, "��ѯ����", "", 0, True)
                Set objControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_QueryCfgUserScheme, "���÷�������", "", 0, False)
                Set objControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_QueryTabDisplayScheme, "��ʾ���÷�����ǩ", "", 0, True)
                
                objControl.Checked = mSysPar.blnQuickTabDisplayScheme
                objControl.CloseSubMenuOnClick = False
            End With
            
            Set cbrFilter.options.Font = CtlFont
        End If
        
        Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
         
        
        Call cbrMain.RecalcLayout
        
        'ˢ���Ŷӽк�ģ�����ݣ�����Ѿ�����
        Call RefreshPacsQueueData(False)
        
        Call CreateAuditorMenu(cbrMain.FindControl(, conMenu_ManagePopup).CommandBar.FindControl(, conMenu_Manage_SendAudit))
        
        If CheckPopedom(mstrPrivs, "���ԤԼ") Then
            'ˢ���Ƿ�����ԤԼ
            Call IsSchedule(mlngCur����ID, mobjCurStudyInfo.lngAdviceId, 0, mblnIsScheduleDept, mblnIsScheduleOrder)
        Else
            mblnIsScheduleDept = False
            mblnIsScheduleOrder = False
        End If
    End If
    
    If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
        glngXWDeptID = mlngCur����ID
    End If
    
    mblnInitOk = True
    
    '�ָ�ģ��ҳǩ��ʾ
    If Not TabWindow.Selected Is Nothing Then
        Call TabWindow_SelectedChanged(TabWindow.Selected)
    End If
Exit Sub
errhandle:
    mblnInitOk = True
    If HintError(err, "Menu_Dept_Select", False) = 1 Then Resume
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
            
            If Val(NVL(rsTemp!�Ƿ���빤����)) = 1 Then
                If blFirst = True Then
                    Set cbrControl = CreateMenu(cbrControls, xtpControlButton, conMenu_Manage_PacsPlugIn * 10000# + j, NVL(rsTemp!��������), "", 2325, True)
                    blFirst = False
                Else
                    Set cbrControl = CreateMenu(cbrControls, xtpControlButton, conMenu_Manage_PacsPlugIn * 10000# + j, NVL(rsTemp!��������), "", 2325, False)
                End If
                
                cbrControl.Parameter = NVL(rsTemp!VBS�ű�)
                cbrControl.DescriptionText = Val(NVL(rsTemp!ִ������))
                cbrControl.Category = Val(NVL(rsTemp!��������)) & "," & Val(NVL(rsTemp!�Ƿ�����Ҽ��˵�)) & "," & Val(NVL(rsTemp!�Ƿ���빤����))
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
                
                If lngAppId <> NVL(rsTemp!ID) Then
                    Set cbrPopControl = CreateMenu(.Controls, xtpControlButtonPopup, conMenu_Manage_PacsPlugLevel2 * 10000# + NVL(rsTemp!ID), NVL(rsTemp!��������), "", , False)
                    lngAppId = NVL(rsTemp!ID)
                Else
                    Set cbrPopControl = cbrMain.FindControl(, conMenu_Manage_PacsPlugLevel2 * 10000# + NVL(rsTemp!ID), , True)
                End If

                If Not cbrPopControl Is Nothing Then
                    If blFirstMenu Then
                        Set cbrControl = CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsPlugIn * 10000# + j, NVL(rsTemp!��������), "", , True)
                    Else
                        Set cbrControl = CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsPlugIn * 10000# + j, NVL(rsTemp!��������), "", , False)
                    End If
                End If
                                
                cbrControl.Parameter = NVL(rsTemp!VBS�ű�)
                cbrControl.DescriptionText = Val(NVL(rsTemp!ִ������))
                cbrControl.Category = Val(NVL(rsTemp!��������)) & "," & Val(NVL(rsTemp!�Ƿ�����Ҽ��˵�)) & "," & Val(NVL(rsTemp!�Ƿ���빤����))
                
                blFirstMenu = False
                
                Call rsTemp.MoveNext
            Wend
        End If
            
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
    If HintError(err, "Menu_View_Refresh_click", False) = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Home_click()
On Error GoTo errhandle
    zlHomePage hwnd
Exit Sub
errhandle:
    If HintError(err, "Menu_Help_Web_Home_click", False) = 1 Then Resume
End Sub

Private Sub Menu_View_StatusBar_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errhandle
    Me.stbThis.Visible = Not Me.stbThis.Visible
    Control.Checked = Not Control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
errhandle:
    If HintError(err, "Menu_View_StatusBar_click", False) = 1 Then Resume
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
    If HintError(err, "Menu_View_ToolBar_Button_click", False) = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Size_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errhandle
    Me.cbrMain.options.LargeIcons = Not Me.cbrMain.options.LargeIcons
    Control.Checked = Not Control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
errhandle:
    If HintError(err, "Menu_View_ToolBar_Size_click", False) = 1 Then Resume
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
    If HintError(err, "Menu_View_ToolBar_Text_click", False) = 1 Then Resume
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
    If HintError(err, "GetDeptName", False) = 1 Then Resume
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
    Dim lngHwnd As Long
    Dim blnIsEmbedReport As Boolean
    
    '�жϵ�ǰ����ģ���Ƿ�Ӱ��ɼ�ģ�飬����ǣ����жϲɼ�ģ���Ƿ��ʼ��������Ѿ���ʼ�������˳��ù��̣�����Ͷ�����г�ʼ��������ʾ
    '��Ϊ��ͬһ����̨�У����ͬʱ�򿪲�����Ƶ�ɼ�ģ�齫���л�������һϵͳ�˳�ʱ���ɼ�ģ��Ҳ�����ͷţ�����л��ص�ǰϵͳ����Ҫ�ж��Ƿ���³�ʼ���ɼ�ģ��
'    Call Form_Resize

    If Not mblnInitOk Then Exit Sub
    
    If TabWindow.Selected Is Nothing Then Exit Sub
    
    'ע���������ʽ���洰����Ƕ������Ƶ�ɼ����������洰�ڲ��л�Ƕ��ʽ��Ƶ�ɼ�����ʾ
    If mstrSelTabName = C_TAB_NAME_Ӱ��ɼ� Then
        'ֻ�й���ģ����Ӱ��ɼ�ʱ������Ҫ�л�Ƕ��ʽ��Ƶ�ɼ�����ʾ
        If mobjWork_ImageCap Is Nothing Then Exit Sub
        
        
        '��������˸����ɼ����ڣ��򲻽���Ƕ��ʽ����
        If mobjWork_ImageCap.VideoDockState Then
            mobjCapLinker.ReportAdviceId = 0
            Exit Sub
        End If
        
        lngHwnd = VerifyModuleObj(C_TAB_NAME_Ӱ��ɼ�)
        
        '�����Ƶ�������������ڣ�����Ҫ����Ƕ��
        If GetAncestor(mobjWork_ImageCap.VideoHwnd, GA_PARENT) = lngHwnd Then
            mobjCapLinker.ReportAdviceId = 0
            Exit Sub
        End If
          
        '�����Ƶ�ɼ�û��Ƕ�뱨�洰�ڣ�����Ƶ�ɼ�Ƕ�뵱ǰ������
        If VideoIsAttachReportWindow = False Then
            Call EmbedWindow(lngHwnd)
            
            mobjCapLinker.ReportAdviceId = 0
            
            '��Ҫ���ô˷�����ʾ����ǰ��Ƶ
            Call mobjWork_ImageCap.zlRefreshVideoWindow
            
            Call mobjWork_ImageCap.zlRestoreWindow(IIf(mobjCurStudyInfo.intStep > 1 And mobjCurStudyInfo.intStep < 5, False, True), True)
        End If
    Else
        '�������ʽ������д����û��Ƕ����Ƶ�ɼ������ڹ���ģ��֮���л�ʱ����ҪǶ����Ƶ�ɼ�
        If Not mobjCapLinker Is Nothing And VideoIsAttachReportWindow = False Then
            mobjCapLinker.ReportAdviceId = 0
            Call ucPacsHelper1.ShowEmbedVideo(mobjCapLinker)
        End If
    End If
    
Exit Sub
errhandle:
    If HintError(err, "Form_Activate", False) = 1 Then Resume
End Sub



Private Sub imgFun_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    timFun.Enabled = False
End Sub

Private Sub mobjCapLinker_OnAfterChange(ByVal blnIsAfter As Boolean)
    Call ResetFloatingVideoState(mobjCurStudyInfo)
End Sub

Private Sub mobjCapLinker_OnLockChange(ByVal blnIsLock As Boolean)
    Call ResetFloatingVideoState(mobjCurStudyInfo)
End Sub

Private Sub mobjPacsQueryWrap_OnQueueRoomChanged()
    Call RefreshPacsQueueData(False)
End Sub

Private Sub mobjPacsQueryWrap_OnSwipeCard()
On Error GoTo errH
    Call VsfListDbClick(True)
    Exit Sub
errH:
    If HintError(err, "mobjPacsQueryWrap_OnSwipeCard", False) = 1 Then Resume
End Sub

Private Sub mobjPacsQueryWrap_OnClearFace()
'�������ݣ�ִ�в�ѯ��û������ʱ����ս���ؼ���ʾ
On Error GoTo errhandle
    Dim i As Long
    
    If vsfList.Rows < 2 Then
        '��û������ʱ��֪ͨˢ�¹���ģ������ص�����
        Set mobjCurStudyInfo = GetNullAdviceInf
        Call RefreshModuleData(mstrSelTabName, mstrSelModuleTag, mobjSelModule)

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
        
        
        Call mobjPacsQueryWrap.FillAppend(0, 0, False, rtxtAppend)
        
        stbThis.Panels(2).Text = "�� " & 0 & " ����¼": stbThis.Panels(2).Alignment = sbrCenter
        stbThis.Panels(3).Text = ""

    End If
    Exit Sub
errhandle:
    If HintError(err, "mobjPacsQueryWrap_OnClearFace", False) = 1 Then Resume
End Sub

Private Sub mobjWork_Report_AfterSetRptQuality(ByVal lngAdviceId As Long, ByVal strValue As String)
    mobjCurStudyInfo.strImageQuality = strValue
    Call UpdateQueryListData(Nothing, lngAdviceId)
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
        Call mobjPacsQueryWrap.Find(True, True)
        TimerRefresh.Enabled = False
        Me.MousePointer = 0
    Else
        Exit Sub
    End If
    Exit Sub
errhandle:
    HintError err, "cmdFind_Click<���Ҳ���>", False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '���ع���ģ��ʱ���������˳�����
    If Not mblnInitOk Then
        Cancel = True
        Exit Sub
    End If
    
    If mblnMenuDownState Then
        If HintMsg("��ǰ������δ��ɣ�ǿ���˳�������ɳ����쳣���Ƿ������", "Form_QueryUnload", vbYesNo) = vbNo Then Cancel = True
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
            Call Menu_RichEPR(conMenu_PacsReport_Write)
            
        Case C_FUNC_STR_�鿴������Ϣ
            frmDegreeCard.ShowMe mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.lngPageID, Me
            
        Case C_FUNC_STR_��Ƭ
            If mobjPacsCore Is Nothing Then
                HintMsg "ͼ��鿴������Ч�����ܽ��д˲�����", "cbrMain_Execute", vbOKOnly
                Exit Sub
            End If
            
            Call OpenViewer(1, mobjPacsCore, mobjCurStudyInfo.lngAdviceId, True, Me, "", mobjCurStudyInfo.blnMoved)
             
        Case C_FUNC_STR_���
            Call Menu_Manage_����������
    End Select
    Exit Sub
errH:
    Call HintError(err, "imgFun_Click", False)
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
    Call HintError(err, "mobjCaptureHot_OnKeyBoardLHook", False)
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
        lngAdviceId = Val(NVL(rsData!node_value))
    End If
    
    
    Select Case strMsgItemIdentity
        Case G_STR_MSG_ZLHIS_CIS_017    '�������
            '������Ϣ��ʾ@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='patient_name'"
            strHint = "���� " & NVL(rsData!node_value) & " ��Ҫ���м�飬�뼰ʱ����"
            
            Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
            
            '�����ݿ���ˢ������
            Call UpdateQueryListData(Nothing, lngAdviceId)
            
        Case G_STR_MSG_ZLHIS_CIS_024    'ҽ������
            '����������ʾ@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='patient_name'"
            strHint = "���� " & NVL(rsData!node_value) & " �ļ��ҽ���ѱ������� "
        
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
            If NVL(rsData!node_value) = -1 Then
                
                '��Ҫ�жϵ�ǰ�û��Ƿ�Ϊ������
                strSQL = "select ������ from Ӱ�����¼ where ҽ��ID=[1]"
                Set rsReport = zlDatabase.OpenSQLRecord(strSQL, "��ѯ������", lngAdviceId)
                If rsReport.RecordCount > 0 Then
                    If NVL(rsReport!������) = UserInfo.���� Then
                        '������Ϣ
                        rsData.Filter = "node_name='patient_name'"
                        strHint = "����" & NVL(rsData!node_value) & "�ı����ѱ����أ���ע�⴦��"
                        
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
            strHint = "���� " & NVL(rsData!node_value) & "��"
            
            rsData.Filter = "node_name='check_item_title'"
            strHint = strHint & "�����Ŀ " & NVL(rsData!node_value) & " ����Σ�������"
            
            Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
        
            '�����б��е���ʾ״̬
            lngRowIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("ҽ��ID"))
            
            If lngRowIndex > 0 Then Call UpdateQueryListData(Nothing, lngAdviceId)
            
    End Select
    
    Exit Sub
errH:
    If HintError(err, "mobjMsgCenter_OnRecevieMsg") = 1 Then Resume
End Sub

Private Sub mobjPacsCore_AfterSaveOuterImage(strStudyUID As String)
    '�������ⲿͼ��ˢ��ͼ��������б�
On Error GoTo errhandle
    
    'û�м�¼���˳�
    If mobjCurStudyInfo.lngAdviceId = 0 Then Exit Sub
    
    '�ǵ�ǰ�ļ�飬��ˢ�¼��������б�
    If mobjCurStudyInfo.strStudyUID = strStudyUID Then
        Call mobjWork_PacsImg.zlRefreshFace(mobjCurStudyInfo, True)
    End If
    
    Exit Sub
errhandle:
    '������
End Sub


Private Sub ReleationImage(ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal intStep As Integer, ByVal lngReleationType As Long, ByVal blnUseMenuReleation As Boolean)
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
    
    Exit Sub
errH:
    If HintError(err, "ReleationImage") = 1 Then Resume
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
                            If Not mobjWork_PacsImg Is Nothing Then Call mobjWork_PacsImg.zlMenu.zlPopupMenu(mstrSelTabName, Popup)
                            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.zlMenu.zlPopupMenu(mstrSelTabName, Popup)
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
    If HintError(err, "mobjQueryShow_OnMouseUp", False) = 1 Then Resume
End Sub


Private Sub LocateMainWorkModuleTab()
On Error GoTo errH
'�ָ���Ҫ����ҳ�棬�����������Ҫ����ҳ�棬�л����ʱ�����л�����Ӧҳ��
    Dim i As Integer
    
    If Len(mSysPar.strFirstTab) <= 0 Then Exit Sub
    If InStr(mstrSelTabName, mSysPar.strFirstTab) > 0 Then Exit Sub
    
    For i = 0 To TabWindow.ItemCount - 1
        If InStr(TabWindow.Item(i).Caption, mSysPar.strFirstTab) > 0 And TabWindow.Item(i).Visible Then
            TabWindow.Item(i).Selected = True
            Exit Sub
        End If
    Next
errH:
End Sub



Private Sub mobjPacsQueryWrap_OnChangeData(ByVal blnRefreshModul As Boolean, ByVal blnIsSelChange As Boolean)
On Error GoTo errH
'blnRefreshModul �Ƿ���Ҫˢ��ģ��

    Dim i As Integer
    Dim intCol As Integer
    Dim lngRow As Long
    Dim lngAdviveID As Long 'ҽ��ID
    Dim strInfo As String
    Dim blnRefreshFace   As Boolean '�Ƿ���Ҫˢ�½���
    Dim strCurModuleTag As String
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''ˢ�±�����Ϣ
    
    If mblnIsPrintMode Then Exit Sub
    
    intCol = vsfList.ColIndex("ҽ��ID")
    
    lngRow = vsfList.RowSel
    If lngRow = -1 Then Exit Sub

    lngAdviveID = Val(vsfList.TextMatrix(lngRow, intCol))
 
    Set mobjCurStudyInfo = mobjPacsQueryWrap.StudyInfo
    mobjCurStudyInfo.lngReportEditState = GetReportEditState(mobjCurStudyInfo)
    
    If blnIsSelChange Then Call LocateMainWorkModuleTab
    
    Call DoLabFlag
    
    mintImgCount = GetScanRequestCount(mobjCurStudyInfo.lngAdviceId)
    
     
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''ˢ�½�����Ϣ
    '������ϸ��Ϣ
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
        If mobjCurStudyInfo.strMarkNum > 0 Then labCollectionInfo = "��:" & mobjCurStudyInfo.strMarkNum & "  "
    ElseIf mobjCurStudyInfo.lngPatientFrom = 2 Then
        If mobjCurStudyInfo.strMarkNum > 0 Then labCollectionInfo = "ס:" & mobjCurStudyInfo.strMarkNum & "  "
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
                HintMsg "δ֪�ļ�����", "mobjPacsQueryWrap_OnSelChange", vbInformation
            End If
    End Select
    
    imgStep.Visible = True
    
    '�ж��Ƿ���Ҫ����ģ��,����༭����Pacs�༭�������Ӳ����༭���������ĵ��༭������ҽ����¼(סԺҽ��������ҽ��)�����ü�¼��סԺ���ã�������ã�
    strCurModuleTag = GetWorkModuleName(mstrSelTabName, mobjCurStudyInfo.lngExeDepartmentId, mobjCurStudyInfo.lngPatientFrom)
    If strCurModuleTag <> "" And strCurModuleTag <> mstrSelModuleTag Then
       Call SelectModule(mstrSelTabName, strCurModuleTag)
       TabWindow.Selected.tag = strCurModuleTag
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''ˢ��ģ����Ϣ
    If blnRefreshModul Then Call RefreshModuleData(mstrSelTabName, mstrSelModuleTag, mobjSelModule)
    
    'ˢ���Ƿ�����ԤԼ
    If CheckPopedom(mstrPrivs, "���ԤԼ") Then
        Call IsSchedule(mlngCur����ID, mobjCurStudyInfo.lngAdviceId, 0, mblnIsScheduleDept, mblnIsScheduleOrder)
    Else
        mblnIsScheduleDept = False
        mblnIsScheduleOrder = False
    End If
    
    
    Exit Sub
errH:
    If HintError(err, "mobjPacsQueryWrap_OnSelChange", False) = 1 Then Resume
End Sub


Private Sub mobjQueue_OnCallAboutLock(ByVal lngType As Long, strLockedName As String, ByVal blnLockPara As Boolean)
On Error GoTo errhandle
'104686��أ����к�������飬
'lngType����  1:�ж��Ƿ������˲��������Ƿ��Ѿ��б������ļ��,����ֱ�ӽ���        2:���²���
'strLockedName   ��="" ������û��Ӱ�죬����˵���Ѿ����ò������ҷ���֮ǰ�����ļ�黼������
'blnLockPara   ���ڸ���PacsMain�еĲ���
            
    If lngType = 2 Then
    '���²���
        mSysPar.blnLockAfterCall = blnLockPara
    End If
    
    Exit Sub
errhandle:
    If HintError(err, "mobjQueue_OnCallAboutLock", False) = 1 Then Resume
End Sub

Private Sub mobjQueue_OnCalled(ByVal lngAdviceId As Long, ByVal strRoom As String, ByVal TCallWay As zlQueueOper.TCallWay)
    Dim intRowIndex As Integer
On Error GoTo errhandle
 
    intRowIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("ҽ��ID"))
    Call QueueDataConsistency(lngAdviceId, strRoom, intRowIndex)
    
    If TCallWay = cwBroadcast Or TCallWay = cwWaitRoom Then Exit Sub
        
    If mSysPar.blnLockAfterCall = False Then Exit Sub
    
    If mobjCapLinker Is Nothing Then Exit Sub
    If mobjWork_ImageCap Is Nothing Then Exit Sub
    
    '�����߼��ж��Ƿ����á�ͬ����λ������б�����δ���ã���Ҫ����ҵ��ID��ȡ��Ҫ�����ļ�飬���Ѿ����ã�ֻ��Ҫ������
    'intRowIndex=-1˵������б���û����ʾ�Ŷ��б������ݣ���Ҫ����������
    If mSysPar.blnSynStudylist Then
        If intRowIndex > 0 And mobjCurStudyInfo.lngAdviceId <> lngAdviceId Then
            'ͬ����λ
            Call mobjPacsQueryWrap.LocateRow(intRowIndex)
        End If
    End If
         
    mobjCapLinker.LockAdviceId = lngAdviceId
    Call mobjWork_ImageCap.ResetLockState(True)
        
    Exit Sub
errhandle:
    If HintError(err, "mobjQueue_OnCalled") = 1 Then Resume
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
    If HintError(err, "mobjQueue_OnQueueQuick", False) = 1 Then Resume
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
        
        Call InitWorkModuleTab
        
        mblnInitOk = True
    
'        Call picWindow_Resize
        picTabFace_Resize
        
        If Not TabWindow.Selected Is Nothing Then Call TabWindow_SelectedChanged(TabWindow.Selected)
        
    End If
    
    Call Unload(frmWorkModule)
Exit Sub
errhandle:
    If HintError(err, "conMenu_WorkModule_Click", False) = 1 Then Resume
End Sub

Private Function ReoprtPrint(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean, _
    Optional ByVal blnIsPrint As Boolean = False, Optional ByVal strSpecifyReportId As String = "", _
    Optional ByVal strPrintFmts As String = "") As Boolean
'�����ӡԤ��
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim blnIsDocEditor As Boolean
    Dim objReportV2 As frmReportV2
    Dim objRichV2 As frmRichReportV2
    
    ReoprtPrint = False
    blnIsDocEditor = False
    
    strSQL = "Select RAWTOHEX(��鱨��ID) as ��鱨��ID From ����ҽ������ Where ҽ��ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��鱨��ID", lngAdviceId)
    If rsData.RecordCount > 0 Then
        If NVL(rsData!��鱨��ID) <> "" Then blnIsDocEditor = True
    End If
    
    '��Ҫ�жϱ���༭������
    If blnIsDocEditor = False Then
        Set objReportV2 = New frmReportV2
        
        Call objReportV2.zlInit(Me, mlngModule, mlngCur����ID, mstrPrivs, Nothing, Nothing, False)
        
        ReoprtPrint = objReportV2.PrintPreview(lngAdviceId, blnMoved, blnIsPrint, Val(strSpecifyReportId), strPrintFmts)
        
        Unload objReportV2
        Set objReportV2 = Nothing
    Else
        Set objRichV2 = New frmRichReportV2
        
        Call objRichV2.zlInit(Me, mlngModule, mlngCur����ID, mstrPrivs)
        
        Call objRichV2.zlRefresh(lngAdviceId, blnMoved, True, strSpecifyReportId)
        
        Call objRichV2.PrintPreview(Not blnIsPrint)
        
        ReoprtPrint = True
        
        Unload objRichV2
        Set objRichV2 = Nothing
    End If
    
    Set objReportV2 = Nothing
    Set objRichV2 = Nothing
End Function



Private Sub cbrMain_Execute(ByVal objControl As XtremeCommandBars.ICommandBarControl)
On Error GoTo errhandle
    Dim Control As XtremeCommandBars.ICommandBarControl
    Dim i As Long
    Dim str��ʦһ As String, str��ʦ�� As String, strִ�м� As String
    Dim intRowIndex As Integer
    Dim strSys1 As String
    Dim strSys2 As String
    Dim bytSize As Byte
    Dim strTmp As String
    Dim objReport As frmReportV2
    
    If mintQueryState <> 1 And objControl.ID <> conMenu_Manage_Query And objControl.ID <> conMenu_Manage_QueryCFG Then
        HintMsg "û�п��ò�ѯ���ã����ڲ�ѯ���������н�����ӡ�", "cbrMain_Execute", vbInformation
        Exit Sub
    End If
    
    If mblnMenuDownState Then Exit Sub
    
    '������Ҫ����id���Ҷ�Ӧ�Ĳ˵���Ŀ����Ϊͨ���󶨿�ݼ�ִ��ʱ����������һ��ֻ��id��û�������κ���Ϣ��control�˵���
    Set Control = cbrMain.FindControl(, objControl.ID, , True)
    If Control Is Nothing Then
        '����ò˵�Ϊ���Ӳ����༭�����Ҽ��˵�������Ҫ�޸��Ҽ��˵���id����Ϣ
        If Not mobjWork_Report Is Nothing Then
            Call mobjWork_Report.ReplacePopupMenu(objControl)
            
            Set Control = cbrMain.FindControl(, objControl.ID, , True)
        End If
        
        If Control Is Nothing Then Exit Sub
    End If
    
    If Control.ID = 0 Then Exit Sub
    
    If Not (Control.ID > conMenu_Manage_PacsPlugIn * 10000# And Control.ID < conMenu_Manage_PacsPlugIn * 10000# + 100) And objControl.ID <> conMenu_Manage_PacsPlugCfg Then
        '�����ִ�в���˵���������Ҫ�㲥�¼�
        Call Notify.Broadcast(BM_SYS__EVENT_MENU, 0, mobjCurStudyInfo.lngAdviceId, objControl.ID, objControl.Category)
    End If
    
    mblnMenuDownState = True
        
    cbrMain.RecalcLayout
    
    If Not mobjSelModule Is Nothing Then
        '���ж��Ƿ������˵�����
        If Not mobjWork_Pathol Is Nothing Then
            If mobjWork_Pathol.zlMenu.zlIsModuleMenu(mstrSelModuleTag, Control) Then
                Call mobjWork_Pathol.zlMenu.zlExecuteMenu(mstrSelModuleTag, Control.ID)
                
                mblnMenuDownState = False
                Exit Sub
            End If
        End If
        
        
        Select Case mstrSelTabName
            Case C_TAB_NAME_Ӱ��ͼ��
                If mobjSelModule.zlMenu.zlIsModuleMenu(mstrSelModuleTag, Control) Then
                    Call mobjSelModule.zlMenu.zlExecuteMenu(mstrSelModuleTag, Control.ID)
                    
                    mblnMenuDownState = False
                    Exit Sub
                End If
                
            Case C_TAB_NAME_Ӱ��ɼ�
            
            Case C_TAB_NAME_�걾����, C_TAB_NAME_����ȡ��, C_TAB_NAME_������Ƭ, C_TAB_NAME_�����ؼ�, C_TAB_NAME_���̱���
                If mobjSelModule.zlMenu.zlIsModuleMenu(Control) Then
                    Call mobjSelModule.zlMenu.zlExecuteMenu(Control.ID)
                    
                    mblnMenuDownState = False
                    Exit Sub
                End If
                
            Case C_TAB_NAME_ҽ����¼, C_TAB_NAME_������¼, C_TAB_NAME_���Ӳ���, C_TAB_NAME_���ü�¼
                If mobjWork_His.zlMenu.zlIsModuleMenu(mstrSelModuleTag, Control) Then
                    If mintChangeUserState = 2 Then  '�������û������������
                        HintMsg "��ͳһ�û����ٲ���", "cbrMain_Execute", vbInformation
                    Else
                        Call mobjWork_His.zlMenu.zlExecuteMenu(mstrSelModuleTag, Control.ID)
                    End If
                    
                    mblnMenuDownState = False
                    Exit Sub
                End If
            Case C_TAB_NAME_��鱨��
                If mobjWork_Report.zlMenu.zlIsModuleMenu(mstrSelModuleTag, Control) Then
                    Call mobjWork_Report.zlMenu.zlExecuteMenu(mstrSelModuleTag, Control.ID)
                    
                    mblnMenuDownState = False
                    Exit Sub
                End If
        End Select
    End If
 
        
    Select Case Control.ID
        Case conMenu_Img_OpenView       '��Ƭ
            If mobjWork_PacsImg Is Nothing Or mstrSelTabName <> C_TAB_NAME_Ӱ��ͼ�� Then
                If mobjPacsCore Is Nothing Then
                    mblnMenuDownState = False
                    HintMsg "ͼ��鿴������Ч�����ܽ��д˲�����", "cbrMain_Execute", vbOKOnly
                    Exit Sub
                End If
                 
                If mobjCurStudyInfo.strStudyUID <> "" Then
                    Call OpenViewer(1, mobjPacsCore, mobjCurStudyInfo.lngAdviceId, False, Me, "", mobjCurStudyInfo.blnMoved)
                Else
                    Call OpenLatestImage(Me, mobjPacsCore, mobjCurStudyInfo, mSysPar.lngAutoImageDays)
                End If
            Else
                Call mobjWork_PacsImg.zlMenu.zlExecuteMenu("", conMenu_Img_Look + mobjWork_PacsImg.zlMenu.zlBaseMenuID)
            End If
            
        Case conMenu_img_ContrastView   '�Աȹ�Ƭ
            If mobjWork_PacsImg Is Nothing Or mstrSelTabName <> C_TAB_NAME_Ӱ��ͼ�� Then
                If mobjPacsCore Is Nothing Then
                    mblnMenuDownState = False
                    HintMsg "ͼ��鿴������Ч�����ܽ��д˲�����", "cbrMain_Execute", vbOKOnly
                    Exit Sub
                End If
                
                Call OpenViewer(1, mobjPacsCore, mobjCurStudyInfo.lngAdviceId, True, Me, "", mobjCurStudyInfo.blnMoved)
            Else
                Call mobjWork_PacsImg.zlMenu.zlExecuteMenu("", conMenu_Img_Contrast + mobjWork_PacsImg.zlMenu.zlBaseMenuID)
            End If
            
        Case conMenu_Check_ViewLink
            Call ViewLinkChecks
        
        Case conMenu_PacsReport_Preview '����Ԥ��
            If Not mobjWork_Report Is Nothing And mstrSelTabName = C_TAB_NAME_��鱨�� Then
                
                strTmp = GetWorkModuleTag(C_TAB_NAME_��鱨��)
                Call mobjWork_Report.zlMenu.zlExecuteMenu(strTmp, conMenu_File_Preview + mobjWork_Report.zlMenu.zlBaseMenuID)
            Else
                Call ReoprtPrint(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.blnMoved, False)
            End If
            
        Case conMenu_PacsReport_Print   '�����ӡ
            If Not mobjWork_Report Is Nothing And mstrSelTabName = C_TAB_NAME_��鱨�� Then
                
                strTmp = GetWorkModuleTag(C_TAB_NAME_��鱨��)
                Call mobjWork_Report.zlMenu.zlExecuteMenu(strTmp, conMenu_File_Print + mobjWork_Report.zlMenu.zlBaseMenuID)
            Else
                Call ReoprtPrint(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.blnMoved, True)
            End If
            
            
        Case conMenu_PacsReport_Write
'            If mobjWork_Report Is Nothing Then
'                '��������ģ�飬���л�
'                For i = 0 To TabWindow.ItemCount - 1
'                    If TabWindow(i).Caption = C_TAB_NAME_��鱨�� Then
'                        TabWindow(i).Selected = True
'                        Exit For
'                    End If
'                Next
'            End If
            
            If mstrSelModuleTag <> C_TAB_NAME_��鱨�� Then
                strTmp = GetWorkModuleName(C_TAB_NAME_��鱨��, mobjCurStudyInfo.lngExeDepartmentId, mobjCurStudyInfo.lngPatientFrom)
            Else
                strTmp = GetWorkModuleTag(C_TAB_NAME_��鱨��)
            End If
            
            '���ñ���༭����װ�Ĳ˵�����д����
            If Not mobjWork_Report Is Nothing And mstrSelTabName = C_TAB_NAME_��鱨�� Then
            
                Select Case strTmp
                    Case C_WORKMODULE_NAME_�ϰ汨��
                        Call mobjWork_Report.zlMenu.zlExecuteMenu(C_WORKMODULE_NAME_�ϰ汨��, conMenu_PacsReport_Open + mobjWork_Report.zlMenu.zlBaseMenuID)
                     
                    Case C_WORKMODULE_NAME_��������
                        '���Ӳ����༭��
                        If Control.Caption = "����" Then    '����
                            Call mobjWork_Report.zlMenu.zlExecuteMenu(C_WORKMODULE_NAME_��������, conMenu_File_Open + mobjWork_Report.zlMenu.zlBaseMenuID)
                        ElseIf Control.Caption = "�޶�" Then    '�޶�
                            Call mobjWork_Report.zlMenu.zlExecuteMenu(C_WORKMODULE_NAME_��������, conMenu_Edit_Audit + mobjWork_Report.zlMenu.zlBaseMenuID)
                        Else                            '��д
                            Call mobjWork_Report.zlMenu.zlExecuteMenu(C_WORKMODULE_NAME_��������, conMenu_Edit_Modify + mobjWork_Report.zlMenu.zlBaseMenuID)
                        End If
                        
                    Case Else
                        '�����ĵ��༭��,����ʾ����ʽ�����ĵ��༭��
                        Call mobjWork_Report.zlMenu.zlExecuteMenu(C_WORKMODULE_NAME_���ܱ���, conMenu_File_Open + mobjWork_Report.zlMenu.zlBaseMenuID)
                End Select
            Else
                If mSysPar.blnReportWithImage And Len(mobjCurStudyInfo.strStudyUID) <= 0 Then
'                    If MsgBoxD(Me, "���μ��δ�ҵ����ͼ���Ƿ�ǿ����д��", vbYesNo, "��ʾ") = vbNo Then
'                        mblnMenuDownState = False
'                        Exit Sub
'                    End If

                    Call MsgBoxD(Me, "���μ��δ�ҵ����ͼ�񣬲�����д��", vbOKOnly, "��ʾ")
                    
                    mblnMenuDownState = False
                    Exit Sub
 
                End If
                
                '��û�н����鱨��ģ��ҳʱ��ִ�еĴ򿪲���
                Select Case strTmp
                    Case C_WORKMODULE_NAME_�ϰ汨�� '�ϰ汨��༭��
                        '�ж��Ƿ�����Ѿ��򿪵ı���༭��
                        If LocateReportWindow(mobjCurStudyInfo.lngAdviceId) Then
                            mblnMenuDownState = False
                            Exit Sub
                        End If
                        
                        Set objReport = New frmReportV2

                        objReport.zlInit Me, mlngModule, GetCurDeptId, mstrPrivs, mobjCapLinker, Nothing
                        objReport.zlRefresh mobjCurStudyInfo

                        Call objReport.Show(0, Me)
                        Call objReport.SetReportTitle(mobjCurStudyInfo)
                        Call objReport.ReSetFormFontSize(FontSize)

                        '����ʽ��ʽ���б���༭ʱ����Ҫ�����㶨λ���༭��
                        objReport.LocateEditBox

                    Case C_WORKMODULE_NAME_�������� '���Ӳ����༭��
                        If mobjRichReportWrap Is Nothing Then
                            Set mobjRichReportWrap = New frmEPREditWrapV2
                        End If

                        If mobjRichReportWrap.InitEprEditor(Nothing, Me, mlngModule, GetCurDeptId) = False Then
                            mblnMenuDownState = False
                            Exit Sub
                        End If

                        If Control.Caption = "����" Then    '����
                            Call mobjRichReportWrap.ExecuteMenu(mobjCurStudyInfo, conMenu_File_Open)
                        ElseIf Control.Caption = "�޶�" Then    '�޶�
                            Call mobjRichReportWrap.ExecuteMenu(mobjCurStudyInfo, conMenu_Edit_Audit)
                        Else '��д
                            Call mobjRichReportWrap.ExecuteMenu(mobjCurStudyInfo, conMenu_Edit_Modify)
                        End If


                        Call IEventNotify_Broadcast(BM_REPORT_EVENT_OPEN, "1", mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo)

                    Case Else   '���ܱ���༭��
                        Call PreviewRichReport(Me, mlngModule, GetCurDeptId, mstrPrivs, mobjCurStudyInfo)
                End Select
            End If

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
            
        Case conMenu_Cap_DevSet         '��Ƶ����
            If Not mobjWork_ImageCap Is Nothing Then
                Call mobjWork_ImageCap.zlShowVideoConfig
                mstrCaptureHot = GetSetting("ZLSOFT", "����ģ��", "�ɼ��ȼ�", "F8")
                mstrCaptureAfterHot = GetSetting("ZLSOFT", "����ģ��", "��̨�ɼ��ȼ�", "F7")
                mstrCaptureAfterTagHot = GetSetting("ZLSOFT", "����ģ��", "��Ǹ����ȼ�", "F6")
            End If
            
        Case conMenu_Manage_ChangeUser
            '�����û�ʱ����Ҫ���жϱ����Ƿ���Ҫ����
            strTmp = GetWorkModuleTag(C_TAB_NAME_��鱨��)
            
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(mobjCurStudyInfo, strTmp, True)
            End If
        
            Call ChangeUser
            
            '�����û�����Ҫˢ�±���༭������Ϊ�û�������ԭ�б���ı༭�û����ߴ����û���Ҫ���и���
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(mobjCurStudyInfo, strTmp, True)
            End If
            
        Case conMenu_Manage_SwitchUser
            '�л��û�ʱ����Ҫ���жϱ����Ƿ���Ҫ����
            strTmp = GetWorkModuleTag(C_TAB_NAME_��鱨��)
            
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(mobjCurStudyInfo, strTmp, True)
            End If
            
            Call SwitchUser
            
            '�л��û�����Ҫˢ�±���༭������Ϊ�û��л���ԭ�б���ı༭�û����ߴ����û���Ҫ���и���
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(mobjCurStudyInfo, strTmp, True)
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
            
        Case conMenu_Cap_StudySyncState '�ɼ�����
            Call LockCapture(mobjCurStudyInfo)
            
'            If Not mobjWork_ImageCap Is Nothing Then
'                If mobjCurStudyInfo.blnMoved Or mobjCapLinker Is Nothing Then
'                    HintMsg "��ǰ���״̬������������", "cbrMain_Execute", vbOKOnly
'                    mblnMenuDownState = False
'                    Exit Sub
'                End If
'
'                mobjCapLinker.ReportAdviceId = 0    '��Ҫ��ձ���id����������ʱ������ʹ�ñ���ҽ��id��������
'                mobjCapLinker.LockAdviceId = mobjCurStudyInfo.lngAdviceId
'
'                Call mobjWork_ImageCap.ResetLockState(True)
'
'            End If
            
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
            If mobjPacsCore Is Nothing Then
                mblnMenuDownState = False
                HintMsg "ͼ��鿴������Ч�����ܽ��д˲�����", "cbrMain_Execute", vbOKOnly
                Exit Sub
            End If
            
            Call OpenViewer(1, mobjPacsCore, mobjCurStudyInfo.lngAdviceId, False, Me, "", mobjCurStudyInfo.blnMoved)
        
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
            Call Menu_Manage_SendAudit(0, Control.Caption)
            
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
            
            zlDatabase.SetPara "��ʾ���÷�����ǩ", IIf(mSysPar.blnQuickTabDisplayScheme, "1", "0"), glngSys, mlngModule
            
            tabScheme.Visible = mSysPar.blnQuickTabDisplayScheme
            tabScheme.tag = IIf(mSysPar.blnQuickTabDisplayScheme, "1", "0")
            
            
            Call AdjustFace(picList.Height, picList.Width)
'----------------------------------------�������������---------------------
        Case conMenu_Manage_PacsPlugCfg
            Call ShowPacsInterfaceCfg
        Case conMenu_Manage_PacsPlugIn * 10000# To conMenu_Manage_PacsPlugIn * 10000# + 100
            Call ExecutePluginInterfaceFun(Control.Caption, Control.Parameter, Control.DescriptionText, False)
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
            mblnIsForceRefresh = True
            
            Call RefreshList
            Call RefreshPacsQueueData(True)
            
            mblnIsForceRefresh = False
        Case comMenu_Cap_Process
            Call Menu_Manage_�����ɼ�
'---------------------------����----------------
        Case conMenu_Tool_Valid         'ͼ��У�Թ���
            
            If Len(Dir(GetAppRootPath & "zlPacsImageValid.exe")) > 0 Then
                If InitRegister Then
                    Shell GetAppRootPath & "zlPacsImageValid.exe   " & gstrServerName & "||" & gstrUserName & "||" & gstrUserPswd & "||" & mlngCur����ID & "||" & mSysPar.lngImageValid & "||" & "" & "||1", 1
                End If
            End If
'--------------------------����-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum
            Call Menu_Help_Web_Forum_click
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
            
            Select Case mstrSelTabName
                Case C_TAB_NAME_�Ŷӽк�
                    If Not mobjQueue Is Nothing Then
                        If mintChangeUserState = 2 Then  '�������û������������
                            HintMsg "��ͳһ�û����ٲ���", "cbrMain_Execute", vbInformation
                        Else
                            mobjQueue.zlExecuteCommandbar Control
                        End If
                    End If
                Case C_TAB_NAME_���ü�¼, C_TAB_NAME_���Ӳ���, C_TAB_NAME_ҽ����¼, C_TAB_NAME_������¼
                    If Not mobjWork_His Is Nothing Then
                        Call mobjWork_His.zlMenu.zlExecuteMenu(mstrSelModuleTag, Control.ID)
                    End If
                Case C_TAB_NAME_��鱨��
                    If Not mobjWork_Report Is Nothing Then
                        Call mobjWork_Report.zlMenu.zlExecuteMenu(mstrSelModuleTag, Control.ID)
                    End If
            End Select
            
    End Select
    
    If Not (Control.ID > conMenu_Manage_PacsPlugIn * 10000# And Control.ID < conMenu_Manage_PacsPlugIn * 10000# + 100) And objControl.ID <> conMenu_Manage_PacsPlugCfg Then
        '�����ִ�в���˵���������Ҫ�㲥�¼�
        Call Notify.Broadcast(BM_SYS__EVENT_MENU, 1, mobjCurStudyInfo.lngAdviceId, objControl.ID, objControl.Category)
    End If
    
    mblnMenuDownState = False
Exit Sub
errhandle:
    mblnMenuDownState = False
    mblnIsForceRefresh = False
    
    If HintError(err, "cbrMain_Execute", False) = 1 Then Resume
End Sub

Private Sub LockCapture(objStudyInfo As clsStudyInfo)
    Dim lngOldReportAdviceId As Long
    If Not mobjWork_ImageCap Is Nothing Then
        If objStudyInfo.blnMoved Or mobjCapLinker Is Nothing Then
            HintMsg "��ǰ���״̬������������", "cbrMain_Execute", vbOKOnly
            Exit Sub
        End If
        
        lngOldReportAdviceId = mobjCapLinker.ReportAdviceId
        
        mobjCapLinker.ReportAdviceId = 0    '��Ҫ��ձ���id����������ʱ������ʹ�ñ���ҽ��id��������
        mobjCapLinker.LockAdviceId = objStudyInfo.lngAdviceId
        
        Call mobjWork_ImageCap.ResetLockState(True)
        
        mobjCapLinker.ReportAdviceId = lngOldReportAdviceId
    End If
End Sub

Private Function LocateReportWindow(ByVal lngAdviceId As Long) As Boolean
'��λ����ʽ���洰��
    Dim objForm As Object
    
    LocateReportWindow = False
    
    '�ж��Ƿ�����Ѿ��򿪵ı���༭��
    For Each objForm In Forms
        If TypeOf objForm Is frmReportV2 Then
            If objForm.AdviceId = lngAdviceId And objForm.IsLinkHelper = False Then
                objForm.WindowState = 0
                objForm.Visible = True
                objForm.ZOrder
                
                Call objForm.Shake
                
                LocateReportWindow = True
                
                Exit Function
            End If
        End If
    Next
End Function

Private Sub ShowPacsInterfaceCfg()
On Error GoTo ErrorHnad
    Dim lngCount As Long
         
    If Not CheckPopedom(mstrPrivs, "������ù���") Then
        HintMsg "��û�иò�����Ȩ�ޣ�����ϵ����Ա��", "ShowPacsInterfaceCfg", vbInformation
        Exit Sub
    End If
    
    If Not ChechHaveTlbinf32 Then
        HintMsg "ϵͳ��ȱ��TLBINF32.DLL�ļ������²�����ù��ܲ�������ʹ�ã�����ϵ���������Ա���(�����������ϵͳĿ¼����Ӳ�ע��TLBINF32.DLL�ļ�)��", "ShowPacsInterfaceCfg", vbInformation
        Exit Sub
    End If
    Call frmPacsInterfaceCfg.ShowPacsInterfaceCfgV2(Me, mlngModule, mstrPrivs, mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.lngPatId)
    
    Call ReCreatCbrMenu(cbrMain)
    
    Exit Sub
ErrorHnad:
    If HintError(err, "ShowPacsInterfaceCfg", False) = 1 Then Resume
End Sub

Private Function ExecutePluginInterfaceFun(ByVal strFuncDes As String, ByVal strVBS As String, Optional ByVal lngTimeTag As Long = 0, _
    Optional ByVal strAttachPar1 As String = "", Optional ByVal strAttachPar2 As String = "", Optional ByVal strAttachPar3 As String = "") As Boolean
'blnAutoDo �Ƿ��Զ�ִ�У�Ӱ���������ʾ��Ϣ����ʽ��
'����vbs�ű�ʵ�ֹ���
    Dim i As Integer
    Dim lngStart As Long, lngEnd As Long
    Dim ary() As String
    Dim strTmpVBS As String, strParaName As String, strParaVal As String
    Dim objCall As Object
    Dim strResult As String
    
'On Error GoTo ErrorHnad
    
    ExecutePluginInterfaceFun = False
    
    ary = Split(strVBS, vbCrLf)
    
    For i = 0 To UBound(ary)
        '����Ԥ����������ڲ���ֵ
        strTmpVBS = ary(i)
        
        Do While InStr(strTmpVBS, "[[") > 0
            lngStart = InStr(strTmpVBS, "[[")
            lngEnd = InStr(strTmpVBS, "]]") + 2
            
            strParaName = Mid(strTmpVBS, lngStart, lngEnd - lngStart)
            
            Select Case strParaName
                Case "[[�������]]"
                    strParaVal = lngTimeTag
                    
                Case "[[���Ӳ���1]]"
                    strParaVal = strAttachPar1
                    
                Case "[[���Ӳ���2]]"
                    strParaVal = strAttachPar2
                    
                Case "[[���Ӳ���3]]"
                    strParaVal = strAttachPar3
                    
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
    
    strResult = ExecuteSub(strVBS)
    
    If strResult = "" Then
        ExecutePluginInterfaceFun = True
    Else
        err.Raise -1, , "��� [" & strFuncDes & "] ��������" & strResult
    End If
    
'    Exit Function
'ErrorHnad:
'    ExecutePluginInterfaceFun = False
'    err.Raise -1, , "���ִ�в�������" & err.Description
End Function

Private Function ExecuteSub(ByVal strVBS As String, Optional ByVal blnCheckVBS As Boolean = False) As String
'����vbs�ű�ʵ�ֹ���
    Dim objCall As Object
    Dim strTempVBS As String
    
On Error GoTo errhandle
    
    ExecuteSub = ""
    
    '�����ű�ִ�ж���
    Set objCall = CreateObject("ScriptControl")
    
    objCall.TimeOut = 30000
    objCall.Language = "vbscript"
    
    Call objCall.AddCode(strVBS)
    
    If blnCheckVBS Then Exit Function
    
    Call objCall.Run(Trim("ExcuteSub"))
    
    ExecuteSub = objCall.Error.Description
    
    Exit Function
errhandle:
    ExecuteSub = err.Description
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
    
    If mobjSelModule Is Nothing Then Exit Sub
    
    Call ReSetModuleFontSize(mstrSelTabName, mstrSelModuleTag, mobjSelModule, gbytFontSize)
End Sub


Private Sub ReSetModuleFontSize(ByVal strSelTabName As String, ByVal strSelModuleTag As String, _
    ByVal objSelModule As Object, ByVal bytFontSize As Byte)
'����:�������ø���ҵ��ģ�鴰��������С
'bytSizeType 0-Ĭ��С���壬1-������,
    Dim bytSizeType As Long
On Error GoTo errhandle
        
    If objSelModule Is Nothing Then Exit Sub

    Select Case strSelTabName
        Case C_TAB_NAME_Ӱ��ͼ��
            Call objSelModule.ReSetFormFontSize(gbytFontSize)
            
        Case C_TAB_NAME_Ӱ��ɼ�
            Call objSelModule.ReSetFormFontSize(gbytFontSize)
            
        Case C_TAB_NAME_��鱨��
            If strSelModuleTag = C_WORKMODULE_NAME_�ϰ汨�� Then
                Call objSelModule.ReSetFormFontSize(gbytFontSize)
            End If
            
        Case C_TAB_NAME_�걾����, C_TAB_NAME_����ȡ��, C_TAB_NAME_������Ƭ, C_TAB_NAME_�����ؼ�, C_TAB_NAME_���̱���
            Call objSelModule.ReSetFormFontSize(gbytFontSize)
            
        Case C_TAB_NAME_���ü�¼, C_TAB_NAME_������¼, C_TAB_NAME_���Ӳ���, C_TAB_NAME_ҽ����¼
            bytSizeType = IIf(bytFontSize = 9, 0, 1)
            If mlngModule = G_LNG_PATHOLSYS_NUM Then
                '����ϵͳʹ�ö�ģ�������ڲ��죬�����Ҫʶ����÷���
                If TypeOf objSelModule Is frmPatholPrice Then
                    Call objSelModule.ReSetFormFontSize(gbytFontSize)
                Else
                    Call objSelModule.SetFontSize(bytSizeType)
                End If
            Else
                Call objSelModule.SetFontSize(bytSizeType)
            End If
            
        Case Else
            
    End Select
Exit Sub
errhandle:
    If HintError(err, "ReSetFormFontSize", False) = 1 Then Resume
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
    
    Call ucPacsHelper1.SetFontSize(gbytFontSize)
    
    If gblUsePacsQuery Then
        Call mobjPacsQueryWrap.CurPacsQuery.RefreshCfgFontSize(gbytFontSize)
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
            objCtrl.Width = TextWidth("����" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = gbytFontSize
            objCtrl.Width = TextWidth("����" & objCtrl.Caption)
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
            If UCase(objCtrl.Name) = "TABWINDOW" Then
                CtlFont.Size = IIf(gbytFontSize >= 15, 13, IIf(gbytFontSize <= 10, 11, 12))
                Set objCtrl.PaintManager.Font = CtlFont
            Else
                CtlFont.Size = gbytFontSize - 1
                Set objCtrl.PaintManager.Font = CtlFont
                
                If UCase(objCtrl.Name) = "TABEXTRA" Then
                    TabExtra.Height = TabExtra.Height - 20
                End If
            End If
        Case UCase("CommandButton"), UCase("PICTUREBOX")
            If UCase(objCtrl.Name) = "PICTABFACE" Then
                objCtrl.FontName = strFontType
                objCtrl.FontSize = IIf(gbytFontSize >= 15, 13, IIf(gbytFontSize <= 10, 11, 12))
            Else
                objCtrl.FontName = strFontType
                objCtrl.FontSize = gbytFontSize
            End If

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

Private Function GetCurDeptId(Optional ByVal lngDeptId As Long = 0) As Long
    Dim blnFromAdvice As Boolean
    
    If lngDeptId <> 0 Then
        GetCurDeptId = lngDeptId
    End If
    
    blnFromAdvice = True
    
    '�ж�ҽ�������Ƿ���Ч
    If mobjCurStudyInfo Is Nothing Then
        blnFromAdvice = False
    Else
        blnFromAdvice = IIf(mobjCurStudyInfo.lngAdviceId <> 0, True, False)
    End If
    
    '�Ƿ��ҽ����ȡ����ID
    If blnFromAdvice Then
        GetCurDeptId = mobjCurStudyInfo.lngExeDepartmentId
    Else
        If mblnAllDepts Then
            GetCurDeptId = UserInfo.����ID
        Else
            GetCurDeptId = mlngCur����ID
        End If
    End If
End Function

Private Function GetCurPatientFrom() As Long
    If mobjCurStudyInfo Is Nothing Then
        GetCurPatientFrom = 0
    Else
        GetCurPatientFrom = mobjCurStudyInfo.lngPatientFrom
    End If
End Function

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
            Select Case mstrSelTabName
                Case C_TAB_NAME_ҽ����¼, C_TAB_NAME_���ü�¼
                    Call mobjWork_His.zlMenu.zlRefreshSubMenu(mstrSelModuleTag, CommandBar)
            End Select
    End Select
    Exit Sub
errhandle:
    If HintError(err, "cbrMain_InitCommandsPopup", False) = 1 Then
        Resume
    End If
End Sub

Private Function GetReportEditState(ByVal objStudyInfo As clsStudyInfo) As Long
'0-��������д��1-������д��2-�������޶���3-�����޶���4-����(�ݶ�)
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strCurModuleTag As String
    Dim lngDeptId As Long
    
    GetReportEditState = 0
    
    If objStudyInfo.lngAdviceId <= 0 Then Exit Function 'ҽ��ID��Чʱ��������༭����
    
    If mblnAllDepts Then
        lngDeptId = UserInfo.����ID
    Else
        lngDeptId = mlngCur����ID
    End If
    
    strSQL = "select ����ID, ������,������,�鵵��,���ʱ��,RawToHex(��鱨��ID) as ��鱨��ID  from ����ҽ������ a , ���Ӳ�����¼ b where a.����ID=b.Id(+) and a.ҽ��ID=[1]"
    If objStudyInfo.blnMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�������", objStudyInfo.lngAdviceId)
    
    strCurModuleTag = GetWorkModuleName(C_TAB_NAME_��鱨��, objStudyInfo.lngExeDepartmentId, objStudyInfo.lngPatientFrom)
    If strCurModuleTag = C_WORKMODULE_NAME_���ܱ��� Then
         If rsData.RecordCount > 0 Then
            GetReportEditState = IIf(NVL(rsData!��鱨��ID) <> "", 4, 0)
         Else
            GetReportEditState = 0
         End If
         
        Exit Function
    End If
    
    If rsData.RecordCount > 0 Then
        '����д����
        '�������Ѿ�ִ����ɣ���ֻ�ܽ��в鿴
        If objStudyInfo.intStep >= 6 Then
            GetReportEditState = 4
            Exit Function
        End If
        
        If NVL(rsData!���ʱ��) <> "" Then
            '��ǩ��
            '�ж��Ƿ����޶�Ȩ��
            If InStr(1, mstrPrivs, "�����޶�") > 0 Then
                If NVL(rsData!����ID) = lngDeptId Or IsContainDept(UserInfo.ID, Val(NVL(rsData!����ID))) Then
                    If NVL(rsData!�鵵��) = "" Then
                        'TASK:����������Ӷ��û�������ж�
                        GetReportEditState = 3
                    Else
                        '�ѹ鵵
                        GetReportEditState = 2
                    End If
                Else
                    GetReportEditState = 2
                End If
            Else
                '�жϱ����˺͵�ǰ�û��Ƿ���ͬ�������ͬ��˵�������һ�α��������Լ�ǩ����
                If NVL(rsData!������) = UserInfo.�û��� Then
                    GetReportEditState = 3
                Else
                    GetReportEditState = 2
                End If
            End If
        Else
            'δǩ��
            '�жϿ���id�뵱ǰ����id�Ƿ���ͬ
            If Val(NVL(rsData!����ID)) <> lngDeptId And IsContainDept(UserInfo.ID, Val(NVL(rsData!����ID))) = False Then Exit Function   '������༭
            If NVL(rsData!������) = UserInfo.���� Or InStr(1, mstrPrivs, "���˱���") > 0 Then
                GetReportEditState = 1
            End If
        End If
    Else
        'δ��д����
        '�жϱ����Ƿ����������༭
        If objStudyInfo.strReportOperation <> "" Then
            If objStudyInfo.strReportOperation <> UserInfo.���� Then 'And InStr(1, mstrPrivs, "���˱���") <= 0 Then
                '�����˱���Ȩ�ޣ��ұ����˱��������༭ʱ����������д����
                Exit Function
            End If
        End If
        
        If InStr(1, mstrPrivs, "������д") > 0 _
            And ((objStudyInfo.intStep > 1 And objStudyInfo.intStep < 6) _
                    Or (objStudyInfo.intStep = 6 And CheckPopedom(mstrPrivs, "��¼����"))) Then
            GetReportEditState = 1
        End If
        
    End If
    
End Function




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
    
    If Not mobjSelModule Is Nothing Then
        
        If Not mobjWork_Pathol Is Nothing Then
            If mobjWork_Pathol.zlMenu.zlIsModuleMenu("", Control) Then
                Call mobjWork_Pathol.zlMenu.zlUpdateMenu("", Control)
                Exit Sub
            End If
        End If
        
        Select Case mstrSelTabName
            Case C_TAB_NAME_Ӱ��ͼ��
                If mobjSelModule.zlMenu.zlIsModuleMenu(mstrSelModuleTag, Control) Then
                    Call mobjSelModule.zlMenu.zlUpdateMenu(mstrSelModuleTag, Control)
                    Exit Sub
                End If
            Case C_TAB_NAME_Ӱ��ɼ�
            
            Case C_TAB_NAME_�걾����, C_TAB_NAME_����ȡ��, C_TAB_NAME_������Ƭ, C_TAB_NAME_�����ؼ�, C_TAB_NAME_���̱���
                '�������ģ��̳е���IWorkMenu�ӿڣ�������ģ������
                If mobjSelModule.zlMenu.zlIsModuleMenu(Control) Then
                    Call mobjSelModule.zlMenu.zlUpdateMenu(Control)
                    Exit Sub
                End If
                
            Case C_TAB_NAME_ҽ����¼, C_TAB_NAME_������¼, C_TAB_NAME_���ü�¼, C_TAB_NAME_���Ӳ���
                If mobjWork_His.zlMenu.zlIsModuleMenu(mstrSelModuleTag, Control) Then
                    Call mobjWork_His.zlMenu.zlUpdateMenu(mstrSelModuleTag, Control)
                    

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
            Case C_TAB_NAME_��鱨��
                If mobjWork_Report.zlMenu.zlIsModuleMenu(mstrSelModuleTag, Control) Then
                    Call mobjWork_Report.zlMenu.zlUpdateMenu(mstrSelModuleTag, Control)
                    Exit Sub
                End If
        End Select
    End If
      
                    
    Select Case Control.ID
        Case conMenu_PacsReport_Preview 'Ԥ��
            If Not mobjWork_Report Is Nothing Then
                If mobjCurStudyInfo.lngAdviceId = mobjWork_Report.AlreadyAdviceId Then
                    Control.ID = conMenu_File_Preview + mobjWork_Report.zlMenu.zlBaseMenuID
                    Call mobjWork_Report.zlMenu.zlUpdateMenu(GetWorkModuleTag(C_TAB_NAME_��鱨��), Control)
                    
                    Control.ID = conMenu_PacsReport_Preview
                Else
                    'δ�л�������ģ�飬����ģ������û��ˢ��
                    Control.Enabled = mobjCurStudyInfo.blnCanPrint
                End If
            Else
                '
                Control.Enabled = mobjCurStudyInfo.blnCanPrint
            End If
        
        Case conMenu_PacsReport_Print   '��ӡ
            If Not mobjWork_Report Is Nothing Then
                If mobjCurStudyInfo.lngAdviceId = mobjWork_Report.AlreadyAdviceId Then
                    Control.ID = conMenu_File_Print + mobjWork_Report.zlMenu.zlBaseMenuID
                    Call mobjWork_Report.zlMenu.zlUpdateMenu(GetWorkModuleTag(C_TAB_NAME_��鱨��), Control)
                    
                    Control.ID = conMenu_PacsReport_Print
                Else
                    'δ�л�������ģ�飬����ģ������û��ˢ��
                    Control.Enabled = mobjCurStudyInfo.blnCanPrint
                End If
            Else
                '
                Control.Enabled = mobjCurStudyInfo.blnCanPrint
            End If
        
        Case conMenu_PacsReport_Write   '��д
            Select Case mobjCurStudyInfo.lngReportEditState
                Case 0
                    Control.Caption = "��д"
                    Control.Enabled = False
                Case 1
                    Control.Caption = "��д"
                    Control.Enabled = True
                Case 2
                    Control.Caption = "�޶�"
                    Control.Enabled = False
                Case 3
                    Control.Caption = "�޶�"
                    Control.Enabled = True
                Case 4
                    Control.Caption = "����"
                    Control.Enabled = True
            End Select
'            If Not mobjWork_Report Is Nothing Then
'                Select Case GetWorkModuleTag(C_TAB_NAME_��鱨��)
'                    Case C_WORKMODULE_NAME_�ϰ汨��
'                        Control.ID = conMenu_PacsReport_Open + mobjWork_Report.zlMenu.zlBaseMenuID
'                        Call mobjWork_Report.zlMenu.zlUpdateMenu(C_WORKMODULE_NAME_�ϰ汨��, Control)
'
'                        Control.ID = conMenu_PacsReport_Write
'                        Control.Visible = True
'                    Case C_WORKMODULE_NAME_��������
'                        '���Ӳ����༭��
'                        Control.ID = conMenu_Edit_Modify + mobjWork_Report.zlMenu.zlBaseMenuID
'                        Call mobjWork_Report.zlMenu.zlUpdateMenu(C_WORKMODULE_NAME_��������, Control)
'
'                        Control.ID = conMenu_PacsReport_Write
'                        Control.Visible = True
'                    Case Else
'                        Control.Visible = False
'                End Select
'            Else
'                '
'                Control.Visible = GetWorkModuleName(C_TAB_NAME_��鱨��, GetCurDeptId, GetCurPatientFrom) <> C_WORKMODULE_NAME_���ܱ��� '  mSysPar.lngReportType <> ReportType.�����ĵ��༭��
'                Control.Enabled = ((intState >= 2 And intState < 6) Or (intState >= 6 And CheckPopedom(mstrPrivs, "��¼����"))) And Not blnNoRecord
'            End If
        
        Case conMenu_Manage_LocateValue
            Control.Enabled = Not blnNoRecord
        Case comMenu_Cap_Process
            Control.Enabled = True 'Not blnNoRecord
        Case conMenu_View_Filter * 10#
            Control.Caption = " " & IIf(mblnAllDepts = True, "ȫ������", Split(mstrCur����, "-")(1)) & " "
            Control.Checked = True

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
        Case conMenu_Cap_StudySyncState
            If Not blnNoRecord Then
                Control.Enabled = (intState = 2 Or intState = 3)
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
            If Control.ID = conMenu_Manage_Disease Then
                Control.Enabled = mobjCurStudyInfo.lngAdviceId > 0
            ElseIf Control.ID = conMenu_Manage_DiseaseQuery Then
                Control.Enabled = mobjCurStudyInfo.lngAdviceId > 0
            Else
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

        Case conMenu_Manage_ReportRelease
            Control.Enabled = IIf(intState >= 4, True, False)
            
            If Not blnNoRecord Then
 
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
            Control.Enabled = Not Control.Enabled
            Control.Enabled = Not Control.Enabled
        
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
        
        Case conMenu_Img_OpenView, conMenu_img_ContrastView 'Ӱ��Ա�,Ӱ���Ƭ

            If Not mobjWork_PacsImg Is Nothing Or mstrSelTabName = C_TAB_NAME_Ӱ��ͼ�� Then
                If Control.ID = conMenu_Img_OpenView Then
                    Control.ID = conMenu_Img_Look
                    Call mobjWork_PacsImg.zlMenu.zlUpdateMenu(mstrSelModuleTag, Control)
                    Control.ID = conMenu_Img_OpenView
                    
                Else
                    Control.ID = conMenu_Img_Contrast
                    Call mobjWork_PacsImg.zlMenu.zlUpdateMenu(mstrSelModuleTag, Control)
                    Control.ID = conMenu_img_ContrastView
                    
                End If
            Else
                If blnNoRecord Then Control.Enabled = False: Exit Sub
                Control.Enabled = mobjCurStudyInfo.strStudyUID <> "" Or mSysPar.lngAutoImageDays > 0
            End If
            
        Case conMenu_Check_ViewLink
            Control.Enabled = Not blnNoRecord
            
        Case conMenu_Manage_RelatingPatiet  '��������
            If blnNoRecord Or (intState < 2 And intState <> -1) Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
            
        Case conMenu_Manage_Change_Undo
        Case conMenu_Manage_More
        Case conMenu_Manage_State
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
        
        Case conMenu_Manage_SetXWParam      '����PACS�������ã�����д˲˵�������ʾ
        Case conMenu_ReportPopup, conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99# '����
        Case conMenu_FilePopup, conMenu_ManagePopup, conMenu_ViewPopup, conMenu_HelpPopup
        Case conMenu_ToolPopup, conMenu_Tool_Valid
        Case conMenu_Help_Help, conMenu_Help_About  '����
        Case conMenu_Help_Web, conMenu_Help_Web_Forum, conMenu_Help_Web_Home, conMenu_Help_Web_Mail '����WEB
        Case conMenu_File_Exit, conMenu_EditPopup
        Case ConMenu_File_ShortcutSet
        Case conMenu_Pathol_WorkModule
        Case conMenu_View_ToolBar
        Case conMenu_Manage_Query
        Case conMenu_Manage_QueryCFG
        Case conMenu_Manage_QueryCfgUserScheme
            Control.Enabled = IIf(mlngCur����ID = 0, False, True)
        Case conMenu_Manage_QueryTabDisplayScheme
            Control.Checked = mSysPar.blnQuickTabDisplayScheme
        Case conMenu_Manage_PacsPlugIn, conMenu_Manage_PacsPlugCfg
        Case conMenu_Manage_PacsPlugIn * 10000# To conMenu_Manage_PacsPlugIn * 10000# + 100#
            '100908             Category������չΪ3��
            'strTmp:����Ƿ�����
            strTmp = IIf(UBound(Split(Control.Category, ",")) = 2, Split(Control.Category, ",")(0), Control.Category)
            Control.Enabled = Val(strTmp)
        Case conMenu_Manage_PacsPlugLevel2 * 10000# To conMenu_Manage_PacsPlugLevel2 * 10000# + 9999#
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
        Case comMenu_Collection_Type * 10000# To comMenu_Collection_Type * 10000# + 9999#  '��̬�ղز˵�
            Control.Enabled = True
        Case conMenu_Collection_ViewShare * 10000# To conMenu_Collection_ViewShare * 10000# + 9999#  '��̬����˵�
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
        Case comMenu_Petition_View, conMenu_Manage_Request
            
        Case conMenu_Manage_CustomQuery * 100# + 1# To conMenu_Manage_CustomQuery * 100# + 99#
            Control.Enabled = True

            If Control.Parameter = mobjPacsQueryWrap.SchemeNo Then
                Control.iconid = 3558
            Else
                Control.iconid = 0
            End If
            
        Case conMenu_Manage_CustomQuery * 100# + 500#
        Case C_LNG_TAB_MENU_ID
            Control.Enabled = True
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
    Exit Sub
errhandle:
    HintMsg err.Description, "cbrMain_Update", infNone
'    Resume
End Sub

Private Sub InitDeptParameter(ByVal lngDeptId As Long)
'����:��ʼ��ģ�鼶����,���������ʱ����һ��
On Error GoTo errH
    Dim rsTemp As ADODB.Recordset
    
    mSysPar.lngListColorMark = NVL(GetDeptPara(lngDeptId, "��ɫ��ʾ����", 0))
    mSysPar.blnNameColColorCfg = GetDeptPara(lngDeptId, "������ɫ����", 0) = "1"         '������ɫ����
    mSysPar.blnOrdinaryNameColColorCfg = GetDeptPara(lngDeptId, "ȱʡ���Ͳ���������ɫ����", 0) = "1"       'ȱʡ���Ͳ���������ɫ����
    mSysPar.lngAutoImageDays = Val(GetDeptPara(lngDeptId, "�Զ�����ʷͼ������", 0))
    
    If mSysPar.blnNameColColorCfg Then
        gstrSQL = "select ���� from �������� where ȱʡ��־=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡȱʡ��������")
        
        If rsTemp.RecordCount > 0 Then mstrDefaultPatientType = NVL(rsTemp!����)
    End If

    
    mSysPar.blnChangeUser = GetDeptPara(lngDeptId, "�������û�", 0) = "1"              '�������û�
    mSysPar.blnSwitchUser = GetDeptPara(lngDeptId, "�����л��û�", 0) = "1"              '�����л��û�
    
    mSysPar.blnIsPetitionScan = IIf(Val(GetDeptPara(lngDeptId, "�������뵥ɨ��", 1)) = 1, True, False)   '��ȡ�������뵥ɨ�����
    mSysPar.strImageLevel = NVL(GetDeptPara(lngDeptId, "Ӱ�������ȼ�", "��,��"))
    mSysPar.strReportLevel = NVL(GetDeptPara(lngDeptId, "���������ȼ�", "��,��"))
    mSysPar.blnֱ�Ӽ�� = (Val(GetDeptPara(lngDeptId, "�ǼǺ�ֱ�Ӽ��", 0)) = 1)         '�ǼǺ�ֱ�Ӽ��
    
    mSysPar.lngReportType = Val(GetDeptPara(lngDeptId, "����༭��", 0))                 '����༭��

'    mSysPar.lngCriticalValues = Val(GetDeptPara(lngDeptId, "Σ������ж�", 0))           'Σ������ж�
    mSysPar.blnIgnoreResult = GetDeptPara(lngDeptId, "���Խ��������", 0) = "1" '        '���Խ��������
    mSysPar.lngConformDetermine = Val(GetDeptPara(lngDeptId, "��������ж�", 0))         '��������ж�
    mSysPar.lngImageLevel = Val(GetDeptPara(lngDeptId, "Ӱ�������ж�", 0))               'Ӱ�������ж�
    mSysPar.lngReportLevel = Val(GetDeptPara(lngDeptId, "���������ж�", 0))
    
    mSysPar.lngHintType = Val(GetDeptPara(lngDeptId, "��Ͻ����ʾ����", 0))
    
    mSysPar.blnReportWithImage = GetDeptPara(lngDeptId, "��ͼ�����д����", 0) = "1" '   '��ͼ�����д����
    mSysPar.blnReportWithResult = GetDeptPara(lngDeptId, "��Ӱ�����Ϊ����", 0) = "1" '  '��Ӱ�����Ϊ����
    mSysPar.blnCompleteCommit = GetDeptPara(lngDeptId, "��˺�ֱ�����", 0) = "1" '      '��˺�ֱ�����
    mSysPar.blnFinallyCompleteCommit = GetDeptPara(lngDeptId, "�����ֱ�����", 0) = "1" '�����ֱ�����
    mSysPar.blnAuditAutoPrint = IIf(Val(GetDeptPara(lngDeptId, "�����ֱ�Ӵ�ӡ", 0)) = 1, True, False) '�����ֱ�Ӵ�ӡ
    mSysPar.blnNoSignFinish = GetDeptPara(mlngCur����ID, "����δǩ�������ӡ���", 0) = "1" '       '����δǩ�������ӡ���
    mSysPar.blnDirectSendRepImg = IIf(Val(GetDeptPara(lngDeptId, "ͬ����ӹ�Ƭ����ͼ", 1)) = 1, True, False)
    
    mSysPar.lngBeforeDays = Val(GetDeptPara(lngDeptId, "Ĭ�Ϲ�������", 2)) '                   'Ĭ�Ϲ�������
    If mSysPar.lngBeforeDays > 15 Or mSysPar.lngBeforeDays <= 0 Then
        mSysPar.lngBeforeDays = 2
    End If
    
    mSysPar.blnWriteCapDoctor = GetDeptPara(lngDeptId, "�ɼ�ͼ����Ϊ��鼼ʦ", 0) = "1"  '�ɼ�ͼ����Ϊ��鼼ʦ
    
    mSysPar.blnPrintCommit = GetDeptPara(lngDeptId, "��ӡ��ֱ�����", 0) = "1" '           '��ӡ��ֱ�����
    mSysPar.blnCanPrint = GetDeptPara(lngDeptId, "ƽ������˲��ܴ򱨸�") = "1"             'ƽ����Ҫ��˲��ܴ�ӡ =true
    mSysPar.blnAutoSendWorkList = GetDeptPara(lngDeptId, "����ʱ�Զ�����WorkList") = "1"   '����ʱ�Զ�����WorkList

    '����������
    mSysPar.blnNameFuzzySearch = GetDeptPara(lngDeptId, "����Ĭ��ģ����ѯ", "1") = "1"     '����Ĭ��ģ����ѯ
    mSysPar.blnNameQueryTimeLimit = GetDeptPara(lngDeptId, "������ѯʱ������", "1") = "1"  '����������ʱ�Ƿ����ʱ������
    
    '�Ƿ�λ����
    mSysPar.blnIsLocateReport = Val(GetDeptPara(lngDeptId, "����л�ʱ��λ����༭", "1")) = 1
    
    If CheckPopedom(mstrPrivs, "�Ŷӽк�") And mlngModule <> G_LNG_PATHSTATION_MODULE And CheckPopedom(";" & GetPrivFunc(glngSys, 1160) & ";", "����") Then      '��Ȩ��ʹ�òŸ��ݲ�������
        mSysPar.blnUseQueue = GetDeptPara(lngDeptId, "�����Ŷӽк�", 0) = "1" '          'Ĭ�ϲ������Ŷӽк�
        
        If mSysPar.blnUseQueue Then
            mSysPar.blnSynStudylist = GetDeptPara(lngDeptId, "ͬ����λ����б�", 0)
            mSysPar.blnAutoInQueue = GetDeptPara(lngDeptId, "�������Զ��Ŷ�", 1)
        End If
    Else
        mSysPar.blnUseQueue = False
    End If
    
    mSysPar.blnRelatingPatient = GetDeptPara(lngDeptId, "������������", 0) = "1"       '�Ƿ�ʹ�ù�
    
    gblnXWLog = (Val(zlDatabase.GetPara("XW��¼�ӿ���־", glngSys, G_LNG_XWPACSVIEW_MODULE, "0")) = 1) '�Ƿ��¼�ӿ���־
    
    Exit Sub
errH:
    If HintError(err, "InitDeptParameter") = 1 Then Resume
End Sub


Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
On Error GoTo errhandle
    '��ֹ����б� �϶�
    Cancel = IIf(((Action = 4 Or Action = 6 Or Action = 5) And Not Pane.hidden), True, False)
errhandle:
End Sub

Private Sub InitQueryWrapComponent()
    Dim objPar As clsQueryPar
    
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
        Set objPar.rtfHisFollow = Nothing
        Set objPar.PicHisFollow = Nothing
        Set objPar.TimerHisFunc = Nothing
        Set objPar.picTemp = picTemp
        
        Set objPar.labPatiInfo = labPatientInfo
         
        Call mobjPacsQueryWrap.Init(mlngCur����ID, UserInfo.ID, mlngModule, 0, mSysPar.blnCanPrint, mobjSquareCard, Me, objPar)
        
        mobjPacsQueryWrap.DefaultLocate = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\", "DEFLOCATE", True)
        
        cmdLocate.BackColor = IIf(mobjPacsQueryWrap.DefaultLocate, &HFF00&, &H8000000F)
        cmdFind.BackColor = IIf(mobjPacsQueryWrap.DefaultLocate = False, &HFF00&, &H8000000F)
    End If
End Sub

Private Sub InitPars()
    Dim bytFontSize As Byte
    
    Dim strTmpImgPath As String
        
    Call WriteLog("InitPars -> Step 1����ʼ��ȡ����...")
    
    '��ȡ�����С
    bytFontSize = Val(zlDatabase.GetPara("��ʾ�����С", glngSys, glngModul))
    gbytFontSize = IIf(bytFontSize = 0, 9, IIf(bytFontSize = 1, 12, 15))
    
    Call WriteLog("InitPars -> Step 2�����뱾��ע������...")
    
    Call InitLocalPars '����ע������
    
    Call WriteLog("InitPars -> Step 3������������̲���...")
    Call InitDeptParameter(mlngCur����ID)
     
    
    ReDim gConnectedShardDir(0) As String   '��ʼ������Ŀ¼���Ӵ�
     
    Call WriteLog("InitPars -> Step 3����ʼ���Զ����ѯ�������...")
     

    Call WriteLog("InitPars -> Step 5��������Ŀ¼...")
    strTmpImgPath = FormatFilePath(GetAppRootPath & "\Apply\TmpImage\")
    ClearCacheFolder strTmpImgPath     '����ʱĿ¼���ˣ�����ո�Ŀ¼
    
    '�ж���ʱĿ¼�Ƿ����
    If Dir(strTmpImgPath, vbDirectory) = "" Then
        Call MkDir(strTmpImgPath)
    End If
    
    Call WriteLog("InitPars -> Step 6����ʼ��˫�û���¼����...")
    '��ʼ��˫�û���½�Ĳ���
    mblnCnOracleIsHIS = True
    mintChangeUserState = 1
    
    mstrHisUserName = UserInfo.����
    mstrOtherUserName = UserInfo.����
    mstrHisUserID = UserInfo.�û���
    mstrOtherUserID = UserInfo.�û���
    
    Set mcnOracleHIS = gcnOracle
    
    Me.stbThis.Panels(4).Text = "����ҽ����" & mstrHisUserName & "   ���ҽ����" & mstrOtherUserName
    
    ReDim mobjPacsReportArry(0) As frmReport
    
    Call WriteLog("InitPars -> Step 7����ȡ�°��Ƭ����״̬...")
    gblnUseXinWangView = False
    
    If mlngModule = G_LNG_PACSSTATION_MODULE Then
        gblnUseXinWangView = IsUseXwViewer
    End If

    
    Call WriteLog("InitPars -> Step End������ִ��...")
End Sub


'Private Sub Form_Load()
'On Error GoTo errHandle
'    '��ʼ����ط�����showstation�е���InitForm���д���......
'    '���ﲻ�ܽ�����صĳ�ʼ����������Ϊ��clsPacsWork��BHCodeMain�����У�������ʾ��ʽ��ʱ�򣬻ᴥ��Load�¼���
'    '��Load�¼��е�ĳЩ������Ҫ��ز���������ȷִ�У������Ҫ��Load�еĴ�����������ȡ����������ShowStation������ִ��...
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'End Sub

Private Function GetWindowCaption() As String
    GetWindowCaption = Mid(Me.Caption & " ", 1, InStr(Me.Caption & " ", " "))
End Function


Private Sub DisposeObj()
    Dim i As Long
    
On Error Resume Next
    Set mobjSelModule = Nothing
    
    TimerRefresh.Enabled = False
    
    For i = 1 To UBound(mAryWorkModule)
        Set mAryWorkModule(i).objModule = Nothing
    Next
    
    If Not mobjPacsQueryWrap Is Nothing Then
        Call mobjPacsQueryWrap.Free
        Set mobjPacsQueryWrap = Nothing
    End If
    
    If Not mobjAppendBill Is Nothing Then
        Set mobjAppendBill = Nothing
    End If
    
    If Not mobjWork_PacsImg Is Nothing Then
        Unload mobjWork_PacsImg
        Set mobjWork_PacsImg = Nothing
    End If
    
    If Not mobjRichReportWrap Is Nothing Then
        Unload mobjRichReportWrap
        Set mobjRichReportWrap = Nothing
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
    
    Set mobjCapLinker = Nothing
    
    If Not gobjMsgCenter Is Nothing Then
        Set gobjMsgCenter = Nothing
    End If
    
    Erase mAryWorkModule
         
    Set mobjSquareCard = Nothing
    
    If Not mobjPublicAdvice Is Nothing Then Set mobjPublicAdvice = Nothing
    
    If err.Number <> 0 Then
        Debug.Print "frmPacsMainV2.DisposeObj Err:" & err.Description
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errhandle
    Dim i As Long
    
    TimFlicker.Enabled = False
    
    Select Case mlngModule
        Case 1290
            Call UnAttachModuleMsgProc(Me.hwnd, mtImage)
            Set gobjImageMainWindow = Nothing
        Case 1291
            Call UnAttachModuleMsgProc(Me.hwnd, mtVideo)
            Set gobjVideoMainWindow = Nothing
        Case 1294
            Call UnAttachModuleMsgProc(Me.hwnd, mtPathol)
            Set gobjPatholMainWindow = Nothing
    End Select
    
    If Not mobjWork_ImageCap Is Nothing Then
        Call mobjWork_ImageCap.zlNotifyQuit
    End If
    
    '�ر���Ϣ����
    If Not gobjMsgCenter Is Nothing Then
        Call gobjMsgCenter.CloseMsgCenter
    End If
 
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\", "�б�����Ϣ�߶�����", mlngMove)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\", "���ؼ���б�", dkpMain.Panes(1).hidden)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\", "���ظ���ģ��", dkpMain.Panes(2).hidden)
    
'    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsList), vsList.Name, mstrCol)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    
    If Me.ScaleWidth > 0 Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "ListWidth", picList.Width / Me.ScaleWidth)
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "HelperWidth", ucPacsHelper1.Width / Me.ScaleWidth)
    End If
    
    '���������С
    zlDatabase.SetPara "��ʾ�����С", IIf(gbytFontSize = 9, 0, IIf(gbytFontSize = 12, 1, IIf(gbytFontSize = 15, 2, gbytFontSize))), glngSys, glngModul
    
    '�ָ���������
    Me.Caption = GetWindowCaption
    
    '����ucpacsHelper������
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\", "HELPER", ucPacsHelper1.GetLayoutStr)
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\", "DEFLOCATE", mobjPacsQueryWrap.DefaultLocate)
    
    Call SaveWinState(Me, App.ProductName)
    
    Call ucPacsHelper1.Destory
    
    Call ResetNullParent
    
    Call dkpMain.CloseAll
    
    Call DisposeObj
    
    '�ָ�����̨�����ݿ�����
    If mblnCnOracleIsHIS = False Then
        Set gcnOracle = mcnOracleHIS
        InitCommon gcnOracle
'        RegCheck
        SetDbUser mstrHisUserID
        Call GetUserInfo
'        Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
    End If
    
    frmTwoUser.intDBState = 1
    
    mblnFormLoadState = False
    mblnIsValid = False
    
    
    Set mobjCurStudyInfo = Nothing
    Set mcnOracleHIS = Nothing
    Set mobjMedicalRecord = Nothing
    Set mfrmRISRequest = Nothing
    Set mobjMsgCenter = Nothing
    Set gobjEvent = Nothing
    
    
    
    '�Ƴ�DataExchangeĿ¼�ļ�
    If DirExists(GetTempImgPath() & "DataExchange\") Then
        Call DeleteFolder(GetTempImgPath() & "DataExchange\", , False)
    End If
    
    Exit Sub
errhandle:
    HintError err, "Form_Unload", False
End Sub

Private Sub ResetNullParent()
    Dim i As Long
    
On Error GoTo errhandle
    For i = 1 To UBound(mAryWorkModule)
        If Not mAryWorkModule(i).objModule Is Nothing Then
            If mAryWorkModule(i).hwnd <> 0 Then
                ShowWindow mAryWorkModule(i).hwnd, 0
                SetParent mAryWorkModule(i).hwnd, 0
            End If
        End If
    Next
Exit Sub
errhandle:
    Debug.Print "frmPacsMainV2.ResetNullParent Err:" & err.Description
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
    
    mlngMove = Val(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\", "�б�����Ϣ�߶�����", 0))
    mblnIsHideStudyList = CBool(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\", "���ؼ���б�", 0))
    mblnIsHideHelper = CBool(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\", "���ظ���ģ��", 0))
    
errContinue2:
    mSysPar.blnLockAfterCall = zlDatabase.GetPara("���к������ɼ�", glngSys, mlngModule, "0")
    mSysPar.strFirstTab = zlDatabase.GetPara("������ҳ", glngSys, mlngModule, "") 'Ϊ�ձ�ʾ��ʹ�ö��ƹ�����ҳ����
    mSysPar.blnAutoOpenReport = (Val(zlDatabase.GetPara("��ʼ����Զ��򿪱���", glngSys, mlngModule, 0)) = 1)
    mSysPar.blnChoosePrintFormat = (Val(zlDatabase.GetPara("������ӡǰѡ���ʽ", glngSys, mlngModule, 0)) = 1)
    mSysPar.strLocalRoom = zlDatabase.GetPara("����ִ�м�����", glngSys, mlngModule, "")
    mSysPar.blnQueueQuick = IIf(Val(zlDatabase.GetPara("�Զ�������ݺ��д���", glngSys, mlngModule, "1")) = 1, True, False)
    mSysPar.lngImageValid = Val(zlDatabase.GetPara("ͼ��У��", glngSys, mlngModule, 0))
    
    mSysPar.blnAutoPrint = Val(zlDatabase.GetPara("�������Զ���ӡ���뵥", glngSys, mlngModule, 0)) '�������Զ���ӡ���뵥
    mSysPar.blnAutoPrintCheck = Val(zlDatabase.GetPara("�Զ�����ظ������ӡ", glngSys, mlngModule, 0))
    
    If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
        '����ǲɼ�ģ�飬����Ҫִ�иò���
        mSysPar.lngVideoStationMoneyExeModle = Val(zlDatabase.GetPara("�ɼ�����ִ��ģʽ", glngSys, mlngModule, 0))
    ElseIf mlngModule = G_LNG_PACSSTATION_MODULE Then
        mSysPar.lngPacsStationMoneyExeModle = Val(zlDatabase.GetPara("ҽ������ִ��ģʽ", glngSys, mlngModule, 0))
    Else
        mSysPar.lngPatholStationMoneyExeModle = Val(zlDatabase.GetPara("�������ִ��ģʽ", glngSys, mlngModule, 0))
    End If
    
    '����ʱ��Ƭ
    mSysPar.blnShowImgAfterReport = (Val(zlDatabase.GetPara("����ʱ��Ƭ", glngSys, mlngModule, 0)) = 1)
    
    '��첡�����ʱ���жϷ���
    mSysPar.blnPEISNoCheckMoneyFinish = (Val(zlDatabase.GetPara("��첡�����ʱ���жϷ���", glngSys, mlngModule, 0)) = 1)

    '��ʾ���÷�����ǩ
    mSysPar.blnQuickTabDisplayScheme = Val(zlDatabase.GetPara("��ʾ���÷�����ǩ", glngSys, mlngModule, 0)) = 1
    
    '�õ�ע����й��ڹ�������ʾ״̬��ֵ�����Ϊ�������9
    mintToolBarWriteReg = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\CommandBars", "cbrMainButtonText", 9))
    
End Sub

Private Function InitDepts() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
On Error GoTo errH
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str����IDs As String, str��Դ As String
    
    mlngCur����ID = 0
    mstrCur���� = ""
    mstrCanUse���� = ""
    mblnAllDepts = False
    
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
        HintMsg "û�з���ҽ��������Ϣ,���ȵ����Ź��������á�", "InitDepts", vbInformation
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
            HintMsg "û�з�������������,����ʹ�ô˹���վ��", "InitDepts", vbInformation
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
    If HintError(err, "InitDepts") = 1 Then Resume
End Function

Private Sub InitLayout()
    Dim dblListWidth As Double
    Dim dblHelperWidth As Double
    
    '��ʼ���沼��
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane
    With Me.dkpMain
        .SetCommandBars cbrMain
        
        .options.HideClient = True
        .options.UseSplitterTracker = False 'ʵʱ�϶�
        .options.ThemedFloatingFrames = True
        .options.AlphaDockingContext = True
    End With
    
'    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    
    dblListWidth = NVL(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "ListWidth", 0.35))
    dblHelperWidth = NVL(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "HelperWidth", 0.25))
    
    If dblListWidth >= 0.7 Or dblListWidth <= 0.05 Then dblListWidth = 0.35
    If dblHelperWidth >= 0.7 Or dblHelperWidth <= 0.05 Then dblHelperWidth = 0.25
    
    'ע����б���Ľ��沼��Pnae�������ԣ������Ĭ�ϵ�Pane����
    If dkpMain.PanesCount <> 3 Then
        dkpMain.DestroyAll
 
        Set Pane1 = dkpMain.CreatePane(1, dblListWidth * 1024, 0, DockLeftOf, Nothing)
        Pane1.title = "����б�"
        Pane1.Handle = picList.hwnd
        Pane1.options = PaneNoCloseable Or PaneNoFloatable
        
        Set Pane2 = dkpMain.CreatePane(2, dblHelperWidth * 1024, 0, DockRightOf, Nothing)
        Pane2.title = "��������"
        Pane2.Handle = picHelper.hwnd
        Pane2.options = PaneNoCaption Or PaneNoCloseable
        
        Set Pane3 = dkpMain.CreatePane(3, (1 - dblListWidth - dblHelperWidth) * 1024, 0, DockRightOf, Pane2)
        Pane3.title = "�Ӵ���"
        Pane3.Handle = picWindow.hwnd
        Pane3.options = PaneNoCaption Or PaneNoCloseable

    End If
    
    If mblnIsHideStudyList Then Call dkpMain.Panes(1).Hide
    If mblnIsHideHelper Then Call dkpMain.Panes(2).Hide
End Sub

Public Sub StyleChange(ByVal lngStyle As TColorStyle)
'��ʽ�ı�
    Dim lngMainColor As Long
    
    Select Case lngStyle
        Case sBlue '��ɫ��ʽ
            lngMainColor = &HFFE8D9
            
            Call CommandBarsGlobalSettings.ColorManager.SetColor(STDCOLOR_BACKGROUND, &HFFE8D9)
            Call CommandBarsGlobalSettings.ColorManager.SetColor(STDCOLOR_BTNFACE, &HFFE8D9)
            Call CommandBarsGlobalSettings.ColorManager.SetColor(STDCOLOR_MENU, &HFFE8D9)
            Call CommandBarsGlobalSettings.ColorManager.SetColor(XPCOLOR_DISABLED, &H8F9296)
            
            picTabFace.BackColor = &HFFE8D9
            
            vsfList.BackColorFixed = &HF9D3B3
            vsfList.BackColor = &HFFE0CC   '&HFFDAC1 '&H00FFD5B9&
            vsfList.ForeColorFixed = &H80000008
            vsfList.GridColor = &HFFECDF  ' &HFFFFFF
            vsfList.BackColorBkg = &HFFFFFF
        
            TabWindow.PaintManager.Appearance = xtpTabAppearancePropertyPageFlat ' xtpTabAppearancePropertyPage2003
            TabWindow.PaintManager.Color = xtpTabColorOffice2003
            TabWindow.PaintManager.ColorSet.ButtonSelected = &HFFC0C0
            TabWindow.PaintManager.ColorSet.ButtonNormal = lngMainColor   '&HE0E0E0
            TabWindow.PaintManager.ColorSet.HeaderFaceDark = &HFFE8D9 '&HFFFFFF
            TabWindow.PaintManager.ColorSet.HeaderFaceLight = &HFFE8D9
        Case sGray '��ɫ��ʽ
            lngMainColor = &HC0C0C0
            Call CommandBarsGlobalSettings.ColorManager.SetColor(STDCOLOR_BTNFACE, &HA5A5A5)
            Call CommandBarsGlobalSettings.ColorManager.SetColor(STDCOLOR_BACKGROUND, &HA5A5A5)
            Call CommandBarsGlobalSettings.ColorManager.SetColor(STDCOLOR_MENU, &HA5A5A5)
            Call CommandBarsGlobalSettings.ColorManager.SetColor(XPCOLOR_DISABLED, &H8F9296)
            
            picTabFace.BackColor = &HBDBDBD
            
            vsfList.BackColorFixed = &HC0C0C0
            vsfList.BackColor = &HFFFFFF
            vsfList.ForeColorFixed = &H80000008
            vsfList.GridColor = &HE0E0E0
            vsfList.BackColorBkg = &HFFFFFF
        
            TabWindow.PaintManager.Appearance = xtpTabAppearancePropertyPageFlat 'xtpTabAppearancePropertyPage2003
            TabWindow.PaintManager.Color = xtpTabColorOffice2003
            TabWindow.PaintManager.ColorSet.ButtonSelected = &HFFC0C0
            TabWindow.PaintManager.ColorSet.ButtonNormal = &HE0E0E0
            TabWindow.PaintManager.ColorSet.HeaderFaceDark = &HBDBDBD          '&HFFFFFF
            TabWindow.PaintManager.ColorSet.HeaderFaceLight = &HBDBDBD
        Case sAshen '�Ұ���ʽ
            lngMainColor = &HE0E0E0
            Call CommandBarsGlobalSettings.ColorManager.SetColor(STDCOLOR_BTNFACE, &HE0E0E0)
            Call CommandBarsGlobalSettings.ColorManager.SetColor(STDCOLOR_BACKGROUND, &HE0E0E0)
            Call CommandBarsGlobalSettings.ColorManager.SetColor(STDCOLOR_MENU, &HE0E0E0)
            Call CommandBarsGlobalSettings.ColorManager.SetColor(XPCOLOR_DISABLED, &H8F9296)
            
            picTabFace.BackColor = &HE6E8EA
            
            vsfList.BackColorFixed = &HE0E0E0
            vsfList.BackColor = &HFFFFFF
            vsfList.ForeColorFixed = &H80000008
            vsfList.GridColor = &HE0E0E0
            vsfList.BackColorBkg = &HC0C0C0    ' &H808080
            
            dkpMain.PanelPaintManager.ColorSet.ControlFace = &HE0E0E0
        
        
            TabWindow.PaintManager.Appearance = xtpTabAppearancePropertyPageFlat ' xtpTabAppearancePropertyPage2003
            TabWindow.PaintManager.Color = xtpTabColorOffice2003
            TabWindow.PaintManager.ColorSet.ButtonSelected = &HFFC0C0
            TabWindow.PaintManager.ColorSet.ButtonNormal = &HE0E0E0
            TabWindow.PaintManager.ColorSet.HeaderFaceDark = &HF4F5F7      '&HE6E8EA
            TabWindow.PaintManager.ColorSet.HeaderFaceLight = &HF4F5F7     ' &HE6E8EA
'            TabWindow.PaintManager.ColorSet.ControlFace = &HE6E8EA
        Case Else
            lngMainColor = &H404040
            Call CommandBarsGlobalSettings.ColorManager.SetColor(STDCOLOR_BTNFACE, &H404040)
            Call CommandBarsGlobalSettings.ColorManager.SetColor(STDCOLOR_BACKGROUND, &H404040)
            Call CommandBarsGlobalSettings.ColorManager.SetColor(STDCOLOR_MENU, &H404040)
            Call CommandBarsGlobalSettings.ColorManager.SetColor(XPCOLOR_DISABLED, &H8F9296)
            
            picTabFace.BackColor = &H747474
            
            vsfList.BackColorFixed = &H404040
            vsfList.BackColor = &H808080
            vsfList.ForeColorFixed = &HFFFFFF
            vsfList.GridColor = &H979797
        
            TabWindow.PaintManager.Appearance = xtpTabAppearancePropertyPageFlat 'xtpTabAppearancePropertyPage2003
            TabWindow.PaintManager.Color = xtpTabColorOffice2003
            TabWindow.PaintManager.ColorSet.HeaderFaceDark = &H747474
            TabWindow.PaintManager.ColorSet.HeaderFaceLight = &H747474
            TabWindow.PaintManager.ColorSet.ButtonSelected = &HC0C0C0
            TabWindow.PaintManager.ColorSet.ButtonNormal = &H808080
    End Select
    
    TabExtra.PaintManager.Appearance = xtpTabAppearancePropertyPage2003 'TabWindow.PaintManager.Appearance
    TabExtra.PaintManager.Color = TabWindow.PaintManager.Color
    TabExtra.PaintManager.ColorSet.HeaderFaceDark = lngMainColor
    TabExtra.PaintManager.ColorSet.HeaderFaceLight = lngMainColor
    TabExtra.PaintManager.ColorSet.ButtonSelected = TabWindow.PaintManager.ColorSet.ButtonSelected
    TabExtra.PaintManager.ColorSet.ButtonNormal = TabWindow.PaintManager.ColorSet.ButtonNormal
            
    Me.BackColor = lngMainColor
     
    picWindow.BackColor = lngMainColor
    picExtra.BackColor = lngMainColor
    rtxtAppend.BackColor = lngMainColor
    picDataSearch.BackColor = lngMainColor
    picDataSearchContainer.BackColor = lngMainColor
    cmdDo.BackColor = lngMainColor
    cmdClear.BackColor = lngMainColor
    cmdMore.BackColor = lngMainColor
    pic�������ڵ�.BackColor = lngMainColor
    picDetail.BackColor = lngMainColor
    picFilter.BackColor = lngMainColor
    PicFucs.BackColor = lngMainColor
    cmdFind.BackColor = lngMainColor
    cmdLocate.BackColor = lngMainColor
End Sub


Private Sub InitCommandBars()
    '���ܴ���������
On Error GoTo errH
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrPopupControl As CommandBarControl
    Dim objCusControl As CommandBarControlCustom
    
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim str3DFuncs() As String
    Dim blnShowCaption As Boolean
    
    Dim rsCollection As ADODB.Recordset
    Dim rsViewShare As ADODB.Recordset
    Dim rsShareCount As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    
    Dim i As Integer
    Dim i3DFunc As Integer
    Dim intTxtLen As Integer
    
    
    mblnMenuDownState = False
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbrMain.VisualTheme = xtpThemeWhidbey
    
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
    
    '�˵�����
    'Begin------------------------�ļ��˵�--------------------------------------Ĭ�Ͽɼ�
    Me.cbrMain.ActiveMenuBar.title = "�˵�"
    Set cbrMenuBar = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_FilePopup, "�ļ�", "", 0, False)
    With cbrMenuBar.CommandBar
        
        Set cbrPopupControl = CreateMenu(.Controls, xtpControlPopup, conMenu_Manage_SetXWParam, "����", "", 181, False)
        
        Set cbrControl = CreateMenu(cbrPopupControl.CommandBar.Controls, xtpControlButton, conMenu_File_PrintSet, "��ӡ����", "", 181, True)
        Set cbrControl = CreateMenu(cbrPopupControl.CommandBar.Controls, xtpControlButton, conMenu_File_Parameter, "��������", "", 181, False)
        Set cbrControl = CreateMenu(cbrPopupControl.CommandBar.Controls, xtpControlButton, ConMenu_File_ShortcutSet, "��ݼ�����", "", 181, False)
        Set cbrControl = CreateMenu(cbrPopupControl.CommandBar.Controls, xtpControlButton, conMenu_Pathol_WorkModule, "վ��ģʽ����", "", 9004, False)
        Set cbrControl = CreateMenu(cbrPopupControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsPlugCfg, "�������", "", 181, False)
        
        '������Ƶ�ɼ����ò˵�
        If mlngModule <> G_LNG_PACSSTATION_MODULE Then
            Set cbrControl = CreateMenu(cbrPopupControl.CommandBar.Controls, xtpControlButton, conMenu_Cap_DevSet, "��Ƶ����", "��Ƶ����", 815, False)
        End If
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_File_Excel, "�嵥��ӡ", "", 103, True)
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_PacsReport_Preview, "Ԥ��", "", 102, True)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_PacsReport_Print, "��ӡ", "", 103, False)
        
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_SwitchUser, "�л��û�", "�л��û�", 3012, True)
        
        If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
            '�����û������˵�,��Ӱ��ɼ�ϵͳ�д˹���
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_ChangeUser, "�����û�", "�������ҽ���ͱ���ҽ��", 3012, False)
        End If
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_File_SendImg, "����ͼ��", "", 3061, True)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_Change_In, "�����б�", "", 0, False)
        
    
        'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
        Set cbrPopupControl = CreateMenu(.Controls, xtpControlPopup, conMenu_HelpPopup, "����", "", 0, True)
        With cbrPopupControl.CommandBar
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Help_Help, "��������", "", 0, False)
            Set cbrControl = CreateMenu(.Controls, xtpControlButtonPopup, conMenu_Help_Web, "WEB�ϵ�����", "", 0, False)
                With cbrControl.CommandBar
                    Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_Help_Web_Forum, "������̳", "", 0, False)
                    Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_Help_Web_Home, "������ҳ", "", 0, False)
                    Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���", "", 0, False)
                End With
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Help_About, "���ڡ�", "", 0, True)
        End With
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_File_Exit, "�˳�", "", 191, True)
    End With



'Begin----------------------���˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_ManagePopup, "���", "", 0, False)
    With cbrMenuBar.CommandBar
    
        Set cbrControl = CreateMenu(.Controls, xtpControlPopup, conMenu_Manage_Request, "���뵥", "���뵥", 0, False)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButtonPopup, conMenu_Manage_RequestPrint, "��ӡ���뵥��", "", 0, False)
        
            '����������뵥ɨ����� ��ѡ������ء�ɨ�����뵥���˵���δ��ѡ�� ������
            If mSysPar.blnIsPetitionScan Then
                Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, comMenu_Petition_Capture, "ɨ�����뵥", "", 5020, , False)
                Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, comMenu_Petition_View, "�鿴���뵥", "�鿴��ɨ������뵥ͼ��", 3935, True)
            End If
            
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_CheckList, "�鿴��������", "�鿴�������뵥", 8044, False)
        
        If InStr(mstrWorkModule, ";��鱨��;") >= 1 Then
            '�м�鱨��ģ�飬���ܽ�����д
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_PacsReport_Write, "��д", "", 2607, True)
        End If
        
        If mlngModule <> G_LNG_PACSSTATION_MODULE Then
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Tool_Analyse, "�߼�����", "�߼�ͼ����", 0, True)
        Else
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Img_OpenView, "��Ƭ", "��Ӱ��ͼ��", 8111, True)
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_img_ContrastView, "�Ա�", "�Ա�Ӱ��ͼ��", 8112, False)
        End If
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Check_ViewLink, "�鿴�������", "�鿴�������", 102, False)
        
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_Regist, "���Ǽ�", "", 2110, True)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_CopyCheck, "���ƵǼ�", "", 0, False)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_ReGet, "�ٻ�ȡ��", "", 0, False)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_Receive, "��鱨��", "", 744, False)
        If mlngModule = G_LNG_VIDEOSTATION_MODULE Then  'ֻ��Ӱ��ɼ�ϵͳ��Ҫ��������
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Cap_StudySyncState, "����", "�ɼ�����", 6884, False)
        End If
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_InQueue, "���", "��ʼ�Ŷ�", 3534, True)
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_Schedule, "���ԤԼ", "", 0, False)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_ScheduleManage, "ԤԼ����", "", 0, False)
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_Transfer, "����Ӱ��", "", 505, True)
                
        If mlngModule = G_LNG_PACSSTATION_MODULE Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_SendArrange, "���Ͱ���", "", 232, False)
        End If
        
        '�����
        Set cbrControl = CreateMenu(.Controls, xtpControlPopup, conMenu_Manage_SendAudit, "�������", "���͵������", 0, False)
        Call CreateAuditorMenu(cbrControl)
        
'        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_LookRelatetion, "�鿴�������", "", , False)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_LookMecRecord, "��������", "", 102, False)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_ReportExecutor, "����ִ��", "ָ����ǰ����ļ�¼��", 5008, True)
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_Complete, "������", "", 225, False, , False)
        
        Set cbrControl = CreateMenu(.Controls, xtpControlPopup, conMenu_Manage_Change_Undo, "��������", "��������", 0, True)
        If Not cbrControl Is Nothing Then
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Redo, "ȡ���Ǽ�", "", 742, False)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Logout, "ȡ������", "", 743, False)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Undone, "ȡ�����", "", 2615, False)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Cancel, "ȡ������", "", 506, False)
        End If
        
        Set cbrControl = CreateMenu(.Controls, xtpControlPopup, conMenu_Manage_State, "�����", "�����", 0, True)

            If mlngModule = G_LNG_PACSSTATION_MODULE Then
                Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlPopup, conMenu_Manage_Release, "���Ŵ���", "�����Ƭ���Ŵ���", 3013, False)
                If Not cbrPopControl Is Nothing Then
                    Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "���淢��", "", 8215, False)
                    Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FilmRelease, "��Ƭ����", "", 8216, False)
                End If
            Else
                Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "���淢��", "", 8215, False)
            End If
            '�����
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlPopup, conMenu_Manage_Result, "������", "", 0, False)
            If Not cbrPopControl Is Nothing Then
                Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Negative, "�������", "", 3506, False)
                Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Positive, "�������", "", 3507, False)
            End If
            '�������
            If mlngModule <> G_LNG_PATHOLSYS_NUM Then
                Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlPopup, conMenu_Manage_FuHeLevel, "�������", "", 0, False)
                If Not cbrPopControl Is Nothing Then
                    Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FuHe, "����", "", 3587, False)
                    Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_JiBenFuHe, "��������", "", 3010, False)
                    Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_BuFuHe, "������", "", 3010, False)
                End If
            End If
                
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlPopup, conMenu_Manage_GChannel, "��ɫͨ��", "", 0, False, , False)
            If Not cbrPopControl Is Nothing Then
                Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_GChannelOk, "���", "", 0, False, , False)
                Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_GChannelCancel, "ȡ��", "", 0, False, , False)
            End If
        
        
        Set cbrControl = CreateMenu(.Controls, xtpControlPopup, conMenu_Manage_More, "�������", "�������", 0, True)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ThingModi, "�޸���Ϣ", "", 0, False)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ModifBaseInfo, "������Ϣ����", "", 4113, False)
            
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ExecOnePart, "�ֲ�λִ��", "�ֲ�λִ�к�ȡ��ҽ��", 0, True)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Review, "������Ϣ", "", 232, False)
    
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_DiseaseRegist, "��Ⱦ���Ǽ�", "��Ⱦ���Ǽ�", 3564, True)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_DiseaseQuery, "��Ⱦ����ѯ", "��Ⱦ����ѯ", 102, False)
            
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsCriticalReg, "Σ�����ߵǼ�", "", 8344, True)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsCriticalManage, "Σ�����߹���", "", 8345, False)
        
        
            If Not (mobjAppendBill Is Nothing) And GetInsidePrivs(pҽ�����ѹ���, True) <> "" Then
                Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_AttachMoney, "���ӷ���", "", 3011, True)
                
                If glngModul = G_LNG_PATHSTATION_MODULE Then
                    Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_CompleteAttach, "��ɲ���", "", 3816, False)
                End If
            End If
        

            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_RelatingPatiet, "��������", "", 803, True)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Burn, "ͼ���¼", "", 0, True)
    
    End With
    
    'Begin-------------------------------------------------------�ղز˵�(Ĭ�Ͽɼ�)----------------------------------------------------------

    gstrSQL = "select a.ID ,a.�ϼ�id,b.���� as ������,a.�ղ���� " & _
                " from Ӱ���ղ���� a,��Ա�� b " & _
                " where a.������ID=" & UserInfo.ID & " and a.������id=b.ID(+) Start With a.�ϼ�id Is Null Connect By Prior a.ID = a.�ϼ�id"
    Set rsCollection = zlDatabase.OpenSQLRecord(gstrSQL, GetWindowCaption)

    gstrSQL = "select a.ID ,a.�ϼ�id,b.���� as ������,a.�ղ����,a.�Ƿ��� " & _
                " from Ӱ���ղ���� a,��Ա�� b " & _
                " where a.������ID<>" & UserInfo.ID & " and a.������id=b.ID(+) Start With a.�ϼ�id Is Null Connect By Prior a.ID = a.�ϼ�id"
    Set rsViewShare = zlDatabase.OpenSQLRecord(gstrSQL, GetWindowCaption)
        
    Set cbrMenuBar = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_Collection, "�ղ�", "", 0, False)
    With cbrMenuBar.CommandBar
        
        '��¡���� ɸѡ����������ݽ����ж�
        Set rsShareCount = zlDatabase.CopyNewRec(rsViewShare)
        rsShareCount.Filter = "�Ƿ���=1"
        
        If rsShareCount.RecordCount <> 0 Then
           '�ݹ鴴������˵�
           mlngShareFatherID = 0
           Set rsTemp = zlDatabase.CopyNewRec(rsViewShare)
           rsViewShare.Filter = "�ϼ�ID=" & NVL(rsViewShare!�ϼ�ID, 1) & " and ������<> '" & UserInfo.���� & "'"
           
           Set cbrControl = CreateMenu(.Controls, xtpControlButtonPopup, conMenu_Collection_ViewShare, "����鿴", "", 0, False)
           Call RecursionCreateShareMenu(rsViewShare, rsTemp, cbrControl)
        End If

        If rsCollection.RecordCount > 0 Then
            '�ݹ鴴���ղ����˵�
                 mlngCollectionFatherID = 0
                 Set rsTemp = zlDatabase.CopyNewRec(rsCollection)
                 rsCollection.Filter = "�ϼ�ID=" & NVL(rsCollection!�ϼ�ID, 1)
                 Call RecursionCreateCollectionMenu(rsCollection, rsTemp, cbrMenuBar)
        End If
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Collection_To, "�ղص�...", "", 0, True)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Collection_Manage, "�ղع���", "", 0, False)
        
    End With
    
    'Begin----------------------�Զ����ѯ�˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_Manage_Query, "��ѯ", "", 0, False)
    
    Call mobjPacsQueryWrap.RefreshCustomQueryMenu(cbrMenuBar, mintQueryState, tabScheme, mSysPar.blnQuickTabDisplayScheme)
    
    Call CheckHaveScheme(False, "")
    
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_QueryCFG, "��ѯ����", "", 0, True)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_QueryCfgUserScheme, "���÷�������", "", 0, False)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_QueryTabDisplayScheme, "��ʾ���÷�����ǩ", "", 0, True)
        cbrControl.Checked = mSysPar.blnQuickTabDisplayScheme
        cbrControl.CloseSubMenuOnClick = False
    End With
    
    
    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
    Set cbrMenuBar = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_ReportPopup, "����", "", 0, True)
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
    
    
    'Begin----------------------�鿴�˵�--------------------------------------
    Set cbrMenuBar = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_ViewPopup, "�鿴", "", 0, False)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateMenu(.Controls, xtpControlPopup, conMenu_View_ToolBar, "������", "", 0, False)
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar '�����˵�
                Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť", "", 0, False): cbrPopControl.Checked = True
                Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ", "", 0, False): cbrPopControl.Checked = True
                Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��", "", 0, False): cbrPopControl.Checked = True
            End With
            
        Set cbrControl = CreateMenu(.Controls, xtpControlButtonPopup, conMenu_View_FontSize, "�����С", "", 0, False)
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar '�����˵�
                Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_View_FontSize_S, "С����", "", 0, False): cbrPopControl.Checked = True
                Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_View_FontSize_M, "������", "", 0, False)
                Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_View_FontSize_L, "������", "", 0, False)
            End With
            
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_View_StatusBar, "״̬��", "", 0, True): cbrControl.Checked = True
        Set cbrControl = CreateMenu(.Controls, xtpControlButtonPopup, conMenu_View_Filter * 10#, "������", "", 0, False)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_View_Refresh, "ˢ��", "", 0, False)
    End With
        
    'Begin----------------------���߲˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_ToolPopup, "����", "", 0, False)
    With cbrMenuBar.CommandBar
        'Begin----------------------���������ܲ���˵�--------------------------------------Ĭ�Ͽɼ�
        Call RefreshCustomPlugInMenu(cbrMenuBar, mlngModule)
    
        '�������߲˵�
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Tool_Valid, "ͼ��У�Թ���", "", 0, True)
    End With
        
    '���ұ���ʾ�����ɼ���ť
    If mlngModule <> G_LNG_PACSSTATION_MODULE And InStr(mstrWorkModule, C_TAB_NAME_Ӱ��ɼ�) > 0 Then
        Set cbrControl = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlButton, comMenu_Cap_Process, "�����ɼ�", "���������ɼ�����", 0, False): cbrControl.flags = xtpFlagRightAlign
    End If
        
    '---------------------�������Ͻǵ�ǰ����----------------------------------
    Set cbrControl = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_View_Filter * 10#, "������", "", 0, False): cbrControl.flags = xtpFlagRightAlign
            
            
    Set objCusControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlCustom, C_LNG_TAB_MENU_ID, "TAB���")
        objCusControl.Handle = picTabFace.hwnd
        objCusControl.flags = xtpFlagControlStretched
        objCusControl.Category = M_STR_MODULE_MENU_TAG
        
        
    '---------------------����������------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True

    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Regist, "�Ǽ�", "���Ǽ�", 211, True)
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Receive, "����", "��鱨��", 744, False)
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Logout, "ȡ��", "ȡ������", 743, False)
    If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Cap_StudySyncState, "����", "�ɼ�����", 6884, False)
    End If
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Schedule, "ԤԼ", "���ԤԼ", 6823, True)
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_InQueue, "���", "��ʼ�Ŷ�", 3534, False)
    
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_PacsReport_Preview, "Ԥ��", "����Ԥ��", 102, True)
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_PacsReport_Print, "��ӡ", "�����ӡ", 103, False)
    
    If InStr(mstrWorkModule, ";��鱨��;") >= 1 Then
        '�м�鱨��ģ�飬���ܽ�����д
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_PacsReport_Write, "��д", "", 2607, True)
    End If
    

    If mlngModule <> G_LNG_PACSSTATION_MODULE Then
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Tool_Analyse, "�߼�", "�߼�ͼ����", 0, True)
    Else
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Img_OpenView, "��Ƭ", "��Ӱ��ͼ��", 8111, True)
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_img_ContrastView, "�Ա�", "�Ա�Ӱ��ͼ��", 8112, False)
    End If
'    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Check_ViewLink, "�鿴����", "", 102, False): cbrControl.ToolTipText = "�鿴�������"
    
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_View_Filter, "����", "����", 0, True)
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_View_Refresh, "ˢ��", "ˢ��", 0, False)
        
    Call AddPlugInToolBarMenu(cbrToolBar.Controls, mlngModule)  '100908
    
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Review, "��ע", "������Ϣ", 232, True)
    
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_Request, "���뵥", "", 3935, False)
    With cbrControl.CommandBar
        If mSysPar.blnIsPetitionScan Then   '���������뵥ɨ�����ܽ��в鿴
            Call CreateMenu(.Controls, xtpControlButton, comMenu_Petition_View, "�鿴���뵥", "�鿴��ɨ������뵥ͼ��", 3935, False)
        End If
        
        Call CreateMenu(.Controls, xtpControlButton, conMenu_Manage_CheckList, "�鿴��������", "�鿴�������뵥", 8044, False)
    End With
    
    If Not (mobjAppendBill Is Nothing) And GetInsidePrivs(pҽ�����ѹ���, True) <> "" Then
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_AttachMoney, "������", "������", 3011, False)
        If glngModul = G_LNG_PATHSTATION_MODULE Then
            Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_CompleteAttach, "��ɲ���", "��ɲ���", 3816, False)
        End If
    End If
    
'    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_Disease, "��Ⱦ��", "��Ⱦ��", 3842, False)
'    If Not cbrControl Is Nothing Then
'        Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_DiseaseRegist, "��Ⱦ���Ǽ�", "��Ⱦ���Ǽ�", 3564, False)
'        Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_DiseaseQuery, "��Ⱦ����ѯ", "��Ⱦ����ѯ", 102, False)
'    End If
    
'    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_SwitchUser, "�л�", "�л��û�", 3012, False, conMenu_Tool_Analyse)
    
    If mlngModule = G_LNG_PACSSTATION_MODULE Then
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_Release, "���Ŵ���", "�����Ƭ���Ŵ���", 3013, False)
        If Not cbrControl Is Nothing Then
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "���淢��", "���淢��", 8215, False)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FilmRelease, "��Ƭ����", "��Ƭ����", 8216, False)
        End If
    Else
        Set cbrPopControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "���淢��", "���淢��", 8215, False)
    End If
    
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_ReportExecutor, "����ִ��", "ָ����ǰ����ļ�¼��", 5008, False)
    
    If mlngModule = G_LNG_PACSSTATION_MODULE Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_SendArrange, "���Ͱ���", "���Ͱ���", 232, False)
    End If
    
'    'Σ�����
'    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_PacsCritical, "Σ��ֵ", "Σ�����", 8338, False)
'    If Not cbrControl Is Nothing Then
'        Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsCriticalReg, "Σ��ֵ�Ǽ�", "Σ��ֵ���ߵǼ�", 8345, False)
'        Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsCriticalManage, "Σ��ֵ����", "Σ��ֵ���߹���", 8338, True)
'    End If
    
    '�����������
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_Result, "���", "�����������", 3506, False)
    If Not cbrControl Is Nothing Then
        Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Negative, "����", "����", 3506, False)
        Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Positive, "����", "����", 3507, False)
    End If
    
    '����ǲ���ϵͳ����û�з��������ť
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_FuHeLevel, "�������", "�������", 8044, False)
        If Not cbrControl Is Nothing Then
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FuHe, "����", "����", 0, False)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_JiBenFuHe, "��������", "��������", 0, False)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_BuFuHe, "������", "������", 0, False)
        End If
    End If
        
'    'ֻ��Ӱ��ɼ�ϵͳ�ž����û���������
'    If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
'        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_ChangeUser, "����", "�������ҽ���ͱ���ҽ��", 3012, False)
'    End If
    
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Complete, "���", "����������", 225, False, , False)
  
  
    '---------------------�����������˵���������---------------------
    If mblnIsHasPatholModule Then
        If mobjWork_Pathol Is Nothing Then
            Set mobjWork_Pathol = New clsWorkModule_PatholV2
            Call mobjWork_Pathol.zlInitModule(Me, mlngModule, mstrPrivs, mlngCur����ID)
        End If
        
        Call mobjWork_Pathol.zlMenu.zlCreateMenu("", Me.cbrMain)
        Call mobjWork_Pathol.zlMenu.zlCreateToolBar("", Me.cbrMain.Item(2))
    End If

    Exit Sub
errH:
    Call HintError(err, "InitCommandBars", False)
End Sub


Private Function CreateMenu(objMenuControl As CommandBarControls, _
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
            If Not (CheckPopedom(mstrPrivs, "��Ⱦ�����Խ���Ǽ�") Or CheckPopedom(mstrPrivs, "��Ⱦ�����Խ����ѯ")) Then blHavePrives = False
            
        Case conMenu_Manage_DiseaseRegist
            If Not CheckPopedom(mstrPrivs, "��Ⱦ�����Խ���Ǽ�") Then blHavePrives = False
            
        Case conMenu_Manage_DiseaseQuery
            If Not CheckPopedom(mstrPrivs, "��Ⱦ�����Խ����ѯ") Then blHavePrives = False
            
        Case conMenu_Manage_PacsCritical, conMenu_Manage_PacsCriticalReg, conMenu_Manage_PacsCriticalManage
            If Not CheckPopedom(mstrPrivs, "Σ��ֵ����") Then blHavePrives = False
            
        Case conMenu_Manage_Undone
            If Not CheckPopedom(mstrPrivs, "ȡ��������") Then blHavePrives = False
            
        Case conMenu_Manage_RelatingPatiet
            If Not (CheckPopedom(mstrPrivs, "��������") And mSysPar.blnRelatingPatient) Then blHavePrives = False
            
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
            Set CreateMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
        Else
            Set CreateMenu = objMenuControl.Add(lngType, lngID, strCaption)
        End If
    
        CreateMenu.ID = lngID '������ﲻָ��id�����ܽ���Щ�˵���ӵ��Ҽ��˵���
        
        If lngIconId <> 0 Then CreateMenu.iconid = lngIconId
        If blnStartGroup Then CreateMenu.BeginGroup = True
        If strToolTip <> "" Then CreateMenu.ToolTipText = strToolTip
        
        If Not blHavePrives Then
            CreateMenu.Visible = False
        End If
        
        CreateMenu.Category = M_STR_MODULE_MENU_TAG
    End If
    Exit Function
err:
    If HintError(err, "CreateMenu", False) = 1 Then Resume
End Function

Private Sub ClearWorkModuleMenu()
    Dim i As Long
    Dim strTabName As String
    Dim strModuleTag As String
    
    For i = 0 To TabWindow.ItemCount - 1
        strTabName = TabWindow.Item(i).Caption
        strModuleTag = TabWindow.Item(i).tag
        
        Select Case strTabName
            Case C_TAB_NAME_Ӱ��ͼ��
                If Not mobjWork_PacsImg Is Nothing Then
                    Call mobjWork_PacsImg.zlMenu.zlClearMenu(strModuleTag)
                    Call mobjWork_PacsImg.zlMenu.zlClearToolBar(strModuleTag)
                End If
                
            Case C_TAB_NAME_Ӱ��ɼ�
            
            Case C_TAB_NAME_��鱨��
                If Not mobjWork_Report Is Nothing Then
                   Call mobjWork_Report.zlMenu.zlClearMenu(strModuleTag)
                   Call mobjWork_Report.zlMenu.zlClearToolBar(strModuleTag)
                End If
                
            Case C_TAB_NAME_ҽ����¼, C_TAB_NAME_������¼, C_TAB_NAME_���ü�¼, C_TAB_NAME_���Ӳ���
                If Not mobjWork_His Is Nothing Then
                    Call mobjWork_His.zlMenu.zlClearMenu(strModuleTag)
                    Call mobjWork_His.zlMenu.zlClearToolBar(strModuleTag)
                End If
                
            Case C_TAB_NAME_�걾����, C_TAB_NAME_����ȡ��, C_TAB_NAME_������Ƭ, C_TAB_NAME_�����ؼ�, C_TAB_NAME_���̱���
                If Not mobjWork_Pathol Is Nothing Then
                    If strModuleTag <> "" Then
                        Call mobjWork_Pathol.zlMenu.zlClearMenu(strModuleTag)
                        Call mobjWork_Pathol.zlMenu.zlClearToolBar(strModuleTag)
                    End If
                End If
                
            Case C_TAB_NAME_�Ŷӽк�
        End Select
    Next
End Sub


Private Sub CreateWorkModuleMenu(ByVal strTabName As String, ByVal strModuleTag As String)
'��������ģ��˵�
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim i As Long
     Dim lngToolCount As Long
    
On Error GoTo err

'    If Not mobjWork_Pathol Is Nothing And mblnIsHasPatholModule Then
'        Call mobjWork_Pathol.zlMenu.zlCreateMenu("", Me.cbrMain)
'        Call mobjWork_Pathol.zlMenu.zlCreateToolBar("", Me.cbrMain.Item(2))
'    End If
    
    
    Select Case strTabName
        Case C_TAB_NAME_Ӱ��ͼ��
            Call mobjWork_PacsImg.zlMenu.zlCreateMenu(strModuleTag, Me.cbrMain)
            Call mobjWork_PacsImg.zlMenu.zlCreateToolBar(strModuleTag, Me.cbrMain.Item(2))
            
        Case C_TAB_NAME_Ӱ��ɼ�
        Case C_TAB_NAME_��鱨��
            Call mobjWork_Report.zlMenu.zlCreateMenu(strModuleTag, Me.cbrMain)
            Call mobjWork_Report.zlMenu.zlCreateToolBar(strModuleTag, Me.cbrMain.Item(2))
            
        Case C_TAB_NAME_ҽ����¼, C_TAB_NAME_���ü�¼, C_TAB_NAME_������¼, C_TAB_NAME_���Ӳ���
            '��Ϊ��PACSϵͳ�� ����ӡ�� �˵����ڱ༭�˵����£������������ļ��˵��£������ڵ��ò����Ĳ˵���������ʱ��
            '���ļ��˵����Ҳ�����ӡ�˵����������PACS�У��嵥��ӡ���ļ��˵��£����Ե��ò����Ĳ˵���������ʱ��
            '�嵥��ӡ��id�ĳɴ�ӡ��id��������󣬻ָ��嵥��ӡԭ����id
            Set cbrControl = Nothing
            
            If strTabName = C_TAB_NAME_���Ӳ��� Then
                Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
                Set cbrControl = cbrMenuBar.CommandBar.Controls.Find(, conMenu_File_Excel)
                
                cbrControl.ID = conMenu_File_Print
                
                Call mobjWork_His.zlMenu.zlCreateMenu(strModuleTag, Me.cbrMain)
                
                If Not cbrControl Is Nothing Then cbrControl.ID = conMenu_File_Excel
                
            Else
                lngToolCount = cbrMain(2).Controls.Count
                
                '�����������˵���ťID,Ĭ�Ͻ���������ť��ӵ����λ�ã�ģ���ڲ�Ĭ�ϸ��ݲ˵�id��Ϊ1��2�İ�ť�������ʼλ��
                For i = 1 To lngToolCount
                    If cbrMain(2).Controls(i).Category = "Main" Then
                        cbrMain(2).Controls(i).ID = CLng(1) & CLng(cbrMain(2).Controls(i).ID)
                        cbrMain(2).Controls(i).Caption = "TMP-" & cbrMain(2).Controls(i).Caption
                    End If
                Next
                
                On Error GoTo errMenu
                    'ģ��˵�������������쳣����Ҫ�ָ�id��caption���ı�Ĳ˵�
                    Call mobjWork_His.zlMenu.zlCreateMenu(strModuleTag, Me.cbrMain)
errMenu:
                
                '�ָ��������˵���ťID
                For i = 1 To lngToolCount
                    If InStr(cbrMain(2).Controls(i).Caption, "TMP-") > 0 Then   '�жϰ�ť�Ƿ���ʱ���ù��˵�ID
                        cbrMain(2).Controls(i).ID = CLng(Mid(cbrMain(2).Controls(i).ID, 2, 255))
                        cbrMain(2).Controls(i).Caption = Replace(cbrMain(2).Controls(i).Caption, "TMP-", "")
                    End If
                Next
                
            End If
                         
            
        Case C_TAB_NAME_�걾����, C_TAB_NAME_����ȡ��, C_TAB_NAME_������Ƭ, C_TAB_NAME_�����ؼ�, C_TAB_NAME_���̱���
            If Len(strModuleTag) > 0 Then Call mobjWork_Pathol.zlMenu.zlCreateMenu(strModuleTag, Me.cbrMain)
        Case C_TAB_NAME_�Ŷӽк�
    End Select
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Find(, conMenu_ReportPopup)
    If Not cbrMenuBar Is Nothing Then
        If cbrMenuBar.CommandBar.Controls.Count <= 0 Then cbrMenuBar.Delete
    End If

    Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
    
    Call cbrMain.RecalcLayout
    
    Exit Sub
err:
    If Not cbrControl Is Nothing Then cbrControl.ID = conMenu_File_Excel
    If HintError(err, "CreateWorkModuleMenu<" & strTabName & "�˵�����>", False) = 1 Then Resume
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
            rsFullADO.Filter = " �ϼ�ID=" & NVL(rsFilterADO!ID)

            If rsFullADO.RecordCount > 0 Then
                Set cbrControl = Nothing
  
                If NVL(rsFilterADO!�Ƿ���) = 1 Or blnIsShare = True Then
                    mlngShareFatherID = NVL(rsFilterADO!ID)
                    '���������˵� ����ϼ�ID=1 ����ʾ����������
                    Set cbrControl = .Add(xtpControlButtonPopup, CLng(conMenu_Collection_ViewShare) * 10000# + j, NVL(rsFilterADO!�ղ����) & Decode(cbrParentControl.ID, conMenu_Collection_ViewShare, "(" & NVL(rsFilterADO!������) & ")", ""), -1, False)
                    cbrControl.DescriptionText = NVL(rsFilterADO!������)
                    cbrControl.Category = M_STR_MODULE_MENU_TAG
                    
                    j = j + 1
                    If i = 1 Then cbrControl.BeginGroup = True
                End If
                
                Set rsFilterTemp = zlDatabase.CopyNewRec(rsFullADO)
                '�����Լ�
                Call RecursionCreateShareMenu(rsFilterTemp, rsFullADO, IIf(cbrControl Is Nothing, cbrParentControl, cbrControl), IIf(NVL(rsFilterADO!�Ƿ���) = 0, False, True))
            Else
            '�����Ӽ��˵�
                If NVL(rsFilterADO!�Ƿ���) = 1 Or blnIsShare = True Then
                    Set cbrControl = .Add(xtpControlButton, CLng(conMenu_Collection_ViewShare) * 10000# + j, NVL(rsFilterADO!�ղ����) & Decode(cbrParentControl.ID, conMenu_Collection_ViewShare, "(" & NVL(rsFilterADO!������) & ")", ""), -1, False)
                    cbrControl.DescriptionText = NVL(rsFilterADO!������)
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

            rsFullADO.Filter = " �ϼ�ID=" & NVL(rsFilterADO!ID)
            mlngCollectionFatherID = NVL(rsFilterADO!ID)
            If rsFullADO.RecordCount > 0 Then
            '���������˵�
                Set cbrControl = .Add(xtpControlButtonPopup, CLng(comMenu_Collection_Type) * 10000# + j, NVL(rsFilterADO!�ղ����), -1, False)
                cbrControl.Category = M_STR_MODULE_MENU_TAG
                
                j = j + 1
                
                Set rsFilterTemp = zlDatabase.CopyNewRec(rsFullADO)
                '�����Լ�
                Call RecursionCreateCollectionMenu(rsFilterTemp, rsFullADO, cbrControl)
                
            Else
            '�����Ӽ��˵�
                Set cbrControl = .Add(xtpControlButton, CLng(comMenu_Collection_Type) * 10000# + j, NVL(rsFilterADO!�ղ����), -1, False)
                cbrControl.Category = M_STR_MODULE_MENU_TAG
                
                j = j + 1
            End If
            If i = 1 Then cbrControl.BeginGroup = True

            If Not rsFilterADO.EOF Then rsFilterADO.MoveNext
        Next
    End With

End Sub


Private Sub ReadBaseModuleName()
    '���õ�ǰ��Ҫ�����Ĺ���ҳ��
    mstrWorkModule = zlDatabase.GetPara("վ��ģ��", glngSys, mlngModule, "")
    mstrWorkModule = IIf(mstrWorkModule <> "", ";" & mstrWorkModule & ";", "")
    
    mstrWorkModule = Replace(Replace(mstrWorkModule, "ģ��", ""), "Ӱ�񱨸�", "��鱨��")
    mstrWorkModule = Replace(mstrWorkModule, "�������", "��鱨��")
    
    If mstrWorkModule = "" Then
        Select Case mlngModule
            Case G_LNG_PACSSTATION_MODULE
                mstrWorkModule = ";Ӱ��ͼ��;��鱨��;ҽ����¼;������¼;���Ӳ���;���ü�¼;"
            
            Case G_LNG_VIDEOSTATION_MODULE
                mstrWorkModule = ";Ӱ��ɼ�;��鱨��;ҽ����¼;������¼;���Ӳ���;���ü�¼;"
            
            Case G_LNG_PATHOLSYS_NUM
                mstrWorkModule = ";�걾����;Ӱ��ɼ�;����ȡ��;������Ƭ;�����ؼ�;���̱���;��鱨��;ҽ����¼;������¼;���Ӳ���;���ü�¼;"
            Case Else
                Exit Sub
        End Select
    End If
    
'    '���Դ���
'    mstrWorkModule = ";Ӱ��ͼ��ģ��;Ӱ��ɼ�ģ��;�걾����ģ��;����ȡ��ģ��;������Ƭģ��;�����ؼ�ģ��;���̱���ģ��;Ӱ�񱨸�ģ��;���ü�¼ģ��;ҽ����¼ģ��;������¼ģ��;"
End Sub



Private Function CreateTabItem(ByVal lngIndex As Long, ByVal strCaption As String, ByVal lngID As Long, _
    Optional ByVal strSelModuleName As String = "") As TabControlItem
    Dim objTabItem As TabControlItem
    
    Set objTabItem = TabWindow.InsertItem(lngIndex, strCaption, picTemp.hwnd, lngID)
'    objTabItem.Tag = strCaption
    
    If strSelModuleName = strCaption Then
        objTabItem.Selected = True
    End If
    
    Set CreateTabItem = objTabItem
End Function

Public Sub InitPacsHelper()
On Error GoTo errhandle
    Call ucPacsHelper1.Init(Me, mlngModule, mlngCur����ID, mstrPrivs, True)
Exit Sub
errhandle:
    Call HintError(err, "InitPacsHelper", False)
End Sub

Public Sub InitWorkModuleTab()
    Dim i As Integer
    Dim strSelModuleName As String
    Dim objTabItem As TabControlItem
    
    mblnIsHasPatholModule = False   '���ñ��������ȻΪfalseʱ�����������ɾ������˵�
    
'    strSelModuleName = "Ӱ��ɼ�" '��ע����ȡ�ϴ�ѡ��Ĺ���ģ��
    
    With TabWindow
        .RemoveAll
        Set .Icons = zlCommFun.GetPubIcons
        
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ButtonMargin.Left = 0
        .PaintManager.ButtonMargin.Top = 0
        .PaintManager.ButtonMargin.Bottom = 0
        .PaintManager.ButtonMargin.Right = 0
        .PaintManager.HeaderMargin.Left = 0
        .PaintManager.HeaderMargin.Top = 0
        .PaintManager.HeaderMargin.Right = 0
        .PaintManager.HeaderMargin.Bottom = 0
        
        
        .PaintManager.ShowIcons = True
        
        .RemoveAll
        
        '��ȡ����ģ������
        Call ReadBaseModuleName
    
        If InStr(mstrWorkModule, ";Ӱ��ͼ��;") > 0 Then
            Call CreateTabItem(0, C_TAB_NAME_Ӱ��ͼ��, 3551, strSelModuleName)
        End If
                        
        If mlngModule <> G_LNG_PACSSTATION_MODULE And CheckPopedom(mstrPrivs, "��Ƶ�ɼ�") _
            And InStr(mstrWorkModule, ";Ӱ��ɼ�;") > 0 Then
            Call CreateTabItem(1, C_TAB_NAME_Ӱ��ɼ�, conMenu_Cap_Dynamic, strSelModuleName)
        Else
            mstrWorkModule = Replace(mstrWorkModule, "Ӱ��ɼ�", "")
        End If
        
        If CheckPopedom(mstrPrivs, "�걾����") And InStr(mstrWorkModule, ";�걾����;") > 0 Then
            Call CreateTabItem(2, C_TAB_NAME_�걾����, G_INT_ICONID_SPECIMEN, strSelModuleName)
            
            mblnIsHasPatholModule = True
        Else
            mstrWorkModule = Replace(mstrWorkModule, "�걾����", "")
        End If
        
        If CheckPopedom(mstrPrivs, "����ȡ��") And InStr(mstrWorkModule, ";����ȡ��;") > 0 Then
            Call CreateTabItem(3, C_TAB_NAME_����ȡ��, G_INT_ICONID_MATERIAL, strSelModuleName)
            
            mblnIsHasPatholModule = True
        Else
            mstrWorkModule = Replace(mstrWorkModule, "����ȡ��", "")
        End If
        
        If CheckPopedom(mstrPrivs, "������Ƭ") And InStr(mstrWorkModule, ";������Ƭ;") > 0 Then
            Call CreateTabItem(4, C_TAB_NAME_������Ƭ, G_INT_ICONID_SLICES, strSelModuleName)
            
            mblnIsHasPatholModule = True
        Else
            mstrWorkModule = Replace(mstrWorkModule, "������Ƭ", "")
        End If
        
        If (CheckPopedom(mstrPrivs, "�����黯") Or CheckPopedom(mstrPrivs, "����Ⱦɫ") Or CheckPopedom(mstrPrivs, "���Ӳ���")) _
            And InStr(mstrWorkModule, ";�����ؼ�;") > 0 Then
            Call CreateTabItem(5, C_TAB_NAME_�����ؼ�, G_INT_ICONID_SPEEXAM, strSelModuleName)
            
            mblnIsHasPatholModule = True
        Else
            mstrWorkModule = Replace(mstrWorkModule, "�����ؼ�", "")
        End If
        
        If (CheckPopedom(mstrPrivs, "��������") Or CheckPopedom(mstrPrivs, "��Ⱦ����") _
            Or CheckPopedom(mstrPrivs, "���ӱ���") Or CheckPopedom(mstrPrivs, "���߱���") _
            Or CheckPopedom(mstrPrivs, "�����ؼ챨�����")) And InStr(mstrWorkModule, ";���̱���;") > 0 Then
            Call CreateTabItem(6, C_TAB_NAME_���̱���, G_INT_ICONID_PROREPORT, strSelModuleName)
            
            mblnIsHasPatholModule = True
        Else
            mstrWorkModule = Replace(mstrWorkModule, "���̱���", "")
        End If
        
        If GetInsidePrivs(p���Ʊ������, True) <> "" And _
             InStr(mstrWorkModule, ";��鱨��;") > 0 Then
            Call CreateTabItem(7, C_TAB_NAME_��鱨��, 10008, strSelModuleName) 'conMenu_Edit_Compend
        Else
            mstrWorkModule = Replace(mstrWorkModule, "��鱨��", "")
        End If
        
        If mobjAppendBill Is Nothing Then   'ʹ�û��ģʽʱ������ʾǶ��Ĳ����ѹ���
            If GetInsidePrivs(pҽ�����ѹ���, True) <> "" And InStr(mstrWorkModule, ";���ü�¼;") > 0 Then
                Call CreateTabItem(8, C_TAB_NAME_���ü�¼, 10007, strSelModuleName)
            Else
                mstrWorkModule = Replace(mstrWorkModule, "���ü�¼", "")
            End If
        End If
        
        If (GetInsidePrivs(pסԺҽ���´�, True) <> "" Or GetInsidePrivs(p����ҽ���´�, True) <> "") _
            And InStr(mstrWorkModule, ";ҽ����¼;") > 0 Then
            Call CreateTabItem(9, C_TAB_NAME_ҽ����¼, 10010, strSelModuleName)
        Else
            mstrWorkModule = Replace(mstrWorkModule, "ҽ����¼", "")
        End If
        
        If (GetInsidePrivs(pסԺ��������, True) <> "" _
            Or GetInsidePrivs(p���ﲡ������, True) <> "" _
            Or GetInsidePrivs(p������Ӳ���, True) <> "" _
            Or GetInsidePrivs(pסԺ���Ӳ���, True) <> "") _
            And InStr(mstrWorkModule, ";������¼;") > 0 Then
            Call CreateTabItem(10, C_TAB_NAME_������¼, 10009, strSelModuleName)
        Else
            mstrWorkModule = Replace(mstrWorkModule, "������¼", "")
        End If

        
        '����Ŷӽк�ҳ��
        If mSysPar.blnUseQueue = True Then
            mstrWorkModule = mstrWorkModule & ";�Ŷӽк�;"
            
            Call CreateTabItem(11, C_TAB_NAME_�Ŷӽк�, 10011, strSelModuleName)
            
'            '��ݽкŽ���
'            If mSysPar.blnQueueQuick Then
'                If Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
'                    Call mobjQueue.OpenQueueQuick(GetSelQueueRooms(True), Me)
'                End If
'            End If
        End If
    
'        If Not GetVideoForm Is Nothing Then Call GetVideoForm.ShowVideoWindow(picVideoContainer)
    End With
    
    '���û��Ĭ��tabҳ����Ĭ����ʾ��һ��tab��ǩ
    If TabWindow.Selected Is Nothing Then
        If TabWindow.ItemCount > 0 Then
            TabWindow.Item(0).Selected = True
        End If
    End If
    
End Sub
 

Private Sub mobjPacsCore_AfterSaveReportImage(strStudyUID As String)
    Dim lngAdviceId As Long
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim aryFiles() As String
    Dim i As Long
    Dim strCurStudyUID As String
    
On Error GoTo errhandle
    strCurStudyUID = strStudyUID
    
    If mSysPar.blnDirectSendRepImg Then strCurStudyUID = Split(strStudyUID & "-", "-")(0)
    
    If mobjCurStudyInfo.strStudyUID = strCurStudyUID Then
        lngAdviceId = mobjCurStudyInfo.lngAdviceId
    Else
        strSQL = "Select ҽ��ID from Ӱ�����¼ where ���UID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯҽ��ID", strCurStudyUID)
        If rsData.RecordCount <= 0 Then Exit Sub
        
        lngAdviceId = Val(NVL(rsData!ҽ��ID))
    End If
    
    If Not mobjWork_Report Is Nothing Then
        Call SyncHelperDataState(lngAdviceId, 0, 0)
       
        If mSysPar.blnDirectSendRepImg Then
            aryFiles = Split(Split(strStudyUID & "-", "-")(1) & ";", ";")
            
            For i = 0 To UBound(aryFiles)
                If Len(aryFiles(i)) > 0 Then
                    Call AddViewImageToReport(lngAdviceId, aryFiles(i))
                End If
            Next i
        End If
    End If
    
    
    
    Exit Sub
    
errhandle:
    If HintError(err, "mobjPacsCore_AfterSaveReportImage") = 1 Then Resume
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
    If HintError(err, "mobjQueue_OnDiagnose", False) = 1 Then Resume
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
    If HintError(err, "mobjQueue_OnCompleted", False) = 1 Then Resume
End Sub

Private Sub mobjQueue_OnSelChange(ByVal lngAdviceId As Long)
'��ѡ��ı��¼�
On Error GoTo errhandle
    Dim lngIndex As Long
    Dim strCurTabName As String
    
    strCurTabName = mstrSelTabName
    
    '�����ǰģ�����Ŷӽкţ����Ŷӽк�ģ���ѡ���иı�󣬲���Ҫ������ǰ���ڵ�ģ��ˢ�·���,ֻ��ˢ���б�����Ϣ��ʾ��pacshelper��
    '���mstrSelTabName����Ϊ�գ��򲻻�ˢ��ģ�����
    If mstrSelTabName = C_TAB_NAME_�Ŷӽк� Then mstrSelTabName = ""
    
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
            Else
                HintMsg "���ͬ����λʧ�ܣ��볢�Բ��ҡ�", "mobjQueue_OnSelChange", vbOKOnly
            End If
        End With
    End If
    
    mstrSelTabName = strCurTabName
Exit Sub
errhandle:
    mstrSelTabName = strCurTabName
    If HintError(err, "mobjQueue_OnSelChange", False) = 1 Then Resume
End Sub
 
  


Public Sub ReportResultHint(ByVal lngOrderID As Long)
On Error GoTo errhandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strResultInput As String
    
    strResultInput = ""
    
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
     
Exit Sub
errhandle:
    If HintError(err, "ReportResultHint") = 1 Then Resume
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
                strSQL = "Zl_Ӱ����_��鼼ʦ( " & lngAdviceId & ",'" & IIf(blnAddImage = True, mstrOtherUserName, "") & "')"
            Else
                strSQL = "Zl_Ӱ����_��鼼ʦ( " & lngAdviceId & ",'" & IIf(blnAddImage = True, mstrHisUserName, "") & "')"
            End If

            zlDatabase.ExecuteProcedure strSQL, GetWindowCaption
        End If
        
        If blnStateChanged Then
            Call UpdateQueryListData(Nothing, lngAdviceId)
        End If
        
    End With
    Exit Sub
errH:
    If HintError(err, "UpdateStudyListState") = 1 Then Resume
End Sub

Private Function ShowBillList(objPopup As CommandBarPopup) As Boolean
'���ܣ���ʾ��ǰִ��ҽ�����Դ�ӡ�����Ƶ����ڲ˵���
    Dim rsTmp As New ADODB.Recordset
    Dim objControl As CommandBarControl
    Dim strSQL As String
        
    On Error GoTo errH
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "ShowBillList", vbInformation
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
    
    '���Ƶ�����Ŀʹ�ó�����Ҫ����סԺ����죬�������Ĭ��Ϊ����
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjCurStudyInfo.lngAdviceId, CLng(Decode(mobjCurStudyInfo.lngPatientFrom, 3, 1, mobjCurStudyInfo.lngPatientFrom))) 'mobjCurStudyInfo.lngPatientFrom
    
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
    If HintError(err, "ShowBillList") = 1 Then Resume
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
        HintMsg M_STR_HINT_NoSelectData, "FuncBillPrint", vbInformation
        Exit Sub
    End If
    
    If ReportPrintSet(gcnOracle, glngSys, objControl.Parameter, Me) Then
        Call ReportOpen(gcnOracle, glngSys, objControl.Parameter, Me, "NO=" & mobjCurStudyInfo.strNO, _
                       "����=" & mobjCurStudyInfo.lngRecordKind, "ҽ��ID=" & mobjCurStudyInfo.lngAdviceId, 1)
    End If
    Exit Sub
errhandle:
    If HintError(err, "FuncBillPrint", False) = 1 Then Resume
End Sub


Public Sub RefreshList()
'blClick �Ƿ���ˢ�´�����ˢ���б�
'ˢ�������б�
    
    If mblnIsLoading = True Then
        HintMsg "���ݼ����У����Ժ�����...", "RefreshList", vbInformation
        Exit Sub
    End If
    
On Error GoTo errhandle
    mblnIsLoading = True
        
    Call mobjPacsQueryWrap.ExecuteQuery(C_QUERY_ˢ��)
    If mobjPacsQueryWrap.SqlScheme.AutoRefreshTimeLen > 0 Then TimerRefresh.Enabled = True
    
    If mobjCurStudyInfo.lngAdviceId > 0 Then
'        RefreshDisPlay���������� 2��ʾ���²���
        Call mobjPacsQueryWrap.RefreshDisplay(vsfList.Row, mobjCurStudyInfo.lngAdviceId, 2)
    End If
    
    'ֱ�ӿ�ʼ��λ
    If vsfList.Rows <= 1 Then
        '��û������ʱ��֪ͨˢ�¹���ģ������ص�����
        Set mobjCurStudyInfo = GetNullAdviceInf
    End If
    
    Call RefreshModuleData(mstrSelTabName, mstrSelModuleTag, mobjSelModule)

    mblnIsLoading = False

    Exit Sub
errhandle:
    mblnIsLoading = False
    If HintError(err, "RefreshList", False) = 1 Then Resume
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


Private Sub picHelper_Resize()
On Error Resume Next
    ucPacsHelper1.Left = 0
    ucPacsHelper1.Top = 0
    ucPacsHelper1.Width = picHelper.ScaleWidth
    ucPacsHelper1.Height = picHelper.ScaleHeight
'    ucPacsHelper1.Move 0, 0, picHelper.ScaleWidth, picHelper.ScaleHeight
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
    HintError err, "cmdLocate_Click<��λ����>", False
End Sub

 
Private Sub picTabFace_Resize()
On Error Resume Next
    Dim i As Long
    Dim lngLen As Long
    
    lngLen = 0
    For i = 0 To TabWindow.ItemCount - 1
        lngLen = lngLen + picTabFace.TextWidth("��") * Len(TabWindow.Item(i).Caption) + 700
    Next
    
    TabWindow.Width = lngLen
 
    If TabWindow.Width < picTabFace.ScaleWidth Then
        TabWindow.Left = picTabFace.ScaleWidth - TabWindow.Width + (TabWindow.ItemCount * 240)
        TabWindow.Top = 0
        TabWindow.Height = picTabFace.Height
    Else
        TabWindow.Move 0, 0, picTabFace.ScaleWidth, picTabFace.ScaleHeight
    End If
     
End Sub

Private Sub picWindow_Resize()
    Dim R As RECT
On Error GoTo errhandle
    If mlngSelHwnd <> 0 Then
        Call MoveWindow(mlngSelHwnd, 0, 0, _
            picWindow.ScaleX(picWindow.Width, vbTwips, vbPixels), _
            picWindow.ScaleY(picWindow.Height, vbTwips, vbPixels), 0)
            
        GetClientRect mlngSelHwnd, R
        InvalidateRect mlngSelHwnd, R, 1
    End If
errhandle:
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
    HintError err, "tabScheme_SelectedChanged<�����л�>", False
End Sub

Private Function InsertWorkModuleInfo(ByVal strModuleName As String, ByVal lngHwnd As Long, ByVal lngDeptId As Long, _
    objModule As Object) As TWorkModuleInfo
'����Ƕ��Ĺ���ģ����Ϣ
    Dim lngBound As Long
    Dim i As Long
    
    For i = 1 To UBound(mAryWorkModule)
        If strModuleName = mAryWorkModule(i).ModuleName Then
            '�ҵ�ģ�����Ҫʹ�þ���ж��Ƿ��Ӧ��ģ��ʵ��
            If lngHwnd = mAryWorkModule(i).hwnd Then
                InsertWorkModuleInfo = mAryWorkModule(i)
            Else
                '�������ͬʱ�������ģ����Ϣ
                mAryWorkModule(i).hwnd = lngHwnd
                mAryWorkModule(i).FontSize = 0
                mAryWorkModule(i).DeptId = lngDeptId
                
                Set mAryWorkModule(i).objModule = objModule
            End If
            
            Exit Function
        End If
    Next
    
    lngBound = UBound(mAryWorkModule) + 1
    ReDim Preserve mAryWorkModule(lngBound)
    
    mAryWorkModule(lngBound).FontSize = 0
    mAryWorkModule(lngBound).hwnd = lngHwnd
    mAryWorkModule(lngBound).DeptId = lngDeptId
    mAryWorkModule(lngBound).ModuleName = strModuleName
    
    Set mAryWorkModule(lngBound).objModule = objModule
    
    InsertWorkModuleInfo = mAryWorkModule(i)
End Function

Private Function GetWorkModuleInfo(ByVal strModuleName As String) As Long
    Dim i As Long
    
    For i = 1 To UBound(mAryWorkModule)
        If strModuleName = mAryWorkModule(i).ModuleName Then
            GetWorkModuleInfo = i
            Exit Function
        End If
    Next
End Function

Private Function GetWorkModuleTag(ByVal strTabName As String) As String
'��ȡ����ģ����,������סԺҽ��������ҽ����סԺ���������ﲡ������������ȹؼ���
    Dim i As Long
    
    GetWorkModuleTag = ""
    
    For i = 0 To TabWindow.ItemCount - 1
        If TabWindow.Item(i).Caption = strTabName Then
            GetWorkModuleTag = TabWindow.Item(i).tag
            Exit Function
        End If
    Next
End Function

Private Function GetWorkModuleName(ByVal strTabName As String, ByVal lngDeptId As Long, Optional ByVal lngPatientFrom As Long = 0) As String
    Dim lngReportType As Long
    
    Select Case strTabName
        Case C_TAB_NAME_Ӱ��ͼ��
            GetWorkModuleName = C_WORKMODULE_NAME_Ӱ��ͼ��
            
        Case C_TAB_NAME_Ӱ��ɼ�
            GetWorkModuleName = C_WORKMODULE_NAME_Ӱ��ɼ�
            
        Case C_TAB_NAME_�걾����
            GetWorkModuleName = C_WORKMODULE_NAME_�걾����
            
        Case C_TAB_NAME_����ȡ��
            GetWorkModuleName = C_WORKMODULE_NAME_����ȡ��
            
        Case C_TAB_NAME_������Ƭ
            GetWorkModuleName = C_WORKMODULE_NAME_������Ƭ
            
        Case C_TAB_NAME_�����ؼ�
            GetWorkModuleName = C_WORKMODULE_NAME_�����ؼ�
        
        Case C_TAB_NAME_���̱���
            GetWorkModuleName = C_WORKMODULE_NAME_���̱���
            
        Case C_TAB_NAME_��鱨��
            lngReportType = Val(GetDeptPara(lngDeptId, "����༭��", 0))
            
            If lngReportType = ReportType.PACS����༭�� Then
                GetWorkModuleName = C_WORKMODULE_NAME_�ϰ汨��
            ElseIf lngReportType = ReportType.���Ӳ����༭�� Then
                GetWorkModuleName = C_WORKMODULE_NAME_��������
            Else
                GetWorkModuleName = C_WORKMODULE_NAME_���ܱ���
            End If
            
        Case C_TAB_NAME_ҽ����¼
            If lngPatientFrom <> 2 Then
                GetWorkModuleName = C_WORKMODULE_NAME_����ҽ��
            Else
                GetWorkModuleName = C_WORKMODULE_NAME_סԺҽ��
            End If
            
        Case C_TAB_NAME_������¼
            If lngPatientFrom <> 2 Then
                GetWorkModuleName = C_WORKMODULE_NAME_���ﲡ��
            Else
                GetWorkModuleName = C_WORKMODULE_NAME_סԺ����
            End If
            
        Case C_TAB_NAME_���Ӳ���
            GetWorkModuleName = C_WORKMODULE_NAME_���Ӳ���
            
        Case C_TAB_NAME_���ü�¼
            GetWorkModuleName = C_WORKMODULE_NAME_���ü�¼
            
        Case C_TAB_NAME_�Ŷӽк�
            GetWorkModuleName = C_WORKMODULE_NAME_�Ŷӽк�
            
    End Select
End Function

Private Function VerifyModuleObj(ByVal strTabName As String, _
    Optional ByRef objModule As Object, Optional ByRef strModuleTag As String, _
    Optional ByVal blnReInit As Boolean = False) As Long
'��֤ģ�����
    Dim lngBound As Long
    Dim lngCurDeptId As Long
    Dim strWorkModuleName As String
    
    VerifyModuleObj = 0
    Set objModule = Nothing
    
    '��ȡ��ǰ�Ŀ���ID
    lngCurDeptId = GetCurDeptId
    
    strWorkModuleName = GetWorkModuleName(strTabName, lngCurDeptId, GetCurPatientFrom)
    strModuleTag = strWorkModuleName
    
    ';�걾����;Ӱ��ɼ�;����ȡ��;������Ƭ;�����ؼ�;���̱���;��鱨��;ҽ����¼;������¼;���ü�¼;
    Select Case strTabName
        Case C_TAB_NAME_Ӱ��ͼ��
            If mobjWork_PacsImg Is Nothing Then
                Set mobjWork_PacsImg = New frmWork_ImageV2
                
                Set mobjWork_PacsImg.PacsCore = mobjPacsCore
                Call mobjWork_PacsImg.zlInitModule(Me, mlngModule, mstrPrivs, lngCurDeptId)
                
                Call InsertWorkModuleInfo(strWorkModuleName, mobjWork_PacsImg.hwnd, lngCurDeptId, mobjWork_PacsImg)
            Else
                If blnReInit Or mobjWork_PacsImg.DeptId <> lngCurDeptId Then Call mobjWork_PacsImg.zlInitModule(Me, mlngModule, mstrPrivs, lngCurDeptId)
            End If
            
            If mobjWork_PacsImg Is Nothing Then Exit Function
            
            Set objModule = mobjWork_PacsImg
            VerifyModuleObj = objModule.hwnd
            
        Case C_TAB_NAME_Ӱ��ɼ�
            If mobjWork_ImageCap Is Nothing Then
                Set mobjWork_ImageCap = New zl9PacsImageCap.clsPacsCaptureV2 ' CreateObject("zl9PacsImageCap.clsPacsCaptureV2")
                With mobjWork_ImageCap
                    .ModuleNo = mlngModule
                    .ParentWindowKey = Me.Name
                    .AllowEventNotify = True
                    
                    Call .zlInitModule(gcnOracle, mobjCapLinker, glngSys, mlngModule, mstrPrivs, lngCurDeptId, Me.hwnd, gblnUseDebugLog)
                End With
                
                Call InsertWorkModuleInfo(strWorkModuleName, mobjWork_ImageCap.ContainerHwnd, lngCurDeptId, mobjWork_ImageCap)
                
                Set ucPacsHelper1.MainVideoWindow = mobjWork_ImageCap
            Else
                If blnReInit Or mobjWork_ImageCap.DeptId <> lngCurDeptId Then Call mobjWork_ImageCap.zlInitModule(gcnOracle, mobjCapLinker, glngSys, mlngModule, mstrPrivs, lngCurDeptId, Me.hwnd, gblnUseDebugLog)
            End If
            
            If mobjWork_ImageCap Is Nothing Then Exit Function
             
            Set objModule = mobjWork_ImageCap
            VerifyModuleObj = objModule.ContainerHwnd
    
        Case C_TAB_NAME_�걾����, C_TAB_NAME_����ȡ��, C_TAB_NAME_������Ƭ, C_TAB_NAME_�����ؼ�, C_TAB_NAME_���̱���
            If mobjWork_Pathol Is Nothing Then
                Set mobjWork_Pathol = New clsWorkModule_PatholV2
                Call mobjWork_Pathol.zlInitModule(Me, mlngModule, mstrPrivs, lngCurDeptId)
            Else
                If blnReInit Or mobjWork_Pathol.DeptId <> lngCurDeptId Then Call mobjWork_Pathol.zlInitModule(Me, mlngModule, mstrPrivs, lngCurDeptId)
            End If
            
            If mobjWork_Pathol Is Nothing Then Exit Function
            
            Set objModule = mobjWork_Pathol.GetModule(strWorkModuleName)
            VerifyModuleObj = objModule.hwnd
            
            Call InsertWorkModuleInfo(strWorkModuleName, objModule.hwnd, lngCurDeptId, objModule)
            
        Case C_TAB_NAME_��鱨��
            If mobjWork_Report Is Nothing Then
                Set mobjWork_Report = New clsWorkModule_ReportV2
                
                Call mobjWork_Report.zlInitModule(Me, mlngModule, mstrPrivs, lngCurDeptId, mobjCapLinker, ucPacsHelper1) 'mlngCur����ID
                
            Else
                If blnReInit Or mobjWork_Report.DeptId <> lngCurDeptId Then
                    Call mobjWork_Report.zlInitModule(Me, mlngModule, mstrPrivs, lngCurDeptId, mobjCapLinker, ucPacsHelper1) 'mlngCur����ID
                    Call mobjWork_Report.ReInit(strWorkModuleName)
                End If
            End If
            
            If mobjWork_Report Is Nothing Then Exit Function
            
            Set objModule = mobjWork_Report.zlGetForm(strWorkModuleName)
            VerifyModuleObj = objModule.hwnd
            
            
             Call InsertWorkModuleInfo(strWorkModuleName, objModule.hwnd, lngCurDeptId, objModule)
            
        Case C_TAB_NAME_ҽ����¼, C_TAB_NAME_������¼, C_TAB_NAME_���Ӳ���, C_TAB_NAME_���ü�¼
            If mobjWork_His Is Nothing Then
                Set mobjWork_His = New clsWorkModule_HisV2
                Call mobjWork_His.zlModule.zlInitModule(Me, mlngModule, mstrPrivs, lngCurDeptId)
            Else
                If blnReInit Or mobjWork_His.zlModule.DeptId <> lngCurDeptId Then Call mobjWork_His.zlModule.zlInitModule(Me, mlngModule, mstrPrivs, lngCurDeptId)
            End If
            
            If mobjWork_His Is Nothing Then Exit Function
            
            Set objModule = mobjWork_His.zlModule.zlGetModule(strWorkModuleName)
            VerifyModuleObj = objModule.hwnd
            
            Call InsertWorkModuleInfo(strWorkModuleName, objModule.hwnd, lngCurDeptId, objModule)
            
        Case C_TAB_NAME_�Ŷӽк�
            If mobjQueue Is Nothing Then
                Set mobjQueue = New frmWork_Queue
                Call mobjQueue.zlInitPacsQueueCfg(mlngModule, lngCurDeptId, zlStr.NeedName(mstrCur����), mstrPrivs, mblnAllDepts, Me)
            Else
                If blnReInit Or mobjQueue.DeptId <> lngCurDeptId Then Call mobjQueue.zlInitPacsQueueCfg(mlngModule, lngCurDeptId, zlStr.NeedName(mstrCur����), mstrPrivs, mblnAllDepts, Me)
            End If
            
            If mobjQueue Is Nothing Then Exit Function
            
            Set objModule = mobjQueue
            VerifyModuleObj = objModule.hwnd
            
            Call InsertWorkModuleInfo(strWorkModuleName, objModule.hwnd, lngCurDeptId, objModule)
            
    End Select
End Function

Private Function GetPatholModuleType(ByVal strModuleName As String) As TPatholModuleType
'��ȡ����ģ������
    Select Case strModuleName
        Case C_TAB_NAME_�걾����
            GetPatholModuleType = pmtSpecimen
        Case C_TAB_NAME_����ȡ��
            GetPatholModuleType = pmtMaterial
        Case C_TAB_NAME_������Ƭ
            GetPatholModuleType = pmtSlices
        Case C_TAB_NAME_�����ؼ�
            GetPatholModuleType = pmtSpeExam
        Case C_TAB_NAME_���̱���
            GetPatholModuleType = pmtProRep
    End Select
End Function

Private Sub EmbedWindow(ByVal lngHwnd As Long)
'Ƕ�봰�ڴ���
    SetParent lngHwnd, picWindow.hwnd
    '��ʾ����
    ShowWindow lngHwnd, 1
    
    Call MoveWindow(lngHwnd, 0, 0, _
            picWindow.ScaleX(picWindow.Width, vbTwips, vbPixels), _
            picWindow.ScaleY(picWindow.Height, vbTwips, vbPixels), 0)
            
'    SetWindowPos lngHwnd, -1, 0, 0, _
'            picWindow.ScaleX(picWindow.Width, vbTwips, vbPixels), _
'            picWindow.ScaleY(picWindow.Height, vbTwips, vbPixels), 3
         
    BringWindowToTop lngHwnd
    
'    '��ʾ����
'    ShowWindow lngHwnd, 1
End Sub

Private Sub AutoHideHelper(ByVal strTabName As String)
    Dim strWorkModuleTag As String
   
   ucPacsHelper1.TabEnable("ͼ��") = True
   ucPacsHelper1.AllowLinkerViewer = True
   
    If strTabName <> C_TAB_NAME_Ӱ��ɼ� Then
        If mlngModule = G_LNG_PACSSTATION_MODULE Then
            'Ӱ��ҽ����Ƕ����Ƶ�ɼ�
            ucPacsHelper1.AllowEmbedVideo = False
            ucPacsHelper1.HideEmbedVideo
        Else
            ucPacsHelper1.AllowEmbedVideo = IIf(Val(GetDeptPara(mlngCur����ID, "��ʾ��Ƶ�ɼ�", "0")) <> 0, True, False) 'True
            
            '�������ʽ���洰��������Ƶ�ɼ����������治�л�Ƕ��ʽ��Ƶ�ɼ�
            If Not mobjCapLinker Is Nothing And Not mobjWork_ImageCap Is Nothing And VideoIsAttachReportWindow = False Then
                Call ucPacsHelper1.ShowEmbedVideo(mobjCapLinker, True)
                '�ָ�֮ǰ�༭�����ڵĽ���
                If strTabName = C_TAB_NAME_��鱨�� Then
                    If Not mobjSelModule Is Nothing Then Call mobjSelModule.SetFocus
                End If
            Else
                '���ؿ���֮ǰ����ƵǶ��ʽ����
                ucPacsHelper1.HideEmbedVideo
            End If
        End If
        
        ucPacsHelper1.TabEnable("�ʾ�") = False
        ucPacsHelper1.AllowWrite = False
            
        If strTabName = C_TAB_NAME_��鱨�� Then
            strWorkModuleTag = GetWorkModuleTag(C_TAB_NAME_��鱨��)
            
            ucPacsHelper1.TabEnable("�ʾ�") = IIf(strWorkModuleTag = C_WORKMODULE_NAME_�ϰ汨��, True, False)
            ucPacsHelper1.AllowWrite = IIf(strWorkModuleTag = C_WORKMODULE_NAME_�ϰ汨��, True, False)
            
            If ucPacsHelper1.tag = "�ʾ�" Then
                Call ucPacsHelper1.LocateTab(ucPacsHelper1.tag)
                If Not mobjSelModule Is Nothing Then Call mobjSelModule.SetFocus
            End If
            
        ElseIf strTabName = C_TAB_NAME_Ӱ��ͼ�� Then
            Call ucPacsHelper1.LocateTab("��ʷ")
            ucPacsHelper1.TabEnable("ͼ��") = False
            
        ElseIf strTabName = C_TAB_NAME_�Ŷӽк� Then
            ucPacsHelper1.AllowLinkerViewer = False
            
'        ElseIf strTabName = C_TAB_NAME_���̱��� Then
'            ucPacsHelper1.AllowLinkerViewer = False
            
        End If
    Else
        ucPacsHelper1.AllowLinkerViewer = False
        ucPacsHelper1.AllowEmbedVideo = False
        ucPacsHelper1.HideEmbedVideo
        ucPacsHelper1.TabEnable("�ʾ�") = False
    End If
End Sub


Private Function ReloadWorkModule(ByVal strTabName As String, _
    Optional ByRef strModuleTag As String = "", _
    Optional ByRef objSelModule As Object, Optional ByVal blnReInit As Boolean = False) As Long
'���ع���ģ��
    ';�걾����;Ӱ��ɼ�;����ȡ��;������Ƭ;�����ؼ�;���̱���;�������;ҽ����¼;������¼;���ü�¼;
 
    Dim lngSelHwnd As Long
    Dim lngModuleInfoIndex As Long
    
    ReloadWorkModule = 0
    
    lngSelHwnd = VerifyModuleObj(strTabName, objSelModule, strModuleTag, blnReInit)
    
    If lngSelHwnd = 0 Then
        HintMsg "��ȡ[" & strModuleTag & "]ģ����ض���ʧ�ܡ�", "ReloadWorkModule<����ģ��>", vbOKOnly
        Exit Function
    End If
     
    Call EmbedWindow(lngSelHwnd)
    
    '�����������ɼ����ڣ����⽹���л����⣬��رչ�Ƭ���ں�Ӱ��ͼ�������ʾ���Ǽ�鱨��ģ��
    For lngModuleInfoIndex = 0 To UBound(mAryWorkModule)
        If mAryWorkModule(lngModuleInfoIndex).hwnd <> lngSelHwnd Then ShowWindow mAryWorkModule(lngModuleInfoIndex).hwnd, 0
    Next
    
    lngModuleInfoIndex = GetWorkModuleInfo(strModuleTag)
    
    If mAryWorkModule(lngModuleInfoIndex).FontSize <> gbytFontSize Then
        Call ReSetModuleFontSize(strTabName, strModuleTag, objSelModule, gbytFontSize)
        mAryWorkModule(lngModuleInfoIndex).FontSize = gbytFontSize
    End If
    
    'ˢ��ģ������
    Call RefreshModuleData(strTabName, strModuleTag, objSelModule)
     
    ''�����˵�
    LockWindowUpdate Me.hwnd
On Error GoTo errhandle
    Call ClearWorkModuleMenu
    Call CreateWorkModuleMenu(strTabName, strModuleTag)
errhandle:
    If err.Number <> 0 Then
        HintError err, "ReloadWorkModule<����ģ��>", False
    End If
    LockWindowUpdate 0
    
    ReloadWorkModule = lngSelHwnd
End Function


Private Sub TabWindow_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errhandle
    Dim strModuleTag As String
 
    If Not mblnInitOk Then Exit Sub
     
     If mstrSelTabName <> Item.Caption And mstrSelTabName = C_TAB_NAME_��鱨�� And mstrSelModuleTag = C_WORKMODULE_NAME_�ϰ汨�� Then
        '�������汣����ʾ
        If Not mobjSelModule Is Nothing Then Call mobjSelModule.PromptSave
     End If
     
    '�л�����ʱ�����ܻὫtabwindow�Ĳ���item.tag��������Ϊ�գ��Ա��л���ǩʱ�����ԶԽ������ˢ�£����鱨��ģ��
    Call SelectModule(Item.Caption, strModuleTag, IIf(Item.tag = "", True, False))
    
    '����ʵ��ʹ�õ�ģ��������סԺҽ��������ҽ����סԺ���������ﲡ�������������
    Item.tag = strModuleTag
    
    Select Case mstrSelTabName
        Case C_TAB_NAME_�Ŷӽк�
            Call RefreshPacsQueueData(True)
 
    
        Case C_TAB_NAME_��鱨��
            If TypeOf mobjSelModule Is frmReportV2 Then
                Call mobjSelModule.LocateEditBox
            End If
            
    End Select

    Exit Sub
errhandle:
    If HintError(err, "TabWindow_SelectedChanged", False) = 1 Then Resume
End Sub

Private Sub SelectModule(ByVal strTabName As String, ByRef strModuleTag As String, Optional ByVal blnReInit As Boolean = False)
    Dim objSel As Object
    
    mlngSelHwnd = ReloadWorkModule(strTabName, strModuleTag, objSel, blnReInit)
    If mlngSelHwnd = 0 Then Exit Sub
    
    mstrSelTabName = strTabName
    mstrSelModuleTag = strModuleTag
    
    Set mobjSelModule = objSel

'    If Not mobjWork_Report Is Nothing And Item.tag = "������д" Then
'        Call mobjWork_Report.AllowLocate(True)
'    End If
    timerHelper.Enabled = True
    
End Sub


 

Private Sub timerCapture_Timer()
On Error GoTo errhandle
    Dim strKeyAlias As String
    
    If Not mblnInitOk Then Exit Sub
    timerCapture.Enabled = False
    
    strKeyAlias = GetKeyAlias(mCaptureMsg.lngMsg, 0)
    
    If strKeyAlias = mstrCaptureHot Or strKeyAlias = mstrCaptureAfterHot Or strKeyAlias = mstrCaptureAfterTagHot Then
        If mobjWork_ImageCap Is Nothing And mlngModule <> G_LNG_PACSSTATION_MODULE Then
            Call VerifyModuleObj(C_TAB_NAME_Ӱ��ɼ�)
            If Not mobjWork_ImageCap Is Nothing Then
                mobjWork_ImageCap.zlRefreshVideoWindow
                Sleep 1000  '�ݶ�һ��
            End If
        End If
    End If
    
    'ʹ���ȼ����вɼ�
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
    If HintError(err, "timerCapture_Timer", False) = 1 Then Resume

End Sub


Private Sub timerHelper_Timer()
On Error GoTo errhandle:
    Call AutoHideHelper(mstrSelTabName)

    timerHelper.Enabled = False
Exit Sub
errhandle:
    timerHelper.Enabled = False
    If HintError(err, "timerHelper_Timer", False) = 1 Then Resume
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
    Exit Sub
errhandle:
    If HintError(err, "timerRefresh_Timer", False) = 1 Then Resume
End Sub


Private Sub ChangeUser()
    Dim strPrivs As String
    Dim strUserID As String
    
'TODO:��Ҫ����
    frmTwoUser.intDBState = mintChangeUserState
    frmTwoUser.strUserNameHIS = mstrHisUserName
    frmTwoUser.strUserIDHIS = mstrHisUserID
    frmTwoUser.Show 1, Me
    
    If frmTwoUser.blnOk = True Then
        If frmTwoUser.intDBState = 1 Then   'ͳһ����ָ���HISԭ�������ݿ����Ӻ��û���
            mstrOtherUserName = mstrHisUserName
            mstrOtherUserID = mstrHisUserID
            
            mblnCnOracleIsHIS = True
            mintChangeUserState = 1
            Set gcnOracle = mcnOracleHIS
            
            InitCommon gcnOracle
            
            SetDbUser mstrHisUserID
'            RegCheck
            Call GetUserInfo
'            strPrivs = ";" & GetPrivFunc(100, mlngModule) & ";"      'Ӱ��ɼ�����վ
            
'            Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
'            Call mobjWork_Report.zlInitModule(Me, mlngModule, strPrivs, mlngCur����ID, Nothing, ucPacsHelper1)
        ElseIf frmTwoUser.intDBState = 2 Then   '�������򽻻����ݿ�����
            '�����ʹ�������ݿ����ӣ��ȼ��Ȩ��
            mstrOtherUserName = frmTwoUser.strUserNameNew
            mstrOtherUserID = frmTwoUser.strUserIDNew
            
            mintChangeUserState = 2
            If frmTwoUser.blnCnOracleIsNew = True Then
                Set gcnOracle = frmTwoUser.cnOracle
                mblnCnOracleIsHIS = False
                
                '��ʼ��zlComLib������ȷ��GetPrivFunc��ȡ������ȷ����Ϣ
                InitCommon gcnOracle
'                RegCheck
                SetDbUser mstrOtherUserID
                
                '�����û�Ȩ��
                strPrivs = GetPrivFunc(100, mlngModule)       'Ӱ��ɼ�����վ
                If strPrivs = "" Then
                    HintMsg "�㲻�߱�ʹ�á�Ӱ��ɼ�����վ��ģ���Ȩ�ޣ�", "ChangeUser", vbInformation
                    
                    '�л���ԭ�����û�
                    Set gcnOracle = mcnOracleHIS
                    
                    InitCommon gcnOracle
'                    RegCheck
                    SetDbUser mstrHisUserID
                
                    mstrOtherUserName = mstrHisUserName
                    mstrOtherUserID = mstrHisUserID
                    mblnCnOracleIsHIS = True
                    mintChangeUserState = 1
                    
                    Exit Sub
                End If
                
                strPrivs = GetPrivFunc(100, 1258)       '���Ʊ������
                If strPrivs = "" Then
                    HintMsg "�㲻�߱�ʹ�á����Ʊ��桱ģ���Ȩ�ޣ�", "ChangeUser", vbInformation
                    
                    '�л���ԭ�����û�
                    Set gcnOracle = mcnOracleHIS
                    
                    InitCommon gcnOracle
'                    RegCheck
                    SetDbUser mstrHisUserID
                    
                    mstrOtherUserName = mstrHisUserName
                    mstrOtherUserID = mstrHisUserID
                    mblnCnOracleIsHIS = True
                    mintChangeUserState = 1
                    
                    Exit Sub
                End If
            Else
                Set gcnOracle = mcnOracleHIS
                
                InitCommon gcnOracle
'                RegCheck
                SetDbUser mstrHisUserID
                
                mblnCnOracleIsHIS = True
            End If
            
            Call GetUserInfo
'            strPrivs = ";" & GetPrivFunc(100, mlngModule) & ";"       'Ӱ��ɼ�����վ
            
'            Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
'            Call mobjWork_Report.zlInitModule(Me, mlngModule, strPrivs, mlngCur����ID, Nothing, ucPacsHelper1)
        End If
        
    End If
    
    If mblnCnOracleIsHIS Then
        Me.stbThis.Panels(4).Text = "����ҽ����" & mstrHisUserName & "   ���ҽ����" & mstrOtherUserName
    Else
        Me.stbThis.Panels(4).Text = "����ҽ����" & mstrOtherUserName & "   ���ҽ����" & mstrHisUserName
    End If
End Sub

Private Sub SwitchUser()
'��ȡ���û�Ȩ��˵����ʹ�� GetPrivFuncByUser ���ұ�֤strDBUser������gstrDBUser��һ���������õ���¼�û�Ȩ�ޣ����� GetPrivFuncByUser ��Ҫ����SetDbUser ֮ǰ
'���� InitCommon ��ִ�� SetDbUser
'����114781�Ķ��㣺�޸��ж��Ƿ��л������û����߼����л��û�������mstrPrivs��ֵ����
    Dim strPrivs As String
 
    Call frmSwitchUser.SetModule(mlngModule)
    frmSwitchUser.Show 1, Me

    If frmSwitchUser.blnOk = False Then Exit Sub
    
'   �����ʹ�������ݿ����ӣ��ȼ��Ȩ��
    mstrOtherUserName = frmSwitchUser.strUserNameNew
    mstrOtherUserID = frmSwitchUser.strUserIDNew

    Set gcnOracle = frmSwitchUser.mcnOracle
    mblnCnOracleIsHIS = False

    If gstrDBUser <> mstrOtherUserID Then

        mstrPrivs = ";" & GetPrivFuncByUser(100, mlngModule, mstrOtherUserID) & ";"
        
        InitCommon gcnOracle
        gstrDBUser = mstrOtherUserID
        
        Call ReCreatCbrMenu(cbrMain)
        
        Call GetUserInfo
    
'        If Not gobjRichEPR Is Nothing Then Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
'        If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.zlInitModule(Me, mlngModule, mstrPrivs, mlngCur����ID, Nothing, ucPacsHelper1)
        
        Me.stbThis.Panels(4).Text = "����ҽ����" & mstrOtherUserName & "   ���ҽ����" & mstrOtherUserName

    End If

End Sub

Private Sub Menu_Manage_���()
On Error GoTo errhandle
    Dim strReview As String
    Dim strDeptName As String

    If mobjCurStudyInfo.lngAdviceId = 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_���", vbInformation
        Exit Sub
    End If
    
    strDeptName = Split(mstrCur����, "-")(1)
    If frmReview.ShowMe(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, Me, strDeptName, strReview) = True Then
            
        mobjCurStudyInfo.strFollowUpDescribe = strReview
        Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
        
    End If

Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_���", False) Then Resume
End Sub

Private Sub Menu_Manage_���淢��()
'���淢��
On Error GoTo errhandle
    Dim strSQL As String

    If mobjCurStudyInfo.lngAdviceId = 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_���淢��", vbInformation
        Exit Sub
    End If
    
 
    strSQL = "Zl_Ӱ�񱨸淢��(" & mobjCurStudyInfo.lngAdviceId & ",'" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "���淢��")

    mobjCurStudyInfo.intReportGiveOut = IIf(mobjCurStudyInfo.intReportGiveOut = 1, 0, 1)
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
  
    
    Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_���淢��") Then Resume
End Sub

Private Sub Menu_Manage_��Ƭ����()
'��Ƭ����
On Error GoTo errhandle
    Dim strSQL As String

    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_��Ƭ����", vbInformation
        Exit Sub
    End If
    
    strSQL = "Zl_Ӱ��Ƭ����(" & mobjCurStudyInfo.lngAdviceId & ",'" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "��Ƭ����")
    
    mobjCurStudyInfo.intFilmGiveOut = IIf(mobjCurStudyInfo.intFilmGiveOut = 1, 0, 1)
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)

    Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_��Ƭ����") Then Resume
End Sub

Private Sub Menu_Manage_���潺Ƭͬʱ����()
'���潺Ƭͬʱ����
On Error GoTo errhandle
    Dim strSQL As String
    

    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_���潺Ƭͬʱ����", vbInformation
        Exit Sub
    End If
 
    If mobjCurStudyInfo.intReportGiveOut = 1 And mobjCurStudyInfo.intFilmGiveOut = 1 Then
        strSQL = "Zl_Ӱ�񱨸淢��(" & mobjCurStudyInfo.lngAdviceId & ",'" & UserInfo.���� & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "���淢��")
        
        strSQL = "Zl_Ӱ��Ƭ����(" & mobjCurStudyInfo.lngAdviceId & ",'" & UserInfo.���� & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "��Ƭ����")
        
        mobjCurStudyInfo.intReportGiveOut = 0
        mobjCurStudyInfo.intFilmGiveOut = 0
    Else
        strSQL = "Zl_Ӱ�񱨸淢��(" & mobjCurStudyInfo.lngAdviceId & ",'" & UserInfo.���� & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "���淢��")
    
        strSQL = "Zl_Ӱ��Ƭ����(" & mobjCurStudyInfo.lngAdviceId & ",'" & UserInfo.���� & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "��Ƭ����")
        
        mobjCurStudyInfo.intReportGiveOut = 1
        mobjCurStudyInfo.intFilmGiveOut = 1
        
    End If
    
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)

Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_���潺Ƭͬʱ����") Then Resume
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
        
        stbThis.Panels(4).Text = "����ҽ����" & strRPTExecutor & "   ���ҽ����" & Split(stbThis.Panels(4).Text, "���ҽ����")(1)
    End If
    
    Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_ReportExecutor") Then Resume
End Sub

Private Function Menu_Manage_SendAudit(ByVal lngAdviceId As Long, ByVal strName As String) As Boolean
    Dim strSQL As String
    Dim lngCurAdviceId As Long

    On Error GoTo errhandle
    
    Menu_Manage_SendAudit = False
    
    lngCurAdviceId = lngAdviceId
    If lngCurAdviceId = 0 Then lngCurAdviceId = mobjCurStudyInfo.lngAdviceId
    
    If lngCurAdviceId > 0 Then
        strSQL = "Zl_Ӱ�����¼_�����������(" & lngCurAdviceId & ",'" & strName & "')"
        zlDatabase.ExecuteProcedure strSQL, "�����������"
        
        If Len(Trim(strName)) > 0 Then
            HintMsg "�ɹ����͵�����ˡ�" & strName & "����", "Menu_Manage_SendAudit", vbInformation
        End If
    Else
        HintMsg "����ѡ��һ����顣", "Menu_Manage_SendAudit", vbInformation
        Exit Function
    End If
    
    Menu_Manage_SendAudit = True
    
    'ͬ��ˢ�¼���б�
    Call UpdateQueryListData(Nothing, lngCurAdviceId)
    
    Exit Function
errhandle:
    If HintError(err, "Menu_Manage_SendAudit") Then Resume
End Function



Private Function GetStudyNumberDisplayName() As String
'��ȡ��������ʾ����
    GetStudyNumberDisplayName = IIf(mlngModule = G_LNG_PATHOLSYS_NUM, "�����", "����")
End Function

Private Function GetScanRequestCount(ByVal lngAdviceId As Long) As Long
'��ȡɨ�����뵥������
On Error GoTo errhandle
    Dim lngCount As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    GetScanRequestCount = 0
    
    If lngAdviceId <= 0 Then Exit Function
    
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
    If HintError(err, "GetScanRequestCount") Then Resume
End Function

Private Function GetStudyReportType(objStudyInfo As clsStudyInfo) As Long
'��ȡ��ǰ��鱨������
    GetStudyReportType = -1
    If objStudyInfo Is Nothing Then Exit Function
    If objStudyInfo.lngExeDepartmentId = mlngCur����ID Then Exit Function
    
    GetStudyReportType = GetDeptPara(objStudyInfo.lngExeDepartmentId, "����༭��", 0)
End Function

Private Sub RefreshModuleData(ByVal strTabName As String, ByVal strWorkModuleTag As String, objSelModule As Object)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ܣ�ˢ��TABҳ��
'������
'       blnRefresh ��ɺ�ȡ�������֪ͨPACS����༭��ˢ��
'''''''''''''''''''''''''''''''''''''''''''''''''
'�ù���ֻ���Ƕ��ʽģ��͸����ɼ�ģ����и��£������絯��ʽ����༭ģ���򲻽���ˢ�²���
    Dim blnIsForceRrfresh As Boolean

On Error GoTo errhandle
    
    If objSelModule Is Nothing Then Exit Sub
    If mobjCurStudyInfo Is Nothing Then Exit Sub
    
    blnIsForceRrfresh = mblnIsForceRefresh
    
    '�����ǰ������ʷ���ݲ鿴״̬����ˢ��ģ������ʱ����Ҫ����ǿ������ˢ�£�������ʷ�����뵱ǰѡ��ļ������һ��ʱ��ģ�鲻ִ��ˢ�²���
    If mblnIsHistoryMode Then blnIsForceRrfresh = True
    If mstrSelModuleTag <> strWorkModuleTag Then blnIsForceRrfresh = True   '������ͬģ�鵥Ԫ�ڣ����ڼ����Ϣ������ͬ���������ģ�鲻��ˢ�£���clsWorkModule_HisV2�����˷��ã�ҽ����������ģ�飬�ҹ�����mobjstudyinfo����
 
    Call SetHistoryViewState(False)
    
    If Not mobjWork_Pathol Is Nothing Then
        Call mobjWork_Pathol.zlRefresh(mobjCurStudyInfo, strWorkModuleTag, blnIsForceRrfresh)
    End If
    
    '����PacsHelper,�ȶ�PacsHelper���и��£��Ա�����ж��Ƿ�����˹������ͼ��
    If Not (dkpMain.Panes(2).hidden = False And ucPacsHelper1.Visible = False) Then
        If strTabName <> C_TAB_NAME_��鱨�� Then
            ucPacsHelper1.AllowWrite = False
        End If
        
        Call ucPacsHelper1.zlRefresh(mobjCurStudyInfo, 0, blnIsForceRrfresh)
    End If
    
        
    Select Case strTabName
        Case C_TAB_NAME_Ӱ��ͼ��
            Call objSelModule.zlRefreshFace(mobjCurStudyInfo, blnIsForceRrfresh)
            
        Case C_TAB_NAME_�걾����, C_TAB_NAME_����ȡ��, C_TAB_NAME_������Ƭ, C_TAB_NAME_�����ؼ�, C_TAB_NAME_���̱���
            Call objSelModule.zlUpdateAdviceInf(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, mobjCurStudyInfo.blnMoved)
            Call objSelModule.zlRefreshFace(blnIsForceRrfresh)
            
        Case C_TAB_NAME_��鱨��
            Call mobjWork_Report.zlRefreshFace(mobjCurStudyInfo, strWorkModuleTag, blnIsForceRrfresh)
            
        Case C_TAB_NAME_���ü�¼, C_TAB_NAME_ҽ����¼, C_TAB_NAME_������¼, C_TAB_NAME_���Ӳ���
            Call mobjWork_His.zlModule.zlRefresh(mobjCurStudyInfo, strWorkModuleTag, blnIsForceRrfresh)
            
        Case C_TAB_NAME_Ӱ��ɼ�
            '�����ǰ��Ƶ�ɼ���Ƕ�뵽����ʽ����༭����ʱ����ִ�к�����������������Ƕ��ʽ��Ƶ�ɼ����л�
            If Not mobjCapLinker Is Nothing And Not mobjWork_ImageCap Is Nothing And VideoIsAttachReportWindow = False Then
                Call objSelModule.zlRefreshVideoWindow
                Call objSelModule.zlRestoreWindow(IIf(mobjCurStudyInfo.intStep >= 2 And mobjCurStudyInfo.intStep < 5, False, True), True)
            End If
            
        Case C_TAB_NAME_�Ŷӽк�
'            Call EmbedWindow(mobjQueue.hwnd)
            
    End Select
    
    If Not mobjWork_PacsImg Is Nothing Then
        If mobjWork_PacsImg.AdviceId <> mobjCurStudyInfo.lngAdviceId Then Set mobjWork_PacsImg.StudyInfo = mobjCurStudyInfo
    End If
    
    If Not mobjWork_Report Is Nothing Then
        If mobjWork_Report.AdviceId <> mobjCurStudyInfo.lngAdviceId Then Set mobjWork_Report.StudyInfo = mobjCurStudyInfo
    End If

    
    '����CapLinker��������
    If Not mobjCapLinker Is Nothing Then mobjCapLinker.MainAdviceId = mobjCurStudyInfo.lngAdviceId
     
'    '����PacsHelper
'    If Not (dkpMain.Panes(2).hidden = False And ucPacsHelper1.Visible = False) Then
'        If strTabName <> C_TAB_NAME_��鱨�� Then
'            ucPacsHelper1.AllowWrite = False
'        End If
'
'        Call ucPacsHelper1.zlRefresh(mobjCurStudyInfo, 0, blnIsForceRrfresh)
'    End If
    
     
    If mobjCurStudyInfo.lngAdviceId <> 0 Then
        '��ʾ�ɴ�ӡ�����Ƶ���:֮���Լ�ʱ����,��Ϊ��ʹ��F2�ȼ�
        Call ShowBillList(cbrMain.FindControl(, conMenu_Manage_RequestPrint, , True))
    End If
    
    '���¸�����Ƶģ��
    Call ResetFloatingVideoState(mobjCurStudyInfo)
    
    
    Exit Sub
errhandle:
    If HintError(err, "RefreshModuleData", False) = 1 Then Resume
End Sub

'Private Sub ResetFloatingReportState(objStudyInfo As clsStudyInfo)
'    Dim objForm As Object
'
'    '���µ�������ģ��
'    If Not mobjWork_Report Is Nothing Then
'        Call mobjWork_Report.RefreshPopupWindow(objStudyInfo)
'    Else
'        For Each objForm In Forms
'            If TypeOf objForm Is frmReportV2 Then
'                If objForm.AdviceId = objStudyInfo.lngAdviceId And objForm.IsLinkHelper = False Then
'                    Call objForm.zlRefresh(objStudyInfo)
'
'                    Exit Sub
'                End If
'            End If
'        Next
'    End If
'
'End Sub

Private Sub ResetFloatingVideoState(objStudyInfo As clsStudyInfo)
'���踡���ɼ�����״̬
    If mobjWork_ImageCap Is Nothing Then Exit Sub
    
    If mobjWork_ImageCap.VideoDockState = False Then Exit Sub
    
    
    If mobjWork_ImageCap.isLock Or mobjWork_ImageCap.IsAfter Then
        Call mobjWork_ImageCap.SetPopupTitle("")
        Exit Sub
    End If
    
    Call mobjWork_ImageCap.SetPopupTitle(objStudyInfo.strPatientName)
    Call mobjWork_ImageCap.zlRestoreWindow(IIf(objStudyInfo.intStep > 1 And objStudyInfo.intStep < 5, False, True), True)
End Sub


Private Sub Menu_Manage_��������()
'��������
On Error GoTo errhandle
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_��������", vbInformation
        Exit Sub
    End If
    
    Call frmReferencePatient.ZlShowMe(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.strPatientName, Me, True, mlngCur����ID)
    
    'ˢ�²����б�
     Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_��������", False) = 1 Then Resume
End Sub


Private Sub Menu_Manage_�����ɼ�()
On Error GoTo errhandle

    If Not GetIsValidOfStorageDevice(mlngCur����ID) Then
      HintMsg "Ӱ��洢�豸δ�������ͣ�ã����飡", "Menu_Manage_�����ɼ�", vbInformation
      Exit Sub
    End If
    
    If mobjWork_ImageCap Is Nothing Then
        Call VerifyModuleObj(C_TAB_NAME_Ӱ��ɼ�)
    End If
    
    If Not mobjWork_ImageCap Is Nothing Then
        mobjCapLinker.ReportAdviceId = 0
        
        Call mobjWork_ImageCap.zlShowPopupVideo(IIf(mobjCurStudyInfo.intStep > 1 And mobjCurStudyInfo.intStep < 5, False, True))
        
        If mobjWork_ImageCap.VideoDockState Then
            Call mobjWork_ImageCap.SetPopupTitle(mobjCurStudyInfo.strPatientName)
        End If
    End If
    
Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_�����ɼ�", False) = 1 Then Resume
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
            HintMsg "���ܴ�����¼�������ڰ�װIMAPI2��¼��������½��롣", "Menu_Manage_ͼ���¼", vbInformation
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
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_��������", vbInformation
        Exit Sub
    End If
    
    If InStr(";" & GetPrivFunc(100, 1259) & ";", ";����;") = 0 Then
        HintMsg "��û�в��ĵ��Ӳ�����Ȩ�ޣ�����ϵ����Ա��", "Menu_Manage_��������", vbInformation
        Exit Sub
    End If
    
    Set mobjMedicalRecord = Nothing
    If mobjMedicalRecord Is Nothing Then
        Set mobjMedicalRecord = DynamicCreate("zlPublicAdvice.clsPublicAdvice", "zlPublicAdvice")
        If mobjMedicalRecord Is Nothing Then Exit Sub
        
        Call mobjMedicalRecord.InitCommon(gcnOracle, glngSys, gstrNodeNo, gfrmMain, glngModul, gstrPrivs, mobjMsgCenter.Msg)
        
        If mobjCurStudyInfo.lngPageID <= 0 Then
            HintMsg "�ò�����δ����������", "Menu_Manage_��������", vbInformation
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
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_�ղص�", vbInformation
        Exit Sub
    End If
    
    gstrSQL = "select �״�ʱ�� from ����ҽ������ where ҽ��ID= " & lngAdviceId & ""
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    '�ж�ѡ�м�¼�Ƿ񱨵������û�б������ܽ����ղز���
    Do While Not rsTemp.EOF
        If NVL(rsTemp!�״�ʱ��) = "" Then
            HintMsg "�ü��δ�����������ղأ�", "Menu_Manage_�ղص�", vbOKOnly
            Exit Sub
        End If
        
        rsTemp.MoveNext
    Loop
    
    Call frmToCollection.ShowToCollectionWind(Me, lngAdviceId, lngSendNo)
    
    Set mobjCurStudyInfo = GetBaseInfo(lngAdviceId, intMovedState + 1)
    
    If mobjCurStudyInfo.lngPatientFrom = 1 Then
        If mobjCurStudyInfo.strMarkNum > 0 Then labCollectionInfo = "��:" & mobjCurStudyInfo.strMarkNum & "  "
    ElseIf mobjCurStudyInfo.lngPatientFrom = 2 Then
        If mobjCurStudyInfo.strMarkNum > 0 Then labCollectionInfo = "ס:" & mobjCurStudyInfo.strMarkNum & "  "
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
    If HintError(err, "Menu_Manage_�ղ�������ʾ", False) = 1 Then Resume
End Sub

Private Sub Menu_Petition_ɨ�����뵥(ByVal intType As Integer)
'intType:0--�鿴���뵥��1--ɨ�����뵥
    Dim objPetitionCap As frmPetitionCapture                  '���뵥
On Error GoTo errFree
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim strPatientDepartment As String
    Dim lngDepID As Long
     
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Petition_ɨ�����뵥", vbInformation
        Exit Sub
    End If
    
    lngDepID = IIf(mlngCur����ID = 0, mobjCurStudyInfo.lngExeDepartmentId, mlngCur����ID)
    With mobjCurStudyInfo
        strSQL = "Select ���� From ���ű� Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˿���", .lngPatDept)
        
        strPatientDepartment = ""
        If rsTemp.RecordCount > 0 Then strPatientDepartment = NVL(rsTemp!����)
    
        Set objPetitionCap = New frmPetitionCapture
        Call objPetitionCap.ShowPetitionCaptureWind( _
                                    mstrPrivs, _
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
        
        If .lngAdviceId > 0 Then Call UpdateQueryListData(Nothing, .lngAdviceId)
    End With
errFree:
    Unload objPetitionCap
    Set objPetitionCap = Nothing
End Sub

'Private Sub Menu_Manage_SetXWParam_click()
''------------------------------------------------
''���ܣ�������PACS�Ĳ������ô���
''���أ�
''------------------------------------------------
'    On Error GoTo err
'
'    Call frmXWSetParams.zlShowMe(Me)
'
'    Exit Sub
'err:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Sub


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
    If HintError(err, "conMenu_File_SendImg_click", False) = 1 Then Resume
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
    
            intExeTime = Val(NVL(rsTemp!�Զ�ִ��ʱ��))
            
            If intExeTime > 0 Then
                strVBS = NVL(rsTemp!VBS�ű�)
                
                mintInterfaceCount = mintInterfaceCount + 1
                mintInterface(mintInterfaceCount).intID = mintInterfaceCount
                mintInterface(mintInterfaceCount).strVBS = strVBS
                mintInterface(mintInterfaceCount).intExeTime = intExeTime
                mintInterface(mintInterfaceCount).strName = NVL(rsTemp!������) & "-" & NVL(rsTemp!������)
            End If
            
            Call rsTemp.MoveNext
        Wend
    End If
        
    Exit Sub
errH:
    If HintError(err, "initInterface<�����ʼ��>") = 1 Then Resume
End Sub

Private Sub ExecutePluginInterface(ByVal intTimeType As Integer, Optional ByVal lngTimeTag As Long = 0, _
    Optional ByVal strAttachPar1 As String, Optional ByVal strAttachPar2 As String, Optional ByVal strAttachPar3 As String)
'���ܣ�����ʱ���Ƿ�����Ҫ�Զ�ִ�еĲ������
'intTime:ִ��ʱ��
'On Error GoTo errH

    Dim i As Integer
        
    If mintInterfaceCount <= 0 Then Exit Sub
    
    For i = 1 To mintInterfaceCount
        If mintInterface(i).intExeTime = intTimeType Then
            Call ExecutePluginInterfaceFun(mintInterface(i).strName, mintInterface(i).strVBS, lngTimeTag, strAttachPar1, strAttachPar2, strAttachPar3)
        End If
    Next

    Exit Sub
'errH:
'    err.Raise -1, , err.Description
'    MsgBoxD Me, "���[" & mintInterface(i).strName & "]ִ���쳣��������Ϣ��" & err.Description, vbInformation, Me.Caption
'    err.Clear
End Sub

Private Function ChechHaveTlbinf32() As Boolean
On Error GoTo errhandle
    Dim objtest As Object
    
    ChechHaveTlbinf32 = False
    Set objtest = CreateObject("TLI.TLIApplication")
    
    If Not objtest Is Nothing Then ChechHaveTlbinf32 = True
    
    Set objtest = Nothing
Exit Function
errhandle:
    ChechHaveTlbinf32 = False
    HintError err, "ChechHaveTlbinf32", False
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
    
    
    With TabExtra
        .RemoveAll
 
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ButtonMargin.Top = 0
        .PaintManager.ButtonMargin.Bottom = 0
        .PaintManager.ShowIcons = True
        .RemoveAll
        
        .InsertItem 1, "���ݼ���", picDataSearchContainer.hwnd, 0
        .Item(TabExtra.ItemCount - 1).tag = "���ݼ���"
        
        .InsertItem 2, "������Ϣ", picExtra.hwnd, 0
        .Item(TabExtra.ItemCount - 1).tag = "������Ϣ"

        
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
            Call mobjPacsQueryWrap.ExecuteMenu(lngShemeNo)
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
    HintError err, "ExecuteDefaultQueryScheme<ִ��Ĭ�Ϸ���>", False
End Sub

Public Sub UpdateQueryListData(ByRef rsData As Recordset, ByVal lngAdviceId As Long, _
    Optional ByVal intSyncDataType As Integer = SyncDataType.rsDataAndrsShow, Optional ByVal blnNoRefreshModule As Boolean = False)
'���²�ѯ�б�ĳһ������
'ͬʱ����¸��л������ݣ�ע��Ҫ���жϸ������Ƿ��ǵ�ǰѡ����
'blIsAdd �Ƿ���������
'lngAdviceID�仯�е�ҽ��ID
'blRaiseEventSelChange �Ƿ񴥷��б�selchange�¼�
On Error GoTo errH
    If Not mobjPacsQueryWrap Is Nothing Then Call mobjPacsQueryWrap.UpdateRow(rsData, lngAdviceId, intSyncDataType, blnNoRefreshModule)
    
    Exit Sub
errH:
    HintError err, "UpdateQueryListData<�����б���>", False
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


Private Sub ucPacsHelper1_OnDockHideClick()
    dkpMain.Panes(2).hidden = Not dkpMain.Panes(2).hidden
End Sub

Private Sub ucPacsHelper1_OnLinkHistoryView(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean, ByVal blnIsDBClick As Boolean)
'��ʷ���������鿴
    Dim objStudyInfo As clsStudyInfo
    Dim strCurModuleTag As String
    
On Error GoTo errhandle
    If mobjSelModule Is Nothing Then Exit Sub
    
    If lngAdviceId <> 0 And blnIsDBClick Then
        Select Case mstrSelTabName
            Case C_TAB_NAME_Ӱ��ͼ��
                If Not mobjWork_PacsImg Is Nothing Then
                    If mobjWork_PacsImg.AdviceId = lngAdviceId Then lngAdviceId = 0
                End If
            
            Case C_TAB_NAME_�걾����, C_TAB_NAME_����ȡ��, C_TAB_NAME_������Ƭ, C_TAB_NAME_�����ؼ�, C_TAB_NAME_���̱���
                If Not mobjWork_Pathol Is Nothing Then
                    If mobjWork_Pathol.AdviceId = lngAdviceId Then lngAdviceId = 0
                End If
                
            Case C_TAB_NAME_��鱨��
                If Not mobjWork_Report Is Nothing Then
                    If mobjWork_Report.AdviceId = lngAdviceId Then lngAdviceId = 0
                End If
                
            Case C_TAB_NAME_���ü�¼, C_TAB_NAME_ҽ����¼, C_TAB_NAME_������¼, C_TAB_NAME_���Ӳ���
                If Not mobjWork_His Is Nothing Then
                    If mobjWork_His.zlModule.AdviceId = lngAdviceId Then lngAdviceId = 0
                End If
        End Select
    End If
    
    If lngAdviceId = mobjCurStudyInfo.lngAdviceId Then
        lngAdviceId = 0
    End If
    
    If lngAdviceId <> 0 Then
        Set objStudyInfo = GetBaseInfo(lngAdviceId)
    Else
        Set objStudyInfo = mobjCurStudyInfo
    End If
    
    Select Case mstrSelTabName
        Case C_TAB_NAME_Ӱ��ͼ��
            If Not mobjWork_PacsImg Is Nothing Then
                Call mobjWork_PacsImg.zlRefreshFace(objStudyInfo, mblnIsForceRefresh, IIf(lngAdviceId = 0, False, True))
            End If
            
        Case C_TAB_NAME_�걾����, C_TAB_NAME_����ȡ��, C_TAB_NAME_������Ƭ, C_TAB_NAME_�����ؼ�, C_TAB_NAME_���̱���
            If Not mobjWork_Pathol Is Nothing Then
                Call mobjWork_Pathol.zlRefresh(objStudyInfo, mstrSelTabName, mblnIsForceRefresh, IIf(lngAdviceId = 0, False, True))
            End If
            
'            Call mobjSelModule.zlUpdateAdviceInf(objStudyInfo.lngAdviceId, objStudyInfo.lngSendNo, objStudyInfo.intStep, objStudyInfo.blnMoved)
'            Call mobjSelModule.zlRefreshFace(mblnIsForceRefresh)
            
        Case C_TAB_NAME_��鱨��
            '�жϱ��������Ƿ��뵱ǰ��ͬ
            strCurModuleTag = GetWorkModuleName(mstrSelTabName, objStudyInfo.lngExeDepartmentId, objStudyInfo.lngPatientFrom)
            If strCurModuleTag <> mstrSelModuleTag Then
               Call SelectModule(mstrSelTabName, strCurModuleTag)
               TabWindow.Selected.tag = strCurModuleTag
            End If
    
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(objStudyInfo, mstrSelModuleTag, mblnIsForceRefresh, IIf(lngAdviceId = 0, False, True))
            End If
            
        Case C_TAB_NAME_���ü�¼, C_TAB_NAME_ҽ����¼, C_TAB_NAME_������¼, C_TAB_NAME_���Ӳ���
            If Not mobjWork_His Is Nothing Then
                Call mobjWork_His.zlModule.zlRefresh(objStudyInfo, mstrSelModuleTag, mblnIsForceRefresh, IIf(lngAdviceId = 0, False, True))
            End If
            
    End Select
    
    Call SetHistoryViewState(IIf(lngAdviceId <> 0, True, False))
Exit Sub
errhandle:
    HintError err, "ucPacsHelper1_OnTabChanged", False
End Sub


Private Sub SetHistoryViewState(ByVal blnIsHistory As Boolean)
    Dim strCap As String
    
    strCap = TabWindow.Selected.Caption
     
    TabWindow.Selected.Caption = ""
    TabWindow.PaintManager.ColorSet.SelectedText = IIf(blnIsHistory, &HC0&, vbBlack)
    TabWindow.Selected.Caption = strCap
    
    mblnIsHistoryMode = blnIsHistory
    
    '���ܸı�caption�����ݣ���Ϊ���ֹ�����Ҫ����caption�����ж�
'    TabWindow.Selected.Caption = Replace(TabWindow.Selected.Caption, C_HISTORY_VIEW_TAG, "") & IIf(blnIsHistory, C_HISTORY_VIEW_TAG, "")
End Sub


Private Sub ucPacsHelper1_OnTabChanged(ByVal strTabName As String)
On Error GoTo errhandle
    '�ж��Ƿ���Ҫ�ָ����ʾ�ģ��ҳ��ʾ�����tagΪ"�ʾ�"�����л�����鱨��ģ��ʱ����Ҫ�ָ����ʾ���ģ��ҳ
    If TabWindow.Selected.Caption = "��鱨��" Then
        If strTabName = "�ʾ�" Then
            ucPacsHelper1.tag = "�ʾ�"
        Else
            ucPacsHelper1.tag = ""
        End If
    End If
Exit Sub
errhandle:
    HintError err, "ucPacsHelper1_OnTabChanged", False
End Sub

 

Private Sub vsfList_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
On Error GoTo errhandle
    Dim blnIsProcessing As Boolean
    
    blnIsProcessing = ucPacsHelper1.Processing
    
    If blnIsProcessing = False Then
        If Not mobjWork_Report Is Nothing Then
            blnIsProcessing = mobjWork_Report.Processing
        End If
    End If
    
    If blnIsProcessing Then
        HintMsg "����δ��ɵ�������ȴ���ɺ����Բ�����", "vsfList_BeforeSelChange", vbOKOnly
        Cancel = True
    End If
Exit Sub
errhandle:
    HintError err, "vsfList_BeforeSelChange", False
End Sub

Private Sub vsfList_DblClick()
On Error GoTo errhandle
    Call VsfListDbClick(False)
Exit Sub
errhandle:
    HintError err, "vsfList_DblClick", False
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
            Set cbrPopControl = CreateMenu(objControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_SendAudit * 10# + i, rsTemp!����, "", 0, False)
        End If
        rsTemp.MoveNext
    Next
    Exit Sub
errH:
    If HintError(err, "CreateAuditorMenu") = 1 Then Resume
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
    Call HintError(err, "Menu_Manage_���ԤԼ", False)
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
    Call HintError(err, "Menu_Manage_ԤԼ����", False)
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
            mstrSelQueueRooms = mstrSelQueueRooms & NVL(rsData!����) & "-" & NVL(rsData!ִ�м�)
            rsData.MoveNext
            
        Wend
        
        GetSelQueueRooms = mstrSelQueueRooms
    Else
        GetSelQueueRooms = mobjPacsQueryWrap.SelQueueRooms
    End If
    
    Exit Function
errH:
    If HintError(err, "GetSelQueueRooms") Then Resume
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
            TimerRefresh.Interval = 10000
        End If
        TimerRefresh.Enabled = True
    End If
    Exit Sub
errH:
    HintError err, "InitAutoRefresh", False
End Sub

Public Function GetBaseInfo(ByVal lngAdviceId As Long, Optional intMovedState As Integer = 0) As clsStudyInfo
    Set GetBaseInfo = mobjPacsQueryWrap.GetBaseInfo(lngAdviceId, intMovedState)
    
    GetBaseInfo.lngReportType = Val(GetDeptPara(GetBaseInfo.lngExeDepartmentId, "����༭��", 0)) + 1
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
        lngSendNo = Val(NVL(rsTemp!���ͺ�))
    End If
        
    Call UpdateQueryListData(Nothing, lngAdviceId)
    
    '�������ݿ�����
    strSQL = "ZL_Ӱ�����¼_���Ͱ���(" & lngAdviceId & "," & lngSendNo & ",null,null,null,'" & strRoom & "',1)"
    Call zlDatabase.ExecuteProcedure(strSQL, "����ִ�м�")
    
    Exit Sub
errH:
    If HintError(err, "QueueDataConsistency") Then Resume
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
    Call CreateWorkModuleMenu(mstrSelTabName, mstrSelModuleTag)
    Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
    
    Call LockWindowUpdate(0)
    
    Exit Sub
errH:
    Call LockWindowUpdate(0)
    HintError err, "ReCreatCbrMenu<���ò˵�>", False
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
                Call Menu_RichEPR(conMenu_PacsReport_Write)
            Case -1, 4, 5               '˫���޶�����,�����ʱ�����趨�Ƿ�򿪹�Ƭվ
                Call Menu_RichEPR(conMenu_Edit_Audit)
            Case 6                  '����
                Call Menu_RichEPR(conMenu_File_Open)
        End Select
    End If

Exit Sub
errhandle:
    Call HintError(err, "VsfListDbClick", False)
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
        
        strResult = mobjPacsQueryWrap.ExecuteMenu(lngID)
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
    HintError err, "ViewLinkChecks", True
End Sub
