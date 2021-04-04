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
   Caption         =   "影像工作站"
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
            Key             =   "PACS报到"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":2474
            Key             =   "观片"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":2BEE
            Key             =   "PACS书写"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":3368
            Key             =   "PACS完成"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":3AE2
            Key             =   "PACS查看病人信息"
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
         ToolTipText     =   "显示全部查询条件"
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
         ToolTipText     =   "重置查询条件"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdDo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "查  询"
         Height          =   735
         Left            =   1920
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPacsMainV2.frx":4C21
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "查询"
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
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
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
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
            Key             =   "申请"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":6C31
            Key             =   "报到"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":7CDB
            Key             =   "检查"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":8D85
            Key             =   "书写"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":9E2F
            Key             =   "诊断"
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":AED9
            Key             =   "审核"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":BF83
            Key             =   "完成"
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":D02D
            Key             =   "驳回"
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":E0D7
            Key             =   "拒绝"
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
            Key             =   "复选留空"
            Object.Tag             =   "90000"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":F71B
            Key             =   "复选选中"
            Object.Tag             =   "90001"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":FCB5
            Key             =   "单选留空"
            Object.Tag             =   "90002"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMainV2.frx":103C7
            Key             =   "单选选中"
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
         ToolTipText     =   "定位"
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
         ToolTipText     =   "查找"
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.PictureBox pic主界面遮挡 
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
               Name            =   "宋体"
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
               Name            =   "宋体"
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
         Begin VB.Label LabFlag急诊 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF0000&
            Caption         =   "急"
            BeginProperty Font 
               Name            =   "宋体"
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
               Name            =   "宋体"
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
               Name            =   "宋体"
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
         Begin VB.Label LabFlag传染病状态 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            Caption         =   "传"
            BeginProperty Font 
               Name            =   "宋体"
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
         Begin VB.Label LabFlag危机状态 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF00FF&
            Caption         =   "危"
            BeginProperty Font 
               Name            =   "宋体"
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
         Begin VB.Label LabFlag绿色通道 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000C000&
            Caption         =   "绿"
            BeginProperty Font 
               Name            =   "宋体"
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
         Begin VB.Label LabFlag婴儿 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "婴"
            BeginProperty Font 
               Name            =   "宋体"
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
         Begin VB.Label LabFlag费用 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "记"
            BeginProperty Font 
               Name            =   "宋体"
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
         ToolTipText     =   "并没有什么用"
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
            Name            =   "宋体"
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
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmPacsMainV2.frx":11365
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         DefaultCardType =   "就诊卡"
         IDkindBorderStyle=   1
         IDKindWidth     =   1800
         FindPatiShowName=   0   'False
         HiddenMoseRightKey=   0   'False
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
'Private Const C_HISTORY_VIEW_TAG = "-历"        '历史内容查看标记

Private Const C_LAYOUT_BASEHEIGHTOFTAB As Long = 5000 ' 其他信息5000
Private Const C_LAYOUT_BASEHEIGHTOFDETAILINFO As Long = 800 ' 详细信息基准高度5000

Private Const C_LNG_TAB_MENU_ID = 123456780

Private Const C_STEPIMG_登记 As String = "申请" '
Private Const C_STEPIMG_报到 As String = "报到" '
Private Const C_STEPIMG_检查 As String = "检查" '
Private Const C_STEPIMG_拒绝 As String = "拒绝" '
Private Const C_STEPIMG_驳回 As String = "驳回" '
Private Const C_STEPIMG_审核 As String = "审核" '
Private Const C_STEPIMG_完成 As String = "完成" '
Private Const C_STEPIMG_诊断 As String = "诊断" '
Private Const C_STEPIMG_书写 As String = "书写" '


'样式样式
Public Enum TColorStyle
    sBlue = 0   '蓝色
    sAshen = 1  '灰白
    sGray = 2   '灰色
    sBlack = 3  '黑色
    sSys = 4    '系统
End Enum

Private Type TWorkModuleInfo
'工作模块信息
    ModuleName As String
    objModule As Object
    hwnd As Long
    FontSize As Double
    DeptId As Long      '所属科室
    tag As String           '保留标志
End Type

Private mAryWorkModule() As TWorkModuleInfo

Private mobjCurStudyInfo As New clsStudyInfo  '用于操作的检查信息
Private mstrFirstTab As String '首次显示的页面
Private mlngMove As Long
Private mintQueryState As Integer '查询方案状态  0 未初始化  ，1 正常  ，2 没有任何有效方案   3：没有已经启用的方案
Private mblHistory As Boolean '是否历次检查
Private mblHaveHistory As Boolean '存在历史检查
Private mintAutoRefreshTimer As Integer '自动刷新计时辅助
Private mintAutoRefreshTimerCount As Integer '自动刷新计时辅助

'---------------------------------------------------
Private Const M_STR_MODULE_MENU_TAG As String = "Main"

'当没有数据时，使用此提示信息
Private Const M_STR_HINT_NoSelectData As String = "请选择需要执行的检查数据。"

'闪烁超时信息
Private Type TFlickerInfo
    LngSchemeNo As Long '当前方案号
    strName As String '闪烁字段名 如： "检查过程"
    strInfo As String '详细信息 如"已登记,申请时间,30|已报到,采样时间,40|"
End Type

'系统参数类型定义
Private Type TSystemPar

    '本地参数
    blnLockAfterCall As Boolean                         '是否呼叫后锁定采集
    strFirstTab As String                               '首次显示的页面
    bln直接检查 As Boolean                               '登记后直接进入检查
    blnWriteCapDoctor As Boolean                        '是否在采集图像后，自动把当前用户填写为检查技师
    blnAutoOpenReport As Boolean                        '开始检查自动打开报告
    blnChoosePrintFormat As Boolean                     '是否报到打印前选择格式
    strLocalRoom As String                              '本机执行间名称
    lngImageValid As Long                               '图像校对
    lngAutoImageDays As Long                            '自动打开历史图像的天数范围
    
    '流程参数
    blnCompleteCommit As Boolean                        '审核后无需再次确认
    blnFinallyCompleteCommit As Boolean                 '终审后直接完成
    blnIgnoreResult As Boolean                          '忽略阴阳性 '=true 忽略
    
    blnReportWithImage As Boolean                       '有图像才能写报告，无图像不可写报告
    blnNoSignFinish As Boolean                              '允许未签名报告打印完成
    blnReportWithResult As Boolean                      '有阴阳性结果才能写报告
    
    blnPrintCommit As Boolean                           '打印后直接完成
    blnCanPrint As Boolean                              '平诊需要审核才能打印 =true
    blnAuditAutoPrint As Boolean                        '终审后直接打印
    lngBeforeDays As Long                               '默认查询的天数
    blnUseQueue As Boolean                              '是否启用排队叫号
    blnSynStudylist As Boolean                          '排队叫号时，点击排队列表或呼叫列表数据后，是否同步定位到检查列表
    blnAutoInQueue As Boolean                           '启用排队叫号后，是否自动入队
    blnQueueQuick As Boolean                            '启用排队叫号后，是否自动弹出快捷叫号窗口
    
    blnRelatingPatient As Boolean                       '是否启用关联病人
    lngConformDetermine As Long                         '符合情况
    strImageLevel As String                             '影像质量等级串
    strReportLevel As String                            '报告质量等级串
    lngImageLevel As Long                               '影像质量判定
    lngReportLevel As Long                              '报告质量判定
    
    lngHintType As Long                                 '诊断结果提示类型
    
    blnIsPetitionScan As Boolean                        '是否启用申请单扫描
    blnChangeUser As Boolean                            '是否启用用户交换
    blnSwitchUser As Boolean                            '是否启用用户切换
    
    lngVideoStationMoneyExeModle As Long                '采集费用执行模式 0-报到时执行，1-检查时执行，2-报告时执行
    lngPacsStationMoneyExeModle As Long                 '医技费用执行模式 0-报到时执行，1-报告时执行
    lngPatholStationMoneyExeModle As Long               '病理费用执行模式 0-报到时执行，1-检查时执行，2-报告时执行
    
    lngListColorMark As Long                            '为0时标记列表前景色，为1时标记列表背景色
    blnNameColColorCfg As Boolean                       '是否根据病人类型设置列表姓名列颜色
    blnOrdinaryNameColColorCfg As Boolean               '缺省类型的病人是否根据病人类型设置姓名颜色
    
    blnAutoSendWorkList As Boolean                      '是否报道时自动发送WorkList
    blnNameFuzzySearch As Boolean                       '是否姓名默认模糊查询
    blnNameQueryTimeLimit As Boolean                    '按姓名过滤时是否进行时间限制
    blnAutoPrint As Boolean                             '报到后自动打印申请单
    blnAutoPrintCheck As Boolean                        '自动避免重复打印
    blnDirectSendRepImg As Boolean                      '直接将观片的图像发送到报告
    
    blnShowImgAfterReport As Boolean                    '报告时观片
    blnIsLocateReport As Boolean
    blnPEISNoCheckMoneyFinish  As Boolean    '体检检查报告完成不判断费用
    blnQuickTabDisplayScheme  As Boolean    '启用快捷tab标签展示方案
    lngReportType As Long
End Type


'视频采集事件信息
Private Type TVideoEventInf
    vetEventType As TVideoEventType
    lngAdviceId As Long
    lngSendNo As Long
    strOtherInf As String
    dcmImage As DicomImage
End Type

'视频采集消息定义
Private Type TCaptureMsgInf
    lngMsg As Long
    lngVirtualKey As Long
    lngScanKey As Long
    lngFlags As Long
End Type


Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long


Private mintInterface() As TInterface   '自动执行的插件
Private mintInterfaceCount As Integer '需要自动执行的插件数量从1 开始计数

Private mintToolBarWriteReg As Integer        '工具栏注册表状态值

Private mstrPrivs As String, mlngModule As Long              '模块号，本模块权限

Private WithEvents mfrmRISRequest As frmRISRequest
Attribute mfrmRISRequest.VB_VarHelpID = -1

'消息处理中心
Private WithEvents mobjMsgCenter As clsPacsMsgProcess
Attribute mobjMsgCenter.VB_VarHelpID = -1

'工作模块的数据刷新模式分三种情况，
'1.工作模块只要存在，强制对其中的数据进行刷新
'2.工作模块在显示时，才对其中的数据进行刷新
'3.工作模块在相关数据变化时且显示的模块是当前模块，才对其中的数据进行刷新

Private mobjWork_PacsImg As frmWork_ImageV2             '影像子窗体
Attribute mobjWork_PacsImg.VB_VarHelpID = -1
Private mobjWork_Pathol As clsWorkModule_PatholV2       '病理相关模块
Private mobjWork_His As clsWorkModule_HisV2             'HIS相关模块
Private mobjWork_Report As clsWorkModule_ReportV2       '报告模块

Private mobjWork_ImageCap As zl9PacsImageCap.clsPacsCaptureV2  '视频采集模块
Private mobjRichReportWrap As frmEPREditWrapV2

Private WithEvents mobjCapLinker As clsCapLinker
Attribute mobjCapLinker.VB_VarHelpID = -1


Private WithEvents mobjPacsCore As zl9PacsCore.clsViewer        '观片站对象
Attribute mobjPacsCore.VB_VarHelpID = -1
Private WithEvents mobjQueue As frmWork_Queue                   '排队叫号
Attribute mobjQueue.VB_VarHelpID = -1



Private mobjSelModule As Object
Private mlngSelHwnd As Long
Private mstrSelTabName As String
Private mstrSelModuleTag As String
Private mobjAppendBill As Object


Private WithEvents mobjPacsQueryWrap As clsPacsQueryWrap      '自定义查询功能封装类
Attribute mobjPacsQueryWrap.VB_VarHelpID = -1

'窗口变量
Private mlngCur科室ID As Long                               '当前科室ID
Private mstrCur科室 As String                               '当前科室 编码-名称
Private mstrCanUse科室 As String                            '当前可用科室  ID_编码-名称

Private mblnInitOk As Boolean   '初始化完成,装载表格
Private mblnAllDepts As Boolean                             '是否选择全部科室
Private mstrCanUse科室IDs As String                         '当前可用的科室ID串，用“，”分隔，可以直接作为SQL查询条件
Private mblnMenuDownState As Boolean                        '避免双击工具栏产生错误
Private mblnIsHasPatholModule As Boolean                   '是否载入了病理模块

Private mblnFormLoadState As Boolean
Private mblnIsScheduleDept As Boolean                       '当前选中科室，是否启用预约
Private mblnIsScheduleOrder As Boolean                      '当前检查是否启用预约，根据预约设备判断

Private mblnIsPrintMode As Boolean                          '是否是清单打印

Private mstrDefaultPatientType As String                    '缺省病人类型

Private mstrRPTExecutor As String                           '保存选择的报告人
Private mblnLockState As Boolean                           '是否有用户处于锁定状态

'流程控制变量
Private mSysPar As TSystemPar                               '系统参数

Private mintImgCount As Integer                             '已扫描申请单数量

Private WithEvents mobjCaptureHot As zl9PacsControl.clsHookKey
Attribute mobjCaptureHot.VB_VarHelpID = -1
Private mVideoEventInf As TVideoEventInf
Private mstrCaptureHot As String                                    '采集热键定义
Private mstrCaptureAfterHot As String                               '后台采集热键定义
Private mstrCaptureAfterTagHot As String                            '标记更新热键定义
Private mCaptureMsg As TCaptureMsgInf
Private mobjSquareCard As Object

'本机本地参数
Private mstrSelQueueRooms As String                         '只处理执行间内的病人
Private mstrAllQueueRooms As String

Private mblnMoved As Boolean                                '当前时间段内是否被转移过
Private mstrWorkModule As String

Private mblnAssignment As Boolean
Private mlngLocateFindType As Long
Private mstrAllExamineRoomCfg As String    '所有科室执行间选择情况
Private mstrCurExamineRoomCfg As String    '当前科室执行间选择情况

'双用户登录
Private mcnOracleHIS As New ADODB.Connection    '记录HIS导航台登陆时使用的数据库联接串
Private mstrHisUserName As String               '记录HIS导航台登陆时使用的用户名
Private mstrHisUserID As String                 '记录HIS导航台登录时使用的用户ID
Private mstrOtherUserName As String             '记录双用户登陆的第二个用户名
Private mstrOtherUserID As String               '记录双用户登录的第二个用户ID
Private mblnCnOracleIsHIS As Boolean            '当前数据库联接是否HIS导航台的连接
Private mintChangeUserState As Integer          '记录用户交换的情况。1- 统一；2-交换

'收藏功能
Private mlngShareFatherID As Long
Private mlngCollectionFatherID As Long
Private mblnIsLoading As Boolean
 
Private mblnIsForceRefresh As Boolean          '是否调用模块强制刷新操作

Private mobjPublicAdvice As Object
Private mobjMedicalRecord As Object
Private mblnIsValid As Boolean                  '窗体界面是否有效

Private mintState As Integer
Private mblnIsHistoryMode As Boolean            '是否历史状态
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



'***********************************************IEventNotify实现***********************************************


Public Function MainWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo errhandle
    Dim strLog As String
    Dim strResult As String
    
    '消息处理
    Select Case uMsg
        Case WM_XWREPORT_IMG
            strLog = Now & " umsg = " & uMsg & ";wparam = " & wParam & ";lparam = " & lParam & vbCrLf
    
            If gblnXWLog Then Call WriteCommLog("XWWindowProc", "XW接口", strLog)
            
            '接收新网发送到系统剪贴板的报告图像
            If lParam <> 0 Then
                If gblnXWLog Then Call WriteCommLog("XWWindowProc", "XW接口", "进入报告图像处理过程。")
    
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
'获取权限串
    IEventNotify_MainPrivs = mstrPrivs
End Function


Public Function IEventNotify_PrintErr(objErr As ErrObject, ByVal lngInfoType As Long, _
    Optional ByVal lngHwnd As Long = 0, Optional ByVal strUnitName As String = "", Optional ByVal strMethodName As String = "") As Long
    IEventNotify_PrintErr = IEventNotify_PrintInfo(objErr.Description, lngInfoType, lngHwnd, strUnitName, strMethodName)
End Function

Public Function IEventNotify_PrintInfo(ByVal strErr As String, ByVal lngInfoType As Long, _
    Optional ByVal lngHwnd As Long = 0, Optional ByVal strUnitName As String = "", Optional ByVal strMethodName As String = "") As Long
'打印错误消息
'样例
'
'测试错误.
'[1290-ZLHIS] [0227 13:36:46] [frmPacsMain.PrintErr]
'
'On Error GoTo errHandle

    Dim strMsg As String
    
    IEventNotify_PrintInfo = 0
    
    strMsg = strErr & vbCrLf & "[" & mlngModule & "-" & UserInfo.用户名 & "] [" & Format(Now, "mmdd hh:mm:ss") & "] [" & strUnitName & "." & strMethodName & "]"
    
    Debug.Print strMsg
    OutputDebugString strMsg
    
    Select Case lngInfoType
        Case infNone
            '不执行任何操作
            
        Case infHint
            If lngHwnd = 0 Then
                MsgBoxD Me, strErr, vbOKOnly, "提示"
            Else
                MsgboxH lngHwnd, strErr, vbOKOnly, "提示"
            End If
            
        Case infWaring
            If lngHwnd = 0 Then
                MsgBoxD Me, strErr, vbOKOnly, "警告"
            Else
                MsgboxH lngHwnd, strErr, vbOKOnly, "警告"
            End If
            
        Case infNormalErr
            If lngHwnd = 0 Then
                MsgBoxD Me, strMsg, vbOKOnly, "错误"
            Else
                MsgboxH lngHwnd, strMsg, vbOKOnly, "错误"
            End If
            
        Case infDataErr
            IEventNotify_PrintInfo = IIf(ErrCenter() = 1, True, False)
            Call SaveErrLog
            
        Case Else
            If lngHwnd = 0 Then
                IEventNotify_PrintInfo = MsgBoxD(Me, strErr, lngInfoType, "提示")
            Else
                IEventNotify_PrintInfo = MsgboxH(lngHwnd, strErr, lngInfoType, "提示")
            End If
            
    End Select
    
'Exit Function
'errHandle:
'    Debug.Print "IEventNotify_PrintErr Err:" & err.Description
End Function


Private Sub IEventNotify_SendRequest(ByVal lngEventNo As Long, Optional ByVal strTag As String = "", _
    Optional data1, Optional data2, Optional data3, Optional strExPro As String = "")
'lngEventNo:事件号
'strTag:事件标记
'data:传递数据
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
        Case WM_LIST_SYNCROW    '同步行显示
            Call UpdateQueryListData(Nothing, curData1)
            
        Case WM_LIST_MOVEUP '上移
            If vsfList.Row > 1 Then vsfList.Row = vsfList.Row - 1
            
        Case WM_LIST_MOVEDOWN '下移
            If vsfList.Row + 1 < vsfList.Rows Then vsfList.Row = vsfList.Row + 1
            
        Case WM_LIST_GETLASTADVICE
            'data1传递当前使用的医嘱ID
            data1 = TraversalAdvice(data1, False, lngSendNo, blnIsMoved)
            data2 = lngSendNo
            data3 = blnIsMoved
        
        Case WM_LIST_GETNEXTADVICE
            'data1传递当前使用的医嘱ID
            data1 = TraversalAdvice(data1, True, lngSendNo, blnIsMoved)
            data2 = lngSendNo
            data3 = blnIsMoved
        
        Case WM_IMG_OPENVIEW
            '观片
            If mobjPacsCore Is Nothing Then
                HintMsg "图像查看对象无效，不能进行此操作。", "cbrMain_Execute", vbOKOnly
                Exit Sub
            End If
            
            If mobjCurStudyInfo.lngAdviceId = Val(curData1) Then
                If mobjCurStudyInfo.strStudyUID <> "" Then
                    Call OpenViewer(1, mobjPacsCore, mobjCurStudyInfo.lngAdviceId, False, Me, "", mobjCurStudyInfo.blnMoved)
                Else
                    '打开关联的检查图像
                    Call OpenLatestImage(Me, mobjPacsCore, mobjCurStudyInfo, mSysPar.lngAutoImageDays)
                End If
            Else
                Call GetSendNo(Val(curData1), strStudyUID)
                
                If Len(strStudyUID) <= 0 Then
                    '打开关联的检查图像
                    Call OpenLatestImage(Me, mobjPacsCore, GetBaseInfo(curData1), mSysPar.lngAutoImageDays)
                Else
                    Call OpenViewer(1, mobjPacsCore, curData1, False, Me)
                End If
            End If
            
        Case WM_IMG_CONTRASTVIEW
            '对比
            If mobjPacsCore Is Nothing Then
                HintMsg "图像查看对象无效，不能进行此操作。", "cbrMain_Execute", vbOKOnly
                Exit Sub
            End If
            
            Call OpenViewer(1, mobjPacsCore, curData1, True, Me)
            
        Case WM_REPORT_VIEW
            '报告预览
            Call ReoprtPrint(curData1, curData2, False, curData3)
            
        Case WM_REPORT_PRINT
            '报告打印
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
    
    strSQL = "Select a.发送号,a.执行过程, b.检查UID, to_Char(B.接收日期,'YYYYMMDD')  as 接收日期 " & _
            " From 病人医嘱发送 a, 影像检查记录 B " & _
            " where a.医嘱ID=b.医嘱ID(+) and a.发送号=b.发送号(+) and a.医嘱Id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询发送号", lngAdviceId)
    
    '由于涉及对转储数据进行观片等操作，因此需要查询转储库的数据
    If rsData.RecordCount <= 0 Then
        '从历史库再进行查询
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
        
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询发送号", lngAdviceId)
    End If
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    strStudyUID = NVL(rsData!检查UID)
    strRecDate = NVL(rsData!接收日期)
    lngStep = Val(NVL(rsData!执行过程))
    
    GetSendNo = Val(NVL(rsData!发送号))
End Function

Private Function GetLastSignInfo(ByVal lngAdviceId As Long, Optional ByRef strCreateUser As String) As TReportSignInfo
'该过程仅支持未转储的数据查询，在resetstate方法中被调用，获取最后签名信息进行回退和审核相关处理
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim reportSignInfo As TReportSignInfo
    Dim lngSendNo As Long
    Dim strStudyUID As String
    Dim dblEPRID As Double
    Dim lngLastVer As Long
    
    reportSignInfo.ID = 0
    
    strCreateUser = UserInfo.姓名
    
    strSQL = "Select a.医嘱id,a.病历id, b.创建人, b.保存人, b.签名级别, b.完成时间, b.最后版本 " _
             & "  From 病人医嘱报告 a, 电子病历记录 b Where a.医嘱id = [1] And a.病历id = b.Id"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "获取签名信息", lngAdviceId)
    If rsData.RecordCount <= 0 Then Exit Function
     
    strCreateUser = NVL(rsData!创建人)
    dblEPRID = Val(NVL(rsData!病历Id))
    lngLastVer = Val(NVL(rsData!最后版本))
             
    '已审核的情况，需要返回签名人。回退内容的情况，保存人不一定是签名人，因此要查找最后一个签名人
    strSQL = "Select ID  From 电子病历内容 Where 文件ID=[1] And 对象类型= 8 And 开始版 = [2] "
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "提取签名级别", dblEPRID, lngLastVer)
    If rsData.RecordCount <= 0 Then
        '判断前次版本是否为签名
        If lngLastVer > 1 Then
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, "提取签名级别", dblEPRID, lngLastVer - 1)
        Else
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, "提取签名级别", dblEPRID, lngLastVer)
        End If
        
        
        If rsData.RecordCount <= 0 Then Exit Function
    End If
    
    Call GetReportSignInfo(Val(NVL(rsData!ID)), reportSignInfo, False)
    
     
    GetLastSignInfo = reportSignInfo
End Function

Public Function IEventNotify_StudyInfo() As clsStudyInfo
'获取检查信息
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
        Case BM_REPORT_EVENT_PRINT  '报告打印事件*****************************************************************************
            'curData1表示医嘱ID
            'curData2表示报告ID 暂未使用
            'curData3编辑器类型 '-1表示病历编辑器
            
            '如果strTag=0, curData3则表示返回是否允许执行打印操作
            
            If Val(curData3) = -1 Then
                If mSysPar.blnIgnoreResult = False Then Call ReportResultHint(curData1)
                
                '如果是转储数据进行打印，则不应该进行标记更新，执行此存储过程不会查询到数据
                strSQL = "ZL_影像报告打印_Update(" & curData1 & ")"
                zlDatabase.ExecuteProcedure strSQL, "更新打印标记"
    
                blnSyncRow = True
                
                If mSysPar.blnPrintCommit = True Then   '打印后直接完成
                    If Menu_Manage_检查最终完成(Val(curData1), False) Then blnSyncRow = False '过程内部会自动对检查列表行进行更新
                End If
                
                
            Else
                '如果不是病历编辑器，需要判断是否打印前
                If Val(strTag) <> 1 Then    '打印前
                    If mSysPar.blnIgnoreResult = False And mSysPar.lngHintType = 2 Then Call ReportResultHint(curData1): blnSyncRow = True
                Else
                    '如果是打印后，均需要验证是否结果提示
                    Call ReportResultHint(curData1)
                    
                    strSQL = "ZL_影像报告打印_Update(" & curData1 & ")"
                    zlDatabase.ExecuteProcedure strSQL, "更新打印标记"
                    
                    blnSyncRow = True
                    
                    If mSysPar.blnPrintCommit = True Then   '打印后直接完成
                        If Menu_Manage_检查最终完成(Val(curData1), False) Then blnSyncRow = False
                    End If
                    
                End If
            End If
            
'        Case BM_RIS_EVENT_COMPLETE  '检查完成事件*****************************************************************************
'            If Val(strTag) = 1 Then
'                '发送检查完成消息
'                Call mobjMsgCenter.Send_Msg_StudyComplete(objStudyInfo.lngAdviceId, strReportId)
'
'                blnSyncRow = True
'            End If
            
'        Case BM_RIS_EVENT_CANCELCOMP    '取消检查完成事件*****************************************************************************
'            If Val(strTag) = 1 Then
'                '发送检查撤销完成消息
'                Call mobjMsgCenter.Send_Msg_CancelComplete(mobjCurStudyInfo.lngAdviceId)
'
'                blnSyncRow = True
'            End If
            
            
        Case BM_REPORT_EVENT_AUDIT   '审核签名通知*****************************************************************************
            'curData1 表示医嘱ID
            'curData2 表示报告ID 暂未使用
            'curData3 表示编辑器类型
            
            If Val(curData3) = -1 Or Val(strTag) = 1 Then
                If mSysPar.blnIgnoreResult = False And mSysPar.lngHintType = 1 Then Call ReportResultHint(curData1)
                 
                '保存审核人
                lngSendNo = GetSendNo(curData1, strStudyUID)
                Call ResetState(curData1, lngSendNo, strStudyUID)
                
                blnOk = False
                If mSysPar.blnAuditAutoPrint Then '终审后直接打印
                    '审核后打印，如果当前报告能够进行审核，说明没有进行转储
                    If ReoprtPrint(Val(curData1), False, True, , strExPro) Then
                        '更新打印状态
                        strSQL = "ZL_影像报告打印_Update(" & curData1 & ")"
                        zlDatabase.ExecuteProcedure strSQL, "更新打印标记"
                        
                        blnOk = True
                    End If
                End If
                
                blnSyncRow = True
                If mSysPar.blnCompleteCommit Then   '如果“审核后直接完成”
                    If Menu_Manage_检查最终完成(Val(curData1), False) Then blnSyncRow = False
                Else
                    If mSysPar.blnPrintCommit = True And blnOk Then   '打印后直接完成
                        If Menu_Manage_检查最终完成(Val(curData1), False) Then blnSyncRow = False
                    End If
                End If
                
                '发送状态同步消息
                Call mobjMsgCenter.Send_Msg_StateSync(curData1)
                
                
            End If
            
        Case BM_REPORT_EVENT_SIGN   '诊断签名通知*****************************************************************************
            'curData1 表示医嘱ID
            'curData2 表示报告ID 暂未使用
            'curData3 表示编辑器类型
            
            If Val(curData3) = -1 Or Val(strTag) = 1 Then
                If mSysPar.blnIgnoreResult = False And mSysPar.lngHintType = 0 Then Call ReportResultHint(curData1)
                
                '保存报告人
                lngSendNo = GetSendNo(curData1, strStudyUID)
                Call ResetState(curData1, lngSendNo, strStudyUID)
                
                '发送状态同步消息
                Call mobjMsgCenter.Send_Msg_StateSync(curData1)
                
                blnSyncRow = True
            End If
            
        Case BM_REPORT_EVENT_SAVE, BM_REPORT_EVENT_POPUPEXIT   '报告保存通知*****************************************************************************
            'curData1 表示医嘱ID
            'curData2 表示报告ID 暂未使用
            'curData3 表示编辑器类型
            
            If Val(curData3) = -1 Or Val(strTag) = 1 Then
                
                '保存报告人
                lngSendNo = GetSendNo(curData1, strStudyUID)
                Call ResetState(curData1, lngSendNo, strStudyUID)
                
                '更新影像检查图像的报告图标记，根据报告中保存的报告图更新影像检查图像记录
                strSQL = "Zl_影像检查图象_报告图(" & curData1 & ")"
                Call zlDatabase.ExecuteProcedure(strSQL, "更新影像报告图标记")
    
                If (mlngModule = G_LNG_VIDEOSTATION_MODULE And mSysPar.lngVideoStationMoneyExeModle = 2) Or _
                   (mlngModule = G_LNG_PATHSTATION_MODULE And mSysPar.lngPatholStationMoneyExeModle = 2) Or _
                   (mlngModule = G_LNG_PACSSTATION_MODULE And mSysPar.lngPacsStationMoneyExeModle = 1) Then
                    '执行费用
                    Call ExecuteExpense(curData1, GetSendNo(curData1), 4)
                End If
                
                If ucPacsHelper1.AdviceId = CLng(curData1) Then
                     Call ucPacsHelper1.SyncReportImgState(GetReportImgs(curData1))
                End If
                
                '发送状态同步消息
                Call mobjMsgCenter.Send_Msg_StateSync(curData1)
                
                If lngEventNo = BM_REPORT_EVENT_POPUPEXIT Then
                    If CLng(curData1) = mobjCurStudyInfo.lngAdviceId And mstrSelTabName = C_TAB_NAME_检查报告 Then
                        '刷新嵌入式报告内容
                        If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.zlRefreshFace(mobjCurStudyInfo, mstrSelModuleTag, True)
                    End If
                End If

                blnSyncRow = True
            End If
            
        Case BM_REPORT_EVENT_DELETE '报告删除通知*****************************************************************************
            'curData1 表示医嘱ID
            'curData2 表示报告ID 暂未使用
            'curData3 表示编辑器类型
            
            If Val(curData3) = -1 Or Val(strTag) = 1 Then
            
                strSQL = "ZL_影像报告标记_Clear(" & curData1 & ")"
                zlDatabase.ExecuteProcedure strSQL, "清空报告标记"
                
                strSQL = "Zl_影像检查图象_报告图(" & curData1 & ")"
                zlDatabase.ExecuteProcedure strSQL, "清空报告图"
                
                '报告缩略图状态同步显示
                Call ucPacsHelper1.ClearReportImgState
                
                blnSyncRow = True
            End If
            
        Case BM_REPORT_EVENT_REJECT '报告驳回通知*****************************************************************************
            'curData1 表示医嘱ID
            'curData2 表示报告ID 暂未使用
            'curData3 表示编辑器类型
            
            If Val(curData3) = -1 Or Val(strTag) = 1 Then
                blnSyncRow = Not Menu_Manage_SendAudit(curData1, "")   '该方法中会刷新列表行
                
                '发送状态同步消息
                Call mobjMsgCenter.Send_Msg_StateSync(curData1)
            End If
            
        Case BM_REPORT_EVENT_BACK '报告回退通知*****************************************************************************
            'curData1 表示医嘱ID
            'curData2 表示报告ID 暂未使用
            'curData3 表示编辑器类型
            
            If Val(curData3) = -1 Or Val(strTag) = 1 Then
                lngSendNo = GetSendNo(curData1, strStudyUID)
                
                Call ResetState(curData1, lngSendNo, strStudyUID)
                blnSyncRow = Menu_Manage_SendAudit(curData1, "")    '该方法中会刷新列表行
            End If
            
        Case BM_REPORT_EVENT_OPEN  '报告打开事件
            'curData1 表示医嘱ID
            'curData2 表示报告ID 暂未使用
            'curData3 表示编辑器类型
            
            If Val(curData3) = -1 Or Val(strTag) = 1 Then
                '医技模块打开报告后自动观片
                If mSysPar.blnShowImgAfterReport = True And mlngModule = G_LNG_PACSSTATION_MODULE Then
                    If mobjPacsCore Is Nothing Then
                        HintMsg "图像查看对象无效，不能进行此操作。", "cbrMain_Execute", vbOKOnly
                        Exit Sub
                    End If
                
                    Call OpenViewer(1, mobjPacsCore, curData1, False, Me)
                End If
            End If
            
        Case BM_REPORT_EVENT_QUALITY    '报告质量标记
            If Val(strTag) = 1 Then
                blnSyncRow = True
            End If
            
        Case BM_IMAGE_EVENT_QUALITYTAG  '影像质量标记
            If Val(strTag) = 1 Then
                blnSyncRow = True
            End If
            
        Case BM_IMAGE_EVENT_GETIMAGE '获取影像
            If Val(strTag) = 1 Then
                blnSyncRow = True
            End If
            
        Case BM_IMAGE_EVENT_TECHDO  '技师执行
            If Val(strTag) = 1 Then
                blnSyncRow = True
            End If
            
        Case BM_IMAGE_EVENT_CHANGEDEVICE   '改变影像设备
            If Val(strTag) = 1 Then
                blnSyncRow = True
            End If
            
        Case BM_IMAGE_EVENT_XWFILMPRINT       '胶片按需打印
            If Val(strTag) = 1 Then
                blnSyncRow = True
            End If
    
        Case BM_IMAGE_EVENT_DEL         '删除图像后通知
            If Val(strTag) = 1 And Val(curData3) = -1 Then      '判断是否为删除最后一张图
                
''                影像医技frmWork_ImageV2删除图像时，需要确保只有删除最后一张图像时才能触发该消息
'                If mobjCurStudyInfo.lngAdviceId = Val(curData1) Then
'                    If mobjCurStudyInfo.intStep = 3 Then
'                        strSQL = "Zl_影像检查_State(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & ",2,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & GetCurDeptId & ")"
'                        zlDatabase.ExecuteProcedure strSQL, "删除最后一张图像"
'
'                        mobjCurStudyInfo.intStep = 2
'                    End If
'                Else
                    lngSendNo = GetSendNo(curData1, strStudyUID, , lngStep)
            
                    '如果检查状态为已检查，则再删除所有图像后，需要对图像进行回退
                    If lngStep = 3 And Len(strStudyUID) <= 0 Then
                        '设置影像检查状态，如果删除最后一个图，且原检查过程为3，则修改为2
                        strSQL = "Zl_影像检查_State(" & Val(curData1) & "," & lngSendNo & ",2,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & GetCurDeptId & ")"
                        zlDatabase.ExecuteProcedure strSQL, "删除最后一张图像"
                    End If
'                End If
                
                '发送状态同步消息
                Call mobjMsgCenter.Send_Msg_StateSync(curData1)
                    
                blnSyncRow = True
            End If
            
            If Val(strTag) = 1 Then Call SyncHelperDataState(curData1, strCurExPro, 0)
            
        Case BM_IMAGE_EVENT_FIRST '首次图像采集通知
            If Val(curData2) = -1 Then      '表示首次图像采集
                Call WriteEprExChangeData(curData1)
                
'                If mobjCurStudyInfo.lngAdviceId = Val(curData1) Then
'                    If mobjCurStudyInfo.intStep < 3 Then
'                        strSQL = "Zl_影像检查_State(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & ",3,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & GetCurDeptId & ")"
'                        zlDatabase.ExecuteProcedure strSQL, "检查首次采集"
'
'                        mobjCurStudyInfo.intStep = 3
'                    End If
'                Else
                    '从数据库获取指定检查状态
                    lngSendNo = GetSendNo(curData1, strStudyUID, , lngStep)
                    
                    If lngStep < 3 Then
                        strSQL = "Zl_影像检查_State(" & Val(curData1) & "," & lngSendNo & ",3,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & GetCurDeptId & ")"
                        zlDatabase.ExecuteProcedure strSQL, "检查首次采集"
                    End If
'                End If
                
                '发送状态同步消息
                Call mobjMsgCenter.Send_Msg_StateSync(curData1)
                
                blnSyncRow = True
            ElseIf Val(curData2) = -2 Or Val(curData2) = -3 Then
                lngSendNo = GetSendNo(curData1, strStudyUID, , lngStep)
                
                If lngStep < 3 Then
                    strSQL = "Zl_影像检查_State(" & Val(curData1) & "," & lngSendNo & ",3,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & GetCurDeptId & ")"
                    zlDatabase.ExecuteProcedure strSQL, "检查首次采集"
                    
                    '发送状态同步消息
                    Call mobjMsgCenter.Send_Msg_StateSync(curData1)
                    
                    blnSyncRow = True
                End If
            End If
            
            '如果是pacs报告编辑器
            Call SyncHelperDataState(curData1, Val(strCurExPro), 0)
            
        Case WM_XWREPORT_IMG    '加入报告图
            Call SyncHelperDataState(curData1, 0, 0)
            
            '同步将观片中另存的报告图添加到报告编辑界面
            If mSysPar.blnDirectSendRepImg Then Call AddViewImageToReport(curData1, curData2)
            
        Case BM_REPORT_EVENT_ADDIMG
            Call AddViewImageToReport(curData1, curData2)
            
        Case BM_REPORT_EVENT_CLOSEEPR   '独立窗口关闭后，需要更新嵌入窗口内容
        
            If mstrSelTabName = C_TAB_NAME_检查报告 And Not mobjWork_Report Is Nothing Then
                If Val(curData1) = mobjWork_Report.StudyInfo.lngAdviceId Then
                    Call mobjWork_Report.zlRefreshFace(mobjWork_Report.StudyInfo, GetWorkModuleTag(C_TAB_NAME_检查报告), True)
                End If
            End If
            
        Case BM_REPORT_EVENT_REFWCHR    '刷新词句   其他窗口的词句字符变更后，会发送此消息进行同步更新
            Call ReinitWordChar(curData1)
            
        Case BM_REPORT_EVENT_REFFRAGMENT    '刷新词句片段  其他窗口的词句片段变更后，会发送此消息进行同步更新
            Call ReinitWordFragment(curData1)
            
        Case BM_PATHOL_EVENT_BASE + wetPatholQuality    '病理质量
            blnSyncRow = True
            
        Case BM_PATHOL_EVENT_BASE + wetPatholBatSlices  '批量制片
            blnSyncRow = True
            
        Case BM_PATHOL_EVENT_BASE + wetPatholBatSpeExm  '批量特检
            blnSyncRow = True
            
        Case BM_PATHOL_EVENT_BASE + wetMaterialSave '材块保存
            blnSyncRow = True
            
        Case BM_PATHOL_EVENT_BASE + wetSlicesSure   '制片确认
            blnSyncRow = True
            
        Case BM_PATHOL_EVENT_BASE + wetSpeExamSure  '特检确认
            blnSyncRow = True
            
        Case BM_PATHOL_EVENT_BASE + wetPatholRequest, BM_PATHOL_EVENT_BASE + wetSpecimenAccept, BM_PATHOL_EVENT_BASE + wetMaterialSure, BM_PATHOL_EVENT_BASE + wetMaterialSave
            blnSyncRow = True
            
            If Not mobjWork_Pathol Is Nothing Then
                If mobjWork_Pathol.AdviceId = Val(curData1) Then
                    '刷新病理其他模块的数据
                    Call ForceRefreshPatholModule
                End If
            End If
        
    End Select
    
    If Val(strTag) = 1 Or Val(curData3) = -1 Then
        '功能操作完成后执行插件
        Call ExecutePlugin(lngEventNo, 1, curData1, curData2, curData3)
    End If
    
    If blnSyncRow Then Call UpdateQueryListData(Nothing, curData1)
'Exit Sub
'errHandle:
'    err.Raise 513, , err.Description
End Sub


Private Function VideoIsAttachReportWindow(Optional ByVal lngVideoRootHwnd As Long = 0)
'判断视频是否嵌入的弹出式报告窗口
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
'强制刷新病理模块
    Dim i As Long
    
    For i = 0 To UBound(mAryWorkModule)
        If InStr(mAryWorkModule(i).ModuleName, "病理") > 0 Then
            If Not mAryWorkModule(i).objModule Is Nothing Then
                Call mAryWorkModule(i).objModule.zlRefreshFace(True)
            End If
        End If
    Next

End Sub

Private Sub ReinitWordChar(ByVal lngSourceHwnd As Long)
'重置常用词句字符
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
'重置常用词句片段
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
'添加观片图像到报告
    Dim objForm As Object
    Dim strFileName As String
    Dim objFSO As New FileSystemObject
    Dim objPopup As Object
    Dim objEmbed As Object
    
    strFileName = objFSO.GetFileName(strImgFile)
    
    For Each objForm In Forms
        If TypeOf objForm Is frmReportV2 Then
            If objForm.AdviceId = lngAdviceId Then
                If objForm.Caption = "报告编辑" Then '嵌入式报告编辑不会调用setTitle方法显示 姓名 性别等患者基本信息
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
'同步helper模块数据显示
    Dim objForm As Object
    
    If lngSourceHwnd < 0 Then Exit Sub
    
    For Each objForm In Forms
        If TypeOf objForm Is frmReportV2 Then
            Call objForm.SyncHelper(lngAdviceId, lngSourceHwnd, lngSyncType)
        End If
    Next
    
    If ucPacsHelper1.hwnd <> lngSourceHwnd And lngAdviceId = mobjCurStudyInfo.lngAdviceId Then
        If lngSyncType = 0 And ucPacsHelper1.SelTabName <> "图像" Then Exit Sub
        If lngSyncType = 1 And ucPacsHelper1.SelTabName <> "词句" Then Exit Sub
        If lngSyncType = 2 And ucPacsHelper1.SelTabName <> "历史" Then Exit Sub
        If lngSyncType = 3 And ucPacsHelper1.SelTabName <> "缓存" Then Exit Sub
        
        Call ucPacsHelper1.zlRefresh(mobjCurStudyInfo, 0, True)
    End If
End Sub

Private Sub WriteEprExChangeData(ByVal lngAdviceId As Long)
'写入和病历编辑器进行图像交换的标记
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
'***********************************************IEventNotify实现***********************************************

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
'获取报告图像UID
'报告保存后调用
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
On Error GoTo errhandle
    GetReportImgs = ""
    
    strSQL = "select 图像UID from 影像检查图象 a, 影像检查序列 b, 影像检查记录 c where a.序列UID =b.序列UID and b.检查UID = c.检查UID and c.医嘱ID=[1] and nvl(a.报告图,0)<>0"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询检查报告图", lngAdviceId)
    If rsData.RecordCount <= 0 Then Exit Function
    
    While Not rsData.EOF
        GetReportImgs = GetReportImgs & ";" & NVL(rsData!图像UID) & ";"
        Call rsData.MoveNext
    Wend
    
Exit Function
errhandle:
    If HintError(err, "GetReportImgs") = 1 Then Resume
End Function

Private Sub ResetState(ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal strStudyUID As String)
'该过程只针对没有转储的数据进行处理
'2-已报到，3-已检查 4-已报告，5-已审核，6-已完成
    Dim curSignInfo As TReportSignInfo
    Dim strSQL As String
    Dim lngState As Long
    Dim strCreateUser As String
    
    curSignInfo = GetLastSignInfo(lngAdviceId, strCreateUser)
    
    If curSignInfo.ID = 0 Then
        '报告处于未签名状态
        '有可能是回退到未签名状态
        '需要判断是否有图像，如果没有图像，则是已报到状态，如果有图像，则是已检查状态
        lngState = 2
        If Len(strStudyUID) > 0 Then lngState = 3
        
        '当没有签名时，写入实际的报告人
        strSQL = "ZL_影像报告保存_Update(" & lngAdviceId & ",'" & IIf(mstrRPTExecutor <> "", mstrRPTExecutor, strCreateUser) & "','')"
        Call zlDatabase.ExecuteProcedure(strSQL, "更新报告人员")
    Else
        If curSignInfo.签名级别 > 1 Then
            '报告处于审核签名状态
            lngState = 5
            
            strSQL = "ZL_影像报告保存_Update(" & lngAdviceId & ",'" & strCreateUser & "','" & curSignInfo.姓名 & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "更新报告人员")
        Else
            '报告处于诊断签名状态
            lngState = 4
            
            '清空复核人
            strSQL = "ZL_影像报告保存_Update(" & lngAdviceId & ",'" & strCreateUser & "','')"
            Call zlDatabase.ExecuteProcedure(strSQL, "更新报告人员")
        End If
    End If
    
    strSQL = "Zl_影像检查_State(" & lngAdviceId & "," & lngSendNo & "," & lngState & ",NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "更新报告状态")
End Sub

Private Sub ExecuteExpense(ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal lngProcState As Long)
'执行费用
    Dim lngID As Long
    Dim strSQL As String
    
    If mblnAllDepts Then
        lngID = UserInfo.部门ID
    Else
        lngID = mlngCur科室ID
    End If
    
    strSQL = "Zl_影像费用执行(" & lngAdviceId & "," & lngSendNo & "," & lngProcState & ",NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & lngID & ")"
    zlDatabase.ExecuteProcedure strSQL, "执行费用"
End Sub

Private Sub ExecutePlugin(ByVal lngEventNo As Long, ByVal lngTimeTag As Long, _
    Optional data1, Optional data2, Optional data3)
'执行插件
'lngTimeTag:0-功能执行前，1-功能执行后
    Dim lngTimeType As Long
    
'On Error GoTo errHandle

    Select Case lngEventNo
        Case BM_RIS_EVENT_REGISTER
            lngTimeType = EInterfaceExeTimeV2.检查登记
        
        Case BM_RIS_EVENT_RECEVIE
            lngTimeType = EInterfaceExeTimeV2.检查报到
            
        Case BM_RIS_EVENT_COMPLETE
            lngTimeType = EInterfaceExeTimeV2.检查完成
            
        Case BM_RIS_EVENT_CANCELREG
            lngTimeType = EInterfaceExeTimeV2.取消登记
            
        Case BM_RIS_EVENT_CANCELREC
            lngTimeType = EInterfaceExeTimeV2.取消报到
        
        Case BM_RIS_EVENT_CANCELCOMP
            lngTimeType = EInterfaceExeTimeV2.取消完成
            
        Case BM_IMAGE_EVENT_CAPTURE
            lngTimeTag = EInterfaceExeTimeV2.图像采集
            
        Case BM_IMAGE_EVENT_DEL
            lngTimeType = EInterfaceExeTimeV2.删除图像
            
        Case BM_REPORT_EVENT_AUDIT
            lngTimeType = EInterfaceExeTimeV2.报告审核
            
        Case BM_REPORT_EVENT_SIGN
            lngTimeType = EInterfaceExeTimeV2.报告签名
            
        Case BM_REPORT_EVENT_SAVE
            lngTimeType = EInterfaceExeTimeV2.报告保存
            
        Case BM_REPORT_EVENT_REJECT
            lngTimeType = EInterfaceExeTimeV2.报告驳回
            
        Case BM_REPORT_EVENT_DELETE
            lngTimeType = EInterfaceExeTimeV2.删除报告
            
        Case BM_REPORT_EVENT_BACK
            lngTimeType = EInterfaceExeTimeV2.报告回退
            
        Case BM_SYS__EVENT_MENU
            lngTimeType = EInterfaceExeTimeV2.菜单执行
            
    End Select
    
    Call ExecutePluginInterface(lngTimeType, lngTimeTag, data1, data2, data3)
'Exit Sub
'errHandle:
'    err.Raise -1, , err.Description
End Sub

Private Function TraversalAdvice(ByVal lngAdviceId As Long, ByVal blnIsMoveDown As Boolean, _
    Optional ByRef lngSendNo As Long = 0, Optional ByRef blnIsMoved As Boolean = False) As Long
'遍历医嘱
    Dim lngRowIndex As Long
    Dim lngNewRow As Long
    Dim lngResult As Long
    Dim lngIdCol As Long
    Dim objBaseInfo As clsStudyInfo
    
    TraversalAdvice = lngAdviceId
    
    lngIdCol = vsfList.ColIndex("医嘱ID")
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
    '创建卡结算部件
    
    strDllName = "zlOneCardComLib.clsOneCardComLib"
    Set mobjSquareCard = CreateObject(strDllName)
    
    'mobjAppendBill: 如果mobjAppendBill不为空，表示使用混合模式
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
    '初始化观片部件
    If mobjPacsCore Is Nothing Then
        Set mobjPacsCore = New zl9PacsCore.clsViewer
        
        '非影像医技工作站观片时，不显示另存报告图按钮
        If mlngModule <> G_LNG_PACSSTATION_MODULE Then
            mobjPacsCore.ReportImgButtonVisible = False
        End If
    End If
    
    '初始化卡结算部件
    If Not mobjSquareCard Is Nothing Then
        mobjSquareCard.zlInitComponents Me, mlngModule, glngSys, gstrDBUser, gcnOracle
    End If
    
    '这句话不能省略，最后一个参数内容随意，只要格式正确即可，后续会被修改
    PatiIdentify.zlInit Me, glngSys, mlngModule, gcnOracle, gstrDBUser, mobjSquareCard, InitCardType("姓名;")
Exit Sub
errhandle:
    HintError err, "InitBaseComponent", False
End Sub

Private Sub StartMsgCenter(ByVal lngDeptId As Long)
'启动消息中心
    If mobjMsgCenter Is Nothing Then
        Set mobjMsgCenter = New clsPacsMsgProcess
    End If
    
    Call mobjMsgCenter.OpenMsgCenter(mlngModule, lngDeptId, mstrPrivs)
End Sub


Private Sub RestoreFormState()
    Dim blnDo As Boolean
    Dim strLayout As String
On Error GoTo errhandle
    '得到个性化风格参数
    blnDo = Val(zlDatabase.GetPara("使用个性化风格")) <> 0
    
     '如果注册表中工具栏相关值为空 并且 已勾选个性化设置，那么向注册表写入工具栏显示模式值
    If mintToolBarWriteReg = 9 Or (mintToolBarWriteReg = 0 And blnDo) Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\CommandBars", "cbrMainButtonText", 3
    End If
    
    '恢复窗体的状态   注：恢复窗体状态 必须放在 向注册表写入工具栏显示模式值 的语句后面，否则会造成工具栏显示模式有误。
    Call RestoreWinState(Me, App.ProductName)
    
    strLayout = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\", "HELPER", "")
    Call ucPacsHelper1.SetLayout(strLayout)
    
     '工具栏--- 文本标签 的设置使用RestoreWinState 恢复不了，还需要单独处理，如未勾选个性化设置，则工具栏默认显示图标和文本
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
    
    '默认报告人为当前登录用户
    mstrRPTExecutor = UserInfo.姓名
    
    Set mrsDeptParas = Nothing  '使科室参数可以重新进行加载
    
    mstrPrivs = gstrPrivs & ";" & GetPrivFunc(100, 9001) & ";"  '用于判断传染病，危急值等菜单权限
    
    Call WriteLog("ShowStation -> Step 0：设置窗体样式及初始化基础部件。")
    
    Call StyleChange(sAshen)
    
    If IsExistsBGServer() = "" Then
        '未检测到后台传输服务程序，则进行提示
        Call HintMsg("未检测到后台服务程序，图像将不能启用后台传输。", "ShowStation", infWaring)
    Else
        'TODO:打开服务处理失败窗口...
    End If
    
    Call DynamicCreateModuleObj
    
    Call InitBaseComponent

    Call WriteLog("ShowStation -> Step 1：进入影像主窗口初始化流程。")

    If Not mblnFormLoadState Then
        If Not InitDepts Then '初始化医技科室
            Unload Me
            Exit Sub
        End If
        
        If mlngModule <> 1290 Then
            ucPacsHelper1.AllowEmbedVideo = IIf(Val(GetDeptPara(mlngCur科室ID, "显示视频采集", "0")) <> 0, True, False)
            
            Set mobjCapLinker = New clsCapLinker
            Set mobjCapLinker.MainHelper = ucPacsHelper1
            
            mobjCapLinker.Init Me, mlngCur科室ID, mstrPrivs
        Else
            ucPacsHelper1.AllowEmbedVideo = False
            Call ucPacsHelper1.HideEmbedVideo
        End If
        
        ReDim mAryWorkModule(0)
    
        Call StartMsgCenter(mlngCur科室ID)  '启动消息中心
        
        Call InitPars                       '初始化参数
        
        Call InitQueryWrapComponent
        
        Call initInterface(mlngModule)
        
        Call ReSetFormFontSize              '设置界面字体大小
        
        Call InitLayout                     '设置界面布局
        
        Call InitPacsHelper                 '初始化pacshelper对象
        Call InitWorkModuleTab              '设置工作模块tab标签
        Call InitCommandBars                '设置系统菜单
        
        Call initTabExtra                   '初始化列表附加信息
         
'        Call RestoreFormState               '恢复窗口状态
        
        mblnFormLoadState = True
    End If
    
    Call WriteLog("ShowStation -> Step 2：显示系统界面。")
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then Set gobjEvent = Me
    
    If Not mobjPacsCore Is Nothing Then mobjPacsCore.DirectSendRepImg = mSysPar.blnDirectSendRepImg
    
    '先显示出当前系统窗体
    Me.Show , Owner
    
    Call RestoreFormState               '恢复窗口状态
    
    If Me.WindowState = 1 Then Me.WindowState = 0
    
    Call WriteLog("ShowStation -> Step 3：刷新数据列表。")
    '刷新检查数据
    
    If mintQueryState = 1 Then
        Call ExecuteDefaultQueryScheme
    End If

    mblnInitOk = True

    Call WriteLog("ShowStation -> Step 4：触发工作模块显示。")
    If Not TabWindow.Selected Is Nothing Then
        Call TabWindow_SelectedChanged(TabWindow.Selected)
    End If
 
    '未避免系统启动后不能看见视频画面，需要重启一次视频预览
    If Not mobjWork_ImageCap Is Nothing Then
        Call WriteLog("ShowStation -> Step 5：重启视频预览。")
        Call mobjWork_ImageCap.zlRePreview
    End If
    
    Call WriteLog("ShowStation -> Step End.：结束影像主窗口初始化流程。")
    
'    Debug.Print "ShowStation耗时" & GetTickCount - t1

    '如果定义了热键，则需要开启热键对象
    If mobjCaptureHot Is Nothing And (mstrCaptureHot <> "" Or mstrCaptureAfterHot <> "" Or mstrCaptureAfterTagHot <> "") Then
        Set mobjCaptureHot = New zl9PacsControl.clsHookKey
        Call mobjCaptureHot.EnableHook(WM_KEYDOWN, True)
    End If
    
    
    '这里需要设置一次字体 用于修改过滤菜单字体 后续可调整为只对菜单控件修改。
    '初始化界面字体 加到这里为了防止在一些特殊操作的时候，会导致字体恢复成初始化
    Call SetFontSize(IIf(gbytFontSize = 12, 1, IIf(gbytFontSize = 15, 2, 0)))
    
    Call StartImageValid
End Sub

Private Sub StartImageValid()
On Error GoTo errhandle
    If mSysPar.lngImageValid > 0 Then
        If Len(Dir(GetAppRootPath & "zlPacsImageValid.exe")) > 0 Then
            If InitRegister Then
                Shell GetAppRootPath & "zlPacsImageValid.exe   " & gstrServerName & "||" & gstrUserName & "||" & gstrUserPswd & "||" & mlngCur科室ID & "||" & mSysPar.lngImageValid & "||" & "" & "||2", 1
            End If
        End If
    End If
Exit Sub
errhandle:
    HintError err, "StartImageValid<启动图像验证>", False
End Sub



Private Sub Menu_File_Excel_click()
'功能:将数据复制到可打印的对象，调用打印
'参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
'       lngSelectedRow,记录调用打印部件前的选中行，在清单关闭后恢复
On Error GoTo errhandle
    Dim bytMode As Byte
    Dim lngSelectedRow As Long
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    mblnInitOk = False
    
    Set objPrint.Body = vsfList
    objPrint.title.Text = "检查病人清单"
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & zlDatabase.Currentdate())
    Call objPrint.BelowAppRows.Add(objAppRow)

    '给 是否是打印清单参数赋值
    mblnIsPrintMode = True
    '得到打印清单前的当前选中行
    lngSelectedRow = vsfList.RowSel
    
    bytMode = zlPrintAsk(objPrint)
    If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    
    '打印货预览结束后 恢复选中行
    vsfList.Row = lngSelectedRow
    mblnIsPrintMode = False
    
    mblnInitOk = True
    
    Exit Sub
errhandle:
    mblnInitOk = True
    If HintError(err, "Menu_File_Excel_click") = 1 Then Resume
End Sub

Private Sub Menu_RichEPR(ByVal cbrID As Long)
'自动打开报告编辑器，同时处理了PACS报告编辑器和电子病历编辑器
On Error GoTo errhandle
    Dim cbrControl As CommandBarControl, i As Long
    Dim strCurModuleTag As String
    
    '如果没有选择行数据，则直接退出执行
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_RichEPR", vbInformation
        Exit Sub
    End If
    
    '报告页面不可见时不执行任何操作
    If TabWindow.Selected.Caption <> C_TAB_NAME_检查报告 Then
        For i = 0 To TabWindow.ItemCount - 1 '循环到了才触发
            If TabWindow(i).Caption = C_TAB_NAME_检查报告 And TabWindow(i).Visible = True Then
                TabWindow(i).Selected = True
                Exit For
            End If
        Next
        
        If TabWindow.Selected.Caption <> C_TAB_NAME_检查报告 Then Exit Sub
    Else
        If TabWindow.Selected.Visible = False Then Exit Sub
    End If

    strCurModuleTag = GetWorkModuleName(mstrSelTabName, mobjCurStudyInfo.lngExeDepartmentId, mobjCurStudyInfo.lngPatientFrom)
    If strCurModuleTag <> mstrSelModuleTag Then
       Call SelectModule(mstrSelTabName, strCurModuleTag)
       TabWindow.Selected.tag = strCurModuleTag
    End If
    
    '找到报告页面，再打开这个报告页面
    '刷新嵌入页面内容
    If Not mobjWork_Report Is Nothing Then
        Call mobjWork_Report.zlRefreshFace(mobjCurStudyInfo, strCurModuleTag, True)
        
        If strCurModuleTag = C_WORKMODULE_NAME_老版报告 Then
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
        .mlng科室ID = mlngCur科室ID
        .mstrPrivs = mstrPrivs
        .Show 1, Me
        
        If .mblnOk Then
            InitLocalPars
            
            If Not mobjWork_Report Is Nothing Then
                '重新加载和报告相关的设置参数
                Call mobjWork_Report.InitReportParameter
            End If
            mSysPar.blnAutoPrint = Val(zlDatabase.GetPara("报到后自动打印申请单", glngSys, mlngModule, 0)) '报到后自动打印申请单
            mSysPar.blnAutoPrintCheck = Val(zlDatabase.GetPara("自动规避重复申请打印", glngSys, mlngModule, 0))
            mSysPar.blnShowImgAfterReport = (Val(zlDatabase.GetPara("报告时观片", glngSys, mlngModule, 0)) = 1)
            mSysPar.blnAutoOpenReport = (Val(zlDatabase.GetPara("开始检查自动打开报告", glngSys, mlngModule, 0)) = 1)
            
            '判断是否需要开启热键
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


'显示快捷方式配置
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
'功能：调用帮助主题
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

Private Sub Menu_Manage_取消关联()
'取消关联的最后结果是，每次取消关联后，图象全部按照序列被拆散成N条临时记录
On Error GoTo errhandle
    Dim lngResult As Long
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_取消关联", vbInformation
        Exit Sub
    End If

    lngResult = -1
    
    '如果是模块号为1298的RIS工作站，则调用新网的数据库查询已匹配的图像记录
    If mlngModule = G_LNG_PACSSTATION_MODULE And mobjCurStudyInfo.intImageLocation = 1 Then
        lngResult = XWShowMatched(Me, mobjCurStudyInfo.lngAdviceId)
    Else
        frmSelectMuli.ShowImageReleation mlngModule, mobjCurStudyInfo.lngAdviceId, mstrPrivs, mobjCurStudyInfo.blnMoved, IIf(mlngModule = G_LNG_PACSSTATION_MODULE, False, True), mlngCur科室ID, 1
        
        If frmSelectMuli.mblnOk = True Then lngResult = 0
    End If
    
    If lngResult <> 0 Then Exit Sub
    
    Call ReleationImage(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, 1, True)

Exit Sub
errhandle:
    Call HintError(err, "Menu_Manage_取消关联", False)
End Sub

Private Sub Menu_Manage_完成病理补费()
'混合模式下使用
    Dim objPatholPrice As New frmPatholPrice
    
    objPatholPrice.zlInitModule -1, mstrPrivs, mlngCur科室ID, Me
    objPatholPrice.zlRefresh mlngCur科室ID, mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.blnMoved
    
    objPatholPrice.Show 1, Me
End Sub

Private Sub Menu_Manage_补附费()
'混合模式下的补附费处理
On Error GoTo errH
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngSystemFrom As Long
    Dim strPar As String
    
    strSQL = "select B.附加标志 From 病人医嘱记录 A, 病人挂号记录 B Where A.挂号单=B.No And A.ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询附加标志", mobjCurStudyInfo.lngAdviceId)
    
    If rsData.RecordCount <= 0 Then
        '弹出老版补费窗口
        lngSystemFrom = 1
    Else
        If Val(NVL(rsData!附加标志)) = 3 Then
            '弹出新版补费
            lngSystemFrom = 2
        Else
            '弹出老版补费窗口
            lngSystemFrom = 1
        End If
    End If
    
    strPar = GetJsonPar(mobjCurStudyInfo.lngAdviceId)
    
    Call mobjAppendBill.EditChargeBill(strPar)
    Exit Sub
errH:
    If HintError(err, "Menu_Manage_补附费") = 1 Then Resume
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
    
    '查询患者来源系统
    strSysFrom = "01"
    strSQL = "Select 附加标志 From 病人挂号记录 Where 病人ID=[1] and No=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询附加标志", mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.strRegNo)
    If rsData.RecordCount > 0 Then
        If Val(NVL(rsData!附加标志)) = 3 Then strSysFrom = "02"
    End If
    
            
    strUserName = gobjRegister.GetUserName
    strUserPswd = gstrInputPwd ' GetLoginPassword 'gobjRegister.GetPassword(App.hInstance)
    
    If strSysFrom = "02" Then
        strSQL = "Select 资源ID From 人员表 Where ID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询人员表资源ID", UserInfo.ID)
        If rsData.RecordCount > 0 Then
            strUerResId = NVL(rsData!资源ID)
        End If
    
        strSQL = "Select a.病人ID," & _
                    " '' As 就诊标识, " & _
                    " Decode(a.病人来源, 4, 2, 2, 1, 0) As 病人来源, " & _
                    "Nvl(a.相关ID, a.ID) As 医嘱编号, b.发送号, " & _
                    " c.资源id As 当前科室标识, " & _
                    " c.编码 As 当前科室编码, c.名称 As 当前科室名称" & _
                    " From 病人医嘱记录 A, 病人医嘱发送 B, 部门表 C " & _
                    " Where a.Id = b.医嘱id And b.执行部门id = c.Id And a.Id = [1]"

    Else
        strNodeNo = gstrNodeNo
        strNodeName = gstrNodeName
    
        strSQL = "Select a.病人ID," & _
                    " To_Char(a.主页id) As 就诊标识, " & _
                    " Decode(a.病人来源, 4, 2, 2, 1, 0) As 病人来源, " & _
                    " b.医嘱id As 医嘱编号, b.发送号, " & _
                    " To_Char(b.执行部门id) As 当前科室标识, " & _
                    " c.编码 As 当前科室编码, c.名称 As 当前科室名称" & _
                    " From 病人医嘱记录 A, 病人医嘱发送 B, 部门表 C " & _
                    " Where a.Id = b.医嘱id And b.执行部门id = c.Id And a.Id = [1]"
                
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询医嘱Json参数", lngAdviceId)
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetJsonPar = "{" & _
            """来源系统"":""" & strSysFrom & """," & _
            """病人来源"":""" & NVL(rsData!病人来源) & """," & _
            """病人标识"":""" & NVL(rsData!病人ID) & """," & _
            IIf(strSysFrom <> "02", """就诊标识"":""" & NVL(rsData!就诊标识) & """,", "") & _
            """医嘱编号"":""" & NVL(rsData!医嘱编号) & """," & _
            """医嘱发送号"":""" & NVL(rsData!发送号) & """," & _
            """当前科室标识"":""" & NVL(rsData!当前科室标识) & """," & _
            """当前科室编码"":""" & NVL(rsData!当前科室编码) & """," & _
            """当前科室名称"":""" & NVL(rsData!当前科室名称) & """," & _
            """操作员标识"":""" & IIf(strSysFrom <> "02", lngUerResId, strUerResId) & """," & _
            """操作员编码"":""" & UserInfo.编号 & """," & _
            """操作员姓名"":""" & UserInfo.姓名 & """," & _
            """院区编码"":""" & strNodeNo & """," & _
            """院区名称"":""" & strNodeName & """," & _
            """用户名"":""" & strUserName & """," & _
            """用户密码"":""" & strUserPswd & """" & _
        "}"
    Exit Function
errH:
    If HintError(err, "GetJsonPar") = 1 Then Resume
End Function

Private Function getRegID(ByVal strRegNo As String) As Long
'功能:获取挂号id
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errhandle
    
    getRegID = 0
    
    strSQL = "select id from 病人挂号记录 where no=[1]"
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
    
    strSQL = "select 综合质量 from 病理检查信息 where 医嘱ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, GetWindowCaption, lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    If NVL(rsData!综合质量) <> "" Then IsAlreadyInputQuality = True
    Exit Function
errH:
    If HintError(err, "IsAlreadyInputQuality") = 1 Then Resume
End Function

Private Function Menu_Manage_检查最终完成(Optional lngAdviceId As Long = 0, Optional blnRefresh As Boolean = True, Optional strReportId As String = "") As Boolean
'可能由其它过程调用，此时传入有医嘱ID，但需要权限判断
On Error GoTo errhandle
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim intState As Integer
    Dim blnAllReportFinished As Boolean
    Dim objStudyInfo As New clsStudyInfo
    Dim intCol As Integer
    Dim lngRow As Long
    Dim lngAdviceIDSub As Long '本过程中的医嘱ID
    Dim lngID As Long
    
    Menu_Manage_检查最终完成 = False
    
    '如果执行过程=6 说明这个检查已经处于完成状态，此时退出本过程并且不需要提示，可能是XX后自动完成操作。
    If lngAdviceId > 0 Then
        strSQL = "select 医嘱ID from 病人医嘱发送 where 医嘱ID=[1] and 执行过程=6"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询是否已经处于完成状态", lngAdviceId)
        If rsData.RecordCount > 0 Then
            Menu_Manage_检查最终完成 = True
            Exit Function
        End If
    End If
    
    If InStr(mstrPrivs, ";检查完成;") <= 0 Then
        HintMsg "没有权限，不允许检查完成", "Menu_Manage_检查最终完成", vbInformation
        Exit Function
    End If
    
    '若未传入医嘱ID,获取选中行医嘱ID
    lngAdviceIDSub = lngAdviceId
    If lngAdviceIDSub = 0 Then
        If vsfList.Rows > 1 Then
            intCol = vsfList.ColIndex("医嘱ID")
            lngRow = vsfList.Row
            lngAdviceIDSub = Val(vsfList.TextMatrix(lngRow, intCol))
            
        End If
    End If
    
    If lngAdviceIDSub = 0 Then
        HintMsg "获取检查数据失败", "Menu_Manage_检查最终完成", vbInformation
        Exit Function
    End If
        
    Set objStudyInfo = mobjPacsQueryWrap.GetBaseInfo(lngAdviceIDSub, GetMovedState(lngRow, vsfList))
    
    If objStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_检查最终完成", vbInformation
        Exit Function
    End If
    
    If Not mSysPar.blnNoSignFinish Then
    '如果勾选允许未签名完成则不必进行下面的判断
        If Is_ExistReportWriting(lngAdviceIDSub) Then
            HintMsg "报告已经修改还未签名，不允许检查完成。", "Menu_Manage_检查最终完成", vbInformation
            Exit Function
        ElseIf objStudyInfo.intStep < 4 Then
            HintMsg "报告还未签名，不允许检查完成。", "Menu_Manage_检查最终完成", vbInformation
            Exit Function
        End If
    End If
    
    '检查完成之前，先判断是否符合条件，以下情况不能完成：
        '1、住院患者，已经出院，且有未审核的划价单，使用“执行后自动审核划价单”功能
        '2、门诊患者，有未交费的单据。
    If objStudyInfo.lngPatientFrom = 2 Then
        '住院患者，判断是否已经出院，且有未审核划价单
        If bln费用未审核出院(objStudyInfo.lngPatId, objStudyInfo.lngPageID, NVL(objStudyInfo.lngAdviceId), objStudyInfo.lngPatientFrom) Then
            '执行后自动审核划价单有效，并且病人已出院，且有未审核的划价单
            HintMsg "该病人已出院，且有未审核的划价单，不能完成！", "Menu_Manage_检查最终完成", vbExclamation
            Exit Function
        End If
    ElseIf objStudyInfo.lngPatientFrom = 4 And mSysPar.blnPEISNoCheckMoneyFinish Then
        '体检完成不判断费用 133458
    Else
        '门诊，外诊患者,判断是否有未缴费用
        If bln未缴费用(objStudyInfo.lngAdviceId) = True Then
            If objStudyInfo.intGreenChannel = 1 Or objStudyInfo.intEmergentTag = 1 Then
                If HintMsg("该患者还有未缴费的项目，是否要完成？", "Menu_Manage_检查最终完成", vbYesNo) = vbNo Then
                    Exit Function
                End If
            Else
                HintMsg "该患者还有未缴费的项目，不能完成。", "Menu_Manage_检查最终完成", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    
    Call Notify.Broadcast(BM_RIS_EVENT_COMPLETE, 0, mobjCurStudyInfo.lngAdviceId)

    '如果是病理系统，检查完成时，则需要弹出质量控制窗口
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        If Not IsAlreadyInputQuality(objStudyInfo.lngAdviceId) Then
            If Not mobjWork_Pathol Is Nothing Then
                Call mobjWork_Pathol.zlMenu.zlExecuteMenu("", conMenu_Pathol_Quality_Manage)
            End If
        End If
        
        If Not IsAlreadyInputQuality(objStudyInfo.lngAdviceId) Then
            HintMsg "未录入检查质量，不能执行完成操作。", "Menu_Manage_检查最终完成", vbInformation
            Exit Function
        End If
    End If
    
    lngID = GetCurDeptId
    
    '清空待处理人
    strSQL = "Zl_影像检查记录_变更待处理人(" & lngAdviceIDSub & ",'')"
    zlDatabase.ExecuteProcedure strSQL, "变更待处理人"
    
    
    If objStudyInfo.lngReportType = 1 Then  'pacs报告编辑器
        strSQL = "ZL_影像检查_STATE(" & objStudyInfo.lngAdviceId & "," & objStudyInfo.lngSendNo & ",6,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & lngID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ",1)"
    Else
        strSQL = "ZL_影像检查_STATE(" & objStudyInfo.lngAdviceId & "," & objStudyInfo.lngSendNo & ",6,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & lngID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ",2)"
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, "改变检查过程")

        
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        gstrSQL = "Zl_病理检查_完成(" & objStudyInfo.lngAdviceId & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "病理检查完成")
    End If
        
    '取消排队信息
    If mSysPar.blnUseQueue = True And Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Call mobjQueue.zlCompletePacsQueue(objStudyInfo.lngAdviceId)
    End If
    
    
    Menu_Manage_检查最终完成 = True
    
    Call UpdateQueryListData(Nothing, objStudyInfo.lngAdviceId)
    
    '最终完成的医嘱与当前列表选择的医嘱相同时，则同步刷新模块数据
    If lngAdviceId <> 0 And lngAdviceId = mobjCurStudyInfo.lngAdviceId Then
        Call RefreshModuleData(mstrSelTabName, mstrSelModuleTag, mobjSelModule)
    End If
    
    '发送检查完成消息
    Call mobjMsgCenter.Send_Msg_StudyComplete(objStudyInfo.lngAdviceId, strReportId)
    
    Call Notify.Broadcast(BM_RIS_EVENT_COMPLETE, 1, mobjCurStudyInfo.lngAdviceId)
    
    
Exit Function
errhandle:
    If HintError(err, "Menu_Manage_检查最终完成") = 1 Then Resume
End Function

Private Sub Menu_Manage_取消检查完成()
On Error GoTo errhandle
    Dim strSQL As String
    Dim intState As Integer

    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_取消检查完成", vbInformation
        Exit Sub
    End If

    If mobjCurStudyInfo.blnMoved Then
        HintMsg "该病人的本次检查数据已经转出到后备数据库，不允许操作。", "Menu_Manage_取消检查完成", vbInformation
        Exit Sub
    End If
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        If CheckIsArchived(mobjCurStudyInfo.lngAdviceId) Then
            HintMsg "该病人的档案已经归档，不允许操作。", "Menu_Manage_取消检查完成", vbInformation
            Exit Sub
        End If
    End If
    
    Call Notify.Broadcast(BM_RIS_EVENT_CANCELCOMP, 0, mobjCurStudyInfo.lngAdviceId)
    
    If mobjCurStudyInfo.lngReportType = 1 Then  '1-pacs报告编辑器，2-病历编辑器
        intState = getStudyState(mobjCurStudyInfo.lngAdviceId, True)
        strSQL = "ZL_影像检查_STATE(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & "," & intState & ",NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ",1)"
    Else
        intState = getStudyState(mobjCurStudyInfo.lngAdviceId, True)
        strSQL = "ZL_影像检查_STATE(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & "," & intState & ",NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & "," & zlStr.To_Date(zlDatabase.Currentdate) & ",2)"
    End If
    
    zlDatabase.ExecuteProcedure strSQL, "取消检查完成"
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        strSQL = "Zl_病理检查_取消完成(" & mobjCurStudyInfo.lngAdviceId & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "病理检查取消完成")
    End If
    
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    Call RefreshModuleData(mstrSelModuleTag, mstrSelModuleTag, mobjSelModule)
    
    '发送检查撤销完成消息
    Call mobjMsgCenter.Send_Msg_CancelComplete(mobjCurStudyInfo.lngAdviceId)
    
    Call Notify.Broadcast(BM_RIS_EVENT_CANCELCOMP, 1, mobjCurStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_取消检查完成") = 1 Then Resume
End Sub


Private Function CheckIsArchived(lngAdviceId As Long) As Boolean
 '检查该病人档案是否已经归档，已归档的检查，需要撤档才能取消完成  0--未归档  1--已归档
 On Error GoTo errhandle
 
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "select distinct c.档案状态 as 状态 from 病理检查信息 a,病理归档信息 b,病理档案信息 c where a.病理医嘱ID = b.病理医嘱ID and b.档案id = c.id and a.医嘱ID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否已归档", lngAdviceId)
    
    If rsTemp.RecordCount < 1 Then
        CheckIsArchived = False
        Exit Function
    End If
    
    CheckIsArchived = IIf(NVL(rsTemp!状态, 0) = 1, True, False)
Exit Function
errhandle:
    If HintError(err, "CheckIsArchived") = 1 Then Resume
End Function

Private Sub Menu_Manage_CriticalMark(ByVal lngID As Long)
'危急值处理
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
        Case conMenu_Manage_PacsCriticalReg     '危急患者登记
            If mobjCurStudyInfo.lngPatientFrom = 1 Then        '门诊
                Call mobjPublicAdvice.ShowAppCritical(Me, True, 0, 1, _
                            mobjCurStudyInfo.lngPatId, 0, mobjCurStudyInfo.strRegNo, mobjCurStudyInfo.lngBaby, lngCriticalId, _
                            mobjCurStudyInfo.lngAdviceId, , , , mlngCur科室ID, gstrUserName, mobjMsgCenter.Msg)
            ElseIf mobjCurStudyInfo.lngPatientFrom = 2 Then    '住院
                Call mobjPublicAdvice.ShowAppCritical(Me, True, 0, 2, _
                            mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.lngPageID, mobjCurStudyInfo.strRegNo, mobjCurStudyInfo.lngBaby, lngCriticalId, _
                            mobjCurStudyInfo.lngAdviceId, , , , mlngCur科室ID, gstrUserName, mobjMsgCenter.Msg)
            Else                                            '外来、体检
                Call mobjPublicAdvice.ShowAppCritical(Me, True, 0, 3, _
                            mobjCurStudyInfo.lngPatId, 0, "", mobjCurStudyInfo.lngBaby, lngCriticalId, _
                            mobjCurStudyInfo.lngAdviceId, , , , mlngCur科室ID, gstrUserName, mobjMsgCenter.Msg)
            End If
    
        Case conMenu_Manage_PacsCriticalManage  '危急患者管理
            If mobjPublicAdvice.ShowQueryCritical(Me, True, 2, 1, mlngCur科室ID, 0, mobjMsgCenter.Msg) = False Then Exit Sub
    End Select

    '查询医嘱危急情况...
    strSQL = "Select ID From 病人危急值记录 Where 医嘱ID=[1] and nvl(状态, 0)<>0"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询危急医嘱状态", mobjCurStudyInfo.lngAdviceId)
    If rsData.RecordCount > 0 Then
        intCritical = 1         '危急
    Else
        intCritical = 0         '不危急
    End If
    
    '更新影像危急状态
    If intCritical = 1 Then
        strSQL = "zl_影像检查_危急更新(" & mobjCurStudyInfo.lngAdviceId & ",1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)

        mobjCurStudyInfo.intDangerState = 1

        Menu_Manage_标记阴阳 conMenu_Manage_Negative
        
        '发送危急值消息
        'Call mobjMsgCenter.Send_Msg_Critical(mobjCurStudyInfo.lngAdviceId)
    ElseIf intCritical = 0 Then
        strSQL = "Zl_影像危急值记录_取消(" & mobjCurStudyInfo.lngAdviceId & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)

        mobjCurStudyInfo.intDangerState = 0
    End If
        
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)

    '如果当前模块是报告模块，则同步更新报告状态的显示
    If mobjWork_Report Is Nothing Then Exit Sub
    If TypeOf mobjSelModule Is frmReportV2 Then Call mobjSelModule.ReadRepStateTag
    
    Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_CriticalMark") = 1 Then Resume
End Sub

Private Sub Menu_Manage_标记阴阳(ByVal lngID As Long)
On Error GoTo errhandle
    Dim strSQL As String
    Dim iResult As Integer
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_标记阴阳", vbInformation
        Exit Sub
    End If
    
    Select Case lngID
        Case conMenu_Manage_Negative
            iResult = 1
        Case conMenu_Manage_Positive
            iResult = 0
    End Select
    
    strSQL = "ZL_影像检查_结果(" & mobjCurStudyInfo.lngAdviceId & "," & iResult & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "结果阴阳性")

    mobjCurStudyInfo.intPositive = iResult
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    '如果当前模块是报告模块，则同步更新报告状态的显示
    If mobjWork_Report Is Nothing Then Exit Sub
    If TypeOf mobjSelModule Is frmReportV2 Then Call mobjSelModule.ReadRepStateTag
    
    
Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_标记阴阳") = 1 Then Resume
End Sub

Private Sub Menu_Manage_绿色通道(ByVal lngID As Long)
On Error GoTo errhandle
    Dim strSQL As String
    Dim intResult As Integer
    Dim blnCanPrint As Boolean
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_绿色通道", vbInformation
        Exit Sub
    End If
    
    Select Case lngID
        Case conMenu_Manage_GChannelOk
            intResult = "1"
        Case conMenu_Manage_GChannelCancel
            intResult = "0"
    End Select
    
    strSQL = "Zl_绿色通道_Update(" & mobjCurStudyInfo.lngAdviceId & ",'" & intResult & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "绿色通道")
    
    mobjCurStudyInfo.intGreenChannel = intResult
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)

Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_绿色通道") = 1 Then Resume
End Sub

Private Sub Menu_Manage_符合情况(ByVal lngID As Long)
On Error GoTo errhandle
    Dim strResult As String
    Dim strSQL As String
    Dim lngColIndex As Long

    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_符合情况", vbInformation
        Exit Sub
    End If

    Select Case lngID
        Case conMenu_Manage_FuHe
            strResult = "符合"
        Case conMenu_Manage_JiBenFuHe
            strResult = "基本符合"
        Case conMenu_Manage_BuFuHe
            strResult = "不符合"
    End Select

    strSQL = "Zl_符合情况_Update(" & mobjCurStudyInfo.lngAdviceId & ",'" & strResult & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "符合情况")
        
    mobjCurStudyInfo.strAccord = strResult
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)

Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_符合情况") = 1 Then Resume
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
        HintMsg "没有选择病人。", "Menu_Manage_CheckList", vbInformation + vbOKOnly
    End If
Exit Sub
errhandle:
    Call HintError(err, "Menu_Manage_CheckList", False)
End Sub

'分部位执行
Private Sub menu_Manage_ExecOnePart()
    Dim frmExecForm As frmExecOnePart
    
    Set frmExecForm = New frmExecOnePart
    
    '显示分部位执行和取消窗口
    Call frmExecForm.ZlShowMe(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.strPatientName, mobjCurStudyInfo.strPatientAge, mobjCurStudyInfo.strPatientSex, mobjCurStudyInfo.strStuStateDesc, Me)
    
    '刷新费用页面
    If TabWindow.Selected.tag = "申请费用" Or TabWindow.Selected.tag = "住院医嘱" Or TabWindow.Selected.tag = "门诊医嘱" Then
        Call RefreshModuleData(mstrSelTabName, mstrSelModuleTag, mobjSelModule)
    End If
End Sub

'传染病登记
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
    
 
    strSQL = "Select  b.内容文本 As 正文 From 电子病历内容 a,电子病历内容 b, 病人医嘱报告 c " & _
             "Where c.医嘱id = [1] And a.内容文本 = '诊断意见' And a.对象类型 = 3 And a.Id = b.父ID " & _
             "And a.文件id = c.病历id And b.对象类型 = 2 And b.终止版 = 0"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "提取报告结果", mobjCurStudyInfo.lngAdviceId)
    
    If rsData.RecordCount > 0 Then strReportResult = NVL(rsData!正文)
    
    If mobjCurStudyInfo.lngPatientFrom = 1 Then        '门诊
        Call mobjPublicAdvice.ShowDisRegist(Me, 0, , mobjCurStudyInfo.lngPatId, , mobjCurStudyInfo.strRegNo, mobjCurStudyInfo.lngAdviceId, mlngCur科室ID, , , , , strReportResult)
    ElseIf mobjCurStudyInfo.lngPatientFrom = 2 Then    '住院
        Call mobjPublicAdvice.ShowDisRegist(Me, 0, , mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.lngPageID, , mobjCurStudyInfo.lngAdviceId, mlngCur科室ID, , , , , strReportResult)
    Else                                            '外来、体检
        Call mobjPublicAdvice.ShowDisRegist(Me, 0, , mobjCurStudyInfo.lngPatId, , , mobjCurStudyInfo.lngAdviceId, mlngCur科室ID, , , , , strReportResult)
    End If
    
    Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_DiseaseRegist") = 1 Then Resume
End Sub

'传染病查询
Private Sub Menu_Manage_DiseaseQuery()
On Error GoTo errhandle
    If mobjPublicAdvice Is Nothing Then
        Set mobjPublicAdvice = DynamicCreate("zlPublicAdvice.clsPublicAdvice", "zlPublicAdvice")
        If mobjPublicAdvice Is Nothing Then Exit Sub
        Call mobjPublicAdvice.InitDisease(gcnOracle, glngSys, gfrmMain, glngModul, gstrPrivs)
    End If
    
    Call mobjPublicAdvice.ShowDisQuery(mlngCur科室ID)

    Exit Sub
errhandle:
    Call HintError(err, "Menu_Manage_DiseaseQuery", False)
End Sub

Private Sub Menu_Manage_修改()
On Error GoTo errhandle
    Dim strOldName As String
    Dim strOldRoom As String
    Dim strQueueName As String
    Dim strCodeNo As String
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_修改", vbInformation
        Exit Sub
    End If
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        With frmRISRequest
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = mobjCurStudyInfo.lngSendNo
            .mlngAdviceId = mobjCurStudyInfo.lngAdviceId
            .mstrPatientName = mobjCurStudyInfo.strPatientName
            .mintEditMode = IIf(mobjCurStudyInfo.intStep > 1, 3, 1)  '0－登记、1－登记后修改、2－报到、3－报到后修改
            .mlngCurDeptId = IIf(mblnAllDepts, mobjCurStudyInfo.lngExeDepartmentId, mlngCur科室ID)
            .mstrCur科室 = zlStr.NeedName(mstrCur科室)
            
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            
            If .mlngResultState <> 0 Then
                strOldName = mobjCurStudyInfo.strPatientName
                strOldRoom = mobjCurStudyInfo.strExeRoom
                
                Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
                
                If mSysPar.blnUseQueue And Not mobjQueue Is Nothing Then
                    '如果是报到后修改，且改变了执行间，则需要重新进行排队
                    If .mintEditMode = 3 And .mlngResultState = 3 Then
                        If .mstrTechnicRoom <> strOldRoom Then
                            If .mstrTechnicRoom = "" Then
                                '如果为空，则需要插入该检查项目对应的项目分组或者科室的队列中
                                Call mobjQueue.zlGetInQueueInf(mobjCurStudyInfo.lngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                            Else
                                '如果不为空，则写入对应的执行间名称
                                strQueueName = .mstrCur科室 & "-" & .mstrTechnicRoom
                                strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                            End If
                            
                            Call mobjQueue.zlUpdatePacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
                        Else
                            '其他方式的修改，则只对排队叫号中的相关信息进行更新
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
            .mintEditMode = IIf(mobjCurStudyInfo.intStep > 1, 3, 1)  '0－登记、1－登记后修改、2－报到、3－报到后修改
            .mlngCurDeptId = IIf(mblnAllDepts, mobjCurStudyInfo.lngExeDepartmentId, mlngCur科室ID)
            .mintImgCount = mintImgCount
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            
            If .RefreshPatiInfor(False) = True Then  '刷新病人
                .mblnOk = False
                .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            End If
            
            If .mblnOk Then Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId) '成功返回
        End With
    End If
Exit Sub
errhandle:
    Call HintError(err, "Menu_Manage_修改", False)
End Sub

Private Sub Menu_Manage_ModifBaseInfo()
'基本信息调整
On Error GoTo errhandle
    Dim zlPubPatient As Object
    
    Dim int场合 As Integer
    Dim str就诊ID As String

    Set zlPubPatient = CreateObject("zlPublicPatient.clsPublicPatient")
    If Not zlPubPatient Is Nothing Then Call zlPubPatient.zlInitCommon(gcnOracle, glngSys)
    
    With mobjCurStudyInfo
        int场合 = Decode(.lngPatientFrom, 1, 1, 2, 2, 3, 3, 4, 4)

        str就诊ID = Decode(.lngPatientFrom, 1, getRegID(.strRegNo), 2, .lngPageID, 3, .lngAdviceId, 4, .strRegNo)

        If zlPubPatient.ModiPatiBaseInfo(Me, mlngModule, .lngPatId, str就诊ID, int场合) Then
            Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
        End If
        
    End With
    
    Set zlPubPatient = Nothing
Exit Sub
errhandle:
    Set zlPubPatient = Nothing
    
    Call HintError(err, "Menu_Manage_ModifBaseInfo", False)
End Sub

Private Sub Menu_Manage_复制登记()
    Dim strQueueName As String
    Dim strCodeNo As String
    Dim lngNewAdviceId As Long
    Dim lngResultState As Long
    
On Error GoTo errhandle
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_复制登记", vbInformation
        Exit Sub
    End If
    
    Call Notify.Broadcast(BM_RIS_EVENT_REGISTER, 0)
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        With frmRISRequest
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = 0
            .mlngAdviceId = 0
            .mintEditMode = 0 '0－登记、1－登记后修改、2－报到、3－报到后修改
            .mlngCurDeptId = IIf(mblnAllDepts, mobjCurStudyInfo.lngExeDepartmentId, mlngCur科室ID)
            .mstrCur科室 = zlStr.NeedName(mstrCur科室)
            .mlngResultState = 0
            
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1), mblnAllDepts, mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo
            
            lngResultState = .mlngResultState
            
            If .mlngResultState <> 0 Then '成功返回
                lngNewAdviceId = .mlngAdviceId
                
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
                
                '如果同时勾选“开始检查自动打开报告”和“登记后自动报到”参数那么会自动打开报告界面
                If mSysPar.blnAutoOpenReport And mSysPar.bln直接检查 Then Call Menu_RichEPR(conMenu_PacsReport_Write)
                
                If .mlngResultState = 2 Then
                    '如果启用排队叫号，则报到后需要插入排队叫号队列......
                    If mSysPar.blnUseQueue And Not mobjQueue Is Nothing Then
                        '设置需要插入的队列名称
                        If .mstrTechnicRoom = "" Then
                            '如果未空，则需要插入该检查项目对应的项目分组或者科室的队列中
                            Call mobjQueue.zlGetInQueueInf(mobjCurStudyInfo.lngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                        Else
                            '如果不为空，则写入对应的执行间名称
                            strQueueName = .mstrCur科室 & "-" & .mstrTechnicRoom
                            strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                        End If
                        
                        Call mobjQueue.zlInPacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
                    End If
                    
                End If
                
                '发送新申请消息
                Call mobjMsgCenter.Send_Msg_Request(.mlngAdviceId)
            End If
        End With
    Else
        With frmPatholRIS
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = 0
            .mlngAdviceId = 0
            .mintEditMode = 0 '0－登记、1－登记后修改、2－报到、3－报到后修改
            .mlngCurDeptId = IIf(mblnAllDepts, mobjCurStudyInfo.lngExeDepartmentId, mlngCur科室ID)
            .mblnOk = False
            lngResultState = 0
            
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            If .CopyCheck(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo) = True Then  '刷新病人
                .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            End If
            
            If .mblnOk Then '成功返回
                lngResultState = 1
                lngNewAdviceId = .mlngAdviceId
                
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
            End If
        End With
    End If
    
    Call Notify.Broadcast(BM_RIS_EVENT_REGISTER, 1, lngNewAdviceId, lngResultState)
Exit Sub
errhandle:
    Call HintError(err, "Menu_Manage_复制登记", False)
End Sub

Private Sub Menu_Manage_登记()
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
            .mintEditMode = 0 '0－登记、1－登记后修改、2－报到、3－报到后修改
            .mlngCurDeptId = IIf(mblnAllDepts, mobjCurStudyInfo.lngExeDepartmentId, mlngCur科室ID)
            .mstrCur科室 = zlStr.NeedName(mstrCur科室)
            .mlngResultState = 0
            
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1), mblnAllDepts
            
            lngResultState = .mlngResultState
            
            If .mlngResultState <> 0 Then '成功返回
                lngNewAdviceId = .mlngAdviceId
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
                
                '如果同时勾选“开始检查自动打开报告”和“登记后自动报到”参数那么会自动打开报告界面
                If mSysPar.blnAutoOpenReport And mSysPar.bln直接检查 Then Call Menu_RichEPR(conMenu_PacsReport_Write)
                
                If .mlngResultState = 2 Then
                    '如果启用排队叫号，则报到后需要插入排队叫号队列......
                    If mSysPar.blnUseQueue And Not mobjQueue Is Nothing Then
                        '设置需要插入的队列名称
                        If .mstrTechnicRoom = "" Then
                            '如果未空，则需要插入该检查项目对应的项目分组或者科室的队列中
                            Call mobjQueue.zlGetInQueueInf(mobjCurStudyInfo.lngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                        Else
                            '如果不为空，则写入对应的执行间名称
                            strQueueName = .mstrCur科室 & "-" & .mstrTechnicRoom
                            strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                        End If
                        
                        Call mobjQueue.zlInPacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
                    End If
                    
                End If
                
                '发送新申请消息
                Call mobjMsgCenter.Send_Msg_Request(.mlngAdviceId)
            End If
            
        End With
    Else
        With frmPatholRIS
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = 0
            .mlngAdviceId = 0
            .mintEditMode = 0 '0－登记、1－登记后修改、2－报到、3－报到后修改
            .mlngCurDeptId = IIf(mblnAllDepts, mobjCurStudyInfo.lngExeDepartmentId, mlngCur科室ID)
            .mintImgCount = 0
            .mblnOk = False
            lngResultState = 0
            
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            
            If .mblnOk Then '成功返回
                lngResultState = 1
                lngNewAdviceId = .mlngAdviceId
            
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
                
                '如果同时勾选“开始检查自动打开报告”和“登记后自动报到”参数那么会自动打开报告界面
                If mSysPar.blnAutoOpenReport And mSysPar.bln直接检查 Then Call Menu_RichEPR(conMenu_PacsReport_Write)
                
                '发送新申请消息
                Call mobjMsgCenter.Send_Msg_Request(.mlngAdviceId)
            End If
        End With
    End If
    
    Call Notify.Broadcast(BM_RIS_EVENT_REGISTER, 1, lngNewAdviceId, lngResultState)
Exit Sub
errhandle:
    Call HintError(err, "Menu_Manage_登记", False)
End Sub

Private Sub Menu_Manage_取消登记()
On Error GoTo errhandle
    Dim strSQL As String
    Dim lngCancelAdviceId As Long
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_取消登记", vbInformation
        Exit Sub
    End If
    
    If HintMsg("确认要取消当前申请吗？" & Chr(10) & Chr(13) & "申请取消后，其对应的医嘱将拒绝执行！", "Menu_Manage_取消登记", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

    lngCancelAdviceId = mobjCurStudyInfo.lngAdviceId
    Call Notify.Broadcast(BM_RIS_EVENT_CANCELREG, 0, lngCancelAdviceId)
    
    strSQL = "ZL_病人医嘱执行_拒绝执行(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & ",null,null," & mlngCur科室ID & ")"
    
    Call zlDatabase.ExecuteProcedure(strSQL, "撤消登记")
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    '发送医嘱撤销消息
    Call mobjMsgCenter.Send_Msg_CancelAdvice(mobjCurStudyInfo.lngAdviceId)
    
    Call Notify.Broadcast(BM_RIS_EVENT_CANCELREG, 1, lngCancelAdviceId)
Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_取消登记") = 1 Then Resume
End Sub

Private Sub Menu_Manage_召回取消()
'功能：召回被取消的登记
On Error GoTo errhandle
    Dim strSQL As String
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_召回取消", vbInformation
        Exit Sub
    End If
    
    If HintMsg("确实要召回被取消登记的项目吗", "Menu_Manage_召回取消", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    strSQL = "ZL_病人医嘱执行_取消拒绝(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & ",null,null," & mlngCur科室ID & ")"
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    '发送状态同步消息
    Call mobjMsgCenter.Send_Msg_StateSync(mobjCurStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_召回取消") = 1 Then Resume
End Sub

Private Sub Menu_Manage_报到()
On Error GoTo errhandle
    Dim blnFocusFind As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim strQueueName As String
    Dim strCodeNo As String
    Dim blnIsCurDayReservations As Boolean '是否今天的预约患者
    Dim strSQL As String
    Dim blnIsClearQueue As Boolean
    Dim lngResultState As Long
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_报到", vbInformation
        Exit Sub
    End If
    
    Call Notify.Broadcast(BM_RIS_EVENT_RECEVIE, 0, mobjCurStudyInfo.lngAdviceId)
    
    blnIsCurDayReservations = False
    blnIsClearQueue = False
    If mblnIsScheduleOrder Then
        '判断是否预约患者
        strSQL = "Select ID,预约开始时间 From 影像预约记录 Where 医嘱Id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检索预约信息", mobjCurStudyInfo.lngAdviceId)
        If rsTemp.RecordCount > 0 Then
            blnIsCurDayReservations = True
            
            '已经预约，则判断预约日期和当前时间是否一致，如果不一致则弹出报到提示
            '如果预约日期和当前日期一致，则直接进入报到
            If Format(NVL(rsTemp!预约开始时间), "yyyy-mm-dd") <> Format(zlDatabase.Currentdate, "yyyy-mm-dd") Then
                If HintMsg("当前患者预约的检查日期为 " & Format(NVL(rsTemp!预约开始时间), "yyyy-mm-dd") & "，与当前时间不一致，是否继续报到？", "Menu_Manage_报到", vbInformation + vbYesNo) = vbNo Then
                    Exit Sub
                Else
                    blnIsClearQueue = True
                    blnIsCurDayReservations = False
                End If
            End If
        End If
    End If
    
    If mobjCurStudyInfo.lngPatientFrom = 4 Then    '如果是体检病人才执行以下过程
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
            .mintEditMode = 2 '0－登记、1－登记后修改、2－报到、3－报到后修改
            .mlngCurDeptId = IIf(mblnAllDepts, mobjCurStudyInfo.lngExeDepartmentId, mlngCur科室ID)
            .mstrCur科室 = zlStr.NeedName(mstrCur科室)
            .mlngResultState = 0
            
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            
            lngResultState = .mlngResultState
            
            If .mlngResultState <> 0 Then  '成功返回
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
                
                If .mblnIsRelationImage = True Then
                    '如果对提前检查的图像进行了自动关联处理，则这里将对影像图像模块进行刷新
                    If Not mobjWork_PacsImg Is Nothing Then
                        Call mobjWork_PacsImg.zlRefreshFace(mobjCurStudyInfo, True)
                    End If
                End If
                
                If mSysPar.blnAutoOpenReport Then Call Menu_RichEPR(conMenu_PacsReport_Write)              '开始检查自动打开报告
                
                If .mlngResultState = 2 Then
                    '如果启用排队叫号，并且报到后自动排队，则报到后需要插入排队叫号队列......
                    If mSysPar.blnUseQueue And mSysPar.blnAutoInQueue And Not mobjQueue Is Nothing Then
                        If blnIsCurDayReservations Then
                            Call mobjQueue.ReservationQueue(.mlngAdviceId)
                        Else
                            If blnIsClearQueue Then
                                '删除之前预约时的排队，后续重新生成队列
                                strSQL = "zl_排队叫号队列_自定义清除(1," & "'业务ID=" & mobjCurStudyInfo.lngAdviceId & "',0)"
                                Call zlDatabase.ExecuteProcedure(strSQL, "删除队列数据")
                            End If
                            
                            '设置需要插入的队列名称
                            If .mstrTechnicRoom = "" Then
                                '如果未空，则需要插入该检查项目对应的项目分组或者科室的队列中
                                Call mobjQueue.zlGetInQueueInf(mobjCurStudyInfo.lngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                            Else
                                '如果不为空，则写入对应的执行间名称
                                strQueueName = .mstrCur科室 & "-" & .mstrTechnicRoom
                                strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                            End If
                            
                            Call mobjQueue.zlInPacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
                        End If
                    End If
                    
                End If
                
                '发送状态同步消息
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
            .mintEditMode = 2 '0－登记、1－登记后修改、2－报到、3－报到后修改
            .mlngCurDeptId = IIf(mblnAllDepts, mobjCurStudyInfo.lngExeDepartmentId, mlngCur科室ID)
            .mintImgCount = mintImgCount
            lngResultState = 0
            
            Call .InitMvar(mobjPacsQueryWrap.PatiColor)
            If .RefreshPatiInfor(True) = True Then  '刷新病人
                .mblnOk = False
                .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            End If
            
            If .mblnOk Then  '成功返回
                lngResultState = 1
                
                Call UpdateQueryListData(Nothing, .mlngAdviceId)
                If mSysPar.blnAutoOpenReport Then Call Menu_RichEPR(conMenu_PacsReport_Write)              '开始检查自动打开报告
                
                '发送状态同步消息
                Call mobjMsgCenter.Send_Msg_StateSync(.mlngAdviceId)
            End If
            
        End With
    End If
    
    Call Notify.Broadcast(BM_RIS_EVENT_RECEVIE, 1, mobjCurStudyInfo.lngAdviceId, lngResultState)
    
Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_报到") = 1 Then Resume
End Sub

'排队叫号入队
Private Sub zlInPacsQueue()
    Dim strQueueName As String
    Dim strCodeNo As String
    
    If mobjQueue Is Nothing Then Exit Sub
    
    '设置需要插入的队列名称
    If Trim(mobjCurStudyInfo.strExeRoom) = "" Then
        '如果未空，则需要插入该检查项目对应的项目分组或者科室的队列中
        Call mobjQueue.zlGetInQueueInf(mobjCurStudyInfo.lngAdviceId, mlngCur科室ID, strQueueName, strCodeNo)
    Else
        '如果不为空，则写入对应的执行间名称
        strQueueName = zlStr.NeedName(mstrCur科室) & "-" & mobjCurStudyInfo.strExeRoom
        strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(mobjCurStudyInfo.strExeRoom, mlngCur科室ID)
    End If
    
    Call mobjQueue.zlInQueue(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.strPatientName, mlngCur科室ID, strQueueName, mobjCurStudyInfo.strExeRoom, strCodeNo)
End Sub




Private Sub Menu_Manage_取消报到()
On Error GoTo errhandle
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim lngResult As Long
    Dim strMsg As String

    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_取消报到", vbInformation
        Exit Sub
    End If
    
  
    If mobjCurStudyInfo.intStep <= 1 Then Call Menu_Manage_取消登记: Exit Sub  '工具栏调用
    '------------------------------------有签名的需要先回退签名后再撤消
    
    Call Notify.Broadcast(BM_RIS_EVENT_CANCELREC, 0, mobjCurStudyInfo.lngAdviceId)
    
    strSQL = "Select Distinct B.完成时间, B.签名级别 From 病人医嘱报告 A, 电子病历记录 B Where A.病历ID=B.Id And A.医嘱ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取是否签名", mobjCurStudyInfo.lngAdviceId)
    
    If Not rsTemp.EOF Then
        If NVL(rsTemp!完成时间, "") <> "" And Val(NVL(rsTemp!签名级别)) > 0 Then '签名保存
            HintMsg "当前病人的检查报告已经签名,若需取消检查,请先回退签名!", "Menu_Manage_取消报到", vbInformation
            Exit Sub
        End If
    End If
    
    '如果检查已取材或者制片，则不能进行取消
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        strSQL = "select count(1) as 数量 from 病理检查信息 a, 病理取材信息 b where a.病理医嘱ID=b.病理医嘱ID and a.医嘱ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, GetWindowCaption, mobjCurStudyInfo.lngAdviceId)
        If rsTemp.RecordCount > 0 Then
            If Val(NVL(rsTemp!数量)) > 0 Then
                HintMsg "该检查已执行取材操作，不能进行取消。", "Menu_Manage_取消报到", vbInformation
                Exit Sub
            End If
        End If
    End If

    If mobjCurStudyInfo.strStudyUID <> "" And Not CheckPopedom(mstrPrivs, "清除图像") Then
        HintMsg "您没有清除检查图像权限,不能请除图像,所以不能取消此项检查!", "Menu_Manage_取消报到", vbInformation
        Exit Sub
    End If
    
    strMsg = "病人信息【姓名：" & mobjCurStudyInfo.strPatientName & "   性别：" & mobjCurStudyInfo.strPatientSex & "   年龄：" & mobjCurStudyInfo.strPatientAge & "   检查号：" & mobjCurStudyInfo.strStudyNum & "】" & vbCrLf & _
             "取消病人本次检查将删除相应的检查图像和检查报告，是否继续？"

    If HintMsg(strMsg, "Menu_Manage_取消报到", vbDefaultButton2 + vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    '取消排队信息
    If mSysPar.blnUseQueue = True And Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Call mobjQueue.zlCancelPacsQueue(mobjCurStudyInfo.lngAdviceId)
    End If
    
    '如果是RIS工作站，而且图像在新网PACS中，则需要先取消关联，然后再调用ZL_影像检查_CANCEL过程取消报到
    If mlngModule = G_LNG_PACSSTATION_MODULE And mobjCurStudyInfo.intImageLocation = 1 Then
        '取消图像关联
        Call XWUnmatchImage(mobjCurStudyInfo.lngAdviceId, 0)
    End If
    
    '取消报告，修改数据库状态，删除“影像检查记录”
    strSQL = "ZL_影像检查_CANCEL(" & mobjCurStudyInfo.lngAdviceId & "," & mobjCurStudyInfo.lngSendNo & ",0," & mlngCur科室ID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        strSQL = "ZL_病理检查_撤销(" & mobjCurStudyInfo.lngAdviceId & ")"
        zlDatabase.ExecuteProcedure strSQL, GetWindowCaption
    End If
    
    '如果图像在中联PACS，则删除影像文件和目录
    If mobjCurStudyInfo.intImageLocation = 0 Then
        RemoveCheckImages mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo
    End If
    
    If TabWindow.Selected.tag = "影像采集" Then
        'TODO:如果进行了自动关联，则需要刷新helper的缩略图像
'        If Not mobjWork_ImageCap Is Nothing Then
'            Call mobjWork_ImageCap.zlRefreshData(True)
'        End If
    End If
    
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
    
    '发送状态回退消息
    Call mobjMsgCenter.Send_Msg_StateCancel(mobjCurStudyInfo.lngAdviceId)
    
    Call Notify.Broadcast(BM_RIS_EVENT_CANCELREC, 1, mobjCurStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_取消报到") = 1 Then Resume
End Sub

Private Sub Menu_Manage_关联影像()
On Error GoTo errhandle
    Dim lngResult As Long
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_关联影像", vbInformation
        Exit Sub
    End If

    lngResult = -1
    '如果是模块号为RIS工作站，则调用新网的数据库查询未匹配的图像记录
    If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
        lngResult = XWShowUnMatched(Me, mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.strImgType)
        
        If lngResult = 0 Then
            '图像关联成功后,使其值为1
            mobjCurStudyInfo.intImageLocation = 1
            Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
        End If
    Else
        frmSelectMuli.ShowImageReleation mlngModule, mobjCurStudyInfo.lngAdviceId, mstrPrivs, mobjCurStudyInfo.blnMoved, IIf(mlngModule = G_LNG_PACSSTATION_MODULE, False, True), mlngCur科室ID, 2, mobjCurStudyInfo.strImgType
        
        If Not frmSelectMuli.mblnOk Then Exit Sub
        lngResult = 0
    End If
    
    If lngResult <> 0 Then Exit Sub
    
    Call ReleationImage(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, 2, True)
Exit Sub
errhandle:
    Call HintError(err, "Menu_Manage_关联影像", False)
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
    
    strFontType = IIf(IsUseClearType = True, "微软雅黑", "宋体")
    
    CtlFont.Name = strFontType
    CtlFont.Size = gbytFontSize
            
    mstrSelQueueRooms = ""
    
    If mlngCur科室ID <> Control.DescriptionText Or (Control.DescriptionText <> 0 And mblnAllDepts = True) Then
        mstrRPTExecutor = UserInfo.姓名
        
        stbThis.Panels(4).Text = "报告医生：" & mstrRPTExecutor & "   检查医生：" & Split(stbThis.Panels(4).Text, "检查医生：")(1)
                
        Set mobjCurStudyInfo = GetNullAdviceInf
        
        '科室切换后，由于没有重新创建菜单和工作模块，也没有调用cbrMain.RecalcLayout，因此需要使用该对象设置科室切换后的科室信息
        Set objDepartmentMenu = cbrMain.FindControl(, conMenu_View_Filter * 10#)
        
        If Control.DescriptionText = 0 Then
            '选择所有科室
            mblnAllDepts = True
            mlngCur科室ID = 0
        
            If Not objDepartmentMenu Is Nothing Then objDepartmentMenu.Caption = "全部科室"
            
            Call mobjPacsQueryWrap.DepChange(mstrCanUse科室IDs, True)
            Set cbrFilter.options.Font = CtlFont
            
            If Not mobjQueue Is Nothing And mlngModule = G_LNG_PACSSTATION_MODULE Then
                mobjQueue.ChangeToAllDept mblnAllDepts
            End If
        Else
            '选择单个科室
            mblnAllDepts = False
            
            mlngCur科室ID = Control.DescriptionText
            mstrCur科室 = Mid(Control.Caption, 1, InStrRev(Control.Caption, "(") - 1)
             
            If Not objDepartmentMenu Is Nothing Then objDepartmentMenu.Caption = mstrCur科室
            
            Call SetParaUseImgSignValid(mlngCur科室ID)
            Call InitDeptParameter(mlngCur科室ID)
            
            Call ucPacsHelper1.Init(Me, mlngModule, mlngCur科室ID, mstrPrivs, True)

            
            If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.zlInitModule(Me, mlngModule, mstrPrivs, mlngCur科室ID)
            If Not mobjWork_PacsImg Is Nothing Then Call mobjWork_PacsImg.zlInitModule(Me, mlngModule, mstrPrivs, mlngCur科室ID)
            If Not mobjWork_His Is Nothing Then Call mobjWork_His.zlModule.zlInitModule(Me, mlngModule, mstrPrivs, mlngCur科室ID)
            If Not mobjWork_ImageCap Is Nothing Then Call mobjWork_ImageCap.zlInitModule(gcnOracle, mobjCapLinker, glngSys, mlngModule, mstrPrivs, mlngCur科室ID, Me.hwnd, gblnUseDebugLog)
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.zlInitModule(Me, mlngModule, mstrPrivs, mlngCur科室ID, mobjCapLinker, ucPacsHelper1)

     
            '如果报告编辑器类型不同，则需要初始化，不同科室可能使用不同的报告编辑器
            If TabWindow.Selected.Caption = C_TAB_NAME_检查报告 Then
                Call SelectModule(C_TAB_NAME_检查报告, strModuleTag, True)
                TabWindow.Selected.tag = strModuleTag
            Else
                '只需要将报告模块的tag设置为空，因为报告模块不同科室可能使用不同的编辑界面
                For i = 1 To TabWindow.ItemCount
                    If TabWindow.Item(i).Caption = C_TAB_NAME_检查报告 Then
                        TabWindow.Item(i).tag = ""
                        Exit For
                    End If
                Next
            End If
            
            
            If Not mobjCapLinker Is Nothing Then Call mobjCapLinker.Init(Me, mlngCur科室ID, mstrPrivs)


            '科室切换后，如果启用了排队叫号，则添加排队叫号页面
            If mSysPar.blnUseQueue = True Then
                Call CreateTabItem(13, C_TAB_NAME_排队叫号, 10011, "")
                
                If mobjQueue Is Nothing Then
'                    mstrWorkModule = mstrWorkModule & ";排队叫号模块;"
'
'                    Set mobjQueue = New frmWork_Queue
'                    Call mobjQueue.zlInitPacsQueueCfg(mlngModule, mlngCur科室ID, zlStr.NeedName(mstrCur科室), mstrPrivs, mblnAllDepts, Me)
'
'                    TabWindow.InsertItem 13, "排队叫号", mobjQueue.hwnd, 10011
'                    TabWindow.Item(TabWindow.ItemCount - 1).tag = "排队叫号"
'
'                    Call picWindow_Resize
                    Call VerifyModuleObj(C_TAB_NAME_排队叫号)
                Else
                    Call mobjQueue.zlInitPacsQueueCfg(mlngModule, mlngCur科室ID, zlStr.NeedName(mstrCur科室), mstrPrivs, mblnAllDepts, Me)
                End If
                
                
                Call picTabFace_Resize
                
                '快捷叫号界面
                If mSysPar.blnQueueQuick Then
                    If Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
                        Call mobjQueue.OpenQueueQuick(GetSelQueueRooms(True), Me)
                    End If
                End If
            Else
                If mSysPar.blnUseQueue = False And Not mobjQueue Is Nothing Then
                    mstrWorkModule = Replace(mstrWorkModule, ";排队叫号模块;", "")
                    
                    For i = 0 To TabWindow.ItemCount - 1
                        If TabWindow.Item(i).tag = "排队叫号" Then
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
            

            '切换消息的接收科室
            Call mobjMsgCenter.ChangeMsgReceiveDept(mlngCur科室ID)
            
            With mobjPacsQueryWrap.CurPacsQuery.GetSqlScheme
                strOldSchemeValue(0) = .Query
                strOldSchemeValue(1) = .FilterCfgCount
                strOldSchemeValue(2) = .Detail
                strOldSchemeValue(3) = .SerachCfgCount
                strOldSchemeValue(4) = .ShowCfgCount
            End With
            
            Call mobjPacsQueryWrap.DepChange(mlngCur科室ID, False)
            
            '判断是否需要切换方案
            Call mobjPacsQueryWrap.CurPacsQuery.LoadQueryScheme(glngSys, mlngModule, mlngCur科室ID, UserInfo.ID)
            
            Call ExecuteDefaultQueryScheme
            
            Set cbrMenuBar = cbrMain.FindControl(, conMenu_Manage_Query)
            
            Call mobjPacsQueryWrap.RefreshCustomQueryMenu(cbrMenuBar, mintQueryState, tabScheme, mSysPar.blnQuickTabDisplayScheme)
            With cbrMenuBar.CommandBar
                Set objControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_QueryCFG, "查询配置", "", 0, True)
                Set objControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_QueryCfgUserScheme, "常用方案调整", "", 0, False)
                Set objControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_QueryTabDisplayScheme, "显示常用方案标签", "", 0, True)
                
                objControl.Checked = mSysPar.blnQuickTabDisplayScheme
                objControl.CloseSubMenuOnClick = False
            End With
            
            Set cbrFilter.options.Font = CtlFont
        End If
        
        Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
         
        
        Call cbrMain.RecalcLayout
        
        '刷新排队叫号模块数据，如果已经启用
        Call RefreshPacsQueueData(False)
        
        Call CreateAuditorMenu(cbrMain.FindControl(, conMenu_ManagePopup).CommandBar.FindControl(, conMenu_Manage_SendAudit))
        
        If CheckPopedom(mstrPrivs, "检查预约") Then
            '刷新是否启用预约
            Call IsSchedule(mlngCur科室ID, mobjCurStudyInfo.lngAdviceId, 0, mblnIsScheduleDept, mblnIsScheduleOrder)
        Else
            mblnIsScheduleDept = False
            mblnIsScheduleOrder = False
        End If
    End If
    
    If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
        glngXWDeptID = mlngCur科室ID
    End If
    
    mblnInitOk = True
    
    '恢复模块页签显示
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
    strSQL = "Select a.id,a.名称 as 程序名称,a.是否启用 as 程序启用,a.执行类型,b.功能序号,b.名称 as 功能名称,b.是否启用 as 功能启用,b.是否加入右键菜单,b.是否加入工具栏,b.vbs脚本 from 影像插件挂接 a, 影像插件功能 b " & _
             "Where a.是否启用=1 and  b.是否启用=1 and a.id = b.插件id And (a.所属模块=0 or a.所属模块=[1]) Order By a.id,b.功能序号"
             
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "创建插件工具栏菜单", lngModule)
    
    If rsTemp.RecordCount > 0 Then

        While Not rsTemp.EOF
                
            j = j + 1
            
            If Val(NVL(rsTemp!是否加入工具栏)) = 1 Then
                If blFirst = True Then
                    Set cbrControl = CreateMenu(cbrControls, xtpControlButton, conMenu_Manage_PacsPlugIn * 10000# + j, NVL(rsTemp!功能名称), "", 2325, True)
                    blFirst = False
                Else
                    Set cbrControl = CreateMenu(cbrControls, xtpControlButton, conMenu_Manage_PacsPlugIn * 10000# + j, NVL(rsTemp!功能名称), "", 2325, False)
                End If
                
                cbrControl.Parameter = NVL(rsTemp!VBS脚本)
                cbrControl.DescriptionText = Val(NVL(rsTemp!执行类型))
                cbrControl.Category = Val(NVL(rsTemp!功能启用)) & "," & Val(NVL(rsTemp!是否加入右键菜单)) & "," & Val(NVL(rsTemp!是否加入工具栏))
            End If
            
            Call rsTemp.MoveNext
        Wend
    End If
            
    Exit Sub
ErrorHand:
    Call err.Raise(0, , "插件菜单添加到工具栏异常-" & err.Description)
End Sub

Private Sub RefreshCustomPlugInMenu(objQueryMenu As Object, ByVal lngModule As Long)
    Dim objCurQueryMenu As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim blFirstMenu As Boolean '是否第一个功能菜单（用于判断是否需要加分割线）
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
    
    strSQL = "Select a.id,a.名称 as 程序名称,a.是否启用 as 程序启用,a.执行类型,b.功能序号,b.名称 as 功能名称,b.是否启用 as 功能启用,b.是否加入右键菜单,b.是否加入工具栏,b.vbs脚本 from 影像插件挂接 a, 影像插件功能 b " & _
             "Where a.id = b.插件id and a.是否启用=1 and b.是否启用=1 And (a.所属模块=0 or a.所属模块=[1]) Order By a.id,b.功能序号"
             
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "创建插件菜单", lngModule)
    
    With objCurQueryMenu.CommandBar
        If rsTemp.RecordCount > 0 Then
            i = 65
            While Not rsTemp.EOF
                j = j + 1
                
                If lngAppId <> NVL(rsTemp!ID) Then
                    Set cbrPopControl = CreateMenu(.Controls, xtpControlButtonPopup, conMenu_Manage_PacsPlugLevel2 * 10000# + NVL(rsTemp!ID), NVL(rsTemp!程序名称), "", , False)
                    lngAppId = NVL(rsTemp!ID)
                Else
                    Set cbrPopControl = cbrMain.FindControl(, conMenu_Manage_PacsPlugLevel2 * 10000# + NVL(rsTemp!ID), , True)
                End If

                If Not cbrPopControl Is Nothing Then
                    If blFirstMenu Then
                        Set cbrControl = CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsPlugIn * 10000# + j, NVL(rsTemp!功能名称), "", , True)
                    Else
                        Set cbrControl = CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsPlugIn * 10000# + j, NVL(rsTemp!功能名称), "", , False)
                    End If
                End If
                                
                cbrControl.Parameter = NVL(rsTemp!VBS脚本)
                cbrControl.DescriptionText = Val(NVL(rsTemp!执行类型))
                cbrControl.Category = Val(NVL(rsTemp!功能启用)) & "," & Val(NVL(rsTemp!是否加入右键菜单)) & "," & Val(NVL(rsTemp!是否加入工具栏))
                
                blFirstMenu = False
                
                Call rsTemp.MoveNext
            Wend
        End If
            
    End With

    Exit Sub
ErrorHnad:
    Call err.Raise(0, , "更新插件菜单异常-" & err.Description)
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
'通过可用的科室串，读取指定科室ID的科室名称
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
    Call mobjPacsQueryWrap.ExecuteQuery(C_QUERY_数据检索)
    If mobjPacsQueryWrap.SqlScheme.AutoRefreshTimeLen > 0 Then TimerRefresh.Enabled = True
End Sub

Private Sub cmdMore_Click()
    Call mobjPacsQueryWrap.ExecuteQuery(C_QUERY_更多过滤)
    If mobjPacsQueryWrap.SqlScheme.AutoRefreshTimeLen > 0 Then TimerRefresh.Enabled = True
End Sub

Private Sub Form_Activate()
On Error GoTo errhandle
    Dim lngHwnd As Long
    Dim blnIsEmbedReport As Boolean
    
    '判断当前工作模块是否影像采集模块，如果是，则判断采集模块是否初始化，如果已经初始化，则退出该过程，否则就对其进行初始化，并显示
    '因为在同一导航台中，如果同时打开病理，视频采集模块将被切换，当另一系统退出时，采集模块也将被释放，因此切换回当前系统后，需要判断是否从新初始化采集模块
'    Call Form_Resize

    If Not mblnInitOk Then Exit Sub
    
    If TabWindow.Selected Is Nothing Then Exit Sub
    
    '注：如果弹出式报告窗口中嵌入了视频采集，则主界面窗口不切换嵌入式视频采集的显示
    If mstrSelTabName = C_TAB_NAME_影像采集 Then
        '只有工作模块是影像采集时，才需要切换嵌入式视频采集的显示
        If mobjWork_ImageCap Is Nothing Then Exit Sub
        
        
        '如果弹出了浮动采集窗口，则不进行嵌入式处理
        If mobjWork_ImageCap.VideoDockState Then
            mobjCapLinker.ReportAdviceId = 0
            Exit Sub
        End If
        
        lngHwnd = VerifyModuleObj(C_TAB_NAME_影像采集)
        
        '如果视频本来就在容器内，则不需要重新嵌入
        If GetAncestor(mobjWork_ImageCap.VideoHwnd, GA_PARENT) = lngHwnd Then
            mobjCapLinker.ReportAdviceId = 0
            Exit Sub
        End If
          
        '如果视频采集没有嵌入报告窗口，则将视频采集嵌入当前窗口中
        If VideoIsAttachReportWindow = False Then
            Call EmbedWindow(lngHwnd)
            
            mobjCapLinker.ReportAdviceId = 0
            
            '需要调用此方法显示出当前视频
            Call mobjWork_ImageCap.zlRefreshVideoWindow
            
            Call mobjWork_ImageCap.zlRestoreWindow(IIf(mobjCurStudyInfo.intStep > 1 And mobjCurStudyInfo.intStep < 5, False, True), True)
        End If
    Else
        '如果弹出式报告书写窗口没有嵌入视频采集，则在工作模块之间切换时，需要嵌入视频采集
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
'处理内容：执行查询后没有数据时，清空界面控件显示
On Error GoTo errhandle
    Dim i As Long
    
    If vsfList.Rows < 2 Then
        '当没有数据时，通知刷新工作模块中相关的数据
        Set mobjCurStudyInfo = GetNullAdviceInf
        Call RefreshModuleData(mstrSelTabName, mstrSelModuleTag, mobjSelModule)

        mblnIsLoading = False
        
        '左下角TAb处理  附加信息  历史检查 状态图
        
        For i = imgState.Count - 1 To 0 Step -1
            imgState(i).Visible = False
        Next
    
        imgStep.Visible = False
        LabFlag费用.Visible = False
        LabFlag婴儿.Visible = False
        LabFlag绿色通道.Visible = False
        LabFlag危机状态.Visible = False
        LabFlag传染病状态.Visible = False
        LabFlag急诊.Visible = False
        
        labCollectionInfo.Visible = False
        labPatientInfo.Visible = False
        labPatientAge.Visible = False
        
        
        Call mobjPacsQueryWrap.FillAppend(0, 0, False, rtxtAppend)
        
        stbThis.Panels(2).Text = "共 " & 0 & " 条记录": stbThis.Panels(2).Alignment = sbrCenter
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
'规则： 数据检索容器宽度大于9000的时候，按钮与左边间隔达到600并且不会增长，中间按比例增加
'初始状态：
On Error GoTo errhandle
    Dim intTMP As Single '不用字体适当增加按钮到查询界面的距离
    Dim lngWidth As Integer '查询界面宽度
    Dim lngBaseWidth As Long '按钮和查询界面的距离
    Dim lngBaseWidthDataSearchContainer As Long '容器基础宽度
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
    HintError err, "cmdFind_Click<查找操作>", False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '加载工作模块时，不允许退出窗口
    If Not mblnInitOk Then
        Cancel = True
        Exit Sub
    End If
    
    If mblnMenuDownState Then
        If HintMsg("当前操作尚未完成，强制退出可能造成程序异常，是否继续？", "Form_QueryUnload", vbYesNo) = vbNo Then Cancel = True
    End If
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace(picList.Height, picList.Width)
End Sub

Private Sub imgFun_Click(Index As Integer)
'目前提供四个： "报到" 取消报到"  修改信息" 书写报告"
On Error GoTo errH
    Dim i As Integer
    
    If mblnMenuDownState Then Exit Sub

    Select Case imgFun(Index).ToolTipText
        Case C_FUNC_STR_报到
            Call Menu_Manage_报到
            
        Case C_FUNC_STR_书写报告
            Call Menu_RichEPR(conMenu_PacsReport_Write)
            
        Case C_FUNC_STR_查看病人信息
            frmDegreeCard.ShowMe mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.lngPageID, Me
            
        Case C_FUNC_STR_观片
            If mobjPacsCore Is Nothing Then
                HintMsg "图像查看对象无效，不能进行此操作。", "cbrMain_Execute", vbOKOnly
                Exit Sub
            End If
            
            Call OpenViewer(1, mobjPacsCore, mobjCurStudyInfo.lngAdviceId, True, Me, "", mobjCurStudyInfo.blnMoved)
             
        Case C_FUNC_STR_完成
            Call Menu_Manage_检查最终完成
    End Select
    Exit Sub
errH:
    Call HintError(err, "imgFun_Click", False)
End Sub

Private Sub mfrmRISRequest_HaveRegist()
    Dim strQueueName As String
    Dim strCodeNo As String
    With mfrmRISRequest
        If .mlngResultState <> 0 Then '成功返回
            '如果启用排队叫号，则报到后需要插入排队叫号队列......
            If mSysPar.blnUseQueue And Not mobjQueue Is Nothing And .mlngResultState = 2 Then
                '设置需要插入的队列名称
                If .mstrTechnicRoom = "" Then
                    '如果未空，则需要插入该检查项目对应的项目分组或者科室的队列中
                    Call mobjQueue.zlGetInQueueInf(mobjCurStudyInfo.lngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                Else
                    '如果不为空，则写入对应的执行间名称
                    strQueueName = .mstrCur科室 & "-" & .mstrTechnicRoom
                    strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                End If
                
                Call mobjQueue.zlInPacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
            End If
            
            '发送新申请消息
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
    
    '不能直接在Hook回调过程中使用ActiveExe对象的相关方法，否则会产生未知界面错误
    timerCapture.Enabled = True
Exit Sub
errhandle:
    Call HintError(err, "mobjCaptureHot_OnKeyBoardLHook", False)
End Sub

Private Sub mobjMsgCenter_OnRecevieMsg(ByVal strMsgItemIdentity As String, ByVal strXmlContext As String, rsData As ADODB.Recordset, objMsgPro As clsMipModule, objXML As clsXML)
'消息接收处理
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
    
    '获取消息中对应的医嘱ID数据
    If strMsgItemIdentity = G_STR_MSG_ZLHIS_PACS_003 Then
        rsData.Filter = "node_name='study_order_id'"
    Else
        rsData.Filter = "node_name='order_id'"
    End If
    
    If rsData.RecordCount > 0 Then
        lngAdviceId = Val(NVL(rsData!node_value))
    End If
    
    
    Select Case strMsgItemIdentity
        Case G_STR_MSG_ZLHIS_CIS_017    '检查申请
            '弹出消息提示@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='patient_name'"
            strHint = "患者 " & NVL(rsData!node_value) & " 需要进行检查，请及时处理。"
            
            Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
            
            '从数据库中刷新数据
            Call UpdateQueryListData(Nothing, lngAdviceId)
            
        Case G_STR_MSG_ZLHIS_CIS_024    '医嘱撤销
            '弹出撤销提示@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='patient_name'"
            strHint = "患者 " & NVL(rsData!node_value) & " 的检查医嘱已被撤销。 "
        
            Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
        
            '从数据库中刷新数据
            Call UpdateQueryListData(Nothing, lngAdviceId)
            
        Case G_STR_MSG_ZLHIS_CIS_025    '危急值阅读
            '由消息平台配置弹出提示
            
        Case G_STR_MSG_ZLHIS_CHARGE_003 '门诊患者划价单据
            '刷新收费状态显示
            '根据单据号查找对应的医嘱ID
            rsData.Filter = "node_name='bill_no'"
            If rsData.RecordCount <= 0 Then
                Exit Sub
            End If
            
             lngRowIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("医嘱ID"))
            If lngRowIndex > 0 Then Call UpdateQueryListData(Nothing, lngAdviceId)
        
        Case G_STR_MSG_ZLHIS_PACS_001   '检查报告完成，检查完成才算检查报告最终完成
            '更新列表中的显示状态
            lngRowIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("医嘱ID"))
            If lngRowIndex > 0 Then Call UpdateQueryListData(Nothing, lngAdviceId)
            
        Case G_STR_MSG_ZLHIS_PACS_002, G_STR_MSG_ZLHIS_PACS_003  '检查状态同步与检查状态回退处理
            '如果报告被驳回，需要弹出提醒@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='study_cur_state'"
            If NVL(rsData!node_value) = -1 Then
                
                '需要判断当前用户是否为报告人
                strSQL = "select 报告人 from 影像检查记录 where 医嘱ID=[1]"
                Set rsReport = zlDatabase.OpenSQLRecord(strSQL, "查询报告人", lngAdviceId)
                If rsReport.RecordCount > 0 Then
                    If NVL(rsReport!报告人) = UserInfo.姓名 Then
                        '弹出消息
                        rsData.Filter = "node_name='patient_name'"
                        strHint = "患者" & NVL(rsData!node_value) & "的报告已被驳回，请注意处理。"
                        
                        Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
                    End If
                End If
            End If
            
            '刷新列表对应显示
            lngRowIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("医嘱ID"))
            
            If lngRowIndex > 0 Then Call UpdateQueryListData(Nothing, lngAdviceId)
        Case G_STR_MSG_ZLHIS_PACS_004   '检查报告撤销
            '更新列表中的显示状态
            lngRowIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("医嘱ID"))
            
            If lngRowIndex > 0 Then Call UpdateQueryListData(Nothing, lngAdviceId)
            
        Case G_STR_MSG_ZLHIS_PACS_005   '检查危急值通知
            '在科室内弹出危急提醒@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='patient_name'"
            strHint = "患者 " & NVL(rsData!node_value) & "的"
            
            rsData.Filter = "node_name='check_item_title'"
            strHint = strHint & "检查项目 " & NVL(rsData!node_value) & " 产生危急情况。"
            
            Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
        
            '更新列表中的显示状态
            lngRowIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("医嘱ID"))
            
            If lngRowIndex > 0 Then Call UpdateQueryListData(Nothing, lngAdviceId)
            
    End Select
    
    Exit Sub
errH:
    If HintError(err, "mobjMsgCenter_OnRecevieMsg") = 1 Then Resume
End Sub

Private Sub mobjPacsCore_AfterSaveOuterImage(strStudyUID As String)
    '保存了外部图像，刷新图像的序列列表
On Error GoTo errhandle
    
    '没有记录则退出
    If mobjCurStudyInfo.lngAdviceId = 0 Then Exit Sub
    
    '是当前的检查，才刷新检查的序列列表
    If mobjCurStudyInfo.strStudyUID = strStudyUID Then
        Call mobjWork_PacsImg.zlRefreshFace(mobjCurStudyInfo, True)
    End If
    
    Exit Sub
errhandle:
    '不处理
End Sub


Private Sub ReleationImage(ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal intStep As Integer, ByVal lngReleationType As Long, ByVal blnUseMenuReleation As Boolean)
On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    If lngReleationType = 1 Then
        If InStr("345", intStep) > 0 Then
            gstrSQL = "Select 检查uid From 影像检查记录 Where  医嘱ID=[1] And 发送号=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngAdviceId, lngSendNo)
            
            If rsTemp.RecordCount > 0 Then
                If IsNull(rsTemp!检查UID) Then
                    '设置影像检查状态，如果当前医嘱已经没有图像，而且检查过程为3，则修改为2
                    If intStep = 3 Then
                        gstrSQL = "Zl_影像检查_State(" & lngAdviceId & "," & lngSendNo & ",2,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
                        zlDatabase.ExecuteProcedure gstrSQL, "取消关联"
                    End If
                End If
            End If
        End If
    Else
        '设置影像检查状态，如果原来的状态是已报到，则修改成已检查，
        If intStep = 2 Then
            '如果病人已经有图像，则修改成已检查
            strSQL = "Select 检查UID From 影像检查记录 Where 医嘱ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否有图像", lngAdviceId)
            
            If Not IsNull(rsTemp!检查UID) Then
                strSQL = "Zl_影像检查_State(" & lngAdviceId & "," & lngSendNo & ",3,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
                zlDatabase.ExecuteProcedure strSQL, "关联影像"
            End If
        End If
    End If
    
    Call UpdateQueryListData(Nothing, lngAdviceId)
    
    Exit Sub
errH:
    If HintError(err, "ReleationImage") = 1 Then Resume
End Sub

Private Sub mobjPacsQueryWrap_OnColStatistics(ByVal strStatisticsInfo As String)
    stbThis.Panels(2).Text = "共 " & vsfList.Rows - 1 & " 条记录": stbThis.Panels(2).Alignment = sbrCenter
    stbThis.Panels(3).Text = strStatisticsInfo
End Sub

Private Sub mobjPacsQueryWrap_OnDoStateImage(ByVal lngRow As Long)
'处理状态图
On Error GoTo errH
    Dim i As Integer, j As Integer, k1 As Integer, k2 As Integer
    Dim objClsRelation As New clsScRowRelation
    Dim intImgCount As Integer
    Dim lngLeft As Long
    Dim strValue As String
    
    '首先清空状态图
    For i = imgState.Count - 1 To 0 Step -1
        imgState(i).Visible = False
    Next
    intImgCount = 0

    If mobjPacsQueryWrap Is Nothing Then Exit Sub
    If mobjPacsQueryWrap.SqlScheme Is Nothing Then Exit Sub
    If mobjPacsQueryWrap.SqlScheme.ShowCfgCount < 1 Then Exit Sub
    
    With mobjPacsQueryWrap.SqlScheme
        
        For i = 1 To .ShowCfgCount 'i 遍历列显示配置
            If .ShowCfg(i).RowRelationCount > 0 Then
                
                For j = 1 To .ShowCfg(i).RowRelationCount 'j遍历行关联
                
                    Set objClsRelation = .ShowCfg(i).RowRelation(j)
                    If Len(objClsRelation.Icon) > 0 And objClsRelation.IsStateIcon Then '首先判断是否配置了显示图标
                        
                        strValue = vsfList.Cell(flexcpText, lngRow, vsfList.ColIndex(.ShowCfg(i).Name))
                        
                        If (LTrim(strValue) = objClsRelation.TiggerData And objClsRelation.TiggerData <> "[非空]" And objClsRelation.TiggerData <> "[空]") _
                    Or (Len(Trim(strValue)) = 0 And objClsRelation.TiggerData = "[空]") Or (Len(Trim(strValue)) > 0 And objClsRelation.TiggerData = "[非空]") Then
                    
                            '添加状态图
                            If intImgCount = 0 Then
                                Set imgState(0).Picture = mobjPacsQueryWrap.GetIcon(objClsRelation.Icon)
                                Call imgState(0).Move(picDetail.Width - imgState(0).Width, C_LAYOUT_BASEHEIGHTOFDETAILINFO - GetMaxImgHeight - 30)
                                imgState(0).Visible = True
                                
                                intImgCount = 1
                            Else
                                If imgState.Count <= intImgCount Then Load imgState(intImgCount)

                                Set imgState(intImgCount).Picture = mobjPacsQueryWrap.GetIcon(objClsRelation.Icon)

'                                重新设置位置
                                lngLeft = 0
                                For k1 = intImgCount To 0 Step -1
                                    '首先计算已经存在的图标的宽度之和
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
'右键菜单处理说明：使用查询类的事件中调用而不是列表控件直接调用原因： 查询类里面的OnMouseUp会处理工功能跟随功能
'pacsmain这边处理弹出右键菜单功能，如果两边都使用vsflist_onMouseUp有位置的风险。
On Error GoTo errH
    Dim Control As CommandBarControl, Menucontrol As CommandBarControl
    Dim controlPlugIn As CommandBarControl
    Dim plugins As CommandBarControl
    Dim Popup As CommandBar
    Dim strTmp As String
    Dim i As Long
    
    If mobjPacsQueryWrap.ShowingRowCount < 1 Then Exit Sub

    If Button = 2 Then
        Set Popup = cbrMain.Add("右键菜单", xtpBarPopup)


        For i = 1 To cbrMain.ActiveMenuBar.Controls.Count
            Set Menucontrol = cbrMain.ActiveMenuBar.Controls(i)

            If (Menucontrol.ID = conMenu_ManagePopup Or Menucontrol.ID = conMenu_Collection) And Menucontrol.type = xtpControlPopup Then
                For Each Control In Menucontrol.CommandBar.Controls
                    '处理右键 "收藏到" 菜单
                    If Control.ID <> conMenu_Collection_ViewShare And Control.ID <> conMenu_Collection_Manage _
                    And Mid(Control.ID, 1, Decode(InStr(Control.ID, "0") - 1, -1, 0, InStr(Control.ID, "0") - 1)) <> comMenu_Collection_Type _
                    And Mid(Control.ID, 1, Decode(InStr(Control.ID, "0") - 1, -1, 0, InStr(Control.ID, "0") - 1)) <> conMenu_Collection_ViewShare Then
                        '在无报告完成之前，插入模块创建的右键菜单
                        If Control.ID = conMenu_Manage_Complete Then
                            If Not mobjWork_PacsImg Is Nothing Then Call mobjWork_PacsImg.zlMenu.zlPopupMenu(mstrSelTabName, Popup)
                            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.zlMenu.zlPopupMenu(mstrSelTabName, Popup)
                        End If

                        Control.Copy Popup
                    End If
                Next
            ElseIf Menucontrol.ID = conMenu_Manage_PacsPlugIn Then
                For Each Control In Menucontrol.CommandBar.Controls '遍历二级菜单
                    If Control.ID >= conMenu_Manage_PacsPlugLevel2 * 10000# And Control.ID <= conMenu_Manage_PacsPlugLevel2 * 10000# + 9999 Then

                        For Each controlPlugIn In Control.CommandBar.Controls

                            If UBound(Split(controlPlugIn.Category, ",")) = 2 Then '遍历末级菜单
                                strTmp = Split(controlPlugIn.Category, ",")(1)
                            Else
                                strTmp = controlPlugIn.Category
                            End If
                            
                            If plugins Is Nothing Then
                                Set plugins = Popup.Controls.Add(xtpControlPopup, conMenu_Manage_PacsPlugIn, "插件功能")
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
'恢复主要工作页面，如果设置了主要工作页面，切换检查时首先切换到对应页面
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
'blnRefreshModul 是否需要刷新模块

    Dim i As Integer
    Dim intCol As Integer
    Dim lngRow As Long
    Dim lngAdviveID As Long '医嘱ID
    Dim strInfo As String
    Dim blnRefreshFace   As Boolean '是否需要刷新界面
    Dim strCurModuleTag As String
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''刷新变量信息
    
    If mblnIsPrintMode Then Exit Sub
    
    intCol = vsfList.ColIndex("医嘱ID")
    
    lngRow = vsfList.RowSel
    If lngRow = -1 Then Exit Sub

    lngAdviveID = Val(vsfList.TextMatrix(lngRow, intCol))
 
    Set mobjCurStudyInfo = mobjPacsQueryWrap.StudyInfo
    mobjCurStudyInfo.lngReportEditState = GetReportEditState(mobjCurStudyInfo)
    
    If blnIsSelChange Then Call LocateMainWorkModuleTab
    
    Call DoLabFlag
    
    mintImgCount = GetScanRequestCount(mobjCurStudyInfo.lngAdviceId)
    
     
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''刷新界面信息
    '更新详细信息
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
        If mobjCurStudyInfo.strMarkNum > 0 Then labCollectionInfo = "门:" & mobjCurStudyInfo.strMarkNum & "  "
    ElseIf mobjCurStudyInfo.lngPatientFrom = 2 Then
        If mobjCurStudyInfo.strMarkNum > 0 Then labCollectionInfo = "住:" & mobjCurStudyInfo.strMarkNum & "  "
    Else
        labCollectionInfo = ""
    End If
    
    labCollectionInfo = labCollectionInfo & mobjCurStudyInfo.strAdviceContext
    labCollectionInfo = labCollectionInfo & IIf(mobjCurStudyInfo.strCollectionInfo = "", "", "  (◆" & mobjCurStudyInfo.strCollectionInfo & ")")
    
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
        Case "已登记"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_登记).Picture        '          "登记"
        Case "已报到"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_报到).Picture        '          "报到"
        Case "已检查"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_检查).Picture        '          "检查"
        Case "已报告"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_诊断).Picture        '          "诊断"
        Case "已审核"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_审核).Picture        '          "审核"
        Case "已完成"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_完成).Picture        '          "完成"
        Case "书写中"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_书写).Picture        '          "书写"
        Case "已驳回"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_驳回).Picture        '          "驳回"
        Case "已拒绝"
            imgStep.Picture = imgList.ListImages.Item(C_STEPIMG_拒绝).Picture        '          "拒绝"
        Case Else
            If App.LogMode = 0 Then
                HintMsg "未知的检查过程", "mobjPacsQueryWrap_OnSelChange", vbInformation
            End If
    End Select
    
    imgStep.Visible = True
    
    '判断是否需要重载模块,报告编辑器（Pacs编辑器，电子病历编辑器，智能文档编辑器），医嘱记录(住院医嘱，门诊医嘱)，费用记录（住院费用，门诊费用）
    strCurModuleTag = GetWorkModuleName(mstrSelTabName, mobjCurStudyInfo.lngExeDepartmentId, mobjCurStudyInfo.lngPatientFrom)
    If strCurModuleTag <> "" And strCurModuleTag <> mstrSelModuleTag Then
       Call SelectModule(mstrSelTabName, strCurModuleTag)
       TabWindow.Selected.tag = strCurModuleTag
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''刷新模块信息
    If blnRefreshModul Then Call RefreshModuleData(mstrSelTabName, mstrSelModuleTag, mobjSelModule)
    
    '刷新是否启用预约
    If CheckPopedom(mstrPrivs, "检查预约") Then
        Call IsSchedule(mlngCur科室ID, mobjCurStudyInfo.lngAdviceId, 0, mblnIsScheduleDept, mblnIsScheduleOrder)
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
'104686相关，呼叫后锁定检查，
'lngType类型  1:判断是否启用了参数并且是否已经有被锁定的检查,若有直接解锁        2:更新参数
'strLockedName   若="" 对流程没有影响，否则说明已经启用参数并且返回之前锁定的检查患者名称
'blnLockPara   用于更新PacsMain中的参数
            
    If lngType = 2 Then
    '更新参数
        mSysPar.blnLockAfterCall = blnLockPara
    End If
    
    Exit Sub
errhandle:
    If HintError(err, "mobjQueue_OnCallAboutLock", False) = 1 Then Resume
End Sub

Private Sub mobjQueue_OnCalled(ByVal lngAdviceId As Long, ByVal strRoom As String, ByVal TCallWay As zlQueueOper.TCallWay)
    Dim intRowIndex As Integer
On Error GoTo errhandle
 
    intRowIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("医嘱ID"))
    Call QueueDataConsistency(lngAdviceId, strRoom, intRowIndex)
    
    If TCallWay = cwBroadcast Or TCallWay = cwWaitRoom Then Exit Sub
        
    If mSysPar.blnLockAfterCall = False Then Exit Sub
    
    If mobjCapLinker Is Nothing Then Exit Sub
    If mobjWork_ImageCap Is Nothing Then Exit Sub
    
    '以下逻辑判断是否启用“同步定位到检查列表”，若未启用，需要根据业务ID获取需要锁定的检查，若已经启用，只需要简单锁定
    'intRowIndex=-1说明检查列表中没有显示排队列表中数据，需要另外获得数据
    If mSysPar.blnSynStudylist Then
        If intRowIndex > 0 And mobjCurStudyInfo.lngAdviceId <> lngAdviceId Then
            '同步定位
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
        '快捷叫号界面
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
    
    '重新配置工作模块页面
    If frmWorkModule.blnIsOk Then
        
        mblnInitOk = False '防止在子窗体加载过程中对子窗体进行刷新
        
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
'报告打印预览
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim blnIsDocEditor As Boolean
    Dim objReportV2 As frmReportV2
    Dim objRichV2 As frmRichReportV2
    
    ReoprtPrint = False
    blnIsDocEditor = False
    
    strSQL = "Select RAWTOHEX(检查报告ID) as 检查报告ID From 病人医嘱报告 Where 医嘱ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询检查报告ID", lngAdviceId)
    If rsData.RecordCount > 0 Then
        If NVL(rsData!检查报告ID) <> "" Then blnIsDocEditor = True
    End If
    
    '需要判断报告编辑器类型
    If blnIsDocEditor = False Then
        Set objReportV2 = New frmReportV2
        
        Call objReportV2.zlInit(Me, mlngModule, mlngCur科室ID, mstrPrivs, Nothing, Nothing, False)
        
        ReoprtPrint = objReportV2.PrintPreview(lngAdviceId, blnMoved, blnIsPrint, Val(strSpecifyReportId), strPrintFmts)
        
        Unload objReportV2
        Set objReportV2 = Nothing
    Else
        Set objRichV2 = New frmRichReportV2
        
        Call objRichV2.zlInit(Me, mlngModule, mlngCur科室ID, mstrPrivs)
        
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
    Dim str技师一 As String, str技师二 As String, str执行间 As String
    Dim intRowIndex As Integer
    Dim strSys1 As String
    Dim strSys2 As String
    Dim bytSize As Byte
    Dim strTmp As String
    Dim objReport As frmReportV2
    
    If mintQueryState <> 1 And objControl.ID <> conMenu_Manage_Query And objControl.ID <> conMenu_Manage_QueryCFG Then
        HintMsg "没有可用查询配置，请在查询方案管理中进行添加。", "cbrMain_Execute", vbInformation
        Exit Sub
    End If
    
    If mblnMenuDownState Then Exit Sub
    
    '这里需要根据id查找对应的菜单项目，因为通过绑定快捷键执行时，产生的是一个只有id而没有其他任何信息的control菜单项
    Set Control = cbrMain.FindControl(, objControl.ID, , True)
    If Control Is Nothing Then
        '如果该菜单为电子病历编辑器的右键菜单，则需要修改右键菜单的id等信息
        If Not mobjWork_Report Is Nothing Then
            Call mobjWork_Report.ReplacePopupMenu(objControl)
            
            Set Control = cbrMain.FindControl(, objControl.ID, , True)
        End If
        
        If Control Is Nothing Then Exit Sub
    End If
    
    If Control.ID = 0 Then Exit Sub
    
    If Not (Control.ID > conMenu_Manage_PacsPlugIn * 10000# And Control.ID < conMenu_Manage_PacsPlugIn * 10000# + 100) And objControl.ID <> conMenu_Manage_PacsPlugCfg Then
        '如果是执行插件菜单本身，则不需要广播事件
        Call Notify.Broadcast(BM_SYS__EVENT_MENU, 0, mobjCurStudyInfo.lngAdviceId, objControl.ID, objControl.Category)
    End If
    
    mblnMenuDownState = True
        
    cbrMain.RecalcLayout
    
    If Not mobjSelModule Is Nothing Then
        '先判断是否病理公共菜单操作
        If Not mobjWork_Pathol Is Nothing Then
            If mobjWork_Pathol.zlMenu.zlIsModuleMenu(mstrSelModuleTag, Control) Then
                Call mobjWork_Pathol.zlMenu.zlExecuteMenu(mstrSelModuleTag, Control.ID)
                
                mblnMenuDownState = False
                Exit Sub
            End If
        End If
        
        
        Select Case mstrSelTabName
            Case C_TAB_NAME_影像图像
                If mobjSelModule.zlMenu.zlIsModuleMenu(mstrSelModuleTag, Control) Then
                    Call mobjSelModule.zlMenu.zlExecuteMenu(mstrSelModuleTag, Control.ID)
                    
                    mblnMenuDownState = False
                    Exit Sub
                End If
                
            Case C_TAB_NAME_影像采集
            
            Case C_TAB_NAME_标本核收, C_TAB_NAME_病理取材, C_TAB_NAME_病理制片, C_TAB_NAME_病理特检, C_TAB_NAME_过程报告
                If mobjSelModule.zlMenu.zlIsModuleMenu(Control) Then
                    Call mobjSelModule.zlMenu.zlExecuteMenu(Control.ID)
                    
                    mblnMenuDownState = False
                    Exit Sub
                End If
                
            Case C_TAB_NAME_医嘱记录, C_TAB_NAME_病历记录, C_TAB_NAME_电子病历, C_TAB_NAME_费用记录
                If mobjWork_His.zlMenu.zlIsModuleMenu(mstrSelModuleTag, Control) Then
                    If mintChangeUserState = 2 Then  '交换了用户，则不允许操作
                        HintMsg "请统一用户后再操作", "cbrMain_Execute", vbInformation
                    Else
                        Call mobjWork_His.zlMenu.zlExecuteMenu(mstrSelModuleTag, Control.ID)
                    End If
                    
                    mblnMenuDownState = False
                    Exit Sub
                End If
            Case C_TAB_NAME_检查报告
                If mobjWork_Report.zlMenu.zlIsModuleMenu(mstrSelModuleTag, Control) Then
                    Call mobjWork_Report.zlMenu.zlExecuteMenu(mstrSelModuleTag, Control.ID)
                    
                    mblnMenuDownState = False
                    Exit Sub
                End If
        End Select
    End If
 
        
    Select Case Control.ID
        Case conMenu_Img_OpenView       '观片
            If mobjWork_PacsImg Is Nothing Or mstrSelTabName <> C_TAB_NAME_影像图像 Then
                If mobjPacsCore Is Nothing Then
                    mblnMenuDownState = False
                    HintMsg "图像查看对象无效，不能进行此操作。", "cbrMain_Execute", vbOKOnly
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
            
        Case conMenu_img_ContrastView   '对比观片
            If mobjWork_PacsImg Is Nothing Or mstrSelTabName <> C_TAB_NAME_影像图像 Then
                If mobjPacsCore Is Nothing Then
                    mblnMenuDownState = False
                    HintMsg "图像查看对象无效，不能进行此操作。", "cbrMain_Execute", vbOKOnly
                    Exit Sub
                End If
                
                Call OpenViewer(1, mobjPacsCore, mobjCurStudyInfo.lngAdviceId, True, Me, "", mobjCurStudyInfo.blnMoved)
            Else
                Call mobjWork_PacsImg.zlMenu.zlExecuteMenu("", conMenu_Img_Contrast + mobjWork_PacsImg.zlMenu.zlBaseMenuID)
            End If
            
        Case conMenu_Check_ViewLink
            Call ViewLinkChecks
        
        Case conMenu_PacsReport_Preview '报告预览
            If Not mobjWork_Report Is Nothing And mstrSelTabName = C_TAB_NAME_检查报告 Then
                
                strTmp = GetWorkModuleTag(C_TAB_NAME_检查报告)
                Call mobjWork_Report.zlMenu.zlExecuteMenu(strTmp, conMenu_File_Preview + mobjWork_Report.zlMenu.zlBaseMenuID)
            Else
                Call ReoprtPrint(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.blnMoved, False)
            End If
            
        Case conMenu_PacsReport_Print   '报告打印
            If Not mobjWork_Report Is Nothing And mstrSelTabName = C_TAB_NAME_检查报告 Then
                
                strTmp = GetWorkModuleTag(C_TAB_NAME_检查报告)
                Call mobjWork_Report.zlMenu.zlExecuteMenu(strTmp, conMenu_File_Print + mobjWork_Report.zlMenu.zlBaseMenuID)
            Else
                Call ReoprtPrint(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.blnMoved, True)
            End If
            
            
        Case conMenu_PacsReport_Write
'            If mobjWork_Report Is Nothing Then
'                '创建报告模块，并切换
'                For i = 0 To TabWindow.ItemCount - 1
'                    If TabWindow(i).Caption = C_TAB_NAME_检查报告 Then
'                        TabWindow(i).Selected = True
'                        Exit For
'                    End If
'                Next
'            End If
            
            If mstrSelModuleTag <> C_TAB_NAME_检查报告 Then
                strTmp = GetWorkModuleName(C_TAB_NAME_检查报告, mobjCurStudyInfo.lngExeDepartmentId, mobjCurStudyInfo.lngPatientFrom)
            Else
                strTmp = GetWorkModuleTag(C_TAB_NAME_检查报告)
            End If
            
            '调用报告编辑器封装的菜单打开书写窗口
            If Not mobjWork_Report Is Nothing And mstrSelTabName = C_TAB_NAME_检查报告 Then
            
                Select Case strTmp
                    Case C_WORKMODULE_NAME_老版报告
                        Call mobjWork_Report.zlMenu.zlExecuteMenu(C_WORKMODULE_NAME_老版报告, conMenu_PacsReport_Open + mobjWork_Report.zlMenu.zlBaseMenuID)
                     
                    Case C_WORKMODULE_NAME_病历报告
                        '电子病历编辑器
                        If Control.Caption = "查阅" Then    '查阅
                            Call mobjWork_Report.zlMenu.zlExecuteMenu(C_WORKMODULE_NAME_病历报告, conMenu_File_Open + mobjWork_Report.zlMenu.zlBaseMenuID)
                        ElseIf Control.Caption = "修订" Then    '修订
                            Call mobjWork_Report.zlMenu.zlExecuteMenu(C_WORKMODULE_NAME_病历报告, conMenu_Edit_Audit + mobjWork_Report.zlMenu.zlBaseMenuID)
                        Else                            '书写
                            Call mobjWork_Report.zlMenu.zlExecuteMenu(C_WORKMODULE_NAME_病历报告, conMenu_Edit_Modify + mobjWork_Report.zlMenu.zlBaseMenuID)
                        End If
                        
                    Case Else
                        '智能文档编辑器,可显示弹出式智能文档编辑器
                        Call mobjWork_Report.zlMenu.zlExecuteMenu(C_WORKMODULE_NAME_智能报告, conMenu_File_Open + mobjWork_Report.zlMenu.zlBaseMenuID)
                End Select
            Else
                If mSysPar.blnReportWithImage And Len(mobjCurStudyInfo.strStudyUID) <= 0 Then
'                    If MsgBoxD(Me, "本次检查未找到相关图像，是否强制书写。", vbYesNo, "提示") = vbNo Then
'                        mblnMenuDownState = False
'                        Exit Sub
'                    End If

                    Call MsgBoxD(Me, "本次检查未找到相关图像，不能书写。", vbOKOnly, "提示")
                    
                    mblnMenuDownState = False
                    Exit Sub
 
                End If
                
                '在没有进入检查报告模块页时，执行的打开操作
                Select Case strTmp
                    Case C_WORKMODULE_NAME_老版报告 '老版报告编辑器
                        '判断是否存在已经打开的报告编辑框
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

                        '弹出式方式进行报告编辑时，需要将焦点定位到编辑框
                        objReport.LocateEditBox

                    Case C_WORKMODULE_NAME_病历报告 '电子病历编辑器
                        If mobjRichReportWrap Is Nothing Then
                            Set mobjRichReportWrap = New frmEPREditWrapV2
                        End If

                        If mobjRichReportWrap.InitEprEditor(Nothing, Me, mlngModule, GetCurDeptId) = False Then
                            mblnMenuDownState = False
                            Exit Sub
                        End If

                        If Control.Caption = "查阅" Then    '查阅
                            Call mobjRichReportWrap.ExecuteMenu(mobjCurStudyInfo, conMenu_File_Open)
                        ElseIf Control.Caption = "修订" Then    '修订
                            Call mobjRichReportWrap.ExecuteMenu(mobjCurStudyInfo, conMenu_Edit_Audit)
                        Else '书写
                            Call mobjRichReportWrap.ExecuteMenu(mobjCurStudyInfo, conMenu_Edit_Modify)
                        End If


                        Call IEventNotify_Broadcast(BM_REPORT_EVENT_OPEN, "1", mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo)

                    Case Else   '智能报告编辑器
                        Call PreviewRichReport(Me, mlngModule, GetCurDeptId, mstrPrivs, mobjCurStudyInfo)
                End Select
            End If

'--------------------------文件------------------
        Case conMenu_File_PrintSet '打印设置
            
            Call zlPrintSet
            
        Case conMenu_File_Excel '清单打印
            Call Menu_File_Excel_click
            
        Case conMenu_File_Parameter '参数设置
            Call Menu_File_Parmeter_click
            
        Case ConMenu_File_ShortcutSet '快捷键设置
            Call Menu_File_ShortcutSet_click
            
        Case conMenu_Pathol_WorkModule  '站点模式设置
            Call conMenu_WorkModule_Click
            
'        Case conMenu_Manage_SetXWParam  '设置新网PACS的参数
'            Call Menu_Manage_SetXWParam_click
            
        Case conMenu_File_SendImg '发送图像
            Call conMenu_File_SendImg_click
            
        Case conMenu_Cap_DevSet         '视频设置
            If Not mobjWork_ImageCap Is Nothing Then
                Call mobjWork_ImageCap.zlShowVideoConfig
                mstrCaptureHot = GetSetting("ZLSOFT", "公共模块", "采集热键", "F8")
                mstrCaptureAfterHot = GetSetting("ZLSOFT", "公共模块", "后台采集热键", "F7")
                mstrCaptureAfterTagHot = GetSetting("ZLSOFT", "公共模块", "标记更新热键", "F6")
            End If
            
        Case conMenu_Manage_ChangeUser
            '交换用户时，需要先判断报告是否需要保存
            strTmp = GetWorkModuleTag(C_TAB_NAME_检查报告)
            
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(mobjCurStudyInfo, strTmp, True)
            End If
        
            Call ChangeUser
            
            '交换用户后，需要刷新报告编辑器，因为用户交换后，原有报告的编辑用户或者创建用户需要进行更新
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(mobjCurStudyInfo, strTmp, True)
            End If
            
        Case conMenu_Manage_SwitchUser
            '切换用户时，需要先判断报告是否需要保存
            strTmp = GetWorkModuleTag(C_TAB_NAME_检查报告)
            
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(mobjCurStudyInfo, strTmp, True)
            End If
            
            Call SwitchUser
            
            '切换用户后，需要刷新报告编辑器，因为用户切换后，原有报告的编辑用户或者创建用户需要进行更新
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(mobjCurStudyInfo, strTmp, True)
            End If
            
        Case conMenu_Manage_Change_In   '隐藏列表
            If dkpMain.Panes(1).hidden = False Then
                dkpMain.Panes(1).Hide
            Else
                dkpMain.ShowPane (1)
            End If
            
        Case conMenu_File_Exit '退出
            mblnMenuDownState = False
            Unload Me
            
'---------------------------检查-----------------
        Case conMenu_Manage_RequestPrint * 10# + 1 To conMenu_Manage_RequestPrint * 10# + 9 '打印诊疗单据
            Call FuncBillPrint(Control)
            
        Case comMenu_Petition_Capture                       '扫描申请单
            Call Menu_Petition_扫描申请单(1)
            
        Case comMenu_Petition_View
            Call Menu_Petition_扫描申请单(0)                '查看申请单
            
        Case conMenu_Manage_Regist                          '登记
            Call Menu_Manage_登记
            
        Case conMenu_Manage_CopyCheck                       '复制登记
            Call Menu_Manage_复制登记
            
        Case conMenu_Manage_Receive                         '报到
            Call Menu_Manage_报到
            
        Case conMenu_Manage_Redo                            '取消登记
            Call Menu_Manage_取消登记
            
        Case conMenu_Manage_ReGet                           '召回取消
            Call Menu_Manage_召回取消
            
        Case conMenu_Manage_ThingModi                       '修改登记
            Call Menu_Manage_修改
        
        Case conMenu_Manage_CheckList                       '查看电子申请单
            Call Menu_Manage_CheckList
            
        Case conMenu_Manage_ExecOnePart                     '分部位执行
            Call menu_Manage_ExecOnePart
            
        Case conMenu_Manage_DiseaseQuery                    '传染病查询
            Call Menu_Manage_DiseaseQuery
            
        Case conMenu_Manage_DiseaseRegist                   '传染病登记
            Call Menu_Manage_DiseaseRegist
        
        Case conMenu_Manage_ModifBaseInfo               '基本信息调整
            Call Menu_Manage_ModifBaseInfo
        
        Case conMenu_Manage_Logout                          '取消报到
            Call Menu_Manage_取消报到
            
        Case conMenu_Cap_StudySyncState '采集锁定
            Call LockCapture(mobjCurStudyInfo)
            
'            If Not mobjWork_ImageCap Is Nothing Then
'                If mobjCurStudyInfo.blnMoved Or mobjCapLinker Is Nothing Then
'                    HintMsg "当前检查状态不允许锁定。", "cbrMain_Execute", vbOKOnly
'                    mblnMenuDownState = False
'                    Exit Sub
'                End If
'
'                mobjCapLinker.ReportAdviceId = 0    '需要清空报告id，避免锁定时，优先使用报告医嘱id进行锁定
'                mobjCapLinker.LockAdviceId = mobjCurStudyInfo.lngAdviceId
'
'                Call mobjWork_ImageCap.ResetLockState(True)
'
'            End If
            
        Case conMenu_Manage_InQueue                         '排队叫号入队
            Call zlInPacsQueue
            
        Case conMenu_Manage_Schedule                        '检查预约
            Call Menu_Manage_检查预约
            
        Case conMenu_Manage_ScheduleManage                  '预约管理
            Call Menu_Manage_预约管理
            
        Case conMenu_Manage_Transfer                        '关联影像
            Call Menu_Manage_关联影像
            
        Case conMenu_Manage_Cancel                          '取消关联
            Call Menu_Manage_取消关联
            
        Case conMenu_Manage_AttachMoney                     '补付费
            Call Menu_Manage_补附费
            
        Case conMenu_Manage_CompleteAttach                  '病理完成补费
            Call Menu_Manage_完成病理补费
            
        Case conMenu_Manage_Review                          '随访
            Call Menu_Manage_随访
            
        Case conMenu_Tool_Analyse
            If mobjPacsCore Is Nothing Then
                mblnMenuDownState = False
                HintMsg "图像查看对象无效，不能进行此操作。", "cbrMain_Execute", vbOKOnly
                Exit Sub
            End If
            
            Call OpenViewer(1, mobjPacsCore, mobjCurStudyInfo.lngAdviceId, False, Me, "", mobjCurStudyInfo.blnMoved)
        
        Case conMenu_Manage_ReportRelease                   '报告发放
            Call Menu_Manage_报告发放
            
        Case conMenu_Manage_FilmRelease                     '胶片发放
            Call Menu_Manage_胶片发放
            
            
        Case conMenu_Manage_SendArrange                     '发送安排
            Call frmSendArrange.ShowMe(Me, mlngCur科室ID, mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, str技师一, str技师二, str执行间)
            If str技师一 <> "" Then
                Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
            End If

        Case conMenu_Manage_ReportExecutor                  '报告执行，即标记报告人
            Call Menu_Manage_ReportExecutor
            
        Case conMenu_Manage_SendAudit * 10# + 1 To conMenu_Manage_SendAudit * 10# + 99    '发送审核
            Call Menu_Manage_SendAudit(0, Control.Caption)
            
        Case conMenu_Manage_PacsCriticalReg, conMenu_Manage_PacsCriticalManage        '危机值处理
            Call Menu_Manage_CriticalMark(Control.ID)
            
        Case conMenu_Manage_Negative, conMenu_Manage_Positive                  '结果阴阳性
            Call Menu_Manage_标记阴阳(Control.ID)
           
        Case conMenu_Manage_FuHe, conMenu_Manage_JiBenFuHe, conMenu_Manage_BuFuHe   '符合情况
            Call Menu_Manage_符合情况(Control.ID)
            
        Case conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel
            Call Menu_Manage_绿色通道(Control.ID)
        
        Case conMenu_Manage_Complete                        '检查完成
            Call Menu_Manage_检查最终完成
                
        Case conMenu_Manage_Undone                          '取消检查完成
            Call Menu_Manage_取消检查完成
            
        Case conMenu_Manage_RelatingPatiet                  '关联病人
            Call Menu_Manage_关联病人
            
        Case conMenu_Manage_Burn                            '图像刻录
            Call Menu_Manage_图像刻录

        Case conMenu_Manage_LookMecRecord                   '病案查阅
            Call Menu_Manage_病案查阅
            
'----------------------------------------收藏---------------------------------------
        Case conMenu_Collection_Manage  '收藏管理
           Call Menu_Manage_收藏管理
        Case conMenu_Collection_To      '收藏到
           Call Menu_Manage_收藏到
        Case comMenu_Collection_Type * 10000# To comMenu_Collection_Type * 10000# + 9999  '动态收藏类别菜单
           Call Menu_Manage_收藏数据显示(Control, 0)
        Case conMenu_Collection_ViewShare * 10000# To conMenu_Collection_ViewShare * 10000# + 9999   '查看共享
           Call Menu_Manage_收藏数据显示(Control, 1)
        Case conMenu_Manage_QueryCFG
            
            strSys1 = "[系统.系统号],[系统.模块号],[系统.科室ID],[系统.用户ID],[系统.用户账号],[系统.用户名称]"
            strSys1 = strSys1 & ",[系统.服务器日期],[系统.服务器时间],[系统.本地日期],[系统.本地时间]"
            strSys1 = strSys1 & ",[系统.开始日期],[系统.结束日期]"


            strSys2 = "[系统.病人ID],[系统.医嘱ID]"

            If gbytFontSize = 9 Then
                bytSize = 0
            ElseIf gbytFontSize = 12 Then
                bytSize = 1
            Else
                bytSize = 2
            End If

            Call mobjPacsQueryWrap.CurPacsQuery.ShowSchemeCfg(mlngModule, strSys1, strSys2, bytSize, Me)
            
        Case conMenu_Manage_QueryCfgUserScheme
            Call mobjPacsQueryWrap.CurPacsQuery.ShowUserScheme(mlngModule, mlngCur科室ID, Me)
        Case conMenu_Manage_QueryTabDisplayScheme
            '更新数据库参数和缓存参数,重新设置选中tab项目
            mSysPar.blnQuickTabDisplayScheme = Not mSysPar.blnQuickTabDisplayScheme
            
            zlDatabase.SetPara "显示常用方案标签", IIf(mSysPar.blnQuickTabDisplayScheme, "1", "0"), glngSys, mlngModule
            
            tabScheme.Visible = mSysPar.blnQuickTabDisplayScheme
            tabScheme.tag = IIf(mSysPar.blnQuickTabDisplayScheme, "1", "0")
            
            
            Call AdjustFace(picList.Height, picList.Width)
'----------------------------------------第三方插件功能---------------------
        Case conMenu_Manage_PacsPlugCfg
            Call ShowPacsInterfaceCfg
        Case conMenu_Manage_PacsPlugIn * 10000# To conMenu_Manage_PacsPlugIn * 10000# + 100
            Call ExecutePluginInterfaceFun(Control.Caption, Control.Parameter, Control.DescriptionText, False)
'-------------------------------------------------------------------
        Case conMenu_View_Filter '过滤
            Call mobjPacsQueryWrap.ExecuteQuery(C_QUERY_更多过滤)
'---------------------------查看----------------
        Case conMenu_View_ToolBar_Button '工具栏
            Call Menu_View_ToolBar_Button_click(Control)
            
        Case conMenu_View_FontSize_S    '小字体
            Call SetFontSize(0)
        Case conMenu_View_FontSize_M    '中字体
            Call SetFontSize(1)
        Case conMenu_View_FontSize_L    '大字体
            Call SetFontSize(2)
            
        Case conMenu_View_ToolBar_Text '按钮文字
            Call Menu_View_ToolBar_Text_click(Control)
        Case conMenu_View_ToolBar_Size '大图标
            Call Menu_View_ToolBar_Size_click(Control)
            
        Case conMenu_View_StatusBar '状态栏
            Call Menu_View_StatusBar_click(Control)
            
        Case conMenu_View_Refresh '刷新
            mblnIsForceRefresh = True
            
            Call RefreshList
            Call RefreshPacsQueueData(True)
            
            mblnIsForceRefresh = False
        Case comMenu_Cap_Process
            Call Menu_Manage_浮动采集
'---------------------------工具----------------
        Case conMenu_Tool_Valid         '图像校对工具
            
            If Len(Dir(GetAppRootPath & "zlPacsImageValid.exe")) > 0 Then
                If InitRegister Then
                    Shell GetAppRootPath & "zlPacsImageValid.exe   " & gstrServerName & "||" & gstrUserName & "||" & gstrUserPswd & "||" & mlngCur科室ID & "||" & mSysPar.lngImageValid & "||" & "" & "||1", 1
                End If
            End If
'--------------------------帮助-----------------
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
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse科室, "|")) + 1
            Call Menu_Dept_Select(Control)
        Case conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99
            If Control.Parameter <> "" Then '执行发布到当前模块的报表
        
                If mobjCurStudyInfo.lngAdviceId <> 0 Then
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                        "执行科室=" & mobjCurStudyInfo.lngExeDepartmentId, "医嘱ID=" & mobjCurStudyInfo.lngAdviceId, "发送号=" & mobjCurStudyInfo.lngSendNo, _
                            "NO=" & mobjCurStudyInfo.strNO, "病人ID=" & mobjCurStudyInfo.lngPatId, "挂号单=" & mobjCurStudyInfo.strRegNo)
                Else
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "", 1)
                End If
                
            End If
        '----------------------------------------自定义查询---------------------------------------
        Case conMenu_Manage_CustomQuery * 100# + 1 To conMenu_Manage_CustomQuery * 100# + 99
            Call ChangeScheme(Control.Caption, Val(Control.Parameter), True)
            
        Case Else
            If mobjCurStudyInfo.lngAdviceId = 0 Then
                mblnMenuDownState = False
                Exit Sub
            End If
            
            Select Case mstrSelTabName
                Case C_TAB_NAME_排队叫号
                    If Not mobjQueue Is Nothing Then
                        If mintChangeUserState = 2 Then  '交换了用户，则不允许操作
                            HintMsg "请统一用户后再操作", "cbrMain_Execute", vbInformation
                        Else
                            mobjQueue.zlExecuteCommandbar Control
                        End If
                    End If
                Case C_TAB_NAME_费用记录, C_TAB_NAME_电子病历, C_TAB_NAME_医嘱记录, C_TAB_NAME_病历记录
                    If Not mobjWork_His Is Nothing Then
                        Call mobjWork_His.zlMenu.zlExecuteMenu(mstrSelModuleTag, Control.ID)
                    End If
                Case C_TAB_NAME_检查报告
                    If Not mobjWork_Report Is Nothing Then
                        Call mobjWork_Report.zlMenu.zlExecuteMenu(mstrSelModuleTag, Control.ID)
                    End If
            End Select
            
    End Select
    
    If Not (Control.ID > conMenu_Manage_PacsPlugIn * 10000# And Control.ID < conMenu_Manage_PacsPlugIn * 10000# + 100) And objControl.ID <> conMenu_Manage_PacsPlugCfg Then
        '如果是执行插件菜单本身，则不需要广播事件
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
            HintMsg "当前检查状态不允许锁定。", "cbrMain_Execute", vbOKOnly
            Exit Sub
        End If
        
        lngOldReportAdviceId = mobjCapLinker.ReportAdviceId
        
        mobjCapLinker.ReportAdviceId = 0    '需要清空报告id，避免锁定时，优先使用报告医嘱id进行锁定
        mobjCapLinker.LockAdviceId = objStudyInfo.lngAdviceId
        
        Call mobjWork_ImageCap.ResetLockState(True)
        
        mobjCapLinker.ReportAdviceId = lngOldReportAdviceId
    End If
End Sub

Private Function LocateReportWindow(ByVal lngAdviceId As Long) As Boolean
'定位弹出式报告窗口
    Dim objForm As Object
    
    LocateReportWindow = False
    
    '判断是否存在已经打开的报告编辑框
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
         
    If Not CheckPopedom(mstrPrivs, "插件配置管理") Then
        HintMsg "您没有该操作的权限，请联系管理员。", "ShowPacsInterfaceCfg", vbInformation
        Exit Sub
    End If
    
    If Not ChechHaveTlbinf32 Then
        HintMsg "系统中缺少TLBINF32.DLL文件，导致插件配置功能不能正常使用，请联系软件技术人员解决(解决方法：在系统目录下添加并注册TLBINF32.DLL文件)。", "ShowPacsInterfaceCfg", vbInformation
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
'blnAutoDo 是否自动执行（影响错误处理提示信息处理方式）
'调用vbs脚本实现功能
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
        '对于预定义参数，内部赋值
        strTmpVBS = ary(i)
        
        Do While InStr(strTmpVBS, "[[") > 0
            lngStart = InStr(strTmpVBS, "[[")
            lngEnd = InStr(strTmpVBS, "]]") + 2
            
            strParaName = Mid(strTmpVBS, lngStart, lngEnd - lngStart)
            
            Select Case strParaName
                Case "[[触发标记]]"
                    strParaVal = lngTimeTag
                    
                Case "[[附加参数1]]"
                    strParaVal = strAttachPar1
                    
                Case "[[附加参数2]]"
                    strParaVal = strAttachPar2
                    
                Case "[[附加参数3]]"
                    strParaVal = strAttachPar3
                    
                Case "[[用户名]]"
                    strParaVal = UserInfo.姓名
                                
                Case "[[账号名]]"
                    strParaVal = UserInfo.用户名
                    
                Case "[[系统号]]"
                    strParaVal = glngSys
                    
                Case "[[模块号]]"
                    strParaVal = mlngModule
                
                Case "[[科室ID]]"
                    strParaVal = mlngCur科室ID
                
                Case "[[病人ID]]"
                    strParaVal = mobjCurStudyInfo.lngPatId
                    
                Case "[[医嘱ID]]"
                    strParaVal = mobjCurStudyInfo.lngAdviceId
                    
                Case "[[发送号]]"
                    strParaVal = mobjCurStudyInfo.lngSendNo
                    
                Case "[[检查号]]"
                    strParaVal = mobjCurStudyInfo.strStudyNum
                    
                Case "[[门诊号]]", "[[住院号]]"
                    strParaVal = mobjCurStudyInfo.strMarkNum
                    
                Case "[[身份证号]]"
                    strParaVal = mobjCurStudyInfo.strIIDNumber
                    
                Case "[[影像类别]]"
                    strParaVal = mobjCurStudyInfo.strImgType
                                        
                Case "[[当前窗口句柄]]"
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
        err.Raise -1, , "插件 [" & strFuncDes & "] 产生错误，" & strResult
    End If
    
'    Exit Function
'ErrorHnad:
'    ExecutePluginInterfaceFun = False
'    err.Raise -1, , "插件执行产生错误，" & err.Description
End Function

Private Function ExecuteSub(ByVal strVBS As String, Optional ByVal blnCheckVBS As Boolean = False) As String
'调用vbs脚本实现功能
    Dim objCall As Object
    Dim strTempVBS As String
    
On Error GoTo errhandle
    
    ExecuteSub = ""
    
    '创建脚本执行对象
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
'刷新排队模块数据
    If mSysPar.blnUseQueue And Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Call mobjQueue.zlRefreshQueueData(GetSelQueueRooms(), blnSetFocus)
    End If
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
    
    '设置字体大小
    gbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, IIf(bytSize = 2, 15, bytSize)))
    
    Call ReSetFormFontSize
    
    If mobjSelModule Is Nothing Then Exit Sub
    
    Call ReSetModuleFontSize(mstrSelTabName, mstrSelModuleTag, mobjSelModule, gbytFontSize)
End Sub


Private Sub ReSetModuleFontSize(ByVal strSelTabName As String, ByVal strSelModuleTag As String, _
    ByVal objSelModule As Object, ByVal bytFontSize As Byte)
'功能:重新设置各个业务模块窗体的字体大小
'bytSizeType 0-默认小字体，1-大字体,
    Dim bytSizeType As Long
On Error GoTo errhandle
        
    If objSelModule Is Nothing Then Exit Sub

    Select Case strSelTabName
        Case C_TAB_NAME_影像图像
            Call objSelModule.ReSetFormFontSize(gbytFontSize)
            
        Case C_TAB_NAME_影像采集
            Call objSelModule.ReSetFormFontSize(gbytFontSize)
            
        Case C_TAB_NAME_检查报告
            If strSelModuleTag = C_WORKMODULE_NAME_老版报告 Then
                Call objSelModule.ReSetFormFontSize(gbytFontSize)
            End If
            
        Case C_TAB_NAME_标本核收, C_TAB_NAME_病理取材, C_TAB_NAME_病理制片, C_TAB_NAME_病理特检, C_TAB_NAME_过程报告
            Call objSelModule.ReSetFormFontSize(gbytFontSize)
            
        Case C_TAB_NAME_费用记录, C_TAB_NAME_病历记录, C_TAB_NAME_电子病历, C_TAB_NAME_医嘱记录
            bytSizeType = IIf(bytFontSize = 9, 0, 1)
            If mlngModule = G_LNG_PATHOLSYS_NUM Then
                '病理系统使用多模块对象存在差异，因此需要识别调用方法
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
'功能:重新设置工作站窗体的字体大小
    On Error Resume Next
    
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim strFontType As String
    Dim i As Integer
    
    Me.FontSize = gbytFontSize
    Set CtlFont = New StdFont
    strFontType = IIf(IsUseClearType = True, "微软雅黑", "宋体")
    CtlFont.Name = strFontType
    
    Call ucPacsHelper1.SetFontSize(gbytFontSize)
    
    If gblUsePacsQuery Then
        Call mobjPacsQueryWrap.CurPacsQuery.RefreshCfgFontSize(gbytFontSize)
    End If
    
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("TabStrip") '页面控件
            objCtrl.Font.Name = strFontType
            objCtrl.Font.Size = gbytFontSize
        Case UCase("Label")
            If objCtrl.Name = "LabFlag费用" Or objCtrl.Name = "LabFlag婴儿" Or objCtrl.Name = "LabFlag绿色通道" _
                Or objCtrl.Name = "LabFlag危机状态" Or objCtrl.Name = "LabFlag传染病状态" _
                Or objCtrl.Name = "labNoScheme" Or objCtrl.Name = "LabFlag急诊" Then
            ElseIf objCtrl.Name = "labCollectionInfo" Then
                objCtrl.Font.Name = strFontType
                objCtrl.Font.Bold = False
                objCtrl.FontSize = gbytFontSize
            Else
                objCtrl.Font.Name = strFontType
                objCtrl.FontSize = gbytFontSize
                objCtrl.Height = TextHeight("啊") + 60
            End If
        Case UCase("vsFlexGrid")
        
            Dim lngRow As Long
            
            objCtrl.Cell(flexcpFontSize, 0, 0, objCtrl.Rows - 1, objCtrl.Cols - 1) = gbytFontSize
            objCtrl.HeadFont.Size = gbytFontSize
            objCtrl.FontSize = gbytFontSize
            objCtrl.RowHeight(0) = TextHeight("啊") + 150
            '根据最大行号修改第一列的宽度
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
                objCtrl.RowHeight(i) = TextHeight("啊") + 120
            Next
            
        Case UCase("ComboBox")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = gbytFontSize
        Case UCase("OptionButton")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = gbytFontSize
            objCtrl.Width = TextWidth("啊啊" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = gbytFontSize
            objCtrl.Width = TextWidth("啊啊" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Name = strFontType
            objCtrl.Font.Size = gbytFontSize
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("啊") * 1.5
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
    
    '判断医嘱对象是否有效
    If mobjCurStudyInfo Is Nothing Then
        blnFromAdvice = False
    Else
        blnFromAdvice = IIf(mobjCurStudyInfo.lngAdviceId <> 0, True, False)
    End If
    
    '是否从医嘱读取科室ID
    If blnFromAdvice Then
        GetCurDeptId = mobjCurStudyInfo.lngExeDepartmentId
    Else
        If mblnAllDepts Then
            GetCurDeptId = UserInfo.部门ID
        Else
            GetCurDeptId = mlngCur科室ID
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
                        '只有医技需要添加“全部科室”的科室选择菜单
                        Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100#, "全部科室")
                    
                        objControl.Category = "Main"
                        objControl.DescriptionText = 0
                        If mblnAllDepts = True Then objControl.Checked = True
                    End If
                    
                    '再添加每一个具体科室
                    For i = 0 To UBound(Split(mstrCanUse科室, "|"))  'mstrCanUse科室=id_编码-名称|id_编码-名称
                        Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100# + i + 1, Split(Split(mstrCanUse科室, "|")(i), "_")(1) & "(&" & i & ")")
                        objControl.Category = "Main"
                        objControl.DescriptionText = Split(Split(mstrCanUse科室, "|")(i), "_")(0)
                        
                        If mblnAllDepts = False And mlngCur科室ID = objControl.DescriptionText Then
                            objControl.Checked = True
                        End If
                    Next
                End If
            End With
        Case Else
            Select Case mstrSelTabName
                Case C_TAB_NAME_医嘱记录, C_TAB_NAME_费用记录
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
'0-不允许书写，1-允许书写，2-不允许修订，3-允许修订，4-查阅(暂定)
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strCurModuleTag As String
    Dim lngDeptId As Long
    
    GetReportEditState = 0
    
    If objStudyInfo.lngAdviceId <= 0 Then Exit Function '医嘱ID无效时，不允许编辑报告
    
    If mblnAllDepts Then
        lngDeptId = UserInfo.部门ID
    Else
        lngDeptId = mlngCur科室ID
    End If
    
    strSQL = "select 科室ID, 创建人,保存人,归档人,完成时间,RawToHex(检查报告ID) as 检查报告ID  from 病人医嘱报告 a , 电子病历记录 b where a.病历ID=b.Id(+) and a.医嘱ID=[1]"
    If objStudyInfo.blnMoved Then
        strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
        strSQL = Replace(strSQL, "电子病历记录", "H电子病历记录")
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询报告情况", objStudyInfo.lngAdviceId)
    
    strCurModuleTag = GetWorkModuleName(C_TAB_NAME_检查报告, objStudyInfo.lngExeDepartmentId, objStudyInfo.lngPatientFrom)
    If strCurModuleTag = C_WORKMODULE_NAME_智能报告 Then
         If rsData.RecordCount > 0 Then
            GetReportEditState = IIf(NVL(rsData!检查报告ID) <> "", 4, 0)
         Else
            GetReportEditState = 0
         End If
         
        Exit Function
    End If
    
    If rsData.RecordCount > 0 Then
        '已书写报告
        '如果检查已经执行完成，则只能进行查看
        If objStudyInfo.intStep >= 6 Then
            GetReportEditState = 4
            Exit Function
        End If
        
        If NVL(rsData!完成时间) <> "" Then
            '已签名
            '判断是否有修订权限
            If InStr(1, mstrPrivs, "报告修订") > 0 Then
                If NVL(rsData!科室ID) = lngDeptId Or IsContainDept(UserInfo.ID, Val(NVL(rsData!科室ID))) Then
                    If NVL(rsData!归档人) = "" Then
                        'TASK:这里可以增加对用户级别的判断
                        GetReportEditState = 3
                    Else
                        '已归档
                        GetReportEditState = 2
                    End If
                Else
                    GetReportEditState = 2
                End If
            Else
                '判断保存人和当前用户是否相同，如果相同，说明是最后一次保存是由自己签名的
                If NVL(rsData!保存人) = UserInfo.用户名 Then
                    GetReportEditState = 3
                Else
                    GetReportEditState = 2
                End If
            End If
        Else
            '未签名
            '判断科室id与当前科室id是否相同
            If Val(NVL(rsData!科室ID)) <> lngDeptId And IsContainDept(UserInfo.ID, Val(NVL(rsData!科室ID))) = False Then Exit Function   '不允许编辑
            If NVL(rsData!保存人) = UserInfo.姓名 Or InStr(1, mstrPrivs, "他人报告") > 0 Then
                GetReportEditState = 1
            End If
        End If
    Else
        '未书写报告
        '判断报告是否被他人锁定编辑
        If objStudyInfo.strReportOperation <> "" Then
            If objStudyInfo.strReportOperation <> UserInfo.姓名 Then 'And InStr(1, mstrPrivs, "他人报告") <= 0 Then
                '无他人报告权限，且被他人报告锁定编辑时，不允许书写报告
                Exit Function
            End If
        End If
        
        If InStr(1, mstrPrivs, "报告书写") > 0 _
            And ((objStudyInfo.intStep > 1 And objStudyInfo.intStep < 6) _
                    Or (objStudyInfo.intStep = 6 And CheckPopedom(mstrPrivs, "补录报告"))) Then
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
      

    '如果该菜单为电子病历编辑器的右键菜单，则需要修改菜单id等信息
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
        intState = mobjCurStudyInfo.intStep   '执行过程
        blnCancel = mobjCurStudyInfo.strStuStateDesc = "已拒绝"
    End If
    
    If Not mobjSelModule Is Nothing Then
        
        If Not mobjWork_Pathol Is Nothing Then
            If mobjWork_Pathol.zlMenu.zlIsModuleMenu("", Control) Then
                Call mobjWork_Pathol.zlMenu.zlUpdateMenu("", Control)
                Exit Sub
            End If
        End If
        
        Select Case mstrSelTabName
            Case C_TAB_NAME_影像图像
                If mobjSelModule.zlMenu.zlIsModuleMenu(mstrSelModuleTag, Control) Then
                    Call mobjSelModule.zlMenu.zlUpdateMenu(mstrSelModuleTag, Control)
                    Exit Sub
                End If
            Case C_TAB_NAME_影像采集
            
            Case C_TAB_NAME_标本核收, C_TAB_NAME_病理取材, C_TAB_NAME_病理制片, C_TAB_NAME_病理特检, C_TAB_NAME_过程报告
                '病理相关模块继承的是IWorkMenu接口，不包含模块名称
                If mobjSelModule.zlMenu.zlIsModuleMenu(Control) Then
                    Call mobjSelModule.zlMenu.zlUpdateMenu(Control)
                    Exit Sub
                End If
                
            Case C_TAB_NAME_医嘱记录, C_TAB_NAME_病历记录, C_TAB_NAME_费用记录, C_TAB_NAME_电子病历
                If mobjWork_His.zlMenu.zlIsModuleMenu(mstrSelModuleTag, Control) Then
                    Call mobjWork_His.zlMenu.zlUpdateMenu(mstrSelModuleTag, Control)
                    

                    '已完成除查阅,以及医嘱中报告查看打印，观片菜单外均不可用
                    If mobjCurStudyInfo.intStep = 6 Then
                        Select Case Control.ID
                            Case conMenu_Edit_MarkMap, conMenu_Tool_PlugIn, conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99, conMenu_Edit_Compend, conMenu_Manage_ReportLisView, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3
                                Control.Enabled = True
                            Case conMenu_Edit_Copy, conMenu_File_ExportToXML, conMenu_Tool_Search, conMenu_File_Open, conMenu_EditPopup, conMenu_Edit_ChargeDelAudit
                                '这几个菜单不控制
                            Case Else
                                Control.Enabled = False
                        End Select
                    End If
                    
                    Exit Sub
                End If
            Case C_TAB_NAME_检查报告
                If mobjWork_Report.zlMenu.zlIsModuleMenu(mstrSelModuleTag, Control) Then
                    Call mobjWork_Report.zlMenu.zlUpdateMenu(mstrSelModuleTag, Control)
                    Exit Sub
                End If
        End Select
    End If
      
                    
    Select Case Control.ID
        Case conMenu_PacsReport_Preview '预览
            If Not mobjWork_Report Is Nothing Then
                If mobjCurStudyInfo.lngAdviceId = mobjWork_Report.AlreadyAdviceId Then
                    Control.ID = conMenu_File_Preview + mobjWork_Report.zlMenu.zlBaseMenuID
                    Call mobjWork_Report.zlMenu.zlUpdateMenu(GetWorkModuleTag(C_TAB_NAME_检查报告), Control)
                    
                    Control.ID = conMenu_PacsReport_Preview
                Else
                    '未切换到报告模块，报告模块数据没有刷新
                    Control.Enabled = mobjCurStudyInfo.blnCanPrint
                End If
            Else
                '
                Control.Enabled = mobjCurStudyInfo.blnCanPrint
            End If
        
        Case conMenu_PacsReport_Print   '打印
            If Not mobjWork_Report Is Nothing Then
                If mobjCurStudyInfo.lngAdviceId = mobjWork_Report.AlreadyAdviceId Then
                    Control.ID = conMenu_File_Print + mobjWork_Report.zlMenu.zlBaseMenuID
                    Call mobjWork_Report.zlMenu.zlUpdateMenu(GetWorkModuleTag(C_TAB_NAME_检查报告), Control)
                    
                    Control.ID = conMenu_PacsReport_Print
                Else
                    '未切换到报告模块，报告模块数据没有刷新
                    Control.Enabled = mobjCurStudyInfo.blnCanPrint
                End If
            Else
                '
                Control.Enabled = mobjCurStudyInfo.blnCanPrint
            End If
        
        Case conMenu_PacsReport_Write   '书写
            Select Case mobjCurStudyInfo.lngReportEditState
                Case 0
                    Control.Caption = "书写"
                    Control.Enabled = False
                Case 1
                    Control.Caption = "书写"
                    Control.Enabled = True
                Case 2
                    Control.Caption = "修订"
                    Control.Enabled = False
                Case 3
                    Control.Caption = "修订"
                    Control.Enabled = True
                Case 4
                    Control.Caption = "查阅"
                    Control.Enabled = True
            End Select
'            If Not mobjWork_Report Is Nothing Then
'                Select Case GetWorkModuleTag(C_TAB_NAME_检查报告)
'                    Case C_WORKMODULE_NAME_老版报告
'                        Control.ID = conMenu_PacsReport_Open + mobjWork_Report.zlMenu.zlBaseMenuID
'                        Call mobjWork_Report.zlMenu.zlUpdateMenu(C_WORKMODULE_NAME_老版报告, Control)
'
'                        Control.ID = conMenu_PacsReport_Write
'                        Control.Visible = True
'                    Case C_WORKMODULE_NAME_病历报告
'                        '电子病历编辑器
'                        Control.ID = conMenu_Edit_Modify + mobjWork_Report.zlMenu.zlBaseMenuID
'                        Call mobjWork_Report.zlMenu.zlUpdateMenu(C_WORKMODULE_NAME_病历报告, Control)
'
'                        Control.ID = conMenu_PacsReport_Write
'                        Control.Visible = True
'                    Case Else
'                        Control.Visible = False
'                End Select
'            Else
'                '
'                Control.Visible = GetWorkModuleName(C_TAB_NAME_检查报告, GetCurDeptId, GetCurPatientFrom) <> C_WORKMODULE_NAME_智能报告 '  mSysPar.lngReportType <> ReportType.报告文档编辑器
'                Control.Enabled = ((intState >= 2 And intState < 6) Or (intState >= 6 And CheckPopedom(mstrPrivs, "补录报告"))) And Not blnNoRecord
'            End If
        
        Case conMenu_Manage_LocateValue
            Control.Enabled = Not blnNoRecord
        Case comMenu_Cap_Process
            Control.Enabled = True 'Not blnNoRecord
        Case conMenu_View_Filter * 10#
            Control.Caption = " " & IIf(mblnAllDepts = True, "全部科室", Split(mstrCur科室, "-")(1)) & " "
            Control.Checked = True

        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse科室, "|")) + 1
            If mblnAllDepts = True Then
                Control.Checked = (Control.DescriptionText = 0)
            Else
                Control.Checked = (Control.DescriptionText = mlngCur科室ID)
            End If
        Case conMenu_View_ToolBar_Button '工具栏
            If cbrMain.Count >= 2 Then
                Control.Checked = Me.cbrMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text '图标文字
            If cbrMain.Count >= 2 Then
                Control.Checked = Not (Me.cbrMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '大图标
            Control.Checked = Me.cbrMain.options.LargeIcons
        Case conMenu_View_StatusBar '状态栏
            Control.Checked = Me.stbThis.Visible
        Case conMenu_View_Filter   '过滤
        
        Case conMenu_View_Refresh  '刷新
        
        Case conMenu_Manage_RequestPrint
            Control.Enabled = Control.CommandBar.Controls.Count > 0 And Not blnNoRecord
            
        Case conMenu_Manage_Regist   '检查登记(&I)
        Case conMenu_Manage_CopyCheck '复制登记
            If Not blnNoRecord Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
        Case conMenu_Manage_Redo   '取消登记(&R)
            If Not blnNoRecord Then
                Control.Enabled = intState <= 1 And intState <> -1 And Not blnCancel
            Else
                Control.Enabled = False
            End If
        Case conMenu_Manage_ReGet   '召回取消
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
        Case conMenu_Manage_ThingModi   '修改信息(&M)
            If Not blnNoRecord Then
                Control.Enabled = intState < 6 And Not blnCancel
            Else
                Control.Enabled = False
            End If
        Case conMenu_Manage_CheckList   '查看申请单
            Control.Visible = True
            If mobjCurStudyInfo.lngAdviceId > 0 And mobjCurStudyInfo.lngPatientFrom <> 3 Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
            
        Case conMenu_Manage_ExecOnePart     '分部位执行
            If Not blnNoRecord Then
                '2, "已报到", 3, "已检查", 4, "已报告", 5, "已审核"
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
        Case conMenu_Manage_ModifBaseInfo '基本信息调整
            If Not blnNoRecord Then
                Control.Enabled = intState < 6 And Not blnCancel And mobjCurStudyInfo.lngPatientFrom <= 2 And mobjCurStudyInfo.lngBaby <= 0
            Else
                Control.Enabled = False
            End If
        Case conMenu_Manage_Receive   '检查报到(&L)
            Control.Enabled = Not Control.Enabled
            Control.Enabled = Not Control.Enabled
            If Not blnNoRecord Then
                Control.Enabled = intState <= 1 And intState <> -1 And Not blnCancel
            Else
                Control.Enabled = False
            End If
        
        Case conMenu_Manage_Logout   '取消报到(&D)
            If blnNoRecord Then
                Control.Enabled = False
            ElseIf Control.Parent Is Nothing Then '当使用热键时，如果不判断parent，将会产生异常
                Exit Sub
            ElseIf Control.Parent.type = xtpControlPopup Then
                Control.ToolTipText = "取消报到"
                Control.Caption = "取消报到(&D)"
                Control.Enabled = (intState = 2 Or intState = 3)

            Else ' 工具栏中的用取消检查代替取消登记,同一按键完成取消登记和取消检查功能
                Control.Enabled = (intState = 2 Or intState = 3) Or (intState <= 1 And intState <> -1 And Not blnCancel) '被拒绝的不能被再次拒绝
                Control.ToolTipText = IIf(intState <= 1 And intState <> -1, "取消登记", "取消报到")
                Control.Caption = "取消"
            End If
                        
             If Control.ToolTipText = "取消登记" Then
                Control.Enabled = Control.Enabled And CheckPopedom(mstrPrivs, "检查登记")
            Else
                Control.Enabled = Control.Enabled And CheckPopedom(mstrPrivs, "取消报到")
            End If
            
        Case conMenu_Manage_InQueue    '排队叫号入队
            Control.Visible = mSysPar.blnUseQueue And Not mSysPar.blnAutoInQueue
            Control.Enabled = (intState >= 2 And intState <= 5)
            
        Case conMenu_Manage_Schedule                        '检查预约
            If mblnIsScheduleDept = False Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Enabled = (intState = 0 Or intState = 1)
                If Control.Enabled = True Then
                    '只有预约项目，才能打开检查预约
                    Control.Enabled = mblnIsScheduleOrder
                End If
            End If
            
        Case conMenu_Manage_ScheduleManage                  '预约管理
                Control.Visible = mblnIsScheduleDept
                Control.Enabled = mblnIsScheduleDept
            
        Case conMenu_Manage_Transfer   '关联影像(&C)
            Control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '在2---5之间可用
            
        Case conMenu_Manage_Cancel   '取消关联(&B)
            If (intState >= 2 And intState <= 5) Or intState = -1 Then
                Control.Enabled = mobjCurStudyInfo.strStudyUID <> ""
            Else
                Control.Enabled = False
            End If
            
        Case conMenu_Manage_AttachMoney, conMenu_Manage_CompleteAttach
            Control.Enabled = intState >= 1 And intState < 6
            
        Case conMenu_Manage_Review  '随访
            If (Not blnNoRecord And intState > 1 And intState <= 6) Or intState = -1 Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
        Case conMenu_Tool_Analyse   '高级图像处理
            If (Not blnNoRecord And intState > 1 And intState < 6) Or intState = -1 Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
        Case conMenu_Manage_LookMecRecord '病案查阅
            If mobjCurStudyInfo.lngPageID > 0 Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
        Case conMenu_Manage_Release     '报告发放,报到后，完成后都可以执行
        

            Control.Enabled = IIf(intState >= 2, True, False)

        
            If Not blnNoRecord Then
              '修改报告发放按钮的标题
                 If Not blnNoRecord Then
                     If mobjCurStudyInfo.intReportGiveOut = 1 And mobjCurStudyInfo.intFilmGiveOut = 1 Then
                         Control.Caption = "收回"
                         Control.ToolTipText = "收回已经发放的报告或胶片"
                         Control.Enabled = Control.Enabled And CheckPopedom(mstrPrivs, "取消发放")
                     Else
                         Control.Caption = "发放"
                         Control.ToolTipText = IIf(Control.ID = conMenu_Manage_Release, "报告或胶片发放", "报告和胶片同时发放")
                     End If
                 End If
            End If
            
            Control.Enabled = Not Control.Enabled
            Control.Enabled = Not Control.Enabled
            
        Case conMenu_Manage_FilmRelease
            Control.Enabled = IIf(intState >= 2, True, False)
            
            If Not blnNoRecord Then
                If mobjCurStudyInfo.intFilmGiveOut = 1 Then
                    Control.Caption = "胶片收回"
                    Control.ToolTipText = "收回已经发放的胶片"
                    Control.Enabled = Control.Enabled And CheckPopedom(mstrPrivs, "取消发放")
                Else
                    Control.Caption = "胶片发放"
                    Control.ToolTipText = "胶片发放"
                    Control.Enabled = Control.Enabled And mobjCurStudyInfo.strStudyUID <> ""
                End If
            End If

        Case conMenu_Manage_ReportRelease
            Control.Enabled = IIf(intState >= 4, True, False)
            
            If Not blnNoRecord Then
 
                If mobjCurStudyInfo.intReportGiveOut = 1 Then
                    Control.Caption = "报告收回"
                    Control.ToolTipText = "收回已经发放的报告"
                    Control.Enabled = Control.Enabled And CheckPopedom(mstrPrivs, "取消发放")
                Else
                    Control.Caption = "报告发放"
                    Control.ToolTipText = "报告发放"
                    Control.Enabled = Control.Enabled And CheckPopedom(mstrPrivs, "报告发放")
                End If
 
            End If
            Control.Enabled = Not Control.Enabled
            Control.Enabled = Not Control.Enabled
        
        Case conMenu_Manage_SendArrange                     '发送安排
            Control.Enabled = IIf(intState >= 2 And intState < 6, True, False)
            
        Case conMenu_Manage_SendAudit               '发送审核
            Control.Enabled = IIf(intState = 4, True, False)
            
        Case conMenu_Manage_ReportExecutor      '报告执行
            Control.Enabled = IIf(intState >= 2 And intState <= 6, True, False)
            
        Case conMenu_Manage_PacsCritical
            Control.Enabled = intState >= 2 Or intState = -1   '在2---6之间可用
            
        Case conMenu_Manage_PacsCriticalReg
            Control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1  '在2---5之间可用
            
        Case conMenu_Manage_PacsCriticalManage
            Control.Enabled = intState >= 2 Or intState = -1   '在2---6之间可用

        Case conMenu_Manage_Result, conMenu_Manage_Negative, conMenu_Manage_Positive   '结果阴阳性(&X)
            If mSysPar.blnIgnoreResult = True Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '在2---5之间可用
                If mobjCurStudyInfo.intDangerState = 1 And Control.ID = conMenu_Manage_Result Then Control.Enabled = False
            End If
            
        Case conMenu_Manage_FuHe, conMenu_Manage_JiBenFuHe, conMenu_Manage_BuFuHe, conMenu_Manage_FuHeLevel '符合情况
            If mSysPar.lngConformDetermine = 0 Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '在2---5之间可用
            End If
        
        Case conMenu_Manage_GChannel, conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel '绿色通道标记/取消
            Control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '在2---5之间可用

        Case conMenu_Manage_Complete   '检查完成(&E)
            Control.Enabled = ((intState = 4 Or intState = 5) Or ((intState = 2 Or intState = 3) And (mSysPar.blnNoSignFinish)))

        Case conMenu_Manage_Undone   '取消完成(&U)
            Control.Enabled = intState = 6

        Case conMenu_File_SendImg  '发送图像
        
        Case conMenu_Img_OpenView, conMenu_img_ContrastView '影像对比,影像观片

            If Not mobjWork_PacsImg Is Nothing Or mstrSelTabName = C_TAB_NAME_影像图像 Then
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
            
        Case conMenu_Manage_RelatingPatiet  '关联病人
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
        Case conMenu_File_PrintSet     '打印设置(&S)
        Case conMenu_File_Excel         '清单打印(&L)
            Control.Enabled = Not blnNoRecord
        Case conMenu_File_Parameter, conMenu_Cap_DevSet
        
        Case conMenu_Manage_ChangeUser  '用户交换
            If mSysPar.blnChangeUser Then
                Control.Visible = True
            Else
                Control.Visible = False
            End If
            
        Case conMenu_Manage_SwitchUser  '切换用户
            If mSysPar.blnSwitchUser Then
                Control.Visible = True
            Else
                Control.Visible = False
            End If
        
        Case conMenu_Manage_SetXWParam      '新网PACS参数设置，如果有此菜单，就显示
        Case conMenu_ReportPopup, conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99# '报表
        Case conMenu_FilePopup, conMenu_ManagePopup, conMenu_ViewPopup, conMenu_HelpPopup
        Case conMenu_ToolPopup, conMenu_Tool_Valid
        Case conMenu_Help_Help, conMenu_Help_About  '帮助
        Case conMenu_Help_Web, conMenu_Help_Web_Forum, conMenu_Help_Web_Home, conMenu_Help_Web_Mail '帮助WEB
        Case conMenu_File_Exit, conMenu_EditPopup
        Case ConMenu_File_ShortcutSet
        Case conMenu_Pathol_WorkModule
        Case conMenu_View_ToolBar
        Case conMenu_Manage_Query
        Case conMenu_Manage_QueryCFG
        Case conMenu_Manage_QueryCfgUserScheme
            Control.Enabled = IIf(mlngCur科室ID = 0, False, True)
        Case conMenu_Manage_QueryTabDisplayScheme
            Control.Checked = mSysPar.blnQuickTabDisplayScheme
        Case conMenu_Manage_PacsPlugIn, conMenu_Manage_PacsPlugCfg
        Case conMenu_Manage_PacsPlugIn * 10000# To conMenu_Manage_PacsPlugIn * 10000# + 100#
            '100908             Category属性扩展为3个
            'strTmp:插件是否启用
            strTmp = IIf(UBound(Split(Control.Category, ",")) = 2, Split(Control.Category, ",")(0), Control.Category)
            Control.Enabled = Val(strTmp)
        Case conMenu_Manage_PacsPlugLevel2 * 10000# To conMenu_Manage_PacsPlugLevel2 * 10000# + 9999#
        Case conMenu_Cap_DevSet     '影像设备设置
        Case conMenu_Manage_Change_In   '隐藏列表
        Case conMenu_Img_3D_MMPR, conMenu_Img_3D_MPR, conMenu_Img_3D_PF, conMenu_Img_3D_SA, conMenu_Img_3D_VA, conMenu_Img_3D_VE '三维重建的几个子菜单不需要设置
        Case conMenu_View_FontSize_S    '小字体
             Control.Checked = gbytFontSize = 9
        Case conMenu_View_FontSize_M    '中字体
             Control.Checked = gbytFontSize = 12
        Case conMenu_View_FontSize_L    '大字体
             Control.Checked = gbytFontSize = 15
        
   '-------------------------------------------------收藏管理部分----------------------------------------------------------
 
        Case conMenu_Collection    '收藏(&C)
            Control.Enabled = True
        Case conMenu_Collection_Manage  '收藏管理菜单
            Control.Enabled = True
        Case conMenu_Collection_ViewShare      '查看共享
            Control.Enabled = True
        Case comMenu_Collection_Type * 10000# To comMenu_Collection_Type * 10000# + 9999#  '动态收藏菜单
            Control.Enabled = True
        Case conMenu_Collection_ViewShare * 10000# To conMenu_Collection_ViewShare * 10000# + 9999#  '动态共享菜单
            Control.Enabled = True
         Case conMenu_Collection_To
            
            
    '-------------------------------------------扫描申请单部分-----------------------------------------------

        '扫描申请单
        Case comMenu_Petition_Capture
            If blnCancel Then
                Control.Enabled = False
            Else
                Control.Enabled = IIf((intState >= 2 And intState <= 5) Or intState = -1, True, False)
            End If
            
        '查看申请单
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
            If Control.Caption = "Toolbar Options" Or Control.Caption = "工具栏选项" Then
                Control.Enabled = True
                Exit Sub
            End If
            
            If blnNoRecord Then
                Control.Enabled = False
                Exit Sub
            End If
                    
            
            '已完成除查阅,以及医嘱中报告查看打印，观片菜单外均不可用
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
'功能:初始化模块级变量,仅窗体加载时调用一次
On Error GoTo errH
    Dim rsTemp As ADODB.Recordset
    
    mSysPar.lngListColorMark = NVL(GetDeptPara(lngDeptId, "颜色显示类型", 0))
    mSysPar.blnNameColColorCfg = GetDeptPara(lngDeptId, "姓名颜色区分", 0) = "1"         '姓名颜色区分
    mSysPar.blnOrdinaryNameColColorCfg = GetDeptPara(lngDeptId, "缺省类型病人姓名颜色区分", 0) = "1"       '缺省类型病人姓名颜色区分
    mSysPar.lngAutoImageDays = Val(GetDeptPara(lngDeptId, "自动打开历史图像天数", 0))
    
    If mSysPar.blnNameColColorCfg Then
        gstrSQL = "select 名称 from 病人类型 where 缺省标志=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取缺省病人类型")
        
        If rsTemp.RecordCount > 0 Then mstrDefaultPatientType = NVL(rsTemp!名称)
    End If

    
    mSysPar.blnChangeUser = GetDeptPara(lngDeptId, "允许交换用户", 0) = "1"              '允许交换用户
    mSysPar.blnSwitchUser = GetDeptPara(lngDeptId, "允许切换用户", 0) = "1"              '允许切换用户
    
    mSysPar.blnIsPetitionScan = IIf(Val(GetDeptPara(lngDeptId, "启用申请单扫描", 1)) = 1, True, False)   '读取启用申请单扫描参数
    mSysPar.strImageLevel = NVL(GetDeptPara(lngDeptId, "影像质量等级", "甲,乙"))
    mSysPar.strReportLevel = NVL(GetDeptPara(lngDeptId, "报告质量等级", "甲,乙"))
    mSysPar.bln直接检查 = (Val(GetDeptPara(lngDeptId, "登记后直接检查", 0)) = 1)         '登记后直接检查
    
    mSysPar.lngReportType = Val(GetDeptPara(lngDeptId, "报告编辑器", 0))                 '报告编辑器

'    mSysPar.lngCriticalValues = Val(GetDeptPara(lngDeptId, "危急情况判断", 0))           '危急情况判断
    mSysPar.blnIgnoreResult = GetDeptPara(lngDeptId, "忽略结果阴阳性", 0) = "1" '        '忽略结果阴阳性
    mSysPar.lngConformDetermine = Val(GetDeptPara(lngDeptId, "符合情况判定", 0))         '符合情况判定
    mSysPar.lngImageLevel = Val(GetDeptPara(lngDeptId, "影像质量判定", 0))               '影像质量判定
    mSysPar.lngReportLevel = Val(GetDeptPara(lngDeptId, "报告质量判定", 0))
    
    mSysPar.lngHintType = Val(GetDeptPara(lngDeptId, "诊断结果提示类型", 0))
    
    mSysPar.blnReportWithImage = GetDeptPara(lngDeptId, "有图像才能写报告", 0) = "1" '   '有图像才能写报告
    mSysPar.blnReportWithResult = GetDeptPara(lngDeptId, "无影像诊断为阴性", 0) = "1" '  '无影像诊断为阴性
    mSysPar.blnCompleteCommit = GetDeptPara(lngDeptId, "审核后直接完成", 0) = "1" '      '审核后直接完成
    mSysPar.blnFinallyCompleteCommit = GetDeptPara(lngDeptId, "终审后直接完成", 0) = "1" '终审后直接完成
    mSysPar.blnAuditAutoPrint = IIf(Val(GetDeptPara(lngDeptId, "终审后直接打印", 0)) = 1, True, False) '终审后直接打印
    mSysPar.blnNoSignFinish = GetDeptPara(mlngCur科室ID, "允许未签名报告打印完成", 0) = "1" '       '允许未签名报告打印完成
    mSysPar.blnDirectSendRepImg = IIf(Val(GetDeptPara(lngDeptId, "同步添加观片报告图", 1)) = 1, True, False)
    
    mSysPar.lngBeforeDays = Val(GetDeptPara(lngDeptId, "默认过滤天数", 2)) '                   '默认过滤天数
    If mSysPar.lngBeforeDays > 15 Or mSysPar.lngBeforeDays <= 0 Then
        mSysPar.lngBeforeDays = 2
    End If
    
    mSysPar.blnWriteCapDoctor = GetDeptPara(lngDeptId, "采集图像者为检查技师", 0) = "1"  '采集图像者为检查技师
    
    mSysPar.blnPrintCommit = GetDeptPara(lngDeptId, "打印后直接完成", 0) = "1" '           '打印后直接完成
    mSysPar.blnCanPrint = GetDeptPara(lngDeptId, "平诊需审核才能打报告") = "1"             '平诊需要审核才能打印 =true
    mSysPar.blnAutoSendWorkList = GetDeptPara(lngDeptId, "报道时自动发送WorkList") = "1"   '报道时自动发送WorkList

    '按姓名过滤
    mSysPar.blnNameFuzzySearch = GetDeptPara(lngDeptId, "姓名默认模糊查询", "1") = "1"     '姓名默认模糊查询
    mSysPar.blnNameQueryTimeLimit = GetDeptPara(lngDeptId, "姓名查询时间限制", "1") = "1"  '按姓名过滤时是否进行时间限制
    
    '是否定位报告
    mSysPar.blnIsLocateReport = Val(GetDeptPara(lngDeptId, "检查切换时定位报告编辑", "1")) = 1
    
    If CheckPopedom(mstrPrivs, "排队叫号") And mlngModule <> G_LNG_PATHSTATION_MODULE And CheckPopedom(";" & GetPrivFunc(glngSys, 1160) & ";", "基本") Then      '有权限使用才根据参数启用
        mSysPar.blnUseQueue = GetDeptPara(lngDeptId, "启动排队叫号", 0) = "1" '          '默认不启用排队叫号
        
        If mSysPar.blnUseQueue Then
            mSysPar.blnSynStudylist = GetDeptPara(lngDeptId, "同步定位检查列表", 0)
            mSysPar.blnAutoInQueue = GetDeptPara(lngDeptId, "报到后自动排队", 1)
        End If
    Else
        mSysPar.blnUseQueue = False
    End If
    
    mSysPar.blnRelatingPatient = GetDeptPara(lngDeptId, "启动关联病人", 0) = "1"       '是否使用关
    
    gblnXWLog = (Val(zlDatabase.GetPara("XW记录接口日志", glngSys, G_LNG_XWPACSVIEW_MODULE, "0")) = 1) '是否记录接口日志
    
    Exit Sub
errH:
    If HintError(err, "InitDeptParameter") = 1 Then Resume
End Sub


Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
On Error GoTo errhandle
    '禁止检查列表 拖动
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
         
        Call mobjPacsQueryWrap.Init(mlngCur科室ID, UserInfo.ID, mlngModule, 0, mSysPar.blnCanPrint, mobjSquareCard, Me, objPar)
        
        mobjPacsQueryWrap.DefaultLocate = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\", "DEFLOCATE", True)
        
        cmdLocate.BackColor = IIf(mobjPacsQueryWrap.DefaultLocate, &HFF00&, &H8000000F)
        cmdFind.BackColor = IIf(mobjPacsQueryWrap.DefaultLocate = False, &HFF00&, &H8000000F)
    End If
End Sub

Private Sub InitPars()
    Dim bytFontSize As Byte
    
    Dim strTmpImgPath As String
        
    Call WriteLog("InitPars -> Step 1：开始读取参数...")
    
    '读取字体大小
    bytFontSize = Val(zlDatabase.GetPara("显示字体大小", glngSys, glngModul))
    gbytFontSize = IIf(bytFontSize = 0, 9, IIf(bytFontSize = 1, 12, 15))
    
    Call WriteLog("InitPars -> Step 2：载入本地注册表参数...")
    
    Call InitLocalPars '本地注册表参数
    
    Call WriteLog("InitPars -> Step 3：载入科室流程参数...")
    Call InitDeptParameter(mlngCur科室ID)
     
    
    ReDim gConnectedShardDir(0) As String   '初始化共享目录连接串
     
    Call WriteLog("InitPars -> Step 3：初始化自定义查询相关配置...")
     

    Call WriteLog("InitPars -> Step 5：清理缓存目录...")
    strTmpImgPath = FormatFilePath(GetAppRootPath & "\Apply\TmpImage\")
    ClearCacheFolder strTmpImgPath     '若临时目录满了，则清空该目录
    
    '判断临时目录是否存在
    If Dir(strTmpImgPath, vbDirectory) = "" Then
        Call MkDir(strTmpImgPath)
    End If
    
    Call WriteLog("InitPars -> Step 6：初始化双用户登录参数...")
    '初始化双用户登陆的参数
    mblnCnOracleIsHIS = True
    mintChangeUserState = 1
    
    mstrHisUserName = UserInfo.姓名
    mstrOtherUserName = UserInfo.姓名
    mstrHisUserID = UserInfo.用户名
    mstrOtherUserID = UserInfo.用户名
    
    Set mcnOracleHIS = gcnOracle
    
    Me.stbThis.Panels(4).Text = "报告医生：" & mstrHisUserName & "   检查医生：" & mstrOtherUserName
    
    ReDim mobjPacsReportArry(0) As frmReport
    
    Call WriteLog("InitPars -> Step 7：读取新版观片启用状态...")
    gblnUseXinWangView = False
    
    If mlngModule = G_LNG_PACSSTATION_MODULE Then
        gblnUseXinWangView = IsUseXwViewer
    End If

    
    Call WriteLog("InitPars -> Step End：结束执行...")
End Sub


'Private Sub Form_Load()
'On Error GoTo errHandle
'    '初始化相关方法在showstation中调用InitForm进行处理......
'    '这里不能进行相关的初始化处理是因为在clsPacsWork的BHCodeMain方法中，设置显示方式的时候，会触发Load事件，
'    '而Load事件中的某些处理需要相关参数才能正确执行，因此需要将Load中的处理方法单独提取出来，放入ShowStation方法中执行...
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
    
    '使用Activex的视频采集方式退出
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
    
    '关闭消息中心
    If Not gobjMsgCenter Is Nothing Then
        Call gobjMsgCenter.CloseMsgCenter
    End If
 
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\", "列表检查信息高度设置", mlngMove)
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\", "隐藏检查列表", dkpMain.Panes(1).hidden)
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\", "隐藏辅助模块", dkpMain.Panes(2).hidden)
    
'    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsList), vsList.Name, mstrCol)
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    
    If Me.ScaleWidth > 0 Then
        Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "ListWidth", picList.Width / Me.ScaleWidth)
        Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "HelperWidth", ucPacsHelper1.Width / Me.ScaleWidth)
    End If
    
    '设置字体大小
    zlDatabase.SetPara "显示字体大小", IIf(gbytFontSize = 9, 0, IIf(gbytFontSize = 12, 1, IIf(gbytFontSize = 15, 2, gbytFontSize))), glngSys, glngModul
    
    '恢复窗口名称
    Me.Caption = GetWindowCaption
    
    '保存ucpacsHelper部件串
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\", "HELPER", ucPacsHelper1.GetLayoutStr)
    
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\", "DEFLOCATE", mobjPacsQueryWrap.DefaultLocate)
    
    Call SaveWinState(Me, App.ProductName)
    
    Call ucPacsHelper1.Destory
    
    Call ResetNullParent
    
    Call dkpMain.CloseAll
    
    Call DisposeObj
    
    '恢复导航台的数据库联接
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
    
    
    
    '移出DataExchange目录文件
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
'按指定格式初始化卡类型
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
'初始化临时本地参数，以个人设置为主,窗体加载，过滤，本地设置等调用

    mstrCaptureHot = GetSetting("ZLSOFT", "公共模块", "采集热键", "F8")
    mstrCaptureAfterHot = GetSetting("ZLSOFT", "公共模块", "后台采集热键", "F7")
    mstrCaptureAfterTagHot = GetSetting("ZLSOFT", "公共模块", "标记更新热键", "F6")
    
    mlngMove = Val(GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\", "列表检查信息高度设置", 0))
    mblnIsHideStudyList = CBool(GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\", "隐藏检查列表", 0))
    mblnIsHideHelper = CBool(GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\", "隐藏辅助模块", 0))
    
errContinue2:
    mSysPar.blnLockAfterCall = zlDatabase.GetPara("呼叫后锁定采集", glngSys, mlngModule, "0")
    mSysPar.strFirstTab = zlDatabase.GetPara("工作首页", glngSys, mlngModule, "") '为空表示不使用定制工作首页功能
    mSysPar.blnAutoOpenReport = (Val(zlDatabase.GetPara("开始检查自动打开报告", glngSys, mlngModule, 0)) = 1)
    mSysPar.blnChoosePrintFormat = (Val(zlDatabase.GetPara("报到打印前选择格式", glngSys, mlngModule, 0)) = 1)
    mSysPar.strLocalRoom = zlDatabase.GetPara("本机执行间名称", glngSys, mlngModule, "")
    mSysPar.blnQueueQuick = IIf(Val(zlDatabase.GetPara("自动弹出快捷呼叫窗口", glngSys, mlngModule, "1")) = 1, True, False)
    mSysPar.lngImageValid = Val(zlDatabase.GetPara("图像校对", glngSys, mlngModule, 0))
    
    mSysPar.blnAutoPrint = Val(zlDatabase.GetPara("报到后自动打印申请单", glngSys, mlngModule, 0)) '报到后自动打印申请单
    mSysPar.blnAutoPrintCheck = Val(zlDatabase.GetPara("自动规避重复申请打印", glngSys, mlngModule, 0))
    
    If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
        '如果是采集模块，才需要执行该参数
        mSysPar.lngVideoStationMoneyExeModle = Val(zlDatabase.GetPara("采集费用执行模式", glngSys, mlngModule, 0))
    ElseIf mlngModule = G_LNG_PACSSTATION_MODULE Then
        mSysPar.lngPacsStationMoneyExeModle = Val(zlDatabase.GetPara("医技费用执行模式", glngSys, mlngModule, 0))
    Else
        mSysPar.lngPatholStationMoneyExeModle = Val(zlDatabase.GetPara("病理费用执行模式", glngSys, mlngModule, 0))
    End If
    
    '报告时观片
    mSysPar.blnShowImgAfterReport = (Val(zlDatabase.GetPara("报告时观片", glngSys, mlngModule, 0)) = 1)
    
    '体检病人完成时不判断费用
    mSysPar.blnPEISNoCheckMoneyFinish = (Val(zlDatabase.GetPara("体检病人完成时不判断费用", glngSys, mlngModule, 0)) = 1)

    '显示常用方案标签
    mSysPar.blnQuickTabDisplayScheme = Val(zlDatabase.GetPara("显示常用方案标签", glngSys, mlngModule, 0)) = 1
    
    '得到注册表中关于工具栏显示状态的值，如果为空则等于9
    mintToolBarWriteReg = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\CommandBars", "cbrMainButtonText", 9))
    
End Sub

Private Function InitDepts() As Boolean
'功能：初始化住院临床科室
On Error GoTo errH
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str科室IDs As String, str来源 As String
    
    mlngCur科室ID = 0
    mstrCur科室 = ""
    mstrCanUse科室 = ""
    mblnAllDepts = False
    
    str来源 = "1,2,3"
    If CheckPopedom(mstrPrivs, "所有科室") Then
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where B.部门ID = A.ID " & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " and (A.站点='" & gstrNodeNo & "' Or A.站点 is Null ) " & _
            " And instr([1],','||B.服务对象||',')> 0 And B.工作性质 IN('检查')" & _
            " Order by A.编码"
    Else
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B,部门人员 C " & _
            " Where B.部门ID = A.ID And A.ID=C.部门ID And C.人员ID=" & UserInfo.ID & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " and (A.站点='" & gstrNodeNo & "' Or A.站点 is Null ) " & _
            " And instr([1],','||B.服务对象||',')>0  And B.工作性质 IN('检查')" & _
            " Order by A.编码"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, GetWindowCaption, CStr("," & str来源 & ","))
    
    If rsTmp.EOF Then
        HintMsg "没有发现医技科室信息,请先到部门管理中设置。", "InitDepts", vbInformation
        Exit Function
    Else
        str科室IDs = GetUser科室IDs
        Do Until rsTmp.EOF
            mstrCanUse科室 = mstrCanUse科室 & "|" & rsTmp!ID & "_" & rsTmp!编码 & "-" & rsTmp!名称
            mstrCanUse科室IDs = mstrCanUse科室IDs & "," & rsTmp!ID
            
            If rsTmp!ID = UserInfo.部门ID Then mlngCur科室ID = rsTmp!ID: mstrCur科室 = rsTmp!编码 & "-" & rsTmp!名称 '提取默认科室
            If InStr("," & str科室IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur科室ID = 0 Then mlngCur科室ID = rsTmp!ID: mstrCur科室 = rsTmp!编码 & "-" & rsTmp!名称 '没有默认科室,取所属检查科室第一个
            rsTmp.MoveNext
        Loop
        
        mstrCanUse科室 = Mid(mstrCanUse科室, 2)
        mstrCanUse科室IDs = Mid(mstrCanUse科室IDs, 2)
        
        If CheckPopedom(mstrPrivs, "所有科室") And mlngCur科室ID = 0 Then
            mlngCur科室ID = Split(Split(mstrCanUse科室, "|")(0), "_")(0)
            mstrCur科室 = Split(Split(mstrCanUse科室, "|")(0), "_")(1)
        End If
        
        If mlngCur科室ID = 0 And Not CheckPopedom(mstrPrivs, "所有科室") Then  '没有所有科室操作权限,而且操作者科室不属于检查类科室
            HintMsg "没有发现你所属科室,不能使用此工作站。", "InitDepts", vbInformation
            Exit Function
        End If
        
        Call SetParaUseImgSignValid(mlngCur科室ID)
        InitDepts = True
    End If
    
    If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
        glngXWDeptID = mlngCur科室ID
    End If
    Exit Function
errH:
    If HintError(err, "InitDepts") = 1 Then Resume
End Function

Private Sub InitLayout()
    Dim dblListWidth As Double
    Dim dblHelperWidth As Double
    
    '初始界面布局
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane
    With Me.dkpMain
        .SetCommandBars cbrMain
        
        .options.HideClient = True
        .options.UseSplitterTracker = False '实时拖动
        .options.ThemedFloatingFrames = True
        .options.AlphaDockingContext = True
    End With
    
'    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    
    dblListWidth = NVL(GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "ListWidth", 0.35))
    dblHelperWidth = NVL(GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "HelperWidth", 0.25))
    
    If dblListWidth >= 0.7 Or dblListWidth <= 0.05 Then dblListWidth = 0.35
    If dblHelperWidth >= 0.7 Or dblHelperWidth <= 0.05 Then dblHelperWidth = 0.25
    
    '注册表中保存的界面布局Pnae数量不对，则加载默认的Pane设置
    If dkpMain.PanesCount <> 3 Then
        dkpMain.DestroyAll
 
        Set Pane1 = dkpMain.CreatePane(1, dblListWidth * 1024, 0, DockLeftOf, Nothing)
        Pane1.title = "检查列表"
        Pane1.Handle = picList.hwnd
        Pane1.options = PaneNoCloseable Or PaneNoFloatable
        
        Set Pane2 = dkpMain.CreatePane(2, dblHelperWidth * 1024, 0, DockRightOf, Nothing)
        Pane2.title = "辅助窗体"
        Pane2.Handle = picHelper.hwnd
        Pane2.options = PaneNoCaption Or PaneNoCloseable
        
        Set Pane3 = dkpMain.CreatePane(3, (1 - dblListWidth - dblHelperWidth) * 1024, 0, DockRightOf, Pane2)
        Pane3.title = "子窗体"
        Pane3.Handle = picWindow.hwnd
        Pane3.options = PaneNoCaption Or PaneNoCloseable

    End If
    
    If mblnIsHideStudyList Then Call dkpMain.Panes(1).Hide
    If mblnIsHideHelper Then Call dkpMain.Panes(2).Hide
End Sub

Public Sub StyleChange(ByVal lngStyle As TColorStyle)
'样式改变
    Dim lngMainColor As Long
    
    Select Case lngStyle
        Case sBlue '蓝色样式
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
        Case sGray '灰色样式
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
        Case sAshen '灰白样式
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
    pic主界面遮挡.BackColor = lngMainColor
    picDetail.BackColor = lngMainColor
    picFilter.BackColor = lngMainColor
    PicFucs.BackColor = lngMainColor
    cmdFind.BackColor = lngMainColor
    cmdLocate.BackColor = lngMainColor
End Sub


Private Sub InitCommandBars()
    '功能创建工具条
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
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '菜单定义
    'Begin------------------------文件菜单--------------------------------------默认可见
    Me.cbrMain.ActiveMenuBar.title = "菜单"
    Set cbrMenuBar = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_FilePopup, "文件", "", 0, False)
    With cbrMenuBar.CommandBar
        
        Set cbrPopupControl = CreateMenu(.Controls, xtpControlPopup, conMenu_Manage_SetXWParam, "配置", "", 181, False)
        
        Set cbrControl = CreateMenu(cbrPopupControl.CommandBar.Controls, xtpControlButton, conMenu_File_PrintSet, "打印设置", "", 181, True)
        Set cbrControl = CreateMenu(cbrPopupControl.CommandBar.Controls, xtpControlButton, conMenu_File_Parameter, "参数设置", "", 181, False)
        Set cbrControl = CreateMenu(cbrPopupControl.CommandBar.Controls, xtpControlButton, ConMenu_File_ShortcutSet, "快捷键设置", "", 181, False)
        Set cbrControl = CreateMenu(cbrPopupControl.CommandBar.Controls, xtpControlButton, conMenu_Pathol_WorkModule, "站点模式设置", "", 9004, False)
        Set cbrControl = CreateMenu(cbrPopupControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsPlugCfg, "插件配置", "", 181, False)
        
        '增加视频采集设置菜单
        If mlngModule <> G_LNG_PACSSTATION_MODULE Then
            Set cbrControl = CreateMenu(cbrPopupControl.CommandBar.Controls, xtpControlButton, conMenu_Cap_DevSet, "视频设置", "视频设置", 815, False)
        End If
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_File_Excel, "清单打印", "", 103, True)
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_PacsReport_Preview, "预览", "", 102, True)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_PacsReport_Print, "打印", "", 103, False)
        
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_SwitchUser, "切换用户", "切换用户", 3012, True)
        
        If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
            '增加用户交换菜单,仅影像采集系统有此功能
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_ChangeUser, "交换用户", "交换检查医生和报告医生", 3012, False)
        End If
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_File_SendImg, "发送图像", "", 3061, True)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_Change_In, "隐藏列表", "", 0, False)
        
    
        'Begin----------------------帮助菜单--------------------------------------默认可见
        Set cbrPopupControl = CreateMenu(.Controls, xtpControlPopup, conMenu_HelpPopup, "帮助", "", 0, True)
        With cbrPopupControl.CommandBar
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Help_Help, "帮助主题", "", 0, False)
            Set cbrControl = CreateMenu(.Controls, xtpControlButtonPopup, conMenu_Help_Web, "WEB上的中联", "", 0, False)
                With cbrControl.CommandBar
                    Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_Help_Web_Forum, "中联论坛", "", 0, False)
                    Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_Help_Web_Home, "中联主页", "", 0, False)
                    Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_Help_Web_Mail, "发送反馈", "", 0, False)
                End With
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Help_About, "关于…", "", 0, True)
        End With
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_File_Exit, "退出", "", 191, True)
    End With



'Begin----------------------检查菜单--------------------------------------默认可见
    Set cbrMenuBar = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_ManagePopup, "检查", "", 0, False)
    With cbrMenuBar.CommandBar
    
        Set cbrControl = CreateMenu(.Controls, xtpControlPopup, conMenu_Manage_Request, "申请单", "申请单", 0, False)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButtonPopup, conMenu_Manage_RequestPrint, "打印申请单据", "", 0, False)
        
            '如果启用申请单扫描参数 勾选，则加载“扫描申请单”菜单，未勾选则 不加载
            If mSysPar.blnIsPetitionScan Then
                Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, comMenu_Petition_Capture, "扫描申请单", "", 5020, , False)
                Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, comMenu_Petition_View, "查看申请单", "查看已扫描的申请单图像", 3935, True)
            End If
            
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_CheckList, "查看电子申请", "查看电子申请单", 8044, False)
        
        If InStr(mstrWorkModule, ";检查报告;") >= 1 Then
            '有检查报告模块，才能进行书写
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_PacsReport_Write, "书写", "", 2607, True)
        End If
        
        If mlngModule <> G_LNG_PACSSTATION_MODULE Then
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Tool_Analyse, "高级处理", "高级图像处理", 0, True)
        Else
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Img_OpenView, "观片", "打开影像图像", 8111, True)
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_img_ContrastView, "对比", "对比影像图像", 8112, False)
        End If
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Check_ViewLink, "查看关联检查", "查看关联检查", 102, False)
        
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_Regist, "检查登记", "", 2110, True)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_CopyCheck, "复制登记", "", 0, False)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_ReGet, "召回取消", "", 0, False)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_Receive, "检查报到", "", 744, False)
        If mlngModule = G_LNG_VIDEOSTATION_MODULE Then  '只有影像采集系统需要锁定功能
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Cap_StudySyncState, "锁定", "采集锁定", 6884, False)
        End If
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_InQueue, "入队", "开始排队", 3534, True)
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_Schedule, "检查预约", "", 0, False)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_ScheduleManage, "预约管理", "", 0, False)
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_Transfer, "关联影像", "", 505, True)
                
        If mlngModule = G_LNG_PACSSTATION_MODULE Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
            Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_SendArrange, "发送安排", "", 232, False)
        End If
        
        '审核人
        Set cbrControl = CreateMenu(.Controls, xtpControlPopup, conMenu_Manage_SendAudit, "发送审核", "发送到审核人", 0, False)
        Call CreateAuditorMenu(cbrControl)
        
'        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_LookRelatetion, "查看关联检查", "", , False)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_LookMecRecord, "病案查阅", "", 102, False)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_ReportExecutor, "报告执行", "指定当前报告的记录人", 5008, True)
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_Complete, "检查完成", "", 225, False, , False)
        
        Set cbrControl = CreateMenu(.Controls, xtpControlPopup, conMenu_Manage_Change_Undo, "撤销回退", "撤销回退", 0, True)
        If Not cbrControl Is Nothing Then
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Redo, "取消登记", "", 742, False)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Logout, "取消报到", "", 743, False)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Undone, "取消完成", "", 2615, False)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Cancel, "取消关联", "", 506, False)
        End If
        
        Set cbrControl = CreateMenu(.Controls, xtpControlPopup, conMenu_Manage_State, "检查标记", "检查标记", 0, True)

            If mlngModule = G_LNG_PACSSTATION_MODULE Then
                Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlPopup, conMenu_Manage_Release, "发放处理", "报告或胶片发放处理", 3013, False)
                If Not cbrPopControl Is Nothing Then
                    Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "报告发放", "", 8215, False)
                    Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FilmRelease, "胶片发放", "", 8216, False)
                End If
            Else
                Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "报告发放", "", 8215, False)
            End If
            '检查结果
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlPopup, conMenu_Manage_Result, "阴阳性", "", 0, False)
            If Not cbrPopControl Is Nothing Then
                Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Negative, "结果阳性", "", 3506, False)
                Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Positive, "结果阴性", "", 3507, False)
            End If
            '符合情况
            If mlngModule <> G_LNG_PATHOLSYS_NUM Then
                Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlPopup, conMenu_Manage_FuHeLevel, "符合情况", "", 0, False)
                If Not cbrPopControl Is Nothing Then
                    Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FuHe, "符合", "", 3587, False)
                    Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_JiBenFuHe, "基本符合", "", 3010, False)
                    Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_BuFuHe, "不符合", "", 3010, False)
                End If
            End If
                
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlPopup, conMenu_Manage_GChannel, "绿色通道", "", 0, False, , False)
            If Not cbrPopControl Is Nothing Then
                Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_GChannelOk, "标记", "", 0, False, , False)
                Call CreateMenu(cbrPopControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_GChannelCancel, "取消", "", 0, False, , False)
            End If
        
        
        Set cbrControl = CreateMenu(.Controls, xtpControlPopup, conMenu_Manage_More, "更多操作", "更多操作", 0, True)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ThingModi, "修改信息", "", 0, False)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ModifBaseInfo, "基本信息调整", "", 4113, False)
            
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ExecOnePart, "分部位执行", "分部位执行和取消医嘱", 0, True)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Review, "附加信息", "", 232, False)
    
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_DiseaseRegist, "传染病登记", "传染病登记", 3564, True)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_DiseaseQuery, "传染病查询", "传染病查询", 102, False)
            
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsCriticalReg, "危急患者登记", "", 8344, True)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsCriticalManage, "危急患者管理", "", 8345, False)
        
        
            If Not (mobjAppendBill Is Nothing) And GetInsidePrivs(p医嘱附费管理, True) <> "" Then
                Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_AttachMoney, "附加费用", "", 3011, True)
                
                If glngModul = G_LNG_PATHSTATION_MODULE Then
                    Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_CompleteAttach, "完成补费", "", 3816, False)
                End If
            End If
        

            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_RelatingPatiet, "关联病人", "", 803, True)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Burn, "图像刻录", "", 0, True)
    
    End With
    
    'Begin-------------------------------------------------------收藏菜单(默认可见)----------------------------------------------------------

    gstrSQL = "select a.ID ,a.上级id,b.姓名 as 创建人,a.收藏类别 " & _
                " from 影像收藏类别 a,人员表 b " & _
                " where a.创建人ID=" & UserInfo.ID & " and a.创建人id=b.ID(+) Start With a.上级id Is Null Connect By Prior a.ID = a.上级id"
    Set rsCollection = zlDatabase.OpenSQLRecord(gstrSQL, GetWindowCaption)

    gstrSQL = "select a.ID ,a.上级id,b.姓名 as 创建人,a.收藏类别,a.是否共享 " & _
                " from 影像收藏类别 a,人员表 b " & _
                " where a.创建人ID<>" & UserInfo.ID & " and a.创建人id=b.ID(+) Start With a.上级id Is Null Connect By Prior a.ID = a.上级id"
    Set rsViewShare = zlDatabase.OpenSQLRecord(gstrSQL, GetWindowCaption)
        
    Set cbrMenuBar = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_Collection, "收藏", "", 0, False)
    With cbrMenuBar.CommandBar
        
        '克隆对象 筛选出共享的数据进行判断
        Set rsShareCount = zlDatabase.CopyNewRec(rsViewShare)
        rsShareCount.Filter = "是否共享=1"
        
        If rsShareCount.RecordCount <> 0 Then
           '递归创建共享菜单
           mlngShareFatherID = 0
           Set rsTemp = zlDatabase.CopyNewRec(rsViewShare)
           rsViewShare.Filter = "上级ID=" & NVL(rsViewShare!上级ID, 1) & " and 创建人<> '" & UserInfo.姓名 & "'"
           
           Set cbrControl = CreateMenu(.Controls, xtpControlButtonPopup, conMenu_Collection_ViewShare, "共享查看", "", 0, False)
           Call RecursionCreateShareMenu(rsViewShare, rsTemp, cbrControl)
        End If

        If rsCollection.RecordCount > 0 Then
            '递归创建收藏类别菜单
                 mlngCollectionFatherID = 0
                 Set rsTemp = zlDatabase.CopyNewRec(rsCollection)
                 rsCollection.Filter = "上级ID=" & NVL(rsCollection!上级ID, 1)
                 Call RecursionCreateCollectionMenu(rsCollection, rsTemp, cbrMenuBar)
        End If
        
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Collection_To, "收藏到...", "", 0, True)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Collection_Manage, "收藏管理", "", 0, False)
        
    End With
    
    'Begin----------------------自定义查询菜单--------------------------------------默认可见
    Set cbrMenuBar = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_Manage_Query, "查询", "", 0, False)
    
    Call mobjPacsQueryWrap.RefreshCustomQueryMenu(cbrMenuBar, mintQueryState, tabScheme, mSysPar.blnQuickTabDisplayScheme)
    
    Call CheckHaveScheme(False, "")
    
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_QueryCFG, "查询配置", "", 0, True)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_QueryCfgUserScheme, "常用方案调整", "", 0, False)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Manage_QueryTabDisplayScheme, "显示常用方案标签", "", 0, True)
        cbrControl.Checked = mSysPar.blnQuickTabDisplayScheme
        cbrControl.CloseSubMenuOnClick = False
    End With
    
    
    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Set cbrMenuBar = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_ReportPopup, "报表", "", 0, True)
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
    
    
    'Begin----------------------查看菜单--------------------------------------
    Set cbrMenuBar = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_ViewPopup, "查看", "", 0, False)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateMenu(.Controls, xtpControlPopup, conMenu_View_ToolBar, "工具栏", "", 0, False)
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar '二级菜单
                Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮", "", 0, False): cbrPopControl.Checked = True
                Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签", "", 0, False): cbrPopControl.Checked = True
                Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_View_ToolBar_Size, "大图标", "", 0, False): cbrPopControl.Checked = True
            End With
            
        Set cbrControl = CreateMenu(.Controls, xtpControlButtonPopup, conMenu_View_FontSize, "字体大小", "", 0, False)
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar '二级菜单
                Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_View_FontSize_S, "小字体", "", 0, False): cbrPopControl.Checked = True
                Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_View_FontSize_M, "中字体", "", 0, False)
                Set cbrPopControl = CreateMenu(.Controls, xtpControlButton, conMenu_View_FontSize_L, "大字体", "", 0, False)
            End With
            
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_View_StatusBar, "状态栏", "", 0, True): cbrControl.Checked = True
        Set cbrControl = CreateMenu(.Controls, xtpControlButtonPopup, conMenu_View_Filter * 10#, "检查科室", "", 0, False)
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_View_Refresh, "刷新", "", 0, False)
    End With
        
    'Begin----------------------工具菜单--------------------------------------默认可见
    Set cbrMenuBar = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_ToolPopup, "工具", "", 0, False)
    With cbrMenuBar.CommandBar
        'Begin----------------------第三方功能插件菜单--------------------------------------默认可见
        Call RefreshCustomPlugInMenu(cbrMenuBar, mlngModule)
    
        '其他工具菜单
        Set cbrControl = CreateMenu(.Controls, xtpControlButton, conMenu_Tool_Valid, "图像校对工具", "", 0, True)
    End With
        
    '最右边显示浮动采集按钮
    If mlngModule <> G_LNG_PACSSTATION_MODULE And InStr(mstrWorkModule, C_TAB_NAME_影像采集) > 0 Then
        Set cbrControl = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlButton, comMenu_Cap_Process, "浮动采集", "弹出独立采集窗口", 0, False): cbrControl.flags = xtpFlagRightAlign
    End If
        
    '---------------------设置右上角当前科室----------------------------------
    Set cbrControl = CreateMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_View_Filter * 10#, "检查科室", "", 0, False): cbrControl.flags = xtpFlagRightAlign
            
            
    Set objCusControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlCustom, C_LNG_TAB_MENU_ID, "TAB面板")
        objCusControl.Handle = picTabFace.hwnd
        objCusControl.flags = xtpFlagControlStretched
        objCusControl.Category = M_STR_MODULE_MENU_TAG
        
        
    '---------------------工具栏定义------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True

    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Regist, "登记", "检查登记", 211, True)
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Receive, "报到", "检查报到", 744, False)
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Logout, "取消", "取消报到", 743, False)
    If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Cap_StudySyncState, "锁定", "采集锁定", 6884, False)
    End If
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Schedule, "预约", "检查预约", 6823, True)
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_InQueue, "入队", "开始排队", 3534, False)
    
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_PacsReport_Preview, "预览", "报告预览", 102, True)
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_PacsReport_Print, "打印", "报告打印", 103, False)
    
    If InStr(mstrWorkModule, ";检查报告;") >= 1 Then
        '有检查报告模块，才能进行书写
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_PacsReport_Write, "书写", "", 2607, True)
    End If
    

    If mlngModule <> G_LNG_PACSSTATION_MODULE Then
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Tool_Analyse, "高级", "高级图像处理", 0, True)
    Else
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Img_OpenView, "观片", "打开影像图像", 8111, True)
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_img_ContrastView, "对比", "对比影像图像", 8112, False)
    End If
'    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Check_ViewLink, "查看关联", "", 102, False): cbrControl.ToolTipText = "查看关联检查"
    
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_View_Filter, "过滤", "过滤", 0, True)
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_View_Refresh, "刷新", "刷新", 0, False)
        
    Call AddPlugInToolBarMenu(cbrToolBar.Controls, mlngModule)  '100908
    
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Review, "备注", "附加信息", 232, True)
    
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_Request, "申请单", "", 3935, False)
    With cbrControl.CommandBar
        If mSysPar.blnIsPetitionScan Then   '启用了申请单扫描后才能进行查看
            Call CreateMenu(.Controls, xtpControlButton, comMenu_Petition_View, "查看申请单", "查看已扫描的申请单图像", 3935, False)
        End If
        
        Call CreateMenu(.Controls, xtpControlButton, conMenu_Manage_CheckList, "查看电子申请", "查看电子申请单", 8044, False)
    End With
    
    If Not (mobjAppendBill Is Nothing) And GetInsidePrivs(p医嘱附费管理, True) <> "" Then
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_AttachMoney, "补附费", "补附费", 3011, False)
        If glngModul = G_LNG_PATHSTATION_MODULE Then
            Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_CompleteAttach, "完成补费", "完成补费", 3816, False)
        End If
    End If
    
'    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_Disease, "传染病", "传染病", 3842, False)
'    If Not cbrControl Is Nothing Then
'        Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_DiseaseRegist, "传染病登记", "传染病登记", 3564, False)
'        Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_DiseaseQuery, "传染病查询", "传染病查询", 102, False)
'    End If
    
'    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_SwitchUser, "切换", "切换用户", 3012, False, conMenu_Tool_Analyse)
    
    If mlngModule = G_LNG_PACSSTATION_MODULE Then
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_Release, "发放处理", "报告或胶片发放处理", 3013, False)
        If Not cbrControl Is Nothing Then
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "报告发放", "报告发放", 8215, False)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FilmRelease, "胶片发放", "胶片发放", 8216, False)
        End If
    Else
        Set cbrPopControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "报告发放", "报告发放", 8215, False)
    End If
    
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_ReportExecutor, "报告执行", "指定当前报告的记录人", 5008, False)
    
    If mlngModule = G_LNG_PACSSTATION_MODULE Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_SendArrange, "发送安排", "发送安排", 232, False)
    End If
    
'    '危急情况
'    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_PacsCritical, "危急值", "危急情况", 8338, False)
'    If Not cbrControl Is Nothing Then
'        Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsCriticalReg, "危急值登记", "危急值患者登记", 8345, False)
'        Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_PacsCriticalManage, "危急值管理", "危急值患者管理", 8338, True)
'    End If
    
    '检查结果阴阳性
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_Result, "结果", "检查结果阴阳性", 3506, False)
    If Not cbrControl Is Nothing Then
        Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Negative, "阳性", "阳性", 3506, False)
        Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Positive, "阴性", "阴性", 3507, False)
    End If
    
    '如果是病理系统，则没有符合情况按钮
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_FuHeLevel, "符合情况", "符合情况", 8044, False)
        If Not cbrControl Is Nothing Then
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FuHe, "符合", "符合", 0, False)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_JiBenFuHe, "基本符合", "基本符合", 0, False)
            Set cbrPopControl = CreateMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_BuFuHe, "不符合", "不符合", 0, False)
        End If
    End If
        
'    '只有影像采集系统才具有用户交换功能
'    If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
'        Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_ChangeUser, "交换", "交换检查医生和报告医生", 3012, False)
'    End If
    
    Set cbrControl = CreateMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Complete, "完成", "检查最终完成", 225, False, , False)
  
  
    '---------------------创建病理公共菜单及工具栏---------------------
    If mblnIsHasPatholModule Then
        If mobjWork_Pathol Is Nothing Then
            Set mobjWork_Pathol = New clsWorkModule_PatholV2
            Call mobjWork_Pathol.zlInitModule(Me, mlngModule, mstrPrivs, mlngCur科室ID)
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
'创建该模块内的菜单
On Error GoTo err
    Dim blHavePrives As Boolean '是否具备相应菜单权限
    'blnIsControlCreate 是否没有权限也允许创建菜单。
    
    '创建菜单前根据ID 和权限判断是否终止创建过程
    '注意  conMenu_Manage_GChannel  conMenu_Manage_Complete conMenu_Manage_Result conMenu_Edit_Audit
    'conMenu_PacsReport_RepFormat 必须创建
    blHavePrives = True
    
    Select Case lngID
        Case conMenu_File_SendImg '发送图像
            If Not CheckPopedom(mstrPrivs, "文件发送") Then blHavePrives = False
            
        Case conMenu_Manage_Regist, conMenu_Manage_CopyCheck, conMenu_Manage_Redo, conMenu_Manage_ThingModi, comMenu_Petition_View
        '检查登记，复制登记，取消登记, 修改信息,查看申请单
            If Not CheckPopedom(mstrPrivs, "检查登记") Then blHavePrives = False
            
        Case conMenu_Manage_Receive '检查报到
            If Not CheckPopedom(mstrPrivs, "检查报到") Then blHavePrives = False
            
        Case conMenu_Manage_Logout '取消报到
            If Not CheckPopedom(mstrPrivs, "取消报到") Then blHavePrives = False
            
        Case conMenu_Manage_Transfer, conMenu_Manage_Cancel '关联影像 取消关联
            If Not CheckPopedom(mstrPrivs, "图像关联") Then blHavePrives = False
            
        Case conMenu_Manage_Review '随访
            If Not CheckPopedom(mstrPrivs, "随访") Then blHavePrives = False
            
        Case conMenu_Manage_Disease
            If Not (CheckPopedom(mstrPrivs, "传染病阳性结果登记") Or CheckPopedom(mstrPrivs, "传染病阳性结果查询")) Then blHavePrives = False
            
        Case conMenu_Manage_DiseaseRegist
            If Not CheckPopedom(mstrPrivs, "传染病阳性结果登记") Then blHavePrives = False
            
        Case conMenu_Manage_DiseaseQuery
            If Not CheckPopedom(mstrPrivs, "传染病阳性结果查询") Then blHavePrives = False
            
        Case conMenu_Manage_PacsCritical, conMenu_Manage_PacsCriticalReg, conMenu_Manage_PacsCriticalManage
            If Not CheckPopedom(mstrPrivs, "危急值处理") Then blHavePrives = False
            
        Case conMenu_Manage_Undone
            If Not CheckPopedom(mstrPrivs, "取消检查完成") Then blHavePrives = False
            
        Case conMenu_Manage_RelatingPatiet
            If Not (CheckPopedom(mstrPrivs, "关联病人") And mSysPar.blnRelatingPatient) Then blHavePrives = False
            
        Case conMenu_Manage_Burn
            If Not CheckPopedom(mstrPrivs, "图像刻录") Then blHavePrives = False
            
        Case conMenu_Tool_Analyse '高级图像处理
            If Not CheckPopedom(";" & GetPrivFunc(glngSys, 1289) & ";", "基本") Then blHavePrives = False
        '------------------
        Case conMenu_Manage_GChannel, conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel
        '绿色通道 标记 取消
            If Not CheckPopedom(mstrPrivs, "绿色通道") Then blHavePrives = False
            
        Case conMenu_Manage_Complete  '检查完成
            If Not CheckPopedom(mstrPrivs, "检查完成") Then blHavePrives = False
            
        Case conMenu_Manage_ModifBaseInfo '基本信息调整
            If Not CheckPopedom(mstrPrivs, "强制修改住院门诊信息") Then blHavePrives = False
            
        Case conMenu_Manage_ExecOnePart '分部位执行
            If Not CheckPopedom(mstrPrivs, "取消报到") Then blHavePrives = False

        Case conMenu_Manage_ConfigQuery, conMenu_Manage_QueryCFG
            If Not CheckPopedom(mstrPrivs, "查询配置") Then blHavePrives = False
        
        Case conMenu_Manage_ReportExecutor '报告执行
            If Not CheckPopedom(mstrPrivs, "报告执行") Then blHavePrives = False
        
        Case conMenu_Manage_Schedule, conMenu_Manage_ScheduleManage       '检查预约,预约管理
            If Not CheckPopedom(mstrPrivs, "检查预约") Then blHavePrives = False
            
        Case Else
    End Select
    
    If blHavePrives Or Not blnIsControlCreate Then
    
        If lngIndex >= 0 Then
            Set CreateMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
        Else
            Set CreateMenu = objMenuControl.Add(lngType, lngID, strCaption)
        End If
    
        CreateMenu.ID = lngID '如果这里不指定id，则不能将有些菜单添加到右键菜单中
        
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
            Case C_TAB_NAME_影像图像
                If Not mobjWork_PacsImg Is Nothing Then
                    Call mobjWork_PacsImg.zlMenu.zlClearMenu(strModuleTag)
                    Call mobjWork_PacsImg.zlMenu.zlClearToolBar(strModuleTag)
                End If
                
            Case C_TAB_NAME_影像采集
            
            Case C_TAB_NAME_检查报告
                If Not mobjWork_Report Is Nothing Then
                   Call mobjWork_Report.zlMenu.zlClearMenu(strModuleTag)
                   Call mobjWork_Report.zlMenu.zlClearToolBar(strModuleTag)
                End If
                
            Case C_TAB_NAME_医嘱记录, C_TAB_NAME_病历记录, C_TAB_NAME_费用记录, C_TAB_NAME_电子病历
                If Not mobjWork_His Is Nothing Then
                    Call mobjWork_His.zlMenu.zlClearMenu(strModuleTag)
                    Call mobjWork_His.zlMenu.zlClearToolBar(strModuleTag)
                End If
                
            Case C_TAB_NAME_标本核收, C_TAB_NAME_病理取材, C_TAB_NAME_病理制片, C_TAB_NAME_病理特检, C_TAB_NAME_过程报告
                If Not mobjWork_Pathol Is Nothing Then
                    If strModuleTag <> "" Then
                        Call mobjWork_Pathol.zlMenu.zlClearMenu(strModuleTag)
                        Call mobjWork_Pathol.zlMenu.zlClearToolBar(strModuleTag)
                    End If
                End If
                
            Case C_TAB_NAME_排队叫号
        End Select
    Next
End Sub


Private Sub CreateWorkModuleMenu(ByVal strTabName As String, ByVal strModuleTag As String)
'创建工作模块菜单
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
        Case C_TAB_NAME_影像图像
            Call mobjWork_PacsImg.zlMenu.zlCreateMenu(strModuleTag, Me.cbrMain)
            Call mobjWork_PacsImg.zlMenu.zlCreateToolBar(strModuleTag, Me.cbrMain.Item(2))
            
        Case C_TAB_NAME_影像采集
        Case C_TAB_NAME_检查报告
            Call mobjWork_Report.zlMenu.zlCreateMenu(strModuleTag, Me.cbrMain)
            Call mobjWork_Report.zlMenu.zlCreateToolBar(strModuleTag, Me.cbrMain.Item(2))
            
        Case C_TAB_NAME_医嘱记录, C_TAB_NAME_费用记录, C_TAB_NAME_病历记录, C_TAB_NAME_电子病历
            '因为在PACS系统中 “打印” 菜单项在编辑菜单组下，而病历中在文件菜单下，所以在调用病历的菜单创建过程时，
            '在文件菜单下找不到打印菜单项而报错，而PACS中，清单打印在文件菜单下，所以调用病历的菜单创建过程时将
            '清单打印的id改成打印的id，创建完后，恢复清单打印原来的id
            Set cbrControl = Nothing
            
            If strTabName = C_TAB_NAME_电子病历 Then
                Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
                Set cbrControl = cbrMenuBar.CommandBar.Controls.Find(, conMenu_File_Excel)
                
                cbrControl.ID = conMenu_File_Print
                
                Call mobjWork_His.zlMenu.zlCreateMenu(strModuleTag, Me.cbrMain)
                
                If Not cbrControl Is Nothing Then cbrControl.ID = conMenu_File_Excel
                
            Else
                lngToolCount = cbrMain(2).Controls.Count
                
                '调整工具栏菜单按钮ID,默认将工具栏按钮添加到最后位置，模块内部默认根据菜单id不为1和2的按钮项计算起始位置
                For i = 1 To lngToolCount
                    If cbrMain(2).Controls(i).Category = "Main" Then
                        cbrMain(2).Controls(i).ID = CLng(1) & CLng(cbrMain(2).Controls(i).ID)
                        cbrMain(2).Controls(i).Caption = "TMP-" & cbrMain(2).Controls(i).Caption
                    End If
                Next
                
                On Error GoTo errMenu
                    '模块菜单创建如果出现异常，需要恢复id和caption被改变的菜单
                    Call mobjWork_His.zlMenu.zlCreateMenu(strModuleTag, Me.cbrMain)
errMenu:
                
                '恢复工具栏菜单按钮ID
                For i = 1 To lngToolCount
                    If InStr(cbrMain(2).Controls(i).Caption, "TMP-") > 0 Then   '判断按钮是否被临时重置过菜单ID
                        cbrMain(2).Controls(i).ID = CLng(Mid(cbrMain(2).Controls(i).ID, 2, 255))
                        cbrMain(2).Controls(i).Caption = Replace(cbrMain(2).Controls(i).Caption, "TMP-", "")
                    End If
                Next
                
            End If
                         
            
        Case C_TAB_NAME_标本核收, C_TAB_NAME_病理取材, C_TAB_NAME_病理制片, C_TAB_NAME_病理特检, C_TAB_NAME_过程报告
            If Len(strModuleTag) > 0 Then Call mobjWork_Pathol.zlMenu.zlCreateMenu(strModuleTag, Me.cbrMain)
        Case C_TAB_NAME_排队叫号
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
    If HintError(err, "CreateWorkModuleMenu<" & strTabName & "菜单创建>", False) = 1 Then Resume
End Sub

Private Sub RecursionCreateShareMenu(rsFilterADO As ADODB.Recordset, rsFullADO As ADODB.Recordset, cbrParentControl As CommandBarControl, Optional blnIsShare As Boolean = False)
'递归循环创建共享菜单
    Dim rsFilterTemp As ADODB.Recordset
    Dim i As Long
    Dim cbrControl As CommandBarControl
    Static j As Long
    
    If rsFilterADO.RecordCount = 0 Then Exit Sub
    rsFilterADO.MoveFirst
    
    With cbrParentControl.CommandBar.Controls
        If mlngShareFatherID <> 0 Then
            Set cbrControl = .Add(xtpControlButton, CLng(conMenu_Collection_ViewShare) * 10000# + mlngShareFatherID, "查看当前收藏", -1, False)
            cbrControl.Category = M_STR_MODULE_MENU_TAG
        End If
        
        For i = 1 To rsFilterADO.RecordCount
            rsFullADO.Filter = " 上级ID=" & NVL(rsFilterADO!ID)

            If rsFullADO.RecordCount > 0 Then
                Set cbrControl = Nothing
  
                If NVL(rsFilterADO!是否共享) = 1 Or blnIsShare = True Then
                    mlngShareFatherID = NVL(rsFilterADO!ID)
                    '创建父级菜单 如果上级ID=1 则显示共享人姓名
                    Set cbrControl = .Add(xtpControlButtonPopup, CLng(conMenu_Collection_ViewShare) * 10000# + j, NVL(rsFilterADO!收藏类别) & Decode(cbrParentControl.ID, conMenu_Collection_ViewShare, "(" & NVL(rsFilterADO!创建人) & ")", ""), -1, False)
                    cbrControl.DescriptionText = NVL(rsFilterADO!创建人)
                    cbrControl.Category = M_STR_MODULE_MENU_TAG
                    
                    j = j + 1
                    If i = 1 Then cbrControl.BeginGroup = True
                End If
                
                Set rsFilterTemp = zlDatabase.CopyNewRec(rsFullADO)
                '调用自己
                Call RecursionCreateShareMenu(rsFilterTemp, rsFullADO, IIf(cbrControl Is Nothing, cbrParentControl, cbrControl), IIf(NVL(rsFilterADO!是否共享) = 0, False, True))
            Else
            '创建子级菜单
                If NVL(rsFilterADO!是否共享) = 1 Or blnIsShare = True Then
                    Set cbrControl = .Add(xtpControlButton, CLng(conMenu_Collection_ViewShare) * 10000# + j, NVL(rsFilterADO!收藏类别) & Decode(cbrParentControl.ID, conMenu_Collection_ViewShare, "(" & NVL(rsFilterADO!创建人) & ")", ""), -1, False)
                    cbrControl.DescriptionText = NVL(rsFilterADO!创建人)
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
'递归循环创建收藏类别菜单
    Dim rsFilterTemp As ADODB.Recordset
    Dim cbrControl As CommandBarControl
    Dim i As Long
    Static j As Long

    If rsFilterADO.RecordCount = 0 Then Exit Sub
    rsFilterADO.MoveFirst

    With cbrMenuBar.CommandBar.Controls
        If mlngCollectionFatherID <> 0 Then
            Set cbrControl = .Add(xtpControlButton, CLng(comMenu_Collection_Type) * 10000# + mlngCollectionFatherID, "查看当前收藏", -1, False)
            cbrControl.Category = M_STR_MODULE_MENU_TAG
        End If

        For i = 1 To rsFilterADO.RecordCount

            rsFullADO.Filter = " 上级ID=" & NVL(rsFilterADO!ID)
            mlngCollectionFatherID = NVL(rsFilterADO!ID)
            If rsFullADO.RecordCount > 0 Then
            '创建父级菜单
                Set cbrControl = .Add(xtpControlButtonPopup, CLng(comMenu_Collection_Type) * 10000# + j, NVL(rsFilterADO!收藏类别), -1, False)
                cbrControl.Category = M_STR_MODULE_MENU_TAG
                
                j = j + 1
                
                Set rsFilterTemp = zlDatabase.CopyNewRec(rsFullADO)
                '调用自己
                Call RecursionCreateCollectionMenu(rsFilterTemp, rsFullADO, cbrControl)
                
            Else
            '创建子级菜单
                Set cbrControl = .Add(xtpControlButton, CLng(comMenu_Collection_Type) * 10000# + j, NVL(rsFilterADO!收藏类别), -1, False)
                cbrControl.Category = M_STR_MODULE_MENU_TAG
                
                j = j + 1
            End If
            If i = 1 Then cbrControl.BeginGroup = True

            If Not rsFilterADO.EOF Then rsFilterADO.MoveNext
        Next
    End With

End Sub


Private Sub ReadBaseModuleName()
    '设置当前需要创建的工作页面
    mstrWorkModule = zlDatabase.GetPara("站点模块", glngSys, mlngModule, "")
    mstrWorkModule = IIf(mstrWorkModule <> "", ";" & mstrWorkModule & ";", "")
    
    mstrWorkModule = Replace(Replace(mstrWorkModule, "模块", ""), "影像报告", "检查报告")
    mstrWorkModule = Replace(mstrWorkModule, "病理诊断", "检查报告")
    
    If mstrWorkModule = "" Then
        Select Case mlngModule
            Case G_LNG_PACSSTATION_MODULE
                mstrWorkModule = ";影像图像;检查报告;医嘱记录;病历记录;电子病历;费用记录;"
            
            Case G_LNG_VIDEOSTATION_MODULE
                mstrWorkModule = ";影像采集;检查报告;医嘱记录;病历记录;电子病历;费用记录;"
            
            Case G_LNG_PATHOLSYS_NUM
                mstrWorkModule = ";标本核收;影像采集;病理取材;病理制片;病理特检;过程报告;检查报告;医嘱记录;病历记录;电子病历;费用记录;"
            Case Else
                Exit Sub
        End Select
    End If
    
'    '测试代码
'    mstrWorkModule = ";影像图像模块;影像采集模块;标本核收模块;病理取材模块;病理制片模块;病理特检模块;过程报告模块;影像报告模块;费用记录模块;医嘱记录模块;病历记录模块;"
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
    Call ucPacsHelper1.Init(Me, mlngModule, mlngCur科室ID, mstrPrivs, True)
Exit Sub
errhandle:
    Call HintError(err, "InitPacsHelper", False)
End Sub

Public Sub InitWorkModuleTab()
    Dim i As Integer
    Dim strSelModuleName As String
    Dim objTabItem As TabControlItem
    
    mblnIsHasPatholModule = False   '当该变量最后仍然为false时，则根据条件删除病理菜单
    
'    strSelModuleName = "影像采集" '从注册表读取上次选择的工作模块
    
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
        
        '读取工作模块配置
        Call ReadBaseModuleName
    
        If InStr(mstrWorkModule, ";影像图像;") > 0 Then
            Call CreateTabItem(0, C_TAB_NAME_影像图像, 3551, strSelModuleName)
        End If
                        
        If mlngModule <> G_LNG_PACSSTATION_MODULE And CheckPopedom(mstrPrivs, "视频采集") _
            And InStr(mstrWorkModule, ";影像采集;") > 0 Then
            Call CreateTabItem(1, C_TAB_NAME_影像采集, conMenu_Cap_Dynamic, strSelModuleName)
        Else
            mstrWorkModule = Replace(mstrWorkModule, "影像采集", "")
        End If
        
        If CheckPopedom(mstrPrivs, "标本核收") And InStr(mstrWorkModule, ";标本核收;") > 0 Then
            Call CreateTabItem(2, C_TAB_NAME_标本核收, G_INT_ICONID_SPECIMEN, strSelModuleName)
            
            mblnIsHasPatholModule = True
        Else
            mstrWorkModule = Replace(mstrWorkModule, "标本核收", "")
        End If
        
        If CheckPopedom(mstrPrivs, "病理取材") And InStr(mstrWorkModule, ";病理取材;") > 0 Then
            Call CreateTabItem(3, C_TAB_NAME_病理取材, G_INT_ICONID_MATERIAL, strSelModuleName)
            
            mblnIsHasPatholModule = True
        Else
            mstrWorkModule = Replace(mstrWorkModule, "病理取材", "")
        End If
        
        If CheckPopedom(mstrPrivs, "病理制片") And InStr(mstrWorkModule, ";病理制片;") > 0 Then
            Call CreateTabItem(4, C_TAB_NAME_病理制片, G_INT_ICONID_SLICES, strSelModuleName)
            
            mblnIsHasPatholModule = True
        Else
            mstrWorkModule = Replace(mstrWorkModule, "病理制片", "")
        End If
        
        If (CheckPopedom(mstrPrivs, "免疫组化") Or CheckPopedom(mstrPrivs, "特殊染色") Or CheckPopedom(mstrPrivs, "分子病理")) _
            And InStr(mstrWorkModule, ";病理特检;") > 0 Then
            Call CreateTabItem(5, C_TAB_NAME_病理特检, G_INT_ICONID_SPEEXAM, strSelModuleName)
            
            mblnIsHasPatholModule = True
        Else
            mstrWorkModule = Replace(mstrWorkModule, "病理特检", "")
        End If
        
        If (CheckPopedom(mstrPrivs, "冰冻报告") Or CheckPopedom(mstrPrivs, "特染报告") _
            Or CheckPopedom(mstrPrivs, "分子报告") Or CheckPopedom(mstrPrivs, "免疫报告") _
            Or CheckPopedom(mstrPrivs, "冰冻特检报告查阅")) And InStr(mstrWorkModule, ";过程报告;") > 0 Then
            Call CreateTabItem(6, C_TAB_NAME_过程报告, G_INT_ICONID_PROREPORT, strSelModuleName)
            
            mblnIsHasPatholModule = True
        Else
            mstrWorkModule = Replace(mstrWorkModule, "过程报告", "")
        End If
        
        If GetInsidePrivs(p诊疗报告管理, True) <> "" And _
             InStr(mstrWorkModule, ";检查报告;") > 0 Then
            Call CreateTabItem(7, C_TAB_NAME_检查报告, 10008, strSelModuleName) 'conMenu_Edit_Compend
        Else
            mstrWorkModule = Replace(mstrWorkModule, "检查报告", "")
        End If
        
        If mobjAppendBill Is Nothing Then   '使用混合模式时，不显示嵌入的补附费管理
            If GetInsidePrivs(p医嘱附费管理, True) <> "" And InStr(mstrWorkModule, ";费用记录;") > 0 Then
                Call CreateTabItem(8, C_TAB_NAME_费用记录, 10007, strSelModuleName)
            Else
                mstrWorkModule = Replace(mstrWorkModule, "费用记录", "")
            End If
        End If
        
        If (GetInsidePrivs(p住院医嘱下达, True) <> "" Or GetInsidePrivs(p门诊医嘱下达, True) <> "") _
            And InStr(mstrWorkModule, ";医嘱记录;") > 0 Then
            Call CreateTabItem(9, C_TAB_NAME_医嘱记录, 10010, strSelModuleName)
        Else
            mstrWorkModule = Replace(mstrWorkModule, "医嘱记录", "")
        End If
        
        If (GetInsidePrivs(p住院病历管理, True) <> "" _
            Or GetInsidePrivs(p门诊病历管理, True) <> "" _
            Or GetInsidePrivs(p门诊电子病历, True) <> "" _
            Or GetInsidePrivs(p住院电子病历, True) <> "") _
            And InStr(mstrWorkModule, ";病历记录;") > 0 Then
            Call CreateTabItem(10, C_TAB_NAME_病历记录, 10009, strSelModuleName)
        Else
            mstrWorkModule = Replace(mstrWorkModule, "病历记录", "")
        End If

        
        '添加排队叫号页面
        If mSysPar.blnUseQueue = True Then
            mstrWorkModule = mstrWorkModule & ";排队叫号;"
            
            Call CreateTabItem(11, C_TAB_NAME_排队叫号, 10011, strSelModuleName)
            
'            '快捷叫号界面
'            If mSysPar.blnQueueQuick Then
'                If Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
'                    Call mobjQueue.OpenQueueQuick(GetSelQueueRooms(True), Me)
'                End If
'            End If
        End If
    
'        If Not GetVideoForm Is Nothing Then Call GetVideoForm.ShowVideoWindow(picVideoContainer)
    End With
    
    '如果没有默认tab页，则默认显示第一个tab标签
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
        strSQL = "Select 医嘱ID from 影像检查记录 where 检查UID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询医嘱ID", strCurStudyUID)
        If rsData.RecordCount <= 0 Then Exit Sub
        
        lngAdviceId = Val(NVL(rsData!医嘱ID))
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
'排队接诊事件
On Error GoTo errhandle
    Dim lngIndex As String
    Dim i As Long
    Dim rsTemp As ADODB.Recordset
    
    lngIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("医嘱ID"))
    If lngIndex = -1 Then

        
        If mSysPar.blnSynStudylist Then
            If vsfList.FindRow(lngAdviceId, 1, vsfList.ColIndex("医嘱ID")) > 0 Then Exit Sub
    
            Set rsTemp = mobjPacsQueryWrap.CurPacsQuery.ExecuteWithAttach("[系统.医嘱ID]", lngAdviceId, 1)
            
            If rsTemp.RecordCount > 0 Then
                Call UpdateQueryListData(rsTemp, lngAdviceId, SyncDataType.NoSync)
                '下面的处理用于保证选中第一行并且更新对应病人信息到主界面
            End If
        End If
    End If
    
    If lngIndex > 0 Then
        Call mobjPacsQueryWrap.LocateRow(lngIndex)
        
        If Trim(strTurnPage) <> "" Then
            '跳转到指定的工作模块

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
'排队完成事件
On Error GoTo errhandle
    Dim lngIndex As String
    lngIndex = vsfList.FindRow(lngAdviceId, , vsfList.ColIndex("医嘱ID"))
    
    If lngIndex > 0 Then
        Call mobjPacsQueryWrap.LocateRow(lngIndex)
    End If
    
Exit Sub
errhandle:
    If HintError(err, "mobjQueue_OnCompleted", False) = 1 Then Resume
End Sub

Private Sub mobjQueue_OnSelChange(ByVal lngAdviceId As Long)
'行选择改变事件
On Error GoTo errhandle
    Dim lngIndex As Long
    Dim strCurTabName As String
    
    strCurTabName = mstrSelTabName
    
    '如果当前模块是排队叫号，则排队叫号模块的选择行改变后，不需要触发当前窗口的模块刷新方法,只需刷新列表附加信息显示及pacshelper等
    '如果mstrSelTabName变量为空，则不会刷新模块对象
    If mstrSelTabName = C_TAB_NAME_排队叫号 Then mstrSelTabName = ""
    
    If mSysPar.blnSynStudylist Then
        With vsfList
            lngIndex = .FindRow(lngAdviceId, 1, .ColIndex("医嘱ID"), False, False)
            .Row = lngIndex
            
            '若定位行未出现在可见剪标范围内，则调整可见区域
            If (lngIndex < .TopRow Or lngIndex > .BottomRow) And lngIndex > 0 Then
                .TopRow = lngIndex
            End If
        
            lngIndex = .FindRow(lngAdviceId, 1, .ColIndex("医嘱ID"))
            
            If lngIndex > 0 Then
                Call mobjPacsQueryWrap.LocateRow(lngIndex)
            Else
                HintMsg "检查同步定位失败，请尝试查找。", "mobjQueue_OnSelChange", vbOKOnly
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
    
    strSQL = "Select B.危急状态, A.结果阳性, B.影像质量, B.报告质量, B.符合情况 " & _
             "From 病人医嘱发送 A, 影像检查记录 B " & _
             "Where A.医嘱id = B.医嘱id and B.医嘱ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取结果阳性", lngOrderID)
    
'    If IsNull(rsTemp!危急状态) And mSysPar.lngCriticalValues <> 0 Then strResultInput = "危急状态|"    '不在报告结果窗口中录入危急值
    If IsNull(rsTemp!结果阳性) And Not mSysPar.blnIgnoreResult Then strResultInput = strResultInput & "结果阳性|"
    If IsNull(rsTemp!影像质量) And mSysPar.strImageLevel <> "" And mSysPar.lngImageLevel <> 0 And CheckPopedom(mstrPrivs, "影像质控") Then strResultInput = strResultInput & "影像质量|"
    If IsNull(rsTemp!报告质量) And mSysPar.strReportLevel <> "" And mSysPar.lngReportLevel <> 0 And CheckPopedom(mstrPrivs, "报告质控") Then strResultInput = strResultInput & "报告质量|"
    If IsNull(rsTemp!符合情况) And mSysPar.lngConformDetermine <> 0 Then strResultInput = strResultInput & "符合情况|"

    If strResultInput <> "" Then Call PromptResult(lngOrderID, mlngModule, Me, mlngCur科室ID, strResultInput)
     
Exit Sub
errhandle:
    If HintError(err, "ReportResultHint") = 1 Then Resume
End Sub

Private Sub UpdateStudyListState(lngAdviceId As Long, strStudyUID As String, blnAddImage As Boolean, blnStateChanged As Boolean)
On Error GoTo errH
    Dim strSQL As String   '兼容老板的处理，图像采集状态相关的操作后刷新选中行（数据库刷新）
    Dim intRowIndex As Integer
    Dim rsData As ADODB.Recordset
    Dim lngCol As Long

    With vsfList

        intRowIndex = .FindRow(lngAdviceId, , .ColIndex("医嘱ID"))
        '根据设置更新影像检查技师
        If mSysPar.blnWriteCapDoctor = True And blnStateChanged = True Then
            If mblnCnOracleIsHIS Then
                strSQL = "Zl_影像检查_检查技师( " & lngAdviceId & ",'" & IIf(blnAddImage = True, mstrOtherUserName, "") & "')"
            Else
                strSQL = "Zl_影像检查_检查技师( " & lngAdviceId & ",'" & IIf(blnAddImage = True, mstrHisUserName, "") & "')"
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
'功能：显示当前执行医嘱可以打印的诊疗单据在菜单上
    Dim rsTmp As New ADODB.Recordset
    Dim objControl As CommandBarControl
    Dim strSQL As String
        
    On Error GoTo errH
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "ShowBillList", vbInformation
        Exit Function
    End If
    
    If Not objPopup Is Nothing Then objPopup.CommandBar.Controls.DeleteAll
    
    strSQL = "Select Distinct C.编号,C.名称,C.说明" & _
        " From 病人医嘱记录 A,病历单据应用 B,病历文件列表 C" & _
        " Where A.ID=[1] And A.相关ID IS NULL" & _
        " And A.诊疗项目ID=B.诊疗项目ID" & _
        " And B.应用场合=[2] And B.病历文件ID=C.ID And C.种类=7" & _
        " Order by C.编号"
        
    If mobjCurStudyInfo.blnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
    End If
    
    '诊疗单据项目使用场合需要区分住院，体检，门诊，外来默认为门诊
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjCurStudyInfo.lngAdviceId, CLng(Decode(mobjCurStudyInfo.lngPatientFrom, 3, 1, mobjCurStudyInfo.lngPatientFrom))) 'mobjCurStudyInfo.lngPatientFrom
    
    If Not rsTmp.EOF Then
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Manage_RequestPrint * 10# + 1, rsTmp!名称 & "(&0)")
            objControl.Parameter = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1" '对应的自定义报表编号
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
'功能：打印诊疗单据
On Error GoTo errhandle
    If objControl.Parameter = "" Then '奇怪，直接按F10时，是一个空的Control
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
                       "性质=" & mobjCurStudyInfo.lngRecordKind, "医嘱ID=" & mobjCurStudyInfo.lngAdviceId, 1)
    End If
    Exit Sub
errhandle:
    If HintError(err, "FuncBillPrint", False) = 1 Then Resume
End Sub


Public Sub RefreshList()
'blClick 是否点击刷新触发的刷新列表
'刷新数据列表
    
    If mblnIsLoading = True Then
        HintMsg "数据加载中，请稍后重试...", "RefreshList", vbInformation
        Exit Sub
    End If
    
On Error GoTo errhandle
    mblnIsLoading = True
        
    Call mobjPacsQueryWrap.ExecuteQuery(C_QUERY_刷新)
    If mobjPacsQueryWrap.SqlScheme.AutoRefreshTimeLen > 0 Then TimerRefresh.Enabled = True
    
    If mobjCurStudyInfo.lngAdviceId > 0 Then
'        RefreshDisPlay第三个参数 2表示更新操作
        Call mobjPacsQueryWrap.RefreshDisplay(vsfList.Row, mobjCurStudyInfo.lngAdviceId, 2)
    End If
    
    '直接开始定位
    If vsfList.Rows <= 1 Then
        '当没有数据时，通知刷新工作模块中相关的数据
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
        '重新设置位置
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
'左下方详细信息高度可以改变
On Error GoTo errhandle
    Dim i As Integer
    
    'Y的值 向上移动为负值 向下移动为正值
    If Button = 1 Then
        '当值达到一定范围就退出函数

        If Y > 0 Then
        '向下拖动的判断
        ElseIf Y < 0 Then
        '向上拖动的判断，主要条件：距离列表头部不大于500
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
    HintError err, "cmdLocate_Click<定位操作>", False
End Sub

 
Private Sub picTabFace_Resize()
On Error Resume Next
    Dim i As Long
    Dim lngLen As Long
    
    lngLen = 0
    For i = 0 To TabWindow.ItemCount - 1
        lngLen = lngLen + picTabFace.TextWidth("啊") * Len(TabWindow.Item(i).Caption) + 700
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
    HintError err, "tabScheme_SelectedChanged<方案切换>", False
End Sub

Private Function InsertWorkModuleInfo(ByVal strModuleName As String, ByVal lngHwnd As Long, ByVal lngDeptId As Long, _
    objModule As Object) As TWorkModuleInfo
'插入嵌入的工作模块信息
    Dim lngBound As Long
    Dim i As Long
    
    For i = 1 To UBound(mAryWorkModule)
        If strModuleName = mAryWorkModule(i).ModuleName Then
            '找到模块后，需要使用句柄判断是否对应的模块实例
            If lngHwnd = mAryWorkModule(i).hwnd Then
                InsertWorkModuleInfo = mAryWorkModule(i)
            Else
                '句柄不相同时，则更新模块信息
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
'获取工作模块标记,保存如住院医嘱，门诊医嘱，住院病历，门诊病历，病历报告等关键字
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
        Case C_TAB_NAME_影像图像
            GetWorkModuleName = C_WORKMODULE_NAME_影像图像
            
        Case C_TAB_NAME_影像采集
            GetWorkModuleName = C_WORKMODULE_NAME_影像采集
            
        Case C_TAB_NAME_标本核收
            GetWorkModuleName = C_WORKMODULE_NAME_标本核收
            
        Case C_TAB_NAME_病理取材
            GetWorkModuleName = C_WORKMODULE_NAME_病理取材
            
        Case C_TAB_NAME_病理制片
            GetWorkModuleName = C_WORKMODULE_NAME_病理制片
            
        Case C_TAB_NAME_病理特检
            GetWorkModuleName = C_WORKMODULE_NAME_病理特检
        
        Case C_TAB_NAME_过程报告
            GetWorkModuleName = C_WORKMODULE_NAME_过程报告
            
        Case C_TAB_NAME_检查报告
            lngReportType = Val(GetDeptPara(lngDeptId, "报告编辑器", 0))
            
            If lngReportType = ReportType.PACS报告编辑器 Then
                GetWorkModuleName = C_WORKMODULE_NAME_老版报告
            ElseIf lngReportType = ReportType.电子病历编辑器 Then
                GetWorkModuleName = C_WORKMODULE_NAME_病历报告
            Else
                GetWorkModuleName = C_WORKMODULE_NAME_智能报告
            End If
            
        Case C_TAB_NAME_医嘱记录
            If lngPatientFrom <> 2 Then
                GetWorkModuleName = C_WORKMODULE_NAME_门诊医嘱
            Else
                GetWorkModuleName = C_WORKMODULE_NAME_住院医嘱
            End If
            
        Case C_TAB_NAME_病历记录
            If lngPatientFrom <> 2 Then
                GetWorkModuleName = C_WORKMODULE_NAME_门诊病历
            Else
                GetWorkModuleName = C_WORKMODULE_NAME_住院病历
            End If
            
        Case C_TAB_NAME_电子病历
            GetWorkModuleName = C_WORKMODULE_NAME_电子病历
            
        Case C_TAB_NAME_费用记录
            GetWorkModuleName = C_WORKMODULE_NAME_费用记录
            
        Case C_TAB_NAME_排队叫号
            GetWorkModuleName = C_WORKMODULE_NAME_排队叫号
            
    End Select
End Function

Private Function VerifyModuleObj(ByVal strTabName As String, _
    Optional ByRef objModule As Object, Optional ByRef strModuleTag As String, _
    Optional ByVal blnReInit As Boolean = False) As Long
'验证模块对象
    Dim lngBound As Long
    Dim lngCurDeptId As Long
    Dim strWorkModuleName As String
    
    VerifyModuleObj = 0
    Set objModule = Nothing
    
    '获取当前的科室ID
    lngCurDeptId = GetCurDeptId
    
    strWorkModuleName = GetWorkModuleName(strTabName, lngCurDeptId, GetCurPatientFrom)
    strModuleTag = strWorkModuleName
    
    ';标本核收;影像采集;病理取材;病理制片;病理特检;过程报告;检查报告;医嘱记录;病历记录;费用记录;
    Select Case strTabName
        Case C_TAB_NAME_影像图像
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
            
        Case C_TAB_NAME_影像采集
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
    
        Case C_TAB_NAME_标本核收, C_TAB_NAME_病理取材, C_TAB_NAME_病理制片, C_TAB_NAME_病理特检, C_TAB_NAME_过程报告
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
            
        Case C_TAB_NAME_检查报告
            If mobjWork_Report Is Nothing Then
                Set mobjWork_Report = New clsWorkModule_ReportV2
                
                Call mobjWork_Report.zlInitModule(Me, mlngModule, mstrPrivs, lngCurDeptId, mobjCapLinker, ucPacsHelper1) 'mlngCur科室ID
                
            Else
                If blnReInit Or mobjWork_Report.DeptId <> lngCurDeptId Then
                    Call mobjWork_Report.zlInitModule(Me, mlngModule, mstrPrivs, lngCurDeptId, mobjCapLinker, ucPacsHelper1) 'mlngCur科室ID
                    Call mobjWork_Report.ReInit(strWorkModuleName)
                End If
            End If
            
            If mobjWork_Report Is Nothing Then Exit Function
            
            Set objModule = mobjWork_Report.zlGetForm(strWorkModuleName)
            VerifyModuleObj = objModule.hwnd
            
            
             Call InsertWorkModuleInfo(strWorkModuleName, objModule.hwnd, lngCurDeptId, objModule)
            
        Case C_TAB_NAME_医嘱记录, C_TAB_NAME_病历记录, C_TAB_NAME_电子病历, C_TAB_NAME_费用记录
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
            
        Case C_TAB_NAME_排队叫号
            If mobjQueue Is Nothing Then
                Set mobjQueue = New frmWork_Queue
                Call mobjQueue.zlInitPacsQueueCfg(mlngModule, lngCurDeptId, zlStr.NeedName(mstrCur科室), mstrPrivs, mblnAllDepts, Me)
            Else
                If blnReInit Or mobjQueue.DeptId <> lngCurDeptId Then Call mobjQueue.zlInitPacsQueueCfg(mlngModule, lngCurDeptId, zlStr.NeedName(mstrCur科室), mstrPrivs, mblnAllDepts, Me)
            End If
            
            If mobjQueue Is Nothing Then Exit Function
            
            Set objModule = mobjQueue
            VerifyModuleObj = objModule.hwnd
            
            Call InsertWorkModuleInfo(strWorkModuleName, objModule.hwnd, lngCurDeptId, objModule)
            
    End Select
End Function

Private Function GetPatholModuleType(ByVal strModuleName As String) As TPatholModuleType
'获取病理模块类型
    Select Case strModuleName
        Case C_TAB_NAME_标本核收
            GetPatholModuleType = pmtSpecimen
        Case C_TAB_NAME_病理取材
            GetPatholModuleType = pmtMaterial
        Case C_TAB_NAME_病理制片
            GetPatholModuleType = pmtSlices
        Case C_TAB_NAME_病理特检
            GetPatholModuleType = pmtSpeExam
        Case C_TAB_NAME_过程报告
            GetPatholModuleType = pmtProRep
    End Select
End Function

Private Sub EmbedWindow(ByVal lngHwnd As Long)
'嵌入窗口处理
    SetParent lngHwnd, picWindow.hwnd
    '显示窗口
    ShowWindow lngHwnd, 1
    
    Call MoveWindow(lngHwnd, 0, 0, _
            picWindow.ScaleX(picWindow.Width, vbTwips, vbPixels), _
            picWindow.ScaleY(picWindow.Height, vbTwips, vbPixels), 0)
            
'    SetWindowPos lngHwnd, -1, 0, 0, _
'            picWindow.ScaleX(picWindow.Width, vbTwips, vbPixels), _
'            picWindow.ScaleY(picWindow.Height, vbTwips, vbPixels), 3
         
    BringWindowToTop lngHwnd
    
'    '显示窗口
'    ShowWindow lngHwnd, 1
End Sub

Private Sub AutoHideHelper(ByVal strTabName As String)
    Dim strWorkModuleTag As String
   
   ucPacsHelper1.TabEnable("图像") = True
   ucPacsHelper1.AllowLinkerViewer = True
   
    If strTabName <> C_TAB_NAME_影像采集 Then
        If mlngModule = G_LNG_PACSSTATION_MODULE Then
            '影像医技不嵌入视频采集
            ucPacsHelper1.AllowEmbedVideo = False
            ucPacsHelper1.HideEmbedVideo
        Else
            ucPacsHelper1.AllowEmbedVideo = IIf(Val(GetDeptPara(mlngCur科室ID, "显示视频采集", "0")) <> 0, True, False) 'True
            
            '如果弹出式报告窗口中有视频采集，则主界面不切换嵌入式视频采集
            If Not mobjCapLinker Is Nothing And Not mobjWork_ImageCap Is Nothing And VideoIsAttachReportWindow = False Then
                Call ucPacsHelper1.ShowEmbedVideo(mobjCapLinker, True)
                '恢复之前编辑器所在的焦点
                If strTabName = C_TAB_NAME_检查报告 Then
                    If Not mobjSelModule Is Nothing Then Call mobjSelModule.SetFocus
                End If
            Else
                '隐藏可能之前的视频嵌入式区域
                ucPacsHelper1.HideEmbedVideo
            End If
        End If
        
        ucPacsHelper1.TabEnable("词句") = False
        ucPacsHelper1.AllowWrite = False
            
        If strTabName = C_TAB_NAME_检查报告 Then
            strWorkModuleTag = GetWorkModuleTag(C_TAB_NAME_检查报告)
            
            ucPacsHelper1.TabEnable("词句") = IIf(strWorkModuleTag = C_WORKMODULE_NAME_老版报告, True, False)
            ucPacsHelper1.AllowWrite = IIf(strWorkModuleTag = C_WORKMODULE_NAME_老版报告, True, False)
            
            If ucPacsHelper1.tag = "词句" Then
                Call ucPacsHelper1.LocateTab(ucPacsHelper1.tag)
                If Not mobjSelModule Is Nothing Then Call mobjSelModule.SetFocus
            End If
            
        ElseIf strTabName = C_TAB_NAME_影像图像 Then
            Call ucPacsHelper1.LocateTab("历史")
            ucPacsHelper1.TabEnable("图像") = False
            
        ElseIf strTabName = C_TAB_NAME_排队叫号 Then
            ucPacsHelper1.AllowLinkerViewer = False
            
'        ElseIf strTabName = C_TAB_NAME_过程报告 Then
'            ucPacsHelper1.AllowLinkerViewer = False
            
        End If
    Else
        ucPacsHelper1.AllowLinkerViewer = False
        ucPacsHelper1.AllowEmbedVideo = False
        ucPacsHelper1.HideEmbedVideo
        ucPacsHelper1.TabEnable("词句") = False
    End If
End Sub


Private Function ReloadWorkModule(ByVal strTabName As String, _
    Optional ByRef strModuleTag As String = "", _
    Optional ByRef objSelModule As Object, Optional ByVal blnReInit As Boolean = False) As Long
'重载工作模块
    ';标本核收;影像采集;病理取材;病理制片;病理特检;过程报告;病理诊断;医嘱记录;病历记录;费用记录;
 
    Dim lngSelHwnd As Long
    Dim lngModuleInfoIndex As Long
    
    ReloadWorkModule = 0
    
    lngSelHwnd = VerifyModuleObj(strTabName, objSelModule, strModuleTag, blnReInit)
    
    If lngSelHwnd = 0 Then
        HintMsg "获取[" & strModuleTag & "]模块相关对象失败。", "ReloadWorkModule<重载模块>", vbOKOnly
        Exit Function
    End If
     
    Call EmbedWindow(lngSelHwnd)
    
    '隐藏其他不可见窗口，避免焦点切换问题，如关闭观片窗口后，影像图像可能显示的是检查报告模块
    For lngModuleInfoIndex = 0 To UBound(mAryWorkModule)
        If mAryWorkModule(lngModuleInfoIndex).hwnd <> lngSelHwnd Then ShowWindow mAryWorkModule(lngModuleInfoIndex).hwnd, 0
    Next
    
    lngModuleInfoIndex = GetWorkModuleInfo(strModuleTag)
    
    If mAryWorkModule(lngModuleInfoIndex).FontSize <> gbytFontSize Then
        Call ReSetModuleFontSize(strTabName, strModuleTag, objSelModule, gbytFontSize)
        mAryWorkModule(lngModuleInfoIndex).FontSize = gbytFontSize
    End If
    
    '刷新模块数据
    Call RefreshModuleData(strTabName, strModuleTag, objSelModule)
     
    ''创建菜单
    LockWindowUpdate Me.hwnd
On Error GoTo errhandle
    Call ClearWorkModuleMenu
    Call CreateWorkModuleMenu(strTabName, strModuleTag)
errhandle:
    If err.Number <> 0 Then
        HintError err, "ReloadWorkModule<重载模块>", False
    End If
    LockWindowUpdate 0
    
    ReloadWorkModule = lngSelHwnd
End Function


Private Sub TabWindow_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errhandle
    Dim strModuleTag As String
 
    If Not mblnInitOk Then Exit Sub
     
     If mstrSelTabName <> Item.Caption And mstrSelTabName = C_TAB_NAME_检查报告 And mstrSelModuleTag = C_WORKMODULE_NAME_老版报告 Then
        '弹出报告保存提示
        If Not mobjSelModule Is Nothing Then Call mobjSelModule.PromptSave
     End If
     
    '切换科室时，可能会将tabwindow的部分item.tag属性设置为空，以便切换标签时，可以对界面进行刷新，如检查报告模块
    Call SelectModule(Item.Caption, strModuleTag, IIf(Item.tag = "", True, False))
    
    '保存实际使用的模块名称如住院医嘱，门诊医嘱，住院病历，门诊病历，病历报告等
    Item.tag = strModuleTag
    
    Select Case mstrSelTabName
        Case C_TAB_NAME_排队叫号
            Call RefreshPacsQueueData(True)
 
    
        Case C_TAB_NAME_检查报告
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

'    If Not mobjWork_Report Is Nothing And Item.tag = "报告填写" Then
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
            Call VerifyModuleObj(C_TAB_NAME_影像采集)
            If Not mobjWork_ImageCap Is Nothing Then
                mobjWork_ImageCap.zlRefreshVideoWindow
                Sleep 1000  '暂定一秒
            End If
        End If
    End If
    
    '使用热键进行采集
    If strKeyAlias = mstrCaptureHot Then
        If Not mobjWork_ImageCap Is Nothing Then
            Call mobjWork_ImageCap.zlCaptureImg
        End If

    '使用热键进行后台采集
    ElseIf strKeyAlias = mstrCaptureAfterHot Then
        If Not mobjWork_ImageCap Is Nothing Then
            Call mobjWork_ImageCap.zlCaptureAfterImg
        End If
    
    '使用热键进行标记更新
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
    '刷新病人列表
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
    
'TODO:需要调整
    frmTwoUser.intDBState = mintChangeUserState
    frmTwoUser.strUserNameHIS = mstrHisUserName
    frmTwoUser.strUserIDHIS = mstrHisUserID
    frmTwoUser.Show 1, Me
    
    If frmTwoUser.blnOk = True Then
        If frmTwoUser.intDBState = 1 Then   '统一，则恢复成HIS原来的数据库联接和用户名
            mstrOtherUserName = mstrHisUserName
            mstrOtherUserID = mstrHisUserID
            
            mblnCnOracleIsHIS = True
            mintChangeUserState = 1
            Set gcnOracle = mcnOracleHIS
            
            InitCommon gcnOracle
            
            SetDbUser mstrHisUserID
'            RegCheck
            Call GetUserInfo
'            strPrivs = ";" & GetPrivFunc(100, mlngModule) & ";"      '影像采集工作站
            
'            Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
'            Call mobjWork_Report.zlInitModule(Me, mlngModule, strPrivs, mlngCur科室ID, Nothing, ucPacsHelper1)
        ElseIf frmTwoUser.intDBState = 2 Then   '交换，则交换数据库联接
            '如果是使用新数据库联接，先检查权限
            mstrOtherUserName = frmTwoUser.strUserNameNew
            mstrOtherUserID = frmTwoUser.strUserIDNew
            
            mintChangeUserState = 2
            If frmTwoUser.blnCnOracleIsNew = True Then
                Set gcnOracle = frmTwoUser.cnOracle
                mblnCnOracleIsHIS = False
                
                '初始化zlComLib部件，确保GetPrivFunc提取的是正确的信息
                InitCommon gcnOracle
'                RegCheck
                SetDbUser mstrOtherUserID
                
                '查找用户权限
                strPrivs = GetPrivFunc(100, mlngModule)       '影像采集工作站
                If strPrivs = "" Then
                    HintMsg "你不具备使用“影像采集工作站”模块的权限！", "ChangeUser", vbInformation
                    
                    '切换回原来的用户
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
                
                strPrivs = GetPrivFunc(100, 1258)       '诊疗报告管理
                If strPrivs = "" Then
                    HintMsg "你不具备使用“诊疗报告”模块的权限！", "ChangeUser", vbInformation
                    
                    '切换回原来的用户
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
'            strPrivs = ";" & GetPrivFunc(100, mlngModule) & ";"       '影像采集工作站
            
'            Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
'            Call mobjWork_Report.zlInitModule(Me, mlngModule, strPrivs, mlngCur科室ID, Nothing, ucPacsHelper1)
        End If
        
    End If
    
    If mblnCnOracleIsHIS Then
        Me.stbThis.Panels(4).Text = "报告医生：" & mstrHisUserName & "   检查医生：" & mstrOtherUserName
    Else
        Me.stbThis.Panels(4).Text = "报告医生：" & mstrOtherUserName & "   检查医生：" & mstrHisUserName
    End If
End Sub

Private Sub SwitchUser()
'获取新用户权限说明：使用 GetPrivFuncByUser 并且保证strDBUser参数与gstrDBUser不一样，否则会得到登录用户权限，所以 GetPrivFuncByUser 需要放在SetDbUser 之前
'其中 InitCommon 会执行 SetDbUser
'问题114781改动点：修改判断是否切换成新用户的逻辑，切换用户后增加mstrPrivs赋值操作
    Dim strPrivs As String
 
    Call frmSwitchUser.SetModule(mlngModule)
    frmSwitchUser.Show 1, Me

    If frmSwitchUser.blnOk = False Then Exit Sub
    
'   如果是使用新数据库联接，先检查权限
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
'        If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.zlInitModule(Me, mlngModule, mstrPrivs, mlngCur科室ID, Nothing, ucPacsHelper1)
        
        Me.stbThis.Panels(4).Text = "报告医生：" & mstrOtherUserName & "   检查医生：" & mstrOtherUserName

    End If

End Sub

Private Sub Menu_Manage_随访()
On Error GoTo errhandle
    Dim strReview As String
    Dim strDeptName As String

    If mobjCurStudyInfo.lngAdviceId = 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_随访", vbInformation
        Exit Sub
    End If
    
    strDeptName = Split(mstrCur科室, "-")(1)
    If frmReview.ShowMe(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, Me, strDeptName, strReview) = True Then
            
        mobjCurStudyInfo.strFollowUpDescribe = strReview
        Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
        
    End If

Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_随访", False) Then Resume
End Sub

Private Sub Menu_Manage_报告发放()
'报告发放
On Error GoTo errhandle
    Dim strSQL As String

    If mobjCurStudyInfo.lngAdviceId = 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_报告发放", vbInformation
        Exit Sub
    End If
    
 
    strSQL = "Zl_影像报告发放(" & mobjCurStudyInfo.lngAdviceId & ",'" & UserInfo.姓名 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "报告发放")

    mobjCurStudyInfo.intReportGiveOut = IIf(mobjCurStudyInfo.intReportGiveOut = 1, 0, 1)
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
  
    
    Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_报告发放") Then Resume
End Sub

Private Sub Menu_Manage_胶片发放()
'胶片发放
On Error GoTo errhandle
    Dim strSQL As String

    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_胶片发放", vbInformation
        Exit Sub
    End If
    
    strSQL = "Zl_影像胶片发放(" & mobjCurStudyInfo.lngAdviceId & ",'" & UserInfo.姓名 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "胶片发放")
    
    mobjCurStudyInfo.intFilmGiveOut = IIf(mobjCurStudyInfo.intFilmGiveOut = 1, 0, 1)
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)

    Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_胶片发放") Then Resume
End Sub

Private Sub Menu_Manage_报告胶片同时发放()
'报告胶片同时发放
On Error GoTo errhandle
    Dim strSQL As String
    

    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_报告胶片同时发放", vbInformation
        Exit Sub
    End If
 
    If mobjCurStudyInfo.intReportGiveOut = 1 And mobjCurStudyInfo.intFilmGiveOut = 1 Then
        strSQL = "Zl_影像报告发放(" & mobjCurStudyInfo.lngAdviceId & ",'" & UserInfo.姓名 & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "报告发放")
        
        strSQL = "Zl_影像胶片发放(" & mobjCurStudyInfo.lngAdviceId & ",'" & UserInfo.姓名 & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "胶片发放")
        
        mobjCurStudyInfo.intReportGiveOut = 0
        mobjCurStudyInfo.intFilmGiveOut = 0
    Else
        strSQL = "Zl_影像报告发放(" & mobjCurStudyInfo.lngAdviceId & ",'" & UserInfo.姓名 & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "报告发放")
    
        strSQL = "Zl_影像胶片发放(" & mobjCurStudyInfo.lngAdviceId & ",'" & UserInfo.姓名 & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "胶片发放")
        
        mobjCurStudyInfo.intReportGiveOut = 1
        mobjCurStudyInfo.intFilmGiveOut = 1
        
    End If
    
    Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)

Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_报告胶片同时发放") Then Resume
End Sub

Private Sub Menu_Manage_ReportExecutor()
    Dim strSQL As String
    
    Dim strRPTExecutor As String
On Error GoTo errhandle
    strRPTExecutor = frmSelectRPTExecutor.GetRPTExecutor(mlngCur科室ID, Me, mstrRPTExecutor)
    
    If strRPTExecutor <> "" Then
        '更新报告人
        strSQL = "ZL_影像报告保存_更新报告人(" & mobjCurStudyInfo.lngAdviceId & ",'" & strRPTExecutor & "')"
        Call zlDatabase.ExecuteProcedure(CStr(strSQL), "更新报告人")
        
        '刷新对应检查的报告人
        mstrRPTExecutor = strRPTExecutor
        
        mobjCurStudyInfo.strReportDoctor = strRPTExecutor
        Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
        
        stbThis.Panels(4).Text = "报告医生：" & strRPTExecutor & "   检查医生：" & Split(stbThis.Panels(4).Text, "检查医生：")(1)
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
        strSQL = "Zl_影像检查记录_变更待处理人(" & lngCurAdviceId & ",'" & strName & "')"
        zlDatabase.ExecuteProcedure strSQL, "变更待处理人"
        
        If Len(Trim(strName)) > 0 Then
            HintMsg "成功发送到审核人【" & strName & "】。", "Menu_Manage_SendAudit", vbInformation
        End If
    Else
        HintMsg "请先选择一条检查。", "Menu_Manage_SendAudit", vbInformation
        Exit Function
    End If
    
    Menu_Manage_SendAudit = True
    
    '同步刷新检查列表
    Call UpdateQueryListData(Nothing, lngCurAdviceId)
    
    Exit Function
errhandle:
    If HintError(err, "Menu_Manage_SendAudit") Then Resume
End Function



Private Function GetStudyNumberDisplayName() As String
'获取检查号码显示名称
    GetStudyNumberDisplayName = IIf(mlngModule = G_LNG_PATHOLSYS_NUM, "病理号", "检查号")
End Function

Private Function GetScanRequestCount(ByVal lngAdviceId As Long) As Long
'获取扫描申请单的数量
On Error GoTo errhandle
    Dim lngCount As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    GetScanRequestCount = 0
    
    If lngAdviceId <= 0 Then Exit Function
    
    '如果启用申请单扫描参数 勾选，则在执行查询得到申请单图像数量，未勾选则 不执行
    If mSysPar.blnIsPetitionScan Then
        '根据医嘱ID查询 影像申请单图像表，得到已扫描张数 传入医嘱附项。并处理 VSList
        strSQL = "select count(*) as 图像数 from 影像申请单图像 where 医嘱ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "得到图像数量", lngAdviceId)
        
        lngCount = Val(rsTemp!图像数)
    Else
        lngCount = 0
    End If
    
    GetScanRequestCount = lngCount
Exit Function
errhandle:
    If HintError(err, "GetScanRequestCount") Then Resume
End Function

Private Function GetStudyReportType(objStudyInfo As clsStudyInfo) As Long
'获取当前检查报告类型
    GetStudyReportType = -1
    If objStudyInfo Is Nothing Then Exit Function
    If objStudyInfo.lngExeDepartmentId = mlngCur科室ID Then Exit Function
    
    GetStudyReportType = GetDeptPara(objStudyInfo.lngExeDepartmentId, "报告编辑器", 0)
End Function

Private Sub RefreshModuleData(ByVal strTabName As String, ByVal strWorkModuleTag As String, objSelModule As Object)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：刷新TAB页面
'参数：
'       blnRefresh 完成和取消完成是通知PACS报告编辑器刷新
'''''''''''''''''''''''''''''''''''''''''''''''''
'该过程只针对嵌入式模块和浮动采集模块进行更新，其余如弹出式报告编辑模块则不进行刷新操作
    Dim blnIsForceRrfresh As Boolean

On Error GoTo errhandle
    
    If objSelModule Is Nothing Then Exit Sub
    If mobjCurStudyInfo Is Nothing Then Exit Sub
    
    blnIsForceRrfresh = mblnIsForceRefresh
    
    '如果当前处于历史数据查看状态，则刷新模块数据时，需要进行强制数据刷新，避免历史数据与当前选择的检查数据一致时，模块不执行刷新操作
    If mblnIsHistoryMode Then blnIsForceRrfresh = True
    If mstrSelModuleTag <> strWorkModuleTag Then blnIsForceRrfresh = True   '避免相同模块单元内，由于检查信息对象相同，造成数据模块不能刷新，如clsWorkModule_HisV2包含了费用，医嘱，病历等模块，且共用了mobjstudyinfo对象
 
    Call SetHistoryViewState(False)
    
    If Not mobjWork_Pathol Is Nothing Then
        Call mobjWork_Pathol.zlRefresh(mobjCurStudyInfo, strWorkModuleTag, blnIsForceRrfresh)
    End If
    
    '更新PacsHelper,先对PacsHelper进行更新，以便后续判断是否加载了关联检查图像
    If Not (dkpMain.Panes(2).hidden = False And ucPacsHelper1.Visible = False) Then
        If strTabName <> C_TAB_NAME_检查报告 Then
            ucPacsHelper1.AllowWrite = False
        End If
        
        Call ucPacsHelper1.zlRefresh(mobjCurStudyInfo, 0, blnIsForceRrfresh)
    End If
    
        
    Select Case strTabName
        Case C_TAB_NAME_影像图像
            Call objSelModule.zlRefreshFace(mobjCurStudyInfo, blnIsForceRrfresh)
            
        Case C_TAB_NAME_标本核收, C_TAB_NAME_病理取材, C_TAB_NAME_病理制片, C_TAB_NAME_病理特检, C_TAB_NAME_过程报告
            Call objSelModule.zlUpdateAdviceInf(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngSendNo, mobjCurStudyInfo.intStep, mobjCurStudyInfo.blnMoved)
            Call objSelModule.zlRefreshFace(blnIsForceRrfresh)
            
        Case C_TAB_NAME_检查报告
            Call mobjWork_Report.zlRefreshFace(mobjCurStudyInfo, strWorkModuleTag, blnIsForceRrfresh)
            
        Case C_TAB_NAME_费用记录, C_TAB_NAME_医嘱记录, C_TAB_NAME_病历记录, C_TAB_NAME_电子病历
            Call mobjWork_His.zlModule.zlRefresh(mobjCurStudyInfo, strWorkModuleTag, blnIsForceRrfresh)
            
        Case C_TAB_NAME_影像采集
            '如果当前视频采集是嵌入到弹出式报告编辑窗口时，则不执行后续操作，否则会造成嵌入式视频采集被切换
            If Not mobjCapLinker Is Nothing And Not mobjWork_ImageCap Is Nothing And VideoIsAttachReportWindow = False Then
                Call objSelModule.zlRefreshVideoWindow
                Call objSelModule.zlRestoreWindow(IIf(mobjCurStudyInfo.intStep >= 2 And mobjCurStudyInfo.intStep < 5, False, True), True)
            End If
            
        Case C_TAB_NAME_排队叫号
'            Call EmbedWindow(mobjQueue.hwnd)
            
    End Select
    
    If Not mobjWork_PacsImg Is Nothing Then
        If mobjWork_PacsImg.AdviceId <> mobjCurStudyInfo.lngAdviceId Then Set mobjWork_PacsImg.StudyInfo = mobjCurStudyInfo
    End If
    
    If Not mobjWork_Report Is Nothing Then
        If mobjWork_Report.AdviceId <> mobjCurStudyInfo.lngAdviceId Then Set mobjWork_Report.StudyInfo = mobjCurStudyInfo
    End If

    
    '更新CapLinker对象属性
    If Not mobjCapLinker Is Nothing Then mobjCapLinker.MainAdviceId = mobjCurStudyInfo.lngAdviceId
     
'    '更新PacsHelper
'    If Not (dkpMain.Panes(2).hidden = False And ucPacsHelper1.Visible = False) Then
'        If strTabName <> C_TAB_NAME_检查报告 Then
'            ucPacsHelper1.AllowWrite = False
'        End If
'
'        Call ucPacsHelper1.zlRefresh(mobjCurStudyInfo, 0, blnIsForceRrfresh)
'    End If
    
     
    If mobjCurStudyInfo.lngAdviceId <> 0 Then
        '显示可打印的诊疗单据:之所以即时加载,是为了使用F2热键
        Call ShowBillList(cbrMain.FindControl(, conMenu_Manage_RequestPrint, , True))
    End If
    
    '更新浮动视频模块
    Call ResetFloatingVideoState(mobjCurStudyInfo)
    
    
    Exit Sub
errhandle:
    If HintError(err, "RefreshModuleData", False) = 1 Then Resume
End Sub

'Private Sub ResetFloatingReportState(objStudyInfo As clsStudyInfo)
'    Dim objForm As Object
'
'    '更新弹出报告模块
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
'重设浮动采集窗口状态
    If mobjWork_ImageCap Is Nothing Then Exit Sub
    
    If mobjWork_ImageCap.VideoDockState = False Then Exit Sub
    
    
    If mobjWork_ImageCap.isLock Or mobjWork_ImageCap.IsAfter Then
        Call mobjWork_ImageCap.SetPopupTitle("")
        Exit Sub
    End If
    
    Call mobjWork_ImageCap.SetPopupTitle(objStudyInfo.strPatientName)
    Call mobjWork_ImageCap.zlRestoreWindow(IIf(objStudyInfo.intStep > 1 And objStudyInfo.intStep < 5, False, True), True)
End Sub


Private Sub Menu_Manage_关联病人()
'关联病人
On Error GoTo errhandle
    
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_关联病人", vbInformation
        Exit Sub
    End If
    
    Call frmReferencePatient.ZlShowMe(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.strPatientName, Me, True, mlngCur科室ID)
    
    '刷新病人列表
     Call UpdateQueryListData(Nothing, mobjCurStudyInfo.lngAdviceId)
Exit Sub
errhandle:
    If HintError(err, "Menu_Manage_关联病人", False) = 1 Then Resume
End Sub


Private Sub Menu_Manage_浮动采集()
On Error GoTo errhandle

    If Not GetIsValidOfStorageDevice(mlngCur科室ID) Then
      HintMsg "影像存储设备未定义或处于停用，请检查！", "Menu_Manage_浮动采集", vbInformation
      Exit Sub
    End If
    
    If mobjWork_ImageCap Is Nothing Then
        Call VerifyModuleObj(C_TAB_NAME_影像采集)
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
    If HintError(err, "Menu_Manage_浮动采集", False) = 1 Then Resume
End Sub


Private Sub Menu_Manage_图像刻录()
'图像刻录
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
            HintMsg "不能创建刻录对象，请在安装IMAPI2刻录组件后重新进入。", "Menu_Manage_图像刻录", vbInformation
            Exit Sub
            
continueBurn:
            
            Set frmBurn = New frmImageBurn
        On Error GoTo errFree
            
            lngCurAdviceId = mobjCurStudyInfo.lngAdviceId
            
            Set frmBurn = New frmImageBurn
            Call frmBurn.ShowBurn(mlngModule, mlngCur科室ID, lngCurAdviceId, mobjCurStudyInfo.blnMoved, Me)
errFree:
            Call Unload(frmBurn)
            Set frmBurn = Nothing
    End If
End Sub

Private Sub Menu_Manage_病案查阅()
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_病案查阅", vbInformation
        Exit Sub
    End If
    
    If InStr(";" & GetPrivFunc(100, 1259) & ";", ";基本;") = 0 Then
        HintMsg "您没有查阅电子病历的权限，请联系管理员。", "Menu_Manage_病案查阅", vbInformation
        Exit Sub
    End If
    
    Set mobjMedicalRecord = Nothing
    If mobjMedicalRecord Is Nothing Then
        Set mobjMedicalRecord = DynamicCreate("zlPublicAdvice.clsPublicAdvice", "zlPublicAdvice")
        If mobjMedicalRecord Is Nothing Then Exit Sub
        
        Call mobjMedicalRecord.InitCommon(gcnOracle, glngSys, gstrNodeNo, gfrmMain, glngModul, gstrPrivs, mobjMsgCenter.Msg)
        
        If mobjCurStudyInfo.lngPageID <= 0 Then
            HintMsg "该病人尚未建立病案。", "Menu_Manage_病案查阅", vbInformation
        Else
            Call mobjMedicalRecord.ShowArchive(Me, mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.lngPageID, True)
            
            Set mobjMedicalRecord = Nothing
        End If
    End If
    
End Sub

Private Sub Menu_Manage_收藏管理()
'收藏管理
On Error GoTo errFree
    Dim frmCollectionManage As New frmCollectionManage
    Dim lngCount As Long

    Call frmCollectionManage.ShowCollectionManageWind(Me)
    
    Call ReCreatCbrMenu(cbrMain)
    
errFree:
    Call Unload(frmCollectionManage)
    Set frmCollectionManage = Nothing
End Sub

Private Sub Menu_Manage_收藏到()
'收藏到
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
        HintMsg M_STR_HINT_NoSelectData, "Menu_Manage_收藏到", vbInformation
        Exit Sub
    End If
    
    gstrSQL = "select 首次时间 from 病人医嘱发送 where 医嘱ID= " & lngAdviceId & ""
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    '判断选中记录是否报到，如果没有报到则不能进行收藏操作
    Do While Not rsTemp.EOF
        If NVL(rsTemp!首次时间) = "" Then
            HintMsg "该检查未报到，不能收藏！", "Menu_Manage_收藏到", vbOKOnly
            Exit Sub
        End If
        
        rsTemp.MoveNext
    Loop
    
    Call frmToCollection.ShowToCollectionWind(Me, lngAdviceId, lngSendNo)
    
    Set mobjCurStudyInfo = GetBaseInfo(lngAdviceId, intMovedState + 1)
    
    If mobjCurStudyInfo.lngPatientFrom = 1 Then
        If mobjCurStudyInfo.strMarkNum > 0 Then labCollectionInfo = "门:" & mobjCurStudyInfo.strMarkNum & "  "
    ElseIf mobjCurStudyInfo.lngPatientFrom = 2 Then
        If mobjCurStudyInfo.strMarkNum > 0 Then labCollectionInfo = "住:" & mobjCurStudyInfo.strMarkNum & "  "
    Else
        labCollectionInfo = ""
    End If
    
    labCollectionInfo = labCollectionInfo & mobjCurStudyInfo.strAdviceContext
    labCollectionInfo = labCollectionInfo & IIf(mobjCurStudyInfo.strCollectionInfo = "", "", "  (◆" & mobjCurStudyInfo.strCollectionInfo & ")")
    
errFree:
    Call Unload(frmToCollection)
    Set frmToCollection = Nothing
End Sub


Private Sub Menu_Manage_收藏数据显示(ByVal Control As XtremeCommandBars.ICommandBarControl, ByVal bytStyle As Byte)
'收藏数据显示方法
On Error GoTo errHand
    Dim strCollectionType As String
    Dim lngFatherID As Long
    Dim strLink As String
    
    '处理收藏类别字符串
    If InStr(Control.Caption, "(") = 0 Then
        strCollectionType = Control.Caption
    Else
        strCollectionType = Mid(Control.Caption, 1, InStr(Control.Caption, "(") - 1)
    End If
    
    '处理父级ID字符串
    If bytStyle = 0 Then
        lngFatherID = CLng(Control.ID) - CLng(comMenu_Collection_Type) * 10000#
    ElseIf bytStyle = 1 Then
        lngFatherID = CLng(Control.ID) - CLng(conMenu_Collection_ViewShare) * 10000#
    End If
    
    If Control.Caption = "查看当前收藏" Then
        strLink = " select 医嘱ID from 影像收藏类别 A ,影像收藏内容 B where A.Id=b.收藏Id and A.ID=" & lngFatherID & " union " & _
                        " select 医嘱ID from 影像收藏类别 A ,影像收藏内容 B,影像收藏类别 C where C.Id=b.收藏Id and A.Id=C.上级id  and A.ID=" & lngFatherID & ""
    Else
        strLink = "select 医嘱ID from 影像收藏类别 A ,影像收藏内容 B where A.Id=b.收藏Id and  A.收藏类别='" & strCollectionType & "'"
    End If
    
    Call mobjPacsQueryWrap.ExecuteWithLink(strLink)
    TimerRefresh.Enabled = False
    
    Exit Sub
errHand:
    If HintError(err, "Menu_Manage_收藏数据显示", False) = 1 Then Resume
End Sub

Private Sub Menu_Petition_扫描申请单(ByVal intType As Integer)
'intType:0--查看申请单；1--扫描申请单
    Dim objPetitionCap As frmPetitionCapture                  '申请单
On Error GoTo errFree
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim strPatientDepartment As String
    Dim lngDepID As Long
     
    If mobjCurStudyInfo.lngAdviceId <= 0 Then
        HintMsg M_STR_HINT_NoSelectData, "Menu_Petition_扫描申请单", vbInformation
        Exit Sub
    End If
    
    lngDepID = IIf(mlngCur科室ID = 0, mobjCurStudyInfo.lngExeDepartmentId, mlngCur科室ID)
    With mobjCurStudyInfo
        strSQL = "Select 名称 From 部门表 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取病人科室", .lngPatDept)
        
        strPatientDepartment = ""
        If rsTemp.RecordCount > 0 Then strPatientDepartment = NVL(rsTemp!名称)
    
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
                                    IIf(Not CheckPopedom(mstrPrivs, "检查登记"), True, IIf(intType = 0, True, False)), _
                                    False, _
                                    .lngAdviceId, _
                                    IIf(.strStuStateDesc = "已拒绝", 1, IIf(.strStuStateDesc = "已完成", 2, 0)))
        
        If .lngAdviceId > 0 Then Call UpdateQueryListData(Nothing, .lngAdviceId)
    End With
errFree:
    Unload objPetitionCap
    Set objPetitionCap = Nothing
End Sub

'Private Sub Menu_Manage_SetXWParam_click()
''------------------------------------------------
''功能：打开新网PACS的参数设置窗口
''返回：
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
'功能：发送图像
'返回：
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
'初始化需要自动执行的插件
On Error GoTo errH

    Dim i As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim intExeTime As Integer
    Dim intType As Integer
    Dim strVBS As String

    mintInterfaceCount = 0
    strSQL = "Select a.名称 as 程序名, b.名称 as 功能名 , b.自动执行时机,b.vbs脚本  from 影像插件挂接 a, 影像插件功能 b " & _
             "Where   b.是否启用=1 and  a.是否启用=1 and a.id = b.插件id And (a.所属模块=0 or a.所属模块=[1]) Order By a.id,b.功能序号"
             
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "初始化插件", lngModule)
    
    If rsTemp.RecordCount > 0 Then
        ReDim mintInterface(rsTemp.RecordCount)

        While Not rsTemp.EOF
    
            intExeTime = Val(NVL(rsTemp!自动执行时机))
            
            If intExeTime > 0 Then
                strVBS = NVL(rsTemp!VBS脚本)
                
                mintInterfaceCount = mintInterfaceCount + 1
                mintInterface(mintInterfaceCount).intID = mintInterfaceCount
                mintInterface(mintInterfaceCount).strVBS = strVBS
                mintInterface(mintInterfaceCount).intExeTime = intExeTime
                mintInterface(mintInterfaceCount).strName = NVL(rsTemp!程序名) & "-" & NVL(rsTemp!功能名)
            End If
            
            Call rsTemp.MoveNext
        Wend
    End If
        
    Exit Sub
errH:
    If HintError(err, "initInterface<插件初始化>") = 1 Then Resume
End Sub

Private Sub ExecutePluginInterface(ByVal intTimeType As Integer, Optional ByVal lngTimeTag As Long = 0, _
    Optional ByVal strAttachPar1 As String, Optional ByVal strAttachPar2 As String, Optional ByVal strAttachPar3 As String)
'功能：检查各时机是否有需要自动执行的插件功能
'intTime:执行时机
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
'    MsgBoxD Me, "插件[" & mintInterface(i).strName & "]执行异常。错误信息：" & err.Description, vbInformation, Me.Caption
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
'字号 目前工作站支持9,12,15三种;lngH 高度；lngW 宽度   C_LAYOUT_LISTLEFT
''主界面控件从上到下 mobjFilterCmdBar，mobjFindPati+mobjFindCmd，mobjList，mobjIconPanel，mobjTab
On Error GoTo errH
    Dim lng快速过滤 As Long
    Dim lng查找病人 As Long
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
    
    '这里是大概规定的分割线有效移动范围
    If mlngMove > 6000 Then mlngMove = 6000
    If mlngMove < -4000 Then mlngMove = -4000

    If Not mobjPacsQueryWrap.blShowPatiIdentify Then
        lng查找病人 = 0
    Else
        If gbytFontSize = 15 Then
            lng查找病人 = 400
        Else
            lng查找病人 = 350
        End If
    End If

    If mobjPacsQueryWrap.SqlScheme Is Nothing Then
        lng快速过滤 = 0
    Else
        If Not mobjPacsQueryWrap.SqlScheme.FilterCfgCount > 0 Then
            lng快速过滤 = 0
        Else
            If gbytFontSize = 15 Then
                lng快速过滤 = 550
            ElseIf gbytFontSize = 12 Then
                lng快速过滤 = 450
            Else
                lng快速过滤 = 400
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
    lngList = lngH - lng查找病人 - lng快速过滤 - lngInfo - lngTab
    If lngList < 0 Then lngList = 0
    
    Call tabScheme.Move(0, 0, lngMoreW + C_LAYOUT_LISTLEFT, lngH)
    
    Call picFilter.Move(lngMoreW + C_LAYOUT_LISTLEFT, 0, lngW - lngMoreW, lng快速过滤)
    Call PatiIdentify.Move(lngMoreW + C_LAYOUT_LISTLEFT, picFilter.Top + picFilter.Height, lngW - lngMoreW - 0.5 * C_LAYOUT_LISTLEFT - cmdLocate.Width - cmdFind.Width, lng查找病人)
    
    Call cmdLocate.Move(lngMoreW + PatiIdentify.Width, PatiIdentify.Top, cmdLocate.Width, lng查找病人)
    Call cmdFind.Move(cmdLocate.Left + cmdLocate.Width, PatiIdentify.Top, cmdFind.Width, lng查找病人)
    
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
    
    Call pic主界面遮挡.Move(0, 0, picList.Width, picList.Height)
    Call labNoScheme.Move((picList.Width - labNoScheme.Width) / 2, (picList.Height - labNoScheme.Height) / 2)
    
    Call DoLabFlag
errH:
End Sub

Private Sub initTabExtra()
'初始化界面左下角Tab控件
' 相关控件： TabExtra  picDataSearch（数据检索） picExtra(附加信息)  picFollowUp(随访)  picEvent(事务)
''数据检索 附加信息 历次检查 随访描述 事务提醒 名称固定若要修改注意查询cls 关联修改
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
        
        .InsertItem 1, "数据检索", picDataSearchContainer.hwnd, 0
        .Item(TabExtra.ItemCount - 1).tag = "数据检索"
        
        .InsertItem 2, "附加信息", picExtra.hwnd, 0
        .Item(TabExtra.ItemCount - 1).tag = "附加信息"

        
'        .InsertItem 4, "随访描述", picFollowUp.hWnd, 0
'        .Item(TabExtra.ItemCount - 1).tag = "随访描述"
'
'        .InsertItem 5, "事务提醒", picEvent.hWnd, 0
'        .Item(TabExtra.ItemCount - 1).tag = "事务提醒"
        
        
        strSelect = mobjPacsQueryWrap.GetTabSelectName(False)
        .Item(0).Selected = True
        
        .Width = Screen.Width
        
        For i = 0 To .ItemCount - 1
            If strSelect = .Item(i).tag And .Item(i).Visible Then
                .Item(i).Selected = True
                Exit For
            End If
        Next
        
        '数据检索 附加信息 历次检查 随访描述 事务提醒
    End With
    
End Sub

Public Sub ExecuteDefaultQueryScheme()
'执行自定义查询默认方案
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
            Call InitAutoRefresh 'ExecuteMenu后必须执行
            gblnXWMoved = mobjPacsQueryWrap.CurPacsQuery.IsMoved 'ExecuteMenu后必须执行
            
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
    HintError err, "ExecuteDefaultQueryScheme<执行默认方案>", False
End Sub

Public Sub UpdateQueryListData(ByRef rsData As Recordset, ByVal lngAdviceId As Long, _
    Optional ByVal intSyncDataType As Integer = SyncDataType.rsDataAndrsShow, Optional ByVal blnNoRefreshModule As Boolean = False)
'更新查询列表某一行数据
'同时会更新改行基本数据，注意要先判断更新行是否是当前选中行
'blIsAdd 是否增加了行
'lngAdviceID变化行的医嘱ID
'blRaiseEventSelChange 是否触发列表selchange事件
On Error GoTo errH
    If Not mobjPacsQueryWrap Is Nothing Then Call mobjPacsQueryWrap.UpdateRow(rsData, lngAdviceId, intSyncDataType, blnNoRefreshModule)
    
    Exit Sub
errH:
    HintError err, "UpdateQueryListData<更新列表行>", False
End Sub

Private Sub DoLabFlag()
    Dim lng标识边长 As Long
    Dim test As Boolean
    Dim lngTop间隔 As Long
    
    lngTop间隔 = 30
    lng标识边长 = 270
    
    Call LabFlag急诊.Move(picDetail.Width - lng标识边长, lngTop间隔, lng标识边长, lng标识边长)
    Call LabFlag传染病状态.Move(picDetail.Width - 2 * lng标识边长, lngTop间隔, lng标识边长, lng标识边长)
    Call LabFlag危机状态.Move(picDetail.Width - 3 * lng标识边长, lngTop间隔, lng标识边长, lng标识边长)
    Call LabFlag绿色通道.Move(picDetail.Width - 4 * lng标识边长, lngTop间隔, lng标识边长, lng标识边长)
    Call LabFlag婴儿.Move(picDetail.Width - 5 * lng标识边长, lngTop间隔, lng标识边长, lng标识边长)
    Call LabFlag费用.Move(picDetail.Width - 6 * lng标识边长, lngTop间隔, lng标识边长, lng标识边长)
    
    If mobjCurStudyInfo.lngAdviceId < 1 Then
        LabFlag费用.Visible = False
        LabFlag婴儿.Visible = False
        LabFlag绿色通道.Visible = False
        LabFlag危机状态.Visible = False
        LabFlag传染病状态.Visible = False
        LabFlag急诊.Visible = False
    Else
        If mobjCurStudyInfo.intEmergentTag Then
            LabFlag急诊.Visible = True
        Else
            LabFlag急诊.Visible = False
            Call LabFlag传染病状态.Move(LabFlag传染病状态.Left + lng标识边长)
            Call LabFlag危机状态.Move(LabFlag危机状态.Left + lng标识边长)
            Call LabFlag绿色通道.Move(LabFlag绿色通道.Left + lng标识边长)
            Call LabFlag婴儿.Move(LabFlag婴儿.Left + lng标识边长)
            Call LabFlag费用.Move(LabFlag费用.Left + lng标识边长)
        End If
    
        If mobjCurStudyInfo.blnIsInfectious Then
            LabFlag传染病状态.Visible = True
        Else
            LabFlag传染病状态.Visible = False
            Call LabFlag危机状态.Move(LabFlag危机状态.Left + lng标识边长)
            Call LabFlag绿色通道.Move(LabFlag绿色通道.Left + lng标识边长)
            Call LabFlag婴儿.Move(LabFlag婴儿.Left + lng标识边长)
            Call LabFlag费用.Move(LabFlag费用.Left + lng标识边长)
        End If
        
        If mobjCurStudyInfo.intDangerState = 1 Then
            LabFlag危机状态.Visible = True
        Else
            LabFlag危机状态.Visible = False
            Call LabFlag绿色通道.Move(LabFlag绿色通道.Left + lng标识边长)
            Call LabFlag婴儿.Move(LabFlag婴儿.Left + lng标识边长)
            Call LabFlag费用.Move(LabFlag费用.Left + lng标识边长)
        End If
        
        If mobjCurStudyInfo.intGreenChannel = 1 Then
            LabFlag绿色通道.Visible = True
        Else
            LabFlag绿色通道.Visible = False
            Call LabFlag婴儿.Move(LabFlag婴儿.Left + lng标识边长)
            Call LabFlag费用.Move(LabFlag费用.Left + lng标识边长)
        End If

        If mobjCurStudyInfo.lngBaby > 0 Then
            LabFlag婴儿.Visible = True
        Else
            LabFlag婴儿.Visible = False
            Call LabFlag费用.Move(LabFlag费用.Left + lng标识边长)
        End If
        
        Select Case mobjCurStudyInfo.lngMoneyState
            Case ChargeState.未收费
                LabFlag费用.Caption = "欠"
'                LabFlag费用.ForeColor = &H80FF&
            Case ChargeState.已收费
                LabFlag费用.Caption = "收"
'                LabFlag费用.ForeColor = &H8000&
            Case ChargeState.无费用
                LabFlag费用.Caption = "无"
'                LabFlag费用.ForeColor = &HC00000
            Case ChargeState.已补缴
                LabFlag费用.Caption = "补"
'                LabFlag费用.ForeColor = &HFF&
            Case ChargeState.已记账
                LabFlag费用.Caption = "记"
'                LabFlag费用.ForeColor = &HFF00FF
            Case ChargeState.已退费
                LabFlag费用.Caption = "退"
'                LabFlag费用.ForeColor = &H80000011
            Case ChargeState.已销账
                LabFlag费用.Caption = "销"
'                LabFlag费用.ForeColor = &H8080&
            Case ChargeState.已调整
                LabFlag费用.Caption = "调"
'                LabFlag费用.ForeColor = &H94
        End Select
        LabFlag费用.Visible = True

    End If
End Sub

Private Sub TimFlicker_Timer()
On Error GoTo errH
'   超时闪烁的处理
    Dim i As Integer, j As Integer
    Dim lngCol As Long, lngColContrast As Long
    Dim strTmp As String
    Dim lngStateColor As Long, lngNextStateColor As Long, lngPreStateColor As Long
    Dim objRowRelation As Object
    
    Static intsta As Integer
    Static TPFlickerInfo As TFlickerInfo '超时闪烁配置
    
    '方案第一次加载时获取超时闪烁相关信息
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
    For i = vsfList.TopRow To vsfList.BottomRow   '遍历可见行  For 1
        For j = 0 To UBound(Split(TPFlickerInfo.strInfo, "|")) - 1 '判断是否满足超时条件 For 2
            strTmp = Split(TPFlickerInfo.strInfo, "|")(j)
            If Split(strTmp, ",")(0) = vsfList.TextMatrix(i, lngCol) Then
                lngColContrast = vsfList.ColIndex(Split(strTmp, ",")(1))
                
                If IsDate(vsfList.TextMatrix(i, lngColContrast)) Then
                
                    If DateDiff("N", vsfList.TextMatrix(i, lngColContrast), Now) >= Val(Split(strTmp, ",")(2)) Then    '若满足设置的超时时间
                    
                        '首先测试闪烁功能
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
                
                Exit For   '若满足超时条件 退出For 2
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
'历史数据联动查看
    Dim objStudyInfo As clsStudyInfo
    Dim strCurModuleTag As String
    
On Error GoTo errhandle
    If mobjSelModule Is Nothing Then Exit Sub
    
    If lngAdviceId <> 0 And blnIsDBClick Then
        Select Case mstrSelTabName
            Case C_TAB_NAME_影像图像
                If Not mobjWork_PacsImg Is Nothing Then
                    If mobjWork_PacsImg.AdviceId = lngAdviceId Then lngAdviceId = 0
                End If
            
            Case C_TAB_NAME_标本核收, C_TAB_NAME_病理取材, C_TAB_NAME_病理制片, C_TAB_NAME_病理特检, C_TAB_NAME_过程报告
                If Not mobjWork_Pathol Is Nothing Then
                    If mobjWork_Pathol.AdviceId = lngAdviceId Then lngAdviceId = 0
                End If
                
            Case C_TAB_NAME_检查报告
                If Not mobjWork_Report Is Nothing Then
                    If mobjWork_Report.AdviceId = lngAdviceId Then lngAdviceId = 0
                End If
                
            Case C_TAB_NAME_费用记录, C_TAB_NAME_医嘱记录, C_TAB_NAME_病历记录, C_TAB_NAME_电子病历
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
        Case C_TAB_NAME_影像图像
            If Not mobjWork_PacsImg Is Nothing Then
                Call mobjWork_PacsImg.zlRefreshFace(objStudyInfo, mblnIsForceRefresh, IIf(lngAdviceId = 0, False, True))
            End If
            
        Case C_TAB_NAME_标本核收, C_TAB_NAME_病理取材, C_TAB_NAME_病理制片, C_TAB_NAME_病理特检, C_TAB_NAME_过程报告
            If Not mobjWork_Pathol Is Nothing Then
                Call mobjWork_Pathol.zlRefresh(objStudyInfo, mstrSelTabName, mblnIsForceRefresh, IIf(lngAdviceId = 0, False, True))
            End If
            
'            Call mobjSelModule.zlUpdateAdviceInf(objStudyInfo.lngAdviceId, objStudyInfo.lngSendNo, objStudyInfo.intStep, objStudyInfo.blnMoved)
'            Call mobjSelModule.zlRefreshFace(mblnIsForceRefresh)
            
        Case C_TAB_NAME_检查报告
            '判断报告类型是否与当前相同
            strCurModuleTag = GetWorkModuleName(mstrSelTabName, objStudyInfo.lngExeDepartmentId, objStudyInfo.lngPatientFrom)
            If strCurModuleTag <> mstrSelModuleTag Then
               Call SelectModule(mstrSelTabName, strCurModuleTag)
               TabWindow.Selected.tag = strCurModuleTag
            End If
    
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(objStudyInfo, mstrSelModuleTag, mblnIsForceRefresh, IIf(lngAdviceId = 0, False, True))
            End If
            
        Case C_TAB_NAME_费用记录, C_TAB_NAME_医嘱记录, C_TAB_NAME_病历记录, C_TAB_NAME_电子病历
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
    
    '不能改变caption的内容，因为部分功能需要根据caption进行判断
'    TabWindow.Selected.Caption = Replace(TabWindow.Selected.Caption, C_HISTORY_VIEW_TAG, "") & IIf(blnIsHistory, C_HISTORY_VIEW_TAG, "")
End Sub


Private Sub ucPacsHelper1_OnTabChanged(ByVal strTabName As String)
On Error GoTo errhandle
    '判断是否需要恢复到词句模块页显示，如果tag为"词句"，则切换到检查报告模块时，需要恢复到词句子模块页
    If TabWindow.Selected.Caption = "检查报告" Then
        If strTabName = "词句" Then
            ucPacsHelper1.tag = "词句"
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
        HintMsg "有尚未完成的事务，请等待完成后重试操作。", "vsfList_BeforeSelChange", vbOKOnly
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
'根据是否已经有方案决定主界面提示信息
'目前有：正常，无方案，无启用方案三种情况。
    
    If blLoadFail Then
        pic主界面遮挡.Visible = True
            labNoScheme.Visible = True
            Call pic主界面遮挡.Move(picList.Left, picList.Top, picList.Width, picList.Height)
            
            If Trim(strHint) <> "" Then
                labNoScheme.Caption = "查询方案加载错误：" & vbLf & strHint
            Else
                labNoScheme.Caption = "查询方案加载错误，请联系软件技术人员"
            End If
    Else
        If mintQueryState = 1 Then
            pic主界面遮挡.Visible = False
            labNoScheme.Visible = False
        Else
            pic主界面遮挡.Visible = True
            labNoScheme.Visible = True
            Call pic主界面遮挡.Move(picList.Left, picList.Top, picList.Width, picList.Height)
            
            If mintQueryState = 2 Then
                labNoScheme.Caption = "没有有效查询方案，请先配置"
            ElseIf mintQueryState = 3 Then
                labNoScheme.Caption = "没有启用方案"
            Else
                labNoScheme.Caption = "查询方案加载错误，请联系软件技术人员"
            End If
        End If
    End If
    
    Call picList_Resize
End Sub

Private Sub vsfList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    timFun.Enabled = True
End Sub


Private Sub CreateAuditorMenu(objControl As CommandBarControl)
'创建审核人菜单
On Error GoTo errH
    Dim cbrPopControl As CommandBarControl
    Dim rsTemp As Recordset
    Dim strSQL As String
    Dim i As Long
    
    If Not objControl Is Nothing Then
        objControl.CommandBar.Controls.DeleteAll
    End If
    
    If mblnAllDepts Then
        strSQL = "Select Distinct a.Id, a.姓名" & vbNewLine & _
            "From 人员表 a, 部门人员 b, 部门性质说明 c" & vbNewLine & _
            "Where a.Id = b.人员id And b.部门id = c.部门id And c.工作性质 = '检查'"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取有审核报告资格的医生")
    Else
        strSQL = "select A.id,A.姓名 from 人员表 A,部门人员 B where B.部门ID=[1] AND A.ID=B.人员ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取有审核报告资格的医生", mlngCur科室ID)
    End If
    
    If rsTemp.RecordCount < 1 Then Exit Sub
    For i = 1 To rsTemp.RecordCount
        If GetUserSignLevel(rsTemp!ID) >= cprSL_主治 Then
            Set cbrPopControl = CreateMenu(objControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_SendAudit * 10# + i, rsTemp!姓名, "", 0, False)
        End If
        rsTemp.MoveNext
    Next
    Exit Sub
errH:
    If HintError(err, "CreateAuditorMenu") = 1 Then Resume
End Sub

Private Sub Menu_Manage_检查预约()
'------------------------------------------------
'功能：打开检查预约窗口
'参数：无
'返回：无
'------------------------------------------------
    On Error GoTo err
    Dim i As Integer
    Dim strIds As String
    Dim lngID() As Long
    Dim blnCheckin As Boolean
    
    blnCheckin = True
    strIds = frmSchSchedule.ZlShowMe(mstrPrivs, mobjCurStudyInfo.lngAdviceId, IIf(mlngCur科室ID = 0, mstrCanUse科室IDs, mlngCur科室ID), Me, blnCheckin)
    If strIds = "" Then Exit Sub
    
    If blnCheckin = True Then
        Call Menu_Manage_报到
    End If
    
    '处理返回值
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
    Call HintError(err, "Menu_Manage_检查预约", False)
End Sub

Private Sub Menu_Manage_预约管理()
'------------------------------------------------
'功能：打开预约管理窗口
'参数：无
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    frmSchManage.ZlShowMe mstrPrivs, IIf(mlngCur科室ID = 0, mstrCanUse科室IDs, mlngCur科室ID), mobjCurStudyInfo.lngAdviceId, Me
    
    Exit Sub
err:
    Call HintError(err, "Menu_Manage_预约管理", False)
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
            If CheckPopedom(mstrPrivs, "所有科室") Then
                strSQL = "select 名称,执行间 from 医技执行房间 a, 部门表 b where a.科室Id=b.Id and instr([1],b.ID)>0 "
                
                strID = mstrCanUse科室IDs
            Else
                '查询对应人员所在科室中所包含的执行间
                strSQL = "select 名称,执行间 from 医技执行房间 a, 部门人员 b,部门表 c where a.科室id=b.部门id and a.科室Id=c.Id and b.人员id = [1]"
                
                strID = UserInfo.ID
            End If
                    
        Else
            strSQL = "Select 名称,执行间 From 医技执行房间 a, 部门表 b Where a.科室Id=b.Id and  科室ID=[1]"
            
            strID = mlngCur科室ID
            
        End If
        
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strID)
        
        
        While rsData.EOF = False
        
            If mstrSelQueueRooms <> "" Then mstrSelQueueRooms = mstrSelQueueRooms & ","
            mstrSelQueueRooms = mstrSelQueueRooms & NVL(rsData!名称) & "-" & NVL(rsData!执行间)
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
'处理自动刷新，必须在切换方案或者执行默认方案后执行
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
    
    GetBaseInfo.lngReportType = Val(GetDeptPara(GetBaseInfo.lngExeDepartmentId, "报告编辑器", 0)) + 1
End Function

Private Sub QueueDataConsistency(ByVal lngAdviceId As Long, ByVal strRoom As String, ByVal intRowIndex As Integer)
'排队数据一致性处理，主要是执行间数据
On Error GoTo errH
    Dim lngSendNo As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    '排队数据一致性处理：判断记录集中是否存在，若存在，更近记录集数据（UpdateSourceData）。
     '判断列表是否已经显示，若显示，则更新列表数据（"执行间"）
     '都需要更新数据库数据，其中发送好来源可能是记录集数据，也可能从数据库中单独查询。
     
    '如果执行间数据没有变化，终止处理
    If intRowIndex > -1 Then
        If mobjPacsQueryWrap.Text(intRowIndex, "执行间") = strRoom Then
            Exit Sub
        End If
    End If

    strSQL = "select 发送号 from 病人医嘱发送 Where 医嘱ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获得发送号", lngAdviceId)
    If rsTemp.RecordCount = 1 Then
        lngSendNo = Val(NVL(rsTemp!发送号))
    End If
        
    Call UpdateQueryListData(Nothing, lngAdviceId)
    
    '更新数据库数据
    strSQL = "ZL_影像检查记录_发送安排(" & lngAdviceId & "," & lngSendNo & ",null,null,null,'" & strRoom & "',1)"
    Call zlDatabase.ExecuteProcedure(strSQL, "更新执行间")
    
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
    HintError err, "ReCreatCbrMenu<重置菜单>", False
End Sub

Private Sub VsfListDbClick(ByVal blnIsLocate As Boolean)
On Error GoTo errhandle
    
    If Not blnIsLocate Then
        If vsfList.MouseRow = -1 Or vsfList.MouseRow = 0 Then Exit Sub
    End If
    
    If mobjCurStudyInfo.lngAdviceId <> 0 Then
        '双击病人检查列表时，如果病人检查状态为 已拒绝，目前不做任何处理
        If mobjCurStudyInfo.strStuStateDesc = "已拒绝" Then Exit Sub
        
        Select Case mobjCurStudyInfo.intStep
            Case 1, 0
                Call Menu_Manage_报到
            Case 2, 3               '双击打开书写报告,报告打开时跟据设定是否打开观片站
                Call Menu_RichEPR(conMenu_PacsReport_Write)
            Case -1, 4, 5               '双击修订报告,报告打开时跟据设定是否打开观片站
                Call Menu_RichEPR(conMenu_Edit_Audit)
            Case 6                  '查阅
                Call Menu_RichEPR(conMenu_File_Open)
        End Select
    End If

Exit Sub
errhandle:
    Call HintError(err, "VsfListDbClick", False)
End Sub


Private Function Is_ExistReportWriting(ByVal lngAdviceId As Long) As Boolean
'是否有报告处于修订状态
On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "Select e.Id,  -null as 对象标记, l.最后版本 As 版本, '正在修订…' As 操作, l.保存人 As 人员" & vbNewLine & _
            "From 电子病历记录 l," & vbNewLine & _
            "    (Select Max(c.开始版) As 开始版, Max(Id + 1) As Id,Max(对象标记+1) as 对象标记" & vbNewLine & _
            "     From 电子病历内容 c ,病人医嘱报告 d" & vbNewLine & _
            "     Where c.文件id = d.病历id  And c.对象类型 = 8 and d.医嘱id=[1]) e ,病人医嘱报告 f" & vbNewLine & _
            "Where L.ID =f.病历id  And L.最后版本 > e.开始版 and f.医嘱id=[1]"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断是否存在修订中的报告", lngAdviceId)
    Is_ExistReportWriting = rsTemp.RecordCount > 0
            
    Exit Function
errH:
    err.Raise -1, , "判断是否存在修订中的报告异常" & vbCrLf & err.Description
End Function

Private Sub ChangeScheme(ByVal strName As String, ByVal lngID As Long, ByVal blnMenuClick As Boolean)
'blnMenuClick 是否菜单点击触发（被动触发指）true: 菜单栏点击触发  false：左侧tab点击触发
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
        Call InitAutoRefresh 'ExecuteMenu后必须执行
        gblnXWMoved = mobjPacsQueryWrap.CurPacsQuery.IsMoved 'ExecuteMenu后必须执行
        
        Call CheckHaveScheme(False, strResult)
        
        dkpMain.FindPane(1).title = strName
        
        If blnMenuClick Then Call mobjPacsQueryWrap.RefreshTabLeft(tabScheme, dkpMain.FindPane(1).title)
        
        Call AdjustFace(picList.Height, picList.Width)
        Call picDataSearchContainer_Resize
        Call ReSetFormFontSize
    End If
    Exit Sub
errH:
    err.Raise -1, , "ChangeScheme异常" & vbCrLf & err.Description
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
'根据科室ID更新一个全局参数
    gblUseImgSignValid = False
    If Len(GetSetting("ZLSOFT", "公共模块\ZL9PACSWork", "启用图像签名验证")) > 0 Then
        gblUseImgSignValid = GetSetting("ZLSOFT", "公共模块\ZL9PACSWork", "启用图像签名验证", "0") = "1" And GetSignVerifyType(lngID) = 1
    Else
        gblUseImgSignValid = GetDeptPara(lngID, "图像签名验证") = "1" And GetSignVerifyType(lngID) = 1
    End If
errH:
End Sub

Private Sub ViewLinkChecks()
'查看关联检查，根据当前病人ID 查到所有的关联医嘱ID,然后使用 ExecuteWithLink(strLink)
'1 全部病人ID的检查信息
'2 如果使用了关联检查，还有关联检查的医嘱ID
On Error GoTo errH
    Dim strSQL As String
    Dim rsData As Recordset
    Dim strAppend As String
    Dim strLink As String
    Dim i As Long
    
    If mobjCurStudyInfo Is Nothing Then Exit Sub
    
    If mSysPar.blnRelatingPatient And mobjCurStudyInfo.lngLinkId > 0 Then
        strLink = "Select A.ID as 医嘱ID From 病人医嘱记录 A Where A.病人id = " & mobjCurStudyInfo.lngPatId & " UNION ALL Select 医嘱ID  from 影像检查记录 Where 关联ID =" & mobjCurStudyInfo.lngLinkId & ""
    Else
        strLink = "Select A.ID as 医嘱ID From 病人医嘱记录 A Where A.病人id = " & mobjCurStudyInfo.lngPatId & ""
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
