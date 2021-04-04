VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmScheduleEdit 
   Caption         =   "体检预约申请"
   ClientHeight    =   7500
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   11580
   Icon            =   "frmScheduleEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ils32 
      Left            =   10365
      Top             =   4305
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleEdit.frx":076A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   6060
      Top             =   6420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleEdit.frx":6FCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleEdit.frx":C036
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleEdit.frx":C330
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleEdit.frx":C8CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleEdit.frx":CE64
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleEdit.frx":CFBE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraGroup 
      Caption         =   "&1.组别"
      Height          =   2280
      Left            =   225
      TabIndex        =   36
      Top             =   2865
      Width           =   2445
      Begin MSComctlLib.ListView lvwGroup 
         Height          =   1380
         Left            =   90
         TabIndex        =   37
         Top             =   300
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   2434
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         Icons           =   "ils32"
         SmallIcons      =   "ils32"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   3
         Left            =   1920
         Picture         =   "frmScheduleEdit.frx":DE10
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "快捷键：F10"
         Top             =   1770
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   2
         Left            =   1545
         Picture         =   "frmScheduleEdit.frx":14662
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "快捷键：F9"
         Top             =   1770
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   0
         Left            =   1170
         Picture         =   "frmScheduleEdit.frx":1AEB4
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "快捷键：F8"
         Top             =   1770
         Width           =   345
      End
   End
   Begin VB.PictureBox picButton 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   435
      ScaleHeight     =   555
      ScaleWidth      =   10650
      TabIndex        =   66
      Top             =   6375
      Width           =   10650
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   8100
         TabIndex        =   55
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   9315
         TabIndex        =   56
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   90
         TabIndex        =   57
         Top             =   135
         Width           =   1100
      End
   End
   Begin VB.Frame fraTitle 
      Height          =   645
      Left            =   -15
      TabIndex        =   60
      Top             =   -90
      Width           =   10635
      Begin VB.PictureBox picNo 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   8745
         ScaleHeight     =   315
         ScaleWidth      =   1815
         TabIndex        =   62
         Top             =   240
         Width           =   1815
         Begin VB.TextBox txt体检号 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
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
            Left            =   825
            Locked          =   -1  'True
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   30
            Width           =   945
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Index           =   37
            Left            =   30
            TabIndex        =   64
            Top             =   45
            Width           =   360
         End
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "团体体检预约申请"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   105
         TabIndex        =   61
         Top             =   255
         Width           =   2040
      End
   End
   Begin VB.Frame fraGroupInfo 
      Height          =   660
      Left            =   75
      TabIndex        =   58
      Top             =   1455
      Width           =   11190
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   10
         Left            =   10350
         Picture         =   "frmScheduleEdit.frx":21706
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "快捷键：F12"
         Top             =   225
         Width           =   345
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   8
         Left            =   8400
         TabIndex        =   28
         Text            =   "cfr@zlsoft.cn"
         Top             =   240
         Width           =   1800
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   11
         Left            =   6000
         TabIndex        =   26
         Text            =   "1399090980"
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   12
         Left            =   4395
         TabIndex        =   24
         Text            =   "空了吹"
         Top             =   240
         Width           =   810
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   13
         Left            =   780
         TabIndex        =   21
         Text            =   "某某市无名有限责任公司"
         Top             =   240
         Width           =   2595
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   4
         Left            =   3405
         Picture         =   "frmScheduleEdit.frx":22548
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   210
         Width           =   345
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "电子邮件"
         Height          =   180
         Index           =   15
         Left            =   7605
         TabIndex        =   27
         Top             =   300
         Width           =   780
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgNew 
         Height          =   240
         Index           =   0
         Left            =   480
         Picture         =   "frmScheduleEdit.frx":22AD2
         Top             =   90
         Width           =   240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人"
         Height          =   180
         Index           =   14
         Left            =   3810
         TabIndex        =   23
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系电话"
         Height          =   180
         Index           =   12
         Left            =   5220
         TabIndex        =   25
         Top             =   315
         Width           =   720
      End
      Begin VB.Label lbl民族 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "团体(&N)"
         Height          =   180
         Left            =   90
         TabIndex        =   20
         Top             =   285
         Width           =   630
      End
   End
   Begin VB.Frame fraSingle 
      Height          =   990
      Left            =   75
      TabIndex        =   0
      Top             =   465
      Width           =   11805
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   14
         Left            =   5010
         TabIndex        =   7
         Text            =   "90"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   14
         Left            =   10890
         Picture         =   "frmScheduleEdit.frx":2305C
         Style           =   1  'Graphical
         TabIndex        =   73
         TabStop         =   0   'False
         ToolTipText     =   "将信息写回IC卡"
         Top             =   225
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   15
         Left            =   10500
         Picture         =   "frmScheduleEdit.frx":298AE
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         ToolTipText     =   "从IC卡读信息"
         Top             =   225
         Width           =   345
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   10
         Left            =   5190
         TabIndex        =   11
         Text            =   "1399090980"
         Top             =   615
         Width           =   1515
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   1
         Left            =   2205
         Picture         =   "frmScheduleEdit.frx":30100
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   225
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   11
         Left            =   11280
         Picture         =   "frmScheduleEdit.frx":3068A
         Style           =   1  'Graphical
         TabIndex        =   68
         TabStop         =   0   'False
         ToolTipText     =   "快捷键：F11"
         Top             =   210
         Width           =   345
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   3
         Left            =   3225
         TabIndex        =   5
         Text            =   "90"
         Top             =   240
         Width           =   1140
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   7
         Left            =   7740
         TabIndex        =   19
         Text            =   "cfr@zlsoft.cn"
         Top             =   600
         Width           =   1800
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   4
         Left            =   7545
         TabIndex        =   9
         Text            =   "123456789012345678901"
         Top             =   240
         Width           =   1995
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   9
         Left            =   2205
         TabIndex        =   15
         Text            =   "90"
         Top             =   615
         Width           =   510
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   5
         Left            =   750
         TabIndex        =   2
         Text            =   "李某某"
         Top             =   255
         Width           =   1425
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   3255
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   615
         Width           =   1110
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   750
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   630
         Width           =   945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "健康号"
         Height          =   180
         Index           =   16
         Left            =   4425
         TabIndex        =   6
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "联系电话"
         Height          =   180
         Index           =   4
         Left            =   4410
         TabIndex        =   10
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   180
         Index           =   0
         Left            =   2640
         TabIndex        =   4
         Top             =   315
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "电子邮件"
         Height          =   180
         Index           =   10
         Left            =   6960
         TabIndex        =   18
         Top             =   690
         Width           =   720
      End
      Begin VB.Image imgNew 
         Height          =   240
         Index           =   1
         Left            =   465
         Picture         =   "frmScheduleEdit.frx":314CC
         Top             =   105
         Width           =   240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   180
         Index           =   7
         Left            =   6765
         TabIndex        =   8
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻"
         Height          =   180
         Index           =   6
         Left            =   2745
         TabIndex        =   16
         Top             =   690
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Index           =   11
         Left            =   1800
         TabIndex        =   14
         Top             =   675
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名(&N)"
         Height          =   180
         Index           =   8
         Left            =   90
         TabIndex        =   1
         Top             =   315
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性  别"
         Height          =   180
         Index           =   9
         Left            =   90
         TabIndex        =   12
         Top             =   690
         Width           =   540
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   615
      Left            =   285
      TabIndex        =   59
      Top             =   2115
      Width           =   10635
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   5655
         TabIndex        =   35
         Text            =   "某某市无名有限责任公司"
         Top             =   210
         Width           =   4890
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   3075
         TabIndex        =   33
         Top             =   210
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1170
         TabIndex        =   31
         Top             =   210
         Width           =   1140
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "地址(&A)"
         Height          =   180
         Index           =   3
         Left            =   4965
         TabIndex        =   34
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "电话(&T)"
         Height          =   180
         Index           =   2
         Left            =   2400
         TabIndex        =   32
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "预 约 人(&L)"
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   30
         Top             =   270
         Width           =   990
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   67
      Top             =   7140
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmScheduleEdit.frx":31A56
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15346
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComDlg.CommonDialog dlg 
      Left            =   2160
      Top             =   6825
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab tbs 
      Height          =   2580
      Left            =   3345
      TabIndex        =   41
      Top             =   2730
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   4551
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   635
      WordWrap        =   0   'False
      TabCaption(0)   =   "&4.体检项目"
      TabPicture(0)   =   "frmScheduleEdit.frx":322EA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(17)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(18)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl(19)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "vsfPrice"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmd(6)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "vsf"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmd(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmd(18)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmd(17)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtSum(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtSum(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtSum(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "&5.受检人员"
      TabPicture(1)   =   "frmScheduleEdit.frx":32306
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vsfPerson"
      Tab(1).Control(1)=   "cmd(8)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmd(9)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmd(13)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmd(12)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmd(16)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   0
         Left            =   1155
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   435
         Width           =   930
      End
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   1
         Left            =   3135
         MaxLength       =   16
         TabIndex        =   78
         Top             =   435
         Width           =   870
      End
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   2
         Left            =   4680
         MaxLength       =   16
         TabIndex        =   77
         Top             =   435
         Width           =   1020
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   17
         Left            =   4620
         Picture         =   "frmScheduleEdit.frx":32322
         Style           =   1  'Graphical
         TabIndex        =   76
         TabStop         =   0   'False
         ToolTipText     =   "全部记帐"
         Top             =   720
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   18
         Left            =   4230
         Picture         =   "frmScheduleEdit.frx":38B74
         Style           =   1  'Graphical
         TabIndex        =   75
         TabStop         =   0   'False
         ToolTipText     =   "全部收费"
         Top             =   720
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   16
         Left            =   -70035
         Picture         =   "frmScheduleEdit.frx":3F3C6
         Style           =   1  'Graphical
         TabIndex        =   74
         TabStop         =   0   'False
         ToolTipText     =   "单位人员选择"
         Top             =   1275
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   12
         Left            =   -70665
         Picture         =   "frmScheduleEdit.frx":45C18
         Style           =   1  'Graphical
         TabIndex        =   71
         TabStop         =   0   'False
         ToolTipText     =   "将信息写回IC卡"
         Top             =   795
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   13
         Left            =   -70545
         Picture         =   "frmScheduleEdit.frx":4C46A
         Style           =   1  'Graphical
         TabIndex        =   70
         TabStop         =   0   'False
         ToolTipText     =   "从IC卡读信息"
         Top             =   1260
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   9
         Left            =   -69630
         Picture         =   "frmScheduleEdit.frx":52CBC
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "导入人员(F7)"
         Top             =   825
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   8
         Left            =   -70080
         Picture         =   "frmScheduleEdit.frx":57D16
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "更多资料(F6)"
         Top             =   810
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   5
         Left            =   5025
         Picture         =   "frmScheduleEdit.frx":58B58
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "多选，快捷键：F3"
         Top             =   705
         Width           =   345
      End
      Begin zl9Medical.VsfGrid vsfPerson 
         Height          =   1545
         Left            =   -74655
         TabIndex        =   45
         Top             =   450
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   2725
      End
      Begin zl9Medical.VsfGrid vsf 
         Height          =   1530
         Left            =   165
         TabIndex        =   42
         Top             =   795
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   2699
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   6
         Left            =   5475
         Picture         =   "frmScheduleEdit.frx":590E2
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "体检类型选择：F4"
         Top             =   720
         Width           =   345
      End
      Begin zl9Medical.VsfGrid vsfPrice 
         Height          =   1635
         Left            =   2985
         TabIndex        =   69
         Top             =   855
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   2884
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "基本价格(&B)"
         Height          =   180
         Index           =   19
         Left            =   150
         TabIndex        =   82
         Top             =   495
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检价格(E)"
         Height          =   180
         Index           =   18
         Left            =   2130
         TabIndex        =   81
         Top             =   495
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "折扣(Z)"
         Height          =   180
         Index           =   17
         Left            =   4035
         TabIndex        =   80
         Top             =   495
         Width           =   630
      End
   End
   Begin VB.Frame fraOther 
      Height          =   570
      Left            =   0
      TabIndex        =   65
      Top             =   5475
      Width           =   10635
      Begin VB.TextBox txt 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   300
         Index           =   31
         Left            =   4455
         TabIndex        =   52
         Top             =   180
         Width           =   480
      End
      Begin VB.CheckBox chk 
         Caption         =   "需要随访(&X)"
         Height          =   195
         Left            =   2475
         TabIndex        =   50
         Top             =   240
         Width           =   1290
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   6
         Left            =   6045
         TabIndex        =   54
         Top             =   180
         Width           =   4335
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   870
         TabIndex        =   49
         Top             =   180
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   91291651
         CurrentDate     =   38545
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检时间"
         Height          =   180
         Index           =   5
         Left            =   75
         TabIndex        =   48
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "期限(&Y)      月"
         Height          =   180
         Index           =   29
         Left            =   3810
         TabIndex        =   51
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "附注(&L)"
         Height          =   180
         Index           =   13
         Left            =   5325
         TabIndex        =   53
         Top             =   240
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmScheduleEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnGroup As Boolean
Private mlngDept As Long
Private mrsItems As New ADODB.Recordset                 '用于暂时保存选择的体检项目
Private mrsPersons As New ADODB.Recordset                 '用于暂时体检人员
Private mrsGroup As New ADODB.Recordset                 '用于暂时体检团体
Private mlngLoop As Long
Private mblnDataChange As Boolean
Private mbytMode As Byte                        '标志,
Private mstrGroup As String
Private mstrSQL As String

Private Enum mCol
    项目 = 1
    执行科室
    检查部位
    采集方式
    采集科室
    检验标本
    基本价格
    体检价格
    折扣
    体检类型
    类别
    结算方式
    执行科室id
    采集方式id
    采集科室id
    检查部位id
    计费明细
    新加
    前景色
    删除
    公共
    
    p计价项目 = 1
    p名称
    p计算单位
    p数次
    p标准单价
    p体检单价
    p折扣
    p标准金额
    p体检金额
    p执行科室
    p执行科室id
    p收费项目id
    p计价性质
    p类别
    p可用库存
End Enum

Private Enum mPersonCol
    姓名 = 1
    门诊号
    健康号
    性别
    年龄
    婚姻状况
    出生日期
    身份证
    民族
    国籍
    学历
    职业
    身份
    联系人姓名
    联系人电话
    电子邮件
    联系人地址
    工作单位
    登记时间
    病人id
    IC卡号
    就诊卡号
    前景色
    
End Enum

Private Enum mColChar
    
    姓名 = 66
    性别
    年龄
    出生日期
    婚姻状况
    身份证号
    门诊号
    健康号
    就诊卡号
    工作单位
    电子邮件
    民族
    学历
    职业
    国籍
    体检组
    
End Enum

'（２）自定义过程或函数************************************************************************************************

Private Property Let DataChange(ByVal vData As Boolean)
        mblnDataChange = vData
End Property

Private Property Get DataChange() As Boolean
        DataChange = mblnDataChange
End Property

Private Function CountGroup() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:按组别统计项目数量、人数（男、女）
    '参数:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strTmp As String
    Dim lngCount1 As Long
    Dim lngCount2 As Long
    
    If mblnGroup Then
        strTmp = """" & lvwGroup.SelectedItem.Text & """组别下"
    End If
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            If vsf.TextMatrix(lngLoop, mCol.类别) = "检查" Then
                lngCount1 = lngCount1 + 1
            Else
                lngCount2 = lngCount2 + 1
            End If
        End If
    Next
    
    strTmp = strTmp & "共有项目" & lngCount1 + lngCount2 & "个(检查:" & lngCount1 & "个,检验:" & lngCount2 & "个)"
    
    If mblnGroup Then
        lngCount1 = 0
        lngCount2 = 0
        
        For lngLoop = 1 To vsfPerson.Rows - 1
            If Trim(vsfPerson.TextMatrix(lngLoop, mPersonCol.姓名)) <> "" Then
                If InStr(vsfPerson.TextMatrix(lngLoop, mPersonCol.性别), "男") > 0 Then
                    lngCount1 = lngCount1 + 1
                Else
                    lngCount2 = lngCount2 + 1
                End If
            End If
        Next
        
        strTmp = strTmp & ";共有人员" & lngCount1 + lngCount2 & "个(男性:" & lngCount1 & "个,女性:" & lngCount2 & "个)"
    End If
    
    stbThis.Panels(2).Text = strTmp
    
End Function

Private Function ChangeTotal(ByVal dbMoney As Double, ByVal dbTmp As Double, Optional ByVal bytMode As Byte = 1) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim db折扣 As Double
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngRow As Long
    Dim dbTotal As Double
    
    If dbMoney = 0 Then Exit Function
    
    If bytMode = 1 Then
        '变化金额
        
        '1.计算折扣
        db折扣 = Format(10 * dbTmp / dbMoney, "0.0000")

    Else
        '变化折扣
        db折扣 = dbTmp

    End If
    
    txtSum(1).Text = Format(dbMoney * db折扣 / 10, "0.00")
    txtSum(2).Text = Format(db折扣, "0.0000")
    dbTotal = 0
    
    For lngLoop = 1 To vsf.Rows - 1
    
        vsf.TextMatrix(lngLoop, mCol.折扣) = db折扣
        vsf.TextMatrix(lngLoop, mCol.体检价格) = Format(Val(vsf.TextMatrix(lngLoop, mCol.基本价格)) * (db折扣 / 10), "0.00")
        
        dbTotal = dbTotal + Val(vsf.TextMatrix(lngLoop, mCol.体检价格))
                    
        varRow = Split(vsf.TextMatrix(lngLoop, mCol.计费明细), ";")
        For lngRow = 0 To UBound(varRow)
            If varRow(lngRow) <> "" Then
                varCol = Split(varRow(lngRow), ":")
                varCol(4) = Format(Val(varCol(3)) * (db折扣 / 10), "0.00000")
                varCol(10) = db折扣
            End If
            varRow(lngRow) = Join(varCol, ":")
        Next
        vsf.TextMatrix(lngLoop, mCol.计费明细) = Join(varRow, ";")
    Next

    '误差处理
    '------------------------------------------------------------------------------------------------------------------
    If dbTotal <> Val(txtSum(1).Text) Then

        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.TextMatrix(lngLoop, mCol.体检价格)) <> 0 Then
            
                vsf.TextMatrix(lngLoop, mCol.体检价格) = Val(vsf.TextMatrix(lngLoop, mCol.体检价格)) + (Val(txtSum(1).Text) - dbTotal)
                
                If Val(vsf.TextMatrix(lngLoop, mCol.基本价格)) <> 0 Then
                    vsf.TextMatrix(lngLoop, mCol.折扣) = Format(10 * Val(vsf.TextMatrix(lngLoop, mCol.体检价格)) / Val(vsf.TextMatrix(lngLoop, mCol.基本价格)), "0.0000")
                Else
                    vsf.TextMatrix(lngLoop, mCol.折扣) = 0
                End If
                
                varRow = Split(vsf.TextMatrix(lngLoop, mCol.计费明细), ";")
                For lngRow = 0 To UBound(varRow)
                    If varRow(lngRow) <> "" Then
                        varCol = Split(varRow(lngRow), ":")
                        If Val(varCol(4)) <> 0 Then
                            varCol(4) = Val(varCol(4)) + (Val(txtSum(1).Text) - dbTotal)
                            If Val(varCol(3)) <> 0 Then
                                varCol(10) = Format(10 * Val(varCol(4)) / Val(varCol(3)), "0.0000")
                            Else
                                varCol(10) = 0
                            End If
                        End If
                    End If
                    varRow(lngRow) = Join(varCol, ":")
                Next
                vsf.TextMatrix(lngLoop, mCol.计费明细) = Join(varRow, ";")
                Exit For
            End If
        Next
    End If

    ChangeTotal = True
    
End Function

Private Function ChangeItem(ByVal dbMoney As Double, ByVal dbTmp As Double, Optional ByVal bytMode As Byte = 1, Optional ByVal blnUpdate As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim dbSum As Double
    Dim db折扣 As Double
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngRow As Long
    
    If blnUpdate Then
        If dbMoney = 0 Then Exit Function
        
        Call WritePrice(vsf.Row)
        
        If bytMode = 1 Then
            '变化金额
            
            '1.计算折扣
            db折扣 = Format(10 * dbTmp / dbMoney, "0.0000")
        Else
            '变化折扣
            db折扣 = dbTmp
            
        End If
        
        vsf.TextMatrix(vsf.Row, mCol.体检价格) = Format(dbMoney * db折扣 / 10, "0.00")
        vsf.TextMatrix(vsf.Row, mCol.折扣) = Format(db折扣, "0.0000")
    End If
    
    '更新总体
    '------------------------------------------------------------------------------------------------------------------
    dbSum = 0
    For lngLoop = 1 To vsf.Rows - 1
       dbSum = dbSum + Val(vsf.TextMatrix(lngLoop, mCol.基本价格))
    Next
    txtSum(0).Text = Format(dbSum, "0.00")
    
    dbSum = 0
    For lngLoop = 1 To vsf.Rows - 1
       dbSum = dbSum + Val(vsf.TextMatrix(lngLoop, mCol.体检价格))
    Next
    txtSum(1).Text = Format(dbSum, "0.00")
    
    If Val(txtSum(0).Text) <> 0 Then
        txtSum(2).Text = Format(10 * Val(txtSum(1).Text) / Val(txtSum(0).Text), "0.0000")
    Else
        txtSum(2).Text = "0.0000"
    End If
    
    '更新价格
    '------------------------------------------------------------------------------------------------------------------
    If blnUpdate Then
        varRow = Split(vsf.TextMatrix(vsf.Row, mCol.计费明细), ";")
        For lngRow = 0 To UBound(varRow)
            If varRow(lngRow) <> "" Then
                varCol = Split(varRow(lngRow), ":")
                varCol(4) = Format(Val(varCol(3)) * (db折扣 / 10), "0.00000")
                varCol(10) = db折扣
            End If
            varRow(lngRow) = Join(varCol, ":")
        Next
        vsf.TextMatrix(vsf.Row, mCol.计费明细) = Join(varRow, ";")
    End If
        
    ChangeItem = True
    
End Function

Private Function ChangePrice(ByVal dbMoney As Double, ByVal dbTmp As Double, Optional ByVal bytMode As Byte = 1) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim dbSum As Double
    Dim db折扣 As Double
    
    If dbMoney = 0 Then Exit Function
    
    If bytMode = 1 Then
        '变化金额
        
        '1.计算折扣
        db折扣 = Format(10 * dbTmp / dbMoney, "0.0000")
    Else
        '变化折扣
        db折扣 = dbTmp
        
    End If
    
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.p体检单价) = Format(dbMoney * db折扣 / 10, "0.00000")
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.p折扣) = Format(db折扣, "0.0000")
    
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.p体检金额) = Val(vsfPrice.TextMatrix(vsfPrice.Row, mCol.p数次)) * Val(vsfPrice.TextMatrix(vsfPrice.Row, mCol.p体检单价))
    
    '更新项目
    '------------------------------------------------------------------------------------------------------------------
    dbSum = 0
    For lngLoop = 1 To vsfPrice.Rows - 1
       dbSum = dbSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.p标准金额))
    Next
    vsf.TextMatrix(vsf.Row, mCol.基本价格) = dbSum
    
    dbSum = 0
    For lngLoop = 1 To vsfPrice.Rows - 1
       dbSum = dbSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.p体检金额))
    Next
    vsf.TextMatrix(vsf.Row, mCol.体检价格) = dbSum
    
    If Val(vsf.TextMatrix(vsf.Row, mCol.基本价格)) <> 0 Then
        vsf.TextMatrix(vsf.Row, mCol.折扣) = Format(10 * Val(vsf.TextMatrix(vsf.Row, mCol.体检价格)) / Val(vsf.TextMatrix(vsf.Row, mCol.基本价格)), "0.0000")
    Else
        vsf.TextMatrix(vsf.Row, mCol.折扣) = "0.0000"
    End If
    
    '更新总体
    '------------------------------------------------------------------------------------------------------------------
    dbSum = 0
    For lngLoop = 1 To vsf.Rows - 1
       dbSum = dbSum + Val(vsf.TextMatrix(lngLoop, mCol.基本价格))
    Next
    txtSum(0).Text = Format(dbSum, "0.00")
    
    dbSum = 0
    For lngLoop = 1 To vsf.Rows - 1
       dbSum = dbSum + Val(vsf.TextMatrix(lngLoop, mCol.体检价格))
    Next
    txtSum(1).Text = Format(dbSum, "0.00")
    
    If Val(txtSum(0).Text) <> 0 Then
        txtSum(2).Text = Format(10 * Val(txtSum(1).Text) / Val(txtSum(0).Text), "0.0000")
    Else
        txtSum(2).Text = "0.0000"
    End If
        
    ChangePrice = True
    
End Function

Private Function SumPrice(ByVal bytMode As Byte) As Single
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim sglSum As Single
    
    For lngLoop = 1 To vsfPrice.Rows - 1
        If bytMode = 2 Then
            sglSum = sglSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.p体检金额))
        Else
            sglSum = sglSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.p标准金额))
        End If
    Next
    SumPrice = sglSum
    
End Function

Private Function GetPatientInfo(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    strSQL = "SELECT A.* FROM 病人信息 A WHERE A.病人id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        
        If mblnGroup Then
                If Val(cmd(4).Tag) <> Val(zlCommFun.NVL(rs("合同单位id"))) And Val(zlCommFun.NVL(rs("合同单位id"))) > 0 And Val(cmd(4).Tag) > 0 Then
                    
                    If MsgBox("不是当前团体的人员，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    
                End If
                
                vsfPerson.EditText = zlCommFun.NVL(rs("姓名"))
                vsfPerson.Cell(flexcpData, vsfPerson.Row, vsfPerson.Col) = zlCommFun.NVL(rs("姓名").Value)
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.姓名) = zlCommFun.NVL(rs("姓名"))
                
                Call SetRowDefault(0, vsfPerson.Row, "缺省信息")
                
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.门诊号) = zlCommFun.NVL(rs("门诊号"))
                
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份证) = zlCommFun.NVL(rs("身份证号"))
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.年龄) = zlCommFun.NVL(rs("年龄"))
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.出生日期) = Format(zlCommFun.NVL(rs("出生日期")), "yyyy-MM-dd")
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.性别) = zlCommFun.NVL(rs("性别").Value)
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.婚姻状况) = zlCommFun.NVL(rs("婚姻状况").Value)
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.病人id) = zlCommFun.NVL(rs("病人id"))
                
                Call FillPatient(Val(zlCommFun.NVL(rs("病人id"))), 2)
                
                vsfPerson.EditMode(mPersonCol.门诊号) = 0
        Else
                    cmd(1).Tag = zlCommFun.NVL(rs("病人id").Value)
                    txt(5).Text = zlCommFun.NVL(rs("姓名").Value)
                    txt(4).Text = zlCommFun.NVL(rs("身份证号").Value)
                    txt(9).Text = zlCommFun.NVL(rs("年龄").Value)
                    
                    txt(3).Text = zlCommFun.NVL(rs("门诊号").Value)
                    
                    zlControl.CboLocate cbo(1), zlCommFun.NVL(rs("性别").Value)
                    zlControl.CboLocate cbo(0), zlCommFun.NVL(rs("婚姻状况").Value)
                    
                    Call FillPatient(Val(cmd(1).Tag))
                    
                    
                    txt(5).Tag = ""
                    imgNew(1).Visible = False
                    
                    txt(3).Locked = (Val(txt(3).Text) > 0 And Val(cmd(1).Tag) > 0)
        End If
        
        DataChange = True
        
    End If
    
    GetPatientInfo = True
    
End Function


Private Function CreatePriceList(ByVal intRow As Integer) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim strKeys As String
    
    strKeys = CStr(Val(vsf.RowData(intRow))) & "'" & CStr(Val(vsf.TextMatrix(intRow, mCol.采集方式id))) & "'" & vsf.TextMatrix(intRow, mCol.检查部位id)
    
    Dim str计价项目 As String
    Dim str计价性质 As String
    
    vsfPrice.Rows = 2
    str计价项目 = vsfPrice.TextMatrix(1, mCol.p计价项目)
    str计价性质 = vsfPrice.TextMatrix(1, mCol.p计价性质)
    
    vsfPrice.Body.Cell(flexcpText, 1, mCol.p计价项目 + 1, 1, vsfPrice.Cols - 1) = ""
    vsfPrice.RowData(1) = 0

    vsfPrice.TextMatrix(1, mCol.p计价项目) = str计价项目
    vsfPrice.TextMatrix(1, mCol.p计价性质) = str计价性质
    
    mstrSQL = GetPublicSQL(SQL.体检项目价表, strKeys)
    
    If vsf.TextMatrix(intRow, mCol.检查部位id) = "" Then
        '检验或单部位检查
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, Val(vsf.RowData(intRow)), Val(vsf.TextMatrix(intRow, mCol.采集方式id)))
    Else
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    End If
    
    If rs.BOF = False Then
        With vsfPrice
            Do While Not rs.EOF
                
                If Val(.TextMatrix(.Rows - 1, mCol.p收费项目id)) > 0 Then
                    .Rows = .Rows + 1
                End If
                
                If zlCommFun.NVL(rs("计价性质")) = 2 Then
                    .TextMatrix(.Rows - 1, mCol.p计价项目) = "采集方式-" & vsf.TextMatrix(vsf.Row, mCol.采集方式)
                ElseIf vsf.TextMatrix(vsf.Row, mCol.类别) = "检验" Then
                    .TextMatrix(.Rows - 1, mCol.p计价项目) = "检验项目-" & vsf.TextMatrix(vsf.Row, mCol.项目)
                Else
                    .TextMatrix(.Rows - 1, mCol.p计价项目) = "检查项目-" & vsf.TextMatrix(vsf.Row, mCol.项目)
                End If
                
                .TextMatrix(.Rows - 1, mCol.p名称) = zlCommFun.NVL(rs("名称"))
                .TextMatrix(.Rows - 1, mCol.p计算单位) = zlCommFun.NVL(rs("计算单位"))
                .TextMatrix(.Rows - 1, mCol.p数次) = zlCommFun.NVL(rs("收费数量"))
                .TextMatrix(.Rows - 1, mCol.p标准单价) = zlCommFun.NVL(rs("现价"))
                .TextMatrix(.Rows - 1, mCol.p体检单价) = zlCommFun.NVL(rs("现价"))
                .TextMatrix(.Rows - 1, mCol.p折扣) = 10
                .TextMatrix(.Rows - 1, mCol.p标准金额) = zlCommFun.NVL(rs("收费数量"), 0) * zlCommFun.NVL(rs("现价"), 0)
                .TextMatrix(.Rows - 1, mCol.p体检金额) = zlCommFun.NVL(rs("收费数量"), 0) * zlCommFun.NVL(rs("现价"), 0)
                .TextMatrix(.Rows - 1, mCol.p收费项目id) = zlCommFun.NVL(rs("ID"))
                
                .TextMatrix(.Rows - 1, mCol.p计价性质) = zlCommFun.NVL(rs("计价性质"))
                .TextMatrix(.Rows - 1, mCol.p类别) = zlCommFun.NVL(rs("类别"))
                
                Call SetRowDefault(zlCommFun.NVL(rs("ID"), 0), .Rows - 1, "收费执行科室")
                
                If InStr("567", .TextMatrix(.Rows - 1, mCol.p类别)) > 0 Then
                    .TextMatrix(.Rows - 1, mCol.p可用库存) = GetStorage(Val(.RowData(.Rows - 1)), Val(.TextMatrix(.Rows - 1, mCol.p执行科室id)))
                    Call PromptStorageWarn(Val(.TextMatrix(.Rows - 1, mCol.p数次)), Val(.TextMatrix(.Rows - 1, mCol.p可用库存)), .TextMatrix(.Rows - 1, mCol.p名称), .TextMatrix(.Rows - 1, mCol.p执行科室), .TextMatrix(.Rows - 1, mCol.p计算单位), 1)
                End If
                
                rs.MoveNext
            Loop
        End With
        
    End If
    
    vsf.TextMatrix(intRow, mCol.基本价格) = SumPrice(1)
    vsf.TextMatrix(intRow, mCol.体检价格) = SumPrice(2)
    
End Function

Public Function ShowEdit(ByVal frmMain As Object, ByVal lngKey As Long, ByVal lngDept As Long, Optional blnGroup As Boolean = False, Optional ByVal bytMode As Byte = 1) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:  frmMain         调用窗体对象
    '       lngKey          预约登记id
    '返回:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    
    mblnOK = False
    mlngKey = lngKey
    mblnGroup = blnGroup
    mlngDept = lngDept
    mbytMode = bytMode
    
    Set mfrmMain = frmMain
    
    Call ClearData
    If InitData = False Then Exit Function
    
    If mlngKey > 0 Then
    
        imgNew(0).Visible = False
        imgNew(1).Visible = False
        
        If ReadData(mlngKey) = False Then Exit Function
'        stbThis.Panels(2).Text = "修改体检预约。"
    Else
        If mblnGroup Then
            Call ReadGroup(0)
        End If
        
        imgNew(0).Visible = True
        imgNew(1).Visible = True
        
'        stbThis.Panels(2).Text = "新开体检预约。"
    End If
    
    Call CountGroup
    
    DataChange = False
    txt(5).Tag = ""
    txt(13).Tag = ""
            
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ClearData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    On Error Resume Next
    
    For lngLoop = 0 To txt.UBound
        txt(lngLoop).Text = ""
        txt(lngLoop).Tag = ""
    Next
    
    On Error GoTo 0
    
    lvwGroup.ListItems.Clear
    Call ResetVsf(vsf)
    Call ResetVsf(vsfPrice)
    Call ResetVsf(vsfPerson)
    
    DataChange = False
    
        
End Function

Private Function InitMaxLength() As Boolean
    
    '设置最大输入长度
    txt(5).MaxLength = GetMaxLength("病人信息", "姓名")
    txt(4).MaxLength = GetMaxLength("病人信息", "身份证号")
    txt(13).MaxLength = GetMaxLength("合约单位", "名称")
    txt(12).MaxLength = GetMaxLength("合约单位", "联系人")
    txt(11).MaxLength = GetMaxLength("合约单位", "联系电话")
    txt(0).MaxLength = GetMaxLength("体检登记记录", "联系人")
    txt(1).MaxLength = GetMaxLength("体检登记记录", "联系电话")
    txt(2).MaxLength = GetMaxLength("体检登记记录", "联系地址")
    txt(6).MaxLength = GetMaxLength("体检登记记录", "附加说明")
    
    txt(10).MaxLength = GetMaxLength("病人信息", "联系人电话")
    
    InitMaxLength = True
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化设置
    '返回:  True        初始化成功
    '       False       初始化失败
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strVsf As String
    
    On Error GoTo errHand
    
    mstrGroup = ""
    
    With vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "名称", 2100, 1, "...", 1
        .NewColumn "执行科室", 1080, 1, " ", 1
        
        .NewColumn "检查部位", 1800, 1, "...", 1
        .NewColumn "采集方式", 1200, 1, " ", 1
        .NewColumn "采集科室", 1080, 1, " ", 1
        
        .NewColumn "检验标本", 900, 1, " ", 1
        .NewColumn "基本价格", 900, 7
        .NewColumn "体检价格", 900, 7, , 1
        .NewColumn "折扣", 900, 7, , 1
        .NewColumn "体检类型", 0, 1
        .NewColumn "类别", 0, 1
        .NewColumn "结算方式", 900, 1, "记帐|收费", 1
        .NewColumn "执行科室id", 0, 1
        .NewColumn "采集方式id", 0, 1
        .NewColumn "采集科室id", 0, 1
        .NewColumn "检查部位id", 0, 1
        .NewColumn "计费明细", 0, 1
        .NewColumn "新加", 0, 1
        .NewColumn "前景色", 0, 1
        .NewColumn "删除", 0, 1
        .NewColumn "公共", 0, 1
        .FixedCols = 1
        
        .SelectMode = True
        
        .Body.ColFormat(mCol.基本价格) = "0.00"
        .Body.ColFormat(mCol.体检价格) = "0.00"
        .Body.ColFormat(mCol.折扣) = "0.0000"
    End With
    
    With vsfPrice
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "计价项目", 2100, 1, " ", 1
        .NewColumn "收费项目", 2700, 1, "...", 1
        .NewColumn "单位", 600, 1
        .NewColumn "数次", 540, 7, , 1
        .NewColumn "标准单价", 900, 7
        .NewColumn "体检单价", 900, 7, , 1
        .NewColumn "折扣", 900, 7, , 1
        .NewColumn "标准价格", 900, 7
        .NewColumn "体检价格", 900, 7
        .NewColumn "执行科室", 1080, 1, " ", 1
        .NewColumn "执行科室id", 0
        .NewColumn "收费项目id", 0
        .NewColumn "计价性质", 0
        .NewColumn "类别", 0
        .NewColumn "", 0
        .FixedCols = 1
        .Body.ColFormat(mCol.p标准单价) = "0.00000"
        .Body.ColFormat(mCol.p体检单价) = "0.00000"
        .Body.ColFormat(mCol.p标准金额) = "0.00"
        .Body.ColFormat(mCol.p体检金额) = "0.00"
        .Body.ColFormat(mCol.p折扣) = "0.0000"
        .SelectMode = True
    End With
    
    With vsfPerson
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "姓名", 990, 1, "...", 1, GetMaxLength("病人信息", "姓名")
        .NewColumn "门诊号", 810, 1
        .NewColumn "健康号", 810, 1, , 1, GetMaxLength("病人信息", "健康号")
        .NewColumn "性别", 750, 1, GetCombList("SELECT 名称 FROM 性别"), 1, GetMaxLength("病人信息", "性别")
        .NewColumn "年龄", 540, 1, , 1, GetMaxLength("病人信息", "年龄")
        .NewColumn "婚姻状况", 900, 1, GetCombList("SELECT 名称 FROM 婚姻状况"), 1, GetMaxLength("病人信息", "婚姻状况")
        .NewColumn "出生日期", 990, 1, , 1
        .NewColumn "身份证", 1800, 1, , 1, GetMaxLength("病人信息", "身份证号")
                
        .NewColumn "民族", 0, 1, , , GetMaxLength("病人信息", "民族")
        .NewColumn "国籍", 0, 1, , , GetMaxLength("病人信息", "国籍")
        .NewColumn "学历", 0, 1, , , GetMaxLength("病人信息", "学历")
        .NewColumn "职业", 0, 1, , , GetMaxLength("病人信息", "职业")
        .NewColumn "身份", 0, 1, , , GetMaxLength("病人信息", "身份")
        .NewColumn "联系人姓名", 0, 1, , , GetMaxLength("病人信息", "联系人姓名")
        .NewColumn "联系人电话", 0, 1, , , GetMaxLength("病人信息", "联系人电话")
        .NewColumn "电子邮件", 0, 1, , , GetMaxLength("病人信息", "电子邮件")
        .NewColumn "联系人地址", 0, 1, , , GetMaxLength("病人信息", "联系人地址")
        .NewColumn "工作单位", 0, 1, , , GetMaxLength("病人信息", "工作单位")
        .NewColumn "登记时间", 0, 1
        .NewColumn "病人id", 0, 1
        .NewColumn "IC卡号", 0, 1
        .NewColumn "就诊卡号", 0, 1
        
        .FixedCols = 1
        .SelectMode = True
        .Body.ColEditMask(mPersonCol.出生日期) = "0000-00-00"
    End With
       
       
    gstrSQL = "SELECT 编码||'-'||名称 AS 名称,0 AS ID,缺省标志 FROM 性别 ORDER BY 编码"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(1), rs)
    
    gstrSQL = "SELECT 编码||'-'||名称 AS 名称,0 AS ID,缺省标志 FROM 婚姻状况 ORDER BY 编码"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(0), rs)
    
    '设置最大输入长度
    Call InitMaxLength
    
    '团体和个人体检
    fraGroupInfo.Visible = False
    fraSingle.Visible = False
    
    If mblnGroup Then
    
        lblTitle.Caption = "团体体检" & IIf(mbytMode = 2, "登记", "预约申请")
        fraGroupInfo.Visible = True
        fraGroup.Visible = True
        tbs.TabVisible(1) = True
        
    Else
        lblTitle.Caption = "个人体检" & IIf(mbytMode = 2, "登记", "预约申请")
                
        tbs.TabVisible(1) = False
        fraSingle.Visible = True
        fraGroup.Visible = False
        
    End If
    
    If mbytMode = 2 Then
        Me.Caption = "体检登记"
        fraInfo.Visible = False
        dtp(0).Enabled = False
        dtp(0).Value = Format(zlDatabase.Currentdate, dtp(0).CustomFormat)
    Else
        dtp(0).Value = Format(DateAdd("d", 7, CDate(zlDatabase.Currentdate)), dtp(0).CustomFormat)
    End If
    
    '初始化数据
    'lvwGroup.TextMatrix(1, 1) = "缺省"
    lvwGroup.ListItems.Add , , "缺省", 1, 1
    
    '1.创建记录集,用于保存选择的体检项目
    Call MedicalItemsRecord(mrsItems)
    
    '2.创建记录集,用于保存选择的体检人员
    Call MedicalItemsRecord(mrsPersons, 2)
    
    InitData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  读取数据
    '参数:  lngKey      体检类型序号
    '返回:  True        读取成功
    '       False       读取失败
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
            
    On Error GoTo errHand
    
    '读取预约基本信息
    gstrSQL = "SELECT * FROM 体检登记记录 A WHERE A.ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        txt体检号.Text = zlCommFun.NVL(rs("体检号").Value)
        txt(0).Text = zlCommFun.NVL(rs("联系人").Value)
        txt(1).Text = zlCommFun.NVL(rs("联系电话").Value)
        txt(2).Text = zlCommFun.NVL(rs("联系地址").Value)
        txt(6).Text = zlCommFun.NVL(rs("附加说明").Value)
        cmd(4).Tag = zlCommFun.NVL(rs("合约单位id").Value, 0)
        dtp(0).Value = Format(zlCommFun.NVL(rs("体检时间").Value), dtp(0).CustomFormat)
        txt(31).Text = zlCommFun.NVL(rs("随访期限").Value, 0)
        chk.Value = IIf(Val(txt(31).Text) > 0, 1, 0)
    End If
                                        
    If mblnGroup Then Call ReadGroup(Val(cmd(4).Tag))
            
    Set rs = zlDatabase.OpenSQLRecord(GetPublicSQL(SQL.体检人员档案), Me.Caption, lngKey)
    If WriteItems(rs, mrsPersons, , 2) = False Then Exit Function
    
    If mrsPersons.RecordCount > 0 And mblnGroup = False Then Call ReadPersons("缺省", 2)
    
    '读取体检组别及体检项目
    
    lvwGroup.ListItems.Clear
    
    gstrSQL = "SELECT A.组别名称 AS 名称, rownum AS ID,1 As 图标 FROM 体检组别 A WHERE A.登记id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        Call FillLvw(lvwGroup, rs)
    Else
        lvwGroup.ListItems.Add , , "缺省", 1, 1
    End If
        
    '读取体检项目
    Set rs = zlDatabase.OpenSQLRecord(GetPublicSQL(SQL.团体体检项目), Me.Caption, mlngKey)
    
    If rs.BOF = False Then
        Do While Not rs.EOF
        
            mrsItems.AddNew
            
            mrsItems("组别").Value = zlCommFun.NVL(rs("组别名称").Value)
            mrsItems("ID").Value = zlCommFun.NVL(rs("ID").Value)
            mrsItems("类别").Value = zlCommFun.NVL(rs("类别").Value)
            mrsItems("名称").Value = zlCommFun.NVL(rs("名称").Value)
            mrsItems("基本价格").Value = Format(zlCommFun.NVL(rs("基本价格").Value), "0.00")
            mrsItems("体检价格").Value = Format(zlCommFun.NVL(rs("体检价格").Value), "0.00")
            mrsItems("体检类型").Value = zlCommFun.NVL(rs("体检类型").Value)
            mrsItems("结算方式").Value = zlCommFun.NVL(rs("结算方式").Value)
            mrsItems("执行科室").Value = zlCommFun.NVL(rs("执行科室").Value)
            mrsItems("采集科室").Value = zlCommFun.NVL(rs("采集科室").Value)
            mrsItems("采集科室id").Value = zlCommFun.NVL(rs("采集科室id").Value)
            mrsItems("执行科室id").Value = zlCommFun.NVL(rs("执行科室id").Value)
            mrsItems("采集方式").Value = zlCommFun.NVL(rs("采集方式").Value)
            mrsItems("采集方式id").Value = zlCommFun.NVL(rs("采集方式id").Value)
            mrsItems("检验标本").Value = zlCommFun.NVL(rs("检验标本").Value)
            mrsItems("检查部位").Value = zlCommFun.NVL(rs("检查部位").Value)
            mrsItems("检查部位id").Value = zlCommFun.NVL(rs("检查部位id").Value)
            mrsItems("折扣").Value = Format(zlCommFun.NVL(rs("折扣").Value), "0.0000")
            mrsItems("计费明细").Value = GetPriceList(zlCommFun.NVL(rs("清单id").Value))
            
            rs.MoveNext
        Loop
    End If
        
    If Not (lvwGroup.SelectedItem Is Nothing) Then Call lvwGroup_ItemClick(lvwGroup.SelectedItem)
    
    ReadData = True
    
    Exit Function
    
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function ReadGroup(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:读取团体基本资料
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHand
    
    gstrSQL = "SELECT A.* FROM 合约单位 A WHERE A.ID=" & lngKey
    
    If mrsGroup.State = adStateOpen Then mrsGroup.Close
    mrsGroup.Open gstrSQL, gcnOracle, adOpenStatic, adLockBatchOptimistic
    If mrsGroup.BOF Then mrsGroup.AddNew
    
    Call ShowGroupInfo
    
    ReadGroup = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ShowGroupInfo() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:填写团体信息到控件
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    
    If mrsGroup.RecordCount > 0 Then
        mrsGroup.MoveFirst
        
        txt(13).Text = zlCommFun.NVL(mrsGroup("名称").Value)
        txt(12).Text = zlCommFun.NVL(mrsGroup("联系人").Value)
        txt(11).Text = zlCommFun.NVL(mrsGroup("电话").Value)
        txt(8).Text = zlCommFun.NVL(mrsGroup("电子邮件").Value)
        cmd(4).Tag = zlCommFun.NVL(mrsGroup("ID").Value)
        
    End If
    
    ShowGroupInfo = True
    
End Function

Private Function SaveGroupInfo() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:更新团体信息
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    If mblnGroup Then
        If mrsGroup.RecordCount > 0 Then
            mrsGroup.MoveFirst
            
            mrsGroup("名称").Value = txt(13).Text
            mrsGroup("联系人").Value = txt(12).Text
            mrsGroup("电话").Value = txt(11).Text
            mrsGroup("电子邮件").Value = txt(8).Text
            mrsGroup("ID").Value = Val(cmd(4).Tag)
            
        End If
    End If
    
    SaveGroupInfo = True
    
End Function

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  检查是否有重复的项目
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) = lngKey And vsf.Row <> lngLoop Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function

Private Function CheckHavePerson(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  检查是否有重复的项目
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsfPerson.Rows - 1
        If Val(vsfPerson.TextMatrix(lngLoop, mPersonCol.病人id)) = lngKey And vsfPerson.Row <> lngLoop And Val(vsfPerson.TextMatrix(lngLoop, mPersonCol.病人id)) > 0 Then
            CheckHavePerson = True
            Exit Function
        End If
    Next
End Function

Private Function ShowOpenList(Optional strText As String, Optional ByVal lngCol As Long = 0) As Byte
    '------------------------------------------------------------------------------------------------------------------
    '功能:  以列表方式显示数据
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strClass As String
    Dim strPath As String
    Dim lngLoop As Long
    Dim strTmp As String
    
    On Error GoTo errHand
    
    ShowOpenList = 2
    
    Select Case lngCol
        Case mCol.项目
            strText = UCase(strText)
            
            strLvw = "编码,1200,0,1;名称,2700,0,0;单位,900,0,0;标本部位,900,0,0;类别,900,0,0"
            strPath = Me.Name & "\体检项目选择"
            
            gstrSQL = GetPublicSQL(SQL.体检项目过滤选择, strText)
            If ParamInfo.项目输入匹配方式 = 1 Then
                strTmp = strText & "%"
            Else
                strTmp = "%" & strText & "%"
            End If
            Dim bytParam1 As Byte
            Dim bytParam2 As Byte
            
            bytParam1 = 1
            bytParam2 = 2
                    
            If mblnGroup = False Then
                Select Case zlCommFun.GetNeedName(cbo(1).Text)
                Case "男"
                    bytParam1 = 1
                    bytParam2 = 1
                Case "女"
                    bytParam1 = 2
                    bytParam2 = 2
                End Select
            End If
            
            If Trim(vsf.TextMatrix(vsf.Row, mCol.类别)) = "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "D", strText & "%", strTmp, bytParam1, bytParam2)
            ElseIf Trim(vsf.TextMatrix(vsf.Row, mCol.类别)) = "检验" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "", strText & "%", strTmp, bytParam1, bytParam2)
            Else
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "D", "", strText & "%", strTmp, bytParam1, bytParam2)
            End If

        Case mCol.检查部位
            
            strText = "'%" & UCase(strText) & "%'"
            
            strLvw = "名称,3300,0,0"
            strPath = Me.Name & "\检查部位选择"
            
            gstrSQL = "select B.标本部位 AS 名称,B.ID,0 AS 选择 from 诊疗项目组合 A,诊疗项目目录 B WHERE (B.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or B.撤档时间 is NULL) AND A.诊疗项目ID=B.ID AND A.诊疗组合ID=" & Val(vsf.RowData(vsf.Row)) & ""
            
            rs.CursorLocation = adUseClient
            If rs.State = adStateOpen Then rs.Close
            rs.Open gstrSQL, gcnOracle, adOpenStatic, adLockOptimistic
            
    End Select
    
    If rs.BOF Then
        ShowOpenList = 0
        Exit Function
    End If
    If rs.RecordCount = 1 And strText <> "'%%'" Then GoTo PointOver
    Call CalcPosition(sglX, sglY, vsf)
    
    If lngCol = mCol.检查部位 Then
        If vsf.TextMatrix(vsf.Row, mCol.检查部位id) <> "" Then
            Do While Not rs.EOF
                If InStr("," & vsf.TextMatrix(vsf.Row, mCol.检查部位id) & ",", "," & rs("ID").Value & ",") > 0 Then rs("选择").Value = 1
                rs.MoveNext
            Loop
        End If
        rs.MoveFirst
        
        If frmSelectDialog.ShowSelect(Me, 2, rs, strLvw, "请从下面选择多个项目,然后回车或双击退出", sglX + 60, sglY + 30, 9000, 4500, 300, , strPath, , False, True) Then GoTo PointOver
        
    Else
                
        If frmSelectDialog.ShowSelect(Me, 2, rs, strLvw, "请从下面选择一个项目", sglX + 60, sglY + 30, 9000, 4500, vsf.Body.RowHeight(1), , strPath, , False) Then GoTo PointOver
        
    End If
        
    Exit Function
    
PointOver:
    Select Case lngCol
        Case mCol.项目
            If CheckHave(zlCommFun.NVL(rs("ID").Value, 0)) Then
                MsgBox "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已经被选择！", vbInformation, gstrSysName
                Exit Function
            End If
            
            vsf.Cell(flexcpText, vsf.Row, mCol.项目 + 1, vsf.Row, vsf.Cols - 1) = ""
            
            vsf.EditText = zlCommFun.NVL(rs("名称").Value)
            vsf.TextMatrix(vsf.Row, mCol.类别) = zlCommFun.NVL(rs("类别").Value)
            vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
            vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
            vsf.RowData(vsf.Row) = zlCommFun.NVL(rs("ID").Value)
            
        Case mCol.检查部位
            
            vsf.TextMatrix(vsf.Row, vsf.Col) = ""
            vsf.TextMatrix(vsf.Row, mCol.检查部位id) = ""
            
            rs.Filter = ""
            rs.Filter = "选择=1"
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                Do While Not rs.EOF
                    vsf.TextMatrix(vsf.Row, vsf.Col) = vsf.TextMatrix(vsf.Row, vsf.Col) & zlCommFun.NVL(rs("名称").Value) & ","
                    vsf.TextMatrix(vsf.Row, mCol.检查部位id) = vsf.TextMatrix(vsf.Row, mCol.检查部位id) & zlCommFun.NVL(rs("ID").Value) & ","
                    rs.MoveNext
                Loop
                
                If vsf.TextMatrix(vsf.Row, mCol.检查部位) <> "" Then vsf.TextMatrix(vsf.Row, mCol.检查部位) = Mid(vsf.TextMatrix(vsf.Row, mCol.检查部位), 1, Len(vsf.TextMatrix(vsf.Row, mCol.检查部位)) - 1)
                If vsf.TextMatrix(vsf.Row, mCol.检查部位id) <> "" Then vsf.TextMatrix(vsf.Row, mCol.检查部位id) = Mid(vsf.TextMatrix(vsf.Row, mCol.检查部位id), 1, Len(vsf.TextMatrix(vsf.Row, mCol.检查部位id)) - 1)
                
            End If

    End Select
    
    ShowOpenList = 1
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function SetRowData(ByVal lngKey As Long, ByVal intRow As Integer, ParamArray arryMode() As Variant) As Boolean
'------------------------------------------------------------------------------------------------------------------
    '功能:设置行数据（随行不同而不同）
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim strCombList As String
    
    On Error Resume Next
    
    For lngLoop = 0 To UBound(arryMode)
        Select Case arryMode(lngLoop)
        Case "收费执行科室"
        
            If InStr("4,5,6,7", vsfPrice.TextMatrix(intRow, mCol.p类别)) > 0 Then
                gstrSQL = GetPublicSQL(SQL.药品执行科室)
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, vsfPrice.TextMatrix(intRow, mCol.p类别))
            Else
                gstrSQL = GetPublicSQL(SQL.收费执行科室, "1")
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, mlngDept, UserInfo.部门ID, "%%")
            End If
            If rs.RecordCount > 1 Then
                vsfPrice.EditMode(mCol.p执行科室) = 1
                vsfPrice.Body.ColComboList(mCol.p执行科室) = vsfPrice.Body.BuildComboList(rs, "名称", "ID")
            Else
                vsfPrice.EditMode(mCol.p执行科室) = 0
                vsfPrice.Body.ColComboList(mCol.p执行科室) = ""
            End If
        
        Case "计价项目"
            
            If Trim(vsf.TextMatrix(intRow, mCol.类别)) = "检查" Then
                strCombList = "检查项目-" & Trim(vsf.TextMatrix(intRow, mCol.项目))
                vsfPrice.EditMode(mCol.p计价项目) = 0
                vsfPrice.Body.ColComboList(mCol.p计价项目) = ""
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p计价项目) = strCombList
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p计价性质) = "1"
            Else
                strCombList = "检验项目-" & Trim(vsf.TextMatrix(intRow, mCol.项目))
                If Val(vsf.TextMatrix(intRow, mCol.采集方式id)) > 0 Then
                    strCombList = strCombList & "|采集方式-" & Trim(vsf.TextMatrix(intRow, mCol.采集方式))
                    vsfPrice.EditMode(mCol.p计价项目) = 1
                    vsfPrice.Body.ColComboList(mCol.p计价项目) = strCombList
                Else
                    vsfPrice.EditMode(mCol.p计价项目) = 0
                    vsfPrice.Body.ColComboList(mCol.p计价项目) = ""
                End If
            End If
            
        Case "诊疗执行科室"
        
            gstrSQL = GetPublicSQL(SQL.诊疗执行科室, "1")
            If gstrSQL <> "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, mlngDept, UserInfo.部门ID, "%%")
                If rs.RecordCount > 1 Then
                    vsf.EditMode(mCol.执行科室) = 1
                    vsf.Body.ColComboList(mCol.执行科室) = vsf.Body.BuildComboList(rs, "名称", "ID")
                Else
                    vsf.EditMode(mCol.执行科室) = 0
                    vsf.Body.ColComboList(mCol.执行科室) = ""
                End If
            End If
        
        Case "采集方式"
        
            gstrSQL = "SELECT A.名称 AS 名称,A.ID FROM 诊疗项目目录 A,诊疗用法用量 B WHERE A.ID=B.用法id AND A.类别='E' AND A.操作类型='6' AND B.项目ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.RecordCount > 1 Then
                vsf.EditMode(mCol.采集方式) = 1
                vsf.Body.ColComboList(mCol.采集方式) = vsf.Body.BuildComboList(rs, "名称", "ID")
            Else
                gstrSQL = "SELECT A.名称 AS 名称,A.ID FROM 诊疗项目目录 A WHERE A.类别='E' AND A.操作类型='6'"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                If rs.RecordCount > 1 Then
                    vsf.EditMode(mCol.采集方式) = 1
                    vsf.Body.ColComboList(mCol.采集方式) = vsf.Body.BuildComboList(rs, "名称", "ID")
                Else
                    vsf.EditMode(mCol.采集方式) = 0
                    vsf.Body.ColComboList(mCol.采集方式) = ""
                End If
            End If
            
        Case "采集科室"
        
            gstrSQL = GetPublicSQL(SQL.诊疗执行科室)
            If gstrSQL <> "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.TextMatrix(intRow, mCol.采集方式id)), mlngDept, UserInfo.部门ID, "%%")
                If rs.RecordCount > 1 Then
                    vsf.EditMode(mCol.采集科室) = 1
                    vsf.Body.ColComboList(mCol.采集科室) = vsf.Body.BuildComboList(rs, "*名称", "ID")
                Else
                    vsf.EditMode(mCol.采集科室) = 0
                    vsf.Body.ColComboList(mCol.采集科室) = ""
                End If
            End If
        
        Case "检验标本"
        
            gstrSQL = "SELECT 1 FROM 诊疗项目目录 WHERE 组合项目=1 AND ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                '是组合项目
                
                gstrSQL = "SELECT DISTINCT A.标本类型 AS 名称 FROM 检验项目参考 A,检验报告项目 B,诊疗项目目录 C " & _
                        "WHERE C.ID<>[1] AND nvl(C.组合项目,0)=0 " & _
                            "AND B.报告项目id=A.项目id and rownum<2"
                            
                gstrSQL = gstrSQL & "AND B.诊疗项目id IN (SELECT C.ID " & _
                             "FROM 检验报告项目 A," & _
                                  "(SELECT 报告项目id FROM 检验报告项目 WHERE 诊疗项目id = [1]) B," & _
                                  "诊疗项目目录 C,诊治所见项目 D,检验项目 E,检验报告项目 F " & _
                            "WHERE A.报告项目id = B.报告项目id AND A.诊疗项目id <> [1] AND " & _
                                  "nvl(C.组合项目,0) = 0 AND A.诊疗项目id = C.ID AND C.ID=F.诊疗项目id AND F.报告项目id=D.ID AND D.ID=E.诊治项目id)  and rownum<2 "
                                          
            Else
                gstrSQL = "SELECT A.标本类型 AS 名称 FROM 检验项目参考 A,检验报告项目 B,诊疗项目目录 C " & _
                        "WHERE C.ID=[1] AND nvl(C.组合项目,0)=0 AND B.诊疗项目id=[1] and B.报告项目id=A.项目id  and rownum<2"
            End If
        
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.RecordCount > 1 Then
                
                vsf.EditMode(mCol.检验标本) = 1
                vsf.Body.ColComboList(mCol.检验标本) = vsf.Body.BuildComboList(rs, "名称", "名称")
                
            Else
                
                '没有对应时，读取所有标本类型
                gstrSQL = "SELECT 名称 FROM 诊疗检验标本 A"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                If rs.RecordCount > 1 Then
                
                    vsf.EditMode(mCol.检验标本) = 1
                    vsf.Body.ColComboList(mCol.检验标本) = vsf.Body.BuildComboList(rs, "名称", "名称")
                Else
                    vsf.EditMode(mCol.检验标本) = 0
                    vsf.Body.ColComboList(mCol.检验标本) = ""
                End If
                
            End If
        
        End Select
    Next
    
    SetRowData = True
    
End Function

Private Function SetRowDefault(ByVal lngKey As Long, ByVal intRow As Integer, ParamArray arryMode() As Variant) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:获取缺省
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim strCombList As String
    
    On Error GoTo errHand
    
    For lngLoop = 0 To UBound(arryMode)
        
        Select Case arryMode(lngLoop)
        Case "缺省信息"
            '先按上行读取
            With vsfPerson
                If vsfPerson.Row > 1 Then
                    .TextMatrix(vsfPerson.Row, mPersonCol.性别) = .TextMatrix(vsfPerson.Row - 1, mPersonCol.性别)
                    .TextMatrix(vsfPerson.Row, mPersonCol.婚姻状况) = .TextMatrix(vsfPerson.Row - 1, mPersonCol.婚姻状况)
                End If
            End With
            
        Case "结算方式"
            
            If mblnGroup Then
                vsf.TextMatrix(vsf.Row, mCol.结算方式) = "记帐"
            Else
                vsf.TextMatrix(vsf.Row, mCol.结算方式) = "收费"
            End If
            
        Case "执行科室"
'            lng开单科室id = mlngDept
            
            gstrSQL = GetPublicSQL(SQL.诊疗执行科室)
            If gstrSQL <> "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, mlngDept, UserInfo.部门ID, "%%")
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Row, mCol.执行科室) = zlCommFun.NVL(rs("名称").Value)
                    vsf.TextMatrix(vsf.Row, mCol.执行科室id) = zlCommFun.NVL(rs("ID").Value)
                Else
                    vsf.TextMatrix(vsf.Row, mCol.执行科室) = gstrDeptName
                    vsf.TextMatrix(vsf.Row, mCol.执行科室id) = UserInfo.部门ID
                End If
            End If
        
        Case "采集方式"
           
            
            gstrSQL = "SELECT A.名称 AS 名称,A.ID FROM 诊疗项目目录 A,诊疗用法用量 B WHERE A.ID=B.用法id AND A.类别='E' AND A.操作类型='6' AND B.项目ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                vsf.TextMatrix(vsf.Row, mCol.采集方式) = zlCommFun.NVL(rs("名称").Value)
                vsf.TextMatrix(vsf.Row, mCol.采集方式id) = zlCommFun.NVL(rs("ID").Value)
            Else
                gstrSQL = "SELECT A.名称 AS 名称,A.ID FROM 诊疗项目目录 A WHERE A.类别='E' AND A.操作类型='6'"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Row, mCol.采集方式) = zlCommFun.NVL(rs("名称").Value)
                    vsf.TextMatrix(vsf.Row, mCol.采集方式id) = zlCommFun.NVL(rs("ID").Value)
                End If
            End If
            
        Case "采集科室"
                    
            gstrSQL = GetPublicSQL(SQL.诊疗执行科室)
            If gstrSQL <> "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.TextMatrix(vsf.Row, mCol.采集方式id)), mlngDept, UserInfo.部门ID, "%%")
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Row, mCol.采集科室) = zlCommFun.NVL(rs("名称").Value)
                    vsf.TextMatrix(vsf.Row, mCol.采集科室id) = zlCommFun.NVL(rs("ID").Value)
                End If
            End If
        
        Case "检验标本"
            
            
            gstrSQL = "SELECT 1 FROM 诊疗项目目录 WHERE 组合项目=1 AND ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                '是组合项目
                
                gstrSQL = "SELECT DISTINCT A.标本类型 AS 名称 FROM 检验项目参考 A,检验报告项目 B,诊疗项目目录 C " & _
                        "WHERE C.ID<>[1] AND nvl(C.组合项目,0)=0 " & _
                            "AND B.报告项目id=A.项目id and rownum<2"
                            
                gstrSQL = gstrSQL & "AND B.诊疗项目id IN (SELECT C.ID " & _
                             "FROM 检验报告项目 A," & _
                                  "(SELECT 报告项目id FROM 检验报告项目 WHERE 诊疗项目id = [1]) B," & _
                                  "诊疗项目目录 C,诊治所见项目 D,检验项目 E,检验报告项目 F " & _
                            "WHERE A.报告项目id = B.报告项目id AND A.诊疗项目id <> [1] AND " & _
                                  "nvl(C.组合项目,0) = 0 AND A.诊疗项目id = C.ID AND C.ID=F.诊疗项目id AND F.报告项目id=D.ID AND D.ID=E.诊治项目id)  and rownum<2 "
                                          
            Else
                gstrSQL = "SELECT A.标本类型 AS 名称 FROM 检验项目参考 A,检验报告项目 B,诊疗项目目录 C " & _
                        "WHERE C.ID=[1] AND nvl(C.组合项目,0)=0 AND B.诊疗项目id=[1] and B.报告项目id=A.项目id  and rownum<2"
            End If
        
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                vsf.TextMatrix(vsf.Row, mCol.检验标本) = rs("名称").Value
            Else
                
                '没有对应时，读取所有标本类型
                gstrSQL = "SELECT 名称 FROM 诊疗检验标本 A where rownum<2"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Row, mCol.检验标本) = rs("名称").Value
                End If
                
            End If
        
        Case "收费执行科室"
            
            If InStr("4,5,6,7", vsfPrice.TextMatrix(intRow, mCol.p类别)) > 0 Then
                gstrSQL = GetPublicSQL(SQL.药品执行科室)
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, vsfPrice.TextMatrix(intRow, mCol.p类别))
            Else
                gstrSQL = GetPublicSQL(SQL.收费执行科室)
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, mlngDept, UserInfo.部门ID, "%%")
            End If
            If rs.BOF = False Then
                vsfPrice.TextMatrix(intRow, mCol.p执行科室) = zlCommFun.NVL(rs("名称").Value)
                vsfPrice.TextMatrix(intRow, mCol.p执行科室id) = zlCommFun.NVL(rs("ID").Value)
            Else
                vsfPrice.TextMatrix(intRow, mCol.p执行科室) = vsf.TextMatrix(vsf.Row, mCol.执行科室)
                vsfPrice.TextMatrix(intRow, mCol.p执行科室id) = vsf.TextMatrix(vsf.Row, mCol.执行科室id)
            End If
            
        Case "计价项目"
        
            If Trim(vsf.TextMatrix(vsf.Row, mCol.类别)) = "检查" Then
                strCombList = "检查项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目))
                vsfPrice.EditMode(mCol.p计价项目) = 0
                vsfPrice.Body.ColComboList(mCol.p计价项目) = ""
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p计价项目) = strCombList
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p计价性质) = "1"
            Else
                strCombList = "检验项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目))
                If Val(vsf.TextMatrix(vsf.Row, mCol.采集方式id)) > 0 Then
                    strCombList = strCombList & "|采集方式-" & Trim(vsf.TextMatrix(vsf.Row, mCol.采集方式))
                    vsfPrice.EditMode(mCol.p计价项目) = 1
                    vsfPrice.Body.ColComboList(mCol.p计价项目) = strCombList
                Else
                    vsfPrice.EditMode(mCol.p计价项目) = 0
                    vsfPrice.Body.ColComboList(mCol.p计价项目) = ""
                End If
            End If

        End Select
    Next
    
    SetRowDefault = True
    
    Exit Function
    
errHand:
    
End Function

Private Function SaveItems(ByVal strGroup As String) As Boolean
    
    Dim lngLoop As Long
    
    On Error GoTo errHand

    '保存所选择的检验项目
    mrsItems.Filter = ""
    mrsItems.Filter = "组别='" & strGroup & "' AND 删除<>'1'"
    
    Call DeleteRecord(mrsItems)
    
    For lngLoop = 1 To vsf.Rows - 1
        
        If Val(vsf.RowData(lngLoop)) > 0 Then
            mrsItems.AddNew
            
            mrsItems("组别").Value = strGroup
            mrsItems("ID").Value = vsf.RowData(lngLoop)
            mrsItems("类别").Value = vsf.TextMatrix(lngLoop, mCol.类别)
            mrsItems("名称").Value = vsf.TextMatrix(lngLoop, mCol.项目)
            mrsItems("执行科室").Value = vsf.TextMatrix(lngLoop, mCol.执行科室)
            mrsItems("检查部位").Value = vsf.TextMatrix(lngLoop, mCol.检查部位)
            mrsItems("采集方式").Value = vsf.TextMatrix(lngLoop, mCol.采集方式)
            mrsItems("采集科室").Value = vsf.TextMatrix(lngLoop, mCol.采集科室)
            mrsItems("检验标本").Value = vsf.TextMatrix(lngLoop, mCol.检验标本)
            mrsItems("体检类型").Value = vsf.TextMatrix(lngLoop, mCol.体检类型)
            mrsItems("基本价格").Value = vsf.TextMatrix(lngLoop, mCol.基本价格)
            mrsItems("体检价格").Value = vsf.TextMatrix(lngLoop, mCol.体检价格)
            mrsItems("结算方式").Value = vsf.TextMatrix(lngLoop, mCol.结算方式)
            mrsItems("执行科室id").Value = vsf.TextMatrix(lngLoop, mCol.执行科室id)
            mrsItems("采集方式id").Value = vsf.TextMatrix(lngLoop, mCol.采集方式id)
            mrsItems("采集科室id").Value = vsf.TextMatrix(lngLoop, mCol.采集科室id)
            mrsItems("检查部位id").Value = vsf.TextMatrix(lngLoop, mCol.检查部位id)
            mrsItems("计费明细").Value = vsf.TextMatrix(lngLoop, mCol.计费明细)
            
            mrsItems("新加").Value = vsf.TextMatrix(lngLoop, mCol.新加)
            mrsItems("前景色").Value = vsf.TextMatrix(lngLoop, mCol.前景色)
            mrsItems("删除").Value = ""
            mrsItems("公共").Value = vsf.TextMatrix(lngLoop, mCol.公共)
            
        End If
    Next
    
    SaveItems = True
    
errHand:

End Function

Private Function WritePersons(ByVal strGroup As String, Optional bytMode As Byte = 1) As Boolean
    
    Dim lngLoop As Long
    
    On Error GoTo errHand

    '保存所选择的检验项目
    If bytMode = 1 Then
        mrsPersons.Filter = ""
        mrsPersons.Filter = "组别='" & strGroup & "' AND 删除<>'1'"
    
        Call DeleteRecord(mrsPersons)
    
        For lngLoop = 1 To vsfPerson.Rows - 1
            
            If vsfPerson.TextMatrix(lngLoop, mPersonCol.姓名) <> "" Then
                mrsPersons.AddNew
                
                mrsPersons("组别").Value = strGroup
                
                With vsfPerson
                    mrsPersons("IC卡号").Value = .TextMatrix(lngLoop, mPersonCol.IC卡号)
                    mrsPersons("健康号").Value = .TextMatrix(lngLoop, mPersonCol.健康号)
                    mrsPersons("病人id").Value = .TextMatrix(lngLoop, mPersonCol.病人id)
                    mrsPersons("姓名").Value = .TextMatrix(lngLoop, mPersonCol.姓名)
                    mrsPersons("门诊号").Value = .TextMatrix(lngLoop, mPersonCol.门诊号)
                    mrsPersons("身份证").Value = .TextMatrix(lngLoop, mPersonCol.身份证)
                    mrsPersons("性别").Value = .TextMatrix(lngLoop, mPersonCol.性别)
                    mrsPersons("出生日期").Value = .TextMatrix(lngLoop, mPersonCol.出生日期)
                    mrsPersons("婚姻状况").Value = .TextMatrix(lngLoop, mPersonCol.婚姻状况)
                    mrsPersons("年龄").Value = .TextMatrix(lngLoop, mPersonCol.年龄)
                    mrsPersons("民族").Value = .TextMatrix(lngLoop, mPersonCol.民族)
                    mrsPersons("国籍").Value = .TextMatrix(lngLoop, mPersonCol.国籍)
                    mrsPersons("学历").Value = .TextMatrix(lngLoop, mPersonCol.学历)
                    mrsPersons("职业").Value = .TextMatrix(lngLoop, mPersonCol.职业)
                    mrsPersons("身份").Value = .TextMatrix(lngLoop, mPersonCol.身份)
                    mrsPersons("联系人姓名").Value = .TextMatrix(lngLoop, mPersonCol.联系人姓名)
                    mrsPersons("联系人电话").Value = .TextMatrix(lngLoop, mPersonCol.联系人电话)
                    mrsPersons("电子邮件").Value = .TextMatrix(lngLoop, mPersonCol.电子邮件)
                    mrsPersons("联系人地址").Value = .TextMatrix(lngLoop, mPersonCol.联系人地址)
                    mrsPersons("工作单位").Value = .TextMatrix(lngLoop, mPersonCol.工作单位)
                    mrsPersons("登记时间").Value = .TextMatrix(lngLoop, mPersonCol.登记时间)
                    mrsPersons("就诊卡号").Value = .TextMatrix(lngLoop, mPersonCol.就诊卡号)
                    mrsPersons("删除").Value = ""
        
                End With
                
            End If
        Next
    End If
    
    If bytMode = 2 Then
        mrsPersons.Filter = ""
        mrsPersons.Filter = "删除<>'1'"
        If mrsPersons.RecordCount = 0 Then mrsPersons.AddNew
                
        mrsPersons("组别").Value = "缺省"
        mrsPersons("病人id").Value = Val(cmd(1).Tag)
        mrsPersons("姓名").Value = txt(5).Text
        mrsPersons("身份证").Value = txt(4).Text
        mrsPersons("性别").Value = zlCommFun.GetNeedName(cbo(1).Text)
        mrsPersons("年龄").Value = txt(9).Text
        mrsPersons("婚姻状况").Value = zlCommFun.GetNeedName(cbo(0).Text)
        mrsPersons("联系人电话").Value = txt(10).Text

        mrsPersons("门诊号").Value = txt(3).Text
        mrsPersons("健康号").Value = txt(14).Text
    End If
    WritePersons = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  校验数据的有效性
    '返回:  True        数据有效
    '       False       数据无效
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    
    If mbytMode = 1 Then
        '检查预约人
        If Trim(txt(0).Text) = "" Then
            ShowSimpleMsg "预约人不能为空值，必须输入！"
            Call LocationObj(txt(0))
            Exit Function
        End If
        
        If mblnGroup = False Then
            If Format(dtp(0).Value, "yyyy-MM-dd") < Format(zlDatabase.Currentdate, "yyyy-MM-dd") Then
                ShowSimpleMsg "预约体检时间不能小于当天！"
                dtp(0).SetFocus
                Exit Function
            End If
        End If
    End If
    
    '检查团体
    If fraGroupInfo.Visible And mblnGroup Then
        If Trim(txt(13).Text) = "" Then
            ShowSimpleMsg "必须要确定团体！"
            Call LocationObj(txt(13))
            Exit Function
        End If
    End If
            
    '检查组别名称是否有效
    For lngLoop = 1 To lvwGroup.ListItems.Count
        If Trim(lvwGroup.ListItems(lngLoop).Text) = "" Then
            ShowSimpleMsg "体检组别不能为空！"
            lvwGroup.SetFocus
            Exit Function
        End If
        
        If StrIsValid(lvwGroup.ListItems(lngLoop).Text, 30) = False Then
            lvwGroup.SetFocus
            Exit Function
        End If
        
    Next
    
    '检查电子邮件
    If mblnGroup Then
        If CheckStrValid(txt(8).Text, CHECKFORMAT.电子邮件) = False Then
        
            ShowSimpleMsg "错误的电子邮件格式。电子邮件格式如下：" & vbCrLf & "1.必须包含@字符；" & vbCrLf & "2.@字符只能在中间，如 xxx@163.com。"
            Call LocationObj(txt(8))
            Exit Function
            
        End If
    Else
        
        If CheckStrValid(txt(7).Text, CHECKFORMAT.电子邮件) = False Then
        
            ShowSimpleMsg "错误的电子邮件格式。电子邮件格式如下：" & vbCrLf & "1.必须包含@字符；" & vbCrLf & "2.@字符只能在中间，如 xxx@163.com。"
            Call LocationObj(txt(7))
            Exit Function
            
        End If
        
    End If
    
    If mbytMode = 2 Then
        
        mrsItems.Filter = ""
        mrsItems.Filter = "ID>0"
        If mrsItems.RecordCount = 0 Then
            ShowSimpleMsg "当前没有体检项目"
            tbs.Tab = 0
            Call tbs_Click(1)
            vsf.SetFocus
            Exit Function
        End If
        
        If mblnGroup = False Then
            If Trim(txt(5).Text) = "" Then
                ShowSimpleMsg "当前体检还没有设置体检人员"
                
                Call LocationObj(txt(5))
                Exit Function
            End If
        Else
                mrsPersons.Filter = ""
                mrsPersons.Filter = "姓名<>''"
                If mrsPersons.RecordCount = 0 Then
                    ShowSimpleMsg "当前体检还没有设置体检人员"
                    tbs.Tab = 1
                    Call tbs_Click(0)
                    vsfPerson.SetFocus
                    Exit Function
                End If
        End If
        
    End If
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.TextMatrix(lngLoop, mCol.体检价格)) < 0 Then
            tbs.Tab = 0
            Call tbs_Click(1)
            
            ShowSimpleMsg "体检价格不能为负数"
            vsf.Row = lngLoop
            vsf.Col = mCol.体检价格
            vsf.ShowCell vsf.Row, vsf.Col
            vsf.SetFocus
            
            Exit Function
        End If
        
'        If Format(Val(vsf.TextMatrix(lngLoop, mCol.体检价格)), "0.00") > Format(Val(vsf.TextMatrix(lngLoop, mCol.基本价格)), "0.00") Then
'
'            tbs.Tab = 0
'            Call tbs_Click(1)
'
'            ShowSimpleMsg "体检价格不能大于基本价格"
'            vsf.Row = lngLoop
'            vsf.Col = mCol.体检价格
'            vsf.ShowCell vsf.Row, vsf.Col
'            vsf.SetFocus
'
'            Exit Function
'        End If
        
    Next
    
    If mblnGroup Then
        For lngLoop = 1 To vsfPerson.Rows - 1
            
            '检查门诊号是否存在
            If Trim(vsfPerson.TextMatrix(lngLoop, mPersonCol.门诊号)) <> "" And Val(vsfPerson.TextMatrix(lngLoop, mPersonCol.病人id)) = 0 Then
                gstrSQL = "Select 1 From 病人信息 Where 门诊号=[1]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(vsfPerson.TextMatrix(lngLoop, mPersonCol.门诊号)))
                If rs.BOF = False Then
                    
                    ShowSimpleMsg "当前门诊号：" & Trim(vsfPerson.TextMatrix(lngLoop, mPersonCol.门诊号)) & "已经存在，不允许重复！"
                    
                    tbs.Tab = 1
                    Call tbs_Click(0)
                    vsfPerson.Row = lngLoop
                    vsfPerson.Col = mPersonCol.门诊号
                    vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                    vsfPerson.SetFocus
                    
                    Exit Function
                End If
            End If
            
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.姓名), GetMaxLength("病人信息", "姓名")) = False Then
                
                tbs.Tab = 1
                Call tbs_Click(0)
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.姓名
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.身份证), GetMaxLength("病人信息", "身份证号")) = False Then
                
                tbs.Tab = 1
                Call tbs_Click(0)
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.身份证
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.婚姻状况), GetMaxLength("病人信息", "婚姻状况")) = False Then
                tbs.Tab = 1
                Call tbs_Click(0)
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.婚姻状况
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.电子邮件), GetMaxLength("体检人员档案", "电子邮件")) = False Then
                tbs.Tab = 1
                Call tbs_Click(0)
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.电子邮件
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.性别), GetMaxLength("病人信息", "性别")) = False Then
                tbs.Tab = 1
                Call tbs_Click(0)
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.性别
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
        
            If Trim(vsfPerson.TextMatrix(lngLoop, mPersonCol.出生日期)) <> "" Then
                
                If CheckStrValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.出生日期), CHECKFORMAT.日期) = False Then
                    
                    tbs.Tab = 1
                    Call tbs_Click(0)
                    
                    ShowSimpleMsg "非法的出生日期！"
                    
                    vsfPerson.Row = lngLoop
                    vsfPerson.Col = mPersonCol.出生日期
                    vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                    vsfPerson.SetFocus
                    
                    Exit Function
                End If
            End If
            
            If CheckStrValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.电子邮件), CHECKFORMAT.电子邮件) = False Then
                    
                    tbs.Tab = 1
                    Call tbs_Click(0)
                    
                    ShowSimpleMsg "电子邮件必须包括 @ 符号，且不在第1位和最后一位上！"
                    
                    vsfPerson.Row = lngLoop
                    vsfPerson.Col = mPersonCol.电子邮件
                    vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                    vsfPerson.SetFocus
                    
                    Exit Function
            End If
                
            If CheckStrValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.身份证), CHECKFORMAT.身份证号) = False Then
                    
                    tbs.Tab = 1
                    Call tbs_Click(0)
                    
                    ShowSimpleMsg "身份证号非法（必须为15位或18位，为0-9、X字符）！"
                    
                    vsfPerson.Row = lngLoop
                    vsfPerson.Col = mPersonCol.身份证
                    vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                    vsfPerson.SetFocus
                    
                    Exit Function
            End If
        Next
    End If
    
    ValidEdit = True
    
End Function

Private Function SaveEdit(ByRef lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  保存数据
    '返回:  True        保存成功
    '       False       保存失败
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim lngRow As Long
    Dim strSQL() As String
    Dim strNow As String
    Dim rsPati As New ADODB.Recordset
    Dim lng病人id As Long
    Dim strGroup As String
    Dim intCount1 As Integer
    Dim str门诊号 As String
    Dim intCount2 As Integer
    Dim bytNew As Byte
    Dim strRegisteDate As String
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    If mblnGroup Then
        If Val(cmd(4).Tag) = 0 Then
            '新增团体
            cmd(4).Tag = zlDatabase.GetNextId("合约单位")
            
            gstrSQL = "zl_合约单位_Insert(" & Val(cmd(4).Tag) & "," & _
                                            "NULL,'" & _
                                            IIf(zlCommFun.NVL(mrsGroup("编码")) = "", GetNextCode("合约单位", "编码", ""), mrsGroup("编码")) & "','" & _
                                            txt(13).Text & "','" & _
                                            zlCommFun.SpellCode(txt(13).Text) & "'," & _
                                            IIf(IsNull(mrsGroup("地址").Value), "NULL", "'" & mrsGroup("地址").Value & "'") & ",'" & _
                                            txt(11).Text & "'," & _
                                            IIf(IsNull(mrsGroup("开户银行").Value), "NULL", "'" & mrsGroup("开户银行").Value & "'") & "," & _
                                            IIf(IsNull(mrsGroup("帐号").Value), "NULL", "'" & mrsGroup("帐号").Value & "'") & ",'" & _
                                            txt(12).Text & "'," & _
                                            "1," & _
                                            IIf(IsNull(mrsGroup("电子邮件").Value), "NULL", "'" & mrsGroup("电子邮件").Value & "'") & "," & _
                                            IIf(IsNull(mrsGroup("说明").Value), "NULL", "'" & mrsGroup("说明").Value & "'") & _
                                            ")"
                                            
            strSQL(ReDimArray(strSQL)) = gstrSQL
        Else
            '修改团体

            gstrSQL = "zl_合约单位_Update(" & Val(cmd(4).Tag) & "," & _
                                            IIf(IsNull(mrsGroup("上级ID").Value), "NULL", mrsGroup("上级ID").Value) & "," & _
                                            IIf(IsNull(mrsGroup("编码").Value), "NULL", "'" & mrsGroup("编码").Value & "'") & ",'" & _
                                            txt(13).Text & "','" & _
                                            zlCommFun.SpellCode(txt(13).Text) & "'," & _
                                            IIf(IsNull(mrsGroup("地址").Value), "NULL", "'" & mrsGroup("地址").Value & "'") & ",'" & _
                                            txt(11).Text & "'," & _
                                            IIf(IsNull(mrsGroup("开户银行").Value), "NULL", "'" & mrsGroup("开户银行").Value & "'") & "," & _
                                            IIf(IsNull(mrsGroup("帐号").Value), "NULL", "'" & mrsGroup("帐号").Value & "'") & ",'" & _
                                            txt(12).Text & _
                                            "',0," & _
                                            IIf(IsNull(mrsGroup("电子邮件").Value), "NULL", "'" & mrsGroup("电子邮件").Value & "'") & "," & _
                                            IIf(IsNull(mrsGroup("说明").Value), "NULL", "'" & mrsGroup("说明").Value & "'") & _
                                            ")"
            strSQL(ReDimArray(strSQL)) = gstrSQL

        End If
    End If
    
    If mlngKey = 0 Then
        
        '取体检号
        txt体检号.Text = GetNextNo(78)
        
        '新增预约
        If Val(tbs.Tag) > 0 Then
            lngKey = Val(tbs.Tag)
        Else
            lngKey = zlDatabase.GetNextId("体检登记记录")
        End If

        strSQL(ReDimArray(strSQL)) = "ZL_体检登记记录_INSERT(" & lngKey & ",'" & _
                                                            txt体检号.Text & "'," & _
                                                            "1," & _
                                                            "1,'" & _
                                                            txt(0).Text & "','" & _
                                                            txt(1).Text & "'," & _
                                                            "NULL,'" & _
                                                            txt(2).Text & "'," & _
                                                            IIf(Val(cmd(4).Tag) = 0, "NULL", Val(cmd(4).Tag)) & "," & _
                                                            "1," & _
                                                            "TO_DATE('" & Format(dtp(0).Value, "yyyy-MM-dd") & " 00:00:00','yyyy-mm-dd hh24:mi:ss')," & _
                                                            mlngDept & ",'" & _
                                                            txt(6).Text & "'," & _
                                                            "TO_DATE('" & strNow & "','yyyy-mm-dd hh24:mi:ss')," & _
                                                            "NULL," & _
                                                            IIf(mblnGroup, 1, 0) & "," & _
                                                            "1," & _
                                                            IIf(chk.Value = 1, Val(txt(31).Text), "NULL") & ")"
        
        '保存操作步骤....
        
        
    Else
        '修改预约
        lngKey = mlngKey
        strSQL(ReDimArray(strSQL)) = "ZL_体检登记记录_UPDATE(" & lngKey & ",'" & _
                                                            txt体检号.Text & "'," & _
                                                            "1," & _
                                                            "1,'" & _
                                                            txt(0).Text & "','" & _
                                                            txt(1).Text & "'," & _
                                                            "NULL,'" & _
                                                            txt(2).Text & "'," & _
                                                            IIf(Val(cmd(4).Tag) = 0, "NULL", Val(cmd(4).Tag)) & "," & _
                                                            "1," & _
                                                            "TO_DATE('" & Format(dtp(0).Value, "yyyy-MM-dd") & " 00:00:00','yyyy-mm-dd hh24:mi:ss')," & _
                                                            mlngDept & ",'" & _
                                                            txt(6).Text & "'," & _
                                                            "TO_DATE('" & strNow & "','yyyy-mm-dd hh24:mi:ss')," & _
                                                            "NULL," & _
                                                            "1," & _
                                                            IIf(chk.Value = 1, Val(txt(31).Text), "NULL") & ")"
                                                            
    End If
    
    strSQL(ReDimArray(strSQL)) = "zl_体检人员档案_Delete(" & lngKey & ")"
    
    strGroup = ""
    mrsPersons.Filter = ""
    If mrsPersons.RecordCount > 0 Then
        mrsPersons.Filter = ""
        If mblnGroup Then mrsPersons.Sort = "组别"
        If mrsPersons.RecordCount > 0 Then mrsPersons.MoveFirst
        
        Dim intCount As Integer

        intCount = -1
        Do While Not mrsPersons.EOF
            
            '检查出生日期
            If mrsPersons("出生日期") <> "" Then
                
                If CheckStrValid(mrsPersons("出生日期"), CHECKFORMAT.日期) = False Then
                    ShowSimpleMsg mrsPersons("姓名").Value & "的出生日期无效！"
                    Exit Function
                End If
            End If
            
            If mblnGroup Then
                If strGroup <> mrsPersons("组别").Value Then strGroup = mrsPersons("组别").Value
            Else
                strGroup = "缺省"
            End If
            
            lng病人id = zlCommFun.NVL(mrsPersons("病人id"), 0)
            bytNew = 0
            If lng病人id = 0 Then
                bytNew = 1
                intCount = intCount + 1
                'lng病人id = GetNextPatientID + intCount
                lng病人id = GetNextNo(1) + intCount
            End If
            
            intCount1 = intCount1 + 1
            
            If zlCommFun.NVL(mrsPersons("门诊号").Value, 0) < 1 Then
                'lng门诊号 = NextNo(3) + intCount2
                str门诊号 = CStr(GetNextNo(3) + intCount2)
                intCount2 = intCount2 + 1
            Else
                str门诊号 = CStr(zlCommFun.NVL(mrsPersons("门诊号").Value, 0))
            End If
            
            If zlCommFun.NVL(mrsPersons("登记时间").Value, "") <> "" Then
                strRegisteDate = mrsPersons("登记时间").Value
                strRegisteDate = "To_Date('" & Format(strRegisteDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
            Else
                strRegisteDate = "Null"
            End If
            
            strSQL(ReDimArray(strSQL)) = "ZL_体检人员档案_INSERT(" & lngKey & "," & _
                                                            lng病人id & "," & _
                                                            "'" & strGroup & "','" & _
                                                            mrsPersons("姓名").Value & "','" & _
                                                            mrsPersons("身份证").Value & "','" & _
                                                            mrsPersons("性别").Value & "'," & _
                                                            IIf(mrsPersons("出生日期").Value = "", "NULL", "TO_DATE('" & mrsPersons("出生日期").Value & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                                                            mrsPersons("婚姻状况").Value & "','" & _
                                                            mrsPersons("民族").Value & "','" & _
                                                            mrsPersons("国籍").Value & "','" & _
                                                            mrsPersons("学历").Value & "','" & _
                                                            mrsPersons("职业").Value & "','" & _
                                                            mrsPersons("联系人姓名").Value & "','" & _
                                                            mrsPersons("联系人电话").Value & "','" & _
                                                            mrsPersons("电子邮件").Value & "','" & _
                                                            mrsPersons("联系人地址").Value & "','" & _
                                                            mrsPersons("工作单位").Value & "','" & _
                                                            mrsPersons("年龄").Value & "'," & _
                                                            Val(str门诊号) & ",'" & _
                                                            mrsPersons("IC卡号").Value & "','" & _
                                                            mrsPersons("健康号").Value & "','" & _
                                                            mrsPersons("就诊卡号").Value & "'," & _
                                                            "1," & _
                                                            IIf(intCount1 = mrsPersons.RecordCount, "1", "0") & ",0," & bytNew & "," & strRegisteDate & _
                                                            ")"
            mrsPersons.MoveNext
        Loop
    End If

    
    '保存选择的体检项目
    strSQL(ReDimArray(strSQL)) = "ZL_体检项目清单_DELETE(" & lngKey & ")"
    strSQL(ReDimArray(strSQL)) = "ZL_体检组别_DELETE(" & lngKey & ")"
        
    For lngLoop = 1 To lvwGroup.ListItems.Count
        If Trim(lvwGroup.ListItems(lngLoop).Text) <> "" Then
            strSQL(ReDimArray(strSQL)) = "ZL_体检组别_INSERT(" & lngKey & ",'" & Trim(lvwGroup.ListItems(lngLoop).Text) & "')"
        End If
    Next
    
    Dim strTmp As String
    Dim varRow As Variant
    Dim varCol As Variant
    
    mrsItems.Filter = ""
    If mrsItems.RecordCount > 0 Then
        
        mrsItems.Filter = ""
        mrsItems.Sort = "组别"
        If mrsItems.RecordCount > 0 Then mrsItems.MoveFirst
        
        strGroup = ""
        
        Do While Not mrsItems.EOF
            
            If strGroup <> mrsItems("组别").Value Then strGroup = mrsItems("组别").Value
                        
            If mrsItems("ID").Value > 0 Then
                
                strTmp = ""
                varRow = Split(mrsItems("计费明细").Value, ";")
                For lngLoop = 0 To UBound(varRow)
                    
                    varCol = Split(varRow(lngLoop), ":")
                    
                    If strTmp <> "" Then strTmp = strTmp & ";"
                    strTmp = strTmp & varCol(5) & ":" & varCol(2) & ":" & varCol(3) & ":" & varCol(4) & ":" & Val(varCol(8)) & ":" & Val(varCol(6))
                    
                Next
                                
                strSQL(ReDimArray(strSQL)) = "ZL_体检项目清单_INSERT(" & lngKey & "," & _
                                            "'" & strGroup & "'," & _
                                            mrsItems("ID").Value & ",'" & _
                                            mrsItems("体检类型").Value & "'," & _
                                            Val(mrsItems("基本价格").Value) & "," & _
                                            Val(mrsItems("体检价格").Value) & "," & _
                                            mrsItems("执行科室id").Value & "," & _
                                            IIf(mrsItems("采集方式id") = "", "NULL", mrsItems("采集方式id")) & "," & _
                                            IIf(mrsItems("采集科室id") = "", "NULL", mrsItems("采集科室id")) & ",'" & _
                                            mrsItems("检验标本").Value & "','" & _
                                            mrsItems("检查部位").Value & "','" & _
                                            mrsItems("检查部位id").Value & "',NULL," & IIf(mrsItems("结算方式").Value = "记帐", "1", "2") & ",'" & _
                                            strTmp & "')"
            End If
            
            mrsItems.MoveNext
        Loop
    End If
    
    strSQL(ReDimArray(strSQL)) = "ZL_体检登记记录_体检类型(" & lngKey & ")"
    
    '如果是体检登记时，还要处理 预约确认
    If mbytMode = 2 Then
        
        strSQL(ReDimArray(strSQL)) = "ZL_体检登记记录_STATE(" & lngKey & ",2)"
        
    End If
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveEdit = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Function ReadItems(ByVal strGroup As String) As Boolean
    
    mrsItems.Filter = ""
    mrsItems.Filter = "组别='" & strGroup & "' AND 删除<>'1'"
    If mrsItems.RecordCount > 0 Then
        mrsItems.MoveFirst
        Call FillGrid(vsf, mrsItems)
    End If
    Call ReadPrice(vsf.Row)
    
    Call vsf_BeforeRowColChange(0, 0, vsf.Row, vsf.Col, False)
    Call vsfPrice_AfterRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col)
    
    Call ChangeItem(Val(vsf.TextMatrix(vsf.Row, mCol.基本价格)), Val(vsf.TextMatrix(vsf.Row, mCol.体检价格)), 1, False)
    
    ReadItems = True
    
End Function

Private Function ReadPersons(ByVal strGroup As String, Optional ByVal bytMode As Byte = 1) As Boolean
    Dim lngLoop As Long
    
    If bytMode = 1 Then
        mrsPersons.Filter = ""
        mrsPersons.Filter = "组别='" & strGroup & "' AND 删除<>'1'"
        If mrsPersons.RecordCount > 0 Then
            mrsPersons.MoveFirst
            Call FillGrid(vsfPerson, mrsPersons)
        End If
        
        For lngLoop = 1 To vsfPerson.Rows - 1
            
            If Val(vsfPerson.TextMatrix(lngLoop, mPersonCol.病人id)) = 0 Then
                vsfPerson.Cell(flexcpForeColor, lngLoop, 0, lngLoop, vsfPerson.Cols - 1) = COLOR.兰色
            End If
        Next
        
    End If
    
    If bytMode = 2 Then
        
        cmd(1).Tag = Val(mrsPersons("病人id").Value)
        
        txt(5).Text = mrsPersons("姓名").Value
        txt(4).Text = mrsPersons("身份证").Value
                
        txt(9).Text = mrsPersons("年龄").Value
        txt(10).Text = mrsPersons("联系人电话").Value
        
        zlControl.CboLocate cbo(1), zlCommFun.NVL(mrsPersons("性别").Value)
        zlControl.CboLocate cbo(0), zlCommFun.NVL(mrsPersons("婚姻状况").Value)
        
        txt(3).Text = zlCommFun.NVL(mrsPersons("门诊号").Value)
        txt(14).Text = zlCommFun.NVL(mrsPersons("健康号").Value)
        
        imgNew(0).Visible = (Val(cmd(1).Tag) = 0)
        
        txt(3).Locked = (Val(txt(3).Text) > 0 And Val(cmd(1).Tag) > 0)
        
    End If
    
    ReadPersons = True
    
End Function

Private Function ReadTemplate(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim rsPrice As New ADODB.Recordset
    
    Dim strKeys As String
    Dim bytParam1 As Byte
    Dim bytParam2 As Byte
    
    bytParam1 = 1
    bytParam2 = 2
            
    If mblnGroup = False Then
        Select Case zlCommFun.GetNeedName(cbo(1).Text)
        Case "男"
            bytParam1 = 1
            bytParam2 = 1
        Case "女"
            bytParam1 = 2
            bytParam2 = 2
        End Select
    End If
        
    On Error GoTo errHand
    
    gstrSQL = "SELECT DISTINCT A.ID,DECODE(A.类别,'C','检验','D','检查') AS 类别,A.编码,A.名称,C.名称 AS 体检类型,D.名称 As 采集方式,B.采集方式id,B.检验标本,B.检查部位,B.检查部位id " & _
                "FROM 诊疗项目目录 A,体检类型目录 B,体检类型 C,诊疗项目目录 D " & _
                "WHERE A.ID=B.诊疗项目ID AND C.序号=B.序号 AND D.ID(+)=B.采集方式id AND B.序号=[1] And Nvl(a.适用性别,0) In (0,[2],[3])"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, bytParam1, bytParam2)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            vsf.Row = vsf.Rows - 1
            If Val(vsf.RowData(vsf.Row)) > 0 Then
                vsf.Rows = vsf.Rows + 1
                vsf.Row = vsf.Rows - 1
            End If
            
            If CheckHave(rs("ID").Value) = False Then
            
                vsf.TextMatrix(vsf.Row, mCol.类别) = zlCommFun.NVL(rs("类别").Value)
                vsf.TextMatrix(vsf.Row, mCol.项目) = zlCommFun.NVL(rs("名称").Value)
                vsf.TextMatrix(vsf.Row, mCol.体检类型) = zlCommFun.NVL(rs("体检类型").Value)
                
                vsf.TextMatrix(vsf.Row, mCol.检验标本) = zlCommFun.NVL(rs("检验标本").Value)
                vsf.TextMatrix(vsf.Row, mCol.检查部位) = zlCommFun.NVL(rs("检查部位").Value)
                vsf.TextMatrix(vsf.Row, mCol.采集方式) = zlCommFun.NVL(rs("采集方式").Value)
                vsf.TextMatrix(vsf.Row, mCol.采集方式id) = zlCommFun.NVL(rs("采集方式id").Value)
                vsf.TextMatrix(vsf.Row, mCol.检查部位id) = zlCommFun.NVL(rs("检查部位id").Value)
                
                vsf.RowData(vsf.Row) = zlCommFun.NVL(rs("ID").Value)
            End If
                        
            If vsf.TextMatrix(vsf.Row, mCol.类别) = "检验" Then
                
                Call SetRowDefault(Val(vsf.RowData(vsf.Row)), vsf.Row, "执行科室")
                
                If Val(vsf.TextMatrix(vsf.Row, mCol.采集方式id)) = 0 Then
                    Call SetRowDefault(Val(vsf.RowData(vsf.Row)), vsf.Row, "采集方式")
                End If
                
                Call SetRowDefault(Val(vsf.RowData(vsf.Row)), vsf.Row, "采集科室")
                
                If Trim(vsf.TextMatrix(vsf.Row, mCol.检验标本)) = "" Then
                    Call SetRowDefault(Val(vsf.RowData(vsf.Row)), vsf.Row, "检验标本")
                End If
                
                Call SetRowDefault(Val(vsf.RowData(vsf.Row)), vsf.Row, "结算方式", "计价项目")
                                
            ElseIf vsf.TextMatrix(vsf.Row, mCol.类别) = "检查" Then
                Call SetRowDefault(Val(vsf.RowData(vsf.Row)), vsf.Row, "执行科室", "结算方式", "计价项目")
            End If
            
            gstrSQL = "Select z.数次,y.名称,y.计算单位,x.现价,x.现价*Nvl(z.折扣,1) As 体检单价,y.id,Nvl(z.计价性质,1) As 计价性质,y.类别,10*Nvl(z.折扣,1) As 折扣 " & _
                        "From " & _
                            "( Select a.序号,a.诊疗项目id,a.收费细目id,Sum(c.现价) As 现价 " & _
                              "From 收费价目 c, " & _
                                   "体检类型计价 a " & _
                              "Where a.收费细目id = c.收费细目id " & _
                                    "and c.执行日期<=SYSDATE and (c.终止日期 IS NULL OR c.终止日期>SYSDATE) " & _
                                    "and A.序号=[2] " & _
                                    "and A.诊疗项目id=[1] " & _
                              "Group by a.序号,a.诊疗项目id,a.收费细目id " & _
                            ") x, " & _
                            "收费项目目录 y, " & _
                            "体检类型计价 z " & _
                        "Where x.收费细目id = y.ID " & _
                              "and z.序号=x.序号 " & _
                              "and z.诊疗项目id=x.诊疗项目id " & _
                              "and z.收费细目id=x.收费细目id "
                        
            Set rsPrice = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)), lngKey)
            If rsPrice.BOF = False Then
                With vsfPrice
                    Do While Not rsPrice.EOF
                        
                        If Val(.TextMatrix(.Rows - 1, mCol.p收费项目id)) > 0 Then
                            .Rows = .Rows + 1
                        End If
                        
                        .TextMatrix(.Rows - 1, mCol.p名称) = zlCommFun.NVL(rsPrice("名称"))
                        .TextMatrix(.Rows - 1, mCol.p计算单位) = zlCommFun.NVL(rsPrice("计算单位"))
                        .TextMatrix(.Rows - 1, mCol.p数次) = zlCommFun.NVL(rsPrice("数次"))
                        .TextMatrix(.Rows - 1, mCol.p标准单价) = zlCommFun.NVL(rsPrice("现价"))
                        .TextMatrix(.Rows - 1, mCol.p体检单价) = zlCommFun.NVL(rsPrice("体检单价"))
                        .TextMatrix(.Rows - 1, mCol.p折扣) = zlCommFun.NVL(rsPrice("折扣"))
                        .TextMatrix(.Rows - 1, mCol.p标准金额) = zlCommFun.NVL(rsPrice("数次"), 0) * zlCommFun.NVL(rsPrice("现价"), 0)
                        .TextMatrix(.Rows - 1, mCol.p体检金额) = zlCommFun.NVL(rsPrice("数次"), 0) * zlCommFun.NVL(rsPrice("体检单价"), 0)
                        .TextMatrix(.Rows - 1, mCol.p收费项目id) = zlCommFun.NVL(rsPrice("ID"))
                        .TextMatrix(.Rows - 1, mCol.p计价性质) = zlCommFun.NVL(rsPrice("计价性质"))
                        .RowData(.Rows - 1) = zlCommFun.NVL(rsPrice("ID"), 0)
                        .TextMatrix(.Rows - 1, mCol.p类别) = zlCommFun.NVL(rsPrice("类别"))
                        
                        If zlCommFun.NVL(rsPrice("计价性质"), 1) = 2 Then
                            .TextMatrix(.Rows - 1, mCol.p计价项目) = "采集方式-" & Trim(vsf.TextMatrix(vsf.Row, mCol.采集方式))
                        ElseIf Trim(vsf.TextMatrix(vsf.Row, mCol.类别)) = "检验" Then
                            .TextMatrix(.Rows - 1, mCol.p计价项目) = "检验项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目))
                        Else
                            .TextMatrix(.Rows - 1, mCol.p计价项目) = "检查项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目))
                        End If
                        
                        Call SetRowDefault(Val(.RowData(.Rows - 1)), vsfPrice.Rows - 1, "收费执行科室")
                        
                        If InStr("567", .TextMatrix(.Rows - 1, mCol.p类别)) > 0 Then
                            .TextMatrix(.Rows - 1, mCol.p可用库存) = GetStorage(Val(.RowData(.Rows - 1)), Val(.TextMatrix(.Rows - 1, mCol.p执行科室id)))
                            Call PromptStorageWarn(Val(.TextMatrix(.Rows - 1, mCol.p数次)), Val(.TextMatrix(.Rows - 1, mCol.p可用库存)), .TextMatrix(.Rows - 1, mCol.p名称), .TextMatrix(.Rows - 1, mCol.p执行科室), .TextMatrix(.Rows - 1, mCol.p计算单位), 1)
                        End If
                        
                        
                        
                        rsPrice.MoveNext
                    Loop
                    
                    
                End With
                
                Call ChangePrice(Val(vsfPrice.TextMatrix(vsfPrice.Row, mCol.p标准单价)), Val(vsfPrice.TextMatrix(vsfPrice.Row, mCol.p体检单价)), 1)
                
                vsf.TextMatrix(vsf.Row, mCol.基本价格) = SumPrice(1)
                vsf.TextMatrix(vsf.Row, mCol.体检价格) = SumPrice(2)
                
            End If
            
            Call vsf_BeforeRowColChange(0, 0, vsf.Row, vsf.Col, False)
            Call vsfPrice_AfterRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col)
            Call WritePrice(vsf.Row)
                                    
            rs.MoveNext
        Loop
    End If
    
    ReadTemplate = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
        
    End If
End Sub

Private Sub chk_Click()
    
    txt(31).Enabled = (chk.Value = 1)
    txt(31).BackColor = IIf(chk.Value = 1, &H80000005, &H8000000F)
    If chk.Value <> 1 Then
        txt(31).Text = ""
    End If
    
End Sub

Private Sub chk_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim lngKey As Long
    Dim strDate As String
    Dim objPoint As POINTAPI
    Dim strTmp As String
    Dim strItem As String
    Dim strValue As String
    Dim strCardNo1 As String
    Dim strCardNo2 As String
    Dim rsPrice As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim lngLoop As Long
    Dim objItem As ListItem
    Dim intRow As Long
    Dim strKeys As String

    Dim clsCard As Object
    Dim strInfo() As String
    
    On Error GoTo errHand
    
    Call ClientToScreen(cmd(Index).hWnd, objPoint)
    
    Select Case Index
    '------------------------------------------------------------------------------------------------------------------
    Case 0
        
        strTmp = ""
        If frmInputBox.ShowInputBox(Me, "请输入体检组别", "输入新的体检组别名称，并为该组别添加体检项目及受检人员。", "组别(&G)", strTmp, 1, 20) Then
        
            '更改了组别名称,检查要点:1.该组名是否已经存在;2.修改人员及项目对应的组别名称
        
            For lngLoop = 1 To lvwGroup.ListItems.Count
                If lngLoop <> lvwGroup.SelectedItem.Index Then
                    If Trim(lvwGroup.ListItems(lngLoop).Text) = Trim(strTmp) Then
                        ShowSimpleMsg "“" & strTmp & "”组别已经存在！"
                        Exit Sub
                    End If
                End If
            Next
            
            Set objItem = lvwGroup.ListItems.Add(, , strTmp, 1, 1)
            objItem.Selected = True
        
            Call lvwGroup_ItemClick(objItem)

        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        If lvwGroup.SelectedItem Is Nothing Then Exit Sub
        
        strTmp = Trim(lvwGroup.SelectedItem.Text)
        If frmInputBox.ShowInputBox(Me, "体检分组", "修改已经存在的体检组别名称。", "新名称(&G)", strTmp, 1, 20) Then
        
            '更改了组别名称,检查要点:1.该组名是否已经存在;2.修改人员及项目对应的组别名称
            
            If Trim(strTmp) = Trim(lvwGroup.SelectedItem.Text) Then Exit Sub
            
            For lngLoop = 1 To lvwGroup.ListItems.Count
                If lngLoop <> lvwGroup.SelectedItem.Index Then
                    If Trim(lvwGroup.ListItems(lngLoop).Text) = Trim(strTmp) Then
                        ShowSimpleMsg "“" & strTmp & "”组别已经存在！"
                        Exit Sub
                    End If
                End If
            Next
            
            '2.修改人员及项目对应的组别名称
            Call WritePrice(vsf.Row)
            Call SaveItems(lvwGroup.SelectedItem.Text)
            Call WritePersons(lvwGroup.SelectedItem.Text)
            
            mrsItems.Filter = ""
            mrsItems.Filter = "组别='" & lvwGroup.SelectedItem.Text & "'"
            If mrsItems.RecordCount > 0 Then
                mrsItems.MoveFirst
                Do While Not mrsItems.EOF
                    mrsItems("组别").Value = strTmp
                    mrsItems.MoveNext
                Loop
            End If
            mrsItems.Filter = ""
        
            mrsPersons.Filter = ""
            mrsPersons.Filter = "组别='" & lvwGroup.SelectedItem.Text & "'"
            If mrsPersons.RecordCount > 0 Then
                mrsPersons.MoveFirst
                Do While Not mrsPersons.EOF
                    mrsPersons("组别").Value = strTmp
                    mrsPersons.MoveNext
                Loop
            End If
            mrsPersons.Filter = ""
                    
            lvwGroup.SelectedItem.Text = strTmp
            
            Call ResetVsf(vsf)
            Call ResetVsf(vsfPrice)
            Call ResetVsf(vsfPerson)
            
            mstrGroup = lvwGroup.SelectedItem.Text
            
            Call ReadItems(mstrGroup)
            Call ReadPersons(mstrGroup)
            Call ReadPrice(vsf.Row)
            Call CountGroup
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 3
        
        If lvwGroup.ListItems.Count = 1 Then
            ShowSimpleMsg "团体体检时至少需要一个组别！"
            Exit Sub
        End If
    
        If MsgBox("删除组别时，将自动处理以下信息：" & vbCrLf & "  1.删除对应的体检项目" & vbCrLf & "  2.删除后需要重新设置其组别或人员划分" & vbCrLf & "继续吗？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
                
        mrsItems.Filter = ""
        mrsItems.Filter = "组别='" & lvwGroup.SelectedItem.Text & "'"
        If mrsItems.RecordCount > 0 Then
            mrsItems.MoveFirst
            Call DeleteRecord(mrsItems)
        End If
        
        Call WritePersons(lvwGroup.SelectedItem.Text)
        
        mrsPersons.Filter = ""
        mrsPersons.Filter = "组别='" & lvwGroup.SelectedItem.Text & "'"
        If mrsPersons.RecordCount > 0 Then
            mrsPersons.MoveFirst
            Do While Not mrsPersons.EOF
                If lvwGroup.SelectedItem.Index = 1 Then
                    mrsPersons("组别").Value = lvwGroup.ListItems(2).Text
                Else
                    mrsPersons("组别").Value = lvwGroup.ListItems(1).Text
                End If
                mrsPersons.MoveNext
            Loop
        End If
        mrsPersons.Filter = ""
    
        Call ResetVsf(vsfPerson)
        Call FillGrid(vsfPerson, mrsPersons)
        
        lngLoop = lvwGroup.SelectedItem.Index
        lvwGroup.ListItems.Remove lngLoop
        Call NextLvwPos(lvwGroup, lngLoop)
        
        If Not (lvwGroup.SelectedItem Is Nothing) Then
            mstrGroup = lvwGroup.SelectedItem.Text
        
            Call ReadItems(mstrGroup)
            Call ReadPersons(mstrGroup)
            Call ReadPrice(vsf.Row)
            
        End If
        Call CountGroup
    '------------------------------------------------------------------------------------------------------------------
    Case 1      '打开病人查找对话框
        If frmPatientFind.ShowFind(Me, lngKey) Then
            If lngKey > 0 Then
                
                gstrSQL = "SELECT A.* FROM 病人信息 A WHERE A.病人id=[1]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
                If rs.BOF = False Then
                    cmd(1).Tag = zlCommFun.NVL(rs("病人id").Value)
                    txt(5).Text = zlCommFun.NVL(rs("姓名").Value)
                    txt(4).Text = zlCommFun.NVL(rs("身份证号").Value)
                    txt(9).Text = zlCommFun.NVL(rs("年龄").Value)
                    
                    txt(3).Text = zlCommFun.NVL(rs("门诊号").Value)
                    txt(14).Text = zlCommFun.NVL(rs("健康号").Value)
                    
                    zlControl.CboLocate cbo(1), zlCommFun.NVL(rs("性别").Value)
                    zlControl.CboLocate cbo(0), zlCommFun.NVL(rs("婚姻状况").Value)
                    
                    Call FillPatient(Val(cmd(1).Tag))
                    
                    DataChange = True
                    
                    txt(5).Tag = ""
                    imgNew(1).Visible = False
                    
                    txt(3).Locked = (Val(txt(3).Text) > 0 And Val(cmd(1).Tag) > 0)
                    
                End If
                
            End If
        End If
        
        LocationObj txt(5)
    '------------------------------------------------------------------------------------------------------------------
    Case 4      '打开团体(合同单位)选择器
        lngKey = Val(cmd(Index).Tag)
        gstrSQL = GetPublicSQL(SQL.体检团体选择)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If ShowTxtSelect(Me, txt(13), "编码,900,0,1;名称,1500,0,1;简码,900,0,1;地址,3000,0,1", Me.Name & "\体检团体选择", "请在下表中选择一个团体/单位。", rsData, rs, 8790, 5100) Then
              
            Call ReadGroup(zlCommFun.NVL(rs("ID").Value, 0))
                        
            If lngKey <> Val(cmd(Index).Tag) Then DataChange = True

            imgNew(0).Visible = False
            
            txt(0).Text = txt(12).Text
            txt(1).Text = txt(11).Text
            
        End If
        
        LocationObj txt(13)
    '------------------------------------------------------------------------------------------------------------------
    Case 5
    
        Dim bytParam1 As Byte
        Dim bytParam2 As Byte
        
        bytParam1 = 1
        bytParam2 = 2
                
        If mblnGroup = False Then
            Select Case zlCommFun.GetNeedName(cbo(1).Text)
            Case "男"
                bytParam1 = 1
                bytParam2 = 1
            Case "女"
                bytParam1 = 2
                bytParam2 = 2
            End Select
        End If
            
        gstrSQL = GetPublicSQL(SQL.体检项目选择)
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, bytParam1, bytParam2)
        
        If ShowTxtSelect(Me, cmd(Index), "编码,1200,0,1;名称,2700,0,0;单位,900,0,0;类别,900,0,0", Me.Name & "\体检项目选择", "请从列表中选择一个体检项目。", rsData, rs, 8790, 5100, True) Then

            rs.Filter = 0
            rs.Filter = "选择=1"
            If rs.RecordCount > 0 Then

                rs.MoveFirst
                Do While Not rs.EOF
                    '选取了一个项目
                    vsf.Row = 0

                    If CheckHave(zlCommFun.NVL(rs("ID").Value)) = False Then

                        If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                            vsf.Rows = vsf.Rows + 1
                        End If
                        intRow = vsf.Rows - 1
                        vsf.Row = vsf.Rows - 1

                        vsf.Cell(flexcpText, intRow, mCol.项目 + 1, intRow, vsf.Cols - 1) = ""

                        vsf.TextMatrix(intRow, mCol.类别) = zlCommFun.NVL(rs("类别").Value)
                        vsf.TextMatrix(intRow, mCol.项目) = zlCommFun.NVL(rs("名称").Value)
                        vsf.RowData(intRow) = zlCommFun.NVL(rs("ID").Value)

                        If vsf.TextMatrix(intRow, mCol.类别) = "检验" Then
                            Call SetRowDefault(Val(vsf.RowData(intRow)), intRow, "执行科室", "采集方式", "采集科室", "检验标本", "结算方式", "计价项目")
                            
                        ElseIf vsf.TextMatrix(intRow, mCol.类别) = "检查" Then
                            Call SetRowDefault(Val(vsf.RowData(intRow)), intRow, "执行科室", "结算方式", "计价项目")
                        End If
                            
                        Call CreatePriceList(intRow)
                        Call vsf_BeforeRowColChange(0, 0, vsf.Row, vsf.Col, False)
                        Call vsfPrice_AfterRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col)
                        
                        Call WritePrice(intRow)

                        DataChange = True
                    End If

                    rs.MoveNext
                Loop
            End If

        End If

        EnterFocus vsf
    '------------------------------------------------------------------------------------------------------------------
    Case 6
    
        gstrSQL = GetPublicSQL(SQL.体检类型分类选择)

        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, IIf(mblnGroup, 2, 1))
    
        If ShowTxtSelect(Me, cmd(Index), "编码,1080,0,1;名称,2400,0,0;简码,900,0,0;说明,1500,0,0", Me.Name & "\体检类型选择", "请从列表中选择一个体检类型。", rsData, rs, 8790, 5100, True) Then

            rs.Filter = 0
            rs.Filter = "选择=1"
            If rs.RecordCount > 0 Then

                If Val(vsf.RowData(1)) > 0 Then
                    If MsgBox("是否要清除已选择的体检项目？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        Call ResetVsf(vsf)
                        Call ResetVsf(vsfPrice)
                    End If
                End If

                rs.MoveFirst

                Do While Not rs.EOF

                    Call ReadTemplate(rs("ID").Value)
                    rs.MoveNext

                Loop

                DataChange = True
            End If

        End If

        EnterFocus vsf
    '------------------------------------------------------------------------------------------------------------------
    Case 8
    
        Dim strParam As String
        Dim varParam As Variant

        strParam = vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.病人id) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.姓名) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份证) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.性别) & "'"
        strParam = strParam & Format(vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.出生日期), "yyyy-MM-dd") & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.婚姻状况) & "'"

        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.民族) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.国籍) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.学历) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.职业) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人姓名) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人电话) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.电子邮件) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人地址) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.工作单位) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.年龄) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.健康号)
        
        If frmPatientEdit.ShowEdit(Me, strParam) Then
            varParam = Split(strParam, "'")

            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.姓名) = varParam(1)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份证) = varParam(2)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.性别) = varParam(3)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.出生日期) = varParam(4)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.婚姻状况) = varParam(5)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.病人id) = Val(varParam(0))

            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.民族) = varParam(6)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.国籍) = varParam(7)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.学历) = varParam(8)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.职业) = varParam(9)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份) = varParam(10)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人姓名) = varParam(11)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人电话) = varParam(12)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.电子邮件) = varParam(13)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人地址) = varParam(14)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.工作单位) = varParam(15)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.年龄) = varParam(16)
            
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.健康号) = varParam(17)
            
'            imgNew(1).Visible = False

        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 9
    
        On Error GoTo 0

        dlg.CancelError = True

        On Error GoTo ErrHandler

        dlg.Flags = &H4 Or &H200000 Or &H800 & &H1000
        dlg.Filter = "体检资料(*.xls)| *.xls"
        dlg.FilterIndex = 0

        dlg.DialogTitle = "体检资料收集"
        dlg.FileName = App.Path & "\体检资料收集.xls"
        dlg.ShowOpen

        If Dir(dlg.FileName) <> "" Then
            If ReadExcelFile(dlg.FileName) Then
                If Not (lvwGroup.SelectedItem Is Nothing) Then
                    mstrGroup = ""
                    Call lvwGroup_ItemClick(lvwGroup.SelectedItem)
                End If
                'Call lvwGroup_AfterDeleteRow(lvwGroup.Row, lvwGroup.Col)
            End If
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 10
                
        Call SaveGroupInfo
        
        If frmGroupEdit.ShowEdit(Me, mrsGroup) Then
            Call ShowGroupInfo
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 11
            
        Call WritePersons("缺省", 2)
        
        strParam = ""
        strParam = strParam & mrsPersons("病人id").Value & "'"
        strParam = strParam & mrsPersons("姓名").Value & "'"
        strParam = strParam & mrsPersons("身份证").Value & "'"
        strParam = strParam & mrsPersons("性别").Value & "'"
        strParam = strParam & mrsPersons("出生日期").Value & "'"
        strParam = strParam & mrsPersons("婚姻状况").Value & "'"
        strParam = strParam & mrsPersons("民族").Value & "'"
        strParam = strParam & mrsPersons("国籍").Value & "'"
        strParam = strParam & mrsPersons("学历").Value & "'"
        strParam = strParam & mrsPersons("职业").Value & "'"
        strParam = strParam & mrsPersons("身份").Value & "'"
        strParam = strParam & mrsPersons("联系人姓名").Value & "'"
        strParam = strParam & mrsPersons("联系人电话").Value & "'"
        strParam = strParam & mrsPersons("电子邮件").Value & "'"
        strParam = strParam & mrsPersons("联系人地址").Value & "'"
        strParam = strParam & mrsPersons("工作单位").Value & "'"
        strParam = strParam & mrsPersons("年龄").Value & "'"
        strParam = strParam & mrsPersons("健康号").Value
        
        
        If frmPatientEdit.ShowEdit(Me, strParam) Then
            varParam = Split(strParam, "'")
            
            mrsPersons("姓名").Value = varParam(1)
            mrsPersons("身份证").Value = varParam(2)
            mrsPersons("性别").Value = varParam(3)
            mrsPersons("出生日期").Value = varParam(4)
            mrsPersons("婚姻状况").Value = varParam(5)
            mrsPersons("民族").Value = varParam(6)
            mrsPersons("国籍").Value = varParam(7)
            mrsPersons("学历").Value = varParam(8)
            mrsPersons("职业").Value = varParam(9)
            mrsPersons("身份").Value = varParam(10)
            mrsPersons("联系人姓名").Value = varParam(11)
            mrsPersons("联系人电话").Value = varParam(12)
            mrsPersons("电子邮件").Value = varParam(13)
            mrsPersons("联系人地址").Value = varParam(14)
            mrsPersons("工作单位").Value = varParam(15)
            mrsPersons("年龄").Value = varParam(16)
            mrsPersons("健康号").Value = varParam(17)
                                    
            Call ReadPersons("缺省", 2)
            
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 12, 14    '写卡
        
        Set clsCard = CreateObject("zl9ICCard.clsICCard")
        If Not (clsCard Is Nothing) Then
            
            ReDim strInfo(1 To 16)
            
            strCardNo1 = clsCard.GetCardNo
            
            strCardNo2 = vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC卡号)
            
            If strCardNo2 <> "" Then
                '病人有卡，但和当前的卡不是同一张卡
                If strCardNo1 <> strCardNo2 Then
                    ShowSimpleMsg "此卡不是当前病人的卡！"
                    Exit Sub
                End If
            Else
                '病人没有卡
                
                If strCardNo1 = "" Then
                
                    '新卡，自动开卡
                    strCardNo1 = "11111111"
                    strCardNo2 = strCardNo1
                    
                    '写卡号
                    If clsCard.SetCardNo(strCardNo1) = False Then Exit Sub
                    vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC卡号) = strCardNo2
                    
                Else
                
                    '不是新卡
                    ShowSimpleMsg "此卡不是新卡，不能进行写入操作！"
                    Exit Sub
                    
                End If
                                                
            End If
            
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC卡号) = strCardNo2
            
            If mblnGroup Then
                strInfo(1) = "姓名=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.姓名)
                strInfo(2) = "身份证号=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份证)
                strInfo(3) = "性别=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.性别)
                strInfo(4) = "出生日期=" & Format(vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.出生日期), "yyyy-MM-dd")
                strInfo(5) = "婚姻状况=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.婚姻状况)
                strInfo(6) = "民族=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.民族)
                strInfo(7) = "国籍=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.国籍)
                strInfo(8) = "学历=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.学历)
                strInfo(9) = "职业=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.职业)
                strInfo(10) = "身份=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份)
                strInfo(11) = "联系人姓名=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人姓名)
                strInfo(12) = "联系人电话=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人电话)
                strInfo(13) = "联系人地址=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人地址)
                strInfo(14) = "电子邮件=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.电子邮件)
                strInfo(15) = "工作单位=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.工作单位)
                strInfo(16) = "年龄=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.年龄)
            Else
                strInfo(1) = "姓名=" & mrsPersons("姓名").Value
                strInfo(2) = "身份证号=" & mrsPersons("身份证").Value
                strInfo(3) = "性别=" & mrsPersons("性别").Value
                strInfo(4) = "出生日期=" & mrsPersons("出生日期").Value
                strInfo(5) = "婚姻状况=" & mrsPersons("婚姻状况").Value
                strInfo(6) = "民族=" & mrsPersons("民族").Value
                strInfo(7) = "国籍=" & mrsPersons("国籍").Value
                strInfo(8) = "学历=" & mrsPersons("学历").Value
                strInfo(9) = "职业=" & mrsPersons("职业").Value
                strInfo(10) = "身份=" & mrsPersons("身份").Value
                strInfo(11) = "联系人姓名=" & mrsPersons("联系人姓名").Value
                strInfo(12) = "联系人电话=" & mrsPersons("联系人电话").Value
                strInfo(13) = "联系人地址=" & mrsPersons("联系人地址").Value
                strInfo(14) = "电子邮件=" & mrsPersons("电子邮件").Value
                strInfo(15) = "工作单位=" & mrsPersons("工作单位").Value
                strInfo(16) = "年龄=" & mrsPersons("年龄").Value
            End If
                        
            If clsCard.SetPatient(strInfo) Then
                ShowSimpleMsg "更新当前病人信息成功！"
            End If
        End If
        
        If mblnGroup Then
            If vsfPerson.Visible Then vsfPerson.SetFocus
        Else
            If txt(5).Visible Then txt(5).SetFocus
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 13, 15    '读卡
        
        Set clsCard = CreateObject("zl9ICCard.clsICCard")
        If Not (clsCard Is Nothing) Then
            
            If mblnGroup = False Then Call WritePersons("缺省", 2)
            
            strCardNo1 = clsCard.GetCardNo
            strCardNo2 = vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC卡号)
            
            If strCardNo2 <> "" Then
                '记录的病人有卡，但和当前的卡不是同一张卡
                If strCardNo1 <> strCardNo2 Then
                    ShowSimpleMsg "此卡不是当前病人的卡！"
                    Exit Sub
                End If
            Else
            
                '病人没有卡，则将当前的卡号付给病人
                strCardNo2 = strCardNo1
                                
            End If
            
            If mblnGroup = False Then
                mrsPersons("IC卡号").Value = strCardNo2
            Else
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC卡号) = strCardNo2
            End If
            
            
            If GetPatientID(strCardNo2) > 0 Then
                
                '在系统中找到了病人
                If mblnGroup Then
                    vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.病人id) = GetPatientID(strCardNo2)
                    Call GetPatientInfo(Val(vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.病人id)))
                Else
                    mrsPersons("IC卡号").Value = GetPatientID(strCardNo2)
                    Call GetPatientInfo(Val(mrsPersons("IC卡号").Value))
                End If
                
            ElseIf clsCard.GetPatient(strInfo) Then
                    
                For lngLoop = LBound(strInfo) To UBound(strInfo)
                    If InStr(strInfo(lngLoop), "=") > 0 Then
                        strItem = Mid(strInfo(lngLoop), 1, InStr(strInfo(lngLoop), "=") - 1)
                        strValue = Mid(strInfo(lngLoop), InStr(strInfo(lngLoop), "=") + 1)
                        
                        Select Case strItem
                        Case "姓名"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.姓名) = strValue
                            Else
                                mrsPersons("姓名").Value = strValue
                            End If
                            
                        Case "身份证号"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份证) = strValue
                            Else
                                mrsPersons("身份证").Value = strValue
                            End If
                            
                        Case "性别"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.性别) = strValue
                            Else
                                mrsPersons("性别").Value = strValue
                            End If
                            
                        Case "出生日期"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.出生日期) = strValue
                            Else
                                mrsPersons("出生日期").Value = strValue
                            End If
                            
                        Case "婚姻状况"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.婚姻状况) = strValue
                            Else
                                mrsPersons("婚姻状况").Value = strValue
                            End If
                            
                        Case "民族"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.民族) = strValue
                            Else
                                mrsPersons("民族").Value = strValue
                            End If
                            
                        Case "国籍"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.国籍) = strValue
                            Else
                                mrsPersons("国籍").Value = strValue
                            End If
                            
                        Case "学历"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.学历) = strValue
                            Else
                                mrsPersons("学历").Value = strValue
                            End If
                            
                        Case "职业"
                        
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.职业) = strValue
                            Else
                                mrsPersons("职业").Value = strValue
                            End If
                            
                        Case "身份"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份) = strValue
                            Else
                                mrsPersons("身份").Value = strValue
                            End If
                            
                        Case "联系人姓名"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人姓名) = strValue
                            Else
                                mrsPersons("联系人姓名").Value = strValue
                            End If
                            
                        Case "联系人电话"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人电话) = strValue
                            Else
                                mrsPersons("联系人电话").Value = strValue
                            End If
                            
                        Case "联系人地址"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人地址) = strValue
                            Else
                                mrsPersons("联系人地址").Value = strValue
                            End If
                            
                        Case "电子邮件"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.电子邮件) = strValue
                            Else
                                mrsPersons("电子邮件").Value = strValue
                            End If
                            
                        Case "工作单位"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.工作单位) = strValue
                            Else
                                mrsPersons("工作单位").Value = strValue
                            End If
                            
                        Case "年龄"
                            
                            If mblnGroup Then
                                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.年龄) = strValue
                            Else
                                mrsPersons("年龄").Value = strValue
                            End If
                            
                        End Select
                    End If
                Next
                
                If mblnGroup = False Then Call ReadPersons("缺省", 2)
            End If
        End If
        
        If mblnGroup Then
            If vsfPerson.Visible Then vsfPerson.SetFocus
        Else
            If txt(5).Visible Then txt(5).SetFocus
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 16
        '选择单位人员

        If frmSelectGroupPerson.ShowFilter(Me, Val(cmd(4).Tag), rs) Then
        
            rs.Filter = 0
            rs.Filter = "选择=1"
            If rs.RecordCount > 0 Then

                If Val(vsfPerson.RowData(1)) > 0 Then
                    If MsgBox("是否要清除已选择的受检人员？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        Call ResetVsf(vsfPerson)
                    End If
                End If

                rs.MoveFirst

                Do While Not rs.EOF

                    If CheckHavePerson(rs("ID").Value) = False Then
                        With vsfPerson

                            .Row = .Rows - 1
                            If Val(.RowData(.Row)) > 0 Then
                                .Rows = .Rows + 1
                                .Row = .Rows - 1
                            End If

                            .TextMatrix(.Row, mPersonCol.姓名) = zlCommFun.NVL(rs("姓名").Value)
                            .TextMatrix(.Row, mPersonCol.门诊号) = zlCommFun.NVL(rs("门诊号").Value)
                            .TextMatrix(.Row, mPersonCol.健康号) = zlCommFun.NVL(rs("健康号").Value)
                            .TextMatrix(.Row, mPersonCol.性别) = zlCommFun.NVL(rs("性别").Value)
                            .TextMatrix(.Row, mPersonCol.年龄) = zlCommFun.NVL(rs("年龄").Value)
                            .TextMatrix(.Row, mPersonCol.婚姻状况) = zlCommFun.NVL(rs("婚姻状况").Value)
                            .TextMatrix(.Row, mPersonCol.出生日期) = zlCommFun.NVL(rs("出生日期").Value)
                            .TextMatrix(.Row, mPersonCol.身份证) = zlCommFun.NVL(rs("身份证号").Value)
                            .TextMatrix(.Row, mPersonCol.民族) = zlCommFun.NVL(rs("民族").Value)
                            .TextMatrix(.Row, mPersonCol.国籍) = zlCommFun.NVL(rs("国籍").Value)
                            .TextMatrix(.Row, mPersonCol.学历) = zlCommFun.NVL(rs("学历").Value)
                            .TextMatrix(.Row, mPersonCol.职业) = zlCommFun.NVL(rs("职业").Value)
                            .TextMatrix(.Row, mPersonCol.身份) = zlCommFun.NVL(rs("身份").Value)
                            .TextMatrix(.Row, mPersonCol.联系人姓名) = zlCommFun.NVL(rs("联系人姓名").Value)
                            .TextMatrix(.Row, mPersonCol.联系人电话) = zlCommFun.NVL(rs("联系人电话").Value)
'                            .TextMatrix(.Row, mPersonCol.电子邮件) = zlCommFun.NVL(rs("电子邮件").Value)
                            .TextMatrix(.Row, mPersonCol.联系人地址) = zlCommFun.NVL(rs("联系人地址").Value)
                            .TextMatrix(.Row, mPersonCol.工作单位) = zlCommFun.NVL(rs("工作单位").Value)
                            .TextMatrix(.Row, mPersonCol.病人id) = zlCommFun.NVL(rs("ID").Value, 0)
                            .TextMatrix(.Row, mPersonCol.IC卡号) = zlCommFun.NVL(rs("IC卡号").Value)
                            .TextMatrix(.Row, mPersonCol.就诊卡号) = zlCommFun.NVL(rs("就诊卡号").Value)

                            .RowData(.Row) = zlCommFun.NVL(rs("ID").Value)
                        End With
                    End If

                    rs.MoveNext

                Loop
                Call CountGroup
                DataChange = True
            End If
        End If
        
        Call EnterFocus(vsfPerson)
        
    '------------------------------------------------------------------------------------------------------------------
    Case 17
        
        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(lngLoop)) > 0 Then
                vsf.TextMatrix(lngLoop, mCol.结算方式) = "记帐"
            End If
        Next
        
    '------------------------------------------------------------------------------------------------------------------
    Case 18
        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(lngLoop)) > 0 Then
                vsf.TextMatrix(lngLoop, mCol.结算方式) = "收费"
            End If
        Next
        
ErrHandler:


    End Select
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub cmdOK_Click()
    
    Dim lngKey As Long
    
    If Trim(lvwGroup.SelectedItem.Text) <> "" Then
        
        Call WritePrice(vsf.Row)
        Call SaveItems(Trim(lvwGroup.SelectedItem.Text))
        
        If mblnGroup = False Then
            Call WritePersons("缺省", 2)
        Else
            Call WritePersons(Trim(lvwGroup.SelectedItem.Text))
        End If
    End If
    
    Call SaveGroupInfo
    
    If ValidEdit = False Then Exit Sub
    
    If SaveEdit(lngKey) Then
        
        mblnOK = True
                
        If mlngKey = 0 And mbytMode = 1 Then
        
            Call ClearData
            
            lvwGroup.ListItems.Add , , "缺省"
            
            DataChange = False
            
            ShowSimpleMsg "预约登记成功，继续下一个预约！"
            
            If mblnGroup Then
                Call LocationObj(txt(13))
            Else
                Call LocationObj(txt(5))
            End If
        
        Else
            DataChange = False
            Unload Me
        End If
    End If
End Sub


Private Sub dtp_Change(Index As Integer)
    DataChange = True
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        
    ElseIf Shift = 0 Then
        Select Case KeyCode
        Case vbKeyF3
        
            If cmd(5).Enabled And tbs.Tab = 0 Then
                Call cmd_Click(5)
            End If
            
        Case vbKeyF4
        
            If cmd(6).Enabled And tbs.Tab = 0 Then
                Call cmd_Click(6)
            End If
                
                
        Case vbKeyF5
        
            If cmd(7).Enabled And tbs.Tab = 0 Then
                Call cmd_Click(7)
            End If
            
        Case vbKeyF6
        
            If cmd(8).Enabled And tbs.Tab = 1 Then
                Call cmd_Click(8)
            End If
            
        Case vbKeyF7
        
            If cmd(9).Enabled And tbs.Tab = 1 Then
                Call cmd_Click(9)
            End If
            
        Case vbKeyF8
        
            If cmd(0).Enabled And mblnGroup Then
                Call cmd_Click(0)
            End If
            
        Case vbKeyF9
    
            If cmd(2).Enabled And mblnGroup Then
                Call cmd_Click(2)
            End If
            
        Case vbKeyF10
        
            If cmd(3).Enabled And mblnGroup Then
                Call cmd_Click(3)
            End If
            
        Case vbKeyF11
        
            If cmd(11).Enabled And mblnGroup = False Then
                Call cmd_Click(11)
            End If
            
        Case vbKeyF12
        
            If cmd(10).Enabled And mblnGroup Then
                Call cmd_Click(10)
            End If
        End Select
    End If
End Sub

Private Sub Form_Load()
    glngFormW = 12000
    glngFormH = 7710
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    
    Call RestoreWinState(Me, App.ProductName)
    
End Sub

Private Sub Form_Resize()
    Dim lngY As Long
    
    On Error Resume Next
    
    With fraTitle
        .Left = 0
        .Top = -90
        .Width = Me.ScaleWidth - .Left
    End With
    
    With fraGroupInfo
        .Left = fraTitle.Left
        .Top = fraTitle.Top + fraTitle.Height - 90
        .Width = fraTitle.Width
    End With
    
    If fraGroupInfo.Visible Then lngY = fraGroupInfo.Top + fraGroupInfo.Height
    
    With fraSingle
        .Left = fraGroupInfo.Left
        .Top = fraGroupInfo.Top
        .Width = fraGroupInfo.Width
    End With
    
    If fraSingle.Visible Then lngY = fraSingle.Top + fraSingle.Height
    
    With fraInfo
        .Left = fraGroupInfo.Left
        .Top = fraGroupInfo.Top + IIf(mblnGroup, fraGroupInfo.Height, fraSingle.Height) - 90
        .Width = fraGroupInfo.Width
    End With
    
    If fraInfo.Visible Then lngY = fraInfo.Top + fraInfo.Height
    
    With fraGroup
        .Left = fraInfo.Left
        .Top = lngY + 60
        .Height = Me.ScaleHeight - .Top - fraOther.Height - picButton.Height + 90 - stbThis.Height
    End With
    
    With tbs
        .Left = IIf(fraGroup.Visible, fraGroup.Left + fraGroup.Width + 45, 0)
        .Top = fraGroup.Top
        .Width = Me.ScaleWidth - .Left - 45
        .Height = fraGroup.Height
    End With
    
    With fraOther
        .Left = fraInfo.Left
        .Top = tbs.Top + tbs.Height - 90
        .Width = fraInfo.Width
    End With
                
    With cmd(11)
        .Left = fraSingle.Width - .Width - 60
    End With
    
    With cmd(14)
        .Left = cmd(11).Left - .Width - 45
    End With
    
    With cmd(15)
        .Left = cmd(14).Left - .Width - 45
    End With
    
    With cmd(10)
        .Left = fraGroupInfo.Width - .Width - 60
    End With
                
    With picButton
        .Left = fraOther.Left
        .Top = fraOther.Top + fraOther.Height
        .Width = fraOther.Width
    End With
    
    With picNo
        .Left = fraTitle.Width - .Width - 45
    End With
    
    With txt(2)
        .Width = fraInfo.Width - .Left - 45
    End With
    
    With lvwGroup
        .Left = 75
        .Top = 225
        .Width = fraGroup.Width - .Left - 75
        .Height = fraGroup.Height - .Top - 60 - cmd(0).Height - 60
    End With
    
    cmd(0).Top = lvwGroup.Top + lvwGroup.Height + 60
    cmd(2).Top = cmd(0).Top
    cmd(3).Top = cmd(0).Top
    
    If mblnGroup Then
        tbs.Tab = 1
        With vsfPerson
            .Left = 90
            .Top = 450
            .Width = tbs.Width - .Left - 90
            .Height = tbs.Height - .Top - 90
        End With
        With cmd(9)
            .Left = tbs.Width - .Width
            .Top = 0
        End With

        With cmd(16)
            .Left = cmd(9).Left - .Width - 45
            .Top = cmd(9).Top
        End With
        
        With cmd(8)
            .Left = cmd(16).Left - .Width - 45
            .Top = cmd(16).Top
        End With
        
        With cmd(12)
            .Left = cmd(8).Left - .Width - 45
            .Top = cmd(8).Top
        End With
        
        
        With cmd(13)
            .Left = cmd(12).Left - .Width - 45
            .Top = cmd(12).Top
        End With
        
    End If
    
    tbs.Tab = 0
    With vsf
        .Left = 90
        .Top = 450 + 300
        .Width = tbs.Width - .Left - 90
        .Height = tbs.Height - .Top - 90 - vsfPrice.Height - 30
    End With
    
    With vsfPrice
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height + 30
        .Width = vsf.Width
    End With
    
    With cmd(6)
        .Left = tbs.Width - .Width
        .Top = 0
    End With

    With cmd(5)
        .Left = cmd(6).Left - .Width - 45
        .Top = cmd(6).Top
    End With
    
    With cmd(17)
        .Left = cmd(5).Left - .Width - 45
        .Top = cmd(6).Top
    End With
    
    With cmd(18)
        .Left = cmd(17).Left - .Width - 45
        .Top = cmd(6).Top
    End With
    
    txt(6).Width = fraOther.Width - txt(6).Left - 45
        
    cmdCancel.Left = picButton.Width - cmdCancel.Width - 60
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If DataChange Then
        Cancel = (MsgBox("数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Call SaveWinState(Me, App.ProductName)
    
    If Not InDesign Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    
End Sub

Private Sub lvwGroup_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If mstrGroup <> Item.Text Then

        On Error Resume Next
        
        Call WritePrice(vsf.Row)
        Call SaveItems(mstrGroup)
        Call WritePersons(mstrGroup)
        
        Call ResetVsf(vsf)
        Call ResetVsf(vsfPerson)
        Call ResetVsf(vsfPrice)
        
        mstrGroup = Item.Text
        
        Call ReadItems(mstrGroup)
        Call ReadPrice(vsf.Row)
        
        Call ReadPersons(mstrGroup)
        
        Call CountGroup
    End If
    
End Sub

Private Function ReadExcelFile(ByVal strExcelFile As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------------------
    Dim objExcel As Object
    Dim ExWorkbook As Object
    Dim ExWorkSheet As Object
    Dim lngLoop As Long
    Dim lngLoop2 As Long
    Dim rsTmp As New ADODB.Recordset
    Dim str姓名 As String
    Dim str身份证号 As String
    Dim lng病人id As Long
    Dim str组别 As String
    Dim str门诊号 As String
    
    On Error GoTo errHand
    
    frmWait.OpenWait Me, "导入体检人员"
    frmWait.WaitInfo = "正在打开""" & strExcelFile & """..."
    
    Set objExcel = CreateObject("Excel.Application")
    Set ExWorkbook = Nothing
    Set ExWorkSheet = Nothing
    
    Set ExWorkbook = objExcel.Workbooks.Open(strExcelFile)
    If ExWorkbook Is Nothing Then Exit Function
    
    Set ExWorkSheet = ExWorkbook.Worksheets("人员资料")
    If ExWorkSheet Is Nothing Then Exit Function
    
    '先删除
    mrsPersons.Filter = ""
    
    Call CopyRecord(mrsPersons, rsTmp)
    Call DeleteRecord(rsTmp)
        
    For lngLoop = 4 To ExWorkSheet.UsedRange.Cells.Rows.Count
        
        str姓名 = Trim(ExWorkSheet.Range(Chr(mColChar.姓名) & lngLoop).Value)
        
        If Trim(str姓名) = "" Then
            frmWait.WaitInfo = "正在搜索人员资料..."
        Else
            frmWait.WaitInfo = "正在导入""" & str姓名 & """资料..."
        
            str身份证号 = Trim(ExWorkSheet.Range(Chr(mColChar.身份证号) & lngLoop).Value)
            str组别 = ExWorkSheet.Range(Chr(mColChar.体检组) & lngLoop).Value
            
            '按姓名和身份证查找人员档案
            str门诊号 = ""
            lng病人id = 0
            Call SearchArchive(str身份证号, lng病人id, str门诊号)
            If Val(str门诊号) = 0 Then
                str门诊号 = Val(ExWorkSheet.Range(Chr(mColChar.门诊号) & lngLoop).Value)
            End If
            rsTmp.AddNew
                                        
            '组别处理
            If str组别 <> "" Then
                For lngLoop2 = 1 To lvwGroup.ListItems.Count
                    If str组别 = lvwGroup.ListItems(lngLoop2).Text Then
                        Exit For
                    End If
                Next
                
                If lngLoop2 = lvwGroup.ListItems.Count + 1 Then
                    
                    '新增组别
                    'lvwGroup.Rows = lvwGroup.Rows + 1
                    'lvwGroup.TextMatrix(lvwGroup.Rows - 1, 1) = str组别
                    
                    lvwGroup.ListItems.Add , , str组别, 1, 1
                End If
            Else
                str组别 = lvwGroup.ListItems(1).Text
            End If
            
            rsTmp("组别").Value = FitlerImport(str组别, 30)
            rsTmp("病人id").Value = lng病人id
            rsTmp("姓名").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.姓名) & lngLoop).Value, 20)
            rsTmp("身份证").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.身份证号) & lngLoop).Value, 18)
            rsTmp("性别").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.性别) & lngLoop).Value, 4)
            rsTmp("出生日期").Value = FitlerImport(Format(ExWorkSheet.Range(Chr(mColChar.出生日期) & lngLoop).Value, "yyyy-MM-dd"), , "日期")
            rsTmp("婚姻状况").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.婚姻状况) & lngLoop).Value, 4)
            rsTmp("民族").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.民族) & lngLoop).Value, 20)
            rsTmp("国籍").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.国籍) & lngLoop).Value, 30)
            rsTmp("学历").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.学历) & lngLoop).Value, 10)
            rsTmp("职业").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.职业) & lngLoop).Value, 20)
            rsTmp("电子邮件").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.电子邮件) & lngLoop).Value, 50)
            rsTmp("工作单位").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.工作单位) & lngLoop).Value, 100)
            rsTmp("年龄").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.年龄) & lngLoop).Value, 10)
            rsTmp("健康号").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.健康号) & lngLoop).Value, 50)
            rsTmp("就诊卡号").Value = FitlerImport(ExWorkSheet.Range(Chr(mColChar.就诊卡号) & lngLoop).Value, 10)
                        
            rsTmp("门诊号").Value = Val(str门诊号)
'            rsTmp("体检时间").Value = Format(DateAdd("d", 7, CDate(zlDatabase.Currentdate)), "yyyy-MM-dd")
            rsTmp("删除").Value = ""
            
        End If
    Next
    
    Call DeleteRecord(mrsPersons)
    Call CopyRecord(rsTmp, mrsPersons)
    
    objExcel.Quit
    ReadExcelFile = True
    
    frmWait.CloseWait
    
    Exit Function
    
errHand:
    objExcel.Quit
    frmWait.CloseWait
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function FitlerImport(ByVal strText As String, Optional ByVal intLen As Integer = 0, Optional ByVal strMode As String = "字符") As String
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    
    Select Case strMode
    Case "字符"
        
        If InStr(strText, "'") > 0 Then strText = ReplaceAll(strText, "'", "")
        
        If intLen > 0 Then
        
            If LenB(StrConv(strText, vbFromUnicode)) > intLen Then
            
                '取值
                strTmp = StrConv(MidB(StrConv(strText, vbFromUnicode), 1, intLen), vbUnicode)
                
                Clipboard.Clear
                Clipboard.SetText strTmp
                strText = Trim(Clipboard.GetText)
                                
            End If
        End If
        
    Case "日期"
        If CheckStrValid(strText, CHECKFORMAT.日期) = False Then strText = ""
    End Select
     
    FitlerImport = strText
    
End Function

Private Function SearchArchive(ByVal str身份证号 As String, ByRef lng病人id As Long, ByRef str门诊号 As String) As Boolean
    
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    If str身份证号 <> "" Then
        strSQL = "SELECT 病人id,门诊号 FROM 病人信息 WHERE 身份证号=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str身份证号)
        If rs.BOF = False Then
            lng病人id = rs("病人id").Value
            str门诊号 = CStr(zlCommFun.NVL(rs("门诊号").Value, 0))
        End If
        
    End If
    
    SearchArchive = True
    
End Function

Private Sub tbs_Click(PreviousTab As Integer)
    tbs.ZOrder 0
    
    If PreviousTab = 0 Then
        vsf.Visible = False
        vsfPerson.Visible = True
        cmd(5).Visible = False
        cmd(6).Visible = False
        cmd(8).Visible = True
        cmd(9).Visible = True
        cmd(12).Visible = True
        cmd(13).Visible = True
        cmd(16).Visible = True
    Else
        vsf.Visible = True
        vsfPerson.Visible = False
        cmd(5).Visible = True
        cmd(6).Visible = True
        cmd(8).Visible = False
        cmd(9).Visible = False
        cmd(12).Visible = False
        cmd(13).Visible = False
        cmd(16).Visible = False
    End If
End Sub

Private Sub txt_Change(Index As Integer)

    DataChange = True
    
    If Index = 13 Or Index = 5 Then
        txt(Index).Tag = "Changed"
                
        cmd(4).Tag = ""
    End If
    
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    Select Case Index
    Case 0, 2, 5, 6, 12, 13
        zlCommFun.OpenIme True
    End Select
End Sub

Private Function FillPatient(ByVal lngKey As Long, Optional ByVal bytMode As Byte = 1)
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    strSQL = GetPublicSQL(SQL.人员档案)
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        If bytMode = 1 Then
            If mrsPersons.RecordCount = 0 Then mrsPersons.AddNew
            mrsPersons("组别").Value = "缺省"
            mrsPersons("姓名").Value = zlCommFun.NVL(rs("姓名").Value)
            mrsPersons("身份证").Value = zlCommFun.NVL(rs("身份证").Value)
            mrsPersons("性别").Value = zlCommFun.NVL(rs("性别").Value)
            mrsPersons("出生日期").Value = zlCommFun.NVL(rs("出生日期").Value)
            mrsPersons("婚姻状况").Value = zlCommFun.NVL(rs("婚姻状况").Value)
            mrsPersons("病人id").Value = zlCommFun.NVL(rs("病人id").Value)
            mrsPersons("民族").Value = zlCommFun.NVL(rs("民族").Value)
            mrsPersons("国籍").Value = zlCommFun.NVL(rs("国籍").Value)
            mrsPersons("学历").Value = zlCommFun.NVL(rs("学历").Value)
            mrsPersons("职业").Value = zlCommFun.NVL(rs("职业").Value)
            mrsPersons("身份").Value = zlCommFun.NVL(rs("身份").Value)
            mrsPersons("联系人姓名").Value = zlCommFun.NVL(rs("联系人姓名").Value)
            mrsPersons("联系人电话").Value = zlCommFun.NVL(rs("联系人电话").Value)
            'mrsPersons("电子邮件").Value = zlCommFun.NVL(rs("电子邮件").Value)
            mrsPersons("联系人地址").Value = zlCommFun.NVL(rs("联系人地址").Value)
            mrsPersons("工作单位").Value = zlCommFun.NVL(rs("工作单位").Value)
        Else
        
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.民族) = zlCommFun.NVL(rs("民族").Value)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.国籍) = zlCommFun.NVL(rs("国籍").Value)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.学历) = zlCommFun.NVL(rs("学历").Value)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.职业) = zlCommFun.NVL(rs("职业").Value)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份) = zlCommFun.NVL(rs("身份").Value)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人姓名) = zlCommFun.NVL(rs("联系人姓名").Value)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人电话) = zlCommFun.NVL(rs("联系人电话").Value)
            'vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.电子邮件) = varParam(13)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人地址) = zlCommFun.NVL(rs("联系人地址").Value)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.工作单位) = zlCommFun.NVL(rs("工作单位").Value)
            
        End If
    End If
    
End Function

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strInput As String
    Dim strText As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim blnCard As Boolean
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    
    If Index = 5 Then
        '就诊卡号

        blnCard = InputIsCard(txt(Index).Text, KeyAscii)

        If blnCard Then
            If Len(txt(Index).Text) = ParamInfo.就诊卡号码长度 - 1 And KeyAscii <> 8 Or KeyAscii = 13 And txt(Index).Text <> "" Then
                If KeyAscii <> 13 Then
                    txt(Index).Text = txt(Index).Text & Chr(KeyAscii)
                    txt(Index).SelStart = Len(txt(Index).Text)
                    KeyAscii = 0
                End If
    
                strInput = strInput & " AND C.就诊卡号=[1] "
    
            End If
        End If
    End If
        
    If KeyAscii <> vbKeyReturn Then
    
        If Index = 4 Then If FilterKeyAscii(KeyAscii, 99, "0123456789X") = 0 Then KeyAscii = 0
        If Index = 14 Then If FilterKeyAscii(KeyAscii, 2) = 0 Then KeyAscii = 0
        If Index = 31 Then If FilterKeyAscii(KeyAscii, 1) = 0 Then KeyAscii = 0
        
        DataChange = True
        
    ElseIf txt(Index).Tag = "Changed" And Index = 5 And KeyAscii = 13 Then
        If InStr(txt(Index).Text, "'") Then
            ShowSimpleMsg "在个人姓名中有非法字符 ' ！"
            Exit Sub
        End If
        
        imgNew(1).Visible = False
        
        Select Case UCase(Left(txt(Index).Text, 1))
        Case "-", "A"                 '病人id
            strInput = strInput & " AND C.病人id=[1]"
        
        Case "+", "B"                 '住院号
            strInput = " AND C.住院号=[1]"

        Case "*", "D"                 '门诊号
            strInput = strInput & " AND C.门诊号=[1]"
            
        Case "/", "C"                 '当前床号
            strInput = strInput & " AND C.当前床号=[1]"
        Case Else
        
            cmd(1).Tag = ""
            imgNew(1).Visible = True
            txt(3).Text = ""
            
            txt(Index).Tag = ""
        
            '预约人缺省为病人本人
            txt(0).Text = txt(5).Text
            
            zlCommFun.PressKey vbKeyTab
            zlCommFun.PressKey vbKeyTab
            
            Exit Sub
            
        End Select
    ElseIf txt(Index).Tag = "Changed" And Index = 13 And KeyAscii = 13 Then
        If InStr(txt(Index).Text, "'") Then
            ShowSimpleMsg "在团体名称中有非法字符 ' ！"
            Exit Sub
        End If
        
        Dim lngKey As Long
        
        lngKey = Val(cmd(4).Tag)
        
        gstrSQL = GetPublicSQL(SQL.团体过滤选择)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "%" & UCase(txt(Index).Text) & "%")
        
        If ShowTxtFilter(Me, txt(Index), "名称,1800,0,0;编码,900,0,0;简码,900,0,0;联系人,900,0,0;电话,1200,0,0", Me.Name & "\团体过滤选择", "请从下面选择一个团体单位", rsData, rs, , , , False) Then
            
            Call ReadGroup(zlCommFun.NVL(rs("ID").Value, 0))
            
            If lngKey <> Val(cmd(4).Tag) Then DataChange = True
            
'            cmd(Index).Tag = lngKey
            
            imgNew(0).Visible = False
            
            txt(0).Text = txt(12).Text
            txt(1).Text = txt(11).Text
            
        Else
            cmd(4).Tag = ""
            imgNew(0).Visible = True
            
            mrsGroup("编码").Value = ""
        End If
        
        txt(Index).Tag = ""
        
        '预约人缺省为团体联系人
        txt(0).Text = txt(12).Text
        txt(1).Text = txt(11).Text
        
        zlCommFun.PressKey vbKeyTab
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    ElseIf KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    End If
    
    If strInput <> "" Then
    
        strText = txt(Index).Text
        
        KeyAscii = 0
        
        gstrSQL = GetPublicSQL(SQL.人员过滤选择, strInput)
        
        If blnCard Then
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(txt(Index).Text))
        Else
            Select Case UCase(Left(txt(Index).Text, 1))
            Case "/", "C"                 '当前床号
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(Mid(txt(Index).Text, 2)))
            Case Else
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(txt(Index).Text, 2)))
            End Select
        End If

        If ShowTxtFilter(Me, txt(Index), "姓名,1200,0,0;性别,810,0,0;出生日期,1200,0,0;婚姻状况,900,0,0;身份证号,1500,0,0", Me.Name & "\人员过滤选择", "请从下面选择一个人员", rsData, rs, , , , False) Then
                                    
            txt(Index).Text = zlCommFun.NVL(rs("姓名"))
            txt(4).Text = zlCommFun.NVL(rs("身份证号"))
            txt(9).Text = zlCommFun.NVL(rs("年龄"))
            txt(3).Text = zlCommFun.NVL(rs("门诊号"))
            txt(14).Text = zlCommFun.NVL(rs("健康号"))
            
            zlControl.CboLocate cbo(1), zlCommFun.NVL(rs("性别").Value)
            zlControl.CboLocate cbo(0), zlCommFun.NVL(rs("婚姻状况").Value)
            
            cmd(1).Tag = zlCommFun.NVL(rs("ID"))
            
            Call FillPatient(Val(cmd(1).Tag))
            
            txt(3).Locked = (Val(txt(3).Text) > 0 And Val(cmd(1).Tag) > 0)
            
        Else
            cmd(1).Tag = ""
            imgNew(1).Visible = True
            txt(3).Text = ""
            txt(4).Text = ""
        End If
        
        txt(Index).Tag = ""
        
        '预约人缺省为病人本人
        txt(0).Text = txt(5).Text

        zlCommFun.PressKey vbKeyTab
        zlCommFun.PressKey vbKeyTab
        
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 0, 2, 5, 6, 12, 13
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    
End Sub

Private Sub txtSum_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txtSum(Index)
        
End Sub

Private Sub txtSum_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        Call WritePrice(vsf.Row)
        
        If Index = 1 Then Call ChangeTotal(Val(txtSum(0).Text), Val(txtSum(1).Text), 1)
        If Index = 2 Then Call ChangeTotal(Val(txtSum(0).Text), Val(txtSum(2).Text), 2)
        
        Call ReadPrice(vsf.Row)
   
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtSum_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtSum(Index).Locked Then
        glngTXTProc = GetWindowLong(txtSum(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtSum(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtSum_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtSum(Index).Locked Then
        Call SetWindowLong(txtSum(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtSum_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txtSum(Index).Text, txtSum(Index).MaxLength)
    
    If Index = 1 Then
        If InStr(txtSum(1).Text, ".") > 0 Then
            If Len(Mid(txtSum(1).Text, InStr(txtSum(1).Text, ".") + 1)) > 2 Then
                MsgBox "只允许输入两位小数位数。", vbExclamation, gstrSysName
                Cancel = True
            End If
        End If
    End If
    
End Sub


Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    If vsf.Rows = 2 And Val(vsf.RowData(1)) = 0 Then
        Call ResetVsf(vsfPrice)
    Else
        Call ReadPrice(vsf.Row)
    End If
    Call CountGroup
    
    Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.基本价格)), Val(vsf.TextMatrix(Row, mCol.体检价格)), 1)
    
    DataChange = True
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case mCol.执行科室
        
        vsf.TextMatrix(Row, mCol.执行科室id) = vsf.Body.ComboData
        vsf.TextMatrix(Row, mCol.执行科室) = vsf.Cell(flexcpTextDisplay, Row, mCol.执行科室)
    
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.采集方式
        
        vsf.TextMatrix(Row, mCol.采集方式id) = vsf.Body.ComboData
        vsf.TextMatrix(Row, mCol.采集方式) = vsf.Cell(flexcpTextDisplay, Row, mCol.采集方式)
    
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.采集科室
    
        vsf.TextMatrix(Row, mCol.采集科室id) = vsf.Body.ComboData
        vsf.TextMatrix(Row, mCol.采集科室) = vsf.Cell(flexcpTextDisplay, Row, mCol.采集科室)
        
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.体检价格
        
        Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.基本价格)), Val(vsf.TextMatrix(Row, mCol.体检价格)), 1)
        Call ReadPrice(Row)

    '------------------------------------------------------------------------------------------------------------------
    Case mCol.折扣
        
        Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.基本价格)), Val(vsf.TextMatrix(Row, mCol.折扣)), 2)
        Call ReadPrice(Row)
        
    End Select
    DataChange = True
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If NewRow = OldRow Then Exit Sub
    
    Call ReadPrice(NewRow)
    
    Call vsfPrice_AfterRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col)
End Sub

Private Function ReadPrice(ByVal intRow As Integer) As Boolean
    '读取对应的计费明细
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngRow As Long
    Dim lngCol As Long
    
    Call ResetVsf(vsfPrice)
    
    If intRow = 0 Then Exit Function
    
    If vsf.TextMatrix(intRow, mCol.计费明细) <> "" Then
        
        varRow = Split(vsf.TextMatrix(intRow, mCol.计费明细), ";")
        
        vsfPrice.Rows = UBound(varRow) + 2
        
        For lngRow = 0 To UBound(varRow)
            If varRow(lngRow) <> "" Then
                varCol = Split(varRow(lngRow), ":")
'                For lngCol = 0 To UBound(varCol)
                    
                    If Val(varCol(6)) = 2 Then
                        vsfPrice.TextMatrix(lngRow + 1, mCol.p计价项目) = "采集方式-" & Trim(vsf.TextMatrix(vsf.Row, mCol.采集方式))
                    ElseIf Trim(vsf.TextMatrix(vsf.Row, mCol.类别)) = "检验" Then
                        vsfPrice.TextMatrix(lngRow + 1, mCol.p计价项目) = "检验项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目))
                    Else
                        vsfPrice.TextMatrix(lngRow + 1, mCol.p计价项目) = "检查项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目))
                    End If
                    
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p名称) = varCol(0)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p计算单位) = varCol(1)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p数次) = varCol(2)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p标准单价) = varCol(3)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p体检单价) = varCol(4)
                    
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p标准金额) = Val(varCol(2)) * Val(varCol(3))
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p体检金额) = Val(varCol(2)) * Val(varCol(4))
                                        
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p收费项目id) = varCol(5)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p计价性质) = varCol(6)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p执行科室) = varCol(7)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p执行科室id) = varCol(8)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p类别) = varCol(9)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p折扣) = varCol(10)
                    
                    vsfPrice.RowData(lngRow + 1) = Val(varCol(5))
                    
'                Next
            End If
        Next
        
    End If
    
    ReadPrice = True
    
End Function

Private Function WritePrice(ByVal intRow As Integer) As Boolean
    Dim strTmp As String
    Dim lngRow As Long
    Dim varCol As Variant
    
    On Error GoTo errHand
    
    If intRow <= 0 Then Exit Function
    
    For lngRow = 1 To vsfPrice.Rows - 1
        If Val(vsfPrice.TextMatrix(lngRow, mCol.p收费项目id)) > 0 Then
            
            varCol = Split(String(11, ":"), ":")
            
            varCol(0) = vsfPrice.TextMatrix(lngRow, mCol.p名称)
            varCol(1) = vsfPrice.TextMatrix(lngRow, mCol.p计算单位)
            varCol(2) = vsfPrice.TextMatrix(lngRow, mCol.p数次)
            varCol(3) = vsfPrice.TextMatrix(lngRow, mCol.p标准单价)
            varCol(4) = vsfPrice.TextMatrix(lngRow, mCol.p体检单价)
            varCol(5) = vsfPrice.TextMatrix(lngRow, mCol.p收费项目id)
            varCol(6) = vsfPrice.TextMatrix(lngRow, mCol.p计价性质)
            
            If Val(varCol(6)) <> 2 Then varCol(6) = 1
                        
            varCol(7) = vsfPrice.TextMatrix(lngRow, mCol.p执行科室)
            varCol(8) = vsfPrice.TextMatrix(lngRow, mCol.p执行科室id)
            varCol(9) = vsfPrice.TextMatrix(lngRow, mCol.p类别)
            varCol(10) = vsfPrice.TextMatrix(lngRow, mCol.p折扣)
            
            If strTmp = "" Then
                strTmp = Join(varCol, ":")
            Else
                strTmp = strTmp & ";" & Join(varCol, ":")
            End If
        End If
    Next
    
    vsf.TextMatrix(intRow, mCol.计费明细) = strTmp
    
    WritePrice = True
    
errHand:
    
End Function


Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsf.RowData(Row)) <= 0)
    Cancel = (Val(vsf.TextMatrix(Row, mCol.执行科室id)) <= 0)
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    On Error GoTo errHand
    
    If NewRow = OldRow Then Exit Sub
    
    '设置编辑状态
    If Val(vsf.TextMatrix(NewRow, mCol.新加)) = 1 Then
        vsf.EditMode(mCol.项目) = 0
        vsf.EditMode(mCol.执行科室) = 0
        vsf.EditMode(mCol.检查部位) = 0
        vsf.EditMode(mCol.采集方式) = 0
        vsf.EditMode(mCol.采集科室) = 0
        vsf.EditMode(mCol.检验标本) = 0
        vsf.EditMode(mCol.结算方式) = 0
        
        vsf.ComboList(mCol.项目) = ""
        vsf.ComboList(mCol.执行科室) = ""
        vsf.ComboList(mCol.检查部位) = ""
        vsf.ComboList(mCol.采集方式) = ""
        vsf.ComboList(mCol.采集科室) = ""
        vsf.ComboList(mCol.检验标本) = ""
        vsf.ComboList(mCol.结算方式) = ""
    Else
        
        vsf.EditMode(mCol.项目) = 1
        vsf.EditMode(mCol.执行科室) = 1
        vsf.EditMode(mCol.结算方式) = 1
        
        vsf.ComboList(mCol.项目) = "..."
        vsf.ComboList(mCol.执行科室) = " "
        vsf.ComboList(mCol.结算方式) = "记帐|收费"
        
        If mblnGroup Then
            vsf.EditMode(mCol.结算方式) = 0
            vsf.ComboList(mCol.结算方式) = ""
        End If
        
        Select Case vsf.TextMatrix(NewRow, mCol.类别)
            Case "检查"
                vsf.EditMode(mCol.采集方式) = 0
                vsf.EditMode(mCol.检验标本) = 0
                vsf.EditMode(mCol.检查部位) = 1
                vsf.EditMode(mCol.采集科室) = 0
                
                vsf.ComboList(mCol.采集科室) = ""
                vsf.ComboList(mCol.采集方式) = ""
                vsf.ComboList(mCol.检验标本) = ""
                vsf.ComboList(mCol.检查部位) = "..."
            Case "检验"
                vsf.EditMode(mCol.采集方式) = 1
                vsf.EditMode(mCol.检验标本) = 1
                vsf.EditMode(mCol.检查部位) = 0
                vsf.EditMode(mCol.采集科室) = 1
                
                vsf.ComboList(mCol.采集科室) = " "
                vsf.ComboList(mCol.采集方式) = " "
                vsf.ComboList(mCol.检验标本) = " "
                vsf.ComboList(mCol.检查部位) = ""
        End Select
    End If
    
    Call WritePrice(OldRow)
    
    If vsf.TextMatrix(NewRow, mCol.类别) = "检验" Then
        Call SetRowData(Val(vsf.RowData(NewRow)), NewRow, "计价项目", "诊疗执行科室", "采集方式", "检验标本")
        Call SetRowData(Val(vsf.RowData(NewRow)), NewRow, "采集科室")
    ElseIf vsf.TextMatrix(NewRow, mCol.类别) = "检查" Then
        Call SetRowData(Val(vsf.RowData(NewRow)), NewRow, "计价项目", "诊疗执行科室")
    End If
    
    Exit Sub
    
errHand:
    
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim bytResult As Byte
    Dim rsPrice As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strKeys As String
    Dim rsData As New ADODB.Recordset
    
    Select Case Col
        Case mCol.项目
            
            gstrSQL = GetPublicSQL(SQL.体检项目选择)
            Dim bytParam1 As Byte
            Dim bytParam2 As Byte
            
            bytParam1 = 1
            bytParam2 = 2
                    
            If mblnGroup = False Then
                Select Case zlCommFun.GetNeedName(cbo(1).Text)
                Case "男"
                    bytParam1 = 1
                    bytParam2 = 1
                Case "女"
                    bytParam1 = 2
                    bytParam2 = 2
                End Select
            End If
            
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, bytParam1, bytParam2)
            
            If ShowGrdSelect(Me, vsf, "编码,1200,0,1;名称,2700,0,0;单位,900,0,0;标本部位,900,0,0;类别,900,0,0", Me.Name & "\体检项目选择", "请从列表中选择一个体检项目。", rsData, rs, 8790, 4500) Then
                '选取了一个项目
                If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                    Exit Sub
                End If
                
                vsf.Cell(flexcpText, Row, mCol.项目 + 1, Row, vsf.Cols - 1) = ""
                
                vsf.EditText = zlCommFun.NVL(rs("名称").Value)
                vsf.TextMatrix(Row, mCol.类别) = zlCommFun.NVL(rs("类别").Value)
                vsf.TextMatrix(Row, mCol.项目) = zlCommFun.NVL(rs("名称").Value)
                vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                  
                If vsf.TextMatrix(Row, mCol.类别) = "检验" Then
                    Call SetRowDefault(Val(vsf.RowData(Row)), Row, "执行科室", "采集方式", "采集科室", "检验标本", "结算方式", "计价项目")
                    
                ElseIf vsf.TextMatrix(Row, mCol.类别) = "检查" Then
                    Call SetRowDefault(Val(vsf.RowData(Row)), Row, "执行科室", "结算方式", "计价项目")
                End If
                
                Call CreatePriceList(Row)
                Call vsf_BeforeRowColChange(0, 0, vsf.Row, vsf.Col, False)
                Call vsfPrice_AfterRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col)
                
                Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.基本价格)), Val(vsf.TextMatrix(Row, mCol.体检价格)), 1)
                
                Call CountGroup
                
                DataChange = True
                
            End If
                
        Case mCol.检查部位
            
            bytResult = ShowOpenList("", mCol.检查部位)
            If bytResult = 0 Then ShowSimpleMsg "没有找到相匹配的项目！"
            If bytResult = 1 Then
                Call CreatePriceList(Row)
                DataChange = True
            End If
            
    End Select
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim bytResult As Byte
    Dim rs As New ADODB.Recordset
    
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." And Col = mCol.项目 Then
            
            If InStr(vsf.EditText, "'") > 0 Then
                KeyCode = 0
                vsf.EditText = ""
                Cancel = True
                Exit Sub
            End If
                        
            bytResult = ShowOpenList(UCase(vsf.EditText), Col)
            
            If bytResult = 0 Then
                '没有匹配的项目
                KeyCode = 0
                Cancel = True
                
                vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                vsf.EditText = vsf.Cell(flexcpData, Row, Col)
                vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                
                MsgBox "没有找到相匹配的体检项目！", vbInformation, gstrSysName
            End If
            
            If bytResult = 1 Then
                '选取了一个项目
                DataChange = True
                
                If Col = mCol.项目 Then
                    
                    If vsf.TextMatrix(Row, mCol.类别) = "检验" Then
                        Call SetRowDefault(Val(vsf.RowData(Row)), Row, "执行科室", "采集方式", "采集科室", "检验标本", "结算方式", "计价项目")
                        
                    ElseIf vsf.TextMatrix(Row, mCol.类别) = "检查" Then
                        Call SetRowDefault(Val(vsf.RowData(Row)), Row, "执行科室", "结算方式", "计价项目")
                    End If
                    
                    Call CreatePriceList(Row)
                    
                    Call vsf_BeforeRowColChange(0, 0, vsf.Row, vsf.Col, False)
                    Call vsfPrice_AfterRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col)
                    Call CountGroup
                    
                    Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.基本价格)), Val(vsf.TextMatrix(Row, mCol.体检价格)), 1)
                    
                End If
            End If
            
            If bytResult = 2 Then
                '取消了本次选择
                KeyCode = 0
                Cancel = True
                
                vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                
            End If
            
        End If
    Else
        DataChange = True
    End If
End Sub

Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    
    If KeyAscii = vbKeyReturn Then
                
        If Col = 1 Then
            If Trim(vsf.TextMatrix(Row, Col)) = "" Then
                
                KeyAscii = 0
                
                If mblnGroup Then
                                            
                    tbs.Tab = 1
                    Call tbs_Click(0)
                    
                Else
                    If dtp(0).Enabled Then
                        dtp(0).SetFocus
                    Else
                        chk.SetFocus
                    End If
                End If
                
                Cancel = True
                
            End If
        End If
    End If
    
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    
    Select Case Col
    Case mPersonCol.门诊号
        '检查门诊号是否存在
        If Trim(vsfPerson.EditText) <> "" Then
            gstrSQL = "Select 1 From 病人信息 Where 门诊号=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(vsfPerson.EditText))
            If rs.BOF = False Then
                '存在
                Cancel = True
                
                vsfPerson.TextMatrix(Row, Col) = vsfPerson.EditText
                
                ShowSimpleMsg "当前门诊号：" & Val(vsfPerson.EditText) & "已经存在，不允许重复！"
                vsfPerson.EditText = ""
                vsfPerson.TextMatrix(Row, Col) = ""
                
            End If
        End If
    End Select
End Sub

Private Sub vsfPerson_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    Call CountGroup
End Sub

Private Sub vsfPerson_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    Select Case Col
    Case mPersonCol.出生日期
    
        If Trim(vsfPerson.TextMatrix(Row, Col)) <> "" Then
            vsfPerson.TextMatrix(Row, Col) = Format(zlCommFun.AddDate(vsfPerson.TextMatrix(Row, Col)), "yyyy-MM-dd")
            If CheckStrValid(vsfPerson.TextMatrix(Row, Col), CHECKFORMAT.日期) = False Then vsfPerson.TextMatrix(Row, Col) = ""
        End If

    Case mPersonCol.电子邮件
    
        If CheckStrValid(vsfPerson.TextMatrix(Row, Col), CHECKFORMAT.电子邮件) = False Then vsfPerson.TextMatrix(Row, Col) = ""
        
    Case mPersonCol.身份证
    
        If CheckStrValid(vsfPerson.TextMatrix(Row, Col), CHECKFORMAT.身份证号) = False Then vsfPerson.TextMatrix(Row, Col) = ""
    
    Case mPersonCol.性别
        
        Call CountGroup
        
    End Select
    
End Sub

Private Sub vsfPerson_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsfPerson.Rows = 2 Then
        vsfPerson.Cell(flexcpForeColor, Row, 0, Row, vsfPerson.Cols - 1) = COLOR.黑色
    End If
End Sub

Private Sub vsfPerson_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    On Error GoTo errHand
    
    If Val(vsfPerson.TextMatrix(NewRow, mPersonCol.病人id)) = 0 Or Val(vsfPerson.TextMatrix(NewRow, mPersonCol.门诊号)) = 0 Then
        vsfPerson.EditMode(mPersonCol.门诊号) = 1
    Else
        vsfPerson.EditMode(mPersonCol.门诊号) = 0
    End If
errHand:
End Sub

Private Sub vsfPerson_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim lngKey As Long
    Dim rs As New ADODB.Recordset
    
    If frmPatientFind.ShowFind(Me, lngKey) Then
        If lngKey > 0 Then
            
            gstrSQL = "SELECT A.* FROM 病人信息 A WHERE A.病人id=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                
                If Val(cmd(4).Tag) <> Val(zlCommFun.NVL(rs("合同单位id"))) And Val(zlCommFun.NVL(rs("合同单位id"))) > 0 And Val(cmd(4).Tag) > 0 Then
                    
                    If MsgBox("不是当前团体的人员，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    
                End If
                
                vsfPerson.EditText = zlCommFun.NVL(rs("姓名"))
                vsfPerson.Cell(flexcpData, Row, vsfPerson.Col) = zlCommFun.NVL(rs("姓名").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.姓名) = zlCommFun.NVL(rs("姓名"))
                
                Call SetRowDefault(0, Row, "缺省信息")
                
                vsfPerson.TextMatrix(Row, mPersonCol.门诊号) = zlCommFun.NVL(rs("门诊号"))
                vsfPerson.TextMatrix(Row, mPersonCol.健康号) = zlCommFun.NVL(rs("健康号"))
                vsfPerson.TextMatrix(Row, mPersonCol.身份证) = zlCommFun.NVL(rs("身份证号"))
                vsfPerson.TextMatrix(Row, mPersonCol.年龄) = zlCommFun.NVL(rs("年龄"))
                vsfPerson.TextMatrix(Row, mPersonCol.出生日期) = Format(zlCommFun.NVL(rs("出生日期")), "yyyy-MM-dd")
                vsfPerson.TextMatrix(Row, mPersonCol.性别) = zlCommFun.NVL(rs("性别").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.婚姻状况) = zlCommFun.NVL(rs("婚姻状况").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.病人id) = zlCommFun.NVL(rs("病人id"))
                
                Call FillPatient(Val(zlCommFun.NVL(rs("病人id"))), 2)
                
                vsfPerson.EditMode(mPersonCol.门诊号) = 0
                Call CountGroup
                DataChange = True
                                
            End If
            
        End If
    End If

End Sub

Private Sub vsfPerson_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    
    Dim strText As String
    Dim strInput As String
    Dim rs As New ADODB.Recordset
    Dim strSvrText As String
    Dim rsData As New ADODB.Recordset
    Dim blnCard As Boolean
    
    If Chr(KeyCode) = "'" Then KeyCode = 0
    
    If Col = mPersonCol.姓名 Then
        
        strText = vsfPerson.EditText
        If KeyCode <> 8 And KeyCode <> 13 Then
            strText = strText & Chr(KeyCode)
        End If
        
        '检查非法字符
        If InStr(strText, "'") > 0 Then
            KeyCode = 0
            ShowSimpleMsg "在个人姓名中有非法字符 ' ！"
            vsfPerson.EditText = ""
            vsfPerson.TextMatrix(Row, Col) = vsfPerson.Cell(flexcpData, Row, Col)
            Cancel = True
            Exit Sub
        End If
                
        '检查是否为就诊卡号码
        blnCard = InputIsCard(vsfPerson.EditText, KeyCode)

        If blnCard And Len(vsfPerson.EditText) = ParamInfo.就诊卡号码长度 - 1 And KeyCode <> 8 And KeyCode <> vbKeyReturn Then
            vsfPerson.Body.EditSelStart = Len(vsfPerson.EditText)
            strInput = strInput & " AND C.就诊卡号=[1] "
        End If

        If KeyCode = vbKeyReturn Then

            If blnCard Then
                '是就诊卡
                strInput = strInput & " AND C.就诊卡号=[1] "
            Else
                '非就诊卡
                blnCard = False
                
                strText = vsfPerson.EditText
                
                Select Case UCase(Left(strText, 1))
                Case "-", "A"                 '病人id,就诊卡号
                    strInput = strInput & " AND C.病人id=[1]"
                Case "+", "B"                 '住院号
                    strInput = " AND C.住院号=[1]"
                Case "*", "D"                 '门诊号
                    strInput = strInput & " AND C.门诊号=[1]"
                Case "/", "C"                 '当前床号
                    strInput = strInput & " AND C.当前床号=[1]"
                Case Else                     '姓名
                    strSvrText = vsfPerson.Cell(flexcpData, Row, Col)
                    vsfPerson.Cell(flexcpData, Row, Col) = strText
                End Select
            End If
                
        End If
    
        
        If strInput <> "" Then
        
            gstrSQL = GetPublicSQL(SQL.人员过滤选择, strInput)
                    
            If blnCard Then
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(strText))
            ElseIf UCase(Left(strText, 1)) = "/" Or UCase(Left(strText, 1)) = "C" Then
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(Mid(strText, 2)))
            Else
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(strText, 2)))
            End If
                    
            If ShowGrdFilter(Me, vsfPerson, "姓名,1200,0,0;性别,810,0,0;出生日期,1200,0,0;婚姻状况,900,0,0;身份证号,1500,0,0", Me.Name & "\人员过滤选择Grid", "请从下面选择一个人员", rsData, rs, , , , False) Then
                                                                        
                vsfPerson.EditText = zlCommFun.NVL(rs("姓名"))
                
                If Val(cmd(4).Tag) <> Val(zlCommFun.NVL(rs("合同单位id"))) And Val(zlCommFun.NVL(rs("合同单位id"))) > 0 And Val(cmd(4).Tag) > 0 Then
                    
                    If MsgBox("病人“" & zlCommFun.NVL(rs("姓名").Value) & "”不是当前团体的人员，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        KeyCode = 0
                        vsfPerson.EditText = ""
                        vsfPerson.TextMatrix(Row, Col) = strSvrText
                        Cancel = True
                        Exit Sub
                    End If
                
                End If
                
                If CheckHavePerson(Val(zlCommFun.NVL(rs("ID")))) Then
                    ShowSimpleMsg "病人“" & zlCommFun.NVL(rs("姓名").Value) & "”已经存在！"
                    KeyCode = 0
                    vsfPerson.EditText = ""
                    vsfPerson.TextMatrix(Row, Col) = strSvrText
                    Cancel = True
                    Exit Sub
                End If

                strText = vsfPerson.EditText
                vsfPerson.Cell(flexcpData, Row, vsfPerson.Col) = zlCommFun.NVL(rs("姓名").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.姓名) = zlCommFun.NVL(rs("姓名"))
                Call SetRowDefault(0, Row, "缺省信息")
                vsfPerson.TextMatrix(Row, mPersonCol.身份证) = zlCommFun.NVL(rs("身份证号"))
                vsfPerson.TextMatrix(Row, mPersonCol.出生日期) = Format(zlCommFun.NVL(rs("出生日期")), "yyyy-MM-dd")
                vsfPerson.TextMatrix(Row, mPersonCol.性别) = zlCommFun.NVL(rs("性别").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.婚姻状况) = zlCommFun.NVL(rs("婚姻状况").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.病人id) = zlCommFun.NVL(rs("ID"))
                vsfPerson.TextMatrix(Row, mPersonCol.年龄) = zlCommFun.NVL(rs("年龄"))
                vsfPerson.TextMatrix(Row, mPersonCol.门诊号) = zlCommFun.NVL(rs("门诊号"))
                vsfPerson.TextMatrix(Row, mPersonCol.健康号) = zlCommFun.NVL(rs("健康号"))
                
                Call FillPatient(Val(zlCommFun.NVL(rs("ID"))), 2)
                
                vsfPerson.Cell(flexcpForeColor, Row, 0, Row, vsfPerson.Cols - 1) = COLOR.黑色
                
                vsfPerson.EditMode(mPersonCol.门诊号) = 0
                Call CountGroup
                
                If blnCard Then
                    vsfPerson.Cell(flexcpData, Row, Col) = strText
                    vsfPerson.TextMatrix(Row, Col) = vsfPerson.Cell(flexcpData, Row, Col)
                    KeyCode = 13
                End If
                
                DataChange = True
            Else
                '取消了本次选择，作为新病人
    
                vsfPerson.EditMode(mPersonCol.门诊号) = 1
                vsfPerson.Cell(flexcpForeColor, Row, 0, Row, vsfPerson.Cols - 1) = COLOR.兰色
                
                vsfPerson.Cell(flexcpData, Row, Col) = vsfPerson.EditText
                vsfPerson.EditText = vsfPerson.Cell(flexcpData, Row, Col)
                vsfPerson.TextMatrix(Row, Col) = vsfPerson.Cell(flexcpData, Row, Col)
                vsfPerson.TextMatrix(Row, mPersonCol.门诊号) = ""
                vsfPerson.TextMatrix(Row, mPersonCol.身份证) = ""
                vsfPerson.TextMatrix(Row, mPersonCol.病人id) = ""
                vsfPerson.TextMatrix(Row, mPersonCol.出生日期) = ""
                vsfPerson.TextMatrix(Row, mPersonCol.年龄) = ""
                
                Call SetRowDefault(0, Row, "缺省信息")
            End If
        ElseIf KeyCode = vbKeyReturn Then
    
            '新病人，允许输入门诊号
            
            vsfPerson.EditMode(mPersonCol.门诊号) = 1
            vsfPerson.Cell(flexcpForeColor, Row, 0, Row, vsfPerson.Cols - 1) = COLOR.兰色
            vsfPerson.TextMatrix(Row, mPersonCol.病人id) = ""
            vsfPerson.TextMatrix(Row, mPersonCol.门诊号) = ""
            vsfPerson.TextMatrix(Row, mPersonCol.身份证) = ""
            vsfPerson.TextMatrix(Row, mPersonCol.出生日期) = ""
            vsfPerson.TextMatrix(Row, mPersonCol.年龄) = ""
            Call SetRowDefault(0, Row, "缺省信息")
        End If
    End If
End Sub

Private Sub vsfPerson_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    
    On Error Resume Next
    
    If KeyAscii = vbKeyReturn Then
        
        If Col = 1 Then
            If Trim(vsfPerson.TextMatrix(Row, Col)) = "" Then
                KeyAscii = 0
                
                If dtp(0).Enabled Then
                    dtp(0).SetFocus
                Else
                    chk.SetFocus
                End If
                
                Cancel = True
                
            End If
        End If
    End If
    
End Sub

Private Sub vsfPerson_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    
    Select Case Col
    Case mPersonCol.门诊号
        '检查门诊号是否存在
        If Val(vsfPerson.EditText) > 0 Then
            gstrSQL = "Select 1 From 病人信息 Where 门诊号=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsfPerson.EditText))
            If rs.BOF = False Then
                '存在
                Cancel = True
                
                vsfPerson.TextMatrix(Row, Col) = vsfPerson.EditText
                
                ShowSimpleMsg "当前门诊号：" & Trim(vsfPerson.EditText) & "已经存在，不允许重复！"
                vsfPerson.EditText = ""
                vsfPerson.TextMatrix(Row, Col) = ""
                
            End If
        End If
    End Select
End Sub

Private Sub vsfPrice_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)

    Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p标准单价)), Val(vsfPrice.TextMatrix(Row, mCol.p体检单价)), 1)
    
    DataChange = True
End Sub

Private Sub vsfPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    With vsfPrice
        Select Case Col
        Case mCol.p计价项目
        
            If Left(.TextMatrix(Row, mCol.p计价项目), 4) = "采集方式" Then
                .TextMatrix(Row, mCol.p计价性质) = "2"
            Else
                .TextMatrix(Row, mCol.p计价性质) = "1"
            End If
            .TextMatrix(Row, mCol.p计价项目) = .Cell(flexcpTextDisplay, Row, mCol.p计价项目)
            
        Case mCol.p数次
            vsfPrice.TextMatrix(Row, mCol.p标准金额) = Val(vsfPrice.TextMatrix(Row, mCol.p标准单价)) * Val(vsfPrice.TextMatrix(Row, mCol.p数次))
            vsfPrice.TextMatrix(Row, mCol.p体检金额) = Val(vsfPrice.TextMatrix(Row, mCol.p体检单价)) * Val(vsfPrice.TextMatrix(Row, mCol.p数次))
            
            Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p标准单价)), Val(vsfPrice.TextMatrix(Row, mCol.p体检单价)), 1)
                    
            If InStr("567", .TextMatrix(Row, mCol.p类别)) > 0 Then
                Call PromptStorageWarn(Val(.TextMatrix(Row, mCol.p数次)), Val(.TextMatrix(Row, mCol.p可用库存)), .TextMatrix(Row, mCol.p名称), .TextMatrix(Row, mCol.p执行科室), .TextMatrix(Row, mCol.p计算单位), 1)
            End If
                
        Case mCol.p体检单价
            
            Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p标准单价)), Val(vsfPrice.TextMatrix(Row, mCol.p体检单价)), 1)
        
        Case mCol.p折扣
            
            Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p标准单价)), Val(vsfPrice.TextMatrix(Row, mCol.p折扣)), 2)
        
        Case mCol.p执行科室
            .TextMatrix(Row, mCol.p执行科室id) = .Body.ComboData
            .TextMatrix(Row, mCol.p执行科室) = .Cell(flexcpTextDisplay, Row, mCol.p执行科室)
            
            If InStr("567", .TextMatrix(Row, mCol.p类别)) > 0 Then
                .TextMatrix(Row, mCol.p可用库存) = GetStorage(Val(.RowData(Row)), Val(.TextMatrix(Row, mCol.p执行科室id)))
                Call PromptStorageWarn(Val(.TextMatrix(Row, mCol.p数次)), Val(.TextMatrix(Row, mCol.p可用库存)), .TextMatrix(Row, mCol.p名称), .TextMatrix(Row, mCol.p执行科室), .TextMatrix(Row, mCol.p计算单位), 1)
            End If
                    
        End Select
    End With
    
    DataChange = True
    
End Sub

Private Sub vsfPrice_AfterNewRow(ByVal Row As Long, Col As Long)
    
    If Row > 1 Then
        vsfPrice.TextMatrix(Row, mCol.p计价项目) = vsfPrice.TextMatrix(Row - 1, mCol.p计价项目)
        
        If Left(vsfPrice.TextMatrix(Row, mCol.p计价项目), 4) = "采集方式" Then
            vsfPrice.TextMatrix(Row, mCol.p计价性质) = "2"
        Else
            vsfPrice.TextMatrix(Row, mCol.p计价性质) = "1"
        End If
        
    End If
    
End Sub

Private Sub vsfPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    Call SetRowData(Val(vsfPrice.RowData(NewRow)), NewRow, "收费执行科室")
End Sub

Private Sub vsfPrice_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str计价项目 As String
    Dim str计价性质 As String
    
    If vsfPrice.Rows = 2 Then
        
        str计价项目 = vsfPrice.TextMatrix(1, mCol.p计价项目)
        str计价性质 = vsfPrice.TextMatrix(1, mCol.p计价性质)
        
        vsfPrice.Body.Cell(flexcpText, 1, mCol.p计价项目 + 1, 1, vsfPrice.Cols - 1) = ""
        vsfPrice.RowData(1) = 0

        vsfPrice.TextMatrix(1, mCol.p计价项目) = str计价项目
        vsfPrice.TextMatrix(1, mCol.p计价性质) = str计价性质
        Call vsfPrice_AfterDeleteRow(1, Col)
        
        Cancel = True
    End If
End Sub

Private Sub vsfPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    If Col = mCol.p名称 Then
        
        
        gstrSQL = GetPublicSQL(SQL.收费项目选择)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        If ShowGrdSelect(Me, vsfPrice, "编码,1200,0,1;名称,2700,0,0;单位,600,0,0;规格,1200,0,0;单价,900,0,0;类别,900,0,0", Me.Name & "\收费项目选择", "请从列表中选择一个收费项目。", rsData, rs, 8790, 5100) Then

            If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                Exit Sub
            End If
            With vsfPrice
                .EditText = zlCommFun.NVL(rs("名称").Value)
                .TextMatrix(Row, mCol.p名称) = zlCommFun.NVL(rs("名称").Value)
                .TextMatrix(Row, mCol.p计算单位) = zlCommFun.NVL(rs("单位").Value)
    
                .TextMatrix(Row, mCol.p标准单价) = zlCommFun.NVL(rs("单价").Value, 0)
                .TextMatrix(Row, mCol.p体检单价) = .TextMatrix(Row, mCol.p标准单价)
    
                .TextMatrix(Row, mCol.p收费项目id) = zlCommFun.NVL(rs("ID").Value, 0)
                If Val(.TextMatrix(Row, mCol.p数次)) < 1 Then .TextMatrix(Row, mCol.p数次) = 1
    
                .TextMatrix(Row, mCol.p标准金额) = Val(.TextMatrix(Row, mCol.p标准单价)) * Val(.TextMatrix(Row, mCol.p数次))
                .TextMatrix(Row, mCol.p体检金额) = .TextMatrix(Row, mCol.p标准金额)
                
                .TextMatrix(Row, mCol.p类别) = zlCommFun.NVL(rs("类别").Value)
                
                .RowData(Row) = zlCommFun.NVL(rs("ID").Value)
    
                Call SetRowDefault(Val(.RowData(Row)), Row, "收费执行科室")
                Call SetRowData(Val(.RowData(Row)), Row, "收费执行科室")
                
                Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p标准单价)), Val(vsfPrice.TextMatrix(Row, mCol.p体检单价)), 1)
                
                If InStr("567", .TextMatrix(Row, mCol.p类别)) > 0 Then
                    .TextMatrix(Row, mCol.p可用库存) = GetStorage(Val(.RowData(Row)), Val(.TextMatrix(Row, mCol.p执行科室id)))
                    Call PromptStorageWarn(Val(.TextMatrix(Row, mCol.p数次)), Val(.TextMatrix(Row, mCol.p可用库存)), .TextMatrix(Row, mCol.p名称), .TextMatrix(Row, mCol.p执行科室), .TextMatrix(Row, mCol.p计算单位), 1)
                End If
'
            End With
            
            DataChange = True

        End If
        
    End If
End Sub

Private Sub vsfPrice_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." Then
            
            If InStr(vsfPrice.EditText, "'") > 0 Then
                KeyCode = 0
                vsfPrice.EditText = ""
                Cancel = True
                Exit Sub
            End If
    
            Select Case Col
                Case mCol.p名称
                    
                    strText = UCase(vsfPrice.EditText)
                    gstrSQL = GetPublicSQL(SQL.收费项目过滤, strText)
                    
                    If ParamInfo.项目输入匹配方式 = 1 Then
                        strTmp = strText & "%"
                    Else
                        strTmp = "%" & strText & "%"
                    End If
                    
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText & "%", strTmp)
                    
                    If ShowGrdFilter(Me, vsfPrice, "编码,1200,0,1;名称,2700,0,0;单位,600,0,0;规格,1200,0,0;单价,900,0,0;类别,900,0,0", Me.Name & "\收费项目过滤", "请从列表中选择一个收费项目。", rsData, rs, 8790, 5100) Then
                        
                        If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                            Exit Sub
                        End If
                        
                        With vsfPrice
                            .EditText = zlCommFun.NVL(rs("名称").Value)
                            .TextMatrix(Row, mCol.p名称) = zlCommFun.NVL(rs("名称").Value)
                            
                            .Cell(flexcpData, Row, mCol.p名称, Row, mCol.p名称) = zlCommFun.NVL(rs("名称").Value)
                            
                            .TextMatrix(Row, mCol.p计算单位) = zlCommFun.NVL(rs("单位").Value)
                            
                            .TextMatrix(Row, mCol.p标准单价) = zlCommFun.NVL(rs("单价").Value, 0)
                            .TextMatrix(Row, mCol.p体检单价) = .TextMatrix(Row, mCol.p标准单价)
                            
                            .TextMatrix(Row, mCol.p收费项目id) = zlCommFun.NVL(rs("ID").Value, 0)
                            If Val(.TextMatrix(Row, mCol.p数次)) < 1 Then .TextMatrix(Row, mCol.p数次) = 1
                            
                            .TextMatrix(Row, mCol.p标准金额) = Val(.TextMatrix(Row, mCol.p标准单价)) * Val(.TextMatrix(Row, mCol.p数次))
                            .TextMatrix(Row, mCol.p体检金额) = .TextMatrix(Row, mCol.p标准金额)
                            .TextMatrix(Row, mCol.p类别) = zlCommFun.NVL(rs("类别").Value)
                            
                            .RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                            
                            Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p标准单价)), Val(vsfPrice.TextMatrix(Row, mCol.p体检单价)), 1)
                            
                            Call SetRowDefault(Val(.RowData(Row)), Row, "收费执行科室")
                            Call SetRowData(Val(.RowData(Row)), Row, "收费执行科室")
                            
                            If InStr("567", .TextMatrix(Row, mCol.p类别)) > 0 Then
                                .TextMatrix(Row, mCol.p可用库存) = GetStorage(Val(.RowData(Row)), Val(.TextMatrix(Row, mCol.p执行科室id)))
                                Call PromptStorageWarn(Val(.TextMatrix(Row, mCol.p数次)), Val(.TextMatrix(Row, mCol.p可用库存)), .TextMatrix(Row, mCol.p名称), .TextMatrix(Row, mCol.p执行科室), .TextMatrix(Row, mCol.p计算单位), 1)
                            End If
                        End With
                        
                        DataChange = True
                    Else
                        KeyCode = 0
                        Cancel = True
                        
                        vsfPrice.Cell(flexcpData, Row, Col) = vsfPrice.Cell(flexcpData, Row, Col)
                        vsfPrice.EditText = vsfPrice.Cell(flexcpData, Row, Col)
                        vsfPrice.TextMatrix(Row, Col) = vsfPrice.Cell(flexcpData, Row, Col)
                        
                    End If
            End Select
        End If
    Else
        DataChange = True
    End If
End Sub








