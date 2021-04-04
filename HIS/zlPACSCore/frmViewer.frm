VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmViewer 
   Caption         =   "ZLPACS Viewer"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   9765
   Icon            =   "frmViewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9765
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPrintInterval 
      Height          =   1455
      Left            =   4200
      ScaleHeight     =   1395
      ScaleWidth      =   1995
      TabIndex        =   13
      Top             =   5280
      Visible         =   0   'False
      Width           =   2055
      Begin VB.OptionButton optPrintStart 
         Caption         =   "偶数起"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   18
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton optPrintStart 
         Caption         =   "奇数起"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdPrintInterval 
         Caption         =   "间隔打印"
         Height          =   350
         Left            =   480
         TabIndex        =   16
         Top             =   960
         Width           =   1100
      End
      Begin MSComCtl2.UpDown udPrintInterval 
         Height          =   300
         Left            =   1621
         TabIndex        =   15
         Top             =   450
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtPrintInterval"
         BuddyDispid     =   196612
         OrigLeft        =   1680
         OrigTop         =   450
         OrigRight       =   1935
         OrigBottom      =   750
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtPrintInterval 
         Height          =   300
         Left            =   720
         TabIndex        =   14
         Text            =   "5"
         Top             =   450
         Width           =   900
      End
      Begin VB.Label lblPrtintInterval 
         Caption         =   "间隔："
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.PictureBox picViewer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4608
      Left            =   1320
      ScaleHeight     =   4605
      ScaleWidth      =   10065
      TabIndex        =   1
      Top             =   720
      Width           =   10068
      Begin VB.PictureBox PicX 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         Height          =   4296
         Index           =   0
         Left            =   744
         MousePointer    =   9  'Size W E
         ScaleHeight     =   4230
         ScaleWidth      =   330
         TabIndex        =   8
         Top             =   156
         Visible         =   0   'False
         Width           =   384
      End
      Begin VB.VScrollBar VScro 
         Height          =   5292
         Index           =   0
         Left            =   5880
         TabIndex        =   7
         Top             =   30
         Visible         =   0   'False
         Width           =   250
      End
      Begin VB.PictureBox PicY 
         Height          =   100
         Index           =   0
         Left            =   360
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   4155
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.PictureBox PicYY 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   108
         Left            =   1440
         ScaleHeight     =   105
         ScaleWidth      =   3780
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   3780
      End
      Begin VB.PictureBox PicXX 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5220
         Left            =   360
         ScaleHeight     =   5220
         ScaleWidth      =   105
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   108
      End
      Begin VB.TextBox txtText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000001&
         Height          =   288
         Left            =   1680
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.PictureBox PicXY 
         BorderStyle     =   0  'None
         Height          =   100
         Index           =   0
         Left            =   240
         MousePointer    =   15  'Size All
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   100
      End
      Begin DicomObjects.DicomViewer Viewer 
         Height          =   2085
         Index           =   0
         Left            =   6120
         TabIndex        =   9
         Tag             =   "0"
         Top             =   720
         Visible         =   0   'False
         Width           =   2370
         _Version        =   262147
         _ExtentX        =   4191
         _ExtentY        =   3683
         _StockProps     =   35
         BackColor       =   12648447
         AutoDisplay     =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid MSFViewer 
         Height          =   1245
         Left            =   1560
         TabIndex        =   10
         Top             =   1680
         Visible         =   0   'False
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   2196
         _Version        =   393216
         FixedRows       =   0
      End
      Begin VB.Label lblChange 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Height          =   180
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   96
      End
   End
   Begin MSComctlLib.ListView lvwSort 
      Height          =   675
      Left            =   660
      TabIndex        =   0
      Top             =   4140
      Visible         =   0   'False
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1191
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog Common 
      Left            =   3075
      Top             =   1065
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   7080
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2302
            MinWidth        =   2293
            Text            =   "中联软件"
            TextSave        =   "中联软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4868
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2831
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4586
            MinWidth        =   4586
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "大写"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   926
            MinWidth        =   706
            Text            =   "数字"
            TextSave        =   "NUM"
            Object.ToolTipText     =   "数字"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListMouse 
      Left            =   2760
      Top             =   5850
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":0CCA
            Key             =   "Stack"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":0E2C
            Key             =   "WindowWL"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1146
            Key             =   "Zoom"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList24 
      Left            =   8400
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   84
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1460
            Key             =   "另存报告图"
            Object.Tag             =   "108"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1BDA
            Key             =   "全序列观片"
            Object.Tag             =   "250"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2354
            Key             =   "框选L"
            Object.Tag             =   "365"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":256E
            Key             =   "框选R"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2788
            Key             =   "缩略图"
            Object.Tag             =   "249"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3162
            Key             =   "打开"
            Object.Tag             =   "102"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":38DC
            Key             =   "胶片打印"
            Object.Tag             =   "406"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":4056
            Key             =   "鼠标"
            Object.Tag             =   "419"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":47D0
            Key             =   "放大镜"
            Object.Tag             =   "402"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":4F4A
            Key             =   "手动调窗L"
            Object.Tag             =   "314"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":56C4
            Key             =   "手动调窗R"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":5E3E
            Key             =   "自适应调窗L"
            Object.Tag             =   "315"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":65B8
            Key             =   "自适应调窗R"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6D32
            Key             =   "漫游L"
            Object.Tag             =   "309"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":74AC
            Key             =   "漫游R"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7C26
            Key             =   "缩放L"
            Object.Tag             =   "311"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":83A0
            Key             =   "缩放R"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":8B1A
            Key             =   "穿梭L"
            Object.Tag             =   "308"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":9294
            Key             =   "穿梭R"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":9A0E
            Key             =   "显示CT值"
            Object.Tag             =   "362"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":A188
            Key             =   "电影"
            Object.Tag             =   "401"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":A902
            Key             =   "序列全选"
            Object.Tag             =   "303"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":B07C
            Key             =   "图像全选"
            Object.Tag             =   "304"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":B7F6
            Key             =   "上一序列"
            Object.Tag             =   "247"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":BF70
            Key             =   "下一个序列"
            Object.Tag             =   "248"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":C6EA
            Key             =   "排版"
            Object.Tag             =   "201"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":CE64
            Key             =   "全屏"
            Object.Tag             =   "246"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":D5DE
            Key             =   "显示隐藏图像信息"
            Object.Tag             =   "203"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":DD58
            Key             =   "全部恢复"
            Object.Tag             =   "312"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":E4D2
            Key             =   "浏览观察模式"
            Object.Tag             =   "202"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":EC4C
            Key             =   "图像属性"
            Object.Tag             =   "112"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":F3C6
            Key             =   "锐化减少"
            Object.Tag             =   "329"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":FB40
            Key             =   "锐化增强"
            Object.Tag             =   "330"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":102BA
            Key             =   "左移减少"
            Object.Tag             =   "333"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":10A34
            Key             =   "左移增强"
            Object.Tag             =   "334"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":111AE
            Key             =   "平滑减少"
            Object.Tag             =   "331"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":11928
            Key             =   "平滑增强"
            Object.Tag             =   "332"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":120A2
            Key             =   "图像复原"
            Object.Tag             =   "335"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1281C
            Key             =   "伪彩"
            Object.Tag             =   "405"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":12F96
            Key             =   "全部定位线"
            Object.Tag             =   "318"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":13710
            Key             =   "首尾定位线"
            Object.Tag             =   "319"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":13E8A
            Key             =   "当前定位线"
            Object.Tag             =   "320"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":14604
            Key             =   "三维鼠标L"
            Object.Tag             =   "321"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":14D7E
            Key             =   "三维鼠标R"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":154F8
            Key             =   "矢冠状重建"
            Object.Tag             =   "403"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":15C72
            Key             =   "水平翻转"
            Object.Tag             =   "323"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":163EC
            Key             =   "垂直翻转"
            Object.Tag             =   "324"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":16B66
            Key             =   "左转90度"
            Object.Tag             =   "325"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":172E0
            Key             =   "右转90度"
            Object.Tag             =   "326"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":17A5A
            Key             =   "反白"
            Object.Tag             =   "327"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":181D4
            Key             =   "DSA"
            Object.Tag             =   "404"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1894E
            Key             =   "文字L"
            Object.Tag             =   "337"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":190C8
            Key             =   "文字R"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":19842
            Key             =   "箭头L"
            Object.Tag             =   "338"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":19FBC
            Key             =   "箭头R"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1A736
            Key             =   "椭圆L"
            Object.Tag             =   "339"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1AEB0
            Key             =   "椭圆R"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1B62A
            Key             =   "角度L"
            Object.Tag             =   "340"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1BDA4
            Key             =   "角度R"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1C51E
            Key             =   "曲线L"
            Object.Tag             =   "341"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1CC98
            Key             =   "曲线R"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1D412
            Key             =   "区域L"
            Object.Tag             =   "342"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1DB8C
            Key             =   "区域R"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1E306
            Key             =   "直线L"
            Object.Tag             =   "343"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1EA80
            Key             =   "直线R"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1F1FA
            Key             =   "矩形L"
            Object.Tag             =   "344"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1F974
            Key             =   "矩形R"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":200EE
            Key             =   "清除所有标注"
            Object.Tag             =   "347"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":20868
            Key             =   "血管测量L"
            Object.Tag             =   "361"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":20FE2
            Key             =   "血管测量R"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2175C
            Key             =   "校准"
            Object.Tag             =   "346"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":21ED6
            Key             =   "区域直方图"
            Object.Tag             =   "345"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":22650
            Key             =   "裁剪L"
            Object.Tag             =   "310"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":22DCA
            Key             =   "裁剪R"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":23544
            Key             =   "图像同步"
            Object.Tag             =   "307"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":23CBE
            Key             =   "序列同步"
            Object.Tag             =   "306"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":24438
            Key             =   "中联图标"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":24BB2
            Key             =   "手工序列同步"
            Object.Tag             =   "363"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2532C
            Key             =   "锁定序列"
            Object.Tag             =   "364"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":25AA6
            Key             =   "心胸比R"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":261A0
            Key             =   "心胸比L"
            Object.Tag             =   "367"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2689A
            Key             =   "滤镜模板"
            Object.Tag             =   "32810"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":27014
            Key             =   "打印序列"
            Object.Tag             =   "40601"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2778E
            Key             =   "斜面重建"
            Object.Tag             =   "420"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList16 
      Left            =   7770
      Top             =   5910
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   83
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":27EA0
            Key             =   "另存报告图"
            Object.Tag             =   "108"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2823A
            Key             =   "全序列观片"
            Object.Tag             =   "250"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":287D4
            Key             =   "框选L"
            Object.Tag             =   "365"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2892E
            Key             =   "框选R"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":28A88
            Key             =   "缩略图"
            Object.Tag             =   "249"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":28E22
            Key             =   "打开"
            Object.Tag             =   "102"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":291BC
            Key             =   "胶片打印"
            Object.Tag             =   "406"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":29556
            Key             =   "鼠标"
            Object.Tag             =   "419"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":298F0
            Key             =   "放大镜"
            Object.Tag             =   "402"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":29C8A
            Key             =   "手动调窗L"
            Object.Tag             =   "314"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2A024
            Key             =   "手动调窗R"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2A3BE
            Key             =   "自适应调窗L"
            Object.Tag             =   "315"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2A758
            Key             =   "自适应调窗R"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2AAF2
            Key             =   "漫游L"
            Object.Tag             =   "309"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2AE8C
            Key             =   "漫游R"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2B226
            Key             =   "缩放L"
            Object.Tag             =   "311"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2B5C0
            Key             =   "缩放R"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2B95A
            Key             =   "穿梭L"
            Object.Tag             =   "308"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2BCF4
            Key             =   "穿梭R"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2C08E
            Key             =   "显示CT值"
            Object.Tag             =   "362"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2C428
            Key             =   "电影"
            Object.Tag             =   "401"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2C7C2
            Key             =   "序列全选"
            Object.Tag             =   "303"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2CB5C
            Key             =   "图像全选"
            Object.Tag             =   "304"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2CEF6
            Key             =   "上一序列"
            Object.Tag             =   "247"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2D290
            Key             =   "下一个序列"
            Object.Tag             =   "248"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2D62A
            Key             =   "排版"
            Object.Tag             =   "201"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2D9C4
            Key             =   "全屏"
            Object.Tag             =   "246"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2DD5E
            Key             =   "显示隐藏图像信息"
            Object.Tag             =   "203"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2E0F8
            Key             =   "全部恢复"
            Object.Tag             =   "312"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2E492
            Key             =   "浏览观察模式"
            Object.Tag             =   "202"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2E82C
            Key             =   "图像属性"
            Object.Tag             =   "112"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2EBC6
            Key             =   "锐化减少"
            Object.Tag             =   "329"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2EF60
            Key             =   "锐化增强"
            Object.Tag             =   "330"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2F2FA
            Key             =   "左移减少"
            Object.Tag             =   "333"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2F694
            Key             =   "左移增强"
            Object.Tag             =   "334"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2FA2E
            Key             =   "平滑减少"
            Object.Tag             =   "331"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2FDC8
            Key             =   "平滑增强"
            Object.Tag             =   "332"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":30162
            Key             =   "图像复原"
            Object.Tag             =   "335"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":304FC
            Key             =   "伪彩"
            Object.Tag             =   "405"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":30896
            Key             =   "全部定位线"
            Object.Tag             =   "318"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":30C30
            Key             =   "首尾定位线"
            Object.Tag             =   "319"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":30FCA
            Key             =   "当前定位线"
            Object.Tag             =   "320"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":31364
            Key             =   "三维鼠标L"
            Object.Tag             =   "321"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":316FE
            Key             =   "三维鼠标R"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":31A98
            Key             =   "矢冠状重建"
            Object.Tag             =   "403"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":31E32
            Key             =   "水平翻转"
            Object.Tag             =   "323"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":321CC
            Key             =   "垂直翻转"
            Object.Tag             =   "324"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":32566
            Key             =   "左转90度"
            Object.Tag             =   "325"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":32900
            Key             =   "右转90度"
            Object.Tag             =   "326"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":32C9A
            Key             =   "反白"
            Object.Tag             =   "327"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":33034
            Key             =   "DSA"
            Object.Tag             =   "404"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":333CE
            Key             =   "文字L"
            Object.Tag             =   "337"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":33768
            Key             =   "文字R"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":33B02
            Key             =   "箭头L"
            Object.Tag             =   "338"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":33E9C
            Key             =   "箭头R"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":34236
            Key             =   "椭圆L"
            Object.Tag             =   "339"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":345D0
            Key             =   "椭圆R"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3496A
            Key             =   "角度L"
            Object.Tag             =   "340"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":34D04
            Key             =   "角度R"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3509E
            Key             =   "曲线L"
            Object.Tag             =   "341"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":35438
            Key             =   "曲线R"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":357D2
            Key             =   "区域L"
            Object.Tag             =   "342"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":35B6C
            Key             =   "区域R"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":35F06
            Key             =   "直线L"
            Object.Tag             =   "343"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":362A0
            Key             =   "直线R"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3663A
            Key             =   "矩形L"
            Object.Tag             =   "344"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":369D4
            Key             =   "矩形R"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":36D6E
            Key             =   "清除所有标注"
            Object.Tag             =   "347"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":37108
            Key             =   "血管测量L"
            Object.Tag             =   "361"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":374A2
            Key             =   "血管测量R"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3783C
            Key             =   "校准"
            Object.Tag             =   "346"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":37BD6
            Key             =   "区域直方图"
            Object.Tag             =   "345"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":37F70
            Key             =   "裁剪L"
            Object.Tag             =   "310"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3830A
            Key             =   "裁剪R"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":386A4
            Key             =   "图像同步"
            Object.Tag             =   "307"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":38A3E
            Key             =   "序列同步"
            Object.Tag             =   "306"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":38DD8
            Key             =   "手工序列同步"
            Object.Tag             =   "363"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":39172
            Key             =   "锁定序列"
            Object.Tag             =   "364"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3950C
            Key             =   "心胸比R"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":39AA6
            Key             =   "心胸比L"
            Object.Tag             =   "367"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3A040
            Key             =   "滤镜模板"
            Object.Tag             =   "32810"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3A3DA
            Key             =   "打印序列"
            Object.Tag             =   "40601"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3A774
            Key             =   "斜面重建"
            Object.Tag             =   "420"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList32 
      Left            =   9000
      Top             =   5910
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   83
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3AAC6
            Key             =   "另存报告图"
            Object.Tag             =   "108"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3B7A0
            Key             =   "全序列观片"
            Object.Tag             =   "250"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3C47A
            Key             =   "框选L"
            Object.Tag             =   "365"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3C794
            Key             =   "框选R"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3CAAE
            Key             =   "缩略图"
            Object.Tag             =   "249"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3D788
            Key             =   "打开"
            Object.Tag             =   "102"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3F272
            Key             =   "胶片打印"
            Object.Tag             =   "406"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3FF4C
            Key             =   "鼠标"
            Object.Tag             =   "419"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":40C26
            Key             =   "放大镜"
            Object.Tag             =   "402"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":41900
            Key             =   "手动调窗L"
            Object.Tag             =   "314"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":425DA
            Key             =   "手动调窗R"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":432B4
            Key             =   "自适应调窗L"
            Object.Tag             =   "315"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":43F8E
            Key             =   "自适应调窗R"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":44C68
            Key             =   "漫游L"
            Object.Tag             =   "309"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":45942
            Key             =   "漫游R"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":4661C
            Key             =   "缩放L"
            Object.Tag             =   "311"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":46EF6
            Key             =   "缩放R"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":477D0
            Key             =   "穿梭L"
            Object.Tag             =   "308"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":484AA
            Key             =   "穿梭R"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":49184
            Key             =   "显示CT值"
            Object.Tag             =   "362"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":49E5E
            Key             =   "电影"
            Object.Tag             =   "401"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":4AB38
            Key             =   "序列全选"
            Object.Tag             =   "303"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":4B812
            Key             =   "图像全选"
            Object.Tag             =   "304"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":4C4EC
            Key             =   "上一序列"
            Object.Tag             =   "247"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":4F8DE
            Key             =   "下一个序列"
            Object.Tag             =   "248"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":52CD0
            Key             =   "排版"
            Object.Tag             =   "201"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":539AA
            Key             =   "全屏"
            Object.Tag             =   "246"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":55394
            Key             =   "显示隐藏图像信息"
            Object.Tag             =   "203"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":5606E
            Key             =   "全部恢复"
            Object.Tag             =   "312"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":56D48
            Key             =   "浏览观察模式"
            Object.Tag             =   "202"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":57A22
            Key             =   "图像属性"
            Object.Tag             =   "112"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":586FC
            Key             =   "锐化减少"
            Object.Tag             =   "329"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":593D6
            Key             =   "锐化增强"
            Object.Tag             =   "330"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":5A0B0
            Key             =   "左移减少"
            Object.Tag             =   "333"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":5AD8A
            Key             =   "左移增强"
            Object.Tag             =   "334"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":5BA64
            Key             =   "平滑减少"
            Object.Tag             =   "331"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":5EE56
            Key             =   "平滑增强"
            Object.Tag             =   "332"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":62248
            Key             =   "图像复原"
            Object.Tag             =   "335"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":62F22
            Key             =   "伪彩"
            Object.Tag             =   "405"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":63BFC
            Key             =   "全部定位线"
            Object.Tag             =   "318"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":648D6
            Key             =   "首尾定位线"
            Object.Tag             =   "319"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":655B0
            Key             =   "当前定位线"
            Object.Tag             =   "320"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6628A
            Key             =   "三维鼠标L"
            Object.Tag             =   "321"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":66F64
            Key             =   "三维鼠标R"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":67C3E
            Key             =   "矢冠状重建"
            Object.Tag             =   "403"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":68918
            Key             =   "水平翻转"
            Object.Tag             =   "323"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":695F2
            Key             =   "垂直翻转"
            Object.Tag             =   "324"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6A2CC
            Key             =   "左转90度"
            Object.Tag             =   "325"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6BDB6
            Key             =   "右转90度"
            Object.Tag             =   "326"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6CA90
            Key             =   "反白"
            Object.Tag             =   "327"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6D76A
            Key             =   "DSA"
            Object.Tag             =   "404"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6E444
            Key             =   "文字L"
            Object.Tag             =   "337"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6F11E
            Key             =   "文字R"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":71580
            Key             =   "箭头L"
            Object.Tag             =   "338"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7225A
            Key             =   "箭头R"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":72F34
            Key             =   "椭圆L"
            Object.Tag             =   "339"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":73C0E
            Key             =   "椭圆R"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":748E8
            Key             =   "角度L"
            Object.Tag             =   "340"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":755C2
            Key             =   "角度R"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7629C
            Key             =   "曲线L"
            Object.Tag             =   "341"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":76F76
            Key             =   "曲线R"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":77C50
            Key             =   "区域L"
            Object.Tag             =   "342"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7892A
            Key             =   "区域R"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":79604
            Key             =   "直线L"
            Object.Tag             =   "343"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7A2DE
            Key             =   "直线R"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7AFB8
            Key             =   "矩形L"
            Object.Tag             =   "344"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7BC92
            Key             =   "矩形R"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7C96C
            Key             =   "清除所有标注"
            Object.Tag             =   "347"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7D646
            Key             =   "血管测量L"
            Object.Tag             =   "361"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":80A38
            Key             =   "血管测量R"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":83E2A
            Key             =   "校准"
            Object.Tag             =   "346"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":84B04
            Key             =   "区域直方图"
            Object.Tag             =   "345"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":857DE
            Key             =   "裁剪L"
            Object.Tag             =   "310"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":864B8
            Key             =   "裁剪R"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":87192
            Key             =   "图像同步"
            Object.Tag             =   "307"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":88B7C
            Key             =   "序列同步"
            Object.Tag             =   "306"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":89856
            Key             =   "手工序列同步"
            Object.Tag             =   "363"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":8A530
            Key             =   "锁定序列"
            Object.Tag             =   "364"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":8B20A
            Key             =   "心胸比R"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":8BAE4
            Key             =   "心胸比L"
            Object.Tag             =   "367"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":8C3BE
            Key             =   "滤镜模板"
            Object.Tag             =   "32810"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":8D098
            Key             =   "打印序列"
            Object.Tag             =   "40601"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":8DD72
            Key             =   "斜面重建"
            Object.Tag             =   "420"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars ComToolBar 
      Left            =   600
      Top             =   5880
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DkpMain 
      Bindings        =   "frmViewer.frx":8E9C4
      Left            =   480
      Top             =   1200
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''MSFViewer结构
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''0列=系列类型 0=来源于PACS图像服务器上的可以保存的序列   1=来源于文件严格区分的序列,2=混合序列,3=类似矢冠状重建产生的序列---以后不用了
''''1列=是否有图像  as Boolean---以后不用了
''''2=检查UID       as string  BMP、jpg文件为空---以后不用了
''''3=当前选择的图像号
''''4=当前选择的图像处于第几帧
''''(OK)5=该序列横向显示图像数目(供序列内单图和多图显示切换用)
''''(OK)6=该序列纵向显示图像数目(供序列内单图和多图显示切换用)
''''7=该序列当前显示第一个图像序号(供序列内单图和多图显示切换用)
''''8=该序列当前显示选择图像序号(供序列内单图和多图显示切换用)
''''9=该序列内图像是否自动同步---以后不用了
''''10=记录该序列当前选中的LABEL序号
''''11=记录该序列所在的横向位置---以后不用了
''''12=记录该序列所在的纵向位置---以后不用了
''''13=记录对应的临时Viewer位置,矢冠状位重建的时候，重建图像所对应的X方向结果图的Viewer index---以后不用了
''''14=记录对应的临时Viewer位置,矢冠状位重建的时候，重建图像所对应的Y方向结果图的Viewer index---以后不用了
''''15=记录当前序列是否被选择，用于自动和手工序列同步   ----以后不用了
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'-----------------------------------------------------
'窗体事件
'-----------------------------------------------------
Public Event AfterSaveReportImage(strCheckUID As String)
Public Event AfterSaveOuterImage(strCheckUID As String)
Public Event AfterSeriesChanged(strStudyUID As String, strSeriesUID As String)

''''''''''''''''''''''''''''''''''''''''''[Viewer控制]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public intSelectedSerial As Integer                             ''''当前操作的序列(Viewer)
Public oldSelectedSerial As Integer                                ''''记录上一次选择的序列

''''''''''''''''''''''''''[通过鼠标调整序列位置用中间变量]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public intFactMoveX As Integer                                  '''''记录鼠标按下横向分隔条的位置
Public intFactMoveY As Integer                                  '''''记录鼠标按下纵向分隔条的位置
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public intOldCountX, intOldCountY As Integer                    '''''记录MPR之前横向的序列布局数和纵向的序列布局数
Public intCountX, intCountY As Integer                          '''''横向的序列排布数和纵向序列排布数
Public intDefaultCountX, intDefaultCountY As Integer            '''''第一个图像的设备类型所决定的横向的序列排布数和纵向序列排布数
Public blnAutoCount                                             '''''第一个图像的设备类型所决定的横向的序列排布数和纵向序列排布数是否自动计算
Public isSelectAllSerial As Boolean                             ''''标注是否选择了所有序列
Public isSelectAllImage As Boolean                              ''''标注是否选择了所有图像
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public intDblClickButton As Integer                             ''''记录鼠标按键
Public SelectedImage As DicomImage                              ''''当前选择的图像
Public oldSelectedImageIndex As Integer                         ''''记录上次选择的图像INDEX
Public SelectedImageIndex As Integer                            ''''本次选择的图像INDEX

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim blnTextInput As Boolean                                     '''''开始输入文字的标识
Dim blnTextInputM As Boolean                                    '''''开始修改文字的标识
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public intClickImageIndex As Integer                            ''''当前点击的图像序列号,0标识没有点中,在MOUSE中填写,供双击使用
Dim lngBaseX As Long                                            ''''弹出菜单记录位置也使用
Dim lngBaseY As Long                                            ''''弹出菜单记录位置也使用
Dim lngBaseXX As Long
Dim lngBaseYY As Long
''''''''''''''''''''''''''''''''''''''''[穿梭播放用变量]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public objStackOldImage As New DicomImage                       '''''''单祯播放记录当前图像
Public intStackIndex As Integer                                '''''''单祯播放记录当前播放图像序号
Public blnStackisFrame As Boolean                               '''''''记录采用多帧播放或是单祯循环播放
Public intStackCurrentlyImage As Integer                        ''''''''记录当前祯数或当前图像号
Public intStackOffset As Integer                                ''''''''记录开始穿梭图像和右上图像的偏移量
Public objStackImages As DicomImages                            ''''''''单祯播放记录用的图像集
Dim blnStackStart As Boolean                                    ''''''''标记穿梭开始
Dim blnMouseStart As Boolean                                    ''''''''记录穿梭开始后，鼠标第一次拖动，需要初始化图像位置等
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public SelectedLabel As DicomLabel                              ''''当前选择的标注
Public SelectedLabelT As DicomLabel                             ''''当前选择的标注当前选择标注对应文字
Public isSelectedLabel As Boolean                               ''''是否选择了标注
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim blnAutoWL As Boolean                                        '''''开始自动窗宽窗位
Dim blnFrameSelectImage As Boolean                              '''''开始框选图象
Dim LabelDrawing As Boolean                                     '''''开始画标注的标识
Dim blnAngle As Boolean                                         '''''开始画角度的标识
Dim intVasMeasure As Integer                                    '''''开始血管狭窄测量的标识，0-表示不在血管狭窄测量状态，1-表示画正常血管测量，2-表示画狭窄血管测量。
Dim intCadioThoracicRatio As Integer                            '''''开始心胸比测量的标志，0-表示不在心胸比测量状态，1-表示画心脏测量，2-表示画胸廓测量。
Dim oldFontSize As Integer                                      '''''标识文字输入点阵前后变换使用
Dim oldTextleft As Integer                                      '''''中间变量
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim blnMoveLabel  As Boolean                                    '''''开始移动标注的标识
Dim blnReSizeLabel As Boolean                                   '''''开始改变标注形状的标识
Dim intReSizeIndex As Integer                                   '''''中间变量，记录句柄序号
Public LngOldColor As Long                                      ''SubChangeColor过程使用
Public DLblOld As DicomLabel                                    ''SubChangeColor过程使用
Public dubCalibrateLength As Double                             '''''校准长度
Public blnForceRefresh As Boolean                               ''''是否强制刷新,供[配置程序使用

''''''''''''''''''''''''''''''''''''''''''三维鼠标''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim obj3dImage() As DicomImage                                  ''''各Viewer当前图像的备份
Dim int3dImageIndex() As Integer                                ''''各Viewer当前图像的INDEX备份
Dim int3dCurrentlyImage() As Integer                            ''''各Viewer左上角图像的INDEX备份
Dim blnIn3dCursor As Boolean                                    ''''是否进入三维鼠标状态

'''''''''''''''''''''''''''''''''''''''''图像拼接'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim fis As New frmImageSpelling
Public blnfis As Boolean                                        ''''是否进入了拼接状态

'''''''''''''''''''''''''''''''''''''''''胶片打印'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public blnPrintFilm As Boolean                                  ''''是否进入了胶片打印状态
Public WithEvents mfrmFilm As frmFilm                                      ''''胶片打印窗体
Attribute mfrmFilm.VB_VarHelpID = -1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim intIntercept As Integer                                     ''''记录显示CT值时换算用的截距
Dim intSlope As Integer                                         ''''记录显示CT值是换算用的斜率
Dim strInstanceUID As String                                    ''''记录显示CT值时的图像实例UID
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim blDicomDown As Boolean                                      ''''是否按下鼠标
Public blnVscroInvoked As Boolean                               ''''记录是否由手工序列同步引起的Vscro改变
Public blnDefaultWW2 As Boolean                                 ''''记录是否使用了默认的第二个窗宽窗位，支持默认双窗口
''''''''''''''''''''''''''''''''''''''''矢冠状位重建''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public blnInMPR As Boolean                                      ''''是否正在进行矢冠状位重建

'观片站注册
Public blnLogined As Boolean                                    ''''是否注册成功，True成功，False失败。
Private mstr启动时间 As String                                   ''''注册成功后的启动时间

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''供技师和医生站调用的本窗体的显示
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(objParent As Object)
    Me.Show , objParent
End Sub

Private Sub cmdPrintInterval_Click()
    Dim intInterval As Integer
    Dim blnStartOdd As Boolean
    
    If optPrintStart(1).Value = True Then
        blnStartOdd = True
    Else
        blnStartOdd = False
    End If
    
    intInterval = Val(txtPrintInterval.Text)
    If intInterval <= 0 Or intInterval >= 100 Then
        intInterval = 1
    End If
    
    '按照设置的间隔添加打印序列
    Call funFilm(Me, False, 4, intInterval, blnStartOdd)
    '关闭弹出菜单
    ComToolBar.ClosePopups
End Sub

Private Sub ComToolBar_ControlSelected(ByVal control As XtremeCommandBars.ICommandBarControl)
    '显示工具栏提示信息
    Me.sbStatusBar.Panels(2).Text = StatusBarTip(control)
End Sub

Private Sub ComToolBar_Customization(ByVal Options As XtremeCommandBars.ICustomizeOptions)
    '设置可以用来自定义的按钮
    Dim Controls As CommandBarControls
    Dim cbrControl As CommandBarControl
    Dim ControlPopup As CommandBarPopup
    
    '隐藏几个页面的显示
    Options.ShowKeyboardPage = False
    Options.ShowOptionsPage = False
    
    Options.CustomIcons.RemoveAll
'    Options.ContextMenu.Title
    
    '添加支持自定义设置的按钮
    Set Controls = ComToolBar.DesignerControls
    
    Controls.DeleteAll
    
    If (Controls.Count = 0) Then
        '主工具栏
        Set cbrControl = Controls.Add(xtpControlButton, ID_File_SAveASReport, "保存报告图")
        cbrControl.Category = "主工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_File_Open, "打开")
        cbrControl.Category = "主工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Tool_FilmPrint, "胶片输出")
        cbrControl.Category = "主工具栏"
        Set ControlPopup = Controls.Add(xtpControlSplitButtonPopup, ID_Tool_Film_AddSeries, "打印序列")
        ControlPopup.Category = "主工具栏"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddSeries, "打印序列"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddImage, "打印图像"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddSelected, "打印所选图"
            Set ControlPopup = ControlPopup.CommandBar.Controls.Add(xtpControlButtonPopup, ID_Tool_Film_AddInterval, "间隔打印")
            ControlPopup.CommandBar.SetPopupToolBar True
            ControlPopup.CommandBar.Title = "间隔打印"
            ControlPopup.ToolTipText = "间隔打印当前序列"
        
        '图像操作
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Eddy_LeftRight, "水平镜象")
        cbrControl.Category = "图像操作"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Eddy_TopButton, "垂直镜象")
        cbrControl.Category = "图像操作"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Eddy_Left90, "左转90度")
        cbrControl.Category = "图像操作"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Eddy_Right90, "右转90度")
        cbrControl.Category = "图像操作"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_ReverseVideo, "反白")
        cbrControl.Category = "图像操作"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Tool_NumberMinusShadow, "DSA数字减影")
        cbrControl.Category = "图像操作"
        
        '测量工具栏
        Set cbrControl = Controls.Add(xtpControlButton, ID_Tool_NothinMouseState, "鼠标")
        cbrControl.Category = "测量工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_ACtive_Mouse_Value, "在鼠标上显示CT值")
        cbrControl.Category = "测量工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_Text, "文字")
        cbrControl.Category = "测量工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_Arrowhead, "箭头")
        cbrControl.Category = "测量工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_Ellipse, "椭圆")
        cbrControl.Category = "测量工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_Angle, "角度")
        cbrControl.Category = "测量工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_Curve, "曲线")
        cbrControl.Category = "测量工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_Area, "区域")
        cbrControl.Category = "测量工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_BeeLine, "直线")
        cbrControl.Category = "测量工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_Rect, "矩形")
        cbrControl.Category = "测量工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_VasMeasure, "血管狭窄测量")
        cbrControl.Category = "测量工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_CadioThoracicRatio, "心胸比测量")
        cbrControl.Category = "测量工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_ClearLbale, "清除标注")
        cbrControl.Category = "测量工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_AdjustLine, "校准")
        cbrControl.Category = "测量工具栏"
        
        '多平面工具栏
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_PointingLine_ALL, "显示所有定位线")
        cbrControl.Category = "多平面工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_PointingLine_FirstLast, "显示首尾定位线")
        cbrControl.Category = "多平面工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_PointingLine_Now, "显示当前定位线")
        cbrControl.Category = "多平面工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_PointingLine_3DLine, "三维鼠标")
        cbrControl.Category = "多平面工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Tool_ArrowyCoronaryReset, "矢/冠状位重建")
        cbrControl.Category = "多平面工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Tool_SlopeReconstruction, "斜面重建")
        cbrControl.Category = "多平面工具栏"
        
        '对象分析
        Set cbrControl = Controls.Add(xtpControlButton, ID_ACtive_FrameSelectImage, "框选图像")
        cbrControl.Category = "对象分析"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Also_Photo, "图像格式同步")
        cbrControl.Category = "对象分析"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Also_Serial, "序列间图像位置同步")
        cbrControl.Category = "对象分析"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Also_ManualSerial, "手工序列同步")
        cbrControl.Category = "对象分析"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Also_LockSerial, "锁定/解锁序列")
        cbrControl.Category = "对象分析"
        Set cbrControl = Controls.Add(xtpControlButton, ID_View_ShowMiniSeries, "显示序列缩略图")
        cbrControl.Category = "对象分析"
        Set cbrControl = Controls.Add(xtpControlButton, ID_View_ViewAllSeries, "全序列观片")
        cbrControl.Category = "对象分析"
        
        '通用工具栏
        Set cbrControl = Controls.Add(xtpControlButton, ID_Tool_Magnifier, "放大镜")
        cbrControl.Category = "通用工具栏"
        
        '手控调窗，包含动态子菜单，暂时先不支持子菜单
'        Set ControlPopup = Controls.Add(xtpControlSplitButtonPopup, ID_Active_AdjustWindow_HandAdjustWindow, "手控调窗")
'        ControlPopup.Category = "通用工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_AdjustWindow_HandAdjustWindow, "手控调窗")
        cbrControl.Category = "通用工具栏"
        
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Cruise, "漫游")
        cbrControl.Category = "通用工具栏"
        Set ControlPopup = Controls.Add(xtpControlSplitButtonPopup, ID_Active_Zoom, "缩放")
        ControlPopup.Category = "通用工具栏"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_AutoShow, "自适应"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_50%, "50%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_100%, "100%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_150%, "150%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_200%, "200%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_250%, "250%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_300%, "300%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_400%, "400%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_Custom, "自定义(&A)"
        Set ControlPopup = Controls.Add(xtpControlSplitButtonPopup, ID_Active_Shuttle, "穿梭")
        ControlPopup.Category = "通用工具栏"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_PhotoNumber, "图像号"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_BedASC, "床位正序"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_BedDESC, "床位逆序"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_CollectionTime, "采集时间"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_PhotoTime, "图像时间"

        Set cbrControl = Controls.Add(xtpControlButton, ID_Tool_Movie, "电影播放")
        cbrControl.Category = "通用工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Select_SelectAllSerial, "选择所有序列")
        cbrControl.Category = "通用工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Acitve_Select_SelectAllPhoto, "选择序列中所有的图像")
        cbrControl.Category = "通用工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_View_UpSeries, "上一序列")
        cbrControl.Category = "通用工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_View_DownSeries, "下一序列")
        cbrControl.Category = "通用工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_View_Typeset, "版面设计")
        cbrControl.Category = "通用工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_View_FullScreen, "全屏显示")
        cbrControl.Category = "通用工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_View_PropertyShow, "显/隐病人信息")
        cbrControl.Category = "通用工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_ReSetAll, "全部恢复")
        cbrControl.Category = "通用工具栏"
        Set cbrControl = Controls.Add(xtpControlButton, ID_View_OneBrowse, "浏览/观察模式")
        cbrControl.Category = "通用工具栏"
        
        '图像增强
        
        '滤镜模版是动态增加的菜单，先不支持
'        Set ControlPopup = Controls.Add(xtpControlSplitButtonPopup, ID_Active_SieveLens_Model, "滤镜模板")
'        ControlPopup.Category = "图像增强"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_SieveLens_LancetMinus, "边缘增强强度减少")
        cbrControl.Category = "图像增强"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_SieveLens_LancetAdd, "边缘增强强度增加")
        cbrControl.Category = "图像增强"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Sievelens_LeftMoveMinus, "边缘增强幅度减少")
        cbrControl.Category = "图像增强"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Sievelens_LeftMoveAdd, "边缘增强幅度增加")
        cbrControl.Category = "图像增强"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_SieveLens_FlatnessMinus, "平滑减少")
        cbrControl.Category = "图像增强"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_SieveLens_FlatnessAdd, "平滑增加")
        cbrControl.Category = "图像增强"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Sievelens_PhotoReset, "图像复原")
        cbrControl.Category = "图像增强"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Tool_BogusColour, "伪彩")
        cbrControl.Category = "图像增强"
        
    End If
End Sub

Private Sub ComToolBar_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim CmdControl As CommandBarControl
    Dim i As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    subMouseRLset control                          ''''处理鼠标左右键分布
    subMnuImageSort control.Id, Me                     ''''排序方式处理
    '''''''''''''''''''''''''''''[功能键设置窗宽窗位处理]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If control.Id >= 349 And control.Id <= 360 Then
        For i = 349 To 360
            If Not ComToolBar.Item(ToolBar_Comm).FindControl(, i, , True) Is Nothing Then
                ComToolBar.Item(ToolBar_Comm).FindControl(, i, , True).Checked = False
                If i = ID_Active_AdjustWindow_HandAdjustWindow_Custom Then
                    ComToolBar.Item(ToolBar_Menu).FindControl(, i, , True).Checked = False
                End If
            End If
        Next
        control.Checked = True
        subFunctionWL ComToolBar.Item(ToolBar_Comm).FindControl(, control.Id, , True), Me
        If control.Id = ID_Active_AdjustWindow_HandAdjustWindow_Custom Then
            ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow_Custom, , True).Checked = True
            ComToolBar.Item(ToolBar_Comm).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow_Custom, , True).Checked = True
        End If
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''''''[处理弹出滤镜模板菜单]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If control.Id >= ID_Active_SieveLens_Model + 1 And control.Id <= ID_Active_SieveLens_Model + 40 Then
        Call subFunctionFilter(control, Me)
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''[处理弹出菜单拷贝序列]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If control.Id >= 500 And control.Id <= 800 And control.Category <> "" Then
        '计算这个Viewer摆放的位置
        Call subIsSerialXY(Me, lngBaseX, lngBaseY, intCol, intRow)
        Call subCreateAndPlaceAViewer(Val(control.Category), intRow, intCol)
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not SelectedImage Is Nothing Then
        If SelectedImage.Attributes(&H28, &H4) = "MONOCHROME2" Or SelectedImage.Attributes(&H28, &H4) = "MONOCHROME1" Then
            blnSelectedImageIfColor = False
        Else
            blnSelectedImageIfColor = True
        End If
    End If
    Select Case control.Id
        ''''''''''''''''''''''''''文件菜单'''''''''''''''''''''''''''''''''''
        Case ID_File_Open                                                               '打开文件
            subOpenFiles Me
            
        Case ID_File_Close                                                              '关闭序列
            subCloseSeries
            
        Case ID_File_DelAllPhoto                                                        '删除所有图像
            subKillPicture
            
        Case ID_File_DelReport                                                          '删除报告图像
            subDelRepImg
            RaiseEvent AfterSaveReportImage(PstrCheckUID)
            
        Case ID_File_SaveFile                                                           '保存文件
            Set frmSave.f = Me: frmSave.Show 1, Me
            
        Case ID_File_SaveASFile                                                         '另存文件
            Set FrmSaveAs.f = Me: FrmSaveAs.Show 1, Me
        
        Case ID_File_SaveToCD                                                           '创建CD
            '程序待定
            Set frmCreateCD.f = Me: frmCreateCD.Show 1, Me
            
        Case ID_File_SAveASReport                                                       '另存报告图像
            subOutPutRptImg
            RaiseEvent AfterSaveReportImage(PstrCheckUID)
            
        Case ID_File_Send_GetHost                                                       '接收主机
            Set frmSendImage.f = Me: frmSendImage.Show 1, Me
            
        Case ID_File_Send_OutPowerPoint                                                 '输出到PowerPoint
            subOutputToPowerPoint Me
            
        Case ID_File_OpenDicomDir                                                       '打开DICOMDIR
            Set frmOpenDicomDir.f = Me
            frmOpenDicomDir.Show 1, Me
            
        Case ID_File_PhotoProperty                                                      '图像属性
            If SelectedImage Is Nothing Then Exit Sub
            Set FrmSfyInfo.img = SelectedImage
            FrmSfyInfo.Show 1, Me
            
        Case ID_File_Exit                                                               '退出
            Unload Me
            
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''视图''''''''''''''''''''''''''''''''''''''''
        Case ID_View_UpSeries                                                           '上一序列
            subChangeASeries 2
            
        Case ID_View_DownSeries                                                         '下一序列
            subChangeASeries 1
        
        Case ID_View_Typeset                                                            '版面安排
            Dim fLayout As New frmSerialLayoutSetup
            fLayout.zlShowMe Me
            
        Case ID_View_OneBrowse                                                          '浏览观察模式
            Call subLookOrBrowsSwitch(Me)
            
        Case ID_View_PropertyShow                                                       '图像上病人信息显示
            Dim v As DicomViewer
            Button_miDispPatientInfo = Not Button_miDispPatientInfo
            ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_PropertyShow).Checked = Button_miDispPatientInfo
            ComToolBar.RecalcLayout
            For Each v In Viewer
            If v.Index <> 0 Then
                subDisplayPatientInfo v
            End If
        Next
            
        Case ID_View_LableShow                                                          '标注显示
            subDispLabelInfo Me
            
        Case ID_View_ShowOverlay                                                        '显示Overlay
            control.Checked = Not control.Checked
            Button_miShowOverlay = control.Checked
            Call ShowOverlay(Me)
            
        Case ID_View_ShowMiniSeries                                                     '显示序列缩略图
            ' 处理菜单、内部变量
            Button_miShowMiniSeries = Not Button_miShowMiniSeries
            Me.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_ShowMiniSeries, , True).Checked = Button_miShowMiniSeries
            Me.ComToolBar.Item(ToolBar_Object).FindControl(, ID_View_ShowMiniSeries, , True).Checked = Button_miShowMiniSeries
            Call subShowMiniImages(Me)
        
        Case ID_View_ViewAllSeries                                                      '全序列观片
            Button_miViewAllSeries = Not Button_miViewAllSeries
            Me.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_ViewAllSeries, , True).Checked = Button_miViewAllSeries
            Me.ComToolBar.Item(ToolBar_Object).FindControl(, ID_View_ViewAllSeries, , True).Checked = Button_miViewAllSeries
            
        Case ID_View_ShowScale_AutoShow                                                 '自适应
            subShowScale ID_View_ShowScale_AutoShow
            If Not SelectedImage Is Nothing Then               '保证图像存在
                SelectedImage.StretchToFit = True
                Viewer(intSelectedSerial).Refresh
                '处理序列内图像同步
                Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
            End If
            
        Case ID_View_ShowScale_50%                                                      '50%
            subShowScale ID_View_ShowScale_50%
            subCenterZoom SelectedImage, Viewer(intSelectedSerial), 0.5
            '处理序列内图像同步
            Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
            
        Case ID_View_ShowScale_100%                                                     '100%
            subShowScale ID_View_ShowScale_100%
            subCenterZoom SelectedImage, Viewer(intSelectedSerial), 1
            '处理序列内图像同步
            Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
        
        Case ID_View_showScale_150%                                                     '150%
            subShowScale ID_View_showScale_150%
            subCenterZoom SelectedImage, Viewer(intSelectedSerial), 1.5
            '处理序列内图像同步
            Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
            
        Case ID_View_ShowScale_200%                                                     '200%
            subShowScale ID_View_ShowScale_200%
            subCenterZoom SelectedImage, Viewer(intSelectedSerial), 2
            '处理序列内图像同步
            Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
            
        Case ID_View_showScale_250%                                                     '250%
            subShowScale ID_View_showScale_250%
            subCenterZoom SelectedImage, Viewer(intSelectedSerial), 2.5
            '处理序列内图像同步
            Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
            
        Case ID_View_showScale_300%                                                     '300%
            subShowScale ID_View_showScale_300%
            subCenterZoom SelectedImage, Viewer(intSelectedSerial), 3
            '处理序列内图像同步
            Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
            
        Case ID_View_showScale_400%                                                     '400%
            subShowScale ID_View_showScale_400%
            subCenterZoom SelectedImage, Viewer(intSelectedSerial), 4
            '处理序列内图像同步
            Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
            
        Case ID_View_ShowScale_Custom                                                   '自定义
            subShowScale ID_View_ShowScale_Custom
            If Not SelectedImage Is Nothing Then
                frmZoomCustom.sRatio = SelectedImage.Zoom
                frmZoomCustom.Show 1, Me
                If frmZoomCustom.bApply Then
                    subCenterZoom SelectedImage, Viewer(intSelectedSerial), IIf(frmZoomCustom.sRatio = 0, 1, frmZoomCustom.sRatio)
                    '处理序列内图像同步
                    Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
                End If
            End If
            
        Case ID_View_FullScreen                                                         '全屏显示
            subFullScreen Me
            
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''动作''''''''''''''''''''''''''''''''''''
        Case ID_Active_Select_OneSelect                                                 '单幅选择
            ZLShowSeriesInfos(intSelectedSerial).ImageInfos(SelectedImageIndex).blnSelected = Not ZLShowSeriesInfos(intSelectedSerial).ImageInfos(SelectedImageIndex).blnSelected
            subDispframe Me, Viewer(intSelectedSerial)
            Viewer(intSelectedSerial).Refresh
        Case ID_Active_Select_SelectAllSerial                                           '选择所有序列
            subSelectAllSerial Me
        Case ID_Acitve_Select_SelectAllPhoto                                            '选择所有图像
            subSelectAllIMage Me
            
        Case ID_Active_Also_Serial                                                      '序列同步
            control.Checked = Not control.Checked
            ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Also_Serial, , True).Checked = control.Checked
            ComToolBar.Item(ToolBar_Object).FindControl(, ID_Active_Also_Serial, , True).Checked = control.Checked
            Button_miSerialPlaceInPhase = control.Checked
            
            '自动选中全部序列
            isSelectAllSerial = Not Button_miSerialPlaceInPhase
            subSelectAllSerial Me
            
            '如果打开了自动序列同步，则关闭手工序列同步和三维鼠标
            If control.Checked = True Then
                '关闭三维鼠标
                Set CmdControl = ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_3DLine)
                CmdControl.Checked = False
                '关闭手工序列同步
                ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Also_ManualSerial, , True).Checked = False
                ComToolBar.Item(ToolBar_Object).FindControl(, ID_Active_Also_ManualSerial, , True).Checked = False
                Button_miSerialManualSyn = False
            End If
            
        Case ID_Active_Also_ManualSerial                                                '手工序列同步
            '开关式按钮,跟序列同步是互斥的
            Button_miSerialManualSyn = Not Button_miSerialManualSyn
            control.Checked = Button_miSerialManualSyn
            ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Also_ManualSerial, , True).Checked = Button_miSerialManualSyn
            ComToolBar.Item(ToolBar_Object).FindControl(, ID_Active_Also_ManualSerial, , True).Checked = Button_miSerialManualSyn
            
            '自动选中全部序列
            isSelectAllSerial = Not Button_miSerialManualSyn
            subSelectAllSerial Me
            
            '如果打开手工序列同步，则关闭自动序列同步和三维鼠标
            If control.Checked = True Then
                '关闭三维鼠标的状态
                Set CmdControl = ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_3DLine)
                CmdControl.Checked = False
                '关闭自动序列同步
                ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Also_Serial, , True).Checked = False
                ComToolBar.Item(ToolBar_Object).FindControl(, ID_Active_Also_Serial, , True).Checked = False
                Button_miSerialPlaceInPhase = False
            End If
                
        Case ID_Active_Also_LockSerial                                                  '锁定序列
            If intSelectedSerial > 0 And intSelectedSerial < Viewer.Count Then
                ZLShowSeriesInfos(intSelectedSerial).Selected = Not ZLShowSeriesInfos(intSelectedSerial).Selected
                subDispframe Me, Viewer(intSelectedSerial)
                Viewer(intSelectedSerial).Refresh
            End If
            
        Case ID_Active_Also_Photo                                                       '图像同步
            '图像内容同步作为一个公共功能进行设置
            Button_miImageInPhase = Not Button_miImageInPhase
            ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Also_Photo, , True).Checked = Button_miImageInPhase
            ComToolBar.Item(ToolBar_Object).FindControl(, ID_Active_Also_Photo, , True).Checked = Button_miImageInPhase
        Case ID_Tool_NothinMouseState
            subSelectLeftorRightBouttom 1, control.Id
            subSelectLeftorRightBouttom 2, control.Id
        Case ID_Active_Shuttle                                                          '穿梭
            subSelectLeftorRightBouttom cMouseUsage("101").lngMouseKey, control.Id
            Button_miStack = True
        Case ID_ACtive_Mouse_Value                                                      '显示CT值
            control.Checked = Not control.Checked
            ComToolBar.Item(ToolBar_Menu).FindControl(, ID_ACtive_Mouse_Value, , True).Checked = control.Checked
            ComToolBar.Item(ToolBar_Scale).FindControl(, ID_ACtive_Mouse_Value, , True).Checked = control.Checked
            Button_miMouseShowValue = control.Checked
        Case ID_Active_Cruise                                                           '漫游
            subSelectLeftorRightBouttom cMouseUsage("103").lngMouseKey, control.Id
            Button_miCruise = True
            
        Case ID_Active_Cut                                                              '裁剪
            If SelectedImage Is Nothing Then Exit Sub
            Button_miCutOut = Not Button_miCutOut
            ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Cut, , True).Checked = Button_miCutOut
            subSelectLeftorRightBouttom cMouseUsage("201").lngMouseKey, control.Id
            subCutOut Me
            
        Case ID_ACtive_FrameSelectImage                                                 '框选图象
            If SelectedImage Is Nothing Then Exit Sub
            Button_miFrameSelectImage = Not Button_miFrameSelectImage
            ComToolBar.Item(ToolBar_Object).FindControl(, ID_ACtive_FrameSelectImage, , True).Checked = Button_miFrameSelectImage
            ComToolBar.Item(ToolBar_Menu).FindControl(, ID_ACtive_FrameSelectImage, , True).Checked = Button_miFrameSelectImage
            '处理相同鼠标键位的按钮状态
            subSelectLeftorRightBouttom cMouseUsage("201").lngMouseKey, control.Id
            
        Case ID_ACtive_SaveInReport                                                     '保存框选的图象进入报告图
            If SelectedImage Is Nothing Or SelectedLabel Is Nothing Or SelectedLabel.LabelType <> doLabelRectangle Then Exit Sub
            SaveFrameSelectImageIntoReport SelectedImage, SelectedLabel
            RaiseEvent AfterSaveReportImage(PstrCheckUID)
            
        Case ID_Active_Zoom                                                             '缩放
            subSelectLeftorRightBouttom cMouseUsage("104").lngMouseKey, control.Id
            Button_miZoom = True
            
        Case ID_Active_ReSetAll                                                         '恢复所有
            If Not SelectedImage Is Nothing Then
                SelectedImage.Mask = 0
                SelectedImage.SetDefaultWindows
                SelectedImage.FlipState = doFlipNormal
                SelectedImage.RotateState = doRotateNormal
                SelectedImage.ScrollX = 0
                SelectedImage.ScrollY = 0
                SelectedImage.StretchToFit = True
                '处理序列内图像同步
                Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_All)
                '处理菜单和工具条
                For i = 349 To 360
                    If Not ComToolBar.Item(ToolBar_Comm).FindControl(, i, , True) Is Nothing Then
                        ComToolBar.Item(ToolBar_Comm).FindControl(, i, , True).Checked = False
                    End If
                Next
                If Not ComToolBar.Item(ToolBar_Comm).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow_ReSet, , True) Is Nothing Then
                    ComToolBar.Item(ToolBar_Comm).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow_ReSet, , True).Checked = True
                End If
                subShowScale ID_View_ShowScale_AutoShow
            End If
            
        Case ID_Active_AdjustWindow_HandAdjustWindow                                    '手动调窗
            subSelectLeftorRightBouttom cMouseUsage("102").lngMouseKey, control.Id
            Button_miWidthLevel = True
            
        Case ID_Active_AdjustWindow_AutoAdjustWindow                                    '自适应调窗
            subSelectLeftorRightBouttom cMouseUsage("105").lngMouseKey, control.Id
            Button_miAutoWidthLevel = True
            
        Case ID_Active_PointingLine_ALL                                                 '所有定位线
            If blnSelectedImageIfColor = False Then
                subCurrentCheck control, Me
            End If
            
        Case ID_Active_PointingLine_FirstLast                                           '首位定位线
            If blnSelectedImageIfColor = False Then
                subCurrentCheck control, Me
            End If
            
        Case ID_Active_PointingLine_Now                                                 '当前定位线
            If blnSelectedImageIfColor = False Then
                subCurrentCheck control, Me
            End If
            
        Case ID_Active_PointingLine_3DLine                                              '3D鼠标
            subSelectLeftorRightBouttom cMouseUsage("106").lngMouseKey, control.Id
            
        Case ID_Active_Eddy_LeftRight                                                   '左右旋转
            subManipulation "FlipHorizontal", Me
            
        Case ID_Active_Eddy_TopButton                                                   '垂直旋转
            subManipulation "FlipVertical", Me
            
        Case ID_Active_Eddy_Left90                                                      '左旋90
            subManipulation "RotateAnticlockwise", Me
            
        Case ID_Active_Eddy_Right90                                                     '右旋90
            subManipulation "RotateClockwise", Me
            
        Case ID_Active_ReverseVideo                                                     '反白
            control.Checked = Not control.Checked
            ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_ReverseVideo, , True).Checked = control.Checked
            ComToolBar.Item(ToolBar_Photo).FindControl(, ID_Active_ReverseVideo, , True).Checked = control.Checked
            subManipulation "Invert", Me
                    
        Case ID_Active_SieveLens_LancetMinus                                            '边缘增强强度减少
            zl9ComLib.zlCommFun.ShowFlash "正在处理图像，请等待！", Me
            SubImageUnsharp "miUnSharpEnhancementDown", Me
            zl9ComLib.zlCommFun.StopFlash
            
        Case ID_Active_SieveLens_LancetAdd                                              '边缘增强强度增加
            zl9ComLib.zlCommFun.ShowFlash "正在处理图像，请等待！", Me
            zl9ComLib.zlCommFun.ShowFlash
            SubImageUnsharp "miUnSharpEnhancementUp", Me
            zl9ComLib.zlCommFun.StopFlash
            
        Case ID_Active_SieveLens_FlatnessMinus                                          '平滑减少
            zl9ComLib.zlCommFun.ShowFlash "正在处理图像，请等待！", Me
            zl9ComLib.zlCommFun.ShowFlash
            SubImageUnsharp "miFilterLengthDown", Me
            zl9ComLib.zlCommFun.StopFlash
            
        Case ID_Active_SieveLens_FlatnessAdd                                            '平滑增加
            zl9ComLib.zlCommFun.ShowFlash "正在处理图像，请等待！", Me
            zl9ComLib.zlCommFun.ShowFlash
            SubImageUnsharp "miFilterLengthUp", Me
            zl9ComLib.zlCommFun.StopFlash
            
        Case ID_Active_Sievelens_LeftMoveMinus                                          '边缘增强幅度减少
            zl9ComLib.zlCommFun.ShowFlash "正在处理图像，请等待！", Me
            zl9ComLib.zlCommFun.ShowFlash
            SubImageUnsharp "miUnSharpLengthDown", Me
            zl9ComLib.zlCommFun.StopFlash
            
        Case ID_Active_Sievelens_LeftMoveAdd                                            '边缘增强幅度增加
            zl9ComLib.zlCommFun.ShowFlash "正在处理图像，请等待！", Me
            zl9ComLib.zlCommFun.ShowFlash
            SubImageUnsharp "miUnSharpLengthUp", Me
            zl9ComLib.zlCommFun.StopFlash
            
        Case ID_Active_Sievelens_PhotoReset                                             '图像还原
            SubImageUnsharp "miRestore", Me
            
        Case ID_Active_Lable_Text                                                       '文字
            subSelectLeftorRightBouttom cMouseUsage("8").lngMouseKey, control.Id
            Button_miLabeltext = True
            
        Case ID_Active_Lable_Arrowhead                                                  '箭头
            subSelectLeftorRightBouttom cMouseUsage("4").lngMouseKey, control.Id
            Button_miLabelArrowhead = True
         
        Case ID_Active_Lable_Ellipse                                                    '椭圆
            subSelectLeftorRightBouttom cMouseUsage("3").lngMouseKey, control.Id
            Button_miLabelEllipse = True
        
        Case ID_Active_Lable_Angle                                                      '角度
            subSelectLeftorRightBouttom cMouseUsage("7").lngMouseKey, control.Id
            Button_miLabelAngle = True
        
        Case ID_Active_Lable_Curve                                                      '曲线
            subSelectLeftorRightBouttom cMouseUsage("6").lngMouseKey, control.Id
            Button_miLabelPolyLine = True
            
        Case ID_Active_Lable_Area                                                       '区域
            subSelectLeftorRightBouttom cMouseUsage("5").lngMouseKey, control.Id
            Button_miLabelPolygon = True
        
        Case ID_Active_Lable_BeeLine                                                    '直线
            subSelectLeftorRightBouttom cMouseUsage("1").lngMouseKey, control.Id
            Button_miLabelLine = True
        
        Case ID_Active_Lable_Rect                                                       '矩形
            subSelectLeftorRightBouttom cMouseUsage("2").lngMouseKey, control.Id
            Button_miLabelRectangle = True
        
        Case ID_Active_Lable_VasMeasure                                                 '血管狭窄测量
            subSelectLeftorRightBouttom cMouseUsage("1").lngMouseKey, control.Id
            Button_miLabelVasMeasure = True
            
        Case ID_Active_Lable_CadioThoracicRatio                                         '心胸比测量
            subSelectLeftorRightBouttom cMouseUsage("1").lngMouseKey, control.Id
            Button_miLabelCadiothoracicRatio = True
        
        Case ID_Active_Lable_AreaBeeLinePhoto                                           '区域直方图
            If Not SelectedLabel Is Nothing And blnSelectedImageIfColor = False Then
                If SelectedLabel.LabelType = 1 Or SelectedLabel.LabelType = 2 Or SelectedLabel.LabelType = 5 Then funcROIHistogram SelectedLabel
                If SelectedLabel.LabelType = 3 Or SelectedLabel.LabelType = 4 Then funcDrawGreyDistribute SelectedImage, SelectedLabel
            Else
                If blnSelectedImageIfColor = True Then
                    MsgBox "该图像是彩色图像，不能够进行直方图计算。", vbInformation, gstrSysName
                End If
            End If
            
        Case ID_Active_Lable_AdjustLine                                                 '校准
            subcalibrate Me
            
        Case ID_Active_Lable_ClearLbale                                                 '清除所有标注
            Call subLabelDeleAll
            
        Case ID_Active_Lable_DelSelectLable                                             '删除当前标注
            Call subDelSelectedLabel
            
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''工具菜单''''''''''''''''''''''''''''''''''''
        Case ID_Tool_Movie                                                              '电影
            If intSelectedSerial <> 0 Then
                If SelectedImage.FrameCount = 1 And ZLShowSeriesInfos(intSelectedSerial).ImageInfos.Count < 2 Then Exit Sub
                Set frmCine.f = Me
                frmCine.Show 1, Me
            End If
            
        Case ID_Tool_Magnifier                                                          '放大镜
            Dim fMagnifier As New FrmMagnify
            Set fMagnifier.f = Me
            fMagnifier.Show , Me
            
        Case ID_Tool_ArrowyCoronaryReset                                                '矢冠状重建
            subSelectOnlyOne ID_Tool_ArrowyCoronaryReset
            Call funViewerMPR(Me)
            
        Case ID_Tool_SlopeReconstruction                                                '斜面重建
            '斜面重建
            Call funMPRslope(Me)
            
        Case ID_Tool_NumberMinusShadow                                                  '数字减影
            subDSA Me
            
        Case ID_Tool_BogusColour                                                        '伪彩
            If blnSelectedImageIfColor = False Then
                subFakeColor Me
            Else
                MsgBox "该图像已经是彩色图像，不能够进行伪彩操作。", vbInformation, gstrSysName
            End If
        
        Case ID_Tool_FilmPrint                                                          '胶片打印
            blnPrintFilm = funFilm(Me, True, 3)
            
        Case ID_Tool_Film_AddSeries                                                     '胶片打印--打印序列
            Call funFilm(Me, False, 1)
            
        Case ID_Tool_Film_AddImage                                                      '胶片打印 - 打印图像
            Call funFilm(Me, False, 2)
            
        Case ID_Tool_Film_AddSelected                                                   '胶片打印 - 打印所选图
            Call funFilm(Me, False, 3)
            
        Case ID_Tool_PhotoUnite                                                         '图像拼接
            If Not blnfis Then
                Set fis.f = Me
                fis.Show , Me
                blnfis = True
                '触发保存外部图像的事件，因为图像拼接可能保存了结果图
                RaiseEvent AfterSaveOuterImage(PstrCheckUID)
            End If
            
        Case ID_Tool_LableTool                                                          '标注工具
            Set frmLabelObject.im = Me.SelectedImage
            Set frmLabelObject.f = Me
            frmLabelObject.Show , Me
            
        Case ID_Tool_LookPhotoOption                                                    '观片选项
            Set frmSysConfig.f = Me
            frmSysConfig.Show 1, Me
            
        Case ID_ToolBar_Left                                                            '工具条靠左
            PutToolbar ComToolBar, 2
            ArrayToolBar ComToolBar, Me.top, Me.left, Me.height, Me.width
            intToolBarPosition = 2
            Call subSaveInterfaceParaIntoDB
            
        Case ID_ToolBar_Right                                                           '工具条靠右
            PutToolbar ComToolBar, 3
            ArrayToolBar ComToolBar, Me.top, Me.left, Me.height, Me.width
            intToolBarPosition = 3
            Call subSaveInterfaceParaIntoDB
        
        Case ID_ToolBar_Top                                                             '工具条靠上
            PutToolbar ComToolBar, 0
            ArrayToolBar ComToolBar, Me.top, Me.left, Me.height, Me.width
            intToolBarPosition = 0
            Call subSaveInterfaceParaIntoDB
            
        Case ID_ToolBar_Button                                                          '工具条靠下
            PutToolbar ComToolBar, 1
            ArrayToolBar ComToolBar, Me.top, Me.left, Me.height, Me.width
            intToolBarPosition = 1
            Call subSaveInterfaceParaIntoDB
            
        Case ID_toolBar_16Icon                                                          '工具条图标16*16显示
            ReplaceToolBarIcon ComToolBar, ImgList16, 16, 16
'            ArrayToolBar ComToolBar, Me.top, Me.left, Me.height, Me.width
            ComToolBar.RecalcLayout
            intToolBarIconSize = 16
            Call subSaveInterfaceParaIntoDB
            
        Case ID_ToolBar_24Icon                                                          '工具条图标24*24显示
            ReplaceToolBarIcon ComToolBar, ImgList24, 24, 24
'            ArrayToolBar ComToolBar, Me.top, Me.left, Me.height, Me.width
            ComToolBar.RecalcLayout
            intToolBarIconSize = 24
            Call subSaveInterfaceParaIntoDB
        
        Case ID_ToolBar_32Icon                                                          '工具条图标32*32显示
            ReplaceToolBarIcon ComToolBar, ImgList32, 32, 32
            'ArrayToolBar ComToolBar, Me.top, Me.left, Me.height, Me.width
            ComToolBar.RecalcLayout
            intToolBarIconSize = 32
            Call subSaveInterfaceParaIntoDB
            
        Case ID_ToolBar_Hide                                                            '工具条隐藏显
            control.Checked = Not control.Checked
            blToolBarHide = Not control.Checked
            blfrmRefresh = False
            For i = 2 To 8
                If i = 8 Then
                    blfrmRefresh = True
                End If
                ComToolBar.Item(i).Visible = Not control.Checked
            Next
            If control.Checked = False Then
                ArrayToolBar ComToolBar, Me.top, Me.left, Me.height, Me.width
            End If
            ComToolBar.RecalcLayout
            Call subSaveInterfaceParaIntoDB
            Me.Refresh
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''帮助'''''''''''''''''''''''''''''''''''''
        Case ID_Help_Help                                                               '帮助
            '功能：调用帮助主题
            Shell "hh.exe  zl9ImgViewer.chm", vbNormalFocus
            'ShowHelp App.ProductName, Me.hwnd, "frmInstrument"
        Case ID_Help_WebZLSOFT_WEB                                                      '中联主页
'            Call zlHomePage(Me.hwnd)
        
        Case ID_Help_WebZLSOFT_Mail                                                     '发送反馈
'            Call zlMailTo(Me.hwnd)
            
        Case ID_Help_UpdateDB                                                           '升级Access数据库
            Set frmUpdateDB.m_cnAccess = cnAccess
            frmUpdateDB.Show 1, Me
        Case ID_Help_About                                                              '关天
'            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End Select
End Sub

Private Sub ComToolBar_GetClientBordersWidth(left As Long, top As Long, Right As Long, Bottom As Long)
    If sbStatusBar.Visible And blfrmRefresh = True Then
        Bottom = sbStatusBar.height
    End If
End Sub

Private Sub ComToolBar_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    
    If CommandBar.Title = "间隔打印" Then
        Dim controlForm As CommandBarControlCustom
        CommandBar.Controls.DeleteAll
        Set controlForm = CommandBar.Controls.Add(xtpControlCustom, 0, "间隔打印")
        controlForm.Handle = picPrintInterval.hwnd
        picPrintInterval.BackColor = ComToolBar.GetSpecialColor(XPCOLOR_MENUBAR_FACE)
        optPrintStart(1).BackColor = picPrintInterval.BackColor
        optPrintStart(2).BackColor = picPrintInterval.BackColor
        lblPrtintInterval.BackColor = picPrintInterval.BackColor
        txtPrintInterval.BackColor = picPrintInterval.BackColor
        cmdPrintInterval.BackColor = picPrintInterval.BackColor
        Exit Sub
    End If
End Sub

Private Sub ComToolBar_Resize()
'    On Error Resume Next
'
'    Dim left As Long
'    Dim top As Long
'    Dim Right As Long
'    Dim Bottom As Long
'
'    If blfrmRefresh = True Then
'        ComToolBar.GetClientRect left, top, Right, Bottom
'        If Right >= left And Bottom >= top Then
'            picViewer.Move left, top, Right - left, Bottom - top
'        Else
'            picViewer.Move 0, 0, 0, 0
'        End If
'    End If
End Sub

Private Sub ComToolBar_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim iIndex As Integer
    
    Select Case control.Id
        Case ID_View_UpSeries, ID_View_DownSeries   '上一序列，下一序列
            control.Visible = IIf(ZLSeriesInfos.Count <= 1, False, True)
        Case ID_Active_AdjustWindow_HandAdjustWindow, ID_Active_AdjustWindow_AutoAdjustWindow, ID_Active_Eddy_LeftRight, ID_Active_Eddy_TopButton, ID_Active_Eddy_Left90, _
            ID_Active_Eddy_Right90, ID_Active_ReverseVideo, ID_Active_Cut, ID_Active_PointingLine_ALL, ID_Active_PointingLine_FirstLast, _
            ID_Active_PointingLine_Now, ID_Active_SieveLens_LancetMinus, ID_Active_SieveLens_LancetAdd, ID_Active_SieveLens_LancetAdd, _
            ID_Active_SieveLens_FlatnessMinus, ID_Active_SieveLens_FlatnessAdd, ID_Active_Sievelens_LeftMoveMinus, ID_Active_Sievelens_LeftMoveAdd, _
            ID_Active_Sievelens_PhotoReset, ID_Active_SieveLens_Model            '图像操作处理
            
            control.Visible = (InStr(mstrPrivs, "图像操作处理") <> 0)
            
        Case ID_Active_Lable_Text, ID_Active_Lable_Arrowhead, ID_Active_Lable_Ellipse, ID_Active_Lable_Angle, ID_Active_Lable_BeeLine, _
            ID_Active_Lable_Rect, ID_Active_Lable_AdjustLine, ID_Active_Lable_ClearLbale, ID_Active_Lable_DelSelectLable, _
            ID_Active_Lable_AreaBeeLinePhoto, ID_Active_Lable_Area, ID_Active_Lable_Curve   '图像标注测量
            
            control.Visible = (InStr(mstrPrivs, "图像标注测量") <> 0)
            
        Case ID_Tool_ArrowyCoronaryReset        '矢冠状重建
            control.Visible = (InStr(mstrPrivs, "矢冠状重建") <> 0)
            
            control.Checked = blnInMPR  '修改重建按钮的状态
        Case ID_Tool_SlopeReconstruction        '斜面重建
            control.Visible = (InStr(mstrPrivs, "矢冠状重建") <> 0)
        Case ID_Tool_BogusColour        '伪彩
            control.Visible = (InStr(mstrPrivs, "伪彩") <> 0)
        Case ID_Active_PointingLine_3DLine  '三维鼠标
            control.Visible = (InStr(mstrPrivs, "三维鼠标") <> 0)
        Case ID_Tool_NumberMinusShadow      '数字减影
            control.Visible = (InStr(mstrPrivs, "数字减影") <> 0)
            If control.Visible = True And Not SelectedImage Is Nothing Then
                control.Enabled = IIf(SelectedImage.FrameCount > 1, True, False)
            End If
        Case ID_File_OpenDicomDir           'DICOM_DIR
            control.Visible = (InStr(mstrPrivs, "DICOM_DIR") <> 0)
        Case ID_Tool_FilmPrint              '胶片排版打印
            control.Visible = (InStr(mstrPrivs, "胶片排版打印") <> 0)
        Case ID_Active_Lable_VasMeasure     '血管狭窄测量
            control.Visible = (InStr(mstrPrivs, "血管狭窄测量") <> 0)
        Case ID_Tool_PhotoUnite             '图像拼接
            control.Visible = (InStr(mstrPrivs, "图像拼接") <> 0)
        Case ID_File_Send_GetHost     '单机版，反权限
            control.Visible = (InStr(mstrPrivs, "单机版") = 0)
        Case ID_File_Open                   '独立观片站，反权限
            control.Visible = (InStr(mstrPrivs, "独立观片站") = 0)
        Case ID_File_DelReport, ID_File_SAveASReport        '保存报告图，删除报告图，需要有“报告图像”权限，且不是“单机版”
            If InStr(mstrPrivs, "单机版") <> 0 Then
                control.Visible = False
            Else
                control.Visible = (InStr(mstrPrivs, "报告图像") <> 0)
            End If
        Case ID_File_SaveFile               '保存图像
            If InStr(mstrPrivs, "单机版") <> 0 Then
                control.Visible = False
            Else
                control.Visible = (InStr(mstrPrivs, "保存图像") <> 0)
            End If
        Case ID_File_SaveFile
            control.Visible = (InStr(mstrPrivs, "保存图像") <> 0)
        Case ID_File_SaveASFile, ID_File_SaveToCD, ID_File_Send_GetHost, ID_File_Send_OutPowerPoint
            control.Visible = (InStr(mstrPrivs, "另存图像") <> 0)
    End Select
End Sub

Private Sub InitFaceSheme()
    Dim Pane1 As Pane
    
    With Me.dkpMain
        .CloseAll
        .SetCommandBars ComToolBar
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = dkpMain.CreatePane(1, 200, 200, DockBottomOf, Nothing)
    Pane1.Handle = picViewer.hwnd
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '处理窗体的键盘事件
    '按下ESC时清除所有可执行鼠标状态
    If KeyCode = vbKeyEscape Then        'ESC
        subSelectLeftorRightBouttom 1, ID_Tool_NothinMouseState
        subSelectLeftorRightBouttom 2, ID_Tool_NothinMouseState
        '如果是全屏，则退出全屏
        If Button_miFullScreen = True Then
            Call subFullScreen(Me)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim strRegPath As String
    
    On Error GoTo err
    
    '初始化参数'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    intSpaceSize = 100   ''''设置viewer之间的间隔
    blnAngle = False
    intVasMeasure = 0
    intCadioThoracicRatio = 0
    
    Call RestoreWinState(Me, App.ProductName)
    
    InitFaceSheme   '初始化界面布局
    
    '''''''''''''''''[信息标注位置设置]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim DG As New DicomGlobal
    DG.DirectionStrings = IIf(blnChinaMark, "右\左\前\后\脚\头", "R\L\A\P\I\S")
    
    '设置系统默认值,默认（未选中，非当前）图像边框的颜色，线型，线宽
    lngDefaultImageBorderColor = vbWhite
    lngDefaultImageBorderLineStyle = 0
    lngDefaultImageBorderLineWidth = 1
    
    ''''''''''''''''''从数据库中获取预设窗宽窗位，预设屏幕布局，填写到系统变量中
    subGetLayoutToVar glngUserID                    '从数据库中读取序列和图像布局到系统变量
    Call subGetWWWLToVal                            '从数据库中读取窗宽窗位到系统变量
    Call subGetFilterToVal                          '从数据库中读取滤镜设置到系统变量
    subGetImageShutterToVar glngUserID              '从数据库中读取图像消隐设置到系统变量
    subGetMouseUsageToVar glngUserID                '从数据库中读取鼠标用法设置的值到系统变量
    subGetInterfaceParaToVar glngUserID             '从数据库读取“影像界面参数表”的内容，并将其保存到系统参数中。
    subGetLabelStoreToVar                           '从数据库读取保存标注的相关信息
    subGetInfoLabelToVar                            '从数据库获取信息标注位置设置数据到系统变量
    subGetDBDicomPrintToVar                         '从数据库获取DICOM打印的打印机设置到系统变量
    subGetParameters                                '从系统参数表读取参数
    '从注册表中，获取胶片打印的标注字体大小
    intFilmFontSize = GetSetting("ZLSOFT", "公共模块\zlPacsCore", "胶片字体", "10")
    
    
    ''''''''''''''''''''''''''处理鼠标滚轮'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If App.LogMode <> 0 Then
        Dim ret As Long
    '    '记录原来的window程序地址
        preWinProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    '    '用自定义程序代替原来的window程序
        ret = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf Wndproc)
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ReDim FilmLayouts(0)        '定义胶片打印图像布局数组
    Set ZLSeriesInfos = New Collection  '初始化图像序列信息集合
    Set ZLShowSeriesInfos = New Collection '初始化图像显示序列信息集合
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''设置状态栏图标'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Set sbStatusBar.Panels(1).Picture = ImgList24.ListImages("中联图标").Picture
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    ''''''''''''''''''''''''''[初始化序列控制]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MSFViewer.Rows = 1
    MSFViewer.Cols = 16
    picViewer.BackColor = lngProgramBackColor
    ''''''''''''''''''''''初使化鼠标左右分布''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''''【初始化菜单按钮】''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    LoadBarSetup Me
    
    ''''''''''''''''''''''''''''''''''[初使化状态栏字体大小]'''''''''''''''''''''''''''''''''''''''''''''
    Me.sbStatusBar.Font.Size = IIf(intStatusBarFontSize < 1, 10, intStatusBarFontSize)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''提取个性化设置'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    strRegPath = "私有模块\ZLHIS\" & App.EXEName & "\frmViewer"
    
    If GetSetting("ZLSOFT", strRegPath, "WindowState", 2) = 1 Then
        '如果观片站的显示状态是最小化，则恢复到默认的位置
        Me.top = 0
        Me.left = 0
        Me.width = 1024 * Screen.TwipsPerPixelX
        Me.height = 768 * Screen.TwipsPerPixelY
        Me.WindowState = 2
    Else
        Me.top = GetSetting("ZLSOFT", strRegPath, "Top", 0)
        Me.left = GetSetting("ZLSOFT", strRegPath, "Left", 0)
        Me.width = GetSetting("ZLSOFT", strRegPath, "Width", 1024 * Screen.TwipsPerPixelX)
        Me.height = GetSetting("ZLSOFT", strRegPath, "Height", 768 * Screen.TwipsPerPixelY)
        Me.WindowState = GetSetting("ZLSOFT", strRegPath, "WindowState", 2)
        '如果窗口的状态不正确，调整窗口返回到正常状态
        If Abs(Me.left) > 200000 Or Abs(Me.top) > 200000 Then
            Me.top = 0
            Me.left = 0
            Me.width = 1024 * Screen.TwipsPerPixelX
            Me.height = 768 * Screen.TwipsPerPixelY
            Me.WindowState = 2
        End If
    End If
    Button_miMouseShowValue = GetSetting("ZLSOFT", strRegPath, "显示鼠标像素值", False)
    Button_miShowMiniSeries = GetSetting("ZLSOFT", strRegPath, "显示序列缩略图", False)
    Button_miViewAllSeries = GetSetting("ZLSOFT", strRegPath, "全序列观片", False)
    Button_miImageInPhase = GetSetting("ZLSOFT", strRegPath, "图像格式同步", True)
    
    
    Button_miDispPatientInfo = True '默认显示图像信息
    Button_miShowOverlay = True     '默认显示Overlay
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    subInitSerial Me    ''对窗体各种内容进行初始化
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '判断观片站是否注册成功，注册不成功则关闭观片站
    gint医技观片站数量 = getLicenseCount(LOGIN_TYPE_医技观片站)
    gint胶片打印机 = getLicenseCount(LOGIN_TYPE_胶片打印机)
    mstr启动时间 = FunLogIn(LOGIN_TYPE_医技观片站)
    If mstr启动时间 = "" Then
        blnLogined = False
    Else
        blnLogined = True
    End If
    
    '读取自定义的工具栏
    ComToolBar.LoadCommandBars "中联观片站", App.Title, "自定义工具栏"

    '允许自定义工具栏
    ComToolBar.EnableCustomization True

    '禁止主工具栏和通用工具栏的内容被自定义
    ComToolBar.Item(7).Customizable = False
    ComToolBar.ActiveMenuBar.Customizable = False

    '统一图标大小
    For i = 9 To ComToolBar.Count
        ComToolBar.Item(i).SetIconSize intToolBarIconSize, intToolBarIconSize
    Next i
    
    Exit Sub
err:
    
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Public Sub InitComButtonChecked()
    '初始化鼠标左右分配的按键按钮
    Dim i As Integer
    
    '先清除所有鼠标状态
    subSelectLeftorRightBouttom 1, ID_Tool_NothinMouseState
    subSelectLeftorRightBouttom 2, ID_Tool_NothinMouseState
        
    ''''''''''''''''''''''''''''''''''[初始化鼠标左右分配的按键按钮]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To cMouseUsage.Count
        If cMouseUsage(i).strProgramName <> "No" Then
            ComToolBar.FindControl(, cMouseUsage(i).ButtomID, , True).Checked = cMouseUsage(i).bSelected
        ElseIf cMouseUsage(i).lngFuncNo = 201 Then  '裁剪和框选按钮
            ComToolBar.FindControl(, ID_ACtive_FrameSelectImage, , True).Checked = cMouseUsage(i).bSelected
        End If
    Next
    
    '穿梭
    If ComToolBar.FindControl(, ID_Active_Shuttle, , True).Checked = True Then
        Button_miStack = True
    End If
    '调窗
    If ComToolBar.FindControl(, ID_Active_AdjustWindow_HandAdjustWindow, , True).Checked = True Then
        Button_miWidthLevel = True
    End If
    '漫游
    If ComToolBar.FindControl(, ID_Active_Cruise, , True).Checked = True Then
        Button_miCruise = True
    End If
    '缩放
    If ComToolBar.FindControl(, ID_Active_Zoom, , True).Checked = True Then
        Button_miZoom = True
    End If
    '自适应调窗
    If ComToolBar.FindControl(, ID_Active_AdjustWindow_AutoAdjustWindow, , True).Checked = True Then
        Button_miAutoWidthLevel = True
    End If
    '框选图像
    If ComToolBar.FindControl(, ID_ACtive_FrameSelectImage, , True).Checked = True Then
        Button_miFrameSelectImage = True
    End If
    '默认为显示CT值
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_ACtive_Mouse_Value, , True).Checked = Button_miMouseShowValue
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_ACtive_Mouse_Value, , True).Checked = Button_miMouseShowValue
    '显示序列缩略图
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_ShowMiniSeries, , True).Checked = Button_miShowMiniSeries
    ComToolBar.Item(ToolBar_Object).FindControl(, ID_View_ShowMiniSeries, , True).Checked = Button_miShowMiniSeries
    '全序列观片
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_ViewAllSeries, , True).Checked = Button_miViewAllSeries
    ComToolBar.Item(ToolBar_Object).FindControl(, ID_View_ViewAllSeries, , True).Checked = Button_miViewAllSeries
    
    ''''''''''''''''''''''''''''''''''初使化工具条风络'''''''''''''''''''''''''''''''''''''''''''''''''
    IntComBarTheme = Me.ComToolBar.VisualTheme
    '''''''''''''''''''''''''''''''''初始化一些默认值''''''''''''''''''''''''''''''''''''''''''''''''''
    ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_PhotoSerial_PhotoNumber, , True).Checked = True
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_PhotoSerial_PhotoNumber, , True).Checked = True
    ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_ShowScale_AutoShow, , True).Checked = True
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_ShowScale_AutoShow, , True).Checked = True
    
    '图像格式同步
    ComToolBar.Item(ToolBar_Object).FindControl(, ID_Active_Also_Photo, , True).Checked = Button_miImageInPhase
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Also_Photo, , True).Checked = Button_miImageInPhase
    
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_LableShow, , True).Checked = True
    Button_miDispLabelInfo = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
    Dim strClearDate As String  '暂存上次清除缓存图像的日期
    
    '如果在MPR状态中，则提示是否保存MPR结果
    If blnInMPR = True Then
        If funViewerMPR(Me) = False Then Cancel = -1
    End If
    
    If Cancel = 0 Then
        
        ComToolBar.SaveCommandBars "中联观片站", App.Title, "自定义工具栏"
        
        If Dir(App.Path & "\temp\*.*") <> "" Then
            Kill App.Path & "\temp\*.*"
        End If
        If Dir(App.Path & "\temp", vbDirectory) <> "" Then
            RmDir App.Path & "\temp"
        End If
        
        '清空按钮
        subSelectLeftorRightBouttom 1, ID_Tool_NothinMouseState
        subSelectLeftorRightBouttom 2, ID_Tool_NothinMouseState
        
        '删除图像
        Call subKillPicture(True)
        
        '保存本机参数
        strRegPath = "私有模块\ZLHIS\" & App.EXEName & "\frmViewer"
        
        SaveSetting "ZLSOFT", strRegPath, "Top", Me.top
        SaveSetting "ZLSOFT", strRegPath, "Left", Me.left
        SaveSetting "ZLSOFT", strRegPath, "Width", Me.width
        SaveSetting "ZLSOFT", strRegPath, "Height", Me.height
        SaveSetting "ZLSOFT", strRegPath, "WindowState", Me.WindowState
        SaveSetting "ZLSOFT", strRegPath, "显示序列缩略图", Button_miShowMiniSeries
        SaveSetting "ZLSOFT", strRegPath, "显示鼠标像素值", Button_miMouseShowValue
        SaveSetting "ZLSOFT", strRegPath, "全序列观片", Button_miViewAllSeries
        SaveSetting "ZLSOFT", strRegPath, "图像格式同步", Button_miImageInPhase
        
        
        If frmMiniSeries.hwnd <> 0 Then
            Unload frmMiniSeries
        End If
        
        '清除缓存图像，每周清除一次
        strClearDate = GetSetting("ZLSOFT", strRegPath, "清除缓存图像", Date)
        If IsDate(strClearDate) = False Then
            strClearDate = Date
        End If
        If DateDiff("d", strClearDate, Date) >= 7 Then
            Call ClearCacheFolder(PstrBufferImagePath)
            SaveSetting "ZLSOFT", strRegPath, "清除缓存图像", Date
        End If
        
        '检查本次退出是否合法注册后的退出
        Call FunLogOut(LOGIN_TYPE_医技观片站, mstr启动时间)
    End If
End Sub

Private Sub mfrmFilm_AfterPrinted(strImageUIDS As String)
    '打印完成的事件，需要重新刷新图片的打印信息
    Dim arrImageUID() As String
    Dim intIndex As Integer
    Dim i As Integer
    Dim j As Integer
    Dim strSeriesUID As String
    Dim intSeriesIndex As Integer
    Dim k As Integer
    Dim blnFindImage As Boolean
    
    On Error GoTo err
    
    If Trim(strImageUIDS) = "" Then Exit Sub
    
    arrImageUID = Split(strImageUIDS, ",")
    If SafeArrayGetDim(arrImageUID) = 0 Then Exit Sub
    
    '逐个查找图像UID
    For intIndex = 1 To UBound(arrImageUID) - 1
        '先在已显示的图像中查找
        blnFindImage = False
        For i = 1 To ZLShowSeriesInfos.Count
            For j = 1 To ZLShowSeriesInfos(i).ImageInfos.Count
                If ZLShowSeriesInfos(i).ImageInfos(j).InstanceUID = arrImageUID(intIndex) Then
                    '从已显示图像集中找到图像，更改打印标记
                    
                    '更改打印标记
                    ZLShowSeriesInfos(i).ImageInfos(j).blnPrinted = True
                    
                    '更改图像显示的标注
                    '自动根据图像大小，判断是否显示病人四角信息,显示或者隐藏图像中的病人信息
                    Call subDisplayPatientInfo(Viewer(i))
                    
                    blnFindImage = True
                    Exit For
                End If
            Next j
            If blnFindImage Then
                Exit For
            End If
        Next i
        
        '在正本中查找
        blnFindImage = False
        For i = 1 To ZLSeriesInfos.Count
            For j = 1 To ZLSeriesInfos(i).ImageInfos.Count
                If ZLSeriesInfos(i).ImageInfos(j).InstanceUID = arrImageUID(intIndex) Then
                    ZLSeriesInfos(i).ImageInfos(j).blnPrinted = True
                    blnFindImage = True
                    Exit For
                End If
            Next j
            If blnFindImage = True Then
                Exit For
            End If
        Next i
        
    Next intIndex
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picViewer_DragDrop(Source As control, x As Single, y As Single)
'2009用

    Dim intSeriesIndex As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    
    If Source.Images.Count <= 0 Then Exit Sub
    intSeriesIndex = Source.Tag
    
    '计算这个Viewer摆放的位置
    Call subIsSerialXY(Me, x, y, intCol, intRow)
    
    Call subCreateAndPlaceAViewer(intSeriesIndex, intRow, intCol)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub picViewer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim imgs As New DicomImages
    Dim img As DicomImage
    Dim i As Integer
    '''''''''''''''''''''''''''''''[如果在空白序列处按动鼠标,则弹出产生序列的弹出菜单]''''''''''''''''''''''''''''''''''''''''''''''''
    If Button = 2 Then
        For i = 1 To ZLSeriesInfos.Count
            '加载一个图像
            Set img = funLoadAImage(i, 1, 0)
            If Not img Is Nothing Then
                imgs.Add img
            End If
        Next i
        lngBaseX = x
        lngBaseY = y
        PopMenu Me, imgs
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub picViewer_Resize()
    If blfrmRefresh = False Then Exit Sub   '刷新工具条的时候，窗体不需要刷新
    
    If intSelectedSerial < 1 Then Exit Sub
    Call subResizeSeries(Me)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PicX_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        PicXX.Visible = True
        PicXX.left = PicX(Index).left
        intFactMoveX = PicX(Index).left '记录鼠标按下横向分隔条的位置
        PicXX.ZOrder
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PicX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim v As DicomViewer
    Dim i As Integer
    If Button = 1 Then
        If Index = 1 Then
            If PicX(Index).left + x < 0 Then
                PicXX.left = 0
            ElseIf PicX(Index).left + x > Me.ScaleWidth - intSpaceSize Then
                PicXX.left = Me.ScaleWidth - intSpaceSize
            Else
                PicXX.left = PicX(Index).left + x
            End If
        Else
            If PicX(Index).left + x < PicX(Index - 1).left + intSpaceSize Then
                PicXX.left = PicX(Index - 1).left + intSpaceSize
            ElseIf PicX(Index).left + x > Me.ScaleWidth - intSpaceSize Then
                PicXX.left = Me.ScaleWidth - intSpaceSize
            Else
                PicXX.left = PicX(Index).left + x
            End If
        End If
        For Each v In Viewer
            If v.left <= PicXX.left And v.left + v.width >= PicXX.left + PicXX.width Then
                v.Refresh
                If VScro(v.Index).Visible Then VScro(v.Index).Refresh
            End If
        Next
        For i = 1 To intMaxAreaY - 1
            PicY(i).Refresh
        Next
        picViewer.Refresh
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PicX_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'------------------------------------------------
'功能：纵向拖动条的鼠标弹起事件，隐藏临时拖动条，显示真实拖动条，
'      并重新摆放viewer的大小,根据viewer大小判定是否自动隐藏病人信息
'参数：Index--自动生成，拖动条序号；
'返回：无
'------------------------------------------------
    Dim i, j, k As Long
    If Button = 1 Then
        PicX(Index).left = PicXX.left
        PicXX.Visible = False
        intFactMoveX = PicX(Index).left - intFactMoveX
        PicX(Index).Tag = PicX(Index).left
        For i = Index + 1 To intMaxAreaX - 1    '移动该分隔条右边的其他分隔条
            If PicX(i).Tag <> "" Then
                PicX(i).left = Val(PicX(i).Tag) + intFactMoveX
                PicX(i).Tag = PicX(i).left
                If PicX(i).left > Me.ScaleWidth - intSpaceSize Then
                    PicX(i).left = Me.ScaleWidth - intSpaceSize
                End If
            End If
        Next
        '重新摆放所有纵横分隔条的交点
        For i = 1 To intMaxAreaX - 1
            For j = 1 To intMaxAreaY - 1
                k = (j - 1) * (intMaxAreaX - 1) + i
                PicXY(k).top = PicY(j).top
                PicXY(k).left = PicX(i).left
            Next
        Next
        '重新摆放所有Viewer
        Call subMoveViewers(Me, 1, Index)
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PicXY_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        PicXX.Visible = True
        PicYY.Visible = True
        PicXX.left = PicXY(Index).left
        PicYY.top = PicXY(Index).top
        intFactMoveX = PicXY(Index).left
        intFactMoveY = PicXY(Index).top
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PicXY_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim indexX, indexY As Integer
    If Button = 1 Then
        indexX = Index Mod (intMaxAreaX - 1)
        If indexX = 0 Then indexX = intMaxAreaX - 1
        indexY = Int((Index - 1) / (intMaxAreaX - 1)) + 1
        PicX_MouseMove CInt(indexX), 1, 1, x, y
        PicY_MouseMove indexY, 1, 1, x, y
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PicXY_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim indexX, indexY As Integer
    If Button = 1 Then
        indexX = Index Mod (intMaxAreaX - 1)
        If indexX = 0 Then indexX = intMaxAreaX - 1
        indexY = Int((Index - 1) / (intMaxAreaX - 1)) + 1
        PicX_MouseUp CInt(indexX), 1, 1, x, y
        PicY_MouseUp indexY, 1, 1, x, y
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PicY_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        PicYY.Visible = True
        PicYY.top = PicY(Index).top
        intFactMoveY = PicY(Index).top
        PicYY.ZOrder
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PicY_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim v As DicomViewer
    Dim i As Integer
    If Button = 1 Then
        If Index = 1 Then
            If PicY(Index).top + y < 0 Then
                PicYY.top = 0
            ElseIf PicY(Index).top + y > Me.picViewer.ScaleHeight - intSpaceSize Then
                PicYY.top = Me.picViewer.ScaleHeight - intSpaceSize
            Else
                PicYY.top = PicY(Index).top + y
            End If
        Else
            If PicY(Index).top + y < PicY(Index - 1).top + intSpaceSize Then
                PicYY.top = PicY(Index - 1).top + intSpaceSize
            ElseIf PicY(Index).top + y > Me.picViewer.ScaleHeight - intSpaceSize Then
                PicYY.top = Me.picViewer.ScaleHeight - intSpaceSize
            Else
                PicYY.top = PicY(Index).top + y
            End If
        End If
        For Each v In Viewer
            If v.top <= PicYY.top And v.top + v.height >= PicYY.top + PicYY.height Then
                v.Refresh
                If VScro(v.Index).Visible Then VScro(v.Index).Refresh
            End If
        Next
        For i = 1 To intMaxAreaX - 1
            PicX(i).Refresh
        Next
        picViewer.Refresh
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PicY_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'------------------------------------------------
'功能：横向拖动条的鼠标弹起事件，隐藏临时拖动条，显示真实拖动条，
'      并重新摆放viewer的大小,根据viewer大小判定是否自动隐藏病人信息
'参数：Index--自动生成，拖动条序号；
'返回：无
'编制人：胡涛
'------------------------------------------------
    Dim i, j, k As Long
    If Button = 1 Then
        PicY(Index).top = PicYY.top
        PicYY.Visible = False
        intFactMoveY = PicY(Index).top - intFactMoveY
        PicY(Index).Tag = PicY(Index).top
        For i = Index + 1 To intMaxAreaY - 1        '移动该分隔条下面的其他分隔条
            If PicY(i).Tag <> "" Then
                PicY(i).top = Val(PicY(i).Tag) + intFactMoveY
                PicY(i).Tag = PicY(i).top
                If PicY(i).top > Me.picViewer.ScaleHeight - intSpaceSize Then
                    PicY(i).top = Me.picViewer.ScaleHeight - intSpaceSize
                End If
            End If
        Next
        '重新摆放所有纵横分隔条的交点
        For i = 1 To intMaxAreaX - 1
            For j = 1 To intMaxAreaY - 1
                k = (j - 1) * (intMaxAreaX - 1) + i
                PicXY(k).top = PicY(j).top
                PicXY(k).left = PicX(i).left
            Next
        Next
        '重新摆放所有Viewer
        Call subMoveViewers(Me, Index, 1)
    End If
End Sub

Private Sub subChangeASeries(intType As Integer)
'------------------------------------------------
'功能： 根据intType的类型，切换序列。
'参数：
'       intType－－切换方式。
'                   如果intType是1则用下一序列取代当前序列。
'                   如果intType=2则用上一序列取代当前序列。
'返回：
'------------------------------------------------
    Dim intSeriesIndex As Integer
    
    If intSelectedSerial = 0 Then Exit Sub
    
    On Error GoTo err
    intSeriesIndex = Viewer(intSelectedSerial).Tag
    If intType = 1 Then     '切换到下一序列
        intSeriesIndex = intSeriesIndex + 1
        If intSeriesIndex > ZLSeriesInfos.Count Then
            intSeriesIndex = 1
        End If
    ElseIf intType = 2 Then '切换到上一序列
        intSeriesIndex = intSeriesIndex - 1
        If intSeriesIndex <= 0 Then
            intSeriesIndex = ZLSeriesInfos.Count
        End If
    End If
    '用新序列的图像代替viewer(intSelectedSerial)中的图像
    Call funcSwapSeries(intSelectedSerial, intSeriesIndex)
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtPrintInterval_GotFocus()
    txtPrintInterval.SelStart = 0
    txtPrintInterval.SelLength = Len(txtPrintInterval.Text)
End Sub

Private Sub Viewer_DragDrop(Index As Integer, Source As control, x As Single, y As Single)
    Dim intSeriesIndex As Integer
    
    On Error GoTo error
    
    If Source.Name = "MiniVeiwer" And Source.Images.Count > 0 Then
        intSeriesIndex = Val(Source.Tag)
        If intSeriesIndex = 0 Then intSeriesIndex = 1
        
        '用新序列的图像代替viewer(index)中的图像
        Call funcSwapSeries(Index, intSeriesIndex)
    End If
    Exit Sub
error:
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Viewer_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '处理Viewer的键盘事件
    
    On Error GoTo err
    
    '处理Del，如果选中标注则删除标注，否则删除当前序列
    If KeyCode = 46 Then        'Delete
        If Not SelectedLabel Is Nothing Then    '如果当前选中了标注则删除当前标注
            Call subDelSelectedLabel
        Else
            subCloseSeries  '如果当前没有选中标注则删除当前序列
        End If
    End If
    
    '处理PageUp和PageDown，整页上下翻图像
    '处理上箭头和下箭头，上下翻单个图像
    ',End and Home 翻到第一个图，和最后一个图
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Or KeyCode = 35 Or KeyCode = 36 Then        ' 上箭头和下箭头,PageUp and PageDown ,End and Home
        If Viewer(intSelectedSerial).Visible = False Then Exit Sub
        If VScro(intSelectedSerial).Visible = False Then Exit Sub
        
        If KeyCode = 38 Then        '上箭头
            If VScro(intSelectedSerial).Value - 1 < 1 Then
                VScro(intSelectedSerial).Value = 1
            Else
                VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Value - 1
            End If
        End If
        If KeyCode = 40 Then        '下箭头
            If VScro(intSelectedSerial).Value + 1 > VScro(intSelectedSerial).Max Then
                VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Max
            Else
                VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Value + 1
            End If
        End If
        
        If KeyCode = 33 Then        'PageUp
            If VScro(intSelectedSerial).Value - VScro(intSelectedSerial).LargeChange < 1 Then
                VScro(intSelectedSerial).Value = 1
            Else
                VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Value - VScro(intSelectedSerial).LargeChange
            End If
        End If
        If KeyCode = 34 Then        'PageDown
            If VScro(intSelectedSerial).Value + VScro(intSelectedSerial).LargeChange > VScro(intSelectedSerial).Max Then
                VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Max
            Else
                VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Value + VScro(intSelectedSerial).LargeChange
            End If
        End If
        
        If KeyCode = 35 Then        'End 最后一个图
            VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Max
        End If
        
        If KeyCode = 36 Then        'Home 第一个图
            VScro(intSelectedSerial).Value = 1
        End If
    End If
    
    '处理 左右箭头，上下翻序列
    If KeyCode = 37 Then    '左箭头
        '按左箭头上一序列
        subChangeASeries 2
    End If
    If KeyCode = 39 Then    '右箭头
        '按下右箭头下一序列
        subChangeASeries 1
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub viewer_DblClick(Index As Integer)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim ls As DicomLabels
    Dim cx As Integer, cy As Integer
    Dim i As Integer
    Dim l As DicomLabel
    Dim im As DicomImage
    Dim oldScrollVisible As Boolean
    Dim CmdControl As CommandBarControl
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo err
    
    Set ls = Viewer(Index).LabelHits(lngBaseXX, lngBaseYY, False, False, True)
    subTakeOut1 ls, SelectedImage, True                                                 ''''从选中的标注中去掉系统标注用于裁减的序号小于6的标注
    If ls.Count > 0 Then     ''双击了一个标注
'        i = funMouseOverPeriod(Viewer(Index), SelectedImage, lngBaseX, lngBaseY)       ''判断鼠标是否越过句柄
'        If i = 0 Then   ''选中一个标注
        If SelectedImage.Labels.IndexOf(ls(1)) > G_INT_SYS_LABEL_COUNT Then   ''''不是双击的类似句柄的LABEL
            ''''''''''处理文字的双击修改'''''''''''''''''''''''''''''''''''
            If Mid(ls(1).Tag, 1, 3) = "TXT" Then
                Set SelectedLabelT = ls(1)
                Set SelectedLabel = Nothing
                isSelectedLabel = True
                SubChangeColor ls(1), Me ''改变显示颜色
                lblChange = SelectedLabelT.Text + " "    '''''通过lblChange的AutoSize属性自动计算text框的大小
                txtText = SelectedLabelT.Text
                oldFontSize = SelectedLabelT.FontSize
                lblChange.FontSize = lblChange.FontSize * IIf(blnLabelTextScaleFontSize, SelectedImage.ActualZoom, 1)
                txtText.FontSize = lblChange.FontSize
                cx = SelectedLabelT.left
                cy = SelectedLabelT.top
                subTextCoordinate SelectedImage, cx, cy, lblChange          '''''根据选装情况决定作坐标转换
                txtText.left = Viewer(Index).left + (cx) * Screen.TwipsPerPixelX * SelectedImage.ActualZoom - SelectedImage.ActualScrollX * Screen.TwipsPerPixelX + Viewer(Index).CellSpacing * Screen.TwipsPerPixelX + SelectedImage.OriginX * Screen.TwipsPerPixelX
                txtText.top = Viewer(Index).top + (cy) * Screen.TwipsPerPixelY * SelectedImage.ActualZoom - SelectedImage.ActualScrollY * Screen.TwipsPerPixelY + Viewer(Index).CellSpacing * Screen.TwipsPerPixelY + SelectedImage.OriginY * Screen.TwipsPerPixelY
                txtText.height = lblChange.height
                txtText.width = lblChange.width
                SelectedLabelT.Visible = False
                txtText.Visible = True
                oldTextleft = txtText.left + lblChange.width
                txtText.SetFocus
                blnTextInputM = True
                Viewer(Index).Refresh
            ElseIf ls(1).LabelType = 1 Or ls(1).LabelType = 2 Or ls(1).LabelType = 4 Then
                funcROIHistogram SelectedLabel      '对于区域类型的标注，画直方图，输入参数为需要画直方图的标注
            ElseIf left(ls(1).Tag, 3) = "VAS" Then      '血管狭窄测量
                '显示测量结果
                
                Set frmVasMeasure.lblText = SelectedLabel.TagObject
                Set frmVasMeasure.f = Me
                frmVasMeasure.Show 1, Me
            ElseIf ls(1).LabelType = 3 Or ls(1).LabelType = 4 Then
                funcDrawGreyDistribute SelectedImage, SelectedLabel '直线和多边线标注
            End If
            Exit Sub
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''多个显示转单个显示'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With Viewer(intSelectedSerial)
        '''''''''处理图片拼接的情况''''''''''''''''''''''''''''
        If blnfis Then
            Call fis.subLoadImage(SelectedImage)
        ElseIf blnPrintFilm Then    '处理胶片打印的情况
            Call AddImgToFilm(SelectedImage, Viewer(Index), ZLShowSeriesInfos(intSelectedSerial).ImageInfos(SelectedImageIndex).blnPrinted)
        ElseIf intClickImageIndex > 0 And intDblClickButton = 1 Then
            
            ''''如果是多行或多列显示则转换为单行单列显示
            If .MultiColumns > 1 Or .MultiRows > 1 Then
                .MultiColumns = 1
                .MultiRows = 1
                MSFViewer.TextMatrix(intSelectedSerial, 3) = .CurrentIndex
                .CurrentIndex = intClickImageIndex
                Set SelectedImage = .CurrentImage
                SelectedImageIndex = .CurrentIndex
                intSelectedSerial = Index
            '''''''''以前使用过多行或多列显示则进行恢复''''''''
            ElseIf MSFViewer.TextMatrix(intSelectedSerial, 5) > 1 Or MSFViewer.TextMatrix(intSelectedSerial, 6) > 1 Then
                .MultiColumns = MSFViewer.TextMatrix(intSelectedSerial, 5)
                .MultiRows = MSFViewer.TextMatrix(intSelectedSerial, 6)
                .CurrentIndex = MSFViewer.TextMatrix(intSelectedSerial, 7)
                MSFViewer.TextMatrix(intSelectedSerial, 3) = MSFViewer.TextMatrix(intSelectedSerial, 8)
                SelectedImageIndex = MSFViewer.TextMatrix(intSelectedSerial, 3)
                Set SelectedImage = .Images(IIf(SelectedImageIndex = 0, .Images.Count, SelectedImageIndex))
            End If
            
            '判断滚动条是否需要显示，如果需要则显示滚动条，设置滚动条的最大值，最小值，LarghChange等
            subDisplayScrollBar Index, Me, False
    
            '''''''''画图像外框
            subDispframe Me, Viewer(intSelectedSerial)
            
            '自动根据图像大小，判断是否显示病人四角信息,显示或者隐藏图像中的病人信息
            Call subDisplayPatientInfo(Viewer(Index))
        End If
    End With
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Viewer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim l As DicomLabel, ls As DicomLabels, ols As DicomLabels, m As Integer
    Dim i As Integer, j As Integer, sj As Single
    
    On Error GoTo err
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    lngBaseXX = x  ''供mouse DblClick使用
    lngBaseYY = y
    intClickImageIndex = Viewer(Index).imageIndex(x, y)         ''''当前点击的图像INDEX
    intDblClickButton = Button                                  ''''供DblClick使用的鼠标按
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '选择一个Viewer
    Call subSelectAViewer(Index, intClickImageIndex)
    
    If SelectedImage Is Nothing Then Exit Sub
    
    
    '''''''''''''''''''''''''''''''''''判断当前选图像是否是彩色图像''''''''''''''''''''''''''''''''''''''''
    If SelectedImage.Attributes(&H28, &H4) = "MONOCHROME2" Or SelectedImage.Attributes(&H28, &H4) = "MONOCHROME1" Then
        blnSelectedImageIfColor = False
    Else
        blnSelectedImageIfColor = True
    End If
    ''''''''''''''''处理鼠标左键,操作图像选择句柄''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Button = 1 And intClickImageIndex <> 0 Then
        Set ls = Viewer(Index).LabelHits(x, y, True, False, False)
        If ls.Count <> 0 Then
            For Each l In ls ''''处理选择句柄
                If Mid(l.Tag, 1, 1) = "B" Then
                    l.Visible = Not l.Visible
                    ZLShowSeriesInfos(Index).ImageInfos(Viewer(Index).Images(intClickImageIndex).Tag).blnSelected = IIf(l.Visible, True, False)
                    Viewer(Index).Refresh
                    Exit Sub
                End If
            Next
        End If
    End If
    '''''''''''''''''''''''''''''''''角度化完毕'''''''''''''''''''''''''''''''''''''''
    If blnAngle Then
        SelectedLabel.XOR = False
        If Int(SelectedLabel.ROILength) = 0 Then
            With SelectedImage.Labels
                .Remove .Count
                .Remove .Count
                .Remove .Count
                isSelectedLabel = False
                Set SelectedLabel = Nothing
            End With
        End If
        blnAngle = False
        Exit Sub
    End If
      
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''3D鼠标的处理
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Button = cMouseUsage("106").lngMouseKey And Shift = cMouseUsage("106").lngShift And Button_mi3dCursor Then
         sub3DCursorStart SelectedImage
         Viewer(intSelectedSerial).Refresh
        Exit Sub
    End If
    ''''''''''''[标注的选择、移动、调整大小]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Button = cMouseUsage("201").lngMouseKey And Shift = cMouseUsage("201").lngShift And intClickImageIndex > 0 And Not LabelDrawing Then
        Set ls = Viewer(Index).LabelHits(x, y, False, False, True)
        Set ols = Viewer(Index).LabelHits(x, y, False, False, True)
        subTakeOut1 ls, SelectedImage, True         '从ls中去掉系统标注中序号为1-5的裁减标注
        subTakeOut1 ols, SelectedImage, False       '从ols中去掉系统标注中序号为2-5的裁减标注
        lngBaseX = Viewer(Index).ImageXPosition(x, y)
        lngBaseY = Viewer(Index).ImageYPosition(x, y)
        If ls.Count > 0 Or ols.Count = 1 Then      ''如果选中一个标注
            i = funMouseOverPeriod(Viewer(Index), SelectedImage, x, y)   ''返回鼠标所越过的句柄编号
            If i = 0 Then   ''没有选中任何句柄，表明选中的是一个标注
                If ls.Count > 0 Then
                    If SelectedImage.Labels.IndexOf(ls(1)) >= G_INT_SYS_LABEL_MPRV And SelectedImage.Labels.IndexOf(ls(1)) <= G_INT_SYS_LABEL_MPR_POINT_O Then
                        '是“矢冠状重建”相关的标注
                        For j = 1 To ls.Count
                            If SelectedImage.Labels.IndexOf(ls(j)) > m Then m = SelectedImage.Labels.IndexOf(ls(j))
                        Next
                        Set SelectedLabel = SelectedImage.Labels(m)     'm为序号最大的标注。
                    ElseIf (SelectedImage.Labels.IndexOf(ls(1)) = G_INT_SYS_LABEL_MPR_RESULT_H) _
                        Or (SelectedImage.Labels.IndexOf(ls(1)) = G_INT_SYS_LABEL_MPR_RESULT_V) Then
                        '是“矢冠状重建”结果图中的横线和竖线
                        Set SelectedLabel = ls(1)
                    ElseIf SelectedImage.Labels.IndexOf(ls(1)) > G_INT_SYS_LABEL_COUNT Then '不是系统标注
                        Set SelectedLabel = ls(1)
                        If left(SelectedLabel.Tag, 3) = "VAS" Then      '血管狭窄测量，让SelectedLabel指向垂直线
                            If Right(SelectedLabel.Tag, 1) = "T" Then
                                Set SelectedLabel = SelectedLabel.TagObject.TagObject.TagObject
                            ElseIf Right(SelectedLabel.Tag, 1) = "1" Then
                                Set SelectedLabel = SelectedLabel.TagObject.TagObject
                            ElseIf Right(SelectedLabel.Tag, 1) = "2" Then
                                Set SelectedLabel = SelectedLabel.TagObject
                            End If
                        ElseIf left(SelectedLabel.Tag, 3) = "CTR" Then '心胸比测量，让SelectLabel指向直线
                            If Right(SelectedLabel.Tag, 1) = "T" Then
                                Set SelectedLabel = SelectedLabel.TagObject
                            End If
                        End If
                        SubChangeColor SelectedLabel, Me ''改变显示颜色
                    End If
                Else
                    Set SelectedLabel = ols(1)      ''指向序号为1的裁剪标注
                End If
                If SelectedLabel Is Nothing Then Exit Sub
                isSelectedLabel = True
                '''''''''''''''''''''''''''''''''''显示用户标注和裁剪标注的句柄
                If SelectedImage.Labels.IndexOf(SelectedLabel) > G_INT_SYS_LABEL_COUNT Or _
                   SelectedImage.Labels.IndexOf(SelectedLabel) <= 6 _
                   Then SubDispPeriod SelectedLabel, SelectedImage, Me    ''显示句柄
                blnMoveLabel = True
                If oldSelectedSerial <> intSelectedSerial Then Viewer(oldSelectedSerial).Refresh  '''刷新原来的序列
                Me.MousePointer = 5
                Viewer(Index).Refresh
                Exit Sub
            Else  ''被选中的标注是句柄，则通过句柄改变标注大小
                blnReSizeLabel = True
                intReSizeIndex = i
                Me.MousePointer = 2
'                viewer(Index).Refresh
                Exit Sub
            End If
        End If
    End If   '''''''''结束[标注的选择、移动、调整大小]''''
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''[文字输入开始]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Button = cMouseUsage("8").lngMouseKey And Shift = cMouseUsage("8").lngShift And Button_miLabeltext Then
        blnTextInput = True
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''[画标注和自适应调窗的处理]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If ((Button = cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift) _
        Or (Button = cMouseUsage("105").lngMouseKey And Shift = cMouseUsage("105").lngShift And Button_miAutoWidthLevel) _
        Or (Button = cMouseUsage("201").lngMouseKey And Shift = cMouseUsage("201").lngShift And Button_miFrameSelectImage)) _
       And intClickImageIndex <> 0 And Not (blnTextInput Or blnTextInputM) Then
       
        If funIsLabelMouse(Me, Button, Shift) Then  ''''判断标注鼠标标志是否按下,并处理各类标注
            '''''''''''''''''''''''''''''''[各类标注处理]'''''''''''''''''''''''''''''''''''''''''
            If blnTextInput Or blnTextInputM Then
                If txtText.Visible Then     '''''如果正在输入文字,点击鼠标的情况
                    txtText_KeyPress 13
                    Exit Sub
                End If
            End If
            ''''''''''''''''在图像中增加两个标注，一个是形状标注一个是相连的文字标注'''''''''
            SubNoDispPeriod SelectedImage, Me   '为指定图像隐藏标注选择句柄
            SelectedImage.Labels.Add GetNewLabel(intSelectLabelStyle, Viewer(Index).ImageXPosition(x, y), Viewer(Index).ImageYPosition(x, y), 0, 0)
            Set SelectedLabel = SelectedImage.Labels(SelectedImage.Labels.Count)
            SelectedImage.Labels.Add GetNewLabel(doLabelText, SelectedLabel.left, SelectedLabel.top, 0, 0)
            ''''''''''''''''''''''''''自动调窗'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Button = cMouseUsage("105").lngMouseKey And Shift = cMouseUsage("105").lngShift And Button_miAutoWidthLevel Then
                SelectedLabel.LineStyle = 2
                SelectedLabel.XOR = False
                blnAutoWL = True  ''自适应调窗开始
            End If
            '''''''''''''''''''框选图象'''''''''''''''''''''''''''''
            If Button = cMouseUsage("201").lngMouseKey And Shift = cMouseUsage("201").lngShift And Button_miFrameSelectImage Then
                SelectedLabel.LineStyle = 2
                SelectedLabel.XOR = False
                blnFrameSelectImage = True      ''框选图象开始
            End If
            
            ''''''''''''''''''''设置SelectedLabelT指向标注文字'''''''''''''''''''''''''
            Set SelectedLabelT = SelectedImage.Labels(SelectedImage.Labels.Count)
            SelectedLabelT.AutoSize = True
            SelectedLabelT.Margin = 0
            ''''''''''''''''''''''''''设置开始进行测量标注的操作'''''''''''''''''''''''''''''''''''''
            If Button_miLabelAngle And Button = cMouseUsage("7").lngMouseKey _
                And Shift = cMouseUsage("7").lngShift Then
                
                blnAngle = True  ''角度开始
            ElseIf Button_miLabelVasMeasure And Button = cMouseUsage("1").lngMouseKey _
                And Shift = cMouseUsage("1").lngShift And intVasMeasure = 0 Then
                
                intVasMeasure = 1      '血管狭窄测量开始
            ElseIf Button_miLabelCadiothoracicRatio And Button = cMouseUsage("1").lngMouseKey _
                And Shift = cMouseUsage("1").lngShift And intCadioThoracicRatio = 0 Then
                
                intCadioThoracicRatio = 1   '心胸比测量开始
            End If
            
'            If Button_miAutoWidthLevel And Button = cMouseUsage("105").lngMouseKey And Shift = cMouseUsage("105").lngShift Then
'                blnAutoWL = True  ''自适应调窗开始
'            End If
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If blnAngle Then          ''''''''如果是角度的处理
                SelectedLabel.TagObject = SelectedLabelT
                SelectedLabel.Tag = "JD1"
                SelectedLabelT.Tag = "JDT"
                SelectedLabelT.left = Viewer(Index).ImageXPosition(x, y)
                SelectedLabelT.top = Viewer(Index).ImageYPosition(x, y)
                SelectedLabelT.AnchorX = Viewer(Index).ImageXPosition(x, y)
                SelectedLabelT.AnchorY = Viewer(Index).ImageYPosition(x, y)
                SelectedLabelT.ShowAnchor = True
                SelectedLabelT.AnchorImageTied = True
                SelectedLabelT.LineStyle = 2        ''''锚点线形
            ElseIf Button_miLabelVasMeasure And (intVasMeasure = 1 Or intVasMeasure = 2) Then '处理血管狭窄测量
                If intVasMeasure = 2 Then       '将正常血管部分和狭窄血管部分的TagObject相连
                    SelectedImage.Labels(SelectedImage.Labels.Count - 2).TagObject = SelectedLabel
                End If
                SelectedLabel.TagObject = SelectedLabelT
                SelectedLabel.Tag = "VAS" & intVasMeasure & "L"
                SelectedLabelT.Tag = "VAS" & intVasMeasure & "T"
                SelectedLabelT.left = SelectedLabel.left + intTextoOffX
                SelectedLabelT.top = SelectedLabel.top + intTextoOffY
                SelectedLabelT.AnchorX = SelectedLabel.left
                SelectedLabelT.AnchorY = SelectedLabel.top
                SelectedLabelT.ShowAnchor = True
                SelectedLabelT.AnchorImageTied = True
                SelectedLabelT.LineStyle = 2        ''''锚点线形
                '再增加两个血管壁标注
                SelectedImage.Labels.Add GetNewLabel(intSelectLabelStyle, SelectedLabel.left, SelectedLabel.top, 0, 0)
                Set l = SelectedImage.Labels(SelectedImage.Labels.Count)
                l.Tag = "VAS" & intVasMeasure & "E1"
                l.XOR = False
                SelectedLabelT.TagObject = l
                SelectedImage.Labels.Add GetNewLabel(intSelectLabelStyle, SelectedLabel.left, SelectedLabel.top, 0, 0)
                l.TagObject = SelectedImage.Labels(SelectedImage.Labels.Count)
                Set l = SelectedImage.Labels(SelectedImage.Labels.Count)
                l.Tag = "VAS" & intVasMeasure & "E2"
                l.XOR = False
                l.TagObject = SelectedImage.Labels(SelectedImage.Labels.Count - (intVasMeasure * 4 - 1)) '将TagObject连成封闭环型
            ElseIf Button_miLabelCadiothoracicRatio And (intCadioThoracicRatio = 1 Or intCadioThoracicRatio = 2) Then  '处理心胸比测量
                If intCadioThoracicRatio = 1 Then   '画心脏部分
                    
                ElseIf intCadioThoracicRatio = 2 Then   '画胸廓部分
                    '将胸廓部分和心脏部分的TagObject相连
                    SelectedImage.Labels(SelectedImage.Labels.Count - 2).TagObject = SelectedLabel
                End If
                SelectedLabel.TagObject = SelectedLabelT
                SelectedLabel.Tag = "CTR" & intCadioThoracicRatio & "L"
                SelectedLabelT.Tag = "CTR" & intCadioThoracicRatio & "T"
                SelectedLabelT.left = SelectedLabel.left + intTextoOffX
                SelectedLabelT.top = SelectedLabel.top + intTextoOffY
                SelectedLabelT.AnchorX = SelectedLabel.left
                SelectedLabelT.AnchorY = SelectedLabel.top
                SelectedLabelT.ShowAnchor = True
                SelectedLabelT.AnchorImageTied = True
                SelectedLabelT.LineStyle = 2        ''''锚点线形
                '将TagObject连成封闭环型
                SelectedLabelT.TagObject = SelectedImage.Labels(SelectedImage.Labels.Count - (intCadioThoracicRatio * 2 - 1))
            
            Else   '''''不是角度的处理
               SelectedLabel.TagObject = SelectedLabelT
               SelectedLabelT.TagObject = SelectedLabel
               SelectedLabelT.Tag = "RIO"
            End If
            ''''''''''''''''''''''''为多边形和多边线增加一个点,避免计算长度出错.''''''''''''''''''''''''''''''''''''''''''''''
            If Button_miLabelPolygon Or Button_miLabelPolygon Then
                SelectedLabel.AddPoint Viewer(Index).ImageXPosition(x, y), Viewer(Index).ImageYPosition(x, y)
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            LabelDrawing = True                     '设置开始画标注的标记为真
            SubChangeColor SelectedLabel, Me        '改变选中LABEL的颜色
            Me.MousePointer = 2
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''判断当前图像是否进入了裁剪状态''''''''''''''''''''''''''''''''''''''''
    If SelectedImage.Labels(1).Visible = True Then
        Button_miCutOut = True
    Else
        Button_miCutOut = False
    End If
    
    blDicomDown = True                      '按下鼠标
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Viewer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim A As Variant
    Dim lngValue As Long, sj As Double
    Dim ww As Long, wl As Long
    Dim tl As Long, tt As Long, tw As Long, th As Long
    Dim dblZoom As Double
    Dim i As Long
    
    If SelectedImage Is Nothing Then Exit Sub
    
    On Error GoTo err
    '''''''''''''''''''''''''''''''''''[在状态栏显示X/Y坐标和图像点值]'''''''''''''''''''''''''''''''''''''
    intClickImageIndex = Viewer(Index).imageIndex(x, y)
    
    '增加条件判断是否按下鼠标,按下不执行计算灰度值提高效率
    If (intClickImageIndex <> 0 And Not blDicomDown) Then
        If Viewer(Index).Images(intClickImageIndex).FrameCount > 1 And Button_miMouseShowValue = False Then
            '多帧图像，而且不随鼠标显示像素值，则不进行计算
            Me.Viewer(Index).ToolTipText = ""
        Else
            '存储像素值的调整解决和斜率
            If strInstanceUID <> Viewer(Index).Images(intClickImageIndex).InstanceUID Then
                strInstanceUID = Viewer(Index).Images(intClickImageIndex).InstanceUID
                intIntercept = 0
                intSlope = 1
                If Not IsNull(Viewer(Index).Images(intClickImageIndex).Attributes(&H28, &H1052).Value) Then
                    intIntercept = Viewer(Index).Images(intClickImageIndex).Attributes(&H28, &H1052)
                End If
                If Not IsNull(Viewer(Index).Images(intClickImageIndex).Attributes(&H28, &H1053)) Then
                    intSlope = Viewer(Index).Images(intClickImageIndex).Attributes(&H28, &H1053)
                End If
            End If
            A = Viewer(Index).Images(intClickImageIndex).Pixels
            If Viewer(Index).ImageXPosition(x, y) > 0 And Viewer(Index).ImageYPosition(x, y) > 0 And Viewer(Index).ImageXPosition(x, y) < Viewer(Index).Images(intClickImageIndex).sizeX And Viewer(Index).ImageYPosition(x, y) < Viewer(Index).Images(intClickImageIndex).sizeY Then
                If Not IsNull(A) Then lngValue = A(Viewer(Index).ImageXPosition(x, y), Viewer(Index).ImageYPosition(x, y), 1)
                sbStatusBar.Panels(4).Text = "行:" & Viewer(Index).ImageXPosition(x, y) & "   列:" & Viewer(Index).ImageYPosition(x, y) & "   值:" & lngValue * intSlope + intIntercept
            End If
            If Button_miMouseShowValue = True Then
                Me.Viewer(Index).ToolTipText = Mid(sbStatusBar.Panels(4).Text, InStr(1, sbStatusBar.Panels(4).Text, "值:") + 2)
            Else
                Me.Viewer(Index).ToolTipText = ""
            End If
        End If
    Else
        Me.Viewer(Index).ToolTipText = ""
    End If
    ''''''在鼠标在未击下任何键的情况下，改变越过当前图像中标注选择句柄的鼠标形状
    If Button = 0 And (Not SelectedLabel Is Nothing And Not blnReSizeLabel Or Button_miCutOut) Then
        If intClickImageIndex = Viewer(Index).Images.IndexOf(SelectedImage) And (isSelectedLabel Or Button_miCutOut) Then
            i = funMouseOverPeriod(Viewer(Index), SelectedImage, x, y)  '返回鼠标所越过的句柄编号
            If i <> 0 Then  '鼠标经过的是句柄之一，改变鼠标形状
                If Mid(SelectedLabel.Tag, 1, 2) = "JD" Then '处理角度的鼠标形状
                      Me.MousePointer = 2
                Else            '处理非角度标注的鼠标形状
                    If (i = 11 Or i = 15) _
                       And (SelectedLabel.width > 0 And SelectedLabel.height > 0 _
                            Or SelectedLabel.width < 0 And SelectedLabel.height < 0) _
                       Or (i = 13 Or i = 17) _
                       And Not (SelectedLabel.width > 0 And SelectedLabel.height > 0 _
                                Or SelectedLabel.width < 0 And SelectedLabel.height < 0) Then
                                
                        Me.MousePointer = 8
                    ElseIf (i = 11 Or i = 15) _
                           And Not (SelectedLabel.width > 0 And SelectedLabel.height > 0 _
                                    Or SelectedLabel.width < 0 And SelectedLabel.height < 0) _
                           Or (i = 13 Or i = 17) _
                           And (SelectedLabel.width > 0 And SelectedLabel.height > 0 _
                                Or SelectedLabel.width < 0 And SelectedLabel.height < 0) Then
                                
                        Me.MousePointer = 6
                    ElseIf i = 12 Or i = 16 Then
                        Me.MousePointer = 9
                    ElseIf i = 14 Or i = 18 Then
                        Me.MousePointer = 7
                    End If
                    If SelectedImage.FlipState = doFlipHorizontal Or SelectedImage.FlipState = doFlipVertical Then
                        If (i = 11 Or i = 13 Or i = 17 Or i = 15) Then Me.MousePointer = IIf(Me.MousePointer = 8, 6, 8)
                    End If
                    If SelectedImage.RotateState = doRotateLeft Or SelectedImage.RotateState = doRotateRight Then
                        If (i = 11 Or i = 13 Or i = 17 Or i = 15) Then Me.MousePointer = IIf(Me.MousePointer = 8, 6, 8)
                        If (i = 12 Or i = 14 Or i = 16 Or i = 18) Then Me.MousePointer = IIf(Me.MousePointer = 9, 7, 9)
                    End If
                End If  '处理鼠标形状结束
            Else            '鼠标经过的不是句柄，将鼠标形状还原成0
                Me.MousePointer = 0
            End If
        End If
    End If          '结束“在鼠标在未击下任何键的情况下，改变越过当前图像中标注选择句柄的鼠标形状”
    '''''''''''''''''''''''''移动标注''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If blnMoveLabel And Not SelectedLabel Is Nothing Then
        subaCorrectCursor Viewer(intSelectedSerial), SelectedImage, x, y    ''''鼠标移动如果超出图像范围则修正其鼠标位置
        '''''''''''''''''''''''''''''''[移动标注]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '先判断是否在做MPR重建,判断当前选中的标注是否是矢冠状重建的5个控制点之一
        If SelectedImage.Labels.IndexOf(SelectedLabel) >= G_INT_SYS_LABEL_MPR_POINT_V1 _
            And SelectedImage.Labels.IndexOf(SelectedLabel) <= G_INT_SYS_LABEL_MPR_POINT_O Then    '是矢冠状重建
            If blnInMPR = False Then
                MsgBox "MPR重建结果已删除，可能是重建过程中出现错误。" & vbCrLf & vbCrLf & "请重新进行重建。", vbInformation, gstrSysName
                blnMoveLabel = False
                Exit Sub
            End If
        End If
        '移动标注的所有操作，包括移动MPR线并显示重建结果图
        subMoveLable SelectedLabel, Viewer(Index).ImageXPosition(x, y) - lngBaseX, Viewer(Index).ImageYPosition(x, y) - lngBaseY, Me, x, y, lngBaseX, lngBaseY
        If SelectedImage.Labels.IndexOf(SelectedLabel) > G_INT_SYS_LABEL_COUNT _
           Or SelectedImage.Labels.IndexOf(SelectedLabel) <= 6 Then   ''''显示用户标注和裁剪标注的选择句柄
            SubDispPeriod SelectedLabel, SelectedImage, Me
        End If
        lngBaseX = Viewer(Index).ImageXPosition(x, y)
        lngBaseY = Viewer(Index).ImageYPosition(x, y)
        Viewer(Index).Refresh
        Exit Sub
    End If
    ''''''''''''''''''''''''''改变标注大小'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If blnReSizeLabel And Not SelectedLabel Is Nothing Then
        subaCorrectCursor Viewer(intSelectedSerial), SelectedImage, x, y                ''''鼠标移动如果超出图像范围则修正其鼠标位置
        '''''''''''''''''''''''''''''''[改变标注大小]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If left(SelectedLabel.Tag, 3) = "VAS" Then      '血管狭窄测量，让SelectedLabel指向垂直线
            If Right(SelectedLabel.Tag, 1) = "T" Then
                Set SelectedLabel = SelectedLabel.TagObject.TagObject.TagObject
            ElseIf Right(SelectedLabel.Tag, 1) = "1" Then
                Set SelectedLabel = SelectedLabel.TagObject.TagObject
            ElseIf Right(SelectedLabel.Tag, 1) = "2" Then
                Set SelectedLabel = SelectedLabel.TagObject
            End If
        ElseIf left(SelectedLabel.Tag, 3) = "CTR" Then  '心胸比测量，让SelectedLabel指向直线
            If Right(SelectedLabel.Tag, 1) = "T" Then
                Set SelectedLabel = SelectedLabel.TagObject
            End If
        End If
        
        subChangeLableSize SelectedLabel, Viewer(Index).ImageXPosition(x, y) - lngBaseX, Viewer(Index).ImageYPosition(x, y) - lngBaseY, intReSizeIndex, Me
        SubDispPeriod SelectedLabel, SelectedImage, Me
        lngBaseX = Viewer(Index).ImageXPosition(x, y)
        lngBaseY = Viewer(Index).ImageYPosition(x, y)
        Viewer(Index).Refresh
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''穿梭''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Button = cMouseUsage("101").lngMouseKey And Shift = cMouseUsage("101").lngShift And Button_miStack Then
        If Abs(y - lngBaseYY) >= lngStackStep Then          ''''穿梭步长控制,计算Y方向的位移作为穿梭的步长
            '调用函数处理穿梭中的鼠标移动
            Call subStackMouseMove(y - lngBaseYY)
            lngBaseYY = y
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''[调窗]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (Button = cMouseUsage("102").lngMouseKey And Shift = cMouseUsage("102").lngShift And Button_miWidthLevel And (intClickImageIndex <> 0 Or blnMouseStart)) _
        Or (Button = 4 And intMouseWheelDrag = 2) Then
        '鼠标左右键，通过按钮可以进行调窗；鼠标中间键拖拽通过设置可以进行调窗
        Dim DicomAttr As DicomAttribute
        Dim DicomDate As DicomDataSets
        Dim VarTmp As Variant
        If Abs(y - lngBaseYY) >= lngWidthLevelStep / 5 Or Abs(x - lngBaseXX) >= lngWidthLevelStep / 5 Then  ''''调窗步长控制
            If Not blnMouseStart Then
                Me.MouseIcon = ImageListMouse.ListImages("WindowWL").Picture
                Me.MousePointer = 99
                blnMouseStart = True
            End If
            '处理特殊的图像VOILUT=0时才能进行调窗
            If SelectedImage.VOILUT = 1 Then
                Set DicomAttr = SelectedImage.Attributes(&H28, &H3010)
                If VarType(DicomAttr) = vbObject Then
                    Set DicomDate = DicomAttr.Value
                    'mindray迈瑞的DR图像，DicomDate(1).Attributes(&H28, &H3002).Value为空
                    If IsNull(DicomDate(1).Attributes(&H28, &H3002).Value) Then
                        subDispWWWL SelectedImage
                    Else
                        VarTmp = DicomDate(1).Attributes(&H28, &H3002).Value
                        '针对吉林省人民医院的DR图像，他们的VarTmp(2)=0，就不采用这种方式来修改窗宽床位了，直接修改VOILUT=1就可以了。
                        If VarTmp(2) = 0 Then
                            subDispWWWL SelectedImage
                        Else
                            SelectedImage.width = VarTmp(1)
                            SelectedImage.Level = VarTmp(2) + (VarTmp(2) / 2)
                        End If
                    End If
                Else
                    subDispWWWL SelectedImage
                End If
                SelectedImage.VOILUT = 0
            End If
            '调窗的调节单位是1
            SelectedImage.width = SelectedImage.width + (x - lngBaseXX) * lngWidthLevelStep / 5
            SelectedImage.Level = SelectedImage.Level + (y - lngBaseYY) * lngWidthLevelStep / 5
            SelectedImage.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & SelectedImage.width & "-L:" & SelectedImage.Level
'            viewer(intSelectedSerial).Refresh
            lngBaseXX = x
            lngBaseYY = y
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''[缩放]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (Button = cMouseUsage("104").lngMouseKey And Shift = cMouseUsage("104").lngShift And Button_miZoom And (intClickImageIndex <> 0 Or blnMouseStart)) _
        Or (Button = 4 And intMouseWheelDrag = 1) Then
        '鼠标左右键，通过按钮可以进行缩放；鼠标中间键拖拽通过设置可以进行漫游
        If Abs(y - lngBaseYY) >= lngZoomStep / 5 Then                                                                 ''''缩放步长控制
            If Not blnMouseStart Then
                Me.MouseIcon = ImageListMouse.ListImages("Zoom").Picture
                Me.MousePointer = 99
                blnMouseStart = True
            End If
            '缩放的调节单位是0.001倍
            dblZoom = SelectedImage.ActualZoom * (1 + (lngBaseYY - y) * lngZoomStep / 5 * 0.001)
            If dblZoom < 0.01 Then dblZoom = 0.01
            If dblZoom > 64 Then dblZoom = 64
            subCenterZoom SelectedImage, Viewer(intSelectedSerial), dblZoom
            If SelectedImage.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then '更新标尺单位
                UpdateRuler SelectedImage, True
            End If
            Viewer(intSelectedSerial).Refresh
            lngBaseXX = x
            lngBaseYY = y
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''[漫游]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (Button = cMouseUsage("103").lngMouseKey And Shift = cMouseUsage("103").lngShift And Button_miCruise And (intClickImageIndex <> 0 Or blnMouseStart)) _
        Or (Button = 4 And intMouseWheelDrag = 0) Then
        '鼠标左右键，通过按钮可以进行漫游；鼠标中间键拖拽通过设置可以进行漫游
        If Abs(y - lngBaseYY) >= lngCruiseStep / 5 Or Abs(x - lngBaseXX) >= lngCruiseStep / 5 Then
            If Not blnMouseStart Then
                Me.MousePointer = 15
                blnMouseStart = True
                subCenterZoom SelectedImage, Viewer(intSelectedSerial), SelectedImage.ActualZoom
            End If
            SelectedImage.ScrollX = SelectedImage.ScrollX - (x - lngBaseXX) * lngCruiseStep / 5
            SelectedImage.ScrollY = SelectedImage.ScrollY - (y - lngBaseYY) * lngCruiseStep / 5
'            viewer(intSelectedSerial).Refresh
            lngBaseXX = x
            lngBaseYY = y
        End If
    End If
    ''''''''''''''''''''''3D 鼠标'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Button = cMouseUsage("106").lngMouseKey And Shift = cMouseUsage("106").lngShift And Button_mi3dCursor Then
        '三维鼠标状态下，移动鼠标的时候，翻动对应的图像
        If Viewer(Index).imageIndex(x, y) = 0 Then Exit Sub
        
        '调用函数处理鼠标移动的操作
        Call sub3DCursorMouseMove(x, y, Viewer(Index))
        Exit Sub
    End If      '鼠标键位和状态栏状态满足3D鼠标的要求
    
    ''''''''''''''''''[画角度的第二根线,Button=0表示没有鼠标键被按下]'''''''''''''''''
    If Button = 0 And intClickImageIndex <> 0 And blnAngle And Not (blnTextInput Or blnTextInputM) _
        And Not SelectedLabel Is Nothing Then
        With SelectedLabel
            .width = Viewer(Index).ImageXPosition(x, y) - .left
            .height = Viewer(Index).ImageYPosition(x, y) - .top
            SelectedLabelT.Text = funROIResultString(SelectedLabel, SelectedImage)  ' "Angle=" & Int(GetAngle(.left, .top, .left + .width, .top + .height, .TagObject.left, .TagObject.top) * 100) / 100  ''注意处理字体属性
            Viewer(Index).Refresh
        End With
    End If
    
    '''''''''''''''''''''''''''''''[画标注]'''''''''''''''''''''''''''''''''''''''''
    If LabelDrawing And Not (blnTextInput Or blnTextInputM) And Not SelectedLabel Is Nothing Then
        subaCorrectCursor Viewer(intSelectedSerial), SelectedImage, x, y    '鼠标移动如果超出图像范围则修正其鼠标位置
        
        '如果是框选图像，则限制标注为正方形
        If Button_miFrameSelectImage = True And Button = cMouseUsage("201").lngMouseKey And blnSquareFrame = True Then
            If Abs(Viewer(Index).ImageXPosition(x, y) - SelectedLabel.left) < Abs(Viewer(Index).ImageYPosition(x, y) - SelectedLabel.top) Then
                SelectedLabel.width = Viewer(Index).ImageXPosition(x, y) - SelectedLabel.left
            Else
                SelectedLabel.width = Viewer(Index).ImageYPosition(x, y) - SelectedLabel.top
            End If
            SelectedLabel.height = SelectedLabel.width
        Else
            SelectedLabel.width = Viewer(Index).ImageXPosition(x, y) - SelectedLabel.left
            SelectedLabel.height = Viewer(Index).ImageYPosition(x, y) - SelectedLabel.top
        End If
        ''''''''''''''''''''''''''''''[多边形和多边线的处理]''''''''''''''''''''''
        If SelectedLabel.LabelType = 4 Or SelectedLabel.LabelType = 5 Then
            SelectedLabel.AddPoint Viewer(Index).ImageXPosition(x, y), Viewer(Index).ImageYPosition(x, y)
        End If
        
        ''血管狭窄测量标注,计算并显示两个血管壁的位置
        If Button_miLabelVasMeasure And (intVasMeasure = 1 Or intVasMeasure = 2) _
           And Button = cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift Then
            
            If funDrawVas(SelectedLabel, SelectedImage, intVasMeasure) = False Then
                intVasMeasure = 0
            End If
            SelectedLabelT.AnchorX = SelectedLabel.left + SelectedLabel.width / 2
            SelectedLabelT.AnchorY = SelectedLabel.top + SelectedLabel.height / 2
            SelectedImage.Refresh False
            Exit Sub
        End If
        '''''''''''''''''''''''''''''''[显示文字]''''''''''''''''''''''''''''''
        With SelectedLabelT
            If Not blnAngle Then
                '''''''''' '''箭头'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If SelectedLabel.LabelType = doLabelArrow Then
                    .Text = " "
                    .left = SelectedLabel.left + SelectedLabel.width
                    .top = SelectedLabel.top + SelectedLabel.height
                    .AnchorX = SelectedLabel.left + SelectedLabel.width
                    .AnchorY = SelectedLabel.top + SelectedLabel.height
                '''''''''''''''''''自动调窗'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf Button = cMouseUsage("105").lngMouseKey And Shift = cMouseUsage("105").lngShift And Button_miAutoWidthLevel Then
                    If SelectedLabel.height <> 0 And SelectedLabel.width <> 0 Then
                        ''''''''''''''处理矩形框宽度和高度为负数的情况''''''''''''''''''''''''''''''''''''''
                        If SelectedLabel.width < 0 Then
                            tl = SelectedLabel.left + SelectedLabel.width
                            tw = -SelectedLabel.width
                        Else
                            tl = SelectedLabel.left
                            tw = SelectedLabel.width
                        End If
                        
                        If SelectedLabel.height < 0 Then
                            tt = SelectedLabel.top + SelectedLabel.height
                            th = -SelectedLabel.height
                        Else
                            tt = SelectedLabel.top
                            th = SelectedLabel.height
                        End If
                        ''''''''''''''计算区域的窗宽窗位并显示''''''''''''''''''''''''''''''''''''''
                        funAutoWinWL SelectedImage, tl, tt, tw, th, ww, wl
                        .Text = "窗宽: " & ww & vbCrLf & "窗位:" & wl
                        SelectedLabel.Tag = ww
                        .Tag = wl
                    End If
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    .left = SelectedLabel.left + SelectedLabel.width + intTextoOffX
                    .top = SelectedLabel.top + SelectedLabel.height + intTextoOffY
                    .AnchorX = SelectedLabel.left + SelectedLabel.width / 2
                    .AnchorY = SelectedLabel.top + SelectedLabel.height / 2
                Else          '不是箭头、自动调窗的标注的处理
                    If SelectedLabel.LabelType = doLabelEllipse Or SelectedLabel.LabelType = doLabelPolygon Or SelectedLabel.LabelType = doLabelRectangle Then
'                        .Text = funROIResultString(SelectedLabel)      ''''区域类型标注的文字处理
                    Else
                        .Text = Int(SelectedLabel.ROILength) & SelectedLabel.ROIDistanceUnits '      ''长度类型的标注文字处理
                    End If
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    .left = SelectedLabel.left + SelectedLabel.width + intTextoOffX
                    If .left > SelectedImage.sizeX * 0.9 Or .left <= SelectedImage.sizeX * 0.1 Then .left = SelectedImage.sizeX / 2
                    .top = SelectedLabel.top + SelectedLabel.height + intTextoOffY
                    If .top > SelectedImage.sizeY * 0.9 Or .top <= SelectedImage.sizeY * 0.1 Then .top = SelectedImage.sizeY / 2
                    
                    .AnchorX = SelectedLabel.left + SelectedLabel.width / 2
                    .AnchorY = SelectedLabel.top + SelectedLabel.height / 2
                End If
                
                .ShowAnchor = True
                .AnchorImageTied = True
                .LineStyle = 2
            End If    'end of “If Not blnAngle Then”
            ''''''''''''''''''''''''''''''[多边形和多边线的处理]''''''''''''''''''''''
            If SelectedLabel.LabelType = 4 Or SelectedLabel.LabelType = 5 Then
                .AnchorX = Viewer(Index).ImageXPosition(x, y)
                .AnchorY = Viewer(Index).ImageYPosition(x, y)
            End If
        End With
        Viewer(Index).Refresh
    End If      'end of "If LabelDrawing And Not (blnTextInput Or blnTextInputM)"
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Viewer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim l As DicomLabel
    Dim i As Integer, v As DicomViewer, j As Integer, ii As Integer, k As Integer
    Dim xx As Integer, Yy As Integer
    
    On Error GoTo err
    
    If SelectedImage Is Nothing Then Exit Sub
    
    i = Viewer(Index).imageIndex(x, y)
    '''''''''''''''''''''''''''''''''''''''''''[穿梭j结束]''''''''''''''''''''''''''''''''''''''''''''
    With Viewer(intSelectedSerial)
        If Button = cMouseUsage("101").lngMouseKey And Shift = cMouseUsage("101").lngShift And Button_miStack And blnStackStart Then  ''''
            If blnStackisFrame Then    ''''多帧图像处理
                j = SelectedImage.Frame - intStackOffset
                SelectedImage.Frame = intStackCurrentlyImage
            Else
                '调用函数结束穿梭
                subStackEnd Viewer(intSelectedSerial), Me
                j = intStackIndex - intStackOffset
            End If
            
            If j > ZLShowSeriesInfos(Index).ImageInfos.Count - Viewer(Index).MultiColumns * Viewer(Index).MultiRows + 1 Then
                j = ZLShowSeriesInfos(Index).ImageInfos.Count - Viewer(Index).MultiColumns * Viewer(Index).MultiRows + 1
            End If
            
            If j < 1 Then j = 1
            
            If VScro(intSelectedSerial).Visible Then VScro(intSelectedSerial).Value = j
            
            blnStackStart = False
        End If
    End With
    ''''''''''''''''''''''''''''''''''''''''''''''''''[3D鼠标结束]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Button = cMouseUsage("106").lngMouseKey And Shift = cMouseUsage("106").lngShift And Button_mi3dCursor And i > 0 Then
        sub3DCursorEnd SelectedImage
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''[文字输入]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If blnTextInput Then
        If txtText.Visible Then     '''''如果正在输入文字,点击鼠标的情况
            txtText_KeyPress 13
        Else
            ''''''''''''''''''''''''''''''''''''''''''[输入文字]'''''''''''''''''''''''''''''''''''''''''''''
            With Viewer(Index)
                ''''''''''''''''''构造文字标签''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                SelectedImage.Labels.Add GetNewLabel(0, .ImageXPosition(x, y), .ImageYPosition(x, y), 0, 0)
                Set SelectedLabelT = SelectedImage.Labels(SelectedImage.Labels.Count)
                SelectedLabelT.Tag = "TXT"
                SelectedLabelT.AutoSize = True
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                lblChange = "  "
                txtText = ""
                oldFontSize = SelectedLabelT.FontSize
                SelectedLabelT.Visible = False
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                lblChange.FontSize = lblChange.FontSize * IIf(blnLabelTextScaleFontSize, SelectedImage.ActualZoom, 1)
                txtText.FontSize = lblChange.FontSize
                xx = Viewer(Index).ImageXPosition(x, y)
                Yy = Viewer(Index).ImageYPosition(x, y)
                subTextCoordinate SelectedImage, xx, Yy, lblChange   '''''计算交换坐标
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                xx = (xx * SelectedImage.ActualZoom - SelectedImage.ActualScrollX) * Screen.TwipsPerPixelX + .left + Viewer(Index).width / Viewer(Index).MultiColumns * (FunImageIsX(i, Viewer(Index)) - 1)
                Yy = (Yy * SelectedImage.ActualZoom - SelectedImage.ActualScrollY) * Screen.TwipsPerPixelY + .top + Viewer(Index).height / Viewer(Index).MultiRows * (FunImageIsY(i, Viewer(Index)) - 1)
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                txtText.Move xx, Yy, lblChange.width, lblChange.height
                txtText.Visible = True
                oldTextleft = xx + lblChange.width
                txtText.SetFocus
            End With
        End If
    End If
    '''''''''''''''''''''''''''''''''''''[标注和自适应调窗的处理]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If LabelDrawing _
       And ((Button = cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift) _
            Or (Button = cMouseUsage("105").lngMouseKey And Shift = cMouseUsage("105").lngShift) _
            Or (Button = cMouseUsage("201").lngMouseKey And Shift = cMouseUsage("201").lngShift And Button_miFrameSelectImage = True)) Then
        LabelDrawing = False
        SelectedLabel.XOR = False
        '''''''''''''自适应调窗'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If blnAutoWL Then  ''''''''''''''''''''''自适应窗宽窗位
            If Int(SelectedLabel.ROILength) > 2 Then
                SelectedImage.VOILUT = 0
                SelectedImage.width = Val(SelectedLabel.Tag)                ''''用于标注调窗范围的矩形和对应文字记录了窗宽窗位
                SelectedImage.Level = Val(SelectedLabelT.Tag)
                SelectedImage.Refresh False
            End If
            SelectedImage.Labels.Remove SelectedImage.Labels.Count
            SelectedImage.Labels.Remove SelectedImage.Labels.Count
            isSelectedLabel = False
            Set SelectedLabel = Nothing
        ElseIf blnFrameSelectImage Then  ''''''''''''''''''框选图象
            '显示图像保存菜单
            If SelectedLabel.width <> 0 And SelectedLabel.height <> 0 Then
                ShowFrameSelectImagePopup Me
            End If
            '删除框选用的临时标注
            SelectedImage.Labels.Remove SelectedImage.Labels.Count
            SelectedImage.Labels.Remove SelectedImage.Labels.Count
            isSelectedLabel = False
            Set SelectedLabel = Nothing
        ElseIf SelectedLabel.LabelType = 4 And UBound(SelectedLabel.Points) = 0 Then '删除长度为0的多边线标注
            '对多边线做单独处理，因为直接调用长度为0的多边线的ROILength会出现即时错误
            SelectedImage.Labels.Remove SelectedImage.Labels.Count
            SelectedImage.Labels.Remove SelectedImage.Labels.Count
            isSelectedLabel = False
            Set SelectedLabel = Nothing
            blnAngle = False    '为何清除角度标志？？“疑问”
        ElseIf Int(SelectedLabel.ROILength) = 0 Or (SelectedLabel.width = 0 And SelectedLabel.height = 0) Then
            ''''删除长度为0的标注,当图像被缩小后，标注的宽高都为0时，ROILength可能不为0。所以两个都要判断
            If left(SelectedLabel.Tag, 3) = "VAS" Then
                '如果是狭窄血管，则将正常血管的链重新连上
                If Mid(SelectedLabel.Tag, 4, 1) = "2" Then
                    SelectedImage.Labels(SelectedImage.Labels.Count - 4).TagObject = SelectedImage.Labels(SelectedImage.Labels.Count).TagObject
                End If
                '删除最后两个长度为0的标注
                SelectedImage.Labels.Remove SelectedImage.Labels.Count
                SelectedImage.Labels.Remove SelectedImage.Labels.Count
            End If
                SelectedImage.Labels.Remove SelectedImage.Labels.Count
                SelectedImage.Labels.Remove SelectedImage.Labels.Count
                isSelectedLabel = False
                Set SelectedLabel = Nothing
                blnAngle = False
                
        ElseIf Button = cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift And blnAngle Then
                '处理角度的第二根线
                SelectedImage.Labels.Add GetNewLabel(3, SelectedLabel.left + SelectedLabel.width, SelectedLabel.top + SelectedLabel.height, 0, 0)
                Set l = SelectedImage.Labels(SelectedImage.Labels.Count)
                l.Tag = "JD2"
                l.ForeColour = lngLabelSelectedColor
                SelectedLabelT.TagObject = l
                l.TagObject = SelectedLabel
                Set SelectedLabel = SelectedImage.Labels(SelectedImage.Labels.Count)
        ElseIf SelectedLabel.LabelType = doLabelArrow Then   '箭头的处理[此段的处理和直接输入文字的处理比较类似,可以考虑合并]
            With Viewer(Index)
                SelectedLabelT.Tag = "TXTA"
                blnTextInput = True
                lblChange = "  "
                txtText = ""
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                oldFontSize = SelectedLabelT.FontSize
                SelectedLabelT.Visible = False
                lblChange.FontSize = lblChange.FontSize * IIf(blnLabelTextScaleFontSize, SelectedImage.ActualZoom, 1)
                txtText.FontSize = lblChange.FontSize
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                xx = Viewer(Index).ImageXPosition(x, y)
                Yy = Viewer(Index).ImageYPosition(x, y)
                subTextCoordinate SelectedImage, xx, Yy, lblChange   '''''计算交换坐标
                xx = (xx * SelectedImage.ActualZoom - SelectedImage.ActualScrollX) * Screen.TwipsPerPixelX + .left + Viewer(Index).width / Viewer(Index).MultiColumns * (FunImageIsX(i, Viewer(Index)) - 1)
                Yy = (Yy * SelectedImage.ActualZoom - SelectedImage.ActualScrollY) * Screen.TwipsPerPixelY + .top + Viewer(Index).height / Viewer(Index).MultiRows * (FunImageIsY(i, Viewer(Index)) - 1)
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                txtText.Move xx, Yy, lblChange.width, lblChange.height
                txtText.Visible = True
                oldTextleft = xx + lblChange.width
                txtText.SetFocus
            End With
        ElseIf Button = cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift And Button_miLabelVasMeasure Then
            '血管狭窄测量的处理
            If intVasMeasure = 1 Then       '准备画第二套血管－狭窄血管
                intVasMeasure = 2
            ElseIf intVasMeasure = 2 Then
                intVasMeasure = 0
                '显示测量结果
                Set frmVasMeasure.lblText = SelectedLabel.TagObject
                Set frmVasMeasure.f = Me
                frmVasMeasure.Show 1, Me
            End If
        ElseIf Button + cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift And Button_miLabelCadiothoracicRatio Then
            '心胸比测量的处理
            If intCadioThoracicRatio = 1 Then   '准备画胸廓线
                intCadioThoracicRatio = 2
            ElseIf intCadioThoracicRatio = 2 Then
                intCadioThoracicRatio = 0
                '显示测量结果
                Call funcGetCadioThoracicRatio(SelectedLabel, SelectedImage)
            End If
        End If
    End If      ''[标注和自适应调窗的处理]的结束
    '''''''''''''''''''''''[矢冠状重建移动标注处理]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If blnMoveLabel Then
        If blnInMPR = True Then
            '''''''''''''如果对应的图像没有进行初始化则进行初始化
            If ZLMPRCube(2).intViewerIndex < Viewer.Count And ZLMPRCube(3).intViewerIndex < Viewer.Count Then
                If Viewer(ZLMPRCube(2).intViewerIndex).Images(1).Labels.Count = 0 Then
                    subInitImageLabels ZLMPRCube(2).intViewerIndex, 1, Viewer(ZLMPRCube(2).intViewerIndex).Images(1), True, True, True
                    subDrawImgShutter Viewer(ZLMPRCube(2).intViewerIndex).Images(1)
                    subDisplayPatientInfo Viewer(ZLMPRCube(2).intViewerIndex)
                End If
                If Viewer(ZLMPRCube(3).intViewerIndex).Images(1).Labels.Count = 0 Then
                    subInitImageLabels ZLMPRCube(3).intViewerIndex, 1, Viewer(ZLMPRCube(3).intViewerIndex).Images(1), True, True, True
                    subDrawImgShutter Viewer(ZLMPRCube(3).intViewerIndex).Images(1)
                    subDisplayPatientInfo Viewer(ZLMPRCube(3).intViewerIndex)
                End If
            End If
        End If
    End If
    '''''重新显示标注的测量文字
    If Not SelectedLabel Is Nothing Then
        '重新显示区域类型标注的测量信息
        If SelectedLabel.LabelType = doLabelEllipse Or SelectedLabel.LabelType = doLabelPolygon Or SelectedLabel.LabelType = doLabelRectangle Then
            Set SelectedLabelT = SelectedImage.Labels(SelectedImage.Labels.IndexOf(SelectedLabel) + 1)
            SelectedLabelT.Text = funROIResultString(SelectedLabel, SelectedImage)     ''''区域类型标注的文字处理
            
            Viewer(Index).Refresh
        End If
        '重新显示血管狭窄测量的信息
        If blnReSizeLabel And left(SelectedLabel.Tag, 3) = "VAS" Then
            Set frmVasMeasure.lblText = SelectedLabel.TagObject
            Set frmVasMeasure.f = Me
            frmVasMeasure.Show 1, Me
        End If
    End If
                    
    ''''''''''''''''''''''''序列内图像内容同步'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If (Button = cMouseUsage("102").lngMouseKey And Shift = cMouseUsage("102").lngShift And Button_miWidthLevel And blnMouseStart) _
        Or (Button = cMouseUsage("105").lngMouseKey And Shift = cMouseUsage("105").lngShift And Button_miAutoWidthLevel) Then
        Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_WINDOW)
    ElseIf (Button = cMouseUsage("103").lngMouseKey And Shift = cMouseUsage("103").lngShift And Button_miCruise And blnMouseStart) _
        Or (Button = cMouseUsage("104").lngMouseKey And Shift = cMouseUsage("104").lngShift And Button_miZoom And blnMouseStart) Then
        Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If blnMoveLabel = False And blnReSizeLabel = False And blnMouseStart = False And blnFrameSelectImage = False And blnAutoWL = False And Me.MousePointer = 0 Then
        If Button = 2 Then      '显示右键菜单
            ShowPopup Me, Viewer(Index).CurrentImage
        ElseIf Button = 1 And Shift = 2 Then         '通过Ctrl+鼠标左键，选择序列
            ZLShowSeriesInfos(Index).Selected = Not ZLShowSeriesInfos(Index).Selected
            subDispframe Me, Viewer(Index)
            Viewer(Index).Refresh
        ElseIf Button = 4 Then      '中间键，切换浏览和观察模式
            Call subLookOrBrowsSwitch(Me)
        End If
    End If
    
    blnMoveLabel = False
    blnReSizeLabel = False
    blnMouseStart = False
    blnAutoWL = False
    blnFrameSelectImage = False
    Me.MousePointer = 0
    blDicomDown = False                         '放开鼠标
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub VScro_Change(Index As Integer)
    Dim intImageIndex As Integer
    Dim iMove As Integer
    
    On Error GoTo err
    
    If Not Viewer(Index).Visible Or blnVscroInvoked = True Then Exit Sub
    blnAngle = False    '如果当前图像发生改变，清除角度测量标记
    intVasMeasure = 0    '如果当前图像发生改变，则给血管狭窄测量标记清零
    intCadioThoracicRatio = 0   '如果当前图像发生改变，则清空心胸比测量标记
    
    '显示Viewer中的图像
    intImageIndex = VScro(Index).Value
    Call subShowALLImage(Me, Viewer(Index), intImageIndex, False)
    
    SelectedImageIndex = Viewer(Index).CurrentIndex
    iMove = SelectedImageIndex - MSFViewer.TextMatrix(Index, 3)
    Set SelectedImage = Viewer(Index).Images(SelectedImageIndex)
    intSelectedSerial = Index
    MSFViewer.TextMatrix(Index, 3) = SelectedImageIndex
    
    '处理手工序列同步和自动序列同步
    If Button_miSerialManualSyn And ZLShowSeriesInfos(Index).Selected = True Then
        '手工序列同步
        subManualSeriesSyn Me, iMove, Index
    ElseIf ZLShowSeriesInfos(Index).ImageInfos(intImageIndex).SliceLocation <> "" And Button_miSerialPlaceInPhase Then
        '自动序列同步
        subSerialPlaceInPhase Val(ZLShowSeriesInfos(Index).ImageInfos(intImageIndex).SliceLocation), Me
    End If
    
    '处理定位线的显示
    If Button_miAllReferLine = True Or Button_miFLReferLine = True Or Button_miCurrentReferLine = True Then
        Call subDisplayReferLine(Viewer(Index), Me, True)    '根据菜单选项，显示三种类型的定位线
    End If
    
    '如果在MPR状态，切换图像后，要刷新MPR结果图的显示
    If blnInMPR = True Then
        Call subMPRChanegImage(Me)
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VScro_GotFocus(Index As Integer)
    On Error GoTo err
    If Viewer(Index).Visible = True Then Viewer(Index).SetFocus
err:
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then  '''ESC和回车键退出输入
        txtText.Visible = False
        blnTextInput = False
        blnTextInputM = False
        txtText.FontSize = oldFontSize
        lblChange.FontSize = oldFontSize
         '''''''''如果什么都没有输入则删除增加的标志，但是这样对于箭头来说在移动和改变大小的时候可能出错的约(检查后不出错，不是指向那里了)
        If Trim(txtText) = "" Then
            If SelectedImage Is Nothing Then Exit Sub
            SelectedImage.Labels.Remove SelectedImage.Labels.Count
            txtText = "1 "              ''''''?????
        Else
            lblChange = Trim(lblChange)
            SelectedLabelT.Text = lblChange
            SelectedLabelT.Visible = True
        End If
        Viewer(intSelectedSerial).Refresh
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtText_Change()
    ''''''''''''''''''''贴在Viewer上的txt输入框的改变''''''''''''''''''''''''''''''''''''''''''''
    If SelectedImage Is Nothing Then Exit Sub
    lblChange = txtText + "  "
    ''''根据图像的反转情况决定是否反向输入
    If (SelectedImage.RotateState = doRotateNormal And (SelectedImage.FlipState = 3 Or SelectedImage.FlipState = 1)) _
        Or (SelectedImage.RotateState = doRotateLeft And (SelectedImage.FlipState = 3 Or SelectedImage.FlipState = 2)) _
        Or (SelectedImage.RotateState = doRotate180 And (SelectedImage.FlipState = 0 Or SelectedImage.FlipState = 2)) _
        Or (SelectedImage.RotateState = doRotateRight And (SelectedImage.FlipState = 0 Or SelectedImage.FlipState = 1)) Then
        txtText.left = oldTextleft - lblChange.width    ''''oldTextleft在启动txTtext的时候填写
        txtText.width = lblChange.width
    Else
        txtText.width = lblChange.width
        txtText.height = lblChange.height
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtText_LostFocus()
    txtText_KeyPress (13)  ''''鼠标点击导致焦点丢失也退出输入
End Sub

Private Sub subOutPutRptImg()
'------------------------------------------------
'功能：保存报告图，搜索当前打开的所有viewer，把被选中的图象保存成报告图，
'      保存报告图的时候，只能把报告图保存到当前最后打开的图像所在的FTP目录中
'参数：
'返回：直接把被选中的图像保存成报告图
'------------------------------------------------
    Dim im As DicomImage
    Dim imgs As New DicomImages
    Dim imTmp As New DicomImage
    Dim strStudyUID As String
    Dim strSQL  As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim j As Integer
    Dim iImageIndex As Integer
    Dim lngImgContainInfo As Long
    
    '读取报告图包含病人信息参数
    lngImgContainInfo = (Val(zlDatabase.GetPara("报告图包含病人信息", glngSys, 1289, 1)))
    
    '将被选中的图像添加到图像集中
    For i = 1 To ZLShowSeriesInfos.Count
        iImageIndex = 1
        For j = 1 To ZLShowSeriesInfos(i).ImageInfos.Count
            If ZLShowSeriesInfos(i).ImageInfos(j).blnSelected = True Then
                Set im = Nothing
                '首先判断图像是否已经装载，如果已经装载，则找到这个图像并显示出来，如果没有装载，则装载该图像
                If ZLShowSeriesInfos(i).ImageInfos(j).blnDisplayed = False Then
                    funcAddAImageA Viewer(i), j
                End If
                
                '查找图像的索引
                While Viewer(i).Images(iImageIndex).Tag < j And iImageIndex < Viewer(i).Images.Count
                    iImageIndex = iImageIndex + 1
                Wend
                
                If iImageIndex <= Viewer(i).Images.Count Then
                    If Viewer(i).Images(iImageIndex).Tag = j Then
                        Set im = Viewer(i).Images(iImageIndex)
                    End If
                End If
                
                If Not im Is Nothing Then
                    If strStudyUID = "" Then
                        '从数据库中读取检查UID
                        strSQL = "select 检查UID FROM 影像检查序列  where 序列UID =[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(im.SeriesUID))
                        If rsTmp.RecordCount = 0 Then
                            strStudyUID = PstrCheckUID  '用默认值
                        Else
                            strStudyUID = rsTmp!检查UID
                        End If
                    End If
                   
                   
                    '先隐藏图片的四角标注
                    If lngImgContainInfo = 0 Then subInitImageLabels intSelectedSerial, 1, im, False
                    
    
                    Set imTmp = im.Capture(False)
                    
                    '再显示图片的四角标注
                    If lngImgContainInfo = 0 Then subInitImageLabels intSelectedSerial, 1, im, True
                    
                    imTmp.VOILUT = 0
                    '不需要设置InstanceUID，在保存报告图的时候会随机分配一个
                    imTmp.SeriesUID = im.SeriesUID
                    imTmp.StudyUID = strStudyUID        '这一步就保证了图像的检查UID跟数据库中的一致，因此可以顺利保存到对一个的记录中
                    
                    imgs.Add imTmp
                End If
            End If
        Next j
    Next i
    
    '保存图像成报告图
    If imgs.Count > 0 Then
        On Error Resume Next
        SaveImages imgs, 1
        If err <> 0 Then
            MsgBox "报告图保存出错", vbExclamation, gstrSysName
        End If
    Else
        MsgBox "没有被选中的图像，请选择图像后再保存", vbExclamation, gstrSysName
    End If
End Sub

Private Sub subCloseSeries(Optional blnNotify As Boolean = True)
    If intSelectedSerial <> 0 Then
        If blnNotify Then
            If MsgBox("你真的要关闭此图像?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Sub
        End If
            Call subUnloadViewer(intSelectedSerial, Me)
    End If
End Sub

Public Sub subKillPicture(Optional ByVal bSilent As Boolean = False)
'------------------------------------------------
'功能：删除全部图像，并且卸载所有的Viewer和滚动条；初始化图像正本集合以及相关的参数
'参数：
'返回：
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    
    If MSFViewer.Rows < 1 Then Exit Sub
    If bSilent = False Then
        If MsgBox("确定要删除所有图像吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    '如果在MPR状态中，则提示是否保存MPR结果
    If blnInMPR = True Then
        If funViewerMPR(Me, bSilent) = False Then Exit Sub
    End If
    
    '删除图像
    For i = 1 To Me.Viewer.Count - 1
        For j = 1 To MSFViewer.Rows - 1
            If MSFViewer.TextMatrix(j, 1) = True Then
                Unload VScro(j)
                Unload Viewer(j)
                MSFViewer.TextMatrix(j, 1) = False
            End If
        Next
    Next
    ReDim aPixels(0)
    Set ZLSeriesInfos = Nothing '这样清空集合，速度更加快
    Set ZLSeriesInfos = New Collection
    Set ZLShowSeriesInfos = Nothing
    Set ZLShowSeriesInfos = New Collection
    
    intSelectedSerial = 0
    Set SelectedImage = Nothing
    oldSelectedImageIndex = 0
    oldSelectedSerial = 0
    SelectedImageIndex = 0
    intClickImageIndex = 0
    MSFViewer.Rows = 1  '清空原图像列表记录的内容
    Me.txtText.Visible = False
    Set SelectedLabel = Nothing
    '重新显示缩略图
    Call subShowMiniImages(Me)
End Sub

Function funcROIHistogram(LLAA As DicomLabel) As Boolean
'------------------------------------------------
'功能：对于区域类型的标注，画直方图，输入参数为需要画直方图的标注。
'参数：LLAA--需要画直方图的标注；
'返回：True--成功，并直接画出直方图窗体；False--输入标注不是区域类型标注，失败。
'上级函数或过程：frmViewer.Viewer_DblClick；
'下级函数或过程：mdlPublic.Max7InArray
'引用的外部参数：frmHistogram窗体
'编制人：黄捷
'------------------------------------------------
    
    '判断标注是否区域类型的标注，若不是，则不画直方图函数返回False，若是，则直接画出直方图，函数返回True。
    Dim i As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If LLAA Is Nothing Then
        funcROIHistogram = False
        Exit Function
    End If
    
    If (LLAA.LabelType <> doLabelEllipse) And (LLAA.LabelType <> doLabelPolygon) _
       And (LLAA.LabelType <> doLabelRectangle) Then
       funcROIHistogram = False
       Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim x As Long
    Dim WHAT As Variant
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo err
    x = 1
    WHAT = LLAA.ROIValues
    '''''画直方图''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim xx() As Long
    Dim lMin As Long
    Dim lMax As Long
    Dim lCount As Long
    
    ''''对数组进行转换，将维度为像素点，内容为灰度值的数组，转换为维度为灰度值，内容为该灰度值数量的数组
    lMax = LLAA.ROIMax      '临时使用lMax和lMin
    lMin = LLAA.ROIMin
    ReDim xx(lMax - lMin + 1)
    '''''初始化数组XX'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For lCount = 1 To (lMax - lMin + 1)
        xx(lCount) = 0
    Next
    ''''处理数组WHAT''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For lCount = 1 To UBound(WHAT)
        xx(WHAT(lCount) - lMin + 1) = xx(WHAT(lCount) - lMin + 1) + 1
    Next
    ''''填写直方图相关信息'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    frmHistogram.lblStart.Caption = LLAA.ROIMin
    frmHistogram.lblEnd.Caption = LLAA.ROIMax
    Max7InArray xx, lMax, lMin
    frmHistogram.Text1 = xx(lMin)
    frmHistogram.Text2 = xx(lMax)
    frmHistogram.Text3 = LLAA.ROIMin + lMin
    frmHistogram.Text4 = LLAA.ROIMin + lMax
    '''''画图表''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With frmHistogram.MSChart1
        .RowCount = 1
        .ColumnCount = UBound(xx)
        For i = 1 To UBound(xx)
            .Column = i
            .Data = xx(i)
            .Plot.SeriesCollection(i).DataPoints(-1).Brush.FillColor.Set 80, 80, 80
        Next
    End With
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    frmHistogram.Show 1, Me
    funcROIHistogram = True
err:
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' 显示线标注的灰度分布直方图，次函数直接显示灰度分布图窗体，
'''' 输入的值为图像及其上面的标注（直线标注或多边线标注），返回值为是否成功
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function funcDrawGreyDistribute(img As DicomImage, la As DicomLabel) As Boolean
    funcDrawGreyDistribute = False
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (img Is Nothing) Or (la Is Nothing) Then Exit Function
    
    '判断标注是否直线或多边线  '判断标注是否贴在图像上 ，不满足条件则推出函数
    If (la.LabelType <> doLabelLine) And (la.LabelType <> doLabelPolyLine) Then
        Exit Function
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (la.ImageTied = False) Then
        Exit Function
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (la.LabelType = doLabelLine) And ((la.width = 0) Or (la.height = 0)) Then
        Exit Function
    End If
    ''获取直线上灰度值，存放到数组中
    Dim aGrey() As Integer
    Dim beginx As Integer
    Dim beginy As Integer
    Dim endx As Integer
    Dim endy As Integer
    Dim i As Integer
    If funGetLinePoints(img, la, aGrey(), beginx, beginy, endx, endy) = False Then Exit Function
    '灰度分布图上显示起始点，结束点坐标
    frmHistogram.lblStart.Caption = "起点：(" & CStr(beginx) & "," & CStr(beginy) & ")"
    frmHistogram.lblEnd.Caption = "终点：(" & CStr(endx) & "," & CStr(endy) & ")"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '灰度分布图上显示直线距离，
    '灰度分布图的x坐标为直线从左到右，从上到下的点，y坐标为该点的灰度值
    '画柱状图
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With frmHistogram.MSChart1
        .AllowSelections = False
        .RowCount = 1
        .ColumnCount = UBound(aGrey)
        For i = 1 To UBound(aGrey)
            .Column = i
            .Data = aGrey(i)
            .Plot.SeriesCollection(i).DataPoints(-1).Brush.FillColor.Set 80, 80, 80
        Next
    End With
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    frmHistogram.frmMaxAndValue.Visible = False
    frmHistogram.Show 1, Me
End Function

Sub subSelectOnlyOne(ButtomID As Long)
    '------------------------------------------------
    '功能：                                     处理按钮按下和弹起的规则
    '参数：
    '       Serial_ID                           设置为选中的排序方式
    '返回：
    '------------------------------------------------
    '测量
    '处理工具条
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_Text, , True).Checked = False                          '文字
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_Arrowhead, , True).Checked = False                     '箭头
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_Ellipse, , True).Checked = False                       '椭圆
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_Angle, , True).Checked = False                         '角度
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_Curve, , True).Checked = False                         '曲线
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_Area, , True).Checked = False                          '区域
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_BeeLine, , True).Checked = False                       '直线
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_Rect, , True).Checked = False                          '矩形
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_VasMeasure, , True).Checked = False                    '血管狭窄测量
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_CadioThoracicRatio, , True).Checked = False            '心胸比
    '处理菜单
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Lable_Text, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Lable_Arrowhead, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Lable_Ellipse, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Lable_Angle, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Lable_Curve, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Lable_Area, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Lable_BeeLine, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Lable_Rect, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Lable_VasMeasure, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Lable_CadioThoracicRatio, , True).Checked = False
    '公共
    '处理工具条
    ComToolBar.Item(ToolBar_Comm).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow, , True).Checked = False       '手动调窗
    ComToolBar.Item(ToolBar_Comm).FindControl(, ID_Active_Cruise, , True).Checked = False                              '漫游
    ComToolBar.Item(ToolBar_Comm).FindControl(, ID_Active_Zoom, , True).Checked = False                                '缩放
    ComToolBar.Item(ToolBar_Comm).FindControl(, ID_Active_Shuttle, , True).Checked = False                             '穿梭
    '处理菜单
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_AdjustWindow_AutoAdjustWindow, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Cruise, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Zoom, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Shuttle, , True).Checked = False
    
    '平面
    '处理工具条
    ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_3DLine, , True).Checked = False                 '三维鼠标
    ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Tool_ArrowyCoronaryReset, , True).Checked = False                   '矢冠状重建
    '处理菜单
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_PointingLine_3DLine, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Tool_ArrowyCoronaryReset, , True).Checked = False
    
    '处理变量
    Button_miWidthLevel = False
    Button_miAutoWidthLevel = False
    Button_miCruise = False
    Button_miZoom = False
    Button_miStack = False
    Button_miLabeltext = False
    Button_miLabelArrowhead = False
    Button_miLabelEllipse = False
    Button_miLabelAngle = False
    Button_miLabelPolyLine = False
    Button_miLabelPolygon = False
    Button_miLabelLine = False
    Button_miLabelRectangle = False
    Button_mi3dCursor = False
    Button_miLabelVasMeasure = False
    Button_miLabelCadiothoracicRatio = False
    
    blnAngle = False    '如果当前图像发生改变，清除角度测量标记
    intVasMeasure = 0   '给血管狭窄测量标注的标记清零
    intCadioThoracicRatio = 0   '给心胸比测量标记清零
    
    '处理测量工具栏
    If ButtomID = ID_Active_Lable_Text Or ButtomID = ID_Active_Lable_Arrowhead Or ButtomID = ID_Active_Lable_Ellipse _
        Or ButtomID = ID_Active_Lable_Angle Or ButtomID = ID_Active_Lable_Curve Or ButtomID = ID_Active_Lable_Area _
        Or ButtomID = ID_Active_Lable_BeeLine Or ButtomID = ID_Active_Lable_Rect Or ButtomID = ID_Active_Lable_VasMeasure _
        Or ButtomID = ID_Active_Lable_CadioThoracicRatio Then
        
        ComToolBar.Item(ToolBar_Scale).FindControl(, ButtomID, , True).Checked = True
        ComToolBar.Item(ToolBar_Menu).FindControl(, ButtomID, , True).Checked = True
        
    End If
    
    '公共工具条
    If ButtomID = ID_Active_AdjustWindow_HandAdjustWindow Or ButtomID = ID_Active_Cruise _
         Or ButtomID = ID_Active_Zoom Or ButtomID = ID_Active_Shuttle Then
         
         ComToolBar.Item(ToolBar_Comm).FindControl(, ButtomID, , True).Checked = True
         ComToolBar.Item(ToolBar_Menu).FindControl(, ButtomID, , True).Checked = True
    End If
    
    '平面工具条
    If ButtomID = ID_Active_PointingLine_3DLine Or ButtomID = ID_Tool_ArrowyCoronaryReset Then
    
        ComToolBar.Item(ToolBar_Plane).FindControl(, ButtomID, , True).Checked = True
        ComToolBar.Item(ToolBar_Menu).FindControl(, ButtomID, , True).Checked = True
    End If
    
    '其他只更新菜单，不更新工具栏
    If ButtomID = ID_Active_AdjustWindow_AutoAdjustWindow Then
        ComToolBar.Item(ToolBar_Menu).FindControl(, ButtomID, , True).Checked = True
    End If
    
    ComToolBar.RecalcLayout
End Sub

Private Sub subShowScale(lngButtonID As Long)
    '------------------------------------------------
    '功能：设置缩放比例按钮的选中和不选中状态
    '参数：
    '       lngButtonID   被选中的按钮ID
    '返回：
    '------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim cbrControl As CommandBarControl
    
    On Error GoTo err
    
    '首先把所有缩放按钮设置成未选中
    
    For i = 1 To ComToolBar.Count
        
        Set cbrControl = ComToolBar.Item(i).FindControl(, ID_View_ShowScale_50%, True)
        If Not cbrControl Is Nothing Then
            cbrControl.Checked = False
        End If
        
        Set cbrControl = ComToolBar.Item(i).FindControl(, ID_View_ShowScale_100%, True)
        If Not cbrControl Is Nothing Then
            cbrControl.Checked = False
        End If
        
        Set cbrControl = ComToolBar.Item(i).FindControl(, ID_View_showScale_150%, True)
        If Not cbrControl Is Nothing Then
            cbrControl.Checked = False
        End If
        
        Set cbrControl = ComToolBar.Item(i).FindControl(, ID_View_ShowScale_200%, True)
        If Not cbrControl Is Nothing Then
            cbrControl.Checked = False
        End If
        
        Set cbrControl = ComToolBar.Item(i).FindControl(, ID_View_showScale_250%, True)
        If Not cbrControl Is Nothing Then
            cbrControl.Checked = False
        End If
        
        Set cbrControl = ComToolBar.Item(i).FindControl(, ID_View_showScale_300%, True)
        If Not cbrControl Is Nothing Then
            cbrControl.Checked = False
        End If
        
        Set cbrControl = ComToolBar.Item(i).FindControl(, ID_View_showScale_400%, True)
        If Not cbrControl Is Nothing Then
            cbrControl.Checked = False
        End If
        
        Set cbrControl = ComToolBar.Item(i).FindControl(, ID_View_ShowScale_AutoShow, True)
        If Not cbrControl Is Nothing Then
            cbrControl.Checked = False
        End If
        
        Set cbrControl = ComToolBar.Item(i).FindControl(, ID_View_ShowScale_Custom, True)
        If Not cbrControl Is Nothing Then
            cbrControl.Checked = False
        End If
    Next i
    
    '将当前比例的按钮设置成选中
    For i = 1 To ComToolBar.Count
        Set cbrControl = ComToolBar.Item(i).FindControl(, lngButtonID, True)
        If Not cbrControl Is Nothing Then
            cbrControl.Checked = True
        End If
    Next i
    
    ComToolBar.RecalcLayout
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Sub subSelectLeftorRightBouttom(LeftOrRigth As Integer, BouttomID As Long)
    '------------------------------------------------
    '功能：                                     测试左键或右键是否有按下的有就弹起按下的按键
    '参数：
    '       LeftOrRigth                         1为左键2为右键
    '返回：
    '------------------------------------------------
    Dim i As Integer
    Dim cbrControl As CommandBarControl
        
    '窗梭
    If cMouseUsage("101").lngMouseKey = LeftOrRigth And ID_Active_Shuttle = BouttomID Then
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Shuttle, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miStack = True
    Else
        If cMouseUsage("101").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Shuttle, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miStack = False
        End If
    End If

    '漫游
    If cMouseUsage("103").lngMouseKey = LeftOrRigth And ID_Active_Cruise = BouttomID Then
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Cruise, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miCruise = True
    Else
        If cMouseUsage("103").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Cruise, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miCruise = False
        End If
    End If
    
    '裁剪
    If cMouseUsage("201").lngMouseKey = LeftOrRigth And ID_Active_Cut = BouttomID Then
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Cut, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miCutOut = True
    Else
        If cMouseUsage("201").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Cut, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miCutOut = False
        End If
    End If
    
    '框选
    If cMouseUsage("201").lngMouseKey = LeftOrRigth And ID_ACtive_FrameSelectImage = BouttomID Then
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_ACtive_FrameSelectImage, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miFrameSelectImage = True
    Else
        If cMouseUsage("201").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_ACtive_FrameSelectImage, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miFrameSelectImage = False
        End If
    End If

    '手动调窗
    If cMouseUsage("102").lngMouseKey = LeftOrRigth And ID_Active_AdjustWindow_HandAdjustWindow = BouttomID Then
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miWidthLevel = True
    Else
        If cMouseUsage("102").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miWidthLevel = False
        End If
    End If
    
    '自适应调窗
    If cMouseUsage("105").lngMouseKey = LeftOrRigth And ID_Active_AdjustWindow_AutoAdjustWindow = BouttomID Then
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_AdjustWindow_AutoAdjustWindow, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miAutoWidthLevel = True
    Else
        If cMouseUsage("105").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_AdjustWindow_AutoAdjustWindow, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miAutoWidthLevel = False
        End If
    End If
    
    '三维鼠标
    If cMouseUsage("106").lngMouseKey = LeftOrRigth And ID_Active_PointingLine_3DLine = BouttomID Then
        Button_mi3dCursor = Not Button_mi3dCursor
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_PointingLine_3DLine, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = Button_mi3dCursor
            End If
        Next i
    Else
        If cMouseUsage("106").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_PointingLine_3DLine, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_mi3dCursor = False
        End If
    End If
    
    '缩放
    If cMouseUsage("104").lngMouseKey = LeftOrRigth And ID_Active_Zoom = BouttomID Then
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Zoom, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miZoom = True
    Else
        If cMouseUsage("104").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Zoom, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miZoom = False
        End If
    End If
    
    '文字
    If cMouseUsage("8").lngMouseKey = LeftOrRigth And ID_Active_Lable_Text = BouttomID Then
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_Text, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miLabeltext = True
    Else
        If cMouseUsage("8").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_Text, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miLabeltext = False
        End If
    End If
    
    '箭头
    If cMouseUsage("4").lngMouseKey = LeftOrRigth And ID_Active_Lable_Arrowhead = BouttomID Then
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_Arrowhead, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miLabelArrowhead = True
    Else
        If cMouseUsage("4").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_Arrowhead, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miLabelArrowhead = False
        End If
    End If
    
    '椭圆
    If cMouseUsage("3").lngMouseKey = LeftOrRigth And ID_Active_Lable_Ellipse = BouttomID Then
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_Ellipse, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miLabelEllipse = True
    Else
        If cMouseUsage("3").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_Ellipse, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miLabelEllipse = False
        End If
    End If
    
    '角度
    If cMouseUsage("7").lngMouseKey = LeftOrRigth And ID_Active_Lable_Angle = BouttomID Then
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_Angle, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miLabelAngle = True
    Else
        If cMouseUsage("7").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_Angle, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miLabelAngle = False
        End If
    End If
    
    '曲线
    If cMouseUsage("6").lngMouseKey = LeftOrRigth And ID_Active_Lable_Curve = BouttomID Then
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_Curve, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miLabelPolyLine = True
    Else
        If cMouseUsage("6").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_Curve, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miLabelPolyLine = False
        End If
    End If
    
    '区域
    If cMouseUsage("5").lngMouseKey = LeftOrRigth And ID_Active_Lable_Area = BouttomID Then
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_Area, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miLabelPolygon = True
    Else
        If cMouseUsage("5").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_Area, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miLabelPolygon = False
        End If
    End If
    
    '直线
    If cMouseUsage("1").lngMouseKey = LeftOrRigth And ID_Active_Lable_BeeLine = BouttomID Then
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_BeeLine, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miLabelLine = True
    Else
        If cMouseUsage("1").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_BeeLine, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miLabelLine = False
        End If
    End If
    
    '血管狭窄测量
    If cMouseUsage("1").lngMouseKey = LeftOrRigth And ID_Active_Lable_VasMeasure = BouttomID Then
        intVasMeasure = 0   '给血管狭窄测量标注的标记清零
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_VasMeasure, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miLabelVasMeasure = True
    Else
        If cMouseUsage("1").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_VasMeasure, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miLabelVasMeasure = False
        End If
    End If
    
    '心胸比测量
    If cMouseUsage("1").lngMouseKey = LeftOrRigth And ID_Active_Lable_CadioThoracicRatio = BouttomID Then
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_CadioThoracicRatio, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miLabelCadiothoracicRatio = True
    Else
        If cMouseUsage("1").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_CadioThoracicRatio, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miLabelCadiothoracicRatio = False
        End If
    End If
    
    '矩形
    If cMouseUsage("2").lngMouseKey = LeftOrRigth And ID_Active_Lable_Rect = BouttomID Then
        For i = 1 To ComToolBar.Count
            Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_Rect, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Checked = True
            End If
        Next i
        Button_miLabelRectangle = True
    Else
        If cMouseUsage("2").lngMouseKey = LeftOrRigth Then
            For i = 1 To ComToolBar.Count
                Set cbrControl = ComToolBar.Item(i).FindControl(, ID_Active_Lable_Rect, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Checked = False
                End If
            Next i
            Button_miLabelRectangle = False
        End If
    End If
    
    subCutOut Me
End Sub

Public Function funFilm(f As frmViewer, blnShowForm As Boolean, intAddType As Integer, _
    Optional intInterval As Integer = 1, Optional blnStartOdd As Boolean = True) As Boolean
'------------------------------------------------
'功能：胶片打印
'参数： f--进行打印的窗体。
'       blnShowForm --- 是否显示打印窗体
'       intAddType  --  添加图像的方式，1-添加序列；2-添加当前图；3-添加所选图;4-间隔添加序列
'       intInterval -- 间隔打印的间隔数，当intAddType=4时使用
'       blnStartOdd -- True 奇数起；False 偶数起，当intAddType=4时使用
'返回：True--成功打开打印窗体；False -- 失败。
'2009用
'------------------------------------------------
    
    On Error GoTo err
    
    '先判断胶片打印机的数量是否超过许可的数量
    If (cDICOMPrinter.Count > gint胶片打印机 And gint胶片打印机 <> -1) Or gint胶片打印机 = 0 Then
        Call MsgBox(LOGIN_TYPE_胶片打印机 & "超过您购买的总数量（" & gint胶片打印机 & "）。请向软件供应商联系。", vbOKOnly, gstrSysName)
        Exit Function
    End If
    
    If mfrmFilm Is Nothing Then
        Set mfrmFilm = New frmFilm
        Set mfrmFilm.f = f
        mfrmFilm.Show , f
        
        If blnShowForm = False Then
            mfrmFilm.Hide
            DoEvents
        End If
    Else
        If blnShowForm Then
            mfrmFilm.Show , f
        End If
    End If
    
    Call subFilmAddImages(intAddType, intInterval, blnStartOdd)
    
    '添加图像成功，使用声音提示,显示胶片窗口的时候，不提示声音
    If blnShowForm = False Then
        Call PrintFilmBeep(1)
    End If
    
    '    挂上截获消息的hook，不能放在mfrmFilm的load事件
    If plngFilmPreWndProc = 0 Then
        plngFilmPreWndProc = FilmHook(mfrmFilm.hwnd)
    End If
        
    funFilm = True
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub subDelRepImg()
    If intSelectedSerial <> 0 Then         '当前被选中的序列
        If Me.MSFViewer.TextMatrix(intSelectedSerial, 1) Then   '该序列有图像
             frmDelRptImg.pSeriesUID = Me.Viewer(intSelectedSerial).Images(1).SeriesUID
             Set frmDelRptImg.f = Me
        End If
     End If
        frmDelRptImg.Show 1, Me '
End Sub

Private Sub SaveFrameSelectImageIntoReport(img As DicomImage, lblFrame As DicomLabel)
    Dim imgResult As DicomImage
    Dim imgs As New DicomImages
    Dim iPlane As Integer
    Dim dblZoom As Double
    Dim iLeft As Integer
    Dim iRight As Integer
    Dim iTop As Integer
    Dim iBottom As Integer
    Dim iMax As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strStudyUID As String
    Dim lngImgContainInfo As Long
    
    If Abs(lblFrame.width) = 0 Or Abs(lblFrame.height) = 0 Or img.Labels.Count < 2 Then
        MsgBox "请选择图像区域后再保存", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    '图象最大宽高=1281
    iMax = 1281
    
    '读取报告图包含病人信息参数
    lngImgContainInfo = (Val(zlDatabase.GetPara("报告图包含病人信息", glngSys, 1289, 1)))
    
    '根据label来提取被框选中的图像
    '图象位数,黑白图像为1，彩色图像为3
    iPlane = 1
    If Not IsNull(img.Attributes(&H28, &H4).Value) And img.Attributes(&H28, &H4).Exists Then
        If img.Attributes(&H28, &H4).Value = "RGB" Then
            iPlane = 3
        End If
    End If
    
    '图象框的位置
    If lblFrame.width >= 0 Then
        iLeft = lblFrame.left
        iRight = iLeft + lblFrame.width
    Else
        iLeft = lblFrame.left + lblFrame.width
        iRight = lblFrame.left
    End If
    
    If lblFrame.height >= 0 Then
        iTop = lblFrame.top
        iBottom = iTop + lblFrame.height
    Else
        iTop = lblFrame.top + lblFrame.height
        iBottom = lblFrame.top
    End If
    
    '控制结果图象的大小在300*300之内
    If (iRight - iLeft) > iMax Or (iBottom - iTop) > iMax Then
        dblZoom = iMax / (iRight - iLeft)
        If dblZoom > iMax / (iBottom - iTop) Then dblZoom = iMax / (iBottom - iTop)
    Else
        dblZoom = 1
    End If
    
    img.Labels(img.Labels.Count).Visible = False
    img.Labels(img.Labels.Count - 1).Visible = False
    
    '先隐藏图片的四角标注
    If lngImgContainInfo = 0 Then subInitImageLabels intSelectedSerial, 1, img, False
    
    If (img.RotateState = doRotateLeft And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipBoth) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipHorizontal) Then
        'X方向对调
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, img.sizeX - iRight, img.sizeX - iLeft, iTop, iBottom)
    ElseIf (img.RotateState = doRotateLeft And img.FlipState = doFlipBoth) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipVertical) Then
        'Y方向对调
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, img.sizeY - iBottom, img.sizeY - iTop)
    ElseIf (img.RotateState = doRotateRight And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateLeft And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipBoth) Then
        'X，Y方向对调
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, img.sizeX - iRight, img.sizeX - iLeft, img.sizeY - iBottom, img.sizeY - iTop)
    Else
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, iTop, iBottom)
    End If
    
    
    '再显示图片的四角标注
    If lngImgContainInfo = 0 Then subInitImageLabels intSelectedSerial, 1, img, True
    
    '不需要设置InstanceUID，在保存报告图的时候会随机分配一个InstanceUID
    imgResult.SeriesUID = img.SeriesUID
    
    '从数据库中读取检查UID
    strSQL = "select 检查UID FROM 影像检查序列  where 序列UID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(imgResult.SeriesUID))
    If rsTemp.RecordCount = 0 Then
        MsgBox "在数据库中无法查到该图像，外部图像不能保存成报告图。", vbExclamation, gstrSysName
        Exit Sub
    Else
        strStudyUID = rsTemp!检查UID
    End If
    imgResult.StudyUID = strStudyUID
    
    imgs.Add imgResult
    '把结果图像保存成报告图
    '保存图像成报告图
    If imgs.Count > 0 Then
        On Error Resume Next
        SaveImages imgs, 1
        If err <> 0 Then
            MsgBox "报告图保存出错", vbExclamation, gstrSysName
        End If
    Else
        MsgBox "没有被选中的图像，请选择图像后再保存", vbExclamation, gstrSysName
    End If
End Sub

Public Function funcSwapSeries(intViewerIndex As Integer, intSeriesIndex As Integer, Optional blnShowLast As Boolean = False) As Boolean
'------------------------------------------------
'功能： 交换序列,用intSeriesIndex指向的序列中的图像，替代viewer(intViewerIndex)中的图像
'参数： intViewerIndex--Viewer的索引
'       intSeriesIndex--图像所在的序列索引
'       blnShowLast --- 可选参数，是否显示最后一个图
'返回：交换成功，则返回True，否则返回False
'时间：2009-7
'------------------------------------------------
    Dim oneImageInfo As clsImageInfo
    Dim i As Integer
    Dim blnSelected As Boolean
    
    On Error GoTo err
    
    '首先判断是否在矢冠状重建状态，而且被代替的Viewer是矢冠状重建的序列或者结果序列，如果是则先退出矢冠状位重建
    If blnInMPR = True And (ZLMPRCube(1).intViewerIndex = intViewerIndex Or ZLMPRCube(2).intViewerIndex = intViewerIndex _
        Or ZLMPRCube(3).intViewerIndex = intViewerIndex) Then
            Call funViewerMPR(Me)
            Exit Function
    End If
    
    '首先根据intViewerIndex，找到旧序列的index
    '把旧序列中的图像清空
    Viewer(intViewerIndex).Images.Clear
    
    '记录旧序列的选择状态
    blnSelected = ZLShowSeriesInfos(intViewerIndex).Selected
    
    '在ZLShowSeriesInfos中用新内容替换旧内容
    Call funCopySeriesInfo(ZLSeriesInfos(intSeriesIndex), ZLShowSeriesInfos(intViewerIndex))
    
    '在新序列中恢复选择状态
    ZLShowSeriesInfos(intViewerIndex).Selected = blnSelected
    
    '复制图像
    Set ZLShowSeriesInfos(intViewerIndex).ImageInfos = Nothing
    For i = 1 To ZLSeriesInfos(intSeriesIndex).ImageInfos.Count
        Set oneImageInfo = funGetNewImageInfo
        Call funCopyImageInfo(ZLSeriesInfos(intSeriesIndex).ImageInfos(i), oneImageInfo)
        ZLShowSeriesInfos(intViewerIndex).ImageInfos.Add oneImageInfo
    Next i
    
    '图像排序,正本图像从数据库中读取，是按照图像号排序的，这里要重新设置显示序列的排序方法
    Call subSortImages(Me, intViewerIndex, funGetImageSort(ZLSeriesInfos(intSeriesIndex).strModality))

    '如果曾经通过双击改变图像显示布局，这里按照新布局显示图像
    If MSFViewer.TextMatrix(intSelectedSerial, 5) > 1 Or MSFViewer.TextMatrix(intSelectedSerial, 6) > 1 Then
        Viewer(intViewerIndex).MultiColumns = MSFViewer.TextMatrix(intSelectedSerial, 5)
        Viewer(intViewerIndex).MultiRows = MSFViewer.TextMatrix(intSelectedSerial, 6)
    End If
    
    '把新序列的图像显示到Viewer中
    If blnShowLast = True Then  '显示最后一个图
        Call subShowALLImage(Me, Viewer(intViewerIndex), ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count, False)
    Else    '显示第一个图
        Call subShowALLImage(Me, Viewer(intViewerIndex), 1, False)
    End If
    
    '设置新Viewer的Tag和滚动条
    Viewer(intViewerIndex).Tag = intSeriesIndex
    
    blnVscroInvoked = True
    Call subDisplayScrollBar(intViewerIndex, Me, False)
    blnVscroInvoked = False
    
    '如果是选择所有序列的状态，将目前序列设置为当前序列，供subDispframe使用
    If isSelectAllSerial Then intSelectedSerial = intViewerIndex
    
    '图像显示完后，添加Viewer中的标注：图象框、右下角的选择标记等
    Call subDispframe(Me, Viewer(intViewerIndex))
    
    '更换序列后，设置当前选中的序列
    SelectedImageIndex = Viewer(intViewerIndex).CurrentIndex
    Set SelectedImage = Viewer(intViewerIndex).Images(SelectedImageIndex)
    intSelectedSerial = intViewerIndex
    MSFViewer.TextMatrix(intViewerIndex, 3) = SelectedImageIndex
    
    '如果序列改变，则触发事件
    If Not SelectedImage Is Nothing Then
        RaiseEvent AfterSeriesChanged(SelectedImage.StudyUID, SelectedImage.SeriesUID)
    End If
    
    funcSwapSeries = True
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function


Public Sub subSelectAViewer(intViewerIndex As Integer, intClickImageIndex As Integer)
'------------------------------------------------
'功能： 选中intViewerIndex指定的Viewer，更新MSFViewer参数信息，重新显示原有序列和新序列的外框
'       重新设置窗宽窗位菜单，同时处理定位线，序列同步等操作
'参数： intViewerIndex--新选择的Viewer的Index
'       intClickImageIndex--新选择的图像的Index
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim blnSeriesChanged As Boolean
    Dim blnImageChanged As Boolean
    Dim intOldSelectedSeries As Integer
    Dim CmdControl As CommandBarControl
    
    '图像发生改变和序列发生改变，应该有不同的处理方法
    
    On Error GoTo err
    
    If Viewer(intViewerIndex).Images.Count = 0 Then Exit Sub
    
    blnSeriesChanged = (intSelectedSerial <> intViewerIndex)
    'blnImageChanged的条件不对，当图像布局为1*1的时候，拖动滚动条时，会改变SelectedImageIndex的值，导致判断错误
    blnImageChanged = (SelectedImageIndex <> intClickImageIndex)
    
    If blnImageChanged = True Or blnSeriesChanged = True Then
        '如果更改了序列或者更改了被选择的图像，则撤销原来的一些标记
        blnAngle = False        '如果当前图像发生改变，则清除角度测量标记
        intVasMeasure = 0       '如果当前图像发生改变，则给血管狭窄测量标记清零
        intCadioThoracicRatio = 0   '如果当前图像发生改变，则给心胸比测量标记清零
    End If
    
    '记录多行或者多列显示时的图像信息
    '5：该序列横向显示图像数目；6：该序列纵向显示图像数目；7：该序列当前显示第一个图像序号；8：该序列当前显示选择图像序号
    If Viewer(intViewerIndex).MultiColumns > 1 Or Viewer(intViewerIndex).MultiRows > 1 Then
        MSFViewer.TextMatrix(intViewerIndex, 5) = Viewer(intViewerIndex).MultiColumns
        MSFViewer.TextMatrix(intViewerIndex, 6) = Viewer(intViewerIndex).MultiRows
        MSFViewer.TextMatrix(intViewerIndex, 7) = Viewer(intViewerIndex).CurrentIndex
        MSFViewer.TextMatrix(intViewerIndex, 8) = intClickImageIndex
    End If
    
    '处理被选中的图像
    If intClickImageIndex <> 0 Then
        '3：当前选择的图像号；4：当前选择的图像处于第几帧
        MSFViewer.TextMatrix(intViewerIndex, 3) = intClickImageIndex
        MSFViewer.TextMatrix(intViewerIndex, 4) = Viewer(intViewerIndex).Images(intClickImageIndex).Frame
        
        Set SelectedImage = Viewer(intViewerIndex).Images(intClickImageIndex)
        SelectedImageIndex = intClickImageIndex
        intSelectedSerial = intViewerIndex
        
        '显示当前显示的图像号和当前序列的图像总数
        sbStatusBar.Panels(3).Text = "当前图像：" & Viewer(intViewerIndex).Images(intClickImageIndex).Tag _
            & "  总数为：" & ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
            
        '处理定位线，根据菜单选项，显示三种类型的定位线
        If Button_miAllReferLine = True Or Button_miFLReferLine = True Or Button_miCurrentReferLine = True Then
            Call subDisplayReferLine(Viewer(intViewerIndex), Me, False)
        End If
    End If
    
    '记录旧的序列号
    If blnSeriesChanged Then
        intOldSelectedSeries = intSelectedSerial
        intSelectedSerial = intViewerIndex
        If intOldSelectedSeries = 0 Then
            intOldSelectedSeries = intSelectedSerial
        ElseIf intOldSelectedSeries >= Viewer.Count Then
            intOldSelectedSeries = intSelectedSerial
        End If
        If intOldSelectedSeries <> intSelectedSerial Then
            '处理旧序列的边框
            subDispframe Me, Viewer(intOldSelectedSeries)
            Viewer(intOldSelectedSeries).Refresh
        End If
    End If
    
    '处理新序列的边框
    subDispframe Me, Viewer(intViewerIndex)
    
    '处理弹出的窗宽窗位菜单
    Call subSetWidthLevelF(Viewer(intViewerIndex).Images(1), Me)
    '处理弹出的图像滤镜菜单
    Call subSetFilterF(Viewer(intViewerIndex).Images(1), Me)
    '处理图像排序菜单勾选项
    Call subSetImageFortF(Me)
    
    '看看是否需要刷新Viewer
    Viewer(intViewerIndex).Refresh
    
    '已经选择了一个序列，则取消序列全选
    isSelectAllSerial = False
    Set CmdControl = ComToolBar.Item(ToolBar_Comm).FindControl(, ID_Active_Select_SelectAllSerial)
    CmdControl.Checked = False
    
    '如果序列改变，触发序列改动事件
    If blnSeriesChanged = True Then
        If Not SelectedImage Is Nothing Then
            RaiseEvent AfterSeriesChanged(SelectedImage.StudyUID, SelectedImage.SeriesUID)
        End If
    End If
    
    '如果在MPR状态，切换图像后，要刷新MPR结果图的显示
    If blnInMPR = True And (blnImageChanged = True Or blnSeriesChanged = True) Then
        Call subMPRChanegImage(Me)
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subCreateAndPlaceAViewer(intSeriesIndex As Integer, ByVal intRow As Integer, ByVal intCol As Integer)
'------------------------------------------------
'功能： 创建一个Viewer，其中放置intSeriesIndex序列的图像，并摆放在picViewer中的lngX和lngY坐标位置
'参数： intSeriesIndex--ZLSeriesInfos中序列的索引
'       intRow -- 新Viewer所在的行数
'       intCol -- 新Viewer所在的列数
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim intCurrentViewer As Integer
    Dim intViewerIndex As Integer
    
    '首先判断这个位置是否有Viewer
    intViewerIndex = (intRow - 1) * intCountX + intCol
    If intViewerIndex >= Viewer.Count Then
        '放置在这里的图像，需要新创建一个Viewer
        intCurrentViewer = funcCeateAViewer(intSeriesIndex, Me)
    Else
        '用现在的Viewer
        Call funcSwapSeries(intViewerIndex, intSeriesIndex)
        intCurrentViewer = intViewerIndex
    End If
    
    '摆放这个Viewer并设置滚动条
    Call subPlaceAViewer(Me, intCurrentViewer, intRow, intCol)
End Sub

Private Sub sub3DCursorStart(thisImage As DicomImage)
'------------------------------------------------
'功能： 开始三维鼠标
'参数： thisImage ---- 三维鼠标进行时鼠标所在的图像
'返回： 是否成功
'时间：2009-7
'------------------------------------------------
    Dim i As Integer
    Dim intViewerIndex As Integer
    Dim intImageIndex As Integer
    Dim sourceFrameOfReferenceUID As String
    Dim destFrameOfReferenceUID As String
    Dim intCurrentImage As Integer
    Dim imgTemp As New DicomImage
    Dim labelRef As DicomLabel
    
    '如果当前图像没有参考帧UID，则不进行三维鼠标操作
    sourceFrameOfReferenceUID = GetImageAttribute(thisImage.Attributes, ATTR_参考帧UID)
    If sourceFrameOfReferenceUID = "" Then
        blnIn3dCursor = False
        Exit Sub
    End If
    
    '循环当前现实的Viewer，找到参与本次三维鼠标的Viewer的Index
    On Error GoTo err
    
    '''''''''''''''''''''''''''定义各种备份变量''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ReDim obj3dImage(Viewer.Count)
    ReDim int3dImageIndex(Viewer.Count)
    ReDim int3dCurrentlyImage(Viewer.Count)
    
    For i = 1 To Viewer.Count - 1
        If i <> intSelectedSerial And Viewer(i).Images.Count > 0 Then '图像所在的Viewer不处理
            intImageIndex = Val(MSFViewer.TextMatrix(i, 3))      '3=当前选择的图像号
            destFrameOfReferenceUID = GetImageAttribute(Viewer(i).Images(intImageIndex).Attributes, ATTR_参考帧UID)
            If sourceFrameOfReferenceUID = destFrameOfReferenceUID Then
                '只有两个序列中的图像的参考帧UID相同，才做三维鼠标
                '''''''''''''''''备份当前图像'''''''''''''''''''''''''''''''''''
                Set obj3dImage(i) = New DicomImage
                Set obj3dImage(i) = Viewer(i).Images(intImageIndex)
                int3dImageIndex(i) = intImageIndex
                int3dCurrentlyImage(i) = Viewer(i).CurrentIndex
                
                '循环这个序列中的所有图像，画定位线
                For intCurrentImage = 1 To ZLShowSeriesInfos(i).ImageInfos.Count
                    '填写虚拟图像的内容，然后计算定位线
                    Call subWriteRefLineImage(imgTemp, intCurrentImage, Viewer(i))
                    If sourceFrameOfReferenceUID = ZLShowSeriesInfos(i).ImageInfos(intCurrentImage).FrameOfReferenceUID Then
                        '画定位线
                        Set labelRef = New DicomLabel
                        Set labelRef = thisImage.ReferenceLine(imgTemp, True)
                        labelRef.ForeColour = vbBlue
                        labelRef.Tag = "3DL" & i & "-" & intCurrentImage
                        labelRef.LineStyle = 2
                        thisImage.Labels.Add labelRef
                    End If
                Next intCurrentImage
            End If
        End If
    Next i
    blnIn3dCursor = True
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Sub sub3DCursorEnd(thisImage As DicomImage)
'------------------------------------------------
'功能： 结束三维鼠标
'参数： thisImage ---- 三维鼠标进行时鼠标所在的图像
'返回： 无
'时间：2009-7
'------------------------------------------------
    Dim k As Integer, i As Integer, v As DicomViewer
    ''''''''''''''''''[删除所有的定位线]''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    subDeleteAppointLabel thisImage, "3D"
    thisImage.Refresh False
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If blnIn3dCursor = False Then Exit Sub
    
    For Each v In Viewer
        i = v.Index
        If i <> intSelectedSerial And i <> 0 And int3dImageIndex(i) <> 0 Then     '这个系列参与了三维鼠标
            k = MSFViewer.TextMatrix(i, 3)        '当前选择的图像号
            Viewer(i).Images.Add obj3dImage(i)
            Viewer(i).Images.Move Viewer(i).Images.Count, k
            subLabelCopyRebuild obj3dImage(i), v.Images(k)           ''''处理标注断链的问题
            Viewer(i).Images.Remove k + 1
'            viewer(i).Refresh
            
            If VScro(i).Visible = True Then VScro(i).Value = int3dImageIndex(i)
        End If
    Next
End Sub

Private Sub subStackMouseMove(lngDeltaY As Long)
'------------------------------------------------
'功能： 鼠标穿梭，移动鼠标时的操作
'参数： lngDeltaY ---- 鼠标在Viewer中的Y方向的位移量，通过
'返回： 无
'时间：2009-7
'------------------------------------------------
    Dim objTempImage As DicomImage
    Dim intNewFrame As Integer
    
    On Error GoTo err
    If Not blnMouseStart Then                           ''''如果没有开始穿梭，则现在开始穿梭
        Me.MouseIcon = ImageListMouse.ListImages("Stack").Picture
        Me.MousePointer = 99
        blnMouseStart = True
        intStackOffset = SelectedImageIndex - Viewer(intSelectedSerial).CurrentIndex    '记录当前图像跟CurrentIndex之间的距离
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If SelectedImage.FrameCount > 1 Then   ''''多帧图像处理
            intStackCurrentlyImage = SelectedImage.Frame
            blnStackisFrame = True
        Else
            blnStackisFrame = False
            '记录穿梭前Viewer的CurrentIndex和当前图像
            Set SelectedLabel = Nothing
            intStackCurrentlyImage = Viewer(intSelectedSerial).CurrentIndex
            Set objStackOldImage = Viewer(intSelectedSerial).Images(SelectedImageIndex)
            intStackIndex = Viewer(intSelectedSerial).Images(SelectedImageIndex).Tag
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If blnStackisFrame And SelectedImage.FrameCount <> 1 Then            ''''多祯图像
        intNewFrame = SelectedImage.Frame + (lngDeltaY) / lngStackStep
        If intNewFrame <= 0 Then intNewFrame = 1
        If intNewFrame > SelectedImage.FrameCount Then intNewFrame = SelectedImage.FrameCount
        SelectedImage.Frame = intNewFrame
    ElseIf Not blnStackisFrame And ZLShowSeriesInfos(intSelectedSerial).ImageInfos.Count <> 1 Then  ''''单祯图像
        '计算新图像的index
        intStackIndex = intStackIndex + (lngDeltaY) / lngStackStep
        If intStackIndex <= 0 Then intStackIndex = 1
        If intStackIndex > ZLShowSeriesInfos(intSelectedSerial).ImageInfos.Count Then intStackIndex = ZLShowSeriesInfos(intSelectedSerial).ImageInfos.Count
        '把指定位置的图像添加到Viewer中
        Set objTempImage = funLoadAImage(intSelectedSerial, intStackIndex, 1)
        If Not objTempImage Is Nothing Then
            Call subInitAImage(objTempImage, intSelectedSerial, Viewer(intSelectedSerial))
            
            Viewer(intSelectedSerial).Images.Add objTempImage
            Viewer(intSelectedSerial).Images.Move Viewer(intSelectedSerial).Images.Count, SelectedImageIndex
            Viewer(intSelectedSerial).Images.Remove SelectedImageIndex + 1
            Viewer(intSelectedSerial).CurrentIndex = intStackCurrentlyImage
        End If
    End If
    blnStackStart = True
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub sub3DCursorMouseMove(x As Long, y As Long, thisViewer As DicomViewer)
'------------------------------------------------
'功能： 三维鼠标，移动鼠标时的操作
'参数： x ---- 鼠标在Viewer中的X方向位置
'       y ---- 鼠标在Viewer中的Y方向位置
'       thisViewer ---- 鼠标所在的Viewer
'返回： 无
'时间：2009-7
'------------------------------------------------
    Dim ls As DicomLabels
    Dim l As DicomLabel
    Dim img As DicomImage
    Dim str3DMoveTag As String  '记录当前三维鼠标操作的标注的一些特征
    Dim objTempImage As DicomImage  '三维鼠标用的临时图像
    Dim j As Integer, k As Integer, ii As Integer
    
    On Error GoTo err
    If blnIn3dCursor = False Then Exit Sub
    
    Set ls = thisViewer.LabelHits(x, y, False, False, True)  'ls是所有被单击的标注的集合
    Set img = thisViewer.Images(thisViewer.imageIndex(x, y))  'img是当前图像
    '标注集合中有标注存在，则进行三维鼠标的翻图操作
    If ls.Count <> 0 Then
        str3DMoveTag = ""   '清空特征标记,目的是避免当鼠标在同一个定位线移动的时候，多次执行切换图像的语句
        '循环所有标注，查找是否有三维鼠标的标注
        For Each l In ls
            '通过条件判断是否是三维鼠标的标注线
            If Mid(l.Tag, 1, 2) = "3D" Then
                k = InStr(l.Tag, "-")
                j = Val(Mid(l.Tag, 4, k - 1))   '标注线对应的图像所在的Viewer的Index
                k = Val(Mid(l.Tag, k + 1))      '标注线对应的图像所在的图像号
                ii = Me.MSFViewer.TextMatrix(j, 3)  '当前选择的图像号，就在这个图像格进行图像的穿梭
                If str3DMoveTag <> Mid(l.Tag, 1, 5) And (ZLShowSeriesInfos.Count >= j) Then
                    '先穿梭图像，实际上就是穿梭的方法
                    Set objTempImage = funLoadAImage(j, k, 1)
                    If Not objTempImage Is Nothing Then
                        Call subInitAImage(objTempImage, j, Viewer(j))
                        
                        Viewer(j).Images.Add objTempImage
                        Viewer(j).Images.Move Viewer(j).Images.Count, ii
                        Viewer(j).Images.Remove ii + 1
                        Viewer(j).CurrentIndex = int3dCurrentlyImage(j)
                    
                        '计算图像上面的红色十字投影的比例
                        Dim cy As Double
                        '定位线是两个图像相交在X，Y，Z面方向上的投影，直接使用图像横向或者纵向的像素数量来做比例
                        If Abs(l.height) > Abs(l.width) Then
                            cy = thisViewer.ImageYPosition(x, y) / img.sizeY '* IIf(l.height < 0, -1, 1)
                        Else
                            cy = thisViewer.ImageXPosition(x, y) / img.sizeX '* IIf(l.width < 0, -1, 1)
                        End If
                        
                        '画定位线和红色十字投影,先画定位线
                        Dim lRefLine As DicomLabel
                        Set lRefLine = New DicomLabel
                        Set lRefLine = Viewer(j).Images(ii).ReferenceLine(SelectedImage, True)
                        lRefLine.ForeColour = vbBlue
                        lRefLine.LineStyle = 2
                        Viewer(j).Images(ii).Labels.Add lRefLine
                        
                        '画十字交叉的横线
                        Set l = GetNewLabel(3, 0, 0, 20, 0)
                        l.ForeColour = vbRed
                        l.Tag = "3DH"
                        l.XOR = False
                        Viewer(j).Images(ii).Labels.Add l
                        l.left = lRefLine.left + cy * lRefLine.width - 10
                        l.top = lRefLine.top + cy * lRefLine.height
                        
                        '画十字交叉的竖线
                        Set l = GetNewLabel(3, 0, 0, 0, 20)
                        l.ForeColour = vbRed
                        l.Tag = "3DV"
                        l.XOR = False
                        Viewer(j).Images(ii).Labels.Add l
                        l.left = lRefLine.left + cy * lRefLine.width
                        l.top = lRefLine.top + cy * lRefLine.height - 10
                        
                        str3DMoveTag = Mid(l.Tag, 1, 5)
                        int3dImageIndex(j) = k
                    End If
                End If
            End If
        Next
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub subLabelDeleAll()
'------------------------------------------------
'功能：删除窗体中被选中图像的所有用户标注
'参数：无
'返回：无，直接删除用户标注
'2009用
'------------------------------------------------
    Dim i As Integer
    If SelectedImage Is Nothing Then Exit Sub
    
    For i = SelectedImage.Labels.Count To G_INT_SYS_LABEL_COUNT + 1 Step -1
        SelectedImage.Labels.Remove i
    Next
    Set SelectedLabel = Nothing
    blnAngle = False                    '给角度测量标记清零
    intVasMeasure = 0                   '给血管狭窄测量标记清零
    intCadioThoracicRatio = 0           '给心胸比测量标记清零
    SubNoDispPeriod SelectedImage, Me      '为指定对象关闭句柄
End Sub

Private Sub subDelSelectedLabel()
'------------------------------------------------
'功能：删除窗体中被选中图像的选中标注
'参数：无
'返回：无，直接删除用户标注
'2009用
'------------------------------------------------
    Dim lblThis As DicomLabel
    Dim lblDel As DicomLabel
    Dim i As Integer
    If SelectedLabel Is Nothing Or SelectedImage Is Nothing Then Exit Sub
    If SelectedImage.Labels.IndexOf(SelectedLabel) <= G_INT_SYS_LABEL_COUNT Then Exit Sub
    If Mid(SelectedLabel.Tag, 1, 2) = "JD" Then '特别处理角度，因为一个角度有三个标注关联，先删除关联的文字标注。
        If Not SelectedLabel.TagObject Is Nothing Then
            If Not SelectedLabel.TagObject.TagObject Is Nothing Then
                If SelectedImage.Labels.IndexOf(SelectedLabel.TagObject.TagObject) <> 0 Then
                    SelectedImage.Labels.Remove SelectedImage.Labels.IndexOf(SelectedLabel.TagObject.TagObject)
                End If
            End If
        End If
        blnAngle = False
    ElseIf left(SelectedLabel.Tag, 3) = "VAS" Then '血管狭窄测量的八个标注,删除其中关联的6个标注
        If Not SelectedLabel.TagObject Is Nothing Then
            If Not SelectedLabel.TagObject.TagObject Is Nothing Then '处理第一个血管壁E1
                Set lblThis = SelectedLabel.TagObject.TagObject
                For i = 1 To 5
                    If Not lblThis.TagObject Is Nothing Then '处理血管垂直线
                        Set lblDel = lblThis
                        Set lblThis = lblThis.TagObject
                        If SelectedImage.Labels.IndexOf(lblDel) <> 0 Then SelectedImage.Labels.Remove SelectedImage.Labels.IndexOf(lblDel)
                    End If
                Next i
                If SelectedImage.Labels.IndexOf(lblThis) <> 0 Then SelectedImage.Labels.Remove SelectedImage.Labels.IndexOf(lblThis)
            End If
        End If
        intVasMeasure = 0
    ElseIf left(SelectedLabel.Tag, 3) = "CTR" Then '心胸比测量的4个标注，删除其中的2个标注
        If Not SelectedLabel.TagObject Is Nothing Then
            If Not SelectedLabel.TagObject.TagObject Is Nothing Then
                If Not SelectedLabel.TagObject.TagObject.TagObject Is Nothing Then
                    Set lblDel = SelectedLabel.TagObject.TagObject.TagObject
                    If SelectedImage.Labels.IndexOf(lblDel) <> 0 Then SelectedImage.Labels.Remove SelectedImage.Labels.IndexOf(lblDel)
                End If
                Set lblDel = SelectedLabel.TagObject.TagObject
                If SelectedImage.Labels.IndexOf(lblDel) <> 0 Then SelectedImage.Labels.Remove SelectedImage.Labels.IndexOf(lblDel)
            End If
        End If
        intCadioThoracicRatio = 0
    End If
    If Not SelectedLabel.TagObject Is Nothing Then  '删除关联的文字标注（对于角度是关联的第二根角度线）
        If SelectedImage.Labels.IndexOf(SelectedLabel.TagObject) <> 0 Then
            SelectedImage.Labels.Remove SelectedImage.Labels.IndexOf(SelectedLabel.TagObject)
        End If
    End If
    If SelectedImage.Labels.IndexOf(SelectedLabel) <> 0 Then    '最后删除被选中的标注本身
        SelectedImage.Labels.Remove SelectedImage.Labels.IndexOf(SelectedLabel)
    End If
    Set SelectedLabel = Nothing
    SubNoDispPeriod SelectedImage, Me
End Sub

Private Sub AddImgToFilm(img As DicomImage, thisViewer As DicomViewer, blnPrinted As Boolean)
'功能：把图像添加到打印预览窗口
'参数： Img -- 需要添加的图像
'       thisViewer -- 图像所在的Viewer，为了缩放图像使用
'       blnPrinted -- 记录图像是否已经被打印过
    
    On Error GoTo err
    
    If mfrmFilm Is Nothing Then Exit Sub
    
    Call mfrmFilm.ZLAddImage(img, blnPrinted, thisViewer.width / thisViewer.MultiColumns, thisViewer.height / thisViewer.MultiRows)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ZoomImage(intZoomDirection As Integer)
'缩放当前图像，提供给鼠标滚轮调用
'参数： intZoomDirection --- 鼠标滚轮方向。1放大，0缩小。
    Dim dblScale As Double
    
    '发生错误，不做任何提示
    On Error Resume Next
    Debug.Print intZoomDirection
    If SelectedImage Is Nothing Then Exit Sub
    If intSelectedSerial = 0 Then Exit Sub
    If Viewer.Count < intSelectedSerial Then Exit Sub
    
    If intZoomDirection = 1 Then
        dblScale = 1 + (lngZoomStep * 0.01)
    Else
        dblScale = 1 - (lngZoomStep * 0.01)
    End If
    
    subCenterZoom SelectedImage, Viewer(intSelectedSerial), SelectedImage.ActualZoom * dblScale
    '处理序列内图像同步
    Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
    Exit Sub
End Sub

Public Sub MouseWheel(intDirection As Integer)
'处理鼠标滚轮的事件
'参数：intDirection --- 鼠标滚轮的方向；1--向上；0--向下
    
    On Error Resume Next
    If intMouseWheelRoll = 0 Then       '翻页
        If Viewer(intSelectedSerial).Visible = False Then Exit Sub
        
        If intDirection = 1 Then '上翻一页
            If VScro(intSelectedSerial).Visible = False Then
                '全序列观片，切换到上一个序列
                If Button_miViewAllSeries = True Then
                    Call subAutoChangeSeries(intSelectedSerial, Val(ZLShowSeriesInfos(intSelectedSerial).SeriesNo), Viewer(intSelectedSerial).Tag, 1)
                End If
            Else
                If VScro(intSelectedSerial).Value = 1 Then
                    '如果向上翻页已经到头了，根据参数判断是否切换到前一个序列
                    If Button_miViewAllSeries = True Then   '全序列观片，切换到上一个序列
                        Call subAutoChangeSeries(intSelectedSerial, Val(ZLShowSeriesInfos(intSelectedSerial).SeriesNo), Viewer(intSelectedSerial).Tag, 1)
                    End If
                Else
                    If VScro(intSelectedSerial).Value - VScro(intSelectedSerial).LargeChange < 1 Then
                            '单序列观片，切换到第一个图
                            VScro(intSelectedSerial).Value = 1
                    Else
                        VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Value - VScro(intSelectedSerial).LargeChange
                    End If
                End If
            End If
        Else        '下翻一页
            If VScro(intSelectedSerial).Visible = False Then
                '全序列观片，切换到下一个序列
                If Button_miViewAllSeries = True Then
                    Call subAutoChangeSeries(intSelectedSerial, Val(ZLShowSeriesInfos(intSelectedSerial).SeriesNo), Viewer(intSelectedSerial).Tag, 2)
                End If
            Else
                If VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Max Then
                    '如果向下翻页已经到尾了，根据参数判断是否切换到后一个序列
                    If Button_miViewAllSeries = True Then       '全序列观片，切换到下一个序列
                        Call subAutoChangeSeries(intSelectedSerial, Val(ZLShowSeriesInfos(intSelectedSerial).SeriesNo), Viewer(intSelectedSerial).Tag, 2)
                    End If
                Else
                    If VScro(intSelectedSerial).Value + VScro(intSelectedSerial).LargeChange > VScro(intSelectedSerial).Max Then
                            '单序列观片，切换到最后一个图
                            VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Max
                    Else
                        VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Value + VScro(intSelectedSerial).LargeChange
                    End If
                End If
            End If
        End If
    ElseIf intMouseWheelRoll = 1 Then   '缩放
        Debug.Print "缩放"
        If intDirection = 1 Then    '放大
            Call ZoomImage(1)
        Else        '缩小
            Call ZoomImage(0)
        End If
    End If
End Sub

Private Sub subAutoChangeSeries(intVieweIndex As Integer, lngCurrentSeriesNo As Long, intCurrentIndex As Integer, intDirection As Integer)
'------------------------------------------------
'功能：自动切换序列
'参数： intVieweIndex --切换序列的Viewer的索引
'       lngCurrentSeriesNo -- Viewer中当前图像的序列号
'       intCurrentIndex --- 图像在ZLSeriesInfos中的序号
'       intDirection -- 切换序列的方向，1-向上切换序列；2-向下切换序列
'返回：无，直接切换Viewer中的序列
'------------------------------------------------
    Dim i As Integer
    Dim lngNextSeriesNo As Long
    Dim intNextIndex As Integer
    Dim lngMax As Long
    
    On Error Resume Next
    
    lngMax = 99999
    '在ZLSeriesInfos中查找下一个序列的索引
    '向上查找
    If intDirection = 1 Then
        lngNextSeriesNo = 0
        intNextIndex = 0
        For i = 1 To ZLSeriesInfos.Count
            '同一次检查的，才参与比较
            If ZLSeriesInfos(i).StudyUID = ZLSeriesInfos(intCurrentIndex).StudyUID Then
                If Val(ZLSeriesInfos(i).SeriesNo) < lngCurrentSeriesNo Then
                    If Val(ZLSeriesInfos(i).SeriesNo) > lngNextSeriesNo Then
                        lngNextSeriesNo = Val(ZLSeriesInfos(i).SeriesNo)
                        intNextIndex = i
                    End If
                End If
            End If
        Next i
    Else
        lngNextSeriesNo = lngMax
        intNextIndex = 0
        For i = 1 To ZLSeriesInfos.Count
            '同一次检查的，才参与比较
            If ZLSeriesInfos(i).StudyUID = ZLSeriesInfos(intCurrentIndex).StudyUID Then
                If Val(ZLSeriesInfos(i).SeriesNo) > lngCurrentSeriesNo Then
                    If Val(ZLSeriesInfos(i).SeriesNo) < lngNextSeriesNo Then
                        lngNextSeriesNo = Val(ZLSeriesInfos(i).SeriesNo)
                        intNextIndex = i
                    End If
                End If
            End If
        Next i
    End If
    '切换到下一个序列
    If intNextIndex <> 0 And intNextIndex <> lngMax Then
        '用新序列的图像代替viewer(index)中的图像
        Call funcSwapSeries(intVieweIndex, intNextIndex, IIf(intDirection = 1, True, False))
    End If
    
    Exit Sub
End Sub

Private Function funMPRslope(frmParent As Object) As Boolean
'------------------------------------------------
'功能： MPR斜面重建
'参数： frmParent -- 父窗体
'返回： True--成功，False---取消退出
'------------------------------------------------
    Dim mfrmSlop As New frmSlopeReconstruction
    
    On Error GoTo err
    
    Call mfrmSlop.zlShowMe(Me)
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function funViewerMPR(thisForm As frmViewer, Optional blnSilent As Boolean = False) As Boolean
'------------------------------------------------
'功能： 封装funMPR的过程，触发保存报告图的事件,对当前窗体中被选中的序列做矢冠状位重建，或者取消矢冠状位重建
'       thisForm.blnInMPR 说明窗体中是否有图像正在进行重建的过程中
'参数： thisForm -- 显示图像的窗体
'       blnSilent -- 静默结束MRP，不提示
'返回： True--成功，False---取消退出
'时间：2009-7
'------------------------------------------------
    funViewerMPR = funMPR(thisForm, blnSilent)
    '如果退出重建状态
    If blnInMPR = False Then
        '触发保存外部图像的事件，因为矢冠状位重建可能保存了结果图
        RaiseEvent AfterSaveOuterImage(PstrCheckUID)
    End If
End Function


Private Sub subFilmAddImages(intType As Integer, Optional intInterval As Integer = 1, Optional blnStartOdd As Boolean = True)
'------------------------------------------------
'功能： 向胶片预览窗口添加图像
'参数： intType -- 添加图像的方式，1-添加序列；2-添加当前图；3-添加所选图;4-间隔添加序列
'       intInterval -- 间隔打印的间隔数，当intAddType=4时使用
'       blnStartOdd -- True 奇数起；False 偶数起，当intAddType=4时使用
'返回： 无
'------------------------------------------------
    Dim i As Integer, j As Integer
    Dim iImageIndex As Integer
    Dim im As DicomImage
    Dim intStart As Integer
    
    On Error GoTo err
    
    If intType = 1 Or intType = 4 Then      '添加序列
        If intSelectedSerial > 0 Then
            If intType = 4 And blnStartOdd = False Then
                intStart = 2
            Else
                intStart = 1
            End If
            
            '如果是直接添加序列，保护intType，强制设置为1
            If intType = 1 Then
                intInterval = 1
            Else
                intInterval = intInterval + 1
            End If
            
            iImageIndex = 1
            For i = intStart To ZLShowSeriesInfos(intSelectedSerial).ImageInfos.Count Step intInterval
                Set im = Nothing
                '首先判断图像是否已经装载，如果已经装载，则找到这个图像并显示出来，如果没有装载，则装载该图像
                If ZLShowSeriesInfos(intSelectedSerial).ImageInfos(i).blnDisplayed = False Then
                    funcAddAImageA Viewer(intSelectedSerial), i
                End If
                
                '查找图像的索引
                While Viewer(intSelectedSerial).Images(iImageIndex).Tag < i And iImageIndex < Viewer(intSelectedSerial).Images.Count
                    iImageIndex = iImageIndex + 1
                Wend
                
                If iImageIndex <= Viewer(intSelectedSerial).Images.Count Then
                    If Viewer(intSelectedSerial).Images(iImageIndex).Tag = i Then
                        Set im = Viewer(intSelectedSerial).Images(iImageIndex)
                    End If
                End If
                
                If Not im Is Nothing Then
                    Call AddImgToFilm(im, Viewer(intSelectedSerial), ZLShowSeriesInfos(intSelectedSerial).ImageInfos(iImageIndex).blnPrinted)
                    DoEvents
                End If
            Next i
        End If
    ElseIf intType = 3 Then '添加所选图
        '把被选择的图像添加到打印预览窗口
        For i = 1 To ZLShowSeriesInfos.Count
            iImageIndex = 1
            For j = 1 To ZLShowSeriesInfos(i).ImageInfos.Count
                If ZLShowSeriesInfos(i).ImageInfos(j).blnSelected = True Then
                    Set im = Nothing
                    '首先判断图像是否已经装载，如果已经装载，则找到这个图像并显示出来，如果没有装载，则装载该图像
                    If ZLShowSeriesInfos(i).ImageInfos(j).blnDisplayed = False Then
                        funcAddAImageA Viewer(i), j
                    End If
                    
                    '查找图像的索引
                    While Viewer(i).Images(iImageIndex).Tag < j And iImageIndex < Viewer(i).Images.Count
                        iImageIndex = iImageIndex + 1
                    Wend
                    
                    If iImageIndex <= Viewer(i).Images.Count Then
                        If Viewer(i).Images(iImageIndex).Tag = j Then
                            Set im = Viewer(i).Images(iImageIndex)
                        End If
                    End If
                    
                    If Not im Is Nothing Then
                        Call AddImgToFilm(im, Viewer(i), ZLShowSeriesInfos(i).ImageInfos(iImageIndex).blnPrinted)
                        DoEvents
                    End If
                End If
            Next j
        Next i
    Else                    '添加当前图
        If Not SelectedImage Is Nothing And intSelectedSerial > 0 Then
            Call AddImgToFilm(SelectedImage, Viewer(intSelectedSerial), ZLShowSeriesInfos(intSelectedSerial).ImageInfos(SelectedImageIndex).blnPrinted)
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
