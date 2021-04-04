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
         Caption         =   "ż����"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   18
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton optPrintStart 
         Caption         =   "������"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdPrintInterval 
         Caption         =   "�����ӡ"
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
         Caption         =   "�����"
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
            Text            =   "�������"
            TextSave        =   "�������"
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
            Text            =   "��д"
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "��д"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   926
            MinWidth        =   706
            Text            =   "����"
            TextSave        =   "NUM"
            Object.ToolTipText     =   "����"
         EndProperty
      EndProperty
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
            Key             =   "��汨��ͼ"
            Object.Tag             =   "108"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1BDA
            Key             =   "ȫ���й�Ƭ"
            Object.Tag             =   "250"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2354
            Key             =   "��ѡL"
            Object.Tag             =   "365"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":256E
            Key             =   "��ѡR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2788
            Key             =   "����ͼ"
            Object.Tag             =   "249"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3162
            Key             =   "��"
            Object.Tag             =   "102"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":38DC
            Key             =   "��Ƭ��ӡ"
            Object.Tag             =   "406"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":4056
            Key             =   "���"
            Object.Tag             =   "419"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":47D0
            Key             =   "�Ŵ�"
            Object.Tag             =   "402"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":4F4A
            Key             =   "�ֶ�����L"
            Object.Tag             =   "314"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":56C4
            Key             =   "�ֶ�����R"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":5E3E
            Key             =   "����Ӧ����L"
            Object.Tag             =   "315"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":65B8
            Key             =   "����Ӧ����R"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6D32
            Key             =   "����L"
            Object.Tag             =   "309"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":74AC
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7C26
            Key             =   "����L"
            Object.Tag             =   "311"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":83A0
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":8B1A
            Key             =   "����L"
            Object.Tag             =   "308"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":9294
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":9A0E
            Key             =   "��ʾCTֵ"
            Object.Tag             =   "362"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":A188
            Key             =   "��Ӱ"
            Object.Tag             =   "401"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":A902
            Key             =   "����ȫѡ"
            Object.Tag             =   "303"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":B07C
            Key             =   "ͼ��ȫѡ"
            Object.Tag             =   "304"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":B7F6
            Key             =   "��һ����"
            Object.Tag             =   "247"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":BF70
            Key             =   "��һ������"
            Object.Tag             =   "248"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":C6EA
            Key             =   "�Ű�"
            Object.Tag             =   "201"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":CE64
            Key             =   "ȫ��"
            Object.Tag             =   "246"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":D5DE
            Key             =   "��ʾ����ͼ����Ϣ"
            Object.Tag             =   "203"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":DD58
            Key             =   "ȫ���ָ�"
            Object.Tag             =   "312"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":E4D2
            Key             =   "����۲�ģʽ"
            Object.Tag             =   "202"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":EC4C
            Key             =   "ͼ������"
            Object.Tag             =   "112"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":F3C6
            Key             =   "�񻯼���"
            Object.Tag             =   "329"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":FB40
            Key             =   "����ǿ"
            Object.Tag             =   "330"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":102BA
            Key             =   "���Ƽ���"
            Object.Tag             =   "333"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":10A34
            Key             =   "������ǿ"
            Object.Tag             =   "334"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":111AE
            Key             =   "ƽ������"
            Object.Tag             =   "331"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":11928
            Key             =   "ƽ����ǿ"
            Object.Tag             =   "332"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":120A2
            Key             =   "ͼ��ԭ"
            Object.Tag             =   "335"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1281C
            Key             =   "α��"
            Object.Tag             =   "405"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":12F96
            Key             =   "ȫ����λ��"
            Object.Tag             =   "318"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":13710
            Key             =   "��β��λ��"
            Object.Tag             =   "319"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":13E8A
            Key             =   "��ǰ��λ��"
            Object.Tag             =   "320"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":14604
            Key             =   "��ά���L"
            Object.Tag             =   "321"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":14D7E
            Key             =   "��ά���R"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":154F8
            Key             =   "ʸ��״�ؽ�"
            Object.Tag             =   "403"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":15C72
            Key             =   "ˮƽ��ת"
            Object.Tag             =   "323"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":163EC
            Key             =   "��ֱ��ת"
            Object.Tag             =   "324"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":16B66
            Key             =   "��ת90��"
            Object.Tag             =   "325"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":172E0
            Key             =   "��ת90��"
            Object.Tag             =   "326"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":17A5A
            Key             =   "����"
            Object.Tag             =   "327"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":181D4
            Key             =   "DSA"
            Object.Tag             =   "404"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1894E
            Key             =   "����L"
            Object.Tag             =   "337"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":190C8
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":19842
            Key             =   "��ͷL"
            Object.Tag             =   "338"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":19FBC
            Key             =   "��ͷR"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1A736
            Key             =   "��ԲL"
            Object.Tag             =   "339"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1AEB0
            Key             =   "��ԲR"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1B62A
            Key             =   "�Ƕ�L"
            Object.Tag             =   "340"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1BDA4
            Key             =   "�Ƕ�R"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1C51E
            Key             =   "����L"
            Object.Tag             =   "341"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1CC98
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1D412
            Key             =   "����L"
            Object.Tag             =   "342"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1DB8C
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1E306
            Key             =   "ֱ��L"
            Object.Tag             =   "343"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1EA80
            Key             =   "ֱ��R"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1F1FA
            Key             =   "����L"
            Object.Tag             =   "344"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":1F974
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":200EE
            Key             =   "������б�ע"
            Object.Tag             =   "347"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":20868
            Key             =   "Ѫ�ܲ���L"
            Object.Tag             =   "361"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":20FE2
            Key             =   "Ѫ�ܲ���R"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2175C
            Key             =   "У׼"
            Object.Tag             =   "346"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":21ED6
            Key             =   "����ֱ��ͼ"
            Object.Tag             =   "345"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":22650
            Key             =   "�ü�L"
            Object.Tag             =   "310"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":22DCA
            Key             =   "�ü�R"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":23544
            Key             =   "ͼ��ͬ��"
            Object.Tag             =   "307"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":23CBE
            Key             =   "����ͬ��"
            Object.Tag             =   "306"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":24438
            Key             =   "����ͼ��"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":24BB2
            Key             =   "�ֹ�����ͬ��"
            Object.Tag             =   "363"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2532C
            Key             =   "��������"
            Object.Tag             =   "364"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":25AA6
            Key             =   "���ر�R"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":261A0
            Key             =   "���ر�L"
            Object.Tag             =   "367"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2689A
            Key             =   "�˾�ģ��"
            Object.Tag             =   "32810"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":27014
            Key             =   "��ӡ����"
            Object.Tag             =   "40601"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2778E
            Key             =   "б���ؽ�"
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
            Key             =   "��汨��ͼ"
            Object.Tag             =   "108"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2823A
            Key             =   "ȫ���й�Ƭ"
            Object.Tag             =   "250"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":287D4
            Key             =   "��ѡL"
            Object.Tag             =   "365"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2892E
            Key             =   "��ѡR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":28A88
            Key             =   "����ͼ"
            Object.Tag             =   "249"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":28E22
            Key             =   "��"
            Object.Tag             =   "102"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":291BC
            Key             =   "��Ƭ��ӡ"
            Object.Tag             =   "406"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":29556
            Key             =   "���"
            Object.Tag             =   "419"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":298F0
            Key             =   "�Ŵ�"
            Object.Tag             =   "402"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":29C8A
            Key             =   "�ֶ�����L"
            Object.Tag             =   "314"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2A024
            Key             =   "�ֶ�����R"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2A3BE
            Key             =   "����Ӧ����L"
            Object.Tag             =   "315"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2A758
            Key             =   "����Ӧ����R"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2AAF2
            Key             =   "����L"
            Object.Tag             =   "309"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2AE8C
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2B226
            Key             =   "����L"
            Object.Tag             =   "311"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2B5C0
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2B95A
            Key             =   "����L"
            Object.Tag             =   "308"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2BCF4
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2C08E
            Key             =   "��ʾCTֵ"
            Object.Tag             =   "362"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2C428
            Key             =   "��Ӱ"
            Object.Tag             =   "401"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2C7C2
            Key             =   "����ȫѡ"
            Object.Tag             =   "303"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2CB5C
            Key             =   "ͼ��ȫѡ"
            Object.Tag             =   "304"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2CEF6
            Key             =   "��һ����"
            Object.Tag             =   "247"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2D290
            Key             =   "��һ������"
            Object.Tag             =   "248"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2D62A
            Key             =   "�Ű�"
            Object.Tag             =   "201"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2D9C4
            Key             =   "ȫ��"
            Object.Tag             =   "246"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2DD5E
            Key             =   "��ʾ����ͼ����Ϣ"
            Object.Tag             =   "203"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2E0F8
            Key             =   "ȫ���ָ�"
            Object.Tag             =   "312"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2E492
            Key             =   "����۲�ģʽ"
            Object.Tag             =   "202"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2E82C
            Key             =   "ͼ������"
            Object.Tag             =   "112"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2EBC6
            Key             =   "�񻯼���"
            Object.Tag             =   "329"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2EF60
            Key             =   "����ǿ"
            Object.Tag             =   "330"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2F2FA
            Key             =   "���Ƽ���"
            Object.Tag             =   "333"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2F694
            Key             =   "������ǿ"
            Object.Tag             =   "334"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2FA2E
            Key             =   "ƽ������"
            Object.Tag             =   "331"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":2FDC8
            Key             =   "ƽ����ǿ"
            Object.Tag             =   "332"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":30162
            Key             =   "ͼ��ԭ"
            Object.Tag             =   "335"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":304FC
            Key             =   "α��"
            Object.Tag             =   "405"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":30896
            Key             =   "ȫ����λ��"
            Object.Tag             =   "318"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":30C30
            Key             =   "��β��λ��"
            Object.Tag             =   "319"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":30FCA
            Key             =   "��ǰ��λ��"
            Object.Tag             =   "320"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":31364
            Key             =   "��ά���L"
            Object.Tag             =   "321"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":316FE
            Key             =   "��ά���R"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":31A98
            Key             =   "ʸ��״�ؽ�"
            Object.Tag             =   "403"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":31E32
            Key             =   "ˮƽ��ת"
            Object.Tag             =   "323"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":321CC
            Key             =   "��ֱ��ת"
            Object.Tag             =   "324"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":32566
            Key             =   "��ת90��"
            Object.Tag             =   "325"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":32900
            Key             =   "��ת90��"
            Object.Tag             =   "326"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":32C9A
            Key             =   "����"
            Object.Tag             =   "327"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":33034
            Key             =   "DSA"
            Object.Tag             =   "404"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":333CE
            Key             =   "����L"
            Object.Tag             =   "337"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":33768
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":33B02
            Key             =   "��ͷL"
            Object.Tag             =   "338"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":33E9C
            Key             =   "��ͷR"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":34236
            Key             =   "��ԲL"
            Object.Tag             =   "339"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":345D0
            Key             =   "��ԲR"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3496A
            Key             =   "�Ƕ�L"
            Object.Tag             =   "340"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":34D04
            Key             =   "�Ƕ�R"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3509E
            Key             =   "����L"
            Object.Tag             =   "341"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":35438
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":357D2
            Key             =   "����L"
            Object.Tag             =   "342"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":35B6C
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":35F06
            Key             =   "ֱ��L"
            Object.Tag             =   "343"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":362A0
            Key             =   "ֱ��R"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3663A
            Key             =   "����L"
            Object.Tag             =   "344"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":369D4
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":36D6E
            Key             =   "������б�ע"
            Object.Tag             =   "347"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":37108
            Key             =   "Ѫ�ܲ���L"
            Object.Tag             =   "361"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":374A2
            Key             =   "Ѫ�ܲ���R"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3783C
            Key             =   "У׼"
            Object.Tag             =   "346"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":37BD6
            Key             =   "����ֱ��ͼ"
            Object.Tag             =   "345"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":37F70
            Key             =   "�ü�L"
            Object.Tag             =   "310"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3830A
            Key             =   "�ü�R"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":386A4
            Key             =   "ͼ��ͬ��"
            Object.Tag             =   "307"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":38A3E
            Key             =   "����ͬ��"
            Object.Tag             =   "306"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":38DD8
            Key             =   "�ֹ�����ͬ��"
            Object.Tag             =   "363"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":39172
            Key             =   "��������"
            Object.Tag             =   "364"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3950C
            Key             =   "���ر�R"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":39AA6
            Key             =   "���ر�L"
            Object.Tag             =   "367"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3A040
            Key             =   "�˾�ģ��"
            Object.Tag             =   "32810"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3A3DA
            Key             =   "��ӡ����"
            Object.Tag             =   "40601"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3A774
            Key             =   "б���ؽ�"
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
            Key             =   "��汨��ͼ"
            Object.Tag             =   "108"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3B7A0
            Key             =   "ȫ���й�Ƭ"
            Object.Tag             =   "250"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3C47A
            Key             =   "��ѡL"
            Object.Tag             =   "365"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3C794
            Key             =   "��ѡR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3CAAE
            Key             =   "����ͼ"
            Object.Tag             =   "249"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3D788
            Key             =   "��"
            Object.Tag             =   "102"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3F272
            Key             =   "��Ƭ��ӡ"
            Object.Tag             =   "406"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":3FF4C
            Key             =   "���"
            Object.Tag             =   "419"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":40C26
            Key             =   "�Ŵ�"
            Object.Tag             =   "402"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":41900
            Key             =   "�ֶ�����L"
            Object.Tag             =   "314"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":425DA
            Key             =   "�ֶ�����R"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":432B4
            Key             =   "����Ӧ����L"
            Object.Tag             =   "315"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":43F8E
            Key             =   "����Ӧ����R"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":44C68
            Key             =   "����L"
            Object.Tag             =   "309"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":45942
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":4661C
            Key             =   "����L"
            Object.Tag             =   "311"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":46EF6
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":477D0
            Key             =   "����L"
            Object.Tag             =   "308"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":484AA
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":49184
            Key             =   "��ʾCTֵ"
            Object.Tag             =   "362"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":49E5E
            Key             =   "��Ӱ"
            Object.Tag             =   "401"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":4AB38
            Key             =   "����ȫѡ"
            Object.Tag             =   "303"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":4B812
            Key             =   "ͼ��ȫѡ"
            Object.Tag             =   "304"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":4C4EC
            Key             =   "��һ����"
            Object.Tag             =   "247"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":4F8DE
            Key             =   "��һ������"
            Object.Tag             =   "248"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":52CD0
            Key             =   "�Ű�"
            Object.Tag             =   "201"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":539AA
            Key             =   "ȫ��"
            Object.Tag             =   "246"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":55394
            Key             =   "��ʾ����ͼ����Ϣ"
            Object.Tag             =   "203"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":5606E
            Key             =   "ȫ���ָ�"
            Object.Tag             =   "312"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":56D48
            Key             =   "����۲�ģʽ"
            Object.Tag             =   "202"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":57A22
            Key             =   "ͼ������"
            Object.Tag             =   "112"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":586FC
            Key             =   "�񻯼���"
            Object.Tag             =   "329"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":593D6
            Key             =   "����ǿ"
            Object.Tag             =   "330"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":5A0B0
            Key             =   "���Ƽ���"
            Object.Tag             =   "333"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":5AD8A
            Key             =   "������ǿ"
            Object.Tag             =   "334"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":5BA64
            Key             =   "ƽ������"
            Object.Tag             =   "331"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":5EE56
            Key             =   "ƽ����ǿ"
            Object.Tag             =   "332"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":62248
            Key             =   "ͼ��ԭ"
            Object.Tag             =   "335"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":62F22
            Key             =   "α��"
            Object.Tag             =   "405"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":63BFC
            Key             =   "ȫ����λ��"
            Object.Tag             =   "318"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":648D6
            Key             =   "��β��λ��"
            Object.Tag             =   "319"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":655B0
            Key             =   "��ǰ��λ��"
            Object.Tag             =   "320"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6628A
            Key             =   "��ά���L"
            Object.Tag             =   "321"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":66F64
            Key             =   "��ά���R"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":67C3E
            Key             =   "ʸ��״�ؽ�"
            Object.Tag             =   "403"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":68918
            Key             =   "ˮƽ��ת"
            Object.Tag             =   "323"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":695F2
            Key             =   "��ֱ��ת"
            Object.Tag             =   "324"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6A2CC
            Key             =   "��ת90��"
            Object.Tag             =   "325"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6BDB6
            Key             =   "��ת90��"
            Object.Tag             =   "326"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6CA90
            Key             =   "����"
            Object.Tag             =   "327"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6D76A
            Key             =   "DSA"
            Object.Tag             =   "404"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6E444
            Key             =   "����L"
            Object.Tag             =   "337"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6F11E
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":71580
            Key             =   "��ͷL"
            Object.Tag             =   "338"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7225A
            Key             =   "��ͷR"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":72F34
            Key             =   "��ԲL"
            Object.Tag             =   "339"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":73C0E
            Key             =   "��ԲR"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":748E8
            Key             =   "�Ƕ�L"
            Object.Tag             =   "340"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":755C2
            Key             =   "�Ƕ�R"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7629C
            Key             =   "����L"
            Object.Tag             =   "341"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":76F76
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":77C50
            Key             =   "����L"
            Object.Tag             =   "342"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7892A
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":79604
            Key             =   "ֱ��L"
            Object.Tag             =   "343"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7A2DE
            Key             =   "ֱ��R"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7AFB8
            Key             =   "����L"
            Object.Tag             =   "344"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7BC92
            Key             =   "����R"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7C96C
            Key             =   "������б�ע"
            Object.Tag             =   "347"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":7D646
            Key             =   "Ѫ�ܲ���L"
            Object.Tag             =   "361"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":80A38
            Key             =   "Ѫ�ܲ���R"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":83E2A
            Key             =   "У׼"
            Object.Tag             =   "346"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":84B04
            Key             =   "����ֱ��ͼ"
            Object.Tag             =   "345"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":857DE
            Key             =   "�ü�L"
            Object.Tag             =   "310"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":864B8
            Key             =   "�ü�R"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":87192
            Key             =   "ͼ��ͬ��"
            Object.Tag             =   "307"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":88B7C
            Key             =   "����ͬ��"
            Object.Tag             =   "306"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":89856
            Key             =   "�ֹ�����ͬ��"
            Object.Tag             =   "363"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":8A530
            Key             =   "��������"
            Object.Tag             =   "364"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":8B20A
            Key             =   "���ر�R"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":8BAE4
            Key             =   "���ر�L"
            Object.Tag             =   "367"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":8C3BE
            Key             =   "�˾�ģ��"
            Object.Tag             =   "32810"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":8D098
            Key             =   "��ӡ����"
            Object.Tag             =   "40601"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":8DD72
            Key             =   "б���ؽ�"
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
''''MSFViewer�ṹ
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''0��=ϵ������ 0=��Դ��PACSͼ��������ϵĿ��Ա��������   1=��Դ���ļ��ϸ����ֵ�����,2=�������,3=����ʸ��״�ؽ�����������---�Ժ�����
''''1��=�Ƿ���ͼ��  as Boolean---�Ժ�����
''''2=���UID       as string  BMP��jpg�ļ�Ϊ��---�Ժ�����
''''3=��ǰѡ���ͼ���
''''4=��ǰѡ���ͼ���ڵڼ�֡
''''(OK)5=�����к�����ʾͼ����Ŀ(�������ڵ�ͼ�Ͷ�ͼ��ʾ�л���)
''''(OK)6=������������ʾͼ����Ŀ(�������ڵ�ͼ�Ͷ�ͼ��ʾ�л���)
''''7=�����е�ǰ��ʾ��һ��ͼ�����(�������ڵ�ͼ�Ͷ�ͼ��ʾ�л���)
''''8=�����е�ǰ��ʾѡ��ͼ�����(�������ڵ�ͼ�Ͷ�ͼ��ʾ�л���)
''''9=��������ͼ���Ƿ��Զ�ͬ��---�Ժ�����
''''10=��¼�����е�ǰѡ�е�LABEL���
''''11=��¼���������ڵĺ���λ��---�Ժ�����
''''12=��¼���������ڵ�����λ��---�Ժ�����
''''13=��¼��Ӧ����ʱViewerλ��,ʸ��״λ�ؽ���ʱ���ؽ�ͼ������Ӧ��X������ͼ��Viewer index---�Ժ�����
''''14=��¼��Ӧ����ʱViewerλ��,ʸ��״λ�ؽ���ʱ���ؽ�ͼ������Ӧ��Y������ͼ��Viewer index---�Ժ�����
''''15=��¼��ǰ�����Ƿ�ѡ�������Զ����ֹ�����ͬ��   ----�Ժ�����
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'-----------------------------------------------------
'�����¼�
'-----------------------------------------------------
Public Event AfterSaveReportImage(strCheckUID As String)
Public Event AfterSaveOuterImage(strCheckUID As String)
Public Event AfterSeriesChanged(strStudyUID As String, strSeriesUID As String)

''''''''''''''''''''''''''''''''''''''''''[Viewer����]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public intSelectedSerial As Integer                             ''''��ǰ����������(Viewer)
Public oldSelectedSerial As Integer                                ''''��¼��һ��ѡ�������

''''''''''''''''''''''''''[ͨ������������λ�����м����]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public intFactMoveX As Integer                                  '''''��¼��갴�º���ָ�����λ��
Public intFactMoveY As Integer                                  '''''��¼��갴������ָ�����λ��
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public intOldCountX, intOldCountY As Integer                    '''''��¼MPR֮ǰ��������в���������������в�����
Public intCountX, intCountY As Integer                          '''''����������Ų��������������Ų���
Public intDefaultCountX, intDefaultCountY As Integer            '''''��һ��ͼ����豸�����������ĺ���������Ų��������������Ų���
Public blnAutoCount                                             '''''��һ��ͼ����豸�����������ĺ���������Ų��������������Ų����Ƿ��Զ�����
Public isSelectAllSerial As Boolean                             ''''��ע�Ƿ�ѡ������������
Public isSelectAllImage As Boolean                              ''''��ע�Ƿ�ѡ��������ͼ��
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public intDblClickButton As Integer                             ''''��¼��갴��
Public SelectedImage As DicomImage                              ''''��ǰѡ���ͼ��
Public oldSelectedImageIndex As Integer                         ''''��¼�ϴ�ѡ���ͼ��INDEX
Public SelectedImageIndex As Integer                            ''''����ѡ���ͼ��INDEX

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim blnTextInput As Boolean                                     '''''��ʼ�������ֵı�ʶ
Dim blnTextInputM As Boolean                                    '''''��ʼ�޸����ֵı�ʶ
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public intClickImageIndex As Integer                            ''''��ǰ�����ͼ�����к�,0��ʶû�е���,��MOUSE����д,��˫��ʹ��
Dim lngBaseX As Long                                            ''''�����˵���¼λ��Ҳʹ��
Dim lngBaseY As Long                                            ''''�����˵���¼λ��Ҳʹ��
Dim lngBaseXX As Long
Dim lngBaseYY As Long
''''''''''''''''''''''''''''''''''''''''[���󲥷��ñ���]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public objStackOldImage As New DicomImage                       '''''''�������ż�¼��ǰͼ��
Public intStackIndex As Integer                                '''''''�������ż�¼��ǰ����ͼ�����
Public blnStackisFrame As Boolean                               '''''''��¼���ö�֡���Ż��ǵ���ѭ������
Public intStackCurrentlyImage As Integer                        ''''''''��¼��ǰ������ǰͼ���
Public intStackOffset As Integer                                ''''''''��¼��ʼ����ͼ�������ͼ���ƫ����
Public objStackImages As DicomImages                            ''''''''�������ż�¼�õ�ͼ��
Dim blnStackStart As Boolean                                    ''''''''��Ǵ���ʼ
Dim blnMouseStart As Boolean                                    ''''''''��¼����ʼ������һ���϶�����Ҫ��ʼ��ͼ��λ�õ�
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public SelectedLabel As DicomLabel                              ''''��ǰѡ��ı�ע
Public SelectedLabelT As DicomLabel                             ''''��ǰѡ��ı�ע��ǰѡ���ע��Ӧ����
Public isSelectedLabel As Boolean                               ''''�Ƿ�ѡ���˱�ע
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim blnAutoWL As Boolean                                        '''''��ʼ�Զ�����λ
Dim blnFrameSelectImage As Boolean                              '''''��ʼ��ѡͼ��
Dim LabelDrawing As Boolean                                     '''''��ʼ����ע�ı�ʶ
Dim blnAngle As Boolean                                         '''''��ʼ���Ƕȵı�ʶ
Dim intVasMeasure As Integer                                    '''''��ʼѪ����խ�����ı�ʶ��0-��ʾ����Ѫ����խ����״̬��1-��ʾ������Ѫ�ܲ�����2-��ʾ����խѪ�ܲ�����
Dim intCadioThoracicRatio As Integer                            '''''��ʼ���رȲ����ı�־��0-��ʾ�������رȲ���״̬��1-��ʾ�����������2-��ʾ������������
Dim oldFontSize As Integer                                      '''''��ʶ�����������ǰ��任ʹ��
Dim oldTextleft As Integer                                      '''''�м����
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim blnMoveLabel  As Boolean                                    '''''��ʼ�ƶ���ע�ı�ʶ
Dim blnReSizeLabel As Boolean                                   '''''��ʼ�ı��ע��״�ı�ʶ
Dim intReSizeIndex As Integer                                   '''''�м��������¼������
Public LngOldColor As Long                                      ''SubChangeColor����ʹ��
Public DLblOld As DicomLabel                                    ''SubChangeColor����ʹ��
Public dubCalibrateLength As Double                             '''''У׼����
Public blnForceRefresh As Boolean                               ''''�Ƿ�ǿ��ˢ��,��[���ó���ʹ��

''''''''''''''''''''''''''''''''''''''''''��ά���''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim obj3dImage() As DicomImage                                  ''''��Viewer��ǰͼ��ı���
Dim int3dImageIndex() As Integer                                ''''��Viewer��ǰͼ���INDEX����
Dim int3dCurrentlyImage() As Integer                            ''''��Viewer���Ͻ�ͼ���INDEX����
Dim blnIn3dCursor As Boolean                                    ''''�Ƿ������ά���״̬

'''''''''''''''''''''''''''''''''''''''''ͼ��ƴ��'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim fis As New frmImageSpelling
Public blnfis As Boolean                                        ''''�Ƿ������ƴ��״̬

'''''''''''''''''''''''''''''''''''''''''��Ƭ��ӡ'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public blnPrintFilm As Boolean                                  ''''�Ƿ�����˽�Ƭ��ӡ״̬
Public WithEvents mfrmFilm As frmFilm                                      ''''��Ƭ��ӡ����
Attribute mfrmFilm.VB_VarHelpID = -1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim intIntercept As Integer                                     ''''��¼��ʾCTֵʱ�����õĽؾ�
Dim intSlope As Integer                                         ''''��¼��ʾCTֵ�ǻ����õ�б��
Dim strInstanceUID As String                                    ''''��¼��ʾCTֵʱ��ͼ��ʵ��UID
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim blDicomDown As Boolean                                      ''''�Ƿ������
Public blnVscroInvoked As Boolean                               ''''��¼�Ƿ����ֹ�����ͬ�������Vscro�ı�
Public blnDefaultWW2 As Boolean                                 ''''��¼�Ƿ�ʹ����Ĭ�ϵĵڶ�������λ��֧��Ĭ��˫����
''''''''''''''''''''''''''''''''''''''''ʸ��״λ�ؽ�''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public blnInMPR As Boolean                                      ''''�Ƿ����ڽ���ʸ��״λ�ؽ�

'��Ƭվע��
Public blnLogined As Boolean                                    ''''�Ƿ�ע��ɹ���True�ɹ���Falseʧ�ܡ�
Private mstr����ʱ�� As String                                   ''''ע��ɹ��������ʱ��

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''����ʦ��ҽ��վ���õı��������ʾ
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
    
    '�������õļ����Ӵ�ӡ����
    Call funFilm(Me, False, 4, intInterval, blnStartOdd)
    '�رյ����˵�
    ComToolBar.ClosePopups
End Sub

Private Sub ComToolBar_ControlSelected(ByVal control As XtremeCommandBars.ICommandBarControl)
    '��ʾ��������ʾ��Ϣ
    Me.sbStatusBar.Panels(2).Text = StatusBarTip(control)
End Sub

Private Sub ComToolBar_Customization(ByVal Options As XtremeCommandBars.ICustomizeOptions)
    '���ÿ��������Զ���İ�ť
    Dim Controls As CommandBarControls
    Dim cbrControl As CommandBarControl
    Dim ControlPopup As CommandBarPopup
    
    '���ؼ���ҳ�����ʾ
    Options.ShowKeyboardPage = False
    Options.ShowOptionsPage = False
    
    Options.CustomIcons.RemoveAll
'    Options.ContextMenu.Title
    
    '���֧���Զ������õİ�ť
    Set Controls = ComToolBar.DesignerControls
    
    Controls.DeleteAll
    
    If (Controls.Count = 0) Then
        '��������
        Set cbrControl = Controls.Add(xtpControlButton, ID_File_SAveASReport, "���汨��ͼ")
        cbrControl.Category = "��������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_File_Open, "��")
        cbrControl.Category = "��������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Tool_FilmPrint, "��Ƭ���")
        cbrControl.Category = "��������"
        Set ControlPopup = Controls.Add(xtpControlSplitButtonPopup, ID_Tool_Film_AddSeries, "��ӡ����")
        ControlPopup.Category = "��������"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddSeries, "��ӡ����"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddImage, "��ӡͼ��"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_Tool_Film_AddSelected, "��ӡ��ѡͼ"
            Set ControlPopup = ControlPopup.CommandBar.Controls.Add(xtpControlButtonPopup, ID_Tool_Film_AddInterval, "�����ӡ")
            ControlPopup.CommandBar.SetPopupToolBar True
            ControlPopup.CommandBar.Title = "�����ӡ"
            ControlPopup.ToolTipText = "�����ӡ��ǰ����"
        
        'ͼ�����
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Eddy_LeftRight, "ˮƽ����")
        cbrControl.Category = "ͼ�����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Eddy_TopButton, "��ֱ����")
        cbrControl.Category = "ͼ�����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Eddy_Left90, "��ת90��")
        cbrControl.Category = "ͼ�����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Eddy_Right90, "��ת90��")
        cbrControl.Category = "ͼ�����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_ReverseVideo, "����")
        cbrControl.Category = "ͼ�����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Tool_NumberMinusShadow, "DSA���ּ�Ӱ")
        cbrControl.Category = "ͼ�����"
        
        '����������
        Set cbrControl = Controls.Add(xtpControlButton, ID_Tool_NothinMouseState, "���")
        cbrControl.Category = "����������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_ACtive_Mouse_Value, "���������ʾCTֵ")
        cbrControl.Category = "����������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_Text, "����")
        cbrControl.Category = "����������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_Arrowhead, "��ͷ")
        cbrControl.Category = "����������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_Ellipse, "��Բ")
        cbrControl.Category = "����������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_Angle, "�Ƕ�")
        cbrControl.Category = "����������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_Curve, "����")
        cbrControl.Category = "����������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_Area, "����")
        cbrControl.Category = "����������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_BeeLine, "ֱ��")
        cbrControl.Category = "����������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_Rect, "����")
        cbrControl.Category = "����������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_VasMeasure, "Ѫ����խ����")
        cbrControl.Category = "����������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_CadioThoracicRatio, "���رȲ���")
        cbrControl.Category = "����������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_ClearLbale, "�����ע")
        cbrControl.Category = "����������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Lable_AdjustLine, "У׼")
        cbrControl.Category = "����������"
        
        '��ƽ�湤����
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_PointingLine_ALL, "��ʾ���ж�λ��")
        cbrControl.Category = "��ƽ�湤����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_PointingLine_FirstLast, "��ʾ��β��λ��")
        cbrControl.Category = "��ƽ�湤����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_PointingLine_Now, "��ʾ��ǰ��λ��")
        cbrControl.Category = "��ƽ�湤����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_PointingLine_3DLine, "��ά���")
        cbrControl.Category = "��ƽ�湤����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Tool_ArrowyCoronaryReset, "ʸ/��״λ�ؽ�")
        cbrControl.Category = "��ƽ�湤����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Tool_SlopeReconstruction, "б���ؽ�")
        cbrControl.Category = "��ƽ�湤����"
        
        '�������
        Set cbrControl = Controls.Add(xtpControlButton, ID_ACtive_FrameSelectImage, "��ѡͼ��")
        cbrControl.Category = "�������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Also_Photo, "ͼ���ʽͬ��")
        cbrControl.Category = "�������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Also_Serial, "���м�ͼ��λ��ͬ��")
        cbrControl.Category = "�������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Also_ManualSerial, "�ֹ�����ͬ��")
        cbrControl.Category = "�������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Also_LockSerial, "����/��������")
        cbrControl.Category = "�������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_View_ShowMiniSeries, "��ʾ��������ͼ")
        cbrControl.Category = "�������"
        Set cbrControl = Controls.Add(xtpControlButton, ID_View_ViewAllSeries, "ȫ���й�Ƭ")
        cbrControl.Category = "�������"
        
        'ͨ�ù�����
        Set cbrControl = Controls.Add(xtpControlButton, ID_Tool_Magnifier, "�Ŵ�")
        cbrControl.Category = "ͨ�ù�����"
        
        '�ֿص�����������̬�Ӳ˵�����ʱ�Ȳ�֧���Ӳ˵�
'        Set ControlPopup = Controls.Add(xtpControlSplitButtonPopup, ID_Active_AdjustWindow_HandAdjustWindow, "�ֿص���")
'        ControlPopup.Category = "ͨ�ù�����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_AdjustWindow_HandAdjustWindow, "�ֿص���")
        cbrControl.Category = "ͨ�ù�����"
        
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Cruise, "����")
        cbrControl.Category = "ͨ�ù�����"
        Set ControlPopup = Controls.Add(xtpControlSplitButtonPopup, ID_Active_Zoom, "����")
        ControlPopup.Category = "ͨ�ù�����"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_AutoShow, "����Ӧ"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_50%, "50%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_100%, "100%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_150%, "150%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_200%, "200%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_250%, "250%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_300%, "300%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_showScale_400%, "400%"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ShowScale_Custom, "�Զ���(&A)"
        Set ControlPopup = Controls.Add(xtpControlSplitButtonPopup, ID_Active_Shuttle, "����")
        ControlPopup.Category = "ͨ�ù�����"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_PhotoNumber, "ͼ���"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_BedASC, "��λ����"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_BedDESC, "��λ����"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_CollectionTime, "�ɼ�ʱ��"
            ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_View_PhotoSerial_PhotoTime, "ͼ��ʱ��"

        Set cbrControl = Controls.Add(xtpControlButton, ID_Tool_Movie, "��Ӱ����")
        cbrControl.Category = "ͨ�ù�����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Select_SelectAllSerial, "ѡ����������")
        cbrControl.Category = "ͨ�ù�����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Acitve_Select_SelectAllPhoto, "ѡ�����������е�ͼ��")
        cbrControl.Category = "ͨ�ù�����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_View_UpSeries, "��һ����")
        cbrControl.Category = "ͨ�ù�����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_View_DownSeries, "��һ����")
        cbrControl.Category = "ͨ�ù�����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_View_Typeset, "�������")
        cbrControl.Category = "ͨ�ù�����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_View_FullScreen, "ȫ����ʾ")
        cbrControl.Category = "ͨ�ù�����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_View_PropertyShow, "��/��������Ϣ")
        cbrControl.Category = "ͨ�ù�����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_ReSetAll, "ȫ���ָ�")
        cbrControl.Category = "ͨ�ù�����"
        Set cbrControl = Controls.Add(xtpControlButton, ID_View_OneBrowse, "���/�۲�ģʽ")
        cbrControl.Category = "ͨ�ù�����"
        
        'ͼ����ǿ
        
        '�˾�ģ���Ƕ�̬���ӵĲ˵����Ȳ�֧��
'        Set ControlPopup = Controls.Add(xtpControlSplitButtonPopup, ID_Active_SieveLens_Model, "�˾�ģ��")
'        ControlPopup.Category = "ͼ����ǿ"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_SieveLens_LancetMinus, "��Ե��ǿǿ�ȼ���")
        cbrControl.Category = "ͼ����ǿ"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_SieveLens_LancetAdd, "��Ե��ǿǿ������")
        cbrControl.Category = "ͼ����ǿ"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Sievelens_LeftMoveMinus, "��Ե��ǿ���ȼ���")
        cbrControl.Category = "ͼ����ǿ"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Sievelens_LeftMoveAdd, "��Ե��ǿ��������")
        cbrControl.Category = "ͼ����ǿ"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_SieveLens_FlatnessMinus, "ƽ������")
        cbrControl.Category = "ͼ����ǿ"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_SieveLens_FlatnessAdd, "ƽ������")
        cbrControl.Category = "ͼ����ǿ"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Active_Sievelens_PhotoReset, "ͼ��ԭ")
        cbrControl.Category = "ͼ����ǿ"
        Set cbrControl = Controls.Add(xtpControlButton, ID_Tool_BogusColour, "α��")
        cbrControl.Category = "ͼ����ǿ"
        
    End If
End Sub

Private Sub ComToolBar_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim CmdControl As CommandBarControl
    Dim i As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    subMouseRLset control                          ''''����������Ҽ��ֲ�
    subMnuImageSort control.Id, Me                     ''''����ʽ����
    '''''''''''''''''''''''''''''[���ܼ����ô���λ����]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    
    ''''''''''''''''''''''''''''''''''[�������˾�ģ��˵�]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If control.Id >= ID_Active_SieveLens_Model + 1 And control.Id <= ID_Active_SieveLens_Model + 40 Then
        Call subFunctionFilter(control, Me)
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''[�������˵���������]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If control.Id >= 500 And control.Id <= 800 And control.Category <> "" Then
        '�������Viewer�ڷŵ�λ��
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
        ''''''''''''''''''''''''''�ļ��˵�'''''''''''''''''''''''''''''''''''
        Case ID_File_Open                                                               '���ļ�
            subOpenFiles Me
            
        Case ID_File_Close                                                              '�ر�����
            subCloseSeries
            
        Case ID_File_DelAllPhoto                                                        'ɾ������ͼ��
            subKillPicture
            
        Case ID_File_DelReport                                                          'ɾ������ͼ��
            subDelRepImg
            RaiseEvent AfterSaveReportImage(PstrCheckUID)
            
        Case ID_File_SaveFile                                                           '�����ļ�
            Set frmSave.f = Me: frmSave.Show 1, Me
            
        Case ID_File_SaveASFile                                                         '����ļ�
            Set FrmSaveAs.f = Me: FrmSaveAs.Show 1, Me
        
        Case ID_File_SaveToCD                                                           '����CD
            '�������
            Set frmCreateCD.f = Me: frmCreateCD.Show 1, Me
            
        Case ID_File_SAveASReport                                                       '��汨��ͼ��
            subOutPutRptImg
            RaiseEvent AfterSaveReportImage(PstrCheckUID)
            
        Case ID_File_Send_GetHost                                                       '��������
            Set frmSendImage.f = Me: frmSendImage.Show 1, Me
            
        Case ID_File_Send_OutPowerPoint                                                 '�����PowerPoint
            subOutputToPowerPoint Me
            
        Case ID_File_OpenDicomDir                                                       '��DICOMDIR
            Set frmOpenDicomDir.f = Me
            frmOpenDicomDir.Show 1, Me
            
        Case ID_File_PhotoProperty                                                      'ͼ������
            If SelectedImage Is Nothing Then Exit Sub
            Set FrmSfyInfo.img = SelectedImage
            FrmSfyInfo.Show 1, Me
            
        Case ID_File_Exit                                                               '�˳�
            Unload Me
            
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''��ͼ''''''''''''''''''''''''''''''''''''''''
        Case ID_View_UpSeries                                                           '��һ����
            subChangeASeries 2
            
        Case ID_View_DownSeries                                                         '��һ����
            subChangeASeries 1
        
        Case ID_View_Typeset                                                            '���氲��
            Dim fLayout As New frmSerialLayoutSetup
            fLayout.zlShowMe Me
            
        Case ID_View_OneBrowse                                                          '����۲�ģʽ
            Call subLookOrBrowsSwitch(Me)
            
        Case ID_View_PropertyShow                                                       'ͼ���ϲ�����Ϣ��ʾ
            Dim v As DicomViewer
            Button_miDispPatientInfo = Not Button_miDispPatientInfo
            ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_PropertyShow).Checked = Button_miDispPatientInfo
            ComToolBar.RecalcLayout
            For Each v In Viewer
            If v.Index <> 0 Then
                subDisplayPatientInfo v
            End If
        Next
            
        Case ID_View_LableShow                                                          '��ע��ʾ
            subDispLabelInfo Me
            
        Case ID_View_ShowOverlay                                                        '��ʾOverlay
            control.Checked = Not control.Checked
            Button_miShowOverlay = control.Checked
            Call ShowOverlay(Me)
            
        Case ID_View_ShowMiniSeries                                                     '��ʾ��������ͼ
            ' ����˵����ڲ�����
            Button_miShowMiniSeries = Not Button_miShowMiniSeries
            Me.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_ShowMiniSeries, , True).Checked = Button_miShowMiniSeries
            Me.ComToolBar.Item(ToolBar_Object).FindControl(, ID_View_ShowMiniSeries, , True).Checked = Button_miShowMiniSeries
            Call subShowMiniImages(Me)
        
        Case ID_View_ViewAllSeries                                                      'ȫ���й�Ƭ
            Button_miViewAllSeries = Not Button_miViewAllSeries
            Me.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_ViewAllSeries, , True).Checked = Button_miViewAllSeries
            Me.ComToolBar.Item(ToolBar_Object).FindControl(, ID_View_ViewAllSeries, , True).Checked = Button_miViewAllSeries
            
        Case ID_View_ShowScale_AutoShow                                                 '����Ӧ
            subShowScale ID_View_ShowScale_AutoShow
            If Not SelectedImage Is Nothing Then               '��֤ͼ�����
                SelectedImage.StretchToFit = True
                Viewer(intSelectedSerial).Refresh
                '����������ͼ��ͬ��
                Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
            End If
            
        Case ID_View_ShowScale_50%                                                      '50%
            subShowScale ID_View_ShowScale_50%
            subCenterZoom SelectedImage, Viewer(intSelectedSerial), 0.5
            '����������ͼ��ͬ��
            Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
            
        Case ID_View_ShowScale_100%                                                     '100%
            subShowScale ID_View_ShowScale_100%
            subCenterZoom SelectedImage, Viewer(intSelectedSerial), 1
            '����������ͼ��ͬ��
            Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
        
        Case ID_View_showScale_150%                                                     '150%
            subShowScale ID_View_showScale_150%
            subCenterZoom SelectedImage, Viewer(intSelectedSerial), 1.5
            '����������ͼ��ͬ��
            Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
            
        Case ID_View_ShowScale_200%                                                     '200%
            subShowScale ID_View_ShowScale_200%
            subCenterZoom SelectedImage, Viewer(intSelectedSerial), 2
            '����������ͼ��ͬ��
            Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
            
        Case ID_View_showScale_250%                                                     '250%
            subShowScale ID_View_showScale_250%
            subCenterZoom SelectedImage, Viewer(intSelectedSerial), 2.5
            '����������ͼ��ͬ��
            Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
            
        Case ID_View_showScale_300%                                                     '300%
            subShowScale ID_View_showScale_300%
            subCenterZoom SelectedImage, Viewer(intSelectedSerial), 3
            '����������ͼ��ͬ��
            Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
            
        Case ID_View_showScale_400%                                                     '400%
            subShowScale ID_View_showScale_400%
            subCenterZoom SelectedImage, Viewer(intSelectedSerial), 4
            '����������ͼ��ͬ��
            Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
            
        Case ID_View_ShowScale_Custom                                                   '�Զ���
            subShowScale ID_View_ShowScale_Custom
            If Not SelectedImage Is Nothing Then
                frmZoomCustom.sRatio = SelectedImage.Zoom
                frmZoomCustom.Show 1, Me
                If frmZoomCustom.bApply Then
                    subCenterZoom SelectedImage, Viewer(intSelectedSerial), IIf(frmZoomCustom.sRatio = 0, 1, frmZoomCustom.sRatio)
                    '����������ͼ��ͬ��
                    Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
                End If
            End If
            
        Case ID_View_FullScreen                                                         'ȫ����ʾ
            subFullScreen Me
            
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''����''''''''''''''''''''''''''''''''''''
        Case ID_Active_Select_OneSelect                                                 '����ѡ��
            ZLShowSeriesInfos(intSelectedSerial).ImageInfos(SelectedImageIndex).blnSelected = Not ZLShowSeriesInfos(intSelectedSerial).ImageInfos(SelectedImageIndex).blnSelected
            subDispframe Me, Viewer(intSelectedSerial)
            Viewer(intSelectedSerial).Refresh
        Case ID_Active_Select_SelectAllSerial                                           'ѡ����������
            subSelectAllSerial Me
        Case ID_Acitve_Select_SelectAllPhoto                                            'ѡ������ͼ��
            subSelectAllIMage Me
            
        Case ID_Active_Also_Serial                                                      '����ͬ��
            control.Checked = Not control.Checked
            ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Also_Serial, , True).Checked = control.Checked
            ComToolBar.Item(ToolBar_Object).FindControl(, ID_Active_Also_Serial, , True).Checked = control.Checked
            Button_miSerialPlaceInPhase = control.Checked
            
            '�Զ�ѡ��ȫ������
            isSelectAllSerial = Not Button_miSerialPlaceInPhase
            subSelectAllSerial Me
            
            '��������Զ�����ͬ������ر��ֹ�����ͬ������ά���
            If control.Checked = True Then
                '�ر���ά���
                Set CmdControl = ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_3DLine)
                CmdControl.Checked = False
                '�ر��ֹ�����ͬ��
                ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Also_ManualSerial, , True).Checked = False
                ComToolBar.Item(ToolBar_Object).FindControl(, ID_Active_Also_ManualSerial, , True).Checked = False
                Button_miSerialManualSyn = False
            End If
            
        Case ID_Active_Also_ManualSerial                                                '�ֹ�����ͬ��
            '����ʽ��ť,������ͬ���ǻ����
            Button_miSerialManualSyn = Not Button_miSerialManualSyn
            control.Checked = Button_miSerialManualSyn
            ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Also_ManualSerial, , True).Checked = Button_miSerialManualSyn
            ComToolBar.Item(ToolBar_Object).FindControl(, ID_Active_Also_ManualSerial, , True).Checked = Button_miSerialManualSyn
            
            '�Զ�ѡ��ȫ������
            isSelectAllSerial = Not Button_miSerialManualSyn
            subSelectAllSerial Me
            
            '������ֹ�����ͬ������ر��Զ�����ͬ������ά���
            If control.Checked = True Then
                '�ر���ά����״̬
                Set CmdControl = ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_3DLine)
                CmdControl.Checked = False
                '�ر��Զ�����ͬ��
                ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Also_Serial, , True).Checked = False
                ComToolBar.Item(ToolBar_Object).FindControl(, ID_Active_Also_Serial, , True).Checked = False
                Button_miSerialPlaceInPhase = False
            End If
                
        Case ID_Active_Also_LockSerial                                                  '��������
            If intSelectedSerial > 0 And intSelectedSerial < Viewer.Count Then
                ZLShowSeriesInfos(intSelectedSerial).Selected = Not ZLShowSeriesInfos(intSelectedSerial).Selected
                subDispframe Me, Viewer(intSelectedSerial)
                Viewer(intSelectedSerial).Refresh
            End If
            
        Case ID_Active_Also_Photo                                                       'ͼ��ͬ��
            'ͼ������ͬ����Ϊһ���������ܽ�������
            Button_miImageInPhase = Not Button_miImageInPhase
            ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Also_Photo, , True).Checked = Button_miImageInPhase
            ComToolBar.Item(ToolBar_Object).FindControl(, ID_Active_Also_Photo, , True).Checked = Button_miImageInPhase
        Case ID_Tool_NothinMouseState
            subSelectLeftorRightBouttom 1, control.Id
            subSelectLeftorRightBouttom 2, control.Id
        Case ID_Active_Shuttle                                                          '����
            subSelectLeftorRightBouttom cMouseUsage("101").lngMouseKey, control.Id
            Button_miStack = True
        Case ID_ACtive_Mouse_Value                                                      '��ʾCTֵ
            control.Checked = Not control.Checked
            ComToolBar.Item(ToolBar_Menu).FindControl(, ID_ACtive_Mouse_Value, , True).Checked = control.Checked
            ComToolBar.Item(ToolBar_Scale).FindControl(, ID_ACtive_Mouse_Value, , True).Checked = control.Checked
            Button_miMouseShowValue = control.Checked
        Case ID_Active_Cruise                                                           '����
            subSelectLeftorRightBouttom cMouseUsage("103").lngMouseKey, control.Id
            Button_miCruise = True
            
        Case ID_Active_Cut                                                              '�ü�
            If SelectedImage Is Nothing Then Exit Sub
            Button_miCutOut = Not Button_miCutOut
            ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Cut, , True).Checked = Button_miCutOut
            subSelectLeftorRightBouttom cMouseUsage("201").lngMouseKey, control.Id
            subCutOut Me
            
        Case ID_ACtive_FrameSelectImage                                                 '��ѡͼ��
            If SelectedImage Is Nothing Then Exit Sub
            Button_miFrameSelectImage = Not Button_miFrameSelectImage
            ComToolBar.Item(ToolBar_Object).FindControl(, ID_ACtive_FrameSelectImage, , True).Checked = Button_miFrameSelectImage
            ComToolBar.Item(ToolBar_Menu).FindControl(, ID_ACtive_FrameSelectImage, , True).Checked = Button_miFrameSelectImage
            '������ͬ����λ�İ�ť״̬
            subSelectLeftorRightBouttom cMouseUsage("201").lngMouseKey, control.Id
            
        Case ID_ACtive_SaveInReport                                                     '�����ѡ��ͼ����뱨��ͼ
            If SelectedImage Is Nothing Or SelectedLabel Is Nothing Or SelectedLabel.LabelType <> doLabelRectangle Then Exit Sub
            SaveFrameSelectImageIntoReport SelectedImage, SelectedLabel
            RaiseEvent AfterSaveReportImage(PstrCheckUID)
            
        Case ID_Active_Zoom                                                             '����
            subSelectLeftorRightBouttom cMouseUsage("104").lngMouseKey, control.Id
            Button_miZoom = True
            
        Case ID_Active_ReSetAll                                                         '�ָ�����
            If Not SelectedImage Is Nothing Then
                SelectedImage.Mask = 0
                SelectedImage.SetDefaultWindows
                SelectedImage.FlipState = doFlipNormal
                SelectedImage.RotateState = doRotateNormal
                SelectedImage.ScrollX = 0
                SelectedImage.ScrollY = 0
                SelectedImage.StretchToFit = True
                '����������ͼ��ͬ��
                Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_All)
                '����˵��͹�����
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
            
        Case ID_Active_AdjustWindow_HandAdjustWindow                                    '�ֶ�����
            subSelectLeftorRightBouttom cMouseUsage("102").lngMouseKey, control.Id
            Button_miWidthLevel = True
            
        Case ID_Active_AdjustWindow_AutoAdjustWindow                                    '����Ӧ����
            subSelectLeftorRightBouttom cMouseUsage("105").lngMouseKey, control.Id
            Button_miAutoWidthLevel = True
            
        Case ID_Active_PointingLine_ALL                                                 '���ж�λ��
            If blnSelectedImageIfColor = False Then
                subCurrentCheck control, Me
            End If
            
        Case ID_Active_PointingLine_FirstLast                                           '��λ��λ��
            If blnSelectedImageIfColor = False Then
                subCurrentCheck control, Me
            End If
            
        Case ID_Active_PointingLine_Now                                                 '��ǰ��λ��
            If blnSelectedImageIfColor = False Then
                subCurrentCheck control, Me
            End If
            
        Case ID_Active_PointingLine_3DLine                                              '3D���
            subSelectLeftorRightBouttom cMouseUsage("106").lngMouseKey, control.Id
            
        Case ID_Active_Eddy_LeftRight                                                   '������ת
            subManipulation "FlipHorizontal", Me
            
        Case ID_Active_Eddy_TopButton                                                   '��ֱ��ת
            subManipulation "FlipVertical", Me
            
        Case ID_Active_Eddy_Left90                                                      '����90
            subManipulation "RotateAnticlockwise", Me
            
        Case ID_Active_Eddy_Right90                                                     '����90
            subManipulation "RotateClockwise", Me
            
        Case ID_Active_ReverseVideo                                                     '����
            control.Checked = Not control.Checked
            ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_ReverseVideo, , True).Checked = control.Checked
            ComToolBar.Item(ToolBar_Photo).FindControl(, ID_Active_ReverseVideo, , True).Checked = control.Checked
            subManipulation "Invert", Me
                    
        Case ID_Active_SieveLens_LancetMinus                                            '��Ե��ǿǿ�ȼ���
            zl9ComLib.zlCommFun.ShowFlash "���ڴ���ͼ����ȴ���", Me
            SubImageUnsharp "miUnSharpEnhancementDown", Me
            zl9ComLib.zlCommFun.StopFlash
            
        Case ID_Active_SieveLens_LancetAdd                                              '��Ե��ǿǿ������
            zl9ComLib.zlCommFun.ShowFlash "���ڴ���ͼ����ȴ���", Me
            zl9ComLib.zlCommFun.ShowFlash
            SubImageUnsharp "miUnSharpEnhancementUp", Me
            zl9ComLib.zlCommFun.StopFlash
            
        Case ID_Active_SieveLens_FlatnessMinus                                          'ƽ������
            zl9ComLib.zlCommFun.ShowFlash "���ڴ���ͼ����ȴ���", Me
            zl9ComLib.zlCommFun.ShowFlash
            SubImageUnsharp "miFilterLengthDown", Me
            zl9ComLib.zlCommFun.StopFlash
            
        Case ID_Active_SieveLens_FlatnessAdd                                            'ƽ������
            zl9ComLib.zlCommFun.ShowFlash "���ڴ���ͼ����ȴ���", Me
            zl9ComLib.zlCommFun.ShowFlash
            SubImageUnsharp "miFilterLengthUp", Me
            zl9ComLib.zlCommFun.StopFlash
            
        Case ID_Active_Sievelens_LeftMoveMinus                                          '��Ե��ǿ���ȼ���
            zl9ComLib.zlCommFun.ShowFlash "���ڴ���ͼ����ȴ���", Me
            zl9ComLib.zlCommFun.ShowFlash
            SubImageUnsharp "miUnSharpLengthDown", Me
            zl9ComLib.zlCommFun.StopFlash
            
        Case ID_Active_Sievelens_LeftMoveAdd                                            '��Ե��ǿ��������
            zl9ComLib.zlCommFun.ShowFlash "���ڴ���ͼ����ȴ���", Me
            zl9ComLib.zlCommFun.ShowFlash
            SubImageUnsharp "miUnSharpLengthUp", Me
            zl9ComLib.zlCommFun.StopFlash
            
        Case ID_Active_Sievelens_PhotoReset                                             'ͼ��ԭ
            SubImageUnsharp "miRestore", Me
            
        Case ID_Active_Lable_Text                                                       '����
            subSelectLeftorRightBouttom cMouseUsage("8").lngMouseKey, control.Id
            Button_miLabeltext = True
            
        Case ID_Active_Lable_Arrowhead                                                  '��ͷ
            subSelectLeftorRightBouttom cMouseUsage("4").lngMouseKey, control.Id
            Button_miLabelArrowhead = True
         
        Case ID_Active_Lable_Ellipse                                                    '��Բ
            subSelectLeftorRightBouttom cMouseUsage("3").lngMouseKey, control.Id
            Button_miLabelEllipse = True
        
        Case ID_Active_Lable_Angle                                                      '�Ƕ�
            subSelectLeftorRightBouttom cMouseUsage("7").lngMouseKey, control.Id
            Button_miLabelAngle = True
        
        Case ID_Active_Lable_Curve                                                      '����
            subSelectLeftorRightBouttom cMouseUsage("6").lngMouseKey, control.Id
            Button_miLabelPolyLine = True
            
        Case ID_Active_Lable_Area                                                       '����
            subSelectLeftorRightBouttom cMouseUsage("5").lngMouseKey, control.Id
            Button_miLabelPolygon = True
        
        Case ID_Active_Lable_BeeLine                                                    'ֱ��
            subSelectLeftorRightBouttom cMouseUsage("1").lngMouseKey, control.Id
            Button_miLabelLine = True
        
        Case ID_Active_Lable_Rect                                                       '����
            subSelectLeftorRightBouttom cMouseUsage("2").lngMouseKey, control.Id
            Button_miLabelRectangle = True
        
        Case ID_Active_Lable_VasMeasure                                                 'Ѫ����խ����
            subSelectLeftorRightBouttom cMouseUsage("1").lngMouseKey, control.Id
            Button_miLabelVasMeasure = True
            
        Case ID_Active_Lable_CadioThoracicRatio                                         '���رȲ���
            subSelectLeftorRightBouttom cMouseUsage("1").lngMouseKey, control.Id
            Button_miLabelCadiothoracicRatio = True
        
        Case ID_Active_Lable_AreaBeeLinePhoto                                           '����ֱ��ͼ
            If Not SelectedLabel Is Nothing And blnSelectedImageIfColor = False Then
                If SelectedLabel.LabelType = 1 Or SelectedLabel.LabelType = 2 Or SelectedLabel.LabelType = 5 Then funcROIHistogram SelectedLabel
                If SelectedLabel.LabelType = 3 Or SelectedLabel.LabelType = 4 Then funcDrawGreyDistribute SelectedImage, SelectedLabel
            Else
                If blnSelectedImageIfColor = True Then
                    MsgBox "��ͼ���ǲ�ɫͼ�񣬲��ܹ�����ֱ��ͼ���㡣", vbInformation, gstrSysName
                End If
            End If
            
        Case ID_Active_Lable_AdjustLine                                                 'У׼
            subcalibrate Me
            
        Case ID_Active_Lable_ClearLbale                                                 '������б�ע
            Call subLabelDeleAll
            
        Case ID_Active_Lable_DelSelectLable                                             'ɾ����ǰ��ע
            Call subDelSelectedLabel
            
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''���߲˵�''''''''''''''''''''''''''''''''''''
        Case ID_Tool_Movie                                                              '��Ӱ
            If intSelectedSerial <> 0 Then
                If SelectedImage.FrameCount = 1 And ZLShowSeriesInfos(intSelectedSerial).ImageInfos.Count < 2 Then Exit Sub
                Set frmCine.f = Me
                frmCine.Show 1, Me
            End If
            
        Case ID_Tool_Magnifier                                                          '�Ŵ�
            Dim fMagnifier As New FrmMagnify
            Set fMagnifier.f = Me
            fMagnifier.Show , Me
            
        Case ID_Tool_ArrowyCoronaryReset                                                'ʸ��״�ؽ�
            subSelectOnlyOne ID_Tool_ArrowyCoronaryReset
            Call funViewerMPR(Me)
            
        Case ID_Tool_SlopeReconstruction                                                'б���ؽ�
            'б���ؽ�
            Call funMPRslope(Me)
            
        Case ID_Tool_NumberMinusShadow                                                  '���ּ�Ӱ
            subDSA Me
            
        Case ID_Tool_BogusColour                                                        'α��
            If blnSelectedImageIfColor = False Then
                subFakeColor Me
            Else
                MsgBox "��ͼ���Ѿ��ǲ�ɫͼ�񣬲��ܹ�����α�ʲ�����", vbInformation, gstrSysName
            End If
        
        Case ID_Tool_FilmPrint                                                          '��Ƭ��ӡ
            blnPrintFilm = funFilm(Me, True, 3)
            
        Case ID_Tool_Film_AddSeries                                                     '��Ƭ��ӡ--��ӡ����
            Call funFilm(Me, False, 1)
            
        Case ID_Tool_Film_AddImage                                                      '��Ƭ��ӡ - ��ӡͼ��
            Call funFilm(Me, False, 2)
            
        Case ID_Tool_Film_AddSelected                                                   '��Ƭ��ӡ - ��ӡ��ѡͼ
            Call funFilm(Me, False, 3)
            
        Case ID_Tool_PhotoUnite                                                         'ͼ��ƴ��
            If Not blnfis Then
                Set fis.f = Me
                fis.Show , Me
                blnfis = True
                '���������ⲿͼ����¼�����Ϊͼ��ƴ�ӿ��ܱ����˽��ͼ
                RaiseEvent AfterSaveOuterImage(PstrCheckUID)
            End If
            
        Case ID_Tool_LableTool                                                          '��ע����
            Set frmLabelObject.im = Me.SelectedImage
            Set frmLabelObject.f = Me
            frmLabelObject.Show , Me
            
        Case ID_Tool_LookPhotoOption                                                    '��Ƭѡ��
            Set frmSysConfig.f = Me
            frmSysConfig.Show 1, Me
            
        Case ID_ToolBar_Left                                                            '����������
            PutToolbar ComToolBar, 2
            ArrayToolBar ComToolBar, Me.top, Me.left, Me.height, Me.width
            intToolBarPosition = 2
            Call subSaveInterfaceParaIntoDB
            
        Case ID_ToolBar_Right                                                           '����������
            PutToolbar ComToolBar, 3
            ArrayToolBar ComToolBar, Me.top, Me.left, Me.height, Me.width
            intToolBarPosition = 3
            Call subSaveInterfaceParaIntoDB
        
        Case ID_ToolBar_Top                                                             '����������
            PutToolbar ComToolBar, 0
            ArrayToolBar ComToolBar, Me.top, Me.left, Me.height, Me.width
            intToolBarPosition = 0
            Call subSaveInterfaceParaIntoDB
            
        Case ID_ToolBar_Button                                                          '����������
            PutToolbar ComToolBar, 1
            ArrayToolBar ComToolBar, Me.top, Me.left, Me.height, Me.width
            intToolBarPosition = 1
            Call subSaveInterfaceParaIntoDB
            
        Case ID_toolBar_16Icon                                                          '������ͼ��16*16��ʾ
            ReplaceToolBarIcon ComToolBar, ImgList16, 16, 16
'            ArrayToolBar ComToolBar, Me.top, Me.left, Me.height, Me.width
            ComToolBar.RecalcLayout
            intToolBarIconSize = 16
            Call subSaveInterfaceParaIntoDB
            
        Case ID_ToolBar_24Icon                                                          '������ͼ��24*24��ʾ
            ReplaceToolBarIcon ComToolBar, ImgList24, 24, 24
'            ArrayToolBar ComToolBar, Me.top, Me.left, Me.height, Me.width
            ComToolBar.RecalcLayout
            intToolBarIconSize = 24
            Call subSaveInterfaceParaIntoDB
        
        Case ID_ToolBar_32Icon                                                          '������ͼ��32*32��ʾ
            ReplaceToolBarIcon ComToolBar, ImgList32, 32, 32
            'ArrayToolBar ComToolBar, Me.top, Me.left, Me.height, Me.width
            ComToolBar.RecalcLayout
            intToolBarIconSize = 32
            Call subSaveInterfaceParaIntoDB
            
        Case ID_ToolBar_Hide                                                            '������������
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
        '''''''''''''''''''''''''''''''''''''����'''''''''''''''''''''''''''''''''''''
        Case ID_Help_Help                                                               '����
            '���ܣ����ð�������
            Shell "hh.exe  zl9ImgViewer.chm", vbNormalFocus
            'ShowHelp App.ProductName, Me.hwnd, "frmInstrument"
        Case ID_Help_WebZLSOFT_WEB                                                      '������ҳ
'            Call zlHomePage(Me.hwnd)
        
        Case ID_Help_WebZLSOFT_Mail                                                     '���ͷ���
'            Call zlMailTo(Me.hwnd)
            
        Case ID_Help_UpdateDB                                                           '����Access���ݿ�
            Set frmUpdateDB.m_cnAccess = cnAccess
            frmUpdateDB.Show 1, Me
        Case ID_Help_About                                                              '����
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
    
    If CommandBar.Title = "�����ӡ" Then
        Dim controlForm As CommandBarControlCustom
        CommandBar.Controls.DeleteAll
        Set controlForm = CommandBar.Controls.Add(xtpControlCustom, 0, "�����ӡ")
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
        Case ID_View_UpSeries, ID_View_DownSeries   '��һ���У���һ����
            control.Visible = IIf(ZLSeriesInfos.Count <= 1, False, True)
        Case ID_Active_AdjustWindow_HandAdjustWindow, ID_Active_AdjustWindow_AutoAdjustWindow, ID_Active_Eddy_LeftRight, ID_Active_Eddy_TopButton, ID_Active_Eddy_Left90, _
            ID_Active_Eddy_Right90, ID_Active_ReverseVideo, ID_Active_Cut, ID_Active_PointingLine_ALL, ID_Active_PointingLine_FirstLast, _
            ID_Active_PointingLine_Now, ID_Active_SieveLens_LancetMinus, ID_Active_SieveLens_LancetAdd, ID_Active_SieveLens_LancetAdd, _
            ID_Active_SieveLens_FlatnessMinus, ID_Active_SieveLens_FlatnessAdd, ID_Active_Sievelens_LeftMoveMinus, ID_Active_Sievelens_LeftMoveAdd, _
            ID_Active_Sievelens_PhotoReset, ID_Active_SieveLens_Model            'ͼ���������
            
            control.Visible = (InStr(mstrPrivs, "ͼ���������") <> 0)
            
        Case ID_Active_Lable_Text, ID_Active_Lable_Arrowhead, ID_Active_Lable_Ellipse, ID_Active_Lable_Angle, ID_Active_Lable_BeeLine, _
            ID_Active_Lable_Rect, ID_Active_Lable_AdjustLine, ID_Active_Lable_ClearLbale, ID_Active_Lable_DelSelectLable, _
            ID_Active_Lable_AreaBeeLinePhoto, ID_Active_Lable_Area, ID_Active_Lable_Curve   'ͼ���ע����
            
            control.Visible = (InStr(mstrPrivs, "ͼ���ע����") <> 0)
            
        Case ID_Tool_ArrowyCoronaryReset        'ʸ��״�ؽ�
            control.Visible = (InStr(mstrPrivs, "ʸ��״�ؽ�") <> 0)
            
            control.Checked = blnInMPR  '�޸��ؽ���ť��״̬
        Case ID_Tool_SlopeReconstruction        'б���ؽ�
            control.Visible = (InStr(mstrPrivs, "ʸ��״�ؽ�") <> 0)
        Case ID_Tool_BogusColour        'α��
            control.Visible = (InStr(mstrPrivs, "α��") <> 0)
        Case ID_Active_PointingLine_3DLine  '��ά���
            control.Visible = (InStr(mstrPrivs, "��ά���") <> 0)
        Case ID_Tool_NumberMinusShadow      '���ּ�Ӱ
            control.Visible = (InStr(mstrPrivs, "���ּ�Ӱ") <> 0)
            If control.Visible = True And Not SelectedImage Is Nothing Then
                control.Enabled = IIf(SelectedImage.FrameCount > 1, True, False)
            End If
        Case ID_File_OpenDicomDir           'DICOM_DIR
            control.Visible = (InStr(mstrPrivs, "DICOM_DIR") <> 0)
        Case ID_Tool_FilmPrint              '��Ƭ�Ű��ӡ
            control.Visible = (InStr(mstrPrivs, "��Ƭ�Ű��ӡ") <> 0)
        Case ID_Active_Lable_VasMeasure     'Ѫ����խ����
            control.Visible = (InStr(mstrPrivs, "Ѫ����խ����") <> 0)
        Case ID_Tool_PhotoUnite             'ͼ��ƴ��
            control.Visible = (InStr(mstrPrivs, "ͼ��ƴ��") <> 0)
        Case ID_File_Send_GetHost     '�����棬��Ȩ��
            control.Visible = (InStr(mstrPrivs, "������") = 0)
        Case ID_File_Open                   '������Ƭվ����Ȩ��
            control.Visible = (InStr(mstrPrivs, "������Ƭվ") = 0)
        Case ID_File_DelReport, ID_File_SAveASReport        '���汨��ͼ��ɾ������ͼ����Ҫ�С�����ͼ��Ȩ�ޣ��Ҳ��ǡ������桱
            If InStr(mstrPrivs, "������") <> 0 Then
                control.Visible = False
            Else
                control.Visible = (InStr(mstrPrivs, "����ͼ��") <> 0)
            End If
        Case ID_File_SaveFile               '����ͼ��
            If InStr(mstrPrivs, "������") <> 0 Then
                control.Visible = False
            Else
                control.Visible = (InStr(mstrPrivs, "����ͼ��") <> 0)
            End If
        Case ID_File_SaveFile
            control.Visible = (InStr(mstrPrivs, "����ͼ��") <> 0)
        Case ID_File_SaveASFile, ID_File_SaveToCD, ID_File_Send_GetHost, ID_File_Send_OutPowerPoint
            control.Visible = (InStr(mstrPrivs, "���ͼ��") <> 0)
    End Select
End Sub

Private Sub InitFaceSheme()
    Dim Pane1 As Pane
    
    With Me.dkpMain
        .CloseAll
        .SetCommandBars ComToolBar
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = dkpMain.CreatePane(1, 200, 200, DockBottomOf, Nothing)
    Pane1.Handle = picViewer.hwnd
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '������ļ����¼�
    '����ESCʱ������п�ִ�����״̬
    If KeyCode = vbKeyEscape Then        'ESC
        subSelectLeftorRightBouttom 1, ID_Tool_NothinMouseState
        subSelectLeftorRightBouttom 2, ID_Tool_NothinMouseState
        '�����ȫ�������˳�ȫ��
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
    
    '��ʼ������'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    intSpaceSize = 100   ''''����viewer֮��ļ��
    blnAngle = False
    intVasMeasure = 0
    intCadioThoracicRatio = 0
    
    Call RestoreWinState(Me, App.ProductName)
    
    InitFaceSheme   '��ʼ�����沼��
    
    '''''''''''''''''[��Ϣ��עλ������]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim DG As New DicomGlobal
    DG.DirectionStrings = IIf(blnChinaMark, "��\��\ǰ\��\��\ͷ", "R\L\A\P\I\S")
    
    '����ϵͳĬ��ֵ,Ĭ�ϣ�δѡ�У��ǵ�ǰ��ͼ��߿����ɫ�����ͣ��߿�
    lngDefaultImageBorderColor = vbWhite
    lngDefaultImageBorderLineStyle = 0
    lngDefaultImageBorderLineWidth = 1
    
    ''''''''''''''''''�����ݿ��л�ȡԤ�贰��λ��Ԥ����Ļ���֣���д��ϵͳ������
    subGetLayoutToVar glngUserID                    '�����ݿ��ж�ȡ���к�ͼ�񲼾ֵ�ϵͳ����
    Call subGetWWWLToVal                            '�����ݿ��ж�ȡ����λ��ϵͳ����
    Call subGetFilterToVal                          '�����ݿ��ж�ȡ�˾����õ�ϵͳ����
    subGetImageShutterToVar glngUserID              '�����ݿ��ж�ȡͼ���������õ�ϵͳ����
    subGetMouseUsageToVar glngUserID                '�����ݿ��ж�ȡ����÷����õ�ֵ��ϵͳ����
    subGetInterfaceParaToVar glngUserID             '�����ݿ��ȡ��Ӱ���������������ݣ������䱣�浽ϵͳ�����С�
    subGetLabelStoreToVar                           '�����ݿ��ȡ�����ע�������Ϣ
    subGetInfoLabelToVar                            '�����ݿ��ȡ��Ϣ��עλ���������ݵ�ϵͳ����
    subGetDBDicomPrintToVar                         '�����ݿ��ȡDICOM��ӡ�Ĵ�ӡ�����õ�ϵͳ����
    subGetParameters                                '��ϵͳ�������ȡ����
    '��ע����У���ȡ��Ƭ��ӡ�ı�ע�����С
    intFilmFontSize = GetSetting("ZLSOFT", "����ģ��\zlPacsCore", "��Ƭ����", "10")
    
    
    ''''''''''''''''''''''''''����������'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If App.LogMode <> 0 Then
        Dim ret As Long
    '    '��¼ԭ����window�����ַ
        preWinProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    '    '���Զ���������ԭ����window����
        ret = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf Wndproc)
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ReDim FilmLayouts(0)        '���彺Ƭ��ӡͼ�񲼾�����
    Set ZLSeriesInfos = New Collection  '��ʼ��ͼ��������Ϣ����
    Set ZLShowSeriesInfos = New Collection '��ʼ��ͼ����ʾ������Ϣ����
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''����״̬��ͼ��'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Set sbStatusBar.Panels(1).Picture = ImgList24.ListImages("����ͼ��").Picture
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    ''''''''''''''''''''''''''[��ʼ�����п���]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MSFViewer.Rows = 1
    MSFViewer.Cols = 16
    picViewer.BackColor = lngProgramBackColor
    ''''''''''''''''''''''��ʹ��������ҷֲ�''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''''����ʼ���˵���ť��''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    LoadBarSetup Me
    
    ''''''''''''''''''''''''''''''''''[��ʹ��״̬�������С]'''''''''''''''''''''''''''''''''''''''''''''
    Me.sbStatusBar.Font.Size = IIf(intStatusBarFontSize < 1, 10, intStatusBarFontSize)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''��ȡ���Ի�����'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    strRegPath = "˽��ģ��\ZLHIS\" & App.EXEName & "\frmViewer"
    
    If GetSetting("ZLSOFT", strRegPath, "WindowState", 2) = 1 Then
        '�����Ƭվ����ʾ״̬����С������ָ���Ĭ�ϵ�λ��
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
        '������ڵ�״̬����ȷ���������ڷ��ص�����״̬
        If Abs(Me.left) > 200000 Or Abs(Me.top) > 200000 Then
            Me.top = 0
            Me.left = 0
            Me.width = 1024 * Screen.TwipsPerPixelX
            Me.height = 768 * Screen.TwipsPerPixelY
            Me.WindowState = 2
        End If
    End If
    Button_miMouseShowValue = GetSetting("ZLSOFT", strRegPath, "��ʾ�������ֵ", False)
    Button_miShowMiniSeries = GetSetting("ZLSOFT", strRegPath, "��ʾ��������ͼ", False)
    Button_miViewAllSeries = GetSetting("ZLSOFT", strRegPath, "ȫ���й�Ƭ", False)
    Button_miImageInPhase = GetSetting("ZLSOFT", strRegPath, "ͼ���ʽͬ��", True)
    
    
    Button_miDispPatientInfo = True 'Ĭ����ʾͼ����Ϣ
    Button_miShowOverlay = True     'Ĭ����ʾOverlay
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    subInitSerial Me    ''�Դ���������ݽ��г�ʼ��
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '�жϹ�Ƭվ�Ƿ�ע��ɹ���ע�᲻�ɹ���رչ�Ƭվ
    gintҽ����Ƭվ���� = getLicenseCount(LOGIN_TYPE_ҽ����Ƭվ)
    gint��Ƭ��ӡ�� = getLicenseCount(LOGIN_TYPE_��Ƭ��ӡ��)
    mstr����ʱ�� = FunLogIn(LOGIN_TYPE_ҽ����Ƭվ)
    If mstr����ʱ�� = "" Then
        blnLogined = False
    Else
        blnLogined = True
    End If
    
    '��ȡ�Զ���Ĺ�����
    ComToolBar.LoadCommandBars "������Ƭվ", App.Title, "�Զ��幤����"

    '�����Զ��幤����
    ComToolBar.EnableCustomization True

    '��ֹ����������ͨ�ù����������ݱ��Զ���
    ComToolBar.Item(7).Customizable = False
    ComToolBar.ActiveMenuBar.Customizable = False

    'ͳһͼ���С
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
    '��ʼ��������ҷ���İ�����ť
    Dim i As Integer
    
    '������������״̬
    subSelectLeftorRightBouttom 1, ID_Tool_NothinMouseState
    subSelectLeftorRightBouttom 2, ID_Tool_NothinMouseState
        
    ''''''''''''''''''''''''''''''''''[��ʼ��������ҷ���İ�����ť]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To cMouseUsage.Count
        If cMouseUsage(i).strProgramName <> "No" Then
            ComToolBar.FindControl(, cMouseUsage(i).ButtomID, , True).Checked = cMouseUsage(i).bSelected
        ElseIf cMouseUsage(i).lngFuncNo = 201 Then  '�ü��Ϳ�ѡ��ť
            ComToolBar.FindControl(, ID_ACtive_FrameSelectImage, , True).Checked = cMouseUsage(i).bSelected
        End If
    Next
    
    '����
    If ComToolBar.FindControl(, ID_Active_Shuttle, , True).Checked = True Then
        Button_miStack = True
    End If
    '����
    If ComToolBar.FindControl(, ID_Active_AdjustWindow_HandAdjustWindow, , True).Checked = True Then
        Button_miWidthLevel = True
    End If
    '����
    If ComToolBar.FindControl(, ID_Active_Cruise, , True).Checked = True Then
        Button_miCruise = True
    End If
    '����
    If ComToolBar.FindControl(, ID_Active_Zoom, , True).Checked = True Then
        Button_miZoom = True
    End If
    '����Ӧ����
    If ComToolBar.FindControl(, ID_Active_AdjustWindow_AutoAdjustWindow, , True).Checked = True Then
        Button_miAutoWidthLevel = True
    End If
    '��ѡͼ��
    If ComToolBar.FindControl(, ID_ACtive_FrameSelectImage, , True).Checked = True Then
        Button_miFrameSelectImage = True
    End If
    'Ĭ��Ϊ��ʾCTֵ
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_ACtive_Mouse_Value, , True).Checked = Button_miMouseShowValue
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_ACtive_Mouse_Value, , True).Checked = Button_miMouseShowValue
    '��ʾ��������ͼ
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_ShowMiniSeries, , True).Checked = Button_miShowMiniSeries
    ComToolBar.Item(ToolBar_Object).FindControl(, ID_View_ShowMiniSeries, , True).Checked = Button_miShowMiniSeries
    'ȫ���й�Ƭ
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_ViewAllSeries, , True).Checked = Button_miViewAllSeries
    ComToolBar.Item(ToolBar_Object).FindControl(, ID_View_ViewAllSeries, , True).Checked = Button_miViewAllSeries
    
    ''''''''''''''''''''''''''''''''''��ʹ������������'''''''''''''''''''''''''''''''''''''''''''''''''
    IntComBarTheme = Me.ComToolBar.VisualTheme
    '''''''''''''''''''''''''''''''''��ʼ��һЩĬ��ֵ''''''''''''''''''''''''''''''''''''''''''''''''''
    ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_PhotoSerial_PhotoNumber, , True).Checked = True
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_PhotoSerial_PhotoNumber, , True).Checked = True
    ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_ShowScale_AutoShow, , True).Checked = True
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_ShowScale_AutoShow, , True).Checked = True
    
    'ͼ���ʽͬ��
    ComToolBar.Item(ToolBar_Object).FindControl(, ID_Active_Also_Photo, , True).Checked = Button_miImageInPhase
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Also_Photo, , True).Checked = Button_miImageInPhase
    
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_LableShow, , True).Checked = True
    Button_miDispLabelInfo = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
    Dim strClearDate As String  '�ݴ��ϴ��������ͼ�������
    
    '�����MPR״̬�У�����ʾ�Ƿ񱣴�MPR���
    If blnInMPR = True Then
        If funViewerMPR(Me) = False Then Cancel = -1
    End If
    
    If Cancel = 0 Then
        
        ComToolBar.SaveCommandBars "������Ƭվ", App.Title, "�Զ��幤����"
        
        If Dir(App.Path & "\temp\*.*") <> "" Then
            Kill App.Path & "\temp\*.*"
        End If
        If Dir(App.Path & "\temp", vbDirectory) <> "" Then
            RmDir App.Path & "\temp"
        End If
        
        '��հ�ť
        subSelectLeftorRightBouttom 1, ID_Tool_NothinMouseState
        subSelectLeftorRightBouttom 2, ID_Tool_NothinMouseState
        
        'ɾ��ͼ��
        Call subKillPicture(True)
        
        '���汾������
        strRegPath = "˽��ģ��\ZLHIS\" & App.EXEName & "\frmViewer"
        
        SaveSetting "ZLSOFT", strRegPath, "Top", Me.top
        SaveSetting "ZLSOFT", strRegPath, "Left", Me.left
        SaveSetting "ZLSOFT", strRegPath, "Width", Me.width
        SaveSetting "ZLSOFT", strRegPath, "Height", Me.height
        SaveSetting "ZLSOFT", strRegPath, "WindowState", Me.WindowState
        SaveSetting "ZLSOFT", strRegPath, "��ʾ��������ͼ", Button_miShowMiniSeries
        SaveSetting "ZLSOFT", strRegPath, "��ʾ�������ֵ", Button_miMouseShowValue
        SaveSetting "ZLSOFT", strRegPath, "ȫ���й�Ƭ", Button_miViewAllSeries
        SaveSetting "ZLSOFT", strRegPath, "ͼ���ʽͬ��", Button_miImageInPhase
        
        
        If frmMiniSeries.hwnd <> 0 Then
            Unload frmMiniSeries
        End If
        
        '�������ͼ��ÿ�����һ��
        strClearDate = GetSetting("ZLSOFT", strRegPath, "�������ͼ��", Date)
        If IsDate(strClearDate) = False Then
            strClearDate = Date
        End If
        If DateDiff("d", strClearDate, Date) >= 7 Then
            Call ClearCacheFolder(PstrBufferImagePath)
            SaveSetting "ZLSOFT", strRegPath, "�������ͼ��", Date
        End If
        
        '��鱾���˳��Ƿ�Ϸ�ע�����˳�
        Call FunLogOut(LOGIN_TYPE_ҽ����Ƭվ, mstr����ʱ��)
    End If
End Sub

Private Sub mfrmFilm_AfterPrinted(strImageUIDS As String)
    '��ӡ��ɵ��¼�����Ҫ����ˢ��ͼƬ�Ĵ�ӡ��Ϣ
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
    
    '�������ͼ��UID
    For intIndex = 1 To UBound(arrImageUID) - 1
        '��������ʾ��ͼ���в���
        blnFindImage = False
        For i = 1 To ZLShowSeriesInfos.Count
            For j = 1 To ZLShowSeriesInfos(i).ImageInfos.Count
                If ZLShowSeriesInfos(i).ImageInfos(j).InstanceUID = arrImageUID(intIndex) Then
                    '������ʾͼ�����ҵ�ͼ�񣬸��Ĵ�ӡ���
                    
                    '���Ĵ�ӡ���
                    ZLShowSeriesInfos(i).ImageInfos(j).blnPrinted = True
                    
                    '����ͼ����ʾ�ı�ע
                    '�Զ�����ͼ���С���ж��Ƿ���ʾ�����Ľ���Ϣ,��ʾ��������ͼ���еĲ�����Ϣ
                    Call subDisplayPatientInfo(Viewer(i))
                    
                    blnFindImage = True
                    Exit For
                End If
            Next j
            If blnFindImage Then
                Exit For
            End If
        Next i
        
        '�������в���
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
'2009��

    Dim intSeriesIndex As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    
    If Source.Images.Count <= 0 Then Exit Sub
    intSeriesIndex = Source.Tag
    
    '�������Viewer�ڷŵ�λ��
    Call subIsSerialXY(Me, x, y, intCol, intRow)
    
    Call subCreateAndPlaceAViewer(intSeriesIndex, intRow, intCol)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub picViewer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim imgs As New DicomImages
    Dim img As DicomImage
    Dim i As Integer
    '''''''''''''''''''''''''''''''[����ڿհ����д��������,�򵯳��������еĵ����˵�]''''''''''''''''''''''''''''''''''''''''''''''''
    If Button = 2 Then
        For i = 1 To ZLSeriesInfos.Count
            '����һ��ͼ��
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
    If blfrmRefresh = False Then Exit Sub   'ˢ�¹�������ʱ�򣬴��岻��Ҫˢ��
    
    If intSelectedSerial < 1 Then Exit Sub
    Call subResizeSeries(Me)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PicX_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        PicXX.Visible = True
        PicXX.left = PicX(Index).left
        intFactMoveX = PicX(Index).left '��¼��갴�º���ָ�����λ��
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
'���ܣ������϶�������굯���¼���������ʱ�϶�������ʾ��ʵ�϶�����
'      �����°ڷ�viewer�Ĵ�С,����viewer��С�ж��Ƿ��Զ����ز�����Ϣ
'������Index--�Զ����ɣ��϶�����ţ�
'���أ���
'------------------------------------------------
    Dim i, j, k As Long
    If Button = 1 Then
        PicX(Index).left = PicXX.left
        PicXX.Visible = False
        intFactMoveX = PicX(Index).left - intFactMoveX
        PicX(Index).Tag = PicX(Index).left
        For i = Index + 1 To intMaxAreaX - 1    '�ƶ��÷ָ����ұߵ������ָ���
            If PicX(i).Tag <> "" Then
                PicX(i).left = Val(PicX(i).Tag) + intFactMoveX
                PicX(i).Tag = PicX(i).left
                If PicX(i).left > Me.ScaleWidth - intSpaceSize Then
                    PicX(i).left = Me.ScaleWidth - intSpaceSize
                End If
            End If
        Next
        '���°ڷ������ݺ�ָ����Ľ���
        For i = 1 To intMaxAreaX - 1
            For j = 1 To intMaxAreaY - 1
                k = (j - 1) * (intMaxAreaX - 1) + i
                PicXY(k).top = PicY(j).top
                PicXY(k).left = PicX(i).left
            Next
        Next
        '���°ڷ�����Viewer
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
'���ܣ������϶�������굯���¼���������ʱ�϶�������ʾ��ʵ�϶�����
'      �����°ڷ�viewer�Ĵ�С,����viewer��С�ж��Ƿ��Զ����ز�����Ϣ
'������Index--�Զ����ɣ��϶�����ţ�
'���أ���
'�����ˣ�����
'------------------------------------------------
    Dim i, j, k As Long
    If Button = 1 Then
        PicY(Index).top = PicYY.top
        PicYY.Visible = False
        intFactMoveY = PicY(Index).top - intFactMoveY
        PicY(Index).Tag = PicY(Index).top
        For i = Index + 1 To intMaxAreaY - 1        '�ƶ��÷ָ�������������ָ���
            If PicY(i).Tag <> "" Then
                PicY(i).top = Val(PicY(i).Tag) + intFactMoveY
                PicY(i).Tag = PicY(i).top
                If PicY(i).top > Me.picViewer.ScaleHeight - intSpaceSize Then
                    PicY(i).top = Me.picViewer.ScaleHeight - intSpaceSize
                End If
            End If
        Next
        '���°ڷ������ݺ�ָ����Ľ���
        For i = 1 To intMaxAreaX - 1
            For j = 1 To intMaxAreaY - 1
                k = (j - 1) * (intMaxAreaX - 1) + i
                PicXY(k).top = PicY(j).top
                PicXY(k).left = PicX(i).left
            Next
        Next
        '���°ڷ�����Viewer
        Call subMoveViewers(Me, Index, 1)
    End If
End Sub

Private Sub subChangeASeries(intType As Integer)
'------------------------------------------------
'���ܣ� ����intType�����ͣ��л����С�
'������
'       intType�����л���ʽ��
'                   ���intType��1������һ����ȡ����ǰ���С�
'                   ���intType=2������һ����ȡ����ǰ���С�
'���أ�
'------------------------------------------------
    Dim intSeriesIndex As Integer
    
    If intSelectedSerial = 0 Then Exit Sub
    
    On Error GoTo err
    intSeriesIndex = Viewer(intSelectedSerial).Tag
    If intType = 1 Then     '�л�����һ����
        intSeriesIndex = intSeriesIndex + 1
        If intSeriesIndex > ZLSeriesInfos.Count Then
            intSeriesIndex = 1
        End If
    ElseIf intType = 2 Then '�л�����һ����
        intSeriesIndex = intSeriesIndex - 1
        If intSeriesIndex <= 0 Then
            intSeriesIndex = ZLSeriesInfos.Count
        End If
    End If
    '�������е�ͼ�����viewer(intSelectedSerial)�е�ͼ��
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
        
        '�������е�ͼ�����viewer(index)�е�ͼ��
        Call funcSwapSeries(Index, intSeriesIndex)
    End If
    Exit Sub
error:
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Viewer_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '����Viewer�ļ����¼�
    
    On Error GoTo err
    
    '����Del�����ѡ�б�ע��ɾ����ע������ɾ����ǰ����
    If KeyCode = 46 Then        'Delete
        If Not SelectedLabel Is Nothing Then    '�����ǰѡ���˱�ע��ɾ����ǰ��ע
            Call subDelSelectedLabel
        Else
            subCloseSeries  '�����ǰû��ѡ�б�ע��ɾ����ǰ����
        End If
    End If
    
    '����PageUp��PageDown����ҳ���·�ͼ��
    '�����ϼ�ͷ���¼�ͷ�����·�����ͼ��
    ',End and Home ������һ��ͼ�������һ��ͼ
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Or KeyCode = 35 Or KeyCode = 36 Then        ' �ϼ�ͷ���¼�ͷ,PageUp and PageDown ,End and Home
        If Viewer(intSelectedSerial).Visible = False Then Exit Sub
        If VScro(intSelectedSerial).Visible = False Then Exit Sub
        
        If KeyCode = 38 Then        '�ϼ�ͷ
            If VScro(intSelectedSerial).Value - 1 < 1 Then
                VScro(intSelectedSerial).Value = 1
            Else
                VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Value - 1
            End If
        End If
        If KeyCode = 40 Then        '�¼�ͷ
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
        
        If KeyCode = 35 Then        'End ���һ��ͼ
            VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Max
        End If
        
        If KeyCode = 36 Then        'Home ��һ��ͼ
            VScro(intSelectedSerial).Value = 1
        End If
    End If
    
    '���� ���Ҽ�ͷ�����·�����
    If KeyCode = 37 Then    '���ͷ
        '�����ͷ��һ����
        subChangeASeries 2
    End If
    If KeyCode = 39 Then    '�Ҽ�ͷ
        '�����Ҽ�ͷ��һ����
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
    subTakeOut1 ls, SelectedImage, True                                                 ''''��ѡ�еı�ע��ȥ��ϵͳ��ע���ڲü������С��6�ı�ע
    If ls.Count > 0 Then     ''˫����һ����ע
'        i = funMouseOverPeriod(Viewer(Index), SelectedImage, lngBaseX, lngBaseY)       ''�ж�����Ƿ�Խ�����
'        If i = 0 Then   ''ѡ��һ����ע
        If SelectedImage.Labels.IndexOf(ls(1)) > G_INT_SYS_LABEL_COUNT Then   ''''����˫�������ƾ����LABEL
            ''''''''''�������ֵ�˫���޸�'''''''''''''''''''''''''''''''''''
            If Mid(ls(1).Tag, 1, 3) = "TXT" Then
                Set SelectedLabelT = ls(1)
                Set SelectedLabel = Nothing
                isSelectedLabel = True
                SubChangeColor ls(1), Me ''�ı���ʾ��ɫ
                lblChange = SelectedLabelT.Text + " "    '''''ͨ��lblChange��AutoSize�����Զ�����text��Ĵ�С
                txtText = SelectedLabelT.Text
                oldFontSize = SelectedLabelT.FontSize
                lblChange.FontSize = lblChange.FontSize * IIf(blnLabelTextScaleFontSize, SelectedImage.ActualZoom, 1)
                txtText.FontSize = lblChange.FontSize
                cx = SelectedLabelT.left
                cy = SelectedLabelT.top
                subTextCoordinate SelectedImage, cx, cy, lblChange          '''''����ѡװ�������������ת��
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
                funcROIHistogram SelectedLabel      '�����������͵ı�ע����ֱ��ͼ���������Ϊ��Ҫ��ֱ��ͼ�ı�ע
            ElseIf left(ls(1).Tag, 3) = "VAS" Then      'Ѫ����խ����
                '��ʾ�������
                
                Set frmVasMeasure.lblText = SelectedLabel.TagObject
                Set frmVasMeasure.f = Me
                frmVasMeasure.Show 1, Me
            ElseIf ls(1).LabelType = 3 Or ls(1).LabelType = 4 Then
                funcDrawGreyDistribute SelectedImage, SelectedLabel 'ֱ�ߺͶ���߱�ע
            End If
            Exit Sub
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''�����ʾת������ʾ'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With Viewer(intSelectedSerial)
        '''''''''����ͼƬƴ�ӵ����''''''''''''''''''''''''''''
        If blnfis Then
            Call fis.subLoadImage(SelectedImage)
        ElseIf blnPrintFilm Then    '����Ƭ��ӡ�����
            Call AddImgToFilm(SelectedImage, Viewer(Index), ZLShowSeriesInfos(intSelectedSerial).ImageInfos(SelectedImageIndex).blnPrinted)
        ElseIf intClickImageIndex > 0 And intDblClickButton = 1 Then
            
            ''''����Ƕ��л������ʾ��ת��Ϊ���е�����ʾ
            If .MultiColumns > 1 Or .MultiRows > 1 Then
                .MultiColumns = 1
                .MultiRows = 1
                MSFViewer.TextMatrix(intSelectedSerial, 3) = .CurrentIndex
                .CurrentIndex = intClickImageIndex
                Set SelectedImage = .CurrentImage
                SelectedImageIndex = .CurrentIndex
                intSelectedSerial = Index
            '''''''''��ǰʹ�ù����л������ʾ����лָ�''''''''
            ElseIf MSFViewer.TextMatrix(intSelectedSerial, 5) > 1 Or MSFViewer.TextMatrix(intSelectedSerial, 6) > 1 Then
                .MultiColumns = MSFViewer.TextMatrix(intSelectedSerial, 5)
                .MultiRows = MSFViewer.TextMatrix(intSelectedSerial, 6)
                .CurrentIndex = MSFViewer.TextMatrix(intSelectedSerial, 7)
                MSFViewer.TextMatrix(intSelectedSerial, 3) = MSFViewer.TextMatrix(intSelectedSerial, 8)
                SelectedImageIndex = MSFViewer.TextMatrix(intSelectedSerial, 3)
                Set SelectedImage = .Images(IIf(SelectedImageIndex = 0, .Images.Count, SelectedImageIndex))
            End If
            
            '�жϹ������Ƿ���Ҫ��ʾ�������Ҫ����ʾ�����������ù����������ֵ����Сֵ��LarghChange��
            subDisplayScrollBar Index, Me, False
    
            '''''''''��ͼ�����
            subDispframe Me, Viewer(intSelectedSerial)
            
            '�Զ�����ͼ���С���ж��Ƿ���ʾ�����Ľ���Ϣ,��ʾ��������ͼ���еĲ�����Ϣ
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
    lngBaseXX = x  ''��mouse DblClickʹ��
    lngBaseYY = y
    intClickImageIndex = Viewer(Index).imageIndex(x, y)         ''''��ǰ�����ͼ��INDEX
    intDblClickButton = Button                                  ''''��DblClickʹ�õ���갴
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ѡ��һ��Viewer
    Call subSelectAViewer(Index, intClickImageIndex)
    
    If SelectedImage Is Nothing Then Exit Sub
    
    
    '''''''''''''''''''''''''''''''''''�жϵ�ǰѡͼ���Ƿ��ǲ�ɫͼ��''''''''''''''''''''''''''''''''''''''''
    If SelectedImage.Attributes(&H28, &H4) = "MONOCHROME2" Or SelectedImage.Attributes(&H28, &H4) = "MONOCHROME1" Then
        blnSelectedImageIfColor = False
    Else
        blnSelectedImageIfColor = True
    End If
    ''''''''''''''''����������,����ͼ��ѡ����''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Button = 1 And intClickImageIndex <> 0 Then
        Set ls = Viewer(Index).LabelHits(x, y, True, False, False)
        If ls.Count <> 0 Then
            For Each l In ls ''''����ѡ����
                If Mid(l.Tag, 1, 1) = "B" Then
                    l.Visible = Not l.Visible
                    ZLShowSeriesInfos(Index).ImageInfos(Viewer(Index).Images(intClickImageIndex).Tag).blnSelected = IIf(l.Visible, True, False)
                    Viewer(Index).Refresh
                    Exit Sub
                End If
            Next
        End If
    End If
    '''''''''''''''''''''''''''''''''�ǶȻ����'''''''''''''''''''''''''''''''''''''''
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
    ''''3D���Ĵ���
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Button = cMouseUsage("106").lngMouseKey And Shift = cMouseUsage("106").lngShift And Button_mi3dCursor Then
         sub3DCursorStart SelectedImage
         Viewer(intSelectedSerial).Refresh
        Exit Sub
    End If
    ''''''''''''[��ע��ѡ���ƶ���������С]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Button = cMouseUsage("201").lngMouseKey And Shift = cMouseUsage("201").lngShift And intClickImageIndex > 0 And Not LabelDrawing Then
        Set ls = Viewer(Index).LabelHits(x, y, False, False, True)
        Set ols = Viewer(Index).LabelHits(x, y, False, False, True)
        subTakeOut1 ls, SelectedImage, True         '��ls��ȥ��ϵͳ��ע�����Ϊ1-5�Ĳü���ע
        subTakeOut1 ols, SelectedImage, False       '��ols��ȥ��ϵͳ��ע�����Ϊ2-5�Ĳü���ע
        lngBaseX = Viewer(Index).ImageXPosition(x, y)
        lngBaseY = Viewer(Index).ImageYPosition(x, y)
        If ls.Count > 0 Or ols.Count = 1 Then      ''���ѡ��һ����ע
            i = funMouseOverPeriod(Viewer(Index), SelectedImage, x, y)   ''���������Խ���ľ�����
            If i = 0 Then   ''û��ѡ���κξ��������ѡ�е���һ����ע
                If ls.Count > 0 Then
                    If SelectedImage.Labels.IndexOf(ls(1)) >= G_INT_SYS_LABEL_MPRV And SelectedImage.Labels.IndexOf(ls(1)) <= G_INT_SYS_LABEL_MPR_POINT_O Then
                        '�ǡ�ʸ��״�ؽ�����صı�ע
                        For j = 1 To ls.Count
                            If SelectedImage.Labels.IndexOf(ls(j)) > m Then m = SelectedImage.Labels.IndexOf(ls(j))
                        Next
                        Set SelectedLabel = SelectedImage.Labels(m)     'mΪ������ı�ע��
                    ElseIf (SelectedImage.Labels.IndexOf(ls(1)) = G_INT_SYS_LABEL_MPR_RESULT_H) _
                        Or (SelectedImage.Labels.IndexOf(ls(1)) = G_INT_SYS_LABEL_MPR_RESULT_V) Then
                        '�ǡ�ʸ��״�ؽ������ͼ�еĺ��ߺ�����
                        Set SelectedLabel = ls(1)
                    ElseIf SelectedImage.Labels.IndexOf(ls(1)) > G_INT_SYS_LABEL_COUNT Then '����ϵͳ��ע
                        Set SelectedLabel = ls(1)
                        If left(SelectedLabel.Tag, 3) = "VAS" Then      'Ѫ����խ��������SelectedLabelָ��ֱ��
                            If Right(SelectedLabel.Tag, 1) = "T" Then
                                Set SelectedLabel = SelectedLabel.TagObject.TagObject.TagObject
                            ElseIf Right(SelectedLabel.Tag, 1) = "1" Then
                                Set SelectedLabel = SelectedLabel.TagObject.TagObject
                            ElseIf Right(SelectedLabel.Tag, 1) = "2" Then
                                Set SelectedLabel = SelectedLabel.TagObject
                            End If
                        ElseIf left(SelectedLabel.Tag, 3) = "CTR" Then '���رȲ�������SelectLabelָ��ֱ��
                            If Right(SelectedLabel.Tag, 1) = "T" Then
                                Set SelectedLabel = SelectedLabel.TagObject
                            End If
                        End If
                        SubChangeColor SelectedLabel, Me ''�ı���ʾ��ɫ
                    End If
                Else
                    Set SelectedLabel = ols(1)      ''ָ�����Ϊ1�Ĳü���ע
                End If
                If SelectedLabel Is Nothing Then Exit Sub
                isSelectedLabel = True
                '''''''''''''''''''''''''''''''''''��ʾ�û���ע�Ͳü���ע�ľ��
                If SelectedImage.Labels.IndexOf(SelectedLabel) > G_INT_SYS_LABEL_COUNT Or _
                   SelectedImage.Labels.IndexOf(SelectedLabel) <= 6 _
                   Then SubDispPeriod SelectedLabel, SelectedImage, Me    ''��ʾ���
                blnMoveLabel = True
                If oldSelectedSerial <> intSelectedSerial Then Viewer(oldSelectedSerial).Refresh  '''ˢ��ԭ��������
                Me.MousePointer = 5
                Viewer(Index).Refresh
                Exit Sub
            Else  ''��ѡ�еı�ע�Ǿ������ͨ������ı��ע��С
                blnReSizeLabel = True
                intReSizeIndex = i
                Me.MousePointer = 2
'                viewer(Index).Refresh
                Exit Sub
            End If
        End If
    End If   '''''''''����[��ע��ѡ���ƶ���������С]''''
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''[�������뿪ʼ]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Button = cMouseUsage("8").lngMouseKey And Shift = cMouseUsage("8").lngShift And Button_miLabeltext Then
        blnTextInput = True
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''[����ע������Ӧ�����Ĵ���]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If ((Button = cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift) _
        Or (Button = cMouseUsage("105").lngMouseKey And Shift = cMouseUsage("105").lngShift And Button_miAutoWidthLevel) _
        Or (Button = cMouseUsage("201").lngMouseKey And Shift = cMouseUsage("201").lngShift And Button_miFrameSelectImage)) _
       And intClickImageIndex <> 0 And Not (blnTextInput Or blnTextInputM) Then
       
        If funIsLabelMouse(Me, Button, Shift) Then  ''''�жϱ�ע����־�Ƿ���,����������ע
            '''''''''''''''''''''''''''''''[�����ע����]'''''''''''''''''''''''''''''''''''''''''
            If blnTextInput Or blnTextInputM Then
                If txtText.Visible Then     '''''���������������,����������
                    txtText_KeyPress 13
                    Exit Sub
                End If
            End If
            ''''''''''''''''��ͼ��������������ע��һ������״��עһ�������������ֱ�ע'''''''''
            SubNoDispPeriod SelectedImage, Me   'Ϊָ��ͼ�����ر�עѡ����
            SelectedImage.Labels.Add GetNewLabel(intSelectLabelStyle, Viewer(Index).ImageXPosition(x, y), Viewer(Index).ImageYPosition(x, y), 0, 0)
            Set SelectedLabel = SelectedImage.Labels(SelectedImage.Labels.Count)
            SelectedImage.Labels.Add GetNewLabel(doLabelText, SelectedLabel.left, SelectedLabel.top, 0, 0)
            ''''''''''''''''''''''''''�Զ�����'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Button = cMouseUsage("105").lngMouseKey And Shift = cMouseUsage("105").lngShift And Button_miAutoWidthLevel Then
                SelectedLabel.LineStyle = 2
                SelectedLabel.XOR = False
                blnAutoWL = True  ''����Ӧ������ʼ
            End If
            '''''''''''''''''''��ѡͼ��'''''''''''''''''''''''''''''
            If Button = cMouseUsage("201").lngMouseKey And Shift = cMouseUsage("201").lngShift And Button_miFrameSelectImage Then
                SelectedLabel.LineStyle = 2
                SelectedLabel.XOR = False
                blnFrameSelectImage = True      ''��ѡͼ��ʼ
            End If
            
            ''''''''''''''''''''����SelectedLabelTָ���ע����'''''''''''''''''''''''''
            Set SelectedLabelT = SelectedImage.Labels(SelectedImage.Labels.Count)
            SelectedLabelT.AutoSize = True
            SelectedLabelT.Margin = 0
            ''''''''''''''''''''''''''���ÿ�ʼ���в�����ע�Ĳ���'''''''''''''''''''''''''''''''''''''
            If Button_miLabelAngle And Button = cMouseUsage("7").lngMouseKey _
                And Shift = cMouseUsage("7").lngShift Then
                
                blnAngle = True  ''�Ƕȿ�ʼ
            ElseIf Button_miLabelVasMeasure And Button = cMouseUsage("1").lngMouseKey _
                And Shift = cMouseUsage("1").lngShift And intVasMeasure = 0 Then
                
                intVasMeasure = 1      'Ѫ����խ������ʼ
            ElseIf Button_miLabelCadiothoracicRatio And Button = cMouseUsage("1").lngMouseKey _
                And Shift = cMouseUsage("1").lngShift And intCadioThoracicRatio = 0 Then
                
                intCadioThoracicRatio = 1   '���رȲ�����ʼ
            End If
            
'            If Button_miAutoWidthLevel And Button = cMouseUsage("105").lngMouseKey And Shift = cMouseUsage("105").lngShift Then
'                blnAutoWL = True  ''����Ӧ������ʼ
'            End If
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If blnAngle Then          ''''''''����ǽǶȵĴ���
                SelectedLabel.TagObject = SelectedLabelT
                SelectedLabel.Tag = "JD1"
                SelectedLabelT.Tag = "JDT"
                SelectedLabelT.left = Viewer(Index).ImageXPosition(x, y)
                SelectedLabelT.top = Viewer(Index).ImageYPosition(x, y)
                SelectedLabelT.AnchorX = Viewer(Index).ImageXPosition(x, y)
                SelectedLabelT.AnchorY = Viewer(Index).ImageYPosition(x, y)
                SelectedLabelT.ShowAnchor = True
                SelectedLabelT.AnchorImageTied = True
                SelectedLabelT.LineStyle = 2        ''''ê������
            ElseIf Button_miLabelVasMeasure And (intVasMeasure = 1 Or intVasMeasure = 2) Then '����Ѫ����խ����
                If intVasMeasure = 2 Then       '������Ѫ�ܲ��ֺ���խѪ�ܲ��ֵ�TagObject����
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
                SelectedLabelT.LineStyle = 2        ''''ê������
                '����������Ѫ�ܱڱ�ע
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
                l.TagObject = SelectedImage.Labels(SelectedImage.Labels.Count - (intVasMeasure * 4 - 1)) '��TagObject���ɷ�ջ���
            ElseIf Button_miLabelCadiothoracicRatio And (intCadioThoracicRatio = 1 Or intCadioThoracicRatio = 2) Then  '�������رȲ���
                If intCadioThoracicRatio = 1 Then   '�����ಿ��
                    
                ElseIf intCadioThoracicRatio = 2 Then   '����������
                    '���������ֺ����ಿ�ֵ�TagObject����
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
                SelectedLabelT.LineStyle = 2        ''''ê������
                '��TagObject���ɷ�ջ���
                SelectedLabelT.TagObject = SelectedImage.Labels(SelectedImage.Labels.Count - (intCadioThoracicRatio * 2 - 1))
            
            Else   '''''���ǽǶȵĴ���
               SelectedLabel.TagObject = SelectedLabelT
               SelectedLabelT.TagObject = SelectedLabel
               SelectedLabelT.Tag = "RIO"
            End If
            ''''''''''''''''''''''''Ϊ����κͶ��������һ����,������㳤�ȳ���.''''''''''''''''''''''''''''''''''''''''''''''
            If Button_miLabelPolygon Or Button_miLabelPolygon Then
                SelectedLabel.AddPoint Viewer(Index).ImageXPosition(x, y), Viewer(Index).ImageYPosition(x, y)
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            LabelDrawing = True                     '���ÿ�ʼ����ע�ı��Ϊ��
            SubChangeColor SelectedLabel, Me        '�ı�ѡ��LABEL����ɫ
            Me.MousePointer = 2
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''�жϵ�ǰͼ���Ƿ�����˲ü�״̬''''''''''''''''''''''''''''''''''''''''
    If SelectedImage.Labels(1).Visible = True Then
        Button_miCutOut = True
    Else
        Button_miCutOut = False
    End If
    
    blDicomDown = True                      '�������
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
    '''''''''''''''''''''''''''''''''''[��״̬����ʾX/Y�����ͼ���ֵ]'''''''''''''''''''''''''''''''''''''
    intClickImageIndex = Viewer(Index).imageIndex(x, y)
    
    '���������ж��Ƿ������,���²�ִ�м���Ҷ�ֵ���Ч��
    If (intClickImageIndex <> 0 And Not blDicomDown) Then
        If Viewer(Index).Images(intClickImageIndex).FrameCount > 1 And Button_miMouseShowValue = False Then
            '��֡ͼ�񣬶��Ҳ��������ʾ����ֵ���򲻽��м���
            Me.Viewer(Index).ToolTipText = ""
        Else
            '�洢����ֵ�ĵ��������б��
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
                sbStatusBar.Panels(4).Text = "��:" & Viewer(Index).ImageXPosition(x, y) & "   ��:" & Viewer(Index).ImageYPosition(x, y) & "   ֵ:" & lngValue * intSlope + intIntercept
            End If
            If Button_miMouseShowValue = True Then
                Me.Viewer(Index).ToolTipText = Mid(sbStatusBar.Panels(4).Text, InStr(1, sbStatusBar.Panels(4).Text, "ֵ:") + 2)
            Else
                Me.Viewer(Index).ToolTipText = ""
            End If
        End If
    Else
        Me.Viewer(Index).ToolTipText = ""
    End If
    ''''''�������δ�����κμ�������£��ı�Խ����ǰͼ���б�עѡ�����������״
    If Button = 0 And (Not SelectedLabel Is Nothing And Not blnReSizeLabel Or Button_miCutOut) Then
        If intClickImageIndex = Viewer(Index).Images.IndexOf(SelectedImage) And (isSelectedLabel Or Button_miCutOut) Then
            i = funMouseOverPeriod(Viewer(Index), SelectedImage, x, y)  '���������Խ���ľ�����
            If i <> 0 Then  '��꾭�����Ǿ��֮һ���ı������״
                If Mid(SelectedLabel.Tag, 1, 2) = "JD" Then '����Ƕȵ������״
                      Me.MousePointer = 2
                Else            '����ǽǶȱ�ע�������״
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
                End If  '���������״����
            Else            '��꾭���Ĳ��Ǿ�����������״��ԭ��0
                Me.MousePointer = 0
            End If
        End If
    End If          '�������������δ�����κμ�������£��ı�Խ����ǰͼ���б�עѡ�����������״��
    '''''''''''''''''''''''''�ƶ���ע''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If blnMoveLabel And Not SelectedLabel Is Nothing Then
        subaCorrectCursor Viewer(intSelectedSerial), SelectedImage, x, y    ''''����ƶ��������ͼ��Χ�����������λ��
        '''''''''''''''''''''''''''''''[�ƶ���ע]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '���ж��Ƿ�����MPR�ؽ�,�жϵ�ǰѡ�еı�ע�Ƿ���ʸ��״�ؽ���5�����Ƶ�֮һ
        If SelectedImage.Labels.IndexOf(SelectedLabel) >= G_INT_SYS_LABEL_MPR_POINT_V1 _
            And SelectedImage.Labels.IndexOf(SelectedLabel) <= G_INT_SYS_LABEL_MPR_POINT_O Then    '��ʸ��״�ؽ�
            If blnInMPR = False Then
                MsgBox "MPR�ؽ������ɾ�����������ؽ������г��ִ���" & vbCrLf & vbCrLf & "�����½����ؽ���", vbInformation, gstrSysName
                blnMoveLabel = False
                Exit Sub
            End If
        End If
        '�ƶ���ע�����в����������ƶ�MPR�߲���ʾ�ؽ����ͼ
        subMoveLable SelectedLabel, Viewer(Index).ImageXPosition(x, y) - lngBaseX, Viewer(Index).ImageYPosition(x, y) - lngBaseY, Me, x, y, lngBaseX, lngBaseY
        If SelectedImage.Labels.IndexOf(SelectedLabel) > G_INT_SYS_LABEL_COUNT _
           Or SelectedImage.Labels.IndexOf(SelectedLabel) <= 6 Then   ''''��ʾ�û���ע�Ͳü���ע��ѡ����
            SubDispPeriod SelectedLabel, SelectedImage, Me
        End If
        lngBaseX = Viewer(Index).ImageXPosition(x, y)
        lngBaseY = Viewer(Index).ImageYPosition(x, y)
        Viewer(Index).Refresh
        Exit Sub
    End If
    ''''''''''''''''''''''''''�ı��ע��С'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If blnReSizeLabel And Not SelectedLabel Is Nothing Then
        subaCorrectCursor Viewer(intSelectedSerial), SelectedImage, x, y                ''''����ƶ��������ͼ��Χ�����������λ��
        '''''''''''''''''''''''''''''''[�ı��ע��С]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If left(SelectedLabel.Tag, 3) = "VAS" Then      'Ѫ����խ��������SelectedLabelָ��ֱ��
            If Right(SelectedLabel.Tag, 1) = "T" Then
                Set SelectedLabel = SelectedLabel.TagObject.TagObject.TagObject
            ElseIf Right(SelectedLabel.Tag, 1) = "1" Then
                Set SelectedLabel = SelectedLabel.TagObject.TagObject
            ElseIf Right(SelectedLabel.Tag, 1) = "2" Then
                Set SelectedLabel = SelectedLabel.TagObject
            End If
        ElseIf left(SelectedLabel.Tag, 3) = "CTR" Then  '���رȲ�������SelectedLabelָ��ֱ��
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
    '''''''''''''''''''''''''''''''''''''''''''����''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Button = cMouseUsage("101").lngMouseKey And Shift = cMouseUsage("101").lngShift And Button_miStack Then
        If Abs(y - lngBaseYY) >= lngStackStep Then          ''''���󲽳�����,����Y�����λ����Ϊ����Ĳ���
            '���ú����������е�����ƶ�
            Call subStackMouseMove(y - lngBaseYY)
            lngBaseYY = y
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''[����]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (Button = cMouseUsage("102").lngMouseKey And Shift = cMouseUsage("102").lngShift And Button_miWidthLevel And (intClickImageIndex <> 0 Or blnMouseStart)) _
        Or (Button = 4 And intMouseWheelDrag = 2) Then
        '������Ҽ���ͨ����ť���Խ��е���������м����קͨ�����ÿ��Խ��е���
        Dim DicomAttr As DicomAttribute
        Dim DicomDate As DicomDataSets
        Dim VarTmp As Variant
        If Abs(y - lngBaseYY) >= lngWidthLevelStep / 5 Or Abs(x - lngBaseXX) >= lngWidthLevelStep / 5 Then  ''''������������
            If Not blnMouseStart Then
                Me.MouseIcon = ImageListMouse.ListImages("WindowWL").Picture
                Me.MousePointer = 99
                blnMouseStart = True
            End If
            '���������ͼ��VOILUT=0ʱ���ܽ��е���
            If SelectedImage.VOILUT = 1 Then
                Set DicomAttr = SelectedImage.Attributes(&H28, &H3010)
                If VarType(DicomAttr) = vbObject Then
                    Set DicomDate = DicomAttr.Value
                    'mindray�����DRͼ��DicomDate(1).Attributes(&H28, &H3002).ValueΪ��
                    If IsNull(DicomDate(1).Attributes(&H28, &H3002).Value) Then
                        subDispWWWL SelectedImage
                    Else
                        VarTmp = DicomDate(1).Attributes(&H28, &H3002).Value
                        '��Լ���ʡ����ҽԺ��DRͼ�����ǵ�VarTmp(2)=0���Ͳ��������ַ�ʽ���޸Ĵ���λ�ˣ�ֱ���޸�VOILUT=1�Ϳ����ˡ�
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
            '�����ĵ��ڵ�λ��1
            SelectedImage.width = SelectedImage.width + (x - lngBaseXX) * lngWidthLevelStep / 5
            SelectedImage.Level = SelectedImage.Level + (y - lngBaseYY) * lngWidthLevelStep / 5
            SelectedImage.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & SelectedImage.width & "-L:" & SelectedImage.Level
'            viewer(intSelectedSerial).Refresh
            lngBaseXX = x
            lngBaseYY = y
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''[����]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (Button = cMouseUsage("104").lngMouseKey And Shift = cMouseUsage("104").lngShift And Button_miZoom And (intClickImageIndex <> 0 Or blnMouseStart)) _
        Or (Button = 4 And intMouseWheelDrag = 1) Then
        '������Ҽ���ͨ����ť���Խ������ţ�����м����קͨ�����ÿ��Խ�������
        If Abs(y - lngBaseYY) >= lngZoomStep / 5 Then                                                                 ''''���Ų�������
            If Not blnMouseStart Then
                Me.MouseIcon = ImageListMouse.ListImages("Zoom").Picture
                Me.MousePointer = 99
                blnMouseStart = True
            End If
            '���ŵĵ��ڵ�λ��0.001��
            dblZoom = SelectedImage.ActualZoom * (1 + (lngBaseYY - y) * lngZoomStep / 5 * 0.001)
            If dblZoom < 0.01 Then dblZoom = 0.01
            If dblZoom > 64 Then dblZoom = 64
            subCenterZoom SelectedImage, Viewer(intSelectedSerial), dblZoom
            If SelectedImage.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then '���±�ߵ�λ
                UpdateRuler SelectedImage, True
            End If
            Viewer(intSelectedSerial).Refresh
            lngBaseXX = x
            lngBaseYY = y
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''[����]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (Button = cMouseUsage("103").lngMouseKey And Shift = cMouseUsage("103").lngShift And Button_miCruise And (intClickImageIndex <> 0 Or blnMouseStart)) _
        Or (Button = 4 And intMouseWheelDrag = 0) Then
        '������Ҽ���ͨ����ť���Խ������Σ�����м����קͨ�����ÿ��Խ�������
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
    ''''''''''''''''''''''3D ���'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Button = cMouseUsage("106").lngMouseKey And Shift = cMouseUsage("106").lngShift And Button_mi3dCursor Then
        '��ά���״̬�£��ƶ�����ʱ�򣬷�����Ӧ��ͼ��
        If Viewer(Index).imageIndex(x, y) = 0 Then Exit Sub
        
        '���ú�����������ƶ��Ĳ���
        Call sub3DCursorMouseMove(x, y, Viewer(Index))
        Exit Sub
    End If      '����λ��״̬��״̬����3D����Ҫ��
    
    ''''''''''''''''''[���Ƕȵĵڶ�����,Button=0��ʾû������������]'''''''''''''''''
    If Button = 0 And intClickImageIndex <> 0 And blnAngle And Not (blnTextInput Or blnTextInputM) _
        And Not SelectedLabel Is Nothing Then
        With SelectedLabel
            .width = Viewer(Index).ImageXPosition(x, y) - .left
            .height = Viewer(Index).ImageYPosition(x, y) - .top
            SelectedLabelT.Text = funROIResultString(SelectedLabel, SelectedImage)  ' "Angle=" & Int(GetAngle(.left, .top, .left + .width, .top + .height, .TagObject.left, .TagObject.top) * 100) / 100  ''ע�⴦����������
            Viewer(Index).Refresh
        End With
    End If
    
    '''''''''''''''''''''''''''''''[����ע]'''''''''''''''''''''''''''''''''''''''''
    If LabelDrawing And Not (blnTextInput Or blnTextInputM) And Not SelectedLabel Is Nothing Then
        subaCorrectCursor Viewer(intSelectedSerial), SelectedImage, x, y    '����ƶ��������ͼ��Χ�����������λ��
        
        '����ǿ�ѡͼ�������Ʊ�עΪ������
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
        ''''''''''''''''''''''''''''''[����κͶ���ߵĴ���]''''''''''''''''''''''
        If SelectedLabel.LabelType = 4 Or SelectedLabel.LabelType = 5 Then
            SelectedLabel.AddPoint Viewer(Index).ImageXPosition(x, y), Viewer(Index).ImageYPosition(x, y)
        End If
        
        ''Ѫ����խ������ע,���㲢��ʾ����Ѫ�ܱڵ�λ��
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
        '''''''''''''''''''''''''''''''[��ʾ����]''''''''''''''''''''''''''''''
        With SelectedLabelT
            If Not blnAngle Then
                '''''''''' '''��ͷ'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If SelectedLabel.LabelType = doLabelArrow Then
                    .Text = " "
                    .left = SelectedLabel.left + SelectedLabel.width
                    .top = SelectedLabel.top + SelectedLabel.height
                    .AnchorX = SelectedLabel.left + SelectedLabel.width
                    .AnchorY = SelectedLabel.top + SelectedLabel.height
                '''''''''''''''''''�Զ�����'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf Button = cMouseUsage("105").lngMouseKey And Shift = cMouseUsage("105").lngShift And Button_miAutoWidthLevel Then
                    If SelectedLabel.height <> 0 And SelectedLabel.width <> 0 Then
                        ''''''''''''''������ο��Ⱥ͸߶�Ϊ���������''''''''''''''''''''''''''''''''''''''
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
                        ''''''''''''''��������Ĵ���λ����ʾ''''''''''''''''''''''''''''''''''''''
                        funAutoWinWL SelectedImage, tl, tt, tw, th, ww, wl
                        .Text = "����: " & ww & vbCrLf & "��λ:" & wl
                        SelectedLabel.Tag = ww
                        .Tag = wl
                    End If
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    .left = SelectedLabel.left + SelectedLabel.width + intTextoOffX
                    .top = SelectedLabel.top + SelectedLabel.height + intTextoOffY
                    .AnchorX = SelectedLabel.left + SelectedLabel.width / 2
                    .AnchorY = SelectedLabel.top + SelectedLabel.height / 2
                Else          '���Ǽ�ͷ���Զ������ı�ע�Ĵ���
                    If SelectedLabel.LabelType = doLabelEllipse Or SelectedLabel.LabelType = doLabelPolygon Or SelectedLabel.LabelType = doLabelRectangle Then
'                        .Text = funROIResultString(SelectedLabel)      ''''�������ͱ�ע�����ִ���
                    Else
                        .Text = Int(SelectedLabel.ROILength) & SelectedLabel.ROIDistanceUnits '      ''�������͵ı�ע���ִ���
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
            End If    'end of ��If Not blnAngle Then��
            ''''''''''''''''''''''''''''''[����κͶ���ߵĴ���]''''''''''''''''''''''
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
    '''''''''''''''''''''''''''''''''''''''''''[����j����]''''''''''''''''''''''''''''''''''''''''''''
    With Viewer(intSelectedSerial)
        If Button = cMouseUsage("101").lngMouseKey And Shift = cMouseUsage("101").lngShift And Button_miStack And blnStackStart Then  ''''
            If blnStackisFrame Then    ''''��֡ͼ����
                j = SelectedImage.Frame - intStackOffset
                SelectedImage.Frame = intStackCurrentlyImage
            Else
                '���ú�����������
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
    ''''''''''''''''''''''''''''''''''''''''''''''''''[3D������]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Button = cMouseUsage("106").lngMouseKey And Shift = cMouseUsage("106").lngShift And Button_mi3dCursor And i > 0 Then
        sub3DCursorEnd SelectedImage
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''[��������]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If blnTextInput Then
        If txtText.Visible Then     '''''���������������,����������
            txtText_KeyPress 13
        Else
            ''''''''''''''''''''''''''''''''''''''''''[��������]'''''''''''''''''''''''''''''''''''''''''''''
            With Viewer(Index)
                ''''''''''''''''''�������ֱ�ǩ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                subTextCoordinate SelectedImage, xx, Yy, lblChange   '''''���㽻������
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
    '''''''''''''''''''''''''''''''''''''[��ע������Ӧ�����Ĵ���]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If LabelDrawing _
       And ((Button = cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift) _
            Or (Button = cMouseUsage("105").lngMouseKey And Shift = cMouseUsage("105").lngShift) _
            Or (Button = cMouseUsage("201").lngMouseKey And Shift = cMouseUsage("201").lngShift And Button_miFrameSelectImage = True)) Then
        LabelDrawing = False
        SelectedLabel.XOR = False
        '''''''''''''����Ӧ����'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If blnAutoWL Then  ''''''''''''''''''''''����Ӧ����λ
            If Int(SelectedLabel.ROILength) > 2 Then
                SelectedImage.VOILUT = 0
                SelectedImage.width = Val(SelectedLabel.Tag)                ''''���ڱ�ע������Χ�ľ��κͶ�Ӧ���ּ�¼�˴���λ
                SelectedImage.Level = Val(SelectedLabelT.Tag)
                SelectedImage.Refresh False
            End If
            SelectedImage.Labels.Remove SelectedImage.Labels.Count
            SelectedImage.Labels.Remove SelectedImage.Labels.Count
            isSelectedLabel = False
            Set SelectedLabel = Nothing
        ElseIf blnFrameSelectImage Then  ''''''''''''''''''��ѡͼ��
            '��ʾͼ�񱣴�˵�
            If SelectedLabel.width <> 0 And SelectedLabel.height <> 0 Then
                ShowFrameSelectImagePopup Me
            End If
            'ɾ����ѡ�õ���ʱ��ע
            SelectedImage.Labels.Remove SelectedImage.Labels.Count
            SelectedImage.Labels.Remove SelectedImage.Labels.Count
            isSelectedLabel = False
            Set SelectedLabel = Nothing
        ElseIf SelectedLabel.LabelType = 4 And UBound(SelectedLabel.Points) = 0 Then 'ɾ������Ϊ0�Ķ���߱�ע
            '�Զ����������������Ϊֱ�ӵ��ó���Ϊ0�Ķ���ߵ�ROILength����ּ�ʱ����
            SelectedImage.Labels.Remove SelectedImage.Labels.Count
            SelectedImage.Labels.Remove SelectedImage.Labels.Count
            isSelectedLabel = False
            Set SelectedLabel = Nothing
            blnAngle = False    'Ϊ������Ƕȱ�־���������ʡ�
        ElseIf Int(SelectedLabel.ROILength) = 0 Or (SelectedLabel.width = 0 And SelectedLabel.height = 0) Then
            ''''ɾ������Ϊ0�ı�ע,��ͼ����С�󣬱�ע�Ŀ�߶�Ϊ0ʱ��ROILength���ܲ�Ϊ0������������Ҫ�ж�
            If left(SelectedLabel.Tag, 3) = "VAS" Then
                '�������խѪ�ܣ�������Ѫ�ܵ�����������
                If Mid(SelectedLabel.Tag, 4, 1) = "2" Then
                    SelectedImage.Labels(SelectedImage.Labels.Count - 4).TagObject = SelectedImage.Labels(SelectedImage.Labels.Count).TagObject
                End If
                'ɾ�������������Ϊ0�ı�ע
                SelectedImage.Labels.Remove SelectedImage.Labels.Count
                SelectedImage.Labels.Remove SelectedImage.Labels.Count
            End If
                SelectedImage.Labels.Remove SelectedImage.Labels.Count
                SelectedImage.Labels.Remove SelectedImage.Labels.Count
                isSelectedLabel = False
                Set SelectedLabel = Nothing
                blnAngle = False
                
        ElseIf Button = cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift And blnAngle Then
                '����Ƕȵĵڶ�����
                SelectedImage.Labels.Add GetNewLabel(3, SelectedLabel.left + SelectedLabel.width, SelectedLabel.top + SelectedLabel.height, 0, 0)
                Set l = SelectedImage.Labels(SelectedImage.Labels.Count)
                l.Tag = "JD2"
                l.ForeColour = lngLabelSelectedColor
                SelectedLabelT.TagObject = l
                l.TagObject = SelectedLabel
                Set SelectedLabel = SelectedImage.Labels(SelectedImage.Labels.Count)
        ElseIf SelectedLabel.LabelType = doLabelArrow Then   '��ͷ�Ĵ���[�˶εĴ����ֱ���������ֵĴ���Ƚ�����,���Կ��Ǻϲ�]
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
                subTextCoordinate SelectedImage, xx, Yy, lblChange   '''''���㽻������
                xx = (xx * SelectedImage.ActualZoom - SelectedImage.ActualScrollX) * Screen.TwipsPerPixelX + .left + Viewer(Index).width / Viewer(Index).MultiColumns * (FunImageIsX(i, Viewer(Index)) - 1)
                Yy = (Yy * SelectedImage.ActualZoom - SelectedImage.ActualScrollY) * Screen.TwipsPerPixelY + .top + Viewer(Index).height / Viewer(Index).MultiRows * (FunImageIsY(i, Viewer(Index)) - 1)
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                txtText.Move xx, Yy, lblChange.width, lblChange.height
                txtText.Visible = True
                oldTextleft = xx + lblChange.width
                txtText.SetFocus
            End With
        ElseIf Button = cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift And Button_miLabelVasMeasure Then
            'Ѫ����խ�����Ĵ���
            If intVasMeasure = 1 Then       '׼�����ڶ���Ѫ�ܣ���խѪ��
                intVasMeasure = 2
            ElseIf intVasMeasure = 2 Then
                intVasMeasure = 0
                '��ʾ�������
                Set frmVasMeasure.lblText = SelectedLabel.TagObject
                Set frmVasMeasure.f = Me
                frmVasMeasure.Show 1, Me
            End If
        ElseIf Button + cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift And Button_miLabelCadiothoracicRatio Then
            '���رȲ����Ĵ���
            If intCadioThoracicRatio = 1 Then   '׼����������
                intCadioThoracicRatio = 2
            ElseIf intCadioThoracicRatio = 2 Then
                intCadioThoracicRatio = 0
                '��ʾ�������
                Call funcGetCadioThoracicRatio(SelectedLabel, SelectedImage)
            End If
        End If
    End If      ''[��ע������Ӧ�����Ĵ���]�Ľ���
    '''''''''''''''''''''''[ʸ��״�ؽ��ƶ���ע����]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If blnMoveLabel Then
        If blnInMPR = True Then
            '''''''''''''�����Ӧ��ͼ��û�н��г�ʼ������г�ʼ��
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
    '''''������ʾ��ע�Ĳ�������
    If Not SelectedLabel Is Nothing Then
        '������ʾ�������ͱ�ע�Ĳ�����Ϣ
        If SelectedLabel.LabelType = doLabelEllipse Or SelectedLabel.LabelType = doLabelPolygon Or SelectedLabel.LabelType = doLabelRectangle Then
            Set SelectedLabelT = SelectedImage.Labels(SelectedImage.Labels.IndexOf(SelectedLabel) + 1)
            SelectedLabelT.Text = funROIResultString(SelectedLabel, SelectedImage)     ''''�������ͱ�ע�����ִ���
            
            Viewer(Index).Refresh
        End If
        '������ʾѪ����խ��������Ϣ
        If blnReSizeLabel And left(SelectedLabel.Tag, 3) = "VAS" Then
            Set frmVasMeasure.lblText = SelectedLabel.TagObject
            Set frmVasMeasure.f = Me
            frmVasMeasure.Show 1, Me
        End If
    End If
                    
    ''''''''''''''''''''''''������ͼ������ͬ��'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If (Button = cMouseUsage("102").lngMouseKey And Shift = cMouseUsage("102").lngShift And Button_miWidthLevel And blnMouseStart) _
        Or (Button = cMouseUsage("105").lngMouseKey And Shift = cMouseUsage("105").lngShift And Button_miAutoWidthLevel) Then
        Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_WINDOW)
    ElseIf (Button = cMouseUsage("103").lngMouseKey And Shift = cMouseUsage("103").lngShift And Button_miCruise And blnMouseStart) _
        Or (Button = cMouseUsage("104").lngMouseKey And Shift = cMouseUsage("104").lngShift And Button_miZoom And blnMouseStart) Then
        Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If blnMoveLabel = False And blnReSizeLabel = False And blnMouseStart = False And blnFrameSelectImage = False And blnAutoWL = False And Me.MousePointer = 0 Then
        If Button = 2 Then      '��ʾ�Ҽ��˵�
            ShowPopup Me, Viewer(Index).CurrentImage
        ElseIf Button = 1 And Shift = 2 Then         'ͨ��Ctrl+��������ѡ������
            ZLShowSeriesInfos(Index).Selected = Not ZLShowSeriesInfos(Index).Selected
            subDispframe Me, Viewer(Index)
            Viewer(Index).Refresh
        ElseIf Button = 4 Then      '�м�����л�����͹۲�ģʽ
            Call subLookOrBrowsSwitch(Me)
        End If
    End If
    
    blnMoveLabel = False
    blnReSizeLabel = False
    blnMouseStart = False
    blnAutoWL = False
    blnFrameSelectImage = False
    Me.MousePointer = 0
    blDicomDown = False                         '�ſ����
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
    blnAngle = False    '�����ǰͼ�����ı䣬����ǶȲ������
    intVasMeasure = 0    '�����ǰͼ�����ı䣬���Ѫ����խ�����������
    intCadioThoracicRatio = 0   '�����ǰͼ�����ı䣬��������رȲ������
    
    '��ʾViewer�е�ͼ��
    intImageIndex = VScro(Index).Value
    Call subShowALLImage(Me, Viewer(Index), intImageIndex, False)
    
    SelectedImageIndex = Viewer(Index).CurrentIndex
    iMove = SelectedImageIndex - MSFViewer.TextMatrix(Index, 3)
    Set SelectedImage = Viewer(Index).Images(SelectedImageIndex)
    intSelectedSerial = Index
    MSFViewer.TextMatrix(Index, 3) = SelectedImageIndex
    
    '�����ֹ�����ͬ�����Զ�����ͬ��
    If Button_miSerialManualSyn And ZLShowSeriesInfos(Index).Selected = True Then
        '�ֹ�����ͬ��
        subManualSeriesSyn Me, iMove, Index
    ElseIf ZLShowSeriesInfos(Index).ImageInfos(intImageIndex).SliceLocation <> "" And Button_miSerialPlaceInPhase Then
        '�Զ�����ͬ��
        subSerialPlaceInPhase Val(ZLShowSeriesInfos(Index).ImageInfos(intImageIndex).SliceLocation), Me
    End If
    
    '����λ�ߵ���ʾ
    If Button_miAllReferLine = True Or Button_miFLReferLine = True Or Button_miCurrentReferLine = True Then
        Call subDisplayReferLine(Viewer(Index), Me, True)    '���ݲ˵�ѡ���ʾ�������͵Ķ�λ��
    End If
    
    '�����MPR״̬���л�ͼ���Ҫˢ��MPR���ͼ����ʾ
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
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then  '''ESC�ͻس����˳�����
        txtText.Visible = False
        blnTextInput = False
        blnTextInputM = False
        txtText.FontSize = oldFontSize
        lblChange.FontSize = oldFontSize
         '''''''''���ʲô��û��������ɾ�����ӵı�־�������������ڼ�ͷ��˵���ƶ��͸ı��С��ʱ����ܳ����Լ(���󲻳�������ָ��������)
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
    ''''''''''''''''''''����Viewer�ϵ�txt�����ĸı�''''''''''''''''''''''''''''''''''''''''''''
    If SelectedImage Is Nothing Then Exit Sub
    lblChange = txtText + "  "
    ''''����ͼ��ķ�ת��������Ƿ�������
    If (SelectedImage.RotateState = doRotateNormal And (SelectedImage.FlipState = 3 Or SelectedImage.FlipState = 1)) _
        Or (SelectedImage.RotateState = doRotateLeft And (SelectedImage.FlipState = 3 Or SelectedImage.FlipState = 2)) _
        Or (SelectedImage.RotateState = doRotate180 And (SelectedImage.FlipState = 0 Or SelectedImage.FlipState = 2)) _
        Or (SelectedImage.RotateState = doRotateRight And (SelectedImage.FlipState = 0 Or SelectedImage.FlipState = 1)) Then
        txtText.left = oldTextleft - lblChange.width    ''''oldTextleft������txTtext��ʱ����д
        txtText.width = lblChange.width
    Else
        txtText.width = lblChange.width
        txtText.height = lblChange.height
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtText_LostFocus()
    txtText_KeyPress (13)  ''''��������½��㶪ʧҲ�˳�����
End Sub

Private Sub subOutPutRptImg()
'------------------------------------------------
'���ܣ����汨��ͼ��������ǰ�򿪵�����viewer���ѱ�ѡ�е�ͼ�󱣴�ɱ���ͼ��
'      ���汨��ͼ��ʱ��ֻ�ܰѱ���ͼ���浽��ǰ���򿪵�ͼ�����ڵ�FTPĿ¼��
'������
'���أ�ֱ�Ӱѱ�ѡ�е�ͼ�񱣴�ɱ���ͼ
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
    
    '��ȡ����ͼ����������Ϣ����
    lngImgContainInfo = (Val(zlDatabase.GetPara("����ͼ����������Ϣ", glngSys, 1289, 1)))
    
    '����ѡ�е�ͼ����ӵ�ͼ����
    For i = 1 To ZLShowSeriesInfos.Count
        iImageIndex = 1
        For j = 1 To ZLShowSeriesInfos(i).ImageInfos.Count
            If ZLShowSeriesInfos(i).ImageInfos(j).blnSelected = True Then
                Set im = Nothing
                '�����ж�ͼ���Ƿ��Ѿ�װ�أ�����Ѿ�װ�أ����ҵ����ͼ����ʾ���������û��װ�أ���װ�ظ�ͼ��
                If ZLShowSeriesInfos(i).ImageInfos(j).blnDisplayed = False Then
                    funcAddAImageA Viewer(i), j
                End If
                
                '����ͼ�������
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
                        '�����ݿ��ж�ȡ���UID
                        strSQL = "select ���UID FROM Ӱ��������  where ����UID =[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(im.SeriesUID))
                        If rsTmp.RecordCount = 0 Then
                            strStudyUID = PstrCheckUID  '��Ĭ��ֵ
                        Else
                            strStudyUID = rsTmp!���UID
                        End If
                    End If
                   
                   
                    '������ͼƬ���ĽǱ�ע
                    If lngImgContainInfo = 0 Then subInitImageLabels intSelectedSerial, 1, im, False
                    
    
                    Set imTmp = im.Capture(False)
                    
                    '����ʾͼƬ���ĽǱ�ע
                    If lngImgContainInfo = 0 Then subInitImageLabels intSelectedSerial, 1, im, True
                    
                    imTmp.VOILUT = 0
                    '����Ҫ����InstanceUID���ڱ��汨��ͼ��ʱ����������һ��
                    imTmp.SeriesUID = im.SeriesUID
                    imTmp.StudyUID = strStudyUID        '��һ���ͱ�֤��ͼ��ļ��UID�����ݿ��е�һ�£���˿���˳�����浽��һ���ļ�¼��
                    
                    imgs.Add imTmp
                End If
            End If
        Next j
    Next i
    
    '����ͼ��ɱ���ͼ
    If imgs.Count > 0 Then
        On Error Resume Next
        SaveImages imgs, 1
        If err <> 0 Then
            MsgBox "����ͼ�������", vbExclamation, gstrSysName
        End If
    Else
        MsgBox "û�б�ѡ�е�ͼ����ѡ��ͼ����ٱ���", vbExclamation, gstrSysName
    End If
End Sub

Private Sub subCloseSeries(Optional blnNotify As Boolean = True)
    If intSelectedSerial <> 0 Then
        If blnNotify Then
            If MsgBox("�����Ҫ�رմ�ͼ��?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Sub
        End If
            Call subUnloadViewer(intSelectedSerial, Me)
    End If
End Sub

Public Sub subKillPicture(Optional ByVal bSilent As Boolean = False)
'------------------------------------------------
'���ܣ�ɾ��ȫ��ͼ�񣬲���ж�����е�Viewer�͹���������ʼ��ͼ�����������Լ���صĲ���
'������
'���أ�
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    
    If MSFViewer.Rows < 1 Then Exit Sub
    If bSilent = False Then
        If MsgBox("ȷ��Ҫɾ������ͼ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    '�����MPR״̬�У�����ʾ�Ƿ񱣴�MPR���
    If blnInMPR = True Then
        If funViewerMPR(Me, bSilent) = False Then Exit Sub
    End If
    
    'ɾ��ͼ��
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
    Set ZLSeriesInfos = Nothing '������ռ��ϣ��ٶȸ��ӿ�
    Set ZLSeriesInfos = New Collection
    Set ZLShowSeriesInfos = Nothing
    Set ZLShowSeriesInfos = New Collection
    
    intSelectedSerial = 0
    Set SelectedImage = Nothing
    oldSelectedImageIndex = 0
    oldSelectedSerial = 0
    SelectedImageIndex = 0
    intClickImageIndex = 0
    MSFViewer.Rows = 1  '���ԭͼ���б��¼������
    Me.txtText.Visible = False
    Set SelectedLabel = Nothing
    '������ʾ����ͼ
    Call subShowMiniImages(Me)
End Sub

Function funcROIHistogram(LLAA As DicomLabel) As Boolean
'------------------------------------------------
'���ܣ������������͵ı�ע����ֱ��ͼ���������Ϊ��Ҫ��ֱ��ͼ�ı�ע��
'������LLAA--��Ҫ��ֱ��ͼ�ı�ע��
'���أ�True--�ɹ�����ֱ�ӻ���ֱ��ͼ���壻False--�����ע�����������ͱ�ע��ʧ�ܡ�
'�ϼ���������̣�frmViewer.Viewer_DblClick��
'�¼���������̣�mdlPublic.Max7InArray
'���õ��ⲿ������frmHistogram����
'�����ˣ��ƽ�
'------------------------------------------------
    
    '�жϱ�ע�Ƿ��������͵ı�ע�������ǣ��򲻻�ֱ��ͼ��������False�����ǣ���ֱ�ӻ���ֱ��ͼ����������True��
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
    '''''��ֱ��ͼ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim xx() As Long
    Dim lMin As Long
    Dim lMax As Long
    Dim lCount As Long
    
    ''''���������ת������ά��Ϊ���ص㣬����Ϊ�Ҷ�ֵ�����飬ת��Ϊά��Ϊ�Ҷ�ֵ������Ϊ�ûҶ�ֵ����������
    lMax = LLAA.ROIMax      '��ʱʹ��lMax��lMin
    lMin = LLAA.ROIMin
    ReDim xx(lMax - lMin + 1)
    '''''��ʼ������XX'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For lCount = 1 To (lMax - lMin + 1)
        xx(lCount) = 0
    Next
    ''''��������WHAT''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For lCount = 1 To UBound(WHAT)
        xx(WHAT(lCount) - lMin + 1) = xx(WHAT(lCount) - lMin + 1) + 1
    Next
    ''''��дֱ��ͼ�����Ϣ'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    frmHistogram.lblStart.Caption = LLAA.ROIMin
    frmHistogram.lblEnd.Caption = LLAA.ROIMax
    Max7InArray xx, lMax, lMin
    frmHistogram.Text1 = xx(lMin)
    frmHistogram.Text2 = xx(lMax)
    frmHistogram.Text3 = LLAA.ROIMin + lMin
    frmHistogram.Text4 = LLAA.ROIMin + lMax
    '''''��ͼ��''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
'''' ��ʾ�߱�ע�ĻҶȷֲ�ֱ��ͼ���κ���ֱ����ʾ�Ҷȷֲ�ͼ���壬
'''' �����ֵΪͼ��������ı�ע��ֱ�߱�ע�����߱�ע��������ֵΪ�Ƿ�ɹ�
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function funcDrawGreyDistribute(img As DicomImage, la As DicomLabel) As Boolean
    funcDrawGreyDistribute = False
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (img Is Nothing) Or (la Is Nothing) Then Exit Function
    
    '�жϱ�ע�Ƿ�ֱ�߻�����  '�жϱ�ע�Ƿ�����ͼ���� ���������������Ƴ�����
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
    ''��ȡֱ���ϻҶ�ֵ����ŵ�������
    Dim aGrey() As Integer
    Dim beginx As Integer
    Dim beginy As Integer
    Dim endx As Integer
    Dim endy As Integer
    Dim i As Integer
    If funGetLinePoints(img, la, aGrey(), beginx, beginy, endx, endy) = False Then Exit Function
    '�Ҷȷֲ�ͼ����ʾ��ʼ�㣬����������
    frmHistogram.lblStart.Caption = "��㣺(" & CStr(beginx) & "," & CStr(beginy) & ")"
    frmHistogram.lblEnd.Caption = "�յ㣺(" & CStr(endx) & "," & CStr(endy) & ")"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '�Ҷȷֲ�ͼ����ʾֱ�߾��룬
    '�Ҷȷֲ�ͼ��x����Ϊֱ�ߴ����ң����ϵ��µĵ㣬y����Ϊ�õ�ĻҶ�ֵ
    '����״ͼ
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
    '���ܣ�                                     ����ť���º͵���Ĺ���
    '������
    '       Serial_ID                           ����Ϊѡ�е�����ʽ
    '���أ�
    '------------------------------------------------
    '����
    '��������
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_Text, , True).Checked = False                          '����
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_Arrowhead, , True).Checked = False                     '��ͷ
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_Ellipse, , True).Checked = False                       '��Բ
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_Angle, , True).Checked = False                         '�Ƕ�
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_Curve, , True).Checked = False                         '����
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_Area, , True).Checked = False                          '����
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_BeeLine, , True).Checked = False                       'ֱ��
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_Rect, , True).Checked = False                          '����
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_VasMeasure, , True).Checked = False                    'Ѫ����խ����
    ComToolBar.Item(ToolBar_Scale).FindControl(, ID_Active_Lable_CadioThoracicRatio, , True).Checked = False            '���ر�
    '����˵�
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
    '����
    '��������
    ComToolBar.Item(ToolBar_Comm).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow, , True).Checked = False       '�ֶ�����
    ComToolBar.Item(ToolBar_Comm).FindControl(, ID_Active_Cruise, , True).Checked = False                              '����
    ComToolBar.Item(ToolBar_Comm).FindControl(, ID_Active_Zoom, , True).Checked = False                                '����
    ComToolBar.Item(ToolBar_Comm).FindControl(, ID_Active_Shuttle, , True).Checked = False                             '����
    '����˵�
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_AdjustWindow_AutoAdjustWindow, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Cruise, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Zoom, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_Shuttle, , True).Checked = False
    
    'ƽ��
    '��������
    ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_3DLine, , True).Checked = False                 '��ά���
    ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Tool_ArrowyCoronaryReset, , True).Checked = False                   'ʸ��״�ؽ�
    '����˵�
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_PointingLine_3DLine, , True).Checked = False
    ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Tool_ArrowyCoronaryReset, , True).Checked = False
    
    '�������
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
    
    blnAngle = False    '�����ǰͼ�����ı䣬����ǶȲ������
    intVasMeasure = 0   '��Ѫ����խ������ע�ı������
    intCadioThoracicRatio = 0   '�����رȲ����������
    
    '�������������
    If ButtomID = ID_Active_Lable_Text Or ButtomID = ID_Active_Lable_Arrowhead Or ButtomID = ID_Active_Lable_Ellipse _
        Or ButtomID = ID_Active_Lable_Angle Or ButtomID = ID_Active_Lable_Curve Or ButtomID = ID_Active_Lable_Area _
        Or ButtomID = ID_Active_Lable_BeeLine Or ButtomID = ID_Active_Lable_Rect Or ButtomID = ID_Active_Lable_VasMeasure _
        Or ButtomID = ID_Active_Lable_CadioThoracicRatio Then
        
        ComToolBar.Item(ToolBar_Scale).FindControl(, ButtomID, , True).Checked = True
        ComToolBar.Item(ToolBar_Menu).FindControl(, ButtomID, , True).Checked = True
        
    End If
    
    '����������
    If ButtomID = ID_Active_AdjustWindow_HandAdjustWindow Or ButtomID = ID_Active_Cruise _
         Or ButtomID = ID_Active_Zoom Or ButtomID = ID_Active_Shuttle Then
         
         ComToolBar.Item(ToolBar_Comm).FindControl(, ButtomID, , True).Checked = True
         ComToolBar.Item(ToolBar_Menu).FindControl(, ButtomID, , True).Checked = True
    End If
    
    'ƽ�湤����
    If ButtomID = ID_Active_PointingLine_3DLine Or ButtomID = ID_Tool_ArrowyCoronaryReset Then
    
        ComToolBar.Item(ToolBar_Plane).FindControl(, ButtomID, , True).Checked = True
        ComToolBar.Item(ToolBar_Menu).FindControl(, ButtomID, , True).Checked = True
    End If
    
    '����ֻ���²˵��������¹�����
    If ButtomID = ID_Active_AdjustWindow_AutoAdjustWindow Then
        ComToolBar.Item(ToolBar_Menu).FindControl(, ButtomID, , True).Checked = True
    End If
    
    ComToolBar.RecalcLayout
End Sub

Private Sub subShowScale(lngButtonID As Long)
    '------------------------------------------------
    '���ܣ��������ű�����ť��ѡ�кͲ�ѡ��״̬
    '������
    '       lngButtonID   ��ѡ�еİ�ťID
    '���أ�
    '------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim cbrControl As CommandBarControl
    
    On Error GoTo err
    
    '���Ȱ��������Ű�ť���ó�δѡ��
    
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
    
    '����ǰ�����İ�ť���ó�ѡ��
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
    '���ܣ�                                     ����������Ҽ��Ƿ��а��µ��о͵����µİ���
    '������
    '       LeftOrRigth                         1Ϊ���2Ϊ�Ҽ�
    '���أ�
    '------------------------------------------------
    Dim i As Integer
    Dim cbrControl As CommandBarControl
        
    '����
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

    '����
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
    
    '�ü�
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
    
    '��ѡ
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

    '�ֶ�����
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
    
    '����Ӧ����
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
    
    '��ά���
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
    
    '����
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
    
    '����
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
    
    '��ͷ
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
    
    '��Բ
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
    
    '�Ƕ�
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
    
    '����
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
    
    '����
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
    
    'ֱ��
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
    
    'Ѫ����խ����
    If cMouseUsage("1").lngMouseKey = LeftOrRigth And ID_Active_Lable_VasMeasure = BouttomID Then
        intVasMeasure = 0   '��Ѫ����խ������ע�ı������
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
    
    '���رȲ���
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
    
    '����
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
'���ܣ���Ƭ��ӡ
'������ f--���д�ӡ�Ĵ��塣
'       blnShowForm --- �Ƿ���ʾ��ӡ����
'       intAddType  --  ���ͼ��ķ�ʽ��1-������У�2-��ӵ�ǰͼ��3-�����ѡͼ;4-����������
'       intInterval -- �����ӡ�ļ��������intAddType=4ʱʹ��
'       blnStartOdd -- True ������False ż���𣬵�intAddType=4ʱʹ��
'���أ�True--�ɹ��򿪴�ӡ���壻False -- ʧ�ܡ�
'2009��
'------------------------------------------------
    
    On Error GoTo err
    
    '���жϽ�Ƭ��ӡ���������Ƿ񳬹���ɵ�����
    If (cDICOMPrinter.Count > gint��Ƭ��ӡ�� And gint��Ƭ��ӡ�� <> -1) Or gint��Ƭ��ӡ�� = 0 Then
        Call MsgBox(LOGIN_TYPE_��Ƭ��ӡ�� & "�������������������" & gint��Ƭ��ӡ�� & "�������������Ӧ����ϵ��", vbOKOnly, gstrSysName)
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
    
    '���ͼ��ɹ���ʹ��������ʾ,��ʾ��Ƭ���ڵ�ʱ�򣬲���ʾ����
    If blnShowForm = False Then
        Call PrintFilmBeep(1)
    End If
    
    '    ���Ͻػ���Ϣ��hook�����ܷ���mfrmFilm��load�¼�
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
    If intSelectedSerial <> 0 Then         '��ǰ��ѡ�е�����
        If Me.MSFViewer.TextMatrix(intSelectedSerial, 1) Then   '��������ͼ��
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
        MsgBox "��ѡ��ͼ��������ٱ���", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    'ͼ�������=1281
    iMax = 1281
    
    '��ȡ����ͼ����������Ϣ����
    lngImgContainInfo = (Val(zlDatabase.GetPara("����ͼ����������Ϣ", glngSys, 1289, 1)))
    
    '����label����ȡ����ѡ�е�ͼ��
    'ͼ��λ��,�ڰ�ͼ��Ϊ1����ɫͼ��Ϊ3
    iPlane = 1
    If Not IsNull(img.Attributes(&H28, &H4).Value) And img.Attributes(&H28, &H4).Exists Then
        If img.Attributes(&H28, &H4).Value = "RGB" Then
            iPlane = 3
        End If
    End If
    
    'ͼ����λ��
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
    
    '���ƽ��ͼ��Ĵ�С��300*300֮��
    If (iRight - iLeft) > iMax Or (iBottom - iTop) > iMax Then
        dblZoom = iMax / (iRight - iLeft)
        If dblZoom > iMax / (iBottom - iTop) Then dblZoom = iMax / (iBottom - iTop)
    Else
        dblZoom = 1
    End If
    
    img.Labels(img.Labels.Count).Visible = False
    img.Labels(img.Labels.Count - 1).Visible = False
    
    '������ͼƬ���ĽǱ�ע
    If lngImgContainInfo = 0 Then subInitImageLabels intSelectedSerial, 1, img, False
    
    If (img.RotateState = doRotateLeft And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipBoth) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipHorizontal) Then
        'X����Ե�
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, img.sizeX - iRight, img.sizeX - iLeft, iTop, iBottom)
    ElseIf (img.RotateState = doRotateLeft And img.FlipState = doFlipBoth) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipVertical) Then
        'Y����Ե�
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, img.sizeY - iBottom, img.sizeY - iTop)
    ElseIf (img.RotateState = doRotateRight And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateLeft And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipBoth) Then
        'X��Y����Ե�
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, img.sizeX - iRight, img.sizeX - iLeft, img.sizeY - iBottom, img.sizeY - iTop)
    Else
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, iTop, iBottom)
    End If
    
    
    '����ʾͼƬ���ĽǱ�ע
    If lngImgContainInfo = 0 Then subInitImageLabels intSelectedSerial, 1, img, True
    
    '����Ҫ����InstanceUID���ڱ��汨��ͼ��ʱ����������һ��InstanceUID
    imgResult.SeriesUID = img.SeriesUID
    
    '�����ݿ��ж�ȡ���UID
    strSQL = "select ���UID FROM Ӱ��������  where ����UID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(imgResult.SeriesUID))
    If rsTemp.RecordCount = 0 Then
        MsgBox "�����ݿ����޷��鵽��ͼ���ⲿͼ���ܱ���ɱ���ͼ��", vbExclamation, gstrSysName
        Exit Sub
    Else
        strStudyUID = rsTemp!���UID
    End If
    imgResult.StudyUID = strStudyUID
    
    imgs.Add imgResult
    '�ѽ��ͼ�񱣴�ɱ���ͼ
    '����ͼ��ɱ���ͼ
    If imgs.Count > 0 Then
        On Error Resume Next
        SaveImages imgs, 1
        If err <> 0 Then
            MsgBox "����ͼ�������", vbExclamation, gstrSysName
        End If
    Else
        MsgBox "û�б�ѡ�е�ͼ����ѡ��ͼ����ٱ���", vbExclamation, gstrSysName
    End If
End Sub

Public Function funcSwapSeries(intViewerIndex As Integer, intSeriesIndex As Integer, Optional blnShowLast As Boolean = False) As Boolean
'------------------------------------------------
'���ܣ� ��������,��intSeriesIndexָ��������е�ͼ�����viewer(intViewerIndex)�е�ͼ��
'������ intViewerIndex--Viewer������
'       intSeriesIndex--ͼ�����ڵ���������
'       blnShowLast --- ��ѡ�������Ƿ���ʾ���һ��ͼ
'���أ������ɹ����򷵻�True�����򷵻�False
'ʱ�䣺2009-7
'------------------------------------------------
    Dim oneImageInfo As clsImageInfo
    Dim i As Integer
    Dim blnSelected As Boolean
    
    On Error GoTo err
    
    '�����ж��Ƿ���ʸ��״�ؽ�״̬�����ұ������Viewer��ʸ��״�ؽ������л��߽�����У�����������˳�ʸ��״λ�ؽ�
    If blnInMPR = True And (ZLMPRCube(1).intViewerIndex = intViewerIndex Or ZLMPRCube(2).intViewerIndex = intViewerIndex _
        Or ZLMPRCube(3).intViewerIndex = intViewerIndex) Then
            Call funViewerMPR(Me)
            Exit Function
    End If
    
    '���ȸ���intViewerIndex���ҵ������е�index
    '�Ѿ������е�ͼ�����
    Viewer(intViewerIndex).Images.Clear
    
    '��¼�����е�ѡ��״̬
    blnSelected = ZLShowSeriesInfos(intViewerIndex).Selected
    
    '��ZLShowSeriesInfos�����������滻������
    Call funCopySeriesInfo(ZLSeriesInfos(intSeriesIndex), ZLShowSeriesInfos(intViewerIndex))
    
    '���������лָ�ѡ��״̬
    ZLShowSeriesInfos(intViewerIndex).Selected = blnSelected
    
    '����ͼ��
    Set ZLShowSeriesInfos(intViewerIndex).ImageInfos = Nothing
    For i = 1 To ZLSeriesInfos(intSeriesIndex).ImageInfos.Count
        Set oneImageInfo = funGetNewImageInfo
        Call funCopyImageInfo(ZLSeriesInfos(intSeriesIndex).ImageInfos(i), oneImageInfo)
        ZLShowSeriesInfos(intViewerIndex).ImageInfos.Add oneImageInfo
    Next i
    
    'ͼ������,����ͼ������ݿ��ж�ȡ���ǰ���ͼ�������ģ�����Ҫ����������ʾ���е����򷽷�
    Call subSortImages(Me, intViewerIndex, funGetImageSort(ZLSeriesInfos(intSeriesIndex).strModality))

    '�������ͨ��˫���ı�ͼ����ʾ���֣����ﰴ���²�����ʾͼ��
    If MSFViewer.TextMatrix(intSelectedSerial, 5) > 1 Or MSFViewer.TextMatrix(intSelectedSerial, 6) > 1 Then
        Viewer(intViewerIndex).MultiColumns = MSFViewer.TextMatrix(intSelectedSerial, 5)
        Viewer(intViewerIndex).MultiRows = MSFViewer.TextMatrix(intSelectedSerial, 6)
    End If
    
    '�������е�ͼ����ʾ��Viewer��
    If blnShowLast = True Then  '��ʾ���һ��ͼ
        Call subShowALLImage(Me, Viewer(intViewerIndex), ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count, False)
    Else    '��ʾ��һ��ͼ
        Call subShowALLImage(Me, Viewer(intViewerIndex), 1, False)
    End If
    
    '������Viewer��Tag�͹�����
    Viewer(intViewerIndex).Tag = intSeriesIndex
    
    blnVscroInvoked = True
    Call subDisplayScrollBar(intViewerIndex, Me, False)
    blnVscroInvoked = False
    
    '�����ѡ���������е�״̬����Ŀǰ��������Ϊ��ǰ���У���subDispframeʹ��
    If isSelectAllSerial Then intSelectedSerial = intViewerIndex
    
    'ͼ����ʾ������Viewer�еı�ע��ͼ������½ǵ�ѡ���ǵ�
    Call subDispframe(Me, Viewer(intViewerIndex))
    
    '�������к����õ�ǰѡ�е�����
    SelectedImageIndex = Viewer(intViewerIndex).CurrentIndex
    Set SelectedImage = Viewer(intViewerIndex).Images(SelectedImageIndex)
    intSelectedSerial = intViewerIndex
    MSFViewer.TextMatrix(intViewerIndex, 3) = SelectedImageIndex
    
    '������иı䣬�򴥷��¼�
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
'���ܣ� ѡ��intViewerIndexָ����Viewer������MSFViewer������Ϣ��������ʾԭ�����к������е����
'       �������ô���λ�˵���ͬʱ����λ�ߣ�����ͬ���Ȳ���
'������ intViewerIndex--��ѡ���Viewer��Index
'       intClickImageIndex--��ѡ���ͼ���Index
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim blnSeriesChanged As Boolean
    Dim blnImageChanged As Boolean
    Dim intOldSelectedSeries As Integer
    Dim CmdControl As CommandBarControl
    
    'ͼ�����ı�����з����ı䣬Ӧ���в�ͬ�Ĵ�����
    
    On Error GoTo err
    
    If Viewer(intViewerIndex).Images.Count = 0 Then Exit Sub
    
    blnSeriesChanged = (intSelectedSerial <> intViewerIndex)
    'blnImageChanged���������ԣ���ͼ�񲼾�Ϊ1*1��ʱ���϶�������ʱ����ı�SelectedImageIndex��ֵ�������жϴ���
    blnImageChanged = (SelectedImageIndex <> intClickImageIndex)
    
    If blnImageChanged = True Or blnSeriesChanged = True Then
        '������������л��߸����˱�ѡ���ͼ������ԭ����һЩ���
        blnAngle = False        '�����ǰͼ�����ı䣬������ǶȲ������
        intVasMeasure = 0       '�����ǰͼ�����ı䣬���Ѫ����խ�����������
        intCadioThoracicRatio = 0   '�����ǰͼ�����ı䣬������رȲ����������
    End If
    
    '��¼���л��߶�����ʾʱ��ͼ����Ϣ
    '5�������к�����ʾͼ����Ŀ��6��������������ʾͼ����Ŀ��7�������е�ǰ��ʾ��һ��ͼ����ţ�8�������е�ǰ��ʾѡ��ͼ�����
    If Viewer(intViewerIndex).MultiColumns > 1 Or Viewer(intViewerIndex).MultiRows > 1 Then
        MSFViewer.TextMatrix(intViewerIndex, 5) = Viewer(intViewerIndex).MultiColumns
        MSFViewer.TextMatrix(intViewerIndex, 6) = Viewer(intViewerIndex).MultiRows
        MSFViewer.TextMatrix(intViewerIndex, 7) = Viewer(intViewerIndex).CurrentIndex
        MSFViewer.TextMatrix(intViewerIndex, 8) = intClickImageIndex
    End If
    
    '����ѡ�е�ͼ��
    If intClickImageIndex <> 0 Then
        '3����ǰѡ���ͼ��ţ�4����ǰѡ���ͼ���ڵڼ�֡
        MSFViewer.TextMatrix(intViewerIndex, 3) = intClickImageIndex
        MSFViewer.TextMatrix(intViewerIndex, 4) = Viewer(intViewerIndex).Images(intClickImageIndex).Frame
        
        Set SelectedImage = Viewer(intViewerIndex).Images(intClickImageIndex)
        SelectedImageIndex = intClickImageIndex
        intSelectedSerial = intViewerIndex
        
        '��ʾ��ǰ��ʾ��ͼ��ź͵�ǰ���е�ͼ������
        sbStatusBar.Panels(3).Text = "��ǰͼ��" & Viewer(intViewerIndex).Images(intClickImageIndex).Tag _
            & "  ����Ϊ��" & ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
            
        '����λ�ߣ����ݲ˵�ѡ���ʾ�������͵Ķ�λ��
        If Button_miAllReferLine = True Or Button_miFLReferLine = True Or Button_miCurrentReferLine = True Then
            Call subDisplayReferLine(Viewer(intViewerIndex), Me, False)
        End If
    End If
    
    '��¼�ɵ����к�
    If blnSeriesChanged Then
        intOldSelectedSeries = intSelectedSerial
        intSelectedSerial = intViewerIndex
        If intOldSelectedSeries = 0 Then
            intOldSelectedSeries = intSelectedSerial
        ElseIf intOldSelectedSeries >= Viewer.Count Then
            intOldSelectedSeries = intSelectedSerial
        End If
        If intOldSelectedSeries <> intSelectedSerial Then
            '��������еı߿�
            subDispframe Me, Viewer(intOldSelectedSeries)
            Viewer(intOldSelectedSeries).Refresh
        End If
    End If
    
    '���������еı߿�
    subDispframe Me, Viewer(intViewerIndex)
    
    '�������Ĵ���λ�˵�
    Call subSetWidthLevelF(Viewer(intViewerIndex).Images(1), Me)
    '��������ͼ���˾��˵�
    Call subSetFilterF(Viewer(intViewerIndex).Images(1), Me)
    '����ͼ������˵���ѡ��
    Call subSetImageFortF(Me)
    
    '�����Ƿ���Ҫˢ��Viewer
    Viewer(intViewerIndex).Refresh
    
    '�Ѿ�ѡ����һ�����У���ȡ������ȫѡ
    isSelectAllSerial = False
    Set CmdControl = ComToolBar.Item(ToolBar_Comm).FindControl(, ID_Active_Select_SelectAllSerial)
    CmdControl.Checked = False
    
    '������иı䣬�������иĶ��¼�
    If blnSeriesChanged = True Then
        If Not SelectedImage Is Nothing Then
            RaiseEvent AfterSeriesChanged(SelectedImage.StudyUID, SelectedImage.SeriesUID)
        End If
    End If
    
    '�����MPR״̬���л�ͼ���Ҫˢ��MPR���ͼ����ʾ
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
'���ܣ� ����һ��Viewer�����з���intSeriesIndex���е�ͼ�񣬲��ڷ���picViewer�е�lngX��lngY����λ��
'������ intSeriesIndex--ZLSeriesInfos�����е�����
'       intRow -- ��Viewer���ڵ�����
'       intCol -- ��Viewer���ڵ�����
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim intCurrentViewer As Integer
    Dim intViewerIndex As Integer
    
    '�����ж����λ���Ƿ���Viewer
    intViewerIndex = (intRow - 1) * intCountX + intCol
    If intViewerIndex >= Viewer.Count Then
        '�����������ͼ����Ҫ�´���һ��Viewer
        intCurrentViewer = funcCeateAViewer(intSeriesIndex, Me)
    Else
        '�����ڵ�Viewer
        Call funcSwapSeries(intViewerIndex, intSeriesIndex)
        intCurrentViewer = intViewerIndex
    End If
    
    '�ڷ����Viewer�����ù�����
    Call subPlaceAViewer(Me, intCurrentViewer, intRow, intCol)
End Sub

Private Sub sub3DCursorStart(thisImage As DicomImage)
'------------------------------------------------
'���ܣ� ��ʼ��ά���
'������ thisImage ---- ��ά������ʱ������ڵ�ͼ��
'���أ� �Ƿ�ɹ�
'ʱ�䣺2009-7
'------------------------------------------------
    Dim i As Integer
    Dim intViewerIndex As Integer
    Dim intImageIndex As Integer
    Dim sourceFrameOfReferenceUID As String
    Dim destFrameOfReferenceUID As String
    Dim intCurrentImage As Integer
    Dim imgTemp As New DicomImage
    Dim labelRef As DicomLabel
    
    '�����ǰͼ��û�вο�֡UID���򲻽�����ά������
    sourceFrameOfReferenceUID = GetImageAttribute(thisImage.Attributes, ATTR_�ο�֡UID)
    If sourceFrameOfReferenceUID = "" Then
        blnIn3dCursor = False
        Exit Sub
    End If
    
    'ѭ����ǰ��ʵ��Viewer���ҵ����뱾����ά����Viewer��Index
    On Error GoTo err
    
    '''''''''''''''''''''''''''������ֱ��ݱ���''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ReDim obj3dImage(Viewer.Count)
    ReDim int3dImageIndex(Viewer.Count)
    ReDim int3dCurrentlyImage(Viewer.Count)
    
    For i = 1 To Viewer.Count - 1
        If i <> intSelectedSerial And Viewer(i).Images.Count > 0 Then 'ͼ�����ڵ�Viewer������
            intImageIndex = Val(MSFViewer.TextMatrix(i, 3))      '3=��ǰѡ���ͼ���
            destFrameOfReferenceUID = GetImageAttribute(Viewer(i).Images(intImageIndex).Attributes, ATTR_�ο�֡UID)
            If sourceFrameOfReferenceUID = destFrameOfReferenceUID Then
                'ֻ�����������е�ͼ��Ĳο�֡UID��ͬ��������ά���
                '''''''''''''''''���ݵ�ǰͼ��'''''''''''''''''''''''''''''''''''
                Set obj3dImage(i) = New DicomImage
                Set obj3dImage(i) = Viewer(i).Images(intImageIndex)
                int3dImageIndex(i) = intImageIndex
                int3dCurrentlyImage(i) = Viewer(i).CurrentIndex
                
                'ѭ����������е�����ͼ�񣬻���λ��
                For intCurrentImage = 1 To ZLShowSeriesInfos(i).ImageInfos.Count
                    '��д����ͼ������ݣ�Ȼ����㶨λ��
                    Call subWriteRefLineImage(imgTemp, intCurrentImage, Viewer(i))
                    If sourceFrameOfReferenceUID = ZLShowSeriesInfos(i).ImageInfos(intCurrentImage).FrameOfReferenceUID Then
                        '����λ��
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
'���ܣ� ������ά���
'������ thisImage ---- ��ά������ʱ������ڵ�ͼ��
'���أ� ��
'ʱ�䣺2009-7
'------------------------------------------------
    Dim k As Integer, i As Integer, v As DicomViewer
    ''''''''''''''''''[ɾ�����еĶ�λ��]''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    subDeleteAppointLabel thisImage, "3D"
    thisImage.Refresh False
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If blnIn3dCursor = False Then Exit Sub
    
    For Each v In Viewer
        i = v.Index
        If i <> intSelectedSerial And i <> 0 And int3dImageIndex(i) <> 0 Then     '���ϵ�в�������ά���
            k = MSFViewer.TextMatrix(i, 3)        '��ǰѡ���ͼ���
            Viewer(i).Images.Add obj3dImage(i)
            Viewer(i).Images.Move Viewer(i).Images.Count, k
            subLabelCopyRebuild obj3dImage(i), v.Images(k)           ''''�����ע����������
            Viewer(i).Images.Remove k + 1
'            viewer(i).Refresh
            
            If VScro(i).Visible = True Then VScro(i).Value = int3dImageIndex(i)
        End If
    Next
End Sub

Private Sub subStackMouseMove(lngDeltaY As Long)
'------------------------------------------------
'���ܣ� ��괩���ƶ����ʱ�Ĳ���
'������ lngDeltaY ---- �����Viewer�е�Y�����λ������ͨ��
'���أ� ��
'ʱ�䣺2009-7
'------------------------------------------------
    Dim objTempImage As DicomImage
    Dim intNewFrame As Integer
    
    On Error GoTo err
    If Not blnMouseStart Then                           ''''���û�п�ʼ���������ڿ�ʼ����
        Me.MouseIcon = ImageListMouse.ListImages("Stack").Picture
        Me.MousePointer = 99
        blnMouseStart = True
        intStackOffset = SelectedImageIndex - Viewer(intSelectedSerial).CurrentIndex    '��¼��ǰͼ���CurrentIndex֮��ľ���
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If SelectedImage.FrameCount > 1 Then   ''''��֡ͼ����
            intStackCurrentlyImage = SelectedImage.Frame
            blnStackisFrame = True
        Else
            blnStackisFrame = False
            '��¼����ǰViewer��CurrentIndex�͵�ǰͼ��
            Set SelectedLabel = Nothing
            intStackCurrentlyImage = Viewer(intSelectedSerial).CurrentIndex
            Set objStackOldImage = Viewer(intSelectedSerial).Images(SelectedImageIndex)
            intStackIndex = Viewer(intSelectedSerial).Images(SelectedImageIndex).Tag
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If blnStackisFrame And SelectedImage.FrameCount <> 1 Then            ''''����ͼ��
        intNewFrame = SelectedImage.Frame + (lngDeltaY) / lngStackStep
        If intNewFrame <= 0 Then intNewFrame = 1
        If intNewFrame > SelectedImage.FrameCount Then intNewFrame = SelectedImage.FrameCount
        SelectedImage.Frame = intNewFrame
    ElseIf Not blnStackisFrame And ZLShowSeriesInfos(intSelectedSerial).ImageInfos.Count <> 1 Then  ''''����ͼ��
        '������ͼ���index
        intStackIndex = intStackIndex + (lngDeltaY) / lngStackStep
        If intStackIndex <= 0 Then intStackIndex = 1
        If intStackIndex > ZLShowSeriesInfos(intSelectedSerial).ImageInfos.Count Then intStackIndex = ZLShowSeriesInfos(intSelectedSerial).ImageInfos.Count
        '��ָ��λ�õ�ͼ����ӵ�Viewer��
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
'���ܣ� ��ά��꣬�ƶ����ʱ�Ĳ���
'������ x ---- �����Viewer�е�X����λ��
'       y ---- �����Viewer�е�Y����λ��
'       thisViewer ---- ������ڵ�Viewer
'���أ� ��
'ʱ�䣺2009-7
'------------------------------------------------
    Dim ls As DicomLabels
    Dim l As DicomLabel
    Dim img As DicomImage
    Dim str3DMoveTag As String  '��¼��ǰ��ά�������ı�ע��һЩ����
    Dim objTempImage As DicomImage  '��ά����õ���ʱͼ��
    Dim j As Integer, k As Integer, ii As Integer
    
    On Error GoTo err
    If blnIn3dCursor = False Then Exit Sub
    
    Set ls = thisViewer.LabelHits(x, y, False, False, True)  'ls�����б������ı�ע�ļ���
    Set img = thisViewer.Images(thisViewer.imageIndex(x, y))  'img�ǵ�ǰͼ��
    '��ע�������б�ע���ڣ��������ά���ķ�ͼ����
    If ls.Count <> 0 Then
        str3DMoveTag = ""   '����������,Ŀ���Ǳ��⵱�����ͬһ����λ���ƶ���ʱ�򣬶��ִ���л�ͼ������
        'ѭ�����б�ע�������Ƿ�����ά���ı�ע
        For Each l In ls
            'ͨ�������ж��Ƿ�����ά���ı�ע��
            If Mid(l.Tag, 1, 2) = "3D" Then
                k = InStr(l.Tag, "-")
                j = Val(Mid(l.Tag, 4, k - 1))   '��ע�߶�Ӧ��ͼ�����ڵ�Viewer��Index
                k = Val(Mid(l.Tag, k + 1))      '��ע�߶�Ӧ��ͼ�����ڵ�ͼ���
                ii = Me.MSFViewer.TextMatrix(j, 3)  '��ǰѡ���ͼ��ţ��������ͼ������ͼ��Ĵ���
                If str3DMoveTag <> Mid(l.Tag, 1, 5) And (ZLShowSeriesInfos.Count >= j) Then
                    '�ȴ���ͼ��ʵ���Ͼ��Ǵ���ķ���
                    Set objTempImage = funLoadAImage(j, k, 1)
                    If Not objTempImage Is Nothing Then
                        Call subInitAImage(objTempImage, j, Viewer(j))
                        
                        Viewer(j).Images.Add objTempImage
                        Viewer(j).Images.Move Viewer(j).Images.Count, ii
                        Viewer(j).Images.Remove ii + 1
                        Viewer(j).CurrentIndex = int3dCurrentlyImage(j)
                    
                        '����ͼ������ĺ�ɫʮ��ͶӰ�ı���
                        Dim cy As Double
                        '��λ��������ͼ���ཻ��X��Y��Z�淽���ϵ�ͶӰ��ֱ��ʹ��ͼ�����������������������������
                        If Abs(l.height) > Abs(l.width) Then
                            cy = thisViewer.ImageYPosition(x, y) / img.sizeY '* IIf(l.height < 0, -1, 1)
                        Else
                            cy = thisViewer.ImageXPosition(x, y) / img.sizeX '* IIf(l.width < 0, -1, 1)
                        End If
                        
                        '����λ�ߺͺ�ɫʮ��ͶӰ,�Ȼ���λ��
                        Dim lRefLine As DicomLabel
                        Set lRefLine = New DicomLabel
                        Set lRefLine = Viewer(j).Images(ii).ReferenceLine(SelectedImage, True)
                        lRefLine.ForeColour = vbBlue
                        lRefLine.LineStyle = 2
                        Viewer(j).Images(ii).Labels.Add lRefLine
                        
                        '��ʮ�ֽ���ĺ���
                        Set l = GetNewLabel(3, 0, 0, 20, 0)
                        l.ForeColour = vbRed
                        l.Tag = "3DH"
                        l.XOR = False
                        Viewer(j).Images(ii).Labels.Add l
                        l.left = lRefLine.left + cy * lRefLine.width - 10
                        l.top = lRefLine.top + cy * lRefLine.height
                        
                        '��ʮ�ֽ��������
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
'���ܣ�ɾ�������б�ѡ��ͼ��������û���ע
'��������
'���أ��ޣ�ֱ��ɾ���û���ע
'2009��
'------------------------------------------------
    Dim i As Integer
    If SelectedImage Is Nothing Then Exit Sub
    
    For i = SelectedImage.Labels.Count To G_INT_SYS_LABEL_COUNT + 1 Step -1
        SelectedImage.Labels.Remove i
    Next
    Set SelectedLabel = Nothing
    blnAngle = False                    '���ǶȲ����������
    intVasMeasure = 0                   '��Ѫ����խ�����������
    intCadioThoracicRatio = 0           '�����رȲ����������
    SubNoDispPeriod SelectedImage, Me      'Ϊָ������رվ��
End Sub

Private Sub subDelSelectedLabel()
'------------------------------------------------
'���ܣ�ɾ�������б�ѡ��ͼ���ѡ�б�ע
'��������
'���أ��ޣ�ֱ��ɾ���û���ע
'2009��
'------------------------------------------------
    Dim lblThis As DicomLabel
    Dim lblDel As DicomLabel
    Dim i As Integer
    If SelectedLabel Is Nothing Or SelectedImage Is Nothing Then Exit Sub
    If SelectedImage.Labels.IndexOf(SelectedLabel) <= G_INT_SYS_LABEL_COUNT Then Exit Sub
    If Mid(SelectedLabel.Tag, 1, 2) = "JD" Then '�ر���Ƕȣ���Ϊһ���Ƕ���������ע��������ɾ�����������ֱ�ע��
        If Not SelectedLabel.TagObject Is Nothing Then
            If Not SelectedLabel.TagObject.TagObject Is Nothing Then
                If SelectedImage.Labels.IndexOf(SelectedLabel.TagObject.TagObject) <> 0 Then
                    SelectedImage.Labels.Remove SelectedImage.Labels.IndexOf(SelectedLabel.TagObject.TagObject)
                End If
            End If
        End If
        blnAngle = False
    ElseIf left(SelectedLabel.Tag, 3) = "VAS" Then 'Ѫ����խ�����İ˸���ע,ɾ�����й�����6����ע
        If Not SelectedLabel.TagObject Is Nothing Then
            If Not SelectedLabel.TagObject.TagObject Is Nothing Then '�����һ��Ѫ�ܱ�E1
                Set lblThis = SelectedLabel.TagObject.TagObject
                For i = 1 To 5
                    If Not lblThis.TagObject Is Nothing Then '����Ѫ�ܴ�ֱ��
                        Set lblDel = lblThis
                        Set lblThis = lblThis.TagObject
                        If SelectedImage.Labels.IndexOf(lblDel) <> 0 Then SelectedImage.Labels.Remove SelectedImage.Labels.IndexOf(lblDel)
                    End If
                Next i
                If SelectedImage.Labels.IndexOf(lblThis) <> 0 Then SelectedImage.Labels.Remove SelectedImage.Labels.IndexOf(lblThis)
            End If
        End If
        intVasMeasure = 0
    ElseIf left(SelectedLabel.Tag, 3) = "CTR" Then '���رȲ�����4����ע��ɾ�����е�2����ע
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
    If Not SelectedLabel.TagObject Is Nothing Then  'ɾ�����������ֱ�ע�����ڽǶ��ǹ����ĵڶ����Ƕ��ߣ�
        If SelectedImage.Labels.IndexOf(SelectedLabel.TagObject) <> 0 Then
            SelectedImage.Labels.Remove SelectedImage.Labels.IndexOf(SelectedLabel.TagObject)
        End If
    End If
    If SelectedImage.Labels.IndexOf(SelectedLabel) <> 0 Then    '���ɾ����ѡ�еı�ע����
        SelectedImage.Labels.Remove SelectedImage.Labels.IndexOf(SelectedLabel)
    End If
    Set SelectedLabel = Nothing
    SubNoDispPeriod SelectedImage, Me
End Sub

Private Sub AddImgToFilm(img As DicomImage, thisViewer As DicomViewer, blnPrinted As Boolean)
'���ܣ���ͼ����ӵ���ӡԤ������
'������ Img -- ��Ҫ��ӵ�ͼ��
'       thisViewer -- ͼ�����ڵ�Viewer��Ϊ������ͼ��ʹ��
'       blnPrinted -- ��¼ͼ���Ƿ��Ѿ�����ӡ��
    
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
'���ŵ�ǰͼ���ṩ�������ֵ���
'������ intZoomDirection --- �����ַ���1�Ŵ�0��С��
    Dim dblScale As Double
    
    '�������󣬲����κ���ʾ
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
    '����������ͼ��ͬ��
    Call subSeriesInPhase(intSelectedSerial, Me, SelectedImage, IMG_SYN_ZOOMPAN)
    Exit Sub
End Sub

Public Sub MouseWheel(intDirection As Integer)
'���������ֵ��¼�
'������intDirection --- �����ֵķ���1--���ϣ�0--����
    
    On Error Resume Next
    If intMouseWheelRoll = 0 Then       '��ҳ
        If Viewer(intSelectedSerial).Visible = False Then Exit Sub
        
        If intDirection = 1 Then '�Ϸ�һҳ
            If VScro(intSelectedSerial).Visible = False Then
                'ȫ���й�Ƭ���л�����һ������
                If Button_miViewAllSeries = True Then
                    Call subAutoChangeSeries(intSelectedSerial, Val(ZLShowSeriesInfos(intSelectedSerial).SeriesNo), Viewer(intSelectedSerial).Tag, 1)
                End If
            Else
                If VScro(intSelectedSerial).Value = 1 Then
                    '������Ϸ�ҳ�Ѿ���ͷ�ˣ����ݲ����ж��Ƿ��л���ǰһ������
                    If Button_miViewAllSeries = True Then   'ȫ���й�Ƭ���л�����һ������
                        Call subAutoChangeSeries(intSelectedSerial, Val(ZLShowSeriesInfos(intSelectedSerial).SeriesNo), Viewer(intSelectedSerial).Tag, 1)
                    End If
                Else
                    If VScro(intSelectedSerial).Value - VScro(intSelectedSerial).LargeChange < 1 Then
                            '�����й�Ƭ���л�����һ��ͼ
                            VScro(intSelectedSerial).Value = 1
                    Else
                        VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Value - VScro(intSelectedSerial).LargeChange
                    End If
                End If
            End If
        Else        '�·�һҳ
            If VScro(intSelectedSerial).Visible = False Then
                'ȫ���й�Ƭ���л�����һ������
                If Button_miViewAllSeries = True Then
                    Call subAutoChangeSeries(intSelectedSerial, Val(ZLShowSeriesInfos(intSelectedSerial).SeriesNo), Viewer(intSelectedSerial).Tag, 2)
                End If
            Else
                If VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Max Then
                    '������·�ҳ�Ѿ���β�ˣ����ݲ����ж��Ƿ��л�����һ������
                    If Button_miViewAllSeries = True Then       'ȫ���й�Ƭ���л�����һ������
                        Call subAutoChangeSeries(intSelectedSerial, Val(ZLShowSeriesInfos(intSelectedSerial).SeriesNo), Viewer(intSelectedSerial).Tag, 2)
                    End If
                Else
                    If VScro(intSelectedSerial).Value + VScro(intSelectedSerial).LargeChange > VScro(intSelectedSerial).Max Then
                            '�����й�Ƭ���л������һ��ͼ
                            VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Max
                    Else
                        VScro(intSelectedSerial).Value = VScro(intSelectedSerial).Value + VScro(intSelectedSerial).LargeChange
                    End If
                End If
            End If
        End If
    ElseIf intMouseWheelRoll = 1 Then   '����
        Debug.Print "����"
        If intDirection = 1 Then    '�Ŵ�
            Call ZoomImage(1)
        Else        '��С
            Call ZoomImage(0)
        End If
    End If
End Sub

Private Sub subAutoChangeSeries(intVieweIndex As Integer, lngCurrentSeriesNo As Long, intCurrentIndex As Integer, intDirection As Integer)
'------------------------------------------------
'���ܣ��Զ��л�����
'������ intVieweIndex --�л����е�Viewer������
'       lngCurrentSeriesNo -- Viewer�е�ǰͼ������к�
'       intCurrentIndex --- ͼ����ZLSeriesInfos�е����
'       intDirection -- �л����еķ���1-�����л����У�2-�����л�����
'���أ��ޣ�ֱ���л�Viewer�е�����
'------------------------------------------------
    Dim i As Integer
    Dim lngNextSeriesNo As Long
    Dim intNextIndex As Integer
    Dim lngMax As Long
    
    On Error Resume Next
    
    lngMax = 99999
    '��ZLSeriesInfos�в�����һ�����е�����
    '���ϲ���
    If intDirection = 1 Then
        lngNextSeriesNo = 0
        intNextIndex = 0
        For i = 1 To ZLSeriesInfos.Count
            'ͬһ�μ��ģ��Ų���Ƚ�
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
            'ͬһ�μ��ģ��Ų���Ƚ�
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
    '�л�����һ������
    If intNextIndex <> 0 And intNextIndex <> lngMax Then
        '�������е�ͼ�����viewer(index)�е�ͼ��
        Call funcSwapSeries(intVieweIndex, intNextIndex, IIf(intDirection = 1, True, False))
    End If
    
    Exit Sub
End Sub

Private Function funMPRslope(frmParent As Object) As Boolean
'------------------------------------------------
'���ܣ� MPRб���ؽ�
'������ frmParent -- ������
'���أ� True--�ɹ���False---ȡ���˳�
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
'���ܣ� ��װfunMPR�Ĺ��̣��������汨��ͼ���¼�,�Ե�ǰ�����б�ѡ�е�������ʸ��״λ�ؽ�������ȡ��ʸ��״λ�ؽ�
'       thisForm.blnInMPR ˵���������Ƿ���ͼ�����ڽ����ؽ��Ĺ�����
'������ thisForm -- ��ʾͼ��Ĵ���
'       blnSilent -- ��Ĭ����MRP������ʾ
'���أ� True--�ɹ���False---ȡ���˳�
'ʱ�䣺2009-7
'------------------------------------------------
    funViewerMPR = funMPR(thisForm, blnSilent)
    '����˳��ؽ�״̬
    If blnInMPR = False Then
        '���������ⲿͼ����¼�����Ϊʸ��״λ�ؽ����ܱ����˽��ͼ
        RaiseEvent AfterSaveOuterImage(PstrCheckUID)
    End If
End Function


Private Sub subFilmAddImages(intType As Integer, Optional intInterval As Integer = 1, Optional blnStartOdd As Boolean = True)
'------------------------------------------------
'���ܣ� ��ƬԤ���������ͼ��
'������ intType -- ���ͼ��ķ�ʽ��1-������У�2-��ӵ�ǰͼ��3-�����ѡͼ;4-����������
'       intInterval -- �����ӡ�ļ��������intAddType=4ʱʹ��
'       blnStartOdd -- True ������False ż���𣬵�intAddType=4ʱʹ��
'���أ� ��
'------------------------------------------------
    Dim i As Integer, j As Integer
    Dim iImageIndex As Integer
    Dim im As DicomImage
    Dim intStart As Integer
    
    On Error GoTo err
    
    If intType = 1 Or intType = 4 Then      '�������
        If intSelectedSerial > 0 Then
            If intType = 4 And blnStartOdd = False Then
                intStart = 2
            Else
                intStart = 1
            End If
            
            '�����ֱ��������У�����intType��ǿ������Ϊ1
            If intType = 1 Then
                intInterval = 1
            Else
                intInterval = intInterval + 1
            End If
            
            iImageIndex = 1
            For i = intStart To ZLShowSeriesInfos(intSelectedSerial).ImageInfos.Count Step intInterval
                Set im = Nothing
                '�����ж�ͼ���Ƿ��Ѿ�װ�أ�����Ѿ�װ�أ����ҵ����ͼ����ʾ���������û��װ�أ���װ�ظ�ͼ��
                If ZLShowSeriesInfos(intSelectedSerial).ImageInfos(i).blnDisplayed = False Then
                    funcAddAImageA Viewer(intSelectedSerial), i
                End If
                
                '����ͼ�������
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
    ElseIf intType = 3 Then '�����ѡͼ
        '�ѱ�ѡ���ͼ����ӵ���ӡԤ������
        For i = 1 To ZLShowSeriesInfos.Count
            iImageIndex = 1
            For j = 1 To ZLShowSeriesInfos(i).ImageInfos.Count
                If ZLShowSeriesInfos(i).ImageInfos(j).blnSelected = True Then
                    Set im = Nothing
                    '�����ж�ͼ���Ƿ��Ѿ�װ�أ�����Ѿ�װ�أ����ҵ����ͼ����ʾ���������û��װ�أ���װ�ظ�ͼ��
                    If ZLShowSeriesInfos(i).ImageInfos(j).blnDisplayed = False Then
                        funcAddAImageA Viewer(i), j
                    End If
                    
                    '����ͼ�������
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
    Else                    '��ӵ�ǰͼ
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
