VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "flash.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmselectinfo 
   BorderStyle     =   0  'None
   Caption         =   "�����Һ�"
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   14.25
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin zl9NewQuery.ctlButton UsrCmd 
      Height          =   720
      Index           =   4
      Left            =   12840
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   6330
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   1270
      Caption         =   "��    ��"
      BackColor       =   64
      ForeColor       =   16777215
      FontSize        =   10.5
      FontBold        =   -1  'True
      ButtonHeight    =   600
      TextAligment    =   0
   End
   Begin VB.Timer tmrLoop 
      Interval        =   50
      Left            =   555
      Top             =   5250
   End
   Begin zl9NewQuery.ctlButton UsrCmd 
      Height          =   720
      Index           =   0
      Left            =   7290
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6375
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1270
      Caption         =   " ���￨�Һ� "
      BackColor       =   64
      ForeColor       =   16777215
      FontSize        =   10.5
      FontBold        =   -1  'True
      ButtonHeight    =   600
      TextAligment    =   0
   End
   Begin zl9NewQuery.ctlButton UsrCmd 
      Height          =   720
      Index           =   1
      Left            =   8625
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6375
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1270
      Caption         =   " ҽ�����Һ�"
      BackColor       =   64
      ForeColor       =   16777215
      FontSize        =   10.5
      FontBold        =   -1  'True
      ButtonHeight    =   600
      TextAligment    =   0
   End
   Begin zl9NewQuery.ctlButton UsrCmd 
      Height          =   720
      Index           =   2
      Left            =   9855
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6375
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1270
      Caption         =   " ���֤�Һ�"
      BackColor       =   64
      ForeColor       =   16777215
      FontSize        =   10.5
      FontBold        =   -1  'True
      ButtonHeight    =   600
      TextAligment    =   0
   End
   Begin zl9NewQuery.ctlButton UsrCmd 
      Height          =   720
      Index           =   3
      Left            =   11085
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6360
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1270
      Caption         =   "�ɣÿ��Һ�"
      BackColor       =   64
      ForeColor       =   16777215
      FontSize        =   10.5
      FontBold        =   -1  'True
      ButtonHeight    =   600
      TextAligment    =   0
   End
   Begin VB.PictureBox PicRB 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   14415
      ScaleHeight     =   555
      ScaleWidth      =   4200
      TabIndex        =   25
      Top             =   7860
      Width           =   4200
      Begin VB.Label Lblbottom 
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ���ڣ�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   15
         TabIndex        =   26
         Top             =   150
         Width           =   3270
      End
   End
   Begin VB.PictureBox PicBack 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   555
      Left            =   960
      ScaleHeight     =   555
      ScaleWidth      =   19395
      TabIndex        =   23
      Top             =   7395
      Width           =   19395
      Begin VB.Label Lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ӭ����Ժ����...����ף�����ջָ�..."
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   480
         Left            =   -15
         TabIndex        =   24
         Top             =   30
         Visible         =   0   'False
         Width           =   8955
      End
   End
   Begin VB.PictureBox PiclB 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   15
      ScaleHeight     =   555
      ScaleWidth      =   1845
      TabIndex        =   22
      Top             =   7830
      Width           =   1845
      Begin VB.Label lblOEM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   615
         TabIndex        =   27
         Top             =   150
         Width           =   165
      End
      Begin VB.Image imgFlag 
         Height          =   450
         Left            =   30
         Stretch         =   -1  'True
         Top             =   60
         Width           =   570
      End
   End
   Begin VB.PictureBox PicOrder 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C000&
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
      Height          =   3855
      Left            =   10395
      ScaleHeight     =   3855
      ScaleWidth      =   4980
      TabIndex        =   3
      Top             =   210
      Width           =   4980
      Begin VB.Label lblintrbottom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmselectinfo.frx":0000
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   75
         TabIndex        =   6
         Top             =   5835
         Width           =   4320
      End
      Begin VB.Label lblintr1 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "�����������������ҺŴ���ϵ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   0
         TabIndex        =   5
         Top             =   5595
         Width           =   3600
      End
      Begin VB.Label lblintr 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmselectinfo.frx":0053
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1920
         Left            =   15
         TabIndex        =   4
         Top             =   2475
         Width           =   14280
      End
      Begin VB.Image Imgback 
         Height          =   3375
         Left            =   1005
         Picture         =   "frmselectinfo.frx":0229
         Top             =   4305
         Width           =   4995
      End
   End
   Begin VB.PictureBox Picright 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   4200
      ScaleHeight     =   870
      ScaleWidth      =   5070
      TabIndex        =   19
      Top             =   15
      Width           =   5070
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Swfright 
         Height          =   900
         Left            =   0
         TabIndex        =   20
         Top             =   -30
         Width           =   4965
         _cx             =   8758
         _cy             =   1587
         FlashVars       =   ""
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   "-1"
         Loop            =   "-1"
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   "-1"
         Base            =   ""
         AllowScriptAccess=   ""
         Scale           =   "ShowAll"
         DeviceFont      =   "0"
         EmbedMovie      =   "0"
         BGColor         =   ""
         SWRemote        =   ""
         MovieData       =   ""
         SeamlessTabbing =   "1"
         Profile         =   "0"
         ProfileAddress  =   ""
         ProfilePort     =   "0"
         AllowNetworking =   "all"
         AllowFullScreen =   "false"
      End
   End
   Begin VB.PictureBox Picleft 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   660
      ScaleHeight     =   870
      ScaleWidth      =   1410
      TabIndex        =   17
      Top             =   15
      Width           =   1410
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Swfleft 
         Height          =   900
         Left            =   0
         TabIndex        =   18
         Top             =   -15
         Width           =   1440
         _cx             =   2540
         _cy             =   1587
         FlashVars       =   ""
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   "-1"
         Loop            =   "-1"
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   "-1"
         Base            =   ""
         AllowScriptAccess=   ""
         Scale           =   "ShowAll"
         DeviceFont      =   "0"
         EmbedMovie      =   "0"
         BGColor         =   ""
         SWRemote        =   ""
         MovieData       =   ""
         SeamlessTabbing =   "1"
         Profile         =   "0"
         ProfileAddress  =   ""
         ProfilePort     =   "0"
         AllowNetworking =   "all"
         AllowFullScreen =   "false"
      End
   End
   Begin VB.Timer Time 
      Interval        =   500
      Left            =   45
      Top             =   2925
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   645
      Top             =   2925
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHHead 
      Height          =   5445
      Left            =   840
      TabIndex        =   2
      Top             =   930
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   9604
      _Version        =   393216
      BackColor       =   16777215
      Rows            =   30
      Cols            =   8
      FixedCols       =   0
      BackColorFixed  =   8388608
      BackColorSel    =   12615808
      BackColorBkg    =   -2147483634
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   0
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BandDisplay     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.PictureBox PicClass 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   0
      Left            =   180
      Picture         =   "frmselectinfo.frx":12DBB
      ScaleHeight     =   1500
      ScaleWidth      =   660
      TabIndex        =   0
      Top             =   885
      Width           =   660
   End
   Begin VB.PictureBox PicMshbak 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5910
      Left            =   1515
      ScaleHeight     =   5910
      ScaleWidth      =   8535
      TabIndex        =   21
      Top             =   675
      Width           =   8535
   End
   Begin zl9NewQuery.ctlButton ctlCmd 
      Height          =   720
      Index           =   0
      Left            =   420
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   6705
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1270
      Caption         =   "��һҳ "
      BackColor       =   16777215
      FontSize        =   10.5
      ButtonHeight    =   600
      TextAligment    =   0
   End
   Begin zl9NewQuery.ctlButton ctlCmd 
      Height          =   720
      Index           =   1
      Left            =   2025
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   6690
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1270
      Caption         =   "��һҳ "
      BackColor       =   16777215
      FontSize        =   10.5
      ButtonHeight    =   600
      TextAligment    =   0
   End
   Begin zl9NewQuery.ctlButton ctlCmd 
      Height          =   720
      Index           =   2
      Left            =   3600
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   6705
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1270
      Caption         =   "��һ�� "
      BackColor       =   16777215
      FontSize        =   10.5
      ButtonHeight    =   600
      TextAligment    =   0
   End
   Begin zl9NewQuery.ctlButton ctlCmd 
      Height          =   720
      Index           =   3
      Left            =   4905
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   6705
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1270
      Caption         =   "��һ�� "
      BackColor       =   16777215
      FontSize        =   10.5
      ButtonHeight    =   600
      TextAligment    =   0
   End
   Begin MSComctlLib.ImageList ilsImage 
      Left            =   0
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
            Picture         =   "frmselectinfo.frx":1618F
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmselectinfo.frx":16529
            Key             =   "down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmselectinfo.frx":168C3
            Key             =   "reset"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmselectinfo.frx":16C5D
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmselectinfo.frx":1D4BF
            Key             =   "close"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picexp 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7590
      Left            =   10875
      ScaleHeight     =   7590
      ScaleWidth      =   5145
      TabIndex        =   7
      Top             =   840
      Width           =   5145
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ר������ʱ�䣺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   150
         TabIndex        =   14
         Top             =   5010
         Width           =   1770
      End
      Begin VB.Label Lbldate 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         Left            =   180
         TabIndex        =   13
         Top             =   5265
         Width           =   3480
      End
      Begin VB.Label LblPerson 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   180
         TabIndex        =   12
         Top             =   2190
         Width           =   3030
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "���˼�飺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   195
         TabIndex        =   11
         Top             =   1995
         Width           =   1395
      End
      Begin VB.Label LblClass 
         BackStyle       =   0  'Transparent
         Caption         =   "ְ�ƣ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1740
         TabIndex        =   10
         Top             =   1320
         Width           =   3120
      End
      Begin VB.Label Lblage 
         BackStyle       =   0  'Transparent
         Caption         =   "�������ڣ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1740
         TabIndex        =   9
         Top             =   810
         Width           =   3120
      End
      Begin VB.Label LblName 
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1740
         TabIndex        =   8
         Top             =   300
         Width           =   3120
      End
      Begin VB.Image Imgman 
         Height          =   1665
         Left            =   135
         Picture         =   "frmselectinfo.frx":23D21
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1485
      End
      Begin VB.Image Imgbak 
         Height          =   3375
         Left            =   420
         Picture         =   "frmselectinfo.frx":28031
         Top             =   3630
         Width           =   4995
      End
   End
   Begin VB.Label LblNoBill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Բ���Ʊ���Ѿ�ʹ���꣬�뵽���ڹҺš�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   840
      Index           =   0
      Left            =   630
      TabIndex        =   16
      Top             =   3705
      Width           =   15960
   End
   Begin VB.Label LblNoBill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Բ���Ʊ���Ѿ�ʹ���꣬�뵽���ڹҺš�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   840
      Index           =   1
      Left            =   45
      TabIndex        =   15
      Top             =   3015
      Width           =   15960
   End
   Begin VB.Label lblHospital 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   870
      TabIndex        =   1
      Top             =   -60
      Width           =   180
   End
   Begin VB.Image Imgnoselect 
      Height          =   600
      Left            =   2895
      Picture         =   "frmselectinfo.frx":3ABC3
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image Imgselected 
      Height          =   360
      Left            =   2160
      Picture         =   "frmselectinfo.frx":3DF97
      Stretch         =   -1  'True
      Top             =   255
      Visible         =   0   'False
      Width           =   660
   End
End
Attribute VB_Name = "frmselectinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private mStrNeed As String                  '�Һŵĺ���
Private mlngindex As Integer                '�ϴε����ͼƬ�ı�־
Private mstrClass As String                 '��ǰѡ�еĹҺ����
Private mlngCurPage As Long                 '��ǰ��ҳ��
Private mlngRecordCount As Long             '�ܹ��ļ�¼����
Private mlngPageSize As Long                'ÿҳ��ʾ�ļ�¼����
Private lngLenGrid(1) As Long               '����ĳ���
Private mlngTime As Long                    'ʱ���־
Private mlngRowP As Long                    '�ϴεĵ��
Private mlngRow As Long                     '��εĵ��
Private mlng����ID As Long
Private mStrBillNo As String             '���ݺ�
Private marr As Variant                  'ҽ��������Ϣ
Private mCurPayNeed As Currency          '������Ҫ�ķ���
Private mStrҽ������ As String           '�������м���ķ�ʽ
Private mStrItem As String               '�����Һŵ���Ŀ
Private mStrDoctorID As String           '�Һŵ�ҽ��ID
Private mStrDepart As String             '���йҺŵĿ���
Private mStrDoctorName As String         '���йҺŵ�ҽ������
Private mDblPrice As Double              '�Һŵļ۸�
Private mStrDetailID As String           '�շ�ϸĿID
Private mStrDepartID As String           '����ID
Private mStr�ű� As String
Private mBlnUseflag As Boolean           '�Ƿ��һ�ν���
Private mblnExit As Boolean
Private mbln����ʾ��Ѻ� As Boolean     '33918
Private mbln���￨�Һ� As Boolean
Private mblnFirst     As Boolean

'######################################################################################################################

Private Sub ctlCmd_CommandClick(Index As Integer)
    
    Dim i As Long
    
    Select Case Index
    Case 0
        mlngTime = Val(GetPara("������ˢ������")) * 2
        '�ж��Ƿ�Ϊ��ҳ
        If mlngCurPage > 1 Then
            mlngCurPage = mlngCurPage - 1
            DispInfo
        Else
            If MSHHead.Visible = True Then MSHHead.SetFocus
        End If
    Case 1
        mlngTime = Val(GetPara("������ˢ������")) * 2
        '����ҳ��Ϊ���������
        If (mlngRecordCount Mod mlngPageSize = 0) And mlngCurPage = mlngRecordCount / mlngPageSize Then
            If MSHHead.Visible = True Then MSHHead.SetFocus
            Exit Sub
        End If
        '����ҳ��Ϊ��ͨ��״��
        If mlngCurPage <= mlngRecordCount / mlngPageSize Then
            mlngCurPage = mlngCurPage + 1
            DispInfo
            If MSHHead.Visible = True Then MSHHead.SetFocus
        Else
            If MSHHead.Visible = True Then MSHHead.SetFocus
        End If
    Case 2
        '�����һ����¼
        
        If MSHHead.Row = 1 And mlngCurPage = 1 Then
            If mlngRecordCount Mod mlngPageSize = 0 Then
               mlngCurPage = mlngRecordCount \ mlngPageSize
            Else
                mlngCurPage = mlngRecordCount \ mlngPageSize + 1
            End If
            DispInfo
            If MSHHead.TextMatrix(1, 0) = "" Then Exit Sub
            MSHHead.Row = mlngRecordCount Mod mlngPageSize
            MSHHead.SetFocus
            Exit Sub
        End If
    
        '�����һ����¼
        If MSHHead.Row = 1 And mlngCurPage > 1 Then
            Call ctlCmd_CommandClick(0)
            MSHHead.Row = mlngPageSize
            MSHHead.SetFocus
            Exit Sub
        End If
    
        If MSHHead.Row > 1 Then
            Call UnSelectRow(MSHHead, RGB(255, 255, 255))
            MSHHead.Row = MSHHead.Row - 1
            MSHHead.SetFocus
        End If
    Case 3
        
        '���Ϊĳҳ�����һ����¼
        If MSHHead.Row = mlngPageSize Then
            Call ctlCmd_CommandClick(1)
            MSHHead.Row = 1
            Exit Sub
        End If
        '�������Ϊ���һ����¼
        If MSHHead.Row < MSHHead.Rows - 1 And MSHHead.TextMatrix(MSHHead.Row + 1, 0) = "" Then
            mlngCurPage = 1
            DispInfo
            MSHHead.SetFocus
            MSHHead.Row = 1
            Exit Sub
        End If
        '��ͨ�����
        If MSHHead.Row < MSHHead.Rows - 1 And MSHHead.TextMatrix(MSHHead.Row, 0) <> "" Then
            Call UnSelectRow(MSHHead, RGB(255, 255, 255))
            MSHHead.Row = MSHHead.Row + 1
            MSHHead.SetFocus
        End If
    
    End Select
    
End Sub

Private Sub Form_Activate()
    '���е��ݵĳ�ʼ��
    Dim i As Integer
    
    If InitBill = False Then Exit Sub
    
    If mblnExit Then
        Unload Me
        Exit Sub
    End If
    '33918
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
        
    'ȷ���������ͼƬλ��
    PicMshbak.BackColor = RGB(0, 204, 255)
    PiclB.Left = 0
    PiclB.Top = frmselectinfo.Height - PiclB.Height
    PicRB.Left = frmselectinfo.Width - Lblbottom.Width
    PicRB.Top = PiclB.Top ' + 100
    picBack.Left = PiclB.Width
    picBack.Top = PiclB.Top
    picBack.Width = PicRB.Left - PiclB.Width

    
    Picleft.Left = 0
    Picleft.Top = 0
    lblHospital.Left = Picleft.Width
    lblHospital.Top = 100
    Picright.Left = Picleft.Width + lblHospital.Width
    Picright.Top = 0
    Picright.Width = frmselectinfo.Width - Picright.Left
    Swfright.Left = 0
    Swfright.Top = 0
    Swfright.Width = Picright.Width
    'ȷ����ʾ�����λ�úͿ��
    MSHHead.Width = frmselectinfo.Width - 4400
    lngLenGrid(0) = MSHHead.Width
    MSHHead.Height = PiclB.Top - MSHHead.Top - 200 - ctlCmd(0).Height
    MSHHead.ColWidth(0) = MSHHead.Width / 4 '- 10
    MSHHead.ColWidth(1) = MSHHead.Width / 3 '- 50
    MSHHead.ColWidth(2) = MSHHead.Width * 5 / 24 '- 10
    lngLenGrid(1) = MSHHead.ColWidth(2)
    MSHHead.ColWidth(3) = MSHHead.Width * 5 / 24 ' - 10
    MSHHead.ColWidth(4) = 0
    MSHHead.ColWidth(5) = 0
    MSHHead.ColWidth(6) = 0
    MSHHead.ColWidth(7) = 0
    MSHHead.ColWidth(8) = 0
    mlngPageSize = MSHHead.Height \ 540 - 1
    MSHHead.Height = 540 * (mlngPageSize + 1)
    'ȷ��ͼƬ�����λ��
    PicOrder.BackColor = RGB(201, 211, 255)
    Picexp.BackColor = RGB(201, 211, 255)
    PicOrder.Left = MSHHead.Left + lngLenGrid(0) - lngLenGrid(1) + 50
    PicOrder.Top = MSHHead.Top - 50
    PicOrder.Width = frmselectinfo.Width - PicOrder.Left - 100
    PicOrder.Height = MSHHead.Height + 100
    '��ר�ҽ���֮�е�ͼƬ���д���
    Picexp.Top = MSHHead.Top - 50
    Picexp.Left = MSHHead.Left + MSHHead.Width + 50
    Picexp.Width = frmselectinfo.Width - Picexp.Left
    Picexp.Height = MSHHead.Height + 100
    '�����ݽ�����ʾ
    DispInfo
    Imgback.Top = PicOrder.Height - Imgback.Height
    Imgback.Left = PicOrder.Width - Imgback.Width '- 500
    Imgbak.Top = Picexp.Height - Imgbak.Height ' - 200
    Imgbak.Left = Picexp.Width - Imgbak.Width  '- 200
    'ȷ������ͼƬ��ť��λ��
    ctlCmd(0).Left = MSHHead.Left - 300
    ctlCmd(0).Top = MSHHead.Top + MSHHead.Height + (PiclB.Top - MSHHead.Top - MSHHead.Height - ctlCmd(0).Height) / 2
    ctlCmd(1).Left = ctlCmd(0).Left + ctlCmd(0).Width + 200
    ctlCmd(1).Top = ctlCmd(0).Top
'    PicUse.Left = MSHHead.Left + lngLenGrid(0) - lngLenGrid(1) + 500
'    PicUse.Top = ctlCmd(0).Top
    ctlCmd(2).Left = ctlCmd(1).Left + ctlCmd(1).Width + 200
    ctlCmd(2).Top = ctlCmd(0).Top
    ctlCmd(3).Left = ctlCmd(2).Left + ctlCmd(2).Width + 200
    ctlCmd(3).Top = ctlCmd(0).Top
        
    UsrCmd(0).ShowPicture = False
    UsrCmd(1).ShowPicture = False
    UsrCmd(2).ShowPicture = False
    UsrCmd(3).ShowPicture = False
    UsrCmd(4).ShowPicture = False
    
    Dim sglX As Single
    
    sglX = MSHHead.Left + lngLenGrid(0) - lngLenGrid(1) + 90
    
    If UsrCmd(0).Visible Then
        UsrCmd(0).Move sglX, ctlCmd(0).Top
        sglX = sglX + UsrCmd(0).Width + 120
    End If
    
    If UsrCmd(1).Visible Then
        UsrCmd(1).Move sglX, ctlCmd(0).Top
        sglX = sglX + UsrCmd(1).Width + 120
    End If
    
    If UsrCmd(2).Visible Then
        UsrCmd(2).Move sglX, ctlCmd(0).Top
        sglX = sglX + UsrCmd(2).Width + 120
    End If
    
    If UsrCmd(3).Visible Then
        UsrCmd(3).Move sglX, ctlCmd(0).Top
        sglX = sglX + UsrCmd(3).Width + 120
    End If
    
    If UsrCmd(4).Visible Then
        UsrCmd(4).Move sglX, ctlCmd(0).Top
        sglX = sglX + UsrCmd(4).Width + 120
    End If
    
    ctlCmd(0).Picture = ilsImage.ListImages("up")
    ctlCmd(1).Picture = ilsImage.ListImages("down")
    ctlCmd(2).Picture = ilsImage.ListImages("up")
    ctlCmd(3).Picture = ilsImage.ListImages("down")

    PicMshbak.Left = MSHHead.Left - 50
    PicMshbak.Top = MSHHead.Top - 50
    PicMshbak.Width = MSHHead.Width + 100
    PicMshbak.Height = MSHHead.Height + 100
    MSHHead.SetFocus
        Call Load���￨
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Dim strTmp As String
    
    mblnFirst = True
    mblnExit = False
    mbln����ʾ��Ѻ� = Val(zlDatabase.GetPara("�ҺŲ���ʾ��Ѻű�", glngSys, 1536, "0")) = 1
    
    If GetPara("�Һ����") = "" Then
       MsgBox "�����ڲ�������֮��������صĲ���", vbInformation, gstrSysName
       mblnExit = True
       Exit Sub
    End If
    
    lblInfo.Caption = GetPara("��ʾ����ʾ��Ϣ")
    
    '������ض���

    mlng����ID = 0
    If Dir(App.Path & "\ͼ��\" & "�����Һ����Ͻ�" & ".swf") <> "" Then
        Call PlayFlash(Picleft, Swfleft, App.Path & "\ͼ��\" & "�����Һ����Ͻ�" & ".swf", Picleft.Width, Picleft.Height)
    End If
    If Dir(App.Path & "\ͼ��\" & "�����Һ����Ͻ�" & ".swf") <> "" Then
        Call PlayFlash(Picright, Swfright, App.Path & "\ͼ��\" & "�����Һ����Ͻ�" & ".swf", Picright.Width, Picright.Height)
    End If
    If Dir(App.Path & "\ͼ��\�����Һ����������±���.pic") <> "" Then
        Imgback.Picture = LoadPicture(App.Path & "\ͼ��\�����Һ����������±���.pic")
    End If
    If Dir(App.Path & "\ͼ��\�����Һ�ר�ҽ������±���.pic") <> "" Then
        Imgbak.Picture = LoadPicture(App.Path & "\ͼ��\�����Һ�ר�ҽ������±���.pic")
    End If
    MSHHead.BackColorFixed = RGB(204, 255, 255)
    mlngindex = 0
    mBlnUseflag = False
    
    '����ǩ�����ݽ��г�ʼ��
    '------------------------------------------------------------------------------------------------------------------
    strTmp = "," & GetPara("�Һ����", "") & ","
    If strTmp = ",���߶�����," Then strTmp = ",���￨,ҽ����,"
    
    UsrCmd(0).Visible = (InStr(strTmp, ",���￨,") > 0)
        mbln���￨�Һ� = (InStr(strTmp, ",���￨,") > 0)
    UsrCmd(1).Visible = (InStr(strTmp, ",ҽ����,") > 0)
    UsrCmd(2).Visible = (InStr(strTmp, ",���֤,") > 0)
    UsrCmd(3).Visible = (InStr(strTmp, ",�ɣÿ�,") > 0)
    UsrCmd(4).Visible = (Val(zlDatabase.GetPara("������ʾ�����Һŷ��ذ�ť", glngSys, 1536, 0)) = 1)
    
    lblHospital.Caption = GetUnitName + Chr(10) + Chr(13) + "���������Һ�ϵͳ"
    
    
    mStrNeed = GetPara("�Һŵĺ���")
    If mStrNeed = "" Then
        mblnExit = True
        MsgBox "�����ڲ�������֮�����ùҺ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Load_Picture    '��ͼƬ����һ����Ҫ��������
    mlngCurPage = 1   '����ǰҳ����Ϊ��һҳ
    frmselectinfo.BackColor = RGB(201, 211, 255)
    lblintr.Caption = "��ϵͳ������" & Mid(strTmp, 2, Len(strTmp) - 2) & "�ķǼ��ﲡ�˹Һ�ʹ�á������������£�" + Chr(10) + Chr(13)
'    If strTmp = "ҽ�������߾���" Then
    lblintr.Caption = lblintr.Caption + "    һ��ѡ��Һſ�Ŀ��ͨ����һҳ����һҳ����һ�С���һ�л�ֱ�ӵ�ѡ��Ŀ��ѡ����ϣ���ĹҺ���Ŀ���������ʵ����������XXX�Һš���" + Chr(10) + Chr(13)
'    Else
'        lblintr.Caption = lblintr.Caption + "    һ��ѡ��Һſ�Ŀ��ͨ����һҳ����һҳ����һ�С���һ�л�ֱ�ӵ�ѡ��Ŀ��ѡ����ϣ���ĹҺ���Ŀ���������Ҫ�Һš���" + Chr(10) + Chr(13)
'    End If
    lblintr.Caption = lblintr.Caption + "    ����ˢ����֤��ݣ����ȷ����ѡ��ĹҺ���Ŀ���밴��Ļ��ʾˢ�����������룻" + Chr(10) + Chr(13)
    lblintr.Caption = lblintr.Caption + "    ����ȷ�ϹҺţ���Ļ��ʾ��ĸ�����Ϣ(�������Ա�����)֮�󣬵㡰ȷ�ϡ�����ִ�йҺ�ͬʱ�Զ������ʻ����շѣ�" + Chr(10) + Chr(13)
    lblintr.Caption = lblintr.Caption + "    �ġ�ȡ�߹Һ�ƾ������Ļ��ʾ�Һųɹ���Ϣ������ȴ���ȡ������ӡ�ĹҺ�ƾ������ָ�����Ҿ��"
    
    '��ҽ���ӿڽ��г�ʼ��
    If strTmp = "ҽ����" Then
'        If gclsInsure.InitInsure(gcnOracle) = False Then
'            mblnExit = True
'            Exit Sub
'        End If
    End If
    If strTmp = "ҽ�������߾���" Then
'        If gclsInsure.InitInsure(gcnOracle) = False Then
'            Picҽ��.Visible = False
'            Pic����.Picture = PicUse.Picture
'        End If
    End If
    
    DoSoftFlag
End Sub
    
Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub MSHHead_Click()
    mlngTime = Val(GetPara("������ˢ������")) * 2
    If (InStr(mstrClass, "ר��") > 0 And (Trim(MSHHead.TextMatrix(MSHHead.Row, 7)) <> "" Or Trim(MSHHead.TextMatrix(MSHHead.Row, 2)) <> "")) Then
        DisExpinfo
        MSHHead.Width = lngLenGrid(0)
        MSHHead.ColWidth(2) = lngLenGrid(1)
        PicOrder.Visible = False
        Picexp.Visible = True
    Else
        DispOrderContent
        PicOrder.Visible = True
        Picexp.Visible = False
        MSHHead.Width = lngLenGrid(0) - lngLenGrid(1)
        MSHHead.ColWidth(2) = 0
        '��ʾ��ͨ�ĹҺ�
    End If
    PicMshbak.Width = MSHHead.Width + 100
End Sub

Private Sub MSHHead_EnterCell()
    If MSHHead.TextMatrix(MSHHead.Row, 0) = "" Then
        If MSHHead.TextMatrix(1, 0) <> "" Then
            MSHHead.Row = 1
            Call SelectRow(MSHHead)
        End If
        Exit Sub
    End If
    Call SelectRow(MSHHead)
    mlngRow = MSHHead.Row
End Sub

Private Sub MSHHead_GotFocus()
    mlngTime = Val(GetPara("������ˢ������")) * 2
    MSHHead_EnterCell
    If (InStr(mstrClass, "ר��") > 0 And (Trim(MSHHead.TextMatrix(MSHHead.Row, 7)) <> "" Or Trim(MSHHead.TextMatrix(MSHHead.Row, 2)) <> "")) Then
        DisExpinfo
        MSHHead.Width = lngLenGrid(0)
        MSHHead.ColWidth(2) = lngLenGrid(1)
        PicOrder.Visible = False
        Picexp.Visible = True
    Else
        DispOrderContent
        PicOrder.Visible = True
        Picexp.Visible = False
        MSHHead.Width = lngLenGrid(0) - lngLenGrid(1)
        MSHHead.ColWidth(2) = 0
        '��ʾ��ͨ�ĹҺ�
    End If
    PicMshbak.Width = MSHHead.Width + 100
End Sub

Private Sub MSHHead_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then Unload Me
End Sub

Private Sub MSHHead_LeaveCell()
    If MSHHead.TextMatrix(MSHHead.Row, 0) = "" Then Exit Sub
    Call UnSelectRow(MSHHead, RGB(255, 255, 255))
End Sub

Private Sub PicBack_Paint()
    
    Call DrawColorToColor(picBack, picBack.BackColor, &HFFC0C0, , True)
    
End Sub

Private Sub PicClass_Click(Index As Integer)
Dim i As Integer
'���Һ������е��
    If IsNull(mlngindex) Then mlngindex = 0
    mlngCurPage = 1
    PicClass(mlngindex).Picture = Imgnoselect.Picture
    PicClass(mlngindex).CurrentX = 200
    PicClass(mlngindex).CurrentY = 200
    For i = 1 To Len(PicClass(mlngindex).Tag)
        PicClass(mlngindex).Print Mid(PicClass(mlngindex).Tag, i, 1)
        PicClass(mlngindex).CurrentY = PicClass(mlngindex).CurrentY + 150
        PicClass(mlngindex).CurrentX = 200
    Next
    PicClass(Index).Picture = Imgselected.Picture
    PicClass(Index).CurrentX = 200
    PicClass(Index).CurrentY = 200
    For i = 1 To Len(PicClass(Index).Tag)
        PicClass(Index).Print Mid(PicClass(Index).Tag, i, 1)
        PicClass(Index).CurrentY = PicClass(Index).CurrentY + 150
        PicClass(Index).CurrentX = 200
    Next
    mlngindex = Index
    mstrClass = PicClass(Index).Tag
    DispInfo
    
    PicMshbak.Left = MSHHead.Left - 50
    PicMshbak.Top = MSHHead.Top - 50
    PicMshbak.Width = MSHHead.Width + 100
    PicMshbak.Height = MSHHead.Height + 100
    '��ͼƬ��������
    For i = 0 To Index
        PicClass(i).ZOrder
    Next
    For i = PicClass.UBound To Index Step -1
        PicClass(i).ZOrder
    Next
    MSHHead.SetFocus
    Exit Sub
End Sub

Private Sub Load_Picture()
'�����Ĺ��ܣ���ʼ���������Խ���ѡ���ͼƬ
Dim rsTmp As New ADODB.Recordset
Dim Intup As Long, i As Integer

    On Error GoTo ErrHandle
    gstrSQL = "select ����,���� from ���� where ����  in(" + mStrNeed + ")"
    If rsTmp.State = adStateOpen Then rsTmp.Close
'    Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
'    RsTmp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
'    Call SQLTest
    PicClass(0).CurrentX = 200
    PicClass(0).CurrentY = 200
        For i = 1 To Len(CStr(rsTmp("����")))
            PicClass(0).Print Mid(CStr(rsTmp("����")), i, 1)
            PicClass(0).CurrentY = PicClass(0).CurrentY + 150
            PicClass(0).CurrentX = 200
        Next
    PicClass(0).Tag = CStr(rsTmp("����"))
    mstrClass = PicClass(0).Tag
    '�����޸ĺ���İ�ť
    rsTmp.MoveNext
    Do While Not rsTmp.EOF
        Intup = PicClass.UBound
        Intup = Intup + 1
        Load PicClass(Intup)
        PicClass(Intup).Visible = True
        PicClass(Intup).Left = PicClass(Intup - 1).Left
        PicClass(Intup).Top = PicClass(Intup - 1).Top + PicClass(Intup - 1).Height - 150
        PicClass(Intup).CurrentX = 350
        PicClass(Intup).CurrentY = 200
        PicClass(Intup).CurrentX = 200
        PicClass(Intup).CurrentY = 200
            For i = 1 To Len(CStr(rsTmp("����")))
                PicClass(Intup).Print Mid(CStr(rsTmp("����")), i, 1)
                PicClass(Intup).CurrentY = PicClass(Intup).CurrentY + 150
                PicClass(Intup).CurrentX = 200
            Next
        PicClass(Intup).Tag = CStr(rsTmp("����"))
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    PicClass(0).Picture = Imgselected.Picture
    
    PicClass(0).CurrentX = 200
    PicClass(0).CurrentY = 200
    For i = 1 To Len(PicClass(0).Tag)
        PicClass(0).Print Mid(PicClass(0).Tag, i, 1)
        PicClass(0).CurrentY = PicClass(0).CurrentY + 150
        PicClass(0).CurrentX = 200
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub PicClass_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub DispInfo()
Dim rsTmp As New ADODB.Recordset
Dim StrTime As String
Dim i As Long
Dim str�ҺŰ��� As String
    On Error GoTo ErrHandle
        '�ҺŰ��� �޺�����Լ�� �ҺŰ��������л�ȡ
    str�ҺŰ��� = "" & _
     "            Select A.ID, A.����, A.����, A.����id, A.��Ŀid, A.ҽ��id, A.ҽ������, A.��������, A. ����, A.��һ, A.�ܶ�, A.����, " & _
     "                   A.���� , A.����, A.����, A.���﷽ʽ,a.��ʼʱ��,a.��ֹʱ��, A.��ſ���, B.�޺���, B.��Լ��,a.ͣ������ " & vbNewLine & _
     "            From �ҺŰ��� A, �ҺŰ������� B " & vbNewLine & _
     "            Where a.ͣ������ Is Null And  sysdate Between Nvl(a.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And " & _
     "                 Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
     "                  And a.ID = B.����id(+) And Decode(To_Char(sysdate, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) = B.������Ŀ(+)" '
     str�ҺŰ��� = str�ҺŰ��� & vbNewLine & _
     "                 And Decode(To_Char(sysdate, 'D'), '1', a.����, '2', a.��һ, '3', a.�ܶ�, '4', a.����, '5', a.����, '6', a.����, '7',a.����, Null) Is Not Null"
    '�����ǰʱ�����ھ����ʱ���
    StrTime = _
    "Select ʱ��� From ʱ��� Where" & _
    " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
    " Between" & _
    " Decode(Sign(��ʼʱ�� - ��ֹʱ��),1,'3000-01-09 '||To_Char(��ʼʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(��ʼʱ��,'HH24:MI:SS'))" & _
    " And" & _
    " '3000-01-10 '||To_Char(��ֹʱ��,'HH24:MI:SS'))" & _
    " Or" & _
    " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
    " Between" & _
    " '3000-01-10 '||To_Char(��ʼʱ��,'HH24:MI:SS')" & _
    " And" & _
    " Decode(Sign(��ʼʱ�� - ��ֹʱ��),1,'3000-01-11 '||To_Char(��ֹʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(��ֹʱ��,'HH24:MI:SS')))"
    '���۸����Ŀ��������ϸ��Ϣ��ʾ
    'gstrSQL = "select M.ID as ID,M.���� as ����,M.���� as �ű�,M.����ID as ����ID,M.��ĿID as ��ĿID,C.���� as ����,N.���� as ����,Nvl(M.ҽ������, ' ') as ҽ������,M.ҽ��id,Decode(To_Char(SysDate,'D'),'1',M.����," & _
                "'2',M.��һ,'3',M.�ܶ�,'4',M.����,'5',M.����,'6',M.����,'7',M.����)  as ʱ��,D.�۸� as �۸�" & _
                " From �ҺŰ��� M,�շ���ĿĿ¼ N,���ű� C,(select �շ�ϸĿID,sum(�ּ�) as �۸� from" & _
                " �շѼ�Ŀ X  where ִ������<=sysdate and (��ֹ����> sysdate or ��ֹ���� is null) group by �շ�ϸĿID) D" & _
                " where M.ID not in (Select A.ID from �ҺŰ��� A,���˹ҺŻ��� B" & _
                " where a.����ID = B.����ID And a.��ĿID = B.��ĿID And" & _
                " Nvl(A.ҽ��ID,0)=Nvl(B.ҽ��ID,0) and (a.����=b.���� or b.���� is NULL )  And " & GetNodeCheckSQL("N.վ��") & " And " & GetNodeCheckSQL("C.վ��") & " And " & _
                " B.����=Trunc(Sysdate)  and �޺���<= B.�ѹ��� and A.�޺���<>0)" & _
                " and  Decode(To_Char(SysDate,'D'),'1',M.����,'2',M.��һ,'3',M.�ܶ�,'4'," & _
                " M.����,'5',M.����,'6',M.����,'7',M.����) in (" + StrTime + ") and M.��ĿID=N.ID" & _
                " and M.����ID=C.ID  and  M.��ĿID = D.�շ�ϸĿID And (M.ҽ��id Is Null Or Exists (Select 1 From ��Ա�� y Where y.ID=M.ҽ��id And " & GetNodeCheckSQL("y.վ��") & " And (y.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or y.����ʱ�� Is Null))) And m.����='" + mstrClass + "'" & _
                " order by ����,�۸� desc"
        
   '����:33918��33645��48446
    '���۸����Ŀ��������ϸ��Ϣ��ʾ
    gstrSQL = "" & _
    "   Select  distinct M.ID as ID,M.���� as ����,M.���� as �ű�,M.����ID as ����ID,M.��ĿID as ��ĿID,C.���� as ����, " & _
    "             N.���� as ����,Nvl(M.ҽ������, ' ') as ҽ������,M.ҽ��id,Decode(To_Char(SysDate,'D'),'1',M.����," & _
    "             '2',M.��һ,'3',M.�ܶ�,'4',M.����,'5',M.����,'6',M.����,'7',M.����)  as ʱ��" & _
    "   From (" & str�ҺŰ��� & ") M,�շ���ĿĿ¼ N,���ű� C " & _
    "   Where M.ID not in (  Select  A.ID from (" & str�ҺŰ��� & ") A,���˹ҺŻ��� B" & _
    "                                   Where  a.����ID = B.����ID And a.��ĿID = B.��ĿID And" & _
    "                                               Nvl(A.ҽ��ID,0)=Nvl(B.ҽ��ID,0)  and (a.����=b.���� or b.���� is NULL )   And " & GetNodeCheckSQL("N.վ��") & " And " & GetNodeCheckSQL("C.վ��") & " And " & _
    "                                               B.����=Trunc(Sysdate)  and �޺���<= B.�ѹ��� and A.�޺���<>0 ) " & _
    "               And  Decode(To_Char(SysDate,'D'),'1',M.����,'2',M.��һ,'3',M.�ܶ�,'4', M.����,'5',M.����,'6',M.����,'7',M.����) in (" + StrTime + ")  " & _
    "               And M.��ĿID=N.ID  and M.����ID=C.ID   " & _
    "               And M.ͣ������ is NULL And (M.ҽ��id Is Null Or Exists (Select 1 From ��Ա�� y Where y.ID=M.ҽ��id And " & GetNodeCheckSQL("y.վ��") & _
    "               And (y.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or y.����ʱ�� Is Null)) ) And nvl(m.ͣ������,to_Date('3000-01-01','yyyy-mm-dd'))>=to_Date('3000-01-01','yyyy-mm-dd')  And m.����='" + mstrClass + "'" & _
    "               And Not Exists(Select 1 From �ҺŰ���ͣ��״̬ Where ����ID=M.ID and Sysdate between ��ʼֹͣʱ�� and ����ֹͣʱ�� )"
                
    gstrSQL = "" & _
    "  Select ID,����,�ű�,����ID,��ĿID,����,����,ҽ������,ҽ��id,ʱ��,sum(nvl(�۸�,0)) as �۸� " & _
    "  From ( With A1 as (" & gstrSQL & ") " & _
    "           Select  A1.*,D.�ּ� as �۸�  From A1,�շѼ�Ŀ D " & _
    "           Where A1.��ĿID=D.�շ�ϸĿID And     D.ִ������<=sysdate and (D.��ֹ����> sysdate or D.��ֹ���� is null)  " & _
    "           Union all " & _
    "           Select  A1.*,D.�ּ� as �۸�  From A1,�շѴ�����Ŀ A,�շѼ�Ŀ D " & _
    "           Where A1.��ĿID=A.����ID and A.����ID=D.�շ�ϸĿID  And  D.ִ������<=sysdate and (D.��ֹ����> sysdate or D.��ֹ���� is null)  " & _
    "       )" & _
    " Group by ID,����,�ű�,����ID,��ĿID,����,����,ҽ������,ҽ��id,ʱ��  " & _
            IIf(mbln����ʾ��Ѻ�, "Having sum(nvl(�۸�,0))<>0 ", "") & _
    "   Order by ����,�۸�"

     Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    mlngRecordCount = rsTmp.RecordCount
    MSHHead.Clear
    MSHHead.TextMatrix(0, 0) = "����"
    MSHHead.TextMatrix(0, 1) = "��Ŀ"
    MSHHead.TextMatrix(0, 2) = "ҽ��"
    MSHHead.TextMatrix(0, 3) = "�۸�"
    '���û���κμ�¼
    If rsTmp.BOF Then
        rsTmp.Close
        Exit Sub
    End If
    i = mlngPageSize * (mlngCurPage - 1) + 1
    rsTmp.Move mlngPageSize * (mlngCurPage - 1)
    '�����ݽ�������д
    Do While i > mlngPageSize * (mlngCurPage - 1) And i <= mlngPageSize * mlngCurPage
    If Not rsTmp.EOF Then
        MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 0) = rsTmp("����")
        MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 1) = rsTmp("����")
        If Not IsNull(rsTmp("ҽ������")) Then MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 2) = rsTmp("ҽ������")
        MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 3) = Format(rsTmp("�۸�"), "0.00") + "Ԫ"
        MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 4) = rsTmp("��ĿID")
        MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 5) = rsTmp("����ID")
        MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 6) = rsTmp("�ű�")
        If Not IsNull(rsTmp("ҽ��ID")) Then MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 7) = rsTmp("ҽ��ID")
        MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 8) = CStr(rsTmp("ID"))
        rsTmp.MoveNext
    End If
    i = i + 1
    Loop
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
    
Private Sub DispOrderContent()
'���Ϊ��ר����𣬾���ʾ��ͨ������
Dim StrTime(6) As String, i As Integer
Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    '��ʱ��μ�¼�����ڴ�֮��
    gstrSQL = "select " & zlGetFeeFields("�ҺŰ���") & " from �ҺŰ��� where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(MSHHead.TextMatrix(MSHHead.Row, 8)))
    
    If Not rsTmp.BOF Then
        If Not IsNull(rsTmp("����")) Then StrTime(0) = CStr(rsTmp("����"))
        If Not IsNull(rsTmp("��һ")) Then StrTime(1) = CStr(rsTmp("��һ"))
        If Not IsNull(rsTmp("�ܶ�")) Then StrTime(2) = CStr(rsTmp("�ܶ�"))
        If Not IsNull(rsTmp("����")) Then StrTime(3) = CStr(rsTmp("����"))
        If Not IsNull(rsTmp("����")) Then StrTime(4) = CStr(rsTmp("����"))
        If Not IsNull(rsTmp("����")) Then StrTime(5) = CStr(rsTmp("����"))
        If Not IsNull(rsTmp("����")) Then StrTime(6) = CStr(rsTmp("����"))
    End If
    rsTmp.Close
    For i = 0 To 6
        If StrTime(i) = "" Then StrTime(i) = "    "
'        StrTime(i) = Left(" ", 2 - Len(StrTime(i))) + StrTime(i)
    Next
    PicOrder.ForeColor = vbRed
    Call DrawCell(Me.PicOrder, MSHHead.TextMatrix(MSHHead.Row, 0) + "." + MSHHead.TextMatrix(MSHHead.Row, 1) + " ����Ӧ��ʱ�䣺", 1, 10, PicOrder.Width - 100, PicOrder.Height - 100, , , , vbRed, RGB(201, 211, 255), frmselectinfo.Font, "0000", , 0, True)
    Call DrawCell(Me.PicOrder, "����һ��" + StrTime(1) + "   ���ڶ���" + StrTime(2), 500, 400, PicOrder.Width - 1000, 300, , , , vbBlack, RGB(201, 211, 255), frmselectinfo.Font, "0000", , 0, True)
    Call DrawCell(Me.PicOrder, "��������" + StrTime(3) + "   �����ģ�" + StrTime(4), 500, 700, PicOrder.Width - 1000, 300, , , , vbBlack, RGB(201, 211, 255), frmselectinfo.Font, "0000", , 0, True)
    Call DrawCell(Me.PicOrder, "�����壺" + StrTime(5) + "   ��������" + StrTime(6), 500, 1000, PicOrder.Width - 1000, 300, , , , vbBlack, RGB(201, 211, 255), frmselectinfo.Font, "0000", , 0, True)
    Call DrawCell(Me.PicOrder, "�����գ�" + StrTime(0), 500, 1300, PicOrder.Width - 1000, 300, , , , vbBlack, RGB(201, 211, 255), frmselectinfo.Font, "0000", , 0, True)
    Call DrawCell(Me.PicOrder, "�Һ�˵��", 0, 1800, PicOrder.Width - 1000, 300, , , , vbRed, RGB(201, 211, 255), frmselectinfo.Font, "0000", , 0, True)
    lblintr.Left = 100
    lblintr.Top = 2100
    lblintr.Width = PicOrder.Width - 200
    lblintr.Height = 3400
    Exit Sub
ErrHandle:
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
End Sub

Private Sub DisExpinfo()
    Dim rsTmp As New ADODB.Recordset
    Dim StrTime(6) As String, i As Integer
    Dim strTempFile As String

        On Error GoTo ErrHandle
        gstrSQL = "select A.ID ,A.����,A.�Ա�,trunc(A.��������) as ����,A.רҵ����ְ�� as רҵ,A.Ƹ�μ���ְ�� as Ƹ��,A.���˼�� as ��� " & _
        " from ��Ա�� A " & _
        " where   (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) and  (A.ID ='" + MSHHead.TextMatrix(MSHHead.Row, 7) + "' or A.���� = '" + MSHHead.TextMatrix(MSHHead.Row, 2) + "')"
        If rsTmp.State = adStateOpen Then rsTmp.Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSQL) 'SQLTest
'        RsTmp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
'        Call SQLTest
        '������ĸ�����Ϣ������ʾ
        If Not rsTmp.BOF Then
            If Not IsNull(rsTmp("����")) Then LblName.Caption = "������" + rsTmp("����") Else LblName.Caption = "������"
            If Not IsNull(rsTmp("����")) Then Lblage.Caption = "���գ�" + Format(rsTmp("����"), "yyyy-MM-dd") Else Lblage.Caption = ""
            If Not IsNull(rsTmp("רҵ")) Then LblClass.Caption = "ְ�ƣ�" + rsTmp("רҵ") Else LblClass.Caption = "ְ�ƣ���"
            If Not IsNull(rsTmp("���")) Then LblPerson.Caption = "ְ�ƣ�" + rsTmp("���") Else LblPerson.Caption = ""
            
        strTempFile = sys.Readlob(glngSys, 16, Val(rsTmp!ID))
        Imgman.Picture = LoadPicture(strTempFile)
        
          '  If Not IsNull(rsTmp("��Ƭ")) Then Imgman.Picture = LoadPicture(zlDatabase.ReadPicture(rsTmp, "��Ƭ")) Else Imgman.Picture = Nothing
         
        Else
            LblName.Caption = "������"
            Lblage.Caption = ""
            LblPerson.Caption = ""
            Imgman.Picture = Nothing
        End If
        rsTmp.Close
        
        '������ĹҺ���Ϣ���н���
        gstrSQL = "select " & zlGetFeeFields("�ҺŰ���") & " From �ҺŰ��� where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(MSHHead.TextMatrix(MSHHead.Row, 8)))
        
        If Not rsTmp.BOF Then
            If Not IsNull(rsTmp("����")) Then StrTime(0) = CStr(rsTmp("����"))
            If Not IsNull(rsTmp("��һ")) Then StrTime(1) = CStr(rsTmp("��һ"))
            If Not IsNull(rsTmp("�ܶ�")) Then StrTime(2) = CStr(rsTmp("�ܶ�"))
            If Not IsNull(rsTmp("����")) Then StrTime(3) = CStr(rsTmp("����"))
            If Not IsNull(rsTmp("����")) Then StrTime(4) = CStr(rsTmp("����"))
            If Not IsNull(rsTmp("����")) Then StrTime(5) = CStr(rsTmp("����"))
            If Not IsNull(rsTmp("����")) Then StrTime(6) = CStr(rsTmp("����"))
        End If
        rsTmp.Close
        For i = 0 To 6
            If StrTime(i) = "" Then StrTime(i) = "    "
    '        StrTime(i) = Left(" ", 2 - Len(StrTime(i))) + StrTime(i)
        Next
        Lbldate.Caption = ""
        Lbldate.Caption = Lbldate.Caption + "����һ��" + StrTime(1) + "   ���ڶ���" + StrTime(2) + Chr(10) + Chr(13)
        Lbldate.Caption = Lbldate.Caption + "��������" + StrTime(3) + "   �����ģ�" + StrTime(4) + Chr(10) + Chr(13)
        Lbldate.Caption = Lbldate.Caption + "�����壺" + StrTime(5) + "   ��������" + StrTime(6) + Chr(10) + Chr(13)
        Lbldate.Caption = Lbldate.Caption + "�����գ�" + StrTime(0)
Exit Sub
ErrHandle:
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
End Sub

Private Sub Picexp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then Unload Me
End Sub


Private Sub PiclB_Paint()
    Call DrawColorToColor(PiclB, PiclB.BackColor, &HFFC0C0, , True)
End Sub

Private Sub PicOrder_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then Unload Me
End Sub

Private Sub PicRB_Paint()
    Call DrawColorToColor(PicRB, PicRB.BackColor, &HFFC0C0, , True)
End Sub

Private Sub Time_Timer()
    Dim Strweek As String
    
    '�������޸�
    
    Select Case Weekday(Now())
        Case 1
            Strweek = "������"
        Case 2
            Strweek = "����һ"
        Case 3
            Strweek = "���ڶ�"
        Case 4
            Strweek = "������"
        Case 5
            Strweek = "������"
        Case 6
            Strweek = "������"
        Case 7
            Strweek = "������"
    End Select
    Lblbottom.Caption = Format(Now(), "YYYY��MM��DD��") + Strweek
    If MSHHead.Visible = False Then Exit Sub
    
    mlngTime = mlngTime - 1
    
    If mlngTime = 0 Then DispInfo
        
    
'    '��ָʾ����ʾ���ж�����ʾ
'    If Lblinfo.Left + Lblinfo.Width > 0 Then
'        Lblinfo.Left = Lblinfo.Left - 100
'    Else
'        Lblinfo.Left = Lblinfo1.Width + Lblinfo1.Left
'    End If
'
'    If Lblinfo1.Left + Lblinfo1.Width > 0 Then
'        Lblinfo1.Left = Lblinfo1.Left - 100
'    Else
'        Lblinfo1.Left = Lblinfo.Width + Lblinfo.Left
'    End If
End Sub

Private Function InitBill() As Boolean
'Ʊ�����ü�鼰��ʼ
Dim i As Integer
        If gblnBill�Һ� Then
            mlng����ID = CheckUsedBill(4, IIf(mlng����ID > 0, mlng����ID, glng�Һ�ID))
            If mlng����ID <= 0 Then
                ctlCmd(0).Visible = False
                ctlCmd(1).Visible = False
'                PicUse.Visible = False
                ctlCmd(2).Visible = False
                ctlCmd(3).Visible = False
                lblHospital.Visible = False
                Picright.Visible = False
                Picleft.Visible = False
                PiclB.Visible = False
                Picexp.Visible = False
                PicOrder.Visible = False
                lblNoBIll(0).Left = 0
                lblNoBIll(0).Top = frmselectinfo.Height / 2 - 1000
                lblNoBIll(1).Top = lblNoBIll(0).Top
                lblNoBIll(1).Left = lblNoBIll(0).Left + lblNoBIll(0).Width
                For i = 0 To PicClass.UBound
                    PicClass(i).Width = 0
                    PicClass(i).Height = 0
                Next
                MSHHead.Visible = False
                PicMshbak.Visible = False
                Picexp.Visible = False
                Lblbottom.Visible = False
                picBack.Visible = False
                PicRB.Visible = False
                lblNoBIll(0).Visible = True
                lblNoBIll(1).Visible = True
                UsrCmd(0).Visible = False
                UsrCmd(1).Visible = False
                UsrCmd(2).Visible = False
                UsrCmd(3).Visible = False
                                If UsrCmd(4).Visible Then
                                        With UsrCmd(4)
                                                .Left = Me.ScaleWidth - .Width - 120
                                                .Top = Me.ScaleHeight - .Height - 120
                                        End With
                End If
                InitBill = False
                Exit Function
            End If
            '�ϸ�ȡ��һ������
             mStrBillNo = GetNextBill(mlng����ID)
        Else
            ctlCmd(0).Visible = False
            ctlCmd(1).Visible = False
'            PicUse.Visible = False
            ctlCmd(2).Visible = False
            ctlCmd(3).Visible = False
            lblHospital.Visible = False
            Picright.Visible = False
            Picleft.Visible = False
            PiclB.Visible = False
            Picexp.Visible = False
            PicOrder.Visible = False
            lblNoBIll(0).Left = 0
            lblNoBIll(0).Top = frmselectinfo.Height / 2 - 1000
            lblNoBIll(1).Top = lblNoBIll(0).Top
            lblNoBIll(1).Left = lblNoBIll(0).Left + lblNoBIll(0).Width
            For i = 0 To PicClass.UBound
                PicClass(i).Width = 0
                PicClass(i).Height = 0
            Next
            MSHHead.Visible = False
            PicMshbak.Visible = False
            Picexp.Visible = False
            Lblbottom.Visible = False
            picBack.Visible = False
            PicRB.Visible = False
            lblNoBIll(0).Visible = True
            lblNoBIll(1).Visible = True
            UsrCmd(0).Visible = False
            UsrCmd(1).Visible = False
            UsrCmd(2).Visible = False
            UsrCmd(3).Visible = False
                        If UsrCmd(4).Visible Then
                With UsrCmd(4)
                    .Left = Me.ScaleWidth - .Width - 120
                    .Top = Me.ScaleHeight - .Height - 120
                End With
            End If
            InitBill = False
            Exit Function
        End If
        lblNoBIll(0).Visible = False
        lblNoBIll(1).Visible = False
        InitBill = True
End Function

Private Sub Timer1_Timer()
'���������ʾ��
    If lblNoBIll(0).Left + lblNoBIll(0).Width > 0 Then
        lblNoBIll(0).Left = lblNoBIll(0).Left - 100
    Else
        lblNoBIll(0).Left = lblNoBIll(1).Left + lblNoBIll(1).Width
    End If
    If lblNoBIll(1).Left + lblNoBIll(1).Width > 0 Then
        lblNoBIll(1).Left = lblNoBIll(1).Left - 100
    Else
        lblNoBIll(1).Left = lblNoBIll(0).Left + lblNoBIll(0).Width
    End If
End Sub

Private Sub DoSoftFlag()
    Dim strTmp As String
    Dim strOEM As String
    On Error Resume Next
    Err.Clear
    
    strTmp = zlRegInfo("��Ʒ����")
    If strTmp <> "-" Then
        lblOEM.Caption = strTmp & "���"
        '����״̬��ͼ���OEM����
        If strTmp = "����" Then
            Set imgFlag.Picture = LoadCustomPicture("Logo")
        Else
            strOEM = GetOEM(strTmp)
            Set imgFlag.Picture = LoadCustomPicture(strOEM)
            If Err <> 0 Then
                Err.Clear
                Set imgFlag.Picture = LoadCustomPicture("Logo")
            End If
        End If
        lblOEM.ToolTipText = ""
    End If
End Sub

Private Function GetOEM(ByVal strAsk As String) As String
    '-------------------------------------------------------------
    '���ܣ�����ÿ�����ߵ�ASCII��
    '������
    '���أ�
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    strCode = "OEM_"
    For intBit = 1 To Len(strAsk)
        'ȡÿ���ֵ�ASCII��
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function

Private Sub tmrLoop_Timer()

    '��ָʾ����ʾ���ж�����ʾ
    If lblInfo.Caption = "" Then Exit Sub
    If lblInfo.Tag = "" Then
        lblInfo.Left = picBack.Width
        lblInfo.Tag = "begined"
        lblInfo.Visible = True
    End If
    
    If lblInfo.Left - 300 + lblInfo.Width < 600 Then
        lblInfo.Left = picBack.Width
    Else
        lblInfo.Left = lblInfo.Left - 15
    End If
    
End Sub

Private Sub UsrCmd_CommandClick(Index As Integer)
    Dim CurLeft As Currency          'ҽ������ʣ��ķ���
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim aryItem As Variant
    Dim lng��ĿID As Long
    
    If Index = 4 Then
        Unload Me
    Else
            '��ȡ���ݽ��г�ʼ��
            '------------------------------------------------------------------------------------------------------------------
            If MSHHead.TextMatrix(MSHHead.Row, 3) = "" Then
                MSHHead.SetFocus
                Exit Sub
            End If
                
            mlngTime = Val(GetPara("������ˢ������")) * 2
            mStrItem = MSHHead.TextMatrix(MSHHead.Row, 1)
            mStrDoctorID = MSHHead.TextMatrix(MSHHead.Row, 7)
            mStrDepart = MSHHead.TextMatrix(MSHHead.Row, 0)
            mStrDoctorName = Trim(MSHHead.TextMatrix(MSHHead.Row, 2))
            mDblPrice = CDbl(Mid(MSHHead.TextMatrix(MSHHead.Row, 3), 1, Len(MSHHead.TextMatrix(MSHHead.Row, 3)) - 1))
            mStrDetailID = CStr(MSHHead.TextMatrix(MSHHead.Row, 4))
            lng��ĿID = CLng(MSHHead.TextMatrix(MSHHead.Row, 4))
            mStrDepartID = CStr(MSHHead.TextMatrix(MSHHead.Row, 5))
            mStr�ű� = CStr(MSHHead.TextMatrix(MSHHead.Row, 6))
            aryItem = GetRegistPrice(lng��ĿID)
                
            Select Case Index
            '------------------------------------------------------------------------------------------------------------------
          
            '------------------------------------------------------------------------------------------------------------------
            Case 1
                frmCheckLogin.ShowLogin Me, "ҽ�����Һ�", mStrDepart, mStrItem, mStrDoctorName, mlng����ID, mStrBillNo, Val(mStrDoctorID), mDblPrice, Val(mStrDetailID), Val(mStrDepartID), mStr�ű�
            '------------------------------------------------------------------------------------------------------------------
            Case 2
                frmCheckLogin.ShowLogin Me, "���֤�Һ�", mStrDepart, mStrItem, mStrDoctorName, mlng����ID, mStrBillNo, Val(mStrDoctorID), mDblPrice, Val(mStrDetailID), Val(mStrDepartID), mStr�ű�
            '------------------------------------------------------------------------------------------------------------------
            Case 3
                frmCheckLogin.ShowLogin Me, "�ɣÿ��Һ�", mStrDepart, mStrItem, mStrDoctorName, mlng����ID, mStrBillNo, Val(mStrDoctorID), mDblPrice, Val(mStrDetailID), Val(mStrDepartID), mStr�ű�
        Case Else
                frmCheckLogin.ShowLogin Me, "���￨�Һ�", mStrDepart, mStrItem, mStrDoctorName, mlng����ID, mStrBillNo, Val(mStrDoctorID), mDblPrice, Val(mStrDetailID), Val(mStrDepartID), mStr�ű�, Val(UsrCmd(Index).Tag)
          End Select
            Call DispInfo
    End If
End Sub

Private Sub Load���￨()
    Dim strSQL As String, lngIndex As Long
    Dim rsTmp As ADODB.Recordset
    If mbln���￨�Һ� = False Then Exit Sub
    
    '95138:���ϴ�,2016/4/12,�Ƿ�ˢ����Ϊ��������
    strSQL = "Select Id,���� From ҽ�ƿ���� Where �Ƿ����� = 1 And (substr(��������,1,1) = 1 or substr(��������, 2,1) = 1) And �Ƿ�̶� =1 Order By ȱʡ��־ Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName)
    If rsTmp.BOF Then Exit Sub
    lngIndex = 5
    Do While Not rsTmp.EOF
        Load UsrCmd(lngIndex)
        With UsrCmd(lngIndex)
            .Visible = True
            .AutoSize = False
            .TextAligment = 1
            .Caption = Nvl(rsTmp!����, "���￨")
            .Tag = Nvl(rsTmp!ID, 0)
                        .Left = UsrCmd(IIf(lngIndex = 5, 1, lngIndex - 1)).Left - UsrCmd(1).Width - 120
            .Top = UsrCmd(0).Top
             .Width = UsrCmd(1).Width
            .ShowPicture = False
        End With
        lngIndex = lngIndex + 1
        rsTmp.MoveNext
    Loop
    UsrCmd(0).Visible = False
End Sub

