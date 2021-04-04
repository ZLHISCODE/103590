VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "flash.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmselectinfo 
   BorderStyle     =   0  'None
   Caption         =   "自助挂号"
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
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
   StartUpPosition =   3  '窗口缺省
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
      Caption         =   "返    回"
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
      Caption         =   " 就诊卡挂号 "
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
      Caption         =   " 医保卡挂号"
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
      Caption         =   " 身份证挂号"
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
      Caption         =   "ＩＣ卡挂号"
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
         Name            =   "宋体"
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
         Caption         =   "当前日期：星期"
         BeginProperty Font 
            Name            =   "宋体"
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
         Name            =   "宋体"
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
         Caption         =   "欢迎到本院就诊...我们祝您早日恢复..."
         BeginProperty Font 
            Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
            Name            =   "宋体"
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
         Caption         =   "出现如下情况，请与挂号处联系："
         BeginProperty Font 
            Name            =   "宋体"
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
            Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      Caption         =   "上一页 "
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
      Caption         =   "下一页 "
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
      Caption         =   "上一条 "
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
      Caption         =   "下一条 "
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
         Name            =   "宋体"
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
         Caption         =   "专家门诊时间："
         BeginProperty Font 
            Name            =   "宋体"
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
            Name            =   "宋体"
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
            Name            =   "宋体"
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
         Caption         =   "个人简介："
         BeginProperty Font 
            Name            =   "宋体"
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
         Caption         =   "职称："
         BeginProperty Font 
            Name            =   "宋体"
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
         Caption         =   "出生日期："
         BeginProperty Font 
            Name            =   "宋体"
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
         Caption         =   "姓名："
         BeginProperty Font 
            Name            =   "宋体"
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
      Caption         =   "对不起，票据已经使用完，请到窗口挂号。"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "对不起，票据已经使用完，请到窗口挂号。"
      BeginProperty Font 
         Name            =   "宋体"
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
         Name            =   "宋体"
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

Private mStrNeed As String                  '挂号的号类
Private mlngindex As Integer                '上次点击的图片的标志
Private mstrClass As String                 '当前选中的挂号类别
Private mlngCurPage As Long                 '当前的页数
Private mlngRecordCount As Long             '总共的记录数量
Private mlngPageSize As Long                '每页显示的记录个数
Private lngLenGrid(1) As Long               '网格的长度
Private mlngTime As Long                    '时间标志
Private mlngRowP As Long                    '上次的点击
Private mlngRow As Long                     '这次的点击
Private mlng领用ID As Long
Private mStrBillNo As String             '单据号
Private marr As Variant                  '医保病人信息
Private mCurPayNeed As Currency          '病人需要的费用
Private mStr医保计算 As String           '用来进行计算的方式
Private mStrItem As String               '用来挂号的项目
Private mStrDoctorID As String           '挂号的医生ID
Private mStrDepart As String             '进行挂号的科室
Private mStrDoctorName As String         '进行挂号的医生姓名
Private mDblPrice As Double              '挂号的价格
Private mStrDetailID As String           '收费细目ID
Private mStrDepartID As String           '科室ID
Private mStr号别 As String
Private mBlnUseflag As Boolean           '是否第一次进入
Private mblnExit As Boolean
Private mbln不显示免费号 As Boolean     '33918
Private mbln就诊卡挂号 As Boolean
Private mblnFirst     As Boolean

'######################################################################################################################

Private Sub ctlCmd_CommandClick(Index As Integer)
    
    Dim i As Long
    
    Select Case Index
    Case 0
        mlngTime = Val(GetPara("主窗体刷新周期")) * 2
        '判断是否为首页
        If mlngCurPage > 1 Then
            mlngCurPage = mlngCurPage - 1
            DispInfo
        Else
            If MSHHead.Visible = True Then MSHHead.SetFocus
        End If
    Case 1
        mlngTime = Val(GetPara("主窗体刷新周期")) * 2
        '处理页面为整数的情况
        If (mlngRecordCount Mod mlngPageSize = 0) And mlngCurPage = mlngRecordCount / mlngPageSize Then
            If MSHHead.Visible = True Then MSHHead.SetFocus
            Exit Sub
        End If
        '处理页面为普通的状况
        If mlngCurPage <= mlngRecordCount / mlngPageSize Then
            mlngCurPage = mlngCurPage + 1
            DispInfo
            If MSHHead.Visible = True Then MSHHead.SetFocus
        Else
            If MSHHead.Visible = True Then MSHHead.SetFocus
        End If
    Case 2
        '如果第一条记录
        
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
    
        '如果第一条记录
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
        
        '如果为某页的最后一条记录
        If MSHHead.Row = mlngPageSize Then
            Call ctlCmd_CommandClick(1)
            MSHHead.Row = 1
            Exit Sub
        End If
        '如果本身为最后一条记录
        If MSHHead.Row < MSHHead.Rows - 1 And MSHHead.TextMatrix(MSHHead.Row + 1, 0) = "" Then
            mlngCurPage = 1
            DispInfo
            MSHHead.SetFocus
            MSHHead.Row = 1
            Exit Sub
        End If
        '普通的情况
        If MSHHead.Row < MSHHead.Rows - 1 And MSHHead.TextMatrix(MSHHead.Row, 0) <> "" Then
            Call UnSelectRow(MSHHead, RGB(255, 255, 255))
            MSHHead.Row = MSHHead.Row + 1
            MSHHead.SetFocus
        End If
    
    End Select
    
End Sub

Private Sub Form_Activate()
    '进行单据的初始化
    Dim i As Integer
    
    If InitBill = False Then Exit Sub
    
    If mblnExit Then
        Unload Me
        Exit Sub
    End If
    '33918
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
        
    '确定最下面的图片位置
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
    '确定显示网格的位置和宽度
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
    '确定图片的相关位置
    PicOrder.BackColor = RGB(201, 211, 255)
    Picexp.BackColor = RGB(201, 211, 255)
    PicOrder.Left = MSHHead.Left + lngLenGrid(0) - lngLenGrid(1) + 50
    PicOrder.Top = MSHHead.Top - 50
    PicOrder.Width = frmselectinfo.Width - PicOrder.Left - 100
    PicOrder.Height = MSHHead.Height + 100
    '对专家介绍之中的图片进行处理
    Picexp.Top = MSHHead.Top - 50
    Picexp.Left = MSHHead.Left + MSHHead.Width + 50
    Picexp.Width = frmselectinfo.Width - Picexp.Left
    Picexp.Height = MSHHead.Height + 100
    '对数据进行显示
    DispInfo
    Imgback.Top = PicOrder.Height - Imgback.Height
    Imgback.Left = PicOrder.Width - Imgback.Width '- 500
    Imgbak.Top = Picexp.Height - Imgbak.Height ' - 200
    Imgbak.Left = Picexp.Width - Imgbak.Width  '- 200
    '确定各个图片按钮的位置
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
        Call Load就诊卡
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Dim strTmp As String
    
    mblnFirst = True
    mblnExit = False
    mbln不显示免费号 = Val(zlDatabase.GetPara("挂号不显示免费号别", glngSys, 1536, "0")) = 1
    
    If GetPara("挂号类别") = "" Then
       MsgBox "请先在参数设置之中设置相关的参数", vbInformation, gstrSysName
       mblnExit = True
       Exit Sub
    End If
    
    lblInfo.Caption = GetPara("显示的提示信息")
    
    '播放相关动画

    mlng领用ID = 0
    If Dir(App.Path & "\图形\" & "自助挂号左上角" & ".swf") <> "" Then
        Call PlayFlash(Picleft, Swfleft, App.Path & "\图形\" & "自助挂号左上角" & ".swf", Picleft.Width, Picleft.Height)
    End If
    If Dir(App.Path & "\图形\" & "自助挂号右上角" & ".swf") <> "" Then
        Call PlayFlash(Picright, Swfright, App.Path & "\图形\" & "自助挂号右上角" & ".swf", Picright.Width, Picright.Height)
    End If
    If Dir(App.Path & "\图形\自助挂号主界面右下背景.pic") <> "" Then
        Imgback.Picture = LoadPicture(App.Path & "\图形\自助挂号主界面右下背景.pic")
    End If
    If Dir(App.Path & "\图形\自助挂号专家介绍右下背景.pic") <> "" Then
        Imgbak.Picture = LoadPicture(App.Path & "\图形\自助挂号专家介绍右下背景.pic")
    End If
    MSHHead.BackColorFixed = RGB(204, 255, 255)
    mlngindex = 0
    mBlnUseflag = False
    
    '将标签的内容进行初始化
    '------------------------------------------------------------------------------------------------------------------
    strTmp = "," & GetPara("挂号类别", "") & ","
    If strTmp = ",两者都可以," Then strTmp = ",就诊卡,医保卡,"
    
    UsrCmd(0).Visible = (InStr(strTmp, ",就诊卡,") > 0)
        mbln就诊卡挂号 = (InStr(strTmp, ",就诊卡,") > 0)
    UsrCmd(1).Visible = (InStr(strTmp, ",医保卡,") > 0)
    UsrCmd(2).Visible = (InStr(strTmp, ",身份证,") > 0)
    UsrCmd(3).Visible = (InStr(strTmp, ",ＩＣ卡,") > 0)
    UsrCmd(4).Visible = (Val(zlDatabase.GetPara("允许显示自助挂号返回按钮", glngSys, 1536, 0)) = 1)
    
    lblHospital.Caption = GetUnitName + Chr(10) + Chr(13) + "病人自助挂号系统"
    
    
    mStrNeed = GetPara("挂号的号类")
    If mStrNeed = "" Then
        mblnExit = True
        MsgBox "请先在参数设置之中设置挂号类别！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Load_Picture    '将图片按照一定的要求进行添加
    mlngCurPage = 1   '将当前页设置为第一页
    frmselectinfo.BackColor = RGB(201, 211, 255)
    lblintr.Caption = "本系统仅供持" & Mid(strTmp, 2, Len(strTmp) - 2) & "的非急诊病人挂号使用。操作步骤如下：" + Chr(10) + Chr(13)
'    If strTmp = "医保卡或者就诊" Then
    lblintr.Caption = lblintr.Caption + "    一、选择挂号科目：通过上一页、下一页、上一行、下一行或直接点选项目表，选中你希望的挂号项目，根据你的实际情况点击“XXX挂号”。" + Chr(10) + Chr(13)
'    Else
'        lblintr.Caption = lblintr.Caption + "    一、选择挂号科目：通过上一页、下一页、上一行、下一行或直接点选项目表，选中你希望的挂号项目，点击“我要挂号”。" + Chr(10) + Chr(13)
'    End If
    lblintr.Caption = lblintr.Caption + "    二、刷卡验证身份：如果确认你选择的挂号项目，请按屏幕提示刷卡、输入密码；" + Chr(10) + Chr(13)
    lblintr.Caption = lblintr.Caption + "    三、确认挂号：屏幕显示你的个人信息(姓名、性别、年龄)之后，点“确认”，将执行挂号同时自动从你帐户中收费；" + Chr(10) + Chr(13)
    lblintr.Caption = lblintr.Caption + "    四、取走挂号凭单：屏幕显示挂号成功信息，请你等待并取走所打印的挂号凭单，到指定科室就诊。"
    
    '将医保接口进行初始化
    If strTmp = "医保卡" Then
'        If gclsInsure.InitInsure(gcnOracle) = False Then
'            mblnExit = True
'            Exit Sub
'        End If
    End If
    If strTmp = "医保卡或者就诊" Then
'        If gclsInsure.InitInsure(gcnOracle) = False Then
'            Pic医保.Visible = False
'            Pic就诊.Picture = PicUse.Picture
'        End If
    End If
    
    DoSoftFlag
End Sub
    
Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub MSHHead_Click()
    mlngTime = Val(GetPara("主窗体刷新周期")) * 2
    If (InStr(mstrClass, "专家") > 0 And (Trim(MSHHead.TextMatrix(MSHHead.Row, 7)) <> "" Or Trim(MSHHead.TextMatrix(MSHHead.Row, 2)) <> "")) Then
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
        '显示普通的挂号
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
    mlngTime = Val(GetPara("主窗体刷新周期")) * 2
    MSHHead_EnterCell
    If (InStr(mstrClass, "专家") > 0 And (Trim(MSHHead.TextMatrix(MSHHead.Row, 7)) <> "" Or Trim(MSHHead.TextMatrix(MSHHead.Row, 2)) <> "")) Then
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
        '显示普通的挂号
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
'将挂号类别进行点击
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
    '将图片进行设置
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
'函数的功能，初始化几个可以进行选择的图片
Dim rsTmp As New ADODB.Recordset
Dim Intup As Long, i As Integer

    On Error GoTo ErrHandle
    gstrSQL = "select 编码,名称 from 号类 where 名称  in(" + mStrNeed + ")"
    If rsTmp.State = adStateOpen Then rsTmp.Close
'    Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
'    RsTmp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
'    Call SQLTest
    PicClass(0).CurrentX = 200
    PicClass(0).CurrentY = 200
        For i = 1 To Len(CStr(rsTmp("名称")))
            PicClass(0).Print Mid(CStr(rsTmp("名称")), i, 1)
            PicClass(0).CurrentY = PicClass(0).CurrentY + 150
            PicClass(0).CurrentX = 200
        Next
    PicClass(0).Tag = CStr(rsTmp("名称"))
    mstrClass = PicClass(0).Tag
    '依次修改后面的按钮
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
            For i = 1 To Len(CStr(rsTmp("名称")))
                PicClass(Intup).Print Mid(CStr(rsTmp("名称")), i, 1)
                PicClass(Intup).CurrentY = PicClass(Intup).CurrentY + 150
                PicClass(Intup).CurrentX = 200
            Next
        PicClass(Intup).Tag = CStr(rsTmp("名称"))
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
Dim str挂号安排 As String
    On Error GoTo ErrHandle
        '挂号安排 限号数限约数 挂号安排限制中获取
    str挂号安排 = "" & _
     "            Select A.ID, A.号码, A.号类, A.科室id, A.项目id, A.医生id, A.医生姓名, A.病案必须, A. 周日, A.周一, A.周二, A.周三, " & _
     "                   A.周四 , A.周五, A.周六, A.分诊方式,a.开始时间,a.终止时间, A.序号控制, B.限号数, B.限约数,a.停用日期 " & vbNewLine & _
     "            From 挂号安排 A, 挂号安排限制 B " & vbNewLine & _
     "            Where a.停用日期 Is Null And  sysdate Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And " & _
     "                 Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
     "                  And a.ID = B.安排id(+) And Decode(To_Char(sysdate, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) = B.限制项目(+)" '
     str挂号安排 = str挂号安排 & vbNewLine & _
     "                 And Decode(To_Char(sysdate, 'D'), '1', a.周日, '2', a.周一, '3', a.周二, '4', a.周三, '5', a.周四, '6', a.周五, '7',a.周六, Null) Is Not Null"
    '求出当前时间属于具体的时间段
    StrTime = _
    "Select 时间段 From 时间段 Where" & _
    " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
    " Between" & _
    " Decode(Sign(开始时间 - 终止时间),1,'3000-01-09 '||To_Char(开始时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(开始时间,'HH24:MI:SS'))" & _
    " And" & _
    " '3000-01-10 '||To_Char(终止时间,'HH24:MI:SS'))" & _
    " Or" & _
    " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & _
    " Between" & _
    " '3000-01-10 '||To_Char(开始时间,'HH24:MI:SS')" & _
    " And" & _
    " Decode(Sign(开始时间 - 终止时间),1,'3000-01-11 '||To_Char(终止时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(终止时间,'HH24:MI:SS')))"
    '将价格和项目及其它明细信息显示
    'gstrSQL = "select M.ID as ID,M.号类 as 号类,M.号码 as 号别,M.科室ID as 科室ID,M.项目ID as 项目ID,C.名称 as 科室,N.名称 as 名称,Nvl(M.医生姓名, ' ') as 医生姓名,M.医生id,Decode(To_Char(SysDate,'D'),'1',M.周日," & _
                "'2',M.周一,'3',M.周二,'4',M.周三,'5',M.周四,'6',M.周五,'7',M.周六)  as 时间,D.价格 as 价格" & _
                " From 挂号安排 M,收费项目目录 N,部门表 C,(select 收费细目ID,sum(现价) as 价格 from" & _
                " 收费价目 X  where 执行日期<=sysdate and (终止日期> sysdate or 终止日期 is null) group by 收费细目ID) D" & _
                " where M.ID not in (Select A.ID from 挂号安排 A,病人挂号汇总 B" & _
                " where a.科室ID = B.科室ID And a.项目ID = B.项目ID And" & _
                " Nvl(A.医生ID,0)=Nvl(B.医生ID,0) and (a.号码=b.号码 or b.号码 is NULL )  And " & GetNodeCheckSQL("N.站点") & " And " & GetNodeCheckSQL("C.站点") & " And " & _
                " B.日期=Trunc(Sysdate)  and 限号数<= B.已挂数 and A.限号数<>0)" & _
                " and  Decode(To_Char(SysDate,'D'),'1',M.周日,'2',M.周一,'3',M.周二,'4'," & _
                " M.周三,'5',M.周四,'6',M.周五,'7',M.周六) in (" + StrTime + ") and M.项目ID=N.ID" & _
                " and M.科室ID=C.ID  and  M.项目ID = D.收费细目ID And (M.医生id Is Null Or Exists (Select 1 From 人员表 y Where y.ID=M.医生id And " & GetNodeCheckSQL("y.站点") & " And (y.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or y.撤档时间 Is Null))) And m.号类='" + mstrClass + "'" & _
                " order by 科室,价格 desc"
        
   '问题:33918和33645和48446
    '将价格和项目及其它明细信息显示
    gstrSQL = "" & _
    "   Select  distinct M.ID as ID,M.号类 as 号类,M.号码 as 号别,M.科室ID as 科室ID,M.项目ID as 项目ID,C.名称 as 科室, " & _
    "             N.名称 as 名称,Nvl(M.医生姓名, ' ') as 医生姓名,M.医生id,Decode(To_Char(SysDate,'D'),'1',M.周日," & _
    "             '2',M.周一,'3',M.周二,'4',M.周三,'5',M.周四,'6',M.周五,'7',M.周六)  as 时间" & _
    "   From (" & str挂号安排 & ") M,收费项目目录 N,部门表 C " & _
    "   Where M.ID not in (  Select  A.ID from (" & str挂号安排 & ") A,病人挂号汇总 B" & _
    "                                   Where  a.科室ID = B.科室ID And a.项目ID = B.项目ID And" & _
    "                                               Nvl(A.医生ID,0)=Nvl(B.医生ID,0)  and (a.号码=b.号码 or b.号码 is NULL )   And " & GetNodeCheckSQL("N.站点") & " And " & GetNodeCheckSQL("C.站点") & " And " & _
    "                                               B.日期=Trunc(Sysdate)  and 限号数<= B.已挂数 and A.限号数<>0 ) " & _
    "               And  Decode(To_Char(SysDate,'D'),'1',M.周日,'2',M.周一,'3',M.周二,'4', M.周三,'5',M.周四,'6',M.周五,'7',M.周六) in (" + StrTime + ")  " & _
    "               And M.项目ID=N.ID  and M.科室ID=C.ID   " & _
    "               And M.停用日期 is NULL And (M.医生id Is Null Or Exists (Select 1 From 人员表 y Where y.ID=M.医生id And " & GetNodeCheckSQL("y.站点") & _
    "               And (y.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or y.撤档时间 Is Null)) ) And nvl(m.停用日期,to_Date('3000-01-01','yyyy-mm-dd'))>=to_Date('3000-01-01','yyyy-mm-dd')  And m.号类='" + mstrClass + "'" & _
    "               And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=M.ID and Sysdate between 开始停止时间 and 结束停止时间 )"
                
    gstrSQL = "" & _
    "  Select ID,号类,号别,科室ID,项目ID,科室,名称,医生姓名,医生id,时间,sum(nvl(价格,0)) as 价格 " & _
    "  From ( With A1 as (" & gstrSQL & ") " & _
    "           Select  A1.*,D.现价 as 价格  From A1,收费价目 D " & _
    "           Where A1.项目ID=D.收费细目ID And     D.执行日期<=sysdate and (D.终止日期> sysdate or D.终止日期 is null)  " & _
    "           Union all " & _
    "           Select  A1.*,D.现价 as 价格  From A1,收费从属项目 A,收费价目 D " & _
    "           Where A1.项目ID=A.主项ID and A.从项ID=D.收费细目ID  And  D.执行日期<=sysdate and (D.终止日期> sysdate or D.终止日期 is null)  " & _
    "       )" & _
    " Group by ID,号类,号别,科室ID,项目ID,科室,名称,医生姓名,医生id,时间  " & _
            IIf(mbln不显示免费号, "Having sum(nvl(价格,0))<>0 ", "") & _
    "   Order by 科室,价格"

     Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    mlngRecordCount = rsTmp.RecordCount
    MSHHead.Clear
    MSHHead.TextMatrix(0, 0) = "科室"
    MSHHead.TextMatrix(0, 1) = "项目"
    MSHHead.TextMatrix(0, 2) = "医生"
    MSHHead.TextMatrix(0, 3) = "价格"
    '如果没有任何记录
    If rsTmp.BOF Then
        rsTmp.Close
        Exit Sub
    End If
    i = mlngPageSize * (mlngCurPage - 1) + 1
    rsTmp.Move mlngPageSize * (mlngCurPage - 1)
    '将数据进行逐步填写
    Do While i > mlngPageSize * (mlngCurPage - 1) And i <= mlngPageSize * mlngCurPage
    If Not rsTmp.EOF Then
        MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 0) = rsTmp("科室")
        MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 1) = rsTmp("名称")
        If Not IsNull(rsTmp("医生姓名")) Then MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 2) = rsTmp("医生姓名")
        MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 3) = Format(rsTmp("价格"), "0.00") + "元"
        MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 4) = rsTmp("项目ID")
        MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 5) = rsTmp("科室ID")
        MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 6) = rsTmp("号别")
        If Not IsNull(rsTmp("医生ID")) Then MSHHead.TextMatrix(i - mlngPageSize * (mlngCurPage - 1), 7) = rsTmp("医生ID")
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
'如果为非专家类别，就显示普通的内容
Dim StrTime(6) As String, i As Integer
Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    '将时间段记录进入内存之中
    gstrSQL = "select " & zlGetFeeFields("挂号安排") & " from 挂号安排 where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(MSHHead.TextMatrix(MSHHead.Row, 8)))
    
    If Not rsTmp.BOF Then
        If Not IsNull(rsTmp("周日")) Then StrTime(0) = CStr(rsTmp("周日"))
        If Not IsNull(rsTmp("周一")) Then StrTime(1) = CStr(rsTmp("周一"))
        If Not IsNull(rsTmp("周二")) Then StrTime(2) = CStr(rsTmp("周二"))
        If Not IsNull(rsTmp("周三")) Then StrTime(3) = CStr(rsTmp("周三"))
        If Not IsNull(rsTmp("周四")) Then StrTime(4) = CStr(rsTmp("周四"))
        If Not IsNull(rsTmp("周五")) Then StrTime(5) = CStr(rsTmp("周五"))
        If Not IsNull(rsTmp("周六")) Then StrTime(6) = CStr(rsTmp("周六"))
    End If
    rsTmp.Close
    For i = 0 To 6
        If StrTime(i) = "" Then StrTime(i) = "    "
'        StrTime(i) = Left(" ", 2 - Len(StrTime(i))) + StrTime(i)
    Next
    PicOrder.ForeColor = vbRed
    Call DrawCell(Me.PicOrder, MSHHead.TextMatrix(MSHHead.Row, 0) + "." + MSHHead.TextMatrix(MSHHead.Row, 1) + " 门诊应诊时间：", 1, 10, PicOrder.Width - 100, PicOrder.Height - 100, , , , vbRed, RGB(201, 211, 255), frmselectinfo.Font, "0000", , 0, True)
    Call DrawCell(Me.PicOrder, "星期一：" + StrTime(1) + "   星期二：" + StrTime(2), 500, 400, PicOrder.Width - 1000, 300, , , , vbBlack, RGB(201, 211, 255), frmselectinfo.Font, "0000", , 0, True)
    Call DrawCell(Me.PicOrder, "星期三：" + StrTime(3) + "   星期四：" + StrTime(4), 500, 700, PicOrder.Width - 1000, 300, , , , vbBlack, RGB(201, 211, 255), frmselectinfo.Font, "0000", , 0, True)
    Call DrawCell(Me.PicOrder, "星期五：" + StrTime(5) + "   星期六：" + StrTime(6), 500, 1000, PicOrder.Width - 1000, 300, , , , vbBlack, RGB(201, 211, 255), frmselectinfo.Font, "0000", , 0, True)
    Call DrawCell(Me.PicOrder, "星期日：" + StrTime(0), 500, 1300, PicOrder.Width - 1000, 300, , , , vbBlack, RGB(201, 211, 255), frmselectinfo.Font, "0000", , 0, True)
    Call DrawCell(Me.PicOrder, "挂号说明", 0, 1800, PicOrder.Width - 1000, 300, , , , vbRed, RGB(201, 211, 255), frmselectinfo.Font, "0000", , 0, True)
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
        gstrSQL = "select A.ID ,A.姓名,A.性别,trunc(A.出生日期) as 生日,A.专业技术职务 as 专业,A.聘任技术职务 as 聘任,A.个人简介 as 简介 " & _
        " from 人员表 A " & _
        " where   (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) and  (A.ID ='" + MSHHead.TextMatrix(MSHHead.Row, 7) + "' or A.姓名 = '" + MSHHead.TextMatrix(MSHHead.Row, 2) + "')"
        If rsTmp.State = adStateOpen Then rsTmp.Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSQL) 'SQLTest
'        RsTmp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
'        Call SQLTest
        '将具体的个人信息进行显示
        If Not rsTmp.BOF Then
            If Not IsNull(rsTmp("姓名")) Then LblName.Caption = "姓名：" + rsTmp("姓名") Else LblName.Caption = "姓名："
            If Not IsNull(rsTmp("生日")) Then Lblage.Caption = "生日：" + Format(rsTmp("生日"), "yyyy-MM-dd") Else Lblage.Caption = ""
            If Not IsNull(rsTmp("专业")) Then LblClass.Caption = "职称：" + rsTmp("专业") Else LblClass.Caption = "职称：无"
            If Not IsNull(rsTmp("简介")) Then LblPerson.Caption = "职称：" + rsTmp("简介") Else LblPerson.Caption = ""
            
        strTempFile = sys.Readlob(glngSys, 16, Val(rsTmp!ID))
        Imgman.Picture = LoadPicture(strTempFile)
        
          '  If Not IsNull(rsTmp("照片")) Then Imgman.Picture = LoadPicture(zlDatabase.ReadPicture(rsTmp, "照片")) Else Imgman.Picture = Nothing
         
        Else
            LblName.Caption = "姓名："
            Lblage.Caption = ""
            LblPerson.Caption = ""
            Imgman.Picture = Nothing
        End If
        rsTmp.Close
        
        '将具体的挂号信息进行介绍
        gstrSQL = "select " & zlGetFeeFields("挂号安排") & " From 挂号安排 where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(MSHHead.TextMatrix(MSHHead.Row, 8)))
        
        If Not rsTmp.BOF Then
            If Not IsNull(rsTmp("周日")) Then StrTime(0) = CStr(rsTmp("周日"))
            If Not IsNull(rsTmp("周一")) Then StrTime(1) = CStr(rsTmp("周一"))
            If Not IsNull(rsTmp("周二")) Then StrTime(2) = CStr(rsTmp("周二"))
            If Not IsNull(rsTmp("周三")) Then StrTime(3) = CStr(rsTmp("周三"))
            If Not IsNull(rsTmp("周四")) Then StrTime(4) = CStr(rsTmp("周四"))
            If Not IsNull(rsTmp("周五")) Then StrTime(5) = CStr(rsTmp("周五"))
            If Not IsNull(rsTmp("周六")) Then StrTime(6) = CStr(rsTmp("周六"))
        End If
        rsTmp.Close
        For i = 0 To 6
            If StrTime(i) = "" Then StrTime(i) = "    "
    '        StrTime(i) = Left(" ", 2 - Len(StrTime(i))) + StrTime(i)
        Next
        Lbldate.Caption = ""
        Lbldate.Caption = Lbldate.Caption + "星期一：" + StrTime(1) + "   星期二：" + StrTime(2) + Chr(10) + Chr(13)
        Lbldate.Caption = Lbldate.Caption + "星期三：" + StrTime(3) + "   星期四：" + StrTime(4) + Chr(10) + Chr(13)
        Lbldate.Caption = Lbldate.Caption + "星期五：" + StrTime(5) + "   星期六：" + StrTime(6) + Chr(10) + Chr(13)
        Lbldate.Caption = Lbldate.Caption + "星期日：" + StrTime(0)
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
    
    '将变量修改
    
    Select Case Weekday(Now())
        Case 1
            Strweek = "星期日"
        Case 2
            Strweek = "星期一"
        Case 3
            Strweek = "星期二"
        Case 4
            Strweek = "星期三"
        Case 5
            Strweek = "星期四"
        Case 6
            Strweek = "星期五"
        Case 7
            Strweek = "星期六"
    End Select
    Lblbottom.Caption = Format(Now(), "YYYY年MM月DD日") + Strweek
    If MSHHead.Visible = False Then Exit Sub
    
    mlngTime = mlngTime - 1
    
    If mlngTime = 0 Then DispInfo
        
    
'    '对指示性提示进行动画显示
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
'票据领用检查及初始
Dim i As Integer
        If gblnBill挂号 Then
            mlng领用ID = CheckUsedBill(4, IIf(mlng领用ID > 0, mlng领用ID, glng挂号ID))
            If mlng领用ID <= 0 Then
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
            '严格：取下一个号码
             mStrBillNo = GetNextBill(mlng领用ID)
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
'对下面的提示性
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
    
    strTmp = zlRegInfo("产品简名")
    If strTmp <> "-" Then
        lblOEM.Caption = strTmp & "软件"
        '处理状态栏图标的OEM策略
        If strTmp = "中联" Then
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
    '功能：返回每个字线的ASCII码
    '参数：
    '返回：
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    strCode = "OEM_"
    For intBit = 1 To Len(strAsk)
        '取每个字的ASCII码
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function

Private Sub tmrLoop_Timer()

    '对指示性提示进行动画显示
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
    Dim CurLeft As Currency          '医保病人剩余的费用
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim aryItem As Variant
    Dim lng项目ID As Long
    
    If Index = 4 Then
        Unload Me
    Else
            '提取数据进行初始化
            '------------------------------------------------------------------------------------------------------------------
            If MSHHead.TextMatrix(MSHHead.Row, 3) = "" Then
                MSHHead.SetFocus
                Exit Sub
            End If
                
            mlngTime = Val(GetPara("主窗体刷新周期")) * 2
            mStrItem = MSHHead.TextMatrix(MSHHead.Row, 1)
            mStrDoctorID = MSHHead.TextMatrix(MSHHead.Row, 7)
            mStrDepart = MSHHead.TextMatrix(MSHHead.Row, 0)
            mStrDoctorName = Trim(MSHHead.TextMatrix(MSHHead.Row, 2))
            mDblPrice = CDbl(Mid(MSHHead.TextMatrix(MSHHead.Row, 3), 1, Len(MSHHead.TextMatrix(MSHHead.Row, 3)) - 1))
            mStrDetailID = CStr(MSHHead.TextMatrix(MSHHead.Row, 4))
            lng项目ID = CLng(MSHHead.TextMatrix(MSHHead.Row, 4))
            mStrDepartID = CStr(MSHHead.TextMatrix(MSHHead.Row, 5))
            mStr号别 = CStr(MSHHead.TextMatrix(MSHHead.Row, 6))
            aryItem = GetRegistPrice(lng项目ID)
                
            Select Case Index
            '------------------------------------------------------------------------------------------------------------------
          
            '------------------------------------------------------------------------------------------------------------------
            Case 1
                frmCheckLogin.ShowLogin Me, "医保卡挂号", mStrDepart, mStrItem, mStrDoctorName, mlng领用ID, mStrBillNo, Val(mStrDoctorID), mDblPrice, Val(mStrDetailID), Val(mStrDepartID), mStr号别
            '------------------------------------------------------------------------------------------------------------------
            Case 2
                frmCheckLogin.ShowLogin Me, "身份证挂号", mStrDepart, mStrItem, mStrDoctorName, mlng领用ID, mStrBillNo, Val(mStrDoctorID), mDblPrice, Val(mStrDetailID), Val(mStrDepartID), mStr号别
            '------------------------------------------------------------------------------------------------------------------
            Case 3
                frmCheckLogin.ShowLogin Me, "ＩＣ卡挂号", mStrDepart, mStrItem, mStrDoctorName, mlng领用ID, mStrBillNo, Val(mStrDoctorID), mDblPrice, Val(mStrDetailID), Val(mStrDepartID), mStr号别
        Case Else
                frmCheckLogin.ShowLogin Me, "就诊卡挂号", mStrDepart, mStrItem, mStrDoctorName, mlng领用ID, mStrBillNo, Val(mStrDoctorID), mDblPrice, Val(mStrDetailID), Val(mStrDepartID), mStr号别, Val(UsrCmd(Index).Tag)
          End Select
            Call DispInfo
    End If
End Sub

Private Sub Load就诊卡()
    Dim strSQL As String, lngIndex As Long
    Dim rsTmp As ADODB.Recordset
    If mbln就诊卡挂号 = False Then Exit Sub
    
    '95138:李南春,2016/4/12,是否刷卡改为读卡性质
    strSQL = "Select Id,名称 From 医疗卡类别 Where 是否启用 = 1 And (substr(读卡性质,1,1) = 1 or substr(读卡性质, 2,1) = 1) And 是否固定 =1 Order By 缺省标志 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName)
    If rsTmp.BOF Then Exit Sub
    lngIndex = 5
    Do While Not rsTmp.EOF
        Load UsrCmd(lngIndex)
        With UsrCmd(lngIndex)
            .Visible = True
            .AutoSize = False
            .TextAligment = 1
            .Caption = Nvl(rsTmp!名称, "就诊卡")
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

