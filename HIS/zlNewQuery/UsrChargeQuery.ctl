VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl UsrChargeQuery 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14685
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   14685
   Begin VB.Timer tmrLoop 
      Left            =   1350
      Top             =   7065
   End
   Begin VB.Timer tmrStop 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   195
      Top             =   2565
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   195
      Top             =   2100
   End
   Begin VB.PictureBox picAsk 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8145
      Index           =   0
      Left            =   2415
      ScaleHeight     =   8145
      ScaleWidth      =   11895
      TabIndex        =   6
      Top             =   -30
      Width           =   11895
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   2250
         TabIndex        =   36
         Top             =   3390
         Width           =   2790
      End
      Begin zl9NewQuery.ctlKeyBoard usrKeyBoard 
         Height          =   3090
         Left            =   1035
         TabIndex        =   35
         Top             =   4125
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   5450
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   540
         Index           =   13
         Left            =   8400
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   45
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   953
         Caption         =   "教您怎么查费用"
         BackColor       =   16185078
         FontSize        =   10.5
         ButtonHeight    =   420
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   11
         Left            =   1110
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1005
         Caption         =   "按就诊卡查"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   12
         Left            =   2700
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1005
         Caption         =   "按住院号查"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   3
         Left            =   4320
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1005
         Caption         =   "按门诊号查"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   14
         Left            =   5940
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1005
         Caption         =   "按病人ID查"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   15
         Left            =   7560
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1005
         Caption         =   "按医保卡查"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   20
         Left            =   1110
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1635
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1005
         Caption         =   "按票据号查"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   21
         Left            =   2745
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1635
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1005
         Caption         =   "按单据号查"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlPicture UsrPic 
         Height          =   3795
         Left            =   4890
         TabIndex        =   24
         Top             =   4125
         Width           =   5175
         _ExtentX        =   8202
         _ExtentY        =   6165
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   25
         Left            =   4335
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1635
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1005
         Caption         =   "按身份证查"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   26
         Left            =   5940
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1005
         Caption         =   "按ＩＣ卡查"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   27
         Left            =   5145
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1005
         Caption         =   " 读  卡  "
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   30
         Left            =   1080
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1005
         Caption         =   "按就诊卡查"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   31
         Left            =   7545
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1635
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1005
         Caption         =   "按就诊卡查"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   32
         Left            =   2730
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1005
         Caption         =   "按就诊卡查"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   33
         Left            =   4290
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   2265
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1005
         Caption         =   "按就诊卡查"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   34
         Left            =   5970
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   2250
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1005
         Caption         =   "按就诊卡查"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   35
         Left            =   7575
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   2250
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1005
         Caption         =   "按就诊卡查"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "您的住院号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1035
         TabIndex        =   38
         Top             =   3450
         Width           =   1155
      End
      Begin VB.Label lblDescrible 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "请您在右下角的数字按钮上依次按您的住院号，然后按""确定""按钮"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1035
         TabIndex        =   37
         Top             =   2985
         Width           =   6090
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   1050
         X2              =   10950
         Y1              =   2910
         Y2              =   2910
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "请选择您的费用查询方式，只须在查询方式上按一下，画上“√”就表示选取中的查询方式。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1110
         TabIndex        =   8
         Top             =   720
         Width           =   8985
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   255
         Picture         =   "UsrChargeQuery.ctx":0000
         Top             =   735
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "选择病人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   120
         Width           =   1260
      End
      Begin VB.Image imgTitle 
         Height          =   660
         Left            =   0
         Picture         =   "UsrChargeQuery.ctx":3072
         Stretch         =   -1  'True
         Top             =   -15
         Width           =   8895
      End
   End
   Begin VB.PictureBox picMsg 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   8955
      TabIndex        =   1
      Top             =   0
      Width           =   8955
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   4
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   900
         Caption         =   "重选病人"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   16
         Left            =   8055
         TabIndex        =   14
         Top             =   30
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   900
         Caption         =   "教您怎么使用？"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名:        性别:    入院时间:          出院时间:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1605
         TabIndex        =   2
         Top             =   180
         Width           =   5250
      End
   End
   Begin MSComctlLib.ImageList ilsImage 
      Left            =   9495
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrChargeQuery.ctx":3899
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrChargeQuery.ctx":3C33
            Key             =   "up"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrChargeQuery.ctx":3FCD
            Key             =   "down"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrChargeQuery.ctx":4367
            Key             =   "menu1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrChargeQuery.ctx":4701
            Key             =   "menu2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrChargeQuery.ctx":5983
            Key             =   "menu3"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrChargeQuery.ctx":5D1D
            Key             =   "menu4"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrChargeQuery.ctx":60B7
            Key             =   "time"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrChargeQuery.ctx":6451
            Key             =   "patient"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrChargeQuery.ctx":67EB
            Key             =   "back"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrChargeQuery.ctx":6B85
            Key             =   "unselect"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrChargeQuery.ctx":6F1F
            Key             =   "select"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrChargeQuery.ctx":72B9
            Key             =   "next"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrChargeQuery.ctx":7653
            Key             =   "finish"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrChargeQuery.ctx":79ED
            Key             =   "clear"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrChargeQuery.ctx":7D87
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTitle 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   585
      Index           =   3
      Left            =   30
      ScaleHeight     =   585
      ScaleWidth      =   8775
      TabIndex        =   18
      Top             =   6105
      Width           =   8775
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   5
         Left            =   30
         TabIndex        =   19
         Top             =   30
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   900
         Caption         =   "日期"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   6
         Left            =   2190
         TabIndex        =   20
         Top             =   30
         Visible         =   0   'False
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   900
         Caption         =   "上一天"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   7
         Left            =   3360
         TabIndex        =   21
         Top             =   30
         Visible         =   0   'False
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   900
         Caption         =   "下一天"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   17
         Left            =   45
         TabIndex        =   28
         Top             =   510
         Visible         =   0   'False
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   900
         Caption         =   "指定月份"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   18
         Left            =   2295
         TabIndex        =   30
         Top             =   510
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   900
         Caption         =   "上月"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   19
         Left            =   3165
         TabIndex        =   31
         Top             =   510
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   900
         Caption         =   "下月"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   22
         Left            =   7230
         TabIndex        =   39
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   900
         Caption         =   "退费费用"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   28
         Left            =   5700
         TabIndex        =   48
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   900
         Caption         =   "收费费用"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   29
         Left            =   4770
         TabIndex        =   49
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   900
         Caption         =   "记帐费用"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin VB.Label lblMont 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "2002-05"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1365
         TabIndex        =   29
         Top             =   675
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblRange 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "2002-05-01"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   915
         TabIndex        =   22
         Top             =   165
         Visible         =   0   'False
         Width           =   1200
      End
   End
   Begin VB.PictureBox picTitle 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   585
      Index           =   2
      Left            =   45
      ScaleHeight     =   585
      ScaleWidth      =   8565
      TabIndex        =   0
      Top             =   1365
      Width           =   8565
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   0
         Left            =   15
         TabIndex        =   3
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   900
         Caption         =   "明细查询"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   1
         Left            =   1410
         TabIndex        =   4
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   900
         Caption         =   "分类查询"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   9
         Left            =   6495
         TabIndex        =   15
         Top             =   30
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   900
         Caption         =   "上翻"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   10
         Left            =   7620
         TabIndex        =   16
         Top             =   30
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   900
         Caption         =   "下翻"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   8
         Left            =   5175
         TabIndex        =   17
         Top             =   45
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   900
         Caption         =   "重读数据"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   510
         Index           =   2
         Left            =   3180
         TabIndex        =   23
         Top             =   30
         Visible         =   0   'False
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   900
         Caption         =   "预缴款查询"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   0
      Left            =   90
      ScaleHeight     =   360
      ScaleWidth      =   9630
      TabIndex        =   12
      Top             =   675
      Width           =   9660
      Begin VB.Label lblWarn 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "报警"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   7800
         TabIndex        =   50
         Top             =   0
         Width           =   750
      End
      Begin VB.Label lblMoneyInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "预缴费用：0.00             已用费用：0.00              剩余费用：0.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   105
         TabIndex        =   13
         Top             =   75
         Width           =   7245
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid msfResult 
      Height          =   3285
      Left            =   270
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2070
      Width           =   4740
      _cx             =   8361
      _cy             =   5794
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   16711680
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16761024
      GridColorFixed  =   16761024
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   345
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      WordWrap        =   -1  'True
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
      WallPaper       =   "UsrChargeQuery.ctx":8121
      WallPaperAlignment=   10
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.PictureBox picTitle 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   1050
      Index           =   1
      Left            =   5205
      ScaleHeight     =   1050
      ScaleWidth      =   4380
      TabIndex        =   40
      Top             =   4650
      Width           =   4380
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   23
         Left            =   3765
         TabIndex        =   42
         Top             =   30
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   741
         Caption         =   "上翻"
         BackColor       =   16777215
         ForeColor       =   12583104
         FontSize        =   10.5
         AutoSize        =   0   'False
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   24
         Left            =   3765
         TabIndex        =   43
         Top             =   540
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   741
         Caption         =   "下翻"
         BackColor       =   16777215
         ForeColor       =   12583104
         FontSize        =   10.5
         AutoSize        =   0   'False
         TextAligment    =   0
      End
      Begin VB.Frame fra 
         BackColor       =   &H00FFC0C0&
         Height          =   30
         Left            =   165
         TabIndex        =   41
         Top             =   960
         Width           =   2400
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目说明:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   4
         Left            =   60
         TabIndex        =   44
         Top             =   75
         Width           =   3225
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "UsrChargeQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private mvarMaxDate As String
Private mvarMinDate As String
Private mvarCurPos As Long
Private mvarRows As Long
Private mvarCurPosPre As Long
Private mvarRowsPre As Long
Private mvar病人id As Long
Private mvar主页id As Long
Private mvar结帐 As Boolean
Private mvar姓名 As String
Private mvarStop As Long                '用户查询信息停留间隔
Private mvarScroll As Long
Private mlngDays As Long
Private mlngLoopCount As Long
Private mbytQueryMode As Byte
Private mstrNO As String
Private mstr费用时间类型 As String
Private mrsDateList As ADODB.Recordset
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Private WithEvents mclsIDKind As clsIDKind
Attribute mclsIDKind.VB_VarHelpID = -1

'######################################################################################################################

Private Sub SelectIDCard(ByVal bytIndex As Byte)
    UsrCmd(3).Check = False
    UsrCmd(11).Check = False
    UsrCmd(30).Check = False
    UsrCmd(31).Check = False
    UsrCmd(32).Check = False
    UsrCmd(33).Check = False
    UsrCmd(34).Check = False
    UsrCmd(35).Check = False
    
    UsrCmd(12).Check = False
    UsrCmd(14).Check = False
    UsrCmd(15).Check = False
    UsrCmd(20).Check = False
    UsrCmd(21).Check = False
    UsrCmd(25).Check = False
    UsrCmd(26).Check = False
    
    UsrCmd(3).Picture = ilsImage.ListImages("unselect")
    UsrCmd(11).Picture = ilsImage.ListImages("unselect")
    UsrCmd(30).Picture = ilsImage.ListImages("unselect")
    UsrCmd(31).Picture = ilsImage.ListImages("unselect")
    UsrCmd(32).Picture = ilsImage.ListImages("unselect")
    UsrCmd(33).Picture = ilsImage.ListImages("unselect")
    UsrCmd(34).Picture = ilsImage.ListImages("unselect")
    UsrCmd(35).Picture = ilsImage.ListImages("unselect")
    
    UsrCmd(12).Picture = ilsImage.ListImages("unselect")
    UsrCmd(14).Picture = ilsImage.ListImages("unselect")
    UsrCmd(15).Picture = ilsImage.ListImages("unselect")
    UsrCmd(20).Picture = ilsImage.ListImages("unselect")
    UsrCmd(21).Picture = ilsImage.ListImages("unselect")
    UsrCmd(25).Picture = ilsImage.ListImages("unselect")
    UsrCmd(26).Picture = ilsImage.ListImages("unselect")
    UsrCmd(22).Picture = ilsImage.ListImages("unselect")
    UsrCmd(22).Tag = ""
     
    UsrCmd(28).Picture = ilsImage.ListImages("select"): UsrCmd(28).Tag = "1"
    UsrCmd(29).Picture = ilsImage.ListImages("select"): UsrCmd(29).Tag = "1"
    
    UsrCmd(bytIndex).Check = True
    UsrCmd(bytIndex).Picture = ilsImage.ListImages("select")
    
        
    Select Case bytIndex
    '------------------------------------------------------------------------------------------------------------------
    Case 14
        
        lblDescrible.Caption = "请您在右下角的数字按钮上依次按您的ID号，然后按“确定”按钮"
        lblCaption.Caption = "您的ID号:"
        
        usrKeyBoard.KeyMode = 2
        DoEvents
        usrKeyBoard.Visible = True
        UsrCmd(27).Visible = False
        UsrPic.Visible = False
        EnterFocus txt(4)
    '------------------------------------------------------------------------------------------------------------------
    Case 3
        
        lblDescrible.Caption = "请您在右下角的数字按钮上依次按您的门诊号，然后按“确定”按钮"
        lblCaption.Caption = "您的门诊号:"
        usrKeyBoard.KeyMode = 2
        DoEvents
        usrKeyBoard.Visible = True
        UsrPic.Visible = False
        UsrCmd(27).Visible = False
        EnterFocus txt(4)
    '------------------------------------------------------------------------------------------------------------------
    Case 11, 30, 31, 32, 33, 34, 35
        
        lblDescrible.Caption = "请您直接在刷卡器上刷您的就诊卡"
        lblCaption.Caption = "请刷卡:"
        usrKeyBoard.KeyMode = 2
        DoEvents
        usrKeyBoard.Visible = False
        UsrCmd(27).Visible = False
        UsrPic.Visible = True
        EnterFocus txt(4)
    '------------------------------------------------------------------------------------------------------------------
    Case 26
        
        lblDescrible.Caption = "请您放入ＩＣ卡，然后按“读卡”按钮"
        lblCaption.Caption = "ＩＣ卡号:"
        
        usrKeyBoard.KeyMode = 2
        usrKeyBoard.Visible = False
         
        UsrPic.Visible = False
        
        UsrCmd(27).Visible = True
        
        EnterFocus txt(4)
    '------------------------------------------------------------------------------------------------------------------
    Case 12, 20, 21, 25
        
        
        If bytIndex = 12 Then
            
            lblDescrible.Caption = "请您在右下角的数字按钮上依次按您的住院号，然后按“确定”按钮"
            lblCaption.Caption = "您的住院号:"
            usrKeyBoard.KeyMode = 2
            
        ElseIf bytIndex = 20 Then
            lblDescrible.Caption = "请您在右下角的数字按钮上依次按您的票据号，然后按“确定”按钮"
            lblCaption.Caption = "您的票据号:"
            usrKeyBoard.KeyMode = 1
        ElseIf bytIndex = 21 Then
            lblDescrible.Caption = "请您在右下角的数字按钮上依次按您的单据号，然后按“确定”按钮"
            lblCaption.Caption = "您的单据号:"
            usrKeyBoard.KeyMode = 1
        ElseIf bytIndex = 25 Then
            lblDescrible.Caption = "请您在右下角的数字按钮上依次按您的身份证号，然后按“确定”按钮"
            lblCaption.Caption = "您的身份证:"
            usrKeyBoard.KeyMode = 1
        End If
        DoEvents
        usrKeyBoard.Visible = True
        UsrCmd(27).Visible = False
        UsrPic.Visible = False
        EnterFocus txt(4)
           
    End Select
End Sub

Public Sub InitLoad()
    '初始化进入
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim IntCount As Integer
    
    Call InitSysPar
    
    picAsk(0).Visible = True
    
    UsrCmd(0).Picture = ilsImage.ListImages("menu1")
    UsrCmd(1).Picture = ilsImage.ListImages("menu2")
    UsrCmd(4).Picture = ilsImage.ListImages("patient")
    
    UsrCmd(5).Picture = ilsImage.ListImages("unselect")
    UsrCmd(5).Tag = ""
    UsrCmd(6).Picture = ilsImage.ListImages("up")
    UsrCmd(7).Picture = ilsImage.ListImages("down")
    UsrCmd(8).Picture = ilsImage.ListImages("refresh")
    UsrCmd(9).Picture = ilsImage.ListImages("up")
    UsrCmd(10).Picture = ilsImage.ListImages("down")
    
    UsrCmd(16).Picture = ilsImage.ListImages("help")
    
    UsrCmd(12).Check = True
    UsrCmd(11).Picture = ilsImage.ListImages("unselect")
    UsrCmd(30).Picture = ilsImage.ListImages("unselect")
    UsrCmd(31).Picture = ilsImage.ListImages("unselect")
    UsrCmd(32).Picture = ilsImage.ListImages("unselect")
    UsrCmd(33).Picture = ilsImage.ListImages("unselect")
    UsrCmd(34).Picture = ilsImage.ListImages("unselect")
    UsrCmd(35).Picture = ilsImage.ListImages("unselect")
    
    UsrCmd(3).Picture = ilsImage.ListImages("unselect")
    UsrCmd(12).Picture = ilsImage.ListImages("select")
    UsrCmd(13).Picture = ilsImage.ListImages("help")
    UsrCmd(27).Picture = ilsImage.ListImages("menu1")
    
    UsrCmd(23).ShowPicture = False
    UsrCmd(24).ShowPicture = False
    
    txt(4).Text = ""
    
    UsrCmd(12).Visible = False
    UsrCmd(3).Visible = False
    
    UsrCmd(11).Visible = False
    UsrCmd(30).Visible = False
    UsrCmd(31).Visible = False
    UsrCmd(32).Visible = False
    UsrCmd(33).Visible = False
    UsrCmd(34).Visible = False
    UsrCmd(35).Visible = False
    
    UsrCmd(14).Visible = False
    UsrCmd(15).Visible = False
    UsrCmd(20).Visible = False
    UsrCmd(21).Visible = False
    UsrCmd(25).Visible = False
    UsrCmd(26).Visible = False
    
    
    strTmp = Trim(zlDatabase.GetPara("查询费用方式", glngSys, 1536, "100000000"))
    '1-就诊卡;2-门诊号;3-住院号;4-病人ID;5-医保卡;6-身份证;7-票据号;8-单据号;9-ＩＣ卡
    Dim strIDKind As String
    
'    strIDKind = "身|身份证号|0|0|18"
    
    Set mclsIDKind = New clsIDKind
    Call mclsIDKind.InitIDKind(1, gcnOracle, glngSys, UserInfo.姓名, UserControl.Extender, 1536, strIDKind)
        
    If Val(Mid(strTmp, 1, 1)) > 0 And mclsIDKind.GetCard(rs) Then
        If rs.RecordCount > 0 Then
            
            IntCount = 0
            Do While Not rs.EOF
                
                IntCount = IntCount + 1
                
                If IntCount < 7 Then
                    Select Case IntCount
                    Case 1
                         UsrCmd(11).Visible = True
                         UsrCmd(11).Caption = rs("全名").Value
                    Case Else
                         UsrCmd(28 + IntCount).Visible = True
                         UsrCmd(28 + IntCount).Caption = rs("全名").Value
                    End Select
                End If
                
                rs.MoveNext
            Loop
            
        End If
    End If
    
'    UsrCmd(11).Visible = IIf(Val(Mid(strTmp, 1, 1)) = 0, False, True)
    
    UsrCmd(12).Visible = IIf(Val(Mid(strTmp, 3, 1)) = 0, False, True)
    UsrCmd(3).Visible = IIf(Val(Mid(strTmp, 2, 1)) = 0, False, True)
    UsrCmd(14).Visible = IIf(Val(Mid(strTmp, 4, 1)) = 0, False, True)
    UsrCmd(20).Visible = IIf(Val(Mid(strTmp, 7, 1)) = 0, False, True)
    UsrCmd(21).Visible = IIf(Val(Mid(strTmp, 8, 1)) = 0, False, True)
    UsrCmd(25).Visible = IIf(Val(Mid(strTmp, 6, 1)) = 0, False, True)
    UsrCmd(26).Visible = IIf(Val(Mid(strTmp, 9, 1)) = 0, False, True)
    UsrCmd(15).Visible = IIf(Val(Mid(strTmp, 5, 1)) = 0, False, True)

    
    If UsrCmd(12).Visible = False And UsrCmd(3).Visible = False And UsrCmd(14).Visible = False And UsrCmd(15).Visible = False _
            And UsrCmd(20).Visible = False And UsrCmd(21).Visible = False And UsrCmd(25).Visible = False And UsrCmd(26).Visible = False _
                Then
        UsrCmd(11).Visible = True
    End If
    
    UsrCmd(9).Enabled = False
    UsrCmd(10).Enabled = False
    
'    On Error Resume Next
'    Set mobjICCard = CreateObject("zlICCard.clsICCard")
'    On Error GoTo 0
    
    tmrStop.Enabled = True
    mvarStop = Val(GetPara("费用查询停留时间", "30"))
    mvarStop = IIf(mvarStop <= 0, 30, mvarStop)
    
    mvarScroll = Val(GetPara("费用查询滚动间隔", "10"))
    mvarScroll = IIf(mvarScroll <= 0, 10, mvarScroll)
    
    gstrSQL = "select A.插图序号 from 咨询段落目录 A where A.页面序号=2 and A.段落序号=1"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "费用查询")
    If gRs.BOF = False Then
        UsrPic.Tag = GetFileName(IIf(IsNull(gRs!插图序号), 0, gRs!插图序号), UsrPic.Width, UsrPic.Height)
        Call UsrPic.ShowPictureByFile(UsrPic.Tag)
    End If
    
    Call AdjustCmdPostion
    
    tmrLoop.Tag = Val(GetPara("返回费用间隔", "0"))
    
    tmrLoop.Interval = 1000
    tmrLoop.Enabled = (Val(tmrLoop.Tag) > 0)
    
    mlngLoopCount = 0
    
    EnterFocus picAsk(0)
    
    Dim intIndex As Integer
    
    If UsrCmd(11).Visible And intIndex = 0 Then intIndex = 11
    If UsrCmd(30).Visible And intIndex = 0 Then intIndex = 30
    If UsrCmd(31).Visible And intIndex = 0 Then intIndex = 31
    If UsrCmd(32).Visible And intIndex = 0 Then intIndex = 32
    If UsrCmd(33).Visible And intIndex = 0 Then intIndex = 33
    If UsrCmd(34).Visible And intIndex = 0 Then intIndex = 34
    If UsrCmd(35).Visible And intIndex = 0 Then intIndex = 35
    
    If UsrCmd(12).Visible And intIndex = 0 Then intIndex = 12
    If UsrCmd(3).Visible And intIndex = 0 Then intIndex = 3
    If UsrCmd(14).Visible And intIndex = 0 Then intIndex = 14
    If UsrCmd(15).Visible And intIndex = 0 Then intIndex = 15
    If UsrCmd(20).Visible And intIndex = 0 Then intIndex = 20
    If UsrCmd(21).Visible And intIndex = 0 Then intIndex = 21
    If UsrCmd(25).Visible And intIndex = 0 Then intIndex = 25
    If UsrCmd(26).Visible And intIndex = 0 Then intIndex = 26
    
    Call UsrCmd_CommandClick(intIndex)
End Sub

Private Function AdjustCmdPostion() As Boolean
    Dim lngX As Long
    Dim lngY As Long
    Dim lngCount As Long
    
    lngX = 1110
    lngY = 1005
    
    If UsrCmd(11).Visible Then
       
       UsrCmd(11).Left = lngX
       UsrCmd(11).Top = lngY
       
       lngCount = lngCount + 1
       
       If lngCount = 5 Or lngCount = 10 Then
            lngX = 1110
            If lngCount = 5 Then lngY = 1635
            If lngCount = 10 Then lngY = 2280
        Else
            lngX = lngX + UsrCmd(11).Width + 60
       End If
       
    End If
    
    If UsrCmd(12).Visible Then
       
       UsrCmd(12).Left = lngX
       UsrCmd(12).Top = lngY
       
       lngCount = lngCount + 1
       
       If lngCount = 5 Or lngCount = 10 Then
            lngX = 1110
            If lngCount = 5 Then lngY = 1635
            If lngCount = 10 Then lngY = 2280
        Else
            lngX = lngX + UsrCmd(12).Width + 60
       End If
       
    End If
    
    If UsrCmd(3).Visible Then
       
       UsrCmd(3).Left = lngX
       UsrCmd(3).Top = lngY
       
       lngCount = lngCount + 1
       
       If lngCount = 5 Or lngCount = 10 Then
            lngX = 1110
            If lngCount = 5 Then lngY = 1635
            If lngCount = 10 Then lngY = 2280
        Else
            lngX = lngX + UsrCmd(3).Width + 60
       End If
       
    End If
    
    If UsrCmd(14).Visible Then
       
       UsrCmd(14).Left = lngX
       UsrCmd(14).Top = lngY
       
       lngCount = lngCount + 1
       
       If lngCount = 5 Or lngCount = 10 Then
            lngX = 1110
            If lngCount = 5 Then lngY = 1635
            If lngCount = 10 Then lngY = 2280
        Else
            lngX = lngX + UsrCmd(14).Width + 60
       End If
       
    End If
    
    If UsrCmd(15).Visible Then
       
       UsrCmd(15).Left = lngX
       UsrCmd(15).Top = lngY
       
       lngCount = lngCount + 1
       
       If lngCount = 5 Or lngCount = 10 Then
            lngX = 1110
            If lngCount = 5 Then lngY = 1635
            If lngCount = 10 Then lngY = 2280
        Else
            lngX = lngX + UsrCmd(15).Width + 60
       End If
       
    End If
    
    If UsrCmd(20).Visible Then
       
       UsrCmd(20).Left = lngX
       UsrCmd(20).Top = lngY
       
       lngCount = lngCount + 1
       
       If lngCount = 5 Or lngCount = 10 Then
            lngX = 1110
            If lngCount = 5 Then lngY = 1635
            If lngCount = 10 Then lngY = 2280
        Else
            lngX = lngX + UsrCmd(20).Width + 60
       End If
       
    End If
    
    If UsrCmd(21).Visible Then
       
       UsrCmd(21).Left = lngX
       UsrCmd(21).Top = lngY
       
       lngCount = lngCount + 1
       
       If lngCount = 5 Or lngCount = 10 Then
            lngX = 1110
            If lngCount = 5 Then lngY = 1635
            If lngCount = 10 Then lngY = 2280
        Else
            lngX = lngX + UsrCmd(21).Width + 60
       End If
       
    End If
    
    If UsrCmd(25).Visible Then
       UsrCmd(25).Left = lngX
       UsrCmd(25).Top = lngY
       
       lngCount = lngCount + 1
       
       If lngCount = 5 Or lngCount = 10 Then
            lngX = 1110
            If lngCount = 5 Then lngY = 1635
            If lngCount = 10 Then lngY = 2280
        Else
            lngX = lngX + UsrCmd(25).Width + 60
       End If
    End If
    
    If UsrCmd(26).Visible Then
       UsrCmd(26).Left = lngX
       UsrCmd(26).Top = lngY
       
       lngCount = lngCount + 1
       
       If lngCount = 5 Or lngCount = 10 Then
            lngX = 1110
            If lngCount = 5 Then lngY = 1635
            If lngCount = 10 Then lngY = 2280
        Else
            lngX = lngX + UsrCmd(26).Width + 60
       End If
    End If
    
    If UsrCmd(31).Visible Then
       UsrCmd(31).Left = lngX
       UsrCmd(31).Top = lngY
       
       lngCount = lngCount + 1
       
       If lngCount = 5 Or lngCount = 10 Then
            lngX = 1110
            If lngCount = 5 Then lngY = 1635
            If lngCount = 10 Then lngY = 2280
        Else
            lngX = lngX + UsrCmd(31).Width + 60
       End If
    End If
    
    If UsrCmd(30).Visible Then
       UsrCmd(30).Left = lngX
       UsrCmd(30).Top = lngY
       
       lngCount = lngCount + 1
       
       If lngCount = 5 Or lngCount = 10 Then
            lngX = 1110
            If lngCount = 5 Then lngY = 1635
            If lngCount = 10 Then lngY = 2280
        Else
            lngX = lngX + UsrCmd(30).Width + 60
       End If
    End If
    
    If UsrCmd(32).Visible Then
       UsrCmd(32).Left = lngX
       UsrCmd(32).Top = lngY
       
       lngCount = lngCount + 1
       
       If lngCount = 5 Or lngCount = 10 Then
            lngX = 1110
            If lngCount = 5 Then lngY = 1635
            If lngCount = 10 Then lngY = 2280
        Else
            lngX = lngX + UsrCmd(32).Width + 60
       End If
    End If
    
    If UsrCmd(33).Visible Then
       UsrCmd(33).Left = lngX
       UsrCmd(33).Top = lngY
       
       lngCount = lngCount + 1
       
       If lngCount = 5 Or lngCount = 10 Then
            lngX = 1110
            If lngCount = 5 Then lngY = 1635
            If lngCount = 10 Then lngY = 2280
        Else
            lngX = lngX + UsrCmd(33).Width + 60
       End If
    End If
    
    If UsrCmd(34).Visible Then
       UsrCmd(34).Left = lngX
       UsrCmd(34).Top = lngY
       
       lngCount = lngCount + 1
       
       If lngCount = 5 Or lngCount = 10 Then
            lngX = 1110
            If lngCount = 5 Then lngY = 1635
            If lngCount = 10 Then lngY = 2280
        Else
            lngX = lngX + UsrCmd(34).Width + 60
       End If
    End If
    
    If UsrCmd(35).Visible Then
       UsrCmd(35).Left = lngX
       UsrCmd(35).Top = lngY
       
       lngCount = lngCount + 1
       
       If lngCount = 5 Or lngCount = 10 Then
            lngX = 1110
            If lngCount = 5 Then lngY = 1635
            If lngCount = 10 Then lngY = 2280
        Else
            lngX = lngX + UsrCmd(35).Width + 60
       End If
    End If
        
End Function



Private Sub lblMsg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    mlngLoopCount = 0
End Sub

Private Sub mclsIDKind_AfterInputComplete(ByVal CardNo As String)
    txt(4).Text = CardNo
    If txt(4).Text <> "" Then
        Call usrKeyBoard_CommandClick("确定")
    End If
End Sub

Private Sub mclsIDKind_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txt(4).Text = "" And Not txt(4).Locked And UserControl.ActiveControl Is txt(4) Then
        txt(4).Text = strID
        Call usrKeyBoard_CommandClick("确定")
    End If
End Sub

Private Sub mfrmCardPass_AfterInputPassword(ByVal strInputPassword As String, blnSucc As Boolean)
    
    
End Sub

Private Sub msfResult_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    Call SelectRow(msfResult, OldRow, NewRow)
    
End Sub

Private Sub msfResult_Click()
    tmrScroll.Enabled = False
    mvarStop = Val(GetPara("费用查询停留时间", "30"))
    mvarStop = IIf(mvarStop <= 0, 30, mvarStop)
End Sub

Private Sub msfResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    mlngLoopCount = 0
End Sub


Private Sub msfResult_RowColChange()
    On Error Resume Next
    
    If UsrCmd(22).Visible = False Then Exit Sub
    
    lbl(4).Caption = msfResult.TextMatrix(msfResult.Row, 8)
    If lbl(4).Caption = "" Then
        picTitle(1).Visible = False
        lbl(4).Caption = "项目说明:"
        
    Else
        picTitle(1).Visible = True
    End If
    msfResult.Move msfResult.Left, msfResult.Top, msfResult.Width, UserControl.Height - msfResult.Top - picTitle(3).Height - IIf(picTitle(1).Visible, picTitle(1).Height, 0)

End Sub

Private Sub picAsk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    mlngLoopCount = 0
End Sub

Private Sub picMsg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    mlngLoopCount = 0
End Sub

Private Sub tmrLoop_Timer()
    
    mlngLoopCount = mlngLoopCount + 1
    
    If mlngLoopCount >= Val(tmrLoop.Tag) Then
        '返回费用登录界面
        
        If picAsk(0).Visible = False Then
            tmrLoop.Enabled = False
        
            Call UsrCmd_CommandClick(4)
        
            tmrLoop.Enabled = (Val(tmrLoop.Tag) > 0)
        End If
        
        mlngLoopCount = 0
    End If
    
End Sub

Private Sub tmrScroll_Timer()
    If mvarScroll > 0 Then
        mvarScroll = mvarScroll - 1
    Else
        If UsrCmd(10).Enabled = False Then
            msfResult.TopRow = 1
            mvarCurPos = 1
            Call EnablePageButton(msfResult, mvarCurPos, mvarRows, UsrCmd(9), UsrCmd(10))
        Else
            Call UsrCmd_CommandClick(10)
        End If
        mvarScroll = Val(GetPara("费用查询滚动间隔", "10"))
        mvarScroll = IIf(mvarScroll <= 0, 10, mvarScroll)
    End If
End Sub

Private Sub tmrStop_Timer()
    If mvarStop > 0 Then
        mvarStop = mvarStop - 1
    Else
        tmrScroll.Enabled = True
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    If mbytQueryMode = 25 Then
        Call mclsIDKind.EnterInputBox("身份证号")
    End If
    
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case mbytQueryMode
    Case 11, 30, 31, 32, 33, 34, 35
        Call mclsIDKind.InputKeyPress(KeyAscii, txt(4))
    End Select
    
    If CheckIsInclude(UCase(Chr(KeyAscii)), "'‘’;；:：?？|,，.。""") = True Then KeyAscii = 0
End Sub

Private Sub txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 4 And mbytQueryMode = 11 And Len(txt(Index).Text) > 0 Then
        
        If CheckIsInclude(UCase(Chr(KeyCode)), "'") = True Then KeyCode = 0
        
        If mclsIDKind.IsCard Then
            If Len(txt(Index).Text) = mclsIDKind.CardLength And KeyCode <> 0 Or KeyCode = 13 Then
                Call usrKeyBoard_CommandClick("确定")
            End If
        Else
            If KeyCode = 13 Then
                Call usrKeyBoard_CommandClick("确定")
            End If
        End If

    ElseIf Index = 2 Then
        If CheckIsInclude(UCase(Chr(KeyCode)), "'") = True Then KeyCode = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    On Error Resume Next
    
    If mbytQueryMode = 25 Then
        Call mclsIDKind.LeaveInputBox
'        If Not (mobjIDCard Is Nothing) Then Call mobjIDCard.SetEnabled(False)
'        Set mobjIDCard = Nothing
    End If
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    mlngLoopCount = 0
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    mlngLoopCount = 0
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
        
    Call ResizeControl(picMsg, 0, 0, UserControl.Width, picMsg.Height)
    Call ResizeControl(picTitle(0), 0, picMsg.Top + picMsg.Height, UserControl.Width, picTitle(0).Height)
    
    '添加欠费提示
    Call ResizeControl(lblWarn, picTitle(0).Width / 4 * 3, 0, lblWarn.Width, lblWarn.Height)
    
    Call ResizeControl(picTitle(2), 0, picTitle(0).Top + picTitle(0).Height, UserControl.Width, picTitle(2).Height)
    
    Call ResizeControl(msfResult, 0, picTitle(2).Top + picTitle(2).Height + 30, UserControl.Width - 15, UserControl.Height - picTitle(2).Top - picTitle(2).Height - picTitle(3).Height - picTitle(1).Height - 30)
    picTitle(1).Move 0, msfResult.Top + msfResult.Height, UserControl.Width
    Fra.Move 0, picTitle(1).Height - 15, picTitle(1).Width
    UsrCmd(23).Move picTitle(1).Width - UsrCmd(23).Width - 120
    UsrCmd(24).Move UsrCmd(23).Left
    
    Call ResizeControl(picTitle(3), 0, picTitle(1).Top + picTitle(1).Height, UserControl.Width, picTitle(3).Height)
'
    UsrCmd(16).Left = picMsg.ScaleWidth - UsrCmd(16).Width - 30
    
    UsrCmd(15).Left = picTitle(1).ScaleWidth - UsrCmd(15).Width - 30
    Call ResizeControl(UsrCmd(14), UsrCmd(15).Left - UsrCmd(14).Width - 90, UsrCmd(15).Top, UsrCmd(14).Width, UsrCmd(14).Height)
    
    Call ResizeControl(UsrCmd(2), UsrCmd(1).Left - UsrCmd(2).Width - 90, UsrCmd(1).Top, UsrCmd(2).Width, UsrCmd(2).Height)
    
    UsrCmd(10).Left = picTitle(2).ScaleWidth - UsrCmd(10).Width - 30
    Call ResizeControl(UsrCmd(9), UsrCmd(10).Left - UsrCmd(9).Width - 90, UsrCmd(10).Top, UsrCmd(9).Width, UsrCmd(9).Height)
    Call ResizeControl(UsrCmd(8), UsrCmd(9).Left - UsrCmd(8).Width - 90, UsrCmd(10).Top, UsrCmd(8).Width, UsrCmd(8).Height)
    
    Call ResizeControl(picAsk(0), 0, 0, UserControl.Width, UserControl.Height)
    Call ResizeControl(imgTitle, 0, 0, UserControl.Width - UsrCmd(13).Width - 120, imgTitle.Height)

    Call ResizeControl(UsrCmd(13), picAsk(0).ScaleWidth - UsrCmd(13).Width - 30, 75, UsrCmd(13).Width, UsrCmd(13).Height)
    
    UsrCmd(22).Left = picTitle(3).ScaleWidth - UsrCmd(22).Width - 30
    UsrCmd(28).Left = UsrCmd(22).Left - 30 - UsrCmd(28).Width
    UsrCmd(29).Left = UsrCmd(28).Left - 30 - UsrCmd(29).Width
    
    lbl(4).Move 0, 0, UsrCmd(23).Left - 30, picTitle(1).Height
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
'    If Not (mobjIDCard Is Nothing) Then Call mobjIDCard.SetEnabled(False)
'    Set mobjIDCard = Nothing
'    Set mobjICCard = Nothing
    
    If Not (mclsIDKind Is Nothing) Then Set mclsIDKind = Nothing
    
End Sub

Private Sub UsrCmd_CommandClick(Index As Integer)
    Dim i As Integer
    Dim strTmp As String
        
    tmrScroll.Enabled = False
    mvarStop = Val(GetPara("费用查询停留时间", "30"))
    mvarStop = IIf(mvarStop <= 0, 30, mvarStop)
    
    Select Case Index
    Case 22
        UsrCmd(Index).Picture = ilsImage.ListImages(IIf(UsrCmd(Index).Tag = "1", "unselect", "select"))
        UsrCmd(Index).Tag = IIf(UsrCmd(Index).Tag = "1", "", "1")
        
        Call UsrCmd_CommandClick(0)
    '------------------------------------------------------------------------------------------------------------------
    Case 28
        
        If UsrCmd(28).Tag = "1" And UsrCmd(29).Tag = "" Then Exit Sub
        
        UsrCmd(Index).Picture = ilsImage.ListImages(IIf(UsrCmd(Index).Tag = "1", "unselect", "select"))
        UsrCmd(Index).Tag = IIf(UsrCmd(Index).Tag = "1", "", "1")
        
        Call UsrCmd_CommandClick(0)
    '------------------------------------------------------------------------------------------------------------------
    Case 29
        
        If UsrCmd(28).Tag = "" And UsrCmd(29).Tag = "1" Then Exit Sub
        
        UsrCmd(Index).Picture = ilsImage.ListImages(IIf(UsrCmd(Index).Tag = "1", "unselect", "select"))
        UsrCmd(Index).Tag = IIf(UsrCmd(Index).Tag = "1", "", "1")
        
        Call UsrCmd_CommandClick(0)
    Case 0, 1               '
        Call DrawMsfHeader(Index)
        mvarCurPos = 1
        mvarRows = 0
                
        UsrCmd(0).State = 0
        UsrCmd(1).State = 0
        UsrCmd(2).State = 0
        UsrCmd(Index).State = -1
        picTitle(1).Visible = (Index = 0)
        
        msfResult.Move msfResult.Left, msfResult.Top, msfResult.Width, UserControl.Height - msfResult.Top - picTitle(3).Height - IIf(picTitle(1).Visible, picTitle(1).Height, 0)
        
        UsrCmd(22).Visible = False
        
        Select Case Index
        Case 0
            UsrCmd(22).Visible = True
            
            If picTitle(3).Visible = False Then
                picTitle(3).Visible = True
                msfResult.Height = msfResult.Height - picTitle(3).Height
            End If
            Call Load费用明细(mvar病人id, mvar主页id)
            
            Call msfResult_RowColChange
            
        Case 1
            If picTitle(3).Visible = False Then
                picTitle(3).Visible = True
                msfResult.Height = msfResult.Height - picTitle(3).Height
            End If
        
            Call Load分类费用(mvar病人id, mvar主页id)
            
        End Select
        Call EnablePageButton(msfResult, mvarCurPos, mvarRows, UsrCmd(9), UsrCmd(10))
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        '刷新病人预交记录
        mvarCurPosPre = 1
        mvarRowsPre = 0
    '------------------------------------------------------------------------------------------------------------------
    Case 4                      '重选病人
        
        txt(4).Text = ""

        UsrCmd(9).Enabled = False
        UsrCmd(10).Enabled = False
        picAsk(0).Visible = True
        
        Call EnterFocus(txt(4))
    '------------------------------------------------------------------------------------------------------------------
    Case 5                      '指定日期
    
        UsrCmd(Index).Tag = IIf(UsrCmd(Index).Tag = "", "1", "")
        UsrCmd(Index).Picture = ilsImage.ListImages(IIf(UsrCmd(Index).Tag = "1", "select", "unselect"))
        
        UsrCmd(6).Visible = IIf(UsrCmd(Index).Tag = "", False, True) And UsrCmd(Index).Visible
        UsrCmd(7).Visible = IIf(UsrCmd(Index).Tag = "", False, True) And UsrCmd(Index).Visible
        
        mrsDateList.Filter = ""
        mrsDateList.Filter = "日期<='" & lblRange.Caption & "'"
        If mrsDateList.RecordCount > 0 Then
            mrsDateList.Sort = "日期 Desc"
            mrsDateList.MoveFirst
            lblRange.Caption = Format(mrsDateList("日期").Value, "yyyy-MM-dd")
        End If
                
        lblRange.Visible = IIf(UsrCmd(Index).Tag = "", False, True) And UsrCmd(Index).Visible
        
        For i = 0 To 1
            If UsrCmd(i).State = -1 Then Call UsrCmd_CommandClick(i)
        Next
        
        Call GetCurrentStateInfo(mvar病人id, mvar主页id, IIf(lblRange.Visible, lblRange.Caption, ""))
        
        Call AdjustEnabled
    '------------------------------------------------------------------------------------------------------------------
    Case 17
        UsrCmd(Index).Tag = IIf(UsrCmd(Index).Tag = "", "1", "")
        UsrCmd(Index).Picture = ilsImage.ListImages(IIf(UsrCmd(Index).Tag = "1", "select", "unselect"))
        
        UsrCmd(18).Visible = IIf(UsrCmd(Index).Tag = "", False, True) And UsrCmd(Index).Visible
        UsrCmd(19).Visible = IIf(UsrCmd(Index).Tag = "", False, True) And UsrCmd(Index).Visible
        ' Visible = IIf(UsrCmd(Index).Tag = "", False, True) And UsrCmd(Index).Visible
        
        For i = 0 To 1
            If UsrCmd(i).State = -1 Then Call UsrCmd_CommandClick(i)
        Next
        Call AdjustEnabled
    '------------------------------------------------------------------------------------------------------------------
    Case 6              '上一天
        mvarCurPos = 1
        mvarRows = 0
        
        mrsDateList.Filter = ""
        mrsDateList.Filter = "日期<'" & lblRange.Caption & "'"
        If mrsDateList.RecordCount > 0 Then
            mrsDateList.Sort = "日期 Desc"
            mrsDateList.MoveFirst
            lblRange.Caption = Format(mrsDateList("日期").Value, "yyyy-MM-dd")
        End If
        
'        lblRange.Caption = Format(CDate(lblRange.Caption) - 1, "YYYY-MM-DD")
        For i = 0 To 1
            If UsrCmd(i).State = -1 Then Call UsrCmd_CommandClick(i)
        Next
        
        Call GetCurrentStateInfo(mvar病人id, mvar主页id, lblRange.Caption)
        
        Call AdjustEnabled
    '------------------------------------------------------------------------------------------------------------------
    Case 7              '下一天
        mvarCurPos = 1
        mvarRows = 0
        
        mrsDateList.Filter = ""
        mrsDateList.Filter = "日期>'" & lblRange.Caption & "'"
        If mrsDateList.RecordCount > 0 Then
            mrsDateList.Sort = "日期"
            mrsDateList.MoveFirst
            lblRange.Caption = Format(mrsDateList("日期").Value, "yyyy-MM-dd")
        End If
        
'        lblRange.Caption = Format(CDate(lblRange.Caption) + 1, "YYYY-MM-DD")
        For i = 0 To 1
            If UsrCmd(i).State = -1 Then Call UsrCmd_CommandClick(i)
        Next
        
        Call GetCurrentStateInfo(mvar病人id, mvar主页id, lblRange.Caption)
        
        Call AdjustEnabled
    '------------------------------------------------------------------------------------------------------------------
    Case 8
        For i = 0 To 1
            If UsrCmd(i).State = -1 Then Call UsrCmd_CommandClick(i)
        Next
    '------------------------------------------------------------------------------------------------------------------
    Case 9
        Call TurnToPage(msfResult, -1, mvarCurPos)
        Call EnablePageButton(msfResult, mvarCurPos, mvarRows, UsrCmd(9), UsrCmd(10))
    '------------------------------------------------------------------------------------------------------------------
    Case 10             '下一页
        Call TurnToPage(msfResult, 1, mvarCurPos)
        Call EnablePageButton(msfResult, mvarCurPos, mvarRows, UsrCmd(9), UsrCmd(10))
    '------------------------------------------------------------------------------------------------------------------
    Case 3, 11, 12, 14, 15, 20, 21, 25, 26, 30, 31, 32, 33, 34, 35
        '11-按就诊卡查询;12-按住院号查询;3-按门诊号查询;14-按病人ID号查询;15-按医保卡查询;20-按票据号查;21-按单据号查;25-按身份证查,26-按ＩＣ卡查
        '11,30,31,32,33,34,35为就诊卡及银行卡等，是由ZLHIS返回的卡类
        
        mbytQueryMode = Index
                
        Select Case Index
        Case 11, 30, 31, 32, 33, 34, 35
            mclsIDKind.LongName = UsrCmd(Index).Caption
            
            If mclsIDKind.IsReadCard Then
                UsrCmd(27).Visible = True
            Else
                UsrCmd(27).Visible = False
            End If
            
        Case Else
        
            UsrCmd(27).Visible = False
            
        End Select
        
        Call SelectIDCard(Index)
        
        If Index = 15 Then
            If CheckIdentify(mvar病人id, mvar主页id) Then
                If mvar病人id <= 0 Then Exit Sub
                Call EnterCharge
            End If
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 27                     '读卡
        
        txt(4).Text = mclsIDKind.ReadICCard
        If txt(4).Text <> "" Then
            Call usrKeyBoard_CommandClick("确定")
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 13, 16
        Call frmHelp.ShowHelp(Me, -2, UserControl.Width, UserControl.Height)
    '------------------------------------------------------------------------------------------------------------------
    Case 23
        If (lbl(4).Top + lbl(4).Height) > picTitle(1).Height Then lbl(4).Top = lbl(4).Top - 210
    '------------------------------------------------------------------------------------------------------------------
    Case 24
        If lbl(4).Top < 0 Then lbl(4).Top = lbl(4).Top + 210
    End Select
End Sub


Private Sub LoadPatient(ByVal 病人id As Long, ByVal 主页id As Long, Optional ByVal str姓名 As String = "")
    '功能:根据病人的住院号获得基本信息
    '参数:str住院号         病人的住院号
    Dim i As Long
    
    On Error GoTo errH
    If 主页id > 0 Then
        gstrSQL = "select A.姓名,A.性别,B.入院日期,B.出院日期 from 病人信息 A,病案主页 B where A.病人id=B.病人id and A.病人id=[1] and B.主页id=[2]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "费用查询", 病人id, 主页id)
        If gRs.RecordCount > 0 Then
            gRs.MoveFirst
            mvar姓名 = Nvl(gRs!姓名)
            lblMsg.Caption = IIf(IsNull(gRs!姓名), "", "姓名:" & gRs!姓名) & IIf(IsNull(gRs!性别), "", "   性别:" & gRs!性别) & IIf(IsNull(gRs!入院日期), "", "   入院日期:" & gRs!入院日期)
            lblRange.Caption = Format(IIf(IsNull(gRs!出院日期), zlDatabase.Currentdate, gRs!入院日期), "YYYY-MM-DD")
            lblRange.Tag = IIf(IsNull(gRs!出院日期), "0", "1")
            Call AdjustEnabled
        End If
    ElseIf 病人id > 0 Then
        gstrSQL = "select A.姓名,A.性别 from 病人信息 A where A.病人id=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "费用查询", 病人id, 主页id)
        If gRs.RecordCount > 0 Then
            gRs.MoveFirst
            mvar姓名 = Nvl(gRs!姓名)
            lblMsg.Caption = IIf(IsNull(gRs!姓名), "", "姓名:" & gRs!姓名) & IIf(IsNull(gRs!性别), "", "   性别:" & gRs!性别) & "(门诊病人)"
            lblRange.Caption = Format(zlDatabase.Currentdate, "YYYY-MM-DD")
            Call AdjustEnabled
        End If
    Else
        lblMsg.Caption = "姓名:" & str姓名
    End If
    
    mlngDays = Val(GetPara("允许查以前的门诊费用", "0"))
    
    
    '11-按就诊卡查询;12-按住院号查询;3-按门诊号查询;14-按病人ID号查询;15-按医保卡查询;20-按票据号查;21-按单据号查;25-按身份证查,26-按ＩＣ卡查
    If 主页id > 0 Then
        gstrSQL = "select max(" & mstr费用时间类型 & ") as MaxDate,min(" & mstr费用时间类型 & ") as MinDate from 住院费用记录 where 病人id=[1] and 主页id=[2] "
    Else
        gstrSQL = " Select max(" & mstr费用时间类型 & ") as MaxDate,min(" & mstr费用时间类型 & ") as MinDate from 门诊费用记录 where 病人id=[1] UNION ALL  " & _
                  " Select max(" & mstr费用时间类型 & ") as MaxDate,min(" & mstr费用时间类型 & ") as MinDate from 住院费用记录 where 病人id=[1]    AND (主页id IS NULL OR 主页id=0) and nvl(门诊标志,0)<>2 "
        gstrSQL = "Select Max(MaxDate) as MaxDate ,min(MinDate) as MinDate From (" & gstrSQL & ")"
    End If
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "费用查询", 病人id, 主页id)
    
    '刘兴洪:因SQL完全相同,因此在大表拆分时,将这个语句屏蔽,同时以上面的SQL语句代替
'    Select Case mbytQueryMode
'    Case 20
'        gstrSQL = "select max(" & mstr费用时间类型 & ") as MaxDate,min(" & mstr费用时间类型 & ") as MinDate from 病人费用记录 where 病人id=[1]" & IIf(主页id > 0, " and 主页id=[2]", " AND (主页id IS NULL OR 主页id=0)")
'    Case 21
'        gstrSQL = "select max(" & mstr费用时间类型 & ") as MaxDate,min(" & mstr费用时间类型 & ") as MinDate from 病人费用记录 where 病人id=[1]" & IIf(主页id > 0, " and 主页id=[2]", " AND (主页id IS NULL OR 主页id=0)")
'    Case Else
'        gstrSQL = "select max(" & mstr费用时间类型 & ") as MaxDate,min(" & mstr费用时间类型 & ") as MinDate from 病人费用记录 where 病人id=[1]" & IIf(主页id > 0, " and 主页id=[2]", " AND (主页id IS NULL OR 主页id=0)")
'    End Select
'
    If gRs.RecordCount > 0 Then
        gRs.MoveFirst
        mvarMaxDate = Format(IIf(IsNull(gRs!MaxDate), zlDatabase.Currentdate, gRs!MaxDate), "YYYY-MM-DD")
        mvarMinDate = Format(IIf(IsNull(gRs!MinDate), zlDatabase.Currentdate, gRs!MinDate), "YYYY-MM-DD")
        If 主页id <= 0 And mlngDays > 0 Then
            mvarMaxDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD")
            mvarMinDate = Format(DateAdd("d", 0 - (mlngDays - 1), CDate(mvarMaxDate)), "YYYY-MM-DD")
        End If
        lblRange.Caption = Format(IIf(lblRange.Tag = "1", mvarMinDate, zlDatabase.Currentdate), "YYYY-MM-DD")
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Function GetStateInfo(ByVal lng病人ID As Long, ByVal lng主页id As Long) As Boolean
    
    '******************************************************************************************************************
    '功能：读取病人费用概况并显示
    '--病人费用概况：汇总反映病人的费用概要情况。
    '--  预交款总额: 从病人预交记录中汇总
    '--  冲预交总额: 从病人结帐记录中汇总
    '--  费用总额:   从病人费用记录中汇总
    '--  结帐总额:   从病人结帐金额中汇总
    '--从上述数据，可以直接计算：
    '--  结帐补退款=结帐总额-冲预交总额
    '--  预交款余额=预交款总额-冲预交总额
    '--  未结帐总额=费用总额-结帐总额
    '--  剩余款总额=预交款余额-未结帐总额
    '******************************************************************************************************************
    
    On Error GoTo errH
    Dim bytDec As Byte
    Dim strDec As String
    Dim strSQL As String
    
    Dim lng级别 As Long
    Dim lng报警级别 As Long
    Dim lng已报警级别 As Long
    Dim sng剩余款额 As Single
    Dim str报警方案 As String
    Dim rsTmp As New ADODB.Recordset
    
    
    '未结费用概况
    
    '费用金额小数点位数
    
    
    strDec = "0.00"
    bytDec = 2
    bytDec = zlDatabase.GetPara(9, glngSys, , 2)
    strDec = "0." & String(bytDec, "0")
    
    strSQL = "Select 预交余额,费用余额,0 as 预结费用 From 病人余额 Where 性质=1 And 病人ID=[1] And 类型=" & IIf(lng主页id = 0, "1 ", "2 ")
    strSQL = strSQL & " Union ALL " & _
        " Select 0 as 预交余额,0 as 费用余额,Sum(B.金额) as 预结费用" & _
        " From 病人信息 A,保险模拟结算 B" & _
        " Where A.病人ID=B.病人ID And A.住院次数=B.主页ID" & _
        " And A.病人ID=[1] "
    
    strSQL = "Select Sum(预交余额) as 预交余额,Sum(费用余额) as 费用余额,Sum(预结费用) as 预结费用 From (" & strSQL & ")"

    Set gRs = zlDatabase.OpenSQLRecord(strSQL, "费用查询", lng病人ID)
    If Not gRs.EOF Then
        sng剩余款额 = IIf(IsNull(gRs!预交余额), 0, gRs!预交余额) - IIf(IsNull(gRs!费用余额), 0, gRs!费用余额) + IIf(IsNull(gRs!预结费用), 0, gRs!预结费用)
        lblMoneyInfo.Caption = "预交余额:" & Format(zlCommFun.Nvl(gRs!预交余额, 0), "0.00") & Space(3) & _
                            "未结费用:" & Format(zlCommFun.Nvl(gRs!费用余额, 0), strDec) & Space(3) & _
                            "预结费用:" & Format(zlCommFun.Nvl(gRs!预结费用, 0), strDec) & Space(3) & _
                            "剩余款额:" & Format(zlCommFun.Nvl(gRs!预交余额, 0) - zlCommFun.Nvl(gRs!费用余额, 0) + zlCommFun.Nvl(gRs!预结费用, 0), "0.00")
        
        '清空上次的提示信息
        lblWarn.BackStyle = 0
        lblWarn.Caption = ""
        
        If mvar主页id > 0 Then
            str报警方案 = ""
            strSQL = "Select zl_PatiWarnScheme([1],[2]) As 报警方案 From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "欠费情况", mvar病人id, mvar主页id)
            If rsTmp.BOF = False Then
                str报警方案 = zlCommFun.Nvl(rsTmp("报警方案").Value)
            End If
                        
            Call 欠费情况(mvar姓名, lng病人ID, mvar主页id, sng剩余款额, str报警方案)
        End If
        
        GetStateInfo = True
        
    Else
        lblMoneyInfo.Caption = "预交余额:0.00" & Space(3) & "预结费用:0.00" & Space(3) & "未结费用:" & strDec & Space(3) & "剩余款额:0.00"
    End If
    lblMoneyInfo.Tag = lblMoneyInfo.Caption
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCurrentStateInfo(ByVal lng病人ID As Long, ByVal lng主页id As Long, ByVal strCurrDate As String) As Boolean
    
    '******************************************************************************************************************
    '功能：读取当日费用状态
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim bytDec As Byte
    Dim strDec As String
    Dim strSQL As String

    
    On Error GoTo errH
    
    '费用金额小数点位数
    
    If strCurrDate = "" Then
        
        lblMoneyInfo.Caption = lblMoneyInfo.Tag & Space(3) & "费用合计:" & Format(msfResult.Tag, strDec)
                            
        GetCurrentStateInfo = True
        Exit Function
    End If

    strDec = "0.00"
    bytDec = 2
    bytDec = zlDatabase.GetPara(9, glngSys, , 2)

    strDec = "0." & String(bytDec, "0")

    strSQL = "" & _
    "       Select 预交余额, 费用余额" & vbNewLine & _
    "       From 病人余额" & vbNewLine & _
    "       Where 性质 = 1 And 病人id = [1] And  类型=" & IIf(lng主页id = 0, "1 ", "2 ") & vbNewLine & _
    "       Union All" & vbNewLine & _
    "       Select 0 - Nvl(A.金额, 0) As 预交余额, 0 As 费用余额" & vbNewLine & _
    "       From 病人预交记录 A" & vbNewLine & _
    "       Where A.病人id = [1] And Nvl(A.主页id, 0) = [2] And A.记录性质 = 1 And A.收款时间 > [3]" & vbNewLine & _
    "       Union All" & vbNewLine & _
    "       Select Nvl(A.冲预交, 0) As 预交余额, 0 As 费用余额" & vbNewLine & _
    "       From 病人预交记录 A, 病人结帐记录 B" & vbNewLine & _
    "       Where A.病人id = [1] And Nvl(A.主页id, 0) = [2] And A.记录性质 = 11 And B.ID(+) = A.结帐id And  Nvl(B.收费时间, A.收款时间) > [3]" & vbNewLine
    
    If lng主页id = 0 Then
        strSQL = strSQL & _
        "       Union All" & vbNewLine & _
        "       Select 0 As 预交余额, 0 - Nvl(实收金额, 0) As 费用余额" & vbNewLine & _
        "       From 门诊费用记录 A" & vbNewLine & _
        "       Where 病人id = [1]  And 记帐费用 = 1 And  登记时间 > [3]" & vbNewLine
    End If
    '97524:李南春,2016/6/16,结帐金额取门诊、住院费用记录中的字段
    strSQL = strSQL & _
    "       Union All" & vbNewLine & _
    "       Select 0 As 预交余额, 0 - Nvl(实收金额, 0) As 费用余额" & vbNewLine & _
    "       From 住院费用记录 A" & vbNewLine & _
    "       Where 病人id = [1] And Nvl(主页id, 0) = [2] And 记帐费用 = 1 And  登记时间 > [3]" & vbNewLine & _
    "       Union All" & vbNewLine & _
    "       Select 0 As 预交余额, Nvl(A.结帐金额, 0) As 费用余额" & vbNewLine & _
    "       From 住院费用记录 A, 病人结帐记录 B" & vbNewLine & _
    "       Where A.病人id = [1] And Nvl(A.主页id, 0) = [2] And A.记帐费用 = 1 And B.ID = A.结帐id And  B.收费时间 > [3] "
    If lng主页id = 0 Then
        strSQL = strSQL & _
        "       Union All" & vbNewLine & _
        "       Select 0 As 预交余额, Nvl(A.结帐金额, 0) As 费用余额" & vbNewLine & _
        "       From 门诊费用记录 A, 病人结帐记录 B" & vbNewLine & _
        "       Where A.病人id = [1] And A.记帐费用 = 1 And B.ID = A.结帐id And  B.收费时间 > [3] "
     End If
     
    strSQL = "" & _
    "Select Sum(预交余额) As 当日预交余额, Sum(费用余额) As 当日费用余额, Sum(预交余额) - Sum(费用余额)" & vbNewLine & _
    "From ( " & strSQL & ") "
    

    strCurrDate = Format(strCurrDate, "yyyy-MM-dd") & " 23:59:59"
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "费用查询", lng病人ID, lng主页id, CDate(strCurrDate))
    If rs.BOF = False Then
        
    
        lblMoneyInfo.Caption = lblMoneyInfo.Tag & Space(3) & "当日剩余款:" & Format(zlCommFun.Nvl(rs!当日预交余额, 0) - zlCommFun.Nvl(rs!当日费用余额, 0), strDec) & Space(3) & _
                            "当日费用合计:" & Format(msfResult.Tag, strDec)
            
        GetCurrentStateInfo = True
        
    Else
        lblMoneyInfo.Caption = lblMoneyInfo.Tag & Space(3) & "当日剩余款:0.00" & Space(3) & "当日费用合计:0.00"
    End If

    Exit Function
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Load预交费用(ByVal lng病人ID As Long, ByVal lng主页id As Long) As Boolean
'功能:装载病人的所有预交费用记录
'参数:lng病人id         病人唯一号
'     lng主页id         病人住院次数
'返回:True成功;False:失败
    Dim sglMoney(1 To 3) As Single
    Dim i As Long
    
    On Error GoTo errH
    
    i = 1
    msfResult.Rows = 2
    Call ClearSpecRowCol(msfResult, 1, Array())
    
    
    gstrSQL = "Select To_Char(A.收款时间,'YYYY-MM-DD') as 日期,A.NO as 单据号," & _
            " B.名称 as 科室,A.结算方式,A.结算号码," & _
            " Ltrim(To_Char(Sum(Nvl(A.金额,0)),'9999999990.00')) as 预交金额," & _
            " Ltrim(To_Char(Sum(Nvl(A.冲预交,0)),'9999999990.00')) as 结帐金额," & _
            " Ltrim(To_Char(Sum(Nvl(A.金额,0)-Nvl(A.冲预交,0)),'9999999990.00')) as 剩余金额," & _
            " A.摘要" & _
            " From 病人预交记录 A,部门表 B" & _
            " Where A.科室ID=B.ID(+) And A.记录性质 IN(1,11) And Nvl(A.金额,0)<>Nvl(A.冲预交,0) And A.病人ID=[1] and A.主页id= [2] " & _
            " Having Sum(Nvl(A.金额, 0) - Nvl(A.冲预交, 0)) <> 0" & _
            " Group by To_Char(收款时间,'YYYY-MM-DD'),A.NO,B.名称,A.结算方式,A.结算号码,A.摘要" & _
            " Order by 日期,单据号"
        
    gstrSQL = "Select To_Char(A.收款时间,'YYYY-MM-DD') as 日期,A.NO as 单据号," & _
            " B.名称 as 科室,A.结算方式,A.结算号码," & _
            " Ltrim(To_Char(Sum(Nvl(A.金额,0)),'9999999990.00')) as 预交金额," & _
            " Ltrim(To_Char(Sum(Nvl(A.冲预交,0)),'9999999990.00')) as 结帐金额," & _
            " Ltrim(To_Char(Sum(Nvl(A.金额,0)-Nvl(A.冲预交,0)),'9999999990.00')) as 剩余金额," & _
            " A.摘要" & _
            " From 病人预交记录 A,部门表 B" & _
            " Where A.科室ID=B.ID(+) And A.记录性质 IN(1,11) And A.病人ID=[1] and A.主页id= [2]" & _
            " Group by To_Char(收款时间,'YYYY-MM-DD'),A.NO,B.名称,A.结算方式,A.结算号码,A.摘要" & _
            " Order by 日期,单据号"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "费用查询", lng病人ID, lng主页id)
    If gRs.RecordCount > 0 Then
        gRs.MoveFirst
        While Not gRs.EOF
            msfResult.TextMatrix(i, 0) = IIf(IsNull(gRs!日期), "", gRs!日期)
            msfResult.TextMatrix(i, 1) = IIf(IsNull(gRs!科室), "", gRs!科室)
            msfResult.TextMatrix(i, 2) = IIf(IsNull(gRs!结算方式), "", gRs!结算方式)
            msfResult.TextMatrix(i, 3) = IIf(IsNull(gRs!结算号码), "", gRs!结算号码)
            msfResult.TextMatrix(i, 4) = IIf(IsNull(gRs!预交金额), "", Format(gRs!预交金额, "0.00"))
            sglMoney(1) = sglMoney(1) + Val(msfResult.TextMatrix(i, 4))
            i = i + 1
            msfResult.Rows = i + 1
            gRs.MoveNext
        Wend
        If msfResult.Rows > 2 Then
            msfResult.MergeRow(msfResult.Rows - 1) = True
            For i = 0 To msfResult.Cols - 2
                msfResult.TextMatrix(msfResult.Rows - 1, i) = "合计"
            Next
            msfResult.TextMatrix(msfResult.Rows - 1, 4) = Format(sglMoney(1), "0.00")
        End If
        
        mvarRows = msfResult.Rows - 1
    End If
    
    msfResult.Rows = msfResult.Rows + 50
    
    Load预交费用 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Load费用明细(ByVal lng病人ID As Long, ByVal lng主页id As Long) As Boolean
    '******************************************************************************************************************
    '功能:装载病人指定日期的未结/已结费用记录
    '参数:lng病人id         病人唯一号
    '     lng主页id         病人住院次数
    '返回:True成功;False:失败
    '******************************************************************************************************************
    Dim i As Long
    Dim strTmp As String
    Dim sglMoney As Single
    Dim v_Class As String
    Dim strBill As String
    Dim strSQL As String
    Dim blnShowExit As Boolean
    Dim str记帐费用 As String
    Dim strFields As String, strWhere As String, strSubTable As String
    
    On Error GoTo errH
    
    '------------------------------------------------------------------------------------------------------------------
    v_Class = GetPara("费用不显明细")
    i = 1
    msfResult.Rows = 2
    Call ClearSpecRowCol(msfResult, 1, Array())
    msfResult.Tag = ""
    
    
    blnShowExit = (UsrCmd(22).Visible And UsrCmd(22).Tag = "1")
    
    Select Case mbytQueryMode
    '------------------------------------------------------------------------------------------------------------------
    Case 20, 21                     '按票据号和单据号查费用
        If mbytQueryMode = 20 Then
            '票据号
            strSQL = " Where 数据性质=1 And ID=(Select 打印ID From (Select 号码,打印ID,票种,性质 From 票据使用明细  Union All Select 号码,打印ID,票种,性质 From H票据使用明细) Where 票种=1 And 性质=1 And 号码=[8] And Rownum=1)"
            strSQL = "Select NO From 票据打印内容 " & strSQL & " Union All Select NO From H票据打印内容" & strSQL
            strBill = " And A.NO In (" & strSQL & ") "
        Else
            '单据号
            strBill = " And A.No=[8] "
        End If
        
        gstrSQL = ""
        
        '对于不显示费用明细的类别,只显示类别的一个总额
        '--------------------------------------------------------------------------------------------------------------
        If v_Class <> "" Then
        
            gstrSQL = "" & _
            " Select   '' As 说明, '' as 收费编码, B.编码, AA.日期, '' as 单据号, '' as 科室, B.名称 AS 类别, '' as 名称," & _
            "       0 as 数次, '' as 数量, '' as 单价, LTRIM(TO_CHAR(AA.实收金额,'99999990.00')) as 金额,1 as 记录状态 " & _
            " From (    Select Trunc(" & mstr费用时间类型 & ") as 日期, 收费类别, Sum(NVL(实收金额,0)) As 实收金额 " & _
            "           From 门诊费用记录 A , Table(Cast(f_Str2list([7]) As zlTools.t_Strlist)) b " & _
            "           Where  A.记录性质=1  And A.门诊标志=1  And A.记录状态<>0 AND A.收费类别=B.Column_Value " & strBill & " " & _
            "           Group by trunc(" & mstr费用时间类型 & ") ,收费类别 " & _
            "       ) AA, 收费项目类别 B " & _
            "Where AA.收费类别=B.编码"
        End If
        
        If v_Class <> "" Then strTmp = " And Not Exists(Select 1 From  Table(Cast(f_Str2list([7]) As zlTools.t_Strlist)) b where A.收费类别 =B.Column_Value) "

        'blnShowExit:显示退费:false=:不显示退费
        strFields = IIf(blnShowExit, "A.付数, A.数次, A.实收金额,A.记录状态 ", "Decode(a.记录状态,1,Nvl(A.付数,1),1) As 付数, Sum(Decode(a.记录状态,1,1, Nvl(A.付数,1))*A.数次) As 数次, Sum(A.实收金额) as 实收金额,1 As 记录状态")

        gstrSQL = IIf(gstrSQL <> "", gstrSQL & " Union ", "") & _
        "Select B.说明, B.编码 As 收费编码,C.编码,A.日期, A.NO as 单据号, D.名称 as 科室, C.名称 AS 类别, B.名称||Decode(nvl(b.费用类型,''),'','','('|| b.费用类型 || ')')  as 名称," & _
        "       A.数次, Decode(A.记录性质,4,'',Decode(A.计算单位,NULL,NULL,A.计算单位)||Decode(A.收费类别,'7','×'||A.付数||'付',NULL)) as 数量 ," & _
        "       LTRIM(TO_CHAR(A.标准单价,'9999990.0000')) as 单价, LTRIM(TO_CHAR(A.实收金额,'99999990.00')) as 金额,A.记录状态 " & _
        "From ( Select trunc(A." & mstr费用时间类型 & ") as 日期, A.NO, A.开单部门ID, A.收费类别, A.收费细目ID, A.计算单位, A.记录性质, A.标准单价," & strFields & _
        "       From   门诊费用记录 A " & _
        "       Where  门诊标志=1  and A.记录状态<>0  and A.记录性质=1  and nvl(A.实收金额,0)<>0 " & strTmp & strBill & _
                IIf(blnShowExit, "          Order By ", " Group By  ") & "  Trunc(A." & mstr费用时间类型 & "),A.NO,A.开单部门ID,A.收费类别,A.收费细目ID,A.计算单位" & IIf(blnShowExit, "", ",A.记录性质,A.标准单价,Decode(a.记录状态,1,Nvl(A.付数,1),1) ") & _
        "       ) A, 收费项目目录 B, 收费项目类别 C, 部门表 D " & _
        " Where A.收费类别=C.编码  And A.收费细目ID=B.ID  And A.开单部门ID=D.ID(+)  "
 
        gstrSQL = "select 说明, 收费编码,编码,日期,单据号,科室,类别,名称,数次,数量,单价,金额,记录状态 from (" & gstrSQL & ") where 金额<>0 order by 编码,日期 Desc,单据号"
    '------------------------------------------------------------------------------------------------------------------
    Case Else
        strWhere = ""
        If UsrCmd(28).Visible And UsrCmd(28).Tag = "1" And UsrCmd(29).Visible And UsrCmd(29).Tag = "1" Then
            
        ElseIf UsrCmd(28).Visible And UsrCmd(28).Tag = "1" Then
            '收费费用
            strWhere = strWhere & " And A.记帐费用<>1"
        ElseIf UsrCmd(29).Visible And UsrCmd(29).Tag = "1" Then
            '记帐费用
             strWhere = strWhere & "And A.记帐费用=1"
        End If
        
        If lblRange.Visible Then strWhere = strWhere & " And (A." & mstr费用时间类型 & ">=[5] And A." & mstr费用时间类型 & "<[6]  ) "
        
        If lng主页id <= 0 And mlngDays > 0 Then strTmp = " And A." & mstr费用时间类型 & " Between [3] And [4]    "
        If v_Class <> "" Then strTmp = strTmp & " And Not Exists(Select 1 From  Table(Cast(f_Str2list([7]) As zlTools.t_Strlist)) b where A.收费类别 =B.Column_Value) "
        
    
         
        'blnShowExit:显示退费:false=:不显示退费
        strFields = IIf(blnShowExit, "A.付数,A.数次,A.实收金额,A.记录状态  ", "Decode(a.记录状态,1,Nvl(A.付数,1),1) As 付数,Sum(Decode(a.记录状态,1,1, Nvl(A.付数,1))*A.数次) As 数次,Sum(A.实收金额) as 实收金额,1 As 记录状态")
        
        gstrSQL = ""
        If v_Class <> "" Then
            
            strSubTable = "" & _
                "        Select Trunc(" & mstr费用时间类型 & ") as 日期, 收费类别, Sum(实收金额) as 实收金额 " & _
                "        From 住院费用记录 A, Table(Cast(f_Str2list([7]) As zlTools.t_Strlist)) b " & _
                "        Where  A.记录状态<>0 And 病人ID=[1] And nvl(主页ID,0)=[2] And a.收费类别=B.Column_Value " & strWhere & _
                "        Group By Trunc(" & mstr费用时间类型 & "),收费类别 "
            If lng主页id = 0 Then
                strSubTable = strSubTable & "" & _
                "        Union ALL  " & _
                "        Select Trunc(" & mstr费用时间类型 & ") as 日期, 收费类别, Sum(实收金额) as 实收金额 " & _
                "        From 门诊费用记录 A, Table(Cast(f_Str2list([7]) As zlTools.t_Strlist)) b " & _
                "        Where  A.记录状态<>0 And 病人ID=[1]  And a.收费类别=B.Column_Value " & strWhere & _
                "        Group By Trunc(" & mstr费用时间类型 & "),收费类别 "
            End If
            
            gstrSQL = "" & _
            " Select '' As 说明, '' as 收费编码, B.编码, AA.日期, '' as 单据号, '' as 科室, B.名称 AS 类别, '' as 名称," & _
            "       0 as 数次, '' as 数量, '' as 单价, LTRIM(TO_CHAR(Sum(AA.实收金额),'99999990.00')) as 金额,1 as 记录状态 " & _
            " From ( " & strSubTable & ")AA, 收费项目类别 B " & _
            " Where AA.收费类别=B.编码 " & _
            " Group by AA.日期,b.编码,b.名称"
        End If
        
        strSubTable = "" & _
        "           Select Trunc(A." & mstr费用时间类型 & ") as 日期, A.NO, A.开单部门ID, A.收费类别, A.收费细目ID, A.计算单位, A.记录性质, A.标准单价,A.付数,A.数次,A.记录状态,A.实收金额 " & _
        "           From 住院费用记录 A  " & _
        "           Where   A.记录状态<>0 And A.病人ID=[1] And nvl(主页ID,0)=[2] " & strTmp & strWhere & _
        "                   And Nvl(A.实收金额,0)<>0 "
        If lng主页id = 0 Then
            strSubTable = strSubTable & "" & _
            "        Union ALL  " & _
        "           Select Trunc(A." & mstr费用时间类型 & ") as 日期, A.NO, A.开单部门ID, A.收费类别, A.收费细目ID, A.计算单位, A.记录性质, A.标准单价,A.付数,A.数次,A.记录状态,A.实收金额 " & _
            "        From 门诊费用记录 A  " & _
            "        Where   A.记录状态<>0 And A.病人ID=[1] " & strTmp & strWhere & _
            "                   And Nvl(A.实收金额,0)<>0 "
        End If
        
        strSubTable = "" & _
        "   Select trunc(日期) as 日期,NO,开单部门ID,收费类别,收费细目ID,计算单位,记录性质,标准单价," & strFields & _
        "   From (" & strSubTable & ") A" & _
            IIf(blnShowExit, "          Order By ", " Group By  ") & " Trunc(A.日期),A.NO,A.开单部门ID,A.收费类别,A.收费细目ID,A.计算单位 " & IIf(blnShowExit, "", ",A.记录性质,A.标准单价,Decode(a.记录状态,1,Nvl(A.付数,1),1) ")

        '48546:单位加上括号
        gstrSQL = IIf(gstrSQL <> "", gstrSQL & " Union ", "") & _
        " Select B.说明,  B.编码 As 收费编码, C.编码,A.日期, A.No As 单据号, D.名称 As 科室, C.名称 AS 类别, B.名称||decode(nvl(B.费用类型,''),'','','('|| B.费用类型 || ')')  As 名称, A.数次, " & _
        "       Decode(A.记录性质,4,'',Decode(A.计算单位,NULL,NULL,'('||A.计算单位||')')||Decode(A.收费类别,'7','×'||A.付数||'付',NULL)) As 数量 , " & _
        "       LTRIM(TO_CHAR(A.标准单价,'9999990.0000')) As 单价, LTRIM(TO_CHAR(A.实收金额,'99999990.00')) As 金额,A.记录状态  " & _
        " From ( " & strSubTable & ") A, 收费项目目录 B, 收费项目类别 C, 部门表 D " & _
        " Where A.收费类别=C.编码 And A.收费细目ID=B.ID And A.开单部门ID=D.ID(+)  "
            
        gstrSQL = "Select 说明, 收费编码,编码,日期,单据号,科室,类别,名称,数次,数量,单价,金额,记录状态 from (" & gstrSQL & ") where 金额<>0 order by 编码,日期 Desc,单据号"
    End Select
    
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "费用查询", lng病人ID, lng主页id, CDate(Format(mvarMinDate, "YYYY-MM-DD") & " 00:00:00"), CDate(Format(mvarMaxDate, "YYYY-MM-DD") + " 23:59:59"), CDate(lblRange.Caption), CDate(lblRange.Caption) + 1, Replace(v_Class, "'", ""), mstrNO)
 
    If gRs.RecordCount > 0 Then
        gRs.MoveFirst
        While Not gRs.EOF
            msfResult.TextMatrix(i, 0) = IIf(IsNull(gRs!日期), "", gRs!日期)
            msfResult.TextMatrix(i, 1) = IIf(IsNull(gRs!科室), "", gRs!科室)
            msfResult.TextMatrix(i, 2) = IIf(IsNull(gRs!类别), "", gRs!类别)
            msfResult.TextMatrix(i, 3) = IIf(IsNull(gRs!名称), "", gRs!名称)
            
            If IsNull(gRs!数次) = False Then
                msfResult.TextMatrix(i, 4) = IIf(gRs!数次 = 0, "", gRs!数次 & IIf(IsNull(gRs!数量), "", gRs!数量))
            End If
            
            msfResult.TextMatrix(i, 5) = IIf(IsNull(gRs!单价), "", Format(gRs!单价, "0.00##"))
            msfResult.TextMatrix(i, 6) = IIf(IsNull(gRs!金额), "", Format(gRs!金额, "0.00"))
            msfResult.TextMatrix(i, 7) = IIf(IsNull(gRs!收费编码), "", gRs!收费编码)
            msfResult.TextMatrix(i, 8) = IIf(IsNull(gRs!说明), "", gRs!说明)
            
            sglMoney = sglMoney + Val(msfResult.TextMatrix(i, 6))
            If Val(IIf(IsNull(gRs!记录状态), "0", gRs!记录状态)) = 2 Then
                '设置前景色
                msfResult.Cell(flexcpForeColor, i, 0, i, msfResult.Cols - 1) = 255
            End If
            
            i = i + 1
            msfResult.Rows = i + 1
            gRs.MoveNext
        Wend
        
        If msfResult.Rows > 2 Then
            msfResult.MergeRow(msfResult.Rows - 1) = True
            For i = 0 To 5
                msfResult.TextMatrix(msfResult.Rows - 1, i) = "合计"
            Next
            msfResult.TextMatrix(msfResult.Rows - 1, 6) = Format(sglMoney, "0.00")
            msfResult.Tag = sglMoney
        End If
        
        mvarRows = msfResult.Rows - 1
    End If
    msfResult.Rows = msfResult.Rows + 50
    
    Load费用明细 = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetDateList(ByVal lng病人ID As Long, ByVal lng主页id As Long, ByVal strMinDate As String, ByVal strMaxDate As String, Optional ByVal strDateType As String = "发生时间") As ADODB.Recordset
    '******************************************************************************************************************
    '功能:获取有费用数据的日期清单
    '参数:lng病人id         病人唯一号
    '     lng主页id         病人住院次数
    '******************************************************************************************************************
    
    Dim strSQL As String
    
    
    strSQL = "" & _
    "   Select To_Char(" & strDateType & ",'yyyy-mm-dd') As 日期 " & _
    "   From 住院费用记录 " & _
    "   Where 病人id=[1] And Nvl(主页id,0)=[2] And " & strDateType & " Between [3] And [4] " & _
    "   Group by To_Char(" & strDateType & ",'yyyy-mm-dd')"
    If lng主页id = 0 Then
        strSQL = strSQL & " UNION " & _
        "   Select To_Char(" & strDateType & ",'yyyy-mm-dd') As 日期 " & _
        "   From 门诊费用记录 " & _
        "   Where 病人id=[1] And " & strDateType & " Between [3] And [4] " & _
        "   Group by To_Char(" & strDateType & ",'yyyy-mm-dd')"
    End If
    Set GetDateList = zlDatabase.OpenSQLRecord(strSQL, "费用查询", lng病人ID, lng主页id, CDate(strMinDate), CDate(strMaxDate & " 23:59:59"))

    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function
 

Private Function Load分类费用(ByVal lng病人ID As Long, ByVal lng主页id As Long) As Boolean
'功能:装载病人指定日期的未结/已结费用记录
'参数:lng病人id         病人唯一号
'     lng主页id         病人住院次数
'返回:True成功;False:失败
    Dim dtStartDate As Date, dtEndDate As Date, dtStartDate1 As Date, dtEndDate1 As Date
    Dim i As Long, j As Long
    Dim strTmp As String, strWhere As String
    Dim sglMoney(1 To 3) As Single
    Dim svrSave As String
    Dim strSQL As String
    Dim strBill As String
    Dim str记帐费用 As String
    
    On Error GoTo errH
    
    dtStartDate = CDate("1901-01-01"): dtEndDate = dtStartDate: dtStartDate1 = dtStartDate: dtEndDate1 = dtStartDate
    i = 1
    msfResult.Rows = 2
    Call ClearSpecRowCol(msfResult, 1, Array())
    
    Select Case mbytQueryMode
    Case 20, 21
        If mbytQueryMode = 20 Then
            strSQL = " Where 数据性质=1 And ID=(Select 打印ID From " & zlGetFullFieldsTable("票据使用明细") & " Where 票种=1 And 性质=1 And 号码=[3] And Rownum=1)"
            strSQL = "Select NO From 票据打印内容 " & strSQL & " Union All Select NO From H票据打印内容" & strSQL
            strBill = " And A.NO In (" & strSQL & ") "
        Else
            strBill = " And A.No=[3] "
        End If
        
        gstrSQL = "" & _
        "SELECT A.收据费目 as 收据费目," & _
        "       Ltrim(To_Char(Sum(Nvl(A.实收金额,0)),'9999999990.00')) as 金额," & _
        "       Ltrim(To_Char(Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)),'9999999990.00')) as 未结帐金额" & _
        " FROM 门诊费用记录 A " & _
        " Where A.记录状态=1 and A.记录性质=1 and 门诊标志=1 " & strBill & _
        " Group by A.收据费目 " & _
        " Order by 收据费目"
    Case Else
         strWhere = ""
        
        If UsrCmd(28).Visible And UsrCmd(28).Tag = "1" And UsrCmd(29).Visible And UsrCmd(29).Tag = "1" Then
            
        ElseIf UsrCmd(28).Visible And UsrCmd(28).Tag = "1" Then
            '收费费用
            strWhere = strWhere & " And A.记帐费用<>1 "
        ElseIf UsrCmd(29).Visible And UsrCmd(29).Tag = "1" Then
            '记帐费用
            strWhere = strWhere & "And A.记帐费用=1 "
        End If
        
        If lng主页id <= 0 And mlngDays > 0 Then strWhere = strWhere & " And A." & mstr费用时间类型 & " between [4] and [5] "
        If lblRange.Visible Then strWhere = strWhere & " And A." & mstr费用时间类型 & ">=[6] and A." & mstr费用时间类型 & "<[7] "

                        
        '江磊改于7月    错误编号2363
        gstrSQL = "" & _
        " SELECT A.收据费目 as 收据费目," & _
        "       Sum(Nvl(A.实收金额,0)) as 金额," & _
        "       Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)) as 未结帐金额" & _
        " FROM 住院费用记录 A  " & _
        " Where A.记录状态<>0   And A.病人ID=[1] and nvl(A.主页id,0)=[2] " & strWhere & _
        " Group by A.收据费目 "
        If lng主页id = 0 Then
            gstrSQL = gstrSQL & " Union ALL " & _
            " SELECT A.收据费目 as 收据费目," & _
            "       Sum(Nvl(A.实收金额,0)) as 金额," & _
            "       Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)) as 未结帐金额" & _
            " FROM 门诊费用记录 A  " & _
            " Where A.记录状态<>0   And A.病人ID=[1] " & strWhere & _
            " Group by A.收据费目 "
        End If
        '77306,李南春,2014/9/1,未结账金额算成了结账金额
        gstrSQL = "" & _
        " SELECT A.收据费目," & _
        "       Ltrim(To_Char(Sum(Nvl(A.金额,0)),'9999999990.00')) as 金额," & _
        "       Ltrim(To_Char(Sum(Nvl(A.未结帐金额,0)),'9999999990.00')) as 未结帐金额" & _
        " FROM (" & gstrSQL & ") A  " & _
        " Group by A.收据费目 " & _
        " Order by 收据费目"
        
    End Select
    
    If IsDate(Format(mvarMinDate, "YYYY-MM-DD")) Then dtStartDate = CDate(Format(mvarMinDate, "YYYY-MM-DD") + " 00:00:00")
    If IsDate(Format(mvarMaxDate, "YYYY-MM-DD")) Then dtEndDate = CDate(Format(mvarMaxDate, "YYYY-MM-DD") + " 23:59:59")
    If lblRange.Visible Then
        dtStartDate1 = CDate(Format(lblRange.Caption, "YYYY-MM-DD HH:MM:SS")): dtEndDate = CDate(Format(CDate(lblRange.Caption) + 1, "YYYY-MM-DD HH:MM:SS"))
    End If
    
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "费用查询", lng病人ID, lng主页id, mstrNO, dtStartDate, dtEndDate, dtStartDate1, dtEndDate1)
    If gRs.RecordCount > 0 Then
        gRs.MoveFirst
        While Not gRs.EOF
            msfResult.TextMatrix(i, 0) = IIf(IsNull(gRs!收据费目), "", gRs!收据费目)
            msfResult.TextMatrix(i, 1) = IIf(IsNull(gRs!金额), "", gRs!金额)
            msfResult.TextMatrix(i, 3) = IIf(IsNull(gRs!未结帐金额), "", gRs!未结帐金额)
            msfResult.TextMatrix(i, 2) = Format(Val(msfResult.TextMatrix(i, 1)) - Val(msfResult.TextMatrix(i, 3)), "0.00")
            
            sglMoney(1) = sglMoney(1) + Val(msfResult.TextMatrix(i, 1))
            sglMoney(2) = sglMoney(2) + Val(msfResult.TextMatrix(i, 2))
            sglMoney(3) = sglMoney(3) + Val(msfResult.TextMatrix(i, 3))
            i = i + 1
            msfResult.Rows = i + 1
            gRs.MoveNext
        Wend
        
        '新加一小计行
        Call SetMsfRowColor(msfResult, i, &H80000018)
        msfResult.MergeRow(i) = True
        For j = 0 To msfResult.Cols - 4
            msfResult.TextMatrix(i, j) = "合计:"
        Next
        
        msfResult.TextMatrix(i, 1) = Format(sglMoney(1), "0.00")
        msfResult.TextMatrix(i, 2) = " " & Format(sglMoney(2), "0.00")
        msfResult.TextMatrix(i, 3) = Format(sglMoney(3), "0.00")
        sglMoney(1) = 0
        sglMoney(2) = 0
        sglMoney(3) = 0
        mvarRows = msfResult.Rows - 1
    End If
    msfResult.Rows = msfResult.Rows + 50
    
    Load分类费用 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub DrawMsfHeader(ByVal bytMode As Byte)
'功能:根据所选功能作出相应的表格
'参数:bytMode       所选取功能索引
    
    With msfResult
        .Rows = 50
        .Cols = 0
        ClearSpecRowCol msfResult, 1, Array()
        Select Case bytMode
        Case 0      '费用明细
            Call AddColumn(msfResult, "日期", 1500, 1)
            Call AddColumn(msfResult, "科室", 1200, 1)
            Call AddColumn(msfResult, "类别", 1200, 1)
            Call AddColumn(msfResult, "名称", 3600, 1)
            Call AddColumn(msfResult, "数量", 1200, 7)
            Call AddColumn(msfResult, "单价", 1200, 7)
            Call AddColumn(msfResult, "金额", 1200, 7)
            Call AddColumn(msfResult, "编码", 1200, 7)
            Call AddColumn(msfResult, "说明", 0, 1)
            Call AddColumn(msfResult, "", 1200, 7)
            Call CalcAutoColWidth(msfResult, 3)
        Case 1      '分类费用
            Call AddColumn(msfResult, "收费费目", 1500, 1)
            Call AddColumn(msfResult, "金额", 1800, 7)
            Call AddColumn(msfResult, "已结帐金额", 1800, 7)
            Call AddColumn(msfResult, "未结帐金额", 1800, 7)
            Call AddColumn(msfResult, "", 1200, 7)
            Call CalcAutoColWidth(msfResult, 4)
        Case 2
            Call AddColumn(msfResult, "日期", 1290, 1)
            Call AddColumn(msfResult, "科室", 1500, 1)
            Call AddColumn(msfResult, "结算方式", 990, 1)
            Call AddColumn(msfResult, "结算号码", 990, 1)
            Call AddColumn(msfResult, "预交金额", 1500, 7)
            Call AddColumn(msfResult, "", 1200, 7)
            Call CalcAutoColWidth(msfResult, 5)
        End Select
    End With
    
    msfResult.Cell(flexcpFontName, 0, 0, 0, msfResult.Cols - 1) = "楷体_GB2312"
    msfResult.Cell(flexcpFontSize, 0, 0, 0, msfResult.Cols - 1) = 14
    msfResult.Cell(flexcpFontBold, 0, 0, 0, msfResult.Cols - 1) = True
End Sub

Private Sub UsrCmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    mlngLoopCount = 0
End Sub

Private Function CheckQueryPass(ByVal lng病人ID As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim IntCount As Integer
    Dim strInputPassword As String
    
    On Error GoTo errH:
    
    strSQL = "Select 查询密码 from 病人信息 where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取信息", lng病人ID)
    If Not rsTemp.EOF Then
        If Trim(Nvl(rsTemp!查询密码)) <> "" Then
            For IntCount = 1 To 3
                With frmCheckQueryPass
                    .Show 1, Me
                    If .mblnOK = False Then
                        Exit Function
                    Else
                        strInputPassword = zlCommFun.zlStringEncode(LCase(.mstrPass))
                        
                        If LCase(Nvl(rsTemp!查询密码)) <> strInputPassword Then
                            If IntCount < 3 Then MsgBox "您输入的密码不对，请重新输入！", vbInformation, gstrSysName
                        Else
                            Exit For
                        End If
                    End If
                End With
            Next
            If IntCount > 3 Then '超过三次
                MsgBox "您输入的密码超过三次均错误，禁止查询！", vbInformation, gstrSysName: Exit Function
            End If
        End If
    End If
    
    CheckQueryPass = True
    Exit Function
    
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub EnterCharge()
'    If mvar病人id > 0 Then
        
        Select Case Val(GetPara("费用时间类型", "0"))
        Case 0
            mstr费用时间类型 = "发生时间"
        Case 1
            mstr费用时间类型 = "登记时间"
        Case Else
            mstr费用时间类型 = "发生时间"
        End Select
    
        picAsk(0).Visible = False
        Call LoadPatient(mvar病人id, mvar主页id, mvar姓名)
        
'        UsrCmd(2).Visible = IIf(mvar主页id > 0, True, False)
            
        UsrCmd(5).Picture = ilsImage.ListImages("unselect")
        UsrCmd(5).Tag = ""
        lblRange.Visible = False
        UsrCmd(6).Visible = False
        UsrCmd(7).Visible = False
        
        UsrCmd(28).Visible = Not (mbytQueryMode = 20 Or mbytQueryMode = 21)
        UsrCmd(29).Visible = Not (mbytQueryMode = 20 Or mbytQueryMode = 21)
        
        Set mrsDateList = GetDateList(mvar病人id, mvar主页id, mvarMinDate, mvarMaxDate, mstr费用时间类型)
        
        '正式进入查询费用
        Call UsrCmd_CommandClick(0)
        
        Call GetCurrentStateInfo(mvar病人id, mvar主页id, "")
        Call GetStateInfo(mvar病人id, mvar主页id)
'    End If

    UsrCmd(5).Enabled = Not (mbytQueryMode = 20 Or mbytQueryMode = 21)
    
End Sub

Private Sub AdjustEnabled()
    UsrCmd(6).Enabled = True
    UsrCmd(7).Enabled = True
    
    If Format(CDate(lblRange.Caption), "YYYY-MM-DD") >= mvarMaxDate Then UsrCmd(7).Enabled = False
    If Format(CDate(lblRange.Caption), "YYYY-MM-DD") <= mvarMinDate Then UsrCmd(6).Enabled = False
    
End Sub

Private Sub SetMsfRowColor(msf As Object, ByVal vRow As Long, ByVal Color As Long)
    Dim i As Long
    Dim svrRow As Long
    Dim svrCol As Long
    Dim blnRedraw As Boolean
    
    blnRedraw = msf.Redraw
    msf.Redraw = False
    svrRow = msf.Row
    svrCol = msf.Col
    msf.Row = vRow
    For i = 0 To msf.Cols - 1
        msf.Col = i
        msf.CellBackColor = Color
    Next
    msf.Row = svrRow
    msf.Col = svrCol
    msf.Redraw = True
    msf.Redraw = blnRedraw
End Sub

Public Sub FirstChar(ByVal ch As String)
    
    On Error Resume Next
    
    txt(4).Text = ch
    EnterFocus txt(4)
    zlCommFun.PressKey vbKeyEnd
    
End Sub

Public Property Let Enabled(ByVal vData As Boolean)
    UserControl.Enabled = vData
End Property

Private Sub usrKeyBoard_CommandClick(Caption As String)
    Dim lngPatientKey As Long
    Dim strPassWord As String
    
    Select Case Caption
    Case "确定"
        
        '按病人住院号查询病人费用
        mvar病人id = 0
        mvar主页id = 0
        mvar姓名 = ""
        
        Call mclsIDKind.LeaveInputBox
        
        Select Case mbytQueryMode
        '--------------------------------------------------------------------------------------------------------------
        Case 12 '住院号
            If IsNumeric(txt(4).Text) Then
                gstrSQL = "select rownum as No,A.入院日期,A.出院日期,A.病人id,A.主页id from 病案主页 A,病人信息 B where A.病人id=B.病人id and B.住院号=[1] order by A.入院日期 desc"
                Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "费用查询", Trim(txt(4).Text))
                If gRs.RecordCount > 0 Then
                    If Not CheckQueryPass(Nvl(gRs!病人id, 0)) Then Exit Sub '密码校验
                    If frmSelect.ShowSelect(gRs, mvar病人id, mvar主页id) Then
                        Call EnterCharge
                    End If
                Else
                    MsgBox "输入了一个不存在的住院号，请重新输入！", vbInformation, gstrSysName
                End If
            Else
                MsgBox "请输入一个正确的住院号！", vbInformation, gstrSysName
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 3      '按病人门诊号查询病人费用
            If IsNumeric(txt(4).Text) Then
                Set gRs = zlDatabase.OpenSQLRecord("select B.病人id from 病人信息 B where B.门诊号=[1]", "费用查询", Trim(txt(4).Text))
                If gRs.RecordCount > 0 Then
                    If Not CheckQueryPass(Nvl(gRs!病人id, 0)) Then Exit Sub '密码校验
                    mvar病人id = IIf(IsNull(gRs!病人id), 0, gRs!病人id)
                    Call EnterCharge
                Else
                    MsgBox "输入了一个不存在的门诊号，请重新输入！", vbInformation, gstrSysName
                End If
            Else
                MsgBox "请输入一个正确的门诊号！", vbInformation, gstrSysName
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 14         '按病人ID号查询病人费用
            If IsNumeric(txt(4).Text) Then
                
                gstrSQL = "" & _
                "   Select to_char(A.入院日期,'yyyy-mm-dd') as 入院日期,to_char(A.出院日期,'yyyy-mm-dd') as 出院日期,A.病人id,A.主页id " & _
                "   From 病案主页 A " & _
                "   Where A.病人id=[1] " & _
                "   Union ALL " & _
                "   Select '门诊费用' as 入院日期,'门诊费用' as 出院日期,病人id, 0 as 主页id " & _
                "   From 门诊费用记录 C " & _
                "   Where 病人id=[1] and rownum<2   "
                
                gstrSQL = "Select 入院日期,出院日期,病人id,主页id From (" & gstrSQL & ") AA  order by AA.入院日期 asc,AA.主页id desc"
                gstrSQL = "SELECT RowNum as No,D.入院日期,D.出院日期,D.病人id,D.主页id FROM (" & gstrSQL & ") D"
                
                Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "费用查询", Val(txt(4).Text))
                If gRs.RecordCount > 0 Then
                    If Not CheckQueryPass(Nvl(gRs!病人id, 0)) Then Exit Sub '密码校验
                    If frmSelect.ShowSelect(gRs, mvar病人id, mvar主页id) Then
                        Call EnterCharge
                    End If
                   
                Else
                    MsgBox "输入了一个不存在的病人ID号，请重新输入！", vbInformation, gstrSysName
                End If
            Else
                MsgBox "请输入一个正确的病人ID号！", vbInformation, gstrSysName
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 11, 30, 31, 32, 33, 34, 35         '按卡号查
            
            gstrSQL = "" & _
            "   Select to_char(A.入院日期,'yyyy-mm-dd') as 入院日期,to_char(A.出院日期,'yyyy-mm-dd') as 出院日期,A.病人id,A.主页id " & _
            "   From 病案主页 A " & _
            "   Where A.病人id=[1] " & _
            "   Union ALL " & _
            "   Select '门诊费用' as 入院日期,'门诊费用' as 出院日期,病人id, 0 as 主页id " & _
            "   From 门诊费用记录 C " & _
            "   Where 病人id=[1] and rownum<2   "
            
            gstrSQL = "Select 入院日期,出院日期,病人id,主页id From (" & gstrSQL & ") AA  order by AA.入院日期 asc,AA.主页id desc"
            gstrSQL = "SELECT RowNum as No,D.入院日期,D.出院日期,D.病人id,D.主页id FROM (" & gstrSQL & ") D"
            
            If mclsIDKind.zlGetPatiIDByCardNo(txt(4).Text, lngPatientKey, strPassWord) Then

                Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "费用查询", lngPatientKey)
                If gRs.RecordCount > 0 Then
                
                    If Val(GetPara("就诊卡密码验证", "1")) = 1 And strPassWord <> "" Then
                        If frmCardPass.ShowCardPass(strPassWord) = False Then
                            Exit Sub
                        End If
                    Else
                        If Not CheckQueryPass(lngPatientKey) Then Exit Sub
                    End If
                    If frmSelect.ShowSelect(gRs, mvar病人id, mvar主页id) Then
                        Call EnterCharge
                    End If
                Else
                    MsgBox "您查找的病人不存在，请重新输入！", vbInformation, gstrSysName
                End If
            Else
                MsgBox "您查找的病人不存在，请重新输入！", vbInformation, gstrSysName
            End If
            
        '--------------------------------------------------------------------------------------------------------------
        Case 20         '按票据号
            
            mstrNO = UCase(Trim(txt(4).Text))
            If mstrNO <> "" Then
                
                gstrSQL = " Where 数据性质=1 And ID=(Select 打印ID From " & zlGetFullFieldsTable("票据使用明细") & " Where 票种=1 And 性质=1 And 号码=[1] And Rownum=1)"
                gstrSQL = "Select NO From 票据打印内容 " & gstrSQL & " Union All Select NO From H票据打印内容 " & gstrSQL
                gstrSQL = "Select 病人id, 0 as 主页id,姓名 From 门诊费用记录 Where 记录状态=1 and 门诊标志=1 and 记录性质=1 and No In (" & gstrSQL & ")"
                
                Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "费用查询", mstrNO)
                If gRs.BOF = False Then
                    mvar病人id = zlCommFun.Nvl(gRs("病人id"), 0)
                    mvar主页id = zlCommFun.Nvl(gRs("主页id"), 0)
                    mvar姓名 = zlCommFun.Nvl(gRs("姓名"), "")
                    If CheckQueryPass(mvar病人id) Then '密码校验
                        Call EnterCharge
                    End If
                Else
                    MsgBox "输入了一个不存在的门诊票据号，请重新输入！", vbInformation, gstrSysName
                    EnterFocus txt(4)
                End If
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 21         '按单据号
            mstrNO = UCase(Trim(txt(4).Text))
            If mstrNO <> "" Then
                gstrSQL = "Select 病人id,0 as 主页id,姓名 From 门诊费用记录 Where 记录状态=1 and 记录性质=1 and 门诊标志=1 and No=[1]"
                Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "费用查询", mstrNO)
                If gRs.BOF = False Then
                    mvar病人id = zlCommFun.Nvl(gRs("病人id"), 0)
                    mvar主页id = zlCommFun.Nvl(gRs("主页id"), 0)
                    mvar姓名 = zlCommFun.Nvl(gRs("姓名"), "")
                    If CheckQueryPass(mvar病人id) Then '密码校验
                        Call EnterCharge
                    End If
                Else
                    MsgBox "输入了一个不存在的门诊单据号，请重新输入！", vbInformation, gstrSysName
                    EnterFocus txt(4)
                End If
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 25             '身份证号
            gstrSQL = "" & _
            "   Select to_char(A.入院日期,'yyyy-mm-dd') as 入院日期,to_char(A.出院日期,'yyyy-mm-dd') as 出院日期,A.病人id,A.主页id,B.卡验证码" & _
            "   From 病案主页 A,病人信息 B " & _
            "   where A.病人id=B.病人id and B.身份证号=[1] " & _
            "   Union ALL " & _
            "   Select '门诊费用' as 入院日期,'门诊费用' as 出院日期,E.病人id,0 as 主页id,E.卡验证码 " & _
            "   From 门诊费用记录 C,病人信息 E " & _
            "   Where C.病人id=E.病人id And E.身份证号=[1] and Rownum<2 "
            
            gstrSQL = " Select AA.入院日期,AA.出院日期,AA.病人id,AA.主页id,AA.卡验证码 From (" & gstrSQL & ") AA  order by AA.入院日期 asc,AA.主页id desc "
            gstrSQL = "" & _
            "   SELECT rownum as No,D.入院日期,D.出院日期,D.病人id,D.主页id,D.卡验证码 " & _
            "   FROM ( " & gstrSQL & ") D"
'
            Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "费用查询", UCase(txt(4).Text))
            If gRs.RecordCount > 0 Then
                txt(4).Text = ""
                If Not CheckQueryPass(Nvl(gRs!病人id, 0)) Then Exit Sub '密码校验
                If frmSelect.ShowSelect(gRs, mvar病人id, mvar主页id) Then
                    Call EnterCharge
                End If
            Else
                MsgBox "不存的身份证号，请按[确定]后重新刷卡！", vbInformation, gstrSysName
                txt(4).Text = ""
                EnterFocus txt(4)
            End If
            
        '--------------------------------------------------------------------------------------------------------------
        Case 26             'ＩＣ卡号
            
            gstrSQL = "" & _
            "   Select to_char(A.入院日期,'yyyy-mm-dd') as 入院日期,to_char(A.出院日期,'yyyy-mm-dd') as 出院日期,A.病人id,A.主页id,B.卡验证码 " & _
            "   From 病案主页 A,病人信息 B " & _
            "   Where A.病人id=B.病人id  and B.IC卡号=[1]" & _
            "   Union ALL " & _
            "   Select  '门诊费用' as 入院日期,'门诊费用' as 出院日期,E.病人id,0 as 主页id,E.卡验证码 " & _
            "   From 门诊费用记录 C,病人信息 E" & _
            "   Where  C.病人id=E.病人id And E.IC卡号=[1] and rownum<2   "
            
            gstrSQL = "Select 入院日期,出院日期,病人id,主页id,卡验证码 From (" & gstrSQL & ") AA  order by AA.入院日期 asc,AA.主页id desc"
            gstrSQL = "SELECT RowNum as No,D.入院日期,D.出院日期,D.病人id,D.主页id,D.卡验证码 FROM (" & gstrSQL & ") D"
            Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "费用查询", UCase(txt(4).Text))
            If gRs.RecordCount > 0 Then
                txt(4).Text = ""
                strPassWord = zlCommFun.Nvl(gRs("卡验证码").Value)
                
                If Val(GetPara("就诊卡密码验证", "1")) = 1 And strPassWord <> "" Then
                    If frmCardPass.ShowCardPass(strPassWord) = False Then
                        Exit Sub
                    End If
                Else
                    If Not CheckQueryPass(Nvl(gRs!病人id, 0)) Then Exit Sub
                End If

                If frmSelect.ShowSelect(gRs, mvar病人id, mvar主页id) Then
                    Call EnterCharge
                End If
            Else
                MsgBox "不存的IC卡号，请按[确定]后重新刷卡！", vbInformation, gstrSysName
                txt(4).Text = ""
                EnterFocus txt(4)
            End If

        End Select
        
        txt(4).Text = ""
        
    Case "清除"
        
        txt(4).Text = ""
            
    Case Else
        
        txt(4).Text = txt(4).Text & Caption
            
        EnterFocus txt(4)
        zlCommFun.PressKey vbKeyEnd
    End Select
End Sub

Private Sub usrKeyBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    mlngLoopCount = 0
End Sub

Public Sub SelectRow(objVsf As Object, ByVal OldRow As Long, ByVal NewRow As Long, Optional ByVal lngBackColor As Long = -1)
    '--------------------------------------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------------------------------------
    Dim lngColor As Long
    
    On Error Resume Next
    
    If OldRow = NewRow Then Exit Sub
    
    If lngBackColor = -1 Then
        lngColor = objVsf.BackColorSel
    Else
        lngColor = lngBackColor
    End If
    
    If OldRow + 1 > objVsf.FixedRows Then
        objVsf.Cell(flexcpBackColor, OldRow, objVsf.FixedCols, OldRow, objVsf.Cols - 1) = objVsf.BackColor
    End If
    
    If NewRow + 1 > objVsf.FixedRows Then
        objVsf.Cell(flexcpBackColor, NewRow, objVsf.FixedCols, NewRow, objVsf.Cols - 1) = lngColor
    End If
    
End Sub

Private Sub 欠费情况(str姓名 As String, lng病人ID As Long, lng主页id As Long, ByVal sng剩余金额 As Single, ByVal str报警方案 As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    Dim strError As String
    Dim sng报警值 As Single
    Dim sng担保额 As Single
    
    '先恢复 lblWarn 的初始状态
    lblWarn.BackStyle = 0
    lblWarn.Caption = ""

    gstrSQL = "Select 报警方法,报警值 From 记帐报警线 A,病案主页 B Where A.适用病人=[3] And A.病区ID = B.当前病区ID And B.病人id =[1] And B.主页id = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng病人ID, lng主页id, str报警方案)
    If rsTmp.BOF Then Exit Sub
    
    sng报警值 = IIf(IsNull(rsTmp!报警值), 0, rsTmp!报警值)
    
    gstrSQL = "Select 担保额 From 病人信息 A Where A.病人ID =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng病人ID)
    If Not rsTmp.BOF Then sng担保额 = zlCommFun.Nvl(rsTmp!担保额, 0)
    
    sng剩余金额 = sng剩余金额 + sng担保额
    
    If sng剩余金额 < 0 Then
        strError = str姓名 & "已欠费" & Abs(FormatEx(sng剩余金额, 2)) & "元，请尽快缴费！"
        lblWarn.BackStyle = 1
        lblWarn.BackColor = &HFFFF&
        lblWarn.Caption = "欠费"
        
    ElseIf sng剩余金额 < sng报警值 Then
        strError = str姓名 & "剩余款总额(" & FormatEx(sng剩余金额, 2) & ")以低于报警值(" & FormatEx(sng报警值, 2) & ")！"
        lblWarn.BackStyle = 1
        lblWarn.BackColor = &HFFFF&
        lblWarn.Caption = "报警"
    End If

    
    Call frmShowMessage.ShowMe(Me, strError)
    
    
End Sub

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
    '******************************************************************************************************************
    '功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
    '参数：vNumber=Single,Double,Currency类型的数字,intBit=最大小数位数
    '******************************************************************************************************************
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = 0
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
        End If
    End If
    FormatEx = strNumber
End Function
