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
            Name            =   "����"
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
         Caption         =   "������ô�����"
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
         Caption         =   "�����￨��"
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
         Caption         =   "��סԺ�Ų�"
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
         Caption         =   "������Ų�"
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
         Caption         =   "������ID��"
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
         Caption         =   "��ҽ������"
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
         Caption         =   "��Ʊ�ݺŲ�"
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
         Caption         =   "�����ݺŲ�"
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
         Caption         =   "�����֤��"
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
         Caption         =   "���ɣÿ���"
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
         Caption         =   " ��  ��  "
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
         Caption         =   "�����￨��"
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
         Caption         =   "�����￨��"
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
         Caption         =   "�����￨��"
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
         Caption         =   "�����￨��"
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
         Caption         =   "�����￨��"
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
         Caption         =   "�����￨��"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   450
         TextAligment    =   0
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "����סԺ��:"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "���������½ǵ����ְ�ť�����ΰ�����סԺ�ţ�Ȼ��""ȷ��""��ť"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��ѡ�����ķ��ò�ѯ��ʽ��ֻ���ڲ�ѯ��ʽ�ϰ�һ�£����ϡ��̡��ͱ�ʾѡȡ�еĲ�ѯ��ʽ��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "ѡ����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��ѡ����"
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
         Caption         =   "������ôʹ�ã�"
         BackColor       =   16777215
         FontSize        =   10.5
         ButtonHeight    =   390
         TextAligment    =   0
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:        �Ա�:    ��Ժʱ��:          ��Ժʱ��:"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����"
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
         Caption         =   "��һ��"
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
         Caption         =   "��һ��"
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
         Caption         =   "ָ���·�"
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
         Caption         =   "����"
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
         Caption         =   "����"
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
         Caption         =   "�˷ѷ���"
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
         Caption         =   "�շѷ���"
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
         Caption         =   "���ʷ���"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "��ϸ��ѯ"
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
         Caption         =   "�����ѯ"
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
         Caption         =   "�Ϸ�"
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
         Caption         =   "�·�"
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
         Caption         =   "�ض�����"
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
         Caption         =   "Ԥ�ɿ��ѯ"
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "Ԥ�ɷ��ã�0.00             ���÷��ã�0.00              ʣ����ã�0.00"
         BeginProperty Font 
            Name            =   "����"
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
         Name            =   "����"
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
         Caption         =   "�Ϸ�"
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
         Caption         =   "�·�"
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
         Caption         =   "��Ŀ˵��:"
         BeginProperty Font 
            Name            =   "����"
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
Private mvar����id As Long
Private mvar��ҳid As Long
Private mvar���� As Boolean
Private mvar���� As String
Private mvarStop As Long                '�û���ѯ��Ϣͣ�����
Private mvarScroll As Long
Private mlngDays As Long
Private mlngLoopCount As Long
Private mbytQueryMode As Byte
Private mstrNO As String
Private mstr����ʱ������ As String
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
        
        lblDescrible.Caption = "���������½ǵ����ְ�ť�����ΰ�����ID�ţ�Ȼ�󰴡�ȷ������ť"
        lblCaption.Caption = "����ID��:"
        
        usrKeyBoard.KeyMode = 2
        DoEvents
        usrKeyBoard.Visible = True
        UsrCmd(27).Visible = False
        UsrPic.Visible = False
        EnterFocus txt(4)
    '------------------------------------------------------------------------------------------------------------------
    Case 3
        
        lblDescrible.Caption = "���������½ǵ����ְ�ť�����ΰ���������ţ�Ȼ�󰴡�ȷ������ť"
        lblCaption.Caption = "���������:"
        usrKeyBoard.KeyMode = 2
        DoEvents
        usrKeyBoard.Visible = True
        UsrPic.Visible = False
        UsrCmd(27).Visible = False
        EnterFocus txt(4)
    '------------------------------------------------------------------------------------------------------------------
    Case 11, 30, 31, 32, 33, 34, 35
        
        lblDescrible.Caption = "����ֱ����ˢ������ˢ���ľ��￨"
        lblCaption.Caption = "��ˢ��:"
        usrKeyBoard.KeyMode = 2
        DoEvents
        usrKeyBoard.Visible = False
        UsrCmd(27).Visible = False
        UsrPic.Visible = True
        EnterFocus txt(4)
    '------------------------------------------------------------------------------------------------------------------
    Case 26
        
        lblDescrible.Caption = "��������ɣÿ���Ȼ�󰴡���������ť"
        lblCaption.Caption = "�ɣÿ���:"
        
        usrKeyBoard.KeyMode = 2
        usrKeyBoard.Visible = False
         
        UsrPic.Visible = False
        
        UsrCmd(27).Visible = True
        
        EnterFocus txt(4)
    '------------------------------------------------------------------------------------------------------------------
    Case 12, 20, 21, 25
        
        
        If bytIndex = 12 Then
            
            lblDescrible.Caption = "���������½ǵ����ְ�ť�����ΰ�����סԺ�ţ�Ȼ�󰴡�ȷ������ť"
            lblCaption.Caption = "����סԺ��:"
            usrKeyBoard.KeyMode = 2
            
        ElseIf bytIndex = 20 Then
            lblDescrible.Caption = "���������½ǵ����ְ�ť�����ΰ�����Ʊ�ݺţ�Ȼ�󰴡�ȷ������ť"
            lblCaption.Caption = "����Ʊ�ݺ�:"
            usrKeyBoard.KeyMode = 1
        ElseIf bytIndex = 21 Then
            lblDescrible.Caption = "���������½ǵ����ְ�ť�����ΰ����ĵ��ݺţ�Ȼ�󰴡�ȷ������ť"
            lblCaption.Caption = "���ĵ��ݺ�:"
            usrKeyBoard.KeyMode = 1
        ElseIf bytIndex = 25 Then
            lblDescrible.Caption = "���������½ǵ����ְ�ť�����ΰ��������֤�ţ�Ȼ�󰴡�ȷ������ť"
            lblCaption.Caption = "�������֤:"
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
    '��ʼ������
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
    
    
    strTmp = Trim(zlDatabase.GetPara("��ѯ���÷�ʽ", glngSys, 1536, "100000000"))
    '1-���￨;2-�����;3-סԺ��;4-����ID;5-ҽ����;6-���֤;7-Ʊ�ݺ�;8-���ݺ�;9-�ɣÿ�
    Dim strIDKind As String
    
'    strIDKind = "��|���֤��|0|0|18"
    
    Set mclsIDKind = New clsIDKind
    Call mclsIDKind.InitIDKind(1, gcnOracle, glngSys, UserInfo.����, UserControl.Extender, 1536, strIDKind)
        
    If Val(Mid(strTmp, 1, 1)) > 0 And mclsIDKind.GetCard(rs) Then
        If rs.RecordCount > 0 Then
            
            IntCount = 0
            Do While Not rs.EOF
                
                IntCount = IntCount + 1
                
                If IntCount < 7 Then
                    Select Case IntCount
                    Case 1
                         UsrCmd(11).Visible = True
                         UsrCmd(11).Caption = rs("ȫ��").Value
                    Case Else
                         UsrCmd(28 + IntCount).Visible = True
                         UsrCmd(28 + IntCount).Caption = rs("ȫ��").Value
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
    mvarStop = Val(GetPara("���ò�ѯͣ��ʱ��", "30"))
    mvarStop = IIf(mvarStop <= 0, 30, mvarStop)
    
    mvarScroll = Val(GetPara("���ò�ѯ�������", "10"))
    mvarScroll = IIf(mvarScroll <= 0, 10, mvarScroll)
    
    gstrSQL = "select A.��ͼ��� from ��ѯ����Ŀ¼ A where A.ҳ�����=2 and A.�������=1"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "���ò�ѯ")
    If gRs.BOF = False Then
        UsrPic.Tag = GetFileName(IIf(IsNull(gRs!��ͼ���), 0, gRs!��ͼ���), UsrPic.Width, UsrPic.Height)
        Call UsrPic.ShowPictureByFile(UsrPic.Tag)
    End If
    
    Call AdjustCmdPostion
    
    tmrLoop.Tag = Val(GetPara("���ط��ü��", "0"))
    
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
        Call usrKeyBoard_CommandClick("ȷ��")
    End If
End Sub

Private Sub mclsIDKind_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txt(4).Text = "" And Not txt(4).Locked And UserControl.ActiveControl Is txt(4) Then
        txt(4).Text = strID
        Call usrKeyBoard_CommandClick("ȷ��")
    End If
End Sub

Private Sub mfrmCardPass_AfterInputPassword(ByVal strInputPassword As String, blnSucc As Boolean)
    
    
End Sub

Private Sub msfResult_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    Call SelectRow(msfResult, OldRow, NewRow)
    
End Sub

Private Sub msfResult_Click()
    tmrScroll.Enabled = False
    mvarStop = Val(GetPara("���ò�ѯͣ��ʱ��", "30"))
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
        lbl(4).Caption = "��Ŀ˵��:"
        
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
        '���ط��õ�¼����
        
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
        mvarScroll = Val(GetPara("���ò�ѯ�������", "10"))
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
        Call mclsIDKind.EnterInputBox("���֤��")
    End If
    
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case mbytQueryMode
    Case 11, 30, 31, 32, 33, 34, 35
        Call mclsIDKind.InputKeyPress(KeyAscii, txt(4))
    End Select
    
    If CheckIsInclude(UCase(Chr(KeyAscii)), "'����;��:��?��|,��.��""") = True Then KeyAscii = 0
End Sub

Private Sub txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 4 And mbytQueryMode = 11 And Len(txt(Index).Text) > 0 Then
        
        If CheckIsInclude(UCase(Chr(KeyCode)), "'") = True Then KeyCode = 0
        
        If mclsIDKind.IsCard Then
            If Len(txt(Index).Text) = mclsIDKind.CardLength And KeyCode <> 0 Or KeyCode = 13 Then
                Call usrKeyBoard_CommandClick("ȷ��")
            End If
        Else
            If KeyCode = 13 Then
                Call usrKeyBoard_CommandClick("ȷ��")
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
    
    '���Ƿ����ʾ
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
    mvarStop = Val(GetPara("���ò�ѯͣ��ʱ��", "30"))
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
            Call Load������ϸ(mvar����id, mvar��ҳid)
            
            Call msfResult_RowColChange
            
        Case 1
            If picTitle(3).Visible = False Then
                picTitle(3).Visible = True
                msfResult.Height = msfResult.Height - picTitle(3).Height
            End If
        
            Call Load�������(mvar����id, mvar��ҳid)
            
        End Select
        Call EnablePageButton(msfResult, mvarCurPos, mvarRows, UsrCmd(9), UsrCmd(10))
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        'ˢ�²���Ԥ����¼
        mvarCurPosPre = 1
        mvarRowsPre = 0
    '------------------------------------------------------------------------------------------------------------------
    Case 4                      '��ѡ����
        
        txt(4).Text = ""

        UsrCmd(9).Enabled = False
        UsrCmd(10).Enabled = False
        picAsk(0).Visible = True
        
        Call EnterFocus(txt(4))
    '------------------------------------------------------------------------------------------------------------------
    Case 5                      'ָ������
    
        UsrCmd(Index).Tag = IIf(UsrCmd(Index).Tag = "", "1", "")
        UsrCmd(Index).Picture = ilsImage.ListImages(IIf(UsrCmd(Index).Tag = "1", "select", "unselect"))
        
        UsrCmd(6).Visible = IIf(UsrCmd(Index).Tag = "", False, True) And UsrCmd(Index).Visible
        UsrCmd(7).Visible = IIf(UsrCmd(Index).Tag = "", False, True) And UsrCmd(Index).Visible
        
        mrsDateList.Filter = ""
        mrsDateList.Filter = "����<='" & lblRange.Caption & "'"
        If mrsDateList.RecordCount > 0 Then
            mrsDateList.Sort = "���� Desc"
            mrsDateList.MoveFirst
            lblRange.Caption = Format(mrsDateList("����").Value, "yyyy-MM-dd")
        End If
                
        lblRange.Visible = IIf(UsrCmd(Index).Tag = "", False, True) And UsrCmd(Index).Visible
        
        For i = 0 To 1
            If UsrCmd(i).State = -1 Then Call UsrCmd_CommandClick(i)
        Next
        
        Call GetCurrentStateInfo(mvar����id, mvar��ҳid, IIf(lblRange.Visible, lblRange.Caption, ""))
        
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
    Case 6              '��һ��
        mvarCurPos = 1
        mvarRows = 0
        
        mrsDateList.Filter = ""
        mrsDateList.Filter = "����<'" & lblRange.Caption & "'"
        If mrsDateList.RecordCount > 0 Then
            mrsDateList.Sort = "���� Desc"
            mrsDateList.MoveFirst
            lblRange.Caption = Format(mrsDateList("����").Value, "yyyy-MM-dd")
        End If
        
'        lblRange.Caption = Format(CDate(lblRange.Caption) - 1, "YYYY-MM-DD")
        For i = 0 To 1
            If UsrCmd(i).State = -1 Then Call UsrCmd_CommandClick(i)
        Next
        
        Call GetCurrentStateInfo(mvar����id, mvar��ҳid, lblRange.Caption)
        
        Call AdjustEnabled
    '------------------------------------------------------------------------------------------------------------------
    Case 7              '��һ��
        mvarCurPos = 1
        mvarRows = 0
        
        mrsDateList.Filter = ""
        mrsDateList.Filter = "����>'" & lblRange.Caption & "'"
        If mrsDateList.RecordCount > 0 Then
            mrsDateList.Sort = "����"
            mrsDateList.MoveFirst
            lblRange.Caption = Format(mrsDateList("����").Value, "yyyy-MM-dd")
        End If
        
'        lblRange.Caption = Format(CDate(lblRange.Caption) + 1, "YYYY-MM-DD")
        For i = 0 To 1
            If UsrCmd(i).State = -1 Then Call UsrCmd_CommandClick(i)
        Next
        
        Call GetCurrentStateInfo(mvar����id, mvar��ҳid, lblRange.Caption)
        
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
    Case 10             '��һҳ
        Call TurnToPage(msfResult, 1, mvarCurPos)
        Call EnablePageButton(msfResult, mvarCurPos, mvarRows, UsrCmd(9), UsrCmd(10))
    '------------------------------------------------------------------------------------------------------------------
    Case 3, 11, 12, 14, 15, 20, 21, 25, 26, 30, 31, 32, 33, 34, 35
        '11-�����￨��ѯ;12-��סԺ�Ų�ѯ;3-������Ų�ѯ;14-������ID�Ų�ѯ;15-��ҽ������ѯ;20-��Ʊ�ݺŲ�;21-�����ݺŲ�;25-�����֤��,26-���ɣÿ���
        '11,30,31,32,33,34,35Ϊ���￨�����п��ȣ�����ZLHIS���صĿ���
        
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
            If CheckIdentify(mvar����id, mvar��ҳid) Then
                If mvar����id <= 0 Then Exit Sub
                Call EnterCharge
            End If
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 27                     '����
        
        txt(4).Text = mclsIDKind.ReadICCard
        If txt(4).Text <> "" Then
            Call usrKeyBoard_CommandClick("ȷ��")
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


Private Sub LoadPatient(ByVal ����id As Long, ByVal ��ҳid As Long, Optional ByVal str���� As String = "")
    '����:���ݲ��˵�סԺ�Ż�û�����Ϣ
    '����:strסԺ��         ���˵�סԺ��
    Dim i As Long
    
    On Error GoTo errH
    If ��ҳid > 0 Then
        gstrSQL = "select A.����,A.�Ա�,B.��Ժ����,B.��Ժ���� from ������Ϣ A,������ҳ B where A.����id=B.����id and A.����id=[1] and B.��ҳid=[2]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "���ò�ѯ", ����id, ��ҳid)
        If gRs.RecordCount > 0 Then
            gRs.MoveFirst
            mvar���� = Nvl(gRs!����)
            lblMsg.Caption = IIf(IsNull(gRs!����), "", "����:" & gRs!����) & IIf(IsNull(gRs!�Ա�), "", "   �Ա�:" & gRs!�Ա�) & IIf(IsNull(gRs!��Ժ����), "", "   ��Ժ����:" & gRs!��Ժ����)
            lblRange.Caption = Format(IIf(IsNull(gRs!��Ժ����), zlDatabase.Currentdate, gRs!��Ժ����), "YYYY-MM-DD")
            lblRange.Tag = IIf(IsNull(gRs!��Ժ����), "0", "1")
            Call AdjustEnabled
        End If
    ElseIf ����id > 0 Then
        gstrSQL = "select A.����,A.�Ա� from ������Ϣ A where A.����id=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "���ò�ѯ", ����id, ��ҳid)
        If gRs.RecordCount > 0 Then
            gRs.MoveFirst
            mvar���� = Nvl(gRs!����)
            lblMsg.Caption = IIf(IsNull(gRs!����), "", "����:" & gRs!����) & IIf(IsNull(gRs!�Ա�), "", "   �Ա�:" & gRs!�Ա�) & "(���ﲡ��)"
            lblRange.Caption = Format(zlDatabase.Currentdate, "YYYY-MM-DD")
            Call AdjustEnabled
        End If
    Else
        lblMsg.Caption = "����:" & str����
    End If
    
    mlngDays = Val(GetPara("�������ǰ���������", "0"))
    
    
    '11-�����￨��ѯ;12-��סԺ�Ų�ѯ;3-������Ų�ѯ;14-������ID�Ų�ѯ;15-��ҽ������ѯ;20-��Ʊ�ݺŲ�;21-�����ݺŲ�;25-�����֤��,26-���ɣÿ���
    If ��ҳid > 0 Then
        gstrSQL = "select max(" & mstr����ʱ������ & ") as MaxDate,min(" & mstr����ʱ������ & ") as MinDate from סԺ���ü�¼ where ����id=[1] and ��ҳid=[2] "
    Else
        gstrSQL = " Select max(" & mstr����ʱ������ & ") as MaxDate,min(" & mstr����ʱ������ & ") as MinDate from ������ü�¼ where ����id=[1] UNION ALL  " & _
                  " Select max(" & mstr����ʱ������ & ") as MaxDate,min(" & mstr����ʱ������ & ") as MinDate from סԺ���ü�¼ where ����id=[1]    AND (��ҳid IS NULL OR ��ҳid=0) and nvl(�����־,0)<>2 "
        gstrSQL = "Select Max(MaxDate) as MaxDate ,min(MinDate) as MinDate From (" & gstrSQL & ")"
    End If
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "���ò�ѯ", ����id, ��ҳid)
    
    '���˺�:��SQL��ȫ��ͬ,����ڴ����ʱ,������������,ͬʱ�������SQL������
'    Select Case mbytQueryMode
'    Case 20
'        gstrSQL = "select max(" & mstr����ʱ������ & ") as MaxDate,min(" & mstr����ʱ������ & ") as MinDate from ���˷��ü�¼ where ����id=[1]" & IIf(��ҳid > 0, " and ��ҳid=[2]", " AND (��ҳid IS NULL OR ��ҳid=0)")
'    Case 21
'        gstrSQL = "select max(" & mstr����ʱ������ & ") as MaxDate,min(" & mstr����ʱ������ & ") as MinDate from ���˷��ü�¼ where ����id=[1]" & IIf(��ҳid > 0, " and ��ҳid=[2]", " AND (��ҳid IS NULL OR ��ҳid=0)")
'    Case Else
'        gstrSQL = "select max(" & mstr����ʱ������ & ") as MaxDate,min(" & mstr����ʱ������ & ") as MinDate from ���˷��ü�¼ where ����id=[1]" & IIf(��ҳid > 0, " and ��ҳid=[2]", " AND (��ҳid IS NULL OR ��ҳid=0)")
'    End Select
'
    If gRs.RecordCount > 0 Then
        gRs.MoveFirst
        mvarMaxDate = Format(IIf(IsNull(gRs!MaxDate), zlDatabase.Currentdate, gRs!MaxDate), "YYYY-MM-DD")
        mvarMinDate = Format(IIf(IsNull(gRs!MinDate), zlDatabase.Currentdate, gRs!MinDate), "YYYY-MM-DD")
        If ��ҳid <= 0 And mlngDays > 0 Then
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



Private Function GetStateInfo(ByVal lng����ID As Long, ByVal lng��ҳid As Long) As Boolean
    
    '******************************************************************************************************************
    '���ܣ���ȡ���˷��øſ�����ʾ
    '--���˷��øſ������ܷ�ӳ���˵ķ��ø�Ҫ�����
    '--  Ԥ�����ܶ�: �Ӳ���Ԥ����¼�л���
    '--  ��Ԥ���ܶ�: �Ӳ��˽��ʼ�¼�л���
    '--  �����ܶ�:   �Ӳ��˷��ü�¼�л���
    '--  �����ܶ�:   �Ӳ��˽��ʽ���л���
    '--���������ݣ�����ֱ�Ӽ��㣺
    '--  ���ʲ��˿�=�����ܶ�-��Ԥ���ܶ�
    '--  Ԥ�������=Ԥ�����ܶ�-��Ԥ���ܶ�
    '--  δ�����ܶ�=�����ܶ�-�����ܶ�
    '--  ʣ����ܶ�=Ԥ�������-δ�����ܶ�
    '******************************************************************************************************************
    
    On Error GoTo errH
    Dim bytDec As Byte
    Dim strDec As String
    Dim strSQL As String
    
    Dim lng���� As Long
    Dim lng�������� As Long
    Dim lng�ѱ������� As Long
    Dim sngʣ���� As Single
    Dim str�������� As String
    Dim rsTmp As New ADODB.Recordset
    
    
    'δ����øſ�
    
    '���ý��С����λ��
    
    
    strDec = "0.00"
    bytDec = 2
    bytDec = zlDatabase.GetPara(9, glngSys, , 2)
    strDec = "0." & String(bytDec, "0")
    
    strSQL = "Select Ԥ�����,�������,0 as Ԥ����� From ������� Where ����=1 And ����ID=[1] And ����=" & IIf(lng��ҳid = 0, "1 ", "2 ")
    strSQL = strSQL & " Union ALL " & _
        " Select 0 as Ԥ�����,0 as �������,Sum(B.���) as Ԥ�����" & _
        " From ������Ϣ A,����ģ����� B" & _
        " Where A.����ID=B.����ID And A.סԺ����=B.��ҳID" & _
        " And A.����ID=[1] "
    
    strSQL = "Select Sum(Ԥ�����) as Ԥ�����,Sum(�������) as �������,Sum(Ԥ�����) as Ԥ����� From (" & strSQL & ")"

    Set gRs = zlDatabase.OpenSQLRecord(strSQL, "���ò�ѯ", lng����ID)
    If Not gRs.EOF Then
        sngʣ���� = IIf(IsNull(gRs!Ԥ�����), 0, gRs!Ԥ�����) - IIf(IsNull(gRs!�������), 0, gRs!�������) + IIf(IsNull(gRs!Ԥ�����), 0, gRs!Ԥ�����)
        lblMoneyInfo.Caption = "Ԥ�����:" & Format(zlCommFun.Nvl(gRs!Ԥ�����, 0), "0.00") & Space(3) & _
                            "δ�����:" & Format(zlCommFun.Nvl(gRs!�������, 0), strDec) & Space(3) & _
                            "Ԥ�����:" & Format(zlCommFun.Nvl(gRs!Ԥ�����, 0), strDec) & Space(3) & _
                            "ʣ����:" & Format(zlCommFun.Nvl(gRs!Ԥ�����, 0) - zlCommFun.Nvl(gRs!�������, 0) + zlCommFun.Nvl(gRs!Ԥ�����, 0), "0.00")
        
        '����ϴε���ʾ��Ϣ
        lblWarn.BackStyle = 0
        lblWarn.Caption = ""
        
        If mvar��ҳid > 0 Then
            str�������� = ""
            strSQL = "Select zl_PatiWarnScheme([1],[2]) As �������� From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Ƿ�����", mvar����id, mvar��ҳid)
            If rsTmp.BOF = False Then
                str�������� = zlCommFun.Nvl(rsTmp("��������").Value)
            End If
                        
            Call Ƿ�����(mvar����, lng����ID, mvar��ҳid, sngʣ����, str��������)
        End If
        
        GetStateInfo = True
        
    Else
        lblMoneyInfo.Caption = "Ԥ�����:0.00" & Space(3) & "Ԥ�����:0.00" & Space(3) & "δ�����:" & strDec & Space(3) & "ʣ����:0.00"
    End If
    lblMoneyInfo.Tag = lblMoneyInfo.Caption
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCurrentStateInfo(ByVal lng����ID As Long, ByVal lng��ҳid As Long, ByVal strCurrDate As String) As Boolean
    
    '******************************************************************************************************************
    '���ܣ���ȡ���շ���״̬
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim bytDec As Byte
    Dim strDec As String
    Dim strSQL As String

    
    On Error GoTo errH
    
    '���ý��С����λ��
    
    If strCurrDate = "" Then
        
        lblMoneyInfo.Caption = lblMoneyInfo.Tag & Space(3) & "���úϼ�:" & Format(msfResult.Tag, strDec)
                            
        GetCurrentStateInfo = True
        Exit Function
    End If

    strDec = "0.00"
    bytDec = 2
    bytDec = zlDatabase.GetPara(9, glngSys, , 2)

    strDec = "0." & String(bytDec, "0")

    strSQL = "" & _
    "       Select Ԥ�����, �������" & vbNewLine & _
    "       From �������" & vbNewLine & _
    "       Where ���� = 1 And ����id = [1] And  ����=" & IIf(lng��ҳid = 0, "1 ", "2 ") & vbNewLine & _
    "       Union All" & vbNewLine & _
    "       Select 0 - Nvl(A.���, 0) As Ԥ�����, 0 As �������" & vbNewLine & _
    "       From ����Ԥ����¼ A" & vbNewLine & _
    "       Where A.����id = [1] And Nvl(A.��ҳid, 0) = [2] And A.��¼���� = 1 And A.�տ�ʱ�� > [3]" & vbNewLine & _
    "       Union All" & vbNewLine & _
    "       Select Nvl(A.��Ԥ��, 0) As Ԥ�����, 0 As �������" & vbNewLine & _
    "       From ����Ԥ����¼ A, ���˽��ʼ�¼ B" & vbNewLine & _
    "       Where A.����id = [1] And Nvl(A.��ҳid, 0) = [2] And A.��¼���� = 11 And B.ID(+) = A.����id And  Nvl(B.�շ�ʱ��, A.�տ�ʱ��) > [3]" & vbNewLine
    
    If lng��ҳid = 0 Then
        strSQL = strSQL & _
        "       Union All" & vbNewLine & _
        "       Select 0 As Ԥ�����, 0 - Nvl(ʵ�ս��, 0) As �������" & vbNewLine & _
        "       From ������ü�¼ A" & vbNewLine & _
        "       Where ����id = [1]  And ���ʷ��� = 1 And  �Ǽ�ʱ�� > [3]" & vbNewLine
    End If
    '97524:���ϴ�,2016/6/16,���ʽ��ȡ���סԺ���ü�¼�е��ֶ�
    strSQL = strSQL & _
    "       Union All" & vbNewLine & _
    "       Select 0 As Ԥ�����, 0 - Nvl(ʵ�ս��, 0) As �������" & vbNewLine & _
    "       From סԺ���ü�¼ A" & vbNewLine & _
    "       Where ����id = [1] And Nvl(��ҳid, 0) = [2] And ���ʷ��� = 1 And  �Ǽ�ʱ�� > [3]" & vbNewLine & _
    "       Union All" & vbNewLine & _
    "       Select 0 As Ԥ�����, Nvl(A.���ʽ��, 0) As �������" & vbNewLine & _
    "       From סԺ���ü�¼ A, ���˽��ʼ�¼ B" & vbNewLine & _
    "       Where A.����id = [1] And Nvl(A.��ҳid, 0) = [2] And A.���ʷ��� = 1 And B.ID = A.����id And  B.�շ�ʱ�� > [3] "
    If lng��ҳid = 0 Then
        strSQL = strSQL & _
        "       Union All" & vbNewLine & _
        "       Select 0 As Ԥ�����, Nvl(A.���ʽ��, 0) As �������" & vbNewLine & _
        "       From ������ü�¼ A, ���˽��ʼ�¼ B" & vbNewLine & _
        "       Where A.����id = [1] And A.���ʷ��� = 1 And B.ID = A.����id And  B.�շ�ʱ�� > [3] "
     End If
     
    strSQL = "" & _
    "Select Sum(Ԥ�����) As ����Ԥ�����, Sum(�������) As ���շ������, Sum(Ԥ�����) - Sum(�������)" & vbNewLine & _
    "From ( " & strSQL & ") "
    

    strCurrDate = Format(strCurrDate, "yyyy-MM-dd") & " 23:59:59"
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "���ò�ѯ", lng����ID, lng��ҳid, CDate(strCurrDate))
    If rs.BOF = False Then
        
    
        lblMoneyInfo.Caption = lblMoneyInfo.Tag & Space(3) & "����ʣ���:" & Format(zlCommFun.Nvl(rs!����Ԥ�����, 0) - zlCommFun.Nvl(rs!���շ������, 0), strDec) & Space(3) & _
                            "���շ��úϼ�:" & Format(msfResult.Tag, strDec)
            
        GetCurrentStateInfo = True
        
    Else
        lblMoneyInfo.Caption = lblMoneyInfo.Tag & Space(3) & "����ʣ���:0.00" & Space(3) & "���շ��úϼ�:0.00"
    End If

    Exit Function
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadԤ������(ByVal lng����ID As Long, ByVal lng��ҳid As Long) As Boolean
'����:װ�ز��˵�����Ԥ�����ü�¼
'����:lng����id         ����Ψһ��
'     lng��ҳid         ����סԺ����
'����:True�ɹ�;False:ʧ��
    Dim sglMoney(1 To 3) As Single
    Dim i As Long
    
    On Error GoTo errH
    
    i = 1
    msfResult.Rows = 2
    Call ClearSpecRowCol(msfResult, 1, Array())
    
    
    gstrSQL = "Select To_Char(A.�տ�ʱ��,'YYYY-MM-DD') as ����,A.NO as ���ݺ�," & _
            " B.���� as ����,A.���㷽ʽ,A.�������," & _
            " Ltrim(To_Char(Sum(Nvl(A.���,0)),'9999999990.00')) as Ԥ�����," & _
            " Ltrim(To_Char(Sum(Nvl(A.��Ԥ��,0)),'9999999990.00')) as ���ʽ��," & _
            " Ltrim(To_Char(Sum(Nvl(A.���,0)-Nvl(A.��Ԥ��,0)),'9999999990.00')) as ʣ����," & _
            " A.ժҪ" & _
            " From ����Ԥ����¼ A,���ű� B" & _
            " Where A.����ID=B.ID(+) And A.��¼���� IN(1,11) And Nvl(A.���,0)<>Nvl(A.��Ԥ��,0) And A.����ID=[1] and A.��ҳid= [2] " & _
            " Having Sum(Nvl(A.���, 0) - Nvl(A.��Ԥ��, 0)) <> 0" & _
            " Group by To_Char(�տ�ʱ��,'YYYY-MM-DD'),A.NO,B.����,A.���㷽ʽ,A.�������,A.ժҪ" & _
            " Order by ����,���ݺ�"
        
    gstrSQL = "Select To_Char(A.�տ�ʱ��,'YYYY-MM-DD') as ����,A.NO as ���ݺ�," & _
            " B.���� as ����,A.���㷽ʽ,A.�������," & _
            " Ltrim(To_Char(Sum(Nvl(A.���,0)),'9999999990.00')) as Ԥ�����," & _
            " Ltrim(To_Char(Sum(Nvl(A.��Ԥ��,0)),'9999999990.00')) as ���ʽ��," & _
            " Ltrim(To_Char(Sum(Nvl(A.���,0)-Nvl(A.��Ԥ��,0)),'9999999990.00')) as ʣ����," & _
            " A.ժҪ" & _
            " From ����Ԥ����¼ A,���ű� B" & _
            " Where A.����ID=B.ID(+) And A.��¼���� IN(1,11) And A.����ID=[1] and A.��ҳid= [2]" & _
            " Group by To_Char(�տ�ʱ��,'YYYY-MM-DD'),A.NO,B.����,A.���㷽ʽ,A.�������,A.ժҪ" & _
            " Order by ����,���ݺ�"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "���ò�ѯ", lng����ID, lng��ҳid)
    If gRs.RecordCount > 0 Then
        gRs.MoveFirst
        While Not gRs.EOF
            msfResult.TextMatrix(i, 0) = IIf(IsNull(gRs!����), "", gRs!����)
            msfResult.TextMatrix(i, 1) = IIf(IsNull(gRs!����), "", gRs!����)
            msfResult.TextMatrix(i, 2) = IIf(IsNull(gRs!���㷽ʽ), "", gRs!���㷽ʽ)
            msfResult.TextMatrix(i, 3) = IIf(IsNull(gRs!�������), "", gRs!�������)
            msfResult.TextMatrix(i, 4) = IIf(IsNull(gRs!Ԥ�����), "", Format(gRs!Ԥ�����, "0.00"))
            sglMoney(1) = sglMoney(1) + Val(msfResult.TextMatrix(i, 4))
            i = i + 1
            msfResult.Rows = i + 1
            gRs.MoveNext
        Wend
        If msfResult.Rows > 2 Then
            msfResult.MergeRow(msfResult.Rows - 1) = True
            For i = 0 To msfResult.Cols - 2
                msfResult.TextMatrix(msfResult.Rows - 1, i) = "�ϼ�"
            Next
            msfResult.TextMatrix(msfResult.Rows - 1, 4) = Format(sglMoney(1), "0.00")
        End If
        
        mvarRows = msfResult.Rows - 1
    End If
    
    msfResult.Rows = msfResult.Rows + 50
    
    LoadԤ������ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Load������ϸ(ByVal lng����ID As Long, ByVal lng��ҳid As Long) As Boolean
    '******************************************************************************************************************
    '����:װ�ز���ָ�����ڵ�δ��/�ѽ���ü�¼
    '����:lng����id         ����Ψһ��
    '     lng��ҳid         ����סԺ����
    '����:True�ɹ�;False:ʧ��
    '******************************************************************************************************************
    Dim i As Long
    Dim strTmp As String
    Dim sglMoney As Single
    Dim v_Class As String
    Dim strBill As String
    Dim strSQL As String
    Dim blnShowExit As Boolean
    Dim str���ʷ��� As String
    Dim strFields As String, strWhere As String, strSubTable As String
    
    On Error GoTo errH
    
    '------------------------------------------------------------------------------------------------------------------
    v_Class = GetPara("���ò�����ϸ")
    i = 1
    msfResult.Rows = 2
    Call ClearSpecRowCol(msfResult, 1, Array())
    msfResult.Tag = ""
    
    
    blnShowExit = (UsrCmd(22).Visible And UsrCmd(22).Tag = "1")
    
    Select Case mbytQueryMode
    '------------------------------------------------------------------------------------------------------------------
    Case 20, 21                     '��Ʊ�ݺź͵��ݺŲ����
        If mbytQueryMode = 20 Then
            'Ʊ�ݺ�
            strSQL = " Where ��������=1 And ID=(Select ��ӡID From (Select ����,��ӡID,Ʊ��,���� From Ʊ��ʹ����ϸ  Union All Select ����,��ӡID,Ʊ��,���� From HƱ��ʹ����ϸ) Where Ʊ��=1 And ����=1 And ����=[8] And Rownum=1)"
            strSQL = "Select NO From Ʊ�ݴ�ӡ���� " & strSQL & " Union All Select NO From HƱ�ݴ�ӡ����" & strSQL
            strBill = " And A.NO In (" & strSQL & ") "
        Else
            '���ݺ�
            strBill = " And A.No=[8] "
        End If
        
        gstrSQL = ""
        
        '���ڲ���ʾ������ϸ�����,ֻ��ʾ����һ���ܶ�
        '--------------------------------------------------------------------------------------------------------------
        If v_Class <> "" Then
        
            gstrSQL = "" & _
            " Select   '' As ˵��, '' as �շѱ���, B.����, AA.����, '' as ���ݺ�, '' as ����, B.���� AS ���, '' as ����," & _
            "       0 as ����, '' as ����, '' as ����, LTRIM(TO_CHAR(AA.ʵ�ս��,'99999990.00')) as ���,1 as ��¼״̬ " & _
            " From (    Select Trunc(" & mstr����ʱ������ & ") as ����, �շ����, Sum(NVL(ʵ�ս��,0)) As ʵ�ս�� " & _
            "           From ������ü�¼ A , Table(Cast(f_Str2list([7]) As zlTools.t_Strlist)) b " & _
            "           Where  A.��¼����=1  And A.�����־=1  And A.��¼״̬<>0 AND A.�շ����=B.Column_Value " & strBill & " " & _
            "           Group by trunc(" & mstr����ʱ������ & ") ,�շ���� " & _
            "       ) AA, �շ���Ŀ��� B " & _
            "Where AA.�շ����=B.����"
        End If
        
        If v_Class <> "" Then strTmp = " And Not Exists(Select 1 From  Table(Cast(f_Str2list([7]) As zlTools.t_Strlist)) b where A.�շ���� =B.Column_Value) "

        'blnShowExit:��ʾ�˷�:false=:����ʾ�˷�
        strFields = IIf(blnShowExit, "A.����, A.����, A.ʵ�ս��,A.��¼״̬ ", "Decode(a.��¼״̬,1,Nvl(A.����,1),1) As ����, Sum(Decode(a.��¼״̬,1,1, Nvl(A.����,1))*A.����) As ����, Sum(A.ʵ�ս��) as ʵ�ս��,1 As ��¼״̬")

        gstrSQL = IIf(gstrSQL <> "", gstrSQL & " Union ", "") & _
        "Select B.˵��, B.���� As �շѱ���,C.����,A.����, A.NO as ���ݺ�, D.���� as ����, C.���� AS ���, B.����||Decode(nvl(b.��������,''),'','','('|| b.�������� || ')')  as ����," & _
        "       A.����, Decode(A.��¼����,4,'',Decode(A.���㵥λ,NULL,NULL,A.���㵥λ)||Decode(A.�շ����,'7','��'||A.����||'��',NULL)) as ���� ," & _
        "       LTRIM(TO_CHAR(A.��׼����,'9999990.0000')) as ����, LTRIM(TO_CHAR(A.ʵ�ս��,'99999990.00')) as ���,A.��¼״̬ " & _
        "From ( Select trunc(A." & mstr����ʱ������ & ") as ����, A.NO, A.��������ID, A.�շ����, A.�շ�ϸĿID, A.���㵥λ, A.��¼����, A.��׼����," & strFields & _
        "       From   ������ü�¼ A " & _
        "       Where  �����־=1  and A.��¼״̬<>0  and A.��¼����=1  and nvl(A.ʵ�ս��,0)<>0 " & strTmp & strBill & _
                IIf(blnShowExit, "          Order By ", " Group By  ") & "  Trunc(A." & mstr����ʱ������ & "),A.NO,A.��������ID,A.�շ����,A.�շ�ϸĿID,A.���㵥λ" & IIf(blnShowExit, "", ",A.��¼����,A.��׼����,Decode(a.��¼״̬,1,Nvl(A.����,1),1) ") & _
        "       ) A, �շ���ĿĿ¼ B, �շ���Ŀ��� C, ���ű� D " & _
        " Where A.�շ����=C.����  And A.�շ�ϸĿID=B.ID  And A.��������ID=D.ID(+)  "
 
        gstrSQL = "select ˵��, �շѱ���,����,����,���ݺ�,����,���,����,����,����,����,���,��¼״̬ from (" & gstrSQL & ") where ���<>0 order by ����,���� Desc,���ݺ�"
    '------------------------------------------------------------------------------------------------------------------
    Case Else
        strWhere = ""
        If UsrCmd(28).Visible And UsrCmd(28).Tag = "1" And UsrCmd(29).Visible And UsrCmd(29).Tag = "1" Then
            
        ElseIf UsrCmd(28).Visible And UsrCmd(28).Tag = "1" Then
            '�շѷ���
            strWhere = strWhere & " And A.���ʷ���<>1"
        ElseIf UsrCmd(29).Visible And UsrCmd(29).Tag = "1" Then
            '���ʷ���
             strWhere = strWhere & "And A.���ʷ���=1"
        End If
        
        If lblRange.Visible Then strWhere = strWhere & " And (A." & mstr����ʱ������ & ">=[5] And A." & mstr����ʱ������ & "<[6]  ) "
        
        If lng��ҳid <= 0 And mlngDays > 0 Then strTmp = " And A." & mstr����ʱ������ & " Between [3] And [4]    "
        If v_Class <> "" Then strTmp = strTmp & " And Not Exists(Select 1 From  Table(Cast(f_Str2list([7]) As zlTools.t_Strlist)) b where A.�շ���� =B.Column_Value) "
        
    
         
        'blnShowExit:��ʾ�˷�:false=:����ʾ�˷�
        strFields = IIf(blnShowExit, "A.����,A.����,A.ʵ�ս��,A.��¼״̬  ", "Decode(a.��¼״̬,1,Nvl(A.����,1),1) As ����,Sum(Decode(a.��¼״̬,1,1, Nvl(A.����,1))*A.����) As ����,Sum(A.ʵ�ս��) as ʵ�ս��,1 As ��¼״̬")
        
        gstrSQL = ""
        If v_Class <> "" Then
            
            strSubTable = "" & _
                "        Select Trunc(" & mstr����ʱ������ & ") as ����, �շ����, Sum(ʵ�ս��) as ʵ�ս�� " & _
                "        From סԺ���ü�¼ A, Table(Cast(f_Str2list([7]) As zlTools.t_Strlist)) b " & _
                "        Where  A.��¼״̬<>0 And ����ID=[1] And nvl(��ҳID,0)=[2] And a.�շ����=B.Column_Value " & strWhere & _
                "        Group By Trunc(" & mstr����ʱ������ & "),�շ���� "
            If lng��ҳid = 0 Then
                strSubTable = strSubTable & "" & _
                "        Union ALL  " & _
                "        Select Trunc(" & mstr����ʱ������ & ") as ����, �շ����, Sum(ʵ�ս��) as ʵ�ս�� " & _
                "        From ������ü�¼ A, Table(Cast(f_Str2list([7]) As zlTools.t_Strlist)) b " & _
                "        Where  A.��¼״̬<>0 And ����ID=[1]  And a.�շ����=B.Column_Value " & strWhere & _
                "        Group By Trunc(" & mstr����ʱ������ & "),�շ���� "
            End If
            
            gstrSQL = "" & _
            " Select '' As ˵��, '' as �շѱ���, B.����, AA.����, '' as ���ݺ�, '' as ����, B.���� AS ���, '' as ����," & _
            "       0 as ����, '' as ����, '' as ����, LTRIM(TO_CHAR(Sum(AA.ʵ�ս��),'99999990.00')) as ���,1 as ��¼״̬ " & _
            " From ( " & strSubTable & ")AA, �շ���Ŀ��� B " & _
            " Where AA.�շ����=B.���� " & _
            " Group by AA.����,b.����,b.����"
        End If
        
        strSubTable = "" & _
        "           Select Trunc(A." & mstr����ʱ������ & ") as ����, A.NO, A.��������ID, A.�շ����, A.�շ�ϸĿID, A.���㵥λ, A.��¼����, A.��׼����,A.����,A.����,A.��¼״̬,A.ʵ�ս�� " & _
        "           From סԺ���ü�¼ A  " & _
        "           Where   A.��¼״̬<>0 And A.����ID=[1] And nvl(��ҳID,0)=[2] " & strTmp & strWhere & _
        "                   And Nvl(A.ʵ�ս��,0)<>0 "
        If lng��ҳid = 0 Then
            strSubTable = strSubTable & "" & _
            "        Union ALL  " & _
        "           Select Trunc(A." & mstr����ʱ������ & ") as ����, A.NO, A.��������ID, A.�շ����, A.�շ�ϸĿID, A.���㵥λ, A.��¼����, A.��׼����,A.����,A.����,A.��¼״̬,A.ʵ�ս�� " & _
            "        From ������ü�¼ A  " & _
            "        Where   A.��¼״̬<>0 And A.����ID=[1] " & strTmp & strWhere & _
            "                   And Nvl(A.ʵ�ս��,0)<>0 "
        End If
        
        strSubTable = "" & _
        "   Select trunc(����) as ����,NO,��������ID,�շ����,�շ�ϸĿID,���㵥λ,��¼����,��׼����," & strFields & _
        "   From (" & strSubTable & ") A" & _
            IIf(blnShowExit, "          Order By ", " Group By  ") & " Trunc(A.����),A.NO,A.��������ID,A.�շ����,A.�շ�ϸĿID,A.���㵥λ " & IIf(blnShowExit, "", ",A.��¼����,A.��׼����,Decode(a.��¼״̬,1,Nvl(A.����,1),1) ")

        '48546:��λ��������
        gstrSQL = IIf(gstrSQL <> "", gstrSQL & " Union ", "") & _
        " Select B.˵��,  B.���� As �շѱ���, C.����,A.����, A.No As ���ݺ�, D.���� As ����, C.���� AS ���, B.����||decode(nvl(B.��������,''),'','','('|| B.�������� || ')')  As ����, A.����, " & _
        "       Decode(A.��¼����,4,'',Decode(A.���㵥λ,NULL,NULL,'('||A.���㵥λ||')')||Decode(A.�շ����,'7','��'||A.����||'��',NULL)) As ���� , " & _
        "       LTRIM(TO_CHAR(A.��׼����,'9999990.0000')) As ����, LTRIM(TO_CHAR(A.ʵ�ս��,'99999990.00')) As ���,A.��¼״̬  " & _
        " From ( " & strSubTable & ") A, �շ���ĿĿ¼ B, �շ���Ŀ��� C, ���ű� D " & _
        " Where A.�շ����=C.���� And A.�շ�ϸĿID=B.ID And A.��������ID=D.ID(+)  "
            
        gstrSQL = "Select ˵��, �շѱ���,����,����,���ݺ�,����,���,����,����,����,����,���,��¼״̬ from (" & gstrSQL & ") where ���<>0 order by ����,���� Desc,���ݺ�"
    End Select
    
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "���ò�ѯ", lng����ID, lng��ҳid, CDate(Format(mvarMinDate, "YYYY-MM-DD") & " 00:00:00"), CDate(Format(mvarMaxDate, "YYYY-MM-DD") + " 23:59:59"), CDate(lblRange.Caption), CDate(lblRange.Caption) + 1, Replace(v_Class, "'", ""), mstrNO)
 
    If gRs.RecordCount > 0 Then
        gRs.MoveFirst
        While Not gRs.EOF
            msfResult.TextMatrix(i, 0) = IIf(IsNull(gRs!����), "", gRs!����)
            msfResult.TextMatrix(i, 1) = IIf(IsNull(gRs!����), "", gRs!����)
            msfResult.TextMatrix(i, 2) = IIf(IsNull(gRs!���), "", gRs!���)
            msfResult.TextMatrix(i, 3) = IIf(IsNull(gRs!����), "", gRs!����)
            
            If IsNull(gRs!����) = False Then
                msfResult.TextMatrix(i, 4) = IIf(gRs!���� = 0, "", gRs!���� & IIf(IsNull(gRs!����), "", gRs!����))
            End If
            
            msfResult.TextMatrix(i, 5) = IIf(IsNull(gRs!����), "", Format(gRs!����, "0.00##"))
            msfResult.TextMatrix(i, 6) = IIf(IsNull(gRs!���), "", Format(gRs!���, "0.00"))
            msfResult.TextMatrix(i, 7) = IIf(IsNull(gRs!�շѱ���), "", gRs!�շѱ���)
            msfResult.TextMatrix(i, 8) = IIf(IsNull(gRs!˵��), "", gRs!˵��)
            
            sglMoney = sglMoney + Val(msfResult.TextMatrix(i, 6))
            If Val(IIf(IsNull(gRs!��¼״̬), "0", gRs!��¼״̬)) = 2 Then
                '����ǰ��ɫ
                msfResult.Cell(flexcpForeColor, i, 0, i, msfResult.Cols - 1) = 255
            End If
            
            i = i + 1
            msfResult.Rows = i + 1
            gRs.MoveNext
        Wend
        
        If msfResult.Rows > 2 Then
            msfResult.MergeRow(msfResult.Rows - 1) = True
            For i = 0 To 5
                msfResult.TextMatrix(msfResult.Rows - 1, i) = "�ϼ�"
            Next
            msfResult.TextMatrix(msfResult.Rows - 1, 6) = Format(sglMoney, "0.00")
            msfResult.Tag = sglMoney
        End If
        
        mvarRows = msfResult.Rows - 1
    End If
    msfResult.Rows = msfResult.Rows + 50
    
    Load������ϸ = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetDateList(ByVal lng����ID As Long, ByVal lng��ҳid As Long, ByVal strMinDate As String, ByVal strMaxDate As String, Optional ByVal strDateType As String = "����ʱ��") As ADODB.Recordset
    '******************************************************************************************************************
    '����:��ȡ�з������ݵ������嵥
    '����:lng����id         ����Ψһ��
    '     lng��ҳid         ����סԺ����
    '******************************************************************************************************************
    
    Dim strSQL As String
    
    
    strSQL = "" & _
    "   Select To_Char(" & strDateType & ",'yyyy-mm-dd') As ���� " & _
    "   From סԺ���ü�¼ " & _
    "   Where ����id=[1] And Nvl(��ҳid,0)=[2] And " & strDateType & " Between [3] And [4] " & _
    "   Group by To_Char(" & strDateType & ",'yyyy-mm-dd')"
    If lng��ҳid = 0 Then
        strSQL = strSQL & " UNION " & _
        "   Select To_Char(" & strDateType & ",'yyyy-mm-dd') As ���� " & _
        "   From ������ü�¼ " & _
        "   Where ����id=[1] And " & strDateType & " Between [3] And [4] " & _
        "   Group by To_Char(" & strDateType & ",'yyyy-mm-dd')"
    End If
    Set GetDateList = zlDatabase.OpenSQLRecord(strSQL, "���ò�ѯ", lng����ID, lng��ҳid, CDate(strMinDate), CDate(strMaxDate & " 23:59:59"))

    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function
 

Private Function Load�������(ByVal lng����ID As Long, ByVal lng��ҳid As Long) As Boolean
'����:װ�ز���ָ�����ڵ�δ��/�ѽ���ü�¼
'����:lng����id         ����Ψһ��
'     lng��ҳid         ����סԺ����
'����:True�ɹ�;False:ʧ��
    Dim dtStartDate As Date, dtEndDate As Date, dtStartDate1 As Date, dtEndDate1 As Date
    Dim i As Long, j As Long
    Dim strTmp As String, strWhere As String
    Dim sglMoney(1 To 3) As Single
    Dim svrSave As String
    Dim strSQL As String
    Dim strBill As String
    Dim str���ʷ��� As String
    
    On Error GoTo errH
    
    dtStartDate = CDate("1901-01-01"): dtEndDate = dtStartDate: dtStartDate1 = dtStartDate: dtEndDate1 = dtStartDate
    i = 1
    msfResult.Rows = 2
    Call ClearSpecRowCol(msfResult, 1, Array())
    
    Select Case mbytQueryMode
    Case 20, 21
        If mbytQueryMode = 20 Then
            strSQL = " Where ��������=1 And ID=(Select ��ӡID From " & zlGetFullFieldsTable("Ʊ��ʹ����ϸ") & " Where Ʊ��=1 And ����=1 And ����=[3] And Rownum=1)"
            strSQL = "Select NO From Ʊ�ݴ�ӡ���� " & strSQL & " Union All Select NO From HƱ�ݴ�ӡ����" & strSQL
            strBill = " And A.NO In (" & strSQL & ") "
        Else
            strBill = " And A.No=[3] "
        End If
        
        gstrSQL = "" & _
        "SELECT A.�վݷ�Ŀ as �վݷ�Ŀ," & _
        "       Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0)),'9999999990.00')) as ���," & _
        "       Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'9999999990.00')) as δ���ʽ��" & _
        " FROM ������ü�¼ A " & _
        " Where A.��¼״̬=1 and A.��¼����=1 and �����־=1 " & strBill & _
        " Group by A.�վݷ�Ŀ " & _
        " Order by �վݷ�Ŀ"
    Case Else
         strWhere = ""
        
        If UsrCmd(28).Visible And UsrCmd(28).Tag = "1" And UsrCmd(29).Visible And UsrCmd(29).Tag = "1" Then
            
        ElseIf UsrCmd(28).Visible And UsrCmd(28).Tag = "1" Then
            '�շѷ���
            strWhere = strWhere & " And A.���ʷ���<>1 "
        ElseIf UsrCmd(29).Visible And UsrCmd(29).Tag = "1" Then
            '���ʷ���
            strWhere = strWhere & "And A.���ʷ���=1 "
        End If
        
        If lng��ҳid <= 0 And mlngDays > 0 Then strWhere = strWhere & " And A." & mstr����ʱ������ & " between [4] and [5] "
        If lblRange.Visible Then strWhere = strWhere & " And A." & mstr����ʱ������ & ">=[6] and A." & mstr����ʱ������ & "<[7] "

                        
        '���ڸ���7��    ������2363
        gstrSQL = "" & _
        " SELECT A.�վݷ�Ŀ as �վݷ�Ŀ," & _
        "       Sum(Nvl(A.ʵ�ս��,0)) as ���," & _
        "       Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)) as δ���ʽ��" & _
        " FROM סԺ���ü�¼ A  " & _
        " Where A.��¼״̬<>0   And A.����ID=[1] and nvl(A.��ҳid,0)=[2] " & strWhere & _
        " Group by A.�վݷ�Ŀ "
        If lng��ҳid = 0 Then
            gstrSQL = gstrSQL & " Union ALL " & _
            " SELECT A.�վݷ�Ŀ as �վݷ�Ŀ," & _
            "       Sum(Nvl(A.ʵ�ս��,0)) as ���," & _
            "       Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)) as δ���ʽ��" & _
            " FROM ������ü�¼ A  " & _
            " Where A.��¼״̬<>0   And A.����ID=[1] " & strWhere & _
            " Group by A.�վݷ�Ŀ "
        End If
        '77306,���ϴ�,2014/9/1,δ���˽������˽��˽��
        gstrSQL = "" & _
        " SELECT A.�վݷ�Ŀ," & _
        "       Ltrim(To_Char(Sum(Nvl(A.���,0)),'9999999990.00')) as ���," & _
        "       Ltrim(To_Char(Sum(Nvl(A.δ���ʽ��,0)),'9999999990.00')) as δ���ʽ��" & _
        " FROM (" & gstrSQL & ") A  " & _
        " Group by A.�վݷ�Ŀ " & _
        " Order by �վݷ�Ŀ"
        
    End Select
    
    If IsDate(Format(mvarMinDate, "YYYY-MM-DD")) Then dtStartDate = CDate(Format(mvarMinDate, "YYYY-MM-DD") + " 00:00:00")
    If IsDate(Format(mvarMaxDate, "YYYY-MM-DD")) Then dtEndDate = CDate(Format(mvarMaxDate, "YYYY-MM-DD") + " 23:59:59")
    If lblRange.Visible Then
        dtStartDate1 = CDate(Format(lblRange.Caption, "YYYY-MM-DD HH:MM:SS")): dtEndDate = CDate(Format(CDate(lblRange.Caption) + 1, "YYYY-MM-DD HH:MM:SS"))
    End If
    
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "���ò�ѯ", lng����ID, lng��ҳid, mstrNO, dtStartDate, dtEndDate, dtStartDate1, dtEndDate1)
    If gRs.RecordCount > 0 Then
        gRs.MoveFirst
        While Not gRs.EOF
            msfResult.TextMatrix(i, 0) = IIf(IsNull(gRs!�վݷ�Ŀ), "", gRs!�վݷ�Ŀ)
            msfResult.TextMatrix(i, 1) = IIf(IsNull(gRs!���), "", gRs!���)
            msfResult.TextMatrix(i, 3) = IIf(IsNull(gRs!δ���ʽ��), "", gRs!δ���ʽ��)
            msfResult.TextMatrix(i, 2) = Format(Val(msfResult.TextMatrix(i, 1)) - Val(msfResult.TextMatrix(i, 3)), "0.00")
            
            sglMoney(1) = sglMoney(1) + Val(msfResult.TextMatrix(i, 1))
            sglMoney(2) = sglMoney(2) + Val(msfResult.TextMatrix(i, 2))
            sglMoney(3) = sglMoney(3) + Val(msfResult.TextMatrix(i, 3))
            i = i + 1
            msfResult.Rows = i + 1
            gRs.MoveNext
        Wend
        
        '�¼�һС����
        Call SetMsfRowColor(msfResult, i, &H80000018)
        msfResult.MergeRow(i) = True
        For j = 0 To msfResult.Cols - 4
            msfResult.TextMatrix(i, j) = "�ϼ�:"
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
    
    Load������� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub DrawMsfHeader(ByVal bytMode As Byte)
'����:������ѡ����������Ӧ�ı��
'����:bytMode       ��ѡȡ��������
    
    With msfResult
        .Rows = 50
        .Cols = 0
        ClearSpecRowCol msfResult, 1, Array()
        Select Case bytMode
        Case 0      '������ϸ
            Call AddColumn(msfResult, "����", 1500, 1)
            Call AddColumn(msfResult, "����", 1200, 1)
            Call AddColumn(msfResult, "���", 1200, 1)
            Call AddColumn(msfResult, "����", 3600, 1)
            Call AddColumn(msfResult, "����", 1200, 7)
            Call AddColumn(msfResult, "����", 1200, 7)
            Call AddColumn(msfResult, "���", 1200, 7)
            Call AddColumn(msfResult, "����", 1200, 7)
            Call AddColumn(msfResult, "˵��", 0, 1)
            Call AddColumn(msfResult, "", 1200, 7)
            Call CalcAutoColWidth(msfResult, 3)
        Case 1      '�������
            Call AddColumn(msfResult, "�շѷ�Ŀ", 1500, 1)
            Call AddColumn(msfResult, "���", 1800, 7)
            Call AddColumn(msfResult, "�ѽ��ʽ��", 1800, 7)
            Call AddColumn(msfResult, "δ���ʽ��", 1800, 7)
            Call AddColumn(msfResult, "", 1200, 7)
            Call CalcAutoColWidth(msfResult, 4)
        Case 2
            Call AddColumn(msfResult, "����", 1290, 1)
            Call AddColumn(msfResult, "����", 1500, 1)
            Call AddColumn(msfResult, "���㷽ʽ", 990, 1)
            Call AddColumn(msfResult, "�������", 990, 1)
            Call AddColumn(msfResult, "Ԥ�����", 1500, 7)
            Call AddColumn(msfResult, "", 1200, 7)
            Call CalcAutoColWidth(msfResult, 5)
        End Select
    End With
    
    msfResult.Cell(flexcpFontName, 0, 0, 0, msfResult.Cols - 1) = "����_GB2312"
    msfResult.Cell(flexcpFontSize, 0, 0, 0, msfResult.Cols - 1) = 14
    msfResult.Cell(flexcpFontBold, 0, 0, 0, msfResult.Cols - 1) = True
End Sub

Private Sub UsrCmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    mlngLoopCount = 0
End Sub

Private Function CheckQueryPass(ByVal lng����ID As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim IntCount As Integer
    Dim strInputPassword As String
    
    On Error GoTo errH:
    
    strSQL = "Select ��ѯ���� from ������Ϣ where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ϣ", lng����ID)
    If Not rsTemp.EOF Then
        If Trim(Nvl(rsTemp!��ѯ����)) <> "" Then
            For IntCount = 1 To 3
                With frmCheckQueryPass
                    .Show 1, Me
                    If .mblnOK = False Then
                        Exit Function
                    Else
                        strInputPassword = zlCommFun.zlStringEncode(LCase(.mstrPass))
                        
                        If LCase(Nvl(rsTemp!��ѯ����)) <> strInputPassword Then
                            If IntCount < 3 Then MsgBox "����������벻�ԣ����������룡", vbInformation, gstrSysName
                        Else
                            Exit For
                        End If
                    End If
                End With
            Next
            If IntCount > 3 Then '��������
                MsgBox "����������볬�����ξ����󣬽�ֹ��ѯ��", vbInformation, gstrSysName: Exit Function
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
'    If mvar����id > 0 Then
        
        Select Case Val(GetPara("����ʱ������", "0"))
        Case 0
            mstr����ʱ������ = "����ʱ��"
        Case 1
            mstr����ʱ������ = "�Ǽ�ʱ��"
        Case Else
            mstr����ʱ������ = "����ʱ��"
        End Select
    
        picAsk(0).Visible = False
        Call LoadPatient(mvar����id, mvar��ҳid, mvar����)
        
'        UsrCmd(2).Visible = IIf(mvar��ҳid > 0, True, False)
            
        UsrCmd(5).Picture = ilsImage.ListImages("unselect")
        UsrCmd(5).Tag = ""
        lblRange.Visible = False
        UsrCmd(6).Visible = False
        UsrCmd(7).Visible = False
        
        UsrCmd(28).Visible = Not (mbytQueryMode = 20 Or mbytQueryMode = 21)
        UsrCmd(29).Visible = Not (mbytQueryMode = 20 Or mbytQueryMode = 21)
        
        Set mrsDateList = GetDateList(mvar����id, mvar��ҳid, mvarMinDate, mvarMaxDate, mstr����ʱ������)
        
        '��ʽ�����ѯ����
        Call UsrCmd_CommandClick(0)
        
        Call GetCurrentStateInfo(mvar����id, mvar��ҳid, "")
        Call GetStateInfo(mvar����id, mvar��ҳid)
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
    Case "ȷ��"
        
        '������סԺ�Ų�ѯ���˷���
        mvar����id = 0
        mvar��ҳid = 0
        mvar���� = ""
        
        Call mclsIDKind.LeaveInputBox
        
        Select Case mbytQueryMode
        '--------------------------------------------------------------------------------------------------------------
        Case 12 'סԺ��
            If IsNumeric(txt(4).Text) Then
                gstrSQL = "select rownum as No,A.��Ժ����,A.��Ժ����,A.����id,A.��ҳid from ������ҳ A,������Ϣ B where A.����id=B.����id and B.סԺ��=[1] order by A.��Ժ���� desc"
                Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "���ò�ѯ", Trim(txt(4).Text))
                If gRs.RecordCount > 0 Then
                    If Not CheckQueryPass(Nvl(gRs!����id, 0)) Then Exit Sub '����У��
                    If frmSelect.ShowSelect(gRs, mvar����id, mvar��ҳid) Then
                        Call EnterCharge
                    End If
                Else
                    MsgBox "������һ�������ڵ�סԺ�ţ����������룡", vbInformation, gstrSysName
                End If
            Else
                MsgBox "������һ����ȷ��סԺ�ţ�", vbInformation, gstrSysName
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 3      '����������Ų�ѯ���˷���
            If IsNumeric(txt(4).Text) Then
                Set gRs = zlDatabase.OpenSQLRecord("select B.����id from ������Ϣ B where B.�����=[1]", "���ò�ѯ", Trim(txt(4).Text))
                If gRs.RecordCount > 0 Then
                    If Not CheckQueryPass(Nvl(gRs!����id, 0)) Then Exit Sub '����У��
                    mvar����id = IIf(IsNull(gRs!����id), 0, gRs!����id)
                    Call EnterCharge
                Else
                    MsgBox "������һ�������ڵ�����ţ����������룡", vbInformation, gstrSysName
                End If
            Else
                MsgBox "������һ����ȷ������ţ�", vbInformation, gstrSysName
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 14         '������ID�Ų�ѯ���˷���
            If IsNumeric(txt(4).Text) Then
                
                gstrSQL = "" & _
                "   Select to_char(A.��Ժ����,'yyyy-mm-dd') as ��Ժ����,to_char(A.��Ժ����,'yyyy-mm-dd') as ��Ժ����,A.����id,A.��ҳid " & _
                "   From ������ҳ A " & _
                "   Where A.����id=[1] " & _
                "   Union ALL " & _
                "   Select '�������' as ��Ժ����,'�������' as ��Ժ����,����id, 0 as ��ҳid " & _
                "   From ������ü�¼ C " & _
                "   Where ����id=[1] and rownum<2   "
                
                gstrSQL = "Select ��Ժ����,��Ժ����,����id,��ҳid From (" & gstrSQL & ") AA  order by AA.��Ժ���� asc,AA.��ҳid desc"
                gstrSQL = "SELECT RowNum as No,D.��Ժ����,D.��Ժ����,D.����id,D.��ҳid FROM (" & gstrSQL & ") D"
                
                Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "���ò�ѯ", Val(txt(4).Text))
                If gRs.RecordCount > 0 Then
                    If Not CheckQueryPass(Nvl(gRs!����id, 0)) Then Exit Sub '����У��
                    If frmSelect.ShowSelect(gRs, mvar����id, mvar��ҳid) Then
                        Call EnterCharge
                    End If
                   
                Else
                    MsgBox "������һ�������ڵĲ���ID�ţ����������룡", vbInformation, gstrSysName
                End If
            Else
                MsgBox "������һ����ȷ�Ĳ���ID�ţ�", vbInformation, gstrSysName
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 11, 30, 31, 32, 33, 34, 35         '�����Ų�
            
            gstrSQL = "" & _
            "   Select to_char(A.��Ժ����,'yyyy-mm-dd') as ��Ժ����,to_char(A.��Ժ����,'yyyy-mm-dd') as ��Ժ����,A.����id,A.��ҳid " & _
            "   From ������ҳ A " & _
            "   Where A.����id=[1] " & _
            "   Union ALL " & _
            "   Select '�������' as ��Ժ����,'�������' as ��Ժ����,����id, 0 as ��ҳid " & _
            "   From ������ü�¼ C " & _
            "   Where ����id=[1] and rownum<2   "
            
            gstrSQL = "Select ��Ժ����,��Ժ����,����id,��ҳid From (" & gstrSQL & ") AA  order by AA.��Ժ���� asc,AA.��ҳid desc"
            gstrSQL = "SELECT RowNum as No,D.��Ժ����,D.��Ժ����,D.����id,D.��ҳid FROM (" & gstrSQL & ") D"
            
            If mclsIDKind.zlGetPatiIDByCardNo(txt(4).Text, lngPatientKey, strPassWord) Then

                Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "���ò�ѯ", lngPatientKey)
                If gRs.RecordCount > 0 Then
                
                    If Val(GetPara("���￨������֤", "1")) = 1 And strPassWord <> "" Then
                        If frmCardPass.ShowCardPass(strPassWord) = False Then
                            Exit Sub
                        End If
                    Else
                        If Not CheckQueryPass(lngPatientKey) Then Exit Sub
                    End If
                    If frmSelect.ShowSelect(gRs, mvar����id, mvar��ҳid) Then
                        Call EnterCharge
                    End If
                Else
                    MsgBox "�����ҵĲ��˲����ڣ����������룡", vbInformation, gstrSysName
                End If
            Else
                MsgBox "�����ҵĲ��˲����ڣ����������룡", vbInformation, gstrSysName
            End If
            
        '--------------------------------------------------------------------------------------------------------------
        Case 20         '��Ʊ�ݺ�
            
            mstrNO = UCase(Trim(txt(4).Text))
            If mstrNO <> "" Then
                
                gstrSQL = " Where ��������=1 And ID=(Select ��ӡID From " & zlGetFullFieldsTable("Ʊ��ʹ����ϸ") & " Where Ʊ��=1 And ����=1 And ����=[1] And Rownum=1)"
                gstrSQL = "Select NO From Ʊ�ݴ�ӡ���� " & gstrSQL & " Union All Select NO From HƱ�ݴ�ӡ���� " & gstrSQL
                gstrSQL = "Select ����id, 0 as ��ҳid,���� From ������ü�¼ Where ��¼״̬=1 and �����־=1 and ��¼����=1 and No In (" & gstrSQL & ")"
                
                Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "���ò�ѯ", mstrNO)
                If gRs.BOF = False Then
                    mvar����id = zlCommFun.Nvl(gRs("����id"), 0)
                    mvar��ҳid = zlCommFun.Nvl(gRs("��ҳid"), 0)
                    mvar���� = zlCommFun.Nvl(gRs("����"), "")
                    If CheckQueryPass(mvar����id) Then '����У��
                        Call EnterCharge
                    End If
                Else
                    MsgBox "������һ�������ڵ�����Ʊ�ݺţ����������룡", vbInformation, gstrSysName
                    EnterFocus txt(4)
                End If
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 21         '�����ݺ�
            mstrNO = UCase(Trim(txt(4).Text))
            If mstrNO <> "" Then
                gstrSQL = "Select ����id,0 as ��ҳid,���� From ������ü�¼ Where ��¼״̬=1 and ��¼����=1 and �����־=1 and No=[1]"
                Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "���ò�ѯ", mstrNO)
                If gRs.BOF = False Then
                    mvar����id = zlCommFun.Nvl(gRs("����id"), 0)
                    mvar��ҳid = zlCommFun.Nvl(gRs("��ҳid"), 0)
                    mvar���� = zlCommFun.Nvl(gRs("����"), "")
                    If CheckQueryPass(mvar����id) Then '����У��
                        Call EnterCharge
                    End If
                Else
                    MsgBox "������һ�������ڵ����ﵥ�ݺţ����������룡", vbInformation, gstrSysName
                    EnterFocus txt(4)
                End If
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 25             '���֤��
            gstrSQL = "" & _
            "   Select to_char(A.��Ժ����,'yyyy-mm-dd') as ��Ժ����,to_char(A.��Ժ����,'yyyy-mm-dd') as ��Ժ����,A.����id,A.��ҳid,B.����֤��" & _
            "   From ������ҳ A,������Ϣ B " & _
            "   where A.����id=B.����id and B.���֤��=[1] " & _
            "   Union ALL " & _
            "   Select '�������' as ��Ժ����,'�������' as ��Ժ����,E.����id,0 as ��ҳid,E.����֤�� " & _
            "   From ������ü�¼ C,������Ϣ E " & _
            "   Where C.����id=E.����id And E.���֤��=[1] and Rownum<2 "
            
            gstrSQL = " Select AA.��Ժ����,AA.��Ժ����,AA.����id,AA.��ҳid,AA.����֤�� From (" & gstrSQL & ") AA  order by AA.��Ժ���� asc,AA.��ҳid desc "
            gstrSQL = "" & _
            "   SELECT rownum as No,D.��Ժ����,D.��Ժ����,D.����id,D.��ҳid,D.����֤�� " & _
            "   FROM ( " & gstrSQL & ") D"
'
            Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "���ò�ѯ", UCase(txt(4).Text))
            If gRs.RecordCount > 0 Then
                txt(4).Text = ""
                If Not CheckQueryPass(Nvl(gRs!����id, 0)) Then Exit Sub '����У��
                If frmSelect.ShowSelect(gRs, mvar����id, mvar��ҳid) Then
                    Call EnterCharge
                End If
            Else
                MsgBox "��������֤�ţ��밴[ȷ��]������ˢ����", vbInformation, gstrSysName
                txt(4).Text = ""
                EnterFocus txt(4)
            End If
            
        '--------------------------------------------------------------------------------------------------------------
        Case 26             '�ɣÿ���
            
            gstrSQL = "" & _
            "   Select to_char(A.��Ժ����,'yyyy-mm-dd') as ��Ժ����,to_char(A.��Ժ����,'yyyy-mm-dd') as ��Ժ����,A.����id,A.��ҳid,B.����֤�� " & _
            "   From ������ҳ A,������Ϣ B " & _
            "   Where A.����id=B.����id  and B.IC����=[1]" & _
            "   Union ALL " & _
            "   Select  '�������' as ��Ժ����,'�������' as ��Ժ����,E.����id,0 as ��ҳid,E.����֤�� " & _
            "   From ������ü�¼ C,������Ϣ E" & _
            "   Where  C.����id=E.����id And E.IC����=[1] and rownum<2   "
            
            gstrSQL = "Select ��Ժ����,��Ժ����,����id,��ҳid,����֤�� From (" & gstrSQL & ") AA  order by AA.��Ժ���� asc,AA.��ҳid desc"
            gstrSQL = "SELECT RowNum as No,D.��Ժ����,D.��Ժ����,D.����id,D.��ҳid,D.����֤�� FROM (" & gstrSQL & ") D"
            Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "���ò�ѯ", UCase(txt(4).Text))
            If gRs.RecordCount > 0 Then
                txt(4).Text = ""
                strPassWord = zlCommFun.Nvl(gRs("����֤��").Value)
                
                If Val(GetPara("���￨������֤", "1")) = 1 And strPassWord <> "" Then
                    If frmCardPass.ShowCardPass(strPassWord) = False Then
                        Exit Sub
                    End If
                Else
                    If Not CheckQueryPass(Nvl(gRs!����id, 0)) Then Exit Sub
                End If

                If frmSelect.ShowSelect(gRs, mvar����id, mvar��ҳid) Then
                    Call EnterCharge
                End If
            Else
                MsgBox "�����IC���ţ��밴[ȷ��]������ˢ����", vbInformation, gstrSysName
                txt(4).Text = ""
                EnterFocus txt(4)
            End If

        End Select
        
        txt(4).Text = ""
        
    Case "���"
        
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

Private Sub Ƿ�����(str���� As String, lng����ID As Long, lng��ҳid As Long, ByVal sngʣ���� As Single, ByVal str�������� As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    Dim strError As String
    Dim sng����ֵ As Single
    Dim sng������ As Single
    
    '�Ȼָ� lblWarn �ĳ�ʼ״̬
    lblWarn.BackStyle = 0
    lblWarn.Caption = ""

    gstrSQL = "Select ��������,����ֵ From ���ʱ����� A,������ҳ B Where A.���ò���=[3] And A.����ID = B.��ǰ����ID And B.����id =[1] And B.��ҳid = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng����ID, lng��ҳid, str��������)
    If rsTmp.BOF Then Exit Sub
    
    sng����ֵ = IIf(IsNull(rsTmp!����ֵ), 0, rsTmp!����ֵ)
    
    gstrSQL = "Select ������ From ������Ϣ A Where A.����ID =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng����ID)
    If Not rsTmp.BOF Then sng������ = zlCommFun.Nvl(rsTmp!������, 0)
    
    sngʣ���� = sngʣ���� + sng������
    
    If sngʣ���� < 0 Then
        strError = str���� & "��Ƿ��" & Abs(FormatEx(sngʣ����, 2)) & "Ԫ���뾡��ɷѣ�"
        lblWarn.BackStyle = 1
        lblWarn.BackColor = &HFFFF&
        lblWarn.Caption = "Ƿ��"
        
    ElseIf sngʣ���� < sng����ֵ Then
        strError = str���� & "ʣ����ܶ�(" & FormatEx(sngʣ����, 2) & ")�Ե��ڱ���ֵ(" & FormatEx(sng����ֵ, 2) & ")��"
        lblWarn.BackStyle = 1
        lblWarn.BackColor = &HFFFF&
        lblWarn.Caption = "����"
    End If

    
    Call frmShowMessage.ShowMe(Me, strError)
    
    
End Sub

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
    '******************************************************************************************************************
    '���ܣ��������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
    '������vNumber=Single,Double,Currency���͵�����,intBit=���С��λ��
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
