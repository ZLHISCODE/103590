VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageCourse 
   AutoRedraw      =   -1  'True
   Caption         =   "病人入出管理"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9135
   Icon            =   "frmManageCourse.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picCard_s 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
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
      Height          =   4140
      Left            =   5850
      MouseIcon       =   "frmManageCourse.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "frmManageCourse.frx":045C
      ScaleHeight     =   4110
      ScaleWidth      =   2805
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1700
      Width           =   2835
      Begin VB.Label lbl病人类型 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
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
         Left            =   795
         TabIndex        =   41
         Top             =   1725
         Width           =   105
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "类型"
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
         Left            =   240
         TabIndex        =   40
         Top             =   1725
         Width           =   420
      End
      Begin VB.Line Line30 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   5030
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line29 
         BorderColor     =   &H80000015&
         X1              =   0
         X2              =   5045
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "诊断"
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
         Left            =   255
         TabIndex        =   35
         Top             =   3870
         Width           =   420
      End
      Begin VB.Label lbl诊断 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
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
         Left            =   795
         TabIndex        =   34
         Top             =   3870
         Width           =   105
      End
      Begin VB.Line Line28 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   5045
         Y1              =   3825
         Y2              =   3825
      End
      Begin VB.Line Line27 
         BorderColor     =   &H80000015&
         X1              =   0
         X2              =   5060
         Y1              =   3810
         Y2              =   3810
      End
      Begin VB.Label lbl医疗付款方式 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
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
         Left            =   795
         TabIndex        =   33
         Top             =   3510
         Width           =   105
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "付款"
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
         Left            =   255
         TabIndex        =   32
         Top             =   3510
         Width           =   420
      End
      Begin VB.Line Line26 
         BorderColor     =   &H80000015&
         X1              =   0
         X2              =   5060
         Y1              =   3435
         Y2              =   3435
      End
      Begin VB.Line Line25 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   5045
         Y1              =   3450
         Y2              =   3450
      End
      Begin VB.Label lbl医保号 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   765
         TabIndex        =   29
         Top             =   1080
         Width           =   105
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "医保号"
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
         Left            =   45
         TabIndex        =   28
         Top             =   1080
         Width           =   630
      End
      Begin VB.Line Line24 
         BorderColor     =   &H80000015&
         X1              =   -45
         X2              =   5045
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Line Line23 
         BorderColor     =   &H80000014&
         X1              =   -30
         X2              =   5045
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Label lblLevel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
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
         Left            =   795
         TabIndex        =   26
         Top             =   390
         Width           =   105
      End
      Begin VB.Line Line22 
         BorderColor     =   &H80000015&
         X1              =   690
         X2              =   690
         Y1              =   330
         Y2              =   645
      End
      Begin VB.Line Line21 
         BorderColor     =   &H80000014&
         X1              =   705
         X2              =   705
         Y1              =   330
         Y2              =   660
      End
      Begin VB.Line Line20 
         BorderColor     =   &H80000015&
         X1              =   0
         X2              =   5000
         Y1              =   315
         Y2              =   315
      End
      Begin VB.Line Line19 
         BorderColor     =   &H80000014&
         X1              =   15
         X2              =   5000
         Y1              =   330
         Y2              =   330
      End
      Begin VB.Line Line18 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   5000
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Line Line17 
         BorderColor     =   &H80000014&
         X1              =   15
         X2              =   5000
         Y1              =   1005
         Y2              =   1005
      End
      Begin VB.Line Line16 
         BorderColor     =   &H80000014&
         X1              =   -75
         X2              =   5000
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line15 
         BorderColor     =   &H80000014&
         X1              =   -30
         X2              =   5000
         Y1              =   2355
         Y2              =   2355
      End
      Begin VB.Line Line14 
         BorderColor     =   &H80000014&
         X1              =   -75
         X2              =   5000
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Line Line13 
         BorderColor     =   &H80000014&
         X1              =   -45
         X2              =   5000
         Y1              =   3060
         Y2              =   3060
      End
      Begin VB.Line Line12 
         BorderColor     =   &H80000014&
         X1              =   1440
         X2              =   1440
         Y1              =   660
         Y2              =   1005
      End
      Begin VB.Line Line11 
         BorderColor     =   &H80000014&
         X1              =   1980
         X2              =   1980
         Y1              =   660
         Y2              =   1005
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000014&
         X1              =   705
         X2              =   705
         Y1              =   1005
         Y2              =   4100
      End
      Begin VB.Line Line9 
         BorderColor     =   &H80000015&
         X1              =   690
         X2              =   690
         Y1              =   990
         Y2              =   4100
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000015&
         X1              =   1965
         X2              =   1965
         Y1              =   645
         Y2              =   990
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000015&
         X1              =   1425
         X2              =   1425
         Y1              =   645
         Y2              =   990
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000015&
         X1              =   -60
         X2              =   5000
         Y1              =   3045
         Y2              =   3045
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000015&
         X1              =   -90
         X2              =   5000
         Y1              =   2685
         Y2              =   2685
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000015&
         X1              =   -45
         X2              =   5000
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000015&
         X1              =   -90
         X2              =   5000
         Y1              =   1665
         Y2              =   1665
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000015&
         X1              =   0
         X2              =   5000
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000015&
         X1              =   -15
         X2              =   5000
         Y1              =   645
         Y2              =   645
      End
      Begin VB.Label lbl医生 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
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
         Left            =   795
         TabIndex        =   25
         Top             =   3135
         Width           =   105
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "医生"
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
         Left            =   255
         TabIndex        =   24
         Top             =   3135
         Width           =   420
      End
      Begin VB.Label lbl护理等级 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
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
         Left            =   795
         TabIndex        =   23
         Top             =   2775
         Width           =   105
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "护理"
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
         Left            =   240
         TabIndex        =   22
         Top             =   2775
         Width           =   420
      End
      Begin VB.Label lbl入院时间 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
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
         Left            =   795
         TabIndex        =   21
         Top             =   2415
         Width           =   105
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "入院"
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
         Left            =   240
         TabIndex        =   20
         Top             =   2415
         Width           =   420
      End
      Begin VB.Label lbl病况 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
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
         Left            =   795
         TabIndex        =   19
         Top             =   2085
         Width           =   105
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "病况"
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
         Left            =   240
         TabIndex        =   18
         Top             =   2055
         Width           =   420
      End
      Begin VB.Label lbl住院号 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
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
         Left            =   795
         TabIndex        =   17
         Top             =   1395
         Width           =   105
      End
      Begin VB.Label lbl标识 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
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
         Left            =   30
         TabIndex        =   16
         Top             =   1395
         Width           =   630
      End
      Begin VB.Label lbl年龄 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
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
         Left            =   2055
         TabIndex        =   15
         Top             =   735
         Width           =   420
      End
      Begin VB.Label lbl性别 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
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
         Left            =   1485
         TabIndex        =   14
         Top             =   735
         Width           =   420
      End
      Begin VB.Label lbl姓名 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
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
         Left            =   60
         TabIndex        =   13
         Top             =   735
         Width           =   1275
      End
      Begin VB.Label lbl床号 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "床号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   450
         TabIndex        =   12
         Top             =   60
         Width           =   570
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "等级"
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
         Left            =   240
         TabIndex        =   27
         Top             =   390
         Width           =   420
      End
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   3240
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   34
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1A4C
            Key             =   "M"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":2326
            Key             =   "M_Change"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":2C00
            Key             =   "KM"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":34DA
            Key             =   "KM_Change"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":3DB4
            Key             =   "F"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":468E
            Key             =   "F_Change"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":4F68
            Key             =   "KF"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":5842
            Key             =   "KF_Change"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":611C
            Key             =   "O"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":69F6
            Key             =   "O_Change"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":72D0
            Key             =   "KO"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":7BAA
            Key             =   "K0_Change"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":8484
            Key             =   "M_Empty"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":879E
            Key             =   "F_Empty"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":8AB8
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":8DD2
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":90EC
            Key             =   "Holding"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":9406
            Key             =   "KHolding"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":9CE0
            Key             =   "Change"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":9FFA
            Key             =   "KChange"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":A8D4
            Key             =   "Out"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":B1AE
            Key             =   "KOut"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":BA88
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":BDA2
            Key             =   "KFamily"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":C67C
            Key             =   "Limit"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":D4C6
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":D62C
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":D792
            Key             =   "MASK_加床"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":DAAC
            Key             =   "MASK_非编"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":DDC6
            Key             =   "MASK_共用"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":E0E0
            Key             =   "MASK_共用_加床"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":E3FA
            Key             =   "MASK_共用_非编"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":E714
            Key             =   "U"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":F566
            Key             =   "KU"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   3240
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":103B8
            Key             =   "M"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":10C92
            Key             =   "M_Change"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1156C
            Key             =   "KM"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":11E46
            Key             =   "KM_Change"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":12720
            Key             =   "F"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":12FFA
            Key             =   "F_Change"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":138D4
            Key             =   "KF"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":141AE
            Key             =   "KF_Change"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":14A88
            Key             =   "O"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":15362
            Key             =   "O_Change"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":15C3C
            Key             =   "KO"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":16516
            Key             =   "KO_Change"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":16DF0
            Key             =   "M_Empty"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1710A
            Key             =   "F_Empty"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":17424
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1773E
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":17A58
            Key             =   "Holding"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":17D72
            Key             =   "KHolding"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1864C
            Key             =   "Change"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":18966
            Key             =   "KChange"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":19240
            Key             =   "Out"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":19B1A
            Key             =   "KOut"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1A3F4
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1A70E
            Key             =   "KFamily"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1AFE8
            Key             =   "MASK_加床"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1B142
            Key             =   "MASK_非编"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1B29C
            Key             =   "MASK_共用"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1B3F6
            Key             =   "MASK_共用_加床"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1B550
            Key             =   "MASK_共用_非编"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1B6AA
            Key             =   "U"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1C4FC
            Key             =   "KU"
         EndProperty
      EndProperty
   End
   Begin VB.Timer timSize 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8445
      Top             =   5775
   End
   Begin MSComctlLib.Toolbar tbrFilter 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   30
      Top             =   780
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   609
      ButtonWidth     =   1984
      ButtonHeight    =   609
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgFilter"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "当天入院"
            Key             =   "curDay"
            Object.ToolTipText     =   "只显示当天入院的病人(F7)"
            Object.Tag             =   "当天入院"
            ImageKey        =   "UnCheck_"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9135
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   7635
      NewRow1         =   0   'False
      BandForeColor2  =   8388608
      Caption2        =   "病区"
      Child2          =   "cboUnit"
      MinWidth2       =   2205
      MinHeight2      =   300
      Width2          =   1215
      NewRow2         =   0   'False
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   6840
         TabIndex        =   5
         Text            =   "cboUnit"
         Top             =   240
         Width           =   2205
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   8
         Top             =   30
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Description     =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "入住"
               Key             =   "In"
               Description     =   "入住"
               Object.ToolTipText     =   "入住"
               Object.Tag             =   "入住"
               ImageKey        =   "In"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "转科"
               Key             =   "Change"
               Description     =   "转科"
               Object.ToolTipText     =   "转科"
               Object.Tag             =   "转科"
               ImageKey        =   "Change"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "换床"
               Key             =   "Move"
               Description     =   "换床"
               Object.ToolTipText     =   "换床"
               Object.Tag             =   "换床"
               ImageKey        =   "Move"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "出院"
               Key             =   "Out"
               Description     =   "出院"
               Object.ToolTipText     =   "出院"
               Object.Tag             =   "出院"
               ImageKey        =   "Out"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "调整"
               Key             =   "Adjust"
               Description     =   "调整"
               Object.ToolTipText     =   "调整病人的身份或在院信息"
               Object.Tag             =   "调整"
               ImageKey        =   "Adjust"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Adjust_"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "撤消"
               Key             =   "Undo"
               Description     =   "撤消"
               Object.ToolTipText     =   "撤消"
               Object.Tag             =   "撤消"
               ImageKey        =   "Undo"
               Style           =   5
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "列表"
               Key             =   "View"
               Description     =   "列表"
               Object.ToolTipText     =   "床位列表显示方式"
               Object.Tag             =   "列表"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Icon"
                     Object.Tag             =   "大图标(&G)"
                     Text            =   "大图标(&G)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Object.Tag             =   "小图标(&M)"
                     Text            =   "小图标(&M)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Object.Tag             =   "列表(&L)"
                     Text            =   "列表(&L)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Object.Tag             =   "详细资料(&D)"
                     Text            =   "详细资料(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   6195
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageCourse.frx":1D34E
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9419
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "病人颜色"
            TextSave        =   "病人颜色"
            Key             =   "PatiColor"
            Object.Tag             =   "PatiColor"
            Object.ToolTipText     =   "病人颜色说明"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
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
   Begin VB.PictureBox picVsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4905
      Left            =   5580
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4905
      ScaleWidth      =   45
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1125
      Width           =   45
   End
   Begin MSComctlLib.ListView lvwOut_s 
      Height          =   2415
      Left            =   5670
      TabIndex        =   3
      Tag             =   "可变化的"
      Top             =   3735
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   4260
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwFamily_s 
      Height          =   2190
      Left            =   5655
      TabIndex        =   1
      Tag             =   "可变化的"
      Top             =   1275
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   3863
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwIn_s 
      Height          =   1410
      Left            =   75
      TabIndex        =   2
      Tag             =   "可变化的"
      Top             =   4740
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   2487
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwBeds_s 
      Height          =   3210
      Left            =   60
      TabIndex        =   0
      Tag             =   "可变化的"
      Top             =   1275
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   5662
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   60
      Top             =   135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1DBE2
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1DDFC
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1E016
            Key             =   "In"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1E230
            Key             =   "Change"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1E44A
            Key             =   "Move"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1E664
            Key             =   "Out"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1E87E
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1EA98
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1ECB2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1EECC
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1F0E6
            Key             =   "Adjust"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   645
      Top             =   135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1F300
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1F51A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1F734
            Key             =   "In"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1F94E
            Key             =   "Change"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1FB68
            Key             =   "Move"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1FD82
            Key             =   "Out"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1FF9C
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":201B6
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":203D0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":205EA
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":20804
            Key             =   "Adjust"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgFilter 
      Left            =   855
      Top             =   1590
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":20A1E
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":20B78
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":20CD2
            Key             =   "UnCheck_"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":20E2C
            Key             =   "Check_"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicOut 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   5655
      MousePointer    =   7  'Size N S
      ScaleHeight     =   225
      ScaleWidth      =   3450
      TabIndex        =   36
      Top             =   3495
      Width           =   3450
      Begin VB.CheckBox chk结清 
         BackColor       =   &H00808080&
         Caption         =   "已结清"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   1350
         TabIndex        =   38
         Top             =   20
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chk结清 
         BackColor       =   &H00808080&
         Caption         =   "未结清"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   2355
         TabIndex        =   37
         Top             =   20
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.Label lblOut 
         BackColor       =   &H00808080&
         Caption         =   " 出院病人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   190
         Left            =   0
         TabIndex        =   39
         Top             =   20
         Width           =   945
      End
   End
   Begin VB.Label lblBed 
      BackColor       =   &H00808080&
      Caption         =   " 病区病床"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   75
      TabIndex        =   31
      Top             =   1035
      Width           =   5475
   End
   Begin VB.Label lblIn 
      BackColor       =   &H00808080&
      Caption         =   " 待入住病人"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   75
      MousePointer    =   7  'Size N S
      TabIndex        =   11
      Top             =   4515
      Width           =   5460
   End
   Begin VB.Label lblFamily 
      BackColor       =   &H00808080&
      Caption         =   " 家庭病床"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5655
      TabIndex        =   10
      Top             =   1035
      Width           =   3450
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFile_PreView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintMed 
         Caption         =   "打印病案(&M)"
      End
      Begin VB.Menu mnuFilePrintCard 
         Caption         =   "打印床头卡(&C)"
      End
      Begin VB.Menu mnuFile_PrintWristlet 
         Caption         =   "打印腕带(&W)"
      End
      Begin VB.Menu mnuFile_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit_In 
         Caption         =   "入住(&I)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEdit_Change 
         Caption         =   "转科(&C)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEdit_ChangeUnit 
         Caption         =   "转病区(&T)"
      End
      Begin VB.Menu mnuEdit_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_ChangeGroup 
         Caption         =   "转医疗小组(&G)"
      End
      Begin VB.Menu mnuEdit_6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Move 
         Caption         =   "换床(&M)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEdit_Swap 
         Caption         =   "床位对换(&S)"
      End
      Begin VB.Menu mnuEdit_AddBeds 
         Caption         =   "包房(&B)"
      End
      Begin VB.Menu mnuEdit_7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Out 
         Caption         =   "出院(&O)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuEdit_PreOut 
         Caption         =   "预出院(&P)"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuEdit_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_ModifOut 
         Caption         =   "修改出院时间(&E)"
      End
      Begin VB.Menu mnuEdit_OutAndModi 
         Caption         =   "出院及调整出院(&J)"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditToInPati 
         Caption         =   "转为住院病人(&K)"
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Level 
         Caption         =   "更改床位等级(&B)"
      End
      Begin VB.Menu mnuEdit_Nurse 
         Caption         =   "更改护理等级(&N)"
      End
      Begin VB.Menu mnuEdit_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Adjust 
         Caption         =   "调整住院信息(&F)"
      End
      Begin VB.Menu mnuEdit_BabyReg 
         Caption         =   "新生儿登记(&Y)"
      End
      Begin VB.Menu mnuEdit_Memo 
         Caption         =   "病人备注信息(&Z)"
      End
      Begin VB.Menu mnuEdit_Recalc 
         Caption         =   "按费别重算费用(&R)"
      End
      Begin VB.Menu mnuEdit_Disease 
         Caption         =   "医保病种选择(&D)"
      End
      Begin VB.Menu mnuEdit_Adjust_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Undo 
         Caption         =   "撤消(&U)"
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "查询(&Q)"
      Begin VB.Menu mnuQuery_Log 
         Caption         =   "病人变动记录(&C)"
      End
      Begin VB.Menu mnuQuery_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQueryInfo 
         Caption         =   "病人信息(&I)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "报表(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolUnit 
            Caption         =   "病区选择(&U)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_Tlb_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_Card 
         Caption         =   "床位卡(&C)"
         Checked         =   -1  'True
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuView_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewColSel 
         Caption         =   "选择列(&C)"
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_ListView 
         Caption         =   "大图标(&G)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuView_ListView 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuView_ListView 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuView_ListView 
         Caption         =   "详细资料(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewFindNext 
         Caption         =   "查找下一个(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuViewreFlash 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmManageCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'常量
Private Const COLOR_FOCUS = &H966334   '&HC0844E
Private Const COLOR_LOST = &H808080   '&H966334
Private Const COL_BEDS = "病床,1170,0,1;姓名,959,0,1;性别,650,2,0;年龄,585,2,0;床位等级,975,0,2;" & "科室,959,0,2;病人ID,750,0,0;住院号,799,0,2;当前病况,929,2,0;入院时间,1620,2,2;" & "护理等级,1000,0,2;住院医师,1000,0,0;房间号,799,2,0;性别分类,1000,2,0;床位编制,1000,2,0;床号,0,0,1;就诊卡号,0,0,1;身份证号,0,0,1;IC卡号,0,0,1;病人类型,1000,0,2"
Private Const COL_FAMILY = "姓名,1000,0,1;性别,650,2,0;年龄,650,2,0;病人ID,799,0,0;" & "住院号,799,0,2;当前科室,1000,0,2;当前病况,1000,2,0;入院时间,1635,2,2;护理等级,1000,2,2;住院医师,1000,0,0;就诊卡号,0,0,1;身份证号,0,0,1;IC卡号,0,0,1;病人类型,1000,0,2"
Private Const COL_IN = "姓名,1000,0,1;性别,555,2,0;年龄,650,2,0;病人ID,799,0,0;" & "住院号,799,0,2;费别,799,2,0;当前病区,1000,0,0;当前科室,1000,0,2;转入科室,1000,0,2;" & "入院时间,1635,2,2;当前病况,615,0,0;护理等级,1440,0,2;就诊卡号,0,0,1;身份证号,0,0,1;IC卡号,0,0,1;病人类型,1000,0,2"
Private Const COL_OUT = "姓名,959,0,1;性别,650,2,0;年龄,650,2,0;病人ID,799,0,0;" & "住院号,799,0,2;出院方式,1000,2,0;出院时间,1665,2,2;入院时间,1635,2,0;出院科室,1000,0,2;" & "出院病床,929,0,2;出院病况,929,2,0;护理等级,1000,2,0;费别,650,2,0;就诊卡号,0,0,1;身份证号,0,0,1;IC卡号,0,0,1;病人类型,1000,0,2;就诊次数,0,0,1"

Private mblnUnload As Boolean
Private mlngPreX As Long, mlngPreY As Long
Private mblnMax As Boolean
Private mblnDropIn As Boolean, mblnDropOut As Boolean '调整列表尺寸后后定位
Private mblnDownIn As Boolean, mblnDownOut As Boolean, mblnDownVsc As Boolean '大小调整
Private mblnBeds As Boolean, mblnFamily As Boolean, mblnIn As Boolean, mblnOut As Boolean '项目点击
Private mlngUnit As Long
Public mstrPrivs As String
Private mlngModul As Long
'统计数据
Private mintBeds_A As Integer, mintChange_A As Integer, mintHolding As Integer
Private mintBeds_B As Integer, mintChange_B As Integer
Private mintIn As Integer, mintChange_C As Integer
Private mintOut As Long
'数据对象:与列表项关联
Public mobjLVW As ListView '当前活动列表
Public mrsBeds As ADODB.Recordset '床位映象表(床位、病人)
Public mrsFamily As ADODB.Recordset '家庭病床病人
Public mrsIn As ADODB.Recordset '入科病人(含入院登记和转科病人)
Public mrsOut As ADODB.Recordset '出院病人
'数据克隆,用于辅助操作
Public mrsCBeds As ADODB.Recordset
Public mrsCFamily As ADODB.Recordset
Public mrsCIn As ADODB.Recordset
Public mrsCOut As ADODB.Recordset
'定位方式
Private mstrSeekKey As String, mstrSeekValue As String
'病人类型及颜色
Private mstrPatiTypeColor As String
Private mstrDeptName As String
Private mlng当前病区id As Long
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub cboUnit_Click()
    If cboUnit.ItemData(cboUnit.ListIndex) = mlngUnit Then Exit Sub
    mlngUnit = cboUnit.ItemData(cboUnit.ListIndex)
    Call LoadList(True, True, True, True, True)
End Sub
'问题28811 by lesfeng 2010-03-30
Private Sub cboUnit_GotFocus()
    With cboUnit
        .SelStart = 0
        .SelLength = Len(.Text)
        mstrDeptName = .Text
    End With
End Sub
'问题28811 by lesfeng 2010-03-30
Private Sub cboUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String
    Dim strSQL As String, intIdx As Long, i As Long
    Dim lngUnit As Long
    
    If KeyCode = vbKeyReturn Then
        If Trim(cboUnit.Text) = "" Then
            If cboUnit.Enabled Then cboUnit.SetFocus
            Exit Sub
        End If
         Set rsTmp = InputGetDept(cboUnit, blnCancel)
        If Not rsTmp Is Nothing Then
            intIdx = cbo.FindIndex(cboUnit, rsTmp!ID)
            If intIdx <> -1 Then
                cboUnit.ListIndex = intIdx
            End If
        Else
            If cboUnit.ListIndex = -1 And cboUnit.ListCount = 0 Then
            Else
                If Not blnCancel Then
                    MsgBox "未找到对应的病区。", vbInformation, gstrSysName
                    cboUnit.SetFocus
                    cboUnit.SelStart = 0
                    cboUnit.Text = mstrDeptName
                    cboUnit.SelLength = Len(cboUnit.Text)
                End If
            End If
        End If
    End If
End Sub
'问题28811 by lesfeng 2010-03-30
Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub
Private Sub chk结清_Click(Index As Integer)
    If chk结清(0).Value = 0 And chk结清(1).Value = 0 Then
        chk结清((Index + 1) Mod 2).Value = 1
    End If
    Call picOut_Click
    LoadList False, False, False, True
End Sub

Private Sub Form_Activate()
    If mblnUnload Then Unload Me: Exit Sub
    mnuView_Card.Checked = picCard_s.Visible
    If mobjLVW Is Nothing Then
        lvwBeds_s.SetFocus
    Else
        If mobjLVW.Visible And mobjLVW.Enabled Then mobjLVW.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long, k As Long
    
    If KeyCode = vbKeyF7 Then
        Call tbrFilter_ButtonClick(tbrFilter.Buttons("curDay"))
    ElseIf Shift = vbAltMask And InStr("0123456789", Chr(KeyCode)) > 0 Then
        j = IIf(KeyCode = vbKey0, 10, Val(Chr(KeyCode)))
        For i = 1 To tbrFilter.Buttons.Count
            If tbrFilter.Buttons(i).Key Like "Nurse*" Then
                k = k + 1
                If k = j Then
                    Call tbrFilter_ButtonClick(tbrFilter.Buttons(i))
                    Exit For
                End If
            End If
        Next
    End If
End Sub

Private Sub Form_Load()
    Dim X As Long, Y As Long, i As Integer
    Dim strLoc As String, blnCard As Boolean
    
    RestoreWinState Me, App.ProductName
    picCard_s.width = 2835
    picCard_s.Height = 3865
    
    If lvwBeds_s.ColumnHeaders.Count = 0 Then Call zlControl.LvwSelectColumns(lvwBeds_s, COL_BEDS, True)
    If lvwFamily_s.ColumnHeaders.Count = 0 Then Call zlControl.LvwSelectColumns(lvwFamily_s, COL_FAMILY, True)
    If lvwIn_s.ColumnHeaders.Count = 0 Then Call zlControl.LvwSelectColumns(lvwIn_s, COL_IN, True)
    If lvwOut_s.ColumnHeaders.Count = 0 Then Call zlControl.LvwSelectColumns(lvwOut_s, COL_OUT, True)
    
    '处理床位卡背景
    For Y = 0 To picCard_s.Height / Screen.TwipsPerPixelY Step 87
        For X = 0 To picCard_s.width / Screen.TwipsPerPixelX Step 50
            BitBlt picCard_s.hDC, X, Y, 50, 87, picCard_s.hDC, 0, 0, SRCCOPY
        Next
    Next
    zlControl.PicShowFlat picCard_s, 1
    
    '处理特殊图标
    Call MakeBedIcon
    
    mblnUnload = False
    mblnBeds = False: mblnFamily = False: mblnIn = False: mblnOut = False
    mlngUnit = 0
   
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    '权限设置
    If InStr(mstrPrivs, "病人出院") = 0 Then
        mnuEdit_Out.Visible = False
        tbr.Buttons("Out").Visible = False
    End If
    If InStr(mstrPrivs, "病人转科") = 0 Then
        mnuEdit_Change.Visible = False
        tbr.Buttons("Change").Visible = False
    End If
    '换床权限控制
    If InStr(mstrPrivs, "换床") = 0 Then
        mnuEdit_Move.Visible = False
        mnuEdit_Swap.Visible = False
        mnuEdit_AddBeds.Visible = False
        tbr.Buttons("Move").Visible = False
        mnuEdit_7.Visible = False
    End If
    
    If InStr(mstrPrivs, "转病区") = 0 Then
        mnuEdit_ChangeUnit.Visible = False
        tbr.Buttons("Change").Visible = mnuEdit_Change.Visible
    End If
    
    If InStr(mstrPrivs, "办理预出院") = 0 Then
        mnuEdit_PreOut.Visible = False
    End If
    If InStr(mstrPrivs, "调整病人信息") = 0 Then
        mnuEdit_Adjust.Visible = False
        'mnuEdit_Adjust_.Visible = False
        tbr.Buttons("Adjust").Visible = False
        'tbr.Buttons("Adjust_").Visible = False
    End If
    If InStr(mstrPrivs, "新生儿登记") = 0 Then
        mnuEdit_BabyReg.Visible = False
    End If
    If InStr(mstrPrivs, "重算费用") = 0 Then
        mnuEdit_Recalc.Visible = False
    End If
    '问题27392 by lesfeng 2010-01-14
    If InStr(mstrPrivs, "调整出院时间") = 0 Then
        mnuEdit_ModifOut.Visible = False
    End If
    '问题27866 by lesfeng 2010-02-05
    If (InStr(mstrPrivs, "病人出院") = 0 Or InStr(mstrPrivs, "调整出院时间") = 0) Then
        mnuEdit_OutAndModi.Visible = False
    End If
    If Not (mnuEdit_ModifOut.Visible Or mnuEdit_OutAndModi.Visible) Then
        mnuEdit_4.Visible = False
    End If
    
    If InStr(mstrPrivs, "调整床位等级") = 0 Then
        mnuEdit_Level.Visible = False
    End If
    
'    If InStr(mstrPrivs, "病人备注编辑") = 0 Then
'        mnuEdit_Memo.Visible = False
'    End If

    If InStr(mstrPrivs, "调整护理等级") = 0 Then
        mnuEdit_Nurse.Visible = False
    End If
    
    If InStr(mstrPrivs, "住院留观转住院") = 0 Then
        mnuEditToInPati.Visible = False
        mnuEdit_2.Visible = False
    End If
                
    Call InitPatiType
    
    If Val(zlDatabase.GetPara("当天入院", glngSys, mlngModul, 0)) = 0 Then
        tbrFilter.Buttons("curDay").Image = "UnCheck_"
    Else
        tbrFilter.Buttons("curDay").Image = "Check_"
    End If
    For i = 1 To tbrFilter.Buttons.Count
        If tbrFilter.Buttons(i).Key Like "Nurse*" Then
            If Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "护理等级" & Replace(tbrFilter.Buttons(i).Key, "Nurse", ""), 1)) <> 0 Then
                tbrFilter.Buttons(i).Image = "Check"
            Else
                tbrFilter.Buttons(i).Image = "UnCheck"
            End If
        End If
    Next
            
    '初始住院病区
    If Not InitUnits Then mblnUnload = True: Exit Sub
    
    '创建消息对象
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPrivs, Me.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '工具条占用高度
    Dim staH As Long '状态栏占用高度
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    If mblnMax Then
        lvwFamily_s.Height = Me.ScaleHeight / 3
        lvwIn_s.Height = Me.ScaleHeight / 4
        
        lvwFamily_s.width = Me.ScaleWidth * 0.35
        lvwOut_s.width = Me.ScaleWidth * 0.35
        mblnMax = False
    End If
    If Me.WindowState = 2 Then mblnMax = True
    
    '靠齐控件宽度和高度
    cbrH = IIf(cbr.Visible, cbr.Height, 0) + tbrFilter.Height
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    With lblBed
        .Left = Me.ScaleLeft
        .Top = Me.ScaleTop + cbrH + 15
        .width = Me.ScaleWidth - lvwFamily_s.width - picVsc.width
    End With
    With lvwBeds_s
        .Left = lblBed.Left
        .Top = lblBed.Top + lblBed.Height + 15
        .width = lblBed.width
        .Height = Me.ScaleHeight - lvwIn_s.Height - lblIn.Height - lblBed.Height - cbrH - staH - 60
    End With
    With lblIn
        .Top = lvwBeds_s.Top + lvwBeds_s.Height + 15
        .Left = lblBed.Left
        .width = lblBed.width
    End With
    With lvwIn_s
        .Top = lblIn.Top + lblIn.Height + 15
        .Left = lblIn.Left
        .width = lvwBeds_s.width
    End With
    With picVsc
        .Top = lblBed.Top
        .Left = lblBed.Left + lblBed.width
        .Height = Me.ScaleHeight - cbrH - staH
    End With
    With lblFamily
        .Top = lblBed.Top
        .Left = picVsc.Left + picVsc.width
        .width = lvwFamily_s.width
    End With
    With lvwFamily_s
        .Top = lblFamily.Top + lblFamily.Height + 15
        .Left = lblFamily.Left
    End With
    With PicOut
        .Left = lblFamily.Left
        .Top = lvwFamily_s.Top + lvwFamily_s.Height + 15
        .width = lvwFamily_s.width
    End With
    With lvwOut_s
        .Left = PicOut.Left
        .Top = PicOut.Top + PicOut.Height + 15
        .width = lvwFamily_s.width
        .Height = Me.ScaleHeight - lblFamily.Height - PicOut.Height - lvwFamily_s.Height - cbrH - staH - 60
    End With
    Me.Refresh
    
    If WindowState = 0 Or WindowState = 2 Then
        timSize.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    Set mrsBeds = Nothing
    Set mrsFamily = Nothing
    Set mrsIn = Nothing
    Set mrsOut = Nothing
    
    Set mrsCBeds = Nothing
    Set mrsCFamily = Nothing
    Set mrsCIn = Nothing
    Set mrsCOut = Nothing
    
    '卸载消息对象
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
    
    mstrSeekKey = "": mstrSeekValue = ""
    
    SaveWinState Me, App.ProductName
    
    zlDatabase.SetPara "当天入院", IIf(tbrFilter.Buttons("curDay").Image = "UnCheck_", 0, 1), glngSys, mlngModul
    For i = 1 To tbrFilter.Buttons.Count
        If tbrFilter.Buttons(i).Key Like "Nurse*" Then
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "护理等级" & Replace(tbrFilter.Buttons(i).Key, "Nurse", ""), IIf(tbrFilter.Buttons(i).Image = "Check", 1, 0)
        End If
    Next
End Sub

Private Sub lblBed_Click()
    lvwBeds_s.SetFocus
End Sub

Private Sub lblFamily_Click()
    lvwFamily_s.SetFocus
End Sub

Private Sub lblIn_Click()
    If mblnDropIn Then
        If lvwBeds_s.Height >= lvwIn_s.Height Then
            lvwBeds_s.SetFocus
        Else
            lvwIn_s.SetFocus
        End If
    Else
        lvwIn_s.SetFocus
    End If
End Sub

Private Sub lblIn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngPreY = Y: mblnDownIn = True: mblnDropIn = False
End Sub

Private Sub lblIn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And mblnDownIn Then
        If lvwIn_s.Height - (Y - mlngPreY) < 600 Or lvwBeds_s.Height + Y - mlngPreY < 600 Then Exit Sub
        lblIn.Top = lblIn.Top + Y - mlngPreY
        lvwBeds_s.Height = lvwBeds_s.Height + Y - mlngPreY
        lvwIn_s.Top = lvwIn_s.Top + Y - mlngPreY
        lvwIn_s.Height = lvwIn_s.Height - (Y - mlngPreY)
        Me.Refresh
        mblnDropIn = True
    End If
End Sub

Private Sub lblIn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnDownIn = False
End Sub

Private Sub lblOut_Click()
 Call picOut_Click
End Sub

Private Sub mclsMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
    Dim strValue As String, strDepts As String
    Dim lngInTime As Long, lngDept As Long, lngUnit As Long, strCurDate As String
    Dim lngPatID As Long, lngPageID As Long
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnExit As Boolean
    
    On Error GoTo ErrHand
    
    If UCase(strMsgItemIdentity) = "ZLHIS_PATIENT_001" Then
        If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
        strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
        strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
        If lngPatID = 0 Or lngPageID = 0 Then Exit Sub
        '检查病区
        If mclsXML.GetSingleNodeValue("in_dept_id", strValue, xsNumber) = False Then Exit Sub
        lngDept = Val(strValue)
        strValue = "": Call mclsXML.GetSingleNodeValue("in_area_id", strValue, xsNumber)
        If Val(strValue) = 0 Then
            strValue = ""
            strSQL = "Select 病区ID From 病区科室对应 where 科室ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取病区信息", lngDept)
            Do While Not rsTmp.EOF
                strValue = strValue & "," & rsTmp!病区ID
            rsTmp.MoveNext
            Loop
            strValue = Mid(strValue, 2)
        End If
        If InStr(1, "," & strValue & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
        
        '检查入院科室是否在待入科病人科室中
        strDepts = zlDatabase.GetPara("待入科病人科室", glngSys, mlngModul, "")
        If strDepts <> "" Then
            strDepts = "," & strDepts & ","
            If InStr(1, strDepts, "," & lngDept & ",") = 0 Then Exit Sub
        End If
        '检查入院时间是否在入院登记天数内
        strValue = "": Call mclsXML.GetSingleNodeValue("in_date", strValue, xsString)
        If IsDate(strValue) Then
            lngInTime = Val(zlDatabase.GetPara("入院天数", glngSys, mlngModul, 3))
            strCurDate = zlDatabase.Currentdate
            If lngInTime <> 0 Then
                strCurDate = Format(DateAdd("D", -1 * lngInTime, CDate(strCurDate)), "YYYY-MM-DD HH:mm:ss")
            Else
                strCurDate = Format(strCurDate, "YYYY-MM-DD")
            End If
            If Format(strValue, "YYYY-MM-DD HH:mm:ss") < Format(strCurDate, "YYYY-MM-DD HH:mm:ss") Then Exit Sub
        End If
        '提取病人信息
        strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsNumber)
        mclsXML.CloseXMLDocument
        mrsIn.Filter = "病人ID=" & lngPatID
        If mrsIn.EOF = True Then
            Call LoadList(False, False, True, False)
            If strValue <> "" Then
                Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新登记的病人:" & strValue, "待办入科提醒")
            End If
        End If
    ElseIf UCase(strMsgItemIdentity) = "ZLHIS_PATIENT_003" Then
        If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
        strValue = "": Call mclsXML.GetSingleNodeValue("send_program", strValue, xsString)
        If strValue <> "" And Val(strValue) = Me.hWnd Then Exit Sub
        strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
        strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
        If lngPatID = 0 Or lngPageID = 0 Then Exit Sub
        '检查病区
        strValue = "": Call mclsXML.GetSingleNodeValue("change_dept_id", strValue, xsNumber)
        lngDept = Val(strValue)
        strValue = "": Call mclsXML.GetSingleNodeValue("change_area_id", strValue, xsNumber)
        lngUnit = Val(strValue)
        
        If lngDept = 0 Then Exit Sub
        
        If lngUnit = 0 Then
            strValue = ""
            strSQL = "Select 病区ID From 病区科室对应 where 科室ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取病区信息", lngDept)
            Do While Not rsTmp.EOF
                strValue = strValue & "," & rsTmp!病区ID
            rsTmp.MoveNext
            Loop
            strValue = Mid(strValue, 2)
        Else
            strValue = lngUnit
        End If
        If InStr(1, "," & strValue & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
        
        '检查入院科室是否在待入科病人科室中
        strDepts = zlDatabase.GetPara("待入科病人科室", glngSys, mlngModul, "")
        If strDepts <> "" Then
            strDepts = "," & strDepts & ","
            If InStr(1, strDepts, "," & lngDept & ",") = 0 Then Exit Sub
        End If
        
        '提取病人信息
        strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsNumber)
        mclsXML.CloseXMLDocument
        mrsIn.Filter = "病人ID=" & lngPatID
        If mrsIn.EOF = True Then
            Call LoadList(True, True, True, False)
            If strValue <> "" Then
                Call mclsMipModule.ShowMessage(strMsgItemIdentity, "有新转入的病人:" & strValue, "待办入科提醒")
            End If
        End If
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mnuEdit_ChangeGroup_Click()
    Dim lng病人ID As Long, lng主页ID As Long
    
    If mobjLVW.SelectedItem Is Nothing Then Exit Sub
    If mobjLVW Is lvwBeds_s Then
        lng病人ID = mrsBeds!病人ID
        lng主页ID = mrsBeds!主页ID
    ElseIf mobjLVW Is lvwFamily_s Then
        lng病人ID = mrsFamily!病人ID
        lng主页ID = mrsFamily!主页ID
    End If
    If ExecPatiChange(EFun.E转医疗小组, Me, mstrPrivs, mlngUnit, lng病人ID, lng主页ID) Then
        Call LoadList(mobjLVW Is lvwBeds_s, mobjLVW Is lvwFamily_s, False, False)
    End If
End Sub

Private Sub mnuEdit_ChangeUnit_Click()
    Call ChangeUnit
End Sub

Private Sub mnuEdit_InUnit_Click()
    Dim strBeds As String, byt入科方式 As Byte, lng床位科室ID As Long
    Dim lng病人ID As Long, lng主页ID As Long
    
    If lvwBeds_s.SelectedItem Is Nothing Then
        strBeds = ""
    ElseIf lvwBeds_s.SelectedItem.Tag <> "空床" Then
        If mrsBeds!病人ID = mrsIn!病人ID Then '病人原住床
            strBeds = Trim(Mid(lvwBeds_s.SelectedItem.Key, 2))
            lng床位科室ID = Val("" & mrsIn!入住科室id)
        Else
            strBeds = ""
        End If
    ElseIf Not (mrsBeds!性别分类 = "不限床" Or (mrsBeds!性别分类 = "男床" And "" & mrsIn!性别 = "男") _
        Or (mrsBeds!性别分类 = "女床" And "" & mrsIn!性别 = "女")) Then
        strBeds = ""
    Else
        strBeds = Trim(Mid(lvwBeds_s.SelectedItem.Key, 2))
        lng床位科室ID = Val("" & mrsBeds!科室ID)
    End If
    byt入科方式 = Val(lvwIn_s.SelectedItem.Tag)
    lng病人ID = mrsIn!病人ID
    lng主页ID = mrsIn!主页ID
    
    Call ExecPatiChange(EFun.E入病区, Me, mstrPrivs, mlngUnit, lng病人ID, lng主页ID, strBeds, lng床位科室ID)
    
    '入科后定位
    If gblnOK Then
        If strBeds <> "" Then
            If InStr(strBeds, ",") Then
                strBeds = Split(strBeds, ",")(0)
                lvwBeds_s.ListItems("_" & strBeds).Selected = True
            Else
                lvwBeds_s.ListItems("_" & strBeds).Selected = True
            End If
        End If
        Call LoadList(True, True, True, False)
    End If
End Sub

Private Sub mnuEdit_Memo_Click()
    Dim lng病人ID As Long, lng主页ID As Long, strBeds As String
    
    If mobjLVW Is lvwBeds_s Then
        lng病人ID = mrsBeds!病人ID
        lng主页ID = mrsBeds!主页ID
    ElseIf mobjLVW Is lvwFamily_s Then
        lng病人ID = mrsFamily!病人ID
        lng主页ID = mrsFamily!主页ID
    ElseIf mobjLVW Is lvwOut_s Then
        lng病人ID = mrsOut!病人ID
        lng主页ID = mrsOut!主页ID
    ElseIf mobjLVW Is lvwIn_s Then
        lng病人ID = mrsIn!病人ID
        lng主页ID = mrsIn!主页ID
    End If
    
    Call ExecPatiChange(EFun.E病人备注编辑, Me, mstrPrivs, lng病人ID, lng主页ID)
'
'    If gblnOK Then Call LoadList(True, True, True, False)
End Sub

'问题27392 by lesfeng 2010-01-14
Private Sub mnuEdit_ModifOut_Click()
    '功能：修改病人出院时间
    Dim lng病人ID As Long, lng主页ID As Long, str姓名 As String
    
    If mobjLVW Is lvwOut_s Then
        lng病人ID = mrsOut!病人ID
        lng主页ID = mrsOut!主页ID
        str姓名 = mrsOut!姓名
        Call ExecPatiChange(EFun.E修改出院时间, Me, mstrPrivs, lng病人ID, lng主页ID)
        If gblnOK Then Call LoadList(False, False, False, True)
    End If
End Sub
'问题27866 by lesfeng 2010-02-05
Private Sub mnuEdit_OutAndModi_Click()
    frmOutAndModi.Show 1, Me
End Sub

Private Sub mnuEdit_Swap_Click()
    Call SwapBeds
End Sub

Private Sub mnuFile_PrintWristlet_Click()
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    '49854:刘鹏飞,2013-10-31,病人腕带打印(排除出院病人)
    If mobjLVW.SelectedItem Is Nothing Then Exit Sub
    If mobjLVW Is lvwBeds_s Then
        lng病人ID = mrsBeds!病人ID
        lng主页ID = mrsBeds!主页ID
    ElseIf mobjLVW Is lvwFamily_s Then
        lng病人ID = mrsFamily!病人ID
        lng主页ID = mrsFamily!主页ID
    ElseIf mobjLVW Is lvwIn_s Then
        lng病人ID = mrsIn!病人ID
        lng主页ID = mrsIn!主页ID
    End If
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_4", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_4", Me, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, 2)
    End If
End Sub

Private Sub mnuFileLocalSet_Click()
    frmSetCourse.mlngModul = mlngModul
    frmSetCourse.mstrPrivs = mstrPrivs
    frmSetCourse.Show 1, Me
    If gblnOK Then
        LoadList False, False, True, True
    End If
End Sub

Private Sub mnuViewFind_Click()
    Dim intIdx As Long
    With frmFindCourse
        .mstrSeekKey = mstrSeekKey
        .mstrSeekValue = mstrSeekValue
        .Show 1, Me
        If .mblnOk Then
            mstrSeekKey = .mstrSeekKey
            mstrSeekValue = .mstrSeekValue
            mlng当前病区id = .mlng病区id
            ' 问题30040 by lesfeng 2010-05-18
            If mstrSeekKey <> "床号" Then
                If mlng当前病区id <> mlngUnit Then
                    If InStr(mstrPrivs, "所有病区") <> 0 Then
                        intIdx = cbo.FindIndex(cboUnit, mlng当前病区id)
                        If intIdx <> -1 Then
                            cboUnit.ListIndex = intIdx
                        End If
'                        mlngUnit = mlng当前病区id
'                        Call LoadList
                    Else
                        MsgBox "你没有‘所有病区’权限，不能查找的 " & mstrSeekKey & "=" & mstrSeekValue & " 的病人！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
            Call SeekPati(True)
        End If
    End With
End Sub
Private Sub SeekPati(ByVal blnFirst As Boolean)
Dim lvwRow As Integer, lvwfor As Integer, intStart As Integer
Dim lvwTemp As ListView, lviTemp As ListItem, intColKey As Integer, lvwValue As String
    If mstrSeekValue = "" Then Exit Sub
    
reFind:
    lvwfor = lvwfor + 1
    If lvwfor > 4 Then
        MsgBox "没有你要查找的 " & mstrSeekKey & "=" & mstrSeekValue & " 的病人！", vbInformation, gstrSysName
        mstrSeekKey = "": mstrSeekValue = ""
        Exit Sub
    End If
    
    '设置当前查找的列表
    If lvwfor = 1 Then
        Set lvwTemp = mobjLVW
    Else '当前列表中没找到，进行列表切换
        Select Case lvwTemp.Name
            Case "lvwIn_s"
                Set lvwTemp = lvwBeds_s
            Case "lvwBeds_s"
                Set lvwTemp = lvwFamily_s
            Case "lvwFamily_s"
                Set lvwTemp = lvwOut_s
            Case "lvwOut_s"
                Set lvwTemp = lvwIn_s
        End Select
    End If
    
    '在当前列表中搜索
    With lvwTemp
        intStart = 1
        If Not blnFirst And lvwfor = 1 Then '如果非第一次查找,起始位置
            If .ListItems.Count > 0 Then intStart = .SelectedItem.Index + 1
        End If
        
        intColKey = GetColNum(lvwTemp, mstrSeekKey) '取列序
        If intColKey <> 0 Or mstrSeekKey = "医保号" Then
            For lvwRow = intStart To .ListItems.Count
                If mstrSeekKey = "医保号" Then '对医保号查找的特殊处理
                    If .Name = "lvwBeds_s" Then '取出病人ID
                        lvwValue = Trim(.ListItems(lvwRow).SubItems(GetColNum(lvwTemp, "病人ID") - 1))
                    Else
                        lvwValue = Trim(Split(.ListItems(lvwRow).Key, "_")(1))
                    End If
                    
                    If lvwValue <> "" Then '取医保号
                        lvwValue = GetInsureInfo(CLng(lvwValue))
                        If InStr(lvwValue, ";") > 0 Then
                            lvwValue = Trim(Split(lvwValue, ";")(1))
                        Else
                            lvwValue = ""
                        End If
                    End If
                Else
                    If intColKey = 1 Then '取出查找列对应值
                        lvwValue = Trim(.ListItems(lvwRow).Text)
                    Else
                        lvwValue = Trim(.ListItems(lvwRow).SubItems(intColKey - 1))
                    End If
                End If
                
                If mstrSeekValue = lvwValue Then '相同则定位并退出
                    Set lviTemp = .ListItems(lvwRow)
                    Select Case .Name
                        Case "lvwIn_s"
                            lvwIn_s.ListItems(lviTemp.Key).Selected = True
                            lvwIn_s.SelectedItem.EnsureVisible
                            Call lvwIn_s_ItemClick(lviTemp)
                            lvwIn_s.SetFocus
                            Call lvwIn_s_GotFocus
                        Case "lvwBeds_s"
                            lvwBeds_s.ListItems(lviTemp.Key).Selected = True
                            lvwBeds_s.SelectedItem.EnsureVisible
                            Call lvwBeds_s_ItemClick(lviTemp)
                            lvwBeds_s.SetFocus
                            Call lvwBeds_s_GotFocus
                        Case "lvwFamily_s"
                            lvwFamily_s.ListItems(lviTemp.Key).Selected = True
                            lvwFamily_s.SelectedItem.EnsureVisible
                            Call lvwFamily_s_ItemClick(lviTemp)
                            lvwFamily_s.SetFocus
                            Call lvwFamily_s_GotFocus
                        Case "lvwOut_s"
                            lvwOut_s.ListItems(lviTemp.Key).Selected = True
                            lvwOut_s.SelectedItem.EnsureVisible
                            Call lvwOut_s_ItemClick(lviTemp)
                            lvwOut_s.SetFocus
                            Call lvwOut_s_GotFocus
                    End Select
                    Exit Sub
                End If
            Next
        End If
    End With
    GoTo reFind
End Sub
Private Sub mnuViewFindNext_Click()
    Call SeekPati(False)
End Sub

Private Sub picOut_Click()
    If mblnDropOut Then
        If lvwFamily_s.Height >= lvwOut_s.Height Then
            lvwFamily_s.SetFocus
        Else
            lvwOut_s.SetFocus
        End If
    Else
        lvwOut_s.SetFocus
    End If
End Sub

Private Sub picOut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngPreY = Y: mblnDownOut = True: mblnDropOut = False
End Sub

Private Sub picOut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And mblnDownOut Then
        If lvwOut_s.Height - (Y - mlngPreY) < 600 Or lvwFamily_s.Height + Y - mlngPreY < 600 Then Exit Sub
        PicOut.Top = PicOut.Top + Y - mlngPreY
        lvwFamily_s.Height = lvwFamily_s.Height + Y - mlngPreY
        lvwOut_s.Top = lvwOut_s.Top + Y - mlngPreY
        lvwOut_s.Height = lvwOut_s.Height - (Y - mlngPreY)
        Me.Refresh
        mblnDropOut = True
    End If
End Sub

Private Sub picOut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnDownOut = False
End Sub

Private Sub lvwBeds_s_DblClick()
    If mblnBeds Then
        If lvwBeds_s.SelectedItem.Tag = "占用" Then
            mnuQueryInfo_Click
        End If
    End If
End Sub

Private Sub lvwBeds_s_GotFocus()
    Call ClearCard
    If Not lvwBeds_s.SelectedItem Is Nothing Then Call lvwBeds_s_ItemClick(lvwBeds_s.SelectedItem)
    Set mobjLVW = lvwBeds_s
    Call SetFocusColor
    
    Call SetMenu
End Sub

Private Sub ClearCard()
    lbl床号.Caption = "床号:"
    lbl姓名.Caption = "姓名"
    lbl性别.Caption = "性别"
    lbl年龄.Caption = "年龄"
    lbl标识.Caption = "住院号"
    lbl住院号.Caption = ""
    lbl医保号.Caption = ""
    lbl医保号.ForeColor = Me.ForeColor
    lbl病况.Caption = ""
    lbl病人类型.Caption = ""
    lbl入院时间.Caption = ""
    lbl护理等级.Caption = ""
    lbl医生.Caption = ""
    lblLevel.Caption = ""
    lbl医疗付款方式.Caption = ""
    lbl诊断.Caption = ""
End Sub

Private Sub lvwBeds_s_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strInfo As String
    If Item Is Nothing Then Exit Sub
    mrsBeds.Filter = "床号='" & Mid(Item.Key, 2) & "'"

    mblnBeds = True
    If Item.Tag = "占用" And Not IsNull(mrsBeds!病人ID) Then
        If Nvl(mrsBeds!房间号) <> "" Then
            lbl床号.Caption = "床号:" & mrsBeds!床号 & "(" & Nvl(mrsBeds!房间号) & ")"
        Else
            lbl床号.Caption = "床号:" & mrsBeds!床号
        End If
        If Not IsNull(mrsBeds!险类) Then strInfo = GetInsureInfo(mrsBeds!病人ID)
        If strInfo <> "" Then
            lbl医保号.Caption = Split(strInfo, ";")(1)
            lbl医保号.ForeColor = vbRed
        Else
            lbl医保号.Caption = "非医保病人"
            lbl医保号.ForeColor = Me.ForeColor
        End If
        
        lbl姓名.Caption = mrsBeds!姓名
        lbl性别.Caption = IIf(IsNull(mrsBeds!性别), "", mrsBeds!性别)
        lbl年龄.Caption = IIf(IsNull(mrsBeds!年龄), "", mrsBeds!年龄)
        
        If Not IsNull(mrsBeds!住院号) Then
            lbl标识.Caption = "住院号"
            lbl住院号.Caption = mrsBeds!住院号
        Else
            lbl标识.Caption = "病人ID"
            lbl住院号.Caption = mrsBeds!病人ID
        End If
        
        lbl病人类型.Caption = Nvl(mrsBeds!病人类型, "普通病人")
        lbl病况.Caption = IIf(IsNull(mrsBeds!当前病况), "", mrsBeds!当前病况)
        lbl入院时间.Caption = Format(mrsBeds!入院时间, "yyyy-MM-dd HH:mm")
        lbl护理等级.Caption = IIf(IsNull(mrsBeds!护理等级), "", mrsBeds!护理等级)
        lbl医生.Caption = IIf(IsNull(mrsBeds!住院医师), "", mrsBeds!住院医师)
        lblLevel.Caption = IIf(IsNull(mrsBeds!床位等级), "", mrsBeds!床位等级)
        lbl医疗付款方式.Caption = "" & mrsBeds!医疗付款方式
        lbl诊断.Caption = GetDiagnostic(mrsBeds!病人ID, Val("" & mrsBeds!主页ID))
        
        stbThis.Panels(2).Text = "住院号:" & IIf(IsNull(mrsBeds!住院号), "", mrsBeds!住院号) & " 入院:" & mrsBeds!入院时间 & _
            " 护理:" & IIf(IsNull(mrsBeds!护理等级), "", mrsBeds!护理等级) & " " & _
            " 科室:" & mrsBeds!科室 & " 等级:" & mrsBeds!床位等级
    Else
        ClearCard
        If Nvl(mrsBeds!房间号) <> "" Then
            lbl床号.Caption = "床号:" & mrsBeds!床号 & "(" & Nvl(mrsBeds!房间号) & ")"
        Else
            lbl床号.Caption = "床号:" & mrsBeds!床号
        End If
        lblLevel.Caption = IIf(IsNull(mrsBeds!床位等级), "", mrsBeds!床位等级)
        stbThis.Panels(2).Text = "科室:" & mrsBeds!科室 & " 等级:" & mrsBeds!床位等级
    End If
    
    Call SetMenu
End Sub

Private Function GetDiagnostic(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetDiagnosticInfo(lng病人ID, lng主页ID, "1,2,3", 2)
    If Not rsTmp Is Nothing Then
        If rsTmp.RecordCount > 0 Then
            rsTmp.Filter = "诊断类型=3"
            If rsTmp.RecordCount > 0 Then
                GetDiagnostic = "" & rsTmp!诊断描述
            Else
                rsTmp.Filter = "诊断类型=2"
                If rsTmp.RecordCount > 0 Then
                    GetDiagnostic = "" & rsTmp!诊断描述
                Else
                    rsTmp.Filter = "诊断类型=1"
                    If rsTmp.RecordCount > 0 Then GetDiagnostic = "" & rsTmp!诊断描述
                End If
            End If
        End If
    End If
End Function




Private Sub lvwBeds_s_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not lvwBeds_s.SelectedItem Is Nothing And KeyCode = vbKeyReturn Then
        mblnBeds = True: Call lvwBeds_s_DblClick
    End If
End Sub

Private Sub lvwBeds_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnBeds = False
    If Button = 2 Then
        PopupMenu mnuEdit, 2
    Else
        If lvwBeds_s.HitTest(X, Y) Is Nothing Then
            stbThis.Panels(2) = "病区病床共 " & mintBeds_A & " 张,病人占用 " & mintHolding & " 张,空床 " & mintBeds_A - mintHolding & " 张,转科病人 " & mintChange_A & " 个"
        End If
    End If
End Sub

Private Sub lvwFamily_s_DblClick()
    If mblnFamily Then mnuQueryInfo_Click
End Sub

Private Sub lvwFamily_s_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If InStr(1, mstrPrivs, "家庭病床") = 0 Then
        Static objIconFam As IPictureDisp
        If Not Source Is lvwIn_s Then
            If State = 0 Then
                Set objIconFam = Source.DragIcon
            ElseIf State = 2 Then
                Set Source.DragIcon = img32.ListImages("Limit").Picture
            ElseIf State = 1 Then
                Set Source.DragIcon = objIconFam
            End If
        End If
    End If
End Sub

Private Sub lvwFamily_s_GotFocus()
    Call ClearCard
    If Not lvwFamily_s.SelectedItem Is Nothing Then Call lvwFamily_s_ItemClick(lvwFamily_s.SelectedItem)
    Set mobjLVW = lvwFamily_s
    Call SetFocusColor
    
    Call SetMenu
End Sub

Private Sub lvwFamily_s_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strInfo As String
    If Item Is Nothing Then Exit Sub
    mrsFamily.Filter = "病人ID=" & Mid(Item.Key, 2)
    
    mblnFamily = True
    lbl床号.Caption = "床号:无"
    lblLevel.Caption = "家庭病床"
    
    strInfo = GetInsureInfo(mrsFamily!病人ID)
    If strInfo <> "" Then
        lbl医保号.Caption = Split(strInfo, ";")(1)
        lbl医保号.ForeColor = vbRed
    Else
        lbl医保号.Caption = "非医保病人"
        lbl医保号.ForeColor = Me.ForeColor
    End If
    
    lbl姓名.Caption = mrsFamily!姓名
    lbl性别.Caption = IIf(IsNull(mrsFamily!性别), "", mrsFamily!性别)
    lbl年龄.Caption = IIf(IsNull(mrsFamily!年龄), "", mrsFamily!年龄)
    
    If Not IsNull(mrsFamily!住院号) Then
        lbl标识.Caption = "住院号"
        lbl住院号.Caption = mrsFamily!住院号
    Else
        lbl标识.Caption = "病人ID"
        lbl住院号.Caption = mrsFamily!病人ID
    End If
    
    lbl病况.Caption = IIf(IsNull(mrsFamily!当前病况), "", mrsFamily!当前病况)
    lbl入院时间.Caption = Format(mrsFamily!入院时间, "yyyy-MM-dd HH:mm")
    lbl护理等级.Caption = IIf(IsNull(mrsFamily!护理等级), "", mrsFamily!护理等级)
    lbl医生.Caption = IIf(IsNull(mrsFamily!住院医师), "", mrsFamily!住院医师)
    lbl医疗付款方式.Caption = "" & mrsFamily!医疗付款方式
    lbl诊断.Caption = GetDiagnostic(mrsFamily!病人ID, Val("" & mrsFamily!主页ID))
    
    stbThis.Panels(2).Text = "住院号:" & IIf(IsNull(mrsFamily!住院号), "", mrsFamily!住院号) & " 护理:" & IIf(IsNull(mrsFamily!护理等级), "", mrsFamily!护理等级) & " 科室:" & mrsFamily!当前科室
    
    Call SetMenu
End Sub

Private Sub lvwFamily_s_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not lvwFamily_s.SelectedItem Is Nothing And KeyCode = vbKeyReturn Then
        mblnFamily = True: Call lvwFamily_s_DblClick
    End If
End Sub

Private Sub lvwFamily_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnFamily = False
    If Button = 2 Then
        PopupMenu mnuEdit, 2
    Else
        If lvwFamily_s.HitTest(X, Y) Is Nothing Then
            stbThis.Panels(2) = "家庭病床共 " & mintBeds_B & " 张,转科病人 " & mintChange_B & " 个"
        End If
    End If
End Sub

Private Sub lvwIn_s_DblClick()
    If mblnIn Then mnuQueryInfo_Click
End Sub

Private Sub lvwIn_s_GotFocus()
    If Not lvwIn_s.SelectedItem Is Nothing Then Call lvwIn_s_ItemClick(lvwIn_s.SelectedItem)
    Set mobjLVW = lvwIn_s
    Call SetFocusColor
    
    Call SetMenu
End Sub

Private Sub lvwIn_s_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item Is Nothing Then Exit Sub
    mrsIn.Filter = "病人ID=" & Mid(Item.Key, 2)
    mblnIn = True
    
    stbThis.Panels(2).Text = "住院号:" & IIf(IsNull(mrsIn!住院号), "", mrsIn!住院号) & " 当前科室:" & mrsIn!当前科室 & _
        IIf(IsNull(mrsIn!转入科室), "", " 转入科室:" & mrsIn!转入科室) & _
        " 入院时间:" & mrsIn!入院时间
    
    Call SetMenu
End Sub

Private Sub lvwIn_s_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not lvwIn_s.SelectedItem Is Nothing And KeyCode = vbKeyReturn Then
        mblnIn = True: Call lvwIn_s_DblClick
    End If
End Sub

Private Sub lvwIn_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnIn = False
    If Button = 2 Then
        PopupMenu mnuEdit, 2
    Else
        If lvwIn_s.HitTest(X, Y) Is Nothing Then
            stbThis.Panels(2) = "入科病人共 " & mintIn & " 个,新入院 " & mintIn - mintChange_C & " 个,它科转入 " & mintChange_C & " 个"
        End If
    End If
End Sub

Private Sub lvwOut_s_DblClick()
    If mblnOut Then mnuQueryInfo_Click
End Sub

Private Sub lvwOut_s_GotFocus()
    If Not lvwOut_s.SelectedItem Is Nothing Then Call lvwOut_s_ItemClick(lvwOut_s.SelectedItem)
    Set mobjLVW = lvwOut_s
    Call SetFocusColor
    
    Call SetMenu
End Sub

Private Sub lvwOut_s_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item Is Nothing Then Exit Sub
    '问题28365 by lesfeng 2010-03-04 病人多次出院时，没有增加主页id 过滤
    mrsOut.Filter = "病人ID=" & Split(Item.Key, "_")(1) & " and 主页id=" & Split(Item.Key, "_")(2)
    Call ClearCard
    mblnOut = True
        
    stbThis.Panels(2).Text = "住院号:" & IIf(IsNull(mrsOut!住院号), "", mrsOut!住院号) & " 出院科室:" & mrsOut!出院科室 & " 出院病床:" & mrsOut!出院病床 & " 出院时间:" & mrsOut!出院时间
    
    Call SetMenu
End Sub

Private Sub lvwOut_s_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not lvwOut_s.SelectedItem Is Nothing And KeyCode = vbKeyReturn Then
        mblnOut = True: Call lvwOut_s_DblClick
    End If
End Sub

Private Sub lvwOut_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnOut = False
    If Button = 2 Then
        PopupMenu mnuEdit, 2
    Else
        If lvwOut_s.HitTest(X, Y) Is Nothing Then
            stbThis.Panels(2) = "出院病人共 " & mintOut & " 个"
        End If
    End If
End Sub

Private Sub mnuEdit_AddBeds_Click()
    If Not mobjLVW Is lvwBeds_s Then Exit Sub
            
    Call ChangeBeds(1)
End Sub

Private Sub mnuEdit_Adjust_Click()
    Dim lng病人ID As Long, lng主页ID As Long
    
    If mobjLVW Is lvwFamily_s Then
        lng病人ID = mrsFamily!病人ID
        lng主页ID = mrsFamily!主页ID
    ElseIf mobjLVW Is lvwBeds_s Then
        lng病人ID = mrsBeds!病人ID
        lng主页ID = mrsBeds!主页ID
    End If
    Call ExecPatiChange(EFun.E调整病人信息, Me, mstrPrivs, mlngUnit, lng病人ID, lng主页ID)
    
    If gblnOK Then Call LoadList(True, True, False, False)
End Sub

Private Sub mnuEdit_BabyReg_Click()
    Dim lng病人ID As Long, lng主页ID As Long
    
    If mobjLVW Is lvwFamily_s Then
        lng病人ID = mrsFamily!病人ID
        lng主页ID = mrsFamily!主页ID
    ElseIf mobjLVW Is lvwBeds_s Then
        lng病人ID = mrsBeds!病人ID
        lng主页ID = mrsBeds!主页ID
    End If
    Call ExecPatiChange(EFun.E新生儿登记, Me, mstrPrivs, lng病人ID, lng主页ID)
End Sub

Private Sub mnuEdit_Disease_Click()
    Dim lng病人ID As Long, lng主页ID As Long, int险类 As Integer
    
    If mobjLVW Is lvwFamily_s Then
        lng病人ID = mrsFamily!病人ID
        lng主页ID = mrsFamily!主页ID
        int险类 = Nvl(mrsFamily!险类, 0)
    ElseIf mobjLVW Is lvwBeds_s Then
        lng病人ID = mrsBeds!病人ID
        lng主页ID = mrsBeds!主页ID
        int险类 = Nvl(mrsBeds!险类, 0)
    ElseIf mobjLVW Is lvwOut_s Then
        lng病人ID = mrsOut!病人ID
        lng主页ID = mrsOut!主页ID
        int险类 = Nvl(mrsOut!险类, 0)
    End If
    Call ExecPatiChange(EFun.E医保病种选择, Me, mstrPrivs, lng病人ID, lng主页ID, int险类)
End Sub

Private Sub mnuEdit_Level_Click()
    
    Call ExecPatiChange(EFun.E更改床位等级, Me, mstrPrivs, mrsBeds!病人ID, mrsBeds!主页ID, mrsBeds!床号)
    
    If gblnOK Then Call LoadList(True, False, False, False)
End Sub

Private Sub mnuEdit_Change_Click()
    Dim lng病人ID As Long, lng主页ID As Long, strBeds As String
    
    If mobjLVW Is lvwBeds_s Then
        lng病人ID = mrsBeds!病人ID
        lng主页ID = mrsBeds!主页ID
    Else
        lng病人ID = mrsFamily!病人ID
        lng主页ID = mrsFamily!主页ID
    End If
    
    Call ExecPatiChange(EFun.E转科, Me, mstrPrivs, mlngUnit, lng病人ID, lng主页ID)
    
    If gblnOK Then Call LoadList(True, True, True, False)
End Sub

Private Sub mnuEdit_In_Click()
    Dim strBeds As String, byt入科方式 As Byte, lng床位科室ID As Long
    Dim lng病人ID As Long, lng主页ID As Long
    
    If lvwBeds_s.SelectedItem Is Nothing Then
        strBeds = ""
    ElseIf lvwBeds_s.SelectedItem.Tag <> "空床" Then
        If mrsBeds!病人ID = mrsIn!病人ID Then '病人原住床
            strBeds = Trim(Mid(lvwBeds_s.SelectedItem.Key, 2))
            lng床位科室ID = Val("" & mrsIn!入住科室id)
        Else
            strBeds = ""
        End If
    ElseIf Not (mrsBeds!性别分类 = "不限床" Or (mrsBeds!性别分类 = "男床" And "" & mrsIn!性别 = "男") _
        Or (mrsBeds!性别分类 = "女床" And "" & mrsIn!性别 = "女")) Then
        strBeds = ""
    Else
        strBeds = Trim(Mid(lvwBeds_s.SelectedItem.Key, 2))
        lng床位科室ID = Val("" & mrsBeds!科室ID)
    End If
    byt入科方式 = Val(lvwIn_s.SelectedItem.Tag)
    lng病人ID = mrsIn!病人ID
    lng主页ID = mrsIn!主页ID
    
    If byt入科方式 <> 2 Then
        Call ExecPatiChange(EFun.E入科, Me, mstrPrivs, mlngUnit, lng病人ID, lng主页ID, strBeds, lng床位科室ID, byt入科方式)
    ElseIf byt入科方式 = 2 Then
        Call ExecPatiChange(EFun.E入病区, Me, mstrPrivs, mlngUnit, lng病人ID, lng主页ID, strBeds, lng床位科室ID)
    End If
    '入科后定位
    If gblnOK Then
        If strBeds <> "" Then
            If InStr(strBeds, ",") Then
                strBeds = Split(strBeds, ",")(0)
                lvwBeds_s.ListItems("_" & strBeds).Selected = True
            Else
                lvwBeds_s.ListItems("_" & strBeds).Selected = True
            End If
        End If
        Call LoadList(True, True, True, False, True)
    End If
End Sub

Private Sub mnuEdit_Move_Click()
    Call ChangeBeds(0)
End Sub

Private Sub ChangeBeds(ByVal bytFun As Byte, Optional ByVal str目标床号 As String)
'参数:bytFun:0-换床,1-包房
    Dim lng病人ID As Long, lng主页ID As Long
    
    If mobjLVW.SelectedItem Is Nothing Then Exit Sub
    If mobjLVW Is lvwBeds_s Or bytFun = 1 Then
        lng病人ID = mrsBeds!病人ID
        lng主页ID = mrsBeds!主页ID
    Else
        lng病人ID = mrsFamily!病人ID
        lng主页ID = mrsFamily!主页ID
    End If
        
    Call ExecPatiChange(EFun.E换床, Me, mstrPrivs, mlngUnit, lng病人ID, lng主页ID, bytFun, str目标床号, "")
        
    If gblnOK Then
        If str目标床号 <> "" Then
            On Error Resume Next '可能目标病床床号在本病区没有，屏避错误
            If InStr(str目标床号, ",") > 0 Then '传出床号
                str目标床号 = Split(str目标床号, ",")(0)
                lvwBeds_s.ListItems("_" & str目标床号).Selected = True
            Else
                lvwBeds_s.ListItems("_" & str目标床号).Selected = True
            End If
            Err.Clear
        End If
    
        Call LoadList(True, True, False, False)
    End If
End Sub

Private Sub ChangeUnit()
    Dim lng病人ID As Long, lng主页ID As Long
    
    If mobjLVW.SelectedItem Is Nothing Then Exit Sub
    If mobjLVW Is lvwBeds_s Then
        lng病人ID = mrsBeds!病人ID
        lng主页ID = mrsBeds!主页ID
    Else
        lng病人ID = mrsFamily!病人ID
        lng主页ID = mrsFamily!主页ID
    End If
        
    Call ExecPatiChange(EFun.E转病区, Me, mstrPrivs, mlngUnit, lng病人ID, lng主页ID)
        
    If gblnOK Then
'
'            On Error Resume Next '可能目标病床床号在本病区没有，屏避错误
'            If InStr(str目标床号, ",") > 0 Then '传出床号
'                str目标床号 = Split(str目标床号, ",")(0)
'                lvwBeds_s.ListItems("_" & str目标床号).Selected = True
'            Else
'                lvwBeds_s.ListItems("_" & str目标床号).Selected = True
'            End If
'            Err.Clear
'        End If
'
        Call LoadList(True, True, True, True)
    End If
End Sub

Private Sub SwapBeds(Optional ByVal str目标床号 As String)
'###########################################################################################################
'## 功能：同病区病人床位对换
'## 参数：病人床位对换的目标床号，可选
'##
'###########################################################################################################
    
    Dim lng病人ID As String, lng主页ID As String, str床号 As String
    Dim rsTmp As ADODB.Recordset
    
    lng病人ID = mrsBeds!病人ID
    lng主页ID = mrsBeds!主页ID
    str床号 = mrsBeds!床号
    
    If Trim(str目标床号) <> "" Then
        '拖动到的目标床位号与当前床位号相同时不执行任何操作
        If str床号 = str目标床号 Then Exit Sub
        '目标床位病人信息与当前床位病人信息相同则不执行任何操作(包床病人)
        If lng病人ID = mrsCBeds!病人ID And lng主页ID = mrsCBeds!主页ID Then Exit Sub
    End If
    
    If ExecPatiChange(EFun.E床位对换, Me, mstrPrivs, mlngUnit, lng病人ID, lng主页ID, str床号, str目标床号) Then
        If str目标床号 <> "" Then
            On Error Resume Next '可能目标病床床号在本病区没有，屏避错误
            If InStr(str目标床号, ",") > 0 Then '传出床号
                str目标床号 = Split(str目标床号, ",")(0)
                lvwBeds_s.ListItems("_" & str目标床号).Selected = True
            Else
                lvwBeds_s.ListItems("_" & str目标床号).Selected = True
            End If
            Err.Clear
        End If
    
        Call LoadList(True, True, False, False)
    End If
End Sub

Private Sub mnuEdit_Nurse_Click()
    On Error Resume Next
    Err.Clear
    
    frmNurse.mblnBed = (mobjLVW Is lvwBeds_s)
    frmNurse.Show 1, Me
    If gblnOK Then Call LoadList(True, True, False, False, True)
End Sub

Private Sub mnuEdit_Out_Click()
    Dim lng病人ID As Long, lng主页ID As Long
    
    If mobjLVW Is lvwFamily_s Then
        lng病人ID = mrsFamily!病人ID
        lng主页ID = mrsFamily!主页ID
    Else
        lng病人ID = mrsBeds!病人ID
        lng主页ID = mrsBeds!主页ID
    End If
    
    Call ExecPatiChange(EFun.E出院, Me, mstrPrivs, lng病人ID, lng主页ID)
    
    If gblnOK Then Call LoadList(True, True, False, True, True)
End Sub

Private Sub mnuEdit_PreOut_Click()
'功能：病人预出院
    Dim lng病人ID As Long, lng主页ID As Long, str姓名 As String
    Dim blnTrue As Boolean
    
    If mobjLVW Is lvwFamily_s Then
        lng病人ID = mrsFamily!病人ID
        lng主页ID = mrsFamily!主页ID
        str姓名 = mrsFamily!姓名
    Else
        lng病人ID = mrsBeds!病人ID
        lng主页ID = mrsBeds!主页ID
        str姓名 = mrsBeds!姓名
    End If
    '--55791:刘鹏飞,2012-11-13,作废出院医嘱才能撤销出院
    On Error Resume Next
    blnTrue = frmPreOut.ShowMe(Me, lng病人ID, lng主页ID, str姓名, mstrPrivs)
    
    If blnTrue = True Then
        Call LoadList(True, True, False, False)
    End If
End Sub

Private Sub mnuEdit_Recalc_Click()
    Dim lng病人ID As Long, lng主页ID As Long, str姓名 As String
    Dim rsTmp As ADODB.Recordset
    
    If mobjLVW Is lvwBeds_s Then
        Set rsTmp = mrsBeds
    ElseIf mobjLVW Is lvwFamily_s Then
        Set rsTmp = mrsFamily
    ElseIf mobjLVW Is lvwOut_s Then
        Set rsTmp = mrsOut
    ElseIf mobjLVW Is lvwIn_s Then
        Set rsTmp = mrsIn
    Else
        Exit Sub
    End If
    
    lng病人ID = rsTmp!病人ID
    lng主页ID = rsTmp!主页ID
    str姓名 = rsTmp!姓名
        
    gblnOK = False
    Call ExecPatiChange(EFun.E重算费用, Me, mstrPrivs, lng病人ID, lng主页ID, str姓名)
    
    If gblnOK Then stbThis.Panels(2).Text = "费用重算操作成功完成!"
End Sub

Private Sub mnuEdit_Undo_Click()
    Dim lng病人ID As Long, lng主页ID As Long
    Dim int险类 As Integer
    
    If mobjLVW Is lvwBeds_s Then
        lng病人ID = mrsBeds!病人ID
        lng主页ID = mrsBeds!主页ID
        int险类 = Nvl(mrsBeds!险类, 0)
    ElseIf mobjLVW Is lvwFamily_s Then
        lng病人ID = mrsFamily!病人ID
        lng主页ID = mrsFamily!主页ID
        int险类 = Nvl(mrsFamily!险类, 0)
    ElseIf mobjLVW Is lvwOut_s Then
        lng病人ID = mrsOut!病人ID
        lng主页ID = mrsOut!主页ID
        int险类 = Nvl(mrsOut!险类, 0)
    ElseIf mobjLVW Is lvwIn_s Then
        Exit Sub
    End If
    
    gblnOK = False
    Call ExecPatiChange(EFun.E撤销, Me, mstrPrivs, mlngUnit, lng病人ID, lng主页ID, int险类, CStr(tbr.Buttons("Undo").ButtonMenus(1).Text))
    
    If gblnOK Then Call LoadList(True, True, True, True, True)
End Sub

Private Sub mnuEditToInPati_Click()
    Dim lng病人ID As Long, lng主页ID As Long
    Dim str住院号 As String, str姓名 As String
        
    If MsgBox("确实要将该住院留观病人转为住院病人吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If mobjLVW Is lvwBeds_s Then
        lng病人ID = mrsBeds!病人ID
        lng主页ID = mrsBeds!主页ID
        
        str住院号 = IIf(IsNull(mrsBeds!住院号), "", mrsBeds!住院号)
        str姓名 = mrsBeds!姓名
    ElseIf mobjLVW Is lvwFamily_s Then
        lng病人ID = mrsFamily!病人ID
        lng主页ID = mrsFamily!主页ID
        
        str住院号 = IIf(IsNull(mrsFamily!住院号), "", mrsFamily!住院号)
        str姓名 = mrsFamily!姓名
    End If
    gblnOK = False
    Call ExecPatiChange(EFun.E转为住院, Me, mstrPrivs, lng病人ID, lng主页ID, str住院号, str姓名)
    
    If gblnOK Then Call LoadList(True, True, False, False)
End Sub

Private Sub mnuFilePrintCard_Click()
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    If mobjLVW.SelectedItem Is Nothing Then Exit Sub
    If mobjLVW Is lvwBeds_s Then
        lng病人ID = mrsBeds!病人ID
        lng主页ID = mrsBeds!主页ID
    ElseIf mobjLVW Is lvwFamily_s Then
        lng病人ID = mrsFamily!病人ID
        lng主页ID = mrsFamily!主页ID
    End If
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_2", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_2", Me, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, 2)
    End If
End Sub

Private Sub mnuFilePrintMed_Click()
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    
    If mobjLVW Is lvwBeds_s Then
        lng病人ID = mrsBeds!病人ID
        lng主页ID = mrsBeds!主页ID
    ElseIf mobjLVW Is lvwFamily_s Then
        lng病人ID = mrsFamily!病人ID
        lng主页ID = mrsFamily!主页ID
    ElseIf mobjLVW Is lvwIn_s Then
        lng病人ID = mrsIn!病人ID
        lng主页ID = mrsIn!主页ID
    ElseIf mobjLVW Is lvwOut_s Then
        lng病人ID = mrsOut!病人ID
        lng主页ID = mrsOut!主页ID
    End If
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_1", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_1", Me, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, 2)
    End If
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub
Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuQuery_Log_Click()
    If mobjLVW Is lvwBeds_s Then
        frmHistory.mlng病人ID = mrsBeds!病人ID
        frmHistory.mlng主页ID = mrsBeds!主页ID
    ElseIf mobjLVW Is lvwOut_s Then
        frmHistory.mlng病人ID = mrsOut!病人ID
        frmHistory.mlng主页ID = mrsOut!主页ID
    ElseIf mobjLVW Is lvwFamily_s Then
        frmHistory.mlng病人ID = mrsFamily!病人ID
        frmHistory.mlng主页ID = mrsFamily!主页ID
    Else
        Exit Sub
    End If

    On Error Resume Next
    Err.Clear
    
    frmHistory.Show 1, Me
End Sub


Private Sub mnuQueryInfo_Click()
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    
    If mobjLVW Is lvwBeds_s Then
        lng病人ID = mrsBeds!病人ID
        lng主页ID = mrsBeds!主页ID
    ElseIf mobjLVW Is lvwFamily_s Then
        lng病人ID = mrsFamily!病人ID
        lng主页ID = mrsFamily!主页ID
    ElseIf mobjLVW Is lvwIn_s Then
        lng病人ID = mrsIn!病人ID
        lng主页ID = mrsIn!主页ID
    Else
        lng病人ID = mrsOut!病人ID
        lng主页ID = mrsOut!主页ID
    End If
    
    On Error Resume Next
    Err.Clear
    
    If CreatePublicPatient() Then
        Call gobjPublicPatient.ReadPatiDegreeCard(Me, lng病人ID, lng主页ID)
    End If
    
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim lng病人ID As Long
    
    If lvwBeds_s Is Me.ActiveControl And Not lvwBeds_s.SelectedItem Is Nothing Then
        lng病人ID = Val("" & mrsBeds!病人ID)
    ElseIf lvwFamily_s Is Me.ActiveControl And Not lvwFamily_s.SelectedItem Is Nothing Then
        lng病人ID = Val("" & mrsFamily!病人ID)
    ElseIf lvwOut_s Is Me.ActiveControl And Not lvwOut_s.SelectedItem Is Nothing Then
        lng病人ID = Val("" & mrsOut!病人ID)
    ElseIf lvwIn_s Is Me.ActiveControl And Not lvwIn_s.SelectedItem Is Nothing Then
        lng病人ID = Val("" & mrsIn!病人ID)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "病区=" & mlngUnit, "病人ID=" & lng病人ID)
End Sub

Private Sub mnuView_Card_Click()
    mnuView_Card.Checked = Not mnuView_Card.Checked
    picCard_s.Visible = mnuView_Card.Checked
    If picCard_s.Visible Then
        If picCard_s.Left <= -picCard_s.width Or picCard_s.Top <= -picCard_s.Height Then
            With picCard_s
                .Left = picVsc.Left - (picCard_s.width - picVsc.width) / 2
                .Top = lvwBeds_s.Top + (lvwBeds_s.Height - picCard_s.Height) / 2
            End With
        End If
    End If
End Sub

Private Sub mnuView_ListView_Click(Index As Integer)
    Call SetView(CByte(Index))
End Sub

Private Sub mnuViewColSel_Click()
    Select Case mobjLVW.Name
        Case "lvwBeds_s"
            If zlControl.LvwSelectColumns(lvwBeds_s, COL_BEDS) Then
                LoadList True, False, False, False
            End If
        Case "lvwFamily_s"
            If zlControl.LvwSelectColumns(lvwFamily_s, COL_FAMILY) Then
                LoadList False, True, False, False
            End If
        Case "lvwIn_s"
            If zlControl.LvwSelectColumns(lvwIn_s, COL_IN) Then
                LoadList False, False, True, False
            End If
        Case "lvwOut_s"
            If zlControl.LvwSelectColumns(lvwOut_s, COL_OUT) Then
                LoadList False, False, False, True
            End If
    End Select
End Sub

Private Sub mnuViewreFlash_Click()
    Call LoadList
    Me.Refresh
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub mnuViewToolUnit_Click()
    mnuViewToolUnit.Checked = Not mnuViewToolUnit.Checked
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = False
    cbr.Bands(2).Visible = Not cbr.Bands(2).Visible
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = True
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Bands(1).Visible = Not cbr.Bands(1).Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub picCard_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Set picCard_s.MouseIcon = img32.ListImages("Down").Picture
        Call MoveObj(picCard_s.hWnd)
        Set picCard_s.MouseIcon = img32.ListImages("Up").Picture
        mobjLVW.SetFocus
    ElseIf Button = 2 Then
        PopupMenu mnuEdit, 2
    End If
End Sub

Private Sub picCard_s_OLECompleteDrag(Effect As Long)
    mobjLVW.SetFocus
End Sub
Private Sub picVsc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngPreX = X: mblnDownVsc = True
End Sub

Private Sub picVsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And mblnDownVsc Then
        If lvwBeds_s.width + X - mlngPreX < 1500 Or lvwFamily_s.width - (X - mlngPreX) < 1000 Then Exit Sub
        picVsc.Left = picVsc.Left + X - mlngPreX
        lblBed.width = lblBed.width + X - mlngPreX
        lvwBeds_s.width = lvwBeds_s.width + X - mlngPreX
        lblIn.width = lblIn.width + X - mlngPreX
        lvwIn_s.width = lvwIn_s.width + X - mlngPreX
        lblFamily.Left = lblFamily.Left + X - mlngPreX
        lblFamily.width = lblFamily.width - (X - mlngPreX)
        lvwFamily_s.Left = lvwFamily_s.Left + X - mlngPreX
        lvwFamily_s.width = lvwFamily_s.width - (X - mlngPreX)
        PicOut.Left = PicOut.Left + X - mlngPreX
        PicOut.width = PicOut.width - (X - mlngPreX)
        lvwOut_s.Left = lvwOut_s.Left + X - mlngPreX
        lvwOut_s.width = lvwOut_s.width - (X - mlngPreX)
        Me.Refresh
    End If
End Sub

Private Sub picVsc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnDownVsc = False
    mobjLVW.SetFocus
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PatiColor" Then
        zlDatabase.ShowPatiColorTip Me
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "View"
            Call SetView((mobjLVW.View + 1) Mod 4)
        Case "In"
            mnuEdit_In_Click
        Case "Change"
            mnuEdit_Change_Click
        Case "Out"
            mnuEdit_Out_Click
        Case "Move"
            mnuEdit_Move_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Adjust"
            mnuEdit_Adjust_Click
        Case "Undo"
            If mnuEdit_Undo.Enabled And mnuEdit_Undo.Visible Then mnuEdit_Undo_Click
    End Select
End Sub

Private Sub SetView(bytStyle As Byte)
'功能：调整列表显示方式
'参数：bytstyle=0-大图标,1-小图标,2-列表,3-详细资料
    mnuView_ListView(0).Checked = False
    mnuView_ListView(1).Checked = False
    mnuView_ListView(2).Checked = False
    mnuView_ListView(3).Checked = False
    mnuView_ListView(bytStyle).Checked = True
    mobjLVW.View = bytStyle
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Icon"
            Call SetView(0)
        Case "Small"
            Call SetView(1)
        Case "List"
            Call SetView(2)
        Case "Detail"
            Call SetView(3)
        Case Else
            mnuEdit_Undo_Click
    End Select
End Sub

Private Sub lvwBeds_s_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwBeds_s.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwBeds_s.SortOrder = lvwDescending
    Else
        lvwBeds_s.SortOrder = lvwAscending
    End If
    lvwBeds_s.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwBeds_s.SelectedItem Is Nothing Then lvwBeds_s.SelectedItem.EnsureVisible
End Sub

Private Sub lvwFamily_s_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwFamily_s.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwFamily_s.SortOrder = lvwDescending
    Else
        lvwFamily_s.SortOrder = lvwAscending
    End If
    lvwFamily_s.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwFamily_s.SelectedItem Is Nothing Then lvwFamily_s.SelectedItem.EnsureVisible
End Sub

Private Sub lvwIn_s_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwIn_s.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwIn_s.SortOrder = lvwDescending
    Else
        lvwIn_s.SortOrder = lvwAscending
    End If
    lvwIn_s.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwIn_s.SelectedItem Is Nothing Then lvwIn_s.SelectedItem.EnsureVisible
End Sub

Private Sub lvwOut_s_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwOut_s.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwOut_s.SortOrder = lvwDescending
    Else
        lvwOut_s.SortOrder = lvwAscending
    End If
    lvwOut_s.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwOut_s.SelectedItem Is Nothing Then lvwOut_s.SelectedItem.EnsureVisible
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub SetFocusColor()
    Dim i As Integer
    
    '设置当前列表突出显示
    lblBed.BackColor = COLOR_LOST
    lblFamily.BackColor = COLOR_LOST
    lblIn.BackColor = COLOR_LOST
    PicOut.BackColor = COLOR_LOST
        chk结清(0).BackColor = COLOR_LOST
        chk结清(1).BackColor = COLOR_LOST
        lblOut.BackColor = COLOR_LOST
    Select Case mobjLVW.Name
        Case "lvwBeds_s"
            lblBed.BackColor = COLOR_FOCUS
        Case "lvwFamily_s"
            lblFamily.BackColor = COLOR_FOCUS
        Case "lvwIn_s"
            lblIn.BackColor = COLOR_FOCUS
        Case "lvwOut_s"
            PicOut.BackColor = COLOR_FOCUS
            chk结清(0).BackColor = COLOR_FOCUS
            chk结清(1).BackColor = COLOR_FOCUS
            lblOut.BackColor = COLOR_FOCUS
    End Select
    '获取当前列表显示方式
    mnuView_ListView(0).Checked = False
    mnuView_ListView(1).Checked = False
    mnuView_ListView(2).Checked = False
    mnuView_ListView(3).Checked = False
    mnuView_ListView(mobjLVW.View).Checked = True
    
    If Not mobjLVW.SelectedItem Is Nothing Then mobjLVW.SelectedItem.EnsureVisible
End Sub

Private Function InitUnits() As Boolean
'功能：初始化住院病区
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, blnLimitUnit As Boolean, strUnitIDs As String
    
    On Error GoTo errH
    
    blnLimitUnit = InStr(mstrPrivs, "所有病区") = 0
    If blnLimitUnit Then
        strUnitIDs = "," & GetUserUnits(False) & ","
    Else
        strUnitIDs = "," & GetUserUnits(True) & ","
    End If
    'by lesfeng 2010-01-12 性能优化
    '目前包含门诊观察室
    gstrSQL = _
        " Select A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where B.部门ID = A.ID" & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.服务对象 IN(1,2,3) And B.工作性质='护理'" & _
        IIf(blnLimitUnit, " And instr([1],',' || A.ID || ',')>0 ", "") & _
        " And (A.站点=[2] Or A.站点 is Null)" & _
        " Order by A.编码"
        '
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strUnitIDs, gstrNodeNo)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = UserInfo.部门ID And cboUnit.ListIndex = -1 Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
        If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0  '调用Click事件
    ElseIf InStr(";" & mstrPrivs, "所有病区") > 0 Then
        MsgBox "没有设置病区,请到部门管理中设置工作性质为护理的部门！", vbExclamation, gstrSysName
        Exit Function
    Else
        MsgBox "你没有 [所有病区] 的权限,并且你所在部门不是病区或不属于病区！", vbExclamation, gstrSysName
        Exit Function
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ReadBedsMap(lngUnitID As Long) As Boolean
'功能：读取指定病区的床位映象表(含床位信息、病人身份信息、病人在院信息),并显示在列表中
    Dim i As Integer, j As Integer, strIcon As String
    Dim objItem As ListItem, blnChange As Boolean
    Dim bytLen As Byte, strChange As String
    Dim strTmp As String
    Dim k As Integer
    Dim strTemp As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    '附加条件
    strTmp = ""
    gstrSQL = ""
    For i = 1 To tbrFilter.Buttons.Count
        If tbrFilter.Buttons(i).Key Like "Nurse*" And tbrFilter.Buttons(i).Image = "Check" Then
            strTmp = strTmp & "," & Val(Replace(tbrFilter.Buttons(i).Key, "Nurse", ""))
        End If
    Next
    strTmp = strTmp & ","
    gstrSQL = " And (instr([2],',' || B.护理等级ID || ',')>0  Or B.护理等级ID is NULL)"
    
    If tbrFilter.Buttons("curDay").Image = "Check_" Then
        gstrSQL = gstrSQL & " And B.入院日期 Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60"
    End If
    
    gstrSQL = _
        "Select A.床号, A.科室id, A.科室, A.房间号, A.性别分类, A.床位编制, A.床位等级id, A.床位等级, A.状态, A.共用, B.主页id," & vbNewLine & _
        "       Nvl(B.状态, 0) As 病人状态, B.当前科室id, B.当前科室, B.病人id, B.住院号, B.姓名, B.性别, B.年龄, B.医疗付款方式," & vbNewLine & _
        "       B.合同单位id, B.当前病况, To_Char(B.入院日期, 'YYYY-MM-DD HH24:MI:SS') As 入院时间, B.护理等级id, B.护理等级," & vbNewLine & _
        "       B.住院医师, B.病人性质, B.险类,B.当前床号 as 主要床号,B.就诊卡号,B.身份证号,B.IC卡号,B.病人类型" & vbNewLine & _
        "From (Select A.床号,A.顺序号, A.科室id, Nvl(C.名称, Decode(A.共用, 1, '<共用病床>', Null)) As 科室, A.房间号, A.性别分类," & vbNewLine & _
        "              A.床位编制, A.等级id As 床位等级id, B.名称 As 床位等级, A.状态, A.共用" & vbNewLine & _
        "       From 床位状况记录 A, 收费项目目录 B, 部门表 C" & vbNewLine & _
        "       Where A.科室id = C.ID(+) And A.等级id = B.ID(+) And A.病区id = [1]) A," & vbNewLine & _
        "     (Select Distinct B.主页id, B.状态, B.出院科室id As 当前科室id, E.名称 As 当前科室, A.病人id, B.住院号, C.床号, NVL(B.姓名,A.姓名) 姓名, NVL(B.性别,A.性别) 性别," & vbNewLine & _
        "              NVL(B.年龄,A.年龄) 年龄, A.医疗付款方式, A.合同单位id, B.当前病况, B.入院日期, B.护理等级id, D.名称 As 护理等级, B.住院医师," & vbNewLine & _
        "              B.病人性质, B.险类,A.当前床号,A.就诊卡号,A.身份证号,A.IC卡号,Nvl(B.病人类型,Decode(B.险类,Null,'普通病人','医保病人')) 病人类型 " & vbNewLine & _
        "       From 病人信息 A, 病案主页 B, 病人变动记录 C, 收费项目目录 D, 部门表 E, 床位状况记录 F" & vbNewLine & _
        "       Where B.病人id = A.病人id And C.病人id = B.病人id And C.主页id = B.主页id And B.护理等级id = D.ID(+) " & gstrSQL & " And" & vbNewLine & _
        "             B.出院科室id = E.ID And B.出院日期 Is Null And Nvl(B.主页id, 0) <> 0 And Nvl(B.状态, 0) In (0, 2, 3) And" & vbNewLine & _
        "             C.开始时间 Is Not Null And C.终止时间 Is Null And C.床号 Is Not Null And F.病人id = B.病人id And" & vbNewLine & _
        "             F.病区id = [1] And F.病人id Is Not Null) B" & vbNewLine & _
        "Where A.床号 = B.床号(+)" & vbNewLine & _
        "Order By A.顺序号,LPad(A.床号, 10, ' ')"
    Set mrsBeds = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngUnitID, strTmp)
    Set mrsCBeds = mrsBeds.Clone
    
    mintBeds_A = 0: mintHolding = 0: mintChange_A = 0
    
    With mrsBeds
        If Not .EOF Then
            bytLen = GetMaxBedLen(lngUnitID)
            For i = 1 To .RecordCount
                If Not (!状态 = "占用" And IsNull(!病人ID)) Then
                    blnChange = False
                    If !病人状态 = 2 Then '转科病人
                        blnChange = True
                        If Not IsNull(!病人ID) Then
                            If InStr(strChange & ",", "," & !病人ID & ",") = 0 Then
                                strChange = strChange & "," & !病人ID
                            End If
                        End If
                    End If
'
'                    If Not (IsNull(!病人id) And IsNull(!主页id)) Then
'                        gstrSQL = "Select 1 From 病人变动记录 Where 病人id=[1] And 主页id=[2] And 开始时间 Is Null And 终止时间 Is Null And 开始原因=15 "
'                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(!病人id), Val(!主页id))
'
'                        If rsTmp.RecordCount > 0 Then
'                            blnChange = True
'                        End If
'                    End If
                    
                    If blnChange Then
                        strIcon = "Change"
                    Else
                        Select Case !状态
                            Case "空床"
                                If !性别分类 = "男床" Then
                                    strIcon = "M_Empty"
                                ElseIf !性别分类 = "女床" Then
                                    strIcon = "F_Empty"
                                Else
                                    strIcon = "Empty"
                                End If
                            Case "修缮"
                                strIcon = "Remedy"
                            Case "占用"
                                If !病人状态 = 3 Then
                                    strIcon = "Out" '预出院病人
                                Else
                                    strIcon = "Holding"
                                End If
                                mintHolding = mintHolding + 1
                        End Select
                    End If
                    
                    '留观病人图标
                    If IIf(IsNull(!病人性质), 0, !病人性质) <> 0 Then strIcon = "K" & strIcon
                    
                    '加床的非编图标
                    If IIf(IsNull(!床位编制), "", !床位编制) = "加床" Then
                        strIcon = "加床_" & strIcon
                    ElseIf IIf(IsNull(!床位编制), "", !床位编制) = "非编" Then
                        strIcon = "非编_" & strIcon
                    End If
                    '共用病床表示
                    If IIf(IsNull(!共用), 0, !共用) = 1 Then
                        strIcon = "共用_" & strIcon
                    End If
                    '问题29710 by lesfeng 2010-05-12 强制将病床置回第一列，因为在其它列时出错
                    strTemp = lvwBeds_s.ColumnHeaders(1).Text
                    For k = 1 To lvwBeds_s.ColumnHeaders.Count
                        If lvwBeds_s.ColumnHeaders(k).Text = "病床" Then Exit For
                    Next
                    If k <> 1 Then
                        lvwBeds_s.ColumnHeaders(1).Text = "病床"
                        lvwBeds_s.ColumnHeaders(k).Text = strTemp
                        lvwBeds_s.ColumnHeaders(1).Key = "_病床1"
                        lvwBeds_s.ColumnHeaders(k).Key = "_" & strTemp
                        lvwBeds_s.ColumnHeaders(1).Key = "_病床"
                    End If
                    
                    '以床位为单位显示,以床号为关键字
                    Set objItem = lvwBeds_s.ListItems.Add(, "_" & !床号, Space(bytLen - Len(!床号)) & !床号 & IIf(IsNull(!姓名), "", ":" & !姓名), strIcon, strIcon)
                   
                    objItem.ForeColor = GetPatiColor(Nvl(mrsBeds!病人类型, "普通病人"))
                    
                    For j = 2 To lvwBeds_s.ColumnHeaders.Count
                        objItem.SubItems(j - 1) = IIf(IsNull(mrsBeds.Fields(lvwBeds_s.ColumnHeaders(j).Text).Value), "", mrsBeds.Fields(lvwBeds_s.ColumnHeaders(j).Text).Value)
                        objItem.ListSubItems(j - 1).ForeColor = objItem.ForeColor
                    Next
                    objItem.Tag = !状态 '用Tag标志床位状态
                    
                    mintBeds_A = mintBeds_A + 1
                End If
                .MoveNext
            Next
            mintChange_A = UBound(Split(Mid(strChange, 2), ",")) + 1
            
            If Not lvwBeds_s.SelectedItem Is Nothing Then
                lvwBeds_s.ListItems(1).Selected = True
                lvwBeds_s.SelectedItem.EnsureVisible
            End If
        End If
    End With
    ReadBedsMap = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ReadNure(lngUnitID As Long) As Boolean
    Dim strTmp As String
    Dim rsNure As New ADODB.Recordset
    Dim i As Integer, lngLen As Integer
    Dim objButton As Button
    
    On Error GoTo errH
    strTmp = "Select c.Id, c.名称, Count(ID) As 数量" & vbNewLine & _
        "From 在院病人 A, 病案主页 B, 收费项目目录 C" & vbNewLine & _
        "Where a.病人id = b.病人id And a.主页id = b.主页id And b.护理等级id = c.Id And a.病区id =[1] And b.状态 In (0, 2, 3)" & vbNewLine & _
        "Group By c.Id, c.名称"
        
    Set rsNure = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, lngUnitID)
    tbrFilter.Buttons.Clear
    Set objButton = tbrFilter.Buttons.Add(, "curDay", "当天入院", , "UnCheck_")
    If rsNure.RecordCount <> 0 Then
        For i = 1 To rsNure.RecordCount
            If LenB(rsNure!名称) > lngLen Then lngLen = LenB(rsNure!名称)
            rsNure.MoveNext
        Next
        rsNure.MoveFirst
    End If
    With rsNure
        If Not .EOF Then
            For i = 1 To .RecordCount
                Set objButton = tbrFilter.Buttons.Add(, "Nurse" & !ID, GetLenText(!名称, lngLen) & "(" & !数量 & ")", , "Check")
                 If i <= 10 Then
                    objButton.ToolTipText = !名称 & "病人(ALT + " & i Mod 10 & ")"
                End If
                .MoveNext
            Next
        End If
    End With
    tbrFilter.Buttons(1).Caption = GetLenText(tbrFilter.Buttons(1).Caption, lngLen)
    ReadNure = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function ReadFamily(lngUnitID As Long) As Boolean
'功能：读取指定病区的家庭病床病人并显示在列表中
'说明：家庭病床床号为空,但又入住了的
    Dim i As Integer, j As Integer, objItem As ListItem
    Dim strChange As String
    Dim strTmp As String
    
    On Error GoTo errH
    
    '附加条件
    strTmp = ""
    gstrSQL = ""
    For i = 1 To tbrFilter.Buttons.Count
        If tbrFilter.Buttons(i).Key Like "Nurse*" And tbrFilter.Buttons(i).Image = "Check" Then
            strTmp = strTmp & "," & Val(Replace(tbrFilter.Buttons(i).Key, "Nurse", ""))
        End If
    Next
    strTmp = strTmp & ","
    gstrSQL = " And (instr([2],','|| B.护理等级ID || ',')>0 Or B.护理等级ID is NULL)"
    
    If tbrFilter.Buttons("curDay").Image = "Check_" Then
        gstrSQL = gstrSQL & " And B.入院日期 Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60"
    End If
    
    '58842,刘鹏飞,2013-02-25,在院病人读取(从在院病人中读取)
    gstrSQL = _
       "Select Nvl(B.状态, 0) As 病人状态, B.出院科室id As 当前科室id, E.名称 As 当前科室, A.病人id, B.住院号, NVL(B.姓名,A.姓名) 姓名, NVL(B.性别,A.性别) 性别," & vbNewLine & _
        "       NVL(B.年龄,A.年龄) 年龄, A.医疗付款方式, A.合同单位id, B.主页id, B.当前病况," & vbNewLine & _
        "       To_Char(B.入院日期, 'YYYY-MM-DD HH24:MI:SS') As 入院时间, B.护理等级id, D.名称 As 护理等级, B.住院医师," & vbNewLine & _
        "       B.病人性质, B.险类,A.就诊卡号,A.身份证号,A.IC卡号,Nvl(B.病人类型,Decode(B.险类,Null,'普通病人','医保病人')) 病人类型 " & vbNewLine & _
        "From 病人信息 A, 病案主页 B, 病人变动记录 C, 收费项目目录 D, 部门表 E,在院病人 F" & vbNewLine & _
        "Where B.病人id = A.病人id And F.病人ID=A.病人ID And C.病人id = B.病人id And C.主页id = B.主页id And B.护理等级id = D.ID(+) And" & vbNewLine & _
        "      B.出院科室id = E.ID And Nvl(B.主页id, 0) <> 0 And Nvl(B.状态, 0) In (0, 2, 3) And" & vbNewLine & _
        "      C.开始时间 Is Not Null And C.终止时间 Is Null And C.床号 Is Null And B.当前病区id+0 = F.病区ID And F.病区ID=[1] And B.出院病床 Is Null" & gstrSQL & vbNewLine & _
        "Order By B.入院日期 Desc, B.住院号 Desc"

    Set mrsFamily = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngUnitID, strTmp)
    Set mrsCFamily = mrsFamily.Clone
    
    mintBeds_B = 0: mintChange_B = 0
    
    With mrsFamily
        If .RecordCount <> 0 Then
            For i = 1 To .RecordCount
                '以病人为单位显示,以病人ID为关键字
                If !病人状态 = 2 Then
                    '转科病人
                    If Nvl(!病人性质, 0) <> 0 Then
                        '留观病人图标
                        Set objItem = lvwFamily_s.ListItems.Add(, "_" & !病人ID, !姓名, "KChange", "KChange")
                    Else
                        Set objItem = lvwFamily_s.ListItems.Add(, "_" & !病人ID, !姓名, "Change", "Change")
                    End If
                    If Not IsNull(!病人ID) Then
                        If InStr(strChange & ",", "," & !病人ID & ",") = 0 Then
                            strChange = strChange & "," & !病人ID
                        End If
                    End If
                ElseIf !病人状态 = 3 Then
                    '预出院病人
                    If Nvl(!病人性质, 0) <> 0 Then
                        Set objItem = lvwFamily_s.ListItems.Add(, "_" & !病人ID, !姓名, "KOut", "KOut")
                    Else
                        Set objItem = lvwFamily_s.ListItems.Add(, "_" & !病人ID, !姓名, "Out", "Out")
                    End If
                Else
                    If Nvl(!病人性质, 0) <> 0 Then
                        Set objItem = lvwFamily_s.ListItems.Add(, "_" & !病人ID, !姓名, "KFamily", "KFamily")
                    Else
                        Set objItem = lvwFamily_s.ListItems.Add(, "_" & !病人ID, !姓名, "Family", "Family")
                    End If
                End If
                
                objItem.ForeColor = GetPatiColor(Nvl(mrsFamily!病人类型, "普通病人"))
                For j = 2 To lvwFamily_s.ColumnHeaders.Count
                    objItem.SubItems(j - 1) = IIf(IsNull(mrsFamily.Fields(lvwFamily_s.ColumnHeaders(j).Text).Value), "", mrsFamily.Fields(lvwFamily_s.ColumnHeaders(j).Text).Value)
                    objItem.ListSubItems(j - 1).ForeColor = objItem.ForeColor
                Next
                mintBeds_B = mintBeds_B + 1
                
                .MoveNext
            Next
            mintChange_B = UBound(Split(Mid(strChange, 2), ",")) + 1

            If Not lvwFamily_s.SelectedItem Is Nothing Then
                lvwFamily_s.ListItems(1).Selected = True
                lvwFamily_s.SelectedItem.EnsureVisible
            End If
        End If
    End With
    ReadFamily = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ReadIn(lngUnitID As Long) As Boolean
'功能：读取指定病区登记为当前病区或尚未登记病区但登记科室属于当前病区的待入科病人,含入院登记病人和转科病人,并显示在列表中
    Dim objItem As ListItem, i As Integer, j As Integer
    Dim strSex As String, strpar1 As String, strpar2 As String
    Dim strDepts As String, lngInTime As Long
    
    On Error GoTo errH
    strDepts = zlDatabase.GetPara("待入科病人科室", glngSys, mlngModul, "")
    If strDepts <> "" Then
        strDepts = "," & strDepts & ","
        strpar1 = " And Instr([2],',' || B.入院科室id || ',')>0 "
        strpar2 = " And Instr([2],',' || C.科室ID || ',')>0 "
    End If
    lngInTime = Val(zlDatabase.GetPara("入院天数", glngSys, mlngModul, 3))
    strpar1 = strpar1 & " And B.入院日期>=" & IIf(lngInTime <> 0, "Sysdate-[3]", "trunc(sysdate)")
    
    '入院病人(状态=1),使用当前病区ID,出院科室ID以便使用索引，不使用入院病区ID,入院科室ID
    '问题29002 by lesfeng 2010-04-09 原And C.科室id = H.科室id 该为 And C.科室id+0 = H.科室id
    '58842,刘鹏飞,2013-02-25,在院病人读取(从在院病人中读取)
    gstrSQL = _
        "Select 0 As 入科标志, A.病人id, B.住院号, NVL(B.姓名,A.姓名) 姓名, NVL(B.性别,A.性别) 性别, NVL(B.年龄,A.年龄) 年龄, B.费别, B.主页id, B.当前病区id, E.名称 As 当前病区," & vbNewLine & _
        "       B.出院科室id, F.名称 As 当前科室, To_Char(B.入院日期, 'YYYY-MM-DD HH24:MI:SS') As 入院时间, B.当前病况," & vbNewLine & _
        "       B.护理等级id, D.名称 As 护理等级, B.出院科室id As 入住科室id, F.名称 As 转入科室, B.责任护士, B.门诊医师," & vbNewLine & _
        "       B.住院医师, B.病人性质, B.险类,A.就诊卡号,A.身份证号,A.IC卡号,Nvl(B.病人类型,Decode(B.险类,Null,'普通病人','医保病人')) 病人类型 " & vbNewLine & _
        "From 病人信息 A, 病案主页 B, 收费项目目录 D, 部门表 E, 部门表 F" & vbNewLine & _
        "Where B.病人id = A.病人id And B.护理等级id = D.ID(+) And B.当前病区id = E.ID(+) And B.出院科室id = F.ID And" & vbNewLine & _
        "      B.出院日期 Is Null And Nvl(B.主页id, 0) <> 0 And B.状态 = 1 And" & vbNewLine & _
        "      (B.当前病区ID+0 = [1] Or B.当前病区ID Is Null And Exists(Select 1 From 病区科室对应 C Where B.出院科室id = C.科室id And C.病区id = [1]))" & strpar1
    '84937:刘鹏飞,性能优化
    '转科病人(存在开始时间为空的入科变动)
    gstrSQL = gstrSQL & vbNewLine & " Union All " & vbNewLine & _
        "Select 1 As 入科标志, A.病人id, B.住院号, NVL(B.姓名,A.姓名) 姓名, NVL(B.性别,A.性别) 性别, NVL(B.年龄,A.年龄) 年龄, B.费别, B.主页id, B.当前病区id, E.名称 As 当前病区," & vbNewLine & _
        "       B.出院科室id, F.名称 As 当前科室, To_Char(B.入院日期, 'YYYY-MM-DD HH24:MI:SS') As 入院时间, B.当前病况," & vbNewLine & _
        "       B.护理等级id, D.名称 As 护理等级, C.科室id As 入住科室id, G.名称 As 转入科室, B.责任护士, B.门诊医师, B.住院医师," & vbNewLine & _
        "       B.病人性质, B.险类,A.就诊卡号,A.身份证号,A.IC卡号,Nvl(B.病人类型,Decode(B.险类,Null,'普通病人','医保病人')) 病人类型 " & vbNewLine & _
        "From 病人信息 A, 病案主页 B, 病人变动记录 C, 收费项目目录 D, 部门表 E, 部门表 F, 部门表 G, 病区科室对应 H" & vbNewLine & _
        "Where A.在院=1 And B.病人id = A.病人id And B.主页ID=A.主页ID And C.病人id = B.病人id And C.主页id = B.主页id And B.护理等级id = D.ID(+) And" & vbNewLine & _
        "      B.当前病区id+0 = E.ID And B.出院科室id+0 = F.ID And Nvl(B.主页id, 0) <> 0 And C.开始原因 = 3 And C.开始时间 Is Null And" & vbNewLine & _
        "      C.终止时间 Is Null And B.状态 = 2 And C.科室id = G.ID And C.科室id+0 = H.科室id And H.病区id = [1] " & strpar2
        
    '转病区病人(存在开始时间为空的入病区变动)
    gstrSQL = gstrSQL & vbNewLine & " Union All " & vbNewLine & _
        "Select 2 As 入科标志, A.病人id, B.住院号, NVL(B.姓名,A.姓名) 姓名, NVL(B.性别,A.性别) 性别, NVL(B.年龄,A.年龄) 年龄, B.费别, B.主页id, B.当前病区id, E.名称 As 当前病区," & vbNewLine & _
        "       B.出院科室id, F.名称 As 当前科室, To_Char(B.入院日期, 'YYYY-MM-DD HH24:MI:SS') As 入院时间, B.当前病况," & vbNewLine & _
        "       B.护理等级id, D.名称 As 护理等级, C.科室id As 入住科室id, G.名称 As 转入科室, B.责任护士, B.门诊医师, B.住院医师," & vbNewLine & _
        "       B.病人性质, B.险类,A.就诊卡号,A.身份证号,A.IC卡号,Nvl(B.病人类型,Decode(B.险类,Null,'普通病人','医保病人')) 病人类型 " & vbNewLine & _
        "From 病人信息 A, 病案主页 B, 病人变动记录 C, 收费项目目录 D, 部门表 E, 部门表 F, 部门表 G, 病区科室对应 H" & vbNewLine & _
        "Where A.在院=1 And B.病人id = A.病人id And B.主页ID=A.主页ID And C.病人id = B.病人id And C.主页id = B.主页id And B.护理等级id = D.ID(+) And" & vbNewLine & _
        "      B.当前病区id+0 = E.ID And B.出院科室id+0 = F.ID And Nvl(B.主页id, 0) <> 0 And C.开始原因 = 15 And C.开始时间 Is Null And" & vbNewLine & _
        "      C.终止时间 Is Null And B.状态 = 2 And C.科室id = G.ID And  C.病区id+0 = H.病区id And C.科室id+0 = H.科室id And H.病区id = [1] " & strpar2 & vbNewLine & _
        "Order By 入科标志 Desc, 入院时间 Desc, 住院号 Desc"
    Set mrsIn = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngUnitID, strDepts, lngInTime)
    Set mrsCIn = mrsIn.Clone
    
    mintIn = 0: mintChange_C = 0
    
    With mrsIn
        If mrsIn.RecordCount <> 0 Then
            For i = 1 To .RecordCount
                If IsNull(!性别) Then
                    strSex = "O"
                Else
                    If InStr(!性别, "男") > 0 Then
                        strSex = "M"
                    ElseIf InStr(!性别, "女") > 0 Then
                        strSex = "F"
                    Else
                        strSex = "O"
                    End If
                End If
                
                '留观病人图标
                If IIf(IsNull(!病人性质), 0, !病人性质) <> 0 Then strSex = "K" & strSex
                
                '以病人ID为关键字
                If !入科标志 = 0 Then
                    Set objItem = lvwIn_s.ListItems.Add(, "_" & !病人ID, !姓名, strSex, strSex)
                ElseIf !入科标志 = 1 Then
                    Set objItem = lvwIn_s.ListItems.Add(, "_" & !病人ID, !姓名, strSex & "_Change", strSex & "_Change")
                    mintChange_C = mintChange_C + 1
                Else
                    Set objItem = lvwIn_s.ListItems.Add(, "_" & !病人ID, !姓名, strSex & "_ChangeUnit", strSex & "_ChangeUnit")
                    'mintChange_C = mintChange_C + 1
                End If
                
                objItem.ForeColor = GetPatiColor(Nvl(mrsIn!病人类型, "普通病人"))
                For j = 2 To lvwIn_s.ColumnHeaders.Count
                    If Not (!入科标志 = 0 And lvwIn_s.ColumnHeaders(j).Text = "转入科室") Then
                        objItem.SubItems(j - 1) = IIf(IsNull(mrsIn.Fields(lvwIn_s.ColumnHeaders(j).Text).Value), "", mrsIn.Fields(lvwIn_s.ColumnHeaders(j).Text).Value)
                        objItem.ListSubItems(j - 1).ForeColor = objItem.ForeColor
                    End If
                Next
                objItem.Tag = !入科标志 '用Tag标志入科病人类别
                
                mintIn = mintIn + 1
                
                .MoveNext
            Next
            
            If Not lvwIn_s.SelectedItem Is Nothing Then
                lvwIn_s.ListItems(1).Selected = True
                lvwIn_s.SelectedItem.EnsureVisible
            End If
        End If
    End With
    ReadIn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadOut(lngUnitID As Long) As Boolean
'功能：读取指定病区出院病人并显示在列表中
    Dim i As Integer, j As Integer, strSex As String
    Dim objItem As ListItem, int结清 As Integer
    Dim lngOutTime As Long, str住院号 As String
    
    '出院病人显示天数
    lngOutTime = Val(zlDatabase.GetPara("出院天数", glngSys, mlngModul, "30"))

    '结清未结清的
    If chk结清(0).Value = 1 And chk结清(1).Value = 1 Then
        int结清 = 0               '都显示
    ElseIf chk结清(0).Value = 0 And chk结清(1).Value = 1 Then
        int结清 = 1               '只显示未结清的
    ElseIf chk结清(0).Value = 1 And chk结清(1).Value = 0 Then
        int结清 = 2              '只显示已结清的
    End If
    
    '50323,刘鹏飞,2012-08-14,判断病人是否已经结清，应该判定某次住院是否存在未结清的费用。
    gstrSQL = " And B.当前病区ID+0=[1] And B.出院日期>=" & IIf(lngOutTime <> 0, "Sysdate-[2]", "trunc(Sysdate)")
    
    '注释原有代码
'    gstrSQL = _
'        "Select A.病人id,NVL(B.姓名,A.姓名) 姓名, NVL(B.性别,A.性别) 性别,A.住院次数,B.住院号, NVL(B.年龄,A.年龄) 年龄, B.费别, B.主页id,B.当前病区ID," & vbNewLine & _
'        "       To_Char(B.入院日期, 'YYYY-MM-DD HH24:MI:SS') As 入院时间," & vbNewLine & _
'        "       To_Char(B.出院日期, 'YYYY-MM-DD HH24:MI:SS') As 出院时间, D.名称 As 出院科室,B.出院科室ID, B.出院病床, B.当前病况 As 出院病况," & vbNewLine & _
'        "       C.名称 As 护理等级, B.出院方式, B.病人性质, B.险类, Decode(Nvl(E.费用余额, 0), 0, '√', Null) As 结清" & vbNewLine & _
'        "       ,A.就诊卡号,A.身份证号,A.IC卡号,Nvl(B.病人类型,Decode(B.险类,Null,'普通病人','医保病人')) 病人类型 " & vbNewLine & _
'        "From 病人信息 A, 病案主页 B, 收费项目目录 C, 部门表 D, (select 病人ID,性质,Nvl(sum(预交余额),0) 预交余额,Nvl(sum(费用余额),0) 费用余额 from 病人余额 group by 病人ID,性质) E" & vbNewLine & _
'        "Where B.病人id = A.病人id And B.出院科室id = D.ID And B.出院日期 Is Not Null And" & vbNewLine & _
'        "      Nvl(B.主页id, 0) <> 0 And B.护理等级id = C.ID(+) And A.病人id = E.病人id(+) And E.性质(+) = 1" & gstrSQL
    '添加新的sql
    '84946：刘鹏飞,SQL优化(将对病人未结费用的查询放在子查询)
    gstrSQL = _
        " Select a.病人id, Nvl(b.姓名, a.姓名) 姓名, Nvl(b.性别, a.性别) 性别, a.住院次数, b.住院号, Nvl(b.年龄, a.年龄) 年龄, b.费别, b.主页id, b.当前病区id," & vbNewLine & _
        "       To_Char(b.入院日期, 'YYYY-MM-DD HH24:MI:SS') As 入院时间, To_Char(b.出院日期, 'YYYY-MM-DD HH24:MI:SS') As 出院时间, d.名称 As 出院科室," & vbNewLine & _
        "       b.出院科室id, b.出院病床, b.当前病况 As 出院病况, c.名称 As 护理等级, b.出院方式, b.病人性质, b.险类," & vbNewLine & _
        "       (Select Decode(Nvl(Sum(金额), 0), 0, '√', Null)" & vbNewLine & _
        "         From 病人未结费用" & vbNewLine & _
        "         Where 来源途径 = 2 And 病人id = b.病人id And 主页id = b.主页id" & vbNewLine & _
        "          ) 结清, a.就诊卡号, a.身份证号, a.Ic卡号, Nvl(b.病人类型, Decode(b.险类, Null, '普通病人', '医保病人')) 病人类型," & vbNewLine & _
        "       a.主页id 就诊次数" & vbNewLine & _
        " From 病人信息 a, 病案主页 b, 收费项目目录 c, 部门表 d" & vbNewLine & _
        " Where b.病人id = a.病人id And b.出院科室id = d.Id And b.出院日期 Is Not Null And Nvl(b.主页id, 0) <> 0 And b.护理等级id = c.Id(+)" & gstrSQL


    gstrSQL = "Select /*+ rule*/ 病人id,姓名,性别,住院次数,住院号, 年龄, 费别, 主页id,当前病区ID,入院时间," & _
                "出院时间 , 出院科室, 出院科室ID, 出院病床, 出院病况, 护理等级, 出院方式, 病人性质, 险类, 结清" & _
                ",就诊卡号,身份证号,IC卡号,病人类型,就诊次数  From (" & gstrSQL & ") Where 1=1" & _
                IIf(int结清 = 0, "", IIf(int结清 = 1, " And 结清 is NULL", " And 结清 is Not NULL")) & _
             " Order by 出院时间 Desc,住院号 Desc"
    
    On Error GoTo errH
    Set mrsOut = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngUnitID, lngOutTime)
    Set mrsCOut = mrsOut.Clone
        
    mintOut = 0
    With mrsOut
        If mrsOut.RecordCount <> 0 Then
            For i = 1 To .RecordCount
                If IsNull(!性别) Then
                    strSex = "O"
                Else
                    If InStr(!性别, "男") > 0 Then
                        strSex = "M"
                    ElseIf InStr(!性别, "女") > 0 Then
                        strSex = "F"
                    Else
                        strSex = "O"
                    End If
                End If
                
                '留观病人图标
                If IIf(IsNull(!病人性质), 0, !病人性质) <> 0 Then strSex = "K" & strSex
                
                '以病人ID 主页为关键字
                Set objItem = lvwOut_s.ListItems.Add(, "_" & !病人ID & "_" & !主页ID, !姓名, strSex, strSex)
                
                objItem.ForeColor = GetPatiColor(Nvl(mrsOut!病人类型, "普通病人"))
                For j = 2 To lvwOut_s.ColumnHeaders.Count
                    objItem.SubItems(j - 1) = IIf(IsNull(mrsOut.Fields(lvwOut_s.ColumnHeaders(j).Text).Value), "", mrsOut.Fields(lvwOut_s.ColumnHeaders(j).Text).Value)
                    objItem.ListSubItems(j - 1).ForeColor = objItem.ForeColor
                Next

                mintOut = mintOut + 1
                
                .MoveNext
            Next
        End If
    End With
    ReadOut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub lvwIn_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lvwIn_s.SelectedItem Is Nothing Then Exit Sub
    If Button = 1 And mblnIn Then
        Set lvwIn_s.DragIcon = lvwIn_s.SelectedItem.CreateDragImage
        lvwIn_s.Drag 1
    End If
End Sub

Private Sub lvwOut_s_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Static objIcon As IPictureDisp
    If Source Is lvwIn_s Or InStr(mstrPrivs, "病人出院") = 0 Then   '入科病人不能拖到出院列表
        If State = 0 Then
            Set objIcon = Source.DragIcon
        ElseIf State = 2 Then
            Set Source.DragIcon = img32.ListImages("Limit").Picture
        ElseIf State = 1 Then
            Set Source.DragIcon = objIcon
        End If
    End If
End Sub

Private Sub lvwOut_s_DragDrop(Source As Control, X As Single, Y As Single)
    If (Source Is lvwBeds_s Or Source Is lvwFamily_s) And InStr(mstrPrivs, "病人出院") > 0 Then
        '病人出院处理
        mnuEdit_Out_Click
    End If
End Sub

Private Sub lvwIn_s_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Static objIcon As IPictureDisp
    '任何对象不许拖入入科病人列表
    If Not Source Is lvwIn_s Then
        If State = 0 Then
            Set objIcon = Source.DragIcon
        ElseIf State = 2 Then
            Set Source.DragIcon = img32.ListImages("Limit").Picture
        ElseIf State = 1 Then
            Set Source.DragIcon = objIcon
        End If
    End If
End Sub

Private Sub lvwBeds_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And mblnBeds Then
        '空床不操作,转科病人不能出院或换床
        If lvwBeds_s.SelectedItem.Tag <> "占用" Or lvwBeds_s.SelectedItem.Icon Like "*Change" Then Exit Sub
        If IIf(IsNull(mrsBeds!性别), "", mrsBeds!性别) = "男" Then
            Set lvwBeds_s.DragIcon = img32.ListImages("M").Picture
        ElseIf IIf(IsNull(mrsBeds!性别), "", mrsBeds!性别) = "女" Then
            Set lvwBeds_s.DragIcon = img32.ListImages("F").Picture
        Else
            Set lvwBeds_s.DragIcon = img32.ListImages("O").Picture
        End If
        lvwBeds_s.Drag 1
    End If
End Sub

Private Sub lvwBeds_s_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Dim objOver As ListItem
    
    If Source Is lvwIn_s Then
        Set objOver = lvwBeds_s.HitTest(X, Y)
        If Not objOver Is Nothing Then
            mrsCBeds.Filter = "床号='" & Mid(objOver.Key, 2) & "'"
            
            If objOver.Tag = "空床" And (mrsCBeds!性别分类 = "不限床" Or _
            (mrsCBeds!性别分类 = "男床" And IIf(IsNull(mrsIn!性别), "", mrsIn!性别) = "男") _
            Or (mrsCBeds!性别分类 = "女床" And IIf(IsNull(mrsIn!性别), "", mrsIn!性别) = "女")) Then
                Set lvwBeds_s.DropHighlight = objOver
                lvwBeds_s.DropHighlight.EnsureVisible
            ElseIf mrsCBeds!病人ID = mrsIn!病人ID _
                And mrsCBeds!共用 = 1 And objOver.Tag <> "空床" Then '当前床位属共用床，并且是病人原住床位，原住床位表明当前公共病区，因为待入科病人只出现在目标科室和公共病区
                Set lvwBeds_s.DropHighlight = objOver
                lvwBeds_s.DropHighlight.EnsureVisible
            End If
        Else
            Set lvwBeds_s.DropHighlight = Nothing
        End If
    ElseIf Source Is lvwBeds_s Then
        Set objOver = lvwBeds_s.HitTest(X, Y)
        If Not objOver Is Nothing And InStr(mstrPrivs, "换床") <> 0 Then
            mrsCBeds.Filter = "床号='" & Mid(objOver.Key, 2) & "'"
            'objOver.Tag = "空床" And
            If mrsBeds!共用 = 1 Then
                If (mrsCBeds!科室ID = mrsBeds!科室ID Or IsNull(mrsCBeds!科室ID) Or mrsCBeds!共用 = 1) _
                    And Nvl(mrsBeds!病人状态, 0) <> 3 And Nvl(mrsCBeds!病人状态, 0) <> 3 And Nvl(mrsCBeds!病人状态, 0) <> 2 Then

                    If mrsBeds!性别分类 = "不限床" Then
                        If Not (mrsCBeds!性别分类 = "不限床" _
                                Or (mrsCBeds!性别分类 = "男床" And IIf(IsNull(mrsBeds!性别), "", mrsBeds!性别) = "男") _
                                Or (mrsCBeds!性别分类 = "女床" And IIf(IsNull(mrsBeds!性别), "", mrsBeds!性别) = "女")) Then
                           Set lvwBeds_s.DropHighlight = Nothing: Exit Sub
                        End If
                    ElseIf mrsBeds!性别分类 = "男床" Then
                        If Not ((mrsCBeds!性别分类 = "不限床" And mrsCBeds!性别 = "男") _
                                Or (mrsCBeds!性别分类 = "男床" And IIf(IsNull(mrsBeds!性别), "", mrsBeds!性别) = "男")) Then
                           Set lvwBeds_s.DropHighlight = Nothing: Exit Sub
                        End If
                    ElseIf mrsBeds!性别分类 = "女床" Then
                        If Not ((mrsCBeds!性别分类 = "不限床" And mrsCBeds!性别 = "女") _
                                Or (mrsCBeds!性别分类 = "女床" And IIf(IsNull(mrsBeds!性别), "", mrsBeds!性别) = "女")) Then
                           Set lvwBeds_s.DropHighlight = Nothing: Exit Sub
                        End If
                    End If
                    
                    Set lvwBeds_s.DropHighlight = objOver
                    lvwBeds_s.DropHighlight.EnsureVisible
                    
                End If
            Else
                If (mrsCBeds!科室ID = mrsBeds!科室ID Or IsNull(mrsCBeds!科室ID) Or (mrsCBeds!共用 = 1 And mrsCBeds!科室ID = mrsBeds!科室ID)) _
                    And Nvl(mrsBeds!病人状态, 0) <> 3 And Nvl(mrsCBeds!病人状态, 0) <> 3 And Nvl(mrsCBeds!病人状态, 0) <> 2 Then
                    
                    If mrsBeds!性别分类 = "不限床" Then
                        If Not (mrsCBeds!性别分类 = "不限床" _
                                Or (mrsCBeds!性别分类 = "男床" And IIf(IsNull(mrsBeds!性别), "", mrsBeds!性别) = "男") _
                                Or (mrsCBeds!性别分类 = "女床" And IIf(IsNull(mrsBeds!性别), "", mrsBeds!性别) = "女")) Then
                           Set lvwBeds_s.DropHighlight = Nothing: Exit Sub
                        End If
                    ElseIf mrsBeds!性别分类 = "男床" Then
                        If Not (mrsCBeds!性别分类 = "不限床" _
                                Or (mrsCBeds!性别分类 = "男床" And IIf(IsNull(mrsBeds!性别), "", mrsBeds!性别) = "男")) Then
                           Set lvwBeds_s.DropHighlight = Nothing: Exit Sub
                        End If
                    ElseIf mrsBeds!性别分类 = "女床" Then
                        If Not (mrsCBeds!性别分类 = "不限床" _
                                Or (mrsCBeds!性别分类 = "女床" And IIf(IsNull(mrsBeds!性别), "", mrsBeds!性别) = "女")) Then
                           Set lvwBeds_s.DropHighlight = Nothing: Exit Sub
                        End If
                    End If
                    
                    Set lvwBeds_s.DropHighlight = objOver
                    lvwBeds_s.DropHighlight.EnsureVisible
                
                End If
            End If
        Else
            Set lvwBeds_s.DropHighlight = Nothing
        End If
    ElseIf Source Is lvwFamily_s Then
        Set objOver = lvwBeds_s.HitTest(X, Y)
        If Not objOver Is Nothing Then
            mrsCBeds.Filter = "床号='" & Mid(objOver.Key, 2) & "'"
        
            If objOver.Tag = "空床" And (mrsCBeds!科室ID = mrsFamily!当前科室id Or IsNull(mrsCBeds!科室ID)) _
                And Nvl(mrsFamily!病人状态, 0) <> 3 And (mrsCBeds!性别分类 = "不限床" _
                Or (mrsCBeds!性别分类 = "男床" And IIf(IsNull(mrsFamily!性别), "", mrsFamily!性别) = "男") _
                Or (mrsCBeds!性别分类 = "女床" And IIf(IsNull(mrsFamily!性别), "", mrsFamily!性别) = "女")) Then
                Set lvwBeds_s.DropHighlight = objOver
                lvwBeds_s.DropHighlight.EnsureVisible
            End If
        Else
            Set lvwBeds_s.DropHighlight = Nothing
        End If
    End If
    If State = 1 Then Set lvwBeds_s.DropHighlight = Nothing
End Sub

Private Sub lvwBeds_s_DragDrop(Source As Control, X As Single, Y As Single)
    Dim str目标床号 As String
    If Source Is lvwIn_s And Not lvwBeds_s.DropHighlight Is Nothing Then
        Set lvwBeds_s.SelectedItem = lvwBeds_s.DropHighlight
        Set lvwBeds_s.DropHighlight = Nothing
        '病人入入住处理(双方的选中项)
        
        Call lvwBeds_s_ItemClick(lvwBeds_s.SelectedItem)
        
        If mrsIn!入科标志 = 2 Then
            Call mnuEdit_InUnit_Click
        Else
            Call mnuEdit_In_Click
        End If
        
    ElseIf (Source Is lvwFamily_s Or Source Is lvwBeds_s) And Not lvwBeds_s.DropHighlight Is Nothing Then
        '病人换床处理
        If InStr(mstrPrivs, "换床") = 0 Then Set lvwBeds_s.DropHighlight = Nothing: Exit Sub
        str目标床号 = Trim(Mid(lvwBeds_s.DropHighlight.Key, 2))
        
        mrsCBeds.Filter = "床号='" & str目标床号 & "'"
        
        If Nvl(mrsCBeds!病人ID, 0) = 0 Then
            Set lvwBeds_s.DropHighlight = Nothing
            Call ChangeBeds(0, str目标床号)
        Else
            Set lvwBeds_s.DropHighlight = Nothing
            Call SwapBeds(str目标床号)
        End If
    End If
End Sub

Private Sub lvwFamily_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And mblnFamily Then
        If lvwFamily_s.SelectedItem Is Nothing Then Exit Sub
        If lvwFamily_s.SelectedItem.Icon Like "*Change" Then Exit Sub
        
        If IIf(IsNull(mrsFamily!性别), "", mrsFamily!性别) = "男" Then
            Set lvwFamily_s.DragIcon = img32.ListImages("M").Picture
        ElseIf IIf(IsNull(mrsFamily!性别), "", mrsFamily!性别) = "女" Then
            Set lvwFamily_s.DragIcon = img32.ListImages("F").Picture
        Else
            Set lvwFamily_s.DragIcon = img32.ListImages("O").Picture
        End If
        lvwFamily_s.Drag 1
    End If
End Sub

Private Sub lvwFamily_s_DragDrop(Source As Control, X As Single, Y As Single)
    If Source Is lvwIn_s Then
        '病人入科处理(家庭病床)
        Dim byt入科方式 As Byte, lng病人ID As Long, lng主页ID As Long
        
        byt入科方式 = Val(lvwIn_s.SelectedItem.Tag)
        lng病人ID = mrsIn!病人ID
        lng主页ID = mrsIn!主页ID
        Call ExecPatiChange(EFun.E入科, Me, mstrPrivs, mlngUnit, lng病人ID, lng主页ID, "家庭病床", 0, byt入科方式)
        
        If gblnOK Then Call LoadList(True, True, True, False)
        
    ElseIf Source Is lvwBeds_s And InStr(1, mstrPrivs, "家庭病床") > 0 And InStr(mstrPrivs, "换床") Then
        If Nvl(mrsBeds!病人状态, 0) <> 3 Then
            '病人换床处理
            
            Call ChangeBeds(0, "家庭病床")
        Else
            MsgBox "预出院病人不能进行换床操作!"
        End If
    End If
End Sub

Private Sub mnuFile_Excel_Click()
    If mobjLVW.ListItems.Count > 100 Then
        If MsgBox("输出到Excel的数据过多,这将耗费许多时间,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Call OutputList(3)
End Sub

Private Sub mnuFile_PreView_Click()
    Call OutputList(2)
End Sub

Private Sub mnuFile_Print_Click()
    Call OutputList(1)
End Sub

Private Sub mnuFile_PrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub OutputList(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrintLvw
    Dim bytR As Byte
    
    On Error GoTo errH
    
    '表头
    Select Case mobjLVW.Name
        Case "lvwBeds_s"
            objOut.Title.Text = "床位映象表"
        Case "lvwFamily_s"
            objOut.Title.Text = "家庭病床表"
        Case "lvwIn_s"
            objOut.Title.Text = "入科病人表"
        Case "lvwOut_s"
            objOut.Title.Text = "出院病人表"
    End Select
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    objOut.UnderAppItems.Add "病区:" & zlCommFun.GetNeedName(cboUnit.Text)
    objOut.BelowAppItems.Add "打印人：" & UserInfo.姓名
    objOut.BelowAppItems.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    
    '表体
    Set objOut.Body.objData = mobjLVW
    
    '输出
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrViewLvw objOut, bytR
    Else
        zlPrintOrViewLvw objOut, bytStyle
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub LoadList(Optional mblnBeds As Boolean = True, Optional mblnFamily As Boolean = True, _
    Optional mblnIn As Boolean = True, Optional mblnOut As Boolean = True, Optional mblnReadNure As Boolean)
'功能：刷新界面列表数据
'参数：缺省刷新所有列表,可以分别指定
    Dim strBeds As String, strFamily As String
    Dim strIn As String, strOut As String, lngUnit As Long
    Dim objFind As ListItem
    
    '记录原位置，清除所有列表
    If mblnBeds Then
        If Not lvwBeds_s.SelectedItem Is Nothing Then strBeds = lvwBeds_s.SelectedItem.Key
        lvwBeds_s.ListItems.Clear
    End If
    If mblnFamily Then
        If Not lvwFamily_s.SelectedItem Is Nothing Then strFamily = lvwFamily_s.SelectedItem.Key
        lvwFamily_s.ListItems.Clear
    End If
    If mblnIn Then
        If Not lvwIn_s.SelectedItem Is Nothing Then strIn = lvwIn_s.SelectedItem.Key
        lvwIn_s.ListItems.Clear
    End If
    If mblnOut Then
        If Not lvwOut_s.SelectedItem Is Nothing Then strOut = lvwOut_s.SelectedItem.Key
        lvwOut_s.ListItems.Clear
    End If
    
    If mblnReadNure = True Then
        '入住、撤销入住、撤销护理等级、调整护理等级、出院时刷新护理等级
        If Not ReadNure(mlngUnit) Then Exit Sub
    End If
    '刷新床位映象表
    If mblnBeds Then Call ReadBedsMap(mlngUnit)
    
    '刷新家庭病床表
    If mblnFamily Then Call ReadFamily(mlngUnit)
    
    '刷新入科病人表
    If mblnIn Then Call ReadIn(mlngUnit)
        
    '刷新出院病人表
    If mblnOut Then Call ReadOut(mlngUnit)
    
    '自动定位到以前位置
    On Error Resume Next
    
    If strBeds <> "" And lvwBeds_s.ListItems.Count > 0 And mblnBeds Then
        Set objFind = lvwBeds_s.ListItems(strBeds)
        If Err.Number = 0 Then
            objFind.Selected = True
            objFind.EnsureVisible
        Else
            Err.Clear
        End If
    End If
    If strFamily <> "" And lvwFamily_s.ListItems.Count > 0 And mblnFamily Then
        Set objFind = lvwFamily_s.ListItems(strFamily)
        If Err.Number = 0 Then
            objFind.Selected = True
            objFind.EnsureVisible
        Else
            Err.Clear
        End If
    End If
    If strIn <> "" And lvwIn_s.ListItems.Count > 0 And mblnIn Then
        Set objFind = lvwIn_s.ListItems(strIn)
        If Err.Number = 0 Then
            objFind.Selected = True
            objFind.EnsureVisible
        Else
            Err.Clear
        End If
    End If
    If strOut <> "" And lvwOut_s.ListItems.Count > 0 And mblnOut Then
        Set objFind = lvwOut_s.ListItems(strOut)
        If Err.Number = 0 Then
            objFind.Selected = True
            objFind.EnsureVisible
        Else
            Err.Clear
        End If
    End If
    
    '83992:每次根据选择项定位记录集中的数据项
    If Not lvwBeds_s.SelectedItem Is Nothing Then mrsBeds.Filter = "床号='" & Mid(lvwBeds_s.SelectedItem.Key, 2) & "'"
    If Not lvwFamily_s.SelectedItem Is Nothing Then mrsFamily.Filter = "病人ID=" & Mid(lvwFamily_s.SelectedItem.Key, 2)
    If Not lvwIn_s.SelectedItem Is Nothing Then mrsIn.Filter = "病人ID=" & Mid(lvwIn_s.SelectedItem.Key, 2)
    If Not lvwOut_s.SelectedItem Is Nothing Then mrsOut.Filter = "病人ID=" & Split(lvwOut_s.SelectedItem.Key, "_")(1) & " and 主页id=" & Split(lvwOut_s.SelectedItem.Key, "_")(2)
    
    If mobjLVW Is lvwBeds_s Then
        Call lvwBeds_s_ItemClick(lvwBeds_s.SelectedItem)
    ElseIf mobjLVW Is lvwFamily_s Then
        Call lvwFamily_s_ItemClick(lvwFamily_s.SelectedItem)
    ElseIf mobjLVW Is lvwIn_s Then
        Call lvwIn_s_ItemClick(lvwIn_s.SelectedItem)
    ElseIf mobjLVW Is lvwOut_s Then
        Call lvwOut_s_ItemClick(lvwOut_s.SelectedItem)
    Else
        If Not lvwBeds_s.SelectedItem Is Nothing Then Call lvwBeds_s_ItemClick(lvwBeds_s.SelectedItem)
    End If
    If Me.Visible Then mobjLVW.SetFocus
End Sub

Private Sub SetMenu()
'功能：根据当前选择列表或床位，设置相应菜单功能的状态。
    Dim lng病人ID As Long, lng主页ID As Long, lng险类 As Long, blnDo As Boolean
    Dim rsTmp As ADODB.Recordset
    
    '住院留察病人转为住院病人
    blnDo = True
    If mobjLVW Is lvwIn_s Or mobjLVW Is lvwOut_s Then
        blnDo = False
    Else
        If mobjLVW Is lvwBeds_s Then
            If lvwBeds_s.SelectedItem Is Nothing Then
                blnDo = False
            ElseIf lvwBeds_s.SelectedItem.Tag <> "占用" Then
                blnDo = False
            ElseIf lvwBeds_s.SelectedItem.Icon Like "*Change" Then
                blnDo = False
            End If
        Else
            If lvwFamily_s.SelectedItem Is Nothing Then
                blnDo = False
            ElseIf lvwFamily_s.SelectedItem.Icon Like "*Change" Then
                blnDo = False
            End If
        End If
        If blnDo Then
            If mobjLVW Is lvwBeds_s Then
                blnDo = IIf(IsNull(mrsBeds!病人性质), 0, mrsBeds!病人性质) = 2
            ElseIf mobjLVW Is lvwFamily_s Then
                blnDo = IIf(IsNull(mrsFamily!病人性质), 0, mrsFamily!病人性质) = 2
            End If
        End If
    End If
    mnuEditToInPati.Enabled = blnDo
    
    '病人入科(入科时选择床号,而不根据界面情况决定)
    blnDo = True
    If Not mobjLVW Is lvwIn_s Then
        blnDo = False
        mnuEdit_In.Enabled = blnDo
    Else
        If lvwIn_s.SelectedItem Is Nothing Then
            blnDo = False
            mnuEdit_In.Enabled = blnDo
        Else
            mnuEdit_In.Enabled = blnDo
        End If
    End If
    tbr.Buttons("In").Enabled = blnDo
    
    '病人备注编辑
    blnDo = True
    If mobjLVW Is lvwBeds_s Then
        If lvwBeds_s.SelectedItem Is Nothing Then
            blnDo = False
        ElseIf lvwBeds_s.SelectedItem.Tag <> "占用" Then
            blnDo = False
        End If
    ElseIf mobjLVW Is lvwFamily_s Then
        If lvwFamily_s.SelectedItem Is Nothing Then blnDo = False
    ElseIf mobjLVW Is lvwOut_s Then
        If lvwOut_s.SelectedItem Is Nothing Then blnDo = False
    ElseIf mobjLVW Is lvwIn_s Then
        If lvwIn_s.SelectedItem Is Nothing Then blnDo = False
    End If
    mnuEdit_Memo.Enabled = blnDo
    
    '病人转科、病人换床、预出院、护理等级、调整信息、新生儿登记(转科和预出院状态不许)
    blnDo = True
    If mobjLVW Is lvwIn_s Or mobjLVW Is lvwOut_s Then
        blnDo = False
    Else
        If mobjLVW Is lvwBeds_s Then
            If lvwBeds_s.SelectedItem Is Nothing Then
                blnDo = False
            ElseIf lvwBeds_s.SelectedItem.Tag <> "占用" Then
                blnDo = False
            ElseIf lvwBeds_s.SelectedItem.Icon Like "*Change" Then
                blnDo = False
            ElseIf lvwBeds_s.SelectedItem.Icon Like "*Out" Then
                blnDo = False
            End If
        Else
            If lvwFamily_s.SelectedItem Is Nothing Then
                blnDo = False
            ElseIf lvwFamily_s.SelectedItem.Icon Like "*Change" Then
                blnDo = False
            ElseIf lvwFamily_s.SelectedItem.Icon Like "*Out" Then
                blnDo = False
            End If
        End If
    End If
    mnuEdit_Change.Enabled = blnDo
    mnuEdit_ChangeUnit.Enabled = blnDo
    tbr.Buttons("Change").Enabled = blnDo
    
    mnuEdit_ChangeGroup.Enabled = blnDo

    mnuEdit_Move.Enabled = blnDo
    mnuEdit_Swap.Enabled = blnDo
    tbr.Buttons("Move").Enabled = blnDo
    mnuEdit_AddBeds.Enabled = blnDo
    
    mnuEdit_PreOut.Enabled = blnDo
    mnuEdit_Nurse.Enabled = blnDo
    
    mnuEdit_Adjust.Enabled = blnDo
    tbr.Buttons("Adjust").Enabled = blnDo
    
    '新生儿登记(产科病人才允许)
    If blnDo Then
        If mobjLVW Is lvwBeds_s Then
            blnDo = is产科(mrsBeds!当前科室id, Nothing)
            blnDo = blnDo And mrsBeds!性别 = "女"
        ElseIf mobjLVW Is lvwFamily_s Then
            blnDo = is产科(mrsFamily!当前科室id, Nothing)
            blnDo = blnDo And mrsFamily!性别 = "女"
        End If
    End If
    mnuEdit_BabyReg.Enabled = blnDo
    
    '重算费用,修改出院时间
    blnDo = True
    If Not mobjLVW Is Nothing Then
        If mobjLVW.SelectedItem Is Nothing Then
            blnDo = False
        Else
            If mobjLVW Is lvwBeds_s Then
                If lvwBeds_s.SelectedItem.Tag <> "占用" Then blnDo = False
                Set rsTmp = mrsBeds
            ElseIf mobjLVW Is lvwFamily_s Then
                Set rsTmp = mrsFamily
            ElseIf mobjLVW Is lvwOut_s Then
                If Nvl(mrsOut!主页ID, 0) <> Nvl(mrsOut!就诊次数, 0) Then blnDo = False
                Set rsTmp = mrsOut
            ElseIf mobjLVW Is lvwIn_s Then
                Set rsTmp = mrsIn
            End If
            If Nvl(rsTmp!险类, 0) <> 0 Then blnDo = False
        End If
    Else
        blnDo = False
    End If
    mnuEdit_Recalc.Enabled = blnDo
    '问题27392 by lesfeng 2010-01-14
    If mobjLVW Is lvwOut_s Then
        mnuEdit_ModifOut.Enabled = blnDo
    Else
        mnuEdit_ModifOut.Enabled = False
    End If
    
    '病人出院(转科状态不许)
    blnDo = True
    If mobjLVW Is lvwIn_s Or mobjLVW Is lvwOut_s Then
        blnDo = False
    Else
        If mobjLVW Is lvwBeds_s Then
            If lvwBeds_s.SelectedItem Is Nothing Then
                blnDo = False
            ElseIf lvwBeds_s.SelectedItem.Tag <> "占用" Then
                blnDo = False
            ElseIf lvwBeds_s.SelectedItem.Icon Like "*Change" Then
                blnDo = False
            End If
        Else
            If lvwFamily_s.SelectedItem Is Nothing Then
                blnDo = False
            ElseIf lvwFamily_s.SelectedItem.Icon Like "*Change" Then
                blnDo = False
            End If
        End If
    End If
    mnuEdit_Out.Enabled = blnDo
    tbr.Buttons("Out").Enabled = blnDo
    
    '床位等级(转科状态和预出院不许)
    blnDo = True
    If Not mobjLVW Is lvwBeds_s Then
        blnDo = False
    Else
        If lvwBeds_s.SelectedItem Is Nothing Then
            blnDo = False
        ElseIf lvwBeds_s.SelectedItem.Tag <> "占用" Then
            blnDo = False
        ElseIf lvwBeds_s.SelectedItem.Icon Like "*Change" Then
            blnDo = False
        ElseIf lvwBeds_s.SelectedItem.Icon Like "*Out" Then
            blnDo = False
        End If
    End If
    mnuEdit_Level.Enabled = blnDo
    
    '转科记录、床位记录、护理记录
    blnDo = True
    If mobjLVW Is lvwIn_s Then
        blnDo = False
    Else
        If mobjLVW Is lvwFamily_s Then
            If lvwFamily_s.SelectedItem Is Nothing Then
                blnDo = False
            End If
        ElseIf mobjLVW Is lvwOut_s Then
            If lvwOut_s.SelectedItem Is Nothing Then
                blnDo = False
            End If
        Else
            If lvwBeds_s.SelectedItem Is Nothing Then
                blnDo = False
            ElseIf lvwBeds_s.SelectedItem.Tag <> "占用" Then
                blnDo = False
            End If
        End If
    End If
    mnuQuery_Log.Enabled = blnDo

    '打印病案、病人信息、费用信息(所有类型病人均可)
    blnDo = True
    If mobjLVW Is lvwBeds_s Then
        If lvwBeds_s.SelectedItem Is Nothing Then
            blnDo = False
        ElseIf lvwBeds_s.SelectedItem.Tag <> "占用" Then
            blnDo = False
        End If
    ElseIf mobjLVW Is lvwFamily_s Then
        If lvwFamily_s.SelectedItem Is Nothing Then
            blnDo = False
        End If
    ElseIf mobjLVW Is lvwOut_s Then
        If lvwOut_s.SelectedItem Is Nothing Then
            blnDo = False
        End If
    ElseIf mobjLVW Is lvwIn_s Then
        If lvwIn_s.SelectedItem Is Nothing Then
            blnDo = False
        End If
    End If
    mnuFilePrintMed.Enabled = blnDo
    mnuQueryInfo.Enabled = blnDo
    
    '49854:刘鹏飞,2013-10-31,病人腕带打印
    '打印病人腕带(出院病人不可)
    mnuFile_PrintWristlet.Visible = (InStr(mstrPrivs, ";腕带打印;") > 0)
    mnuFile_PrintWristlet.Enabled = mnuFile_PrintWristlet.Visible And blnDo And (Not mobjLVW Is lvwOut_s)
    
    '病种选择
    blnDo = True
    If mobjLVW Is lvwIn_s Then
        blnDo = False
    ElseIf mobjLVW Is lvwOut_s Then
        If lvwOut_s.SelectedItem Is Nothing Then
            blnDo = False
        ElseIf Nvl(mrsOut!险类, 0) = 0 Then
            blnDo = False
        End If
    ElseIf mobjLVW Is lvwBeds_s Then
        If lvwBeds_s.SelectedItem Is Nothing Then
            blnDo = False
        ElseIf lvwBeds_s.SelectedItem.Tag <> "占用" Then
            blnDo = False
        ElseIf Nvl(mrsBeds!险类, 0) = 0 Then
            blnDo = False
        End If
    ElseIf mobjLVW Is lvwFamily_s Then
        If lvwFamily_s.SelectedItem Is Nothing Then
            blnDo = False
        ElseIf Nvl(mrsFamily!险类, 0) = 0 Then
            blnDo = False
        End If
    End If
    mnuEdit_Disease.Enabled = blnDo
    
    '撤消状态设置
    blnDo = True
    If mobjLVW Is lvwBeds_s Then
        If lvwBeds_s.SelectedItem Is Nothing Then
            blnDo = False
        ElseIf lvwBeds_s.SelectedItem.Tag <> "占用" Then
            blnDo = False
        Else
            lng病人ID = mrsBeds!病人ID
            lng主页ID = mrsBeds!主页ID
        End If
    ElseIf mobjLVW Is lvwFamily_s Then
        If lvwFamily_s.SelectedItem Is Nothing Then
            blnDo = False
        Else
            lng病人ID = mrsFamily!病人ID
            lng主页ID = mrsFamily!主页ID
        End If
    ElseIf mobjLVW Is lvwOut_s Then
        If lvwOut_s.SelectedItem Is Nothing Then
            blnDo = False
        Else
            lng病人ID = mrsOut!病人ID
            lng主页ID = mrsOut!主页ID
            If Nvl(mrsOut!主页ID, 0) <> Nvl(mrsOut!就诊次数, 0) Then blnDo = False
        End If
    ElseIf mobjLVW Is lvwIn_s Then
        blnDo = False
    End If
    tbr.Buttons("Undo").ButtonMenus.Clear
    mnuEdit_Undo.Caption = "撤消(&U)"
    tbr.Buttons("Undo").Enabled = False
    mnuEdit_Undo.Enabled = False
    If lng病人ID > 0 And lng主页ID > 0 And blnDo Then Call SetUndoLog(lng病人ID, lng主页ID)
    
    '查找
    If lvwBeds_s.ListItems.Count > 0 Or lvwIn_s.ListItems.Count > 0 Or lvwFamily_s.ListItems.Count > 0 Or lvwOut_s.ListItems.Count > 0 Then
        mnuViewFind.Enabled = True
        mnuViewFindNext.Enabled = (mstrSeekKey = "姓名" Or mstrSeekKey = "住院号")
    Else
        mnuViewFind.Enabled = False
        mnuViewFindNext.Enabled = False
    End If
    
End Sub

Private Sub SetUndoLog(lng病人ID As Long, lng主页ID As Long)
'功能：根据病人类型显示病人可撤消操作
'说明：1.调用该函数之前设置功能的初始状态
'      2.不能撤消入院(入院同时入科的撤消到入院状态)
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, blnExist As Boolean
    Dim objMenu As ButtonMenu
    
    Set rsTmp = GetPatiLog(lng病人ID, lng主页ID)
    If rsTmp Is Nothing Then Exit Sub
    
    mnuEdit_Undo.Enabled = True
    tbr.Buttons("Undo").Enabled = True
        
    '如果是出院,则多加一条撤消出院
    If Not IsNull(rsTmp!终止时间) And rsTmp!终止原因 = 1 Then
        Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "出院")
        If tbr.Buttons("Undo").ButtonMenus.Count > 1 Then objMenu.Enabled = False
        mnuEdit_Undo.Caption = "撤消出院(&U)"
        blnExist = True
        
        If InStr(";" & mstrPrivs & ";", ";撤消出院;") = 0 Then
            objMenu.Enabled = False
            mnuEdit_Undo.Enabled = False
        End If
    End If
    '问题28386 by lesfeng 2010-03-06 调整开始原因为2\3的入科分别为入院入科\转科入科
    For i = 1 To rsTmp.RecordCount
        If IsNull(rsTmp!开始时间) And rsTmp!开始原因 = 3 Then
            Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "转科")
        ElseIf IsNull(rsTmp!开始时间) And rsTmp!开始原因 = 15 Then
            Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "转病区")
        Else
            Select Case rsTmp!开始原因
                Case 1 '入院
                    '非lvwIN中的病人当前可撤消的为入院变动则一定是入院同时入科,处理为撤消入科
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "入院入住")
                    '不是当前可撤消的入院变动,则为单独的入院登记
                    If tbr.Buttons("Undo").ButtonMenus.Count > 1 Then objMenu.Visible = False
                Case 2 '入院入科
'                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "入科")
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "入住")
                Case 3 '转科入科
'                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "入科")
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "转科入住")
                Case 4
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "换床")
                Case 5
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "床位等级变动")
                Case 6
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "护理等级变动")
                Case 7
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "经治医师改变")
                Case 8
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "责任护士改变")
                Case 9
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "转为住院病人")
                Case 10
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "预出院")
                Case 11
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "主治医师变动")
                Case 12
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "主任医师变动")
                Case 13
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "病况变动")
                Case 14
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "转医疗小组")
                Case 15
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "转病区入住")
            End Select
        End If
        If tbr.Buttons("Undo").ButtonMenus.Count > 1 Then objMenu.Enabled = False
        If Not blnExist And i = 1 Then mnuEdit_Undo.Caption = "撤消" & objMenu.Text & "(&U)"
        
        If InStr(mstrPrivs, "撤消入科") = 0 And (objMenu.Text = "入住" Or objMenu.Text = "入院入住" Or objMenu.Text = "转科入住" Or objMenu.Text = "转病区入住") Then
            objMenu.Enabled = False
        End If
        
        If InStr(mstrPrivs, "换床") = 0 And (objMenu.Text = "换床") Then
            objMenu.Enabled = False
        End If
        
        If InStr(mstrPrivs, "住院留观转住院") = 0 And objMenu.Text = "转为住院病人" Then
            objMenu.Enabled = False
        End If
        
        If InStr(mstrPrivs, "撤销预出院") = 0 And objMenu.Text = "预出院" Then
            objMenu.Enabled = False
        End If
        
        rsTmp.MoveNext
    Next
    
    If tbr.Buttons("Undo").ButtonMenus(1).Enabled = False Then '奇怪,用Not方式无效
        mnuEdit_Undo.Enabled = False
    End If
End Sub
Private Function InitPatiType() As Boolean
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errH
    mstrPatiTypeColor = ""
    gstrSQL = "select 名称,颜色 from 病人类型"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人类型")
    Do Until rsTemp.EOF
        mstrPatiTypeColor = mstrPatiTypeColor & rsTemp!名称 & "," & Nvl(rsTemp!颜色, 0) & "|"
        rsTemp.MoveNext
    Loop
    If Len(mstrPatiTypeColor) > 0 Then
        mstrPatiTypeColor = Mid(mstrPatiTypeColor, 1, Len(mstrPatiTypeColor) - 1)
    Else
        mstrPatiTypeColor = "普通病人,0|医保病人,255"
    End If
    InitPatiType = True
    Exit Function
errH:
    mstrPatiTypeColor = "普通病人,0|医保病人,255"
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function GetPatiColor(ByVal strPatiType) As Long
Dim arrType As Variant, i As Integer
    arrType = Split(mstrPatiTypeColor, "|")
    For i = LBound(arrType) To UBound(arrType)
        If Split(arrType(i), ",")(0) = strPatiType Then
            GetPatiColor = Split(arrType(i), ",")(1)
            Exit Function
        End If
    Next
End Function
Private Function GetLenText(ByVal strText As String, ByVal lngLen As Long) As String
'参数：将文字右添空格到指定长度
    Dim i As Long
    
    i = zlCommFun.ActualLen(strText)
    If i < lngLen Then
        i = lngLen - i
    Else
        i = i - lngLen
    End If
    GetLenText = strText & Space(i)
End Function

Private Sub MakeBedIcon()
    Dim i As Integer, k As Integer
    
    k = img32.ListImages.Count
    For i = 13 To 22
        img32.ListImages.Add , "加床_" & img32.ListImages(i).Key, img32.Overlay("MASK_加床", i)
        img32.ListImages.Add , "非编_" & img32.ListImages(i).Key, img32.Overlay("MASK_非编", i)
        img32.ListImages.Add , "共用_" & img32.ListImages(i).Key, img32.Overlay("MASK_共用", i)
        img32.ListImages.Add , "共用_加床_" & img32.ListImages(i).Key, img32.Overlay("MASK_共用_加床", i)
        img32.ListImages.Add , "共用_非编_" & img32.ListImages(i).Key, img32.Overlay("MASK_共用_非编", i)
    Next
    
    img32.ListImages.Add , "M_ChangeUnit", img32.Overlay("M", "U")
    img32.ListImages.Add , "KM_ChangeUnit", img32.Overlay("KM", "KU")
    img32.ListImages.Add , "F_ChangeUnit", img32.Overlay("F", "U")
    img32.ListImages.Add , "FM_ChangeUnit", img32.Overlay("KF", "KU")

    k = img16.ListImages.Count
    For i = 13 To 22
        img16.ListImages.Add , "加床_" & img16.ListImages(i).Key, img16.Overlay("MASK_加床", i)
        img16.ListImages.Add , "非编_" & img16.ListImages(i).Key, img16.Overlay("MASK_非编", i)
        img16.ListImages.Add , "共用_" & img16.ListImages(i).Key, img16.Overlay("MASK_共用", i)
        img16.ListImages.Add , "共用_加床_" & img16.ListImages(i).Key, img16.Overlay("MASK_共用_加床", i)
        img16.ListImages.Add , "共用_非编_" & img16.ListImages(i).Key, img16.Overlay("MASK_共用_非编", i)
    Next
    
    img16.ListImages.Add , "M_ChangeUnit", img16.Overlay("M", "U")
    img16.ListImages.Add , "KM_ChangeUnit", img16.Overlay("KM", "KU")
    img16.ListImages.Add , "F_ChangeUnit", img16.Overlay("F", "U")
    img16.ListImages.Add , "FM_ChangeUnit", img16.Overlay("KF", "KU")
End Sub

Private Sub tbrFilter_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim blnCheck As Boolean, i As Long
    
    If Button.Key = "curDay" Then
        Button.Image = IIf(Button.Image = "UnCheck_", "Check_", "UnCheck_")
        Call LoadList(True, True, False, False)
    ElseIf Button.Key Like "Nurse*" Then
        '不准全部清除
        blnCheck = False
        For i = 1 To tbrFilter.Buttons.Count
            If tbrFilter.Buttons(i).Key Like "Nurse*" _
                And tbrFilter.Buttons(i).Key <> Button.Key Then
                If tbrFilter.Buttons(i).Image = "Check" Then
                    blnCheck = True: Exit For
                End If
            End If
        Next
        If blnCheck Then
            Button.Image = IIf(Button.Image = "UnCheck", "Check", "UnCheck")
            Call LoadList(True, True, False, False)
        Else
            Button.Image = "Check"
            Exit Sub
        End If
    End If
End Sub

Private Sub tbrFilter_Change()
    Caption = Timer
End Sub

Private Sub timSize_Timer()
    Call Form_Resize
    timSize.Enabled = False
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub
'问题28811 by lesfeng 2010-03-30
Private Function InputGetDept(ByRef cboToDept As ComboBox, ByRef blnCancel As Boolean) As ADODB.Recordset
    '选择病区科室
'    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim vRect As RECT
    Dim lngHeigth As Long
    Dim strInput As String
    Dim strInputN As String
    Dim strno As String
    Dim blnLimitUnit As Boolean
    Dim strUnitIDs As String
    
    On Error GoTo errH
    
    cboToDept.Text = Replace(UCase(cboToDept.Text), "'", "")
    strInput = UCase(cboToDept.Text)
    strInputN = gstrLike & strInput & "%"
    strno = strInput & "%"
    
    If zlCommFun.IsCharChinese(strInput) Or InStr(1, strInput, "-", 0) <> 0 Then
        strSQL = strSQL & " And (A.名称 Like [3] or A.编码||'-'||A.名称 Like [3])" '输入汉字时只匹配名称
    Else
        strSQL = strSQL & " And (A.编码 Like [4] Or A.名称 Like [3] Or A.简码 Like [3])"
    End If
    
    blnLimitUnit = InStr(mstrPrivs, "所有病区") = 0
    If blnLimitUnit Then
        strUnitIDs = "," & GetUserUnits(False) & ","
    Else
        strUnitIDs = "," & GetUserUnits(True) & ","
    End If
    '目前包含门诊观察室
    strSQL = _
        " Select A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where B.部门ID = A.ID" & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.服务对象 IN(1,2,3) And B.工作性质='护理'" & _
        IIf(blnLimitUnit, " And instr([1],',' || A.ID || ',')>0 ", "") & _
        " And (A.站点=[2] Or A.站点 is Null) " & strSQL & _
        " Order by A.编码"
        '
    vRect = zlControl.GetControlRect(cboToDept.hWnd)
    lngHeigth = cboToDept.Height

    Set InputGetDept = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "科室选择", False, cboToDept.Tag, "", False, False, True, vRect.Left - 15, vRect.Top, lngHeigth, blnCancel, False, False, strUnitIDs, gstrNodeNo, strInputN, strno)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


