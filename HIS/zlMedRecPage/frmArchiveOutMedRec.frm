VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArchiveOutMedRec 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "门诊首页"
   ClientHeight    =   7260
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   8055
   Icon            =   "frmArchiveOutMedRec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   120
      ScaleHeight     =   7095
      ScaleWidth      =   7875
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   7875
      Begin MSComctlLib.ImageList imgSize 
         Left            =   480
         Top             =   7200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   9
         ImageHeight     =   9
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmArchiveOutMedRec.frx":000C
               Key             =   "-"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmArchiveOutMedRec.frx":04F6
               Key             =   "+"
            EndProperty
         EndProperty
      End
      Begin VB.VScrollBar vsc 
         Height          =   5985
         Left            =   7560
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.HScrollBar hsc 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   6720
         Visible         =   0   'False
         Width           =   7290
      End
      Begin VB.Frame fraVH 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7560
         TabIndex        =   2
         Top             =   6720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame fraBack 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   6555
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   7290
         Begin VB.Frame fraMain 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   基本信息"
            ForeColor       =   &H00FF0000&
            Height          =   5715
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Tag             =   "5715"
            Top             =   120
            Width           =   7095
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   172
               Left            =   1635
               Locked          =   -1  'True
               TabIndex        =   99
               Top             =   5280
               Width           =   2265
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   122
               Left            =   4110
               Locked          =   -1  'True
               TabIndex        =   86
               Top             =   1755
               Width           =   2265
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   200
               Left            =   7935
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   58
               Top             =   2040
               Width           =   1455
            End
            Begin VB.PictureBox picSize 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   0
               Left            =   180
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   24
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   56
               Top             =   4260
               Width           =   1395
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   23
               Left            =   1095
               Locked          =   -1  'True
               TabIndex        =   55
               Top             =   3915
               Width           =   5400
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   0
               Left            =   5340
               Locked          =   -1  'True
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   600
               Width           =   1140
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   18
               Left            =   1050
               Locked          =   -1  'True
               TabIndex        =   53
               Top             =   1375
               Width           =   2175
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   19
               Left            =   5505
               Locked          =   -1  'True
               TabIndex        =   52
               Top             =   995
               Width           =   975
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   16
               Left            =   3120
               Locked          =   -1  'True
               TabIndex        =   51
               Top             =   995
               Width           =   1215
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   8
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   50
               Top             =   995
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   6
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   49
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   5
               Left            =   3210
               Locked          =   -1  'True
               TabIndex        =   48
               Top             =   240
               Width           =   1230
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   155
               Left            =   3480
               Locked          =   -1  'True
               TabIndex        =   47
               Top             =   4260
               Width           =   3030
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   14
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   2100
               Width           =   2145
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   22
               Left            =   5085
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   3540
               Width           =   1395
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   21
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   44
               Top             =   3540
               Width           =   3030
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   20
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   43
               Top             =   3180
               Width           =   5400
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   27
               Left            =   5085
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   2820
               Width           =   1395
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   26
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   41
               Top             =   2820
               Width           =   3030
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   25
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   40
               Top             =   2460
               Width           =   5280
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   17
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   39
               Top             =   1755
               Width           =   2115
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   7
               Left            =   3210
               Locked          =   -1  'True
               TabIndex        =   38
               Top             =   600
               Width           =   1230
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   4
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   37
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   154
               Left            =   5340
               Locked          =   -1  'True
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   240
               Width           =   1140
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   32
               Left            =   4110
               Locked          =   -1  'True
               TabIndex        =   35
               Top             =   1375
               Width           =   2385
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   15
               Left            =   4110
               Locked          =   -1  'True
               TabIndex        =   34
               Top             =   2100
               Width           =   2385
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   156
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   33
               Top             =   4620
               Width           =   1515
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   77
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   32
               Top             =   4980
               Width           =   1515
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   80
               Left            =   3405
               Locked          =   -1  'True
               TabIndex        =   31
               Top             =   4980
               Width           =   3120
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   82
               Left            =   3405
               Locked          =   -1  'True
               TabIndex        =   30
               Top             =   4620
               Width           =   3120
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   1
               X1              =   1560
               X2              =   3930
               Y1              =   5550
               Y2              =   5550
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "监护人身份证号"
               Height          =   180
               Index           =   200
               Left            =   255
               TabIndex        =   100
               Top             =   5340
               Width           =   1260
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "其他证件"
               Height          =   180
               Index           =   122
               Left            =   3315
               TabIndex        =   87
               Top             =   1755
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   122
               X1              =   4155
               X2              =   6525
               Y1              =   1950
               Y2              =   1950
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "户口邮编"
               Height          =   180
               Index           =   24
               Left            =   255
               TabIndex        =   85
               Top             =   4275
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   24
               X1              =   1020
               X2              =   2490
               Y1              =   4455
               Y2              =   4455
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "户口地址"
               Height          =   180
               Index           =   23
               Left            =   255
               TabIndex        =   84
               Top             =   3915
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   23
               X1              =   1020
               X2              =   6495
               Y1              =   4110
               Y2              =   4110
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   22
               X1              =   5010
               X2              =   6480
               Y1              =   3735
               Y2              =   3735
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   27
               X1              =   5010
               X2              =   6480
               Y1              =   3015
               Y2              =   3015
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   155
               X1              =   3405
               X2              =   6510
               Y1              =   4455
               Y2              =   4455
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   21
               X1              =   1005
               X2              =   4110
               Y1              =   3735
               Y2              =   3735
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   26
               X1              =   1005
               X2              =   4110
               Y1              =   3015
               Y2              =   3015
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   20
               X1              =   1005
               X2              =   6480
               Y1              =   3375
               Y2              =   3375
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   25
               X1              =   1005
               X2              =   6480
               Y1              =   2655
               Y2              =   2655
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   14
               X1              =   1005
               X2              =   3240
               Y1              =   2295
               Y2              =   2295
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   17
               X1              =   1005
               X2              =   3240
               Y1              =   1935
               Y2              =   1935
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   18
               X1              =   990
               X2              =   3240
               Y1              =   1575
               Y2              =   1575
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   16
               X1              =   3120
               X2              =   4320
               Y1              =   1185
               Y2              =   1185
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   19
               X1              =   5280
               X2              =   6480
               Y1              =   1185
               Y2              =   1185
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   8
               X1              =   1005
               X2              =   2640
               Y1              =   1185
               Y2              =   1185
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   0
               X1              =   5280
               X2              =   6480
               Y1              =   795
               Y2              =   795
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   154
               X1              =   5280
               X2              =   6480
               Y1              =   435
               Y2              =   435
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   7
               X1              =   3150
               X2              =   4435
               Y1              =   795
               Y2              =   795
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   5
               X1              =   3150
               X2              =   4435
               Y1              =   435
               Y2              =   435
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   6
               X1              =   1005
               X2              =   2670
               Y1              =   795
               Y2              =   795
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   4
               X1              =   1005
               X2              =   2670
               Y1              =   435
               Y2              =   435
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "监护人"
               Height          =   180
               Index           =   155
               Left            =   2820
               TabIndex        =   83
               Top             =   4275
               Width           =   540
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出生日期"
               Height          =   180
               Index           =   6
               Left            =   240
               TabIndex        =   82
               Top             =   615
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出生地"
               Height          =   180
               Index           =   14
               Left            =   420
               TabIndex        =   81
               Top             =   2115
               Width           =   540
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "付费方式"
               Height          =   180
               Index           =   0
               Left            =   4500
               TabIndex        =   80
               Top             =   615
               Width           =   720
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "家庭邮编"
               Height          =   180
               Index           =   22
               Left            =   4215
               TabIndex        =   79
               Top             =   3555
               Width           =   720
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "家庭电话"
               Height          =   180
               Index           =   21
               Left            =   240
               TabIndex        =   78
               Top             =   3555
               Width           =   720
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "家庭地址"
               Height          =   180
               Index           =   20
               Left            =   240
               TabIndex        =   77
               Top             =   3195
               Width           =   720
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "单位邮编"
               Height          =   180
               Index           =   27
               Left            =   4215
               TabIndex        =   76
               Top             =   2835
               Width           =   720
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "单位电话"
               Height          =   180
               Index           =   26
               Left            =   240
               TabIndex        =   75
               Top             =   2835
               Width           =   720
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "单位名称"
               Height          =   180
               Index           =   25
               Left            =   240
               TabIndex        =   74
               Top             =   2475
               Width           =   720
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "身份证号"
               Height          =   180
               Index           =   17
               Left            =   240
               TabIndex        =   73
               Top             =   1755
               Width           =   720
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "国籍"
               Height          =   180
               Index           =   8
               Left            =   600
               TabIndex        =   72
               Top             =   995
               Width           =   360
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "民族"
               Height          =   180
               Index           =   16
               Left            =   2730
               TabIndex        =   71
               Top             =   995
               Width           =   360
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "职业"
               Height          =   180
               Index           =   18
               Left            =   600
               TabIndex        =   70
               Top             =   1375
               Width           =   360
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "婚姻状况"
               Height          =   180
               Index           =   19
               Left            =   4500
               TabIndex        =   69
               Top             =   995
               Width           =   720
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "年龄"
               Height          =   180
               Index           =   7
               Left            =   2730
               TabIndex        =   68
               Top             =   615
               Width           =   360
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "性别"
               Height          =   180
               Index           =   5
               Left            =   2730
               TabIndex        =   67
               Top             =   255
               Width           =   360
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "姓名"
               Height          =   180
               Index           =   4
               Left            =   600
               TabIndex        =   66
               Top             =   255
               Width           =   360
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门诊号"
               Height          =   180
               Index           =   154
               Left            =   4680
               TabIndex        =   65
               Top             =   255
               Width           =   540
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   32
               X1              =   4110
               X2              =   6480
               Y1              =   1570
               Y2              =   1570
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "区域"
               Height          =   180
               Index           =   32
               Left            =   3675
               TabIndex        =   64
               Top             =   1375
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   15
               X1              =   4110
               X2              =   6480
               Y1              =   2295
               Y2              =   2295
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "籍贯"
               Height          =   180
               Index           =   15
               Left            =   3675
               TabIndex        =   63
               Top             =   2115
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   156
               X1              =   1020
               X2              =   2520
               Y1              =   4815
               Y2              =   4815
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "文化程度"
               Height          =   180
               Index           =   156
               Left            =   255
               TabIndex        =   62
               Top             =   4635
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   77
               X1              =   1020
               X2              =   2520
               Y1              =   5175
               Y2              =   5175
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "血型"
               Height          =   180
               Index           =   77
               Left            =   615
               TabIndex        =   61
               Top             =   4995
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   80
               X1              =   3405
               X2              =   6480
               Y1              =   5175
               Y2              =   5175
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rh"
               Height          =   180
               Index           =   80
               Left            =   3180
               TabIndex        =   60
               Top             =   4995
               Width           =   180
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   82
               X1              =   3405
               X2              =   6480
               Y1              =   4815
               Y2              =   4815
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "生育状况"
               Height          =   180
               Index           =   82
               Left            =   2640
               TabIndex        =   59
               Top             =   4635
               Width           =   720
            End
         End
         Begin VB.Frame fraMain 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   就诊信息"
            ForeColor       =   &H00FF0000&
            Height          =   6405
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Tag             =   "6405"
            Top             =   120
            Width           =   7095
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "无过敏记录"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   31
               Left            =   1200
               TabIndex        =   27
               Top             =   1028
               Width           =   1290
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   162
               Left            =   3525
               Locked          =   -1  'True
               TabIndex        =   95
               Top             =   5520
               Width           =   960
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   163
               Left            =   5205
               Locked          =   -1  'True
               TabIndex        =   94
               Top             =   5520
               Width           =   960
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   10
               Left            =   3510
               Locked          =   -1  'True
               TabIndex        =   90
               Top             =   5205
               Width           =   960
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   9
               Left            =   1350
               Locked          =   -1  'True
               TabIndex        =   89
               Top             =   5205
               Width           =   1515
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   160
               Left            =   5190
               Locked          =   -1  'True
               TabIndex        =   88
               Top             =   5205
               Width           =   960
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   201
               Left            =   7935
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   16
               Top             =   2040
               Width           =   1455
            End
            Begin VB.PictureBox picSize 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   1
               Left            =   180
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   405
               Index           =   157
               Left            =   600
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   14
               Top             =   495
               Width           =   6285
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "复诊(&R)"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   28
               Left            =   240
               TabIndex        =   13
               Top             =   4560
               Width           =   930
            End
            Begin VB.CheckBox chkInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "传染病上传(&U)"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   29
               Left            =   1440
               TabIndex        =   12
               Top             =   4560
               Width           =   1470
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   79
               Left            =   1425
               Locked          =   -1  'True
               TabIndex        =   11
               Top             =   4905
               Width           =   1470
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   159
               Left            =   3960
               Locked          =   -1  'True
               TabIndex        =   10
               Top             =   4905
               Width           =   1830
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   161
               Left            =   1350
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   5520
               Width           =   1125
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   158
               Left            =   3870
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   4515
               Width           =   1965
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   92
               Left            =   1350
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   6135
               Width           =   4965
            End
            Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
               Height          =   1035
               Left            =   255
               TabIndex        =   17
               Top             =   2535
               Width           =   6705
               _cx             =   11827
               _cy             =   1826
               Appearance      =   0
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
               BackColorFixed  =   -2147483643
               ForeColorFixed  =   -2147483630
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   225
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveOutMedRec.frx":09E0
               ScrollTrack     =   -1  'True
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   115
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
            Begin VSFlex8Ctl.VSFlexGrid vsAller 
               Height          =   915
               Left            =   255
               TabIndex        =   18
               Top             =   1260
               Width           =   6705
               _cx             =   11827
               _cy             =   1614
               Appearance      =   0
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
               BackColorFixed  =   -2147483643
               ForeColorFixed  =   -2147483630
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   8421504
               GridColorFixed  =   8421504
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   225
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveOutMedRec.frx":0A7F
               ScrollTrack     =   -1  'True
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
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   91
               Left            =   1350
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   5820
               Width           =   4845
            End
            Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
               Height          =   675
               Left            =   240
               TabIndex        =   98
               Top             =   3720
               Width           =   6705
               _cx             =   11827
               _cy             =   1191
               Appearance      =   0
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
               BackColorFixed  =   -2147483643
               ForeColorFixed  =   -2147483630
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   5
               FixedRows       =   0
               FixedCols       =   0
               RowHeightMin    =   225
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveOutMedRec.frx":0AF5
               ScrollTrack     =   -1  'True
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   115
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
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "呼吸"
               Height          =   180
               Index           =   162
               Left            =   3120
               TabIndex        =   97
               Top             =   5520
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   162
               X1              =   3525
               X2              =   4560
               Y1              =   5715
               Y2              =   5715
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   163
               X1              =   5160
               X2              =   6240
               Y1              =   5715
               Y2              =   5715
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "血压"
               Height          =   180
               Index           =   163
               Left            =   4695
               TabIndex        =   96
               Top             =   5520
               Width           =   360
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "体重"
               Height          =   180
               Index           =   10
               Left            =   3105
               TabIndex        =   93
               Top             =   5205
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   10
               X1              =   3510
               X2              =   4545
               Y1              =   5400
               Y2              =   5400
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "身高"
               Height          =   180
               Index           =   9
               Left            =   960
               TabIndex        =   92
               Top             =   5205
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   9
               X1              =   1350
               X2              =   2850
               Y1              =   5400
               Y2              =   5400
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   160
               X1              =   5145
               X2              =   6225
               Y1              =   5400
               Y2              =   5400
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "体温"
               Height          =   180
               Index           =   160
               Left            =   4680
               TabIndex        =   91
               Top             =   5205
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   157
               X1              =   375
               X2              =   6960
               Y1              =   900
               Y2              =   900
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "就诊摘要："
               Height          =   180
               Index           =   157
               Left            =   255
               TabIndex        =   28
               Top             =   240
               Width           =   900
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "过敏记录："
               Height          =   180
               Index           =   42
               Left            =   240
               TabIndex        =   26
               Top             =   1035
               Width           =   900
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "诊断记录："
               Height          =   180
               Index           =   43
               Left            =   240
               TabIndex        =   25
               Top             =   2325
               Width           =   900
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   79
               X1              =   1350
               X2              =   3000
               Y1              =   5100
               Y2              =   5100
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "发病时间"
               Height          =   180
               Index           =   79
               Left            =   600
               TabIndex        =   24
               Top             =   4905
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   159
               X1              =   3960
               X2              =   5970
               Y1              =   5100
               Y2              =   5100
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "发病地址"
               Height          =   180
               Index           =   159
               Left            =   3120
               TabIndex        =   23
               Top             =   4905
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   161
               X1              =   1350
               X2              =   2880
               Y1              =   5715
               Y2              =   5715
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "脉搏"
               Height          =   180
               Index           =   161
               Left            =   960
               TabIndex        =   22
               Top             =   5520
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   158
               X1              =   3960
               X2              =   5970
               Y1              =   4710
               Y2              =   4710
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "去向"
               Height          =   180
               Index           =   158
               Left            =   3480
               TabIndex        =   21
               Top             =   4567
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   91
               X1              =   1350
               X2              =   6240
               Y1              =   6000
               Y2              =   6000
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医学警示"
               Height          =   180
               Index           =   91
               Left            =   600
               TabIndex        =   20
               Top             =   5820
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   92
               X1              =   1350
               X2              =   6360
               Y1              =   6330
               Y2              =   6330
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "其他医学警示"
               Height          =   180
               Index           =   92
               Left            =   240
               TabIndex        =   19
               Top             =   6135
               Width           =   1080
            End
         End
      End
   End
End
Attribute VB_Name = "frmArchiveOutMedRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'说明：为了保持界面的可维护性，在新增控件时，注意保持每个信息条目包含的lblInfo，linInfo,txtinfo 的index相同，
'      若这组信息条目包含2个lblinfo则另外一个lblinfo的index为txtinfo.index+100
Private Sub chkInfo_Click(Index As Integer)
    Call ArchivechkInfoClick(Index)
End Sub

Private Sub Form_Activate()
    Call Form_Resize
    gOldwinproc = GetWindowLong(picBack.hwnd, GWL_WNDPROC)
    SetWindowLong picBack.hwnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub

Private Sub Form_Deactivate()
    SetWindowLong picBack.hwnd, GWL_WNDPROC, gOldwinproc
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call ArchiveFormKeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
    Call ArchiveFormLoad
End Sub

Private Sub Form_Resize()
    Call ArchiveFormResize
End Sub

Private Sub hsc_Change()
    Call hsc_Scroll
End Sub

Private Sub picSize_Click(Index As Integer)
    Call ArchivepicSizeClick(Index)
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtInfo(Index))
End Sub

Private Sub vsAller_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= 0 And NewCol >= 0 Then Call vsAller.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsc_Change()
    Call vsc_Scroll
End Sub

Private Sub hsc_Scroll()
    fraBack.Left = hsc.Value * Screen.TwipsPerPixelX
End Sub

Private Sub vsc_Scroll()
    fraBack.Top = vsc.Value * Screen.TwipsPerPixelY
End Sub

Private Sub vsDiagXY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= 0 And NewCol >= 0 Then Call vsDiagXY.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= 0 And NewCol >= 0 Then Call vsDiagZY.ShowCell(NewRow, NewCol)
End Sub



