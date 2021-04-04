VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArchiveOutMedRec 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "门诊首页"
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   8055
   Icon            =   "frmArchiveOutMedRec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   120
      ScaleHeight     =   6615
      ScaleWidth      =   7875
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   7875
      Begin MSComctlLib.ImageList imgSize 
         Left            =   600
         Top             =   6000
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
         Top             =   6240
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
         Top             =   6240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame fraBack 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   5985
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   7290
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   基本信息"
            ForeColor       =   &H00FF0000&
            Height          =   5850
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Tag             =   "5850"
            Top             =   120
            Width           =   7095
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   38
               Left            =   4110
               Locked          =   -1  'True
               TabIndex        =   89
               Top             =   1755
               Width           =   2265
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   130
               Left            =   7935
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   59
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
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   21
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   57
               Top             =   4260
               Width           =   1395
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   20
               Left            =   1095
               Locked          =   -1  'True
               TabIndex        =   56
               Top             =   3915
               Width           =   5400
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   5
               Left            =   5340
               Locked          =   -1  'True
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   600
               Width           =   1140
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   9
               Left            =   1050
               Locked          =   -1  'True
               TabIndex        =   54
               Top             =   1375
               Width           =   2175
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   8
               Left            =   5505
               Locked          =   -1  'True
               TabIndex        =   53
               Top             =   995
               Width           =   975
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   7
               Left            =   3120
               Locked          =   -1  'True
               TabIndex        =   52
               Top             =   995
               Width           =   1215
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   6
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   51
               Top             =   995
               Width           =   1335
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   3
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   50
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   1
               Left            =   3210
               Locked          =   -1  'True
               TabIndex        =   49
               Top             =   240
               Width           =   1230
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   22
               Left            =   3480
               Locked          =   -1  'True
               TabIndex        =   48
               Top             =   4260
               Width           =   3030
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   12
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   47
               Top             =   2100
               Width           =   2145
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   19
               Left            =   5085
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   3540
               Width           =   1395
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   18
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   3540
               Width           =   3030
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   17
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   44
               Top             =   3180
               Width           =   5400
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   16
               Left            =   5085
               Locked          =   -1  'True
               TabIndex        =   43
               Top             =   2820
               Width           =   1395
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   15
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   2820
               Width           =   3030
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   14
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   41
               Top             =   2460
               Width           =   5280
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   10
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   40
               Top             =   1755
               Width           =   2115
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   4
               Left            =   3210
               Locked          =   -1  'True
               TabIndex        =   39
               Top             =   600
               Width           =   1230
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   0
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   38
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   2
               Left            =   5340
               Locked          =   -1  'True
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   240
               Width           =   1140
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   11
               Left            =   4110
               Locked          =   -1  'True
               TabIndex        =   36
               Top             =   1375
               Width           =   2385
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   13
               Left            =   4110
               Locked          =   -1  'True
               TabIndex        =   35
               Top             =   2100
               Width           =   2385
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   23
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   34
               Top             =   4620
               Width           =   1515
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   25
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   33
               Top             =   4980
               Width           =   1515
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   26
               Left            =   3405
               Locked          =   -1  'True
               TabIndex        =   32
               Top             =   4980
               Width           =   3120
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   27
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   31
               Top             =   5340
               Width           =   1515
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   24
               Left            =   3405
               Locked          =   -1  'True
               TabIndex        =   30
               Top             =   4620
               Width           =   3120
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   28
               Left            =   3405
               Locked          =   -1  'True
               TabIndex        =   29
               Top             =   5340
               Width           =   3120
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "其他证件"
               Height          =   180
               Index           =   38
               Left            =   3315
               TabIndex        =   90
               Top             =   1755
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   38
               X1              =   4155
               X2              =   6525
               Y1              =   1950
               Y2              =   1950
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "户口邮编"
               Height          =   180
               Index           =   21
               Left            =   255
               TabIndex        =   88
               Top             =   4275
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   21
               X1              =   1020
               X2              =   2490
               Y1              =   4455
               Y2              =   4455
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "户口地址"
               Height          =   180
               Index           =   20
               Left            =   255
               TabIndex        =   87
               Top             =   3915
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   20
               X1              =   1020
               X2              =   6495
               Y1              =   4110
               Y2              =   4110
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   19
               X1              =   5010
               X2              =   6480
               Y1              =   3735
               Y2              =   3735
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   16
               X1              =   5010
               X2              =   6480
               Y1              =   3015
               Y2              =   3015
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   22
               X1              =   3405
               X2              =   6510
               Y1              =   4455
               Y2              =   4455
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   18
               X1              =   1005
               X2              =   4110
               Y1              =   3735
               Y2              =   3735
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   15
               X1              =   1005
               X2              =   4110
               Y1              =   3015
               Y2              =   3015
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   17
               X1              =   1005
               X2              =   6480
               Y1              =   3375
               Y2              =   3375
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   14
               X1              =   1005
               X2              =   6480
               Y1              =   2655
               Y2              =   2655
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   12
               X1              =   1005
               X2              =   3240
               Y1              =   2295
               Y2              =   2295
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   10
               X1              =   1005
               X2              =   3240
               Y1              =   1935
               Y2              =   1935
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   9
               X1              =   990
               X2              =   3240
               Y1              =   1575
               Y2              =   1575
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   7
               X1              =   3120
               X2              =   4320
               Y1              =   1185
               Y2              =   1185
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   8
               X1              =   5280
               X2              =   6480
               Y1              =   1185
               Y2              =   1185
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   6
               X1              =   1005
               X2              =   2640
               Y1              =   1185
               Y2              =   1185
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   5
               X1              =   5280
               X2              =   6480
               Y1              =   795
               Y2              =   795
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   2
               X1              =   5280
               X2              =   6480
               Y1              =   435
               Y2              =   435
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   4
               X1              =   3150
               X2              =   4435
               Y1              =   795
               Y2              =   795
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   1
               X1              =   3150
               X2              =   4435
               Y1              =   435
               Y2              =   435
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   3
               X1              =   1005
               X2              =   2670
               Y1              =   795
               Y2              =   795
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   0
               X1              =   1005
               X2              =   2670
               Y1              =   435
               Y2              =   435
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "监护人"
               Height          =   180
               Index           =   22
               Left            =   2820
               TabIndex        =   86
               Top             =   4275
               Width           =   540
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出生日期"
               Height          =   180
               Index           =   3
               Left            =   240
               TabIndex        =   85
               Top             =   615
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出生地点"
               Height          =   180
               Index           =   12
               Left            =   240
               TabIndex        =   84
               Top             =   2115
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "付款方式"
               Height          =   180
               Index           =   5
               Left            =   4500
               TabIndex        =   83
               Top             =   615
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "家庭邮编"
               Height          =   180
               Index           =   19
               Left            =   4215
               TabIndex        =   82
               Top             =   3555
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "家庭电话"
               Height          =   180
               Index           =   18
               Left            =   240
               TabIndex        =   81
               Top             =   3555
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "家庭地址"
               Height          =   180
               Index           =   17
               Left            =   240
               TabIndex        =   80
               Top             =   3195
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "单位邮编"
               Height          =   180
               Index           =   16
               Left            =   4215
               TabIndex        =   79
               Top             =   2835
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "单位电话"
               Height          =   180
               Index           =   15
               Left            =   240
               TabIndex        =   78
               Top             =   2835
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "单位名称"
               Height          =   180
               Index           =   14
               Left            =   240
               TabIndex        =   77
               Top             =   2475
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "身份证号"
               Height          =   180
               Index           =   10
               Left            =   240
               TabIndex        =   76
               Top             =   1755
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "国籍"
               Height          =   180
               Index           =   6
               Left            =   600
               TabIndex        =   75
               Top             =   995
               Width           =   360
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "民族"
               Height          =   180
               Index           =   7
               Left            =   2730
               TabIndex        =   74
               Top             =   995
               Width           =   360
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "职业"
               Height          =   180
               Index           =   9
               Left            =   600
               TabIndex        =   73
               Top             =   1375
               Width           =   360
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "婚姻状况"
               Height          =   180
               Index           =   8
               Left            =   4500
               TabIndex        =   72
               Top             =   995
               Width           =   720
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "年龄"
               Height          =   180
               Index           =   4
               Left            =   2730
               TabIndex        =   71
               Top             =   615
               Width           =   360
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "性别"
               Height          =   180
               Index           =   1
               Left            =   2730
               TabIndex        =   70
               Top             =   255
               Width           =   360
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "姓名"
               Height          =   180
               Index           =   0
               Left            =   600
               TabIndex        =   69
               Top             =   255
               Width           =   360
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门诊号"
               Height          =   180
               Index           =   2
               Left            =   4680
               TabIndex        =   68
               Top             =   255
               Width           =   540
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   11
               X1              =   4110
               X2              =   6480
               Y1              =   1570
               Y2              =   1570
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "区域"
               Height          =   180
               Index           =   11
               Left            =   3675
               TabIndex        =   67
               Top             =   1375
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   13
               X1              =   4110
               X2              =   6480
               Y1              =   2295
               Y2              =   2295
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "籍贯"
               Height          =   180
               Index           =   13
               Left            =   3675
               TabIndex        =   66
               Top             =   2115
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   23
               X1              =   1020
               X2              =   2520
               Y1              =   4815
               Y2              =   4815
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "文化程度"
               Height          =   180
               Index           =   23
               Left            =   255
               TabIndex        =   65
               Top             =   4635
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   25
               X1              =   1020
               X2              =   2520
               Y1              =   5175
               Y2              =   5175
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "血型"
               Height          =   180
               Index           =   25
               Left            =   615
               TabIndex        =   64
               Top             =   4995
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   26
               X1              =   3405
               X2              =   6480
               Y1              =   5175
               Y2              =   5175
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rh"
               Height          =   180
               Index           =   26
               Left            =   3180
               TabIndex        =   63
               Top             =   4995
               Width           =   180
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   27
               X1              =   1020
               X2              =   2520
               Y1              =   5535
               Y2              =   5535
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "身高"
               Height          =   180
               Index           =   27
               Left            =   615
               TabIndex        =   62
               Top             =   5355
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   24
               X1              =   3405
               X2              =   6480
               Y1              =   4815
               Y2              =   4815
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "生育状况"
               Height          =   180
               Index           =   24
               Left            =   2640
               TabIndex        =   61
               Top             =   4635
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   28
               X1              =   3405
               X2              =   6480
               Y1              =   5535
               Y2              =   5535
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "体重"
               Height          =   180
               Index           =   28
               Left            =   3000
               TabIndex        =   60
               Top             =   5355
               Width           =   360
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   就诊信息"
            ForeColor       =   &H00FF0000&
            Height          =   5805
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Tag             =   "6400"
            Top             =   120
            Width           =   7095
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   0
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
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   405
               Index           =   29
               Left            =   600
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   14
               Top             =   495
               Width           =   6285
            End
            Begin VB.CheckBox chkEdit 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "复诊(&R)"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   13
               Top             =   3960
               Width           =   930
            End
            Begin VB.CheckBox chkEdit 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "传染病上传(&U)"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   1440
               TabIndex        =   12
               Top             =   3960
               Width           =   1470
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   30
               Left            =   1425
               Locked          =   -1  'True
               TabIndex        =   11
               Top             =   4305
               Width           =   1725
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   31
               Left            =   4440
               Locked          =   -1  'True
               TabIndex        =   10
               Top             =   4305
               Width           =   1830
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   32
               Left            =   1350
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   4635
               Width           =   1725
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   33
               Left            =   4350
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   4635
               Width           =   1965
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   34
               Left            =   1350
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   4965
               Width           =   4965
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   35
               Left            =   1350
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   5295
               Width           =   4965
            End
            Begin VSFlex8Ctl.VSFlexGrid vsDiag 
               Height          =   1275
               Left            =   255
               TabIndex        =   17
               Top             =   2535
               Width           =   6705
               _cx             =   11827
               _cy             =   2249
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
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   29
               X1              =   375
               X2              =   6960
               Y1              =   900
               Y2              =   900
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "就诊摘要："
               Height          =   180
               Index           =   29
               Left            =   255
               TabIndex        =   27
               Top             =   240
               Width           =   900
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "过敏记录："
               Height          =   180
               Index           =   37
               Left            =   240
               TabIndex        =   26
               Top             =   1035
               Width           =   900
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "诊断记录："
               Height          =   180
               Index           =   36
               Left            =   240
               TabIndex        =   25
               Top             =   2325
               Width           =   900
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   30
               X1              =   1350
               X2              =   3155
               Y1              =   4500
               Y2              =   4500
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "发病时间"
               Height          =   180
               Index           =   30
               Left            =   600
               TabIndex        =   24
               Top             =   4305
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   31
               X1              =   4350
               X2              =   6360
               Y1              =   4500
               Y2              =   4500
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "发病地址"
               Height          =   180
               Index           =   31
               Left            =   3600
               TabIndex        =   23
               Top             =   4305
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   32
               X1              =   1350
               X2              =   3155
               Y1              =   4830
               Y2              =   4830
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "血压"
               Height          =   180
               Index           =   32
               Left            =   960
               TabIndex        =   22
               Top             =   4635
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   33
               X1              =   4350
               X2              =   6360
               Y1              =   4830
               Y2              =   4830
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "去向"
               Height          =   180
               Index           =   33
               Left            =   3960
               TabIndex        =   21
               Top             =   4635
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   34
               X1              =   1350
               X2              =   6360
               Y1              =   5160
               Y2              =   5160
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医学警示"
               Height          =   180
               Index           =   34
               Left            =   600
               TabIndex        =   20
               Top             =   4965
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   35
               X1              =   1350
               X2              =   6360
               Y1              =   5490
               Y2              =   5490
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "其他医学警示"
               Height          =   180
               Index           =   35
               Left            =   240
               TabIndex        =   19
               Top             =   5295
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
Private mlng病人ID As Long
Private mlng挂号ID As Long
Private mblnMoved As Boolean
Private mblnCheck As Boolean

Private Enum 基本信息
    txt姓名 = 0
    txt性别 = 1
    txt门诊号 = 2
    txt出生日期 = 3
    txt年龄 = 4
    txt付款方式 = 5
    txt国籍 = 6
    txt民族 = 7
    txt婚姻 = 8
    txt职业 = 9
    txt身份证号 = 10
    txt其他证件 = 38
    txt区域 = 11
    txt出生地点 = 12
    txt籍贯 = 13
    txt工作单位 = 14
    txt单位电话 = 15
    txt单位邮编 = 16
    txt家庭地址 = 17
    txt家庭电话 = 18
    txt家庭邮编 = 19
    txt户口地址 = 20
    txt户口地址邮编 = 21
    txt监护人 = 22
    txt文化程度 = 23
    txt生育状况 = 24
    txt血型 = 25
    txtRh = 26
    txt身高 = 27
    txt体重 = 28
End Enum

Private Enum 就诊信息
    txt就诊摘要 = 29
    chk复诊 = 0
    chk传染病上传 = 1
    txt发病时间 = 30
    txt发病地址 = 31
    txt血压 = 32
    txt去向 = 33
    txt医学警示 = 34
    txt其他医学警示 = 35
End Enum

Private Enum COL_Aller
    col过敏时间 = 0
    col过敏药物 = 1
    col过敏反应 = 2
End Enum

Private Enum COL_Diag
    col类型 = 0
    col诊断编码 = 1
    col诊断描述 = 2
    col发病时间 = 3
    col疑诊 = 4
End Enum

Public Function zlRefresh(ByVal lng病人ID As Long, ByVal lng挂号id As Long, ByVal blnMoved As Boolean) As Boolean
'功能：刷新或清除医嘱清单
    mlng病人ID = lng病人ID: mlng挂号ID = lng挂号id: mblnMoved = blnMoved
    
    Call SetPageHeight
    Call SetScrollbar
    
'    Call ClearPageData
    If mlng病人ID <> 0 Then Call LoadMedRec
    
    Call Form_Resize
    zlRefresh = True

End Function

Private Function LoadMedRec() As Boolean
'功能：读取门诊首页的各种信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim lngRow As Long, bln中医 As Boolean
    
    mblnCheck = True
    
    On Error GoTo errH
    
    '基本信息
    strSql = "Select B.执行部门ID as 科室ID,B.摘要,B.复诊," & _
        " B.传染病上传,B.发病时间,B.发病地址,A.险类,A.门诊号,A.姓名,A.性别,A.年龄,A.出生日期,A.医疗付款方式," & _
        " A.国籍,A.民族,A.婚姻状况,A.职业,A.身份证号,A.出生地点,A.监护人,A.家庭地址,A.家庭电话," & _
        " A.家庭地址邮编,A.工作单位,A.单位电话,A.单位邮编,a.户口地址,a.户口地址邮编,A.区域,A.籍贯,A.其他证件 " & _
        " From 病人信息 A,病人挂号记录 B Where A.病人ID=B.病人ID And B.ID=[1] And B.记录性质=1 And B.记录状态=1"
    If mblnMoved Then
        strSql = Replace(strSql, "病人挂号记录", "H病人挂号记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng挂号ID)
    If rsTmp.EOF Then Exit Function
    
    bln中医 = Have部门性质(rsTmp!科室ID, "中医科")
        
    txtEdit(txt姓名).Text = Nvl(rsTmp!姓名)
    txtEdit(txt性别).Text = Nvl(rsTmp!性别)
    txtEdit(txt年龄).Text = Nvl(rsTmp!年龄)
    txtEdit(txt门诊号).Text = Nvl(rsTmp!门诊号)
    
    txtEdit(txt出生日期).Text = Format(rsTmp!出生日期, "yyyy-MM-dd")
    If Format(rsTmp!出生日期, "HH:mm") <> "00:00" Then
        txtEdit(txt出生日期).Text = Format(rsTmp!出生日期, "yyyy-MM-dd HH:mm")
    End If
    
    txtEdit(txt付款方式) = Nvl(rsTmp!医疗付款方式)
    txtEdit(txt国籍) = Nvl(rsTmp!国籍)
    txtEdit(txt民族) = Nvl(rsTmp!民族)
    txtEdit(txt婚姻) = Nvl(rsTmp!婚姻状况)
    txtEdit(txt职业) = Nvl(rsTmp!职业)
    txtEdit(txt身份证号).Text = Nvl(rsTmp!身份证号)
    txtEdit(txt其他证件).Text = Nvl(rsTmp!其他证件)
    txtEdit(txt区域).Text = Nvl(rsTmp!区域)
    txtEdit(txt出生地点).Text = Nvl(rsTmp!出生地点)
    txtEdit(txt籍贯).Text = Nvl(rsTmp!籍贯)
    txtEdit(txt工作单位).Text = Nvl(rsTmp!工作单位)
    txtEdit(txt单位电话).Text = Nvl(rsTmp!单位电话)
    txtEdit(txt单位邮编).Text = Nvl(rsTmp!单位邮编)
    txtEdit(txt家庭地址).Text = Nvl(rsTmp!家庭地址)
    txtEdit(txt家庭电话).Text = Nvl(rsTmp!家庭电话)
    txtEdit(txt家庭邮编).Text = Nvl(rsTmp!家庭地址邮编)
    txtEdit(txt户口地址).Text = Nvl(rsTmp!户口地址)
    txtEdit(txt户口地址邮编).Text = Nvl(rsTmp!户口地址邮编)
    txtEdit(txt监护人).Text = Nvl(rsTmp!监护人)
    txtEdit(txt就诊摘要).Text = Nvl(rsTmp!摘要)
    
    chkEdit(chk复诊).Value = Nvl(rsTmp!复诊, 0)
    chkEdit(chk传染病上传).Value = Nvl(rsTmp!传染病上传, 0)

    txtEdit(txt发病时间).Text = Format(rsTmp!发病时间, "yyyy-MM-dd")
    If Format(rsTmp!发病时间, "HH:mm") <> "00:00" Then
        txtEdit(txt发病时间).Text = Format(rsTmp!发病时间, "yyyy-MM-dd HH:mm")
    End If
    txtEdit(txt发病地址).Text = Nvl(rsTmp!发病地址)
    '附加信息
    strSql = "Select 信息名,信息值 From 病人信息从表 Where 病人ID=[1] And (就诊ID=[2] Or 就诊ID is Null) Order by Nvl(就诊ID,999999999)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng挂号ID)
    rsTmp.Filter = "信息名='身高'"
    If Not rsTmp.EOF Then txtEdit(txt身高).Text = Nvl(rsTmp!信息值) & IIf(Nvl(rsTmp!信息值) = "", "", " cm")
    rsTmp.Filter = "信息名='体重'"
    If Not rsTmp.EOF Then txtEdit(txt体重).Text = Nvl(rsTmp!信息值) & IIf(Nvl(rsTmp!信息值) = "", "", " Kg")
    rsTmp.Filter = "信息名='文化程度'"
    If Not rsTmp.EOF Then txtEdit(txt文化程度) = Decode(Val(Nvl(rsTmp!信息值)), 1, "研究生及以上", 2, "大学", 3, "大专", 4, "中专", 6, "高中", 7, "初中", 8, "小学（包括村学）", 9, "文盲和半文盲")
    rsTmp.Filter = "信息名='生育状况'"
    If Not rsTmp.EOF Then txtEdit(txt生育状况) = Decode(Val(Nvl(rsTmp!信息值)), 0, "未生育", 1, "生育1胎", 2, "生育2胎及以上", 4, "不详")
    rsTmp.Filter = "信息名='收缩压'"
    If Not rsTmp.EOF Then txtEdit(txt血压).Text = IIf(Nvl(rsTmp!信息值) = "", "   ", Nvl(rsTmp!信息值))
    rsTmp.Filter = "信息名='舒张压'"
    If Not rsTmp.EOF Then txtEdit(txt血压).Text = txtEdit(txt血压).Text & " / " & IIf(Nvl(rsTmp!信息值) = "", "   ", Nvl(rsTmp!信息值))
    
    If txtEdit(txt血压).Text <> "    /    " Then
        rsTmp.Filter = "信息名='血压单位'"
        If Not rsTmp.EOF Then
            txtEdit(txt血压).Text = txtEdit(txt血压).Text & " " & Nvl(rsTmp!信息值)
        Else
            txtEdit(txt血压).Text = txtEdit(txt血压).Text & " mmHg"
        End If
    End If
    rsTmp.Filter = "信息名='去向'"
    If Not rsTmp.EOF Then txtEdit(txt去向).Text = Nvl(rsTmp!信息值)
    rsTmp.Filter = "信息名='血型'"
    If Not rsTmp.EOF Then txtEdit(txt血型) = Nvl(rsTmp!信息值)
    rsTmp.Filter = "信息名='RH'"
    If Not rsTmp.EOF Then txtEdit(txtRh) = Nvl(rsTmp!信息值)
    rsTmp.Filter = "信息名='医学警示'"
    If Not rsTmp.EOF Then txtEdit(txt医学警示).Text = Nvl(rsTmp!信息值)
    rsTmp.Filter = "信息名='其他医学警示'"
    If Not rsTmp.EOF Then txtEdit(txt其他医学警示).Text = Nvl(rsTmp!信息值)
    
    '过敏信息:本次挂号的,过敏的
    strSql = "Select 记录来源,Nvl(过敏时间,记录时间) as 过敏时间,药物ID,药物名,过敏反应 From 病人过敏记录 A" & _
        " Where 结果=1 And 病人ID=[1] And 主页ID=[2]" & _
        " And Not Exists(Select 药物ID From 病人过敏记录" & _
            " Where (Nvl(药物ID,0)=Nvl(A.药物ID,0) Or Nvl(药物名,'Null')=Nvl(A.药物名,'Null'))" & _
            " And Nvl(结果,0)=0 And 记录时间>A.记录时间 And 病人ID=[1] And 主页ID=[2])" & _
        " Order by Nvl(过敏时间,记录时间),药物名"
    If mblnMoved Then
        strSql = Replace(strSql, "病人过敏记录", "H病人过敏记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng挂号ID)
    
    vsAller.Rows = vsAller.FixedRows
    If Not rsTmp.EOF Then
        rsTmp.Filter = "记录来源=3" '首页本身填写的
        If rsTmp.EOF Then rsTmp.Filter = "记录来源<>3" '其它来源的作为缺省显示
        With vsAller
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                '其它来源的可能有重复
                lngRow = -1
                If Not IsNull(rsTmp!药物ID) Then
                    lngRow = .FindRow(CLng(rsTmp!药物ID))
                ElseIf Not IsNull(rsTmp!药物名) Then
                    lngRow = .FindRow(CStr(rsTmp!药物名), , 1)
                End If
                If lngRow = -1 Then
                    .RowData(i) = CLng(Nvl(rsTmp!药物ID, 0))
                    .TextMatrix(i, col过敏时间) = Format(rsTmp!过敏时间, "yyyy-MM-dd HH:mm")
                    .TextMatrix(i, col过敏药物) = Nvl(rsTmp!药物名)
                    .TextMatrix(i, col过敏反应) = Nvl(rsTmp!过敏反应)
                End If
                rsTmp.MoveNext
            Next
            .Row = 1: .Col = 1
        End With
    End If
    strSql = "Select a.记录来源,a.诊断类型,a.疾病ID,a.诊断ID,a.证候ID,a.诊断描述,a.是否疑诊,b.编码 as 疾病编码,c.编码 as 诊断编码,d.编码 as 证候编码,A.发病时间,A.诊断次序 From 病人诊断记录  A, 疾病编码目录 B, 疾病诊断目录 C,疾病编码目录 D" & _
            " Where  a.疾病id = b.Id(+) And a.诊断id = c.Id(+) And a.证候ID=d.ID(+) And a.记录来源 IN(1,3) And a.诊断类型 IN(1,11)" & _
            " And a.取消时间 is Null And a.病人ID=[1] And a.主页ID=[2]" & _
            " Order by a.诊断类型,a.诊断次序,a.编码序号"
    If mblnMoved Then
        strSql = Replace(strSql, "病人诊断记录", "H病人诊断记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng挂号ID)
    
     vsDiag.Rows = vsDiag.FixedRows
    If Not rsTmp.EOF Then
        With vsDiag
            '西医诊断
            rsTmp.Filter = "诊断类型=1 And 记录来源=3" '首页本身填写的
            If rsTmp.EOF Then rsTmp.Filter = "诊断类型=1 And 记录来源<>3" '其它来源的作为缺省显示
            
            Do While Not rsTmp.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, col类型) = "西医"
                If Mid(rsTmp!诊断描述, 1, 1) <> "(" Or (Val(rsTmp!诊断id & "") = 0 And Val(rsTmp!疾病id & "") = 0) Then '中医的诊断描述后面加了（候症），所以只判断第一个字符
                    '由于疾病编码和诊断可以对应，如果两个都不为空的时候，先判断疾病编码，先取疾病编码
                    If Val(rsTmp!疾病id & "") <> 0 Then
                        .TextMatrix(.Rows - 1, col诊断编码) = Nvl(rsTmp!疾病编码)
                    ElseIf Val(rsTmp!诊断id & "") <> 0 Then
                        .TextMatrix(.Rows - 1, col诊断编码) = Nvl(rsTmp!诊断编码)
                    Else
                        .TextMatrix(.Rows - 1, col诊断编码) = ""
                    End If
                    .TextMatrix(.Rows - 1, col诊断描述) = rsTmp!诊断描述
                Else
                    .TextMatrix(.Rows - 1, col诊断编码) = Mid(rsTmp!诊断描述, 2, InStr(rsTmp!诊断描述, ")") - 2)
                    .TextMatrix(.Rows - 1, col诊断描述) = Mid(rsTmp!诊断描述, InStr(rsTmp!诊断描述, ")") + 1)
                End If
                .TextMatrix(.Rows - 1, col发病时间) = Nvl(rsTmp!发病时间)
                .TextMatrix(.Rows - 1, col疑诊) = IIf(Nvl(rsTmp!是否疑诊, 0) = 1, "？", "")
                rsTmp.MoveNext
            Loop

            '中医诊断
            rsTmp.Filter = "诊断类型=11 And 记录来源=3"
            If rsTmp.EOF Then rsTmp.Filter = "诊断类型=11 And 记录来源<>3"
            If rsTmp.EOF Then .ColHidden(col类型) = True
            Do While Not rsTmp.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, col类型) = "中医"
                If Mid(rsTmp!诊断描述, 1, 1) <> "(" Or (Val(rsTmp!诊断id & "") = 0 And Val(rsTmp!疾病id & "") = 0) Then '中医的诊断描述后面加了（候症），所以只判断第一个字符
                    '由于疾病编码和诊断可以对应，如果两个都不为空的时候，先判断疾病编码，先取疾病编码
                    If Val(rsTmp!疾病id & "") <> 0 Then
                        .TextMatrix(.Rows - 1, col诊断编码) = Nvl(rsTmp!疾病编码)
                    ElseIf Val(rsTmp!诊断id & "") <> 0 Then
                        .TextMatrix(.Rows - 1, col诊断编码) = Nvl(rsTmp!诊断编码)
                    Else
                        .TextMatrix(.Rows - 1, col诊断编码) = ""
                    End If
                    .TextMatrix(.Rows - 1, col诊断描述) = rsTmp!诊断描述
                Else
                    .TextMatrix(.Rows - 1, col诊断编码) = Mid(rsTmp!诊断描述, 2, InStr(rsTmp!诊断描述, ")") - 2)
                    .TextMatrix(.Rows - 1, col诊断描述) = Mid(rsTmp!诊断描述, InStr(rsTmp!诊断描述, ")") + 1)
                End If
                .TextMatrix(.Rows - 1, col发病时间) = Nvl(rsTmp!发病时间)
                    
                .TextMatrix(.Rows - 1, col疑诊) = IIf(Nvl(rsTmp!是否疑诊, 0) = 1, "？", "")
                rsTmp.MoveNext
            Loop
            
            .Cell(flexcpForeColor, .FixedRows, col疑诊, .Rows - 1, col疑诊) = vbRed
            .Row = .FixedRows: .Col = col诊断描述
        End With
    End If
    
    mblnCheck = False
    LoadMedRec = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub chkEdit_Click(Index As Integer)
    If Not mblnCheck Then
        mblnCheck = True
        chkEdit(Index).Value = IIf(chkEdit(Index).Value = 1, 0, 1)
        mblnCheck = False
    End If
End Sub

Private Sub Form_Activate()
    Call Form_Resize
End Sub

Private Sub Form_Load()
    Me.BackColor = fraBack.BackColor
    '滚动条尺寸
    vsc.Width = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    hsc.Height = GetSystemMetrics(SM_CXHSCROLL) * Screen.TwipsPerPixelY
    fraVH.Width = vsc.Width: fraVH.Height = hsc.Height
    fraBack.Left = 0: fraBack.Top = 0
    picBack.BackColor = fraBack.BackColor
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picBack.Left = 0
    picBack.Top = 0
    picBack.Width = Me.ScaleWidth
    picBack.Height = Me.ScaleHeight
    
    Call SetScrollbar
    
    If hsc.Visible Then
        hsc.Left = 0
        hsc.Top = picBack.ScaleHeight - hsc.Height
        hsc.Width = picBack.ScaleWidth - IIf(vsc.Visible, vsc.Width, 0)
    End If
    If vsc.Visible Then
        vsc.Top = 0
        vsc.Left = picBack.ScaleWidth - vsc.Width
        vsc.Height = picBack.ScaleHeight - IIf(hsc.Visible, hsc.Height, 0)
    End If
    If fraVH.Visible Then
        fraVH.Left = vsc.Left
        fraVH.Top = hsc.Top
        fraVH.Refresh
    End If
End Sub

Private Sub hsc_Change()
    Call hsc_Scroll
End Sub

Private Sub picSize_Click(Index As Integer)
    picSize(Index).Tag = IIf(Val(picSize(Index).Tag) = 0, 1, 0)
    Call SetPageHeight
    Call Form_Resize
    If Not vsc.Visible Then fraBack.Top = 0
    If Not hsc.Visible Then fraBack.Left = 0
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

Private Sub txtEdit_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtEdit(Index))
End Sub

Private Sub SetPageHeight()
'功能：根据页面收缩与展开状态设置界面尺寸
'说明：Tag=1表示收缩
    Dim i As Long, intCurIdx As Integer
    
    For i = 0 To fraInfo.UBound
        If Val(picSize(i).Tag) = 0 Then
            fraInfo(i).Height = Val(fraInfo(i).Tag)
            Set picSize(i).Picture = imgSize.ListImages("-").Picture
        Else
            fraInfo(i).Height = 225
            Set picSize(i).Picture = imgSize.ListImages("+").Picture
        End If
    Next
    
    intCurIdx = 0
    For i = 1 To fraInfo.UBound
        If fraInfo(i).Enabled Then
            fraInfo(i).Top = fraInfo(intCurIdx).Top + fraInfo(intCurIdx).Height + 100
            intCurIdx = i
        End If
    Next
    fraBack.Height = fraInfo(intCurIdx).Top + fraInfo(intCurIdx).Height + fraInfo(0).Top
End Sub

Private Sub SetScrollbar()
'功能：根据当前窗体尺寸设置滚动条可见性及相关属性
    If fraBack.Width + IIf(vsc.Visible, vsc.Width, 0) <= picBack.ScaleWidth Then
        hsc.Visible = False
    Else
        hsc.Min = 0
        hsc.SmallChange = 5
        hsc.LargeChange = 50
        If Not hsc.Visible Then hsc.Value = 0
        hsc.Visible = True
    End If
    
    If fraBack.Height + IIf(hsc.Visible, hsc.Height, 0) <= picBack.ScaleHeight Then
        vsc.Visible = False
    Else
        vsc.Min = 0
        vsc.SmallChange = 5
        vsc.LargeChange = 50
        If Not vsc.Visible Then vsc.Value = 0
        vsc.Visible = True
    End If
    
    If hsc.Visible Then
        hsc.Max = (picBack.ScaleWidth - fraBack.Width - IIf(vsc.Visible, vsc.Width, 0)) / Screen.TwipsPerPixelX
    End If
    
    If vsc.Visible Then
        vsc.Max = (picBack.ScaleHeight - fraBack.Height - IIf(hsc.Visible, hsc.Height, 0)) / Screen.TwipsPerPixelY
    End If
    
    fraVH.Visible = vsc.Visible And hsc.Visible
End Sub

