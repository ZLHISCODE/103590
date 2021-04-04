VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArchiveInMedRec_HN 
   BorderStyle     =   0  'None
   ClientHeight    =   7665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7515
      Left            =   120
      ScaleHeight     =   7515
      ScaleWidth      =   10245
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   120
      Width           =   10245
      Begin VB.HScrollBar hsc 
         Height          =   255
         Left            =   90
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   7200
         Visible         =   0   'False
         Width           =   9705
      End
      Begin VB.Frame fraVH 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9840
         TabIndex        =   84
         Top             =   7200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.VScrollBar vsc 
         Height          =   6975
         Left            =   9840
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComctlLib.ImageList imgSize 
         Left            =   960
         Top             =   5190
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
               Picture         =   "frmArchiveInMedRec_HN.frx":0000
               Key             =   "-"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmArchiveInMedRec_HN.frx":04EA
               Key             =   "+"
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraBack 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6975
         Left            =   90
         TabIndex        =   85
         Top             =   120
         Width           =   9705
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   住院情况 "
            ForeColor       =   &H00FF0000&
            Height          =   6090
            Index           =   4
            Left            =   120
            TabIndex        =   122
            Tag             =   "6090"
            Top             =   120
            Width           =   9495
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "疑难病例"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   9
               Left            =   6780
               TabIndex        =   337
               Top             =   1763
               Width           =   1020
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   92
               Left            =   4170
               Locked          =   -1  'True
               TabIndex        =   305
               Top             =   2467
               Width           =   2160
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   91
               Left            =   975
               Locked          =   -1  'True
               TabIndex        =   303
               Top             =   2467
               Width           =   1800
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   117
               Left            =   4200
               Locked          =   -1  'True
               TabIndex        =   274
               Top             =   5640
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   116
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   272
               Top             =   5640
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   94
               Left            =   3720
               Locked          =   -1  'True
               TabIndex        =   269
               Top             =   2820
               Width           =   5295
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   93
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   268
               Top             =   2820
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   82
               Left            =   7530
               Locked          =   -1  'True
               TabIndex        =   266
               Top             =   1062
               Width           =   1560
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   79
               Left            =   7530
               Locked          =   -1  'True
               TabIndex        =   264
               Top             =   711
               Width           =   1560
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   76
               Left            =   7530
               Locked          =   -1  'True
               TabIndex        =   262
               Top             =   360
               Width           =   1560
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   81
               Left            =   4170
               Locked          =   -1  'True
               TabIndex        =   260
               Top             =   1062
               Width           =   1440
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   78
               Left            =   4170
               Locked          =   -1  'True
               TabIndex        =   258
               Top             =   711
               Width           =   1440
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   75
               Left            =   4200
               Locked          =   -1  'True
               TabIndex        =   256
               Top             =   360
               Width           =   1440
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   74
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   254
               Top             =   360
               Width           =   1440
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   101
               Left            =   2190
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   251
               Text            =   "无"
               Top             =   3525
               Width           =   360
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   102
               Left            =   3120
               Locked          =   -1  'True
               TabIndex        =   250
               Top             =   3525
               Width           =   5940
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   90
               Left            =   7530
               Locked          =   -1  'True
               TabIndex        =   213
               Top             =   2467
               Width           =   1080
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   103
               Left            =   1155
               Locked          =   -1  'True
               TabIndex        =   210
               Top             =   3885
               Width           =   720
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   100
               Left            =   8010
               Locked          =   -1  'True
               TabIndex        =   208
               Top             =   3180
               Width           =   435
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   99
               Left            =   7080
               Locked          =   -1  'True
               TabIndex        =   206
               Top             =   3180
               Width           =   435
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   98
               Left            =   6120
               Locked          =   -1  'True
               TabIndex        =   203
               Top             =   3180
               Width           =   480
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   97
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   201
               Top             =   3180
               Width           =   435
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   95
               Left            =   2880
               Locked          =   -1  'True
               TabIndex        =   196
               Top             =   3180
               Width           =   480
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   114
               Left            =   4200
               Locked          =   -1  'True
               TabIndex        =   187
               Top             =   5280
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   115
               Left            =   7530
               Locked          =   -1  'True
               TabIndex        =   80
               Top             =   5280
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   113
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   79
               Top             =   5280
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   84
               Left            =   4170
               Locked          =   -1  'True
               TabIndex        =   64
               Top             =   1413
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   89
               Left            =   4170
               Locked          =   -1  'True
               TabIndex        =   69
               Top             =   2115
               Width           =   1080
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   88
               Left            =   915
               TabIndex        =   68
               Top             =   2115
               Width           =   1080
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   87
               Left            =   7530
               Locked          =   -1  'True
               TabIndex        =   67
               Top             =   2115
               Width           =   1080
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   86
               Left            =   4170
               Locked          =   -1  'True
               TabIndex        =   66
               Top             =   1764
               Width           =   1080
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   85
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   65
               Top             =   1764
               Width           =   1080
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "示教病案"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   13
               Left            =   6780
               TabIndex        =   60
               Top             =   1406
               Width           =   1020
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   104
               Left            =   7530
               Locked          =   -1  'True
               TabIndex        =   59
               Top             =   3885
               Width           =   1440
            End
            Begin VB.CheckBox chkInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "随诊"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   15
               Left            =   3990
               TabIndex        =   58
               Top             =   3960
               Width           =   660
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   77
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   62
               Top             =   711
               Width           =   1440
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   80
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   63
               Top             =   1062
               Width           =   1440
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   105
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   71
               Top             =   4230
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   107
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   73
               Top             =   4575
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   110
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   76
               Top             =   4935
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   108
               Left            =   4200
               Locked          =   -1  'True
               TabIndex        =   74
               Top             =   4575
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   111
               Left            =   4200
               Locked          =   -1  'True
               TabIndex        =   77
               Top             =   4935
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   106
               Left            =   7530
               Locked          =   -1  'True
               TabIndex        =   72
               Top             =   4230
               Width           =   1575
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   109
               Left            =   7530
               Locked          =   -1  'True
               TabIndex        =   75
               Top             =   4560
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   112
               Left            =   7560
               Locked          =   -1  'True
               TabIndex        =   78
               Top             =   4935
               Width           =   1335
            End
            Begin VB.PictureBox picSize 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   4
               Left            =   180
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   123
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   83
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   70
               Top             =   1413
               Width           =   1455
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "科研病案"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   14
               Left            =   8115
               TabIndex        =   61
               Top             =   1406
               Width           =   1020
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   96
               Left            =   3690
               Locked          =   -1  'True
               TabIndex        =   199
               Top             =   3120
               Width           =   435
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   92
               X1              =   4170
               X2              =   6360
               Y1              =   2640
               Y2              =   2640
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "其他医学警示"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   92
               Left            =   3060
               TabIndex        =   306
               Top             =   2467
               Width           =   1080
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   91
               X1              =   975
               X2              =   2880
               Y1              =   2640
               Y2              =   2640
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医学警示"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   91
               Left            =   240
               TabIndex        =   304
               Top             =   2467
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   117
               X1              =   4200
               X2              =   5625
               Y1              =   5820
               Y2              =   5820
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病案质量"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   117
               Left            =   3420
               TabIndex        =   275
               Top             =   5640
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   116
               X1              =   915
               X2              =   2340
               Y1              =   5820
               Y2              =   5820
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "质控日期"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   116
               Left            =   180
               TabIndex        =   273
               Top             =   5640
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "转入机构"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   94
               Left            =   2940
               TabIndex        =   271
               Top             =   2820
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出院方式"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   93
               Left            =   180
               TabIndex        =   270
               Top             =   2820
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   93
               X1              =   915
               X2              =   2400
               Y1              =   3000
               Y2              =   3000
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   94
               X1              =   3690
               X2              =   9120
               Y1              =   3000
               Y2              =   3000
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   82
               X1              =   7530
               X2              =   9120
               Y1              =   1245
               Y2              =   1245
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "生育状况"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   82
               Left            =   6780
               TabIndex        =   267
               Top             =   1065
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   79
               X1              =   7530
               X2              =   9120
               Y1              =   885
               Y2              =   885
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "发病时间"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   79
               Left            =   6780
               TabIndex        =   265
               Top             =   705
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   76
               X1              =   7530
               X2              =   9120
               Y1              =   540
               Y2              =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输血前9项检查"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   76
               Left            =   6330
               TabIndex        =   263
               Top             =   360
               Width           =   1170
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   81
               X1              =   4170
               X2              =   5640
               Y1              =   1245
               Y2              =   1245
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HIV-Ab"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   81
               Left            =   3600
               TabIndex        =   261
               Top             =   1065
               Width           =   540
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   78
               X1              =   4170
               X2              =   5640
               Y1              =   885
               Y2              =   885
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HCV-Ab"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   78
               Left            =   3600
               TabIndex        =   259
               Top             =   705
               Width           =   540
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   75
               X1              =   4200
               X2              =   5670
               Y1              =   540
               Y2              =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HBsAg"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   75
               Left            =   3690
               TabIndex        =   257
               Top             =   360
               Width           =   450
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   74
               X1              =   915
               X2              =   2400
               Y1              =   540
               Y2              =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病例分型"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   74
               Left            =   180
               TabIndex        =   255
               Top             =   360
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   102
               X1              =   3000
               X2              =   9120
               Y1              =   3705
               Y2              =   3705
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出院31天内再入院计划"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   101
               Left            =   180
               TabIndex        =   253
               Top             =   3525
               Width           =   1800
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   101
               X1              =   2070
               X2              =   2650
               Y1              =   3705
               Y2              =   3705
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "目的"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   102
               Left            =   2715
               TabIndex        =   252
               Top             =   3525
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ml"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   190
               Left            =   8640
               TabIndex        =   215
               Top             =   2467
               Width           =   180
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "自体回收"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   90
               Left            =   6780
               TabIndex        =   214
               Top             =   2467
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   90
               X1              =   7530
               X2              =   8640
               Y1              =   2640
               Y2              =   2640
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   103
               X1              =   1155
               X2              =   1965
               Y1              =   4065
               Y2              =   4065
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "呼吸机使用"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   103
               Left            =   180
               TabIndex        =   212
               Top             =   3885
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "小时"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   203
               Left            =   2145
               TabIndex        =   211
               Top             =   3885
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "分钟"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   100
               Left            =   8475
               TabIndex        =   209
               Top             =   3180
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   100
               X1              =   7875
               X2              =   8470
               Y1              =   3360
               Y2              =   3360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "小时"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   99
               Left            =   7515
               TabIndex        =   207
               Top             =   3180
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   99
               X1              =   6930
               X2              =   7500
               Y1              =   3360
               Y2              =   3360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "天"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   98
               Left            =   6780
               TabIndex        =   205
               Top             =   3180
               Width           =   180
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   98
               X1              =   6120
               X2              =   6700
               Y1              =   3360
               Y2              =   3360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院后"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   198
               Left            =   5580
               TabIndex        =   204
               Top             =   3180
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "分钟"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   97
               Left            =   5115
               TabIndex        =   202
               Top             =   3180
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   97
               X1              =   4485
               X2              =   5055
               Y1              =   3360
               Y2              =   3360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   96
               X1              =   3600
               X2              =   4200
               Y1              =   3360
               Y2              =   3360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "天"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   95
               Left            =   3420
               TabIndex        =   198
               Top             =   3180
               Width           =   180
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   95
               X1              =   2880
               X2              =   3435
               Y1              =   3360
               Y2              =   3360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "颅脑损伤患者昏迷时间;   入院前"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   195
               Left            =   180
               TabIndex        =   197
               Top             =   3180
               Width           =   2700
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "责任护士"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   114
               Left            =   3420
               TabIndex        =   188
               Top             =   5280
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   114
               X1              =   4200
               X2              =   5625
               Y1              =   5460
               Y2              =   5460
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   115
               X1              =   7530
               X2              =   9120
               Y1              =   5460
               Y2              =   5460
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   113
               X1              =   915
               X2              =   2340
               Y1              =   5460
               Y2              =   5460
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "质控医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   113
               Left            =   180
               TabIndex        =   172
               Top             =   5280
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "质控护士"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   115
               Left            =   6780
               TabIndex        =   171
               Top             =   5280
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输血反应"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   84
               Left            =   3420
               TabIndex        =   170
               Top             =   1410
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   84
               X1              =   4170
               X2              =   5595
               Y1              =   1590
               Y2              =   1590
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "实习医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   112
               Left            =   6780
               TabIndex        =   144
               Top             =   4935
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "研究生医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   111
               Left            =   3240
               TabIndex        =   143
               Top             =   4935
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "进修医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   110
               Left            =   180
               TabIndex        =   142
               Top             =   4935
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   109
               Left            =   6780
               TabIndex        =   141
               Top             =   4575
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "主治医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   108
               Left            =   3420
               TabIndex        =   140
               Top             =   4575
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "主任(副主任)医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   107
               Left            =   180
               TabIndex        =   139
               Top             =   4575
               Width           =   1440
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "科主任"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   106
               Left            =   6960
               TabIndex        =   138
               Top             =   4230
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门诊医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   105
               Left            =   180
               TabIndex        =   137
               Top             =   4230
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输其他"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   89
               Left            =   3600
               TabIndex        =   136
               Top             =   2115
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ml"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   188
               Left            =   2160
               TabIndex        =   135
               Top             =   2115
               Width           =   180
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输全血"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   88
               Left            =   360
               TabIndex        =   134
               Top             =   2115
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ml"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   187
               Left            =   8640
               TabIndex        =   133
               Top             =   2115
               Width           =   180
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输血浆"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   87
               Left            =   6960
               TabIndex        =   132
               Top             =   2115
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "单位"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   186
               Left            =   5400
               TabIndex        =   131
               Top             =   1770
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输血小板"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   86
               Left            =   3420
               TabIndex        =   130
               Top             =   1770
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "单位"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   185
               Left            =   2160
               TabIndex        =   129
               Top             =   1764
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输红细胞"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   85
               Left            =   180
               TabIndex        =   128
               Top             =   1764
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rh"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   80
               Left            =   720
               TabIndex        =   127
               Top             =   1062
               Width           =   180
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "血型"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   77
               Left            =   540
               TabIndex        =   126
               Top             =   711
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "随诊期限"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   104
               Left            =   6780
               TabIndex        =   125
               Top             =   3885
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   104
               X1              =   7530
               X2              =   9105
               Y1              =   4065
               Y2              =   4065
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   77
               X1              =   915
               X2              =   2400
               Y1              =   885
               Y2              =   885
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   85
               X1              =   915
               X2              =   2085
               Y1              =   1950
               Y2              =   1950
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   88
               X1              =   915
               X2              =   2085
               Y1              =   2310
               Y2              =   2310
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   80
               X1              =   915
               X2              =   2400
               Y1              =   1245
               Y2              =   1245
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   86
               X1              =   4170
               X2              =   5340
               Y1              =   1950
               Y2              =   1950
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   89
               X1              =   4170
               X2              =   5595
               Y1              =   2310
               Y2              =   2310
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   87
               X1              =   7530
               X2              =   8640
               Y1              =   2310
               Y2              =   2310
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   105
               X1              =   915
               X2              =   2340
               Y1              =   4410
               Y2              =   4410
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   107
               X1              =   1680
               X2              =   3105
               Y1              =   4755
               Y2              =   4755
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   110
               X1              =   915
               X2              =   2340
               Y1              =   5115
               Y2              =   5115
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   108
               X1              =   4200
               X2              =   5625
               Y1              =   4755
               Y2              =   4755
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   111
               X1              =   4200
               X2              =   5625
               Y1              =   5115
               Y2              =   5115
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   106
               X1              =   7530
               X2              =   9120
               Y1              =   4410
               Y2              =   4410
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   109
               X1              =   7530
               X2              =   9120
               Y1              =   4755
               Y2              =   4755
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   112
               X1              =   7530
               X2              =   9120
               Y1              =   5115
               Y2              =   5115
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   83
               X1              =   915
               X2              =   2400
               Y1              =   1590
               Y2              =   1590
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输液反应"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   83
               Left            =   180
               TabIndex        =   124
               Top             =   1413
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "小时"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   96
               Left            =   4155
               TabIndex        =   200
               Top             =   3180
               Width           =   360
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   其他 "
            ForeColor       =   &H00FF0000&
            Height          =   6810
            Index           =   6
            Left            =   120
            TabIndex        =   280
            Tag             =   "6810"
            Top             =   120
            Width           =   9495
            Begin VB.Frame fraHNAddtion 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "附加"
               ForeColor       =   &H80000008&
               Height          =   1455
               Left            =   120
               TabIndex        =   309
               Top             =   180
               Width           =   9255
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   129
                  Left            =   3930
                  Locked          =   -1  'True
                  TabIndex        =   336
                  Top             =   720
                  Width           =   495
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   133
                  Left            =   8040
                  Locked          =   -1  'True
                  TabIndex        =   335
                  Top             =   1080
                  Width           =   735
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   132
                  Left            =   8040
                  Locked          =   -1  'True
                  TabIndex        =   333
                  Top             =   720
                  Width           =   855
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   131
                  Left            =   8040
                  Locked          =   -1  'True
                  TabIndex        =   331
                  Top             =   360
                  Width           =   735
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   130
                  Left            =   1560
                  Locked          =   -1  'True
                  TabIndex        =   329
                  Top             =   1080
                  Width           =   495
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   125
                  Left            =   2760
                  Locked          =   -1  'True
                  TabIndex        =   327
                  Top             =   360
                  Width           =   615
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   124
                  Left            =   2160
                  Locked          =   -1  'True
                  TabIndex        =   325
                  Top             =   360
                  Width           =   375
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   123
                  Left            =   1440
                  Locked          =   -1  'True
                  MaxLength       =   9
                  TabIndex        =   323
                  Text            =   "无"
                  Top             =   360
                  Width           =   375
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   1310
                  Left            =   3930
                  Locked          =   -1  'True
                  MaxLength       =   9
                  TabIndex        =   321
                  Top             =   600
                  Width           =   495
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   128
                  Left            =   3210
                  Locked          =   -1  'True
                  TabIndex        =   319
                  Top             =   720
                  Width           =   495
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   127
                  Left            =   2400
                  Locked          =   -1  'True
                  TabIndex        =   317
                  Top             =   720
                  Width           =   495
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   126
                  Left            =   1080
                  Locked          =   -1  'True
                  TabIndex        =   315
                  Top             =   720
                  Width           =   975
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "使用抗生素"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   21
                  Left            =   4560
                  TabIndex        =   312
                  Top             =   1073
                  Width           =   1275
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "细菌培养标本送检"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   19
                  Left            =   4560
                  TabIndex        =   311
                  Top             =   713
                  Width           =   1755
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "单病种管理"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   20
                  Left            =   4560
                  TabIndex        =   310
                  Top             =   353
                  Width           =   1275
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   129
                  X1              =   3840
                  X2              =   4440
                  Y1              =   900
                  Y2              =   900
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   133
                  X1              =   7920
                  X2              =   8880
                  Y1              =   1260
                  Y2              =   1260
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   132
                  X1              =   7920
                  X2              =   8880
                  Y1              =   900
                  Y2              =   900
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   131
                  X1              =   7920
                  X2              =   8880
                  Y1              =   600
                  Y2              =   600
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   128
                  X1              =   3120
                  X2              =   3720
                  Y1              =   900
                  Y2              =   900
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   127
                  X1              =   2280
                  X2              =   2880
                  Y1              =   900
                  Y2              =   900
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   126
                  X1              =   960
                  X2              =   2160
                  Y1              =   900
                  Y2              =   900
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   125
                  X1              =   2760
                  X2              =   3480
                  Y1              =   535
                  Y2              =   535
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   124
                  X1              =   2040
                  X2              =   2520
                  Y1              =   535
                  Y2              =   535
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   123
                  X1              =   1320
                  X2              =   1800
                  Y1              =   535
                  Y2              =   535
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   130
                  X1              =   1560
                  X2              =   2160
                  Y1              =   1255
                  Y2              =   1255
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "实施&DRGs管理"
                  Height          =   180
                  Index           =   133
                  Left            =   6840
                  TabIndex        =   334
                  Top             =   1080
                  Width           =   1080
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "法定传染病"
                  Height          =   180
                  Index           =   132
                  Left            =   7020
                  TabIndex        =   332
                  Top             =   720
                  Width           =   900
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "实施临床路径管理"
                  Height          =   180
                  Index           =   131
                  Left            =   6480
                  TabIndex        =   330
                  Top             =   360
                  Width           =   1440
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "小时"
                  Enabled         =   0   'False
                  Height          =   180
                  Index           =   125
                  Left            =   3480
                  TabIndex        =   328
                  Top             =   360
                  Width           =   360
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "天"
                  Enabled         =   0   'False
                  Height          =   180
                  Index           =   224
                  Left            =   2580
                  TabIndex        =   326
                  Top             =   360
                  Width           =   180
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "有"
                  Enabled         =   0   'False
                  Height          =   180
                  Index           =   124
                  Left            =   1860
                  TabIndex        =   324
                  Top             =   360
                  Width           =   180
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "实施重症监护"
                  Height          =   180
                  Index           =   123
                  Left            =   240
                  TabIndex        =   322
                  Top             =   360
                  Width           =   1080
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "N"
                  Height          =   180
                  Index           =   129
                  Left            =   3720
                  TabIndex        =   320
                  Top             =   720
                  Width           =   90
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "M"
                  Height          =   180
                  Index           =   128
                  Left            =   3000
                  TabIndex        =   318
                  Top             =   720
                  Width           =   90
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "T"
                  Height          =   180
                  Index           =   127
                  Left            =   2190
                  TabIndex        =   316
                  Top             =   720
                  Width           =   90
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "肿瘤分期"
                  Height          =   180
                  Index           =   126
                  Left            =   240
                  TabIndex        =   314
                  Top             =   720
                  Width           =   720
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "新生儿&Apgar评分       分"
                  Height          =   180
                  Index           =   130
                  Left            =   240
                  TabIndex        =   313
                  Top             =   1080
                  Width           =   2160
               End
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "3.彩色多普勒"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   18
               Left            =   2280
               TabIndex        =   298
               Top             =   5280
               Width           =   1395
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "2.MRI"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   17
               Left            =   1200
               TabIndex        =   297
               Top             =   5280
               Width           =   915
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "1.CT"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   296
               Top             =   5280
               Width           =   915
            End
            Begin VB.Frame fraSplit 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00FFFFFF&
               Height          =   75
               Index           =   0
               Left            =   1200
               TabIndex        =   294
               Top             =   4980
               Width           =   4455
            End
            Begin VB.Frame fraSplit 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00FFFFFF&
               Height          =   75
               Index           =   1
               Left            =   1200
               TabIndex        =   292
               Top             =   1973
               Width           =   4455
            End
            Begin VB.Frame fraAdvEvent 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "不良事件"
               ForeColor       =   &H80000008&
               Height          =   2595
               Left            =   5760
               TabIndex        =   286
               Top             =   3975
               Width           =   3615
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   121
                  Left            =   1440
                  Locked          =   -1  'True
                  TabIndex        =   302
                  Top             =   2220
                  Width           =   1995
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   120
                  Left            =   1440
                  Locked          =   -1  'True
                  TabIndex        =   301
                  Top             =   1860
                  Width           =   1995
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   119
                  Left            =   2640
                  Locked          =   -1  'True
                  TabIndex        =   300
                  Top             =   1500
                  Width           =   915
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   118
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   299
                  Top             =   1500
                  Width           =   975
               End
               Begin VB.ListBox lstAdvEvent 
                  Height          =   960
                  ItemData        =   "frmArchiveInMedRec_HN.frx":09D4
                  Left            =   120
                  List            =   "frmArchiveInMedRec_HN.frx":09D6
                  TabIndex        =   287
                  Top             =   240
                  Width           =   3405
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   121
                  X1              =   1440
                  X2              =   3600
                  Y1              =   2400
                  Y2              =   2400
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   120
                  X1              =   1440
                  X2              =   3600
                  Y1              =   2040
                  Y2              =   2040
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   119
                  X1              =   2655
                  X2              =   3605
                  Y1              =   1680
                  Y2              =   1680
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   118
                  X1              =   1200
                  X2              =   2260
                  Y1              =   1680
                  Y2              =   1680
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "分期"
                  Height          =   180
                  Index           =   119
                  Left            =   2280
                  TabIndex        =   291
                  Top             =   1500
                  Width           =   360
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "压疮发生期间"
                  Height          =   180
                  Index           =   118
                  Left            =   120
                  TabIndex        =   290
                  Top             =   1500
                  Width           =   1080
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "跌倒或坠床原因"
                  Height          =   180
                  Index           =   121
                  Left            =   120
                  TabIndex        =   289
                  Top             =   2220
                  Width           =   1260
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "跌倒或坠床伤害"
                  Height          =   180
                  Index           =   120
                  Left            =   120
                  TabIndex        =   288
                  Top             =   1860
                  Width           =   1260
               End
            End
            Begin VB.Frame fraInfection 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "感染因素"
               ForeColor       =   &H80000008&
               Height          =   1815
               Left            =   5760
               TabIndex        =   284
               Top             =   1920
               Width           =   3615
               Begin VB.ListBox lstInfection 
                  Height          =   1320
                  ItemData        =   "frmArchiveInMedRec_HN.frx":09D8
                  Left            =   120
                  List            =   "frmArchiveInMedRec_HN.frx":09DA
                  TabIndex        =   285
                  Top             =   240
                  Width           =   3405
               End
            End
            Begin VB.PictureBox picSize 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   6
               Left            =   180
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   281
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfMain 
               Height          =   2490
               Left            =   120
               TabIndex        =   283
               Top             =   2160
               Width           =   5565
               _cx             =   9816
               _cy             =   4392
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
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
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   ""
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
            Begin VSFlex8Ctl.VSFlexGrid vsTSJC 
               Height          =   930
               Left            =   120
               TabIndex        =   295
               Top             =   5640
               Width           =   5565
               _cx             =   9816
               _cy             =   1640
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
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
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   2
               FixedRows       =   0
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_HN.frx":09DC
               ScrollTrack     =   -1  'True
               ScrollBars      =   0
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
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "特殊检查情况"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   1330
               Left            =   120
               TabIndex        =   293
               Top             =   4920
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病案附加项目"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   1280
               Left            =   120
               TabIndex        =   282
               Top             =   1920
               Width           =   1080
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   放疗与化疗 "
            ForeColor       =   &H00FF0000&
            Height          =   5010
            Index           =   5
            Left            =   120
            TabIndex        =   224
            Tag             =   "5010"
            Top             =   120
            Width           =   9495
            Begin VB.PictureBox picSize 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   5
               Left            =   180
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   225
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VSFlex8Ctl.VSFlexGrid vsChemotherapy 
               Height          =   1635
               Left            =   120
               TabIndex        =   278
               Top             =   480
               Width           =   9240
               _cx             =   16298
               _cy             =   2884
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
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
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_HN.frx":0A4A
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
            Begin VSFlex8Ctl.VSFlexGrid vsRadiotherapy 
               Height          =   2205
               Left            =   120
               TabIndex        =   279
               Top             =   2640
               Width           =   9240
               _cx             =   16298
               _cy             =   3889
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
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
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_HN.frx":0B60
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
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "放疗记录信息"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   1250
               Left            =   120
               TabIndex        =   277
               Top             =   2400
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "化疗记录信息"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   1240
               Left            =   120
               TabIndex        =   276
               Top             =   240
               Width           =   1080
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   过敏与手术 "
            ForeColor       =   &H00FF0000&
            Height          =   3345
            Index           =   3
            Left            =   120
            TabIndex        =   145
            Tag             =   "3705"
            Top             =   120
            Width           =   9495
            Begin VB.PictureBox picSize 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   3
               Left            =   180
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   146
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VSFlex8Ctl.VSFlexGrid vsOPS 
               Height          =   1335
               Left            =   165
               TabIndex        =   57
               Top             =   1800
               Width           =   9180
               _cx             =   16192
               _cy             =   2355
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
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
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   14
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmArchiveInMedRec_HN.frx":0C86
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
            Begin VSFlex8Ctl.VSFlexGrid vsAller 
               Height          =   1335
               Left            =   165
               TabIndex        =   56
               Top             =   300
               Width           =   9180
               _cx             =   16192
               _cy             =   2355
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
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
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_HN.frx":0E2A
               ScrollTrack     =   -1  'True
               ScrollBars      =   0
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
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   中医诊断 "
            ForeColor       =   &H00FF0000&
            Height          =   4170
            Index           =   2
            Left            =   120
            TabIndex        =   147
            Tag             =   "4170"
            Top             =   120
            Width           =   9495
            Begin VB.Frame fraSub 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   " 治疗方法 "
               ForeColor       =   &H00404040&
               Height          =   1320
               Index           =   2
               Left            =   4320
               TabIndex        =   154
               Top             =   2580
               Width           =   4905
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   73
                  Left            =   3990
                  Locked          =   -1  'True
                  TabIndex        =   220
                  Top             =   960
                  Width           =   555
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   72
                  Left            =   3990
                  Locked          =   -1  'True
                  TabIndex        =   218
                  Top             =   645
                  Width           =   555
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   71
                  Left            =   3990
                  Locked          =   -1  'True
                  TabIndex        =   216
                  Top             =   330
                  Width           =   555
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   68
                  Left            =   1185
                  Locked          =   -1  'True
                  TabIndex        =   53
                  Top             =   330
                  Width           =   1035
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   69
                  Left            =   1185
                  Locked          =   -1  'True
                  TabIndex        =   54
                  Top             =   645
                  Width           =   1035
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   70
                  Left            =   1545
                  Locked          =   -1  'True
                  TabIndex        =   55
                  Top             =   960
                  Width           =   1035
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   73
                  X1              =   3960
                  X2              =   4545
                  Y1              =   1140
                  Y2              =   1140
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "辨证施护"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   73
                  Left            =   3240
                  TabIndex        =   221
                  Top             =   960
                  Width           =   720
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   72
                  X1              =   3960
                  X2              =   4545
                  Y1              =   825
                  Y2              =   825
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "使用中医诊疗技术"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   72
                  Left            =   2520
                  TabIndex        =   219
                  Top             =   645
                  Width           =   1440
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   71
                  X1              =   3960
                  X2              =   4545
                  Y1              =   510
                  Y2              =   510
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "使用中医诊疗设备"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   71
                  Left            =   2520
                  TabIndex        =   217
                  Top             =   330
                  Width           =   1440
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "治疗类别"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   68
                  Left            =   315
                  TabIndex        =   157
                  Top             =   330
                  Width           =   720
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "抢救方法"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   69
                  Left            =   315
                  TabIndex        =   156
                  Top             =   645
                  Width           =   720
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "自制中药制剂"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   70
                  Left            =   315
                  TabIndex        =   155
                  Top             =   960
                  Width           =   1080
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   68
                  X1              =   1095
                  X2              =   2220
                  Y1              =   510
                  Y2              =   510
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   69
                  X1              =   1095
                  X2              =   2220
                  Y1              =   825
                  Y2              =   825
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   70
                  X1              =   1455
                  X2              =   2580
                  Y1              =   1140
                  Y2              =   1140
               End
            End
            Begin VB.Frame fraSub 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   " 住院期间病情 "
               ForeColor       =   &H00404040&
               Height          =   1320
               Index           =   0
               Left            =   165
               TabIndex        =   153
               Top             =   2580
               Width           =   1485
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "疑难"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   8
                  Left            =   405
                  TabIndex        =   49
                  Top             =   960
                  Width           =   660
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "急症"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   7
                  Left            =   405
                  TabIndex        =   48
                  Top             =   645
                  Width           =   660
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "危重"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   6
                  Left            =   405
                  TabIndex        =   47
                  Top             =   330
                  Width           =   660
               End
            End
            Begin VB.PictureBox picSize 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   2
               Left            =   180
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   152
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.Frame fraSub 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   " 准确度 "
               ForeColor       =   &H00404040&
               Height          =   1320
               Index           =   1
               Left            =   2032
               TabIndex        =   148
               Top             =   2580
               Width           =   1905
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   67
                  Left            =   600
                  Locked          =   -1  'True
                  TabIndex        =   52
                  Top             =   960
                  Width           =   1035
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   66
                  Left            =   600
                  Locked          =   -1  'True
                  TabIndex        =   51
                  Top             =   645
                  Width           =   1035
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   65
                  Left            =   600
                  Locked          =   -1  'True
                  TabIndex        =   50
                  Top             =   330
                  Width           =   1035
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   67
                  X1              =   630
                  X2              =   1755
                  Y1              =   1140
                  Y2              =   1140
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   66
                  X1              =   630
                  X2              =   1755
                  Y1              =   825
                  Y2              =   825
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   65
                  X1              =   630
                  X2              =   1755
                  Y1              =   510
                  Y2              =   510
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "方药"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   67
                  Left            =   210
                  TabIndex        =   151
                  Top             =   960
                  Width           =   360
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "治法"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   66
                  Left            =   210
                  TabIndex        =   150
                  Top             =   645
                  Width           =   360
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "辨证"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   65
                  Left            =   210
                  TabIndex        =   149
                  Top             =   330
                  Width           =   360
               End
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   63
               Left            =   1425
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   2190
               Width           =   915
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   64
               Left            =   4020
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   2190
               Width           =   915
            End
            Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
               Height          =   1710
               Left            =   165
               TabIndex        =   44
               Top             =   270
               Width           =   9180
               _cx             =   16192
               _cy             =   3016
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
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
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   5
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_HN.frx":0E97
               ScrollTrack     =   -1  'True
               ScrollBars      =   0
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
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院与出院"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   64
               Left            =   3000
               TabIndex        =   159
               Top             =   2190
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门诊与出院"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   63
               Left            =   390
               TabIndex        =   158
               Top             =   2190
               Width           =   900
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   63
               X1              =   1335
               X2              =   2465
               Y1              =   2370
               Y2              =   2370
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   64
               X1              =   3930
               X2              =   5015
               Y1              =   2370
               Y2              =   2370
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   西医诊断 "
            ForeColor       =   &H00FF0000&
            Height          =   5355
            Index           =   1
            Left            =   120
            TabIndex        =   160
            Tag             =   "5355"
            Top             =   120
            Width           =   9495
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   56
               Left            =   7155
               Locked          =   -1  'True
               TabIndex        =   248
               Top             =   3876
               Width           =   1980
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   49
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   247
               Top             =   3120
               Width           =   1515
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   59
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   245
               Top             =   4627
               Width           =   4410
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   57
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   243
               Top             =   4248
               Width           =   1635
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   53
               Left            =   7080
               Locked          =   -1  'True
               TabIndex        =   241
               Top             =   3504
               Width           =   1875
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   48
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   239
               Top             =   3132
               Width           =   1695
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   45
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   237
               Top             =   2760
               Width           =   1660
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "医院感染作病原学检查"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   3
               Left            =   6960
               TabIndex        =   226
               Top             =   4241
               Width           =   2150
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   47
               Left            =   7080
               Locked          =   -1  'True
               TabIndex        =   222
               Top             =   2760
               Width           =   2115
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   62
               Left            =   4335
               Locked          =   -1  'True
               TabIndex        =   194
               Top             =   5010
               Width           =   3690
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   58
               Left            =   3720
               Locked          =   -1  'True
               TabIndex        =   191
               Top             =   4248
               Width           =   2970
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "新发肿瘤"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   5
               Left            =   1680
               TabIndex        =   190
               Top             =   4620
               Width           =   1020
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "死亡患者尸检"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   4
               Left            =   240
               TabIndex        =   189
               Top             =   4620
               Width           =   1485
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   61
               Left            =   2910
               Locked          =   -1  'True
               TabIndex        =   38
               Top             =   5010
               Width           =   510
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   60
               Left            =   975
               Locked          =   -1  'True
               TabIndex        =   37
               Top             =   5010
               Width           =   870
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "是否确诊"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   2
               Left            =   2760
               TabIndex        =   35
               Top             =   2753
               Width           =   1020
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   46
               Left            =   4695
               Locked          =   -1  'True
               TabIndex        =   36
               Top             =   2760
               Width           =   1680
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
               TabIndex        =   161
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   50
               Left            =   7080
               Locked          =   -1  'True
               TabIndex        =   41
               Top             =   3132
               Width           =   1755
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   55
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   43
               Top             =   3876
               Width           =   1455
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   52
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   40
               Top             =   3504
               Width           =   1575
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   54
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   3876
               Width           =   1755
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   51
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   39
               Top             =   3504
               Width           =   1755
            End
            Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
               Height          =   2385
               Left            =   135
               TabIndex        =   34
               Top             =   270
               Width           =   9240
               _cx             =   16298
               _cy             =   4207
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
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
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   9
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_HN.frx":0F9A
               ScrollTrack     =   -1  'True
               ScrollBars      =   0
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
               Index           =   56
               X1              =   7065
               X2              =   9240
               Y1              =   4050
               Y2              =   4050
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "术前与术后"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   56
               Left            =   6180
               TabIndex        =   249
               Top             =   3876
               Width           =   900
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   59
               X1              =   4605
               X2              =   9000
               Y1              =   4800
               Y2              =   4800
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医院感染病原学诊断"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   59
               Left            =   2960
               TabIndex        =   246
               Top             =   4627
               Width           =   1620
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "死亡时间"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   57
               Left            =   240
               TabIndex        =   244
               Top             =   4248
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   57
               X1              =   960
               X2              =   2760
               Y1              =   4420
               Y2              =   4420
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门诊与入院"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   53
               Left            =   6180
               TabIndex        =   242
               Top             =   3504
               Width           =   900
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   53
               X1              =   7080
               X2              =   9240
               Y1              =   3690
               Y2              =   3690
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "最高诊断依据"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   49
               Left            =   2960
               TabIndex        =   240
               Top             =   3135
               Width           =   1080
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   49
               X1              =   4080
               X2              =   5640
               Y1              =   3315
               Y2              =   3315
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   48
               X1              =   960
               X2              =   2745
               Y1              =   3310
               Y2              =   3310
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "分化程度"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   48
               Left            =   240
               TabIndex        =   238
               Top             =   3132
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   45
               X1              =   960
               X2              =   2745
               Y1              =   2940
               Y2              =   2940
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院情况"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   45
               Left            =   240
               TabIndex        =   236
               Top             =   2760
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   47
               X1              =   7080
               X2              =   9315
               Y1              =   2940
               Y2              =   2940
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病理号"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   47
               Left            =   6540
               TabIndex        =   223
               Top             =   2760
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "抢救原因"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   62
               Left            =   3480
               TabIndex        =   195
               Top             =   5010
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   62
               X1              =   4245
               X2              =   8245
               Y1              =   5190
               Y2              =   5190
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "死亡原因"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   58
               Left            =   2960
               TabIndex        =   192
               Top             =   4248
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   58
               X1              =   3720
               X2              =   6840
               Y1              =   4420
               Y2              =   4420
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "成功次数"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   61
               Left            =   2055
               TabIndex        =   169
               Top             =   5010
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "抢救次数"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   60
               Left            =   240
               TabIndex        =   168
               Top             =   5010
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "确诊日期"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   46
               Left            =   3855
               TabIndex        =   167
               Top             =   2760
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   46
               X1              =   4605
               X2              =   6390
               Y1              =   2940
               Y2              =   2940
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   60
               X1              =   960
               X2              =   1845
               Y1              =   5190
               Y2              =   5190
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   61
               X1              =   2820
               X2              =   3420
               Y1              =   5190
               Y2              =   5190
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   50
               X1              =   7080
               X2              =   9240
               Y1              =   3310
               Y2              =   3310
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   55
               X1              =   4080
               X2              =   5640
               Y1              =   4050
               Y2              =   4050
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   52
               X1              =   4080
               X2              =   5640
               Y1              =   3690
               Y2              =   3690
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   54
               X1              =   960
               X2              =   2760
               Y1              =   4050
               Y2              =   4050
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   51
               X1              =   960
               X2              =   2760
               Y1              =   3690
               Y2              =   3690
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门诊与出院"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   51
               Left            =   60
               TabIndex        =   166
               Top             =   3504
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院与出院"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   52
               Left            =   3140
               TabIndex        =   165
               Top             =   3504
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "放射与病理"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   50
               Left            =   6180
               TabIndex        =   164
               Top             =   3132
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "临床与病理"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   54
               Left            =   60
               TabIndex        =   163
               Top             =   3876
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "临床与尸检"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   55
               Left            =   3140
               TabIndex        =   162
               Top             =   3876
               Width           =   900
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   基本信息 "
            ForeColor       =   &H00FF0000&
            Height          =   6195
            Index           =   0
            Left            =   120
            TabIndex        =   86
            Tag             =   "6195"
            Top             =   120
            Width           =   9495
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   122
               Left            =   4815
               Locked          =   -1  'True
               TabIndex        =   307
               Top             =   2160
               Width           =   2805
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   10
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   235
               Top             =   1065
               Width           =   425
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   9
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   234
               Top             =   1065
               Width           =   425
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   32
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   231
               Top             =   3945
               Width           =   3075
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   37
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   229
               Top             =   5025
               Width           =   1455
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "再入院"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   0
               Left            =   5640
               TabIndex        =   227
               Top             =   338
               Width           =   915
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "入院前经外院治疗"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   1
               Left            =   6480
               TabIndex        =   193
               Top             =   4658
               Width           =   1740
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   23
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   184
               Top             =   2865
               Width           =   2805
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   24
               Left            =   4815
               Locked          =   -1  'True
               TabIndex        =   183
               Top             =   2865
               Width           =   1530
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   15
               Left            =   4815
               Locked          =   -1  'True
               TabIndex        =   181
               Top             =   1785
               Width           =   1530
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   13
               Left            =   6960
               Locked          =   -1  'True
               TabIndex        =   178
               Top             =   1425
               Width           =   1740
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   12
               Left            =   4455
               Locked          =   -1  'True
               TabIndex        =   177
               Top             =   1425
               Width           =   1050
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   11
               Left            =   2400
               Locked          =   -1  'True
               TabIndex        =   175
               Top             =   1425
               Width           =   810
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   3
               Left            =   7260
               Locked          =   -1  'True
               TabIndex        =   173
               Top             =   345
               Width           =   1395
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   0
               Left            =   1305
               Locked          =   -1  'True
               TabIndex        =   1
               Top             =   345
               Width           =   900
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   40
               Left            =   6120
               Locked          =   -1  'True
               TabIndex        =   29
               Top             =   5385
               Width           =   2010
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   39
               Left            =   3525
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   5385
               Width           =   1650
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   38
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   27
               Top             =   5385
               Width           =   1530
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   44
               Left            =   7200
               Locked          =   -1  'True
               TabIndex        =   33
               Top             =   5745
               Width           =   1695
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   43
               Left            =   5160
               Locked          =   -1  'True
               TabIndex        =   32
               Top             =   5745
               Width           =   945
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   42
               Left            =   3270
               Locked          =   -1  'True
               TabIndex        =   31
               Top             =   5745
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   41
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   30
               Top             =   5745
               Width           =   1455
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   36
               Left            =   5160
               Locked          =   -1  'True
               TabIndex        =   26
               Top             =   4665
               Width           =   945
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   35
               Left            =   3270
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   4665
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   34
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   24
               Top             =   4665
               Width           =   1455
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   31
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   23
               Top             =   3945
               Width           =   4200
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   30
               Left            =   4815
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   3585
               Width           =   1530
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   28
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   20
               Top             =   3585
               Width           =   1035
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   27
               Left            =   6960
               Locked          =   -1  'True
               TabIndex        =   19
               Top             =   3225
               Width           =   1695
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   26
               Left            =   4815
               Locked          =   -1  'True
               TabIndex        =   18
               Top             =   3225
               Width           =   1530
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   25
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   17
               Top             =   3225
               Width           =   2805
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   22
               Left            =   6960
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   2505
               Width           =   1815
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   21
               Left            =   4815
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   2505
               Width           =   1530
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   20
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   14
               Top             =   2505
               Width           =   2805
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   17
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   13
               Top             =   2145
               Width           =   2805
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   14
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   12
               Top             =   1785
               Width           =   2805
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   7
               Left            =   6330
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   705
               Width           =   690
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   6
               Left            =   4545
               Locked          =   -1  'True
               TabIndex        =   5
               Top             =   705
               Width           =   1260
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   4
               Left            =   1180
               Locked          =   -1  'True
               TabIndex        =   3
               Top             =   705
               Width           =   860
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   2
               Left            =   4635
               Locked          =   -1  'True
               TabIndex        =   2
               Top             =   345
               Width           =   285
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   1
               Left            =   3210
               Locked          =   -1  'True
               TabIndex        =   0
               Top             =   345
               Width           =   1050
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   5
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   4
               Top             =   705
               Width           =   645
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   19
               Left            =   6240
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   1065
               Width           =   975
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   16
               Left            =   7560
               Locked          =   -1  'True
               TabIndex        =   10
               Top             =   1065
               Width           =   1215
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   8
               Left            =   7650
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   705
               Width           =   1095
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   18
               Left            =   4095
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   1065
               Width           =   1530
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   33
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   11
               Top             =   4305
               Width           =   1500
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   29
               Left            =   2880
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   3585
               Width           =   1095
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
               TabIndex        =   87
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   122
               X1              =   4815
               X2              =   7695
               Y1              =   2340
               Y2              =   2340
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "其他证件"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   122
               Left            =   4080
               TabIndex        =   308
               Top             =   2160
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   10
               X1              =   2670
               X2              =   3120
               Y1              =   1240
               Y2              =   1240
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   9
               X1              =   1060
               X2              =   1580
               Y1              =   1240
               Y2              =   1240
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   32
               X1              =   5760
               X2              =   8880
               Y1              =   4125
               Y2              =   4125
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   37
               X1              =   1080
               X2              =   2670
               Y1              =   5200
               Y2              =   5200
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "身高      cm"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   10
               Left            =   2265
               TabIndex        =   233
               Top             =   1065
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "体重      kg"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   9
               Left            =   720
               TabIndex        =   232
               Top             =   1065
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "区域"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   32
               Left            =   5400
               TabIndex        =   230
               Top             =   3945
               Width           =   360
            End
            Begin VB.Label lblInfo 
               BackStyle       =   0  'Transparent
               Caption         =   "入科时间"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   37
               Left            =   330
               TabIndex        =   228
               Top             =   5025
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   24
               X1              =   4845
               X2              =   6380
               Y1              =   3040
               Y2              =   3040
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   23
               X1              =   1080
               X2              =   3960
               Y1              =   3040
               Y2              =   3040
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "户口地址"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   23
               Left            =   330
               TabIndex        =   186
               Top             =   2865
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "邮编"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   24
               Left            =   4440
               TabIndex        =   185
               Top             =   2865
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   15
               X1              =   4845
               X2              =   6375
               Y1              =   1960
               Y2              =   1960
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   16
               X1              =   7560
               X2              =   8760
               Y1              =   1245
               Y2              =   1245
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "籍贯"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   15
               Left            =   4440
               TabIndex        =   182
               Top             =   1785
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   13
               X1              =   6960
               X2              =   8760
               Y1              =   1605
               Y2              =   1605
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   12
               X1              =   4365
               X2              =   5520
               Y1              =   1600
               Y2              =   1600
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "新生儿体重"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   12
               Left            =   3480
               TabIndex        =   180
               Top             =   1425
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "新生儿入院体重"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   13
               Left            =   5700
               TabIndex        =   179
               Top             =   1425
               Width           =   1260
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   11
               X1              =   2280
               X2              =   3360
               Y1              =   1600
               Y2              =   1600
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "（年龄不足一周岁的） 年龄"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   11
               Left            =   90
               TabIndex        =   176
               Top             =   1425
               Width           =   2250
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   3
               X1              =   7170
               X2              =   8760
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病案号"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   3
               Left            =   6600
               TabIndex        =   174
               Top             =   345
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院天数"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   44
               Left            =   6480
               TabIndex        =   121
               Top             =   5745
               Width           =   720
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "───→"
               ForeColor       =   &H00808080&
               Height          =   180
               Index           =   40
               Left            =   5280
               TabIndex        =   120
               Top             =   5385
               Width           =   720
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "───→"
               ForeColor       =   &H00808080&
               Height          =   180
               Index           =   39
               Left            =   2760
               TabIndex        =   119
               Top             =   5385
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "转科情况"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   38
               Left            =   360
               TabIndex        =   118
               Top             =   5385
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病房"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   43
               Left            =   4680
               TabIndex        =   117
               Top             =   5745
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "科室"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   42
               Left            =   2805
               TabIndex        =   116
               Top             =   5745
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出院时间"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   41
               Left            =   330
               TabIndex        =   115
               Top             =   5745
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病房"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   36
               Left            =   4680
               TabIndex        =   114
               Top             =   4665
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "科室"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   35
               Left            =   2805
               TabIndex        =   113
               Top             =   4665
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院时间"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   34
               Left            =   330
               TabIndex        =   112
               Top             =   4665
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "联系人地址"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   31
               Left            =   150
               TabIndex        =   111
               Top             =   3945
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "电话"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   30
               Left            =   4440
               TabIndex        =   110
               Top             =   3585
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "关系"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   29
               Left            =   2400
               TabIndex        =   109
               Top             =   3585
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "联系人姓名"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   28
               Left            =   150
               TabIndex        =   108
               Top             =   3585
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "邮编"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   27
               Left            =   6600
               TabIndex        =   107
               Top             =   3225
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "电话"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   26
               Left            =   4440
               TabIndex        =   106
               Top             =   3225
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "工作单位"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   25
               Left            =   330
               TabIndex        =   105
               Top             =   3225
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "邮编"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   22
               Left            =   6600
               TabIndex        =   104
               Top             =   2505
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "电话"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   21
               Left            =   4440
               TabIndex        =   103
               Top             =   2505
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "现住址"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   20
               Left            =   510
               TabIndex        =   102
               Top             =   2505
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "身份证号"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   17
               Left            =   330
               TabIndex        =   101
               Top             =   2145
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出生地点"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   14
               Left            =   330
               TabIndex        =   100
               Top             =   1785
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "民族"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   16
               Left            =   7200
               TabIndex        =   99
               Top             =   1065
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "国籍"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   8
               Left            =   7170
               TabIndex        =   98
               Top             =   690
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院途径"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   33
               Left            =   330
               TabIndex        =   97
               Top             =   4305
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "职业"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   18
               Left            =   3720
               TabIndex        =   96
               Top             =   1065
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "婚姻"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   19
               Left            =   5880
               TabIndex        =   95
               Top             =   1065
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "年龄"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   7
               Left            =   5940
               TabIndex        =   94
               Top             =   690
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出生日期"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   6
               Left            =   3690
               TabIndex        =   93
               Top             =   690
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "性别"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   5
               Left            =   2265
               TabIndex        =   92
               Top             =   690
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "姓名"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   4
               Left            =   690
               TabIndex        =   91
               Top             =   705
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医疗付费方式"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   0
               Left            =   90
               TabIndex        =   90
               Top             =   345
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "第    次住院"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   2
               Left            =   4425
               TabIndex        =   89
               Top             =   345
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "健康卡号"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   1
               Left            =   2370
               TabIndex        =   88
               Top             =   345
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   1
               X1              =   3120
               X2              =   4320
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   2
               X1              =   4635
               X2              =   4925
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   0
               X1              =   1215
               X2              =   2280
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   4
               X1              =   1080
               X2              =   2040
               Y1              =   880
               Y2              =   880
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   7
               X1              =   6330
               X2              =   7080
               Y1              =   880
               Y2              =   880
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   8
               X1              =   7560
               X2              =   8760
               Y1              =   885
               Y2              =   885
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   5
               X1              =   2670
               X2              =   3480
               Y1              =   880
               Y2              =   880
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   19
               X1              =   6240
               X2              =   7200
               Y1              =   1245
               Y2              =   1245
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   6
               X1              =   4455
               X2              =   5760
               Y1              =   880
               Y2              =   880
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   18
               X1              =   4125
               X2              =   5655
               Y1              =   1240
               Y2              =   1240
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   33
               X1              =   1080
               X2              =   2670
               Y1              =   4480
               Y2              =   4480
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   14
               X1              =   1080
               X2              =   3975
               Y1              =   1960
               Y2              =   1960
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   17
               X1              =   1080
               X2              =   3960
               Y1              =   2320
               Y2              =   2320
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   20
               X1              =   1080
               X2              =   3975
               Y1              =   2680
               Y2              =   2680
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   25
               X1              =   1080
               X2              =   3975
               Y1              =   3400
               Y2              =   3400
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   31
               X1              =   1080
               X2              =   5280
               Y1              =   4120
               Y2              =   4120
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   21
               X1              =   4845
               X2              =   6380
               Y1              =   2680
               Y2              =   2680
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   26
               X1              =   4845
               X2              =   6360
               Y1              =   3400
               Y2              =   3400
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   30
               X1              =   4845
               X2              =   6360
               Y1              =   3760
               Y2              =   3760
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   22
               X1              =   6960
               X2              =   8760
               Y1              =   2685
               Y2              =   2685
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   27
               X1              =   6960
               X2              =   8760
               Y1              =   3400
               Y2              =   3400
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   28
               X1              =   1080
               X2              =   2205
               Y1              =   3760
               Y2              =   3760
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   29
               X1              =   2790
               X2              =   3975
               Y1              =   3760
               Y2              =   3760
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   34
               X1              =   1080
               X2              =   2700
               Y1              =   4840
               Y2              =   4840
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   41
               X1              =   1080
               X2              =   2700
               Y1              =   5920
               Y2              =   5920
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   35
               X1              =   3195
               X2              =   4680
               Y1              =   4845
               Y2              =   4845
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   42
               X1              =   3195
               X2              =   4560
               Y1              =   5920
               Y2              =   5920
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   36
               X1              =   5080
               X2              =   6080
               Y1              =   4840
               Y2              =   4840
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   43
               X1              =   5160
               X2              =   6190
               Y1              =   5920
               Y2              =   5920
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   44
               X1              =   7200
               X2              =   8880
               Y1              =   5925
               Y2              =   5925
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   38
               X1              =   1080
               X2              =   2700
               Y1              =   5560
               Y2              =   5560
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   39
               X1              =   3480
               X2              =   5160
               Y1              =   5560
               Y2              =   5560
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   40
               X1              =   6120
               X2              =   8880
               Y1              =   5565
               Y2              =   5565
            End
         End
      End
   End
End
Attribute VB_Name = "frmArchiveInMedRec_HN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'说明：为了保持界面的可维护性，在新增控件时，注意保持每个信息条目包含的lblInfo，linInfo,txtinfo 的index相同，
'      若这组信息条目包含2个lblinfo则另外一个lblinfo的index为txtinfo.index+100

'上次刷新数据时的病人信息
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mblnMoved As Boolean
Private mblnCheck As Boolean
Private mbln共享 As Boolean

Private Enum Fra菜单
    FRA_基本信息 = 0
    FRA_西医诊断 = 1
    FRA_中医诊断 = 2
    FRA_过敏与手术 = 3
    FRA_住院情况 = 4
    FRA_放疗与化疗 = 5
    FRA_其他 = 6
End Enum

Private Enum 基本信息
    txt付款方式 = 0
    txt健康卡号 = 1
    txt住院次数 = 2
    chk再入院 = 0
    txt病案号 = 3
    txt姓名 = 4
    txt性别 = 5
    txt出生日期 = 6
    txt年龄 = 7
    txt国籍 = 8
    txt体重 = 9
    txt身高 = 10
    txt不足周岁年龄 = 11
    txt新生儿体重 = 12
    txt新生儿入院体重 = 13
    txt出生地点 = 14
    txt籍贯 = 15
    txt民族 = 16
    txt身份证号 = 17
    txt职业 = 18
    txt婚姻 = 19
    txt家庭地址 = 20
    txt家庭电话 = 21
    txt家庭邮编 = 22
    txt户口地址 = 23
    txt户口邮编 = 24
    txt工作单位 = 25
    txt单位电话 = 26
    txt单位邮编 = 27
    txt联系人姓名 = 28
    txt联系人关系 = 29
    txt联系人电话 = 30
    txt联系人地址 = 31
    txt区域 = 32
    txt入院途径 = 33
    txt入院时间 = 34
    txt入院科室 = 35
    txt入院病室 = 36
    chk入院前经外院治疗 = 1
    txt入科时间 = 37
    txt转科1 = 38
    txt转科2 = 39
    txt转科3 = 40
    txt出院时间 = 41
    txt出院科室 = 42
    txt出院病室 = 43
    txt住院天数 = 44
    txt其他证件 = 122
End Enum

Private Enum 西医诊断
    txt入院情况 = 45
    chk是否确诊 = 2
    txt确诊日期 = 46
    txt病理号 = 47
    txt分化程度 = 48
    txt最高诊断依据 = 49
    txt放射与病理 = 50
    txt门诊与出院 = 51
    txt入院与出院 = 52
    txt门诊与入院 = 53
    txt临床与病理 = 54
    txt临床与尸检 = 55
    txt术前与术后 = 56
    txt死亡时间 = 57
    txt死亡原因 = 58
    chk医院感染作病原学检查 = 3
    chk死亡患者尸检 = 4
    chk新发肿瘤 = 5
    txt医院感染病原学诊断 = 59
    txt抢救次数 = 60
    txt成功次数 = 61
    txt抢救原因 = 62
End Enum

Private Enum 中医诊断
    txt中医门诊与出院 = 63
    txt中医入院与出院 = 64
    chk危重 = 6
    chk急症 = 7
    chk疑难 = 8
    txt辨证 = 65
    txt治法 = 66
    txt方药 = 67
    txt治疗类别 = 68
    txt抢救方法 = 69
    txt自制中药 = 70
    txt中医设备 = 71
    txt中医技术 = 72
    txt辨证施护 = 73
End Enum

Private Enum 住院情况
    txt病例分型 = 74
    txtHBsAg = 75
    txt输血前9项检查 = 76
    txt血型 = 77
    txtHCVAb = 78
    txt发病时间 = 79
    txtRh = 80
    txtHIVAb = 81
    txt生育状况 = 82
    txt输液反应 = 83
    txt输血反应 = 84
    chk示教病案 = 13
    chk科研病案 = 14
    chk疑难病例 = 9
    txt输红细胞 = 85
    txt输血小板 = 86
    txt输血浆 = 87
    txt输全血 = 88
    txt输其他 = 89
    txt自体回收 = 90
    txt医学警示 = 91
    txt其他医学警示 = 92
    txt出院方式 = 93
    txt转入机构 = 94
    txt入院前天 = 95
    txt入院前小时 = 96
    txt入院前分钟 = 97
    txt入院后天 = 98
    txt入院后小时 = 99
    txt入院后分钟 = 100
    txt再入院天数 = 101
    txt31天目的 = 102
    txt呼吸机小时 = 103
    chk随诊 = 15
    txt随诊期限 = 104
    txt门诊医师 = 105
    txt科主任 = 106
    txt主任医师 = 107
    txt主治医师 = 108
    txt住院医师 = 109
    txt进修医师 = 110
    txt研究生医师 = 111
    txt实习医师 = 112
    txt质控医师 = 113
     txt责任护士 = 114
    txt质控护士 = 115
    txt质控日期 = 116
    txt病案质量 = 117
End Enum

Private Enum 诊断情况
    col诊断类型 = 0
    col诊断描述 = 1
    col中医证候 = 2
    col备注 = 3
    col入院病情 = 4
    col出院情况 = 5
    colzy类型 = 6
    col是否未治 = 6
    col是否疑诊 = 7
    col类型 = 8
End Enum

Private Enum 手术情况
    col手术日期 = 0
    COL手术情况 = 1
    col手术名称 = 2
    col再次手术 = 3
    col主刀医师 = 4
    col助产护士 = 5
    col助手1 = 6
    col助手2 = 7
    col麻醉方式 = 8
    colASA分级 = 9
    colNNIS分级 = 10
    col手术分级 = 11
    col麻醉医师 = 12
    col切口愈合 = 13
End Enum

Private Enum 过敏情况
    col过敏时间 = 0
    col过敏药物 = 1
    col过敏反应 = 2
End Enum

Private Enum 化疗记录
    col化疗编码 = 0
    COL化疗开始日期 = 1
    col化疗结束日期 = 2
    col化疗疗程数 = 3
    col化疗方案 = 4
    col化疗总量 = 5
    col化疗效果 = 6
End Enum

Private Enum 放疗记录
    col放疗编码 = 0
    COL放疗开始日期 = 1
    col放疗结束日期 = 2
    col设野部位 = 3
    col放射剂量 = 4
    col放射累计剂量 = 5
    col放疗效果 = 6
End Enum

Private Enum 其他
     chkCT = 16
     chkMRI = 17
     chk彩色多普勒 = 18
     txt压疮发生期间 = 118
     txt压疮分期 = 119
     txt跌倒或坠床伤害 = 120
     txt跌倒或坠床原因 = 121
     txt重症监护 = 123
     txt重症监护天数 = 124
     txt重症监护小时 = 125
     txt肿瘤分期 = 126
     txt肿瘤分期T = 127
     txt肿瘤分期M = 128
     txt肿瘤分期N = 129
     chk使用抗生素 = 21
     chk单病种 = 20
     chk细菌培养标本送检 = 19
     txtApgar = 130
     txt临床路径管理 = 131
     txt传染病 = 132
     txtDrGs管理 = 133
End Enum


Public Function zlRefresh(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal blnMoved As Boolean) As Boolean
'功能：刷新或清除医嘱清单
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, lng科室ID As Long
    Dim bln中医 As Boolean
    
    mlng病人ID = lng病人ID: mlng主页ID = lng主页ID: mblnMoved = blnMoved
    
    On Error GoTo errH
    
    StrSQL = "Select 出院科室ID From 病案主页 Where 病人id=[1] And 主页id=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Not rsTmp.EOF Then lng科室ID = Nvl(rsTmp!出院科室ID, 0)
    bln中医 = Have部门性质(lng科室ID, "中医科")
    fraInfo(FRA_中医诊断).Visible = bln中医
    fraInfo(FRA_中医诊断).Enabled = bln中医
    mbln共享 = CheckShare(300) '病案系统
    fraInfo(FRA_放疗与化疗).Visible = mbln共享
    fraInfo(FRA_放疗与化疗).Enabled = mbln共享
    
    Call SetPageHeight
    Call SetScrollbar
    
    Call ClearPageData
    If mlng病人ID <> 0 Then Call LoadPageData
    
    Call Form_Resize
    zlRefresh = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub chkInfo_Click(Index As Integer)
    If Not mblnCheck Then
        mblnCheck = True
        chkInfo(Index).Value = IIf(chkInfo(Index).Value = 1, 0, 1)
        mblnCheck = False
    End If
End Sub

Private Sub Form_Activate()
    Call Form_Resize
End Sub

Private Sub Form_Load()
    '滚动条尺寸
    vsc.Width = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    hsc.Height = GetSystemMetrics(SM_CXHSCROLL) * Screen.TwipsPerPixelY
    fraVH.Width = vsc.Width: fraVH.Height = hsc.Height
    fraBack.Left = 0: fraBack.Top = 0
    picBack.BackColor = fraBack.BackColor
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

Private Sub txtInfo_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtInfo(Index))
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
    Call vsDiagXY.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsDiagZY.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsfMain__AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsfMain.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsOPS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsOPS.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsRadiotherapy_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsRadiotherapy.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsChemotherapy_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsChemotherapy.ShowCell(NewRow, NewCol)
End Sub

Private Sub ClearPageData()
'功能：清除首页中的内容
    Dim objTmp As Object
    Dim i As Long, j As Long
    
    mblnCheck = True
    
    For Each objTmp In Me.Controls
        If TypeName(objTmp) = "TextBox" Then
            objTmp.Text = ""
        ElseIf TypeName(objTmp) = "CheckBox" Then
            objTmp.Value = 0
        End If
    Next
    
    With vsDiagXY
        .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
        For i = .Rows - 1 To .FixedRows Step -1
            If .TextMatrix(i, col诊断类型) = "" Then
                .RemoveItem i
            End If
        Next
    End With
    With vsDiagZY
        .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
        For i = .Rows - 1 To .FixedRows Step -1
            If .TextMatrix(i, col诊断类型) = "" Then
                .RemoveItem i
            End If
        Next
    End With
    With vsOPS
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    End With
    With vsAller
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    End With
    
    With vsChemotherapy
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    End With
    
    With vsfMain
        .Rows = .FixedRows
        .Rows = .FixedRows + 10
        .Cols = .FixedCols
        .Cols = .FixedCols + 10
    End With
    
    With vsRadiotherapy
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    End With
    
    With vsTSJC
        .Cols = .FixedCols
        .Cols = .FixedCols + 1
    End With
    
    lstAdvEvent.Clear
    
    lstInfection.Clear
        
    mblnCheck = False
End Sub

Private Function GetRow(ByVal lng诊断类型 As Long) As Long
'功能：返回指定诊断类型的第一诊断行
    If InStr(",11,12,13,", "," & lng诊断类型 & ",") > 0 Then
        GetRow = vsDiagZY.FindRow(CStr(lng诊断类型), , colzy类型)
    Else
        GetRow = vsDiagXY.FindRow(CStr(lng诊断类型), , col类型)
    End If
End Function

Private Function LoadPageData() As Boolean
'功能：读取病人的首页信息
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Long, j As Long
    Dim lngRow As Long, varTmp As Variant
    Dim strTmp As String
    Dim bln首页诊断 As Boolean, bln分化程度 As Boolean
    Dim bln压疮 As Boolean, bln跌倒坠床 As Boolean
    
    On Error GoTo errH

    Screen.MousePointer = 11
    mblnCheck = True
    
    '初始化病案附加项目
    Call FillVsf
    
    '病人信息部份
    '---------------------------------------------------------------
    StrSQL = "Select 住院号,姓名,性别,出生日期,出生地点,身份证号,其他证件,区域,民族,健康号,籍贯 From 病人信息 Where 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID)

    txtInfo(txt健康卡号).Text = Nvl(rsTmp!健康号)
    txtInfo(txt住院次数).Text = mlng主页ID
    txtInfo(txt姓名).Text = Nvl(rsTmp!姓名)
    txtInfo(txt性别).Text = Nvl(rsTmp!性别)
    txtInfo(txt再入院天数).Text = "无"
    If Format(rsTmp!出生日期, "HH:mm") <> "00:00" Then
        txtInfo(txt出生日期).Text = Format(rsTmp!出生日期, "yyyy-MM-dd HH:mm")
    Else
        txtInfo(txt出生日期).Text = Format(rsTmp!出生日期, "yyyy-MM-dd")
    End If

    txtInfo(txt出生地点).Text = Nvl(rsTmp!出生地点)
    txtInfo(txt籍贯).Text = Nvl(rsTmp!籍贯)
    txtInfo(txt身份证号).Text = Nvl(rsTmp!身份证号)
    txtInfo(txt民族).Text = Nvl(rsTmp!民族)
    txtInfo(txt区域).Text = Nvl(rsTmp!区域)
    txtInfo(txt其他证件).Text = Nvl(rsTmp!其他证件)
    '病案号读取
    StrSQL = "select 病案号 from 住院病案记录 where 病人ID=[1] and 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If rsTmp.RecordCount <> 0 Then
        txtInfo(txt病案号).Text = Nvl(rsTmp!病案号)
    End If
    '病案主页部份
    '---------------------------------------------------------------
    StrSQL = "Select A.*,B.名称 as 入院科室,C.名称 as 出院科室" & _
        " From 病案主页 A,部门表 B,部门表 C" & _
        " Where A.入院科室ID=B.ID And A.出院科室ID=C.ID" & _
        " And A.病人ID=[1] And A.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)

    txtInfo(txt付款方式).Text = Nvl(rsTmp!医疗付款方式)
    '留观病人无住院号
    If Nvl(rsTmp!病人性质, 0) <> 0 Then
        lblInfo(txt健康卡号).Visible = False
        txtInfo(txt健康卡号).Visible = False
    End If
    chkInfo(chk再入院).Value = Nvl(rsTmp!再入院, 0)
    txtInfo(txt病案号).Text = Nvl(rsTmp!病案号)
    
    txtInfo(txt年龄).Text = Nvl(rsTmp!年龄)
    txtInfo(txt国籍).Text = Nvl(rsTmp!国籍)
    '身高体重
    txtInfo(txt身高).Text = IIf(rsTmp!身高 & "" = "0", "", rsTmp!身高 & "")
    txtInfo(txt体重).Text = IIf(rsTmp!体重 & "" = "0", "", rsTmp!体重 & "")
    txtInfo(txt职业).Text = Nvl(rsTmp!职业)
    txtInfo(txt婚姻).Text = Nvl(rsTmp!婚姻状况)
    txtInfo(txt家庭地址).Text = Nvl(rsTmp!家庭地址)
    txtInfo(txt家庭电话).Text = Nvl(rsTmp!家庭电话)
    txtInfo(txt家庭邮编).Text = Nvl(rsTmp!家庭地址邮编)
    txtInfo(txt户口地址).Text = Nvl(rsTmp!户口地址)
    txtInfo(txt户口邮编).Text = Nvl(rsTmp!户口地址邮编)
    
    txtInfo(txt工作单位).Text = Nvl(rsTmp!单位地址)
    txtInfo(txt单位电话).Text = Nvl(rsTmp!单位电话)
    txtInfo(txt单位邮编).Text = Nvl(rsTmp!单位邮编)
    txtInfo(txt联系人姓名).Text = Nvl(rsTmp!联系人姓名)
    txtInfo(txt联系人关系).Text = Nvl(rsTmp!联系人关系)
    txtInfo(txt联系人电话).Text = Nvl(rsTmp!联系人电话)
    txtInfo(txt联系人地址).Text = Nvl(rsTmp!联系人地址)
    If Not IsNull(rsTmp!区域) Then
        txtInfo(txt区域).Text = Nvl(rsTmp!区域)
    End If

    txtInfo(txt入院途径).Text = Nvl(rsTmp!入院方式)
    txtInfo(txt入院时间).Text = Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm")
    txtInfo(txt入院科室).Text = rsTmp!入院科室
    
    txtInfo(txt出院时间).Text = Format(Nvl(rsTmp!出院日期), "yyyy-MM-dd HH:mm")
    txtInfo(txt出院科室).Text = rsTmp!出院科室
    If Not IsNull(rsTmp!出院日期) Then
        txtInfo(txt住院天数).Text = DateDiff("d", rsTmp!入院日期, rsTmp!出院日期)
    Else
        txtInfo(txt住院天数).Text = DateDiff("d", rsTmp!入院日期, zlDatabase.Currentdate)
    End If
    If Val(txtInfo(txt住院天数).Text) = 0 Then txtInfo(txt住院天数).Text = "1"
    
     txtInfo(txt入院情况).Text = Nvl(rsTmp!入院病况)
    chkInfo(chk是否确诊).Value = Nvl(rsTmp!是否确诊, 0)
    If chkInfo(chk是否确诊).Value = 1 Then
        txtInfo(txt确诊日期).Text = Format(Nvl(rsTmp!确诊日期), "yyyy-MM-dd HH:mm")
    End If
    chkInfo(chk死亡患者尸检).Value = Nvl(rsTmp!尸检标志, 0)
    chkInfo(chk新发肿瘤).Value = Nvl(rsTmp!新发肿瘤, 0)
    txtInfo(txt抢救次数).Text = Nvl(rsTmp!抢救次数)
    If Val(txtInfo(txt抢救次数).Text) <> 0 Then
        txtInfo(txt成功次数).Text = Nvl(rsTmp!成功次数)
    End If
    
    txtInfo(txt治疗类别).Text = Nvl(rsTmp!中医治疗类别)
    
    txtInfo(txt血型).Text = Nvl(rsTmp!血型)
    chkInfo(chk随诊).Value = IIf(Nvl(rsTmp!随诊标志, 0) = 0, 0, 1)
    If chkInfo(chk随诊).Value = 1 Then
        txtInfo(txt随诊期限).Text = IIf(Nvl(rsTmp!随诊标志, 0) = 9, "", Nvl(rsTmp!随诊期限, 0)) & _
            Decode(Nvl(rsTmp!随诊标志, 0), 1, "月", 2, "年", 3, "周", 4, "天", 9, "终身")
    End If
    txtInfo(txt门诊医师).Text = Nvl(rsTmp!门诊医师)
    txtInfo(txt住院医师).Text = Nvl(rsTmp!住院医师)
    txtInfo(txt责任护士).Text = Nvl(rsTmp!责任护士)
    '入科时间
    If Nvl(rsTmp!状态, 0) = 1 Then
        txtInfo(txt入科时间).Text = "尚未入科"
    Else
        StrSQL = "Select 开始时间 From 病人变动记录" & _
            " Where 病人ID=[1] And 主页ID=[2] And 开始原因 IN(2,1) And 开始时间 is Not Null Order by 开始原因 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
        If Not rsTmp.EOF Then
            txtInfo(txt入科时间).Text = Format(rsTmp!开始时间, "yyyy-MM-dd HH:mm")
        End If
    End If
    
    '病案从表部份
    '---------------------------------------------------------------
    StrSQL = "Select a.病人ID,a.主页ID,a.信息名,a.信息值,b.编码 From 病案主页从表 a " & _
            ",病案项目 b" & " where a.信息名=b.名称(+) And a.病人ID=[1] And a.主页ID=[2] Order by a.信息名"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    For i = 1 To rsTmp.RecordCount
        Select Case UCase(Nvl(rsTmp!信息名))
            Case "不足周岁年龄"
                txtInfo(txt不足周岁年龄).Text = Nvl(rsTmp!信息值)
            Case "新生儿出生体重"
                txtInfo(txt新生儿体重).Text = Nvl(rsTmp!信息值) & IIf(Nvl(rsTmp!信息值) = "", "", " 克")
            Case "新生儿入院体重"
                txtInfo(txt新生儿入院体重).Text = Nvl(rsTmp!信息值) & IIf(Nvl(rsTmp!信息值) = "", "", " 克")
            Case "籍贯"
                txtInfo(txt籍贯).Text = Nvl(rsTmp!信息值)
            Case "入院病室"
                txtInfo(txt入院病室).Text = Nvl(rsTmp!信息值)
            Case "入院前经外院治疗"
                chkInfo(chk入院前经外院治疗).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "转科记录"
                varTmp = Split(Nvl(rsTmp!信息值), ",")
                If UBound(varTmp) >= 0 Then txtInfo(txt转科1).Text = varTmp(0)
                If UBound(varTmp) >= 1 Then txtInfo(txt转科2).Text = varTmp(1)
                If UBound(varTmp) >= 2 Then txtInfo(txt转科3).Text = varTmp(2)
            Case "出院病室"
                txtInfo(txt出院病室).Text = Nvl(rsTmp!信息值)
            Case "病理号"
                txtInfo(txt病理号).Text = Nvl(rsTmp!信息值)
            Case "分化程度"
                If Nvl(rsTmp!信息值) <> "" Then
                    txtInfo(txt分化程度).Text = Nvl(rsTmp!信息值)
                End If
            Case "最高诊断依据"
                If Nvl(rsTmp!信息值) <> "" Then
                    txtInfo(txt最高诊断依据).Text = Nvl(rsTmp!信息值)
                End If
            Case "病原学检查"
                chkInfo(chk医院感染作病原学检查).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "死亡时间"
                If Not (IsNull(rsTmp!信息值) Or Not IsDate(rsTmp!信息值)) Then
                    txtInfo(txt死亡时间).Text = rsTmp!信息值
                End If
            Case "死亡根本原因"
                txtInfo(txt死亡原因).Text = Nvl(rsTmp!信息值)
            Case "抢救病因"
                txtInfo(txt抢救原因).Text = Nvl(rsTmp!信息值)
            Case "中医危重"
                chkInfo(chk危重).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "中医急症"
                chkInfo(chk急症).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "中医疑难"
                chkInfo(chk疑难).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "中医抢救方法"
                txtInfo(txt抢救方法).Text = Nvl(rsTmp!信息值)
            Case "自制中药制剂"
                txtInfo(txt自制中药).Text = Nvl(rsTmp!信息值)
            Case "中医设备"
                txtInfo(txt中医设备).Text = Nvl(rsTmp!信息值)
            Case "中医技术"
                txtInfo(txt中医技术).Text = Nvl(rsTmp!信息值)
            Case "辨证施护"
                txtInfo(txt辨证施护).Text = Nvl(rsTmp!信息值)
            Case "病例分型"
                txtInfo(txt病例分型).Text = GetNameByCode("病例分型", Nvl(rsTmp!信息值))
            Case UCase("HBsAg")
                txtInfo(txtHBsAg).Text = Nvl(rsTmp!信息值)
            Case UCase("HCV-Ab")
                txtInfo(txtHCVAb).Text = Nvl(rsTmp!信息值)
            Case UCase("HIV-Ab")
                txtInfo(txtHIVAb).Text = Nvl(rsTmp!信息值)
            Case UCase("Rh")
            Case UCase("Rh")
                txtInfo(txtRh).Text = Nvl(rsTmp!信息值)
            Case "输血检查"
                txtInfo(txt输血前9项检查).Text = Nvl(rsTmp!信息值)
            Case "发病时间"
                If Nvl(rsTmp!信息值) <> "" Then
                    If Format(rsTmp!信息值, "HH:mm") <> "00:00" Then
                        txtInfo(txt发病时间).Text = Format(rsTmp!信息值, "yyyy-MM-dd HH:mm")
                    Else
                        txtInfo(txt发病时间).Text = Format(rsTmp!信息值, "yyyy-MM-dd")
                    End If
                End If
            Case "生育状况"
                txtInfo(txt生育状况).Text = Decode(Val(Nvl(rsTmp!信息值, 0)), 0, "未生育", 1, "生育1胎", 2, "生育2胎及以上", 4, "4-不详")
            Case "输液反应"
                txtInfo(txt输液反应).Text = Nvl(rsTmp!信息值)
            Case "输血反应"
                txtInfo(txt输血反应).Text = Decode(Val(Nvl(rsTmp!信息值, 0)), 0, "无", 1, "有", 2, "未输")
            Case "示教病案"
                chkInfo(chk示教病案).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "科研病案"
                chkInfo(chk科研病案).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "疑难病历"
                chkInfo(chk疑难病例).Value = Val(Nvl(rsTmp!信息值))
            Case "输红细胞"
                txtInfo(txt输红细胞).Text = Nvl(rsTmp!信息值)
            Case "输血小板"
                txtInfo(txt输血小板).Text = Nvl(rsTmp!信息值)
            Case "输血浆"
                txtInfo(txt输血浆).Text = Nvl(rsTmp!信息值)
            Case "输全血"
                txtInfo(txt输全血).Text = Nvl(rsTmp!信息值)
            Case "输其他"
                txtInfo(txt输其他).Text = Nvl(rsTmp!信息值)
            Case "自体回收"
                txtInfo(txt自体回收).Text = Nvl(rsTmp!信息值)
            Case "医学警示"
                txtInfo(txt医学警示).Text = Nvl(rsTmp!信息值)
            Case "其他医学警示"
                txtInfo(txt其他医学警示).Text = Nvl(rsTmp!信息值)
            Case "出院方式"
                txtInfo(txt出院方式).Text = Nvl(rsTmp!信息值)
            Case "出院转入"
                txtInfo(txt转入机构).Text = Nvl(rsTmp!信息值)
            Case "昏迷时间"
                '保存格式:入院前(天，小时,分钟)|入院后(天，小时,分钟)
                txtInfo(txt入院前天).Text = Split(Split(Nvl(rsTmp!信息值), "|")(0) & ",", ",")(0)
                txtInfo(txt入院前小时).Text = Split(Split(Nvl(rsTmp!信息值), "|")(0) & ",", ",")(1)
                txtInfo(txt入院前分钟).Text = Split(Split(Nvl(rsTmp!信息值), "|")(0) & ",", ",")(2)
                txtInfo(txt入院后天).Text = Split(Split(Nvl(rsTmp!信息值), "|")(1) & ",", ",")(0)
                txtInfo(txt入院后小时).Text = Split(Split(Nvl(rsTmp!信息值) & "|", "|")(1) & ",", ",")(1)
                txtInfo(txt入院后分钟).Text = Split(Split(Nvl(rsTmp!信息值) & "|", "|")(1) & ",", ",")(2)
            Case "再入院计划天数"
                lblInfo(txt再入院天数).Caption = "出院" & IIf(Nvl(rsTmp!信息值, "0") = "0", "31", "7") & "天内再入院计划"
            Case "31天内再住院"
                If Nvl(rsTmp!信息值) <> "" Then
                    txtInfo(txt31天目的).Text = Nvl(rsTmp!信息值)
                    txtInfo(txt再入院天数).Text = "有"
                Else
                    txtInfo(txt再入院天数).Text = "无"
                End If
            Case "呼吸机使用时间"
                txtInfo(txt呼吸机小时).Text = Nvl(rsTmp!信息值)
            Case "科主任"
                txtInfo(txt科主任).Text = Nvl(rsTmp!信息值)
            Case "主任医师"
                txtInfo(txt主任医师).Text = Nvl(rsTmp!信息值)
            Case "主治医师"
                txtInfo(txt主治医师).Text = Nvl(rsTmp!信息值)
            Case "进修医师"
                txtInfo(txt进修医师).Text = Nvl(rsTmp!信息值)
            Case "研究生实习医师"
                txtInfo(txt研究生医师).Text = Nvl(rsTmp!信息值)
            Case "实习医师"
                txtInfo(txt实习医师).Text = Nvl(rsTmp!信息值)
            Case "质控医师"
                txtInfo(txt质控医师).Text = Nvl(rsTmp!信息值)
            Case "质控护士"
                txtInfo(txt质控护士).Text = Nvl(rsTmp!信息值)
            Case "病案质量"
                txtInfo(txt病案质量).Text = Nvl(rsTmp!信息值)
            Case "主页质量日期"
                txtInfo(txt质控日期).Text = Nvl(rsTmp!信息值)
            Case "CT"
                chkInfo(chkCT).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "MRI"
                chkInfo(chkMRI).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "彩色多普勒"
                chkInfo(chk彩色多普勒).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "特殊检查4"
                vsTSJC.TextMatrix(vsTSJC.FixedRows + 0, 1) = Nvl(rsTmp!信息值)
                vsTSJC.Cell(flexcpData, vsTSJC.FixedRows + 0, 1) = vsTSJC.TextMatrix(vsTSJC.FixedRows + 0, 1)
            Case "特殊检查5"
                vsTSJC.TextMatrix(vsTSJC.FixedRows + 1, 1) = Nvl(rsTmp!信息值)
                vsTSJC.Cell(flexcpData, vsTSJC.FixedRows + 1, 1) = vsTSJC.TextMatrix(vsTSJC.FixedRows + 1, 1)
            Case "特殊检查6"
                vsTSJC.TextMatrix(vsTSJC.FixedRows + 2, 1) = Nvl(rsTmp!信息值)
                vsTSJC.Cell(flexcpData, vsTSJC.FixedRows + 2, 1) = vsTSJC.TextMatrix(vsTSJC.FixedRows + 2, 1)
            Case "压疮发生期间"
                txtInfo(txt压疮发生期间).Text = Nvl(rsTmp!信息值, " ")
            Case "压疮分期"
                txtInfo(txt压疮分期).Text = Nvl(rsTmp!信息值, " ")
            Case "跌倒或坠床伤害"
                txtInfo(txt跌倒或坠床伤害).Text = Nvl(rsTmp!信息值, " ")
            Case "跌倒或坠床原因"
                txtInfo(txt跌倒或坠床原因).Text = Nvl(rsTmp!信息值, " ")
            Case "重症监护天数"
                txtInfo(txt重症监护天数).Text = Nvl(rsTmp!信息值, "")
            Case "重症监护小时"
                txtInfo(txt重症监护小时).Text = Nvl(rsTmp!信息值, "")
            Case "单病种"
                chkInfo(chk单病种).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "临床路径"
                txtInfo(txt临床路径管理).Text = Decode(Val(rsTmp!信息值 & ""), 1, "未进入", 2, "变异退出", 3, "完成")
            Case "DRGS"
                txtInfo(txtDrGs管理).Text = Decode(Val(rsTmp!信息值 & ""), 1, "无", 2, "按病种", 3, "按费用", 4, "两者都有")
            Case "抗生素"
                chkInfo(chk使用抗生素).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "标本送检"
                chkInfo(chk细菌培养标本送检).Value = Val(Nvl(rsTmp!信息值, 0))
            Case "传染病"
                txtInfo(txt传染病).Text = Decode(Val(rsTmp!信息值 & ""), 1, "甲类", 2, "乙类", 3, "丙类")
            Case "肿瘤分期"
                txtInfo(txt肿瘤分期).Text = Decode(Val(rsTmp!信息值 & ""), 1, "0期", 2, "I期", 3, "Ⅱ期", 4, "Ⅲ期", 5, "Ⅳ期", 6, "不详")
            Case "肿瘤T"
                txtInfo(txt肿瘤分期T).Text = Nvl(rsTmp!信息值, "")
            Case "肿瘤M"
                txtInfo(txt肿瘤分期M).Text = Nvl(rsTmp!信息值, "")
            Case "肿瘤N"
                txtInfo(txt肿瘤分期N).Text = Nvl(rsTmp!信息值, "")
            Case "APGAR"
                txtInfo(txtApgar).Text = Nvl(rsTmp!信息值, "")
            Case Else
                If Not (Left(Nvl(rsTmp!信息名), 3) = "抗生素" And Not IsNull(rsTmp!信息值)) Then
                    '附加项目
                    If Not IsNull(rsTmp("编码")) Then
                        With vsfMain
                            For j = 0 To vsfMain.Cols - 1 Step 3
                                lngRow = vsfMain.FindRow(rsTmp("信息名"), , j)
                                If lngRow >= 0 Then
                                    If vsfMain.TextMatrix(lngRow, j) = rsTmp("信息名") Then
                                        If vsfMain.TextMatrix(lngRow, j + 2) = "是否" Then
                                            vsfMain.Cell(flexcpChecked, lngRow, j + 1) = IIf(rsTmp("信息值") = 0, 2, 1)
                                            Exit For
                                        Else
                                            vsfMain.TextMatrix(lngRow, j + 1) = rsTmp("信息值") & ""
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next j
                        End With
                    End If
                End If
        End Select
        rsTmp.MoveNext
    Next
    '重症监护设置
    If Val(txtInfo(txt重症监护天数).Text) <> 0 And Val(txtInfo(txt重症监护小时).Text) <> 0 Then
        txtInfo(txt重症监护) = "是"
        If Val(txtInfo(txt重症监护天数).Text) <> 0 Then
            txtInfo(txt重症监护天数).Enabled = True
            lblInfo(txt重症监护天数).Enabled = True
            lblInfo(txt重症监护天数 + 100).Enabled = True
        End If
        If Val(txtInfo(txt重症监护小时).Text) <> 0 Then
            txtInfo(txt重症监护小时).Enabled = True
            lblInfo(txt重症监护小时).Enabled = True
        End If
    End If
    
    '自动提取转科科室及入出病室(房间号)
    '---------------------------------------------------------------
    If txtInfo(txt转科1).Text = "" And txtInfo(txt转科2).Text = "" And txtInfo(txt转科3).Text = "" Then
        StrSQL = _
            " Select B.名称" & _
            " From 病人变动记录 A,部门表 B" & _
            " Where A.病人ID=[1] And A.主页ID=[2]" & _
            " And A.科室ID=B.ID And A.开始原因=3 And A.开始时间 is Not NULL" & _
            " Order by A.开始时间"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
        For i = 1 To rsTmp.RecordCount
            If i = 1 Then
                txtInfo(txt转科1).Text = rsTmp!名称
            ElseIf i = 2 Then
                txtInfo(txt转科2).Text = rsTmp!名称
            ElseIf i = 3 Then
                txtInfo(txt转科3).Text = rsTmp!名称
                Exit For
            End If
            rsTmp.MoveNext
        Next
    End If

    If txtInfo(txt入院病室).Text = "" Then
        StrSQL = "Select B.房间号" & _
            " From 病案主页 A,床位状况记录 B" & _
            " Where A.病人ID=[1] And A.主页ID=[2]" & _
            " And A.入院病区ID=B.病区ID And A.入院病床=B.床号"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
        If Not rsTmp.EOF Then txtInfo(txt入院病室).Text = Nvl(rsTmp!房间号)
    End If

    If txtInfo(txt出院病室).Text = "" Then
        StrSQL = "Select B.房间号" & _
            " From 病案主页 A,床位状况记录 B" & _
            " Where A.病人ID=[1] And A.主页ID=[2]" & _
            " And A.当前病区ID=B.病区ID And A.出院病床=B.床号"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
        If Not rsTmp.EOF Then txtInfo(txt出院病室).Text = Nvl(rsTmp!房间号)
    End If
    
    '西医诊断
    '---------------------------------------------------------------
'    str治疗结果 = Get治疗结果
'    vsDiagXY.ColData(col出院情况) = str治疗结果

    '判断首页是否填过诊断
    StrSQL = "Select 1 From 病人诊断记录 Where 病人ID=[1] And 主页ID=[2] And 记录来源=3  And RowNum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    bln首页诊断 = rsTmp.RecordCount > 0
    If bln首页诊断 Then
        strTmp = " and a.记录来源=3 "
    Else
        strTmp = " And a.记录来源 IN(1,2,3,4) "
    End If
    '缺省表格初始化
    With vsDiagXY
        '1-西医门诊诊断;2-西医入院诊断;3-出院诊断(其他诊断);5-院内感染;6-病理诊断;7-损伤中毒码;10-并发症
        .TextMatrix(1, col类型) = 1
        .TextMatrix(2, col类型) = 2
        .TextMatrix(3, col类型) = 3
        .TextMatrix(4, col类型) = 3
        .TextMatrix(5, col类型) = 5
        .TextMatrix(6, col类型) = 10
        .TextMatrix(7, col类型) = 6
        .TextMatrix(8, col类型) = 7
    End With

    '读取各种来源的诊断
    StrSQL = "Select a.备注,a.ID,a.病人ID,a.主页ID,a.医嘱ID,a.记录来源,a.诊断次序,a.编码序号,a.病历ID,a.诊断类型,a.疾病ID,a.入院病情," & _
        " a.诊断ID,a.证候ID,a.诊断描述,a.出院情况,a.是否未治,a.是否疑诊,a.记录日期,a.记录人,a.取消时间,a.取消人,a.病例ID, b.编码 As 疾病编码, c.编码 As 诊断编码 " & _
        " From 病人诊断记录 A, 疾病编码目录 B, 疾病诊断目录 C" & _
        " Where a.疾病id = b.Id(+) And a.诊断id = c.Id(+)  And a.诊断类型 IN(1,2,3,5,6,7,10,21)" & _
        strTmp & _
        " And a.取消时间 is Null And a.病人ID=[1] And a.主页ID=[2]" & _
        " Order by a.诊断类型,a.记录来源 Desc,a.诊断次序,a.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Not rsTmp.EOF Then
        With vsDiagXY
            StrSQL = "1,2,3,5,6,7,10,21"
            For i = 0 To UBound(Split(StrSQL, ","))
                rsTmp.Filter = "记录来源=3 And 诊断类型=" & Split(StrSQL, ",")(i)
                If Val(Split(StrSQL, ",")(i)) <> 21 Then
                    If rsTmp.EOF Then
                        rsTmp.Filter = "记录来源=2 And 诊断类型=" & Split(StrSQL, ",")(i)
                    End If
                    If rsTmp.EOF Then
                        rsTmp.Filter = "记录来源=1 And 诊断类型=" & Split(StrSQL, ",")(i)
                    End If
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "记录来源=4 And 诊断类型=" & Split(StrSQL, ",")(i)
                End If

                If Val(Split(StrSQL, ",")(i)) = 21 Then
                    '21-病原学诊断
                    If Not rsTmp.EOF Then
                        txtInfo(txt医院感染病原学诊断).Text = Nvl(rsTmp!诊断描述)
                    End If
                Else
                    Do While Not rsTmp.EOF
                        '确定当前显示行
                        lngRow = .FindRow(CStr(Split(StrSQL, ",")(i)), , col类型)
                        For j = lngRow To .Rows - 1
                            If Val(.TextMatrix(j, col类型)) = Val(Split(StrSQL, ",")(i)) Then
                                lngRow = j
                                If .TextMatrix(j, col诊断描述) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, col诊断描述) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, col类型) = Split(StrSQL, ",")(i)
                        End If

                        '分化程度和最高诊断依据
                        If Val("" & rsTmp!诊断类型) = 3 And Val("" & rsTmp!诊断次序) = 1 Then
                            If Trim(Nvl(rsTmp!疾病编码)) = "" Then
                                bln分化程度 = False
                            Else
                                bln分化程度 = ((InStr("C", UCase(Left(Nvl(rsTmp!疾病编码), 1)))) > 0) Or ((InStr("D0", UCase(Left(Nvl(rsTmp!疾病编码), 2)))) > 0) Or ((InStr("D32.,D33.,", UCase(Left(Nvl(rsTmp!疾病编码), 4)))) > 0)
                            End If
                        End If

                        txtInfo(txt分化程度).Enabled = bln分化程度
                        lblInfo(txt分化程度).Enabled = bln分化程度
                        lblInfo(txt最高诊断依据).Enabled = bln分化程度
                        txtInfo(txt最高诊断依据).Enabled = bln分化程度
                        .TextMatrix(lngRow, col诊断描述) = Nvl(rsTmp!诊断描述)
                        .TextMatrix(lngRow, col备注) = Nvl(rsTmp!备注)
                        .TextMatrix(lngRow, col出院情况) = Nvl(rsTmp!出院情况)
                        .TextMatrix(lngRow, col入院病情) = Nvl(rsTmp!入院病情)
                        .TextMatrix(lngRow, col是否未治) = IIf(Nvl(rsTmp!是否未治, 0) = 1, "√", "")
                        .TextMatrix(lngRow, col是否疑诊) = IIf(Nvl(rsTmp!是否疑诊, 0) = 1, "？", "")
                        rsTmp.MoveNext
                    Loop
                End If
            Next
        End With
    End If

    vsDiagXY.Cell(flexcpForeColor, 1, col是否疑诊, vsDiagXY.Rows - 1, col是否疑诊) = vbRed
    vsDiagXY.Cell(flexcpBackColor, GetRow(3), vsDiagXY.FixedRows, GetRow(3), vsDiagXY.Cols - 1) = &HC0FFC0
    vsDiagXY.Row = 1: vsDiagXY.Col = col诊断描述
    If vsDiagXY.TextMatrix(GetRow(6), col诊断描述) <> "" Then
        txtInfo(txt病理号).Enabled = True
        txtInfo(txt病理号).BackColor = vbWindowBackground
    End If

    '诊断符合情况
    '---------------------------------------------------------------
    StrSQL = "Select 符合类型,符合情况 From 诊断符合情况 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    Do While Not rsTmp.EOF
        Select Case rsTmp!符合类型
        Case 1 '门诊与出院
            txtInfo(txt门诊与出院).Text = Decode(Nvl(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 2 '入院与出院
            txtInfo(txt入院与出院).Text = Decode(Nvl(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 3 '放射与病理
            txtInfo(txt放射与病理).Text = Decode(Nvl(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 4 '临床与病理
            txtInfo(txt临床与病理).Text = Decode(Nvl(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 5 '临床与尸检
            txtInfo(txt临床与尸检).Text = Decode(Nvl(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 6 '术前与术后
            txtInfo(txt术前与术后).Text = Decode(Nvl(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 7 '门诊与入院
             txtInfo(txt门诊与入院).Text = Decode(Nvl(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 11 '中医门诊与出院
            txtInfo(txt中医门诊与出院).Text = Decode(Nvl(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 12 '中医入院与出院
            txtInfo(txt中医入院与出院).Text = Decode(Nvl(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 13 '中医辨证
            txtInfo(txt辨证).Text = Decode(Nvl(rsTmp!符合情况, 0), 1, "准确", 2, "基本准确", 3, "重大缺陷", 4, "错误", "")
        Case 14 '中医治法
            txtInfo(txt治法).Text = Decode(Nvl(rsTmp!符合情况, 0), 1, "准确", 2, "基本准确", 3, "重大缺陷", 4, "错误", "")
        Case 15 '中医方药
            txtInfo(txt方药).Text = Decode(Nvl(rsTmp!符合情况, 0), 1, "准确", 2, "基本准确", 3, "重大缺陷", 4, "错误", "")
        End Select
        rsTmp.MoveNext
    Loop

    '中医诊断
    '---------------------------------------------------------------
    '缺省表格初始化
    With vsDiagZY
        '11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断(主要诊断、其它诊断)
        .TextMatrix(1, colzy类型) = 11
        .TextMatrix(2, colzy类型) = 12
        .TextMatrix(3, colzy类型) = 13
        .TextMatrix(4, colzy类型) = 13
    End With

    If bln首页诊断 Then
        strTmp = " and a.记录来源=3 "
    Else
        strTmp = " And a.记录来源 IN(1,2,3,4) "
    End If

    '读取各种来源的诊断
    StrSQL = "Select a.备注, a.Id, a.病人id, a.主页id, a.医嘱id, a.记录来源, a.诊断次序, a.编码序号, a.病历id, a.诊断类型,a.入院病情," & _
        " a.疾病id, a.诊断id, a.证候id, a.诊断描述,a.出院情况, a.是否未治, a.是否疑诊, a.记录日期, a.记录人, a.取消时间," & _
        " a.取消人, a.病例id, b.编码 As 疾病编码, c.编码 As 诊断编码,d.编码 as 证候编码 From 病人诊断记录 A, 疾病编码目录 B, 疾病诊断目录 C,疾病编码目录 D" & _
        " Where a.疾病id = b.Id(+) And a.诊断id = c.Id(+) And a.证候ID=d.ID(+) And a.诊断类型 IN(11,12,13)" & _
        strTmp & _
        " And 取消时间 Is Null And 病人ID=[1] And 主页ID=[2]" & _
        " Order by a.诊断类型,a.记录来源 Desc,a.诊断次序,a.编码序号,a.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    
    If Not rsTmp.EOF Then
        With vsDiagZY
            StrSQL = "11,12,13"
            For i = 0 To UBound(Split(StrSQL, ","))
                rsTmp.Filter = "记录来源=3 And 诊断类型=" & Split(StrSQL, ",")(i)
                If rsTmp.EOF Then
                    rsTmp.Filter = "记录来源=2 And 诊断类型=" & Split(StrSQL, ",")(i)
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "记录来源=1 And 诊断类型=" & Split(StrSQL, ",")(i)
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "记录来源=4 And 诊断类型=" & Split(StrSQL, ",")(i)
                End If

                Do While Not rsTmp.EOF
                    '确定当前显示行
                    lngRow = .FindRow(CStr(Split(StrSQL, ",")(i)), , colzy类型)
                    For j = lngRow To .Rows - 1
                        If Val(.TextMatrix(j, colzy类型)) = Val(Split(StrSQL, ",")(i)) Then
                            lngRow = j
                            If .TextMatrix(j, col诊断描述) = "" Then Exit For
                        Else
                            Exit For
                        End If
                    Next
                    If .TextMatrix(lngRow, col诊断描述) <> "" Then
                        lngRow = lngRow + 1: .AddItem "", lngRow
                        .TextMatrix(lngRow, colzy类型) = Split(StrSQL, ",")(i)
                    End If
                    .TextMatrix(lngRow, col备注) = Nvl(rsTmp!备注)
                    .TextMatrix(lngRow, col诊断描述) = Nvl(rsTmp!诊断描述)
                    .TextMatrix(lngRow, col出院情况) = Nvl(rsTmp!出院情况)
                    .TextMatrix(lngRow, col入院病情) = Nvl(rsTmp!入院病情)
                    '取证候名称
                    If InStr(.TextMatrix(lngRow, col诊断描述), "(") > 0 And InStr(.TextMatrix(lngRow, col诊断描述), ")") > 0 Then
                        strTmp = Mid(.TextMatrix(lngRow, col诊断描述), InStrRev(.TextMatrix(lngRow, col诊断描述), "(") + 1)
                        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                        '先取证候
                        .TextMatrix(lngRow, col中医证候) = strTmp
                        '去掉诊断描述的证候
                        .TextMatrix(lngRow, col诊断描述) = Mid(.TextMatrix(lngRow, col诊断描述), 1, InStrRev(.TextMatrix(lngRow, col诊断描述), "(") - 1)
                    Else
                       .TextMatrix(lngRow, col中医证候) = ""
                    End If
                    
                    rsTmp.MoveNext
                Loop
            Next
        End With
    End If
    vsDiagZY.Cell(flexcpBackColor, GetRow(13), vsDiagZY.FixedRows, GetRow(13), vsDiagZY.Cols - 1) = &HC0FFC0
    vsDiagZY.Row = 1: vsDiagZY.Col = col诊断描述

    '过敏信息:本次住院的,过敏的
    '---------------------------------------------------------------
    StrSQL = "Select 记录来源,NVL(过敏时间,记录时间) as 过敏时间,药物ID,药物名,过敏反应 From 病人过敏记录 A" & _
        " Where 结果=1 And 病人ID=[1] And 主页ID=[2]" & _
        " And Not Exists(Select 药物ID From 病人过敏记录" & _
            " Where (Nvl(药物ID,0)=Nvl(A.药物ID,0) Or Nvl(药物名,'Null')=Nvl(A.药物名,'Null'))" & _
            " And Nvl(结果,0)=0 And 记录时间>A.记录时间 And 病人ID=[1] And 主页ID=[2])" & _
        " Order by NVL(过敏时间,记录时间),药物名"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Not rsTmp.EOF Then
        rsTmp.Filter = "记录来源=3" '首页本身填写的
        If rsTmp.EOF Then rsTmp.Filter = "记录来源<>3" '其它来源的作为缺省显示
        With vsAller
            .Rows = rsTmp.RecordCount + 1 '固定行+新行
            For i = 1 To rsTmp.RecordCount
                '其它来源的可能有重复
                lngRow = -1
                If Not IsNull(rsTmp!药物ID) Then
                    lngRow = .FindRow(CLng(rsTmp!药物ID))
                ElseIf Not IsNull(rsTmp!药物名) Then
                    lngRow = .FindRow(CStr(rsTmp!药物名), , col过敏药物)
                End If
                If lngRow = -1 Then
                    .RowData(i) = CLng(Nvl(rsTmp!药物ID, 0))
                    .TextMatrix(i, col过敏时间) = Format(rsTmp!过敏时间, "yyyy-MM-dd HH:mm")
                    .TextMatrix(i, col过敏药物) = Nvl(rsTmp!药物名)
                    .TextMatrix(i, col过敏反应) = Nvl(rsTmp!过敏反应)
                End If
                rsTmp.MoveNext
            Next
            If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        End With
    End If
    vsAller.Row = 1: vsAller.Col = col过敏药物

    '手术情况
    '---------------------------------------------------------------
    '首读取首页整理保存的内容
    StrSQL = "Select a.助产护士,a.手术情况,a.切口,a.愈合,a.手术日期,a.已行手术,a.主刀医师,a.第一助手,a.第二助手,a.麻醉类型,a.麻醉医师,a.ASA分级,a.再次手术,a.NNIS分级,decode(a.手术级别,1,'一级手术',2,'二级手术',3,'三级手术',4,'四级手术',' ') as 手术级别" & _
            " From 病人手麻记录  A,疾病编码目录 B,诊疗项目目录 C Where c.ID(+)=a.诊疗项目ID And A.手术操作ID=B.ID(+) and 病人ID=[1] And 主页ID=[2] And 记录来源=3 Order by A.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If rsTmp.EOF Then '没有时读取其它来源的诊断
        '病历：病历作废时填写取消
        StrSQL = "Select Max(记录日期) From 病人手麻记录" & _
            " Where 病人ID=" & mlng病人ID & " And 主页ID=" & mlng主页ID & _
            " And 记录来源=1 And 取消时间 is NULL"
         StrSQL = "Select a.助产护士,a.手术情况,a.切口,a.愈合,a.手术日期,a.已行手术,a.主刀医师,a.第一助手,a.第二助手,a.麻醉类型,a.麻醉医师,a.ASA分级,a.再次手术,a.NNIS分级,decode(a.手术级别,1,'一级手术',2,'二级手术',3,'三级手术',4,'四级手术',' ') as 手术级别" & _
            " From 病人手麻记录  A,疾病编码目录 B,诊疗项目目录 C Where c.ID(+)=a.诊疗项目ID And " & _
            " A.手术操作ID=B.ID(+) and 病人ID=[1] And 主页ID=[2]" & _
            " And 记录来源=1 And 取消时间 is NULL And 记录日期=(" & StrSQL & ")" & _
            " Order by A.ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
        If rsTmp.EOF Then '病案
            StrSQL = "Select a.助产护士,a.手术情况,a.切口,a.愈合,a.手术日期,a.已行手术,a.主刀医师,a.第一助手,a.第二助手,a.麻醉类型,a.麻醉医师,a.ASA分级,a.再次手术,a.NNIS分级,decode(a.手术级别,1,'一级手术',2,'二级手术',3,'三级手术',4,'四级手术',' ') as 手术级别" & _
                " From 病人手麻记录  A,疾病编码目录 B,诊疗项目目录 C Where c.ID(+)=a.诊疗项目ID And  A.手术操作ID=B.ID(+) and 病人ID=[1] And 主页ID=[2] And 记录来源=4 Order by A.ID"
            Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
        End If
    End If
    If Not rsTmp.EOF Then
        With vsOPS
            .Rows = .FixedRows + rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col手术日期) = Format(Nvl(rsTmp!手术日期), "yyyy-MM-dd")
                .TextMatrix(i, col手术名称) = Nvl(rsTmp!已行手术)
                .TextMatrix(i, col主刀医师) = Nvl(rsTmp!主刀医师)
                .TextMatrix(i, col助产护士) = Nvl(rsTmp!助产护士)
                .TextMatrix(i, col助手1) = Nvl(rsTmp!第一助手)
                .TextMatrix(i, col助手2) = Nvl(rsTmp!第二助手)
                .TextMatrix(i, col麻醉方式) = Nvl(rsTmp!麻醉类型)
                .TextMatrix(i, col麻醉医师) = Nvl(rsTmp!麻醉医师)
                If Not IsNull(rsTmp!切口) And Not IsNull(rsTmp!愈合) Then
                    .TextMatrix(i, col切口愈合) = rsTmp!切口 & "/" & rsTmp!愈合
                End If
                .TextMatrix(i, COL手术情况) = Nvl(rsTmp!手术情况)
                .TextMatrix(i, colASA分级) = Decode(Nvl(rsTmp!asa分级), "I级", "P1", "II级", "P2", "III级", "P3", "IV级", "P4", "V级", "P5", Nvl(rsTmp!asa分级))
                .TextMatrix(i, colNNIS分级) = Nvl(rsTmp!NNIS分级)
                .TextMatrix(i, col手术分级) = Nvl(rsTmp!手术级别)
                .TextMatrix(i, col再次手术) = IIf(Val(rsTmp!再次手术 & "") = 1, -1, 0)
                rsTmp.MoveNext
            Next
        End With
    End If

    
    If mbln共享 Then
        '放疗化疗
        Call Load化疗与放疗(mlng病人ID, mlng主页ID)
    End If
    
    '附加信息
    '---------------------------------------------------------------
    '不良事件
    lstAdvEvent.Clear
    
    
    bln压疮 = False: bln跌倒坠床 = False
    StrSQL = "Select 编码, 名称" & vbNewLine & _
            "From 不良事件 A," & vbNewLine & _
            "     (Select Decode(信息值, Null, Null, ',' || 信息值 || ',') 信息值" & vbNewLine & _
            "       From 病案主页从表" & vbNewLine & _
            "       Where 病人id = [1] And 主页id = [2] And 信息名 = '不良事件') B" & vbNewLine & _
            "Where Instr(b.信息值 , chr(44)|| a.编码 ||chr(44) ) > 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    For i = 1 To rsTmp.RecordCount
        lstAdvEvent.AddItem Nvl(rsTmp!名称)
        If Nvl(rsTmp!名称) = "压疮" Then
            bln压疮 = True
        ElseIf Nvl(rsTmp!名称) = "医院内跌倒/坠床" Then '压疮 跌倒坠床
            bln跌倒坠床 = True
        End If
        rsTmp.MoveNext
    Next

    txtInfo(txt压疮发生期间).Enabled = bln压疮
    txtInfo(txt压疮分期).Enabled = bln压疮
    lblInfo(txt压疮发生期间).Enabled = bln压疮
    lblInfo(txt压疮分期).Enabled = bln压疮

    txtInfo(txt跌倒或坠床原因).Enabled = bln跌倒坠床
    txtInfo(txt跌倒或坠床伤害).Enabled = bln跌倒坠床
    lblInfo(txt跌倒或坠床原因).Enabled = bln跌倒坠床
    lblInfo(txt跌倒或坠床伤害).Enabled = bln跌倒坠床
    
    '感染因素
    lstInfection.Clear
    StrSQL = "Select 编码, 名称" & vbNewLine & _
        "From 感染因素 A," & vbNewLine & _
        "     (Select Decode(信息值, Null, Null, ',' || 信息值 || ',') 信息值" & vbNewLine & _
        "       From 病案主页从表" & vbNewLine & _
        "       Where 病人id = [1] And 主页id = [2] And 信息名 = '感染因素') B" & vbNewLine & _
        "Where Instr(b.信息值 , chr(44)|| a.编码 ||chr(44) ) > 0"

    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng主页ID)
    For i = 1 To rsTmp.RecordCount
        lstInfection.AddItem Nvl(rsTmp!名称)
        rsTmp.MoveNext
    Next
    
    mblnCheck = False
    Screen.MousePointer = 0
    LoadPageData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Load化疗与放疗(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载放疗与化疗信息
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-21 15:55:27
    '问题:13999
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim StrSQL As String
    
    Err = 0: On Error GoTo Errhand:
    StrSQL = " " & _
    "   Select A.病人id, A.主页id, A.序号, A.疾病id, A.开始日期, A.结束日期, A.疗程数, A.总量, A.化疗方案, A.化疗效果, " & _
    "          B.编码 || '-' || B.名称 As 疾病信息 " & _
    "   From 病案化疗记录 A, 疾病编码目录 B " & _
    "   Where A.疾病id = B.Id And a.病人id=[1] And a.主页id=[2] " & _
    "   Order By 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng病人ID, lng主页ID)
    With vsChemotherapy
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Clear 1
        If rsTemp.RecordCount <> 0 Then .Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("化学治疗编码")) = Nvl(rsTemp!疾病信息)
            .TextMatrix(lngRow, .ColIndex("开始日期")) = Format(rsTemp!开始日期, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("结束日期")) = Format(rsTemp!结束日期, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("疗程数")) = Format(Val(Nvl(rsTemp!疗程数)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("总量")) = Format(Val(Nvl(rsTemp!总量)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("化疗方案")) = Trim(Nvl(rsTemp!化疗方案))
            .TextMatrix(lngRow, .ColIndex("化疗效果")) = Trim(Nvl(rsTemp!化疗效果))
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    StrSQL = " " & _
    "   Select A.病人id, A.主页id, A.序号, A.疾病id, A.开始日期, A.结束日期,A.设野部位, A.放射剂量, A.累计量, A.放疗效果, " & _
    "          B.编码 || '-' || B.名称 As 疾病信息 " & _
    "   From 病案放疗记录 A, 疾病编码目录 B " & _
    "   Where A.疾病id = B.Id And a.病人id=[1] And a.主页id=[2] " & _
    "   Order By 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng病人ID, lng主页ID)
    With vsRadiotherapy
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Clear 1
        If rsTemp.RecordCount <> 0 Then .Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("放射治疗编码")) = Nvl(rsTemp!疾病信息)
            .TextMatrix(lngRow, .ColIndex("开始日期")) = Format(rsTemp!开始日期, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("结束日期")) = Format(rsTemp!结束日期, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("放射剂量")) = Format(Val(Nvl(rsTemp!放射剂量)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("累计量")) = Format(Val(Nvl(rsTemp!累计量)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("设野部位")) = Trim(Nvl(rsTemp!设野部位))
            .TextMatrix(lngRow, .ColIndex("放疗效果")) = Trim(Nvl(rsTemp!放疗效果))
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Load化疗与放疗 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub FillVsf()
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim lngCol As Long
    Dim StrSQL As String
    
    On Error GoTo errH
    StrSQL = "select 名称,内容 from 病案项目 order by 编码"
    vsfMain.Clear
    
    Call zlDatabase.OpenRecordset(rsTemp, StrSQL, Me.Caption)
    If rsTemp.RecordCount = 0 Then vsfMain.Rows = 1: vsfMain.Cols = 1: Exit Sub
    If (rsTemp.RecordCount Mod 2) <> 0 Then
        vsfMain.Rows = rsTemp.RecordCount \ 2 + 2
    Else
        vsfMain.Rows = rsTemp.RecordCount \ 2 + 1
    End If
    With vsfMain
        .Cols = 6
        For lngRow = 0 To 3 Step 3
            .TextMatrix(0, lngRow) = "项目"
            .TextMatrix(0, lngRow + 1) = "内容"
            .ColWidth(0 + lngRow) = 1500
            .ColWidth(1 + lngRow) = 1200
            .ColWidth(2 + lngRow) = 0
        Next lngRow
        .Cell(flexcpAlignment, 0, 0, 0, vsfMain.Cols - 1) = 4
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, 0) = &HFCE7D8
        .Cell(flexcpBackColor, 1, 3, .Rows - 1, 3) = &HFCE7D8
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
    End With
    lngRow = 1
    lngCol = 0
    While Not rsTemp.EOF
        If lngCol < 4 Then
            With vsfMain
                .TextMatrix(lngRow, lngCol + 0) = rsTemp!名称
                .TextMatrix(lngRow, lngCol + 2) = rsTemp!内容 & ""
                If InStr(rsTemp!内容, "是否") > 0 Then
                    vsfMain.TextMatrix(lngRow, lngCol + 1) = "是"
                    vsfMain.Cell(flexcpChecked, lngRow, lngCol + 1) = 2
                End If
            End With
            lngCol = lngCol + 3
            rsTemp.MoveNext
        Else
            lngCol = 0
            lngRow = lngRow + 1
        End If
    Wend
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

