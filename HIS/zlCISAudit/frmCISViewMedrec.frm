VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCISViewMedrec 
   Caption         =   "病案主页查阅"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9765
   Icon            =   "frmCISViewMedrec.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7320
      Left            =   0
      ScaleHeight     =   7320
      ScaleWidth      =   8640
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8640
      Begin VB.VScrollBar vsc 
         Height          =   6945
         Left            =   8205
         TabIndex        =   187
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.HScrollBar hsc 
         Height          =   255
         Left            =   30
         TabIndex        =   186
         TabStop         =   0   'False
         Top             =   7005
         Visible         =   0   'False
         Width           =   8115
      End
      Begin VB.Frame fraVH 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8190
         TabIndex        =   185
         Top             =   7035
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame fraBack 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7200
         Left            =   90
         TabIndex        =   1
         Top             =   75
         Width           =   8175
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   基本信息 "
            ForeColor       =   &H00FF0000&
            Height          =   765
            Index           =   0
            Left            =   90
            TabIndex        =   112
            Tag             =   "4545"
            Top             =   60
            Width           =   7830
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "再入院"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   10
               Left            =   4185
               TabIndex        =   149
               Top             =   338
               Width           =   900
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
               TabIndex        =   148
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   34
               Left            =   6525
               Locked          =   -1  'True
               MaxLength       =   6
               TabIndex        =   147
               Top             =   3555
               Width           =   1095
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   33
               Left            =   2880
               Locked          =   -1  'True
               MaxLength       =   6
               TabIndex        =   146
               Top             =   2805
               Width           =   1095
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   32
               Left            =   6105
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   145
               Top             =   1410
               Width           =   1500
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   31
               Left            =   6105
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   144
               Top             =   1095
               Width           =   1500
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   30
               Left            =   1170
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   143
               Top             =   1410
               Width           =   1410
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   29
               Left            =   3480
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   142
               Top             =   1410
               Width           =   1410
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   28
               Left            =   3480
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   141
               Top             =   1095
               Width           =   1410
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   27
               Left            =   3480
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   140
               Top             =   780
               Width           =   1410
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   0
               Left            =   1170
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   139
               Top             =   345
               Width           =   1410
            End
            Begin VB.TextBox txtInfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   2
               Left            =   3195
               Locked          =   -1  'True
               TabIndex        =   138
               Top             =   345
               Width           =   285
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   3
               Left            =   1170
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   137
               Top             =   780
               Width           =   1410
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   4
               Left            =   6105
               Locked          =   -1  'True
               TabIndex        =   136
               Top             =   780
               Width           =   1500
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   5
               Left            =   1170
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   135
               Top             =   1095
               Width           =   1410
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   6
               Left            =   1170
               Locked          =   -1  'True
               MaxLength       =   30
               TabIndex        =   134
               Top             =   1725
               Width           =   2805
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   7
               Left            =   5295
               Locked          =   -1  'True
               MaxLength       =   18
               TabIndex        =   133
               Top             =   1725
               Width           =   2325
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   1
               Left            =   1170
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   132
               Top             =   2175
               Width           =   2805
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   8
               Left            =   4665
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   131
               Top             =   2175
               Width           =   1185
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   9
               Left            =   6525
               Locked          =   -1  'True
               MaxLength       =   6
               TabIndex        =   130
               Top             =   2175
               Width           =   1095
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   10
               Left            =   1170
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   129
               Top             =   2490
               Width           =   2805
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   11
               Left            =   4665
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   128
               Top             =   2490
               Width           =   1185
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   12
               Left            =   6525
               Locked          =   -1  'True
               MaxLength       =   6
               TabIndex        =   127
               Top             =   2490
               Width           =   1095
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   13
               Left            =   1170
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   126
               Top             =   2805
               Width           =   1035
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   14
               Left            =   4665
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   125
               Top             =   2805
               Width           =   1185
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   15
               Left            =   1170
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   124
               Top             =   3120
               Width           =   4680
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   16
               Left            =   1125
               Locked          =   -1  'True
               TabIndex        =   123
               Top             =   3555
               Width           =   1455
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   17
               Left            =   3270
               Locked          =   -1  'True
               TabIndex        =   122
               Top             =   3555
               Width           =   975
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   18
               Left            =   4905
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   121
               Top             =   3555
               Width           =   945
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   19
               Left            =   1125
               Locked          =   -1  'True
               TabIndex        =   120
               Top             =   4185
               Width           =   1455
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   20
               Left            =   3270
               Locked          =   -1  'True
               TabIndex        =   119
               Top             =   4185
               Width           =   975
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   21
               Left            =   4905
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   118
               Top             =   4185
               Width           =   945
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   22
               Left            =   6885
               Locked          =   -1  'True
               TabIndex        =   117
               Top             =   4185
               Width           =   735
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   23
               Left            =   1170
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   116
               Top             =   3870
               Width           =   1530
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   24
               Left            =   3645
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   115
               Top             =   3870
               Width           =   1530
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   25
               Left            =   6090
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   114
               Top             =   3870
               Width           =   1530
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   26
               Left            =   6105
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   113
               Top             =   345
               Width           =   1500
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   34
               X1              =   6000
               X2              =   7620
               Y1              =   4050
               Y2              =   4050
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   33
               X1              =   3555
               X2              =   5175
               Y1              =   4050
               Y2              =   4050
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   32
               X1              =   1080
               X2              =   2700
               Y1              =   4050
               Y2              =   4050
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   31
               X1              =   6795
               X2              =   7620
               Y1              =   4365
               Y2              =   4365
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   30
               X1              =   6435
               X2              =   7620
               Y1              =   3735
               Y2              =   3735
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   29
               X1              =   4830
               X2              =   5860
               Y1              =   4365
               Y2              =   4365
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   28
               X1              =   4830
               X2              =   5860
               Y1              =   3735
               Y2              =   3735
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   27
               X1              =   3195
               X2              =   4255
               Y1              =   4365
               Y2              =   4365
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   26
               X1              =   3195
               X2              =   4255
               Y1              =   3735
               Y2              =   3735
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   25
               X1              =   1080
               X2              =   2600
               Y1              =   4365
               Y2              =   4365
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   24
               X1              =   1080
               X2              =   2600
               Y1              =   3735
               Y2              =   3735
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   23
               X1              =   2790
               X2              =   3975
               Y1              =   2985
               Y2              =   2985
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   22
               X1              =   1080
               X2              =   2205
               Y1              =   2985
               Y2              =   2985
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   21
               X1              =   6435
               X2              =   7620
               Y1              =   2670
               Y2              =   2670
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   20
               X1              =   6435
               X2              =   7620
               Y1              =   2355
               Y2              =   2355
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   19
               X1              =   4575
               X2              =   5850
               Y1              =   2985
               Y2              =   2985
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   18
               X1              =   4575
               X2              =   5850
               Y1              =   2670
               Y2              =   2670
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   17
               X1              =   4575
               X2              =   5850
               Y1              =   2355
               Y2              =   2355
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   16
               X1              =   1080
               X2              =   5850
               Y1              =   3300
               Y2              =   3300
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   15
               X1              =   1080
               X2              =   3975
               Y1              =   2670
               Y2              =   2670
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   14
               X1              =   1080
               X2              =   3975
               Y1              =   2355
               Y2              =   2355
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   13
               X1              =   5205
               X2              =   7620
               Y1              =   1905
               Y2              =   1905
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   12
               X1              =   1080
               X2              =   3975
               Y1              =   1905
               Y2              =   1905
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   11
               X1              =   6015
               X2              =   7605
               Y1              =   1590
               Y2              =   1590
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   10
               X1              =   6015
               X2              =   7605
               Y1              =   1275
               Y2              =   1275
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   9
               X1              =   6015
               X2              =   7605
               Y1              =   960
               Y2              =   960
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   8
               X1              =   3390
               X2              =   4890
               Y1              =   1590
               Y2              =   1590
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   7
               X1              =   3390
               X2              =   4890
               Y1              =   1275
               Y2              =   1275
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   6
               X1              =   3390
               X2              =   4890
               Y1              =   960
               Y2              =   960
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   5
               X1              =   1080
               X2              =   2580
               Y1              =   1590
               Y2              =   1590
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   4
               X1              =   1080
               X2              =   2580
               Y1              =   1275
               Y2              =   1275
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   3
               X1              =   1080
               X2              =   2580
               Y1              =   960
               Y2              =   960
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   2
               X1              =   6015
               X2              =   7605
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   1
               X1              =   3195
               X2              =   3485
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   0
               X1              =   1080
               X2              =   2580
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院号"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   0
               Left            =   510
               TabIndex        =   184
               Top             =   345
               Width           =   540
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "第    次住院"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   3
               Left            =   2985
               TabIndex        =   183
               Top             =   345
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "付款方式"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   2
               Left            =   5250
               TabIndex        =   182
               Top             =   345
               Width           =   720
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
               TabIndex        =   181
               Top             =   780
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "性别"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   5
               Left            =   2985
               TabIndex        =   180
               Top             =   780
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
               Left            =   5250
               TabIndex        =   179
               Top             =   780
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "年龄"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   7
               Left            =   690
               TabIndex        =   178
               Top             =   1095
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "婚姻"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   8
               Left            =   2985
               TabIndex        =   177
               Top             =   1095
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "职业"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   9
               Left            =   5610
               TabIndex        =   176
               Top             =   1095
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "区域"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   11
               Left            =   5610
               TabIndex        =   175
               Top             =   1410
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "国籍"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   12
               Left            =   690
               TabIndex        =   174
               Top             =   1410
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "民族"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   13
               Left            =   2985
               TabIndex        =   173
               Top             =   1410
               Width           =   360
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
               TabIndex        =   172
               Top             =   1725
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "身份证号"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   15
               Left            =   4455
               TabIndex        =   171
               Top             =   1725
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "家庭地址"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   1
               Left            =   330
               TabIndex        =   170
               Top             =   2175
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "电话"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   16
               Left            =   4185
               TabIndex        =   169
               Top             =   2175
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "邮编"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   17
               Left            =   6045
               TabIndex        =   168
               Top             =   2175
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "工作单位"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   18
               Left            =   330
               TabIndex        =   167
               Top             =   2490
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "电话"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   19
               Left            =   4185
               TabIndex        =   166
               Top             =   2490
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "邮编"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   20
               Left            =   6045
               TabIndex        =   165
               Top             =   2490
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "联系人姓名"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   21
               Left            =   150
               TabIndex        =   164
               Top             =   2805
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "关系"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   22
               Left            =   2400
               TabIndex        =   163
               Top             =   2805
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "电话"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   23
               Left            =   4185
               TabIndex        =   162
               Top             =   2805
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "联系人地址"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   24
               Left            =   150
               TabIndex        =   161
               Top             =   3120
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院时间"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   25
               Left            =   330
               TabIndex        =   160
               Top             =   3555
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "科室"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   26
               Left            =   2805
               TabIndex        =   159
               Top             =   3555
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病室"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   27
               Left            =   4440
               TabIndex        =   158
               Top             =   3555
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病情"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   28
               Left            =   6045
               TabIndex        =   157
               Top             =   3555
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出院时间"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   29
               Left            =   330
               TabIndex        =   156
               Top             =   4185
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "科室"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   30
               Left            =   2805
               TabIndex        =   155
               Top             =   4185
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病室"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   31
               Left            =   4440
               TabIndex        =   154
               Top             =   4185
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "转科情况"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   33
               Left            =   330
               TabIndex        =   153
               Top             =   3870
               Width           =   720
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "───→"
               ForeColor       =   &H00808080&
               Height          =   180
               Index           =   34
               Left            =   2760
               TabIndex        =   152
               Top             =   3870
               Width           =   720
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "───→"
               ForeColor       =   &H00808080&
               Height          =   180
               Index           =   35
               Left            =   5235
               TabIndex        =   151
               Top             =   3870
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院天数"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   32
               Left            =   6045
               TabIndex        =   150
               Top             =   4185
               Width           =   720
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   住院情况 "
            ForeColor       =   &H00FF0000&
            Height          =   885
            Index           =   4
            Left            =   90
            TabIndex        =   63
            Tag             =   "3120"
            Top             =   1095
            Width           =   7830
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   64
               Left            =   1140
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   88
               Top             =   300
               Width           =   3690
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "科研病案"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   11
               Left            =   6570
               TabIndex        =   87
               Top             =   660
               Width           =   1020
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   63
               Left            =   6225
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   86
               Top             =   1650
               Width           =   1335
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "新发肿瘤"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   5
               Left            =   6570
               TabIndex        =   85
               Top             =   345
               Width           =   1020
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
               TabIndex        =   84
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   56
               Left            =   6225
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   83
               Top             =   2715
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   55
               Left            =   6225
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   82
               Top             =   2400
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   54
               Left            =   1155
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   81
               Top             =   2400
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   53
               Left            =   3750
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   80
               Top             =   2715
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   52
               Left            =   3750
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   79
               Top             =   2400
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   51
               Left            =   1155
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   78
               Top             =   2715
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   50
               Left            =   6225
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   77
               Top             =   2070
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   49
               Left            =   1155
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   76
               Top             =   2085
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   48
               Left            =   3750
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   75
               Top             =   1020
               Width           =   1080
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   47
               Left            =   1155
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   74
               Top             =   1020
               Width           =   1080
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "尸检"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   6
               Left            =   5415
               TabIndex        =   73
               Top             =   345
               Width           =   660
            End
            Begin VB.CheckBox chkInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "随诊"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   7
               Left            =   645
               TabIndex        =   72
               Top             =   660
               Width           =   660
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   46
               Left            =   2475
               Locked          =   -1  'True
               MaxLength       =   5
               TabIndex        =   71
               Top             =   645
               Width           =   1080
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "示教病案"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   8
               Left            =   5415
               TabIndex        =   70
               Top             =   660
               Width           =   1020
            End
            Begin VB.CheckBox chkInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "输血反应"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   9
               Left            =   5340
               TabIndex        =   69
               Top             =   1020
               Width           =   1020
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   45
               Left            =   1155
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   68
               Top             =   1335
               Width           =   1080
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   44
               Left            =   3750
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   67
               Top             =   1335
               Width           =   1080
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   43
               Left            =   6225
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   66
               Top             =   1335
               Width           =   1080
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   42
               Left            =   1155
               MaxLength       =   10
               TabIndex        =   65
               Top             =   1650
               Width           =   1080
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   41
               Left            =   3750
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   64
               Top             =   1650
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "死亡原因"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   66
               Left            =   285
               TabIndex        =   111
               Top             =   300
               Width           =   720
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   61
               X1              =   1050
               X2              =   5050
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输液反应"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   65
               Left            =   5370
               TabIndex        =   110
               Top             =   1650
               Width           =   720
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   60
               X1              =   6135
               X2              =   7560
               Y1              =   1830
               Y2              =   1830
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   56
               X1              =   6135
               X2              =   7560
               Y1              =   2895
               Y2              =   2895
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   55
               X1              =   6135
               X2              =   7560
               Y1              =   2580
               Y2              =   2580
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   54
               X1              =   1065
               X2              =   2490
               Y1              =   2580
               Y2              =   2580
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   53
               X1              =   3660
               X2              =   5085
               Y1              =   2895
               Y2              =   2895
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   52
               X1              =   3660
               X2              =   5085
               Y1              =   2580
               Y2              =   2580
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   51
               X1              =   1065
               X2              =   2490
               Y1              =   2895
               Y2              =   2895
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   50
               X1              =   6135
               X2              =   7560
               Y1              =   2250
               Y2              =   2250
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   49
               X1              =   1065
               X2              =   2490
               Y1              =   2265
               Y2              =   2265
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   48
               X1              =   6135
               X2              =   7305
               Y1              =   1515
               Y2              =   1515
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   47
               X1              =   3660
               X2              =   4830
               Y1              =   1830
               Y2              =   1830
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   46
               X1              =   3660
               X2              =   4830
               Y1              =   1515
               Y2              =   1515
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   45
               X1              =   3660
               X2              =   4830
               Y1              =   1200
               Y2              =   1200
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   44
               X1              =   1065
               X2              =   2235
               Y1              =   1830
               Y2              =   1830
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   43
               X1              =   1065
               X2              =   2235
               Y1              =   1515
               Y2              =   1515
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   42
               X1              =   1065
               X2              =   2235
               Y1              =   1200
               Y2              =   1200
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   41
               X1              =   2385
               X2              =   3555
               Y1              =   825
               Y2              =   825
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "随诊期限"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   44
               Left            =   1635
               TabIndex        =   109
               Top             =   660
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "血型"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   45
               Left            =   660
               TabIndex        =   108
               Top             =   1020
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rh"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   46
               Left            =   3450
               TabIndex        =   107
               Top             =   1020
               Width           =   180
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输红细胞"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   47
               Left            =   300
               TabIndex        =   106
               Top             =   1335
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "单位"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   48
               Left            =   2280
               TabIndex        =   105
               Top             =   1335
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输血小板"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   49
               Left            =   2910
               TabIndex        =   104
               Top             =   1335
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "袋"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   50
               Left            =   4875
               TabIndex        =   103
               Top             =   1335
               Width           =   180
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输血浆"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   51
               Left            =   5550
               TabIndex        =   102
               Top             =   1335
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ml"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   52
               Left            =   7380
               TabIndex        =   101
               Top             =   1335
               Width           =   180
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输全血"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   53
               Left            =   480
               TabIndex        =   100
               Top             =   1650
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ml"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   54
               Left            =   2280
               TabIndex        =   99
               Top             =   1650
               Width           =   180
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输其他"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   55
               Left            =   3090
               TabIndex        =   98
               Top             =   1650
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ml"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   56
               Left            =   4890
               TabIndex        =   97
               Top             =   1650
               Width           =   180
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门诊医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   57
               Left            =   300
               TabIndex        =   96
               Top             =   2085
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "科主任"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   58
               Left            =   480
               TabIndex        =   95
               Top             =   2400
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "主任(副主任)医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   59
               Left            =   4650
               TabIndex        =   94
               Top             =   2070
               Width           =   1440
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "主治医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   60
               Left            =   2910
               TabIndex        =   93
               Top             =   2400
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   61
               Left            =   5370
               TabIndex        =   92
               Top             =   2400
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "进修医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   62
               Left            =   300
               TabIndex        =   91
               Top             =   2715
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "研究生医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   63
               Left            =   2730
               TabIndex        =   90
               Top             =   2715
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "实习医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   64
               Left            =   5370
               TabIndex        =   89
               Top             =   2715
               Width           =   720
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "   过敏与手术 "
            ForeColor       =   &H00FF0000&
            Height          =   765
            Index           =   3
            Left            =   90
            TabIndex        =   49
            Tag             =   "3795"
            Top             =   3450
            Width           =   7830
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   75
               Left            =   960
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   57
               Top             =   1710
               Width           =   780
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   74
               Left            =   2805
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   56
               Top             =   1710
               Width           =   780
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   73
               Left            =   4680
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   55
               Top             =   1710
               Width           =   780
            End
            Begin VB.CheckBox chkInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "手术、治疗、检查、诊断为本院第一例"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   1
               Left            =   4305
               TabIndex        =   54
               Top             =   3495
               Width           =   3360
            End
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
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   62
               Left            =   2415
               Locked          =   -1  'True
               MaxLength       =   16
               TabIndex        =   52
               Top             =   2730
               Width           =   1695
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   61
               Left            =   5535
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   51
               Top             =   2730
               Width           =   510
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   60
               Left            =   7110
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   50
               Top             =   2730
               Width           =   510
            End
            Begin VSFlex8Ctl.VSFlexGrid vsOPS 
               Height          =   1335
               Left            =   165
               TabIndex        =   58
               Top             =   2085
               Width           =   7500
               _cx             =   13229
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
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmCISViewMedrec.frx":6852
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
            Begin VSFlex8Ctl.VSFlexGrid vsAller 
               Height          =   1335
               Left            =   165
               TabIndex        =   59
               Top             =   300
               Width           =   7500
               _cx             =   13229
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
               Cols            =   2
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmCISViewMedrec.frx":694D
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
               Caption         =   "HIV-Ab"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   77
               Left            =   4005
               TabIndex        =   62
               Top             =   1710
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HCV-Ab"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   76
               Left            =   2145
               TabIndex        =   61
               Top             =   1710
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HBsAg"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   75
               Left            =   375
               TabIndex        =   60
               Top             =   1710
               Width           =   450
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   72
               X1              =   870
               X2              =   1865
               Y1              =   1890
               Y2              =   1890
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   71
               X1              =   2715
               X2              =   3715
               Y1              =   1890
               Y2              =   1890
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   70
               X1              =   4590
               X2              =   5615
               Y1              =   1890
               Y2              =   1890
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   中医诊断 "
            ForeColor       =   &H00FF0000&
            Height          =   840
            Index           =   2
            Left            =   90
            TabIndex        =   24
            Tag             =   "4050"
            Top             =   4440
            Width           =   7830
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   72
               Left            =   4020
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   45
               Top             =   2190
               Width           =   915
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   71
               Left            =   1425
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   44
               Top             =   2190
               Width           =   915
            End
            Begin VB.Frame fraSub 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   " 准确度 "
               ForeColor       =   &H00404040&
               Height          =   1320
               Index           =   1
               Left            =   2160
               TabIndex        =   37
               Top             =   2580
               Width           =   2385
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   57
                  Left            =   840
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   40
                  Top             =   330
                  Width           =   1035
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   58
                  Left            =   840
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   39
                  Top             =   645
                  Width           =   1035
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   59
                  Left            =   840
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   38
                  Top             =   960
                  Width           =   1035
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "辨证"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   40
                  Left            =   330
                  TabIndex        =   43
                  Top             =   330
                  Width           =   360
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "治法"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   39
                  Left            =   330
                  TabIndex        =   42
                  Top             =   645
                  Width           =   360
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "方药"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   38
                  Left            =   330
                  TabIndex        =   41
                  Top             =   960
                  Width           =   360
               End
               Begin VB.Line lineInfo 
                  BorderColor     =   &H00808080&
                  Index           =   59
                  X1              =   750
                  X2              =   1875
                  Y1              =   510
                  Y2              =   510
               End
               Begin VB.Line lineInfo 
                  BorderColor     =   &H00808080&
                  Index           =   58
                  X1              =   750
                  X2              =   1875
                  Y1              =   825
                  Y2              =   825
               End
               Begin VB.Line lineInfo 
                  BorderColor     =   &H00808080&
                  Index           =   57
                  X1              =   750
                  X2              =   1875
                  Y1              =   1140
                  Y2              =   1140
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
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.Frame fraSub 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   " 住院期间病情 "
               ForeColor       =   &H00404040&
               Height          =   1320
               Index           =   0
               Left            =   165
               TabIndex        =   32
               Top             =   2580
               Width           =   1845
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "危重"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   2
                  Left            =   525
                  TabIndex        =   35
                  Top             =   330
                  Width           =   660
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "急症"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   3
                  Left            =   525
                  TabIndex        =   34
                  Top             =   645
                  Width           =   660
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "疑难"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   4
                  Left            =   525
                  TabIndex        =   33
                  Top             =   960
                  Width           =   660
               End
            End
            Begin VB.Frame fraSub 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   " 治疗方法 "
               ForeColor       =   &H00404040&
               Height          =   1320
               Index           =   2
               Left            =   4680
               TabIndex        =   25
               Top             =   2580
               Width           =   2985
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   40
                  Left            =   1545
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   28
                  Top             =   960
                  Width           =   1035
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   39
                  Left            =   1545
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   27
                  Top             =   645
                  Width           =   1035
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   38
                  Left            =   1545
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   26
                  Top             =   330
                  Width           =   1035
               End
               Begin VB.Line lineInfo 
                  BorderColor     =   &H00808080&
                  Index           =   40
                  X1              =   1455
                  X2              =   2580
                  Y1              =   1140
                  Y2              =   1140
               End
               Begin VB.Line lineInfo 
                  BorderColor     =   &H00808080&
                  Index           =   39
                  X1              =   1455
                  X2              =   2580
                  Y1              =   825
                  Y2              =   825
               End
               Begin VB.Line lineInfo 
                  BorderColor     =   &H00808080&
                  Index           =   38
                  X1              =   1455
                  X2              =   2580
                  Y1              =   510
                  Y2              =   510
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "自制中药制剂"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   41
                  Left            =   315
                  TabIndex        =   31
                  Top             =   960
                  Width           =   1080
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "抢救方法"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   42
                  Left            =   675
                  TabIndex        =   30
                  Top             =   645
                  Width           =   720
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "治疗类别"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   43
                  Left            =   675
                  TabIndex        =   29
                  Top             =   330
                  Width           =   720
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
               Height          =   1830
               Left            =   135
               TabIndex        =   46
               Top             =   285
               Width           =   7500
               _cx             =   13229
               _cy             =   3228
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
               Rows            =   6
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmCISViewMedrec.frx":699E
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
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   69
               X1              =   3930
               X2              =   5015
               Y1              =   2370
               Y2              =   2370
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   68
               X1              =   1335
               X2              =   2465
               Y1              =   2370
               Y2              =   2370
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门诊与出院"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   74
               Left            =   390
               TabIndex        =   48
               Top             =   2190
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院与出院"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   73
               Left            =   3000
               TabIndex        =   47
               Top             =   2190
               Width           =   900
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "   西医诊断 "
            ForeColor       =   &H00FF0000&
            Height          =   1125
            Index           =   1
            Left            =   90
            TabIndex        =   2
            Tag             =   "3900"
            Top             =   2160
            Width           =   7830
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   70
               Left            =   1335
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   13
               Top             =   3180
               Width           =   915
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   69
               Left            =   1335
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   12
               Top             =   3495
               Width           =   915
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   68
               Left            =   3930
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   11
               Top             =   3180
               Width           =   915
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   67
               Left            =   3930
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   10
               Top             =   3495
               Width           =   915
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   66
               Left            =   6405
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   9
               Top             =   3180
               Width           =   915
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   65
               Left            =   6405
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   8
               Top             =   3495
               Width           =   915
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
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   37
               Left            =   2415
               Locked          =   -1  'True
               MaxLength       =   16
               TabIndex        =   6
               Top             =   2730
               Width           =   1695
            End
            Begin VB.CheckBox chkInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "是否确诊"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   5
               Top             =   2730
               Width           =   1020
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   36
               Left            =   5535
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   4
               Top             =   2730
               Width           =   510
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   35
               Left            =   7110
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   3
               Top             =   2730
               Width           =   510
            End
            Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
               Height          =   2355
               Left            =   135
               TabIndex        =   14
               Top             =   270
               Width           =   7500
               _cx             =   13229
               _cy             =   4154
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
               Rows            =   7
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmCISViewMedrec.frx":6A3C
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
               Caption         =   "术前与术后"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   72
               Left            =   5370
               TabIndex        =   23
               Top             =   3495
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "临床与尸检"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   71
               Left            =   2910
               TabIndex        =   22
               Top             =   3495
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "临床与病理"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   70
               Left            =   300
               TabIndex        =   21
               Top             =   3495
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "放射与病理"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   69
               Left            =   5370
               TabIndex        =   20
               Top             =   3180
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院与出院"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   68
               Left            =   2910
               TabIndex        =   19
               Top             =   3180
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门诊与出院"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   67
               Left            =   300
               TabIndex        =   18
               Top             =   3180
               Width           =   900
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   67
               X1              =   1245
               X2              =   2375
               Y1              =   3360
               Y2              =   3360
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   66
               X1              =   1245
               X2              =   2375
               Y1              =   3675
               Y2              =   3675
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   65
               X1              =   3840
               X2              =   4925
               Y1              =   3360
               Y2              =   3360
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   64
               X1              =   3840
               X2              =   4925
               Y1              =   3675
               Y2              =   3675
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   63
               X1              =   6315
               X2              =   7475
               Y1              =   3360
               Y2              =   3360
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   62
               X1              =   6315
               X2              =   7475
               Y1              =   3675
               Y2              =   3675
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   37
               X1              =   7020
               X2              =   7620
               Y1              =   2910
               Y2              =   2910
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   36
               X1              =   5445
               X2              =   6045
               Y1              =   2910
               Y2              =   2910
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   35
               X1              =   2325
               X2              =   4110
               Y1              =   2910
               Y2              =   2910
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "确诊日期"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   37
               Left            =   1575
               TabIndex        =   17
               Top             =   2730
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "抢救次数"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   10
               Left            =   4695
               TabIndex        =   16
               Top             =   2730
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "成功次数"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   36
               Left            =   6255
               TabIndex        =   15
               Top             =   2730
               Width           =   720
            End
         End
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
               Picture         =   "frmCISViewMedrec.frx":6B1E
               Key             =   "-"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCISViewMedrec.frx":7008
               Key             =   "+"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCISViewMedrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public mfrmParent As Object
'Public mstrPrivs As String

'上次刷新数据时的病人信息
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng病区ID As Long
Private mlng科室ID As Long
Private mbln出院 As Boolean
Private mblnMoved As Boolean
Private mbln中医 As Boolean
Private mblnCheck As Boolean
Private mblnOpenForm As Boolean
Private Enum 基本信息
    txt付款方式 = 26
    txt性别 = 27
    txt婚姻 = 28
    txt职业 = 31
    txt入院病情 = 34
    txt区域 = 32
    txt国籍 = 30
    txt民族 = 29
    txt联系人关系 = 33
    txt住院号 = 0
    txt家庭地址 = 1
    txt住院次数 = 2
    txt姓名 = 3
    txt出生日期 = 4
    txt年龄 = 5
    txt出生地点 = 6
    txt身份证号 = 7
    txt家庭电话 = 8
    txt家庭邮编 = 9
    txt工作单位 = 10
    txt单位电话 = 11
    txt单位邮编 = 12
    txt联系人姓名 = 13
    txt联系人电话 = 14
    txt联系人地址 = 15
    txt入院时间 = 16
    txt入院科室 = 17
    txt入院病室 = 18
    txt出院时间 = 19
    txt出院科室 = 20
    txt出院病室 = 21
    txt住院天数 = 22
    txt转科1 = 23
    txt转科2 = 24
    txt转科3 = 25
    chk再入院 = 10
End Enum
Private Enum 西医诊断
    chk是否确诊 = 0
    txt抢救次数 = 36
    txt确诊日期 = 37
    txt成功次数 = 35
    txt门诊与出院 = 70
    txt入院与出院 = 68
    txt放射与病理 = 66
    txt临床与病理 = 69
    txt临床与尸检 = 67
    txt术前与术后 = 65
End Enum
Private Enum 中医诊断
    chk危重 = 2
    chk急症 = 3
    chk疑难 = 4
    txt辨证 = 57
    txtHBsAg = 75
    txtHCVAb = 74
    txtHIVAb = 73
    txt治法 = 58
    txt方药 = 59
    txt自制中药 = 40
    txt抢救方法 = 39
    txt治疗类别 = 38
    txt中医门诊与出院 = 71
    txt中医入院与出院 = 72
End Enum
Private Enum 过敏与手术
    chk首例 = 1
End Enum
Private Enum 住院情况
    chk尸检 = 6
    chk随诊 = 7
    chk新发肿瘤 = 5
    chk示教病案 = 8
    chk科研病案 = 11
    chk输血反应 = 9
    txt死亡原因 = 64
    txt随诊期限 = 46
    txt输红细胞 = 45
    txt输血小板 = 44
    txt输血浆 = 43
    txt输全血 = 42
    txt输其他 = 41
    txt血型 = 47
    txtRh = 48
    txt输液反应 = 63
    txt门诊医师 = 49
    txt科主任 = 54
    txt主任医师 = 50
    txt主治医师 = 52
    txt住院医师 = 55
    txt进修医师 = 51
    txt研究生医师 = 53
    txt实习医师 = 56
End Enum
Private Enum 诊断情况
    col诊断类型 = 0
    col诊断描述 = 1
    col出院情况 = 2
    col是否未治 = 3
    col是否疑诊 = 4
End Enum
Private Enum 手术情况
    col手术日期 = 0
    col手术名称 = 1
    col主刀医师 = 2
    col助手1 = 3
    col助手2 = 4
    col麻醉方式 = 5
    col麻醉医师 = 6
    col切口愈合 = 7
End Enum

Private Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'功能：显示或隐藏一个窗体的标题栏
'参数：blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(objForm.hWnd, vRect)
    lngStyle = GetWindowLong(objForm.hWnd, GWL_STYLE)
    If blnCaption Then
        lngStyle = lngStyle Or WS_CAPTION Or WS_THICKFRAME
        If objForm.ControlBox Then lngStyle = lngStyle Or WS_SYSMENU
        If objForm.MaxButton Then lngStyle = lngStyle Or WS_MAXIMIZEBOX
        If objForm.MinButton Then lngStyle = lngStyle Or WS_MINIMIZEBOX
    Else
        If blnBorder Then
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        Else
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
        End If
    End If
    SetWindowLong objForm.hWnd, GWL_STYLE, lngStyle
    SetWindowPos objForm.hWnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub

Public Function zlRefresh(ByVal frmMain As Object, lng病人ID As Long, lng主页ID As Long, lng病区ID As Long, lng科室ID As Long, Optional bln出院 As Boolean) As Boolean
'功能：刷新或清除医嘱清单
    
    mlng病人ID = lng病人ID: mlng主页ID = lng主页ID
    mlng病区ID = lng病区ID: mlng科室ID = lng科室ID
    mbln出院 = bln出院: mblnMoved = False
        
    '可能科室切换
    mbln中医 = Have部门性质(mlng科室ID, "中医科")
    fraInfo(2).Visible = mbln中医
    fraInfo(2).Enabled = mbln中医 '标志不操作
    Call SetPageHeight
    Call SetScrollbar
    
    Call ClearPageData
    If mlng病人ID <> 0 Then Call LoadPageData

    Me.Show 1, frmMain
    
End Function

Private Sub chkInfo_Click(Index As Integer)
    If Not mblnCheck Then
        mblnCheck = True
        chkInfo(Index).Value = IIf(chkInfo(Index).Value = 1, 0, 1)
        mblnCheck = False
    End If
End Sub

Private Sub Form_Load()
    '滚动条尺寸
    vsc.Width = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    hsc.Height = GetSystemMetrics(SM_CXHSCROLL) * Screen.TwipsPerPixelY
    fraVH.Width = vsc.Width: fraVH.Height = hsc.Height
    fraBack.Left = 0: fraBack.Top = 0
    picBack.BackColor = fraBack.BackColor
    fraInfo(1).Left = fraInfo(0).Left
'    '初始化系统参数
'    Call InitSysPar
    Call RestoreWinState(Me, App.ProductName)
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
    If fraBack.Width <= picBack.ScaleWidth Then
        hsc.Visible = False
    Else
        hsc.Min = 0
        hsc.SmallChange = 5
        hsc.LargeChange = 50
        If Not hsc.Visible Then hsc.Value = 0
        hsc.Visible = True
    End If
    
    If fraBack.Height <= picBack.ScaleHeight Then
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

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
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

Private Sub txtInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtInfo(Index).Locked Then
        glngTXTProc = GetWindowLong(txtInfo(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtInfo(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtInfo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtInfo(Index).Locked Then
        Call SetWindowLong(txtInfo(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
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
        For i = .FixedRows To .Rows - 1
            For j = .FixedCols To .Cols - 1
                .TextMatrix(i, j) = ""
            Next
        Next
        For i = .Rows - 4 To 4 Step -1
            .RemoveItem i
        Next
    End With
    With vsDiagZY
        For i = .FixedRows To .Rows - 1
            For j = .FixedCols To .Cols - 1
                .TextMatrix(i, j) = ""
            Next
        Next
        .Rows = .FixedRows + 5
    End With
    With vsOPS
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    End With
    With vsAller
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    End With
    
    mblnCheck = False
End Sub

Private Function LoadPageData() As Boolean
'功能：读取病人的首页信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim lngRow As Long, varTmp As Variant
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    mblnCheck = True
    
    '病人信息部份
    '---------------------------------------------------------------
    strSQL = "Select 住院号,性别,姓名,出生日期,出生地点,身份证号,民族 From 病人信息 Where 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    
    txtInfo(txt住院号).Text = NVL(rsTmp!住院号)
    txtInfo(txt住院次数).Text = mlng主页ID
    txtInfo(txt姓名).Text = NVL(rsTmp!姓名)
    txtInfo(txt性别).Text = NVL(rsTmp!性别)
    
    If Format(rsTmp!出生日期, "HH:mm") <> "00:00" Then
        txtInfo(txt出生日期).Text = Format(rsTmp!出生日期, "yyyy-MM-dd HH:mm")
    Else
        txtInfo(txt出生日期).Text = Format(rsTmp!出生日期, "yyyy-MM-dd")
    End If
    
    txtInfo(txt出生地点).Text = NVL(rsTmp!出生地点)
    txtInfo(txt身份证号).Text = NVL(rsTmp!身份证号)
    txtInfo(txt民族).Text = NVL(rsTmp!民族)
    
    '病案主页部份
    '---------------------------------------------------------------
    strSQL = "Select A.病人ID,A.主页ID,A.住院号,A.病人性质,A.医疗付款方式,A.费别,A.再入院,A.入院病区ID,A.入院科室ID,A.入院日期,A.入院病况,A.入院方式,A.入院属性,A.二级院转入,A.住院目的,A.入院病床,A.是否陪伴,A.当前病况,A.当前病区ID,A.护理等级ID,A.出院科室ID,A.出院病床,A.出院日期,A.住院天数,A.出院方式,A.是否确诊,A.确诊日期,A.新发肿瘤,A.血型,A.抢救次数,A.成功次数,A.随诊标志,A.随诊期限,A.尸检标志,A.门诊医师,A.责任护士,A.住院医师,A.病案号,A.编目员编号,A.编目员姓名,A.编目日期,A.状态,A.费用和,A.年龄,A.婚姻状况,A.职业,A.国籍,A.学历,A.单位电话,A.单位邮编,A.单位地址,A.区域,A.家庭地址,A.家庭电话,A.家庭地址邮编,A.联系人姓名,A.联系人关系,A.联系人地址,A.联系人电话,A.中医治疗类别,A.险类,A.审核标志,A.审核人,A.审核日期,A.是否上传,A.数据转出,A.登记人,A.登记时间,A.备注,A.社区,A.病案状态,A.封存时间,A.病人类型,B.名称 as 入院科室,C.名称 as 出院科室" & _
        " From 病案主页 A,部门表 B,部门表 C" & _
        " Where A.入院科室ID=B.ID And A.出院科室ID=C.ID" & _
        " And A.病人ID=[1] And A.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    mblnMoved = NVL(rsTmp!数据转出, 0) = 1
    
    '留观病人无住院号
    If NVL(rsTmp!病人性质, 0) <> 0 Then
        lblInfo(0).Visible = False
        txtInfo(txt住院号).Visible = False
    End If
        
    txtInfo(txt付款方式).Text = NVL(rsTmp!医疗付款方式)
    txtInfo(txt年龄).Text = NVL(rsTmp!年龄)
    txtInfo(txt婚姻).Text = NVL(rsTmp!婚姻状况)
    txtInfo(txt职业).Text = NVL(rsTmp!职业)
    txtInfo(txt国籍).Text = NVL(rsTmp!国籍)
    txtInfo(txt区域).Text = NVL(rsTmp!区域)
    txtInfo(txt家庭地址).Text = NVL(rsTmp!家庭地址)
    txtInfo(txt家庭电话).Text = NVL(rsTmp!家庭电话)
    txtInfo(txt家庭邮编).Text = NVL(rsTmp!家庭地址邮编)
    txtInfo(txt工作单位).Text = NVL(rsTmp!单位地址)
    txtInfo(txt单位电话).Text = NVL(rsTmp!单位电话)
    txtInfo(txt单位邮编).Text = NVL(rsTmp!单位邮编)
    txtInfo(txt联系人姓名).Text = NVL(rsTmp!联系人姓名)
    txtInfo(txt联系人关系).Text = NVL(rsTmp!联系人关系)
    txtInfo(txt联系人电话).Text = NVL(rsTmp!联系人电话)
    txtInfo(txt联系人地址).Text = NVL(rsTmp!联系人地址)
    chkInfo(chk再入院).Value = NVL(rsTmp!再入院, 0)
    txtInfo(txt入院时间).Text = Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm")
    txtInfo(txt入院科室).Text = rsTmp!入院科室
    txtInfo(txt入院病情).Text = NVL(rsTmp!入院病况)
    txtInfo(txt出院时间).Text = Format(NVL(rsTmp!出院日期), "yyyy-MM-dd HH:mm")
    txtInfo(txt出院科室).Text = rsTmp!出院科室
    If Not IsNull(rsTmp!出院日期) Then
        txtInfo(txt住院天数).Text = DateDiff("d", rsTmp!入院日期, rsTmp!出院日期)
    Else
        txtInfo(txt住院天数).Text = DateDiff("d", rsTmp!入院日期, zlDatabase.Currentdate)
    End If
    If Val(txtInfo(txt住院天数).Text) = 0 Then txtInfo(txt住院天数).Text = "1"
    chkInfo(chk是否确诊).Value = NVL(rsTmp!是否确诊, 0)
    If chkInfo(chk是否确诊).Value = 1 Then
        txtInfo(txt确诊日期).Text = Format(NVL(rsTmp!确诊日期), "yyyy-MM-dd HH:mm")
    End If
    txtInfo(txt抢救次数).Text = NVL(rsTmp!抢救次数)
    If Val(txtInfo(txt抢救次数).Text) <> 0 Then
        txtInfo(txt成功次数).Text = NVL(rsTmp!成功次数)
    End If
    chkInfo(chk新发肿瘤).Value = NVL(rsTmp!新发肿瘤, 0)
    
    txtInfo(txt治疗类别).Text = NVL(rsTmp!中医治疗类别)
    chkInfo(chk尸检).Value = NVL(rsTmp!尸检标志, 0)
    chkInfo(chk随诊).Value = IIf(NVL(rsTmp!随诊标志, 0) = 0, 0, 1)
    If chkInfo(chk随诊).Value = 1 Then
        txtInfo(txt随诊期限).Text = NVL(rsTmp!随诊期限, 0) & Decode(NVL(rsTmp!随诊标志, 0), 1, "月", 2, "年", 3, "周")
    End If
    txtInfo(txt门诊医师).Text = NVL(rsTmp!门诊医师)
    txtInfo(txt住院医师).Text = NVL(rsTmp!住院医师)
    txtInfo(txt血型).Text = NVL(rsTmp!血型)
    
    '病案从表部份
    '---------------------------------------------------------------
    strSQL = "Select 病人ID,主页ID,信息名,信息值 From 病案主页从表 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    For i = 1 To rsTmp.RecordCount
        Select Case UCase(NVL(rsTmp!信息名))
            Case "入院病室"
                txtInfo(txt入院病室).Text = NVL(rsTmp!信息值)
            Case "出院病室"
                txtInfo(txt出院病室).Text = NVL(rsTmp!信息值)
            Case "转科记录"
                varTmp = Split(NVL(rsTmp!信息值), ",")
                If UBound(varTmp) >= 0 Then txtInfo(txt转科1).Text = varTmp(0)
                If UBound(varTmp) >= 1 Then txtInfo(txt转科2).Text = varTmp(1)
                If UBound(varTmp) >= 2 Then txtInfo(txt转科3).Text = varTmp(2)
            Case "死亡根本原因"
                txtInfo(txt死亡原因).Text = NVL(rsTmp!信息值)
            Case UCase("HBsAg")
                txtInfo(txtHBsAg).Text = NVL(rsTmp!信息值)
            Case UCase("HCV-Ab")
                txtInfo(txtHCVAb).Text = NVL(rsTmp!信息值)
            Case UCase("HIV-Ab")
                txtInfo(txtHIVAb).Text = NVL(rsTmp!信息值)
            Case "首例"
                chkInfo(chk首例).Value = Val(NVL(rsTmp!信息值, 0))
            Case "中医危重"
                chkInfo(chk危重).Value = Val(NVL(rsTmp!信息值, 0))
            Case "中医急症"
                chkInfo(chk急症).Value = Val(NVL(rsTmp!信息值, 0))
            Case "中医疑难"
                chkInfo(chk疑难).Value = Val(NVL(rsTmp!信息值, 0))
            Case "中医抢救方法"
                txtInfo(txt抢救方法).Text = NVL(rsTmp!信息值)
            Case "自制中药制剂"
                txtInfo(txt自制中药).Text = NVL(rsTmp!信息值)
            Case "示教病案"
                chkInfo(chk示教病案).Value = Val(NVL(rsTmp!信息值, 0))
            Case "科研病案"
                chkInfo(chk科研病案).Value = Val(NVL(rsTmp!信息值, 0))
            Case UCase("Rh")
                txtInfo(txtRh).Text = NVL(rsTmp!信息值)
            Case "输血反应"
                chkInfo(chk输血反应).Value = Val(NVL(rsTmp!信息值, 0))
            Case "输红细胞"
                txtInfo(txt输红细胞).Text = NVL(rsTmp!信息值)
            Case "输血小板"
                txtInfo(txt输血小板).Text = NVL(rsTmp!信息值)
            Case "输血浆"
                txtInfo(txt输血浆).Text = NVL(rsTmp!信息值)
            Case "输全血"
                txtInfo(txt输全血).Text = NVL(rsTmp!信息值)
            Case "输其他"
                txtInfo(txt输其他).Text = NVL(rsTmp!信息值)
            Case "输液反应"
                txtInfo(txt输液反应).Text = NVL(rsTmp!信息值)
            Case "科主任"
                txtInfo(txt科主任).Text = NVL(rsTmp!信息值)
            Case "主任医师"
                txtInfo(txt主任医师).Text = NVL(rsTmp!信息值)
            Case "主治医师"
                txtInfo(txt主治医师).Text = NVL(rsTmp!信息值)
            Case "进修医师"
                txtInfo(txt进修医师).Text = NVL(rsTmp!信息值)
            Case "研究生实习医师"
                txtInfo(txt研究生医师).Text = NVL(rsTmp!信息值)
            Case "实习医师"
                txtInfo(txt实习医师).Text = NVL(rsTmp!信息值)
        End Select
        rsTmp.MoveNext
    Next
    
    '诊断符合情况
    '---------------------------------------------------------------
    strSQL = "Select 符合类型,符合情况 From 诊断符合情况 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    Do While Not rsTmp.EOF
        Select Case rsTmp!符合类型
        Case 1 '门诊与出院
            txtInfo(txt门诊与出院).Text = Decode(NVL(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 2 '入院与出院
            txtInfo(txt入院与出院).Text = Decode(NVL(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 3 '放射与病理
            txtInfo(txt放射与病理).Text = Decode(NVL(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 4 '临床与病理
            txtInfo(txt临床与病理).Text = Decode(NVL(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 5 '临床与尸检
            txtInfo(txt临床与尸检).Text = Decode(NVL(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 6 '术前与术后
            txtInfo(txt术前与术后).Text = Decode(NVL(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 11 '中医门诊与出院
            txtInfo(txt中医门诊与出院).Text = Decode(NVL(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 12 '中医入院与出院
            txtInfo(txt中医入院与出院).Text = Decode(NVL(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
        Case 13 '中医辨证
            txtInfo(txt辨证).Text = Decode(NVL(rsTmp!符合情况, 0), 1, "准确", 2, "基本准确", 3, "重大缺陷", 4, "错误", "")
        Case 14 '中医治法
            txtInfo(txt治法).Text = Decode(NVL(rsTmp!符合情况, 0), 1, "准确", 2, "基本准确", 3, "重大缺陷", 4, "错误", "")
        Case 15 '中医方药
            txtInfo(txt方药).Text = Decode(NVL(rsTmp!符合情况, 0), 1, "准确", 2, "基本准确", 3, "重大缺陷", 4, "错误", "")
        End Select
        rsTmp.MoveNext
    Loop
    
    '自动提取转科科室及入出病室(房间号)
    '---------------------------------------------------------------
    If txtInfo(txt转科1).Text = "" And txtInfo(txt转科2).Text = "" And txtInfo(txt转科3).Text = "" Then
        strSQL = _
            " Select B.名称" & _
            " From 病人变动记录 A,部门表 B" & _
            " Where A.病人ID=[1] And A.主页ID=[2]" & _
            " And A.科室ID=B.ID And A.开始原因=3 And A.开始时间 is Not NULL" & _
            " Order by A.开始时间"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
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
        strSQL = "Select B.房间号" & _
            " From 病案主页 A,床位状况记录 B" & _
            " Where A.病人ID=[1] And A.主页ID=[2]" & _
            " And A.入院病区ID=B.病区ID And A.入院病床=B.床号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
        If Not rsTmp.EOF Then txtInfo(txt入院病室).Text = NVL(rsTmp!房间号)
    End If
    If txtInfo(txt出院病室).Text = "" Then
        strSQL = "Select B.房间号" & _
            " From 病案主页 A,床位状况记录 B" & _
            " Where A.病人ID=[1] And A.主页ID=[2]" & _
            " And A.当前病区ID=B.病区ID And A.出院病床=B.床号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
        If Not rsTmp.EOF Then txtInfo(txt出院病室).Text = NVL(rsTmp!房间号)
    End If
    
    '过敏信息:本次住院的,过敏的
    '---------------------------------------------------------------
    strSQL = "Select 记录来源,Decode(过敏时间,Null ,记录时间,过敏时间) as 过敏时间,药物ID,药物名 From 病人过敏记录 A" & _
        " Where 结果=1 And 病人ID=[1] And 主页ID=[2]" & _
        " And Not Exists(Select 药物ID From 病人过敏记录" & _
            " Where (Nvl(药物ID,0)=Nvl(A.药物ID,0) Or Nvl(药物名,'Null')=Nvl(A.药物名,'Null'))" & _
            " And Nvl(结果,0)=0 And 记录时间>=A.记录时间 And 病人ID=[1] And 主页ID=[2])" & _
        " Order by 过敏时间,药物名"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Not rsTmp.EOF Then
        rsTmp.Filter = "记录来源=3" '首页本身填写的
        If rsTmp.EOF Then rsTmp.Filter = "记录来源<>3" '其它来源的作为缺省显示
        With vsAller
            .Rows = rsTmp.RecordCount + 1 '固定行
            For i = 1 To rsTmp.RecordCount
                '其它来源的可能有重复
                lngRow = -1
                If Not IsNull(rsTmp!药物ID) Then
                    lngRow = .FindRow(CLng(rsTmp!药物ID))
                ElseIf Not IsNull(rsTmp!药物名) Then
                    lngRow = .FindRow(CStr(rsTmp!药物名), , 1)
                End If
                If lngRow = -1 Then
                    .RowData(i) = CLng(NVL(rsTmp!药物ID, 0))
                    .TextMatrix(i, 0) = Format(rsTmp!过敏时间, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, i, 0) = Format(rsTmp!过敏时间, "yyyy-MM-dd HH:mm:ss") '用于保存
                    .TextMatrix(i, 1) = NVL(rsTmp!药物名)
                    .Cell(flexcpData, i, 1) = .TextMatrix(i, 1) '用于输入恢复
                End If
                rsTmp.MoveNext
            Next
            If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        End With
    End If
    vsAller.Row = 1: vsAller.Col = 1
    
    '西医诊断
    '---------------------------------------------------------------
    strSQL = "Select 取消人,病例ID,备注,ID,病人ID,主页ID,医嘱ID,记录来源,诊断次序,编码序号,病历ID,诊断类型,疾病ID,诊断ID,证候ID,诊断描述,出院情况,是否未治,是否疑诊,记录日期,记录人,取消时间 From 病人诊断记录" & _
        " Where 记录来源 IN(1,2,3) And 诊断类型 IN(1,2,3,5,6,7)" & _
        " And 取消时间 is Null And 病人ID=[1] And 主页ID=[2]" & _
        " Order by 诊断类型,记录来源 Desc,诊断次序,ID"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人诊断记录", "H病人诊断记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Not rsTmp.EOF Then
        With vsDiagXY
            strSQL = "1,2,3,5,6,7"
            For i = 0 To UBound(Split(strSQL, ","))
                rsTmp.Filter = "记录来源=3 And 诊断类型=" & Split(strSQL, ",")(i)
                If rsTmp.EOF Then
                    rsTmp.Filter = "记录来源=2 And 诊断类型=" & Split(strSQL, ",")(i)
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "记录来源=1 And 诊断类型=" & Split(strSQL, ",")(i)
                End If
                Do While Not rsTmp.EOF
                    '1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码
                    lngRow = Decode(NVL(rsTmp!诊断类型, 0), 1, 1, 2, 2, 5, .Rows - 3, 6, .Rows - 2, 7, .Rows - 1)
                    
                    '出院诊断(主要诊断)可能有多个
                    If NVL(rsTmp!诊断类型, 0) = 3 Then
                        If .TextMatrix(3, col诊断描述) = "" Then
                            lngRow = 3
                        Else
                            .AddItem "", .Rows - 3
                            lngRow = .Rows - 4
                        End If
                    End If
                    
                    .TextMatrix(lngRow, col诊断描述) = NVL(rsTmp!诊断描述)
                    .TextMatrix(lngRow, col出院情况) = NVL(rsTmp!出院情况)
                    .TextMatrix(lngRow, col是否未治) = IIf(NVL(rsTmp!是否未治, 0) = 1, "√", "")
                    .TextMatrix(lngRow, col是否疑诊) = IIf(NVL(rsTmp!是否疑诊, 0) = 1, "？", "")
                    rsTmp.MoveNext
                Loop
            Next
        End With
    End If
    vsDiagXY.Cell(flexcpForeColor, 1, col是否疑诊, vsDiagXY.Rows - 1, col是否疑诊) = vbRed
    vsDiagXY.Cell(flexcpBackColor, 3, col诊断描述, 3, col是否疑诊) = &HC0FFC0
    vsDiagXY.Row = 1: vsDiagXY.Col = col诊断描述
        
    '中医诊断
    '---------------------------------------------------------------
    strSQL = "Select 取消人,病例ID,备注,ID,病人ID,主页ID,医嘱ID,记录来源,诊断次序,编码序号,病历ID,诊断类型,疾病ID,诊断ID,证候ID,诊断描述,出院情况,是否未治,是否疑诊,记录日期,记录人,取消时间 From 病人诊断记录" & _
        " Where 记录来源 IN(1,2,3) And 诊断类型 IN(11,12,13)" & _
        " And 取消时间 Is Null And 病人ID=[1] And 主页ID=[2]" & _
        " Order by 诊断类型,记录来源 Desc,诊断次序,ID"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人诊断记录", "H病人诊断记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Not rsTmp.EOF Then
        With vsDiagZY
            strSQL = "11,12,13"
            For i = 0 To UBound(Split(strSQL, ","))
                rsTmp.Filter = "记录来源=3 And 诊断类型=" & Split(strSQL, ",")(i)
                If rsTmp.EOF Then
                    rsTmp.Filter = "记录来源=2 And 诊断类型=" & Split(strSQL, ",")(i)
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "记录来源=1 And 诊断类型=" & Split(strSQL, ",")(i)
                End If
                Do While Not rsTmp.EOF
                    '11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断
                    lngRow = Decode(NVL(rsTmp!诊断类型, 0), 11, 1, 12, 2)
                    
                    '出院诊断(主要诊断)可能有多个
                    If NVL(rsTmp!诊断类型, 0) = 13 Then
                        For j = 3 To .Rows - 1
                            If .TextMatrix(j, col诊断描述) = "" Then
                                lngRow = j: Exit For
                            End If
                        Next
                        If j > .Rows - 1 Then
                            .AddItem "": lngRow = .Rows - 1
                        End If
                    End If
                    
                    .TextMatrix(lngRow, col诊断描述) = NVL(rsTmp!诊断描述)
                    .TextMatrix(lngRow, col出院情况) = NVL(rsTmp!出院情况)
                    rsTmp.MoveNext
                Loop
            Next
        End With
    End If
    vsDiagZY.Cell(flexcpBackColor, 3, col诊断描述, 3, col出院情况) = &HC0FFC0
    vsDiagZY.Row = 1: vsDiagZY.Col = col诊断描述
    
    '手术情况
    '---------------------------------------------------------------
    strSQL = "Select 记录来源,手术日期,手术开始时间,手术结束时间,拟行手术,手术操作ID,诊疗项目ID,已行手术,主刀医师,第一助手,第二助手,手术护士,麻醉开始时间,麻醉结束时间,麻醉方式,麻醉类型,麻醉质量,输液总量,麻醉医师,输氧开始时间,输氧结束时间,切口,愈合,记录日期,记录人,取消时间,取消人,助产护士,ID,病人ID,主页ID From 病人手麻记录" & _
        " Where 病人ID=[1] And 主页ID=[2]" & _
        " And 记录来源=3 Order by ID"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人手麻记录", "H病人手麻记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Not rsTmp.EOF Then
        With vsOPS
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col手术日期) = Format(NVL(rsTmp!手术日期), "yyyy-MM-dd")
                .TextMatrix(i, col手术名称) = NVL(rsTmp!已行手术)
                .TextMatrix(i, col主刀医师) = NVL(rsTmp!主刀医师)
                .TextMatrix(i, col助手1) = NVL(rsTmp!第一助手)
                .TextMatrix(i, col助手2) = NVL(rsTmp!第二助手)
                .TextMatrix(i, col麻醉方式) = GetItemField("诊疗项目目录", Val(NVL(rsTmp!麻醉方式, 0)), "名称")
                .TextMatrix(i, col麻醉医师) = NVL(rsTmp!麻醉医师)
                If Not IsNull(rsTmp!切口) And Not IsNull(rsTmp!愈合) Then
                    .TextMatrix(i, col切口愈合) = rsTmp!切口 & "/" & rsTmp!愈合
                End If
                rsTmp.MoveNext
            Next
        End With
    End If
    
    mblnCheck = False
    Screen.MousePointer = 0
    LoadPageData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsDiagXY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsDiagXY.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsDiagZY.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsOPS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsOPS.ShowCell(NewRow, NewCol)
End Sub




