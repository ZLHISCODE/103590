VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCISViewMedrec 
   Caption         =   "������ҳ����"
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
   StartUpPosition =   3  '����ȱʡ
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
            Caption         =   "   ������Ϣ "
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
               Caption         =   "����Ժ"
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
               Caption         =   "סԺ��"
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
               Caption         =   "��    ��סԺ"
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
               Caption         =   "���ʽ"
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
               Caption         =   "����"
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
               Caption         =   "�Ա�"
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
               Caption         =   "��������"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
               Caption         =   "ְҵ"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
               Caption         =   "�����ص�"
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
               Caption         =   "���֤��"
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
               Caption         =   "��ͥ��ַ"
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
               Caption         =   "�绰"
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
               Caption         =   "�ʱ�"
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
               Caption         =   "������λ"
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
               Caption         =   "�绰"
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
               Caption         =   "�ʱ�"
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
               Caption         =   "��ϵ������"
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
               Caption         =   "��ϵ"
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
               Caption         =   "�绰"
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
               Caption         =   "��ϵ�˵�ַ"
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
               Caption         =   "��Ժʱ��"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
               Caption         =   "��Ժʱ��"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
               Caption         =   "ת�����"
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
               Caption         =   "��������"
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
               Caption         =   "��������"
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
               Caption         =   "סԺ����"
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
            Caption         =   "   סԺ��� "
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
               Caption         =   "���в���"
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
               Caption         =   "�·�����"
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
               Caption         =   "ʬ��"
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
               Caption         =   "����"
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
               Caption         =   "ʾ�̲���"
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
               Caption         =   "��Ѫ��Ӧ"
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
               Caption         =   "����ԭ��"
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
               Caption         =   "��Һ��Ӧ"
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
               Caption         =   "��������"
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
               Caption         =   "Ѫ��"
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
               Caption         =   "���ϸ��"
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
               Caption         =   "��λ"
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
               Caption         =   "��ѪС��"
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
               Caption         =   "��"
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
               Caption         =   "��Ѫ��"
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
               Caption         =   "��ȫѪ"
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
               Caption         =   "������"
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
               Caption         =   "����ҽʦ"
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
               Caption         =   "������"
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
               Caption         =   "����(������)ҽʦ"
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
               Caption         =   "����ҽʦ"
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
               Caption         =   "סԺҽʦ"
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
               Caption         =   "����ҽʦ"
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
               Caption         =   "�о���ҽʦ"
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
               Caption         =   "ʵϰҽʦ"
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
            Caption         =   "   ���������� "
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
               Caption         =   "���������ơ���顢���Ϊ��Ժ��һ��"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
            Caption         =   "   ��ҽ��� "
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
               Caption         =   " ׼ȷ�� "
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
                  Caption         =   "��֤"
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
                  Caption         =   "�η�"
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
                  Caption         =   "��ҩ"
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
               Caption         =   " סԺ�ڼ䲡�� "
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
                  Caption         =   "Σ��"
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
                  Caption         =   "��֢"
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
                  Caption         =   "����"
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
               Caption         =   " ���Ʒ��� "
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
                  Caption         =   "������ҩ�Ƽ�"
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
                  Caption         =   "���ȷ���"
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
                  Caption         =   "�������"
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
                  Name            =   "����"
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
               Caption         =   "�������Ժ"
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
               Caption         =   "��Ժ���Ժ"
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
            Caption         =   "   ��ҽ��� "
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
               Caption         =   "�Ƿ�ȷ��"
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
                  Name            =   "����"
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
               Caption         =   "��ǰ������"
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
               Caption         =   "�ٴ���ʬ��"
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
               Caption         =   "�ٴ��벡��"
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
               Caption         =   "�����벡��"
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
               Caption         =   "��Ժ���Ժ"
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
               Caption         =   "�������Ժ"
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
               Caption         =   "ȷ������"
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
               Caption         =   "���ȴ���"
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
               Caption         =   "�ɹ�����"
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

'�ϴ�ˢ������ʱ�Ĳ�����Ϣ
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng����ID As Long
Private mlng����ID As Long
Private mbln��Ժ As Boolean
Private mblnMoved As Boolean
Private mbln��ҽ As Boolean
Private mblnCheck As Boolean
Private mblnOpenForm As Boolean
Private Enum ������Ϣ
    txt���ʽ = 26
    txt�Ա� = 27
    txt���� = 28
    txtְҵ = 31
    txt��Ժ���� = 34
    txt���� = 32
    txt���� = 30
    txt���� = 29
    txt��ϵ�˹�ϵ = 33
    txtסԺ�� = 0
    txt��ͥ��ַ = 1
    txtסԺ���� = 2
    txt���� = 3
    txt�������� = 4
    txt���� = 5
    txt�����ص� = 6
    txt���֤�� = 7
    txt��ͥ�绰 = 8
    txt��ͥ�ʱ� = 9
    txt������λ = 10
    txt��λ�绰 = 11
    txt��λ�ʱ� = 12
    txt��ϵ������ = 13
    txt��ϵ�˵绰 = 14
    txt��ϵ�˵�ַ = 15
    txt��Ժʱ�� = 16
    txt��Ժ���� = 17
    txt��Ժ���� = 18
    txt��Ժʱ�� = 19
    txt��Ժ���� = 20
    txt��Ժ���� = 21
    txtסԺ���� = 22
    txtת��1 = 23
    txtת��2 = 24
    txtת��3 = 25
    chk����Ժ = 10
End Enum
Private Enum ��ҽ���
    chk�Ƿ�ȷ�� = 0
    txt���ȴ��� = 36
    txtȷ������ = 37
    txt�ɹ����� = 35
    txt�������Ժ = 70
    txt��Ժ���Ժ = 68
    txt�����벡�� = 66
    txt�ٴ��벡�� = 69
    txt�ٴ���ʬ�� = 67
    txt��ǰ������ = 65
End Enum
Private Enum ��ҽ���
    chkΣ�� = 2
    chk��֢ = 3
    chk���� = 4
    txt��֤ = 57
    txtHBsAg = 75
    txtHCVAb = 74
    txtHIVAb = 73
    txt�η� = 58
    txt��ҩ = 59
    txt������ҩ = 40
    txt���ȷ��� = 39
    txt������� = 38
    txt��ҽ�������Ժ = 71
    txt��ҽ��Ժ���Ժ = 72
End Enum
Private Enum ����������
    chk���� = 1
End Enum
Private Enum סԺ���
    chkʬ�� = 6
    chk���� = 7
    chk�·����� = 5
    chkʾ�̲��� = 8
    chk���в��� = 11
    chk��Ѫ��Ӧ = 9
    txt����ԭ�� = 64
    txt�������� = 46
    txt���ϸ�� = 45
    txt��ѪС�� = 44
    txt��Ѫ�� = 43
    txt��ȫѪ = 42
    txt������ = 41
    txtѪ�� = 47
    txtRh = 48
    txt��Һ��Ӧ = 63
    txt����ҽʦ = 49
    txt������ = 54
    txt����ҽʦ = 50
    txt����ҽʦ = 52
    txtסԺҽʦ = 55
    txt����ҽʦ = 51
    txt�о���ҽʦ = 53
    txtʵϰҽʦ = 56
End Enum
Private Enum ������
    col������� = 0
    col������� = 1
    col��Ժ��� = 2
    col�Ƿ�δ�� = 3
    col�Ƿ����� = 4
End Enum
Private Enum �������
    col�������� = 0
    col�������� = 1
    col����ҽʦ = 2
    col����1 = 3
    col����2 = 4
    col����ʽ = 5
    col����ҽʦ = 6
    col�п����� = 7
End Enum

Private Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'���ܣ���ʾ������һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
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

Public Function zlRefresh(ByVal frmMain As Object, lng����ID As Long, lng��ҳID As Long, lng����ID As Long, lng����ID As Long, Optional bln��Ժ As Boolean) As Boolean
'���ܣ�ˢ�»����ҽ���嵥
    
    mlng����ID = lng����ID: mlng��ҳID = lng��ҳID
    mlng����ID = lng����ID: mlng����ID = lng����ID
    mbln��Ժ = bln��Ժ: mblnMoved = False
        
    '���ܿ����л�
    mbln��ҽ = Have��������(mlng����ID, "��ҽ��")
    fraInfo(2).Visible = mbln��ҽ
    fraInfo(2).Enabled = mbln��ҽ '��־������
    Call SetPageHeight
    Call SetScrollbar
    
    Call ClearPageData
    If mlng����ID <> 0 Then Call LoadPageData

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
    '�������ߴ�
    vsc.Width = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    hsc.Height = GetSystemMetrics(SM_CXHSCROLL) * Screen.TwipsPerPixelY
    fraVH.Width = vsc.Width: fraVH.Height = hsc.Height
    fraBack.Left = 0: fraBack.Top = 0
    picBack.BackColor = fraBack.BackColor
    fraInfo(1).Left = fraInfo(0).Left
'    '��ʼ��ϵͳ����
'    Call InitSysPar
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub SetPageHeight()
'���ܣ�����ҳ��������չ��״̬���ý���ߴ�
'˵����Tag=1��ʾ����
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
'���ܣ����ݵ�ǰ����ߴ����ù������ɼ��Լ��������
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
'���ܣ������ҳ�е�����
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
'���ܣ���ȡ���˵���ҳ��Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim lngRow As Long, varTmp As Variant
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    mblnCheck = True
    
    '������Ϣ����
    '---------------------------------------------------------------
    strSQL = "Select סԺ��,�Ա�,����,��������,�����ص�,���֤��,���� From ������Ϣ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    
    txtInfo(txtסԺ��).Text = NVL(rsTmp!סԺ��)
    txtInfo(txtסԺ����).Text = mlng��ҳID
    txtInfo(txt����).Text = NVL(rsTmp!����)
    txtInfo(txt�Ա�).Text = NVL(rsTmp!�Ա�)
    
    If Format(rsTmp!��������, "HH:mm") <> "00:00" Then
        txtInfo(txt��������).Text = Format(rsTmp!��������, "yyyy-MM-dd HH:mm")
    Else
        txtInfo(txt��������).Text = Format(rsTmp!��������, "yyyy-MM-dd")
    End If
    
    txtInfo(txt�����ص�).Text = NVL(rsTmp!�����ص�)
    txtInfo(txt���֤��).Text = NVL(rsTmp!���֤��)
    txtInfo(txt����).Text = NVL(rsTmp!����)
    
    '������ҳ����
    '---------------------------------------------------------------
    strSQL = "Select A.����ID,A.��ҳID,A.סԺ��,A.��������,A.ҽ�Ƹ��ʽ,A.�ѱ�,A.����Ժ,A.��Ժ����ID,A.��Ժ����ID,A.��Ժ����,A.��Ժ����,A.��Ժ��ʽ,A.��Ժ����,A.����Ժת��,A.סԺĿ��,A.��Ժ����,A.�Ƿ����,A.��ǰ����,A.��ǰ����ID,A.����ȼ�ID,A.��Ժ����ID,A.��Ժ����,A.��Ժ����,A.סԺ����,A.��Ժ��ʽ,A.�Ƿ�ȷ��,A.ȷ������,A.�·�����,A.Ѫ��,A.���ȴ���,A.�ɹ�����,A.�����־,A.��������,A.ʬ���־,A.����ҽʦ,A.���λ�ʿ,A.סԺҽʦ,A.������,A.��ĿԱ���,A.��ĿԱ����,A.��Ŀ����,A.״̬,A.���ú�,A.����,A.����״��,A.ְҵ,A.����,A.ѧ��,A.��λ�绰,A.��λ�ʱ�,A.��λ��ַ,A.����,A.��ͥ��ַ,A.��ͥ�绰,A.��ͥ��ַ�ʱ�,A.��ϵ������,A.��ϵ�˹�ϵ,A.��ϵ�˵�ַ,A.��ϵ�˵绰,A.��ҽ�������,A.����,A.��˱�־,A.�����,A.�������,A.�Ƿ��ϴ�,A.����ת��,A.�Ǽ���,A.�Ǽ�ʱ��,A.��ע,A.����,A.����״̬,A.���ʱ��,A.��������,B.���� as ��Ժ����,C.���� as ��Ժ����" & _
        " From ������ҳ A,���ű� B,���ű� C" & _
        " Where A.��Ժ����ID=B.ID And A.��Ժ����ID=C.ID" & _
        " And A.����ID=[1] And A.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    mblnMoved = NVL(rsTmp!����ת��, 0) = 1
    
    '���۲�����סԺ��
    If NVL(rsTmp!��������, 0) <> 0 Then
        lblInfo(0).Visible = False
        txtInfo(txtסԺ��).Visible = False
    End If
        
    txtInfo(txt���ʽ).Text = NVL(rsTmp!ҽ�Ƹ��ʽ)
    txtInfo(txt����).Text = NVL(rsTmp!����)
    txtInfo(txt����).Text = NVL(rsTmp!����״��)
    txtInfo(txtְҵ).Text = NVL(rsTmp!ְҵ)
    txtInfo(txt����).Text = NVL(rsTmp!����)
    txtInfo(txt����).Text = NVL(rsTmp!����)
    txtInfo(txt��ͥ��ַ).Text = NVL(rsTmp!��ͥ��ַ)
    txtInfo(txt��ͥ�绰).Text = NVL(rsTmp!��ͥ�绰)
    txtInfo(txt��ͥ�ʱ�).Text = NVL(rsTmp!��ͥ��ַ�ʱ�)
    txtInfo(txt������λ).Text = NVL(rsTmp!��λ��ַ)
    txtInfo(txt��λ�绰).Text = NVL(rsTmp!��λ�绰)
    txtInfo(txt��λ�ʱ�).Text = NVL(rsTmp!��λ�ʱ�)
    txtInfo(txt��ϵ������).Text = NVL(rsTmp!��ϵ������)
    txtInfo(txt��ϵ�˹�ϵ).Text = NVL(rsTmp!��ϵ�˹�ϵ)
    txtInfo(txt��ϵ�˵绰).Text = NVL(rsTmp!��ϵ�˵绰)
    txtInfo(txt��ϵ�˵�ַ).Text = NVL(rsTmp!��ϵ�˵�ַ)
    chkInfo(chk����Ժ).Value = NVL(rsTmp!����Ժ, 0)
    txtInfo(txt��Ժʱ��).Text = Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm")
    txtInfo(txt��Ժ����).Text = rsTmp!��Ժ����
    txtInfo(txt��Ժ����).Text = NVL(rsTmp!��Ժ����)
    txtInfo(txt��Ժʱ��).Text = Format(NVL(rsTmp!��Ժ����), "yyyy-MM-dd HH:mm")
    txtInfo(txt��Ժ����).Text = rsTmp!��Ժ����
    If Not IsNull(rsTmp!��Ժ����) Then
        txtInfo(txtסԺ����).Text = DateDiff("d", rsTmp!��Ժ����, rsTmp!��Ժ����)
    Else
        txtInfo(txtסԺ����).Text = DateDiff("d", rsTmp!��Ժ����, zlDatabase.Currentdate)
    End If
    If Val(txtInfo(txtסԺ����).Text) = 0 Then txtInfo(txtסԺ����).Text = "1"
    chkInfo(chk�Ƿ�ȷ��).Value = NVL(rsTmp!�Ƿ�ȷ��, 0)
    If chkInfo(chk�Ƿ�ȷ��).Value = 1 Then
        txtInfo(txtȷ������).Text = Format(NVL(rsTmp!ȷ������), "yyyy-MM-dd HH:mm")
    End If
    txtInfo(txt���ȴ���).Text = NVL(rsTmp!���ȴ���)
    If Val(txtInfo(txt���ȴ���).Text) <> 0 Then
        txtInfo(txt�ɹ�����).Text = NVL(rsTmp!�ɹ�����)
    End If
    chkInfo(chk�·�����).Value = NVL(rsTmp!�·�����, 0)
    
    txtInfo(txt�������).Text = NVL(rsTmp!��ҽ�������)
    chkInfo(chkʬ��).Value = NVL(rsTmp!ʬ���־, 0)
    chkInfo(chk����).Value = IIf(NVL(rsTmp!�����־, 0) = 0, 0, 1)
    If chkInfo(chk����).Value = 1 Then
        txtInfo(txt��������).Text = NVL(rsTmp!��������, 0) & Decode(NVL(rsTmp!�����־, 0), 1, "��", 2, "��", 3, "��")
    End If
    txtInfo(txt����ҽʦ).Text = NVL(rsTmp!����ҽʦ)
    txtInfo(txtסԺҽʦ).Text = NVL(rsTmp!סԺҽʦ)
    txtInfo(txtѪ��).Text = NVL(rsTmp!Ѫ��)
    
    '�����ӱ���
    '---------------------------------------------------------------
    strSQL = "Select ����ID,��ҳID,��Ϣ��,��Ϣֵ From ������ҳ�ӱ� Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    For i = 1 To rsTmp.RecordCount
        Select Case UCase(NVL(rsTmp!��Ϣ��))
            Case "��Ժ����"
                txtInfo(txt��Ժ����).Text = NVL(rsTmp!��Ϣֵ)
            Case "��Ժ����"
                txtInfo(txt��Ժ����).Text = NVL(rsTmp!��Ϣֵ)
            Case "ת�Ƽ�¼"
                varTmp = Split(NVL(rsTmp!��Ϣֵ), ",")
                If UBound(varTmp) >= 0 Then txtInfo(txtת��1).Text = varTmp(0)
                If UBound(varTmp) >= 1 Then txtInfo(txtת��2).Text = varTmp(1)
                If UBound(varTmp) >= 2 Then txtInfo(txtת��3).Text = varTmp(2)
            Case "��������ԭ��"
                txtInfo(txt����ԭ��).Text = NVL(rsTmp!��Ϣֵ)
            Case UCase("HBsAg")
                txtInfo(txtHBsAg).Text = NVL(rsTmp!��Ϣֵ)
            Case UCase("HCV-Ab")
                txtInfo(txtHCVAb).Text = NVL(rsTmp!��Ϣֵ)
            Case UCase("HIV-Ab")
                txtInfo(txtHIVAb).Text = NVL(rsTmp!��Ϣֵ)
            Case "����"
                chkInfo(chk����).Value = Val(NVL(rsTmp!��Ϣֵ, 0))
            Case "��ҽΣ��"
                chkInfo(chkΣ��).Value = Val(NVL(rsTmp!��Ϣֵ, 0))
            Case "��ҽ��֢"
                chkInfo(chk��֢).Value = Val(NVL(rsTmp!��Ϣֵ, 0))
            Case "��ҽ����"
                chkInfo(chk����).Value = Val(NVL(rsTmp!��Ϣֵ, 0))
            Case "��ҽ���ȷ���"
                txtInfo(txt���ȷ���).Text = NVL(rsTmp!��Ϣֵ)
            Case "������ҩ�Ƽ�"
                txtInfo(txt������ҩ).Text = NVL(rsTmp!��Ϣֵ)
            Case "ʾ�̲���"
                chkInfo(chkʾ�̲���).Value = Val(NVL(rsTmp!��Ϣֵ, 0))
            Case "���в���"
                chkInfo(chk���в���).Value = Val(NVL(rsTmp!��Ϣֵ, 0))
            Case UCase("Rh")
                txtInfo(txtRh).Text = NVL(rsTmp!��Ϣֵ)
            Case "��Ѫ��Ӧ"
                chkInfo(chk��Ѫ��Ӧ).Value = Val(NVL(rsTmp!��Ϣֵ, 0))
            Case "���ϸ��"
                txtInfo(txt���ϸ��).Text = NVL(rsTmp!��Ϣֵ)
            Case "��ѪС��"
                txtInfo(txt��ѪС��).Text = NVL(rsTmp!��Ϣֵ)
            Case "��Ѫ��"
                txtInfo(txt��Ѫ��).Text = NVL(rsTmp!��Ϣֵ)
            Case "��ȫѪ"
                txtInfo(txt��ȫѪ).Text = NVL(rsTmp!��Ϣֵ)
            Case "������"
                txtInfo(txt������).Text = NVL(rsTmp!��Ϣֵ)
            Case "��Һ��Ӧ"
                txtInfo(txt��Һ��Ӧ).Text = NVL(rsTmp!��Ϣֵ)
            Case "������"
                txtInfo(txt������).Text = NVL(rsTmp!��Ϣֵ)
            Case "����ҽʦ"
                txtInfo(txt����ҽʦ).Text = NVL(rsTmp!��Ϣֵ)
            Case "����ҽʦ"
                txtInfo(txt����ҽʦ).Text = NVL(rsTmp!��Ϣֵ)
            Case "����ҽʦ"
                txtInfo(txt����ҽʦ).Text = NVL(rsTmp!��Ϣֵ)
            Case "�о���ʵϰҽʦ"
                txtInfo(txt�о���ҽʦ).Text = NVL(rsTmp!��Ϣֵ)
            Case "ʵϰҽʦ"
                txtInfo(txtʵϰҽʦ).Text = NVL(rsTmp!��Ϣֵ)
        End Select
        rsTmp.MoveNext
    Next
    
    '��Ϸ������
    '---------------------------------------------------------------
    strSQL = "Select ��������,������� From ��Ϸ������ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    Do While Not rsTmp.EOF
        Select Case rsTmp!��������
        Case 1 '�������Ժ
            txtInfo(txt�������Ժ).Text = Decode(NVL(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 2 '��Ժ���Ժ
            txtInfo(txt��Ժ���Ժ).Text = Decode(NVL(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 3 '�����벡��
            txtInfo(txt�����벡��).Text = Decode(NVL(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 4 '�ٴ��벡��
            txtInfo(txt�ٴ��벡��).Text = Decode(NVL(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 5 '�ٴ���ʬ��
            txtInfo(txt�ٴ���ʬ��).Text = Decode(NVL(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 6 '��ǰ������
            txtInfo(txt��ǰ������).Text = Decode(NVL(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 11 '��ҽ�������Ժ
            txtInfo(txt��ҽ�������Ժ).Text = Decode(NVL(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 12 '��ҽ��Ժ���Ժ
            txtInfo(txt��ҽ��Ժ���Ժ).Text = Decode(NVL(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 13 '��ҽ��֤
            txtInfo(txt��֤).Text = Decode(NVL(rsTmp!�������, 0), 1, "׼ȷ", 2, "����׼ȷ", 3, "�ش�ȱ��", 4, "����", "")
        Case 14 '��ҽ�η�
            txtInfo(txt�η�).Text = Decode(NVL(rsTmp!�������, 0), 1, "׼ȷ", 2, "����׼ȷ", 3, "�ش�ȱ��", 4, "����", "")
        Case 15 '��ҽ��ҩ
            txtInfo(txt��ҩ).Text = Decode(NVL(rsTmp!�������, 0), 1, "׼ȷ", 2, "����׼ȷ", 3, "�ش�ȱ��", 4, "����", "")
        End Select
        rsTmp.MoveNext
    Loop
    
    '�Զ���ȡת�ƿ��Ҽ��������(�����)
    '---------------------------------------------------------------
    If txtInfo(txtת��1).Text = "" And txtInfo(txtת��2).Text = "" And txtInfo(txtת��3).Text = "" Then
        strSQL = _
            " Select B.����" & _
            " From ���˱䶯��¼ A,���ű� B" & _
            " Where A.����ID=[1] And A.��ҳID=[2]" & _
            " And A.����ID=B.ID And A.��ʼԭ��=3 And A.��ʼʱ�� is Not NULL" & _
            " Order by A.��ʼʱ��"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
        For i = 1 To rsTmp.RecordCount
            If i = 1 Then
                txtInfo(txtת��1).Text = rsTmp!����
            ElseIf i = 2 Then
                txtInfo(txtת��2).Text = rsTmp!����
            ElseIf i = 3 Then
                txtInfo(txtת��3).Text = rsTmp!����
                Exit For
            End If
            rsTmp.MoveNext
        Next
    End If
    
    If txtInfo(txt��Ժ����).Text = "" Then
        strSQL = "Select B.�����" & _
            " From ������ҳ A,��λ״����¼ B" & _
            " Where A.����ID=[1] And A.��ҳID=[2]" & _
            " And A.��Ժ����ID=B.����ID And A.��Ժ����=B.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If Not rsTmp.EOF Then txtInfo(txt��Ժ����).Text = NVL(rsTmp!�����)
    End If
    If txtInfo(txt��Ժ����).Text = "" Then
        strSQL = "Select B.�����" & _
            " From ������ҳ A,��λ״����¼ B" & _
            " Where A.����ID=[1] And A.��ҳID=[2]" & _
            " And A.��ǰ����ID=B.����ID And A.��Ժ����=B.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If Not rsTmp.EOF Then txtInfo(txt��Ժ����).Text = NVL(rsTmp!�����)
    End If
    
    '������Ϣ:����סԺ��,������
    '---------------------------------------------------------------
    strSQL = "Select ��¼��Դ,Decode(����ʱ��,Null ,��¼ʱ��,����ʱ��) as ����ʱ��,ҩ��ID,ҩ���� From ���˹�����¼ A" & _
        " Where ���=1 And ����ID=[1] And ��ҳID=[2]" & _
        " And Not Exists(Select ҩ��ID From ���˹�����¼" & _
            " Where (Nvl(ҩ��ID,0)=Nvl(A.ҩ��ID,0) Or Nvl(ҩ����,'Null')=Nvl(A.ҩ����,'Null'))" & _
            " And Nvl(���,0)=0 And ��¼ʱ��>=A.��¼ʱ�� And ����ID=[1] And ��ҳID=[2])" & _
        " Order by ����ʱ��,ҩ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then
        rsTmp.Filter = "��¼��Դ=3" '��ҳ������д��
        If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ<>3" '������Դ����Ϊȱʡ��ʾ
        With vsAller
            .Rows = rsTmp.RecordCount + 1 '�̶���
            For i = 1 To rsTmp.RecordCount
                '������Դ�Ŀ������ظ�
                lngRow = -1
                If Not IsNull(rsTmp!ҩ��ID) Then
                    lngRow = .FindRow(CLng(rsTmp!ҩ��ID))
                ElseIf Not IsNull(rsTmp!ҩ����) Then
                    lngRow = .FindRow(CStr(rsTmp!ҩ����), , 1)
                End If
                If lngRow = -1 Then
                    .RowData(i) = CLng(NVL(rsTmp!ҩ��ID, 0))
                    .TextMatrix(i, 0) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, i, 0) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss") '���ڱ���
                    .TextMatrix(i, 1) = NVL(rsTmp!ҩ����)
                    .Cell(flexcpData, i, 1) = .TextMatrix(i, 1) '��������ָ�
                End If
                rsTmp.MoveNext
            Next
            If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        End With
    End If
    vsAller.Row = 1: vsAller.Col = 1
    
    '��ҽ���
    '---------------------------------------------------------------
    strSQL = "Select ȡ����,����ID,��ע,ID,����ID,��ҳID,ҽ��ID,��¼��Դ,��ϴ���,�������,����ID,�������,����ID,���ID,֤��ID,�������,��Ժ���,�Ƿ�δ��,�Ƿ�����,��¼����,��¼��,ȡ��ʱ�� From ������ϼ�¼" & _
        " Where ��¼��Դ IN(1,2,3) And ������� IN(1,2,3,5,6,7)" & _
        " And ȡ��ʱ�� is Null And ����ID=[1] And ��ҳID=[2]" & _
        " Order by �������,��¼��Դ Desc,��ϴ���,ID"
    If mblnMoved Then
        strSQL = Replace(strSQL, "������ϼ�¼", "H������ϼ�¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then
        With vsDiagXY
            strSQL = "1,2,3,5,6,7"
            For i = 0 To UBound(Split(strSQL, ","))
                rsTmp.Filter = "��¼��Դ=3 And �������=" & Split(strSQL, ",")(i)
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=2 And �������=" & Split(strSQL, ",")(i)
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=1 And �������=" & Split(strSQL, ",")(i)
                End If
                Do While Not rsTmp.EOF
                    '1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���
                    lngRow = Decode(NVL(rsTmp!�������, 0), 1, 1, 2, 2, 5, .Rows - 3, 6, .Rows - 2, 7, .Rows - 1)
                    
                    '��Ժ���(��Ҫ���)�����ж��
                    If NVL(rsTmp!�������, 0) = 3 Then
                        If .TextMatrix(3, col�������) = "" Then
                            lngRow = 3
                        Else
                            .AddItem "", .Rows - 3
                            lngRow = .Rows - 4
                        End If
                    End If
                    
                    .TextMatrix(lngRow, col�������) = NVL(rsTmp!�������)
                    .TextMatrix(lngRow, col��Ժ���) = NVL(rsTmp!��Ժ���)
                    .TextMatrix(lngRow, col�Ƿ�δ��) = IIf(NVL(rsTmp!�Ƿ�δ��, 0) = 1, "��", "")
                    .TextMatrix(lngRow, col�Ƿ�����) = IIf(NVL(rsTmp!�Ƿ�����, 0) = 1, "��", "")
                    rsTmp.MoveNext
                Loop
            Next
        End With
    End If
    vsDiagXY.Cell(flexcpForeColor, 1, col�Ƿ�����, vsDiagXY.Rows - 1, col�Ƿ�����) = vbRed
    vsDiagXY.Cell(flexcpBackColor, 3, col�������, 3, col�Ƿ�����) = &HC0FFC0
    vsDiagXY.Row = 1: vsDiagXY.Col = col�������
        
    '��ҽ���
    '---------------------------------------------------------------
    strSQL = "Select ȡ����,����ID,��ע,ID,����ID,��ҳID,ҽ��ID,��¼��Դ,��ϴ���,�������,����ID,�������,����ID,���ID,֤��ID,�������,��Ժ���,�Ƿ�δ��,�Ƿ�����,��¼����,��¼��,ȡ��ʱ�� From ������ϼ�¼" & _
        " Where ��¼��Դ IN(1,2,3) And ������� IN(11,12,13)" & _
        " And ȡ��ʱ�� Is Null And ����ID=[1] And ��ҳID=[2]" & _
        " Order by �������,��¼��Դ Desc,��ϴ���,ID"
    If mblnMoved Then
        strSQL = Replace(strSQL, "������ϼ�¼", "H������ϼ�¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then
        With vsDiagZY
            strSQL = "11,12,13"
            For i = 0 To UBound(Split(strSQL, ","))
                rsTmp.Filter = "��¼��Դ=3 And �������=" & Split(strSQL, ",")(i)
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=2 And �������=" & Split(strSQL, ",")(i)
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=1 And �������=" & Split(strSQL, ",")(i)
                End If
                Do While Not rsTmp.EOF
                    '11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���
                    lngRow = Decode(NVL(rsTmp!�������, 0), 11, 1, 12, 2)
                    
                    '��Ժ���(��Ҫ���)�����ж��
                    If NVL(rsTmp!�������, 0) = 13 Then
                        For j = 3 To .Rows - 1
                            If .TextMatrix(j, col�������) = "" Then
                                lngRow = j: Exit For
                            End If
                        Next
                        If j > .Rows - 1 Then
                            .AddItem "": lngRow = .Rows - 1
                        End If
                    End If
                    
                    .TextMatrix(lngRow, col�������) = NVL(rsTmp!�������)
                    .TextMatrix(lngRow, col��Ժ���) = NVL(rsTmp!��Ժ���)
                    rsTmp.MoveNext
                Loop
            Next
        End With
    End If
    vsDiagZY.Cell(flexcpBackColor, 3, col�������, 3, col��Ժ���) = &HC0FFC0
    vsDiagZY.Row = 1: vsDiagZY.Col = col�������
    
    '�������
    '---------------------------------------------------------------
    strSQL = "Select ��¼��Դ,��������,������ʼʱ��,��������ʱ��,��������,��������ID,������ĿID,��������,����ҽʦ,��һ����,�ڶ�����,������ʿ,����ʼʱ��,�������ʱ��,����ʽ,��������,��������,��Һ����,����ҽʦ,������ʼʱ��,��������ʱ��,�п�,����,��¼����,��¼��,ȡ��ʱ��,ȡ����,������ʿ,ID,����ID,��ҳID From ���������¼" & _
        " Where ����ID=[1] And ��ҳID=[2]" & _
        " And ��¼��Դ=3 Order by ID"
    If mblnMoved Then
        strSQL = Replace(strSQL, "���������¼", "H���������¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then
        With vsOPS
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col��������) = Format(NVL(rsTmp!��������), "yyyy-MM-dd")
                .TextMatrix(i, col��������) = NVL(rsTmp!��������)
                .TextMatrix(i, col����ҽʦ) = NVL(rsTmp!����ҽʦ)
                .TextMatrix(i, col����1) = NVL(rsTmp!��һ����)
                .TextMatrix(i, col����2) = NVL(rsTmp!�ڶ�����)
                .TextMatrix(i, col����ʽ) = GetItemField("������ĿĿ¼", Val(NVL(rsTmp!����ʽ, 0)), "����")
                .TextMatrix(i, col����ҽʦ) = NVL(rsTmp!����ҽʦ)
                If Not IsNull(rsTmp!�п�) And Not IsNull(rsTmp!����) Then
                    .TextMatrix(i, col�п�����) = rsTmp!�п� & "/" & rsTmp!����
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




