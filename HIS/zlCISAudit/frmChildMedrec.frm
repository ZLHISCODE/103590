VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChildMedrec 
   BorderStyle     =   0  'None
   ClientHeight    =   8700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   255
      ScaleHeight     =   8295
      ScaleWidth      =   9735
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   9735
      Begin VB.Frame fraBack 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7545
         Left            =   195
         TabIndex        =   4
         Top             =   135
         Width           =   8715
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   基本信息 "
            ForeColor       =   &H00FF0000&
            Height          =   1335
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Tag             =   "6600"
            Top             =   30
            Width           =   7830
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "入院前经外院治疗"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   15
               Left            =   5265
               TabIndex        =   204
               Top             =   5610
               Width           =   1800
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   90
               Left            =   3465
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   202
               Top             =   5580
               Width           =   1410
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   89
               Left            =   1155
               Locked          =   -1  'True
               TabIndex        =   200
               Top             =   5580
               Width           =   1455
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   88
               Left            =   6105
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   198
               Top             =   2385
               Width           =   1500
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   87
               Left            =   6105
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   196
               Top             =   2055
               Width           =   1500
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   86
               Left            =   6105
               Locked          =   -1  'True
               TabIndex        =   194
               Top             =   4485
               Width           =   1500
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   85
               Left            =   6525
               Locked          =   -1  'True
               MaxLength       =   6
               TabIndex        =   192
               Top             =   3105
               Width           =   1095
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   84
               Left            =   1140
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   190
               Top             =   3090
               Width           =   4710
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   83
               Left            =   6945
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   189
               Top             =   1425
               Width           =   645
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   82
               Left            =   4005
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   187
               Top             =   1425
               Width           =   855
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   81
               Left            =   1170
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   185
               Top             =   1410
               Width           =   1410
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   80
               Left            =   4425
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   183
               Top             =   1110
               Width           =   450
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   79
               Left            =   3420
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   181
               Top             =   1110
               Width           =   495
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   78
               Left            =   4620
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   179
               Top             =   6300
               Width           =   2985
            End
            Begin VB.OptionButton optExits 
               BackColor       =   &H00FFFFFF&
               Caption         =   "有"
               ForeColor       =   &H00404040&
               Height          =   240
               Index           =   1
               Left            =   3360
               TabIndex        =   178
               Top             =   6300
               Width           =   525
            End
            Begin VB.OptionButton optExits 
               BackColor       =   &H00FFFFFF&
               Caption         =   "无"
               ForeColor       =   &H00404040&
               Height          =   240
               Index           =   0
               Left            =   2775
               TabIndex        =   177
               Top             =   6300
               Value           =   -1  'True
               Width           =   585
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   77
               Left            =   1125
               Locked          =   -1  'True
               TabIndex        =   175
               Top             =   6270
               Width           =   1455
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   76
               Left            =   6105
               Locked          =   -1  'True
               TabIndex        =   172
               Top             =   4845
               Width           =   1500
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
               TabIndex        =   42
               Top             =   345
               Width           =   1500
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
               TabIndex        =   41
               Top             =   5250
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
               TabIndex        =   40
               Top             =   5250
               Width           =   1530
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   23
               Left            =   1140
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   39
               Top             =   5250
               Width           =   1530
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
               TabIndex        =   38
               Top             =   5955
               Width           =   735
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
               TabIndex        =   37
               Top             =   5925
               Width           =   945
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   20
               Left            =   3270
               Locked          =   -1  'True
               TabIndex        =   36
               Top             =   5925
               Width           =   975
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   19
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   35
               Top             =   5925
               Width           =   1455
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   18
               Left            =   1125
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   34
               Top             =   4845
               Width           =   1455
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   17
               Left            =   3480
               Locked          =   -1  'True
               TabIndex        =   33
               Top             =   4485
               Width           =   1410
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
               TabIndex        =   32
               Top             =   4485
               Width           =   1455
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
               TabIndex        =   31
               Top             =   4050
               Width           =   4680
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
               TabIndex        =   30
               Top             =   3735
               Width           =   1185
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
               TabIndex        =   29
               Top             =   3735
               Width           =   1035
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
               TabIndex        =   28
               Top             =   3420
               Width           =   1095
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
               TabIndex        =   27
               Top             =   3420
               Width           =   1185
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
               TabIndex        =   26
               Top             =   3420
               Width           =   2805
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
               TabIndex        =   25
               Top             =   2805
               Width           =   1095
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
               TabIndex        =   24
               Top             =   2805
               Width           =   1185
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
               TabIndex        =   23
               Top             =   2805
               Width           =   2805
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   7
               Left            =   1170
               Locked          =   -1  'True
               MaxLength       =   18
               TabIndex        =   22
               Top             =   2385
               Width           =   2805
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
               TabIndex        =   21
               Top             =   2055
               Width           =   2805
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
               TabIndex        =   20
               Top             =   1095
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
               TabIndex        =   19
               Top             =   780
               Width           =   1500
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
               TabIndex        =   18
               Top             =   780
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
               TabIndex        =   17
               Top             =   345
               Width           =   285
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
               TabIndex        =   16
               Top             =   345
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
               TabIndex        =   15
               Top             =   780
               Width           =   1410
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   28
               Left            =   6105
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   14
               Top             =   1110
               Width           =   1500
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   29
               Left            =   3465
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   13
               Top             =   1740
               Width           =   1410
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
               TabIndex        =   12
               Top             =   1740
               Width           =   1410
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
               TabIndex        =   11
               Top             =   1740
               Width           =   1500
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   32
               Left            =   6525
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   10
               Top             =   4065
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
               TabIndex        =   9
               Top             =   3735
               Width           =   1095
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   34
               Left            =   3480
               Locked          =   -1  'True
               MaxLength       =   6
               TabIndex        =   8
               Top             =   4845
               Width           =   1410
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
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "再入院"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   10
               Left            =   4185
               TabIndex        =   6
               Top             =   338
               Width           =   900
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   87
               X1              =   3375
               X2              =   4875
               Y1              =   5760
               Y2              =   5760
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "转入"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   92
               Left            =   2970
               TabIndex        =   203
               Top             =   5580
               Width           =   360
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   86
               X1              =   1095
               X2              =   2615
               Y1              =   5760
               Y2              =   5760
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出院方式"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   91
               Left            =   345
               TabIndex        =   201
               Top             =   5580
               Width           =   720
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   85
               X1              =   6015
               X2              =   7605
               Y1              =   2565
               Y2              =   2565
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "其他证件"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   90
               Left            =   5250
               TabIndex        =   199
               Top             =   2385
               Width           =   720
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   84
               X1              =   6015
               X2              =   7605
               Y1              =   2235
               Y2              =   2235
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "籍贯"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   89
               Left            =   5610
               TabIndex        =   197
               Top             =   2055
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院途径"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   88
               Left            =   5250
               TabIndex        =   195
               Top             =   4485
               Width           =   720
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   83
               X1              =   6015
               X2              =   7605
               Y1              =   4665
               Y2              =   4665
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   82
               X1              =   6435
               X2              =   7620
               Y1              =   3285
               Y2              =   3285
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "邮编"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   87
               Left            =   6045
               TabIndex        =   193
               Top             =   3105
               Width           =   360
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   81
               X1              =   1080
               X2              =   5850
               Y1              =   3285
               Y2              =   3285
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "户口地址"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   86
               Left            =   330
               TabIndex        =   191
               Top             =   3105
               Width           =   720
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   80
               X1              =   6870
               X2              =   7600
               Y1              =   1620
               Y2              =   1620
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "新生儿入院体重"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   85
               Left            =   5595
               TabIndex        =   188
               Top             =   1425
               Width           =   1260
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   79
               X1              =   3915
               X2              =   4890
               Y1              =   1605
               Y2              =   1605
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "新生儿体重"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   84
               Left            =   2985
               TabIndex        =   186
               Top             =   1425
               Width           =   900
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   78
               X1              =   1080
               X2              =   2580
               Y1              =   1590
               Y2              =   1590
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "婴幼儿年龄"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   83
               Left            =   150
               TabIndex        =   184
               Top             =   1410
               Width           =   900
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   77
               X1              =   4395
               X2              =   4895
               Y1              =   1290
               Y2              =   1290
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "体重"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   82
               Left            =   4020
               TabIndex        =   182
               Top             =   1110
               Width           =   360
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   76
               X1              =   3390
               X2              =   3890
               Y1              =   1290
               Y2              =   1290
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "身高"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   81
               Left            =   2985
               TabIndex        =   180
               Top             =   1110
               Width           =   360
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   75
               X1              =   4545
               X2              =   7620
               Y1              =   6480
               Y2              =   6480
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "，目的："
               ForeColor       =   &H00404040&
               Height          =   240
               Index           =   80
               Left            =   3885
               TabIndex        =   176
               Top             =   6315
               Width           =   720
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   74
               X1              =   1080
               X2              =   2600
               Y1              =   6450
               Y2              =   6450
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "再入院计划"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   79
               Left            =   150
               TabIndex        =   174
               Top             =   6270
               Width           =   900
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   73
               X1              =   6015
               X2              =   7605
               Y1              =   5025
               Y2              =   5025
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入科时间"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   78
               Left            =   5250
               TabIndex        =   173
               Top             =   4845
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
               TabIndex        =   77
               Top             =   5925
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
               TabIndex        =   76
               Top             =   5250
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
               TabIndex        =   75
               Top             =   5250
               Width           =   720
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
               TabIndex        =   74
               Top             =   5250
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病房"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   31
               Left            =   4440
               TabIndex        =   73
               Top             =   5925
               Width           =   360
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
               TabIndex        =   72
               Top             =   5925
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
               TabIndex        =   71
               Top             =   5925
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病情"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   28
               Left            =   2985
               TabIndex        =   70
               Top             =   4845
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病房"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   27
               Left            =   675
               TabIndex        =   69
               Top             =   4845
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "科室"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   26
               Left            =   2985
               TabIndex        =   68
               Top             =   4485
               Width           =   360
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
               TabIndex        =   67
               Top             =   4485
               Width           =   720
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
               TabIndex        =   66
               Top             =   4050
               Width           =   900
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
               TabIndex        =   65
               Top             =   3735
               Width           =   360
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
               TabIndex        =   64
               Top             =   3735
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
               TabIndex        =   63
               Top             =   3735
               Width           =   900
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
               TabIndex        =   62
               Top             =   3420
               Width           =   360
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
               TabIndex        =   61
               Top             =   3420
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
               TabIndex        =   60
               Top             =   3420
               Width           =   720
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
               TabIndex        =   59
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
               Index           =   16
               Left            =   4185
               TabIndex        =   58
               Top             =   2805
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "现住址"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   1
               Left            =   495
               TabIndex        =   57
               Top             =   2805
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "身份证号"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   15
               Left            =   330
               TabIndex        =   56
               Top             =   2385
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
               TabIndex        =   55
               Top             =   2055
               Width           =   720
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
               TabIndex        =   54
               Top             =   1740
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
               TabIndex        =   53
               Top             =   1740
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
               Left            =   6045
               TabIndex        =   52
               Top             =   4065
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
               TabIndex        =   51
               Top             =   1740
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
               Left            =   5610
               TabIndex        =   50
               Top             =   1110
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
               Left            =   690
               TabIndex        =   49
               Top             =   1095
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
               TabIndex        =   48
               Top             =   780
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
               Left            =   2985
               TabIndex        =   47
               Top             =   780
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
               TabIndex        =   46
               Top             =   780
               Width           =   360
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
               TabIndex        =   45
               Top             =   345
               Width           =   720
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "第    次住院"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   3
               Left            =   2985
               TabIndex        =   44
               Top             =   345
               Width           =   1080
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
               TabIndex        =   43
               Top             =   345
               Width           =   540
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   0
               X1              =   1080
               X2              =   2580
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
               Index           =   2
               X1              =   6015
               X2              =   7605
               Y1              =   525
               Y2              =   525
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
               Index           =   4
               X1              =   1080
               X2              =   2580
               Y1              =   1275
               Y2              =   1275
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   5
               X1              =   1080
               X2              =   2580
               Y1              =   1920
               Y2              =   1920
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
               Index           =   7
               X1              =   6015
               X2              =   7605
               Y1              =   1290
               Y2              =   1290
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   8
               X1              =   3390
               X2              =   4890
               Y1              =   1920
               Y2              =   1920
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
               Index           =   10
               X1              =   6015
               X2              =   7605
               Y1              =   1920
               Y2              =   1920
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   11
               X1              =   6435
               X2              =   7620
               Y1              =   4245
               Y2              =   4245
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   12
               X1              =   1080
               X2              =   3975
               Y1              =   2235
               Y2              =   2235
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   13
               X1              =   1080
               X2              =   3975
               Y1              =   2565
               Y2              =   2565
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   14
               X1              =   1080
               X2              =   3975
               Y1              =   2985
               Y2              =   2985
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   15
               X1              =   1080
               X2              =   3975
               Y1              =   3600
               Y2              =   3600
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   16
               X1              =   1080
               X2              =   5850
               Y1              =   4230
               Y2              =   4230
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   17
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
               Y1              =   3600
               Y2              =   3600
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   19
               X1              =   4575
               X2              =   5850
               Y1              =   3915
               Y2              =   3915
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   20
               X1              =   6435
               X2              =   7620
               Y1              =   2985
               Y2              =   2985
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   21
               X1              =   6435
               X2              =   7620
               Y1              =   3600
               Y2              =   3600
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   22
               X1              =   1080
               X2              =   2205
               Y1              =   3915
               Y2              =   3915
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   23
               X1              =   2790
               X2              =   3975
               Y1              =   3915
               Y2              =   3915
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   24
               X1              =   1080
               X2              =   2600
               Y1              =   4665
               Y2              =   4665
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   25
               X1              =   1080
               X2              =   2600
               Y1              =   6105
               Y2              =   6105
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   26
               X1              =   3390
               X2              =   4890
               Y1              =   4665
               Y2              =   4665
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   27
               X1              =   3195
               X2              =   4255
               Y1              =   6105
               Y2              =   6105
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   28
               X1              =   1065
               X2              =   2600
               Y1              =   5025
               Y2              =   5025
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   29
               X1              =   4830
               X2              =   5860
               Y1              =   6120
               Y2              =   6120
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   30
               X1              =   3390
               X2              =   4890
               Y1              =   5025
               Y2              =   5025
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   31
               X1              =   6795
               X2              =   7620
               Y1              =   6135
               Y2              =   6135
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   32
               X1              =   1080
               X2              =   2700
               Y1              =   5430
               Y2              =   5430
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   33
               X1              =   3555
               X2              =   5175
               Y1              =   5430
               Y2              =   5430
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   34
               X1              =   6000
               X2              =   7620
               Y1              =   5430
               Y2              =   5430
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "   西医诊断 "
            ForeColor       =   &H00FF0000&
            Height          =   270
            Index           =   1
            Left            =   240
            TabIndex        =   156
            Tag             =   "4440"
            Top             =   1740
            Width           =   7830
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   95
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   219
               Top             =   3780
               Width           =   1050
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   64
               Left            =   3930
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   217
               Top             =   3795
               Width           =   2280
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "死亡患者尸检"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   6
               Left            =   6315
               TabIndex        =   216
               Top             =   3810
               Width           =   1440
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "新发肿瘤"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   5
               Left            =   6315
               TabIndex        =   215
               Top             =   3150
               Width           =   1020
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   94
               Left            =   4410
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   211
               Top             =   4110
               Width           =   3000
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   93
               Left            =   6345
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   209
               Top             =   3465
               Width           =   1155
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   92
               Left            =   3930
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   207
               Top             =   3465
               Width           =   915
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   91
               Left            =   1335
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   205
               Top             =   3450
               Width           =   915
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   35
               Left            =   2895
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   164
               Top             =   4110
               Width           =   510
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   36
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   163
               Top             =   4110
               Width           =   510
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
               TabIndex        =   162
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
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
               TabIndex        =   161
               Top             =   2805
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
               TabIndex        =   160
               Top             =   3120
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
               TabIndex        =   159
               Top             =   2805
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
               TabIndex        =   158
               Top             =   3120
               Width           =   915
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   70
               Left            =   1335
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   157
               Top             =   2805
               Width           =   915
            End
            Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
               Height          =   2385
               Left            =   165
               TabIndex        =   270
               Top             =   285
               Width           =   7500
               _cx             =   13229
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
               Cols            =   12
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmChildMedrec.frx":0000
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
               Index           =   92
               X1              =   1245
               X2              =   2375
               Y1              =   3960
               Y2              =   3960
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "死亡时间"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   97
               Left            =   480
               TabIndex        =   220
               Top             =   3780
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "死亡原因"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   66
               Left            =   3105
               TabIndex        =   218
               Top             =   3795
               Width           =   720
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   61
               X1              =   3840
               X2              =   6210
               Y1              =   3975
               Y2              =   3975
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   91
               X1              =   4305
               X2              =   7470
               Y1              =   4290
               Y2              =   4290
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "抢救原因"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   96
               Left            =   3570
               TabIndex        =   212
               Top             =   4125
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "感染病原学诊断"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   95
               Left            =   5025
               TabIndex        =   210
               Top             =   3465
               Width           =   1260
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   90
               X1              =   6315
               X2              =   7475
               Y1              =   3645
               Y2              =   3645
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "最高诊断依据"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   94
               Left            =   2730
               TabIndex        =   208
               Top             =   3465
               Width           =   1080
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   89
               X1              =   3825
               X2              =   4955
               Y1              =   3645
               Y2              =   3645
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "分化程度"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   93
               Left            =   480
               TabIndex        =   206
               Top             =   3450
               Width           =   720
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   88
               X1              =   1245
               X2              =   2375
               Y1              =   3630
               Y2              =   3630
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "成功次数"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   36
               Left            =   2040
               TabIndex        =   171
               Top             =   4110
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
               Left            =   480
               TabIndex        =   170
               Top             =   4110
               Width           =   720
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   36
               X1              =   1230
               X2              =   1830
               Y1              =   4290
               Y2              =   4290
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   37
               X1              =   2805
               X2              =   3405
               Y1              =   4290
               Y2              =   4290
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   63
               X1              =   6315
               X2              =   7475
               Y1              =   2985
               Y2              =   2985
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   64
               X1              =   3840
               X2              =   4925
               Y1              =   3300
               Y2              =   3300
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   65
               X1              =   3840
               X2              =   4925
               Y1              =   2985
               Y2              =   2985
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   66
               X1              =   1245
               X2              =   2375
               Y1              =   3300
               Y2              =   3300
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   67
               X1              =   1245
               X2              =   2375
               Y1              =   2985
               Y2              =   2985
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
               TabIndex        =   169
               Top             =   2790
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
               TabIndex        =   168
               Top             =   2805
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
               TabIndex        =   167
               Top             =   2805
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
               TabIndex        =   166
               Top             =   3120
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
               TabIndex        =   165
               Top             =   3120
               Width           =   900
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "   过敏与手术 "
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   3
            Left            =   240
            TabIndex        =   118
            Tag             =   "3975"
            Top             =   2700
            Width           =   7830
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   111
               Left            =   7080
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   267
               Top             =   3255
               Width           =   270
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   110
               Left            =   6060
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   264
               Top             =   3255
               Width           =   270
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   109
               Left            =   5025
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   259
               Top             =   3270
               Width           =   270
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   108
               Left            =   3570
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   258
               Top             =   3270
               Width           =   270
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   107
               Left            =   2100
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   255
               Top             =   3270
               Width           =   270
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   105
               Left            =   615
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   250
               Top             =   3285
               Width           =   270
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   65
               Left            =   1215
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   213
               Top             =   3630
               Width           =   915
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
               TabIndex        =   125
               Top             =   2730
               Width           =   510
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
               TabIndex        =   124
               Top             =   2730
               Width           =   510
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
               TabIndex        =   123
               Top             =   2730
               Width           =   1695
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
               TabIndex        =   122
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
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
               TabIndex        =   121
               Top             =   1710
               Visible         =   0   'False
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
               TabIndex        =   120
               Top             =   1710
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   75
               Left            =   960
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   119
               Top             =   1710
               Visible         =   0   'False
               Width           =   780
            End
            Begin VSFlex8Ctl.VSFlexGrid vsOPS 
               Height          =   1335
               Left            =   165
               TabIndex        =   126
               Top             =   1830
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
               Cols            =   13
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmChildMedrec.frx":01A5
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
               TabIndex        =   127
               Top             =   315
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
               FormatString    =   $"frmChildMedrec.frx":032D
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
               Caption         =   "天"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   124
               Left            =   7380
               TabIndex        =   266
               Top             =   3255
               Width           =   180
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   108
               X1              =   7035
               X2              =   7335
               Y1              =   3450
               Y2              =   3450
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CCU"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   123
               Left            =   6750
               TabIndex        =   265
               Top             =   3255
               Width           =   270
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "天"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   122
               Left            =   6360
               TabIndex        =   263
               Top             =   3255
               Width           =   180
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   107
               X1              =   6015
               X2              =   6315
               Y1              =   3450
               Y2              =   3450
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ICU"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   121
               Left            =   5700
               TabIndex        =   262
               Top             =   3255
               Width           =   270
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "天"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   120
               Left            =   5340
               TabIndex        =   261
               Top             =   3255
               Width           =   180
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   106
               X1              =   4995
               X2              =   5295
               Y1              =   3450
               Y2              =   3450
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "三级护理"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   119
               Left            =   4230
               TabIndex        =   260
               Top             =   3270
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "天"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   118
               Left            =   3870
               TabIndex        =   257
               Top             =   3270
               Width           =   180
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   105
               X1              =   3525
               X2              =   3825
               Y1              =   3465
               Y2              =   3465
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "二级护理"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   117
               Left            =   2775
               TabIndex        =   256
               Top             =   3270
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "天"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   116
               Left            =   2400
               TabIndex        =   254
               Top             =   3270
               Width           =   180
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   104
               X1              =   2055
               X2              =   2355
               Y1              =   3465
               Y2              =   3465
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "一级护理"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   115
               Left            =   1290
               TabIndex        =   253
               Top             =   3270
               Width           =   735
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "天"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   112
               Left            =   930
               TabIndex        =   252
               Top             =   3270
               Width           =   180
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   102
               X1              =   585
               X2              =   885
               Y1              =   3465
               Y2              =   3465
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "特护"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   111
               Left            =   180
               TabIndex        =   251
               Top             =   3285
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "术前与术后"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   72
               Left            =   180
               TabIndex        =   214
               Top             =   3630
               Width           =   900
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   62
               X1              =   1125
               X2              =   2285
               Y1              =   3810
               Y2              =   3810
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   70
               Visible         =   0   'False
               X1              =   4590
               X2              =   5615
               Y1              =   1890
               Y2              =   1890
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   71
               Visible         =   0   'False
               X1              =   2715
               X2              =   3715
               Y1              =   1890
               Y2              =   1890
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   72
               Visible         =   0   'False
               X1              =   870
               X2              =   1865
               Y1              =   1890
               Y2              =   1890
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
               TabIndex        =   130
               Top             =   1710
               Visible         =   0   'False
               Width           =   450
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
               TabIndex        =   129
               Top             =   1710
               Visible         =   0   'False
               Width           =   540
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
               TabIndex        =   128
               Top             =   1710
               Visible         =   0   'False
               Width           =   540
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   中医诊断 "
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   2
            Left            =   240
            TabIndex        =   131
            Tag             =   "4050"
            Top             =   2145
            Width           =   7830
            Begin VB.Frame fraSub 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   " 治疗方法 "
               ForeColor       =   &H00404040&
               Height          =   1320
               Index           =   2
               Left            =   4680
               TabIndex        =   146
               Top             =   2580
               Width           =   2985
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   38
                  Left            =   1545
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   149
                  Top             =   330
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
                  TabIndex        =   148
                  Top             =   645
                  Width           =   1035
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   40
                  Left            =   1545
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   147
                  Top             =   960
                  Width           =   1035
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
                  TabIndex        =   152
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
                  Index           =   42
                  Left            =   675
                  TabIndex        =   151
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
                  Index           =   41
                  Left            =   315
                  TabIndex        =   150
                  Top             =   960
                  Width           =   1080
               End
               Begin VB.Line lineInfo 
                  BorderColor     =   &H00808080&
                  Index           =   38
                  X1              =   1455
                  X2              =   2580
                  Y1              =   510
                  Y2              =   510
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
                  Index           =   40
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
               TabIndex        =   142
               Top             =   2580
               Width           =   1845
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "疑难"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   4
                  Left            =   525
                  TabIndex        =   145
                  Top             =   960
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
                  TabIndex        =   144
                  Top             =   645
                  Width           =   660
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "危重"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   2
                  Left            =   525
                  TabIndex        =   143
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
               TabIndex        =   141
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
               Left            =   2160
               TabIndex        =   134
               Top             =   2580
               Width           =   2385
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   59
                  Left            =   840
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   137
                  Top             =   960
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
                  TabIndex        =   136
                  Top             =   645
                  Width           =   1035
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   57
                  Left            =   840
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   135
                  Top             =   330
                  Width           =   1035
               End
               Begin VB.Line lineInfo 
                  BorderColor     =   &H00808080&
                  Index           =   57
                  X1              =   750
                  X2              =   1875
                  Y1              =   1140
                  Y2              =   1140
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
                  Index           =   59
                  X1              =   750
                  X2              =   1875
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
                  Index           =   38
                  Left            =   330
                  TabIndex        =   140
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
                  Index           =   39
                  Left            =   330
                  TabIndex        =   139
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
                  Index           =   40
                  Left            =   330
                  TabIndex        =   138
                  Top             =   330
                  Width           =   360
               End
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
               TabIndex        =   133
               Top             =   2190
               Width           =   915
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   72
               Left            =   4020
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   132
               Top             =   2190
               Width           =   915
            End
            Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
               Height          =   1830
               Left            =   165
               TabIndex        =   153
               Top             =   270
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
               Cols            =   11
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmChildMedrec.frx":037E
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
               Index           =   73
               Left            =   3000
               TabIndex        =   155
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
               Index           =   74
               Left            =   390
               TabIndex        =   154
               Top             =   2190
               Width           =   900
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   68
               X1              =   1335
               X2              =   2465
               Y1              =   2370
               Y2              =   2370
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   69
               X1              =   3930
               X2              =   5015
               Y1              =   2370
               Y2              =   2370
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   住院情况 "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   78
            Tag             =   "4170"
            Top             =   3210
            Width           =   7830
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   112
               Left            =   3765
               MaxLength       =   10
               TabIndex        =   268
               Top             =   2400
               Width           =   1080
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   106
               Left            =   1155
               MaxLength       =   10
               TabIndex        =   247
               Top             =   2385
               Width           =   1080
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   104
               Left            =   6225
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   243
               Top             =   3810
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   103
               Left            =   3750
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   242
               Top             =   3810
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   102
               Left            =   1155
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   241
               Top             =   3810
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   101
               Left            =   6210
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   238
               Top             =   1410
               Width           =   1080
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   100
               Left            =   4215
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   236
               Top             =   1035
               Width           =   450
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   99
               Left            =   3075
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   234
               Top             =   1035
               Width           =   450
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   98
               Left            =   2100
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   231
               Top             =   1020
               Width           =   450
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   97
               Left            =   4215
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   229
               Top             =   720
               Width           =   450
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   96
               Left            =   3075
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   227
               Top             =   720
               Width           =   450
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   46
               Left            =   2100
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   224
               Top             =   705
               Width           =   450
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   37
               Left            =   1155
               Locked          =   -1  'True
               MaxLength       =   30
               TabIndex        =   221
               Top             =   360
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
               MaxLength       =   30
               TabIndex        =   96
               Top             =   2055
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
               TabIndex        =   95
               Top             =   2055
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
               TabIndex        =   94
               Top             =   1740
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
               TabIndex        =   93
               Top             =   1740
               Width           =   1080
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
               TabIndex        =   92
               Top             =   1740
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
               TabIndex        =   91
               Top             =   1425
               Width           =   1080
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
               TabIndex        =   90
               Top             =   1425
               Width           =   1080
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
               TabIndex        =   89
               Top             =   2850
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   50
               Left            =   6240
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   88
               Top             =   2835
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
               TabIndex        =   87
               Top             =   3480
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
               TabIndex        =   86
               Top             =   3165
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
               TabIndex        =   85
               Top             =   3480
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
               TabIndex        =   84
               Top             =   3165
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
               TabIndex        =   83
               Top             =   3165
               Width           =   1335
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
               TabIndex        =   82
               Top             =   3480
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
               TabIndex        =   81
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
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
               TabIndex        =   80
               Top             =   2055
               Width           =   1335
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "科研病案"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   11
               Left            =   3360
               TabIndex        =   79
               Top             =   360
               Width           =   1020
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输血反应"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   125
               Left            =   2910
               TabIndex        =   269
               Top             =   2400
               Width           =   720
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   109
               X1              =   3675
               X2              =   4845
               Y1              =   2580
               Y2              =   2580
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ml"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   114
               Left            =   2265
               TabIndex        =   249
               Top             =   2415
               Width           =   180
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   103
               X1              =   1065
               X2              =   2235
               Y1              =   2565
               Y2              =   2565
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "自体回收"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   113
               Left            =   300
               TabIndex        =   248
               Top             =   2385
               Width           =   720
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   101
               X1              =   6135
               X2              =   7560
               Y1              =   3990
               Y2              =   3990
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   100
               X1              =   3660
               X2              =   5085
               Y1              =   3990
               Y2              =   3990
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   99
               X1              =   1065
               X2              =   2490
               Y1              =   3990
               Y2              =   3990
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "质控医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   110
               Left            =   300
               TabIndex        =   246
               Top             =   3810
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "质控护士"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   109
               Left            =   2910
               TabIndex        =   245
               Top             =   3810
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "责任护士"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   108
               Left            =   5370
               TabIndex        =   244
               Top             =   3810
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "小时"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   107
               Left            =   7335
               TabIndex        =   240
               Top             =   1395
               Width           =   360
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   98
               X1              =   6120
               X2              =   7290
               Y1              =   1590
               Y2              =   1590
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "呼吸机使用"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   106
               Left            =   5175
               TabIndex        =   239
               Top             =   1410
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "分钟"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   105
               Left            =   4785
               TabIndex        =   237
               Top             =   1035
               Width           =   360
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   97
               X1              =   4125
               X2              =   4695
               Y1              =   1215
               Y2              =   1215
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "小时"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   104
               Left            =   3645
               TabIndex        =   235
               Top             =   1035
               Width           =   360
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   96
               X1              =   2985
               X2              =   3555
               Y1              =   1215
               Y2              =   1215
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "天"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   103
               Left            =   2685
               TabIndex        =   233
               Top             =   1035
               Width           =   180
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院后"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   102
               Left            =   1365
               TabIndex        =   232
               Top             =   1035
               Width           =   540
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   95
               X1              =   2010
               X2              =   2580
               Y1              =   1200
               Y2              =   1200
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "分钟"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   101
               Left            =   4785
               TabIndex        =   230
               Top             =   720
               Width           =   360
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   94
               X1              =   4125
               X2              =   4695
               Y1              =   900
               Y2              =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "小时"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   100
               Left            =   3645
               TabIndex        =   228
               Top             =   720
               Width           =   360
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   93
               X1              =   2985
               X2              =   3555
               Y1              =   900
               Y2              =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "天"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   99
               Left            =   2685
               TabIndex        =   226
               Top             =   720
               Width           =   180
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院前"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   98
               Left            =   1365
               TabIndex        =   225
               Top             =   720
               Width           =   540
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   41
               X1              =   2010
               X2              =   2580
               Y1              =   885
               Y2              =   885
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "昏迷时间"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   44
               Left            =   300
               TabIndex        =   223
               Top             =   720
               Width           =   720
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   35
               X1              =   1065
               X2              =   2235
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
               Index           =   37
               Left            =   300
               TabIndex        =   222
               Top             =   345
               Width           =   720
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
               TabIndex        =   117
               Top             =   3480
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
               TabIndex        =   116
               Top             =   3480
               Width           =   900
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
               TabIndex        =   115
               Top             =   3480
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
               TabIndex        =   114
               Top             =   3165
               Width           =   720
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
               TabIndex        =   113
               Top             =   3165
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "主任(副主任)医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   59
               Left            =   4665
               TabIndex        =   112
               Top             =   2835
               Width           =   1440
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
               TabIndex        =   111
               Top             =   3165
               Width           =   540
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
               TabIndex        =   110
               Top             =   2850
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ml"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   56
               Left            =   4875
               TabIndex        =   109
               Top             =   2055
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
               TabIndex        =   108
               Top             =   2055
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
               TabIndex        =   107
               Top             =   2055
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
               TabIndex        =   106
               Top             =   2055
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
               TabIndex        =   105
               Top             =   1740
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
               TabIndex        =   104
               Top             =   1740
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "单位"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   50
               Left            =   4860
               TabIndex        =   103
               Top             =   1740
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
               TabIndex        =   102
               Top             =   1740
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
               TabIndex        =   101
               Top             =   1740
               Width           =   360
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
               TabIndex        =   100
               Top             =   1740
               Width           =   720
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
               TabIndex        =   99
               Top             =   1425
               Width           =   180
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
               TabIndex        =   98
               Top             =   1425
               Width           =   360
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   42
               X1              =   1065
               X2              =   2235
               Y1              =   1605
               Y2              =   1605
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   43
               X1              =   1065
               X2              =   2235
               Y1              =   1920
               Y2              =   1920
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   44
               X1              =   1065
               X2              =   2235
               Y1              =   2235
               Y2              =   2235
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   45
               X1              =   3660
               X2              =   4830
               Y1              =   1605
               Y2              =   1605
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   46
               X1              =   3660
               X2              =   4830
               Y1              =   1920
               Y2              =   1920
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   47
               X1              =   3660
               X2              =   4830
               Y1              =   2235
               Y2              =   2235
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   48
               X1              =   6135
               X2              =   7305
               Y1              =   1920
               Y2              =   1920
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   49
               X1              =   1065
               X2              =   2490
               Y1              =   3030
               Y2              =   3030
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   50
               X1              =   6150
               X2              =   7575
               Y1              =   3015
               Y2              =   3015
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   51
               X1              =   1065
               X2              =   2490
               Y1              =   3660
               Y2              =   3660
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   52
               X1              =   3660
               X2              =   5085
               Y1              =   3345
               Y2              =   3345
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   53
               X1              =   3660
               X2              =   5085
               Y1              =   3660
               Y2              =   3660
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   54
               X1              =   1065
               X2              =   2490
               Y1              =   3345
               Y2              =   3345
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   55
               X1              =   6135
               X2              =   7560
               Y1              =   3345
               Y2              =   3345
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   56
               X1              =   6135
               X2              =   7560
               Y1              =   3660
               Y2              =   3660
            End
            Begin VB.Line lineInfo 
               BorderColor     =   &H00808080&
               Index           =   60
               X1              =   6135
               X2              =   7560
               Y1              =   2235
               Y2              =   2235
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
               TabIndex        =   97
               Top             =   2055
               Width           =   720
            End
         End
      End
      Begin VB.Frame fraVH 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9105
         TabIndex        =   3
         Top             =   7590
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.HScrollBar hsc 
         Height          =   255
         Left            =   105
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   7680
         Visible         =   0   'False
         Width           =   8115
      End
      Begin VB.VScrollBar vsc 
         Height          =   5475
         Left            =   8850
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
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
               Picture         =   "frmChildMedrec.frx":04F0
               Key             =   "-"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmChildMedrec.frx":09DA
               Key             =   "+"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmChildMedrec"
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

Private Const ColorUnEditCell = &H8000000B  '灰蓝色

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
    txt入科时间 = 76
    txt再入院计划 = 77
    txt再入院目的 = 78
    txt身高 = 79
    txt体重 = 80
    txt婴幼儿年龄 = 81
    txt新生儿体重 = 82
    txt新生儿入院体重 = 83
    txt户口地址 = 84
    txt户口地址邮编 = 85
    txt入院途径 = 86
    txt籍贯 = 87
    txt其他证件 = 88
    txt出院方式 = 89
    txt出院转入 = 90
    chk外院治疗 = 15
    
    txt特护 = 105
    txt一级护理 = 107
    txt二级护理 = 108
    txt三级护理 = 109
    txtICU = 110
    txtCCU = 111
    
    txt输血反应 = 112
End Enum
Private Enum 西医诊断
'    chk是否确诊 = 0
    txt抢救次数 = 36
'    txt确诊日期 = 37
    txt成功次数 = 35
    txt门诊与出院 = 70
    txt入院与出院 = 68
    txt放射与病理 = 66
    txt临床与病理 = 69
    txt临床与尸检 = 67
    txt术前与术后 = 65
    txt分化程度 = 91
    txt最高诊断依据 = 92
    txt感染病原学诊断 = 93
    txt抢救原因 = 94
    txt死亡时间 = 95
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

''Private Enum 过敏与手术
''    chk首例手术 = 1
''    chk首例治疗 = 12
''    chk首例检查 = 13
''    chk首例诊断 = 14
''End Enum

Private Enum 住院情况
    chk尸检 = 6
    chk随诊 = 7
    chk新发肿瘤 = 5
'    chk示教病案 = 8
    chk科研病案 = 11
'    chk输血反应 = 9
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
    txt入院前天 = 46
    txt入院前小时 = 96
    txt入院前分钟 = 97
    
    txt入院后天 = 98
    txt入院后小时 = 99
    txt入院后分钟 = 100
    txt呼吸机使用 = 101
    
    txt质控医师 = 102
    txt质控护士 = 103
    txt责任护士 = 104
    
    txt病例分型 = 37
    txt自体回收 = 106
End Enum

Private Enum 诊断情况
    col诊断类型 = 0
    col诊断编码 = 1
    col诊断描述 = 2
    col备注 = 3
    col入院病情 = 4
    col出院情况 = 5
    col是否未治 = 6
    col是否疑诊 = 7
    col增加 = 8
    col诊断ID = 9
    col疾病ID = 10
    col类型 = 11 '1-西医门诊诊断;2-西医入院诊断;3-出院诊断(其他诊断);5-院内感染;6-病理诊断;7-损伤中毒码;10-并发症
    
    colzy增加 = 6
    colzy诊断ID = 7
    colzy疾病ID = 8
    colzy证候ID = 9
    colzy类型 = 10
End Enum
Private Enum 手术情况
    col手术日期 = 0
    col手术名称 = 1
    col再次手术 = 2
    col主刀医师 = 3
    col助产护士 = 4
    col助手1 = 5
    col助手2 = 6
    col麻醉方式 = 7
    colASA分级 = 8
    colNNIS分级 = 9
    col手术分级 = 10
    col麻醉医师 = 11
    col切口愈合 = 12
End Enum

Public Function zlRefresh(lng病人ID As Long, lng主页ID As Long, lng病区ID As Long, lng科室ID As Long, Optional bln出院 As Boolean) As Boolean
'功能：刷新或清除医嘱清单
    
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    mlng病人ID = lng病人ID: mlng主页ID = lng主页ID
    mlng病区ID = lng病区ID: mlng科室ID = lng科室ID
    mbln出院 = bln出院: mblnMoved = False
        
    strSQL = "Select 出院科室ID,当前病区ID From 病案主页 Where 病人id=[1] And 主页id=[2]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If rs.BOF = False Then
        mlng科室ID = zlCommFun.NVL(rs("出院科室ID").Value, 0)
        mlng病区ID = zlCommFun.NVL(rs("当前病区ID").Value, 0)
    End If
    
    '可能科室切换
    mbln中医 = Have部门性质(mlng科室ID, "中医科")
    fraInfo(2).Visible = mbln中医
    fraInfo(2).Enabled = mbln中医 '标志不操作
    Call SetPageHeight
    Call SetScrollbar
    
    Call ClearPageData
    If mlng病人ID <> 0 Then Call LoadPageData
    Call Form_Resize
    Exit Function
errHand:
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

Private Sub Form_Load()
    '滚动条尺寸
    vsc.Width = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    hsc.Height = GetSystemMetrics(SM_CXHSCROLL) * Screen.TwipsPerPixelY
    fraVH.Width = vsc.Width: fraVH.Height = hsc.Height
    fraBack.Left = 0: fraBack.Top = 0
    picBack.BackColor = fraBack.BackColor
    
    optExits(0).Enabled = False
    optExits(1).Enabled = False
'    '初始化系统参数
'    Call InitSysPar
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
    optExits(0).Value = True
    optExits(1).Value = False
    
    mblnCheck = False
End Sub

Private Function LoadPageData() As Boolean
'功能：读取病人的首页信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim lngRow As Long, varTmp As Variant
    Dim strTmp As String
    Dim str治疗结果 As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    mblnCheck = True
    
    '病人信息部份
    '---------------------------------------------------------------
    strSQL = "Select 住院号,性别,姓名,出生日期,出生地点,身份证号,其他证件,民族,籍贯 From 病人信息 Where 病人ID=[1]"
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
    txtInfo(txt籍贯).Text = NVL(rsTmp!籍贯)
    txtInfo(txt其他证件).Text = NVL(rsTmp!其他证件)
    
    '病案主页部份
    '---------------------------------------------------------------
    strSQL = "Select A.病人ID,A.主页ID,A.住院号,A.病人性质,A.医疗付款方式,A.费别,A.再入院,A.入院病区ID,A.入院科室ID,A.入院日期,A.入院病况,A.入院方式,A.入院属性,A.二级院转入,A.住院目的,A.入院病床,A.是否陪伴,A.当前病况,A.当前病区ID,A.护理等级ID,A.出院科室ID,A.出院病床,A.出院日期,A.住院天数,A.出院方式,A.是否确诊,A.确诊日期,A.新发肿瘤,A.血型,A.抢救次数,A.成功次数,A.随诊标志,A.随诊期限,A.尸检标志,A.门诊医师,A.责任护士,A.住院医师,A.病案号,A.编目员编号,A.编目员姓名,A.编目日期,A.状态,A.费用和,A.年龄,A.婚姻状况,A.职业,A.国籍,A.学历,A.单位电话,A.单位邮编,A.单位地址,A.区域,A.家庭地址,A.家庭电话,A.家庭地址邮编,A.联系人姓名,A.联系人关系,A.联系人地址,A.联系人电话,A.中医治疗类别,A.险类,A.审核标志,A.审核人,A.审核日期,A.是否上传,A.数据转出,A.登记人,A.登记时间,A.备注,A.社区,A.病案状态,A.封存时间,A.病人类型,A.身高,A.体重,A.户口地址,A.户口地址邮编,B.名称 as 入院科室,C.名称 as 出院科室" & _
        " From 病案主页 A,部门表 B,部门表 C" & _
        " Where A.入院科室ID=B.ID And A.出院科室ID=C.ID" & _
        " And A.病人ID=[1] And A.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    mblnMoved = NVL(rsTmp!数据转出, 0) = 1
    
    '留观病人无住院号
    If NVL(rsTmp!病人性质, 0) <> 0 Then
        lblInfo(0).Visible = False
        txtInfo(txt住院号).Visible = False
    Else
        lblInfo(0).Visible = True
        txtInfo(txt住院号).Visible = True
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
'    chkInfo(chk是否确诊).Value = NVL(rsTmp!是否确诊, 0)
'    If chkInfo(chk是否确诊).Value = 1 Then
'        txtInfo(txt确诊日期).Text = Format(NVL(rsTmp!确诊日期), "yyyy-MM-dd HH:mm")
'    End If
    txtInfo(txt抢救次数).Text = NVL(rsTmp!抢救次数)
    If Val(txtInfo(txt抢救次数).Text) <> 0 Then
        txtInfo(txt成功次数).Text = NVL(rsTmp!成功次数)
    End If
    chkInfo(chk新发肿瘤).Value = NVL(rsTmp!新发肿瘤, 0)
    
    txtInfo(txt治疗类别).Text = NVL(rsTmp!中医治疗类别)
    chkInfo(chk尸检).Value = NVL(rsTmp!尸检标志, 0)
'    chkInfo(chk随诊).Value = IIf(NVL(rsTmp!随诊标志, 0) = 0, 0, 1)
'    If chkInfo(chk随诊).Value = 1 Then
'        txtInfo(txt随诊期限).Text = NVL(rsTmp!随诊期限, 0) & Decode(NVL(rsTmp!随诊标志, 0), 1, "月", 2, "年", 3, "周")
'    End If
    txtInfo(txt门诊医师).Text = NVL(rsTmp!门诊医师)
    txtInfo(txt住院医师).Text = NVL(rsTmp!住院医师)
    txtInfo(txt血型).Text = NVL(rsTmp!血型)
    
    If NVL(rsTmp!身高) = "" Or NVL(rsTmp!身高) = 0 Then
        txtInfo(txt身高).Text = ""
    Else
        txtInfo(txt身高).Text = NVL(rsTmp!身高) & "cm"
    End If
    
    If NVL(rsTmp!体重) = "" Or NVL(rsTmp!体重) = 0 Then
        txtInfo(txt体重).Text = ""
    Else
        txtInfo(txt体重).Text = NVL(rsTmp!体重) & "kg"
    End If
    
    txtInfo(txt户口地址).Text = NVL(rsTmp!户口地址)
    txtInfo(txt户口地址邮编).Text = NVL(rsTmp!户口地址邮编)
    
    txtInfo(txt入院途径).Text = NVL(rsTmp!入院方式)
    txtInfo(txt出院方式).Text = NVL(rsTmp!出院方式)
    
    txtInfo(txt责任护士).Text = NVL(rsTmp!责任护士)
   
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
                
'''                chkInfo(chk首例手术).Value = 0
'''                chkInfo(chk首例治疗).Value = 0
'''                chkInfo(chk首例检查).Value = 0
'''                chkInfo(chk首例诊断).Value = 0
'''
'''                strTmp = NVL(rsTmp("信息值").Value, "0000")
'''                If Len(strTmp) >= 1 Then chkInfo(chk首例手术).Value = Val(Mid(strTmp, 1, 1))
'''                If Len(strTmp) >= 2 Then chkInfo(chk首例治疗).Value = Val(Mid(strTmp, 2, 1))
'''                If Len(strTmp) >= 3 Then chkInfo(chk首例检查).Value = Val(Mid(strTmp, 3, 1))
'''                If Len(strTmp) >= 4 Then chkInfo(chk首例诊断).Value = Val(Mid(strTmp, 4, 1))
                
                                
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
'                chkInfo(chk示教病案).Value = Val(NVL(rsTmp!信息值, 0))
            Case "科研病案"
                chkInfo(chk科研病案).Value = Val(NVL(rsTmp!信息值, 0))
            Case UCase("Rh")
                txtInfo(txtRh).Text = NVL(rsTmp!信息值)
            Case "输血反应"
                txtInfo(txt输血反应).Text = Decode(Val(NVL(rsTmp!信息值, 0)), 0, "无", 1, "有", 2, "未输")
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
            Case "再入院计划天数"
                If NVL(rsTmp!信息值, 0) = 0 Then
                    txtInfo(txt再入院计划).Text = "31天再住院计划"
                Else
                    txtInfo(txt再入院计划).Text = "7天再住院计划"
                End If
            Case "31天内再住院"
                If NVL(rsTmp!信息值) = "" Then
                    optExits(0).Value = True
                    txtInfo(txt再入院目的).Text = ""
                Else
                    optExits(1).Value = True
                    txtInfo(txt再入院目的).Text = NVL(rsTmp!信息值)
                End If
            Case "不足周岁年龄"
                txtInfo(txt婴幼儿年龄).Text = NVL(rsTmp!信息值)
            Case "新生儿出生体重"
                If NVL(rsTmp!信息值) = "" Then
                    txtInfo(txt新生儿体重).Text = ""
                Else
                    txtInfo(txt新生儿体重).Text = NVL(rsTmp!信息值) & "克"
                End If
            Case "新生儿入院体重"
                If NVL(rsTmp!信息值) = "" Then
                    txtInfo(txt新生儿入院体重).Text = ""
                Else
                    txtInfo(txt新生儿入院体重).Text = NVL(rsTmp!信息值) & "克"
                End If
            Case "籍贯"
                '如果已经有了籍贯,就不在从从表中读取。
                If txtInfo(txt籍贯).Text = "" Then
                    txtInfo(txt籍贯).Text = NVL(rsTmp!信息值)
                End If
'            Case "其他证件"
'                txtInfo(txt其他证件).Text = NVL(rsTmp!信息值)
            Case "出院转入"
                txtInfo(txt出院转入).Text = NVL(rsTmp!信息值)
            Case "入院前经外院治疗"
                chkInfo(chk外院治疗).Value = Val(NVL(rsTmp!信息值, 0))
            Case "抢救病因"
                txtInfo(txt抢救原因).Text = NVL(rsTmp!信息值)
            Case "死亡时间"
                txtInfo(txt死亡时间).Text = NVL(rsTmp!信息值)
            Case "昏迷时间"
                Call SetInfoTime(NVL(rsTmp!信息值))
            Case "呼吸机使用时间"
                txtInfo(txt呼吸机使用).Text = NVL(rsTmp!信息值)
            Case "质控医师"
                txtInfo(txt质控医师).Text = NVL(rsTmp!信息值)
            Case "质控护士"
                txtInfo(txt质控护士).Text = NVL(rsTmp!信息值)
            Case "病例分型"
                txtInfo(txt病例分型).Text = Get病例分型(NVL(rsTmp!信息值))
            Case "分化程度"
                txtInfo(txt分化程度).Text = NVL(rsTmp!信息值)
            Case "最高诊断依据"
                txtInfo(txt最高诊断依据).Text = NVL(rsTmp!信息值)
            Case "自体回收"
                txtInfo(txt自体回收).Text = NVL(rsTmp!信息值)
            Case "特级护理天数"
                txtInfo(txt特护).Text = NVL(rsTmp!信息值)
            Case "一级护理天数"
                txtInfo(txt一级护理).Text = NVL(rsTmp!信息值)
            Case "二级护理天数"
                txtInfo(txt二级护理).Text = NVL(rsTmp!信息值)
            Case "三级护理天数"
                txtInfo(txt三级护理).Text = NVL(rsTmp!信息值)
            Case "ICU天数"
                txtInfo(txtICU).Text = NVL(rsTmp!信息值)
            Case "CCU天数"
                txtInfo(txtCCU).Text = NVL(rsTmp!信息值)
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
    '读取入科时间
    strSQL = _
            " Select A.开始时间" & _
            " From 病人变动记录 A,部门表 B" & _
            " Where A.病人ID=[1] And A.主页ID=[2]" & _
            " And A.科室ID=B.ID And A.开始原因=2 And A.开始时间 is Not NULL" & _
            " Order by A.开始时间"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If rsTmp Is Nothing Then
        txtInfo(txt入科时间).Text = txtInfo(txt入院时间).Text
    ElseIf rsTmp.EOF Or rsTmp.BOF Then
        txtInfo(txt入科时间).Text = txtInfo(txt入院时间).Text
    Else
        txtInfo(txt入科时间).Text = Format("" & rsTmp!开始时间, "yyyy-mm-dd hh:mm")
    End If
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
   str治疗结果 = Get治疗结果
    vsDiagXY.ColData(col出院情况) = str治疗结果
  
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
    strSQL = "Select a.备注,a.ID,a.病人ID,a.主页ID,a.医嘱ID,a.记录来源,a.诊断次序,a.编码序号,a.病历ID,a.诊断类型,a.疾病ID,a.入院病情," & _
        " a.诊断ID,a.证候ID,a.诊断描述,a.出院情况,a.是否未治,a.是否疑诊,a.记录日期,a.记录人,a.取消时间,a.取消人,a.病例ID, b.编码 As 疾病编码, c.编码 As 诊断编码 " & _
        " From 病人诊断记录 A, 疾病编码目录 B, 疾病诊断目录 C" & _
        " Where a.疾病id = b.Id(+) And a.诊断id = c.Id(+) And a.记录来源 IN(1,2,3,4) And a.诊断类型 IN(1,2,3,5,6,7,10,21)" & _
        " And a.取消时间 is Null And a.病人ID=[1] And a.主页ID=[2]" & _
        " Order by a.诊断类型,a.记录来源 Desc,a.诊断次序,a.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Not rsTmp.EOF Then
        With vsDiagXY
            strSQL = "1,2,3,5,6,7,10,21"
            For i = 0 To UBound(Split(strSQL, ","))
                rsTmp.Filter = "记录来源=3 And 诊断类型=" & Split(strSQL, ",")(i)
                If Val(Split(strSQL, ",")(i)) <> 21 Then
                    If rsTmp.EOF Then
                        rsTmp.Filter = "记录来源=2 And 诊断类型=" & Split(strSQL, ",")(i)
                    End If
                    If rsTmp.EOF Then
                        rsTmp.Filter = "记录来源=1 And 诊断类型=" & Split(strSQL, ",")(i)
                    End If
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "记录来源=4 And 诊断类型=" & Split(strSQL, ",")(i)
                End If
                
                If Val(Split(strSQL, ",")(i)) = 21 Then
                    '21-病原学诊断
                    If Not rsTmp.EOF Then
                        txtInfo(txt感染病原学诊断).Text = NVL(rsTmp!诊断描述)
                        txtInfo(txt感染病原学诊断).Tag = txtInfo(txt感染病原学诊断).Text
                    End If
                Else
                    Do While Not rsTmp.EOF
                        '确定当前显示行
                        lngRow = .FindRow(CStr(Split(strSQL, ",")(i)), , col类型)
                        For j = lngRow To .Rows - 1
                            If Val(.TextMatrix(j, col类型)) = Val(Split(strSQL, ",")(i)) Then
                                lngRow = j
                                If .TextMatrix(j, col诊断描述) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, col诊断描述) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, col类型) = Split(strSQL, ",")(i)
                        End If
                        
                        If IsNull(rsTmp!诊断描述) Then
                            .TextMatrix(lngRow, col诊断编码) = ""
                            .TextMatrix(lngRow, col诊断描述) = ""
                        Else
                            If Mid(rsTmp!诊断描述, 1, 1) <> "(" Or (Val(rsTmp!诊断id & "") = 0 And Val(rsTmp!疾病id & "") = 0) Then '中医的诊断描述后面加了（候症），所以只判断第一个字符
                                '由于疾病编码和诊断可以对应，如果两个都不为空的时候，先判断疾病编码，先取疾病编码
                                If Val(rsTmp!疾病id & "") <> 0 Then
                                    .TextMatrix(lngRow, col诊断编码) = NVL(rsTmp!疾病编码)
                                ElseIf Val(rsTmp!诊断id & "") <> 0 Then
                                    .TextMatrix(lngRow, col诊断编码) = NVL(rsTmp!诊断编码)
                                Else
                                    .TextMatrix(lngRow, col诊断编码) = ""
                                End If
                                .TextMatrix(lngRow, col诊断描述) = rsTmp!诊断描述
                            Else
                                .TextMatrix(lngRow, col诊断编码) = Mid(rsTmp!诊断描述, 2, InStr(rsTmp!诊断描述, ")") - 2)
                                .TextMatrix(lngRow, col诊断描述) = Mid(rsTmp!诊断描述, InStr(rsTmp!诊断描述, ")") + 1)
                            End If
                        End If
                        If Not IsNull(rsTmp!疾病id) Or Not IsNull(rsTmp!诊断id) Then
                            .Cell(flexcpData, lngRow, col诊断描述) = Get诊断描述(Val("" & rsTmp!诊断id), Val("" & rsTmp!疾病id))    '获取原始名称以便修改时判断
                        Else
                            .Cell(flexcpData, lngRow, col诊断描述) = .TextMatrix(lngRow, col诊断描述)
                        End If
                        
                        .TextMatrix(lngRow, col备注) = NVL(rsTmp!备注)
                        .TextMatrix(lngRow, col出院情况) = NVL(rsTmp!出院情况)
                        .TextMatrix(lngRow, col入院病情) = NVL(rsTmp!入院病情)
                        .TextMatrix(lngRow, col是否未治) = IIf(NVL(rsTmp!是否未治, 0) = 1, "√", "")
                        .TextMatrix(lngRow, col是否疑诊) = IIf(NVL(rsTmp!是否疑诊, 0) = 1, "？", "")
                        .TextMatrix(lngRow, col诊断ID) = NVL(rsTmp!诊断id, 0)
                        .TextMatrix(lngRow, col疾病ID) = NVL(rsTmp!疾病id, 0)
                        rsTmp.MoveNext
                    Loop
                End If
            Next
        End With
    End If
    
    vsDiagXY.Cell(flexcpForeColor, 1, col是否疑诊, vsDiagXY.Rows - 1, col是否疑诊) = vbRed
    vsDiagXY.Cell(flexcpBackColor, GetRow(3), vsDiagXY.FixedRows, GetRow(3), vsDiagXY.Cols - 1) = &HC0FFC0
    vsDiagXY.Cell(flexcpBackColor, 1, col诊断编码, vsDiagXY.Rows - 1, col诊断编码) = ColorUnEditCell      '灰蓝色
    vsDiagXY.Row = 1: vsDiagXY.Col = col诊断描述
    Call vsDiagXY_AfterRowColChange(-1, -1, vsDiagXY.Row, vsDiagXY.Col)
        
    '中医诊断
    '---------------------------------------------------------------
    strSQL = "Select 取消人,病例ID,备注,ID,病人ID,主页ID,医嘱ID,记录来源,诊断次序,编码序号,病历ID,诊断类型,疾病ID,诊断ID,证候ID,诊断描述,出院情况,是否未治,是否疑诊,记录日期,记录人,取消时间,入院病情 From 病人诊断记录" & _
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
                    .TextMatrix(lngRow, col备注) = NVL(rsTmp!备注)
                    .TextMatrix(lngRow, col入院病情) = NVL(rsTmp!入院病情)
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
    strSQL = "Select 记录来源,手术日期,手术开始时间,手术结束时间,拟行手术,手术操作ID,诊疗项目ID,已行手术,主刀医师,再次手术,第一助手,第二助手,手术护士,麻醉开始时间,麻醉结束时间,麻醉方式,麻醉类型,麻醉质量,输液总量,麻醉医师,输氧开始时间,输氧结束时间,切口,愈合,记录日期,记录人,取消时间,取消人,助产护士,ID,病人ID,主页ID,decode(ASA分级,'I级','P1','II级','P2','III级','P3','IV级','P4','V级','P5',ASA分级) as ASA分级,NNIS分级,decode(手术级别,1,'一级手术',2,'二级手术',3,'三级手术',4,'四级手术',' ') as 手术级别 From 病人手麻记录" & _
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
                .TextMatrix(i, col再次手术) = NVL(rsTmp!再次手术, -1)
                .TextMatrix(i, col主刀医师) = NVL(rsTmp!主刀医师)
                .TextMatrix(i, col助产护士) = NVL(rsTmp!助产护士)
                .TextMatrix(i, col助手1) = NVL(rsTmp!第一助手)
                .TextMatrix(i, col助手2) = NVL(rsTmp!第二助手)
                .TextMatrix(i, col麻醉方式) = GetItemField("诊疗项目目录", Val(NVL(rsTmp!麻醉方式, 0)), "名称")
                .TextMatrix(i, colASA分级) = NVL(rsTmp!asa分级)
                .TextMatrix(i, colNNIS分级) = NVL(rsTmp!NNIS分级)
                .TextMatrix(i, col手术分级) = NVL(rsTmp!手术级别)
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
    If ErrCenter() = 1 Then
        Resume
    End If
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

Private Sub SetInfoTime(ByVal strVal As String)
    '设置时间
    Dim strTemp As String
    Dim strVval() As String
    Dim i As Integer
    Dim iCount As Integer
    txtInfo(txt入院前天).Text = ""
    txtInfo(txt入院前小时).Text = ""
    txtInfo(txt入院前分钟).Text = ""
    
    txtInfo(txt入院后天).Text = ""
    txtInfo(txt入院后小时).Text = ""
    txtInfo(txt入院后分钟).Text = ""
    
    If Len(strVal) > 0 Then
    i = InStrRev(strVal, "|", -1)
        If i > 0 Then
            strTemp = Left(strVal, i - 1)
            strVval = Split(strTemp, ",")
            For iCount = 0 To UBound(strVval)
                Select Case iCount
                Case 0
                    txtInfo(txt入院前天).Text = strVval(iCount)
                Case 1
                    txtInfo(txt入院前小时).Text = strVval(iCount)
                Case 2
                    txtInfo(txt入院前分钟).Text = strVval(iCount)
                End Select
            Next
            
            strTemp = Right(strVal, i - 1)
            strVval = Split(strTemp, ",")
            For iCount = 0 To UBound(strVval)
                Select Case iCount
                Case 0
                    txtInfo(txt入院后天).Text = strVval(iCount)
                Case 1
                    txtInfo(txt入院后小时).Text = strVval(iCount)
                Case 2
                    txtInfo(txt入院后分钟).Text = strVval(iCount)
                End Select
            Next
            
        End If
    End If
End Sub

Private Function Get治疗结果() As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    strSQL = "Select 编码,名称,简码 From 治疗结果 Order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "|" & rsTmp!编码 & "-" & rsTmp!名称
        rsTmp.MoveNext
    Loop
    If strSQL = "" Then
        Get治疗结果 = "1-治愈|2-好转|3-未愈|4-死亡|5-其他"
    Else
        Get治疗结果 = Mid(strSQL, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get诊断描述(ByVal lng诊断ID As Long, ByVal lng疾病ID As Long) As String
'功能：根据诊断ID或疾病ID获取字典表中的名称（病人诊断记录中的名称可以是修改后的,允许加前缀或后缀），以便再次修改时判断
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    If lng诊断ID <> 0 Then
        strSQL = "Select 名称 From 疾病诊断目录 Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng诊断ID)
        If rsTmp.RecordCount > 0 Then Get诊断描述 = "" & rsTmp!名称
    ElseIf lng疾病ID <> 0 Then
        strSQL = "Select 名称 From 疾病编码目录 Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng疾病ID)
        If rsTmp.RecordCount > 0 Then Get诊断描述 = "" & rsTmp!名称
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetRow(ByVal lng诊断类型 As Long) As Long
'功能：返回指定诊断类型的第一诊断行
    If InStr(",11,12,13,", "," & lng诊断类型 & ",") > 0 Then
        GetRow = vsDiagZY.FindRow(CStr(lng诊断类型), , colzy类型)
    Else
        GetRow = vsDiagXY.FindRow(CStr(lng诊断类型), , col类型)
    End If
End Function

Private Function Get病例分型(ByVal str编码 As String) As String
'功能:返回指定的病例分型
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    If str编码 <> "" Then
        strSQL = "Select 编码 || '-' ||名称 AS 名称 From 临床病例分型 Where 编码=[1] Order by 编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str编码)
        If rsTmp.RecordCount > 0 Then Get病例分型 = "" & rsTmp!名称
    Else
        Get病例分型 = ""
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

