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
   StartUpPosition =   3  '����ȱʡ
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
            Caption         =   "   ������Ϣ "
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
               Caption         =   "��Ժǰ����Ժ����"
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
               Caption         =   "��"
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
               Caption         =   "��"
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
               Caption         =   "����Ժ"
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
               Caption         =   "ת��"
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
               Caption         =   "��Ժ��ʽ"
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
               Caption         =   "����֤��"
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
               Caption         =   "����"
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
               Caption         =   "��Ժ;��"
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
               Caption         =   "�ʱ�"
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
               Caption         =   "���ڵ�ַ"
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
               Caption         =   "��������Ժ����"
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
               Caption         =   "����������"
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
               Caption         =   "Ӥ�׶�����"
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
               Caption         =   "����"
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
               Caption         =   "���"
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
               Caption         =   "��Ŀ�ģ�"
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
               Caption         =   "����Ժ�ƻ�"
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
               Caption         =   "���ʱ��"
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
               Caption         =   "סԺ����"
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
               Caption         =   "��������"
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
               Caption         =   "��������"
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
               Caption         =   "ת�����"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
               Caption         =   "��Ժʱ��"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
               Caption         =   "��Ժʱ��"
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
               Caption         =   "��ϵ�˵�ַ"
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
               Caption         =   "�绰"
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
               Caption         =   "��ϵ"
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
               Caption         =   "��ϵ������"
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
               Caption         =   "�ʱ�"
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
               Caption         =   "�绰"
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
               Caption         =   "������λ"
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
               Caption         =   "�ʱ�"
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
               Caption         =   "�绰"
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
               Caption         =   "��סַ"
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
               Caption         =   "���֤��"
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
               Caption         =   "�����ص�"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
               Caption         =   "ְҵ"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
               Caption         =   "��������"
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
               Caption         =   "�Ա�"
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
               Caption         =   "����"
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
               Caption         =   "���ʽ"
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
               Caption         =   "��    ��סԺ"
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
               Caption         =   "סԺ��"
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
            Caption         =   "   ��ҽ��� "
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
               Caption         =   "��������ʬ��"
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
               Caption         =   "�·�����"
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
               Caption         =   "����ʱ��"
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
               Caption         =   "����ԭ��"
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
               Caption         =   "����ԭ��"
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
               Caption         =   "��Ⱦ��ԭѧ���"
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
               Caption         =   "����������"
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
               Caption         =   "�ֻ��̶�"
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
               Caption         =   "�ɹ�����"
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
               Caption         =   "���ȴ���"
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
               Caption         =   "�������Ժ"
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
               Caption         =   "��Ժ���Ժ"
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
               Caption         =   "�����벡��"
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
               Caption         =   "�ٴ��벡��"
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
               Caption         =   "�ٴ���ʬ��"
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
            Caption         =   "   ���������� "
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
               Caption         =   "��"
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
               Caption         =   "��"
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
               Caption         =   "��"
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
               Caption         =   "��������"
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
               Caption         =   "��"
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
               Caption         =   "��������"
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
               Caption         =   "��"
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
               Caption         =   "һ������"
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
               Caption         =   "��"
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
               Caption         =   "�ػ�"
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
               Caption         =   "��ǰ������"
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
            Caption         =   "   ��ҽ��� "
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
               Caption         =   " ���Ʒ��� "
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
                  Caption         =   "�������"
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
                  Caption         =   "���ȷ���"
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
                  Caption         =   "������ҩ�Ƽ�"
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
               Caption         =   " סԺ�ڼ䲡�� "
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
                  Caption         =   "����"
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
                  Caption         =   "��֢"
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
                  Caption         =   "Σ��"
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
               Caption         =   " ׼ȷ�� "
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
                  Caption         =   "��ҩ"
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
                  Caption         =   "�η�"
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
                  Caption         =   "��֤"
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
               Caption         =   "��Ժ���Ժ"
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
               Caption         =   "�������Ժ"
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
            Caption         =   "   סԺ��� "
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
               Caption         =   "���в���"
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
               Caption         =   "��Ѫ��Ӧ"
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
               Caption         =   "�������"
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
               Caption         =   "�ʿ�ҽʦ"
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
               Caption         =   "�ʿػ�ʿ"
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
               Caption         =   "���λ�ʿ"
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
               Caption         =   "Сʱ"
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
               Caption         =   "������ʹ��"
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
               Caption         =   "����"
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
               Caption         =   "Сʱ"
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
               Caption         =   "��"
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
               Caption         =   "��Ժ��"
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
               Caption         =   "����"
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
               Caption         =   "Сʱ"
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
               Caption         =   "��"
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
               Caption         =   "��Ժǰ"
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
               Caption         =   "����ʱ��"
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
               Caption         =   "��������"
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
               Caption         =   "ʵϰҽʦ"
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
               Caption         =   "�о���ҽʦ"
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
               Caption         =   "����ҽʦ"
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
               Caption         =   "סԺҽʦ"
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
               Caption         =   "����ҽʦ"
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
               Caption         =   "����(������)ҽʦ"
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
               Caption         =   "������"
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
               Caption         =   "����ҽʦ"
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
               Caption         =   "������"
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
               Caption         =   "��ȫѪ"
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
               Caption         =   "��Ѫ��"
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
               Caption         =   "��λ"
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
               Caption         =   "��ѪС��"
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
               Caption         =   "��λ"
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
               Caption         =   "���ϸ��"
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
               Caption         =   "Ѫ��"
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
               Caption         =   "��Һ��Ӧ"
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

'�ϴ�ˢ������ʱ�Ĳ�����Ϣ
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng����ID As Long
Private mlng����ID As Long
Private mbln��Ժ As Boolean
Private mblnMoved As Boolean
Private mbln��ҽ As Boolean
Private mblnCheck As Boolean

Private Const ColorUnEditCell = &H8000000B  '����ɫ

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
    txt���ʱ�� = 76
    txt����Ժ�ƻ� = 77
    txt����ԺĿ�� = 78
    txt��� = 79
    txt���� = 80
    txtӤ�׶����� = 81
    txt���������� = 82
    txt��������Ժ���� = 83
    txt���ڵ�ַ = 84
    txt���ڵ�ַ�ʱ� = 85
    txt��Ժ;�� = 86
    txt���� = 87
    txt����֤�� = 88
    txt��Ժ��ʽ = 89
    txt��Ժת�� = 90
    chk��Ժ���� = 15
    
    txt�ػ� = 105
    txtһ������ = 107
    txt�������� = 108
    txt�������� = 109
    txtICU = 110
    txtCCU = 111
    
    txt��Ѫ��Ӧ = 112
End Enum
Private Enum ��ҽ���
'    chk�Ƿ�ȷ�� = 0
    txt���ȴ��� = 36
'    txtȷ������ = 37
    txt�ɹ����� = 35
    txt�������Ժ = 70
    txt��Ժ���Ժ = 68
    txt�����벡�� = 66
    txt�ٴ��벡�� = 69
    txt�ٴ���ʬ�� = 67
    txt��ǰ������ = 65
    txt�ֻ��̶� = 91
    txt���������� = 92
    txt��Ⱦ��ԭѧ��� = 93
    txt����ԭ�� = 94
    txt����ʱ�� = 95
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

''Private Enum ����������
''    chk�������� = 1
''    chk�������� = 12
''    chk������� = 13
''    chk������� = 14
''End Enum

Private Enum סԺ���
    chkʬ�� = 6
    chk���� = 7
    chk�·����� = 5
'    chkʾ�̲��� = 8
    chk���в��� = 11
'    chk��Ѫ��Ӧ = 9
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
    txt��Ժǰ�� = 46
    txt��ԺǰСʱ = 96
    txt��Ժǰ���� = 97
    
    txt��Ժ���� = 98
    txt��Ժ��Сʱ = 99
    txt��Ժ����� = 100
    txt������ʹ�� = 101
    
    txt�ʿ�ҽʦ = 102
    txt�ʿػ�ʿ = 103
    txt���λ�ʿ = 104
    
    txt�������� = 37
    txt������� = 106
End Enum

Private Enum ������
    col������� = 0
    col��ϱ��� = 1
    col������� = 2
    col��ע = 3
    col��Ժ���� = 4
    col��Ժ��� = 5
    col�Ƿ�δ�� = 6
    col�Ƿ����� = 7
    col���� = 8
    col���ID = 9
    col����ID = 10
    col���� = 11 '1-��ҽ�������;2-��ҽ��Ժ���;3-��Ժ���(�������);5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���;10-����֢
    
    colzy���� = 6
    colzy���ID = 7
    colzy����ID = 8
    colzy֤��ID = 9
    colzy���� = 10
End Enum
Private Enum �������
    col�������� = 0
    col�������� = 1
    col�ٴ����� = 2
    col����ҽʦ = 3
    col������ʿ = 4
    col����1 = 5
    col����2 = 6
    col����ʽ = 7
    colASA�ּ� = 8
    colNNIS�ּ� = 9
    col�����ּ� = 10
    col����ҽʦ = 11
    col�п����� = 12
End Enum

Public Function zlRefresh(lng����ID As Long, lng��ҳID As Long, lng����ID As Long, lng����ID As Long, Optional bln��Ժ As Boolean) As Boolean
'���ܣ�ˢ�»����ҽ���嵥
    
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    mlng����ID = lng����ID: mlng��ҳID = lng��ҳID
    mlng����ID = lng����ID: mlng����ID = lng����ID
    mbln��Ժ = bln��Ժ: mblnMoved = False
        
    strSQL = "Select ��Ժ����ID,��ǰ����ID From ������ҳ Where ����id=[1] And ��ҳid=[2]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rs.BOF = False Then
        mlng����ID = zlCommFun.NVL(rs("��Ժ����ID").Value, 0)
        mlng����ID = zlCommFun.NVL(rs("��ǰ����ID").Value, 0)
    End If
    
    '���ܿ����л�
    mbln��ҽ = Have��������(mlng����ID, "��ҽ��")
    fraInfo(2).Visible = mbln��ҽ
    fraInfo(2).Enabled = mbln��ҽ '��־������
    Call SetPageHeight
    Call SetScrollbar
    
    Call ClearPageData
    If mlng����ID <> 0 Then Call LoadPageData
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
    '�������ߴ�
    vsc.Width = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    hsc.Height = GetSystemMetrics(SM_CXHSCROLL) * Screen.TwipsPerPixelY
    fraVH.Width = vsc.Width: fraVH.Height = hsc.Height
    fraBack.Left = 0: fraBack.Top = 0
    picBack.BackColor = fraBack.BackColor
    
    optExits(0).Enabled = False
    optExits(1).Enabled = False
'    '��ʼ��ϵͳ����
'    Call InitSysPar
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
        .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
        For i = .Rows - 1 To .FixedRows Step -1
            If .TextMatrix(i, col�������) = "" Then
                .RemoveItem i
            End If
        Next
    End With
    With vsDiagZY
        .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
        For i = .Rows - 1 To .FixedRows Step -1
            If .TextMatrix(i, col�������) = "" Then
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
'���ܣ���ȡ���˵���ҳ��Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim lngRow As Long, varTmp As Variant
    Dim strTmp As String
    Dim str���ƽ�� As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    mblnCheck = True
    
    '������Ϣ����
    '---------------------------------------------------------------
    strSQL = "Select סԺ��,�Ա�,����,��������,�����ص�,���֤��,����֤��,����,���� From ������Ϣ Where ����ID=[1]"
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
    txtInfo(txt����).Text = NVL(rsTmp!����)
    txtInfo(txt����֤��).Text = NVL(rsTmp!����֤��)
    
    '������ҳ����
    '---------------------------------------------------------------
    strSQL = "Select A.����ID,A.��ҳID,A.סԺ��,A.��������,A.ҽ�Ƹ��ʽ,A.�ѱ�,A.����Ժ,A.��Ժ����ID,A.��Ժ����ID,A.��Ժ����,A.��Ժ����,A.��Ժ��ʽ,A.��Ժ����,A.����Ժת��,A.סԺĿ��,A.��Ժ����,A.�Ƿ����,A.��ǰ����,A.��ǰ����ID,A.����ȼ�ID,A.��Ժ����ID,A.��Ժ����,A.��Ժ����,A.סԺ����,A.��Ժ��ʽ,A.�Ƿ�ȷ��,A.ȷ������,A.�·�����,A.Ѫ��,A.���ȴ���,A.�ɹ�����,A.�����־,A.��������,A.ʬ���־,A.����ҽʦ,A.���λ�ʿ,A.סԺҽʦ,A.������,A.��ĿԱ���,A.��ĿԱ����,A.��Ŀ����,A.״̬,A.���ú�,A.����,A.����״��,A.ְҵ,A.����,A.ѧ��,A.��λ�绰,A.��λ�ʱ�,A.��λ��ַ,A.����,A.��ͥ��ַ,A.��ͥ�绰,A.��ͥ��ַ�ʱ�,A.��ϵ������,A.��ϵ�˹�ϵ,A.��ϵ�˵�ַ,A.��ϵ�˵绰,A.��ҽ�������,A.����,A.��˱�־,A.�����,A.�������,A.�Ƿ��ϴ�,A.����ת��,A.�Ǽ���,A.�Ǽ�ʱ��,A.��ע,A.����,A.����״̬,A.���ʱ��,A.��������,A.���,A.����,A.���ڵ�ַ,A.���ڵ�ַ�ʱ�,B.���� as ��Ժ����,C.���� as ��Ժ����" & _
        " From ������ҳ A,���ű� B,���ű� C" & _
        " Where A.��Ժ����ID=B.ID And A.��Ժ����ID=C.ID" & _
        " And A.����ID=[1] And A.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    mblnMoved = NVL(rsTmp!����ת��, 0) = 1
    
    '���۲�����סԺ��
    If NVL(rsTmp!��������, 0) <> 0 Then
        lblInfo(0).Visible = False
        txtInfo(txtסԺ��).Visible = False
    Else
        lblInfo(0).Visible = True
        txtInfo(txtסԺ��).Visible = True
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
'    chkInfo(chk�Ƿ�ȷ��).Value = NVL(rsTmp!�Ƿ�ȷ��, 0)
'    If chkInfo(chk�Ƿ�ȷ��).Value = 1 Then
'        txtInfo(txtȷ������).Text = Format(NVL(rsTmp!ȷ������), "yyyy-MM-dd HH:mm")
'    End If
    txtInfo(txt���ȴ���).Text = NVL(rsTmp!���ȴ���)
    If Val(txtInfo(txt���ȴ���).Text) <> 0 Then
        txtInfo(txt�ɹ�����).Text = NVL(rsTmp!�ɹ�����)
    End If
    chkInfo(chk�·�����).Value = NVL(rsTmp!�·�����, 0)
    
    txtInfo(txt�������).Text = NVL(rsTmp!��ҽ�������)
    chkInfo(chkʬ��).Value = NVL(rsTmp!ʬ���־, 0)
'    chkInfo(chk����).Value = IIf(NVL(rsTmp!�����־, 0) = 0, 0, 1)
'    If chkInfo(chk����).Value = 1 Then
'        txtInfo(txt��������).Text = NVL(rsTmp!��������, 0) & Decode(NVL(rsTmp!�����־, 0), 1, "��", 2, "��", 3, "��")
'    End If
    txtInfo(txt����ҽʦ).Text = NVL(rsTmp!����ҽʦ)
    txtInfo(txtסԺҽʦ).Text = NVL(rsTmp!סԺҽʦ)
    txtInfo(txtѪ��).Text = NVL(rsTmp!Ѫ��)
    
    If NVL(rsTmp!���) = "" Or NVL(rsTmp!���) = 0 Then
        txtInfo(txt���).Text = ""
    Else
        txtInfo(txt���).Text = NVL(rsTmp!���) & "cm"
    End If
    
    If NVL(rsTmp!����) = "" Or NVL(rsTmp!����) = 0 Then
        txtInfo(txt����).Text = ""
    Else
        txtInfo(txt����).Text = NVL(rsTmp!����) & "kg"
    End If
    
    txtInfo(txt���ڵ�ַ).Text = NVL(rsTmp!���ڵ�ַ)
    txtInfo(txt���ڵ�ַ�ʱ�).Text = NVL(rsTmp!���ڵ�ַ�ʱ�)
    
    txtInfo(txt��Ժ;��).Text = NVL(rsTmp!��Ժ��ʽ)
    txtInfo(txt��Ժ��ʽ).Text = NVL(rsTmp!��Ժ��ʽ)
    
    txtInfo(txt���λ�ʿ).Text = NVL(rsTmp!���λ�ʿ)
   
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
                
'''                chkInfo(chk��������).Value = 0
'''                chkInfo(chk��������).Value = 0
'''                chkInfo(chk�������).Value = 0
'''                chkInfo(chk�������).Value = 0
'''
'''                strTmp = NVL(rsTmp("��Ϣֵ").Value, "0000")
'''                If Len(strTmp) >= 1 Then chkInfo(chk��������).Value = Val(Mid(strTmp, 1, 1))
'''                If Len(strTmp) >= 2 Then chkInfo(chk��������).Value = Val(Mid(strTmp, 2, 1))
'''                If Len(strTmp) >= 3 Then chkInfo(chk�������).Value = Val(Mid(strTmp, 3, 1))
'''                If Len(strTmp) >= 4 Then chkInfo(chk�������).Value = Val(Mid(strTmp, 4, 1))
                
                                
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
'                chkInfo(chkʾ�̲���).Value = Val(NVL(rsTmp!��Ϣֵ, 0))
            Case "���в���"
                chkInfo(chk���в���).Value = Val(NVL(rsTmp!��Ϣֵ, 0))
            Case UCase("Rh")
                txtInfo(txtRh).Text = NVL(rsTmp!��Ϣֵ)
            Case "��Ѫ��Ӧ"
                txtInfo(txt��Ѫ��Ӧ).Text = Decode(Val(NVL(rsTmp!��Ϣֵ, 0)), 0, "��", 1, "��", 2, "δ��")
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
            Case "����Ժ�ƻ�����"
                If NVL(rsTmp!��Ϣֵ, 0) = 0 Then
                    txtInfo(txt����Ժ�ƻ�).Text = "31����סԺ�ƻ�"
                Else
                    txtInfo(txt����Ժ�ƻ�).Text = "7����סԺ�ƻ�"
                End If
            Case "31������סԺ"
                If NVL(rsTmp!��Ϣֵ) = "" Then
                    optExits(0).Value = True
                    txtInfo(txt����ԺĿ��).Text = ""
                Else
                    optExits(1).Value = True
                    txtInfo(txt����ԺĿ��).Text = NVL(rsTmp!��Ϣֵ)
                End If
            Case "������������"
                txtInfo(txtӤ�׶�����).Text = NVL(rsTmp!��Ϣֵ)
            Case "��������������"
                If NVL(rsTmp!��Ϣֵ) = "" Then
                    txtInfo(txt����������).Text = ""
                Else
                    txtInfo(txt����������).Text = NVL(rsTmp!��Ϣֵ) & "��"
                End If
            Case "��������Ժ����"
                If NVL(rsTmp!��Ϣֵ) = "" Then
                    txtInfo(txt��������Ժ����).Text = ""
                Else
                    txtInfo(txt��������Ժ����).Text = NVL(rsTmp!��Ϣֵ) & "��"
                End If
            Case "����"
                '����Ѿ����˼���,�Ͳ��ڴӴӱ��ж�ȡ��
                If txtInfo(txt����).Text = "" Then
                    txtInfo(txt����).Text = NVL(rsTmp!��Ϣֵ)
                End If
'            Case "����֤��"
'                txtInfo(txt����֤��).Text = NVL(rsTmp!��Ϣֵ)
            Case "��Ժת��"
                txtInfo(txt��Ժת��).Text = NVL(rsTmp!��Ϣֵ)
            Case "��Ժǰ����Ժ����"
                chkInfo(chk��Ժ����).Value = Val(NVL(rsTmp!��Ϣֵ, 0))
            Case "���Ȳ���"
                txtInfo(txt����ԭ��).Text = NVL(rsTmp!��Ϣֵ)
            Case "����ʱ��"
                txtInfo(txt����ʱ��).Text = NVL(rsTmp!��Ϣֵ)
            Case "����ʱ��"
                Call SetInfoTime(NVL(rsTmp!��Ϣֵ))
            Case "������ʹ��ʱ��"
                txtInfo(txt������ʹ��).Text = NVL(rsTmp!��Ϣֵ)
            Case "�ʿ�ҽʦ"
                txtInfo(txt�ʿ�ҽʦ).Text = NVL(rsTmp!��Ϣֵ)
            Case "�ʿػ�ʿ"
                txtInfo(txt�ʿػ�ʿ).Text = NVL(rsTmp!��Ϣֵ)
            Case "��������"
                txtInfo(txt��������).Text = Get��������(NVL(rsTmp!��Ϣֵ))
            Case "�ֻ��̶�"
                txtInfo(txt�ֻ��̶�).Text = NVL(rsTmp!��Ϣֵ)
            Case "����������"
                txtInfo(txt����������).Text = NVL(rsTmp!��Ϣֵ)
            Case "�������"
                txtInfo(txt�������).Text = NVL(rsTmp!��Ϣֵ)
            Case "�ؼ���������"
                txtInfo(txt�ػ�).Text = NVL(rsTmp!��Ϣֵ)
            Case "һ����������"
                txtInfo(txtһ������).Text = NVL(rsTmp!��Ϣֵ)
            Case "������������"
                txtInfo(txt��������).Text = NVL(rsTmp!��Ϣֵ)
            Case "������������"
                txtInfo(txt��������).Text = NVL(rsTmp!��Ϣֵ)
            Case "ICU����"
                txtInfo(txtICU).Text = NVL(rsTmp!��Ϣֵ)
            Case "CCU����"
                txtInfo(txtCCU).Text = NVL(rsTmp!��Ϣֵ)
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
    '��ȡ���ʱ��
    strSQL = _
            " Select A.��ʼʱ��" & _
            " From ���˱䶯��¼ A,���ű� B" & _
            " Where A.����ID=[1] And A.��ҳID=[2]" & _
            " And A.����ID=B.ID And A.��ʼԭ��=2 And A.��ʼʱ�� is Not NULL" & _
            " Order by A.��ʼʱ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTmp Is Nothing Then
        txtInfo(txt���ʱ��).Text = txtInfo(txt��Ժʱ��).Text
    ElseIf rsTmp.EOF Or rsTmp.BOF Then
        txtInfo(txt���ʱ��).Text = txtInfo(txt��Ժʱ��).Text
    Else
        txtInfo(txt���ʱ��).Text = Format("" & rsTmp!��ʼʱ��, "yyyy-mm-dd hh:mm")
    End If
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
   str���ƽ�� = Get���ƽ��
    vsDiagXY.ColData(col��Ժ���) = str���ƽ��
  
    'ȱʡ����ʼ��
    With vsDiagXY
        '1-��ҽ�������;2-��ҽ��Ժ���;3-��Ժ���(�������);5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���;10-����֢
        .TextMatrix(1, col����) = 1
        .TextMatrix(2, col����) = 2
        .TextMatrix(3, col����) = 3
        .TextMatrix(4, col����) = 3
        .TextMatrix(5, col����) = 5
        .TextMatrix(6, col����) = 10
        .TextMatrix(7, col����) = 6
        .TextMatrix(8, col����) = 7
    End With
    
    '��ȡ������Դ�����
    strSQL = "Select a.��ע,a.ID,a.����ID,a.��ҳID,a.ҽ��ID,a.��¼��Դ,a.��ϴ���,a.�������,a.����ID,a.�������,a.����ID,a.��Ժ����," & _
        " a.���ID,a.֤��ID,a.�������,a.��Ժ���,a.�Ƿ�δ��,a.�Ƿ�����,a.��¼����,a.��¼��,a.ȡ��ʱ��,a.ȡ����,a.����ID, b.���� As ��������, c.���� As ��ϱ��� " & _
        " From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C" & _
        " Where a.����id = b.Id(+) And a.���id = c.Id(+) And a.��¼��Դ IN(1,2,3,4) And a.������� IN(1,2,3,5,6,7,10,21)" & _
        " And a.ȡ��ʱ�� is Null And a.����ID=[1] And a.��ҳID=[2]" & _
        " Order by a.�������,a.��¼��Դ Desc,a.��ϴ���,a.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then
        With vsDiagXY
            strSQL = "1,2,3,5,6,7,10,21"
            For i = 0 To UBound(Split(strSQL, ","))
                rsTmp.Filter = "��¼��Դ=3 And �������=" & Split(strSQL, ",")(i)
                If Val(Split(strSQL, ",")(i)) <> 21 Then
                    If rsTmp.EOF Then
                        rsTmp.Filter = "��¼��Դ=2 And �������=" & Split(strSQL, ",")(i)
                    End If
                    If rsTmp.EOF Then
                        rsTmp.Filter = "��¼��Դ=1 And �������=" & Split(strSQL, ",")(i)
                    End If
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=4 And �������=" & Split(strSQL, ",")(i)
                End If
                
                If Val(Split(strSQL, ",")(i)) = 21 Then
                    '21-��ԭѧ���
                    If Not rsTmp.EOF Then
                        txtInfo(txt��Ⱦ��ԭѧ���).Text = NVL(rsTmp!�������)
                        txtInfo(txt��Ⱦ��ԭѧ���).Tag = txtInfo(txt��Ⱦ��ԭѧ���).Text
                    End If
                Else
                    Do While Not rsTmp.EOF
                        'ȷ����ǰ��ʾ��
                        lngRow = .FindRow(CStr(Split(strSQL, ",")(i)), , col����)
                        For j = lngRow To .Rows - 1
                            If Val(.TextMatrix(j, col����)) = Val(Split(strSQL, ",")(i)) Then
                                lngRow = j
                                If .TextMatrix(j, col�������) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, col�������) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, col����) = Split(strSQL, ",")(i)
                        End If
                        
                        If IsNull(rsTmp!�������) Then
                            .TextMatrix(lngRow, col��ϱ���) = ""
                            .TextMatrix(lngRow, col�������) = ""
                        Else
                            If Mid(rsTmp!�������, 1, 1) <> "(" Or (Val(rsTmp!���id & "") = 0 And Val(rsTmp!����id & "") = 0) Then '��ҽ���������������ˣ���֢��������ֻ�жϵ�һ���ַ�
                                '���ڼ����������Ͽ��Զ�Ӧ�������������Ϊ�յ�ʱ�����жϼ������룬��ȡ��������
                                If Val(rsTmp!����id & "") <> 0 Then
                                    .TextMatrix(lngRow, col��ϱ���) = NVL(rsTmp!��������)
                                ElseIf Val(rsTmp!���id & "") <> 0 Then
                                    .TextMatrix(lngRow, col��ϱ���) = NVL(rsTmp!��ϱ���)
                                Else
                                    .TextMatrix(lngRow, col��ϱ���) = ""
                                End If
                                .TextMatrix(lngRow, col�������) = rsTmp!�������
                            Else
                                .TextMatrix(lngRow, col��ϱ���) = Mid(rsTmp!�������, 2, InStr(rsTmp!�������, ")") - 2)
                                .TextMatrix(lngRow, col�������) = Mid(rsTmp!�������, InStr(rsTmp!�������, ")") + 1)
                            End If
                        End If
                        If Not IsNull(rsTmp!����id) Or Not IsNull(rsTmp!���id) Then
                            .Cell(flexcpData, lngRow, col�������) = Get�������(Val("" & rsTmp!���id), Val("" & rsTmp!����id))    '��ȡԭʼ�����Ա��޸�ʱ�ж�
                        Else
                            .Cell(flexcpData, lngRow, col�������) = .TextMatrix(lngRow, col�������)
                        End If
                        
                        .TextMatrix(lngRow, col��ע) = NVL(rsTmp!��ע)
                        .TextMatrix(lngRow, col��Ժ���) = NVL(rsTmp!��Ժ���)
                        .TextMatrix(lngRow, col��Ժ����) = NVL(rsTmp!��Ժ����)
                        .TextMatrix(lngRow, col�Ƿ�δ��) = IIf(NVL(rsTmp!�Ƿ�δ��, 0) = 1, "��", "")
                        .TextMatrix(lngRow, col�Ƿ�����) = IIf(NVL(rsTmp!�Ƿ�����, 0) = 1, "��", "")
                        .TextMatrix(lngRow, col���ID) = NVL(rsTmp!���id, 0)
                        .TextMatrix(lngRow, col����ID) = NVL(rsTmp!����id, 0)
                        rsTmp.MoveNext
                    Loop
                End If
            Next
        End With
    End If
    
    vsDiagXY.Cell(flexcpForeColor, 1, col�Ƿ�����, vsDiagXY.Rows - 1, col�Ƿ�����) = vbRed
    vsDiagXY.Cell(flexcpBackColor, GetRow(3), vsDiagXY.FixedRows, GetRow(3), vsDiagXY.Cols - 1) = &HC0FFC0
    vsDiagXY.Cell(flexcpBackColor, 1, col��ϱ���, vsDiagXY.Rows - 1, col��ϱ���) = ColorUnEditCell      '����ɫ
    vsDiagXY.Row = 1: vsDiagXY.Col = col�������
    Call vsDiagXY_AfterRowColChange(-1, -1, vsDiagXY.Row, vsDiagXY.Col)
        
    '��ҽ���
    '---------------------------------------------------------------
    strSQL = "Select ȡ����,����ID,��ע,ID,����ID,��ҳID,ҽ��ID,��¼��Դ,��ϴ���,�������,����ID,�������,����ID,���ID,֤��ID,�������,��Ժ���,�Ƿ�δ��,�Ƿ�����,��¼����,��¼��,ȡ��ʱ��,��Ժ���� From ������ϼ�¼" & _
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
                    .TextMatrix(lngRow, col��ע) = NVL(rsTmp!��ע)
                    .TextMatrix(lngRow, col��Ժ����) = NVL(rsTmp!��Ժ����)
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
    strSQL = "Select ��¼��Դ,��������,������ʼʱ��,��������ʱ��,��������,��������ID,������ĿID,��������,����ҽʦ,�ٴ�����,��һ����,�ڶ�����,������ʿ,����ʼʱ��,�������ʱ��,����ʽ,��������,��������,��Һ����,����ҽʦ,������ʼʱ��,��������ʱ��,�п�,����,��¼����,��¼��,ȡ��ʱ��,ȡ����,������ʿ,ID,����ID,��ҳID,decode(ASA�ּ�,'I��','P1','II��','P2','III��','P3','IV��','P4','V��','P5',ASA�ּ�) as ASA�ּ�,NNIS�ּ�,decode(��������,1,'һ������',2,'��������',3,'��������',4,'�ļ�����',' ') as �������� From ���������¼" & _
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
                .TextMatrix(i, col�ٴ�����) = NVL(rsTmp!�ٴ�����, -1)
                .TextMatrix(i, col����ҽʦ) = NVL(rsTmp!����ҽʦ)
                .TextMatrix(i, col������ʿ) = NVL(rsTmp!������ʿ)
                .TextMatrix(i, col����1) = NVL(rsTmp!��һ����)
                .TextMatrix(i, col����2) = NVL(rsTmp!�ڶ�����)
                .TextMatrix(i, col����ʽ) = GetItemField("������ĿĿ¼", Val(NVL(rsTmp!����ʽ, 0)), "����")
                .TextMatrix(i, colASA�ּ�) = NVL(rsTmp!asa�ּ�)
                .TextMatrix(i, colNNIS�ּ�) = NVL(rsTmp!NNIS�ּ�)
                .TextMatrix(i, col�����ּ�) = NVL(rsTmp!��������)
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
    '����ʱ��
    Dim strTemp As String
    Dim strVval() As String
    Dim i As Integer
    Dim iCount As Integer
    txtInfo(txt��Ժǰ��).Text = ""
    txtInfo(txt��ԺǰСʱ).Text = ""
    txtInfo(txt��Ժǰ����).Text = ""
    
    txtInfo(txt��Ժ����).Text = ""
    txtInfo(txt��Ժ��Сʱ).Text = ""
    txtInfo(txt��Ժ�����).Text = ""
    
    If Len(strVal) > 0 Then
    i = InStrRev(strVal, "|", -1)
        If i > 0 Then
            strTemp = Left(strVal, i - 1)
            strVval = Split(strTemp, ",")
            For iCount = 0 To UBound(strVval)
                Select Case iCount
                Case 0
                    txtInfo(txt��Ժǰ��).Text = strVval(iCount)
                Case 1
                    txtInfo(txt��ԺǰСʱ).Text = strVval(iCount)
                Case 2
                    txtInfo(txt��Ժǰ����).Text = strVval(iCount)
                End Select
            Next
            
            strTemp = Right(strVal, i - 1)
            strVval = Split(strTemp, ",")
            For iCount = 0 To UBound(strVval)
                Select Case iCount
                Case 0
                    txtInfo(txt��Ժ����).Text = strVval(iCount)
                Case 1
                    txtInfo(txt��Ժ��Сʱ).Text = strVval(iCount)
                Case 2
                    txtInfo(txt��Ժ�����).Text = strVval(iCount)
                End Select
            Next
            
        End If
    End If
End Sub

Private Function Get���ƽ��() As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    strSQL = "Select ����,����,���� From ���ƽ�� Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "|" & rsTmp!���� & "-" & rsTmp!����
        rsTmp.MoveNext
    Loop
    If strSQL = "" Then
        Get���ƽ�� = "1-����|2-��ת|3-δ��|4-����|5-����"
    Else
        Get���ƽ�� = Mid(strSQL, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get�������(ByVal lng���ID As Long, ByVal lng����ID As Long) As String
'���ܣ��������ID�򼲲�ID��ȡ�ֵ���е����ƣ�������ϼ�¼�е����ƿ������޸ĺ��,�����ǰ׺���׺�����Ա��ٴ��޸�ʱ�ж�
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    If lng���ID <> 0 Then
        strSQL = "Select ���� From �������Ŀ¼ Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng���ID)
        If rsTmp.RecordCount > 0 Then Get������� = "" & rsTmp!����
    ElseIf lng����ID <> 0 Then
        strSQL = "Select ���� From ��������Ŀ¼ Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        If rsTmp.RecordCount > 0 Then Get������� = "" & rsTmp!����
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetRow(ByVal lng������� As Long) As Long
'���ܣ�����ָ��������͵ĵ�һ�����
    If InStr(",11,12,13,", "," & lng������� & ",") > 0 Then
        GetRow = vsDiagZY.FindRow(CStr(lng�������), , colzy����)
    Else
        GetRow = vsDiagXY.FindRow(CStr(lng�������), , col����)
    End If
End Function

Private Function Get��������(ByVal str���� As String) As String
'����:����ָ���Ĳ�������
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    If str���� <> "" Then
        strSQL = "Select ���� || '-' ||���� AS ���� From �ٴ��������� Where ����=[1] Order by ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����)
        If rsTmp.RecordCount > 0 Then Get�������� = "" & rsTmp!����
    Else
        Get�������� = ""
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

