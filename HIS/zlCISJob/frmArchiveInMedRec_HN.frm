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
   StartUpPosition =   3  '����ȱʡ
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
            Caption         =   "   סԺ��� "
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
               Caption         =   "���Ѳ���"
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
               Text            =   "��"
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
               Caption         =   "ʾ�̲���"
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
               Caption         =   "����"
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
               Caption         =   "���в���"
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
               Caption         =   "����ҽѧ��ʾ"
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
               Caption         =   "ҽѧ��ʾ"
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
               Caption         =   "��������"
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
               Caption         =   "�ʿ�����"
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
               Caption         =   "ת�����"
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
               Caption         =   "��Ժ��ʽ"
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
               Caption         =   "����״��"
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
               Caption         =   "����ʱ��"
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
               Caption         =   "��Ѫǰ9����"
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
               Caption         =   "��������"
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
               Caption         =   "��Ժ31��������Ժ�ƻ�"
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
               Caption         =   "Ŀ��"
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
               Caption         =   "�������"
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
               Caption         =   "������ʹ��"
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
               Caption         =   "Сʱ"
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
               Caption         =   "����"
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
               Caption         =   "Сʱ"
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
               Caption         =   "��"
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
               Caption         =   "��Ժ��"
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
               Caption         =   "����"
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
               Caption         =   "��"
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
               Caption         =   "­�����˻��߻���ʱ��;   ��Ժǰ"
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
               Caption         =   "���λ�ʿ"
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
               Caption         =   "�ʿ�ҽʦ"
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
               Caption         =   "�ʿػ�ʿ"
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
               Caption         =   "��Ѫ��Ӧ"
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
               Caption         =   "ʵϰҽʦ"
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
               Caption         =   "�о���ҽʦ"
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
               Caption         =   "����ҽʦ"
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
               Caption         =   "סԺҽʦ"
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
               Caption         =   "����ҽʦ"
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
               Caption         =   "����(������)ҽʦ"
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
               Caption         =   "������"
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
               Caption         =   "����ҽʦ"
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
               Caption         =   "������"
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
               Caption         =   "��ȫѪ"
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
               Caption         =   "��Ѫ��"
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
               Caption         =   "��λ"
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
               Caption         =   "��ѪС��"
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
               Caption         =   "��λ"
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
               Caption         =   "���ϸ��"
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
               Caption         =   "Ѫ��"
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
               Caption         =   "��������"
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
               Caption         =   "��Һ��Ӧ"
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
               Caption         =   "Сʱ"
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
            Caption         =   "   ���� "
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
               Caption         =   "����"
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
                  Text            =   "��"
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
                  Caption         =   "ʹ�ÿ�����"
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
                  Caption         =   "ϸ�������걾�ͼ�"
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
                  Caption         =   "�����ֹ���"
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
                  Caption         =   "ʵʩ&DRGs����"
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
                  Caption         =   "������Ⱦ��"
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
                  Caption         =   "ʵʩ�ٴ�·������"
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
                  Caption         =   "Сʱ"
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
                  Caption         =   "��"
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
                  Caption         =   "��"
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
                  Caption         =   "ʵʩ��֢�໤"
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
                  Caption         =   "��������"
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
                  Caption         =   "������&Apgar����       ��"
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
               Caption         =   "3.��ɫ������"
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
               Caption         =   "�����¼�"
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
                  Caption         =   "����"
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
                  Caption         =   "ѹ�������ڼ�"
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
                  Caption         =   "������׹��ԭ��"
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
                  Caption         =   "������׹���˺�"
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
               Caption         =   "��Ⱦ����"
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
               Caption         =   "���������"
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
               Caption         =   "����������Ŀ"
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
            Caption         =   "   �����뻯�� "
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
               Caption         =   "���Ƽ�¼��Ϣ"
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
               Caption         =   "���Ƽ�¼��Ϣ"
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
            Caption         =   "   ���������� "
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
            Caption         =   "   ��ҽ��� "
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
               Caption         =   " ���Ʒ��� "
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
                  Caption         =   "��֤ʩ��"
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
                  Caption         =   "ʹ����ҽ���Ƽ���"
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
                  Caption         =   "ʹ����ҽ�����豸"
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
                  Caption         =   "�������"
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
                  Caption         =   "���ȷ���"
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
                  Caption         =   "������ҩ�Ƽ�"
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
               Caption         =   " סԺ�ڼ䲡�� "
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
                  Caption         =   "����"
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
                  Caption         =   "��֢"
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
                  Caption         =   "Σ��"
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
               Caption         =   " ׼ȷ�� "
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
                  Caption         =   "��ҩ"
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
                  Caption         =   "�η�"
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
                  Caption         =   "��֤"
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
               Caption         =   "��Ժ���Ժ"
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
               Caption         =   "�������Ժ"
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
            Caption         =   "   ��ҽ��� "
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
               Caption         =   "ҽԺ��Ⱦ����ԭѧ���"
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
               Caption         =   "�·�����"
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
               Caption         =   "��������ʬ��"
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
               Caption         =   "�Ƿ�ȷ��"
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
               Caption         =   "��ǰ������"
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
               Caption         =   "ҽԺ��Ⱦ��ԭѧ���"
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
               Caption         =   "����ʱ��"
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
               Caption         =   "��������Ժ"
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
               Caption         =   "����������"
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
               Caption         =   "�ֻ��̶�"
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
               Caption         =   "��Ժ���"
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
               Caption         =   "�����"
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
               Caption         =   "����ԭ��"
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
               Caption         =   "����ԭ��"
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
               Caption         =   "�ɹ�����"
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
               Caption         =   "���ȴ���"
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
               Caption         =   "ȷ������"
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
               Caption         =   "�������Ժ"
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
               Caption         =   "��Ժ���Ժ"
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
               Caption         =   "�����벡��"
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
               Caption         =   "�ٴ��벡��"
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
               Caption         =   "�ٴ���ʬ��"
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
            Caption         =   "   ������Ϣ "
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
               Caption         =   "����Ժ"
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
               Caption         =   "��Ժǰ����Ժ����"
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
               Caption         =   "����֤��"
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
               Caption         =   "���      cm"
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
               Caption         =   "����      kg"
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
               Caption         =   "����"
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
               Caption         =   "���ʱ��"
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
               Caption         =   "���ڵ�ַ"
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
               Caption         =   "�ʱ�"
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
               Caption         =   "����"
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
               Caption         =   "����������"
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
               Caption         =   "��������Ժ����"
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
               Caption         =   "�����䲻��һ����ģ� ����"
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
               Caption         =   "������"
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
               Caption         =   "סԺ����"
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
               Caption         =   "��������"
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
               Caption         =   "��������"
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
               Caption         =   "ת�����"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
               Caption         =   "��Ժʱ��"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
               Caption         =   "��Ժʱ��"
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
               Caption         =   "��ϵ�˵�ַ"
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
               Caption         =   "�绰"
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
               Caption         =   "��ϵ"
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
               Caption         =   "��ϵ������"
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
               Caption         =   "�ʱ�"
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
               Caption         =   "�绰"
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
               Caption         =   "������λ"
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
               Caption         =   "�ʱ�"
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
               Caption         =   "�绰"
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
               Caption         =   "��סַ"
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
               Caption         =   "���֤��"
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
               Caption         =   "�����ص�"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
               Caption         =   "��Ժ;��"
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
               Caption         =   "ְҵ"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
               Caption         =   "��������"
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
               Caption         =   "�Ա�"
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
               Caption         =   "����"
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
               Caption         =   "ҽ�Ƹ��ѷ�ʽ"
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
               Caption         =   "��    ��סԺ"
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
               Caption         =   "��������"
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

'˵����Ϊ�˱��ֽ���Ŀ�ά���ԣ��������ؼ�ʱ��ע�Ᵽ��ÿ����Ϣ��Ŀ������lblInfo��linInfo,txtinfo ��index��ͬ��
'      ��������Ϣ��Ŀ����2��lblinfo������һ��lblinfo��indexΪtxtinfo.index+100

'�ϴ�ˢ������ʱ�Ĳ�����Ϣ
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mblnMoved As Boolean
Private mblnCheck As Boolean
Private mbln���� As Boolean

Private Enum Fra�˵�
    FRA_������Ϣ = 0
    FRA_��ҽ��� = 1
    FRA_��ҽ��� = 2
    FRA_���������� = 3
    FRA_סԺ��� = 4
    FRA_�����뻯�� = 5
    FRA_���� = 6
End Enum

Private Enum ������Ϣ
    txt���ʽ = 0
    txt�������� = 1
    txtסԺ���� = 2
    chk����Ժ = 0
    txt������ = 3
    txt���� = 4
    txt�Ա� = 5
    txt�������� = 6
    txt���� = 7
    txt���� = 8
    txt���� = 9
    txt��� = 10
    txt������������ = 11
    txt���������� = 12
    txt��������Ժ���� = 13
    txt�����ص� = 14
    txt���� = 15
    txt���� = 16
    txt���֤�� = 17
    txtְҵ = 18
    txt���� = 19
    txt��ͥ��ַ = 20
    txt��ͥ�绰 = 21
    txt��ͥ�ʱ� = 22
    txt���ڵ�ַ = 23
    txt�����ʱ� = 24
    txt������λ = 25
    txt��λ�绰 = 26
    txt��λ�ʱ� = 27
    txt��ϵ������ = 28
    txt��ϵ�˹�ϵ = 29
    txt��ϵ�˵绰 = 30
    txt��ϵ�˵�ַ = 31
    txt���� = 32
    txt��Ժ;�� = 33
    txt��Ժʱ�� = 34
    txt��Ժ���� = 35
    txt��Ժ���� = 36
    chk��Ժǰ����Ժ���� = 1
    txt���ʱ�� = 37
    txtת��1 = 38
    txtת��2 = 39
    txtת��3 = 40
    txt��Ժʱ�� = 41
    txt��Ժ���� = 42
    txt��Ժ���� = 43
    txtסԺ���� = 44
    txt����֤�� = 122
End Enum

Private Enum ��ҽ���
    txt��Ժ��� = 45
    chk�Ƿ�ȷ�� = 2
    txtȷ������ = 46
    txt����� = 47
    txt�ֻ��̶� = 48
    txt���������� = 49
    txt�����벡�� = 50
    txt�������Ժ = 51
    txt��Ժ���Ժ = 52
    txt��������Ժ = 53
    txt�ٴ��벡�� = 54
    txt�ٴ���ʬ�� = 55
    txt��ǰ������ = 56
    txt����ʱ�� = 57
    txt����ԭ�� = 58
    chkҽԺ��Ⱦ����ԭѧ��� = 3
    chk��������ʬ�� = 4
    chk�·����� = 5
    txtҽԺ��Ⱦ��ԭѧ��� = 59
    txt���ȴ��� = 60
    txt�ɹ����� = 61
    txt����ԭ�� = 62
End Enum

Private Enum ��ҽ���
    txt��ҽ�������Ժ = 63
    txt��ҽ��Ժ���Ժ = 64
    chkΣ�� = 6
    chk��֢ = 7
    chk���� = 8
    txt��֤ = 65
    txt�η� = 66
    txt��ҩ = 67
    txt������� = 68
    txt���ȷ��� = 69
    txt������ҩ = 70
    txt��ҽ�豸 = 71
    txt��ҽ���� = 72
    txt��֤ʩ�� = 73
End Enum

Private Enum סԺ���
    txt�������� = 74
    txtHBsAg = 75
    txt��Ѫǰ9���� = 76
    txtѪ�� = 77
    txtHCVAb = 78
    txt����ʱ�� = 79
    txtRh = 80
    txtHIVAb = 81
    txt����״�� = 82
    txt��Һ��Ӧ = 83
    txt��Ѫ��Ӧ = 84
    chkʾ�̲��� = 13
    chk���в��� = 14
    chk���Ѳ��� = 9
    txt���ϸ�� = 85
    txt��ѪС�� = 86
    txt��Ѫ�� = 87
    txt��ȫѪ = 88
    txt������ = 89
    txt������� = 90
    txtҽѧ��ʾ = 91
    txt����ҽѧ��ʾ = 92
    txt��Ժ��ʽ = 93
    txtת����� = 94
    txt��Ժǰ�� = 95
    txt��ԺǰСʱ = 96
    txt��Ժǰ���� = 97
    txt��Ժ���� = 98
    txt��Ժ��Сʱ = 99
    txt��Ժ����� = 100
    txt����Ժ���� = 101
    txt31��Ŀ�� = 102
    txt������Сʱ = 103
    chk���� = 15
    txt�������� = 104
    txt����ҽʦ = 105
    txt������ = 106
    txt����ҽʦ = 107
    txt����ҽʦ = 108
    txtסԺҽʦ = 109
    txt����ҽʦ = 110
    txt�о���ҽʦ = 111
    txtʵϰҽʦ = 112
    txt�ʿ�ҽʦ = 113
     txt���λ�ʿ = 114
    txt�ʿػ�ʿ = 115
    txt�ʿ����� = 116
    txt�������� = 117
End Enum

Private Enum ������
    col������� = 0
    col������� = 1
    col��ҽ֤�� = 2
    col��ע = 3
    col��Ժ���� = 4
    col��Ժ��� = 5
    colzy���� = 6
    col�Ƿ�δ�� = 6
    col�Ƿ����� = 7
    col���� = 8
End Enum

Private Enum �������
    col�������� = 0
    COL������� = 1
    col�������� = 2
    col�ٴ����� = 3
    col����ҽʦ = 4
    col������ʿ = 5
    col����1 = 6
    col����2 = 7
    col����ʽ = 8
    colASA�ּ� = 9
    colNNIS�ּ� = 10
    col�����ּ� = 11
    col����ҽʦ = 12
    col�п����� = 13
End Enum

Private Enum �������
    col����ʱ�� = 0
    col����ҩ�� = 1
    col������Ӧ = 2
End Enum

Private Enum ���Ƽ�¼
    col���Ʊ��� = 0
    COL���ƿ�ʼ���� = 1
    col���ƽ������� = 2
    col�����Ƴ��� = 3
    col���Ʒ��� = 4
    col�������� = 5
    col����Ч�� = 6
End Enum

Private Enum ���Ƽ�¼
    col���Ʊ��� = 0
    COL���ƿ�ʼ���� = 1
    col���ƽ������� = 2
    col��Ұ��λ = 3
    col������� = 4
    col�����ۼƼ��� = 5
    col����Ч�� = 6
End Enum

Private Enum ����
     chkCT = 16
     chkMRI = 17
     chk��ɫ������ = 18
     txtѹ�������ڼ� = 118
     txtѹ������ = 119
     txt������׹���˺� = 120
     txt������׹��ԭ�� = 121
     txt��֢�໤ = 123
     txt��֢�໤���� = 124
     txt��֢�໤Сʱ = 125
     txt�������� = 126
     txt��������T = 127
     txt��������M = 128
     txt��������N = 129
     chkʹ�ÿ����� = 21
     chk������ = 20
     chkϸ�������걾�ͼ� = 19
     txtApgar = 130
     txt�ٴ�·������ = 131
     txt��Ⱦ�� = 132
     txtDrGs���� = 133
End Enum


Public Function zlRefresh(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal blnMoved As Boolean) As Boolean
'���ܣ�ˢ�»����ҽ���嵥
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, lng����ID As Long
    Dim bln��ҽ As Boolean
    
    mlng����ID = lng����ID: mlng��ҳID = lng��ҳID: mblnMoved = blnMoved
    
    On Error GoTo errH
    
    StrSQL = "Select ��Ժ����ID From ������ҳ Where ����id=[1] And ��ҳid=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then lng����ID = Nvl(rsTmp!��Ժ����ID, 0)
    bln��ҽ = Have��������(lng����ID, "��ҽ��")
    fraInfo(FRA_��ҽ���).Visible = bln��ҽ
    fraInfo(FRA_��ҽ���).Enabled = bln��ҽ
    mbln���� = CheckShare(300) '����ϵͳ
    fraInfo(FRA_�����뻯��).Visible = mbln����
    fraInfo(FRA_�����뻯��).Enabled = mbln����
    
    Call SetPageHeight
    Call SetScrollbar
    
    Call ClearPageData
    If mlng����ID <> 0 Then Call LoadPageData
    
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
    '�������ߴ�
    vsc.Width = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    hsc.Height = GetSystemMetrics(SM_CXHSCROLL) * Screen.TwipsPerPixelY
    fraVH.Width = vsc.Width: fraVH.Height = hsc.Height
    fraBack.Left = 0: fraBack.Top = 0
    picBack.BackColor = fraBack.BackColor
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

Private Function GetRow(ByVal lng������� As Long) As Long
'���ܣ�����ָ��������͵ĵ�һ�����
    If InStr(",11,12,13,", "," & lng������� & ",") > 0 Then
        GetRow = vsDiagZY.FindRow(CStr(lng�������), , colzy����)
    Else
        GetRow = vsDiagXY.FindRow(CStr(lng�������), , col����)
    End If
End Function

Private Function LoadPageData() As Boolean
'���ܣ���ȡ���˵���ҳ��Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Long, j As Long
    Dim lngRow As Long, varTmp As Variant
    Dim strTmp As String
    Dim bln��ҳ��� As Boolean, bln�ֻ��̶� As Boolean
    Dim blnѹ�� As Boolean, bln����׹�� As Boolean
    
    On Error GoTo errH

    Screen.MousePointer = 11
    mblnCheck = True
    
    '��ʼ������������Ŀ
    Call FillVsf
    
    '������Ϣ����
    '---------------------------------------------------------------
    StrSQL = "Select סԺ��,����,�Ա�,��������,�����ص�,���֤��,����֤��,����,����,������,���� From ������Ϣ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID)

    txtInfo(txt��������).Text = Nvl(rsTmp!������)
    txtInfo(txtסԺ����).Text = mlng��ҳID
    txtInfo(txt����).Text = Nvl(rsTmp!����)
    txtInfo(txt�Ա�).Text = Nvl(rsTmp!�Ա�)
    txtInfo(txt����Ժ����).Text = "��"
    If Format(rsTmp!��������, "HH:mm") <> "00:00" Then
        txtInfo(txt��������).Text = Format(rsTmp!��������, "yyyy-MM-dd HH:mm")
    Else
        txtInfo(txt��������).Text = Format(rsTmp!��������, "yyyy-MM-dd")
    End If

    txtInfo(txt�����ص�).Text = Nvl(rsTmp!�����ص�)
    txtInfo(txt����).Text = Nvl(rsTmp!����)
    txtInfo(txt���֤��).Text = Nvl(rsTmp!���֤��)
    txtInfo(txt����).Text = Nvl(rsTmp!����)
    txtInfo(txt����).Text = Nvl(rsTmp!����)
    txtInfo(txt����֤��).Text = Nvl(rsTmp!����֤��)
    '�����Ŷ�ȡ
    StrSQL = "select ������ from סԺ������¼ where ����ID=[1] and ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTmp.RecordCount <> 0 Then
        txtInfo(txt������).Text = Nvl(rsTmp!������)
    End If
    '������ҳ����
    '---------------------------------------------------------------
    StrSQL = "Select A.*,B.���� as ��Ժ����,C.���� as ��Ժ����" & _
        " From ������ҳ A,���ű� B,���ű� C" & _
        " Where A.��Ժ����ID=B.ID And A.��Ժ����ID=C.ID" & _
        " And A.����ID=[1] And A.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)

    txtInfo(txt���ʽ).Text = Nvl(rsTmp!ҽ�Ƹ��ʽ)
    '���۲�����סԺ��
    If Nvl(rsTmp!��������, 0) <> 0 Then
        lblInfo(txt��������).Visible = False
        txtInfo(txt��������).Visible = False
    End If
    chkInfo(chk����Ժ).Value = Nvl(rsTmp!����Ժ, 0)
    txtInfo(txt������).Text = Nvl(rsTmp!������)
    
    txtInfo(txt����).Text = Nvl(rsTmp!����)
    txtInfo(txt����).Text = Nvl(rsTmp!����)
    '�������
    txtInfo(txt���).Text = IIf(rsTmp!��� & "" = "0", "", rsTmp!��� & "")
    txtInfo(txt����).Text = IIf(rsTmp!���� & "" = "0", "", rsTmp!���� & "")
    txtInfo(txtְҵ).Text = Nvl(rsTmp!ְҵ)
    txtInfo(txt����).Text = Nvl(rsTmp!����״��)
    txtInfo(txt��ͥ��ַ).Text = Nvl(rsTmp!��ͥ��ַ)
    txtInfo(txt��ͥ�绰).Text = Nvl(rsTmp!��ͥ�绰)
    txtInfo(txt��ͥ�ʱ�).Text = Nvl(rsTmp!��ͥ��ַ�ʱ�)
    txtInfo(txt���ڵ�ַ).Text = Nvl(rsTmp!���ڵ�ַ)
    txtInfo(txt�����ʱ�).Text = Nvl(rsTmp!���ڵ�ַ�ʱ�)
    
    txtInfo(txt������λ).Text = Nvl(rsTmp!��λ��ַ)
    txtInfo(txt��λ�绰).Text = Nvl(rsTmp!��λ�绰)
    txtInfo(txt��λ�ʱ�).Text = Nvl(rsTmp!��λ�ʱ�)
    txtInfo(txt��ϵ������).Text = Nvl(rsTmp!��ϵ������)
    txtInfo(txt��ϵ�˹�ϵ).Text = Nvl(rsTmp!��ϵ�˹�ϵ)
    txtInfo(txt��ϵ�˵绰).Text = Nvl(rsTmp!��ϵ�˵绰)
    txtInfo(txt��ϵ�˵�ַ).Text = Nvl(rsTmp!��ϵ�˵�ַ)
    If Not IsNull(rsTmp!����) Then
        txtInfo(txt����).Text = Nvl(rsTmp!����)
    End If

    txtInfo(txt��Ժ;��).Text = Nvl(rsTmp!��Ժ��ʽ)
    txtInfo(txt��Ժʱ��).Text = Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm")
    txtInfo(txt��Ժ����).Text = rsTmp!��Ժ����
    
    txtInfo(txt��Ժʱ��).Text = Format(Nvl(rsTmp!��Ժ����), "yyyy-MM-dd HH:mm")
    txtInfo(txt��Ժ����).Text = rsTmp!��Ժ����
    If Not IsNull(rsTmp!��Ժ����) Then
        txtInfo(txtסԺ����).Text = DateDiff("d", rsTmp!��Ժ����, rsTmp!��Ժ����)
    Else
        txtInfo(txtסԺ����).Text = DateDiff("d", rsTmp!��Ժ����, zlDatabase.Currentdate)
    End If
    If Val(txtInfo(txtסԺ����).Text) = 0 Then txtInfo(txtסԺ����).Text = "1"
    
     txtInfo(txt��Ժ���).Text = Nvl(rsTmp!��Ժ����)
    chkInfo(chk�Ƿ�ȷ��).Value = Nvl(rsTmp!�Ƿ�ȷ��, 0)
    If chkInfo(chk�Ƿ�ȷ��).Value = 1 Then
        txtInfo(txtȷ������).Text = Format(Nvl(rsTmp!ȷ������), "yyyy-MM-dd HH:mm")
    End If
    chkInfo(chk��������ʬ��).Value = Nvl(rsTmp!ʬ���־, 0)
    chkInfo(chk�·�����).Value = Nvl(rsTmp!�·�����, 0)
    txtInfo(txt���ȴ���).Text = Nvl(rsTmp!���ȴ���)
    If Val(txtInfo(txt���ȴ���).Text) <> 0 Then
        txtInfo(txt�ɹ�����).Text = Nvl(rsTmp!�ɹ�����)
    End If
    
    txtInfo(txt�������).Text = Nvl(rsTmp!��ҽ�������)
    
    txtInfo(txtѪ��).Text = Nvl(rsTmp!Ѫ��)
    chkInfo(chk����).Value = IIf(Nvl(rsTmp!�����־, 0) = 0, 0, 1)
    If chkInfo(chk����).Value = 1 Then
        txtInfo(txt��������).Text = IIf(Nvl(rsTmp!�����־, 0) = 9, "", Nvl(rsTmp!��������, 0)) & _
            Decode(Nvl(rsTmp!�����־, 0), 1, "��", 2, "��", 3, "��", 4, "��", 9, "����")
    End If
    txtInfo(txt����ҽʦ).Text = Nvl(rsTmp!����ҽʦ)
    txtInfo(txtסԺҽʦ).Text = Nvl(rsTmp!סԺҽʦ)
    txtInfo(txt���λ�ʿ).Text = Nvl(rsTmp!���λ�ʿ)
    '���ʱ��
    If Nvl(rsTmp!״̬, 0) = 1 Then
        txtInfo(txt���ʱ��).Text = "��δ���"
    Else
        StrSQL = "Select ��ʼʱ�� From ���˱䶯��¼" & _
            " Where ����ID=[1] And ��ҳID=[2] And ��ʼԭ�� IN(2,1) And ��ʼʱ�� is Not Null Order by ��ʼԭ�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If Not rsTmp.EOF Then
            txtInfo(txt���ʱ��).Text = Format(rsTmp!��ʼʱ��, "yyyy-MM-dd HH:mm")
        End If
    End If
    
    '�����ӱ���
    '---------------------------------------------------------------
    StrSQL = "Select a.����ID,a.��ҳID,a.��Ϣ��,a.��Ϣֵ,b.���� From ������ҳ�ӱ� a " & _
            ",������Ŀ b" & " where a.��Ϣ��=b.����(+) And a.����ID=[1] And a.��ҳID=[2] Order by a.��Ϣ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    For i = 1 To rsTmp.RecordCount
        Select Case UCase(Nvl(rsTmp!��Ϣ��))
            Case "������������"
                txtInfo(txt������������).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��������������"
                txtInfo(txt����������).Text = Nvl(rsTmp!��Ϣֵ) & IIf(Nvl(rsTmp!��Ϣֵ) = "", "", " ��")
            Case "��������Ժ����"
                txtInfo(txt��������Ժ����).Text = Nvl(rsTmp!��Ϣֵ) & IIf(Nvl(rsTmp!��Ϣֵ) = "", "", " ��")
            Case "����"
                txtInfo(txt����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Ժ����"
                txtInfo(txt��Ժ����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Ժǰ����Ժ����"
                chkInfo(chk��Ժǰ����Ժ����).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "ת�Ƽ�¼"
                varTmp = Split(Nvl(rsTmp!��Ϣֵ), ",")
                If UBound(varTmp) >= 0 Then txtInfo(txtת��1).Text = varTmp(0)
                If UBound(varTmp) >= 1 Then txtInfo(txtת��2).Text = varTmp(1)
                If UBound(varTmp) >= 2 Then txtInfo(txtת��3).Text = varTmp(2)
            Case "��Ժ����"
                txtInfo(txt��Ժ����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "�����"
                txtInfo(txt�����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "�ֻ��̶�"
                If Nvl(rsTmp!��Ϣֵ) <> "" Then
                    txtInfo(txt�ֻ��̶�).Text = Nvl(rsTmp!��Ϣֵ)
                End If
            Case "����������"
                If Nvl(rsTmp!��Ϣֵ) <> "" Then
                    txtInfo(txt����������).Text = Nvl(rsTmp!��Ϣֵ)
                End If
            Case "��ԭѧ���"
                chkInfo(chkҽԺ��Ⱦ����ԭѧ���).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "����ʱ��"
                If Not (IsNull(rsTmp!��Ϣֵ) Or Not IsDate(rsTmp!��Ϣֵ)) Then
                    txtInfo(txt����ʱ��).Text = rsTmp!��Ϣֵ
                End If
            Case "��������ԭ��"
                txtInfo(txt����ԭ��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "���Ȳ���"
                txtInfo(txt����ԭ��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��ҽΣ��"
                chkInfo(chkΣ��).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "��ҽ��֢"
                chkInfo(chk��֢).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "��ҽ����"
                chkInfo(chk����).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "��ҽ���ȷ���"
                txtInfo(txt���ȷ���).Text = Nvl(rsTmp!��Ϣֵ)
            Case "������ҩ�Ƽ�"
                txtInfo(txt������ҩ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��ҽ�豸"
                txtInfo(txt��ҽ�豸).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��ҽ����"
                txtInfo(txt��ҽ����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��֤ʩ��"
                txtInfo(txt��֤ʩ��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��������"
                txtInfo(txt��������).Text = GetNameByCode("��������", Nvl(rsTmp!��Ϣֵ))
            Case UCase("HBsAg")
                txtInfo(txtHBsAg).Text = Nvl(rsTmp!��Ϣֵ)
            Case UCase("HCV-Ab")
                txtInfo(txtHCVAb).Text = Nvl(rsTmp!��Ϣֵ)
            Case UCase("HIV-Ab")
                txtInfo(txtHIVAb).Text = Nvl(rsTmp!��Ϣֵ)
            Case UCase("Rh")
            Case UCase("Rh")
                txtInfo(txtRh).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Ѫ���"
                txtInfo(txt��Ѫǰ9����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����ʱ��"
                If Nvl(rsTmp!��Ϣֵ) <> "" Then
                    If Format(rsTmp!��Ϣֵ, "HH:mm") <> "00:00" Then
                        txtInfo(txt����ʱ��).Text = Format(rsTmp!��Ϣֵ, "yyyy-MM-dd HH:mm")
                    Else
                        txtInfo(txt����ʱ��).Text = Format(rsTmp!��Ϣֵ, "yyyy-MM-dd")
                    End If
                End If
            Case "����״��"
                txtInfo(txt����״��).Text = Decode(Val(Nvl(rsTmp!��Ϣֵ, 0)), 0, "δ����", 1, "����1̥", 2, "����2̥������", 4, "4-����")
            Case "��Һ��Ӧ"
                txtInfo(txt��Һ��Ӧ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Ѫ��Ӧ"
                txtInfo(txt��Ѫ��Ӧ).Text = Decode(Val(Nvl(rsTmp!��Ϣֵ, 0)), 0, "��", 1, "��", 2, "δ��")
            Case "ʾ�̲���"
                chkInfo(chkʾ�̲���).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "���в���"
                chkInfo(chk���в���).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "���Ѳ���"
                chkInfo(chk���Ѳ���).Value = Val(Nvl(rsTmp!��Ϣֵ))
            Case "���ϸ��"
                txtInfo(txt���ϸ��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��ѪС��"
                txtInfo(txt��ѪС��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Ѫ��"
                txtInfo(txt��Ѫ��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��ȫѪ"
                txtInfo(txt��ȫѪ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "������"
                txtInfo(txt������).Text = Nvl(rsTmp!��Ϣֵ)
            Case "�������"
                txtInfo(txt�������).Text = Nvl(rsTmp!��Ϣֵ)
            Case "ҽѧ��ʾ"
                txtInfo(txtҽѧ��ʾ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����ҽѧ��ʾ"
                txtInfo(txt����ҽѧ��ʾ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Ժ��ʽ"
                txtInfo(txt��Ժ��ʽ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Ժת��"
                txtInfo(txtת�����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����ʱ��"
                '�����ʽ:��Ժǰ(�죬Сʱ,����)|��Ժ��(�죬Сʱ,����)
                txtInfo(txt��Ժǰ��).Text = Split(Split(Nvl(rsTmp!��Ϣֵ), "|")(0) & ",", ",")(0)
                txtInfo(txt��ԺǰСʱ).Text = Split(Split(Nvl(rsTmp!��Ϣֵ), "|")(0) & ",", ",")(1)
                txtInfo(txt��Ժǰ����).Text = Split(Split(Nvl(rsTmp!��Ϣֵ), "|")(0) & ",", ",")(2)
                txtInfo(txt��Ժ����).Text = Split(Split(Nvl(rsTmp!��Ϣֵ), "|")(1) & ",", ",")(0)
                txtInfo(txt��Ժ��Сʱ).Text = Split(Split(Nvl(rsTmp!��Ϣֵ) & "|", "|")(1) & ",", ",")(1)
                txtInfo(txt��Ժ�����).Text = Split(Split(Nvl(rsTmp!��Ϣֵ) & "|", "|")(1) & ",", ",")(2)
            Case "����Ժ�ƻ�����"
                lblInfo(txt����Ժ����).Caption = "��Ժ" & IIf(Nvl(rsTmp!��Ϣֵ, "0") = "0", "31", "7") & "��������Ժ�ƻ�"
            Case "31������סԺ"
                If Nvl(rsTmp!��Ϣֵ) <> "" Then
                    txtInfo(txt31��Ŀ��).Text = Nvl(rsTmp!��Ϣֵ)
                    txtInfo(txt����Ժ����).Text = "��"
                Else
                    txtInfo(txt����Ժ����).Text = "��"
                End If
            Case "������ʹ��ʱ��"
                txtInfo(txt������Сʱ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "������"
                txtInfo(txt������).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����ҽʦ"
                txtInfo(txt����ҽʦ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����ҽʦ"
                txtInfo(txt����ҽʦ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����ҽʦ"
                txtInfo(txt����ҽʦ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "�о���ʵϰҽʦ"
                txtInfo(txt�о���ҽʦ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "ʵϰҽʦ"
                txtInfo(txtʵϰҽʦ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "�ʿ�ҽʦ"
                txtInfo(txt�ʿ�ҽʦ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "�ʿػ�ʿ"
                txtInfo(txt�ʿػ�ʿ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��������"
                txtInfo(txt��������).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��ҳ��������"
                txtInfo(txt�ʿ�����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "CT"
                chkInfo(chkCT).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "MRI"
                chkInfo(chkMRI).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "��ɫ������"
                chkInfo(chk��ɫ������).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "������4"
                vsTSJC.TextMatrix(vsTSJC.FixedRows + 0, 1) = Nvl(rsTmp!��Ϣֵ)
                vsTSJC.Cell(flexcpData, vsTSJC.FixedRows + 0, 1) = vsTSJC.TextMatrix(vsTSJC.FixedRows + 0, 1)
            Case "������5"
                vsTSJC.TextMatrix(vsTSJC.FixedRows + 1, 1) = Nvl(rsTmp!��Ϣֵ)
                vsTSJC.Cell(flexcpData, vsTSJC.FixedRows + 1, 1) = vsTSJC.TextMatrix(vsTSJC.FixedRows + 1, 1)
            Case "������6"
                vsTSJC.TextMatrix(vsTSJC.FixedRows + 2, 1) = Nvl(rsTmp!��Ϣֵ)
                vsTSJC.Cell(flexcpData, vsTSJC.FixedRows + 2, 1) = vsTSJC.TextMatrix(vsTSJC.FixedRows + 2, 1)
            Case "ѹ�������ڼ�"
                txtInfo(txtѹ�������ڼ�).Text = Nvl(rsTmp!��Ϣֵ, " ")
            Case "ѹ������"
                txtInfo(txtѹ������).Text = Nvl(rsTmp!��Ϣֵ, " ")
            Case "������׹���˺�"
                txtInfo(txt������׹���˺�).Text = Nvl(rsTmp!��Ϣֵ, " ")
            Case "������׹��ԭ��"
                txtInfo(txt������׹��ԭ��).Text = Nvl(rsTmp!��Ϣֵ, " ")
            Case "��֢�໤����"
                txtInfo(txt��֢�໤����).Text = Nvl(rsTmp!��Ϣֵ, "")
            Case "��֢�໤Сʱ"
                txtInfo(txt��֢�໤Сʱ).Text = Nvl(rsTmp!��Ϣֵ, "")
            Case "������"
                chkInfo(chk������).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "�ٴ�·��"
                txtInfo(txt�ٴ�·������).Text = Decode(Val(rsTmp!��Ϣֵ & ""), 1, "δ����", 2, "�����˳�", 3, "���")
            Case "DRGS"
                txtInfo(txtDrGs����).Text = Decode(Val(rsTmp!��Ϣֵ & ""), 1, "��", 2, "������", 3, "������", 4, "���߶���")
            Case "������"
                chkInfo(chkʹ�ÿ�����).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "�걾�ͼ�"
                chkInfo(chkϸ�������걾�ͼ�).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "��Ⱦ��"
                txtInfo(txt��Ⱦ��).Text = Decode(Val(rsTmp!��Ϣֵ & ""), 1, "����", 2, "����", 3, "����")
            Case "��������"
                txtInfo(txt��������).Text = Decode(Val(rsTmp!��Ϣֵ & ""), 1, "0��", 2, "I��", 3, "����", 4, "����", 5, "����", 6, "����")
            Case "����T"
                txtInfo(txt��������T).Text = Nvl(rsTmp!��Ϣֵ, "")
            Case "����M"
                txtInfo(txt��������M).Text = Nvl(rsTmp!��Ϣֵ, "")
            Case "����N"
                txtInfo(txt��������N).Text = Nvl(rsTmp!��Ϣֵ, "")
            Case "APGAR"
                txtInfo(txtApgar).Text = Nvl(rsTmp!��Ϣֵ, "")
            Case Else
                If Not (Left(Nvl(rsTmp!��Ϣ��), 3) = "������" And Not IsNull(rsTmp!��Ϣֵ)) Then
                    '������Ŀ
                    If Not IsNull(rsTmp("����")) Then
                        With vsfMain
                            For j = 0 To vsfMain.Cols - 1 Step 3
                                lngRow = vsfMain.FindRow(rsTmp("��Ϣ��"), , j)
                                If lngRow >= 0 Then
                                    If vsfMain.TextMatrix(lngRow, j) = rsTmp("��Ϣ��") Then
                                        If vsfMain.TextMatrix(lngRow, j + 2) = "�Ƿ�" Then
                                            vsfMain.Cell(flexcpChecked, lngRow, j + 1) = IIf(rsTmp("��Ϣֵ") = 0, 2, 1)
                                            Exit For
                                        Else
                                            vsfMain.TextMatrix(lngRow, j + 1) = rsTmp("��Ϣֵ") & ""
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
    '��֢�໤����
    If Val(txtInfo(txt��֢�໤����).Text) <> 0 And Val(txtInfo(txt��֢�໤Сʱ).Text) <> 0 Then
        txtInfo(txt��֢�໤) = "��"
        If Val(txtInfo(txt��֢�໤����).Text) <> 0 Then
            txtInfo(txt��֢�໤����).Enabled = True
            lblInfo(txt��֢�໤����).Enabled = True
            lblInfo(txt��֢�໤���� + 100).Enabled = True
        End If
        If Val(txtInfo(txt��֢�໤Сʱ).Text) <> 0 Then
            txtInfo(txt��֢�໤Сʱ).Enabled = True
            lblInfo(txt��֢�໤Сʱ).Enabled = True
        End If
    End If
    
    '�Զ���ȡת�ƿ��Ҽ��������(�����)
    '---------------------------------------------------------------
    If txtInfo(txtת��1).Text = "" And txtInfo(txtת��2).Text = "" And txtInfo(txtת��3).Text = "" Then
        StrSQL = _
            " Select B.����" & _
            " From ���˱䶯��¼ A,���ű� B" & _
            " Where A.����ID=[1] And A.��ҳID=[2]" & _
            " And A.����ID=B.ID And A.��ʼԭ��=3 And A.��ʼʱ�� is Not NULL" & _
            " Order by A.��ʼʱ��"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
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
        StrSQL = "Select B.�����" & _
            " From ������ҳ A,��λ״����¼ B" & _
            " Where A.����ID=[1] And A.��ҳID=[2]" & _
            " And A.��Ժ����ID=B.����ID And A.��Ժ����=B.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If Not rsTmp.EOF Then txtInfo(txt��Ժ����).Text = Nvl(rsTmp!�����)
    End If

    If txtInfo(txt��Ժ����).Text = "" Then
        StrSQL = "Select B.�����" & _
            " From ������ҳ A,��λ״����¼ B" & _
            " Where A.����ID=[1] And A.��ҳID=[2]" & _
            " And A.��ǰ����ID=B.����ID And A.��Ժ����=B.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If Not rsTmp.EOF Then txtInfo(txt��Ժ����).Text = Nvl(rsTmp!�����)
    End If
    
    '��ҽ���
    '---------------------------------------------------------------
'    str���ƽ�� = Get���ƽ��
'    vsDiagXY.ColData(col��Ժ���) = str���ƽ��

    '�ж���ҳ�Ƿ�������
    StrSQL = "Select 1 From ������ϼ�¼ Where ����ID=[1] And ��ҳID=[2] And ��¼��Դ=3  And RowNum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    bln��ҳ��� = rsTmp.RecordCount > 0
    If bln��ҳ��� Then
        strTmp = " and a.��¼��Դ=3 "
    Else
        strTmp = " And a.��¼��Դ IN(1,2,3,4) "
    End If
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
    StrSQL = "Select a.��ע,a.ID,a.����ID,a.��ҳID,a.ҽ��ID,a.��¼��Դ,a.��ϴ���,a.�������,a.����ID,a.�������,a.����ID,a.��Ժ����," & _
        " a.���ID,a.֤��ID,a.�������,a.��Ժ���,a.�Ƿ�δ��,a.�Ƿ�����,a.��¼����,a.��¼��,a.ȡ��ʱ��,a.ȡ����,a.����ID, b.���� As ��������, c.���� As ��ϱ��� " & _
        " From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C" & _
        " Where a.����id = b.Id(+) And a.���id = c.Id(+)  And a.������� IN(1,2,3,5,6,7,10,21)" & _
        strTmp & _
        " And a.ȡ��ʱ�� is Null And a.����ID=[1] And a.��ҳID=[2]" & _
        " Order by a.�������,a.��¼��Դ Desc,a.��ϴ���,a.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then
        With vsDiagXY
            StrSQL = "1,2,3,5,6,7,10,21"
            For i = 0 To UBound(Split(StrSQL, ","))
                rsTmp.Filter = "��¼��Դ=3 And �������=" & Split(StrSQL, ",")(i)
                If Val(Split(StrSQL, ",")(i)) <> 21 Then
                    If rsTmp.EOF Then
                        rsTmp.Filter = "��¼��Դ=2 And �������=" & Split(StrSQL, ",")(i)
                    End If
                    If rsTmp.EOF Then
                        rsTmp.Filter = "��¼��Դ=1 And �������=" & Split(StrSQL, ",")(i)
                    End If
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=4 And �������=" & Split(StrSQL, ",")(i)
                End If

                If Val(Split(StrSQL, ",")(i)) = 21 Then
                    '21-��ԭѧ���
                    If Not rsTmp.EOF Then
                        txtInfo(txtҽԺ��Ⱦ��ԭѧ���).Text = Nvl(rsTmp!�������)
                    End If
                Else
                    Do While Not rsTmp.EOF
                        'ȷ����ǰ��ʾ��
                        lngRow = .FindRow(CStr(Split(StrSQL, ",")(i)), , col����)
                        For j = lngRow To .Rows - 1
                            If Val(.TextMatrix(j, col����)) = Val(Split(StrSQL, ",")(i)) Then
                                lngRow = j
                                If .TextMatrix(j, col�������) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, col�������) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, col����) = Split(StrSQL, ",")(i)
                        End If

                        '�ֻ��̶Ⱥ�����������
                        If Val("" & rsTmp!�������) = 3 And Val("" & rsTmp!��ϴ���) = 1 Then
                            If Trim(Nvl(rsTmp!��������)) = "" Then
                                bln�ֻ��̶� = False
                            Else
                                bln�ֻ��̶� = ((InStr("C", UCase(Left(Nvl(rsTmp!��������), 1)))) > 0) Or ((InStr("D0", UCase(Left(Nvl(rsTmp!��������), 2)))) > 0) Or ((InStr("D32.,D33.,", UCase(Left(Nvl(rsTmp!��������), 4)))) > 0)
                            End If
                        End If

                        txtInfo(txt�ֻ��̶�).Enabled = bln�ֻ��̶�
                        lblInfo(txt�ֻ��̶�).Enabled = bln�ֻ��̶�
                        lblInfo(txt����������).Enabled = bln�ֻ��̶�
                        txtInfo(txt����������).Enabled = bln�ֻ��̶�
                        .TextMatrix(lngRow, col�������) = Nvl(rsTmp!�������)
                        .TextMatrix(lngRow, col��ע) = Nvl(rsTmp!��ע)
                        .TextMatrix(lngRow, col��Ժ���) = Nvl(rsTmp!��Ժ���)
                        .TextMatrix(lngRow, col��Ժ����) = Nvl(rsTmp!��Ժ����)
                        .TextMatrix(lngRow, col�Ƿ�δ��) = IIf(Nvl(rsTmp!�Ƿ�δ��, 0) = 1, "��", "")
                        .TextMatrix(lngRow, col�Ƿ�����) = IIf(Nvl(rsTmp!�Ƿ�����, 0) = 1, "��", "")
                        rsTmp.MoveNext
                    Loop
                End If
            Next
        End With
    End If

    vsDiagXY.Cell(flexcpForeColor, 1, col�Ƿ�����, vsDiagXY.Rows - 1, col�Ƿ�����) = vbRed
    vsDiagXY.Cell(flexcpBackColor, GetRow(3), vsDiagXY.FixedRows, GetRow(3), vsDiagXY.Cols - 1) = &HC0FFC0
    vsDiagXY.Row = 1: vsDiagXY.Col = col�������
    If vsDiagXY.TextMatrix(GetRow(6), col�������) <> "" Then
        txtInfo(txt�����).Enabled = True
        txtInfo(txt�����).BackColor = vbWindowBackground
    End If

    '��Ϸ������
    '---------------------------------------------------------------
    StrSQL = "Select ��������,������� From ��Ϸ������ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    Do While Not rsTmp.EOF
        Select Case rsTmp!��������
        Case 1 '�������Ժ
            txtInfo(txt�������Ժ).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 2 '��Ժ���Ժ
            txtInfo(txt��Ժ���Ժ).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 3 '�����벡��
            txtInfo(txt�����벡��).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 4 '�ٴ��벡��
            txtInfo(txt�ٴ��벡��).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 5 '�ٴ���ʬ��
            txtInfo(txt�ٴ���ʬ��).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 6 '��ǰ������
            txtInfo(txt��ǰ������).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 7 '��������Ժ
             txtInfo(txt��������Ժ).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 11 '��ҽ�������Ժ
            txtInfo(txt��ҽ�������Ժ).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 12 '��ҽ��Ժ���Ժ
            txtInfo(txt��ҽ��Ժ���Ժ).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 13 '��ҽ��֤
            txtInfo(txt��֤).Text = Decode(Nvl(rsTmp!�������, 0), 1, "׼ȷ", 2, "����׼ȷ", 3, "�ش�ȱ��", 4, "����", "")
        Case 14 '��ҽ�η�
            txtInfo(txt�η�).Text = Decode(Nvl(rsTmp!�������, 0), 1, "׼ȷ", 2, "����׼ȷ", 3, "�ش�ȱ��", 4, "����", "")
        Case 15 '��ҽ��ҩ
            txtInfo(txt��ҩ).Text = Decode(Nvl(rsTmp!�������, 0), 1, "׼ȷ", 2, "����׼ȷ", 3, "�ش�ȱ��", 4, "����", "")
        End Select
        rsTmp.MoveNext
    Loop

    '��ҽ���
    '---------------------------------------------------------------
    'ȱʡ����ʼ��
    With vsDiagZY
        '11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���(��Ҫ��ϡ��������)
        .TextMatrix(1, colzy����) = 11
        .TextMatrix(2, colzy����) = 12
        .TextMatrix(3, colzy����) = 13
        .TextMatrix(4, colzy����) = 13
    End With

    If bln��ҳ��� Then
        strTmp = " and a.��¼��Դ=3 "
    Else
        strTmp = " And a.��¼��Դ IN(1,2,3,4) "
    End If

    '��ȡ������Դ�����
    StrSQL = "Select a.��ע, a.Id, a.����id, a.��ҳid, a.ҽ��id, a.��¼��Դ, a.��ϴ���, a.�������, a.����id, a.�������,a.��Ժ����," & _
        " a.����id, a.���id, a.֤��id, a.�������,a.��Ժ���, a.�Ƿ�δ��, a.�Ƿ�����, a.��¼����, a.��¼��, a.ȡ��ʱ��," & _
        " a.ȡ����, a.����id, b.���� As ��������, c.���� As ��ϱ���,d.���� as ֤����� From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C,��������Ŀ¼ D" & _
        " Where a.����id = b.Id(+) And a.���id = c.Id(+) And a.֤��ID=d.ID(+) And a.������� IN(11,12,13)" & _
        strTmp & _
        " And ȡ��ʱ�� Is Null And ����ID=[1] And ��ҳID=[2]" & _
        " Order by a.�������,a.��¼��Դ Desc,a.��ϴ���,a.�������,a.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    
    If Not rsTmp.EOF Then
        With vsDiagZY
            StrSQL = "11,12,13"
            For i = 0 To UBound(Split(StrSQL, ","))
                rsTmp.Filter = "��¼��Դ=3 And �������=" & Split(StrSQL, ",")(i)
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=2 And �������=" & Split(StrSQL, ",")(i)
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=1 And �������=" & Split(StrSQL, ",")(i)
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=4 And �������=" & Split(StrSQL, ",")(i)
                End If

                Do While Not rsTmp.EOF
                    'ȷ����ǰ��ʾ��
                    lngRow = .FindRow(CStr(Split(StrSQL, ",")(i)), , colzy����)
                    For j = lngRow To .Rows - 1
                        If Val(.TextMatrix(j, colzy����)) = Val(Split(StrSQL, ",")(i)) Then
                            lngRow = j
                            If .TextMatrix(j, col�������) = "" Then Exit For
                        Else
                            Exit For
                        End If
                    Next
                    If .TextMatrix(lngRow, col�������) <> "" Then
                        lngRow = lngRow + 1: .AddItem "", lngRow
                        .TextMatrix(lngRow, colzy����) = Split(StrSQL, ",")(i)
                    End If
                    .TextMatrix(lngRow, col��ע) = Nvl(rsTmp!��ע)
                    .TextMatrix(lngRow, col�������) = Nvl(rsTmp!�������)
                    .TextMatrix(lngRow, col��Ժ���) = Nvl(rsTmp!��Ժ���)
                    .TextMatrix(lngRow, col��Ժ����) = Nvl(rsTmp!��Ժ����)
                    'ȡ֤������
                    If InStr(.TextMatrix(lngRow, col�������), "(") > 0 And InStr(.TextMatrix(lngRow, col�������), ")") > 0 Then
                        strTmp = Mid(.TextMatrix(lngRow, col�������), InStrRev(.TextMatrix(lngRow, col�������), "(") + 1)
                        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                        '��ȡ֤��
                        .TextMatrix(lngRow, col��ҽ֤��) = strTmp
                        'ȥ�����������֤��
                        .TextMatrix(lngRow, col�������) = Mid(.TextMatrix(lngRow, col�������), 1, InStrRev(.TextMatrix(lngRow, col�������), "(") - 1)
                    Else
                       .TextMatrix(lngRow, col��ҽ֤��) = ""
                    End If
                    
                    rsTmp.MoveNext
                Loop
            Next
        End With
    End If
    vsDiagZY.Cell(flexcpBackColor, GetRow(13), vsDiagZY.FixedRows, GetRow(13), vsDiagZY.Cols - 1) = &HC0FFC0
    vsDiagZY.Row = 1: vsDiagZY.Col = col�������

    '������Ϣ:����סԺ��,������
    '---------------------------------------------------------------
    StrSQL = "Select ��¼��Դ,NVL(����ʱ��,��¼ʱ��) as ����ʱ��,ҩ��ID,ҩ����,������Ӧ From ���˹�����¼ A" & _
        " Where ���=1 And ����ID=[1] And ��ҳID=[2]" & _
        " And Not Exists(Select ҩ��ID From ���˹�����¼" & _
            " Where (Nvl(ҩ��ID,0)=Nvl(A.ҩ��ID,0) Or Nvl(ҩ����,'Null')=Nvl(A.ҩ����,'Null'))" & _
            " And Nvl(���,0)=0 And ��¼ʱ��>A.��¼ʱ�� And ����ID=[1] And ��ҳID=[2])" & _
        " Order by NVL(����ʱ��,��¼ʱ��),ҩ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then
        rsTmp.Filter = "��¼��Դ=3" '��ҳ������д��
        If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ<>3" '������Դ����Ϊȱʡ��ʾ
        With vsAller
            .Rows = rsTmp.RecordCount + 1 '�̶���+����
            For i = 1 To rsTmp.RecordCount
                '������Դ�Ŀ������ظ�
                lngRow = -1
                If Not IsNull(rsTmp!ҩ��ID) Then
                    lngRow = .FindRow(CLng(rsTmp!ҩ��ID))
                ElseIf Not IsNull(rsTmp!ҩ����) Then
                    lngRow = .FindRow(CStr(rsTmp!ҩ����), , col����ҩ��)
                End If
                If lngRow = -1 Then
                    .RowData(i) = CLng(Nvl(rsTmp!ҩ��ID, 0))
                    .TextMatrix(i, col����ʱ��) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                    .TextMatrix(i, col����ҩ��) = Nvl(rsTmp!ҩ����)
                    .TextMatrix(i, col������Ӧ) = Nvl(rsTmp!������Ӧ)
                End If
                rsTmp.MoveNext
            Next
            If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        End With
    End If
    vsAller.Row = 1: vsAller.Col = col����ҩ��

    '�������
    '---------------------------------------------------------------
    '�׶�ȡ��ҳ�����������
    StrSQL = "Select a.������ʿ,a.�������,a.�п�,a.����,a.��������,a.��������,a.����ҽʦ,a.��һ����,a.�ڶ�����,a.��������,a.����ҽʦ,a.ASA�ּ�,a.�ٴ�����,a.NNIS�ּ�,decode(a.��������,1,'һ������',2,'��������',3,'��������',4,'�ļ�����',' ') as ��������" & _
            " From ���������¼  A,��������Ŀ¼ B,������ĿĿ¼ C Where c.ID(+)=a.������ĿID And A.��������ID=B.ID(+) and ����ID=[1] And ��ҳID=[2] And ��¼��Դ=3 Order by A.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTmp.EOF Then 'û��ʱ��ȡ������Դ�����
        '��������������ʱ��дȡ��
        StrSQL = "Select Max(��¼����) From ���������¼" & _
            " Where ����ID=" & mlng����ID & " And ��ҳID=" & mlng��ҳID & _
            " And ��¼��Դ=1 And ȡ��ʱ�� is NULL"
         StrSQL = "Select a.������ʿ,a.�������,a.�п�,a.����,a.��������,a.��������,a.����ҽʦ,a.��һ����,a.�ڶ�����,a.��������,a.����ҽʦ,a.ASA�ּ�,a.�ٴ�����,a.NNIS�ּ�,decode(a.��������,1,'һ������',2,'��������',3,'��������',4,'�ļ�����',' ') as ��������" & _
            " From ���������¼  A,��������Ŀ¼ B,������ĿĿ¼ C Where c.ID(+)=a.������ĿID And " & _
            " A.��������ID=B.ID(+) and ����ID=[1] And ��ҳID=[2]" & _
            " And ��¼��Դ=1 And ȡ��ʱ�� is NULL And ��¼����=(" & StrSQL & ")" & _
            " Order by A.ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If rsTmp.EOF Then '����
            StrSQL = "Select a.������ʿ,a.�������,a.�п�,a.����,a.��������,a.��������,a.����ҽʦ,a.��һ����,a.�ڶ�����,a.��������,a.����ҽʦ,a.ASA�ּ�,a.�ٴ�����,a.NNIS�ּ�,decode(a.��������,1,'һ������',2,'��������',3,'��������',4,'�ļ�����',' ') as ��������" & _
                " From ���������¼  A,��������Ŀ¼ B,������ĿĿ¼ C Where c.ID(+)=a.������ĿID And  A.��������ID=B.ID(+) and ����ID=[1] And ��ҳID=[2] And ��¼��Դ=4 Order by A.ID"
            Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        End If
    End If
    If Not rsTmp.EOF Then
        With vsOPS
            .Rows = .FixedRows + rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col��������) = Format(Nvl(rsTmp!��������), "yyyy-MM-dd")
                .TextMatrix(i, col��������) = Nvl(rsTmp!��������)
                .TextMatrix(i, col����ҽʦ) = Nvl(rsTmp!����ҽʦ)
                .TextMatrix(i, col������ʿ) = Nvl(rsTmp!������ʿ)
                .TextMatrix(i, col����1) = Nvl(rsTmp!��һ����)
                .TextMatrix(i, col����2) = Nvl(rsTmp!�ڶ�����)
                .TextMatrix(i, col����ʽ) = Nvl(rsTmp!��������)
                .TextMatrix(i, col����ҽʦ) = Nvl(rsTmp!����ҽʦ)
                If Not IsNull(rsTmp!�п�) And Not IsNull(rsTmp!����) Then
                    .TextMatrix(i, col�п�����) = rsTmp!�п� & "/" & rsTmp!����
                End If
                .TextMatrix(i, COL�������) = Nvl(rsTmp!�������)
                .TextMatrix(i, colASA�ּ�) = Decode(Nvl(rsTmp!asa�ּ�), "I��", "P1", "II��", "P2", "III��", "P3", "IV��", "P4", "V��", "P5", Nvl(rsTmp!asa�ּ�))
                .TextMatrix(i, colNNIS�ּ�) = Nvl(rsTmp!NNIS�ּ�)
                .TextMatrix(i, col�����ּ�) = Nvl(rsTmp!��������)
                .TextMatrix(i, col�ٴ�����) = IIf(Val(rsTmp!�ٴ����� & "") = 1, -1, 0)
                rsTmp.MoveNext
            Next
        End With
    End If

    
    If mbln���� Then
        '���ƻ���
        Call Load���������(mlng����ID, mlng��ҳID)
    End If
    
    '������Ϣ
    '---------------------------------------------------------------
    '�����¼�
    lstAdvEvent.Clear
    
    
    blnѹ�� = False: bln����׹�� = False
    StrSQL = "Select ����, ����" & vbNewLine & _
            "From �����¼� A," & vbNewLine & _
            "     (Select Decode(��Ϣֵ, Null, Null, ',' || ��Ϣֵ || ',') ��Ϣֵ" & vbNewLine & _
            "       From ������ҳ�ӱ�" & vbNewLine & _
            "       Where ����id = [1] And ��ҳid = [2] And ��Ϣ�� = '�����¼�') B" & vbNewLine & _
            "Where Instr(b.��Ϣֵ , chr(44)|| a.���� ||chr(44) ) > 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    For i = 1 To rsTmp.RecordCount
        lstAdvEvent.AddItem Nvl(rsTmp!����)
        If Nvl(rsTmp!����) = "ѹ��" Then
            blnѹ�� = True
        ElseIf Nvl(rsTmp!����) = "ҽԺ�ڵ���/׹��" Then 'ѹ�� ����׹��
            bln����׹�� = True
        End If
        rsTmp.MoveNext
    Next

    txtInfo(txtѹ�������ڼ�).Enabled = blnѹ��
    txtInfo(txtѹ������).Enabled = blnѹ��
    lblInfo(txtѹ�������ڼ�).Enabled = blnѹ��
    lblInfo(txtѹ������).Enabled = blnѹ��

    txtInfo(txt������׹��ԭ��).Enabled = bln����׹��
    txtInfo(txt������׹���˺�).Enabled = bln����׹��
    lblInfo(txt������׹��ԭ��).Enabled = bln����׹��
    lblInfo(txt������׹���˺�).Enabled = bln����׹��
    
    '��Ⱦ����
    lstInfection.Clear
    StrSQL = "Select ����, ����" & vbNewLine & _
        "From ��Ⱦ���� A," & vbNewLine & _
        "     (Select Decode(��Ϣֵ, Null, Null, ',' || ��Ϣֵ || ',') ��Ϣֵ" & vbNewLine & _
        "       From ������ҳ�ӱ�" & vbNewLine & _
        "       Where ����id = [1] And ��ҳid = [2] And ��Ϣ�� = '��Ⱦ����') B" & vbNewLine & _
        "Where Instr(b.��Ϣֵ , chr(44)|| a.���� ||chr(44) ) > 0"

    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    For i = 1 To rsTmp.RecordCount
        lstInfection.AddItem Nvl(rsTmp!����)
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

Private Function Load���������(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ط����뻯����Ϣ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-21 15:55:27
    '����:13999
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim StrSQL As String
    
    Err = 0: On Error GoTo Errhand:
    StrSQL = " " & _
    "   Select A.����id, A.��ҳid, A.���, A.����id, A.��ʼ����, A.��������, A.�Ƴ���, A.����, A.���Ʒ���, A.����Ч��, " & _
    "          B.���� || '-' || B.���� As ������Ϣ " & _
    "   From �������Ƽ�¼ A, ��������Ŀ¼ B " & _
    "   Where A.����id = B.Id And a.����id=[1] And a.��ҳid=[2] " & _
    "   Order By ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng����ID, lng��ҳID)
    With vsChemotherapy
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Clear 1
        If rsTemp.RecordCount <> 0 Then .Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("��ѧ���Ʊ���")) = Nvl(rsTemp!������Ϣ)
            .TextMatrix(lngRow, .ColIndex("��ʼ����")) = Format(rsTemp!��ʼ����, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!��������, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("�Ƴ���")) = Format(Val(Nvl(rsTemp!�Ƴ���)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("����")) = Format(Val(Nvl(rsTemp!����)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("���Ʒ���")) = Trim(Nvl(rsTemp!���Ʒ���))
            .TextMatrix(lngRow, .ColIndex("����Ч��")) = Trim(Nvl(rsTemp!����Ч��))
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    StrSQL = " " & _
    "   Select A.����id, A.��ҳid, A.���, A.����id, A.��ʼ����, A.��������,A.��Ұ��λ, A.�������, A.�ۼ���, A.����Ч��, " & _
    "          B.���� || '-' || B.���� As ������Ϣ " & _
    "   From �������Ƽ�¼ A, ��������Ŀ¼ B " & _
    "   Where A.����id = B.Id And a.����id=[1] And a.��ҳid=[2] " & _
    "   Order By ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng����ID, lng��ҳID)
    With vsRadiotherapy
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Clear 1
        If rsTemp.RecordCount <> 0 Then .Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("�������Ʊ���")) = Nvl(rsTemp!������Ϣ)
            .TextMatrix(lngRow, .ColIndex("��ʼ����")) = Format(rsTemp!��ʼ����, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!��������, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("�������")) = Format(Val(Nvl(rsTemp!�������)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("�ۼ���")) = Format(Val(Nvl(rsTemp!�ۼ���)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("��Ұ��λ")) = Trim(Nvl(rsTemp!��Ұ��λ))
            .TextMatrix(lngRow, .ColIndex("����Ч��")) = Trim(Nvl(rsTemp!����Ч��))
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Load��������� = True
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
    StrSQL = "select ����,���� from ������Ŀ order by ����"
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
            .TextMatrix(0, lngRow) = "��Ŀ"
            .TextMatrix(0, lngRow + 1) = "����"
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
                .TextMatrix(lngRow, lngCol + 0) = rsTemp!����
                .TextMatrix(lngRow, lngCol + 2) = rsTemp!���� & ""
                If InStr(rsTemp!����, "�Ƿ�") > 0 Then
                    vsfMain.TextMatrix(lngRow, lngCol + 1) = "��"
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

