VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArchiveInMedRec_YN 
   BorderStyle     =   0  'None
   ClientHeight    =   7695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   10275
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
      ScaleWidth      =   10125
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   120
      Width           =   10125
      Begin VB.VScrollBar vsc 
         Height          =   6975
         Left            =   9840
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   255
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
      Begin VB.HScrollBar hsc 
         Height          =   255
         Left            =   90
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   7200
         Visible         =   0   'False
         Width           =   9645
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
               Picture         =   "frmArchiveInMedRec_YN.frx":0000
               Key             =   "-"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmArchiveInMedRec_YN.frx":04EA
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
         Width           =   9645
         Begin VB.Frame fraMain 
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
               Index           =   172
               Left            =   1425
               Locked          =   -1  'True
               TabIndex        =   355
               Top             =   4320
               Width           =   2550
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   166
               Left            =   2970
               Locked          =   -1  'True
               TabIndex        =   352
               Top             =   3585
               Width           =   975
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   165
               Left            =   6990
               Locked          =   -1  'True
               TabIndex        =   350
               Top             =   4320
               Width           =   1740
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   164
               Left            =   7560
               Locked          =   -1  'True
               TabIndex        =   348
               Top             =   735
               Width           =   1095
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   122
               Left            =   4815
               Locked          =   -1  'True
               TabIndex        =   297
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
               TabIndex        =   232
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
               TabIndex        =   231
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
               TabIndex        =   228
               Top             =   3945
               Width           =   3075
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "再入院"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   0
               Left            =   5640
               TabIndex        =   226
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
               TabIndex        =   192
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
               Top             =   5025
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
               Top             =   5025
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
               Top             =   5025
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
               Top             =   5385
               Width           =   1575
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
               Top             =   5385
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
               Top             =   5385
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
               Top             =   5385
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
               Width           =   555
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
               Left            =   6240
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
               Left            =   6960
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   1785
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
               Left            =   4800
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
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   3585
               Width           =   615
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
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "监护人身份证号"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   200
               Left            =   150
               TabIndex        =   356
               Top             =   4320
               Width           =   1260
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   37
               X1              =   1425
               X2              =   3960
               Y1              =   4500
               Y2              =   4500
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   166
               X1              =   2880
               X2              =   3960
               Y1              =   3765
               Y2              =   3765
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   165
               X1              =   6990
               X2              =   8760
               Y1              =   4500
               Y2              =   4500
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "转入"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   165
               Left            =   6600
               TabIndex        =   351
               Top             =   4320
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   164
               X1              =   7560
               X2              =   8760
               Y1              =   915
               Y2              =   915
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院号"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   164
               Left            =   7020
               TabIndex        =   349
               Top             =   720
               Width           =   540
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
               TabIndex        =   298
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
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "身高      cm"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   10
               Left            =   2265
               TabIndex        =   230
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
               TabIndex        =   229
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
               TabIndex        =   227
               Top             =   3945
               Width           =   360
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
               Top             =   5385
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
               Top             =   5025
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
               Top             =   5025
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
               Top             =   5025
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
               Top             =   5385
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
               Top             =   5385
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
               Left            =   1800
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
               Caption         =   "出生地"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   14
               Left            =   510
               TabIndex        =   100
               Top             =   1785
               Width           =   540
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
               Left            =   6600
               TabIndex        =   98
               Top             =   1770
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
               Left            =   4080
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
               Left            =   5880
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
               X1              =   6240
               X2              =   6990
               Y1              =   880
               Y2              =   880
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   8
               X1              =   6960
               X2              =   8160
               Y1              =   1965
               Y2              =   1965
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
               X1              =   4800
               X2              =   6390
               Y1              =   4485
               Y2              =   4485
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
               Y1              =   4125
               Y2              =   4125
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
               X2              =   1680
               Y1              =   3765
               Y2              =   3765
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   29
               X1              =   2190
               X2              =   2880
               Y1              =   3765
               Y2              =   3765
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
               Y1              =   5565
               Y2              =   5565
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
               Y1              =   5565
               Y2              =   5565
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
               Y1              =   5565
               Y2              =   5565
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   44
               X1              =   7200
               X2              =   8880
               Y1              =   5565
               Y2              =   5565
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   38
               X1              =   1080
               X2              =   2700
               Y1              =   5205
               Y2              =   5205
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   39
               X1              =   3480
               X2              =   5160
               Y1              =   5205
               Y2              =   5205
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   40
               X1              =   6120
               X2              =   8880
               Y1              =   5205
               Y2              =   5205
            End
         End
         Begin VB.Frame fraMain 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   西医诊断 "
            ForeColor       =   &H00FF0000&
            Height          =   5775
            Index           =   1
            Left            =   120
            TabIndex        =   160
            Tag             =   "5775"
            Top             =   240
            Width           =   9495
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   168
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   353
               Top             =   4627
               Width           =   555
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   145
               Left            =   975
               Locked          =   -1  'True
               TabIndex        =   346
               Top             =   5400
               Width           =   3870
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   146
               Left            =   6375
               Locked          =   -1  'True
               TabIndex        =   344
               Top             =   5400
               Width           =   870
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   56
               Left            =   7155
               Locked          =   -1  'True
               TabIndex        =   245
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
               TabIndex        =   244
               Top             =   3120
               Width           =   1515
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   59
               Left            =   5400
               Locked          =   -1  'True
               TabIndex        =   242
               Top             =   4627
               Width           =   3570
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   57
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   240
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
               TabIndex        =   238
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
               TabIndex        =   236
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
               TabIndex        =   234
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
               TabIndex        =   225
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
               TabIndex        =   221
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
               TabIndex        =   193
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
               TabIndex        =   190
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
               Left            =   2520
               TabIndex        =   189
               Top             =   4620
               Width           =   1020
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
               FormatString    =   $"frmArchiveInMedRec_YN.frx":09D4
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
               Index           =   168
               X1              =   1320
               X2              =   2160
               Y1              =   4805
               Y2              =   4805
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "死亡患者尸检"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   168
               Left            =   240
               TabIndex        =   354
               Top             =   4627
               Width           =   1080
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   145
               X1              =   960
               X2              =   4920
               Y1              =   5580
               Y2              =   5580
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "感染部位"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   145
               Left            =   240
               TabIndex        =   347
               Top             =   5400
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   146
               X1              =   6360
               X2              =   7245
               Y1              =   5580
               Y2              =   5580
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "感染与死亡关系"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   146
               Left            =   4980
               TabIndex        =   345
               Top             =   5400
               Width           =   1260
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
               TabIndex        =   246
               Top             =   3876
               Width           =   900
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   59
               X1              =   5400
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
               Left            =   3675
               TabIndex        =   243
               Top             =   4620
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
               TabIndex        =   241
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
               TabIndex        =   239
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
               TabIndex        =   237
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
               TabIndex        =   235
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
               TabIndex        =   233
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
               TabIndex        =   222
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
               TabIndex        =   194
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
               TabIndex        =   191
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
         Begin VB.Frame fraMain 
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
               TabIndex        =   343
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
               TabIndex        =   295
               Top             =   2467
               Width           =   2280
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
               TabIndex        =   293
               Top             =   2467
               Width           =   1920
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   117
               Left            =   4200
               Locked          =   -1  'True
               TabIndex        =   271
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
               TabIndex        =   269
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
               TabIndex        =   266
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
               TabIndex        =   265
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
               TabIndex        =   263
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
               TabIndex        =   261
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
               TabIndex        =   259
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
               TabIndex        =   257
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
               TabIndex        =   255
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
               TabIndex        =   253
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
               TabIndex        =   251
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
               TabIndex        =   248
               Tag             =   "无"
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
               TabIndex        =   247
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
               TabIndex        =   212
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
               TabIndex        =   209
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
               TabIndex        =   207
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
               TabIndex        =   205
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
               TabIndex        =   202
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
               TabIndex        =   200
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
               TabIndex        =   195
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
               Index           =   10
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
               Index           =   12
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
               Index           =   11
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
               TabIndex        =   198
               Top             =   3120
               Width           =   435
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   92
               X1              =   4170
               X2              =   6480
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
               TabIndex        =   296
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
               TabIndex        =   294
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
               TabIndex        =   272
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
               TabIndex        =   270
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
               TabIndex        =   268
               Top             =   2820
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "离院方式"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   93
               Left            =   180
               TabIndex        =   267
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
               TabIndex        =   264
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
               TabIndex        =   262
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
               TabIndex        =   260
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
               TabIndex        =   258
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
               TabIndex        =   256
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
               TabIndex        =   254
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
               TabIndex        =   252
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
               TabIndex        =   250
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
               TabIndex        =   249
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
               TabIndex        =   214
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
               TabIndex        =   213
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
               TabIndex        =   211
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
               TabIndex        =   210
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
               TabIndex        =   208
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
               TabIndex        =   206
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
               TabIndex        =   204
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
               TabIndex        =   203
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
               TabIndex        =   201
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
               TabIndex        =   197
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
               TabIndex        =   196
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
               TabIndex        =   199
               Top             =   3180
               Width           =   360
            End
         End
         Begin VB.Frame fraMain 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   放疗与化疗 "
            ForeColor       =   &H00FF0000&
            Height          =   5010
            Index           =   5
            Left            =   120
            TabIndex        =   223
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
               TabIndex        =   224
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VSFlex8Ctl.VSFlexGrid vsChemoth 
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
               FormatString    =   $"frmArchiveInMedRec_YN.frx":0B25
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
            Begin VSFlex8Ctl.VSFlexGrid vsRadioth 
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
               FormatString    =   $"frmArchiveInMedRec_YN.frx":0C3B
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
            Begin VB.Label lblVsTitle 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "放疗记录信息"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   274
               Top             =   2400
               Width           =   1080
            End
            Begin VB.Label lblVsTitle 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "化疗记录信息"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   273
               Top             =   240
               Width           =   1080
            End
         End
         Begin VB.Frame fraMain 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   附页1 "
            ForeColor       =   &H00FF0000&
            Height          =   5370
            Index           =   6
            Left            =   120
            TabIndex        =   275
            Tag             =   "5370"
            Top             =   120
            Width           =   9495
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
               TabIndex        =   342
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "住院期间出现危重"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   18
               Left            =   7560
               TabIndex        =   317
               Top             =   2640
               Width           =   1755
            End
            Begin VB.Frame fraPath 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "临床路径信息"
               ForeColor       =   &H80000008&
               Height          =   2595
               Left            =   240
               TabIndex        =   309
               Top             =   2520
               Width           =   2895
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "变异"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   21
                  Left            =   120
                  TabIndex        =   316
                  Top             =   1629
                  Width           =   675
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "完成路径"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   20
                  Left            =   120
                  TabIndex        =   315
                  Top             =   708
                  Width           =   1395
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "进入路径"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   19
                  Left            =   120
                  TabIndex        =   314
                  Top             =   240
                  Width           =   1395
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   141
                  Left            =   1260
                  Locked          =   -1  'True
                  TabIndex        =   311
                  Top             =   2100
                  Width           =   1395
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   140
                  Left            =   1320
                  Locked          =   -1  'True
                  TabIndex        =   310
                  Top             =   1176
                  Width           =   1335
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   141
                  X1              =   1260
                  X2              =   2640
                  Y1              =   2280
                  Y2              =   2280
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   140
                  X1              =   1320
                  X2              =   2640
                  Y1              =   1350
                  Y2              =   1350
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "退出原因"
                  Enabled         =   0   'False
                  Height          =   180
                  Index           =   140
                  Left            =   600
                  TabIndex        =   313
                  Top             =   1176
                  Width           =   720
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "变异原因"
                  Enabled         =   0   'False
                  Height          =   180
                  Index           =   141
                  Left            =   480
                  TabIndex        =   312
                  Top             =   2100
                  Width           =   720
               End
            End
            Begin VB.Frame fraAdvEvent 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "不良事件"
               ForeColor       =   &H80000008&
               Height          =   2595
               Left            =   3360
               TabIndex        =   299
               Top             =   2520
               Width           =   3855
               Begin VB.ListBox lstAdvEvent 
                  Height          =   960
                  ItemData        =   "frmArchiveInMedRec_YN.frx":0D61
                  Left            =   120
                  List            =   "frmArchiveInMedRec_YN.frx":0D63
                  TabIndex        =   304
                  Top             =   240
                  Width           =   3405
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
                  TabIndex        =   303
                  Top             =   1500
                  Width           =   975
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
                  TabIndex        =   302
                  Top             =   1500
                  Width           =   915
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
                  Index           =   121
                  Left            =   1440
                  Locked          =   -1  'True
                  TabIndex        =   300
                  Top             =   2220
                  Width           =   1995
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "跌倒或坠床伤害"
                  Height          =   180
                  Index           =   120
                  Left            =   120
                  TabIndex        =   308
                  Top             =   1860
                  Width           =   1260
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "跌倒或坠床原因"
                  Height          =   180
                  Index           =   121
                  Left            =   120
                  TabIndex        =   307
                  Top             =   2220
                  Width           =   1260
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "压疮发生期间"
                  Height          =   180
                  Index           =   118
                  Left            =   120
                  TabIndex        =   306
                  Top             =   1500
                  Width           =   1080
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "分期"
                  Height          =   180
                  Index           =   119
                  Left            =   2280
                  TabIndex        =   305
                  Top             =   1500
                  Width           =   360
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   118
                  X1              =   1200
                  X2              =   2260
                  Y1              =   1680
                  Y2              =   1680
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
                  Index           =   120
                  X1              =   1440
                  X2              =   3600
                  Y1              =   2040
                  Y2              =   2040
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   121
                  X1              =   1440
                  X2              =   3600
                  Y1              =   2400
                  Y2              =   2400
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid vsKSS 
               Height          =   1815
               Left            =   120
               TabIndex        =   276
               Top             =   600
               Width           =   9120
               _cx             =   16087
               _cy             =   3201
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
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmArchiveInMedRec_YN.frx":0D65
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
            Begin VB.Label lblVsTitle 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "抗菌药物使用情况（按DDD数降序排列）"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   2
               Left            =   120
               TabIndex        =   277
               Top             =   240
               Width           =   3150
            End
         End
         Begin VB.Frame fraMain 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   附页2 "
            ForeColor       =   &H00FF0000&
            Height          =   6810
            Index           =   7
            Left            =   120
            TabIndex        =   280
            Tag             =   "6810"
            Top             =   120
            Width           =   9495
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   153
               Left            =   7440
               Locked          =   -1  'True
               TabIndex        =   336
               Top             =   6360
               Width           =   975
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   152
               Left            =   7440
               Locked          =   -1  'True
               TabIndex        =   334
               Top             =   5563
               Width           =   975
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   151
               Left            =   7440
               Locked          =   -1  'True
               TabIndex        =   332
               Top             =   5166
               Width           =   975
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   150
               Left            =   7440
               Locked          =   -1  'True
               TabIndex        =   330
               Top             =   4769
               Width           =   975
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   149
               Left            =   7440
               Locked          =   -1  'True
               TabIndex        =   327
               Top             =   4365
               Width           =   975
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   148
               Left            =   2400
               Locked          =   -1  'True
               TabIndex        =   325
               Top             =   1533
               Width           =   2535
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "非预期的重返重症医学科"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   27
               Left            =   3360
               TabIndex        =   324
               Top             =   1132
               Width           =   2355
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "发生人工气道脱出"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   26
               Left            =   1200
               TabIndex        =   323
               Top             =   1132
               Width           =   1755
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "住院期间使用物理约束"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   23
               Left            =   5880
               TabIndex        =   322
               Top             =   3960
               Width           =   2235
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   147
               Left            =   2400
               Locked          =   -1  'True
               TabIndex        =   320
               Top             =   746
               Width           =   2415
            End
            Begin VB.Frame fraSplit 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00FFFFFF&
               Height          =   75
               Index           =   9
               Left            =   2400
               TabIndex        =   318
               Top             =   420
               Width           =   6855
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "3.彩色多普勒"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   15
               Left            =   2280
               TabIndex        =   292
               Top             =   5280
               Width           =   1395
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "2.MRI"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   14
               Left            =   1200
               TabIndex        =   291
               Top             =   5280
               Width           =   915
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "1.CT"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   13
               Left            =   120
               TabIndex        =   290
               Top             =   5280
               Width           =   915
            End
            Begin VB.Frame fraSplit 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00FFFFFF&
               Height          =   75
               Index           =   8
               Left            =   1200
               TabIndex        =   288
               Top             =   4980
               Width           =   4455
            End
            Begin VB.Frame fraSplit 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00FFFFFF&
               Height          =   75
               Index           =   7
               Left            =   1200
               TabIndex        =   286
               Top             =   1973
               Width           =   4455
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
                  ItemData        =   "frmArchiveInMedRec_YN.frx":0E4C
                  Left            =   120
                  List            =   "frmArchiveInMedRec_YN.frx":0E4E
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
               Index           =   7
               Left            =   240
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
               Rows            =   2
               Cols            =   9
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
               TabIndex        =   289
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
               FormatString    =   $"frmArchiveInMedRec_YN.frx":0E50
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
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "产科新生儿情况"
               Height          =   180
               Index           =   307
               Left            =   5880
               TabIndex        =   338
               Top             =   5955
               Width           =   1260
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "离院方式"
               Height          =   180
               Index           =   153
               Left            =   6720
               TabIndex        =   337
               Top             =   6360
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   153
               X1              =   7440
               X2              =   8500
               Y1              =   6535
               Y2              =   6535
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "约束原因"
               Height          =   180
               Index           =   152
               Left            =   6720
               TabIndex        =   335
               Top             =   5563
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   152
               X1              =   7440
               X2              =   8500
               Y1              =   5740
               Y2              =   5740
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "约束工具"
               Height          =   180
               Index           =   151
               Left            =   6720
               TabIndex        =   333
               Top             =   5166
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   151
               X1              =   7440
               X2              =   8500
               Y1              =   5340
               Y2              =   5340
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "约束方式"
               Height          =   180
               Index           =   150
               Left            =   6720
               TabIndex        =   331
               Top             =   4769
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   150
               X1              =   7440
               X2              =   8500
               Y1              =   4945
               Y2              =   4945
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "小时"
               Height          =   180
               Index           =   249
               Left            =   8520
               TabIndex        =   329
               Top             =   4365
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "约束总时间"
               Height          =   180
               Index           =   149
               Left            =   6540
               TabIndex        =   328
               Top             =   4365
               Width           =   900
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   149
               X1              =   7440
               X2              =   8500
               Y1              =   4540
               Y2              =   4540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "重返间隔时间"
               Height          =   180
               Index           =   148
               Left            =   1320
               TabIndex        =   326
               Top             =   1530
               Width           =   1080
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   148
               X1              =   2400
               X2              =   5040
               Y1              =   1710
               Y2              =   1710
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "重症监护室名称"
               Height          =   180
               Index           =   147
               Left            =   1140
               TabIndex        =   321
               Top             =   750
               Width           =   1260
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   147
               X1              =   2400
               X2              =   4920
               Y1              =   930
               Y2              =   930
            End
            Begin VB.Label lblVsTitle 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入住重症监护室（ICU）情况"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   6
               Left            =   120
               TabIndex        =   319
               Top             =   360
               Width           =   2250
            End
            Begin VB.Label lblVsTitle 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "特殊检查情况"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   5
               Left            =   120
               TabIndex        =   287
               Top             =   4920
               Width           =   1080
            End
            Begin VB.Label lblVsTitle 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病案附加项目"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   4
               Left            =   120
               TabIndex        =   282
               Top             =   1920
               Width           =   1080
            End
         End
         Begin VB.Frame fraMain 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   过敏与手术 "
            ForeColor       =   &H00FF0000&
            Height          =   3585
            Index           =   3
            Left            =   120
            TabIndex        =   145
            Tag             =   "3585"
            Top             =   120
            Width           =   9495
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "发生术后猝死"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   25
               Left            =   3840
               TabIndex        =   340
               Top             =   3248
               Width           =   1380
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "发生围术期死亡"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   24
               Left            =   2145
               TabIndex        =   339
               Top             =   3248
               Width           =   1620
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
               Cols            =   29
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmArchiveInMedRec_YN.frx":0EBE
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
               FormatString    =   $"frmArchiveInMedRec_YN.frx":131F
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
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "手术及操作相关情况："
               Height          =   180
               Index           =   300
               Left            =   165
               TabIndex        =   341
               Top             =   3255
               Width           =   1800
            End
         End
         Begin VB.Frame fraMain 
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
                  TabIndex        =   219
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
                  TabIndex        =   217
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
                  TabIndex        =   215
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
                  TabIndex        =   220
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
                  TabIndex        =   218
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
                  TabIndex        =   216
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
               FormatString    =   $"frmArchiveInMedRec_YN.frx":138C
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
      End
   End
End
Attribute VB_Name = "frmArchiveInMedRec_YN"
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

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= 0 And NewCol >= 0 Then Call vsfMain.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsKSS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= 0 And NewCol >= 0 Then Call vsKSS.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsOPS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= 0 And NewCol >= 0 Then Call vsOPS.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsRadioth_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= 0 And NewCol >= 0 Then Call vsRadioth.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsChemoth_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= 0 And NewCol >= 0 Then Call vsChemoth.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsTSJC_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= 0 And NewCol >= 0 Then Call vsTSJC.ShowCell(NewRow, NewCol)
End Sub



