VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArchiveInMedRec_SC 
   BorderStyle     =   0  'None
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8370
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8235
      Left            =   120
      ScaleHeight     =   8235
      ScaleWidth      =   10245
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   120
      Width           =   10245
      Begin VB.Frame fraVH 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9840
         TabIndex        =   61
         Top             =   7920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.VScrollBar vsc 
         Height          =   7800
         Left            =   9840
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.HScrollBar hsc 
         Height          =   255
         Left            =   90
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   7920
         Visible         =   0   'False
         Width           =   9675
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
               Picture         =   "frmArchiveInMedRec_SC.frx":0000
               Key             =   "-"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmArchiveInMedRec_SC.frx":04EA
               Key             =   "+"
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraBack 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7800
         Left            =   90
         TabIndex        =   62
         Top             =   0
         Width           =   9675
         Begin VB.Frame fraMain 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   基本信息 "
            ForeColor       =   &H00FF0000&
            Height          =   6195
            Index           =   0
            Left            =   120
            TabIndex        =   63
            Tag             =   "6195"
            Top             =   240
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
               TabIndex        =   352
               Top             =   4680
               Width           =   2430
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
               TabIndex        =   344
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
               TabIndex        =   342
               Top             =   4680
               Width           =   1860
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
               TabIndex        =   340
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   138
               Left            =   6960
               Locked          =   -1  'True
               TabIndex        =   205
               Top             =   4305
               Width           =   1860
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   137
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   203
               Top             =   4305
               Width           =   1860
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
               TabIndex        =   201
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
               TabIndex        =   162
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
               TabIndex        =   161
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
               TabIndex        =   158
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
               TabIndex        =   156
               Top             =   5385
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
               TabIndex        =   155
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
               Left            =   6780
               TabIndex        =   141
               Top             =   5018
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
               TabIndex        =   135
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
               TabIndex        =   134
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
               TabIndex        =   132
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
               TabIndex        =   129
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
               TabIndex        =   128
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
               TabIndex        =   126
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
               TabIndex        =   124
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
               Left            =   5280
               Locked          =   -1  'True
               TabIndex        =   29
               Top             =   5385
               Width           =   1290
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   39
               Left            =   3165
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   5385
               Width           =   1290
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
               Width           =   1290
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
               Width           =   975
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
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   26
               Top             =   5025
               Width           =   945
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   35
               Left            =   3150
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   5025
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   34
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   24
               Top             =   5025
               Width           =   1335
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
               Width           =   675
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
               Top             =   4665
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
               TabIndex        =   64
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
               TabIndex        =   353
               Top             =   4680
               Width           =   1260
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   37
               X1              =   1425
               X2              =   3960
               Y1              =   4860
               Y2              =   4860
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
               X2              =   8880
               Y1              =   4860
               Y2              =   4860
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
               TabIndex        =   343
               Top             =   4680
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   164
               X1              =   7560
               X2              =   8760
               Y1              =   900
               Y2              =   900
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
               TabIndex        =   341
               Top             =   720
               Width           =   540
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   138
               X1              =   6960
               X2              =   8880
               Y1              =   4485
               Y2              =   4485
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "QQ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   138
               Left            =   6780
               TabIndex        =   206
               Top             =   4305
               Width           =   180
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   137
               X1              =   1080
               X2              =   3000
               Y1              =   4485
               Y2              =   4485
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Email"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   137
               Left            =   600
               TabIndex        =   204
               Top             =   4305
               Width           =   450
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
               TabIndex        =   202
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
               TabIndex        =   160
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
               TabIndex        =   159
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
               TabIndex        =   157
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
               TabIndex        =   137
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
               TabIndex        =   136
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
               TabIndex        =   133
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
               TabIndex        =   131
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
               TabIndex        =   130
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
               TabIndex        =   127
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
               TabIndex        =   125
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
               TabIndex        =   98
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
               Left            =   4500
               TabIndex        =   97
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
               Left            =   2400
               TabIndex        =   96
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
               TabIndex        =   95
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
               TabIndex        =   94
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
               TabIndex        =   93
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
               TabIndex        =   92
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
               Left            =   4560
               TabIndex        =   91
               Top             =   5025
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
               Left            =   2685
               TabIndex        =   90
               Top             =   5025
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
               TabIndex        =   89
               Top             =   5025
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
               TabIndex        =   88
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
               TabIndex        =   87
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
               TabIndex        =   86
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
               TabIndex        =   85
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
               TabIndex        =   84
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
               TabIndex        =   83
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
               TabIndex        =   82
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
               TabIndex        =   81
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
               TabIndex        =   80
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
               TabIndex        =   79
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
               TabIndex        =   78
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
               TabIndex        =   77
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
               TabIndex        =   76
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
               TabIndex        =   75
               Top             =   1785
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
               TabIndex        =   74
               Top             =   4665
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
               TabIndex        =   73
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
               TabIndex        =   72
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
               TabIndex        =   71
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
               TabIndex        =   70
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
               TabIndex        =   69
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
               TabIndex        =   68
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
               TabIndex        =   67
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
               TabIndex        =   66
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
               TabIndex        =   65
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
               Y1              =   1960
               Y2              =   1960
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
               Y1              =   4845
               Y2              =   4845
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
               X2              =   1800
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
               X2              =   2640
               Y1              =   5205
               Y2              =   5205
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
               X1              =   3075
               X2              =   4560
               Y1              =   5205
               Y2              =   5205
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
               X1              =   4965
               X2              =   5965
               Y1              =   5205
               Y2              =   5205
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
               X2              =   2400
               Y1              =   5565
               Y2              =   5565
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   39
               X1              =   3120
               X2              =   4440
               Y1              =   5565
               Y2              =   5565
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   40
               X1              =   5280
               X2              =   6600
               Y1              =   5565
               Y2              =   5565
            End
         End
         Begin VB.Frame fraMain 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   西医诊断 "
            ForeColor       =   &H00FF0000&
            Height          =   6195
            Index           =   1
            Left            =   120
            TabIndex        =   114
            Tag             =   "6195"
            Top             =   120
            Width           =   9495
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   168
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   350
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
               TabIndex        =   335
               Top             =   5760
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
               TabIndex        =   333
               Top             =   5760
               Width           =   870
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "住院期间告病重或病危"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   18
               Left            =   240
               TabIndex        =   208
               Top             =   5400
               Width           =   2325
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "疑难病例"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   9
               Left            =   3000
               TabIndex        =   207
               Top             =   5400
               Width           =   1020
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   56
               Left            =   7155
               Locked          =   -1  'True
               TabIndex        =   175
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
               TabIndex        =   174
               Top             =   3120
               Width           =   1515
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   59
               Left            =   5280
               Locked          =   -1  'True
               TabIndex        =   172
               Top             =   4627
               Width           =   3810
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   57
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   170
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
               TabIndex        =   168
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
               TabIndex        =   166
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
               TabIndex        =   164
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
               TabIndex        =   154
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
               TabIndex        =   150
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
               TabIndex        =   142
               Top             =   5010
               Width           =   4770
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   58
               Left            =   3720
               Locked          =   -1  'True
               TabIndex        =   139
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
               TabIndex        =   138
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
               TabIndex        =   115
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
               FormatString    =   $"frmArchiveInMedRec_SC.frx":09D4
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
               TabIndex        =   351
               Top             =   4627
               Width           =   1080
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   145
               X1              =   960
               X2              =   4920
               Y1              =   5940
               Y2              =   5940
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
               TabIndex        =   336
               Top             =   5760
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   146
               X1              =   6360
               X2              =   7245
               Y1              =   5940
               Y2              =   5940
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
               TabIndex        =   334
               Top             =   5760
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
               TabIndex        =   176
               Top             =   3876
               Width           =   900
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   59
               X1              =   5280
               X2              =   9120
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
               Left            =   3555
               TabIndex        =   173
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
               TabIndex        =   171
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
               TabIndex        =   169
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
               TabIndex        =   167
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
               TabIndex        =   165
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
               TabIndex        =   163
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
               TabIndex        =   151
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
               TabIndex        =   143
               Top             =   5010
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   62
               X1              =   4245
               X2              =   9120
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
               TabIndex        =   140
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
               TabIndex        =   123
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
               TabIndex        =   122
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
               TabIndex        =   121
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
               TabIndex        =   120
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
               TabIndex        =   119
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
               TabIndex        =   118
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
               TabIndex        =   117
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
               TabIndex        =   116
               Top             =   3876
               Width           =   900
            End
         End
         Begin VB.Frame fraMain 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   住院情况 "
            ForeColor       =   &H00FF0000&
            Height          =   7650
            Index           =   4
            Left            =   120
            TabIndex        =   216
            Tag             =   "7650"
            Top             =   120
            Width           =   9495
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   126
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   345
               Top             =   5211
               Width           =   1440
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   109
               Left            =   7650
               Locked          =   -1  'True
               TabIndex        =   337
               Top             =   6015
               Width           =   1575
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "示教病案"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   10
               Left            =   540
               TabIndex        =   332
               Top             =   1469
               Width           =   1020
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   143
               Left            =   4800
               Locked          =   -1  'True
               TabIndex        =   268
               Top             =   4087
               Width           =   720
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   144
               Left            =   6360
               Locked          =   -1  'True
               TabIndex        =   267
               Top             =   4087
               Width           =   2400
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "会诊情况"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   22
               Left            =   180
               TabIndex        =   266
               Top             =   4080
               Width           =   1020
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   92
               Left            =   5415
               Locked          =   -1  'True
               TabIndex        =   265
               Top             =   2964
               Width           =   3480
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
               TabIndex        =   264
               Top             =   2964
               Width           =   2640
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   117
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   263
               Top             =   7080
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   116
               Left            =   1035
               Locked          =   -1  'True
               TabIndex        =   262
               Top             =   7080
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
               TabIndex        =   261
               Top             =   3708
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
               TabIndex        =   260
               Top             =   3708
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   82
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   259
               Top             =   1104
               Width           =   1200
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   79
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   258
               Top             =   732
               Width           =   1200
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   76
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   257
               Top             =   360
               Width           =   1200
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   74
               Left            =   960
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
               IMEMode         =   3  'DISABLE
               Index           =   101
               Left            =   2190
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   255
               Tag             =   "无"
               Text            =   "无"
               Top             =   4839
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
               TabIndex        =   254
               Top             =   4839
               Width           =   5940
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   90
               Left            =   3690
               Locked          =   -1  'True
               TabIndex        =   253
               Top             =   2592
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
               TabIndex        =   252
               Top             =   5211
               Width           =   720
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   100
               Left            =   8070
               Locked          =   -1  'True
               TabIndex        =   251
               Top             =   4467
               Width           =   435
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   99
               Left            =   7140
               Locked          =   -1  'True
               TabIndex        =   250
               Top             =   4467
               Width           =   435
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   98
               Left            =   6180
               Locked          =   -1  'True
               TabIndex        =   249
               Top             =   4467
               Width           =   480
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   97
               Left            =   4620
               Locked          =   -1  'True
               TabIndex        =   248
               Top             =   4467
               Width           =   435
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   95
               Left            =   2940
               Locked          =   -1  'True
               TabIndex        =   247
               Top             =   4467
               Width           =   480
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   114
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   246
               Top             =   6720
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   113
               Left            =   1035
               Locked          =   -1  'True
               TabIndex        =   245
               Top             =   6720
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   84
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   244
               Top             =   1476
               Width           =   1200
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   89
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   243
               Top             =   2220
               Width           =   1200
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   88
               Left            =   915
               TabIndex        =   242
               Top             =   2220
               Width           =   1440
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   87
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   241
               Top             =   2592
               Width           =   1440
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   86
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   240
               Top             =   1848
               Width           =   1200
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   85
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   239
               Top             =   1848
               Width           =   1440
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   104
               Left            =   7890
               Locked          =   -1  'True
               TabIndex        =   238
               Top             =   5211
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
               Left            =   6240
               TabIndex        =   237
               Top             =   5204
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
               TabIndex        =   236
               Top             =   732
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
               TabIndex        =   235
               Top             =   1104
               Width           =   1440
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   105
               Left            =   975
               Locked          =   -1  'True
               TabIndex        =   234
               Top             =   5583
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
               TabIndex        =   233
               Top             =   6015
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
               TabIndex        =   232
               Top             =   6375
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   108
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   231
               Top             =   6015
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   111
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   230
               Top             =   6375
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   106
               Left            =   7650
               Locked          =   -1  'True
               TabIndex        =   229
               Top             =   5583
               Width           =   1575
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   112
               Left            =   7650
               Locked          =   -1  'True
               TabIndex        =   228
               Top             =   6360
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   115
               Left            =   7680
               Locked          =   -1  'True
               TabIndex        =   227
               Top             =   6735
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
               TabIndex        =   226
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "科研病案"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   11
               Left            =   1860
               TabIndex        =   225
               Top             =   1469
               Width           =   1020
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   96
               Left            =   3750
               Locked          =   -1  'True
               TabIndex        =   224
               Top             =   4467
               Width           =   435
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   142
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   223
               Top             =   4087
               Width           =   720
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "进入路径"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   19
               Left            =   1560
               TabIndex        =   222
               Top             =   3329
               Width           =   1020
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "完成路径"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   20
               Left            =   2640
               TabIndex        =   221
               Top             =   3329
               Width           =   1020
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "变异"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   21
               Left            =   6000
               TabIndex        =   220
               Top             =   3329
               Width           =   780
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   140
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   219
               Top             =   3336
               Width           =   1335
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   139
               Left            =   6060
               Locked          =   -1  'True
               TabIndex        =   218
               Top             =   2592
               Width           =   1080
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   141
               Left            =   7560
               Locked          =   -1  'True
               TabIndex        =   217
               Top             =   3330
               Width           =   1335
            End
            Begin VSFlex8Ctl.VSFlexGrid vsTSJC 
               Height          =   2055
               Left            =   6480
               TabIndex        =   269
               Top             =   480
               Width           =   2415
               _cx             =   4260
               _cy             =   3625
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
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   8
               Cols            =   2
               FixedRows       =   0
               FixedCols       =   1
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_SC.frx":0B25
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
               Index           =   126
               X1              =   4320
               X2              =   5895
               Y1              =   5385
               Y2              =   5385
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "TNM分期"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   126
               Left            =   3630
               TabIndex        =   346
               Top             =   5205
               Width           =   630
            End
            Begin VB.Label lblVsTitle 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "检查情况"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   5
               Left            =   6450
               TabIndex        =   339
               Top             =   240
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   109
               X1              =   7650
               X2              =   9240
               Y1              =   6195
               Y2              =   6195
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   109
               Left            =   6900
               TabIndex        =   338
               Top             =   6015
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "院内会诊         次   外院会诊          次，其他"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   142
               Left            =   2010
               TabIndex        =   322
               Top             =   4087
               Width           =   4320
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   92
               X1              =   5415
               X2              =   9000
               Y1              =   3145
               Y2              =   3145
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "其他医学警示"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   92
               Left            =   4320
               TabIndex        =   321
               Top             =   2964
               Width           =   1080
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   91
               X1              =   975
               X2              =   3840
               Y1              =   3145
               Y2              =   3145
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医学警示"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   91
               Left            =   180
               TabIndex        =   320
               Top             =   2964
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   117
               X1              =   4320
               X2              =   5745
               Y1              =   7260
               Y2              =   7260
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病案质量"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   117
               Left            =   3540
               TabIndex        =   319
               Top             =   7080
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   116
               X1              =   1035
               X2              =   2460
               Y1              =   7260
               Y2              =   7260
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
               TabIndex        =   318
               Top             =   7080
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
               TabIndex        =   317
               Top             =   3708
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
               TabIndex        =   316
               Top             =   3708
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   93
               X1              =   915
               X2              =   2400
               Y1              =   3885
               Y2              =   3885
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   94
               X1              =   3690
               X2              =   9120
               Y1              =   3885
               Y2              =   3885
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   82
               X1              =   4320
               X2              =   5640
               Y1              =   1290
               Y2              =   1290
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "生育状况"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   82
               Left            =   3540
               TabIndex        =   315
               Top             =   1104
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   79
               X1              =   4320
               X2              =   5640
               Y1              =   915
               Y2              =   915
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "发病时间"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   79
               Left            =   3540
               TabIndex        =   314
               Top             =   732
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   76
               X1              =   4320
               X2              =   5640
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
               Left            =   3090
               TabIndex        =   313
               Top             =   360
               Width           =   1170
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   74
               X1              =   915
               X2              =   2400
               Y1              =   535
               Y2              =   535
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
               TabIndex        =   312
               Top             =   360
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   102
               X1              =   3000
               X2              =   9120
               Y1              =   5020
               Y2              =   5020
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
               TabIndex        =   311
               Top             =   4839
               Width           =   1800
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   101
               X1              =   2070
               X2              =   2650
               Y1              =   5020
               Y2              =   5020
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
               TabIndex        =   310
               Top             =   4839
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
               Left            =   4800
               TabIndex        =   309
               Top             =   2595
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
               Left            =   2940
               TabIndex        =   308
               Top             =   2595
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   90
               X1              =   3690
               X2              =   4800
               Y1              =   2775
               Y2              =   2775
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   103
               X1              =   1155
               X2              =   2040
               Y1              =   5385
               Y2              =   5385
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
               TabIndex        =   307
               Top             =   5211
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
               TabIndex        =   306
               Top             =   5211
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
               Left            =   8535
               TabIndex        =   305
               Top             =   4467
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   100
               X1              =   7935
               X2              =   8530
               Y1              =   4645
               Y2              =   4645
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "小时"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   99
               Left            =   7575
               TabIndex        =   304
               Top             =   4467
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   99
               X1              =   6990
               X2              =   7560
               Y1              =   4645
               Y2              =   4645
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "天"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   98
               Left            =   6840
               TabIndex        =   303
               Top             =   4467
               Width           =   180
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   98
               X1              =   6180
               X2              =   6760
               Y1              =   4645
               Y2              =   4645
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院后"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   198
               Left            =   5640
               TabIndex        =   302
               Top             =   4467
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
               Left            =   5175
               TabIndex        =   301
               Top             =   4467
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   97
               X1              =   4545
               X2              =   5115
               Y1              =   4645
               Y2              =   4645
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   96
               X1              =   3660
               X2              =   4260
               Y1              =   4645
               Y2              =   4645
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "天"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   95
               Left            =   3480
               TabIndex        =   300
               Top             =   4467
               Width           =   180
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   95
               X1              =   2940
               X2              =   3495
               Y1              =   4645
               Y2              =   4645
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
               TabIndex        =   299
               Top             =   4467
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
               Left            =   3540
               TabIndex        =   298
               Top             =   6720
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   114
               X1              =   4320
               X2              =   5745
               Y1              =   6900
               Y2              =   6900
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   113
               X1              =   1035
               X2              =   2460
               Y1              =   6900
               Y2              =   6900
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
               TabIndex        =   297
               Top             =   6699
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
               Left            =   3540
               TabIndex        =   296
               Top             =   1476
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   84
               X1              =   4320
               X2              =   5640
               Y1              =   1650
               Y2              =   1650
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "质控护士"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   115
               Left            =   6900
               TabIndex        =   295
               Top             =   6735
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "主诊医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   111
               Left            =   3540
               TabIndex        =   294
               Top             =   6375
               Width           =   720
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
               TabIndex        =   293
               Top             =   6327
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "实习医师"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   112
               Left            =   6900
               TabIndex        =   292
               Top             =   6375
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
               Left            =   3540
               TabIndex        =   291
               Top             =   6015
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
               TabIndex        =   290
               Top             =   5955
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
               Left            =   7080
               TabIndex        =   289
               Top             =   5583
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
               TabIndex        =   288
               Top             =   5583
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
               Left            =   3720
               TabIndex        =   287
               Top             =   2220
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
               Left            =   2520
               TabIndex        =   286
               Top             =   2220
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
               TabIndex        =   285
               Top             =   2220
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
               Left            =   2520
               TabIndex        =   284
               Top             =   2640
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
               Left            =   360
               TabIndex        =   283
               Top             =   2592
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
               Left            =   5640
               TabIndex        =   282
               Top             =   1848
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
               Left            =   3540
               TabIndex        =   281
               Top             =   1848
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
               Left            =   2520
               TabIndex        =   280
               Top             =   1845
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
               TabIndex        =   279
               Top             =   1848
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
               TabIndex        =   278
               Top             =   1104
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
               TabIndex        =   277
               Top             =   732
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
               Left            =   7140
               TabIndex        =   276
               Top             =   5211
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   104
               X1              =   7770
               X2              =   9345
               Y1              =   5406
               Y2              =   5406
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   77
               X1              =   915
               X2              =   2400
               Y1              =   910
               Y2              =   910
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   85
               X1              =   915
               X2              =   2400
               Y1              =   2025
               Y2              =   2025
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   88
               X1              =   915
               X2              =   2400
               Y1              =   2400
               Y2              =   2400
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   80
               X1              =   915
               X2              =   2400
               Y1              =   1290
               Y2              =   1290
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   86
               X1              =   4320
               X2              =   5640
               Y1              =   2025
               Y2              =   2025
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   89
               X1              =   4320
               X2              =   5640
               Y1              =   2400
               Y2              =   2400
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   87
               X1              =   915
               X2              =   2400
               Y1              =   2775
               Y2              =   2775
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   105
               X1              =   975
               X2              =   2400
               Y1              =   5760
               Y2              =   5760
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   107
               X1              =   1680
               X2              =   3105
               Y1              =   6195
               Y2              =   6195
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   110
               X1              =   915
               X2              =   2340
               Y1              =   6555
               Y2              =   6555
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   108
               X1              =   4320
               X2              =   5745
               Y1              =   6195
               Y2              =   6195
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   111
               X1              =   4320
               X2              =   5745
               Y1              =   6555
               Y2              =   6555
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   106
               X1              =   7650
               X2              =   9240
               Y1              =   5760
               Y2              =   5760
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   112
               X1              =   7650
               X2              =   9240
               Y1              =   6555
               Y2              =   6555
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   115
               X1              =   7650
               X2              =   9240
               Y1              =   6915
               Y2              =   6915
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "小时"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   96
               Left            =   4215
               TabIndex        =   275
               Top             =   4467
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "临床路径信息："
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   153
               Left            =   180
               TabIndex        =   274
               Top             =   3336
               Width           =   1260
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "退出原因"
               Height          =   180
               Index           =   140
               Left            =   3720
               TabIndex        =   273
               Top             =   3336
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "变异原因"
               Height          =   180
               Index           =   141
               Left            =   6840
               TabIndex        =   272
               Top             =   3336
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   139
               X1              =   6060
               X2              =   7230
               Y1              =   2770
               Y2              =   2770
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输白蛋白"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   139
               Left            =   5280
               TabIndex        =   271
               Top             =   2592
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "g"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   239
               Left            =   7320
               TabIndex        =   270
               Top             =   2595
               Width           =   90
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   140
               X1              =   4440
               X2              =   5880
               Y1              =   3510
               Y2              =   3510
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   141
               X1              =   7560
               X2              =   9000
               Y1              =   3525
               Y2              =   3525
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   142
               X1              =   2760
               X2              =   3480
               Y1              =   4265
               Y2              =   4265
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   143
               X1              =   4800
               X2              =   5520
               Y1              =   4265
               Y2              =   4265
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   144
               X1              =   6360
               X2              =   8880
               Y1              =   4265
               Y2              =   4265
            End
         End
         Begin VB.Frame fraMain 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   附页2"
            ForeColor       =   &H00FF0000&
            Height          =   6090
            Index           =   7
            Left            =   120
            TabIndex        =   186
            Tag             =   "6090"
            Top             =   240
            Width           =   9495
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   167
               Left            =   6855
               Locked          =   -1  'True
               TabIndex        =   348
               Top             =   2160
               Width           =   975
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "上一次住本院与本次住院是因同一疾病(主要诊断)"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   30
               Left            =   5040
               TabIndex        =   347
               Top             =   2520
               Width           =   4260
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "住院期间身体约束"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   23
               Left            =   5520
               TabIndex        =   331
               Top             =   1440
               Width           =   1740
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   136
               Left            =   7935
               Locked          =   -1  'True
               TabIndex        =   329
               Top             =   1800
               Width           =   1455
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   135
               Left            =   6255
               Locked          =   -1  'True
               TabIndex        =   327
               Top             =   1080
               Width           =   1455
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   134
               Left            =   6255
               Locked          =   -1  'True
               TabIndex        =   325
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   83
               Left            =   6255
               Locked          =   -1  'True
               TabIndex        =   323
               Top             =   360
               Width           =   1455
            End
            Begin VB.Frame fraSplit 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00FFFFFF&
               Height          =   75
               Index           =   7
               Left            =   1320
               TabIndex        =   196
               Top             =   3240
               Width           =   7815
            End
            Begin VB.Frame fraAdvEvent 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "不良事件"
               ForeColor       =   &H80000008&
               Height          =   2715
               Left            =   240
               TabIndex        =   190
               Top             =   240
               Width           =   4335
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   121
                  Left            =   1440
                  Locked          =   -1  'True
                  TabIndex        =   200
                  Top             =   2220
                  Width           =   1995
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   120
                  Left            =   1440
                  Locked          =   -1  'True
                  TabIndex        =   199
                  Top             =   1860
                  Width           =   1995
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   119
                  Left            =   2880
                  Locked          =   -1  'True
                  TabIndex        =   198
                  Top             =   1500
                  Width           =   915
               End
               Begin VB.TextBox txtInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   118
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   197
                  Top             =   1500
                  Width           =   975
               End
               Begin VB.ListBox lstAdvEvent 
                  Height          =   960
                  ItemData        =   "frmArchiveInMedRec_SC.frx":0B9D
                  Left            =   120
                  List            =   "frmArchiveInMedRec_SC.frx":0B9F
                  TabIndex        =   191
                  Top             =   240
                  Width           =   3765
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   121
                  X1              =   1440
                  X2              =   3840
                  Y1              =   2400
                  Y2              =   2400
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   120
                  X1              =   1440
                  X2              =   3840
                  Y1              =   2040
                  Y2              =   2040
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   119
                  X1              =   2895
                  X2              =   3845
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
                  Enabled         =   0   'False
                  Height          =   180
                  Index           =   119
                  Left            =   2520
                  TabIndex        =   195
                  Top             =   1500
                  Width           =   360
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "压疮发生期间"
                  Enabled         =   0   'False
                  Height          =   180
                  Index           =   118
                  Left            =   120
                  TabIndex        =   194
                  Top             =   1500
                  Width           =   1080
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "跌倒或坠床原因"
                  Enabled         =   0   'False
                  Height          =   180
                  Index           =   121
                  Left            =   120
                  TabIndex        =   193
                  Top             =   2220
                  Width           =   1260
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "跌倒或坠床伤害"
                  Enabled         =   0   'False
                  Height          =   180
                  Index           =   120
                  Left            =   120
                  TabIndex        =   192
                  Top             =   1860
                  Width           =   1260
               End
            End
            Begin VB.PictureBox picSize 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   7
               Left            =   180
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   187
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfMain 
               Height          =   2490
               Left            =   120
               TabIndex        =   189
               Top             =   3480
               Width           =   9165
               _cx             =   16166
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
               Cols            =   12
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
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "距上一次住本院的时间            天"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   167
               Left            =   4980
               TabIndex        =   349
               Top             =   2160
               Width           =   3060
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   167
               X1              =   6840
               X2              =   7860
               Y1              =   2340
               Y2              =   2340
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   136
               X1              =   7935
               X2              =   9420
               Y1              =   1980
               Y2              =   1980
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "离院时透析（血透、腹透）尿素氮值"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   136
               Left            =   5040
               TabIndex        =   330
               Top             =   1800
               Width           =   2880
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   135
               X1              =   6255
               X2              =   7740
               Y1              =   1260
               Y2              =   1260
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "临床表现"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   135
               Left            =   5520
               TabIndex        =   328
               Top             =   1080
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   134
               X1              =   6255
               X2              =   7740
               Y1              =   900
               Y2              =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "引发药物"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   134
               Left            =   5520
               TabIndex        =   326
               Top             =   720
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   83
               X1              =   6255
               X2              =   7740
               Y1              =   540
               Y2              =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输液反应"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   83
               Left            =   5520
               TabIndex        =   324
               Top             =   360
               Width           =   720
            End
            Begin VB.Label lblVsTitle 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病案附加项目"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   4
               Left            =   240
               TabIndex        =   188
               Top             =   3180
               Width           =   1080
            End
         End
         Begin VB.Frame fraMain 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   过敏与手术 "
            ForeColor       =   &H00FF0000&
            Height          =   3345
            Index           =   3
            Left            =   120
            TabIndex        =   99
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
               TabIndex        =   100
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
               Cols            =   24
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmArchiveInMedRec_SC.frx":0BA1
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
               FormatString    =   $"frmArchiveInMedRec_SC.frx":0E7D
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
         Begin VB.Frame fraMain 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   附页1 "
            ForeColor       =   &H00FF0000&
            Height          =   7170
            Index           =   6
            Left            =   120
            TabIndex        =   179
            Tag             =   "7170"
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
               TabIndex        =   180
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VSFlex8Ctl.VSFlexGrid vsKSS 
               Height          =   1095
               Left            =   120
               TabIndex        =   181
               Top             =   480
               Width           =   9120
               _cx             =   16087
               _cy             =   1931
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
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_SC.frx":0EEA
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
            Begin VSFlex8Ctl.VSFlexGrid vsFlxAddICU 
               Height          =   1305
               Left            =   120
               TabIndex        =   209
               Top             =   2040
               Width           =   9120
               _cx             =   16087
               _cy             =   2302
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
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_SC.frx":0FD1
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
            Begin VSFlex8Ctl.VSFlexGrid vsICUInstruments 
               Height          =   1305
               Left            =   120
               TabIndex        =   211
               Top             =   3840
               Width           =   9120
               _cx             =   16087
               _cy             =   2302
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
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_SC.frx":109B
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
            Begin VSFlex8Ctl.VSFlexGrid vsInfect 
               Height          =   1305
               Left            =   120
               TabIndex        =   213
               Top             =   5520
               Width           =   3360
               _cx             =   5927
               _cy             =   2302
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
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_SC.frx":115E
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
            Begin VSFlex8Ctl.VSFlexGrid vsSample 
               Height          =   1305
               Left            =   4200
               TabIndex        =   215
               Top             =   5520
               Width           =   5040
               _cx             =   8890
               _cy             =   2302
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
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_SC.frx":11CE
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
               Caption         =   "标本来源"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   9
               Left            =   4200
               TabIndex        =   214
               Top             =   5280
               Width           =   720
            End
            Begin VB.Label lblVsTitle 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医院感染情况"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   8
               Left            =   120
               TabIndex        =   212
               Top             =   5280
               Width           =   1080
            End
            Begin VB.Label lblVsTitle 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "患者入住重症监护室期间器械使用情况"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   7
               Left            =   120
               TabIndex        =   210
               Top             =   3600
               Width           =   3060
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
               TabIndex        =   183
               Top             =   240
               Width           =   3150
            End
            Begin VB.Label lblVsTitle 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "患者入住重症监护病房记录"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   6
               Left            =   120
               TabIndex        =   182
               Top             =   1800
               Width           =   2160
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
            TabIndex        =   152
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
               TabIndex        =   153
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VSFlex8Ctl.VSFlexGrid vsChemoth 
               Height          =   1635
               Left            =   120
               TabIndex        =   184
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
               FormatString    =   $"frmArchiveInMedRec_SC.frx":123F
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
               TabIndex        =   185
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
               FormatString    =   $"frmArchiveInMedRec_SC.frx":1355
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
               TabIndex        =   178
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
               TabIndex        =   177
               Top             =   240
               Width           =   1080
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
            TabIndex        =   101
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
               TabIndex        =   108
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
                  TabIndex        =   148
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
                  TabIndex        =   146
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
                  TabIndex        =   144
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
                  TabIndex        =   149
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
                  TabIndex        =   147
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
                  TabIndex        =   145
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
                  TabIndex        =   111
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
                  TabIndex        =   110
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
                  TabIndex        =   109
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
               TabIndex        =   107
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
               TabIndex        =   106
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
               TabIndex        =   102
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
                  TabIndex        =   105
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
                  TabIndex        =   104
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
                  TabIndex        =   103
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
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmArchiveInMedRec_SC.frx":147B
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
               TabIndex        =   113
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
               TabIndex        =   112
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
Attribute VB_Name = "frmArchiveInMedRec_SC"
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
    Call vsAller.ShowCell(NewRow, NewCol)
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

Private Sub vsFlxAddICU_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= 0 And NewCol >= 0 Then Call vsFlxAddICU.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= 0 And NewCol >= 0 Then Call vsfMain.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsICUInstruments_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= 0 And NewCol >= 0 Then Call vsICUInstruments.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsInfect_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= 0 And NewCol >= 0 Then Call vsInfect.ShowCell(NewRow, NewCol)
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

Private Sub vsSample_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= 0 And NewCol >= 0 Then Call vsSample.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsTSJC_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= 0 And NewCol >= 0 Then Call vsTSJC.ShowCell(NewRow, NewCol)
End Sub



