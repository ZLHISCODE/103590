VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Begin VB.Form frmReport 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   16755
   ClientLeft      =   26535
   ClientTop       =   -2100
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   ScaleHeight     =   16755
   ScaleWidth      =   14670
   ShowInTaskbar   =   0   'False
   Begin zlSubclass.Subclass Subclass 
      Left            =   900
      Top             =   3495
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.HScrollBar hsbReport 
      Height          =   255
      LargeChange     =   500
      Left            =   0
      Max             =   100
      SmallChange     =   10
      TabIndex        =   0
      Top             =   0
      Width           =   8535
   End
   Begin VB.VScrollBar vsbReport 
      Height          =   7335
      LargeChange     =   50
      Left            =   0
      Max             =   100
      SmallChange     =   10
      TabIndex        =   157
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   15640
      Left            =   1800
      ScaleHeight     =   15615
      ScaleWidth      =   11865
      TabIndex        =   158
      Top             =   -480
      WhatsThisHelpID =   -480
      Width           =   11895
      Begin MSComCtl2.MonthView MView 
         Height          =   2220
         Left            =   10920
         TabIndex        =   220
         Tag             =   "0"
         Top             =   1200
         Visible         =   0   'False
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   3916
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Appearance      =   1
         StartOfWeek     =   251199490
         CurrentDate     =   41981
      End
      Begin VB.PictureBox picPane 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1065
         Index           =   0
         Left            =   960
         ScaleHeight     =   1065
         ScaleWidth      =   9870
         TabIndex        =   216
         Top             =   1080
         Width           =   9875
         Begin VB.TextBox txtNumber 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   1200
            TabIndex        =   1
            Top             =   660
            Width           =   1455
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckType 
            Height          =   270
            Index           =   0
            Left            =   6225
            TabIndex        =   2
            Top             =   615
            Width           =   1575
            _ExtentX        =   56118
            _ExtentY        =   476
            Checked         =   -1  'True
            BackColor       =   -2147483643
            Caption         =   "1�� ���α���"
            BoxVisible      =   0   'False
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckType 
            Height          =   270
            Index           =   1
            Left            =   7815
            TabIndex        =   3
            Top             =   615
            Width           =   1575
            _ExtentX        =   132318
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "2����������"
            BoxVisible      =   0   'False
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "�������"
            Height          =   180
            Index           =   1
            Left            =   5295
            TabIndex        =   219
            Top             =   660
            Width           =   900
         End
         Begin VB.Line LineNumber 
            X1              =   1215
            X2              =   2640
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�л����񹲺͹���Ⱦ�����濨"
            BeginProperty Font 
               Name            =   "����"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2640
            TabIndex        =   218
            Top             =   0
            Width           =   4875
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��Ƭ��ţ�"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   217
            Top             =   660
            Width           =   900
         End
      End
      Begin VB.PictureBox picPane 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5415
         Index           =   1
         Left            =   990
         ScaleHeight     =   5415
         ScaleWidth      =   9870
         TabIndex        =   176
         Top             =   2355
         Width           =   9875
         Begin VB.TextBox txtAddress 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   1620
            TabIndex        =   33
            Top             =   1170
            Width           =   4185
         End
         Begin VB.TextBox txtPhone 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   6720
            TabIndex        =   34
            Top             =   1170
            Width           =   1815
         End
         Begin VB.TextBox txtDiagnose 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   3
            Left            =   4035
            TabIndex        =   77
            Top             =   4770
            Width           =   525
         End
         Begin VB.TextBox txtDeath 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   0
            Left            =   1215
            TabIndex        =   78
            Top             =   5130
            Width           =   1095
         End
         Begin VB.TextBox txtDeath 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   1
            Left            =   2505
            TabIndex        =   79
            Top             =   5130
            Width           =   615
         End
         Begin VB.TextBox txtDeath 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   2
            Left            =   3300
            TabIndex        =   80
            Top             =   5130
            Width           =   525
         End
         Begin VB.TextBox txtDiagnose 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   0
            Left            =   1215
            TabIndex        =   74
            Top             =   4770
            Width           =   1095
         End
         Begin VB.TextBox txtDiagnose 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   1
            Left            =   2505
            TabIndex        =   75
            Top             =   4770
            Width           =   615
         End
         Begin VB.TextBox txtDiagnose 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   2
            Left            =   3300
            TabIndex        =   76
            Top             =   4770
            Width           =   525
         End
         Begin VB.TextBox txtAttack 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   0
            Left            =   1200
            TabIndex        =   71
            Top             =   4410
            Width           =   1095
         End
         Begin VB.TextBox txtAttack 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   1
            Left            =   2505
            TabIndex        =   72
            Top             =   4410
            Width           =   615
         End
         Begin VB.TextBox txtAttack 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   2
            Left            =   3300
            TabIndex        =   73
            Top             =   4410
            Width           =   525
         End
         Begin VB.TextBox txtAddInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   0
            Left            =   1545
            TabIndex        =   41
            Top             =   1875
            Width           =   690
         End
         Begin VB.TextBox txtAddInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   1
            Left            =   2415
            TabIndex        =   42
            Top             =   1875
            Width           =   885
         End
         Begin VB.TextBox txtAddInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   3
            Left            =   5040
            TabIndex        =   44
            Top             =   1875
            Width           =   885
         End
         Begin VB.TextBox txtAddInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   4
            Left            =   7050
            TabIndex        =   45
            Top             =   1875
            Width           =   915
         End
         Begin VB.TextBox txtAddInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   5
            Left            =   8160
            TabIndex        =   46
            Top             =   1875
            Width           =   690
         End
         Begin VB.TextBox txtParentName 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   3735
            TabIndex        =   5
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   0
            Left            =   1455
            TabIndex        =   6
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   1
            Left            =   1755
            TabIndex        =   7
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   2
            Left            =   2055
            TabIndex        =   8
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   3
            Left            =   2355
            TabIndex        =   9
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   4
            Left            =   2655
            TabIndex        =   10
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   5
            Left            =   2955
            TabIndex        =   11
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   6
            Left            =   3255
            TabIndex        =   12
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   7
            Left            =   3555
            TabIndex        =   13
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   8
            Left            =   3870
            TabIndex        =   14
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   9
            Left            =   4170
            TabIndex        =   15
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   10
            Left            =   4470
            TabIndex        =   16
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   11
            Left            =   4770
            TabIndex        =   17
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   12
            Left            =   5070
            TabIndex        =   18
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   13
            Left            =   5370
            TabIndex        =   19
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   14
            Left            =   5670
            TabIndex        =   20
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   15
            Left            =   5970
            TabIndex        =   21
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   16
            Left            =   6270
            TabIndex        =   22
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtIDCard 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   17
            Left            =   6585
            TabIndex        =   23
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox txtBirth 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   0
            Left            =   1215
            TabIndex        =   26
            Top             =   825
            Width           =   1095
         End
         Begin VB.TextBox txtBirth 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   1
            Left            =   2490
            TabIndex        =   27
            Top             =   825
            Width           =   615
         End
         Begin VB.TextBox txtBirth 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   2
            Left            =   3285
            TabIndex        =   28
            Top             =   825
            Width           =   525
         End
         Begin VB.TextBox txtAge 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   6600
            TabIndex        =   29
            Top             =   825
            Width           =   525
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   750
            TabIndex        =   4
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox txtAddInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   2
            Left            =   3525
            TabIndex        =   43
            Top             =   1875
            Width           =   855
         End
         Begin zlDisReportCardEx.uCheckNorm ucCaseType2 
            Height          =   270
            Index           =   1
            Left            =   2310
            TabIndex        =   70
            Top             =   4020
            Width           =   690
            _ExtentX        =   134144
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "����"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCaseType1 
            Height          =   270
            Index           =   3
            Left            =   5440
            TabIndex        =   68
            Top             =   3660
            Width           =   1350
            _ExtentX        =   135308
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��ԭЯ����"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCaseType1 
            Height          =   270
            Index           =   2
            Left            =   4245
            TabIndex        =   67
            Top             =   3660
            Width           =   1170
            _ExtentX        =   134990
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "ȷ�ﲡ����"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCaseType1 
            Height          =   270
            Index           =   1
            Left            =   2700
            TabIndex        =   66
            Top             =   3660
            Width           =   1530
            _ExtentX        =   135625
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�ٴ���ϲ�����"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   17
            Left            =   1420
            TabIndex        =   64
            Top             =   3255
            Width           =   810
            _ExtentX        =   131815
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "����"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   16
            Left            =   60
            TabIndex        =   63
            Top             =   3255
            Width           =   1350
            _ExtentX        =   132768
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "������ ����"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   15
            Left            =   8235
            TabIndex        =   62
            Top             =   2895
            Width           =   1350
            _ExtentX        =   132768
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���񼰴�ҵ��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   14
            Left            =   7020
            TabIndex        =   61
            Top             =   2895
            Width           =   1170
            _ExtentX        =   134990
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "������Ա��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   13
            Left            =   5790
            TabIndex        =   60
            Top             =   2895
            Width           =   1170
            _ExtentX        =   132450
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�ɲ�ְԱ��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   12
            Left            =   4605
            TabIndex        =   59
            Top             =   2895
            Width           =   1170
            _ExtentX        =   132450
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��(��)��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   11
            Left            =   3795
            TabIndex        =   58
            Top             =   2895
            Width           =   810
            _ExtentX        =   131815
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "����"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   10
            Left            =   2970
            TabIndex        =   57
            Top             =   2895
            Width           =   810
            _ExtentX        =   131815
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "ũ��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   9
            Left            =   2160
            TabIndex        =   56
            Top             =   2895
            Width           =   810
            _ExtentX        =   131815
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�񹤡�"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   7
            Left            =   60
            TabIndex        =   54
            Top             =   2895
            Width           =   1170
            _ExtentX        =   132450
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "ҽ����Ա��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   6
            Left            =   7900
            TabIndex        =   53
            Top             =   2535
            Width           =   1170
            _ExtentX        =   134990
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��ҵ����"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   5
            Left            =   6540
            TabIndex        =   52
            Top             =   2520
            Width           =   1350
            _ExtentX        =   132768
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "����ʳƷҵ��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   4
            Left            =   4995
            TabIndex        =   51
            Top             =   2520
            Width           =   1530
            _ExtentX        =   133085
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "����Ա����ķ��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   3
            Left            =   4180
            TabIndex        =   50
            Top             =   2520
            Width           =   810
            _ExtentX        =   131815
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��ʦ��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   2
            Left            =   2445
            TabIndex        =   49
            Top             =   2520
            Width           =   1710
            _ExtentX        =   133403
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "ѧ��(����Сѧ)��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   1
            Left            =   1275
            TabIndex        =   48
            Top             =   2520
            Width           =   1170
            _ExtentX        =   132450
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "ɢ�Ӷ�ͯ��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucAge 
            Height          =   270
            Index           =   0
            Left            =   8145
            TabIndex        =   30
            Top             =   780
            Width           =   465
            _ExtentX        =   130784
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucSex 
            Height          =   270
            Index           =   0
            Left            =   7590
            TabIndex        =   24
            Top             =   405
            Width           =   570
            _ExtentX        =   337979
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucSex 
            Height          =   270
            Index           =   1
            Left            =   8175
            TabIndex        =   25
            Top             =   405
            Width           =   570
            _ExtentX        =   337979
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "Ů"
         End
         Begin zlDisReportCardEx.uCheckNorm ucAge 
            Height          =   270
            Index           =   1
            Left            =   8610
            TabIndex        =   31
            Top             =   780
            Width           =   465
            _ExtentX        =   130784
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucAge 
            Height          =   270
            Index           =   2
            Left            =   9075
            TabIndex        =   32
            Top             =   780
            Width           =   555
            _ExtentX        =   130942
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��)"
         End
         Begin zlDisReportCardEx.uCheckNorm ucFrom 
            Height          =   270
            Index           =   0
            Left            =   1215
            TabIndex        =   35
            Top             =   1485
            Width           =   855
            _ExtentX        =   82788
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "������"
         End
         Begin zlDisReportCardEx.uCheckNorm ucFrom 
            Height          =   270
            Index           =   1
            Left            =   2115
            TabIndex        =   36
            Top             =   1485
            Width           =   1425
            _ExtentX        =   339487
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "������������"
         End
         Begin zlDisReportCardEx.uCheckNorm ucFrom 
            Height          =   270
            Index           =   2
            Left            =   3615
            TabIndex        =   37
            Top             =   1485
            Width           =   1380
            _ExtentX        =   339619
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��ʡ��������"
         End
         Begin zlDisReportCardEx.uCheckNorm ucFrom 
            Height          =   270
            Index           =   3
            Left            =   5055
            TabIndex        =   38
            Top             =   1485
            Width           =   675
            _ExtentX        =   338376
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��ʡ"
         End
         Begin zlDisReportCardEx.uCheckNorm ucFrom 
            Height          =   270
            Index           =   4
            Left            =   5775
            TabIndex        =   39
            Top             =   1485
            Width           =   885
            _ExtentX        =   338746
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�۰�̨"
         End
         Begin zlDisReportCardEx.uCheckNorm ucFrom 
            Height          =   270
            Index           =   5
            Left            =   6735
            TabIndex        =   40
            Top             =   1485
            Width           =   675
            _ExtentX        =   338376
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�⼮"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   0
            Left            =   60
            TabIndex        =   47
            Top             =   2520
            Width           =   1170
            _ExtentX        =   134990
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���ж�ͯ��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCheckJob 
            Height          =   270
            Index           =   8
            Left            =   1275
            TabIndex        =   55
            Top             =   2895
            Width           =   810
            _ExtentX        =   131815
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���ˡ�"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCaseType1 
            Height          =   270
            Index           =   0
            Left            =   1515
            TabIndex        =   65
            Top             =   3660
            Width           =   1170
            _ExtentX        =   134990
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���Ʋ�����"
         End
         Begin zlDisReportCardEx.uCheckNorm ucCaseType2 
            Height          =   270
            Index           =   0
            Left            =   1515
            TabIndex        =   69
            Top             =   4020
            Width           =   810
            _ExtentX        =   134355
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���ԡ�"
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "(���͸���*��Ѫ���没*������)"
            Height          =   180
            Index           =   2
            Left            =   2970
            TabIndex        =   215
            Top             =   4065
            Width           =   2520
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "(1)"
            Height          =   180
            Index           =   3
            Left            =   1125
            TabIndex        =   214
            Top             =   3705
            Width           =   270
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��(��)"
            Height          =   180
            Index           =   16
            Left            =   4395
            TabIndex        =   213
            Top             =   1875
            Width           =   540
         End
         Begin VB.Line Line1 
            Index           =   23
            X1              =   3300
            X2              =   3845
            Y1              =   5310
            Y2              =   5310
         End
         Begin VB.Line Line1 
            Index           =   22
            X1              =   2505
            X2              =   3120
            Y1              =   5310
            Y2              =   5310
         End
         Begin VB.Line Line1 
            Index           =   21
            X1              =   4035
            X2              =   4595
            Y1              =   4950
            Y2              =   4950
         End
         Begin VB.Line Line1 
            Index           =   20
            X1              =   3300
            X2              =   3835
            Y1              =   4950
            Y2              =   4950
         End
         Begin VB.Line Line1 
            Index           =   19
            X1              =   2505
            X2              =   3120
            Y1              =   4950
            Y2              =   4950
         End
         Begin VB.Line Line1 
            Index           =   18
            X1              =   3300
            X2              =   3835
            Y1              =   4590
            Y2              =   4590
         End
         Begin VB.Line Line1 
            Index           =   17
            X1              =   2505
            X2              =   3120
            Y1              =   4590
            Y2              =   4590
         End
         Begin VB.Line Line1 
            Index           =   16
            X1              =   1215
            X2              =   2390
            Y1              =   5310
            Y2              =   5310
         End
         Begin VB.Line Line1 
            Index           =   15
            X1              =   1215
            X2              =   2310
            Y1              =   4950
            Y2              =   4950
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   1215
            X2              =   2310
            Y1              =   4590
            Y2              =   4590
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   1620
            X2              =   5770
            Y1              =   1365
            Y2              =   1365
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��ϵ�绰��"
            Height          =   180
            Index           =   11
            Left            =   5835
            TabIndex        =   212
            Top             =   1170
            Width           =   900
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "������λ(ѧУ)��"
            Height          =   180
            Index           =   10
            Left            =   120
            TabIndex        =   211
            Top             =   1170
            Width           =   1440
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��������*��"
            Height          =   180
            Index           =   21
            Left            =   120
            TabIndex        =   210
            Top             =   3705
            Width           =   990
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "(2)"
            Height          =   180
            Index           =   22
            Left            =   1125
            TabIndex        =   209
            Top             =   4065
            Width           =   270
         End
         Begin VB.Line Line1 
            Index           =   11
            X1              =   3480
            X2              =   4380
            Y1              =   2055
            Y2              =   2055
         End
         Begin VB.Label lblDiagnose 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "ʱ"
            Height          =   180
            Index           =   3
            Left            =   4605
            TabIndex        =   208
            Top             =   4770
            Width           =   180
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "�������� ��"
            Height          =   180
            Index           =   24
            Left            =   120
            TabIndex        =   207
            Top             =   5130
            Width           =   990
         End
         Begin VB.Label lblDeath 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   2310
            TabIndex        =   206
            Top             =   5130
            Width           =   180
         End
         Begin VB.Label lblDeath 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   3120
            TabIndex        =   205
            Top             =   5130
            Width           =   180
         End
         Begin VB.Label lblDeath 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   2
            Left            =   3825
            TabIndex        =   204
            Top             =   5130
            Width           =   180
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "���ʱ��*��"
            Height          =   180
            Index           =   23
            Left            =   120
            TabIndex        =   203
            Top             =   4770
            Width           =   990
         End
         Begin VB.Label lblDiagnose 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   2310
            TabIndex        =   202
            Top             =   4770
            Width           =   180
         End
         Begin VB.Label lblDiagnose 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   3120
            TabIndex        =   201
            Top             =   4770
            Width           =   180
         End
         Begin VB.Label lblDiagnose 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   2
            Left            =   3825
            TabIndex        =   200
            Top             =   4770
            Width           =   180
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��������*��"
            Height          =   180
            Index           =   25
            Left            =   120
            TabIndex        =   199
            Top             =   4410
            Width           =   990
         End
         Begin VB.Label lblAttack 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   2310
            TabIndex        =   198
            Top             =   4410
            Width           =   180
         End
         Begin VB.Label lblAttack 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   3120
            TabIndex        =   197
            Top             =   4410
            Width           =   180
         End
         Begin VB.Label lblAttack 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   2
            Left            =   3825
            TabIndex        =   196
            Top             =   4410
            Width           =   180
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��סַ(����)*��"
            Height          =   180
            Index           =   13
            Left            =   120
            TabIndex        =   195
            Top             =   1875
            Width           =   1350
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   1545
            X2              =   2215
            Y1              =   2055
            Y2              =   2055
         End
         Begin VB.Line Line1 
            Index           =   10
            X1              =   2415
            X2              =   3315
            Y1              =   2055
            Y2              =   2055
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "ʡ"
            Height          =   180
            Index           =   14
            Left            =   2235
            TabIndex        =   194
            Top             =   1875
            Width           =   180
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   15
            Left            =   3300
            TabIndex        =   193
            Tag             =   "301,281"
            Top             =   1875
            Width           =   180
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   5040
            X2              =   5940
            Y1              =   2055
            Y2              =   2055
         End
         Begin VB.Line Line1 
            Index           =   13
            X1              =   7035
            X2              =   7935
            Y1              =   2055
            Y2              =   2055
         End
         Begin VB.Line Line1 
            Index           =   14
            X1              =   8160
            X2              =   8865
            Y1              =   2055
            Y2              =   2055
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��(�򡢽ֵ�)"
            Height          =   180
            Index           =   17
            Left            =   5940
            TabIndex        =   192
            Top             =   1875
            Width           =   1080
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   18
            Left            =   7965
            TabIndex        =   191
            Top             =   1875
            Width           =   180
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "(���ƺ�)"
            Height          =   180
            Index           =   19
            Left            =   8850
            TabIndex        =   190
            Top             =   1875
            Width           =   720
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��Ⱥ����*��"
            Height          =   180
            Index           =   20
            Left            =   120
            TabIndex        =   189
            Top             =   2235
            Width           =   990
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��������*��"
            Height          =   180
            Index           =   12
            Left            =   120
            TabIndex        =   188
            Top             =   1530
            Width           =   990
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����*��"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   187
            Top             =   120
            Width           =   630
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   740
            X2              =   2175
            Y1              =   300
            Y2              =   300
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "(�����ҳ�������"
            Height          =   180
            Index           =   5
            Left            =   2340
            TabIndex        =   186
            Top             =   120
            Width           =   1350
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   3735
            X2              =   5195
            Y1              =   300
            Y2              =   300
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   ")"
            Height          =   180
            Index           =   6
            Left            =   5295
            TabIndex        =   185
            Top             =   120
            Width           =   90
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��Ч֤����*��"
            Height          =   180
            Index           =   7
            Left            =   120
            TabIndex        =   184
            Top             =   465
            Width           =   1170
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "�Ա�*:"
            Height          =   180
            Index           =   8
            Left            =   6960
            TabIndex        =   183
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��������*��"
            Height          =   180
            Index           =   9
            Left            =   120
            TabIndex        =   182
            Top             =   825
            Width           =   990
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   1215
            X2              =   2295
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Label lblBirth 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   2310
            TabIndex        =   181
            Top             =   825
            Width           =   180
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   2460
            X2              =   3125
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Label lblBirth 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   3105
            TabIndex        =   180
            Top             =   825
            Width           =   180
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   3285
            X2              =   3780
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Label lblBirth 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   2
            Left            =   3870
            TabIndex        =   179
            Top             =   825
            Width           =   180
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "(��������ڲ��꣬ʵ�����䣺"
            Height          =   180
            Index           =   26
            Left            =   4140
            TabIndex        =   178
            Top             =   825
            Width           =   2430
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   6600
            X2              =   7160
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "���䵥λ:"
            Height          =   180
            Index           =   27
            Left            =   7290
            TabIndex        =   177
            Top             =   825
            Width           =   810
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   6720
            X2              =   8535
            Y1              =   1365
            Y2              =   1365
         End
      End
      Begin VB.PictureBox picPane 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4520
         Index           =   2
         Left            =   990
         ScaleHeight     =   4515
         ScaleWidth      =   9870
         TabIndex        =   172
         Top             =   8000
         Width           =   9875
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   4
            Left            =   2535
            TabIndex        =   95
            Top             =   1352
            Width           =   2250
            _ExtentX        =   111495
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�˸�Ⱦ���²��������С�"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   3
            Left            =   1155
            TabIndex        =   94
            Top             =   1352
            Width           =   1350
            _ExtentX        =   109908
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��������ס�"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucPTB 
            Height          =   270
            Index           =   2
            Left            =   2760
            TabIndex        =   112
            Top             =   2115
            Width           =   1650
            _ExtentX        =   3334
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�޲�ԭѧ���)��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucPTB 
            Height          =   270
            Index           =   1
            Left            =   1400
            TabIndex        =   111
            Top             =   2115
            Width           =   1365
            _ExtentX        =   3254
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��ԭѧ���ԡ�"
         End
         Begin zlDisReportCardEx.uCheckNorm ucPTB 
            Height          =   270
            Index           =   0
            Left            =   60
            TabIndex        =   110
            Top             =   2115
            Width           =   1395
            _ExtentX        =   3307
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��ԭѧ���ԡ�"
         End
         Begin zlDisReportCardEx.uCheckNorm ucAIDS 
            Height          =   270
            Index           =   0
            Left            =   4065
            TabIndex        =   86
            Top             =   975
            Width           =   760
            _ExtentX        =   108876
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "HIV)��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   17
            Left            =   1020
            TabIndex        =   119
            Top             =   2490
            Width           =   1530
            _ExtentX        =   110225
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���������˷硢"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   16
            Left            =   60
            TabIndex        =   118
            Top             =   2490
            Width           =   810
            _ExtentX        =   108955
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�׺�"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   15
            Left            =   8835
            TabIndex        =   117
            Top             =   2115
            Width           =   1005
            _ExtentX        =   109299
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���տȡ�"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   14
            Left            =   6915
            TabIndex        =   116
            Top             =   2115
            Width           =   1905
            _ExtentX        =   110887
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�������Լ���Ĥ�ס�"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucTyphia 
            Height          =   270
            Index           =   1
            Left            =   5820
            TabIndex        =   115
            Top             =   2115
            Width           =   1080
            _ExtentX        =   109432
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���˺�)��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucTyphia 
            Height          =   270
            Index           =   0
            Left            =   4980
            TabIndex        =   114
            Top             =   2115
            Width           =   810
            _ExtentX        =   108955
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�˺���"
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   13
            Left            =   4400
            TabIndex        =   113
            Top             =   2115
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�˺�("
            CheckType       =   1
            BoxVisible      =   0   'False
            CheckedVisible  =   0   'False
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   11
            Left            =   4910
            TabIndex        =   105
            Top             =   1732
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "����("
            CheckType       =   1
            BoxVisible      =   0   'False
            CheckedVisible  =   0   'False
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   12
            Left            =   7685
            TabIndex        =   108
            Top             =   1732
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�ν��("
            CheckType       =   1
            BoxVisible      =   0   'False
            CheckedVisible  =   0   'False
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   9
            Left            =   8145
            TabIndex        =   99
            Top             =   1352
            Width           =   1830
            _ExtentX        =   108215
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�������������ס�"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   8
            Left            =   7140
            TabIndex        =   98
            Top             =   1352
            Width           =   1005
            _ExtentX        =   109299
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��Ȯ����"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   6
            Left            =   4800
            TabIndex        =   96
            Top             =   1352
            Width           =   810
            _ExtentX        =   108955
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   7
            Left            =   5595
            TabIndex        =   97
            Top             =   1352
            Width           =   1530
            _ExtentX        =   110225
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�����Գ�Ѫ�ȡ�"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   1
            Left            =   1955
            TabIndex        =   84
            Top             =   975
            Width           =   700
            _ExtentX        =   1244
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���̲�("
            CheckType       =   1
            BoxVisible      =   0   'False
            CheckedVisible  =   0   'False
         End
         Begin zlDisReportCardEx.uCheckNorm ucAIDS 
            Height          =   270
            Index           =   1
            Left            =   2700
            TabIndex        =   85
            Top             =   975
            Width           =   1350
            _ExtentX        =   109485
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���̲����ˡ�"
         End
         Begin zlDisReportCardEx.uCheckNorm ucHepatitis 
            Height          =   270
            Index           =   5
            Left            =   60
            TabIndex        =   93
            Top             =   1352
            Width           =   1080
            _ExtentX        =   109432
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "δ����)��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucHepatitis 
            Height          =   270
            Index           =   4
            Left            =   9125
            TabIndex        =   92
            Top             =   975
            Width           =   810
            _ExtentX        =   108955
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���͡�"
         End
         Begin zlDisReportCardEx.uCheckNorm ucHepatitis 
            Height          =   270
            Index           =   2
            Left            =   7505
            TabIndex        =   90
            Top             =   975
            Width           =   810
            _ExtentX        =   108955
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���͡�"
         End
         Begin zlDisReportCardEx.uCheckNorm ucHepatitis 
            Height          =   270
            Index           =   1
            Left            =   6695
            TabIndex        =   89
            Top             =   975
            Width           =   810
            _ExtentX        =   108955
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���͡�"
         End
         Begin zlDisReportCardEx.uCheckNorm ucHepatitis 
            Height          =   270
            Index           =   0
            Left            =   5885
            TabIndex        =   88
            Top             =   975
            Width           =   810
            _ExtentX        =   108955
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���͡�"
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   2
            Left            =   4825
            TabIndex        =   87
            Top             =   975
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�����Ը���("
            CheckType       =   1
            BoxVisible      =   0   'False
            CheckedVisible  =   0   'False
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   5
            Left            =   7380
            TabIndex        =   135
            Top             =   2865
            Width           =   1770
            _ExtentX        =   109802
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�˸�ȾH7N9������"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucAnthrax 
            Height          =   270
            Index           =   2
            Left            =   3815
            TabIndex        =   104
            Top             =   1732
            Width           =   1080
            _ExtentX        =   111548
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "δ����)��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucAnthrax 
            Height          =   270
            Index           =   0
            Left            =   1665
            TabIndex        =   102
            Top             =   1732
            Width           =   1005
            _ExtentX        =   111416
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��̿�ҡ�"
         End
         Begin zlDisReportCardEx.uCheckNorm ucAnthrax 
            Height          =   270
            Index           =   1
            Left            =   2655
            TabIndex        =   103
            Top             =   1732
            Width           =   1170
            _ExtentX        =   111707
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "Ƥ��̿�ҡ�"
         End
         Begin zlDisReportCardEx.uCheckNorm ucDysentery 
            Height          =   270
            Index           =   1
            Left            =   6420
            TabIndex        =   107
            Top             =   1732
            Width           =   1265
            _ExtentX        =   109749
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���װ���)��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucDysentery 
            Height          =   270
            Index           =   0
            Left            =   5415
            TabIndex        =   106
            Top             =   1732
            Width           =   1005
            _ExtentX        =   111416
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "ϸ���ԡ�"
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   21
            Left            =   1020
            TabIndex        =   129
            Top             =   2865
            Width           =   1530
            _ExtentX        =   110225
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���������岡��"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucSyphilis 
            Height          =   270
            Index           =   4
            Left            =   60
            TabIndex        =   128
            Top             =   2865
            Width           =   900
            _ExtentX        =   111231
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "����)��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucSyphilis 
            Height          =   270
            Index           =   3
            Left            =   8715
            TabIndex        =   127
            Top             =   2490
            Width           =   810
            _ExtentX        =   108955
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "̥����"
         End
         Begin zlDisReportCardEx.uCheckNorm ucSyphilis 
            Height          =   270
            Index           =   2
            Left            =   7905
            TabIndex        =   126
            Top             =   2490
            Width           =   810
            _ExtentX        =   108955
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���ڡ�"
         End
         Begin zlDisReportCardEx.uCheckNorm ucSyphilis 
            Height          =   270
            Index           =   1
            Left            =   7095
            TabIndex        =   125
            Top             =   2490
            Width           =   810
            _ExtentX        =   108955
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���ڡ�"
         End
         Begin zlDisReportCardEx.uCheckNorm ucSyphilis 
            Height          =   270
            Index           =   0
            Left            =   6285
            TabIndex        =   124
            Top             =   2490
            Width           =   810
            _ExtentX        =   108955
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���ڡ�"
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   25
            Left            =   4950
            TabIndex        =   122
            Top             =   2490
            Width           =   810
            _ExtentX        =   108955
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�ܲ���"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   19
            Left            =   3585
            TabIndex        =   121
            Top             =   2490
            Width           =   1350
            _ExtentX        =   109908
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��³�Ͼ�����"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   18
            Left            =   2565
            TabIndex        =   120
            Top             =   2490
            Width           =   1005
            _ExtentX        =   109299
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�ɺ��ȡ�"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucMalaria 
            Height          =   270
            Index           =   0
            Left            =   4275
            TabIndex        =   132
            Top             =   2865
            Width           =   1005
            _ExtentX        =   109299
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "����ű��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucMalaria 
            Height          =   270
            Index           =   1
            Left            =   5280
            TabIndex        =   133
            Top             =   2865
            Width           =   1005
            _ExtentX        =   109299
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "����ű��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucMalaria 
            Height          =   270
            Index           =   2
            Left            =   6285
            TabIndex        =   134
            Top             =   2865
            Width           =   1080
            _ExtentX        =   109432
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "δ����)��"
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousC 
            Height          =   270
            Index           =   6
            Left            =   60
            TabIndex        =   142
            Top             =   3870
            Width           =   1005
            _ExtentX        =   111416
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���Ȳ���"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousC 
            Height          =   270
            Index           =   5
            Left            =   6960
            TabIndex        =   141
            Top             =   3540
            Width           =   2520
            _ExtentX        =   15452
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�����Ժ͵ط��԰����˺���"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousC 
            Height          =   270
            Index           =   4
            Left            =   5820
            TabIndex        =   140
            Top             =   3540
            Width           =   1005
            _ExtentX        =   111416
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��粡��"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousC 
            Height          =   270
            Index           =   3
            Left            =   3885
            TabIndex        =   139
            Top             =   3540
            Width           =   1890
            _ExtentX        =   112977
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���Գ�Ѫ�Խ�Ĥ�ס�"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousC 
            Height          =   270
            Index           =   2
            Left            =   3075
            TabIndex        =   138
            Top             =   3540
            Width           =   810
            _ExtentX        =   111072
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousC 
            Height          =   270
            Index           =   1
            Left            =   1515
            TabIndex        =   137
            Top             =   3540
            Width           =   1530
            _ExtentX        =   75935
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�����������ס�"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousC 
            Height          =   270
            Index           =   8
            Left            =   2280
            TabIndex        =   144
            Top             =   3870
            Width           =   1005
            _ExtentX        =   111416
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "˿�没��"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousC 
            Height          =   270
            Index           =   9
            Left            =   3360
            TabIndex        =   145
            Top             =   3870
            Width           =   5850
            _ExtentX        =   206322
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�����ҡ�ϸ���ԺͰ��װ����������˺��͸��˺�����ĸ�Ⱦ�Ը�к����"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousA 
            Height          =   270
            Index           =   1
            Left            =   1035
            TabIndex        =   82
            Top             =   262
            Width           =   825
            _ExtentX        =   195395
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "����"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   0
            Left            =   60
            TabIndex        =   83
            Top             =   975
            Width           =   1890
            _ExtentX        =   30427
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "��Ⱦ�Էǵ��ͷ��ס�"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   10
            Left            =   1080
            TabIndex        =   101
            Top             =   1732
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "̿��("
            CheckType       =   1
            BoxVisible      =   0   'False
            CheckedVisible  =   0   'False
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   22
            Left            =   2580
            TabIndex        =   130
            Top             =   2865
            Width           =   1170
            _ExtentX        =   109590
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "Ѫ���没��"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   23
            Left            =   3765
            TabIndex        =   131
            Top             =   2865
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "ű��("
            CheckType       =   1
            BoxVisible      =   0   'False
            CheckedVisible  =   0   'False
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousC 
            Height          =   270
            Index           =   0
            Left            =   60
            TabIndex        =   136
            Top             =   3540
            Width           =   1395
            _ExtentX        =   5001
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�����Ը�ð��"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousC 
            Height          =   270
            Index           =   7
            Left            =   1200
            TabIndex        =   143
            Top             =   3870
            Width           =   1005
            _ExtentX        =   111416
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���没��"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousA 
            Height          =   270
            Index           =   0
            Left            =   60
            TabIndex        =   81
            Top             =   262
            Width           =   825
            _ExtentX        =   108982
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���ߡ�"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucHepatitis 
            Height          =   270
            Index           =   3
            Left            =   8315
            TabIndex        =   91
            Top             =   975
            Width           =   810
            _ExtentX        =   108955
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "���͡�"
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   24
            Left            =   60
            TabIndex        =   100
            Top             =   1732
            Width           =   1005
            _ExtentX        =   109299
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "�Ǹ��ȡ�"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousB 
            Height          =   270
            Index           =   20
            Left            =   5760
            TabIndex        =   123
            Top             =   2490
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "÷��("
            CheckType       =   1
            BoxVisible      =   0   'False
            CheckedVisible  =   0   'False
         End
         Begin zlDisReportCardEx.uCheckNorm ucInfectiousC 
            Height          =   270
            Index           =   10
            Left            =   60
            TabIndex        =   146
            Top             =   4200
            Width           =   1170
            _ExtentX        =   110014
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "����ڲ���"
            CheckType       =   1
         End
         Begin zlDisReportCardEx.uCheckNorm ucPTB 
            Height          =   270
            Index           =   3
            Left            =   8415
            TabIndex        =   109
            Top             =   1725
            Width           =   1410
            _ExtentX        =   6297
            _ExtentY        =   476
            BackColor       =   -2147483643
            Caption         =   "����ƽ��ҩ��"
         End
         Begin VB.Line LineNew 
            Index           =   1
            X1              =   0
            X2              =   9970
            Y1              =   592
            Y2              =   592
         End
         Begin VB.Line LineNew 
            Index           =   2
            X1              =   0
            X2              =   9982
            Y1              =   3222
            Y2              =   3222
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "���ഫȾ��*��"
            Height          =   180
            Index           =   29
            Left            =   105
            TabIndex        =   175
            Top             =   3307
            Width           =   1170
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "���ഫȾ��*��"
            Height          =   180
            Index           =   28
            Left            =   105
            TabIndex        =   174
            Top             =   676
            Width           =   1170
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "���ഫȾ��*��"
            Height          =   180
            Index           =   30
            Left            =   105
            TabIndex        =   173
            Top             =   0
            Width           =   1170
         End
      End
      Begin VB.PictureBox picPane 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   3
         Left            =   990
         ScaleHeight     =   2415
         ScaleWidth      =   9870
         TabIndex        =   160
         Top             =   12675
         Width           =   9875
         Begin VB.TextBox txtEnter 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   1
            Left            =   7620
            TabIndex        =   154
            ToolTipText     =   "�ʱ�������ʱ�ɳ����Զ�����"
            Top             =   1620
            Width           =   450
         End
         Begin VB.TextBox txtEnter 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   0
            Left            =   6330
            TabIndex        =   153
            ToolTipText     =   "�ʱ�������ʱ�ɳ����Զ�����"
            Top             =   1620
            Width           =   1095
         End
         Begin VB.TextBox txtEnter 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   2
            Left            =   8250
            TabIndex        =   155
            ToolTipText     =   "�ʱ�������ʱ�ɳ����Զ�����"
            Top             =   1620
            Width           =   525
         End
         Begin VB.TextBox txtRemarks 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   156
            Top             =   2130
            Width           =   9060
         End
         Begin VB.TextBox txtDoctor 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   1170
            TabIndex        =   152
            ToolTipText     =   "�ҽ�������ʱ�ɳ����Զ�����"
            Top             =   1620
            Width           =   3255
         End
         Begin VB.TextBox txtDocNumber 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   6330
            TabIndex        =   151
            Top             =   1260
            Width           =   2520
         End
         Begin VB.TextBox txtUnit 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   1170
            TabIndex        =   150
            Top             =   1260
            Width           =   3255
         End
         Begin VB.TextBox txtReason 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   6330
            TabIndex        =   149
            Top             =   915
            Width           =   2520
         End
         Begin VB.TextBox txtIName 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   1170
            TabIndex        =   148
            Top             =   915
            Width           =   3255
         End
         Begin VB.TextBox txtImportant 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   105
            MultiLine       =   -1  'True
            TabIndex        =   147
            Top             =   240
            Width           =   9500
         End
         Begin VB.Line Line5 
            X1              =   8250
            X2              =   8685
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line4 
            X1              =   7620
            X2              =   8055
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line1 
            Index           =   24
            X1              =   6360
            X2              =   7425
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label lblAttack 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   3
            Left            =   7455
            TabIndex        =   171
            Top             =   1620
            Width           =   180
         End
         Begin VB.Label lblAttack 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   4
            Left            =   8085
            TabIndex        =   170
            Top             =   1620
            Width           =   180
         End
         Begin VB.Label lblAttack 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Index           =   5
            Left            =   8760
            TabIndex        =   169
            Top             =   1620
            Width           =   180
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��ע��"
            Height          =   180
            Index           =   31
            Left            =   225
            TabIndex        =   168
            Top             =   2130
            Width           =   540
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "�����*��"
            Height          =   180
            Index           =   32
            Left            =   5265
            TabIndex        =   167
            ToolTipText     =   "�ʱ�������ʱ�ɳ����Զ�����"
            Top             =   1620
            Width           =   990
         End
         Begin VB.Line Line1 
            Index           =   25
            X1              =   1155
            X2              =   4470
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "�ҽ��*��"
            Height          =   180
            Index           =   33
            Left            =   105
            TabIndex        =   166
            ToolTipText     =   "�ҽ�������ʱ�ɳ����Զ�����"
            Top             =   1620
            Width           =   990
         End
         Begin VB.Line Line1 
            Index           =   26
            X1              =   6360
            X2              =   8905
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��ϵ�绰��"
            Height          =   180
            Index           =   34
            Left            =   5265
            TabIndex        =   165
            Top             =   1260
            Width           =   900
         End
         Begin VB.Line Line1 
            Index           =   27
            X1              =   1140
            X2              =   4455
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "���浥λ��"
            Height          =   180
            Index           =   35
            Left            =   105
            TabIndex        =   164
            Top             =   1260
            Width           =   900
         End
         Begin VB.Line Line1 
            Index           =   28
            X1              =   6360
            X2              =   8875
            Y1              =   1095
            Y2              =   1095
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "�˿�ԭ��"
            Height          =   180
            Index           =   36
            Left            =   5265
            TabIndex        =   163
            Top             =   915
            Width           =   900
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����������"
            Height          =   180
            Index           =   37
            Left            =   105
            TabIndex        =   162
            Top             =   915
            Width           =   900
         End
         Begin VB.Line Line1 
            Index           =   29
            X1              =   1140
            X2              =   4455
            Y1              =   1095
            Y2              =   1095
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "�������������Լ��ص��⴫Ⱦ����"
            Height          =   180
            Index           =   38
            Left            =   105
            TabIndex        =   161
            Top             =   0
            Width           =   2880
         End
         Begin VB.Line LineNew 
            Index           =   4
            X1              =   0
            X2              =   9875
            Y1              =   1965
            Y2              =   1965
         End
         Begin VB.Line LineNew 
            Index           =   5
            X1              =   0
            X2              =   9875
            Y1              =   645
            Y2              =   645
         End
      End
      Begin zlDisReportCardEx.uCheckNorm ucTmp 
         Height          =   270
         Index           =   0
         Left            =   11040
         TabIndex        =   221
         Top             =   3960
         Visible         =   0   'False
         Width           =   675
         _ExtentX        =   17701
         _ExtentY        =   476
         BackColor       =   -2147483643
         Caption         =   "��չ"
         CheckType       =   1
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   11040
         Picture         =   "frmReport.frx":0000
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape Shape1 
         Height          =   12900
         Left            =   960
         Top             =   2235
         Width           =   9930
      End
      Begin VB.Line LineNew 
         Index           =   0
         X1              =   990
         X2              =   10870
         Y1              =   7875
         Y2              =   7875
      End
      Begin VB.Line LineNew 
         Index           =   3
         X1              =   990
         X2              =   10870
         Y1              =   12555
         Y2              =   12555
      End
   End
   Begin VB.PictureBox picShadow 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1770
      Left            =   750
      ScaleHeight     =   1770
      ScaleWidth      =   1140
      TabIndex        =   159
      Top             =   660
      Width           =   1140
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���幦�ܣ���Ⱦ�����濨��ʽ�޸�
'ע��������֧����Ӵ���ԭ�е�����ؼ��ļ����뱣�棬����Ҫʹ�������ؼ�����ͨ��LoadData,CheckData,SaveData����������
'��ӡ�ؼ�˵����Ŀǰ��֧��Textbox,Line,uCheckNorm�����ֿؼ��Ĵ�ӡ,��Ҫ��ӡ�Ŀؼ���Ҫ����picPane�ؼ���
'������һ����չ��ѡ��,����֧�ֱ�����չ��Ϣ,Ĭ��Ϊ���ز�ʹ��

Public mclsReport As Object '��Ⱦ�����濨��������

Private Sub Form_Load()
    Call mclsReport.FormLoad
End Sub

Private Sub Form_Resize()
    Call mclsReport.FormResize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call mclsReport.FormUnload
    Set mclsReport = Nothing
End Sub

Private Sub hsbReport_Change()
    Call mclsReport.hsbReportChange
End Sub

Private Sub lblAttack_Click(Index As Integer)
    Call mclsReport.lblAttackClick(Index)
End Sub

Private Sub lblDeath_Click(Index As Integer)
    Call mclsReport.lblDeathClick(Index)
End Sub

Private Sub lblBirth_Click(Index As Integer)
    Call mclsReport.lblBirthClick(Index)
End Sub

Private Sub lblDiagnose_Click(Index As Integer)
    Call mclsReport.lblDiagnoseClick(Index)
End Sub

Private Sub lblAttack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call mclsReport.lblAttackMouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub lblDeath_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call mclsReport.lblDeathMouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub lblBirth_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call mclsReport.lblBirthMouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub lblDiagnose_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call mclsReport.lblDiagnoseMouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub MView_DateClick(ByVal DateClicked As Date)
    Call mclsReport.MViewDateClick(DateClicked)
End Sub

Private Sub MView_LostFocus()
    Call mclsReport.MViewLostFocus
End Sub

Private Sub picReport_GotFocus()
    Call mclsReport.picReportGotFocus
End Sub

Private Sub Subclass_WndProc(Msg As Long, wParam As Long, lParam As Long, result As Long)
    Call mclsReport.SubclassWndProc(Msg, wParam, lParam, result)
End Sub

Private Sub txtAge_Change()
    Call mclsReport.txtAgeChange
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    Call mclsReport.txtAgeKeyPress(KeyAscii)
End Sub

Private Sub vsbReport_Change()
    Call mclsReport.vsbReportChange
End Sub

Private Sub txtAttack_KeyPress(Index As Integer, KeyAscii As Integer)
    Call mclsReport.txtAttackKeyPress(Index, KeyAscii)
End Sub

Private Sub txtBirth_KeyPress(Index As Integer, KeyAscii As Integer)
    Call mclsReport.txtBirthKeyPress(Index, KeyAscii)
End Sub

Private Sub txtDeath_KeyPress(Index As Integer, KeyAscii As Integer)
    Call mclsReport.txtDeathKeyPress(Index, KeyAscii)
End Sub

Private Sub txtDiagnose_KeyPress(Index As Integer, KeyAscii As Integer)
    Call mclsReport.txtDiagnoseKeyPress(Index, KeyAscii)
End Sub

Private Sub txtIDCard_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsReport.txtIDCardKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub txtIDCard_KeyPress(Index As Integer, KeyAscii As Integer)
    Call mclsReport.txtIDCardKeyPress(Index, KeyAscii)
End Sub

Private Sub ucCaseType1_Change(Index As Integer)
    Call mclsReport.ucCaseType1Change(Index)
End Sub

Private Sub ucCaseType2_Change(Index As Integer)
    Call mclsReport.ucCaseType2Change(Index)
End Sub

Public Function LoadData(lngFileID As Long, lngPatiID As Long, lngPageID As Long, bytType As Byte, bytFrom As Byte, lngDeptID As Long, lngtBabyNo As Long) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'����: ��Ⱦ�����濨�Զ����������
'���: lngPatiID = ����id
'      lngPageID = ��ҳID(���ﴫ�Һ�ID)
'      bytType=�༭��ʽ0-������1-�޸ģ�����������ȡ����
'      bytFrom=������Դ1-���� 2-סԺ
'      lngDeptID = ��ǰ����ID
'      lngFileID=�ļ�ID,��Դ�ڵ��Ӳ�����¼.ID
'      bytBabyNo = Ӥ��ID
'����:��������ֵ = ���سɹ�����True,ʧ�ܷ���False
'����:��͢��
'����:2017-08-15 09:30:21
'---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    LoadData = True
'    txtName.Text = "��͢��"
    Exit Function
errHandle:
    LoadData = False
End Function

Public Function CheckData(lngFileID As Long, lngPatiID As Long, lngPageID As Long, bytType As Byte, bytFrom As Byte, lngDeptID As Long, lngtBabyNo As Long, ByRef strTmp As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'����: ��Ⱦ�����濨�Զ���������
'���: lngPatiID = ����id
'      lngPageID = ��ҳID(���ﴫ�Һ�ID)
'      bytType=�༭��ʽ0-������1-�޸ģ�����������ȡ����
'      bytFrom=������Դ1-���� 2-סԺ
'      lngDeptID = ��ǰ����ID
'      lngFileID=�ļ�ID,��Դ�ڵ��Ӳ�����¼.ID
'      bytBabyNo = Ӥ��ID
'����:��������ֵ = ���ͨ������True,��ͨ������False
'      strTmp = ���ؼ������ʾ��Ϣ����ʽ��AAAAA$BBBBB$CCCCC$
'����:��͢��
'����:2017-08-15 09:30:21
'---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    CheckData = True
'    If txtName.Text = "" Then
'           strTmp = strTmp & "���ֲ���Ϊ��$"
'
'    End If
'    strTmp = "���Բ���$12345$"
    Exit Function
errHandle:
    CheckData = False
End Function
 

Public Function SaveData(lngFileID As Long, lngPatiID As Long, lngPageID As Long, bytType As Byte, bytFrom As Byte, lngDeptID As Long, lngtBabyNo As Long) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'����: ��Ⱦ�����濨�Զ��屣������
'���: lngPatiID = ����id
'      lngPageID = ��ҳID(���ﴫ�Һ�ID)
'      bytType=�༭��ʽ0-������1-�޸ģ�����������ȡ����
'      bytFrom=������Դ1-���� 2-סԺ
'      lngDeptID = ��ǰ����ID
'      lngFileID=�ļ�ID,��Դ�ڵ��Ӳ�����¼.ID
'      bytBabyNo = Ӥ��ID
'����:��������ֵ = ����ɹ�����True,ʧ�ܷ���False
'ע���ѳ�ʼ���������Ӷ���gcnOracle��SaveData����������д���ʱ�Ĵ���
'����:��͢��
'����:2017-08-15 09:30:21
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo errHandle

'    strSql = "Zl_��Ⱦ�����濨����_Update('6077428728059','1135877','1','0','3','4','��͢��','1','����','1','0','')"
'    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    SaveData = True
    Exit Function
errHandle:
    SaveData = False
End Function

