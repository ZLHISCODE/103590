VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl PaneTwo 
   Appearance      =   0  'Flat
   BackColor       =   &H0080C0FF&
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9825
   ScaleHeight     =   5595
   ScaleWidth      =   9825
   Begin MSComCtl2.MonthView MView 
      Height          =   2220
      Left            =   9120
      TabIndex        =   117
      Top             =   4200
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483643
      Appearance      =   1
      StartOfWeek     =   185270274
      CurrentDate     =   41981
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1600
      TabIndex        =   29
      Tag             =   "137,236"
      Top             =   1236
      Width           =   4185
   End
   Begin zlDisReportCard.uCheckNorm ucCaseType2 
      Height          =   270
      Index           =   1
      Left            =   2265
      TabIndex        =   68
      Tag             =   "215,394"
      Top             =   4200
      Width           =   690
      _ExtentX        =   23654
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "慢性"
   End
   Begin VB.TextBox txtNumber 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   6705
      TabIndex        =   30
      Tag             =   "455,236"
      Top             =   1236
      Width           =   1815
   End
   Begin zlDisReportCard.uCheckNorm ucCaseType1 
      Height          =   270
      Index           =   3
      Left            =   5355
      TabIndex        =   66
      Tag             =   "452,371"
      Top             =   3840
      Width           =   1350
      _ExtentX        =   24818
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "病原携带者"
   End
   Begin zlDisReportCard.uCheckNorm ucCaseType1 
      Height          =   270
      Index           =   2
      Left            =   4185
      TabIndex        =   65
      Tag             =   "339,371"
      Top             =   3840
      Width           =   1170
      _ExtentX        =   24500
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "确诊病例、"
   End
   Begin zlDisReportCard.uCheckNorm ucCaseType1 
      Height          =   270
      Index           =   1
      Left            =   2670
      TabIndex        =   64
      Tag             =   "238,371"
      Top             =   3840
      Width           =   1530
      _ExtentX        =   25135
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "临床诊断病例、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   17
      Left            =   3968
      TabIndex        =   62
      Tag             =   "677,350"
      Top             =   3320
      Width           =   810
      _ExtentX        =   21325
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "不详"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   16
      Left            =   2618
      TabIndex        =   61
      Tag             =   "594,350"
      Top             =   3320
      Width           =   1350
      _ExtentX        =   22278
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "其他（ ）、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   15
      Left            =   1268
      TabIndex        =   60
      Tag             =   "506,350"
      Top             =   3320
      Width           =   1350
      _ExtentX        =   22278
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "家务及待业、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   14
      Left            =   105
      TabIndex        =   59
      Tag             =   "431,350"
      Top             =   3320
      Width           =   1163
      _ExtentX        =   24500
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "离退人员、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   13
      Left            =   6840
      TabIndex        =   57
      Tag             =   "356,350"
      Top             =   2950
      Width           =   1170
      _ExtentX        =   21960
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "干部职员、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   12
      Left            =   5670
      TabIndex        =   56
      Tag             =   "281,350"
      Top             =   2950
      Width           =   1170
      _ExtentX        =   21960
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "渔(船)民、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   11
      Left            =   4860
      TabIndex        =   55
      Tag             =   "230,350"
      Top             =   2950
      Width           =   810
      _ExtentX        =   21325
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "牧民、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   10
      Left            =   4050
      TabIndex        =   54
      Tag             =   "179,350"
      Top             =   2950
      Width           =   810
      _ExtentX        =   21325
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "农民、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   9
      Left            =   3240
      TabIndex        =   53
      Tag             =   "128,350"
      Top             =   2950
      Width           =   810
      _ExtentX        =   21325
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "民工、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   7
      Left            =   1268
      TabIndex        =   51
      Tag             =   "654,326"
      Top             =   2950
      Width           =   1163
      _ExtentX        =   21960
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "医务人员、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   6
      Left            =   105
      TabIndex        =   50
      Tag             =   "579,326"
      Top             =   2950
      Width           =   1163
      _ExtentX        =   24500
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "商业服务、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   5
      Left            =   6470
      TabIndex        =   48
      Tag             =   "492,326"
      Top             =   2580
      Width           =   1350
      _ExtentX        =   22278
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "餐饮食品业、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   4
      Left            =   4945
      TabIndex        =   47
      Tag             =   "391,326"
      Top             =   2580
      Width           =   1525
      _ExtentX        =   22595
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "保育员及保姆、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   3
      Left            =   4140
      TabIndex        =   46
      Tag             =   "340,326"
      Top             =   2580
      Width           =   805
      _ExtentX        =   21325
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "教师、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   2
      Left            =   2430
      TabIndex        =   45
      Tag             =   "227,326"
      Top             =   2580
      Width           =   1710
      _ExtentX        =   22913
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "学生(大中小学)、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   1
      Left            =   1268
      TabIndex        =   44
      Tag             =   "152,326"
      Top             =   2580
      Width           =   1163
      _ExtentX        =   21960
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "散居儿童、"
   End
   Begin VB.TextBox txtDiagnose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   3
      Left            =   3900
      TabIndex        =   76
      Tag             =   "314,445"
      Top             =   4950
      Width           =   525
   End
   Begin VB.TextBox txtDeath 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   77
      Tag             =   "554,445"
      Top             =   5310
      Width           =   1095
   End
   Begin VB.TextBox txtDeath 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   2370
      TabIndex        =   78
      Tag             =   "620,445"
      Top             =   5310
      Width           =   615
   End
   Begin VB.TextBox txtDeath 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   2
      Left            =   3165
      TabIndex        =   79
      Tag             =   "660,445"
      Top             =   5310
      Width           =   525
   End
   Begin VB.TextBox txtDiagnose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   73
      Tag             =   "170,445"
      Top             =   4950
      Width           =   1095
   End
   Begin VB.TextBox txtDiagnose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   2370
      TabIndex        =   74
      Tag             =   "235,445"
      Top             =   4950
      Width           =   615
   End
   Begin VB.TextBox txtDiagnose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   2
      Left            =   3165
      TabIndex        =   75
      Tag             =   "275,445"
      Top             =   4950
      Width           =   525
   End
   Begin VB.TextBox txtAttack 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   70
      Tag             =   "170,422"
      Top             =   4590
      Width           =   1095
   End
   Begin VB.TextBox txtAttack 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   2370
      TabIndex        =   71
      Tag             =   "235,422"
      Top             =   4590
      Width           =   615
   End
   Begin VB.TextBox txtAttack 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   2
      Left            =   3165
      TabIndex        =   72
      Tag             =   "275,422"
      Top             =   4590
      Width           =   525
   End
   Begin VB.TextBox txtAddInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   0
      Left            =   1530
      TabIndex        =   37
      Tag             =   "179,281"
      Top             =   1940
      Width           =   690
   End
   Begin VB.TextBox txtAddInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   2400
      TabIndex        =   38
      Tag             =   "239,281"
      Top             =   1940
      Width           =   885
   End
   Begin VB.TextBox txtAddInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   3
      Left            =   5025
      TabIndex        =   40
      Tag             =   "407,281"
      Top             =   1940
      Width           =   885
   End
   Begin VB.TextBox txtAddInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   4
      Left            =   7035
      TabIndex        =   41
      Tag             =   "551,281"
      Top             =   1940
      Width           =   915
   End
   Begin VB.TextBox txtAddInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   5
      Left            =   8145
      TabIndex        =   42
      Tag             =   "617,281"
      Top             =   1940
      Width           =   690
   End
   Begin zlDisReportCard.uCheckNorm ucAge 
      Height          =   270
      Index           =   0
      Left            =   8115
      TabIndex        =   26
      Tag             =   "599,210"
      Top             =   839
      Width           =   465
      _ExtentX        =   20294
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "岁"
   End
   Begin VB.TextBox txtParentName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   3645
      TabIndex        =   1
      Tag             =   "317,164"
      Top             =   180
      Width           =   1455
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Tag             =   "137,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   1
      Left            =   1620
      TabIndex        =   3
      Tag             =   "157,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   2
      Left            =   1920
      TabIndex        =   4
      Tag             =   "177,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   3
      Left            =   2220
      TabIndex        =   5
      Tag             =   "197,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   4
      Left            =   2520
      TabIndex        =   6
      Tag             =   "217,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   5
      Left            =   2820
      TabIndex        =   7
      Tag             =   "237,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   6
      Left            =   3120
      TabIndex        =   8
      Tag             =   "257,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   7
      Left            =   3420
      TabIndex        =   9
      Tag             =   "277,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   8
      Left            =   3735
      TabIndex        =   10
      Tag             =   "297,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   9
      Left            =   4035
      TabIndex        =   11
      Tag             =   "317,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   10
      Left            =   4335
      TabIndex        =   12
      Tag             =   "337,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   11
      Left            =   4635
      TabIndex        =   13
      Tag             =   "357,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   12
      Left            =   4935
      TabIndex        =   14
      Tag             =   "377,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   13
      Left            =   5235
      TabIndex        =   15
      Tag             =   "397,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   14
      Left            =   5535
      TabIndex        =   16
      Tag             =   "417,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   15
      Left            =   5835
      TabIndex        =   17
      Tag             =   "437,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   16
      Left            =   6135
      TabIndex        =   18
      Tag             =   "457,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtIDCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   17
      Left            =   6450
      TabIndex        =   19
      Tag             =   "477,183"
      Top             =   472
      Width           =   240
   End
   Begin VB.TextBox txtBirth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   22
      Tag             =   "170,212"
      Top             =   884
      Width           =   1095
   End
   Begin VB.TextBox txtBirth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   2355
      TabIndex        =   23
      Tag             =   "235,212"
      Top             =   884
      Width           =   615
   End
   Begin VB.TextBox txtBirth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   2
      Left            =   3150
      TabIndex        =   24
      Tag             =   "275,212"
      Top             =   884
      Width           =   525
   End
   Begin VB.TextBox txtAge 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   6585
      TabIndex        =   25
      Tag             =   "479,212"
      Top             =   884
      Width           =   525
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   735
      TabIndex        =   0
      Tag             =   "119,164"
      Top             =   180
      Width           =   1455
   End
   Begin zlDisReportCard.uCheckNorm ucSex 
      Height          =   270
      Index           =   0
      Left            =   7575
      TabIndex        =   20
      Tag             =   "539,183"
      Top             =   465
      Width           =   570
      _ExtentX        =   227489
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "男"
   End
   Begin zlDisReportCard.uCheckNorm ucSex 
      Height          =   270
      Index           =   1
      Left            =   8160
      TabIndex        =   21
      Tag             =   "583,183"
      Top             =   465
      Width           =   570
      _ExtentX        =   227489
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "女"
   End
   Begin zlDisReportCard.uCheckNorm ucAge 
      Height          =   270
      Index           =   1
      Left            =   8573
      TabIndex        =   27
      Tag             =   "625,210"
      Top             =   839
      Width           =   465
      _ExtentX        =   20294
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "月"
   End
   Begin zlDisReportCard.uCheckNorm ucAge 
      Height          =   270
      Index           =   2
      Left            =   9030
      TabIndex        =   28
      Tag             =   "651,210"
      Top             =   839
      Width           =   555
      _ExtentX        =   20452
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "天)"
   End
   Begin zlDisReportCard.uCheckNorm ucFrom 
      Height          =   270
      Index           =   0
      Left            =   1080
      TabIndex        =   31
      Tag             =   "149,256"
      Top             =   1545
      Width           =   885
      _ExtentX        =   228256
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "本县区"
   End
   Begin zlDisReportCard.uCheckNorm ucFrom 
      Height          =   270
      Index           =   1
      Left            =   2070
      TabIndex        =   32
      Tag             =   "217,256"
      Top             =   1543
      Width           =   1425
      _ExtentX        =   228997
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "本市其他县区"
   End
   Begin zlDisReportCard.uCheckNorm ucFrom 
      Height          =   270
      Index           =   2
      Left            =   3630
      TabIndex        =   33
      Tag             =   "321,256"
      Top             =   1545
      Width           =   1380
      _ExtentX        =   229129
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "本省其它地市"
   End
   Begin zlDisReportCard.uCheckNorm ucFrom 
      Height          =   270
      Index           =   3
      Left            =   5190
      TabIndex        =   34
      Tag             =   "431,256"
      Top             =   1545
      Width           =   675
      _ExtentX        =   227886
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "外省"
   End
   Begin zlDisReportCard.uCheckNorm ucFrom 
      Height          =   270
      Index           =   4
      Left            =   6150
      TabIndex        =   35
      Tag             =   "493,256"
      Top             =   1545
      Width           =   885
      _ExtentX        =   228256
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "港澳台"
   End
   Begin zlDisReportCard.uCheckNorm ucFrom 
      Height          =   270
      Index           =   5
      Left            =   7230
      TabIndex        =   36
      Tag             =   "561,256"
      Top             =   1545
      Width           =   675
      _ExtentX        =   227886
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "外籍"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   43
      Tag             =   "77,326"
      Top             =   2580
      Width           =   1163
      _ExtentX        =   24500
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "幼托儿童、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   8
      Left            =   2430
      TabIndex        =   52
      Tag             =   "77,350"
      Top             =   2950
      Width           =   810
      _ExtentX        =   21325
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "工人、"
   End
   Begin zlDisReportCard.uCheckNorm ucCaseType1 
      Height          =   270
      Index           =   0
      Left            =   1500
      TabIndex        =   63
      Tag             =   "163,371"
      Top             =   3840
      Width           =   1170
      _ExtentX        =   24500
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "疑似病例、"
   End
   Begin zlDisReportCard.uCheckNorm ucCaseType2 
      Height          =   270
      Index           =   0
      Left            =   1500
      TabIndex        =   67
      Tag             =   "163,394"
      Top             =   4200
      Width           =   810
      _ExtentX        =   23865
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "急性、"
   End
   Begin VB.TextBox txtAddInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   2
      Left            =   3510
      TabIndex        =   39
      Tag             =   "311,281"
      Top             =   1940
      Width           =   855
   End
   Begin zlDisReportCard.uCheckNorm ucCaseType2 
      Height          =   270
      Index           =   2
      Left            =   5640
      TabIndex        =   69
      Tag             =   "215,394"
      Top             =   4200
      Width           =   855
      _ExtentX        =   23945
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "未分型"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   18
      Left            =   7820
      TabIndex        =   49
      Tag             =   "492,326"
      Top             =   2580
      Width           =   1710
      _ExtentX        =   22913
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "公共场所服务员、"
   End
   Begin zlDisReportCard.uCheckNorm ucCheckJob 
      Height          =   270
      Index           =   19
      Left            =   8010
      TabIndex        =   58
      Tag             =   "356,350"
      Top             =   2950
      Width           =   1770
      _ExtentX        =   23019
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "海员及长途驾驶员、"
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(乙型肝炎*、血吸虫病*、丙肝)、"
      Height          =   180
      Index           =   1
      Left            =   2940
      TabIndex        =   119
      Tag             =   "255,397"
      Top             =   4245
      Width           =   2700
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8895
      Picture         =   "PaneTwo.ctx":0000
      Top             =   1380
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(1)"
      Height          =   180
      Index           =   0
      Left            =   1110
      TabIndex        =   118
      Tag             =   "143,375"
      Top             =   3885
      Width           =   270
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "县(区)"
      Height          =   180
      Index           =   16
      Left            =   4380
      TabIndex        =   96
      Tag             =   "360,281"
      Top             =   1935
      Width           =   540
   End
   Begin VB.Line Line1 
      Index           =   23
      Tag             =   "653,455,684"
      X1              =   3165
      X2              =   3710
      Y1              =   5490
      Y2              =   5490
   End
   Begin VB.Line Line1 
      Index           =   22
      Tag             =   "611,455,645"
      X1              =   2370
      X2              =   2985
      Y1              =   5490
      Y2              =   5490
   End
   Begin VB.Line Line1 
      Index           =   21
      Tag             =   "311,455,330"
      X1              =   3900
      X2              =   4515
      Y1              =   5130
      Y2              =   5130
   End
   Begin VB.Line Line1 
      Index           =   20
      Tag             =   "269,455,300"
      X1              =   3165
      X2              =   3700
      Y1              =   5130
      Y2              =   5130
   End
   Begin VB.Line Line1 
      Index           =   19
      Tag             =   "227,455,254"
      X1              =   2370
      X2              =   2985
      Y1              =   5130
      Y2              =   5130
   End
   Begin VB.Line Line1 
      Index           =   18
      Tag             =   "269,432,300"
      X1              =   3165
      X2              =   3700
      Y1              =   4770
      Y2              =   4770
   End
   Begin VB.Line Line1 
      Index           =   17
      Tag             =   "227,432,254"
      X1              =   2370
      X2              =   2985
      Y1              =   4770
      Y2              =   4770
   End
   Begin VB.Line Line1 
      Index           =   16
      Tag             =   "527,455,600"
      X1              =   1080
      X2              =   2255
      Y1              =   5490
      Y2              =   5490
   End
   Begin VB.Line Line1 
      Index           =   15
      Tag             =   "143,455,216"
      X1              =   1080
      X2              =   2175
      Y1              =   5130
      Y2              =   5130
   End
   Begin VB.Line Line1 
      Index           =   0
      Tag             =   "143,432,216"
      X1              =   1080
      X2              =   2175
      Y1              =   4770
      Y2              =   4770
   End
   Begin VB.Line Line1 
      Index           =   7
      Tag             =   "137,245,394"
      X1              =   1600
      X2              =   5750
      Y1              =   1425
      Y2              =   1425
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "联系电话："
      Height          =   180
      Index           =   11
      Left            =   5820
      TabIndex        =   116
      Tag             =   "396,236"
      Top             =   1230
      Width           =   900
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "工作单位(学校)："
      Height          =   180
      Index           =   10
      Left            =   105
      TabIndex        =   115
      Tag             =   "78,236"
      Top             =   1230
      Width           =   1440
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "病例分类*："
      Height          =   180
      Index           =   21
      Left            =   105
      TabIndex        =   114
      Tag             =   "78,375"
      Top             =   3885
      Width           =   990
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(2)"
      Height          =   180
      Index           =   22
      Left            =   1110
      TabIndex        =   113
      Tag             =   "143,398"
      Top             =   4245
      Width           =   270
   End
   Begin VB.Line Line1 
      Index           =   11
      Tag             =   "311,292,351"
      X1              =   3465
      X2              =   4365
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Label lblDiagnose 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "时"
      Height          =   180
      Index           =   3
      Left            =   4470
      TabIndex        =   112
      Tag             =   "331,445 "
      Top             =   4950
      Width           =   180
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "死亡日期 ："
      Height          =   180
      Index           =   24
      Left            =   105
      TabIndex        =   111
      Tag             =   "462,445"
      Top             =   5310
      Width           =   990
   End
   Begin VB.Label lblDeath 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "年"
      Height          =   180
      Index           =   0
      Left            =   2175
      TabIndex        =   110
      Tag             =   "600,445"
      Top             =   5310
      Width           =   180
   End
   Begin VB.Label lblDeath 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "月"
      Height          =   180
      Index           =   1
      Left            =   2985
      TabIndex        =   109
      Tag             =   "642,445"
      Top             =   5310
      Width           =   180
   End
   Begin VB.Label lblDeath 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "日"
      Height          =   180
      Index           =   2
      Left            =   3690
      TabIndex        =   108
      Tag             =   "686,445"
      Top             =   5310
      Width           =   180
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "诊断时间*："
      Height          =   180
      Index           =   23
      Left            =   105
      TabIndex        =   107
      Tag             =   "78,445"
      Top             =   4950
      Width           =   990
   End
   Begin VB.Label lblDiagnose 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "年"
      Height          =   180
      Index           =   0
      Left            =   2175
      TabIndex        =   106
      Tag             =   "216,445"
      Top             =   4950
      Width           =   180
   End
   Begin VB.Label lblDiagnose 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "月"
      Height          =   180
      Index           =   1
      Left            =   2985
      TabIndex        =   105
      Tag             =   "258,445"
      Top             =   4950
      Width           =   180
   End
   Begin VB.Label lblDiagnose 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "日"
      Height          =   180
      Index           =   2
      Left            =   3690
      TabIndex        =   104
      Tag             =   "302,445"
      Top             =   4950
      Width           =   180
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "发病日期*："
      Height          =   180
      Index           =   25
      Left            =   105
      TabIndex        =   103
      Tag             =   "78,422"
      Top             =   4590
      Width           =   990
   End
   Begin VB.Label lblAttack 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "年"
      Height          =   180
      Index           =   0
      Left            =   2175
      TabIndex        =   102
      Tag             =   "216,422"
      Top             =   4590
      Width           =   180
   End
   Begin VB.Label lblAttack 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "月"
      Height          =   180
      Index           =   1
      Left            =   2985
      TabIndex        =   101
      Tag             =   "258,422 "
      Top             =   4590
      Width           =   180
   End
   Begin VB.Label lblAttack 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "日"
      Height          =   180
      Index           =   2
      Left            =   3690
      TabIndex        =   100
      Tag             =   "302,422"
      Top             =   4590
      Width           =   180
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "现住址(详填)*："
      Height          =   180
      Index           =   13
      Left            =   105
      TabIndex        =   99
      Tag             =   "78,281"
      Top             =   1935
      Width           =   1350
   End
   Begin VB.Line Line1 
      Index           =   9
      Tag             =   "179,292,226"
      X1              =   1530
      X2              =   2200
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Line Line1 
      Index           =   10
      Tag             =   "239,292,300"
      X1              =   2400
      X2              =   3300
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "省"
      Height          =   180
      Index           =   14
      Left            =   2220
      TabIndex        =   98
      Tag             =   "229,281"
      Top             =   1935
      Width           =   180
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "市"
      Height          =   180
      Index           =   15
      Left            =   3285
      TabIndex        =   97
      Tag             =   "301,281"
      Top             =   1940
      Width           =   180
   End
   Begin VB.Line Line1 
      Index           =   12
      Tag             =   "407,292,470"
      X1              =   5025
      X2              =   5925
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Line Line1 
      Index           =   13
      Tag             =   "551,292,606"
      X1              =   7020
      X2              =   7920
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Line Line1 
      Index           =   14
      Tag             =   "617,292,670"
      X1              =   8145
      X2              =   8850
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "乡(镇、街道)"
      Height          =   180
      Index           =   17
      Left            =   5925
      TabIndex        =   95
      Tag             =   "470,281"
      Top             =   1935
      Width           =   1080
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "村"
      Height          =   180
      Index           =   18
      Left            =   7950
      TabIndex        =   94
      Tag             =   "606,281"
      Top             =   1940
      Width           =   180
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(门牌号)"
      Height          =   180
      Index           =   19
      Left            =   8835
      TabIndex        =   93
      Tag             =   "672,281"
      Top             =   1935
      Width           =   720
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "人群分类*："
      Height          =   180
      Index           =   20
      Left            =   105
      TabIndex        =   92
      Tag             =   "79,305"
      Top             =   2292
      Width           =   990
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "病人属于*："
      Height          =   180
      Index           =   12
      Left            =   105
      TabIndex        =   91
      Tag             =   "78,258"
      Top             =   1588
      Width           =   990
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "姓名*："
      Height          =   180
      Index           =   2
      Left            =   105
      TabIndex        =   80
      Tag             =   "78,164"
      Top             =   180
      Width           =   630
   End
   Begin VB.Line Line1 
      Index           =   1
      Tag             =   "119,175,210"
      X1              =   705
      X2              =   2265
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(患儿家长姓名："
      Height          =   180
      Index           =   3
      Left            =   2325
      TabIndex        =   81
      Tag             =   "228,164"
      Top             =   180
      Width           =   1350
   End
   Begin VB.Line Line1 
      Index           =   2
      Tag             =   "317,175,390"
      X1              =   3645
      X2              =   5205
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   ")"
      Height          =   180
      Index           =   4
      Left            =   5235
      TabIndex        =   82
      Tag             =   "403,164"
      Top             =   180
      Width           =   90
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "有效证件号*："
      Height          =   180
      Index           =   5
      Left            =   105
      TabIndex        =   83
      Tag             =   "78,187"
      Top             =   525
      Width           =   1170
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "性别*:"
      Height          =   180
      Index           =   6
      Left            =   6945
      TabIndex        =   84
      Tag             =   "498,187"
      Top             =   510
      Width           =   540
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "出生日期*："
      Height          =   180
      Index           =   7
      Left            =   105
      TabIndex        =   90
      Tag             =   "79,212"
      Top             =   884
      Width           =   990
   End
   Begin VB.Line Line1 
      Index           =   3
      Tag             =   "143,222,214"
      X1              =   1080
      X2              =   2280
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblBirth 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "年"
      Height          =   180
      Index           =   0
      Left            =   2175
      TabIndex        =   89
      Tag             =   "216,212"
      Top             =   884
      Width           =   180
   End
   Begin VB.Line Line1 
      Index           =   4
      Tag             =   "227,222,254"
      X1              =   2325
      X2              =   2985
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblBirth 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "月"
      Height          =   180
      Index           =   1
      Left            =   2970
      TabIndex        =   88
      Tag             =   "258,212"
      Top             =   884
      Width           =   180
   End
   Begin VB.Line Line1 
      Index           =   5
      Tag             =   "269,222,300"
      X1              =   3135
      X2              =   3735
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblBirth 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "日"
      Height          =   180
      Index           =   2
      Left            =   3735
      TabIndex        =   87
      Tag             =   "301,212"
      Top             =   884
      Width           =   180
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(如出生日期不详，实足年龄："
      Height          =   180
      Index           =   8
      Left            =   4125
      TabIndex        =   86
      Tag             =   "318,212"
      Top             =   884
      Width           =   2430
   End
   Begin VB.Line Line1 
      Index           =   6
      Tag             =   "479,222,526"
      X1              =   6525
      X2              =   7125
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "年龄单位:"
      Height          =   180
      Index           =   9
      Left            =   7290
      TabIndex        =   85
      Tag             =   "540,212"
      Top             =   884
      Width           =   810
   End
   Begin VB.Line Line1 
      Index           =   8
      Tag             =   "455,245,598"
      X1              =   6705
      X2              =   8520
      Y1              =   1425
      Y2              =   1425
   End
End
Attribute VB_Name = "PaneTwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDateType As Byte                           '0表示出生日期，1表示发病日期，2表示诊断日期，3表示死亡日期
Private mblnFirst As Boolean
Private mbln身份证必填 As Boolean                   '身份证信息必填 参数：传染病报告身份证号码必填
Private mcolLoadData As Collection
Public Event ClickPositives(blnSelected As Boolean)  '选择了阳性检测结果时触发

Private Sub lblAttack_Click(Index As Integer)
    mDateType = 1
    Call ShowMView(lblAttack(Index).Left, lblAttack(Index).Top)
End Sub

Public Function HaveChanged() As Boolean
'判断控件显示信息是否发生变化
    Dim objCtl As Control
    Dim i As Integer
    i = 0
    HaveChanged = False
    If mcolLoadData Is Nothing Then
        Set mcolLoadData = New Collection
    End If
    If mcolLoadData.Count <= 0 Then
        Exit Function
    End If
    For Each objCtl In UserControl.Controls
        Select Case TypeName(objCtl)
            Case "TextBox"
                If objCtl.Text <> mcolLoadData("K" & i) Then
                    HaveChanged = True
                    Exit Function
                End If
            Case "uCheckNorm"
                If IIf(objCtl.Checked = True, 1, 0) <> mcolLoadData("K" & i) Then
                    HaveChanged = True
                    Exit Function
                End If
        End Select
        i = i + 1
    Next
End Function

Private Sub SaveLoadData()
'功能：保存控件显示信息
    Dim objCtl As Control
    Dim i As Integer
    i = 0
    Set mcolLoadData = New Collection
    For Each objCtl In UserControl.Controls
        Select Case TypeName(objCtl)
            Case "TextBox"
                Call mcolLoadData.Add(objCtl.Text, "K" & i)
            Case "uCheckNorm"
                Call mcolLoadData.Add(IIf(objCtl.Checked = True, 1, 0), "K" & i)
        End Select
        i = i + 1
    Next
End Sub

Public Sub SetCaption身份证(ByVal blnHave As Boolean)
    mbln身份证必填 = blnHave
    If mbln身份证必填 Then
        lblReport(5).Caption = "有效证件号*："
    Else
        lblReport(5).Caption = "有效证件号："
    End If
End Sub

Public Sub ClearMe()
    Dim objCtl As Control
    
    On Error GoTo errHand
    For Each objCtl In UserControl.Controls
        Call ClearInfo(objCtl)
    Next
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub PrintTwo()
    Dim objCtl As Control
    For Each objCtl In UserControl.Controls
        Call PrintInfo(objCtl)
    Next
End Sub

Public Sub LoadData(colData As Collection, bytType As Byte, ByVal strChkType As String)
    Dim strTmp As String
    Dim i As Integer
    Dim strInfo() As String
    Dim objCtl As Control
    Dim dteTmp As Date
    
    On Error GoTo errHand
    mblnFirst = True
    If bytType = 1 Then  '修改
        txtName.Text = CStr(colData("K3"))          '姓名
        txtParentName.Text = CStr(colData("K4"))    '家长姓名
        
        '身份证号
        strTmp = CStr(colData("K5"))
        For i = 1 To 18
            txtIDCard(i - 1).Text = Mid(strTmp, i, 1)
        Next
        
        '性别，年龄单位，家人属于，人群分类，病例分类
        For Each objCtl In UserControl.Controls
            If TypeName(objCtl) = "uCheckNorm" Then
                strTmp = Trim(objCtl.Caption)
                strTmp = Replace(strTmp, "(", "")
                strTmp = Replace(strTmp, ")", "")
                strTmp = Replace(strTmp, "、", "")
                
                If objCtl.Name = "ucCheckJob" Then          '人群分类
                    If InStr(strChkType, "14," & strTmp) > 0 And Trim(strTmp) <> "" Then
                        objCtl.Checked = True
                    End If
                ElseIf objCtl.Name = "ucCaseType1" Then     '病例分类1
                    If InStr(strChkType, "15," & strTmp) > 0 And Trim(strTmp) <> "" Then
                        objCtl.Checked = True
                        
                        If strTmp = "疑似病例" Then
                            gbytDiseaseType = 0
                        ElseIf strTmp = "临床诊断病例" Then
                            gbytDiseaseType = 1
                        ElseIf strTmp = "确诊病例" Then
                            gbytDiseaseType = 2
                        ElseIf strTmp = "病原携带者" Then
                            gbytDiseaseType = 3
                        End If
                    End If
                ElseIf objCtl.Name = "ucCaseType2" Then     '病例分类2
                    If InStr(strChkType, "16," & strTmp) > 0 And Trim(strTmp) <> "" Then
                        objCtl.Checked = True
                        
                        If strTmp = "急性" Then
                            gbytAcute = 0
                        ElseIf strTmp = "慢性" Then
                            gbytAcute = 1
                        ElseIf strTmp = "未分型" Then
                            gbytAcute = 2
                        End If
                    End If
                ElseIf InStr(strChkType, strTmp) > 0 And Trim(strTmp) <> "" Then
                    objCtl.Checked = True
                End If
            End If
        Next
        
        '出生日期
        strTmp = CStr(colData("K7"))
        strInfo = Split(strTmp, "-")
        For i = 0 To UBound(strInfo)
            txtBirth(i).Text = IIf(val(strInfo(i)) = 0, "", strInfo(i))
        Next
        
        txtAge.Text = CStr(colData("K8"))           '年龄
        txtAddress.Text = CStr(colData("K10"))      '地址
        txtNumber.Text = CStr(colData("K11"))       '电话
        
        '住址
        strTmp = CStr(colData("K13"))
        strInfo = Split(strTmp, ";")
        For i = 0 To UBound(strInfo) - 1
            txtAddInfo(i).Text = strInfo(i)
        Next
        
        '发病日期
        strInfo = Split(CStr(colData("K17")), "-")
        For i = 0 To UBound(strInfo)
            txtAttack(i) = IIf(val(strInfo(i)) = 0, "", strInfo(i))
        Next
        
        '诊断时间
        strInfo = Split(CStr(colData("K18")), " ")
        If UBound(strInfo) > 0 Then
            txtDiagnose(3) = IIf(val(strInfo(1)) = 0, "", strInfo(1))
        End If
        If UBound(strInfo) >= 0 Then
            strInfo = Split(strInfo(0), "-")
            For i = 0 To UBound(strInfo)
                txtDiagnose(i) = IIf(val(strInfo(i)) = 0, "", strInfo(i))
            Next
        End If
        
        '死亡日期
        strInfo = Split(CStr(colData("K19")), "-")
        For i = 0 To UBound(strInfo)
            txtDeath(i) = IIf(val(strInfo(i)) = 0, "", strInfo(i))
        Next
    Else   '新增
        txtName.Text = CStr(colData("K0"))              '姓名
        txtParentName.Text = CStr(colData("KParent"))   '家长姓名
        '身份证号
        strTmp = CStr(colData("K1"))
        For i = 1 To 18
            txtIDCard(i - 1).Text = Mid(strTmp, i, 1)
        Next
        
        ucSex(IIf(CStr(colData("K2")) = "男", 0, 1)).Checked = True
        
        strTmp = Format(CStr(colData("K3")), "yyyy-mm-dd")
        If strTmp <> "年月日" Then
            strInfo = Split(strTmp, "-")
            For i = 0 To UBound(strInfo)
                txtBirth(i).Text = strInfo(i)
            Next
        End If
        
        strTmp = Trim(CStr(colData("K4")))
        i = InStr("大岁月日天", Right(strTmp, 1))
        If i > 1 Then
            i = IIf(i > 4, 4, i)
            txtAge.Text = val(CStr(colData("K4")))
            ucAge(i - 2).Checked = True
        Else
            txtAge.Text = val(CStr(colData("K4")))
            ucAge(0).Checked = True
            If val(txtAge.Text) = 0 Then
                txtAge.Text = ""
                ucAge(0).Checked = False
            End If
        End If
        
        txtAddress.Text = CStr(colData("K5"))
        
        If CStr(colData("K6")) <> "" Then
            strTmp = CStr(colData("K6"))
        ElseIf CStr(colData("K7")) <> "" Then
            strTmp = CStr(colData("K7"))
        Else
            strTmp = CStr(colData("K8"))
        End If
        txtNumber.Text = strTmp
        
        lblReport(13).ToolTipText = CStr(colData("K13"))
        
        strTmp = Format(CStr(colData("K14")), "yyyy-mm-dd")
        If strTmp <> "年月日" Then
            strInfo = Split(strTmp, "-")
            For i = 0 To UBound(strInfo)
                txtAttack(i).Text = strInfo(i)
            Next
        End If
        
        strTmp = Format(CStr(colData("K15")), "yyyy-mm-dd-hh")
        If strTmp <> "年月日" Then
            strInfo = Split(strTmp, "-")
            For i = 0 To UBound(strInfo)
                txtDiagnose(i).Text = strInfo(i)
            Next
        End If
        
        strTmp = Format(CStr(colData("K17")), "yyyy-mm-dd")
        If strTmp <> "年月日" Then
            strInfo = Split(strTmp, "-")
            For i = 0 To UBound(strInfo)
                txtDeath(i).Text = strInfo(i)
            Next
        End If
    End If
    mblnFirst = False
    Call SaveLoadData
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub
Public Function MakeSaveSql(arrSql() As Variant, colCls As Collection, strFileId As String) As Boolean
    Dim strObjNo As String
    Dim strContent As String
    Dim strReportInfo As String
    
    Dim i As Integer
    Dim strTmp As String
    Dim strTmp1 As String
    
    On Error GoTo errHand
    strObjNo = "3$4$5$6$7$8$9$10$11$12$13$14$15$16$17$18$19"
    
    '姓名、患者父母姓名
    strContent = txtName.Text & "$" & txtParentName.Text & "$"
    
    '身份证号码
    strTmp = ""
    For i = 0 To 17
        strTmp = Trim(strTmp) & Trim(txtIDCard(i).Text)
    Next
    strContent = strContent & strTmp & "$"
    
    '性别
    strTmp = IIf(ucSex(0).Checked = True, ucSex(0).Caption, IIf(ucSex(1).Checked = True, ucSex(1).Caption, ""))
    strContent = strContent & strTmp & "$"
    
    '出生日期
    strTmp = IIf(Trim(txtBirth(0).Text) = "", 0, Trim(txtBirth(0).Text)) & "-" & IIf(Trim(txtBirth(1).Text) = "", 0, Trim(txtBirth(1).Text)) & "-" & IIf(Trim(txtBirth(2).Text) = "", 0, Trim(txtBirth(2).Text))
    If Trim(strTmp) = "--" Then
        strTmp = ""
    End If
    strContent = strContent & strTmp & "$"
    
    '年龄
    strContent = strContent & Trim(txtAge.Text) & "$"
    
    '年龄单位
    strTmp = IIf(ucAge(0).Checked = True, ucAge(0).Caption, IIf(ucAge(1).Checked = True, ucAge(1).Caption, IIf(ucAge(2).Checked = True, ucAge(2).Caption, "")))
    strContent = strContent & strTmp & "$"
    
    '工作单位
    strContent = strContent & Trim(txtAddress.Text) & "$"
    
    '联系电话
    strContent = strContent & Trim(txtNumber.Text) & "$"
    
    '病人属于
    For i = 0 To 5
        If ucFrom(i).Checked = True Then
            strTmp = ucFrom(i).Caption
            Exit For
        End If
        strTmp = ""
    Next
    strContent = strContent & strTmp & "$"
    
    '现居住
    strTmp = ""
    For i = 0 To 5
        strTmp = Trim(strTmp) & Trim(txtAddInfo(i).Text) & ";"
    Next
    strContent = strContent & strTmp & "$"
    
    '患者职业
    For i = 0 To 19
        If ucCheckJob(i).Checked = True Then
            strTmp = ucCheckJob(i).Caption
            Exit For
        End If
        strTmp = ""
    Next
    strContent = strContent & strTmp & "$"
    
    '病例分类1
    For i = 0 To 3
        If ucCaseType1(i).Checked = True Then
            strTmp = ucCaseType1(i).Caption
            Exit For
        End If
        strTmp = ""
    Next
    strContent = strContent & strTmp & "$"
    
    '病例分类2
    For i = 0 To 2
        If ucCaseType2(i).Checked = True Then
            strTmp = ucCaseType2(i).Caption
            Exit For
        End If
        strTmp = ""
    Next
    strContent = strContent & strTmp & "$"
    
    '发病日期
    strTmp = IIf(Trim(txtAttack(0).Text) = "", 0, Trim(txtAttack(0).Text)) & "-" & IIf(Trim(txtAttack(1).Text) = "", 0, Trim(txtAttack(1).Text)) & "-" & IIf(Trim(txtAttack(2).Text) = "", 0, Trim(txtAttack(2).Text))
    If Trim(strTmp) = "--" Then
        strTmp = ""
    End If
    strContent = strContent & strTmp & "$"
    
    '诊断日期
    strTmp = IIf(Trim(txtDiagnose(0).Text) = "", 0, Trim(txtDiagnose(0).Text)) & "-" & IIf(Trim(txtDiagnose(1).Text) = "", 0, Trim(txtDiagnose(1).Text)) & "-" & IIf(Trim(txtDiagnose(2).Text) = "", 0, Trim(txtDiagnose(2).Text)) & " " & IIf(Trim(txtDiagnose(3).Text) = "", 0, Trim(txtDiagnose(3).Text))
    strContent = strContent & strTmp & "$"
    
    '死亡日期
    strTmp = IIf(Trim(txtDeath(0).Text) = "", 0, Trim(txtDeath(0).Text)) & "-" & IIf(Trim(txtDeath(1).Text) = "", 0, Trim(txtDeath(1).Text)) & "-" & IIf(Trim(txtDeath(2).Text) = "", 0, Trim(txtDeath(2).Text))
    If Trim(strTmp) = "--" Then
        strTmp = ""
    End If
    strContent = strContent & strTmp & "$"
    
    strReportInfo = strObjNo & "|" & strContent
    MakeSaveSql = GetSaveSql(arrSql, colCls, strFileId, strReportInfo)
    Call SaveLoadData
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Public Function CheckValidity(ByRef strMsg As String) As Boolean
'功能：检查输入的合法性
    Dim strBirth As String      '生日，例如：1991-01-01
    Dim blnIsChild As Boolean   '判断是否为14岁以下的患者
    Dim intAge As Integer
    Dim blnDate As Boolean      '判断日期是否输入完整
    Dim i As Integer
    Dim strTmp As String

    On Error GoTo errHand
    CheckValidity = False
    '检查姓名
    If Trim(txtName.Text) = "" Then
        strMsg = strMsg & "<姓名>为必选项，请检查！$"
    End If
    '检查身份证号
    If mbln身份证必填 Then
        For i = 0 To 17
            strTmp = strTmp & txtIDCard(i).Text
        Next
        If Trim(strTmp) = "" Then
            strMsg = strMsg & "<身份证号>为必选项，请检查！$"
        End If
    End If
    
    '检查性别，没有选择时必须选
    If ucSex(0).Checked = False And ucSex(1).Checked = False Then
        strMsg = strMsg & "<性别>为必选项，请检查！$"
    End If
    
    '检查出生日期 允许全空 不允许部份空或数值无效
    txtBirth(0).Text = Trim(txtBirth(0).Text): txtBirth(1).Text = Trim(txtBirth(1).Text): txtBirth(2).Text = Trim(txtBirth(2).Text)
    strBirth = txtBirth(0).Text & "-" & txtBirth(1).Text & "-" & txtBirth(2).Text
    
    If txtBirth(0).Text <> "" Or txtBirth(1).Text <> "" Or txtBirth(2).Text <> "" Then
        If Not IsDate(strBirth) Then
            strMsg = strMsg & "<出生日期>不完整或不是有效日期，请检查！$"
        Else
            If DateDiff("yyyy", strBirth, Now()) <= 14 And Trim(txtParentName.Text) = "" Then
                blnIsChild = True
            End If
        End If
    End If
    
    If Trim(txtBirth(0).Text) = "" And Trim(txtAge.Text) = "" Then
        strMsg = strMsg & "<出生日期>与<年龄>必需填写一项，请检查！$"
    End If
    
    '检查年龄，如果小于14，必须输入父母的名字,必须输入父母的联系电话
    If txtBirth(0).Text = "" Then
        intAge = val(txtAge.Text) * IIf(ucSex(0).Checked, 365, IIf(ucSex(1).Checked, 30, 1))
        If intAge <= (14 * 365) And Trim(txtParentName.Text) = "" Then
            blnIsChild = True
        End If
    End If
    If blnIsChild = True Then
        If Trim(txtNumber.Text) = "" Then
            strMsg = strMsg & "14岁以下患者要求填写<家长联系电话>，请检查！$"
        Else
            strMsg = strMsg & "14岁以下患者要求填写<家长姓名>，请检查！$"
        End If
    End If
    
    '检查病人属于
    For i = 0 To 5
        If ucFrom(i).Checked = True Then
            Exit For
        End If
        If i = 5 Then
            strMsg = strMsg & "<病人属于>为必选项，请检查！$"
        End If
    Next
    
    '检查地址
    For i = 0 To 5
        If Trim(txtAddInfo(i).Text) <> "" Then
            Exit For
        End If
        If i = 5 Then
            strMsg = strMsg & "<现住址>为必选项，请检查！$"
        End If
    Next
    
    '检查职业，必须选择一项
    For i = 0 To 19
        If ucCheckJob(i).Checked = True Then
            Exit For
        End If
        If i = 19 Then
            strMsg = strMsg & "<人群分类>为必选项，请检查！$"
        End If
    Next
    
    '检查病例分类
    For i = 0 To 3
        If ucCaseType1(i).Checked = True Then
            Exit For
        End If
        If i = 3 Then
            If ucCaseType2(0).Checked = False And ucCaseType2(1).Checked = False Then
                strMsg = strMsg & "<病例分类>为必选项，请检查！$"
            End If
        End If
    Next
    
    '发病日期必须填写
    txtAttack(0).Text = Trim(txtAttack(0).Text): txtAttack(1).Text = Trim(txtAttack(1).Text): txtAttack(2).Text = Trim(txtAttack(2).Text)
    If Not IsDate(txtAttack(0).Text & "-" & txtAttack(1).Text & "-" & txtAttack(2).Text) Then
        strMsg = strMsg & "<发病日期>不完整或不是有效日期，请检查！$"
    End If
    
    '诊断时间必须完整并精确到小时
    txtDiagnose(0).Text = Trim(txtDiagnose(0).Text): txtDiagnose(1).Text = Trim(txtDiagnose(1).Text):
    txtDiagnose(2).Text = Trim(txtDiagnose(2).Text): txtDiagnose(3).Text = Trim(txtDiagnose(3).Text)
    If (Not IsDate(txtDiagnose(0).Text & "-" & txtDiagnose(1).Text & "-" & txtDiagnose(2).Text)) Or (Decode(txtDiagnose(3).Text, "", "-1", txtDiagnose(3).Text) < 0) Or (Decode(txtDiagnose(3).Text, "", "-1", txtDiagnose(3).Text) > 23) Then
        strMsg = strMsg & "<诊断时间>不完整，或不是有效日期，或未精确到小时，请检查！$"
    End If
    
    '死亡日期 允许全空 不允许部份空或数值无效
    txtDeath(0).Text = Trim(txtDeath(0).Text): txtDeath(1).Text = Trim(txtDeath(1).Text): txtDeath(2).Text = Trim(txtDeath(2).Text)
    If txtDeath(0).Text <> "" Or txtDeath(1).Text <> "" Or txtDeath(2).Text <> "" Then
        If Not IsDate(txtDeath(0).Text & "-" & txtDeath(1).Text & "-" & txtDeath(2).Text) Then
            strMsg = strMsg & "<死亡日期>不完整或不是有效日期，请检查！$"
        End If
    End If
    
    CheckValidity = True
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Private Sub ShowMView(x As Long, y As Long)
'功能：显示日期选择控件
    
    MView.Left = x
    If mDateType = 0 Then
        MView.Top = y + 200
    ElseIf mDateType = 3 Then
        MView.Top = y - MView.Height
        MView.Left = txtDeath(0).Left
    Else
        MView.Top = y - MView.Height
    End If
    MView.Visible = True
    Call MView.SetFocus
End Sub

Private Sub lblAttack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Set lblAttack(Index).MouseIcon = Image1.Picture
    lblAttack(Index).MousePointer = vbCustom
End Sub

Private Sub lblBirth_Click(Index As Integer)
    mDateType = 0
    Call ShowMView(lblBirth(Index).Left, lblBirth(Index).Top)
End Sub

Private Sub lblBirth_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Set lblBirth(Index).MouseIcon = Image1.Picture
    lblBirth(Index).MousePointer = vbCustom
End Sub

Private Sub lblDeath_Click(Index As Integer)
    mDateType = 3
    Call ShowMView(lblDeath(Index).Left, lblDeath(Index).Top)
End Sub

Private Sub lblDeath_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Set lblDeath(Index).MouseIcon = Image1.Picture
    lblDeath(Index).MousePointer = vbCustom
End Sub

Private Sub lblDiagnose_Click(Index As Integer)
    If Index <> 3 Then
        mDateType = 2
        Call ShowMView(lblDiagnose(Index).Left, lblDiagnose(Index).Top)
    End If
End Sub

Private Sub lblDiagnose_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index <> 3 Then
        Set lblDiagnose(Index).MouseIcon = Image1.Picture
        lblDiagnose(Index).MousePointer = vbCustom
    End If
End Sub

Private Sub MView_DateClick(ByVal DateClicked As Date)
    MView.Visible = False
    Select Case mDateType
        Case 0
            txtBirth(0).Text = MView.Year
            txtBirth(1).Text = MView.Month
            txtBirth(2).Text = MView.Day
        Case 1
            txtAttack(0).Text = MView.Year
            txtAttack(1).Text = MView.Month
            txtAttack(2).Text = MView.Day
        Case 2
            txtDiagnose(0).Text = MView.Year
            txtDiagnose(1).Text = MView.Month
            txtDiagnose(2).Text = MView.Day
        Case 3
            txtDeath(0).Text = MView.Year
            txtDeath(1).Text = MView.Month
            txtDeath(2).Text = MView.Day
    End Select
End Sub

Private Sub MView_LostFocus()
    MView.Visible = False
End Sub

Private Sub txtAge_Change()
    If mblnFirst = False Then
        ucAge(1).Checked = IIf(Trim(txtAge.Text) = "", False, True)
    End If
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    Call CheckVal(KeyAscii)
End Sub

Private Sub txtAttack_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckVal(KeyAscii)
End Sub

Private Sub txtBirth_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckVal(KeyAscii)
    txtAge.Text = ""
End Sub

Private Sub txtDeath_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckVal(KeyAscii)
End Sub

Private Sub txtDiagnose_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckVal(KeyAscii)
End Sub

Private Sub txtIDCard_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Then
        SendKeys "+{TAB}"
    ElseIf KeyCode = vbKeyRight Then
        SendKeys "{TAB}"
    Else
        txtIDCard(Index).SelStart = 0
        txtIDCard(Index).SelLength = Len(txtIDCard(Index).Text)
    End If
End Sub

Private Sub txtIDCard_KeyPress(Index As Integer, KeyAscii As Integer)
    If zlStr.IsNumOrChar(Chr(KeyAscii)) Then
        SendKeys "{TAB}"
    ElseIf Not (KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If
End Sub

Private Sub ucCaseType1_Change(Index As Integer)
    gbytDiseaseType = 5
    If ucCaseType1(Index).Checked = True Then
        gbytDiseaseType = Index
    End If
    If Index = 4 Then
        RaiseEvent ClickPositives(ucCaseType1(4).Checked)
    End If
End Sub

Private Sub ucCaseType2_Change(Index As Integer)
    gbytAcute = 3
    If ucCaseType2(Index).Checked = True Then
        gbytAcute = Index
    End If
End Sub

Private Sub UserControl_Initialize()
    UserControl.BackColor = vbWindowBackground
End Sub
