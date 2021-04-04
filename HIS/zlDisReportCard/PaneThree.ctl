VERSION 5.00
Begin VB.UserControl PaneThree 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9825
   ScaleHeight     =   4560
   ScaleWidth      =   9825
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   4
      Left            =   2535
      TabIndex        =   12
      Tag             =   "164,561"
      Top             =   1360
      Width           =   2250
      _ExtentX        =   19209
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "人感染高致病性禽流感、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   3
      Left            =   1185
      TabIndex        =   11
      Tag             =   "77,561"
      Top             =   1360
      Width           =   1350
      _ExtentX        =   17621
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "脊髓灰质炎、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucPTB 
      Height          =   270
      Index           =   2
      Left            =   1110
      TabIndex        =   25
      Tag             =   "77,608"
      Top             =   2120
      Width           =   810
      _ExtentX        =   16669
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "菌阴、"
   End
   Begin zlDisReportCard.uCheckNorm ucPTB 
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   24
      Tag             =   "641,584"
      Top             =   2120
      Width           =   1005
      _ExtentX        =   17013
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "仅培阳、"
   End
   Begin zlDisReportCard.uCheckNorm ucPTB 
      Height          =   270
      Index           =   0
      Left            =   8435
      TabIndex        =   23
      Tag             =   "591,584"
      Top             =   1740
      Width           =   810
      _ExtentX        =   16669
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "涂阳、"
   End
   Begin zlDisReportCard.uCheckNorm ucAIDS 
      Height          =   270
      Index           =   0
      Left            =   4085
      TabIndex        =   4
      Tag             =   "259,538"
      Top             =   980
      Width           =   760
      _ExtentX        =   16589
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "HIV)、"
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   17
      Left            =   105
      TabIndex        =   32
      Tag             =   "588,608"
      Top             =   2500
      Width           =   1530
      _ExtentX        =   17939
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "新生儿破伤风、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   16
      Left            =   8325
      TabIndex        =   31
      Tag             =   "537,608"
      Top             =   2120
      Width           =   810
      _ExtentX        =   16669
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "白喉、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   15
      Left            =   7320
      TabIndex        =   30
      Tag             =   "474,608"
      Top             =   2120
      Width           =   1005
      _ExtentX        =   17013
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "百日咳、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   14
      Left            =   5415
      TabIndex        =   29
      Tag             =   "349,608"
      Top             =   2120
      Width           =   1905
      _ExtentX        =   18600
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "流行性脑脊髓膜炎、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucTyphia 
      Height          =   270
      Index           =   1
      Left            =   4335
      TabIndex        =   28
      Tag             =   "280,608"
      Top             =   2120
      Width           =   1080
      _ExtentX        =   17145
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "副伤寒)、"
   End
   Begin zlDisReportCard.uCheckNorm ucTyphia 
      Height          =   270
      Index           =   0
      Left            =   3525
      TabIndex        =   27
      Tag             =   "229,608"
      Top             =   2120
      Width           =   810
      _ExtentX        =   16669
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "伤寒、"
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   13
      Left            =   3000
      TabIndex        =   69
      Tag             =   "195,608"
      Top             =   2120
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "伤寒("
      CheckType       =   1
      BoxVisible      =   0   'False
      CheckedVisible  =   0   'False
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   11
      Left            =   4890
      TabIndex        =   67
      Tag             =   "366,584"
      Top             =   1740
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "痢疾("
      CheckType       =   1
      BoxVisible      =   0   'False
      CheckedVisible  =   0   'False
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   12
      Left            =   7685
      TabIndex        =   68
      Tag             =   "544,584"
      Top             =   1740
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "肺结核("
      CheckType       =   1
      BoxVisible      =   0   'False
      CheckedVisible  =   0   'False
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   9
      Left            =   8115
      TabIndex        =   16
      Tag             =   "617,561"
      Top             =   1360
      Width           =   1830
      _ExtentX        =   15928
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "流行性乙型脑炎、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   8
      Left            =   7110
      TabIndex        =   15
      Tag             =   "552,561"
      Top             =   1360
      Width           =   1005
      _ExtentX        =   17013
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "狂犬病、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   7
      Left            =   5595
      TabIndex        =   14
      Tag             =   "452,561"
      Top             =   1360
      Width           =   1530
      _ExtentX        =   17939
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "流行性出血热、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   6
      Left            =   4785
      TabIndex        =   13
      Tag             =   "400,561"
      Top             =   1360
      Width           =   810
      _ExtentX        =   16669
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "麻疹、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   1
      Left            =   1990
      TabIndex        =   65
      Tag             =   "202,538"
      Top             =   980
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "艾滋病("
      CheckType       =   1
      BoxVisible      =   0   'False
      CheckedVisible  =   0   'False
   End
   Begin zlDisReportCard.uCheckNorm ucAIDS 
      Height          =   270
      Index           =   1
      Left            =   2725
      TabIndex        =   3
      Tag             =   "304,538"
      Top             =   975
      Width           =   1350
      _ExtentX        =   17198
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "艾滋病病人、"
   End
   Begin zlDisReportCard.uCheckNorm ucHepatitis 
      Height          =   270
      Index           =   5
      Left            =   105
      TabIndex        =   10
      Tag             =   "647,538"
      Top             =   1360
      Width           =   1080
      _ExtentX        =   17145
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "未分型)、"
   End
   Begin zlDisReportCard.uCheckNorm ucHepatitis 
      Height          =   270
      Index           =   4
      Left            =   9120
      TabIndex        =   9
      Tag             =   "595,538"
      Top             =   980
      Width           =   810
      _ExtentX        =   16669
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "戊型、"
   End
   Begin zlDisReportCard.uCheckNorm ucHepatitis 
      Height          =   270
      Index           =   2
      Left            =   7500
      TabIndex        =   7
      Tag             =   "544,538"
      Top             =   980
      Width           =   810
      _ExtentX        =   16669
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "丙型、"
   End
   Begin zlDisReportCard.uCheckNorm ucHepatitis 
      Height          =   270
      Index           =   1
      Left            =   6690
      TabIndex        =   6
      Tag             =   "493,538"
      Top             =   980
      Width           =   810
      _ExtentX        =   16669
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "乙型、"
   End
   Begin zlDisReportCard.uCheckNorm ucHepatitis 
      Height          =   270
      Index           =   0
      Left            =   5880
      TabIndex        =   5
      Tag             =   "442,538"
      Top             =   980
      Width           =   810
      _ExtentX        =   16669
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "甲型、"
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   2
      Left            =   4845
      TabIndex        =   66
      Tag             =   "357,538"
      Top             =   980
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "病毒性肝炎("
      CheckType       =   1
      BoxVisible      =   0   'False
      CheckedVisible  =   0   'False
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   5
      Left            =   6405
      TabIndex        =   46
      Tag             =   "300,561"
      Top             =   2880
      Width           =   1770
      _ExtentX        =   17515
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "人感染H7N9禽流感"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucAnthrax 
      Height          =   270
      Index           =   2
      Left            =   3810
      TabIndex        =   20
      Tag             =   "297,584"
      Top             =   1740
      Width           =   1080
      _ExtentX        =   19262
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "未分型)、"
   End
   Begin zlDisReportCard.uCheckNorm ucAnthrax 
      Height          =   270
      Index           =   0
      Left            =   1635
      TabIndex        =   18
      Tag             =   "159,584"
      Top             =   1740
      Width           =   1005
      _ExtentX        =   19129
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "肺炭疽、"
   End
   Begin zlDisReportCard.uCheckNorm ucAnthrax 
      Height          =   270
      Index           =   1
      Left            =   2640
      TabIndex        =   19
      Tag             =   "221,584"
      Top             =   1740
      Width           =   1170
      _ExtentX        =   19420
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "皮肤炭疽、"
   End
   Begin zlDisReportCard.uCheckNorm ucDysentery 
      Height          =   270
      Index           =   1
      Left            =   6420
      TabIndex        =   22
      Tag             =   "464,584"
      Top             =   1740
      Width           =   1265
      _ExtentX        =   17463
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "阿米巴性)、"
   End
   Begin zlDisReportCard.uCheckNorm ucDysentery 
      Height          =   270
      Index           =   0
      Left            =   5415
      TabIndex        =   21
      Tag             =   "401,584"
      Top             =   1740
      Width           =   1005
      _ExtentX        =   19129
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "细菌性、"
   End
   Begin zlDisReportCard.uCheckNorm ucPTB 
      Height          =   270
      Index           =   3
      Left            =   1920
      TabIndex        =   26
      Tag             =   "128,608"
      Top             =   2120
      Width           =   1080
      _ExtentX        =   17145
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "未痰检)、"
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   21
      Left            =   105
      TabIndex        =   41
      Tag             =   "563,631"
      Top             =   2880
      Width           =   1530
      _ExtentX        =   17939
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "钩端螺旋体病、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucSyphilis 
      Height          =   270
      Index           =   4
      Left            =   8565
      TabIndex        =   40
      Tag             =   "513,631"
      Top             =   2520
      Width           =   900
      _ExtentX        =   18944
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "隐性)、"
   End
   Begin zlDisReportCard.uCheckNorm ucSyphilis 
      Height          =   270
      Index           =   3
      Left            =   7755
      TabIndex        =   39
      Tag             =   "462,631"
      Top             =   2505
      Width           =   810
      _ExtentX        =   16669
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "胎传、"
   End
   Begin zlDisReportCard.uCheckNorm ucSyphilis 
      Height          =   270
      Index           =   2
      Left            =   6945
      TabIndex        =   38
      Tag             =   "411,631"
      Top             =   2505
      Width           =   810
      _ExtentX        =   16669
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "Ⅲ期、"
   End
   Begin zlDisReportCard.uCheckNorm ucSyphilis 
      Height          =   270
      Index           =   1
      Left            =   6135
      TabIndex        =   37
      Tag             =   "360,631"
      Top             =   2505
      Width           =   810
      _ExtentX        =   16669
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "Ⅱ期、"
   End
   Begin zlDisReportCard.uCheckNorm ucSyphilis 
      Height          =   270
      Index           =   0
      Left            =   5325
      TabIndex        =   36
      Tag             =   "309,631"
      Top             =   2505
      Width           =   810
      _ExtentX        =   16669
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "Ⅰ期、"
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   25
      Left            =   3985
      TabIndex        =   35
      Tag             =   "227,631"
      Top             =   2505
      Width           =   810
      _ExtentX        =   16669
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "淋病、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   19
      Left            =   2635
      TabIndex        =   34
      Tag             =   "140,631"
      Top             =   2505
      Width           =   1350
      _ExtentX        =   17621
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "布鲁氏菌病、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   18
      Left            =   1630
      TabIndex        =   33
      Tag             =   "77,631"
      Top             =   2505
      Width           =   1005
      _ExtentX        =   17013
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "猩红热、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucMalaria 
      Height          =   270
      Index           =   0
      Left            =   3315
      TabIndex        =   43
      Tag             =   "125,654"
      Top             =   2880
      Width           =   1005
      _ExtentX        =   17013
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "间日疟、"
   End
   Begin zlDisReportCard.uCheckNorm ucMalaria 
      Height          =   270
      Index           =   1
      Left            =   4320
      TabIndex        =   44
      Tag             =   "188,654"
      Top             =   2880
      Width           =   1005
      _ExtentX        =   17013
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "恶性疟、"
   End
   Begin zlDisReportCard.uCheckNorm ucMalaria 
      Height          =   270
      Index           =   2
      Left            =   5325
      TabIndex        =   45
      Tag             =   "247,654"
      Top             =   2880
      Width           =   1080
      _ExtentX        =   17145
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "未分型)、"
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   6
      Left            =   705
      TabIndex        =   53
      Tag             =   "665,702"
      Top             =   3885
      Width           =   1005
      _ExtentX        =   19129
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "黑热病、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   5
      Left            =   7930
      TabIndex        =   52
      Tag             =   "504,702"
      Top             =   3555
      Width           =   2400
      _ExtentX        =   21590
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "流行性和地方性斑疹伤寒、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   4
      Left            =   6975
      TabIndex        =   51
      Tag             =   "441,702"
      Top             =   3555
      Width           =   1005
      _ExtentX        =   19129
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "麻风病、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   3
      Left            =   5105
      TabIndex        =   50
      Tag             =   "316,702"
      Top             =   3555
      Width           =   1890
      _ExtentX        =   20690
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "急性出血性结膜炎、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   2
      Left            =   4325
      TabIndex        =   49
      Tag             =   "265,702"
      Top             =   3555
      Width           =   810
      _ExtentX        =   18785
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "风疹、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   1
      Left            =   2820
      TabIndex        =   48
      Tag             =   "164,702"
      Top             =   3555
      Width           =   1525
      _ExtentX        =   20055
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "流行性腮腺炎、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   8
      Left            =   2715
      TabIndex        =   55
      Tag             =   "140,725"
      Top             =   3885
      Width           =   1005
      _ExtentX        =   19129
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "丝虫病，"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   9
      Left            =   3720
      TabIndex        =   56
      Tag             =   "203,725"
      Top             =   3885
      Width           =   5850
      _ExtentX        =   114035
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "除霍乱、细菌性和阿米巴性痢疾、伤寒和副伤寒以外的感染性腹泻病、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   10
      Left            =   105
      TabIndex        =   57
      Tag             =   "590,725"
      Top             =   4245
      Width           =   1215
      _ExtentX        =   17383
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "手足口病。"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousA 
      Height          =   270
      Index           =   1
      Left            =   1050
      TabIndex        =   1
      Tag             =   "129,490"
      Top             =   301
      Width           =   825
      _ExtentX        =   103108
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "霍乱"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Tag             =   "77,538"
      Top             =   980
      Width           =   1885
      _ExtentX        =   18574
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "传染性非典型肺炎、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   10
      Left            =   1110
      TabIndex        =   61
      Tag             =   "77,584"
      Top             =   1740
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "炭疽("
      CheckType       =   1
      BoxVisible      =   0   'False
      CheckedVisible  =   0   'False
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   22
      Left            =   1635
      TabIndex        =   42
      Tag             =   "660,631"
      Top             =   2880
      Width           =   1170
      _ExtentX        =   17304
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "血吸虫病、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   23
      Left            =   2805
      TabIndex        =   70
      Tag             =   "78,654"
      Top             =   2880
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "疟疾("
      CheckType       =   1
      BoxVisible      =   0   'False
      CheckedVisible  =   0   'False
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   47
      Tag             =   "77,702"
      Top             =   3555
      Width           =   1170
      _ExtentX        =   17304
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "流行性感冒"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousC 
      Height          =   270
      Index           =   7
      Left            =   1710
      TabIndex        =   54
      Tag             =   "77,725"
      Top             =   3885
      Width           =   1005
      _ExtentX        =   19129
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "包虫病、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousA 
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Tag             =   "77,490"
      Top             =   270
      Width           =   825
      _ExtentX        =   16695
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "鼠疫、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm ucHepatitis 
      Height          =   270
      Index           =   3
      Left            =   8310
      TabIndex        =   8
      Tag             =   "595,538"
      Top             =   980
      Width           =   810
      _ExtentX        =   16669
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "丁型、"
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   24
      Left            =   105
      TabIndex        =   17
      Tag             =   "617,561"
      Top             =   1740
      Width           =   1005
      _ExtentX        =   17013
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "登革热、"
      CheckType       =   1
   End
   Begin zlDisReportCard.uCheckNorm shanghan 
      Height          =   270
      Left            =   105
      TabIndex        =   63
      Tag             =   "77,584"
      Top             =   3885
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "伤寒、"
      CheckType       =   1
      BoxVisible      =   0   'False
      CheckedVisible  =   0   'False
   End
   Begin zlDisReportCard.uCheckNorm ucInfectiousB 
      Height          =   270
      Index           =   20
      Left            =   4795
      TabIndex        =   64
      Tag             =   "227,631"
      Top             =   2505
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "梅毒("
      CheckType       =   1
      BoxVisible      =   0   'False
      CheckedVisible  =   0   'False
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(含甲型H1N1流感)、"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1280
      TabIndex        =   62
      Top             =   3600
      Width           =   1620
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   0
      X2              =   9982
      Y1              =   665
      Y2              =   665
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   9982
      Y1              =   3255
      Y2              =   3255
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "丙类传染病*："
      Height          =   180
      Index           =   29
      Left            =   105
      TabIndex        =   60
      Tag             =   "78,680"
      Top             =   3285
      Width           =   1170
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "乙类传染病*："
      Height          =   180
      Index           =   28
      Left            =   105
      TabIndex        =   59
      Tag             =   "79,517"
      Top             =   684
      Width           =   1170
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "甲类传染病*："
      Height          =   180
      Index           =   27
      Left            =   105
      TabIndex        =   58
      Tag             =   "79,468"
      Top             =   8
      Width           =   1170
   End
End
Attribute VB_Name = "PaneThree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mcolLoadData As Collection  '保存控件显示信息

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
    Err = 0
End Sub

Public Sub PrintThree()
    Dim objCtl As Control
    For Each objCtl In UserControl.Controls
        Call PrintInfo(objCtl)
    Next
End Sub

Public Sub LoadData(colData As Collection, bytType As Byte, ByVal strChkType As String)
    Dim strTmp As String
    Dim strName As String
    Dim i As Integer
    Dim strInfo() As String
    Dim objCtl As Control
    
    On Error GoTo errHand
    If bytType = 1 Then    '修改
        For Each objCtl In UserControl.Controls
            If TypeName(objCtl) = "uCheckNorm" Then
                strTmp = Trim(objCtl.Caption)
                strTmp = Replace(strTmp, "(", "")
                strTmp = Replace(strTmp, ")", "")
                strTmp = Replace(strTmp, "、", "")
                Select Case objCtl.Name
                    Case "ucAIDS"           '艾滋病
                        If InStr(strChkType, "22," & strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                    Case "ucHepatitis"        '病毒性肝炎
                        If InStr(strChkType, "23," & strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                    Case "ucAnthrax"          '炭疽
                        If InStr(strChkType, "24," & strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                    Case "ucDysentery"        '痢疾
                        If InStr(strChkType, "25," & strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                    Case "ucPTB"              '肺结核
                        If InStr(strChkType, "26," & strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                    Case "ucTyphia"          '伤寒
                        If InStr(strChkType, "27," & strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                    Case "ucSyphilis"         '梅毒
                        If InStr(strChkType, "28," & strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                    Case "ucMalaria"          '疟疾
                        If InStr(strChkType, "29," & strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                    Case "ucInfectiousA"     '甲类传染病
                        If InStr(strChkType, "20," & strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                    Case Else                 '
                        If InStr(strChkType, strTmp) > 0 And Trim(strTmp) <> "" Then
                            objCtl.Checked = True
                        End If
                End Select
            End If
        Next
    Else  '新增
        strTmp = CStr(colData("K16"))
        
        '甲类传染病
        ucInfectiousA(0).Checked = IIf(InStr(strTmp, "鼠疫") <> 0, True, False)
        ucInfectiousA(1).Checked = IIf(InStr(strTmp, "霍乱") <> 0, True, False)
        
        '乙类传染病
        For i = 0 To 25
            strName = Trim(ucInfectiousB(i).Caption)
            strName = Mid(strName, 1, Len(strName) - 1)
            If InStr(strTmp, strName) <> 0 Then
                If ucInfectiousB(i).CheckedVisible Then ucInfectiousB(i).Checked = True
            End If
        Next
        
        '丙类传染病
        For i = 0 To 10
            strName = Trim(ucInfectiousC(i).Caption)
            strName = Mid(strName, 1, Len(strName) - 1)
            If InStr(strTmp, strName) <> 0 Then
                ucInfectiousC(i).Checked = True
            End If
        Next
        
        '艾滋病
        ucAIDS(0).Checked = IIf(InStr(strTmp, "艾滋病病人") <> 0, True, False)
        ucAIDS(1).Checked = IIf(InStr(strTmp, "AIDS") <> 0, True, False)
        
        '病毒性肝炎
        For i = 0 To 5
            strName = Trim(ucHepatitis(i).Caption)
            strName = Mid(strName, 1, Len(strName) - IIf(i = 5, 2, 1))
            If InStr(strTmp, strName) <> 0 Then
                ucHepatitis(i).Checked = True
            End If
        Next
        
        '炭疽
        ucAnthrax(0).Checked = IIf(InStr(strTmp, "肺炭疽") <> 0, True, False)
        ucAnthrax(1).Checked = IIf(InStr(strTmp, "皮肤炭疽") <> 0, True, False)
        ucAnthrax(2).Checked = IIf(InStr(strTmp, "未分型") <> 0, True, False)
        
        '痢疾
        ucDysentery(0).Checked = IIf(InStr(strTmp, "细菌性") <> 0, True, False)
        ucDysentery(1).Checked = IIf(InStr(strTmp, "阿米巴性") <> 0, True, False)
        
        '肺结核
        For i = 0 To 3
            strName = Trim(ucPTB(i).Caption)
            strName = Mid(strName, 1, Len(strName) - IIf(i = 3, 2, 1))
            If InStr(strTmp, strName) <> 0 Then
                ucPTB(i).Checked = True
            End If
        Next
        
        '伤寒
        ucTyphia(0).Checked = IIf(InStr(strTmp, "伤寒") <> 0, True, False)
        ucTyphia(1).Checked = IIf(InStr(strTmp, "副伤寒") <> 0, True, False)
        
        '梅毒
        For i = 0 To 4
            strName = Trim(ucSyphilis(i).Caption)
            strName = Mid(strName, 1, Len(strName) - IIf(i = 4, 2, 1))
            If InStr(strTmp, strName) <> 0 Then
                ucSyphilis(i).Checked = True
            End If
        Next
        
        '疟疾
        ucMalaria(0).Checked = IIf(InStr(strTmp, "间日疟") <> 0, True, False)
        ucMalaria(1).Checked = IIf(InStr(strTmp, "恶性疟") <> 0, True, False)
        ucMalaria(2).Checked = IIf(InStr(strTmp, "未分型") <> 0, True, False)
    End If
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
    Dim strTmpB As String
    On Error GoTo errHand
    strObjNo = "20$21$22$23$24$25$26$27$28$29$30"
    
    '甲类传染病
    strTmp = IIf(ucInfectiousA(0).Checked = True, ucInfectiousA(0).Caption & ";", "")
    strTmp = Trim(strTmp) & Trim(IIf(ucInfectiousA(1).Checked = True, ucInfectiousA(1).Caption, ""))
    strContent = strContent & strTmp & "$"
    
    '艾滋病
    strTmp = Decode(True, ucAIDS(0).Checked, ucAIDS(0).Caption, ucAIDS(1).Checked, ucAIDS(1).Caption, "")
    strTmpB = strTmpB & strTmp & "$"
    ucInfectiousB(1).Checked = IIf(strTmp <> "", True, False)
    
    '病毒性肝炎
    strTmp = Decode(True, ucHepatitis(0).Checked, ucHepatitis(0).Caption, ucHepatitis(1).Checked, ucHepatitis(1).Caption, ucHepatitis(2).Checked, ucHepatitis(2).Caption, ucHepatitis(3).Checked, ucHepatitis(3).Caption, ucHepatitis(4).Checked, ucHepatitis(4).Caption, ucHepatitis(5).Checked, ucHepatitis(5).Caption, "")
    strTmpB = strTmpB & strTmp & "$"
    ucInfectiousB(2).Checked = IIf(strTmp <> "", True, False)
    
    '炭疽
    strTmp = Decode(True, ucAnthrax(0).Checked, ucAnthrax(0).Caption, ucAnthrax(1).Checked, ucAnthrax(1).Caption, ucAnthrax(2).Checked, ucAnthrax(2).Caption, "")
    strTmpB = strTmpB & strTmp & "$"
    ucInfectiousB(10).Checked = IIf(strTmp <> "", True, False)
    
    '痢疾
    strTmp = Decode(True, ucDysentery(0).Checked, ucDysentery(0).Caption, ucDysentery(1).Checked, ucDysentery(1).Caption, "")
    strTmpB = strTmpB & strTmp & "$"
    ucInfectiousB(11).Checked = IIf(strTmp <> "", True, False)
    
    '肺结核
    strTmp = Decode(True, ucPTB(0).Checked, ucPTB(0).Caption, ucPTB(1).Checked, ucPTB(1).Caption, ucPTB(2).Checked, ucPTB(2).Caption, ucPTB(3).Checked, ucPTB(3).Caption, "")
    strTmpB = strTmpB & strTmp & "$"
    ucInfectiousB(12).Checked = IIf(strTmp <> "", True, False)
    
    '伤寒
    strTmp = Decode(True, ucTyphia(0).Checked, ucTyphia(0).Caption, ucTyphia(1).Checked, ucTyphia(1).Caption, "")
    strTmpB = strTmpB & strTmp & "$"
    ucInfectiousB(13).Checked = IIf(strTmp <> "", True, False)
    
    '梅毒
    strTmp = Decode(True, ucSyphilis(0).Checked, ucSyphilis(0).Caption, ucSyphilis(1).Checked, ucSyphilis(1).Caption, ucSyphilis(2).Checked, ucSyphilis(2).Caption, ucSyphilis(3).Checked, ucSyphilis(3).Caption, ucSyphilis(4).Checked, ucSyphilis(4).Caption, "")
    strTmpB = strTmpB & strTmp & "$"
    ucInfectiousB(20).Checked = IIf(strTmp <> "", True, False)
    
    '疟疾
    strTmp = Decode(True, ucMalaria(0).Checked, ucMalaria(0).Caption, ucMalaria(1).Checked, ucMalaria(1).Caption, ucMalaria(2).Checked, ucMalaria(2).Caption, "")
    strTmpB = strTmpB & strTmp & "$"
    ucInfectiousB(23).Checked = IIf(strTmp <> "", True, False)
    
    '乙类传染病
    strTmp = ""
    For i = 0 To ucInfectiousB.UBound
        If ucInfectiousB(i).Checked = True Then
            strTmp = strTmp & ";" & ucInfectiousB(i).Caption
        End If
    Next
    If strTmp <> "" Then
        strTmp = Mid(strTmp, 2)
    End If
    strContent = strContent & strTmp & "$"
    
    '连接乙类单个疾病
    strContent = strContent & strTmpB
    
    '丙类传染病
    strTmp = ""
    For i = 0 To ucInfectiousC.UBound
        If ucInfectiousC(i).Checked = True Then
            If i = 5 Then
                strTmp = strTmp & ";" & ucInfectiousC(i).Caption & "伤寒、"
            Else
                strTmp = strTmp & ";" & ucInfectiousC(i).Caption
            End If
        End If
    Next
    If strTmp <> "" Then
        strTmp = Mid(strTmp, 2)
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
'检查输入合法性
    On Error GoTo errHand
    CheckValidity = False
    '1.检查病例分类为"病原携带者"时，病种是否是"霍乱、脊髓灰质炎、艾滋病"
    If gbytDiseaseType = 3 Then
        If ucInfectiousA(1).Checked = False And (ucAIDS(0).Checked = False And ucAIDS(1).Checked = False) And ucInfectiousB(3).Checked = False Then
            strMsg = strMsg & "需报告病原携带者的法定传染病病种包括<霍乱>、<脊髓灰质炎>、<艾滋病>，请检查！$"
        End If
    End If
    
    '2."梅毒"、"淋病"的病例分类只能为"确诊病例"和"疑似病例"
    If ucInfectiousB(25).Checked Or ucSyphilis(0).Checked Or ucSyphilis(1).Checked Or ucSyphilis(2).Checked Or ucSyphilis(3).Checked Or ucSyphilis(4).Checked Then
        If gbytDiseaseType <> 0 And gbytDiseaseType <> 2 Then
            strMsg = strMsg & "<梅毒、淋病>的病例分类只能为<确诊病例>和<疑似病例>！$"
        End If
    End If
    
    '3.乙肝、血吸虫病例须分急性或慢性填写
    If ucHepatitis(1).Checked = True Or ucInfectiousB(22).Checked = True Then
        If gbytAcute <> 0 And gbytAcute <> 1 Then
            strMsg = strMsg & "<乙肝>、<血吸虫病例>须分<急性>或<慢性>，请检查！$"
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

Private Sub UserControl_Initialize()
    UserControl.BackColor = vbWindowBackground
End Sub
