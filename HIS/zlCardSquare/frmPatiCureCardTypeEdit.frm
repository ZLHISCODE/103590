VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmPatiCureCardTypeEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "医疗卡类别编辑"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9285
   Icon            =   "frmPatiCureCardTypeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   9285
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picProperty 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1965
      Index           =   0
      Left            =   8700
      ScaleHeight     =   1965
      ScaleWidth      =   7875
      TabIndex        =   83
      Top             =   4020
      Width           =   7875
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2040
         TabIndex        =   92
         Top             =   210
         Width           =   5430
         Begin VB.OptionButton OptSendCardLen 
            Caption         =   "禁止发卡"
            Height          =   285
            Index           =   0
            Left            =   1500
            TabIndex        =   51
            Top             =   0
            Value           =   -1  'True
            Width           =   1110
         End
         Begin VB.OptionButton OptSendCardLen 
            Caption         =   "不限制"
            Height          =   285
            Index           =   1
            Left            =   -15
            TabIndex        =   50
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton OptSendCardLen 
            Caption         =   "提醒发卡"
            Height          =   285
            Index           =   2
            Left            =   3255
            TabIndex        =   52
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.OptionButton OptSendCard 
         Caption         =   "不限制"
         Height          =   285
         Index           =   0
         Left            =   2025
         TabIndex        =   53
         Top             =   840
         Width           =   960
      End
      Begin VB.OptionButton OptSendCard 
         Caption         =   "同一个病人只发一张卡"
         Height          =   285
         Index           =   1
         Left            =   3550
         TabIndex        =   54
         Top             =   840
         Width           =   2115
      End
      Begin VB.OptionButton OptSendCard 
         Caption         =   "同一个病人发多张卡时提醒"
         Height          =   285
         Index           =   2
         Left            =   2025
         TabIndex        =   55
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Frame Frame3 
         Height          =   60
         Left            =   0
         TabIndex        =   86
         Top             =   585
         Width           =   7875
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "同一病人发卡时控制"
         Height          =   180
         Left            =   150
         TabIndex        =   93
         Top             =   900
         Width           =   1620
      End
      Begin VB.Label lbl卡号限制 
         AutoSize        =   -1  'True
         Caption         =   "发卡卡号长度不足时"
         Height          =   180
         Left            =   150
         TabIndex        =   91
         Top             =   240
         Width           =   1620
      End
   End
   Begin VB.PictureBox picProperty 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1965
      Index           =   1
      Left            =   8100
      ScaleHeight     =   1965
      ScaleWidth      =   7875
      TabIndex        =   76
      Top             =   1305
      Width           =   7875
      Begin VB.Frame fraRule 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         TabIndex        =   88
         Top             =   150
         Width           =   3375
         Begin VB.OptionButton optRule 
            Caption         =   "任意字符"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   56
            Top             =   30
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton optRule 
            Caption         =   "输入字符只能为数字"
            Height          =   180
            Index           =   1
            Left            =   1335
            TabIndex        =   57
            Top             =   30
            Width           =   2070
         End
      End
      Begin VB.Frame Frame4 
         Height          =   30
         Left            =   0
         TabIndex        =   80
         Top             =   1455
         Width           =   7890
      End
      Begin VB.Frame Frame2 
         Height          =   60
         Left            =   -15
         TabIndex        =   79
         Top             =   945
         Width           =   7890
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   450
         Left            =   1410
         TabIndex        =   87
         Top             =   540
         Width           =   5430
         Begin VB.OptionButton optPassConfine 
            Caption         =   "不输入密码禁止"
            Height          =   210
            Index           =   2
            Left            =   3600
            TabIndex        =   61
            Top             =   135
            Value           =   -1  'True
            Width           =   1890
         End
         Begin VB.OptionButton optPassConfine 
            Caption         =   "不输入密码提醒"
            Height          =   210
            Index           =   1
            Left            =   1500
            TabIndex        =   60
            Top             =   150
            Width           =   1890
         End
         Begin VB.OptionButton optPassConfine 
            Caption         =   "不限制"
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   59
            Top             =   135
            Width           =   1200
         End
      End
      Begin VB.TextBox txtPassByIDCard 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1815
         TabIndex        =   69
         Text            =   "0"
         Top             =   1620
         Width           =   300
      End
      Begin VB.CheckBox chkPassByIDCard 
         Caption         =   "缺省以身份证后    位为缺省密码:提示缺省密码位数根据密码长度自动获取"
         Height          =   300
         Left            =   240
         TabIndex        =   66
         Top             =   1605
         Width           =   6900
      End
      Begin VB.Frame fraSplit 
         Height          =   60
         Left            =   0
         TabIndex        =   78
         Top             =   465
         Width           =   7875
      End
      Begin VB.TextBox txtPassInput 
         Height          =   270
         Left            =   6045
         TabIndex        =   65
         Text            =   "0"
         Top             =   1125
         Width           =   300
      End
      Begin VB.OptionButton optPassInput 
         Caption         =   "必须输入    位密码以上"
         Height          =   210
         Index           =   2
         Left            =   5025
         TabIndex        =   64
         Top             =   1155
         Width           =   2295
      End
      Begin VB.OptionButton optPassInput 
         Caption         =   "输入不固定"
         Height          =   210
         Index           =   0
         Left            =   1560
         TabIndex        =   62
         Top             =   1140
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.TextBox txtPasLen 
         Height          =   270
         Left            =   6135
         MaxLength       =   2
         TabIndex        =   58
         Text            =   "10"
         Top             =   120
         Width           =   300
      End
      Begin VB.OptionButton optPassInput 
         Caption         =   "固定输入10位"
         Height          =   210
         Index           =   1
         Left            =   2910
         TabIndex        =   63
         Top             =   1155
         Width           =   1665
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "密码输入规则："
         Height          =   180
         Left            =   240
         TabIndex        =   95
         Top             =   1140
         Width           =   1260
      End
      Begin VB.Label lbl密码输入限制 
         AutoSize        =   -1  'True
         Caption         =   "密码输入限制："
         Height          =   180
         Left            =   255
         TabIndex        =   82
         Top             =   690
         Width           =   1140
      End
      Begin VB.Label lbl密码规则 
         AutoSize        =   -1  'True
         Caption         =   "密码构成规则："
         Height          =   180
         Left            =   255
         TabIndex        =   81
         Top             =   195
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "密码最大长度    位"
         Height          =   180
         Left            =   5040
         TabIndex        =   77
         Top             =   165
         Width           =   1620
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8100
      TabIndex        =   68
      Top             =   870
      Width           =   1100
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   7590
      Index           =   0
      Left            =   135
      TabIndex        =   70
      Top             =   75
      Width           =   7920
      Begin VB.Frame fraboard 
         Caption         =   "自助系统软键盘控制"
         Height          =   735
         Left            =   0
         TabIndex        =   94
         Top             =   4500
         Width           =   5235
         Begin VB.OptionButton opt键盘控制 
            Caption         =   "禁止使用软键盘"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton opt键盘控制 
            Caption         =   "使用数字软键盘"
            Height          =   180
            Index           =   1
            Left            =   1935
            TabIndex        =   31
            Top             =   360
            Width           =   1635
         End
         Begin VB.OptionButton opt键盘控制 
            Caption         =   "使用字符软键盘"
            Height          =   180
            Index           =   2
            Left            =   3615
            TabIndex        =   32
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame FraReadCard 
         Caption         =   "读卡性质"
         Height          =   735
         Left            =   0
         TabIndex        =   90
         Top             =   3660
         Width           =   5235
         Begin VB.CheckBox chk读卡性质 
            Caption         =   "接触式读卡"
            Height          =   180
            Index           =   2
            Left            =   2220
            TabIndex        =   28
            Top             =   360
            Width           =   1350
         End
         Begin VB.CheckBox chk读卡性质 
            Caption         =   "刷卡"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Value           =   1  'Checked
            Width           =   690
         End
         Begin VB.CheckBox chk读卡性质 
            Caption         =   "扫描卡"
            Height          =   180
            Index           =   1
            Left            =   1095
            TabIndex        =   27
            Top             =   360
            Width           =   870
         End
         Begin VB.CheckBox chk读卡性质 
            Caption         =   "非接触式读卡"
            Height          =   180
            Index           =   3
            Left            =   3780
            TabIndex        =   29
            Top             =   360
            Width           =   1410
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "消费刷卡控制"
         Height          =   1995
         Left            =   5325
         TabIndex        =   89
         Top             =   3240
         Width           =   2555
         Begin VB.CheckBox chk缺省退现 
            Caption         =   "缺省退现"
            Height          =   300
            Left            =   1485
            TabIndex        =   44
            Top             =   260
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chk退款验卡 
            Caption         =   "退款时需要验卡(&D)"
            Height          =   180
            Left            =   150
            TabIndex        =   48
            Top             =   1420
            Width           =   2010
         End
         Begin VB.CheckBox chk持卡消费 
            Caption         =   "必须持卡消费(&P)"
            Height          =   195
            Left            =   150
            TabIndex        =   47
            Top             =   1160
            Width           =   1875
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "允许退现(&R)"
            Height          =   300
            Index           =   3
            Left            =   150
            TabIndex        =   43
            Top             =   260
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "部分退款(&S)"
            Height          =   240
            Index           =   4
            Left            =   150
            TabIndex        =   45
            Top             =   580
            Width           =   1515
         End
         Begin VB.CheckBox chk转帐及代扣 
            Caption         =   "支持转帐及代扣(&H)"
            Height          =   195
            Left            =   150
            TabIndex        =   46
            Top             =   880
            Width           =   1875
         End
         Begin VB.CheckBox chk发送调用接口 
            Caption         =   "医嘱发送生成支付条码(&A)"
            Height          =   195
            Left            =   150
            TabIndex        =   49
            Top             =   1700
            Width           =   2370
         End
      End
      Begin VB.PictureBox picExpend 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Left            =   0
         ScaleHeight     =   2355
         ScaleWidth      =   7860
         TabIndex        =   84
         Top             =   5220
         Width           =   7860
         Begin XtremeSuiteControls.TabControl tbPage 
            Height          =   390
            Left            =   0
            TabIndex        =   85
            Top             =   0
            Width           =   270
            _Version        =   589884
            _ExtentX        =   476
            _ExtentY        =   688
            _StockProps     =   64
         End
      End
      Begin VB.CommandButton cmdInsureSel 
         Caption         =   "&P"
         Height          =   270
         Left            =   4920
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   2925
         Width           =   270
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   9
         Left            =   1005
         MaxLength       =   50
         TabIndex        =   20
         Tag             =   "备注"
         Top             =   2910
         Width           =   4200
      End
      Begin VB.ComboBox cbo类别 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   150
         Width           =   1410
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "&P"
         Height          =   270
         Left            =   4920
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   1740
         Width           =   270
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   8
         Left            =   1005
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   14
         Tag             =   "医疗卡费"
         Top             =   1725
         Width           =   4200
      End
      Begin VB.TextBox txt结束位置 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   2970
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "1"
         Top             =   3300
         Width           =   360
      End
      Begin MSComCtl2.UpDown upd开始位置 
         Height          =   300
         Left            =   2205
         TabIndex        =   23
         Top             =   3300
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt开始位置"
         BuddyDispid     =   196653
         OrigLeft        =   1455
         OrigTop         =   2550
         OrigRight       =   1710
         OrigBottom      =   2940
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.TextBox txt开始位置 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "1"
         Top             =   3300
         Width           =   315
      End
      Begin VB.Frame Frame1 
         Caption         =   "医疗卡属性"
         Height          =   3015
         Left            =   5325
         TabIndex        =   71
         Top             =   60
         Width           =   2550
         Begin VB.CheckBox chkOpenEnter 
            Caption         =   "设备启用回车(&0)"
            Height          =   180
            Left            =   150
            TabIndex        =   42
            Top             =   2650
            Width           =   1875
         End
         Begin VB.CheckBox chkCertificate 
            Caption         =   "证    件(&9)"
            Height          =   180
            Left            =   150
            TabIndex        =   41
            Top             =   2400
            Width           =   1320
         End
         Begin VB.CheckBox chkWriteCard 
            Caption         =   "允许写卡(&5)"
            Height          =   180
            Left            =   150
            TabIndex        =   37
            Top             =   1350
            Width           =   1695
         End
         Begin VB.CheckBox chkSendCard 
            Caption         =   "允许发卡(&4)"
            Height          =   180
            Left            =   150
            TabIndex        =   36
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CheckBox chkMakeCard 
            Caption         =   "允许制卡(&3)"
            Height          =   180
            Left            =   150
            TabIndex        =   35
            Top             =   810
            Width           =   1695
         End
         Begin VB.CheckBox chk模糊查找 
            Caption         =   "支持模糊查找(&8)"
            Height          =   180
            Left            =   150
            TabIndex        =   40
            Top             =   2145
            Width           =   1695
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "卡号重复使用(&6)"
            Height          =   180
            Index           =   6
            Left            =   150
            TabIndex        =   38
            Top             =   1605
            Width           =   1665
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "缺省刷卡类别(&7)"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   39
            Top             =   1875
            Width           =   1695
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "存在帐户(&2)"
            Height          =   180
            Index           =   2
            Left            =   150
            TabIndex        =   34
            Top             =   540
            Width           =   1305
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "严格控制(&1)"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   33
            Top             =   285
            Width           =   1320
         End
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   6
         Left            =   1005
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "部件"
         Top             =   2130
         Width           =   4200
      End
      Begin VB.ComboBox cbo结算方式 
         Height          =   300
         Left            =   3675
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1315
         Width           =   1515
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   3675
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "前缀文本"
         Top             =   915
         Width           =   1515
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   1005
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "短名"
         Top             =   915
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   1005
         MaxLength       =   100
         TabIndex        =   4
         Tag             =   "名称"
         Top             =   525
         Width           =   4185
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   7
         Left            =   1005
         MaxLength       =   50
         TabIndex        =   18
         Tag             =   "备注"
         Top             =   2520
         Width           =   4200
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   3675
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "编码"
         Top             =   150
         Width           =   1515
      End
      Begin MSComCtl2.UpDown upd结束位置 
         Height          =   300
         Left            =   3330
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3300
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt结束位置"
         BuddyDispid     =   196651
         OrigLeft        =   1455
         OrigTop         =   2550
         OrigRight       =   1710
         OrigBottom      =   2940
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.CheckBox chk卡号 
         Caption         =   "卡号从        位至        位加密显示(&M)"
         Height          =   180
         Left            =   990
         TabIndex        =   21
         Top             =   3360
         Width           =   3885
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   4
         Left            =   1005
         MaxLength       =   3
         TabIndex        =   10
         Tag             =   "卡号长度"
         Top             =   1315
         Width           =   1395
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "险类(&X)"
         Height          =   180
         Index           =   9
         Left            =   330
         TabIndex        =   19
         Top             =   2970
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "卡类别(&G)"
         Height          =   180
         Left            =   165
         TabIndex        =   74
         Top             =   210
         Width           =   810
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "医疗卡费(&F)"
         Height          =   180
         Index           =   5
         Left            =   -15
         TabIndex        =   13
         Top             =   1785
         Width           =   990
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "部件(M)"
         Height          =   180
         Index           =   8
         Left            =   345
         TabIndex        =   15
         Top             =   2190
         Width           =   630
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "结算方式(&J)"
         Height          =   180
         Index           =   7
         Left            =   2670
         TabIndex        =   11
         Top             =   1375
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "卡号长度(&L)"
         Height          =   180
         Index           =   6
         Left            =   -15
         TabIndex        =   9
         Top             =   1375
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "前缀文本(&T)"
         Height          =   180
         Index           =   3
         Left            =   2670
         TabIndex        =   7
         Top             =   975
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "短名(&D)"
         Height          =   180
         Index           =   1
         Left            =   345
         TabIndex        =   5
         Top             =   975
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   2
         Left            =   345
         TabIndex        =   3
         Top             =   585
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "编码(&B)"
         Height          =   180
         Index           =   0
         Left            =   3030
         TabIndex        =   1
         Top             =   210
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "说明(&E)"
         Height          =   180
         Index           =   4
         Left            =   345
         TabIndex        =   17
         Top             =   2580
         Width           =   630
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   4635
      Index           =   1
      Left            =   90
      TabIndex        =   72
      Top             =   150
      Width           =   7575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8100
      TabIndex        =   67
      Top             =   330
      Width           =   1100
   End
End
Attribute VB_Name = "frmPatiCureCardTypeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------------------
'入口参数
Public Enum gCardTypeEdit
    edT_增加 = 0
    edt_修改 = 1
    edt_删除 = 2
    edt_停用 = 3
    edt_启用 = 4
    dt_查看 = 5
End Enum
Private mlngModule As Long
Private mEditType As gCardTypeEdit
Private mlngCardTypeID As Long
'-----------------------------------------------------------------------------------------
Private mintSucces As Integer
Private mstrPrivs As String
Private mblnFirst As Boolean
Private Enum mtxtIdx
     idx_编码 = 0
     idx_短名 = 1
     idx_名称 = 2
     idx_前缀文本 = 3
     idx_卡号长度 = 4
     idx_部件 = 6
     idx_备注 = 7
     idx_医疗卡费 = 8
     idx_险类 = 9
End Enum
Private Enum mchkIdx
    idx_缺省 = 0
    idx_严格控制 = 1
'    idx_刷卡方式 = 2
    'idx_自制卡 = 3
    idx_存在帐户 = 2
    idx_允许退现 = 3
    idx_部分退款 = 4
    idx_卡号重复使用 = 6
End Enum
Private Enum mlblIdx
   idx_lbl结算方式 = 7
End Enum
'问题号:57326
Private Enum moptIdx
   idx_不限制 = 0
   idx_只发一张卡 = 1
   idx_发多张卡并提醒 = 2
End Enum

Private Enum moptLenIdx
   idx_卡号不足禁止 = 0
   idx_卡号不限制 = 1
   idx_卡号不足提醒 = 2
End Enum

Private Enum constOpt
    禁止 = 0
    数字 = 1
    字符 = 2
End Enum

Private Enum constChk
    刷卡 = 0
    扫描卡 = 1
    接触式读卡 = 2
    非接触式读卡 = 3
End Enum

Private Enum mPageIndex
    读卡设置 = 1
    密码设置 = 2
End Enum

Private mbln固定 As Boolean
Private mblnLoadCard As Boolean

Private Sub SetCtrlEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的编辑属性
    '编制:刘兴洪
    '日期:2011-06-28 03:50:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnEdit As Boolean
    Dim blnModify As Boolean
    Dim bln院内卡 As Boolean
    
    bln院内卡 = cbo类别.Text = "院内卡"
    blnModify = mEditType = edt_修改 And mbln固定
    blnEdit = mEditType = edT_增加 Or (mEditType = edt_修改)
    
    For i = 0 To txtEdit.UBound
        If i <> 5 Then
            txtEdit(i).Enabled = blnEdit And Not mbln固定
            Select Case i
            Case mtxtIdx.idx_卡号长度, mtxtIdx.idx_医疗卡费, mtxtIdx.idx_部件
                txtEdit(i).Enabled = IIf(mEditType = edt_修改, blnEdit Or blnModify, blnEdit)
             
            End Select
            txtEdit(i).BackColor = IIf(txtEdit(i).Enabled, -2147483643, Me.BackColor)
        End If
    Next
    
    chkEdit(mchkIdx.idx_存在帐户).Enabled = blnEdit And Not bln院内卡
    chkEdit(mchkIdx.idx_严格控制).Enabled = blnEdit And bln院内卡
    chkEdit(mchkIdx.idx_卡号重复使用).Enabled = blnEdit And bln院内卡
    chkEdit(mchkIdx.idx_允许退现).Enabled = blnEdit And Not bln院内卡
    chk缺省退现.Enabled = chkEdit(mchkIdx.idx_允许退现).value = 1
    chkEdit(mchkIdx.idx_部分退款).Enabled = blnEdit And Not bln院内卡
'    chkEdit(mchkIdx.idx_刷卡方式).Enabled = blnEdit
    chkEdit(mchkIdx.idx_缺省).Enabled = blnEdit
    chk卡号.Enabled = blnEdit Or blnModify
    '105718:李南春，2017/8/16，卡号密文前后都是0则不显示密文
    txt开始位置.Enabled = (blnEdit Or blnModify) And chk卡号.value = 1
    txt结束位置.Enabled = (blnEdit Or blnModify) And chk卡号.value = 1
    upd结束位置.Enabled = (blnEdit Or blnModify) And chk卡号.value = 1
    upd开始位置.Enabled = (blnEdit Or blnModify) And chk卡号.value = 1
    txt开始位置.BackColor = IIf(upd结束位置.Enabled, -2147483643, Me.BackColor)
    txt结束位置.BackColor = IIf(upd结束位置.Enabled, -2147483643, Me.BackColor)
    
    lblEdit(idx_lbl结算方式).Enabled = blnEdit And Not bln院内卡
    txtPasLen.Enabled = blnEdit Or blnModify
    optPassInput(0).Enabled = blnEdit Or blnModify
    optPassInput(1).Enabled = blnEdit Or blnModify
    optPassInput(2).Enabled = blnEdit Or blnModify
    optRule(0).Enabled = blnEdit Or blnModify
    optRule(1).Enabled = blnEdit Or blnModify
    txtPassInput.Enabled = blnEdit Or blnModify
    cbo类别.Enabled = Not mbln固定 And blnEdit
    chk模糊查找.Enabled = blnEdit Or blnModify '47522
    '问题号;56508
    chkMakeCard.Enabled = (chk读卡性质(接触式读卡).value = 1 Or chk读卡性质(非接触式读卡).value = 1) And blnEdit
    chkSendCard.Enabled = Not bln院内卡
    chkOpenEnter.Enabled = (chk读卡性质(刷卡).value = 1 Or chk读卡性质(扫描卡).value = 1) And blnEdit
    opt键盘控制(禁止).Enabled = (chk读卡性质(刷卡).value = 1 Or chk读卡性质(扫描卡).value = 1) And blnEdit
    opt键盘控制(数字).Enabled = (chk读卡性质(刷卡).value = 1 Or chk读卡性质(扫描卡).value = 1) And blnEdit
    opt键盘控制(字符).Enabled = (chk读卡性质(刷卡).value = 1 Or chk读卡性质(扫描卡).value = 1) And blnEdit
    
    txtPasLen.BackColor = IIf(txtPasLen.Enabled, -2147483643, Me.BackColor)
    txtPassInput.BackColor = IIf(txtPassInput.Enabled, -2147483643, Me.BackColor)
    cbo结算方式.BackColor = IIf(cbo结算方式.Enabled, -2147483643, Me.BackColor)
    cbo类别.BackColor = IIf(cbo类别.Enabled, -2147483643, Me.BackColor)
    cmdSel.Enabled = blnEdit Or blnModify
    chk转帐及代扣.Enabled = chkEdit(mchkIdx.idx_存在帐户).value = 1 And chkEdit(mchkIdx.idx_存在帐户).Enabled
    '90875:李南春,2016/11/8,增加医疗卡证件类型,不可编辑
    chkCertificate.Enabled = False
    '104238:李南春，2017/2/15，医疗卡类别增加发卡卡号控制
    OptSendCardLen(0).Enabled = blnEdit Or blnModify
    OptSendCardLen(1).Enabled = blnEdit Or blnModify
    OptSendCardLen(2).Enabled = blnEdit Or blnModify
 End Sub
 Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据的有效性
    '返回:数据有效，返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-28 03:58:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    For i = 0 To txtEdit.UBound
        If i <> 5 Then
            If i <> mtxtIdx.idx_医疗卡费 Then
                If i = mtxtIdx.idx_编码 Or i = mtxtIdx.idx_名称 Then
                    If Trim(txtEdit(i).Text) = "" Then
                        MsgBox txtEdit(i).Tag & " 必须输入,请检查", vbOKOnly + vbInformation, gstrSysName
                        If txtEdit(i).Enabled And txtEdit(i).Visible Then txtEdit(i).SetFocus
                        Exit Function
                    End If
                End If
                If zlCommFun.ActualLen(Trim(txtEdit(i).Text)) > txtEdit(i).MaxLength And txtEdit(i).MaxLength <> 0 Then
                    MsgBox txtEdit(i).Tag & " 最多能输入" & txtEdit(i).MaxLength \ 2 & "个汉字或" & txtEdit(i).MaxLength & "个字符,请检查", vbOKOnly + vbInformation, gstrSysName
                    If txtEdit(i).Enabled And txtEdit(i).Visible Then txtEdit(i).SetFocus
                    Exit Function
                End If
                If InStr(1, Trim(txtEdit(i).Text), "'") > 0 Then
                    MsgBox txtEdit(i).Tag & " 不能输入单引号,请检查", vbOKOnly + vbInformation, gstrSysName
                    If txtEdit(i).Enabled And txtEdit(i).Visible Then txtEdit(i).SetFocus
                    Exit Function
                End If
            End If
        End If
    Next
    If cbo类别.Text <> "院内卡" Then
        '三方卡
        If Trim(cbo结算方式.Text) = "" And chkEdit(mchkIdx.idx_存在帐户).value = 1 Then
            MsgBox "注意:" & vbCrLf & "    如果是院外卡且存在帐户的,必须设置结算方式!", vbInformation + vbOKOnly, gstrSysName
            If cbo结算方式.Enabled And cbo结算方式.Visible Then cbo结算方式.SetFocus
            Exit Function
        End If
        '99858:李南春,2016/9/2,三方账户必须设置接口部件
        If Trim(txtEdit(mtxtIdx.idx_部件).Text) = "" And chkEdit(mchkIdx.idx_存在帐户).value = 1 Then
            MsgBox "注意:" & vbCrLf & "    如果是院外卡且存在帐户的,必须设置接口部件!", vbInformation + vbOKOnly, gstrSysName
            If txtEdit(mtxtIdx.idx_部件).Enabled And txtEdit(mtxtIdx.idx_部件).Visible Then txtEdit(mtxtIdx.idx_部件).SetFocus
            Exit Function
        End If
     Else
        '问题:48090
        If Trim(txtEdit(mtxtIdx.idx_医疗卡费).Text) = "" Then
           MsgBox "注意:" & vbCrLf & "    如果是院内卡,必须设置医疗卡费!", vbInformation + vbOKOnly, gstrSysName
           txtEdit(mtxtIdx.idx_医疗卡费).SetFocus
           Exit Function
        End If
    End If
    
    If Val(txtPasLen.Text) = 0 Then
        MsgBox "注意:" & vbCrLf & "    密码长度不能设置为零!", vbInformation + vbOKOnly, gstrSysName
        If txtPasLen.Enabled And txtPasLen.Visible Then txtPasLen.SetFocus
        Exit Function
    End If
    If Val(txtPasLen.Text) > 50 Then
        MsgBox "注意:" & vbCrLf & "    密码长度不能大于50!", vbInformation + vbOKOnly, gstrSysName
        If txtPasLen.Enabled And txtPasLen.Visible Then txtPasLen.SetFocus
        Exit Function
    End If
    If optPassInput(2).value Then
        If Val(txtPasLen.Text) < Val(txtPassInput.Text) Then
            MsgBox "注意:" & vbCrLf & "    必须输入的密码长度不能大于总的密码长度(" & Val(txtPasLen.Text) & ")位!", vbInformation + vbOKOnly, gstrSysName
            If txtPassInput.Enabled And txtPassInput.Visible Then txtPassInput.SetFocus
            Exit Function
        End If
    End If
    '问题:46851
    If Val(txtEdit(mtxtIdx.idx_卡号长度).Text) > 50 Then
            MsgBox "注意:" & vbCrLf & "    卡号最长只能设置50位!", vbInformation + vbOKOnly, gstrSysName
            If txtEdit(mtxtIdx.idx_卡号长度).Enabled And txtEdit(mtxtIdx.idx_卡号长度).Visible Then txtEdit(mtxtIdx.idx_卡号长度).SetFocus
            Exit Function
    End If
    '82412:李南春,2015/01/30,结算方式重复性检查
    If Replace(cbo结算方式.Text, Chr(32), "") = "" Then isValied = True: Exit Function
    strSQL = "" & _
            " Select 名称 from 医疗卡类别 where not ID =[1] and 结算方式=[2]" & _
            " Union All " & _
            " Select 名称 from 消费卡类别目录 where Not 编号=[1] And 结算方式=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCardTypeID, cbo结算方式.Text)
    If Not rsTemp.EOF Then
        MsgBox "注意:" & vbCrLf & "    结算方式『" & cbo结算方式.Text & "』已被" & NVL(rsTemp!名称) & "使用" & vbCrLf & "    重复使用会造成财务扎帐紊乱，请重新选定一种结算方式", vbInformation + vbOKOnly, gstrSysName
        If cbo结算方式.Visible And cbo结算方式.Enabled Then cbo结算方式.SetFocus
        Exit Function
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 Private Function CheckDepent() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据的关联性
    '返回:数据存在关联，返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-28 03:43:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    '82412:李南春,2015/01/30,医疗卡结算方式调整
    strSQL = "Select 名称 From 结算方式 Where 性质 =8 and nvl(应付款,0)=0 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With cbo结算方式
        .Clear
        .AddItem ""
        .ListIndex = .NewIndex
        Do While Not rsTemp.EOF
            .AddItem NVL(rsTemp!名称)
            rsTemp.MoveNext
        Loop
    End With
    CheckDepent = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
Private Function LoadCardData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载卡片数据
    '返回:加载成功，返回true，否则返回False
    '编制:刘兴洪
    '日期:2011-06-28 02:57:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, i As Long, varTemp As Variant
    Dim strValue As String
    On Error GoTo errHandle
    Call ClearCtrlData
    
    mblnLoadCard = True
    mbln固定 = False
    
    If mEditType = edT_增加 Then
        txtEdit(mtxtIdx.idx_编码).Text = zlDatabase.GetMax("医疗卡类别", "编码", txtEdit(mtxtIdx.idx_编码).MaxLength)
        If txtEdit(mtxtIdx.idx_名称).Enabled And txtEdit(mtxtIdx.idx_名称).Visible Then txtEdit(mtxtIdx.idx_名称).SetFocus
        '问题号:50172
        txtPassByIDCard.Text = txtPasLen.Text
        '问题号;56508
        chkSendCard.value = IIf(chkSendCard.Enabled, 0, 1)
        '问题号:57326
        OptSendCard(moptIdx.idx_不限制).value = 1
        
        OptSendCardLen(moptLenIdx.idx_卡号不足禁止).value = 1
        '106838:李南春，2017/4/11，更新加载完成标志
        mblnLoadCard = False
        LoadCardData = True
        Exit Function
    End If
    '问题号:57326
    '问题号:57697
    '问题号:51072
    '问题号:56508
    '77872,李南春,2014/10/28:是否支持转帐及代扣
    '85565:李南春,2015/7/8,读卡性质以及键盘控制
    '90875:李南春,2016/11/8,增加医疗卡证件类型
    strSQL = "" & _
    "   Select A.Id, A.名称, A.编码, A.短名, A.前缀文本, A.卡号长度,  nvl(A.缺省标志,0) as 缺省标志,  " & _
    "            nvl(A.是否固定,0) as 是否固定,  nvl(A.是否严格控制,0)  as  是否严格控制, " & _
    "            nvl(A.是否自制,0)  as    是否自制," & _
    "            nvl(A.是否存在帐户,0) as   是否存在帐户,  nvl(A.是否退现,0)  as    是否退现, " & _
    "           nvl(A.是否全退,0)  as    是否全退," & _
    "           A.部件,A.特定项目, A.结算方式,A.卡号密文,nvl(A.是否重复使用,0)  as 是否重复使用,  " & _
    "           nvl(A.密码长度,10) as 密码长度,nvl(密码长度限制,0) as 密码长度限制,nvl(密码规则,0) as 密码规则," & _
    "           nvl(A.是否启用,0)  as 是否启用, A.备注,C.名称 as 卡费,C.ID as 细目ID,nvl(是否模糊查找,0) as 是否模糊查找," & _
    "           nvl(A.密码输入限制,0) as 密码输入限制,nvl(A.是否缺省密码,0) as 是否缺省密码,nvl(A.是否制卡,0) as 是否制卡,nvl(A.是否发卡,0) as 是否发卡,nvl(A.是否写卡,0) as 是否写卡, " & _
    "           nvl(A.险类,0) as 险类,nvl(A.发卡性质,0) as 发卡性质, " & _
    "           nvl(A.是否转帐及代扣,0) as 是否转帐及代扣,nvl(A.是否证件,0) as 是否证件, " & _
    "           A.读卡性质, nvl(A.键盘控制方式,0) as 键盘控制方式, " & _
    "           nvl(A.是否持卡消费,0) as 是否持卡消费,nvl(A.发送调用接口,0) as 发送调用接口, " & _
    "           Nvl(a.是否退款验卡,0) As 是否退款验卡," & _
    "           A.设备是否启用回车 as 启用回车,nvl(A.发卡控制,0) as 发卡控制, " & _
    "           Nvl(a.是否缺省退现,0) As 是否缺省退现" & _
    "    From 医疗卡类别 A,收费特定项目 B,收费项目目录 C" & _
    "    Where  A.ID=[1]  And A.特定项目=B.特定项目(+) and B.收费细目ID=C.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCardTypeID)
    If rsTemp.EOF Then
        MsgBox "未找到医疗卡类别信息，可能已经被他人删除！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    With cbo结算方式
        For i = 0 To .ListCount - 1
            If Trim(.List(i)) = Trim(rsTemp!结算方式) Then
                .ListIndex = i: i = -1: Exit For
            End If
        Next
        If i >= 0 And NVL(rsTemp!结算方式) <> "" Then
            .AddItem NVL(rsTemp!结算方式): .ListIndex = .NewIndex
        End If
    End With
    
    txtEdit(mtxtIdx.idx_编码).Text = NVL(rsTemp!编码)
    txtEdit(mtxtIdx.idx_名称).Text = NVL(rsTemp!名称)
    txtEdit(mtxtIdx.idx_短名).Text = NVL(rsTemp!短名)
    upd开始位置.Max = IIf(Val(NVL(rsTemp!卡号长度)) = 0, 1, Val(NVL(rsTemp!卡号长度)))
    upd结束位置.Max = upd开始位置.Max
    txtEdit(mtxtIdx.idx_卡号长度).Text = IIf(Val(NVL(rsTemp!卡号长度)) = 0, 1, Val(NVL(rsTemp!卡号长度)))
    txtEdit(mtxtIdx.idx_前缀文本) = NVL(rsTemp!前缀文本)
    txtEdit(mtxtIdx.idx_备注) = NVL(rsTemp!备注)
    txtEdit(mtxtIdx.idx_部件) = NVL(rsTemp!部件)
    'txtEdit(mtxtIdx.idx_特定项目) = Nvl(rsTemp!特定项目)
    txtEdit(mtxtIdx.idx_医疗卡费) = NVL(rsTemp!卡费)
    txtEdit(mtxtIdx.idx_医疗卡费).Tag = Val(NVL(rsTemp!细目ID))
    varTemp = Split(NVL(rsTemp!卡号密文) & "-", "-")
    If Val(varTemp(0)) = 0 Or Val(varTemp(1)) = 0 Then
        upd开始位置.value = IIf(Val(varTemp(0)) = 0, IIf(Val(varTemp(1)) = 0, 1, Val(varTemp(1))), Val(varTemp(0)))
        upd结束位置.value = upd结束位置.Max
        chk卡号.value = IIf(Val(varTemp(0)) = 0 And Val(varTemp(1)) = 0, 0, 1)
    Else
        upd开始位置.value = Val(varTemp(0))
        upd结束位置.value = Val(varTemp(1))
        chk卡号.value = 1
    End If
    chkEdit(mchkIdx.idx_严格控制).value = IIf(Val(NVL(rsTemp!是否严格控制)) = 1, 1, 0)
'    chkEdit(mchkIdx.idx_刷卡方式).value = IIf(Val(Nvl(rsTemp!是否刷卡)) = 1, 1, 0)
    'chkEdit(mchkIdx.idx_自制卡).Value = IIf(Val(Nvl(rsTemp!是否自制)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_部分退款).value = IIf(Val(NVL(rsTemp!是否全退)) = 1, 0, 1)
    chkEdit(mchkIdx.idx_存在帐户).value = IIf(Val(NVL(rsTemp!是否存在帐户)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_卡号重复使用).value = IIf(Val(NVL(rsTemp!是否重复使用)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_缺省).value = IIf(Val(NVL(rsTemp!缺省标志)) = 1, 1, 0)
    chk模糊查找.value = IIf(Val(NVL(rsTemp!是否模糊查找)) = 1, 1, 0)
    '问题号;56508
    chkMakeCard.value = IIf(Val(NVL(rsTemp!是否制卡)) = 1, 1, 0)
    chkSendCard.value = IIf(Val(NVL(rsTemp!是否发卡)) = 1, 1, 0)
    chkWriteCard.value = IIf(Val(NVL(rsTemp!是否写卡)) = 1, 1, 0)
    
    chkEdit(mchkIdx.idx_允许退现).value = IIf(Val(NVL(rsTemp!是否退现)) = 1, 1, 0)
    If chkEdit(mchkIdx.idx_允许退现).value = 1 Then
        chk缺省退现.value = IIf(Val(NVL(rsTemp!是否缺省退现)) = 1, 1, 0)
        chk缺省退现.Enabled = True
    Else
        chk缺省退现.value = 0
        chk缺省退现.Enabled = False
    End If
    txtPasLen.Text = Val(NVL(rsTemp!密码长度))
    For i = 0 To cbo类别.ListCount - 1
        If cbo类别.List(i) = IIf(Val(NVL(rsTemp!是否自制)) = 0, "院外卡", "院内卡") Then
            cbo类别.ListIndex = i: Exit For
        End If
    Next

    Select Case Val(NVL(rsTemp!密码长度限制))
    Case 0
            optPassInput(0).value = True
    Case 1
            optPassInput(1).value = True
    Case Else
            optPassInput(2).value = True
            txtPassInput.Text = Abs(Val(NVL(rsTemp!密码长度限制)))
    End Select
     optRule(0).value = IIf(Val(NVL(rsTemp!密码规则)) = 0, True, False)
     optRule(1).value = IIf(Val(NVL(rsTemp!密码规则)) = 1, True, False)
    '问题号:51072
    Select Case Val(NVL(rsTemp!密码输入限制))
    Case 0
            optPassConfine(0).value = True
    Case 1
            optPassConfine(1).value = True
    Case Else
            optPassConfine(2).value = True
    End Select
    '问题号:50172
    chkPassByIDCard.value = rsTemp!是否缺省密码
    txtPassByIDCard.Text = txtPasLen.Text
    
    If Val(NVL(rsTemp!是否固定)) = 1 Then
        '固定，只能查看
        mbln固定 = True
    End If
    
    '问题号:57697
    txtEdit(mtxtIdx.idx_险类).Tag = NVL(rsTemp!险类, 0)
    txtEdit(mtxtIdx.idx_险类).Text = Get险类名称(CStr(txtEdit(mtxtIdx.idx_险类).Tag))
    
    '问题号:57326
    OptSendCard(Val(NVL(rsTemp!发卡性质))).value = 1
    
    '77872,李南春,2014/9/15:是否支持转帐及代扣
    chk转帐及代扣.Enabled = chkEdit(mchkIdx.idx_存在帐户).value = 1
    If chk转帐及代扣.Enabled Then chk转帐及代扣.value = IIf(Val(NVL(rsTemp!是否转帐及代扣)) = 1, 1, 0)
    chk持卡消费.value = IIf(Val(NVL(rsTemp!是否持卡消费)) = 1, 1, 0)
    chk发送调用接口.value = IIf(Val(NVL(rsTemp!发送调用接口)) = 1, 1, 0)
    chk退款验卡.value = IIf(Val(NVL(rsTemp!是否退款验卡)) = 1, 1, 0)
    
    strValue = NVL(rsTemp!读卡性质, "1000")
    chk读卡性质(刷卡).value = Mid(strValue, 1, 1)
    chk读卡性质(扫描卡).value = Mid(strValue, 2, 1)
    chk读卡性质(接触式读卡).value = Mid(strValue, 3, 1)
    chk读卡性质(非接触式读卡).value = Mid(strValue, 4, 1)
    
    Select Case Val(NVL(rsTemp!键盘控制方式))
    Case 0
            opt键盘控制(禁止).value = True
    Case 1
            opt键盘控制(数字).value = True
    Case Else
            opt键盘控制(字符).value = True
    End Select
    
    '90875:李南春,2016/11/8,增加医疗卡证件类型
    chkCertificate.value = IIf(Val(NVL(rsTemp!是否证件)) = 1, 1, 0)
    
    '103310:李南春,2016/12/7,启用回车后增加卡号长度
    chkOpenEnter.Enabled = chk读卡性质(刷卡).value = 1 Or chk读卡性质(扫描卡).value = 1
    chkOpenEnter.value = IIf(Val(NVL(rsTemp!启用回车)) = 1, 1, 0)
    
    '104238:李南春，2017/2/15，医疗卡类别增加发卡卡号控制
    OptSendCardLen(Val(NVL(rsTemp!发卡控制))).value = 1
    
    If mEditType = dt_查看 Then
        cmdOK.Visible = False
    End If
    mblnLoadCard = False
    LoadCardData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function Get险类名称(str序号 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取险类对应名称
    '编制:王吉
    '日期:2013-01-29 02:54:36
    '问题号:57697
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    On Error GoTo Errhand:
        strSQL = "Select 名称 From 保险类别 Where 序号=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str序号)
        If rsTemp.EOF = False Then
            Get险类名称 = rsTemp!名称
        End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ClearCtrlData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除控件数据
    '编制:刘兴洪
    '日期:2011-06-28 02:54:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To txtEdit.UBound
        If i <> 5 Then '5-暂没有
            txtEdit(i).Text = ""
        End If
    Next
    For i = 0 To chkEdit.UBound
        If i <> 5 Then
            chkEdit(i).value = 0
        End If
    Next
    chkEdit(idx_允许退现).value = IIf(chkEdit(idx_允许退现).Enabled, 1, 0)
    For i = 0 To chk读卡性质.UBound
        If i <> 3 Then
            chk读卡性质(i).value = 0
        End If
    Next
    cbo结算方式.ListIndex = 0
    chk卡号.value = 0
    chk转帐及代扣.value = 0
    chkCertificate.value = 0
End Sub

Private Sub SetDefaultLen()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省的编辑长度
    '编制:刘兴洪
    '日期:2011-06-28 02:50:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select  名称, 编码, 短名, 前缀文本, 卡号长度 ,部件,特定项目,结算方式,备注" & _
    "    From 医疗卡类别" & _
    "    Where ID=-1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    txtEdit(mtxtIdx.idx_编码).MaxLength = rsTemp.Fields("编码").DefinedSize
    txtEdit(mtxtIdx.idx_名称).MaxLength = rsTemp.Fields("名称").DefinedSize
    txtEdit(mtxtIdx.idx_短名).MaxLength = rsTemp.Fields("短名").DefinedSize
    txtEdit(mtxtIdx.idx_前缀文本).MaxLength = 2 '  rsTemp.Fields("前缀文本").DefinedSize
    txtEdit(mtxtIdx.idx_部件).MaxLength = rsTemp.Fields("部件").DefinedSize
    txtEdit(mtxtIdx.idx_备注).MaxLength = rsTemp.Fields("备注").DefinedSize
    txtEdit(mtxtIdx.idx_卡号长度).MaxLength = 2
   ' txtEdit(mtxtIdx.idx_特定项目).MaxLength = rsTemp.Fields("特定项目").DefinedSize
   

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Function zlEditCard(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal EditType As gCardTypeEdit, Optional lngCardTypeID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医疗卡类别编辑
    '入参:EditType-编辑类型
    '        lngCardTypeID-增加时为0
    '出参:
    '返回:只要成功一次,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-06-27 20:43:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mEditType = EditType: mlngModule = lngModule: mlngCardTypeID = lngCardTypeID
    mintSucces = 0: mstrPrivs = strPrivs
    Me.Show 1, frmMain
    zlEditCard = mintSucces > 0
End Function

Private Sub cbo类别_Click()
    Call SetCtrlEnable
End Sub

Private Sub chkEdit_Click(Index As Integer)
'修改日期:2012-11-29
'问题号:56508
'85565
    Select Case Index
'        Case 2 '磁卡
'            If chkEdit(Index).value = 1 Then
'                chkMakeCard.value = 0
'                chkMakeCard.Enabled = False
'            Else
'                chkMakeCard.Enabled = True
'            End If
'        Case 3 '是否退现
'            If chkEdit(Index).value = 1 Then
'                chkEdit(idx_部分退款).value = 0
'                chkEdit(idx_部分退款).Enabled = False
'            Else
'                chkEdit(idx_部分退款).Enabled = True
'            End If
        Case 2 '存在帐户
            If chkEdit(Index).value = 0 Then
                chk转帐及代扣.value = 0
                chk转帐及代扣.Enabled = False
            Else
                chk转帐及代扣.Enabled = True
            End If
        Case 3
            If chkEdit(Index).value = 0 Then
                chk缺省退现.value = 0
                chk缺省退现.Enabled = False
            Else
                chk缺省退现.Enabled = True
            End If
    End Select
End Sub

Private Sub chk读卡性质_Click(Index As Integer)
    '至少保留一项读卡方式
    Dim i As Integer, blnCancel As Boolean
    If mblnLoadCard Then Exit Sub
    For i = 0 To chk读卡性质.UBound
        If chk读卡性质(i).value = 1 Then
            blnCancel = True: Exit For
        End If
    Next
    
    If Not blnCancel Then
        chk读卡性质(Index).value = 1
    End If
    
    If chk读卡性质(刷卡).value <> 1 And chk读卡性质(扫描卡).value <> 1 Then
        opt键盘控制(禁止).value = True
        opt键盘控制(禁止).Enabled = False: opt键盘控制(数字).Enabled = False: opt键盘控制(字符).Enabled = False
        
        chkOpenEnter.value = 0
        chkOpenEnter.Enabled = False
    Else
        opt键盘控制(禁止).Enabled = True: opt键盘控制(数字).Enabled = True: opt键盘控制(字符).Enabled = True
        chkOpenEnter.Enabled = True
    End If
    
    If chk读卡性质(接触式读卡).value = 1 Or chk读卡性质(非接触式读卡).value = 1 Then
        chkMakeCard.Enabled = True
    Else
        chkMakeCard.value = 0
        chkMakeCard.Enabled = False
    End If
End Sub

''Private Sub chkEdit_Click(Index As Integer)
''    If Index = mchkIdx.idx_自制卡 Then
''        chkEdit(mchkIdx.idx_存在帐户).Enabled = chkEdit(mchkIdx.idx_自制卡).Value = 0
''        chkEdit(mchkIdx.idx_允许退现).Enabled = chkEdit(mchkIdx.idx_存在帐户).Enabled
''    End If
''End Sub

Private Sub chk卡号_Click()
    Dim blnEnable As Boolean
    blnEnable = chk卡号.Enabled And chk卡号.value = 1
    txt开始位置.Enabled = blnEnable
    txt结束位置.Enabled = blnEnable
    upd结束位置.Enabled = blnEnable
    upd开始位置.Enabled = blnEnable
    '105718:李南春，2017/8/16，卡号密文前后都是0则不显示密文
    txt开始位置.BackColor = IIf(blnEnable, -2147483643, Me.BackColor)
    txt结束位置.BackColor = IIf(blnEnable, -2147483643, Me.BackColor)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInsureSel_Click()
    '问题号:57697
     If Select险类 = False Then Exit Sub
End Sub

Private Sub cmdOK_Click()
    If isValied = False Then Exit Sub
    If SaveData = False Then Exit Sub
    mintSucces = mintSucces + 1
    If mEditType = edT_增加 Then
        Call LoadCardData: Exit Sub
    End If
    Unload Me
End Sub
Private Function Select险类() As Boolean
    Dim blnCancel As Boolean, lngH As Long
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rsTemp  As ADODB.Recordset, strSQL As String
    '问题号:57697
    On Error GoTo ErrHandl:
    strSQL = "Select 序号 as Id,名称,说明,医院编码,是否固定,是否禁止,具有中心,医保部件,外挂,项目提示,医保包 From 保险类别"
    vRect = zlControl.GetControlRect(txtEdit(mtxtIdx.idx_险类).hWnd)
    lngH = txtEdit(mtxtIdx.idx_险类).Height
    sngX = vRect.Left - 15
    sngY = vRect.Top
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "医疗保险类别", False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False)
    If blnCancel = False Then
            txtEdit(mtxtIdx.idx_险类).Text = NVL(rsTemp!名称, "")
            txtEdit(mtxtIdx.idx_险类).Tag = NVL(rsTemp!id, "")
    End If
    Select险类 = Not blnCancel
    
    Exit Function
ErrHandl:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Select卡费(ByVal strInput As String) As Boolean
    Dim blnCancel As Boolean, lngH As Long
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rsTemp  As ADODB.Recordset, strSQL As String
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
 
    strSQL = "select id,编码,名称,计算单位,说明 from 收费项目目录 where 类别='Z' Order by 编码"
    vRect = zlControl.GetControlRect(txtEdit(mtxtIdx.idx_医疗卡费).hWnd)
    lngH = txtEdit(mtxtIdx.idx_医疗卡费).Height
    sngX = vRect.Left - 15
    sngY = vRect.Top
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "医疗卡卡费项目选择", False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strInput)
    If blnCancel = True Then
        If txtEdit(mtxtIdx.idx_医疗卡费).Enabled Then txtEdit(mtxtIdx.idx_医疗卡费).SetFocus
        zlControl.TxtSelAll txtEdit(mtxtIdx.idx_医疗卡费)
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgbox "没有找到满足条件的卡费项目,请检查!"
        If txtEdit(mtxtIdx.idx_医疗卡费).Enabled Then txtEdit(mtxtIdx.idx_医疗卡费).SetFocus
        If UCase(TypeName(txtEdit(mtxtIdx.idx_医疗卡费))) = UCase("TextBox") Then zlControl.TxtSelAll txtEdit(mtxtIdx.idx_医疗卡费)
        Exit Function
    End If
    If IsCtrlSetFocus(txtEdit(mtxtIdx.idx_医疗卡费)) Then txtEdit(mtxtIdx.idx_医疗卡费).SetFocus
    txtEdit(mtxtIdx.idx_医疗卡费).Text = NVL(rsTemp!名称)
    txtEdit(mtxtIdx.idx_医疗卡费).Tag = NVL(rsTemp!id)
    zlCommFun.PressKey vbKeyTab
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub cmdSel_Click()
    If Select卡费("") = False Then Exit Sub
    
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If CheckDepent = False Then Unload Me: Exit Sub
    If LoadCardData = False Then Unload Me: Exit Sub
    Call InitPage
    Call SetCtrlEnable
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnFirst = True
    With cbo类别
        .Clear
        .AddItem "院内卡": .ListIndex = .NewIndex
        .AddItem "院外卡"
    End With
    If mEditType = edT_增加 Then chk持卡消费.value = 1
    Call SetDefaultLen
End Sub

Private Sub optPassInput_Click(Index As Integer)
    txtPassInput.Enabled = optPassInput(2).value
    txtPassInput.BackColor = IIf(txtPassInput.Enabled, -2147483643, Me.BackColor)
End Sub

Private Sub picExpend_Resize()
    Err = 0: On Error Resume Next
    With picExpend
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If Index = mtxtIdx.idx_卡号长度 Then
        upd结束位置.Max = Val(txtEdit(Index))
        upd开始位置.Max = Val(txtEdit(Index))
        If upd结束位置.value > Val(txtEdit(Index)) Then upd结束位置.value = Val(txtEdit(Index))
        If upd开始位置.value > Val(txtEdit(Index)) Then upd开始位置.value = Val(txtEdit(Index))
    End If
    If Index = mtxtIdx.idx_名称 Then
        If Trim(txtEdit(mtxtIdx.idx_短名)) = "" And txtEdit(Index).Text <> "" Then txtEdit(mtxtIdx.idx_短名) = Left(txtEdit(Index), 1)
    End If
    If Index = mtxtIdx.idx_医疗卡费 Then
        txtEdit(Index).Tag = ""
    End If
    '问题号:57697
    If Index = mtxtIdx.idx_险类 Then
        If txtEdit(Index).Text = "" Then
            txtEdit(Index).Tag = ""
        End If
    End If
End Sub
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存数据
    '返回:保存成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-28 04:13:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lngID As Long
    Dim strValue As String
    If mEditType = edT_增加 Then
        lngID = zlDatabase.GetNextId("医疗卡类别")
    Else
        lngID = mlngCardTypeID
    End If
    
    On Error GoTo errHandle
    ' Zl_医疗卡类别_Update
    strSQL = "Zl_医疗卡类别_Update("
    '  Id_In           In 医疗卡类别.ID%Type,
    strSQL = strSQL & "" & lngID & ","
    '  编码_In         In 医疗卡类别.编码%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_编码).Text) & "',"
    '  名称_In         In 医疗卡类别.名称%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_名称).Text) & "',"
    '  短名_In         In 医疗卡类别.短名%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_短名).Text) & "',"
    '  前缀文本_In     In 医疗卡类别.前缀文本%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_前缀文本).Text) & "',"
    '  卡号长度_In     In 医疗卡类别.卡号长度%Type,
    strSQL = strSQL & "" & Val(txtEdit(mtxtIdx.idx_卡号长度).Text) & ","
    '  缺省标志_In     In 医疗卡类别.缺省标志%Type,
    strSQL = strSQL & "" & chkEdit(mchkIdx.idx_缺省).value & ","
    '  是否固定_In     In 医疗卡类别.是否固定%Type,
    strSQL = strSQL & "0,"
    '  是否严格控制_In In 医疗卡类别.是否严格控制%Type,
    strSQL = strSQL & "" & IIf(chkEdit(mchkIdx.idx_严格控制).value = 1 And chkEdit(mchkIdx.idx_严格控制).Enabled, 1, 0) & ","
    '  是否自制_In     In 医疗卡类别.是否自制%Type,
    strSQL = strSQL & "" & IIf(cbo类别.Text = "院内卡", 1, 0) & ","
    '  是否存在帐户_In In 医疗卡类别.是否存在帐户%Type,
    strSQL = strSQL & "" & IIf(chkEdit(mchkIdx.idx_存在帐户).Enabled And chkEdit(mchkIdx.idx_存在帐户).value = 1, 1, 0) & ","
    '  是否全退_In     In 医疗卡类别.是否全退%Type,
    strSQL = strSQL & "" & IIf(chkEdit(mchkIdx.idx_部分退款).Enabled And chkEdit(mchkIdx.idx_部分退款).value = 1, 0, 1) & ","
    '  部件_In         In 医疗卡类别.部件%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_部件).Text) & "',"
    '  备注_In         In 医疗卡类别.备注%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_备注).Text) & "',"
    '  特定项目_In     In 医疗卡类别.特定项目%Type,
    If Trim(txtEdit(mtxtIdx.idx_名称).Text) = "就诊卡" Then
        strSQL = strSQL & "'就诊卡',"
    Else
        strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_编码).Text) & "',"
    End If
    '    收费细目id_In   In 收费项目目录.ID%Type,
    strSQL = strSQL & "" & IIf(Val(txtEdit(mtxtIdx.idx_医疗卡费).Tag) = 0, "NULL", Val(txtEdit(mtxtIdx.idx_医疗卡费).Tag)) & ","
    '  结算方式_In     In 医疗卡类别.结算方式%Type,
    strSQL = strSQL & "'" & Trim(cbo结算方式.Text) & "',"
    '  是否启用_In     In 医疗卡类别.是否启用%Type,
    strSQL = strSQL & "1,"
    '  卡号密文_In     In 医疗卡类别.卡号密文%Type,
    strSQL = strSQL & "" & IIf(chk卡号.value = 1, "'" & upd开始位置.value & "-" & upd结束位置.value & "'", "NULL") & ","
    '  是否重复使用_In In 医疗卡类别.是否重复使用%Type,
    strSQL = strSQL & "" & IIf(chkEdit(mchkIdx.idx_卡号重复使用).Enabled And chkEdit(mchkIdx.idx_卡号重复使用).value = 1, 1, 0) & ","
    '密码长度_In     In 医疗卡类别.密码长度%Type,
    strSQL = strSQL & "" & Val(txtPasLen.Text) & ","
    '密码长度限制_In In 医疗卡类别.密码长度限制%Type,
    If optPassInput(0).value Then
        strSQL = strSQL & "" & 0 & ","
    ElseIf optPassInput(1).value Then
        strSQL = strSQL & "" & 1 & ","
    Else
        strSQL = strSQL & "" & -1 * Val(txtPassInput.Text) & ","
    End If
    '密码规则_In     In 医疗卡类别.密码规则%Type,
    If optRule(0).value Then
        strSQL = strSQL & "" & 0 & ","
    Else
        strSQL = strSQL & "" & 1 & ","
    End If
    strSQL = strSQL & "" & IIf(chkEdit(mchkIdx.idx_允许退现).Enabled And chkEdit(mchkIdx.idx_允许退现).value = 1, 1, 0) & ","
    '  操作方式_In     In Integer := 0
    strSQL = strSQL & "" & IIf(mEditType = edT_增加, 0, 1) & ","
    '是否模糊查找_In     In 医疗卡类别.是否模糊查找%Type:=0
    strSQL = strSQL & "" & IIf(chk模糊查找.value = 1, 1, 0) & ","
    '问题号:51072
    '密码输入限制_In     In 医疗卡类别.密码输入限制%Type:=0
    If optPassConfine(0).value Then
         strSQL = strSQL & "" & 0 & ","
    ElseIf optPassConfine(1) Then
         strSQL = strSQL & "" & 1 & ","
    ElseIf optPassConfine(2) Then
         strSQL = strSQL & "" & 2 & ","
    End If
    '是否缺省密码_In     In 医疗卡类别.是否缺省密码%Type:=0
    strSQL = strSQL & "" & IIf(chkPassByIDCard.value, 1, 0) & ","
    '问题号:56508
    '是否制卡_In
    strSQL = strSQL & "" & chkMakeCard & ","
    '是否发卡_In
    strSQL = strSQL & "" & IIf(chkSendCard.Enabled, chkSendCard, 0) & ","
    '是否写卡_In
    strSQL = strSQL & "" & chkWriteCard & ","
    '问题号:57697
    '险类_In
    strSQL = strSQL & "" & IIf(CStr(txtEdit(mtxtIdx.idx_险类).Tag) = "", 0, Val(txtEdit(mtxtIdx.idx_险类).Tag)) & ","
    '问题号:57326
    If OptSendCard(moptIdx.idx_不限制).value Then
        strSQL = strSQL & "" & 0 & ","
    ElseIf OptSendCard(moptIdx.idx_只发一张卡).value Then
        strSQL = strSQL & "" & 1 & ","
    ElseIf OptSendCard(moptIdx.idx_发多张卡并提醒).value Then
        strSQL = strSQL & "" & 2 & ","
    End If
    '77872,李南春,2014/10/28:是否支持转帐及代扣
    '是否转帐及代扣_In  In 医疗卡类别.是否转帐及代扣%Type:=0
    strSQL = strSQL & "" & IIf(chk转帐及代扣.Enabled And chk转帐及代扣.value = 1, 1, 0) & ","
    
    '85565,李南春,2015/7/9:读卡性质及键盘控制方式
    strValue = IIf(chk读卡性质(刷卡).value = 1, "1", "0")
    strValue = strValue & IIf(chk读卡性质(扫描卡).value = 1, "1", "0")
    strValue = strValue & IIf(chk读卡性质(接触式读卡).value = 1, "1", "0")
    strValue = strValue & IIf(chk读卡性质(非接触式读卡).value = 1, "1", "0")
    strSQL = strSQL & "'" & strValue & "',"
    
    If opt键盘控制(禁止).value Then
        strSQL = strSQL & "" & 0 & ","
    ElseIf opt键盘控制(数字).value Then
        strSQL = strSQL & "" & 1 & ","
    ElseIf opt键盘控制(字符).value Then
        strSQL = strSQL & "" & 2 & ","
    End If
    '是否证件_In  In 医疗卡类别.是否证件%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '是否持卡消费_In  In 医疗卡类别.是否持卡消费%Type:=0
    strSQL = strSQL & "" & IIf(chk持卡消费.Enabled And chk持卡消费.value = 1, 1, 0) & ","
    '发送调用接口_In  In 医疗卡类别.发送调用接口%Type:=0
    strSQL = strSQL & "" & IIf(chk发送调用接口.Enabled And chk发送调用接口.value = 1, 1, 0) & ","
    '是否退款验卡_In   In 医疗卡类别.是否退款验卡%Type := 0
    strSQL = strSQL & "" & IIf(chk退款验卡.Enabled And chk退款验卡.value = 1, 1, 0) & ","
    '设备是否启用回车_In  In 医疗卡类别.设备是否启用回车%Type:=0
    strSQL = strSQL & "" & IIf(chkOpenEnter.Enabled And chkOpenEnter.value = 1, 1, 0) & ","
    '发卡卡号控制_In   In 医疗卡类别.发卡控制%Type := 0
    If OptSendCardLen(moptLenIdx.idx_卡号不足禁止).value Then
        strSQL = strSQL & "" & 0 & ","
    ElseIf OptSendCardLen(moptLenIdx.idx_卡号不限制).value Then
        strSQL = strSQL & "" & 1 & ","
    ElseIf OptSendCardLen(moptLenIdx.idx_卡号不足提醒).value Then
        strSQL = strSQL & "" & 2 & ","
    End If
    '是否缺省退现_In   In 医疗卡类别.是否缺省退现%Type := 0
    strSQL = strSQL & "" & IIf(chk缺省退现.Enabled And chk缺省退现.value = 1, 1, 0) & ")"
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    Select Case Index
    Case mtxtIdx.idx_名称, mtxtIdx.idx_备注, mtxtIdx.idx_短名
        zlCommFun.OpenIme True
    Case Else
    End Select
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> mtxtIdx.idx_医疗卡费 Then Exit Sub
    If KeyCode = vbKeyDelete Then txtEdit(Index).Text = ""
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = mtxtIdx.idx_卡号长度 Or Index = mtxtIdx.idx_编码 Then
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m数字式
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
        zlCommFun.OpenIme False
End Sub
Private Sub txtPasLen_Change()
    optPassInput(1).Caption = "固定输入" & Val(txtPasLen.Text) & "位"
    '问题号:51072
    txtPassByIDCard.Text = txtPasLen.Text
End Sub
Private Sub upd结束位置_Change()
     If upd结束位置.value < upd开始位置.value Then upd开始位置.value = upd结束位置.value
     If upd开始位置.value = 0 And upd结束位置.value = 0 Then chk卡号.value = 0
End Sub

Private Sub upd开始位置_Change()
     If upd结束位置.value < upd开始位置.value Then upd结束位置.value = upd开始位置.value
     If upd开始位置.value = 0 And upd结束位置.value = 0 Then chk卡号.value = 0
End Sub

Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化分页控件
    '编制:李南春
    '问题号:85565
    '日期:2015/7/8 17:19:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Dim intEditType As Integer '进入窗体时的操作类型
    
    Err = 0: On Error GoTo Errhand:
    
    Set objItem = tbPage.InsertItem(mPageIndex.读卡设置, "读卡设置", picProperty(0).hWnd, 0)
    objItem.Tag = mPageIndex.读卡设置
    
    Set objItem = tbPage.InsertItem(mPageIndex.密码设置, "密码设置", picProperty(1).hWnd, 0)
    objItem.Tag = mPageIndex.密码设置

    With tbPage
       tbPage.Item(0).Selected = True
       .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
       .PaintManager.BoldSelected = True
       .PaintManager.Layout = xtpTabLayoutAutoSize
       .PaintManager.StaticFrame = False
       .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Call picExpend_Resize
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub
