VERSION 5.00
Begin VB.Form frm医保帐户补入院 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人补办医保入院登记"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frm医保帐户补入院.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "登记(&X)"
      Height          =   350
      Left            =   6090
      TabIndex        =   9
      Top             =   6015
      Width           =   1100
   End
   Begin VB.Frame fra费用信息 
      Caption         =   "【费用信息】"
      ForeColor       =   &H00C00000&
      Height          =   705
      Left            =   75
      TabIndex        =   28
      Top             =   1815
      Width           =   8745
      Begin VB.TextBox txt费用余额 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txt预交余额 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txt担保额 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7380
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   1080
      End
      Begin VB.TextBox txt担保人 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未结费用"
         Height          =   180
         Left            =   2370
         TabIndex        =   31
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预交余额"
         Height          =   180
         Left            =   375
         TabIndex        =   29
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保额"
         Height          =   180
         Left            =   6765
         TabIndex        =   35
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保人"
         Height          =   180
         Left            =   4695
         TabIndex        =   33
         Top             =   300
         Width           =   540
      End
   End
   Begin VB.Frame fra基本信息 
      Caption         =   "【基本信息】"
      ForeColor       =   &H00C00000&
      Height          =   3345
      Left            =   75
      TabIndex        =   59
      Top             =   2580
      Width           =   8745
      Begin VB.TextBox txt医疗付款 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt出生日期 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   570
         Width           =   1140
      End
      Begin VB.TextBox txt联系人关系 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1890
         Width           =   2000
      End
      Begin VB.TextBox txt身份 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txt职业 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txt婚姻状况 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txt国籍 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt学历 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   240
         Width           =   1140
      End
      Begin VB.TextBox txt民族 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt出生地点 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   900
         Width           =   3225
      End
      Begin VB.TextBox txt家庭地址 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1230
         Width           =   3150
      End
      Begin VB.TextBox txt户口邮编 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1230
         Width           =   1170
      End
      Begin VB.TextBox txt联系人姓名 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1170
      End
      Begin VB.TextBox txt联系人地址 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1890
         Width           =   3225
      End
      Begin VB.TextBox txt联系人电话 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2220
         Width           =   2000
      End
      Begin VB.TextBox txt工作单位 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   2220
         Width           =   3225
      End
      Begin VB.TextBox txt单位电话 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   2550
         Width           =   2000
      End
      Begin VB.TextBox txt单位邮编 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   2550
         Width           =   1170
      End
      Begin VB.TextBox txt单位开户行 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   2880
         Width           =   3135
      End
      Begin VB.TextBox txt单位帐号 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   2880
         Width           =   3225
      End
      Begin VB.TextBox txt家庭电话 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2000
      End
      Begin VB.TextBox txt身份证号 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   900
         Width           =   3150
      End
      Begin VB.Label lbl医疗付款 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医疗付款"
         Height          =   180
         Left            =   345
         TabIndex        =   81
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbl出生日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         Height          =   180
         Left            =   6570
         TabIndex        =   80
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl出生地点 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生地点"
         Height          =   180
         Left            =   4470
         TabIndex        =   79
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl身份证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   180
         Left            =   345
         TabIndex        =   78
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl身份 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份"
         Height          =   180
         Left            =   4830
         TabIndex        =   77
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lbl职业 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职业"
         Height          =   180
         Left            =   2685
         TabIndex        =   76
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lbl民族 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民族"
         Height          =   180
         Left            =   4830
         TabIndex        =   75
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lbl国籍 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国籍"
         Height          =   180
         Left            =   2685
         TabIndex        =   74
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lbl学历 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "学历"
         Height          =   180
         Left            =   6930
         TabIndex        =   73
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lvl婚姻状况 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况"
         Height          =   180
         Left            =   345
         TabIndex        =   72
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl家庭地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址"
         Height          =   180
         Left            =   345
         TabIndex        =   71
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label lbl家庭电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭电话"
         Height          =   180
         Left            =   345
         TabIndex        =   70
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label lbl户口邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "户口邮编"
         Height          =   180
         Left            =   4470
         TabIndex        =   69
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label lbl联系人姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人姓名"
         Height          =   180
         Left            =   4290
         TabIndex        =   68
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label lbl联系人关系 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人关系"
         Height          =   180
         Left            =   165
         TabIndex        =   67
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label lbl联系人地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人地址"
         Height          =   180
         Left            =   4290
         TabIndex        =   66
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label lbl联系人电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人电话"
         Height          =   180
         Left            =   165
         TabIndex        =   65
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label lbl工作单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工作单位"
         Height          =   180
         Left            =   4470
         TabIndex        =   64
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label lbl单位电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位电话"
         Height          =   180
         Left            =   345
         TabIndex        =   63
         Top             =   2610
         Width           =   720
      End
      Begin VB.Label lbl单位邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位邮编"
         Height          =   180
         Left            =   4470
         TabIndex        =   62
         Top             =   2610
         Width           =   720
      End
      Begin VB.Label lbl单位开户行 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位开户行"
         Height          =   180
         Left            =   165
         TabIndex        =   61
         Top             =   2940
         Width           =   900
      End
      Begin VB.Label lbl单位帐号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位帐号"
         Height          =   180
         Left            =   4470
         TabIndex        =   60
         Top             =   2940
         Width           =   720
      End
   End
   Begin VB.Frame fra在院信息 
      Caption         =   "【住院信息】"
      ForeColor       =   &H00C00000&
      Height          =   1695
      Left            =   75
      TabIndex        =   5
      Top             =   30
      Width           =   8730
      Begin VB.CommandButton cmdTurn 
         Caption         =   "门诊费用转住院(&T)"
         Height          =   300
         Left            =   5280
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "热键:F12(医保病人验证)"
         Top             =   225
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.ComboBox cob住院次数 
         Height          =   300
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1140
      End
      Begin VB.TextBox txt诊断 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   1260
         Width           =   3180
      End
      Begin VB.CommandButton cmdYB 
         Caption         =   "验证(&V)"
         Height          =   300
         Left            =   4440
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "热键:F12(医保病人验证)"
         Top             =   225
         Width           =   825
      End
      Begin VB.TextBox txt护理 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   885
         Width           =   1065
      End
      Begin VB.TextBox txt床号 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   885
         Width           =   1170
      End
      Begin VB.TextBox txt入院时间 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   885
         Width           =   1110
      End
      Begin VB.TextBox txt科室 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   885
         Width           =   1110
      End
      Begin VB.TextBox txt医保号 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   1260
         Width           =   3225
      End
      Begin VB.TextBox txt住院号 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   1170
      End
      Begin VB.TextBox txt费别 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   555
         Width           =   1110
      End
      Begin VB.TextBox txt性别 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox txt病人ID 
         Height          =   300
         Left            =   1110
         TabIndex        =   0
         Top             =   225
         Width           =   1110
      End
      Begin VB.TextBox txt姓名 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   555
         Width           =   1110
      End
      Begin VB.TextBox txt年龄 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   555
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院诊断"
         Height          =   180
         Left            =   330
         TabIndex        =   83
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "护理"
         Height          =   180
         Left            =   4800
         TabIndex        =   24
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         Height          =   180
         Left            =   2715
         TabIndex        =   22
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院时间"
         Height          =   180
         Left            =   6540
         TabIndex        =   26
         Top             =   945
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         Height          =   180
         Left            =   690
         TabIndex        =   20
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医保号"
         Height          =   180
         Left            =   4620
         TabIndex        =   8
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   2535
         TabIndex        =   7
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   690
         TabIndex        =   12
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   2715
         TabIndex        =   14
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl年龄 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   4800
         TabIndex        =   16
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         Height          =   180
         Left            =   6900
         TabIndex        =   18
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lbl病人ID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人ID"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   510
         TabIndex        =   6
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   780
      TabIndex        =   11
      Top             =   6015
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7380
      TabIndex        =   10
      Top             =   6015
      Width           =   1100
   End
End
Attribute VB_Name = "frm医保帐户补入院"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;99-所有交易增加附加参数(最新版)

Private mlng病人ID As Long '要修改或查看的病人ID
Private mlng主页ID As Long '要修改或查看的主页ID
Private mstr医保号 As String
Public mint险类 As Integer
Private mstrNOS As String   '选择转入的单据,票据,结帐ID,险类(非医保为零):H0000001,F000023,81235,901;H0000002,F000045,81263,901;...

Private Function ReadCard() As Boolean
'功能：读取指定病人信息,并显示在界面上
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrH
    #If gverControl < 6 Then
        gstrSQL = "Select * From 病人信息 Where 病人ID=" & mlng病人ID
    #Else
        gstrSQL = "Select A.病人id, A.门诊号, A.住院号, A.就诊卡号, A.卡验证码, A.费别, A.医疗付款方式, A.姓名, A.性别, A.年龄, A.出生日期, A.出生地点, A.身份证号, A.其他证件, A.身份, A.职业, A.民族, A.国籍, A.区域, A.学历, A.婚姻状况, A.家庭地址," & vbNewLine & _
            "      A.家庭电话, A.家庭地址邮编 As 户口邮编, A.监护人, A.联系人姓名, A.联系人关系, A.联系人地址, A.联系人电话, A.合同单位id, A.工作单位, A.单位电话, A.单位邮编, A.单位开户行, A.单位帐号, A.担保人, A.担保额, A.担保性质, A.就诊时间, A.就诊状态," & vbNewLine & _
            "      A.就诊诊室, A.住院次数, A.当前科室id, A.当前病区id, A.当前床号, A.入院时间, A.出院时间, A.在院, A.Ic卡号, A.健康号, A.医保号, A.险类, A.查询密码, A.登记时间, A.停用时间, A.锁定" & vbNewLine & _
            "From 病人信息 A Where A.病人ID =" & mlng病人ID
    #End If
    rsTmp.CursorLocation = adUseClient
    
    Call OpenRecordset(rsTmp, Me.Caption)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp.RecordCount <> 1 Then Exit Function
    
    '住院信息
    txt病人ID.Locked = True
    txt病人ID.Text = mlng病人ID
    txt病人ID.Locked = False
    
    txt姓名.Text = rsTmp!姓名
    txt住院号.Text = IIf(IsNull(rsTmp!住院号), "", rsTmp!住院号)
    
    '基本信息
    txt性别.Text = IIf(IsNull(rsTmp!性别), "", rsTmp!性别)
    txt年龄.Text = IIf(IsNull(rsTmp!年龄), "", rsTmp!年龄)
    txt费别.Text = IIf(IsNull(rsTmp!费别), "", rsTmp!费别)
    txt医疗付款.Text = IIf(IsNull(rsTmp!医疗付款方式), "", rsTmp!医疗付款方式)
    txt国籍.Text = IIf(IsNull(rsTmp!国籍), "", rsTmp!国籍)
    txt民族.Text = IIf(IsNull(rsTmp!民族), "", rsTmp!民族)
    txt学历.Text = IIf(IsNull(rsTmp!学历), "", rsTmp!学历)
    txt婚姻状况.Text = IIf(IsNull(rsTmp!婚姻状况), "", rsTmp!婚姻状况)
    txt职业.Text = IIf(IsNull(rsTmp!职业), "", rsTmp!职业)
    txt身份.Text = IIf(IsNull(rsTmp!身份), "", rsTmp!身份)
    txt出生日期.Text = Format(IIf(IsNull(rsTmp!出生日期), "", rsTmp!出生日期), "yyyy-MM-dd")
    txt身份证号.Text = IIf(IsNull(rsTmp!身份证号), "", rsTmp!身份证号)
    txt出生地点.Text = IIf(IsNull(rsTmp!出生地点), "", rsTmp!出生地点)
    txt家庭地址.Text = IIf(IsNull(rsTmp!家庭地址), "", rsTmp!家庭地址)
    txt家庭电话.Text = IIf(IsNull(rsTmp!家庭电话), "", rsTmp!家庭电话)
    txt户口邮编.Text = IIf(IsNull(rsTmp!户口邮编), "", rsTmp!户口邮编)
    txt联系人姓名.Text = IIf(IsNull(rsTmp!联系人姓名), "", rsTmp!联系人姓名)
    txt联系人关系.Text = IIf(IsNull(rsTmp!联系人关系), "", rsTmp!联系人关系)
    txt联系人地址.Text = IIf(IsNull(rsTmp!联系人地址), "", rsTmp!联系人地址)
    txt联系人电话.Text = IIf(IsNull(rsTmp!联系人电话), "", rsTmp!联系人电话)
    txt工作单位.Text = IIf(IsNull(rsTmp!工作单位), "", rsTmp!工作单位)
    txt单位电话.Text = IIf(IsNull(rsTmp!单位电话), "", rsTmp!单位电话)
    txt单位邮编.Text = IIf(IsNull(rsTmp!单位邮编), "", rsTmp!单位邮编)
    txt单位开户行.Text = IIf(IsNull(rsTmp!单位开户行), "", rsTmp!单位开户行)
    txt单位帐号.Text = IIf(IsNull(rsTmp!单位帐号), "", rsTmp!单位帐号)
        
    '费用信息
    txt担保人.Text = IIf(IsNull(rsTmp!担保人), "", rsTmp!担保人)
    txt担保额.Text = Format(IIf(IsNull(rsTmp!担保额), "", rsTmp!担保额), "0.00")
    
    #If gverControl >= 5 Then
        gstrSQL = "Select * From 病人余额 Where 性质=1 And 类型=2 And 病人ID=" & mlng病人ID
    #Else
        gstrSQL = "Select * From 病人余额 Where 性质=1 And 病人ID=" & mlng病人ID
    #End If
    Call OpenRecordset(rsTmp, Me.Caption)
    
    If Not rsTmp.EOF Then
        txt费用余额.Text = Format(IIf(IsNull(rsTmp!费用余额), 0, rsTmp!费用余额), "0.00")
        txt预交余额.Text = Format(IIf(IsNull(rsTmp!预交余额), 0, rsTmp!预交余额), "0.00")
    End If
    
    
    '病人医保信息
    txt医保号.Text = ""
    mstr医保号 = ""
    
    
    '病案主页信息
    gstrSQL = "Select A.入院日期,A.出院病床,b.名称 as 入院科室,C.名称 as 护理等级,A.入院科室ID" & _
              " From 病案主页 A,部门表 B,护理等级 C" & _
              " Where A.病人ID=" & mlng病人ID & " And A.主页ID=" & mlng主页ID & _
              "       and A.入院科室ID=B.ID and A.护理等级ID=C.序号(+) "
    Call OpenRecordset(rsTmp, Me.Caption)
    '2006-06-13 未入科病人无科室信息
    txt科室.Text = IIf(IsNull(rsTmp!入院科室), "无", rsTmp!入院科室)
    txt科室.Tag = Val("" & rsTmp!入院科室ID)
    txt护理.Text = IIf(IsNull(rsTmp!护理等级), "无", rsTmp!护理等级)
    txt床号.Text = IIf(IsNull(rsTmp!出院病床), "", rsTmp!出院病床)
    txt入院时间.Text = Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm")
    
    '曾明春(2006-2-17):入院诊断,HIS+补充填写入院诊断时诊断类型为2
    gstrSQL = "Select 描述信息" & _
              " From 诊断情况" & _
              " Where 病人ID=" & mlng病人ID & " And 主页ID=" & mlng主页ID & " and 诊断类型 in (1,2) "
    Call OpenRecordset(rsTmp, Me.Caption)
    If rsTmp.EOF = False Then
        txt诊断.Text = Nvl(rsTmp("描述信息"))
    End If
    
    If gclsInsure.GetCapability(support必须录入入出诊断, 0, mint险类) = True Then
        txt诊断.Locked = False
        txt诊断.BackColor = txt病人ID.BackColor
    Else
        txt诊断.Locked = True
        txt诊断.BackColor = txt住院号.BackColor
    End If
    
    ReadCard = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
'补办入院登记
    
    If mlng病人ID = 0 Then
        MsgBox "请先确定等补办入院登记的病人。", vbInformation, gstrSysName
        txt病人ID.SetFocus
        Exit Sub
    End If
    If mstr医保号 = "" Then
        MsgBox "请先验证该病人是否可以进行医保入院。", vbInformation, gstrSysName
        cmdYB.SetFocus
        Exit Sub
    End If
    If txt诊断.Locked = False And txt诊断.Text = "" Then
        MsgBox "请填写入院诊断。", vbInformation, gstrSysName
        txt诊断.SetFocus
        Exit Sub
    End If
    If zlCommFun.StrIsValid(txt诊断.Text, txt诊断.MaxLength, txt诊断.hwnd, "入院诊断") = False Then
        Exit Sub
    End If
    
    If mint险类 = 106 Then '内江医保需要判断补办销期
       If Not 判断补办效期_成都内江(mlng病人ID, mlng主页ID) Then
          Exit Sub
       End If
    End If
    
        '门诊费用转住院
    If mstrNOS <> "" Then
        If Not frmChargeTurn.ExecuteTurn(mstrNOS, txt住院号.Text, mlng主页ID, CDate(txt入院时间.Text), Val(txt科室.Tag)) Then
            Exit Sub
        End If
    End If
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    gstrSQL = "zl_病案主页_补办医保入院(" & mlng病人ID & "," & mlng主页ID & "," & mint险类 & ",'" & txt诊断 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'$IF HIS9.19
#If gverControl = 0 Then
    If gclsInsure.ComeInSwap(mlng病人ID, mlng主页ID, mstr医保号) = False Then
        '登记失败
        gcnOracle.RollbackTrans
        Exit Sub
    End If
#Else
'$ELSE  HIS+
    If gclsInsure.ComeInSwap(mlng病人ID, mlng主页ID, mstr医保号, mint险类) = False Then
        '登记失败
        gcnOracle.RollbackTrans
        Exit Sub
    End If
#End If
'$END IF
gcnOracle.CommitTrans
    MsgBox "病人" & txt姓名.Text & "补办医保入院成功！" & IIf(mint险类 > 900, vbCrLf & "病人费用明细中医保数据已经按医保规则重算。", "") _
        , vbInformation, gstrSysName
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdTurn_Click()
    gstrDec = "0.00"
    Call frmChargeTurn.ShowME(Me, Val(txt病人ID.Text), mstrNOS)
End Sub

Private Sub cmdYB_Click()
'验证医保病人身份
    Dim lng病人ID As Long, int险类 As Integer
    Dim strYBPati As String, arr信息 As Variant
    
    If mlng病人ID = 0 Then
        MsgBox "请先确定等补办入院登记的病人。", vbInformation, gstrSysName
        txt病人ID.SetFocus: Exit Sub
    End If
    lng病人ID = mlng病人ID
    int险类 = mint险类
'$IF HIS9.19
#If gverControl = 0 Then
    strYBPati = gclsInsure.Identify(1, lng病人ID)
'$ELSE
#Else
    strYBPati = gclsInsure.Identify(1, lng病人ID, int险类)
#End If
'$END IF
    If strYBPati = "" Then
        MsgBox "该病人身份验证失败。", vbInformation, gstrSysName
        cmdYB.SetFocus: Exit Sub
    End If
    
    arr信息 = Split(strYBPati, ";")
    If lng病人ID <> 0 Then mlng病人ID = lng病人ID
    
    '空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID,...
    If UBound(arr信息) >= 8 Then
        '可能身份合并后病人ID发生变化
        If Val(arr信息(8)) <> Val(txt病人ID.Text) Then
            txt病人ID.Text = "-" & Val(arr信息(8))
            Call GetPatient
            Call ReadCard
        End If
        
        txt医保号.Text = arr信息(1)
        mstr医保号 = txt医保号.Text
        
        txt姓名.Text = arr信息(3)
        txt性别.Text = arr信息(4)
        txt出生日期.Text = arr信息(5)
        txt身份证号.Text = arr信息(6)
        
        If IsZLHIS10 Then cmdTurn.Visible = True
        
        cmdOK.SetFocus
    End If
End Sub

Private Sub cob住院次数_Click()
 If mint险类 <> TYPE_大连开发区 And mint险类 <> TYPE_大连市 Then Exit Sub
    mlng主页ID = cob住院次数.ItemData(cob住院次数.ListIndex)
  Call ReadCard
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Call cmdYB_Click
    End If
End Sub

Private Sub Form_Load()
    mstrNOS = ""
    mlng病人ID = 0
    mlng主页ID = 0
End Sub

Private Sub txt病人ID_Change()
    If txt病人ID.Locked = False Then
        mlng病人ID = 0
        mlng主页ID = 0
    End If
End Sub

Private Sub txt病人ID_GotFocus()
    zlControl.TxtSelAll txt病人ID
End Sub

Private Sub txt病人ID_KeyPress(KeyAscii As Integer)
    Dim lng病人ID  As Long
    
    '转换成大写(汉字不可处理)
    If KeyAscii > 0 Then KeyAscii = asc(UCase(Chr(KeyAscii)))
    
    If InStr("0123456789", Chr(KeyAscii)) = 0 And (txt病人ID.Text = "" Or txt病人ID.SelLength = Len(txt病人ID.Text)) Then
        txt病人ID.MaxLength = 15
    End If
    
    If Len(Trim(Me.txt病人ID.Text)) = 0 And KeyAscii = 13 Then
        If frm医保病人选择.Get病人(lng病人ID) = True Then
            txt病人ID.Text = "A" & lng病人ID
        End If
    End If
    Me.Refresh
    
    '刷卡完毕或输入号码后回车
    If (KeyAscii = 13 And Trim(txt病人ID.Text) <> "") Then
        If Val(txt病人ID.Text) = mlng病人ID And mlng病人ID > 0 Then
            If mstr医保号 = "" Then
                cmdYB.SetFocus
            Else
                cmdOK.SetFocus
            End If
            Exit Sub
        End If
        
        If KeyAscii <> 13 Then
            txt病人ID.Text = txt病人ID.Text & Chr(KeyAscii)
            txt病人ID.SelStart = Len(txt病人ID.Text)
        End If
        KeyAscii = 0
        
        '20040923:刘兴宏加入
        Call Load住院次数
        
        If Not GetPatient() Then
            MsgBox "没有发现该病人的住院信息,请重新输入！", vbInformation, gstrSysName
            txt病人ID.Text = ""
            txt病人ID.SetFocus
            Exit Sub
        Else
            Call ReadCard
            cmdYB.SetFocus
        End If
    End If

End Sub

Private Function GetPatient() As Boolean
''功能：读取病人信息
'返回:是否读取成功,成功时rsInfo中包含病人信息,失败时rsInfo=Close
    Dim rsInfo As New ADODB.Recordset
    Dim strCode As String
    Dim lng住院次数 As Long
    '刘兴宏:2004/09/23:取消了出院病人和险类的限制
    Dim bln大连 As Boolean
    bln大连 = (mint险类 = TYPE_大连开发区 Or mint险类 = TYPE_大连市)
    
    strCode = Trim(txt病人ID.Text)
    On Error GoTo ErrH
    If bln大连 Then
        lng住院次数 = cob住院次数.ItemData(cob住院次数.ListIndex)
    Else
        lng住院次数 = 1
    End If
    If (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '病人ID
        gstrSQL = _
            "Select C.病人ID,C.主页ID" & _
            " From 病人信息 A,病案主页 C" & _
            " Where A.病人ID=C.病人ID " & IIf(bln大连, " And C.主页id=" & lng住院次数, " And Nvl(A.住院次数,0)=C.主页ID") & _
            "   And A.病人ID=" & Val(Mid(strCode, 2)) & _
            "     " & IIf(bln大连, "", "  and C.险类 is null and C.出院日期 is null")
    ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then '住院号
        gstrSQL = _
            "Select C.病人ID,C.主页ID" & _
            " From 病人信息 A,病案主页 C" & _
            " Where A.病人ID=C.病人ID " & IIf(bln大连, " And C.主页id=" & lng住院次数, " And Nvl(A.住院次数,0)=C.主页ID") & _
            "       And A.住院号=" & Mid(strCode, 2) & _
            "      " & IIf(bln大连, "", " and C.险类 is null and C.出院日期 is null")
    ElseIf (Left(strCode, 1) = "C" Or Left(strCode, 1) = ";") And IsNumeric(Split(Mid(strCode, 2), "?")(0)) Then '住院号
        gstrSQL = _
            "Select C.病人ID,C.主页ID" & _
            " From 病人信息 A,病案主页 C" & _
            " Where A.病人ID=C.病人ID " & IIf(bln大连, " And C.主页id=" & lng住院次数, " And Nvl(A.住院次数,0)=C.主页ID") & _
            "       And A.就诊卡号='" & Split(Mid(strCode, 2), "?")(0) & _
            "'      " & IIf(bln大连, "", " and C.险类 is null and C.出院日期 is null")
    Else '当作姓名
        gstrSQL = _
            "Select C.病人ID,C.主页ID" & _
            " From 病人信息 A,病案主页 C" & _
            " Where A.病人ID=C.病人ID " & IIf(bln大连, " And C.主页id=" & lng住院次数, " And Nvl(A.住院次数,0)=C.主页ID") & _
            "       And A.姓名='" & strCode & _
            "'    " & IIf(bln大连, "", "   and C.险类 is null and C.出院日期 is null")
    End If
    
    rsInfo.CursorLocation = adUseClient
    Call OpenRecordset(rsInfo, Me.Caption)
    '可以改时间
    txt入院时间.Locked = Not bln大连
    '读取失败
    If rsInfo.EOF Then
        Exit Function
    End If
        
    mlng病人ID = rsInfo("病人ID")
    mlng主页ID = rsInfo("主页ID")
    
    GetPatient = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function Load住院次数() As Boolean
     Dim rsInfo  As New ADODB.Recordset
     Dim strCode  As String
     Dim bln大连 As Boolean
    '加载住院次数
    bln大连 = (mint险类 = TYPE_大连开发区 Or mint险类 = TYPE_大连市)
    
    strCode = Trim(txt病人ID.Text)
    If (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '病人ID
        gstrSQL = _
            "Select C.病人ID,C.主页ID" & _
            " From 病案主页 C" & _
            " Where 病人ID=" & Val(Mid(strCode, 2)) & _
            "   order by C.主页id Desc"
            
    ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then '住院号
        gstrSQL = _
            "Select C.病人ID,C.主页ID" & _
            " From 病人信息 A,病案主页 C" & _
            " Where A.病人ID=C.病人ID  And A.住院号=" & Mid(strCode, 2) & _
            "      " & IIf(bln大连, "", " and C.险类 is null and C.出院日期 is null") & _
            "   order by C.主页id Desc"
    Else '当作姓名
        gstrSQL = _
            "Select C.病人ID,C.主页ID" & _
            " From 病人信息 A,病案主页 C" & _
            " Where A.病人ID=C.病人ID  And A.姓名='" & strCode & _
            "'    " & IIf(bln大连, "", "   and C.险类 is null and C.出院日期 is null") & _
            "   order by C.主页id Desc"

    End If
    zlDatabase.OpenRecordset rsInfo, gstrSQL, "获取住院次数"
    
     With rsInfo
        cob住院次数.Clear
        Do While Not rsInfo.EOF
            cob住院次数.AddItem Nvl(!主页ID, 0) & "次"
            cob住院次数.ItemData(cob住院次数.NewIndex) = Nvl(!主页ID, 0)
            .MoveNext
        Loop
        If cob住院次数.ListCount <> 0 Then cob住院次数.ListIndex = 0
        cob住院次数.Enabled = bln大连
     End With
     
    
End Function

