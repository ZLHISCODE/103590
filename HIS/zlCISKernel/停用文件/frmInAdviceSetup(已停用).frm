VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInAdviceSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "住院医嘱选项"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9675
   Icon            =   "frmInAdviceSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab tabPar 
      Height          =   8340
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   14711
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   617
      WordWrap        =   0   'False
      TabCaption(0)   =   "医嘱下达(&1)"
      TabPicture(0)   =   "frmInAdviceSetup.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl可用药房"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl卫材"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl缺省药房"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "vsfDrugStore"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra入院诊断"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra医嘱下达"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra输液配置"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cbo卫材"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraPurMed"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra期效"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "医嘱处理(&2)"
      TabPicture(1)   =   "frmInAdviceSetup.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra超期收回"
      Tab(1).Control(1)=   "fra后续处理"
      Tab(1).Control(2)=   "fra校对病区"
      Tab(1).Control(3)=   "fraBat"
      Tab(1).Control(4)=   "fraBaby"
      Tab(1).Control(5)=   "fraAdvicePrint"
      Tab(1).Control(6)=   "fraBillPrint"
      Tab(1).Control(7)=   "fraBloodPrint"
      Tab(1).ControlCount=   8
      Begin VB.Frame fraBloodPrint 
         Caption         =   "输血申请单打印模式"
         Height          =   855
         Left            =   -68040
         TabIndex        =   75
         Top             =   5205
         Width           =   2295
         Begin VB.OptionButton optBloodPrintType 
            Caption         =   "新开时打印"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   77
            Top             =   600
            Width           =   1440
         End
         Begin VB.OptionButton optBloodPrintType 
            Caption         =   "发送时打印"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   76
            Top             =   285
            Value           =   -1  'True
            Width           =   1545
         End
      End
      Begin VB.Frame fra期效 
         Caption         =   "启用输液配制中心的医嘱期效"
         Height          =   1185
         Left            =   6600
         TabIndex        =   50
         Top             =   480
         Width           =   2655
         Begin VB.OptionButton opt期效 
            Caption         =   "长嘱和临嘱"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   53
            Top             =   840
            Value           =   -1  'True
            Width           =   1200
         End
         Begin VB.OptionButton opt期效 
            Caption         =   "临嘱"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   52
            Top             =   570
            Width           =   680
         End
         Begin VB.OptionButton opt期效 
            Caption         =   "长嘱"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   51
            Top             =   300
            Width           =   680
         End
      End
      Begin VB.Frame fraBillPrint 
         Caption         =   "医嘱发送后,诊疗单据"
         Height          =   1980
         Left            =   -68055
         TabIndex        =   66
         Top             =   6195
         Width           =   2295
         Begin VB.OptionButton optPrint 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   2
            Left            =   150
            TabIndex        =   69
            Top             =   1140
            Value           =   -1  'True
            Width           =   1080
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   68
            Top             =   760
            Width           =   1440
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   67
            Top             =   400
            Width           =   840
         End
      End
      Begin VB.Frame fraAdvicePrint 
         Caption         =   "医嘱单打印模式"
         Height          =   885
         Left            =   -70800
         TabIndex        =   63
         Top             =   5205
         Width           =   2535
         Begin VB.OptionButton optPrintType 
            Caption         =   "校对后打印"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   65
            Top             =   285
            Value           =   -1  'True
            Width           =   1545
         End
         Begin VB.OptionButton optPrintType 
            Caption         =   "新开时打印"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   64
            Top             =   600
            Width           =   1440
         End
      End
      Begin VB.Frame fraPurMed 
         Caption         =   "抗菌药物缺省用药目的"
         Height          =   1185
         Left            =   4320
         TabIndex        =   59
         Top             =   480
         Width           =   2205
         Begin VB.OptionButton optPurMed 
            Caption         =   "下达时确定"
            Height          =   180
            Index           =   0
            Left            =   255
            TabIndex        =   78
            Top             =   270
            Width           =   1635
         End
         Begin VB.OptionButton optPurMed 
            Caption         =   "预防"
            Height          =   180
            Index           =   1
            Left            =   255
            TabIndex        =   61
            Top             =   585
            Width           =   680
         End
         Begin VB.OptionButton optPurMed 
            Caption         =   "治疗"
            Height          =   180
            Index           =   2
            Left            =   255
            TabIndex        =   60
            Top             =   870
            Value           =   -1  'True
            Width           =   680
         End
      End
      Begin VB.Frame fraBaby 
         Caption         =   "医嘱处理缺省范围(含提醒)"
         Height          =   1200
         Left            =   -70800
         TabIndex        =   55
         Top             =   6975
         Width           =   2535
         Begin VB.OptionButton optBaby 
            Caption         =   "婴儿医嘱"
            Height          =   180
            Index           =   2
            Left            =   150
            TabIndex        =   58
            Top             =   900
            Width           =   1440
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "全部医嘱"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   57
            Top             =   285
            Value           =   -1  'True
            Width           =   1545
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "病人医嘱"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   56
            Top             =   592
            Width           =   1440
         End
      End
      Begin VB.ComboBox cbo卫材 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1425
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   500
         Width           =   2790
      End
      Begin VB.Frame fra输液配置 
         Caption         =   " 在输液配制中心发药的病人科室 "
         Height          =   2655
         Left            =   4320
         TabIndex        =   42
         Top             =   1800
         Width           =   4935
         Begin VB.CheckBox chkWithSelf 
            Caption         =   "自备药、不取药、离院带药允许发送到配制中心"
            Height          =   195
            Left            =   210
            TabIndex        =   83
            Top             =   2370
            Width           =   4485
         End
         Begin VB.CheckBox chk静脉营养 
            Caption         =   "配制中心不接收的静脉营养医嘱在病区配置"
            Height          =   195
            Left            =   210
            TabIndex        =   79
            Top             =   2115
            Width           =   4485
         End
         Begin VB.CommandButton cmdAllDel 
            Caption         =   "全清"
            Height          =   350
            Left            =   3720
            TabIndex        =   45
            Top             =   720
            Width           =   1100
         End
         Begin VB.CommandButton cmdAllSelect 
            Caption         =   "全选"
            Height          =   350
            Left            =   3720
            TabIndex        =   44
            Top             =   240
            Width           =   1100
         End
         Begin VB.ListBox lstDept 
            ForeColor       =   &H80000012&
            Height          =   1740
            IMEMode         =   3  'DISABLE
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   43
            Top             =   300
            Width           =   3465
         End
      End
      Begin VB.Frame fraBat 
         Caption         =   " 允许批量进行 "
         Height          =   690
         Left            =   -70800
         TabIndex        =   33
         Top             =   6195
         Width           =   2535
         Begin VB.CheckBox chkBat 
            Caption         =   "暂停/启用"
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   35
            Top             =   360
            Width           =   1110
         End
         Begin VB.CheckBox chkBat 
            Caption         =   "校对"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   660
         End
      End
      Begin VB.Frame fra校对病区 
         Caption         =   " 自动校对医嘱的病区 "
         Height          =   4530
         Left            =   -70800
         TabIndex        =   29
         Top             =   480
         Width           =   5040
         Begin VB.ListBox lst校对病区 
            ForeColor       =   &H80000012&
            Height          =   4050
            Left            =   165
            Style           =   1  'Checkbox
            TabIndex        =   32
            Top             =   270
            Width           =   3510
         End
         Begin VB.CommandButton cmd校对病区ALL 
            Caption         =   "全选"
            Height          =   350
            Left            =   3840
            TabIndex        =   31
            ToolTipText     =   "Ctrl+A"
            Top             =   240
            Width           =   1100
         End
         Begin VB.CommandButton cmd校对病区Clear 
            Caption         =   "全清"
            Height          =   350
            Left            =   3840
            TabIndex        =   30
            ToolTipText     =   "Ctrl+R"
            Top             =   720
            Width           =   1100
         End
      End
      Begin VB.Frame fra后续处理 
         Caption         =   " 后续处理与控制 "
         Height          =   4530
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   3945
         Begin MSComCtl2.DTPicker dtpEnd 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "yyyy-M-d"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   300
            Left            =   2535
            TabIndex        =   85
            Top             =   4155
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   97910786
            CurrentDate     =   42017
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "长嘱口服药发送结束时间"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   84
            Top             =   4245
            Width           =   2325
         End
         Begin VB.CheckBox chkLimit 
            Caption         =   "药品长嘱的给药途径发送次数以结束时间为准计算"
            Height          =   420
            Left            =   210
            TabIndex        =   74
            Top             =   2865
            Width           =   3225
         End
         Begin VB.CheckBox chk关闭医嘱 
            Caption         =   "发送完成后关闭医嘱窗体"
            Height          =   195
            Left            =   210
            TabIndex        =   47
            Top             =   2130
            Width           =   3165
         End
         Begin VB.CheckBox chkShort 
            Caption         =   "临嘱"
            Height          =   255
            Left            =   1200
            TabIndex        =   41
            Top             =   1155
            Width           =   975
         End
         Begin VB.CheckBox chkLong 
            Caption         =   "长嘱"
            Height          =   255
            Left            =   480
            TabIndex        =   40
            Top             =   1155
            Width           =   975
         End
         Begin VB.CheckBox chk皮试 
            Caption         =   "填写皮试结果时验证身份"
            Height          =   195
            Left            =   210
            TabIndex        =   37
            Top             =   3945
            Width           =   2355
         End
         Begin VB.CheckBox chk校对签名 
            Caption         =   "校对和确认停止时使用电子签名"
            Height          =   195
            Left            =   210
            TabIndex        =   36
            Top             =   3645
            Width           =   3165
         End
         Begin VB.CheckBox chk医技 
            Caption         =   "允许对医技下达的医嘱进行后续处理"
            Height          =   195
            Left            =   210
            TabIndex        =   28
            Top             =   3360
            Width           =   3180
         End
         Begin VB.CheckBox chk打印 
            Caption         =   "校对,确认停止,重整医嘱后进行打印"
            Height          =   405
            Left            =   210
            TabIndex        =   27
            Top             =   525
            Width           =   3180
         End
         Begin VB.CheckBox chk校对 
            Caption         =   "新开医嘱后自动校对计价"
            Height          =   195
            Left            =   210
            TabIndex        =   26
            Top             =   330
            Width           =   3180
         End
         Begin VB.CheckBox chk执行 
            Caption         =   "发送时将本科执行的项目填为已执行"
            Height          =   195
            Left            =   210
            TabIndex        =   25
            Top             =   930
            Width           =   3180
         End
         Begin VB.CheckBox chk医保审批 
            Caption         =   "发送时对医保病人检查项目是否审批"
            Height          =   195
            Left            =   210
            TabIndex        =   24
            Top             =   1830
            Value           =   1  'Checked
            Width           =   3180
         End
         Begin VB.CheckBox chkAutoVerify 
            Caption         =   "无须校对即可发送医嘱"
            Height          =   180
            Left            =   210
            TabIndex        =   23
            Top             =   1530
            Width           =   3180
         End
         Begin VB.CheckBox chkTurnCheck 
            Caption         =   "存在未校对医嘱或待发送的医嘱时禁止发送转科、出院、转院、死亡医嘱"
            Height          =   405
            Left            =   210
            TabIndex        =   22
            Top             =   2415
            Width           =   3180
         End
      End
      Begin VB.Frame fra超期收回 
         Caption         =   " 超期收回 "
         Height          =   2970
         Left            =   -74880
         TabIndex        =   16
         Top             =   5205
         Width           =   3945
         Begin VB.OptionButton optRoll 
            Caption         =   "销帐申请"
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   73
            Top             =   300
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optRoll 
            Caption         =   "负数记帐"
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   72
            Top             =   300
            Width           =   1095
         End
         Begin VB.CheckBox chk超期审核 
            Caption         =   "超期收回时自动审核本科执行的销帐申请"
            Height          =   195
            Left            =   210
            TabIndex        =   19
            Top             =   600
            Width           =   3680
         End
         Begin VB.CheckBox chkAutoRoll 
            Caption         =   "确认停止后自动执行超期收回"
            Height          =   195
            Left            =   210
            TabIndex        =   18
            Top             =   900
            Width           =   3180
         End
         Begin VB.ListBox lst发药类型 
            Columns         =   3
            ForeColor       =   &H80000012&
            Height          =   1320
            IMEMode         =   3  'DISABLE
            ItemData        =   "frmInAdviceSetup.frx":0044
            Left            =   210
            List            =   "frmInAdviceSetup.frx":0046
            Style           =   1  'Checkbox
            TabIndex        =   17
            Top             =   1500
            Width           =   3525
         End
         Begin VB.Label lblRoll 
            Caption         =   "费用收回模式"
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label lblSend 
            Caption         =   "以下发药方式的西药一但发药就不收回"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   1230
            Width           =   3255
         End
      End
      Begin VB.Frame fra医嘱下达 
         Caption         =   " 检查与控制 "
         Height          =   3585
         Left            =   4320
         TabIndex        =   9
         Top             =   4575
         Width           =   4935
         Begin VB.CommandButton cmdBloodTip 
            Caption         =   "输血申请注意事项设置"
            Height          =   350
            Left            =   225
            TabIndex        =   86
            Top             =   3180
            Width           =   2490
         End
         Begin VB.OptionButton optSTCheck 
            Caption         =   "配制中心接收的药品"
            Height          =   255
            Index           =   1
            Left            =   2970
            TabIndex        =   82
            Top             =   1845
            Width           =   1920
         End
         Begin VB.OptionButton optSTCheck 
            Caption         =   "所有药品"
            Height          =   255
            Index           =   0
            Left            =   1900
            TabIndex        =   81
            Top             =   1845
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.CheckBox chkST 
            Caption         =   "自动增加皮试并根据结果限制医嘱发送"
            Height          =   225
            Left            =   210
            TabIndex        =   62
            Top             =   1575
            Width           =   3555
         End
         Begin VB.CheckBox chk待入住病人医嘱下达 
            Caption         =   "允许给待入住病人下达医嘱"
            Height          =   195
            Left            =   210
            TabIndex        =   54
            Top             =   2400
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.CheckBox chk停嘱审核 
            Caption         =   $"frmInAdviceSetup.frx":0048
            Height          =   195
            Left            =   210
            TabIndex        =   46
            Top             =   2640
            Width           =   3180
         End
         Begin VB.CommandButton cmdAdviceSortSet 
            Caption         =   "排序规则设置(&S)"
            Height          =   350
            Left            =   3120
            TabIndex        =   39
            Top             =   2800
            Width           =   1695
         End
         Begin VB.CheckBox chkAdviceSort 
            Caption         =   "保存医嘱时自动排序"
            Height          =   255
            Left            =   210
            TabIndex        =   38
            Top             =   2880
            Width           =   1935
         End
         Begin VB.CheckBox chk术后医嘱 
            Caption         =   "手术执行完成后才允许下达术后医嘱"
            Height          =   195
            Left            =   210
            TabIndex        =   15
            Top             =   1350
            Width           =   3180
         End
         Begin VB.CheckBox chk先输单量 
            Caption         =   "下达临嘱时先输入单量"
            Height          =   195
            Left            =   210
            TabIndex        =   14
            Top             =   500
            Width           =   3180
         End
         Begin VB.CheckBox chk出院诊断 
            Caption         =   "下达出院医嘱时检查出院诊断的填写"
            Height          =   195
            Left            =   210
            TabIndex        =   13
            Top             =   990
            Width           =   3180
         End
         Begin VB.CheckBox chk一次性 
            Caption         =   "临嘱的执行频率缺省为一次性"
            Height          =   195
            Left            =   210
            TabIndex        =   12
            Top             =   285
            Width           =   3180
         End
         Begin VB.CheckBox chk天数 
            Caption         =   "下达药品临嘱时可以指定用药天数"
            Height          =   195
            Left            =   210
            TabIndex        =   11
            Top             =   750
            Width           =   3180
         End
         Begin VB.CheckBox chkStopNurseGrade 
            Caption         =   "允许单独停止护理等级医嘱"
            Height          =   195
            Left            =   210
            TabIndex        =   10
            ToolTipText     =   "不允许时，只能通过转科、出院，或下达新的医嘱来停止护士等级"
            Top             =   2160
            Value           =   1  'Checked
            Width           =   3180
         End
         Begin VB.Label lblSTCheck 
            Caption         =   "限制医嘱的类型："
            Height          =   255
            Left            =   480
            TabIndex        =   80
            Top             =   1845
            Width           =   1455
         End
      End
      Begin VB.Frame fra入院诊断 
         Height          =   1480
         Left            =   120
         TabIndex        =   4
         Top             =   6680
         Width           =   4095
         Begin VB.ListBox lst入院诊断 
            Columns         =   3
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   1110
            IMEMode         =   3  'DISABLE
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   6
            Top             =   260
            Width           =   3900
         End
         Begin VB.CheckBox chk入院诊断 
            Caption         =   "下达这些类别的医嘱时检查是否填写诊断"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   0
            Width           =   3720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
         Height          =   4785
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   4095
         _cx             =   7223
         _cy             =   8440
         Appearance      =   1
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmInAdviceSetup.frx":0066
         ScrollTrack     =   0   'False
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
      Begin VB.Label lbl缺省药房 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省和可用药房"
         Height          =   180
         Left            =   120
         TabIndex        =   70
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label lbl卫材 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省发料部门"
         Height          =   180
         Left            =   120
         TabIndex        =   49
         Top             =   560
         Width           =   1080
      End
      Begin VB.Label lbl可用药房 
         Caption         =   $"frmInAdviceSetup.frx":00EF
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   4095
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   530
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   9675
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   8505
      Width           =   9675
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   7320
         TabIndex        =   0
         Top             =   60
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   8490
         TabIndex        =   1
         Top             =   60
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmInAdviceSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mMainPrivs As String
Public mint场合 As Integer  '调用场合:0-医生站调用,1-护士站调用,2-医技站调用
Private Const VsPubBackColor = &HFAEADA
Private mblnTmp As Boolean

Private Enum mCtlID
    chk长嘱口服药发送结束时间 = 0
End Enum

Private Sub chkAdviceSort_Click()
    cmdAdviceSortSet.Enabled = chkAdviceSort.value = 1 And chkAdviceSort.Enabled
End Sub

Private Sub chkInfo_Click(Index As Integer)
    Select Case Index
    Case chk长嘱口服药发送结束时间
        dtpEnd.Enabled = chkInfo(Index).value = 1
    End Select
End Sub

Private Sub chkLong_Click()
    mblnTmp = True
    If chkShort.value = 0 Then chk执行.value = chkLong.value
    mblnTmp = False
End Sub

Private Sub chkShort_Click()
    mblnTmp = True
    If chkLong.value = 0 Then chk执行.value = chkShort.value
    mblnTmp = False
End Sub

Private Sub chkST_Click()
    If chkST.value Then
        optSTCheck(0).Enabled = True
        optSTCheck(1).Enabled = True
    Else
        optSTCheck(0).Enabled = False
        optSTCheck(1).Enabled = False
    End If
End Sub

Private Sub cmdBloodTip_Click()
    Dim strPar As String
    strPar = cmdBloodTip.Tag
    Call frmInputBox.InputBox(Me, "输血申请注意事项", "内容：", 4000, 6, True, True, strPar)
    cmdBloodTip.Tag = strPar
End Sub

Private Sub optRoll_Click(Index As Integer)
    '负数记帐时，不使用自动审核申请单
    If Index = 0 Then
        chk超期审核.Enabled = False
        chk超期审核.value = 0
    Else
        chk超期审核.Enabled = True
    End If
End Sub

Private Sub chk入院诊断_Click()
    lst入院诊断.Enabled = chk入院诊断.value = 1 And lst入院诊断.Tag = ""
End Sub

Private Sub chk校对_Click()
    fra校对病区.Enabled = chk校对.value = 1
    cmd校对病区ALL.Enabled = fra校对病区.Enabled
    cmd校对病区Clear.Enabled = fra校对病区.Enabled
End Sub

Private Sub chk执行_Click()
    If mblnTmp Then Exit Sub
    chkLong.value = chk执行.value
    chkShort.value = chk执行.value
End Sub

Private Sub cmdAdviceSortSet_Click()
    frmPathSetup.mbytFun = 1
    frmPathSetup.Show vbModal, Me
End Sub

Private Sub cmdAllDel_Click()
    Dim i As Long
    
    For i = 0 To lstDept.ListCount - 1
        lstDept.Selected(i) = False
    Next
End Sub

Private Sub cmdAllSelect_Click()
    Dim i As Long, Y As Long
    
    Y = lstDept.ListIndex
    For i = 0 To lstDept.ListCount - 1
        lstDept.Selected(i) = True
    Next
    lstDept.ListIndex = Y
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim str诊断 As String, str病区 As String, i As Long
    Dim strValue As String, bytType As Long
    Dim arr可用药房(3) As String, arr缺省药房(3) As String, arrTmp() As String
    Dim str输液配置中心 As String, bln输液全选 As Boolean
    Dim blnSetup As Boolean, blnSetPara As Boolean
    
     '不检查是否指定了缺省药房，因为可能没有参数设置权限，参数类型是可自定义的。
     
    If fra入院诊断.Visible And chk入院诊断.value = 1 Then
        For i = 0 To lst入院诊断.ListCount - 1
            If lst入院诊断.Selected(i) Then
                str诊断 = str诊断 & Chr(lst入院诊断.ItemData(i))
            End If
        Next
        If str诊断 = "" Then
            MsgBox "请至少选择一种要检查入院诊断的医嘱类别。", vbInformation, gstrSysName
            lst入院诊断.SetFocus: Exit Sub
        End If
    End If
    If fra校对病区.Visible And fra校对病区.Enabled And chk校对.value = 1 Then
        For i = 0 To lst校对病区.ListCount - 1
            If lst校对病区.Selected(i) Then
                str病区 = str病区 & "," & lst校对病区.ItemData(i)
            End If
        Next
        str病区 = Mid(str病区, 2)
        If str病区 = "" Then
            MsgBox "请至少选择一个要自动对医嘱进行校对计价的病区。", vbInformation, gstrSysName
            lst校对病区.SetFocus: Exit Sub
        End If
    End If
    
    blnSetup = InStr(GetInsidePrivs(p住院医嘱下达), ";医嘱选项设置;") > 0
    blnSetPara = InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱选项设置;") > 0
    
    Call zlDatabase.SetPara("临嘱缺省一次性", chk一次性.value, glngSys, p住院医嘱下达, blnSetup)
    Call zlDatabase.SetPara("医嘱执行天数", chk天数.value, glngSys, p住院医嘱下达, blnSetup)
    Call zlDatabase.SetPara("自动处理皮试", chkST.value, glngSys, p住院医嘱下达, blnSetup)
    Call zlDatabase.SetPara("根据皮试结果限制医嘱发送类型", IIF(chkST.value = 1, IIF(optSTCheck(0).value, 0, 1), 0), glngSys, p住院医嘱下达, blnSetup)
    Call zlDatabase.SetPara("临嘱先输入单量", chk先输单量.value, glngSys, p住院医嘱下达, blnSetup)
    Call zlDatabase.SetPara("要求输入入院诊断", str诊断, glngSys, p住院医嘱下达, blnSetup)
    Call zlDatabase.SetPara("手术完成后下达术后医嘱", chk术后医嘱.value, glngSys, p住院医嘱下达, blnSetup)
    Call zlDatabase.SetPara("医嘱自动排序", chkAdviceSort.value, glngSys, p住院医嘱下达, blnSetup)
    Call zlDatabase.SetPara("允许给待入住病人下达医嘱", chk待入住病人医嘱下达.value, glngSys, p住院医嘱下达, blnSetup)
    '医嘱单打印模式
    Call zlDatabase.SetPara("医嘱单打印模式", IIF(optPrintType(1).value, 1, 0), glngSys, p住院医嘱下达, blnSetup)
    Call zlDatabase.SetPara("输血申请单打印模式", IIF(optBloodPrintType(1).value, 1, 2), glngSys, p住院医嘱发送, blnSetup)
    If mint场合 <> 2 Then
        Call zlDatabase.SetPara("要求输入出院诊断", chk出院诊断.value, glngSys, p住院医嘱下达, blnSetup)
        Call zlDatabase.SetPara("单独停止护理等级", chkStopNurseGrade.value, glngSys, p住院医嘱下达, blnSetup)
        Call zlDatabase.SetPara("实习医生停止医嘱需要审核", chk停嘱审核.value, glngSys, p住院医嘱下达, blnSetup)
    End If
    
    If chk关闭医嘱.Enabled = True Then
        Call zlDatabase.SetPara("发送完成后关闭医嘱窗体", chk关闭医嘱.value, glngSys, p住院医嘱下达, blnSetup)
    End If
    
    If mint场合 = 1 Then
        If chk校对.value = 0 Then
            Call zlDatabase.SetPara("自动完成校对计价", "", glngSys, p住院医嘱发送, blnSetPara)
        ElseIf UBound(Split(str病区, ",")) + 1 = lst校对病区.ListCount Then
            Call zlDatabase.SetPara("自动完成校对计价", "*", glngSys, p住院医嘱发送, blnSetPara)
        Else
            Call zlDatabase.SetPara("自动完成校对计价", str病区, glngSys, p住院医嘱发送, blnSetPara)
        End If
        
        Call zlDatabase.SetPara("本科执行自动完成", chkLong.value & chkShort.value, glngSys, p住院医嘱发送, blnSetPara)
        Call zlDatabase.SetPara("自动进入医嘱打印", chk打印.value, glngSys, p住院医嘱发送, blnSetPara)
        Call zlDatabase.SetPara("医技医嘱后续处理", chk医技.value, glngSys, p住院医嘱发送, blnSetPara)
        Call zlDatabase.SetPara("皮试验证身份", chk皮试.value, glngSys, p住院医嘱发送, blnSetPara)
        Call zlDatabase.SetPara("超期收回产生负数费用", IIF(optRoll(0).value, 1, 0), glngSys, p住院医嘱发送, blnSetPara)
        Call zlDatabase.SetPara("超期收回费用本科自动审核", chk超期审核.value, glngSys, p住院医嘱发送, blnSetPara)
        Call zlDatabase.SetPara("校对医嘱电子签名", chk校对签名.value, glngSys, p住院医嘱发送, blnSetPara)
        Call zlDatabase.SetPara("检查医保审批", chk医保审批.value, glngSys, p住院医嘱发送, blnSetPara)
        Call zlDatabase.SetPara("发送前自动校对", chkAutoVerify.value, glngSys, p住院医嘱发送, blnSetPara)
        Call zlDatabase.SetPara("停止后自动超期收回", chkAutoRoll.value, glngSys, p住院医嘱发送, blnSetPara)
        Call zlDatabase.SetPara("特殊医嘱发送前检查未生效医嘱", chkTurnCheck.value, glngSys, p住院医嘱发送, blnSetPara)
        Call zlDatabase.SetPara("药嘱发送限制结束时间", chkLimit.value, glngSys, p住院医嘱发送, blnSetPara)
        
        '批量操作
        Call zlDatabase.SetPara("批量医嘱校对", chkBat(0).value, glngSys, p住院医嘱发送, blnSetPara)
        Call zlDatabase.SetPara("批量医嘱启停", chkBat(1).value, glngSys, p住院医嘱发送, blnSetPara)
        
        '发药类型不收回
        strValue = ""
        For i = 0 To lst发药类型.ListCount - 1
            If lst发药类型.Selected(i) Then
                strValue = strValue & "," & ZLCommFun.GetNeedName(lst发药类型.List(i))
            End If
        Next
        strValue = Mid(strValue, 2)
        Call zlDatabase.SetPara("发药后不收回", strValue, glngSys, p住院医嘱发送, blnSetPara)
        
        If chkInfo(chk长嘱口服药发送结束时间).value = 1 Then
            strValue = "1|" & dtpEnd.value
        Else
            strValue = 0
        End If
        Call zlDatabase.SetPara("长嘱口服药发送结束时间", strValue, glngSys, p住院医嘱发送, blnSetPara)
    End If
    
    '在输液配置中心发药的病人科室
    bln输液全选 = True
    If lstDept.Enabled Then
        For i = 0 To lstDept.ListCount - 1
            If lstDept.Selected(i) Then
                str输液配置中心 = str输液配置中心 & "," & lstDept.ItemData(i)
            Else
                bln输液全选 = False
            End If
        Next
    End If
    
    If str输液配置中心 <> "" Then
        str输液配置中心 = Mid(str输液配置中心, 2)
    Else
        str输液配置中心 = ","
    End If
    If bln输液全选 Then str输液配置中心 = "*"
    Call zlDatabase.SetPara("在输液配置中心发药的病人科室", str输液配置中心, glngSys, p住院医嘱下达, blnSetup)
    
    '启用输液配置中心的医嘱期效
    If fra期效.Enabled Then
        For i = 0 To 2
            If opt期效(i).value Then
                Call zlDatabase.SetPara("启用输液配置中心的医嘱期效", i & "", glngSys, p住院医嘱下达, blnSetup)
                Exit For
            End If
        Next
    End If
    Call zlDatabase.SetPara("配置中心不接收的静脉营养医嘱在病区配置", chk静脉营养.value, glngSys, p住院医嘱下达, blnSetup)
    Call zlDatabase.SetPara("特殊性质药品允许发送到配制中心", chkWithSelf.value, glngSys, p住院医嘱下达, blnSetup)

    '抗菌药物缺省用药目的
    For i = 0 To 2
        If optPurMed(i).value Then
            Call zlDatabase.SetPara("抗菌药物缺省用药目的", i & "", glngSys, p住院医嘱下达, blnSetup)
            Exit For
        End If
    Next
    
     '药房
    With vsfDrugStore
        For i = .FixedRows To .Rows - 1
            Select Case .TextMatrix(i, .ColIndex("类别"))
            Case "西药房"
                bytType = 0
            Case "成药房"
                bytType = 1
            Case "中药房"
                bytType = 2
            End Select
            If .TextMatrix(i, .ColIndex("可用")) <> 0 Then arr可用药房(bytType) = arr可用药房(bytType) & "," & .RowData(i)
            If .TextMatrix(i, .ColIndex("缺省")) = "√" Then arr缺省药房(bytType) = .RowData(i)
        Next
    End With
    arrTmp = Split("西药房,成药房,中药房", ",")
    For bytType = 0 To UBound(arrTmp)
        Call zlDatabase.SetPara("住院可用" & arrTmp(bytType), Mid(arr可用药房(bytType), 2), glngSys, p住院医嘱下达, blnSetup)
        Call zlDatabase.SetPara("住院缺省" & arrTmp(bytType), arr缺省药房(bytType), glngSys, p住院医嘱下达, blnSetup)
    Next
        
    Call zlDatabase.SetPara("住院缺省发料部门", IIF(cbo卫材.ListIndex = 0, "0", cbo卫材.ItemData(cbo卫材.ListIndex)), glngSys, p住院医嘱下达, blnSetup)
    Call zlDatabase.SetPara("输血申请注意事项", cmdBloodTip.Tag, glngSys, p住院医嘱下达, blnSetup)
    
    '单据打印:0-不打印,1-手工打印,2-自动打印
    If mint场合 <> 2 Then
        Call zlDatabase.SetPara("住院发送单据打印", IIF(optPrint(0).value, 0, IIF(optPrint(1).value, 1, 2)), glngSys, p住院医嘱发送, blnSetPara)
        Call zlDatabase.SetPara("医嘱处理范围", IIF(optBaby(0).value, 0, IIF(optBaby(1).value, 1, 2)), glngSys, p住院医嘱发送, blnSetPara)
    End If

    gblnOK = True
    Unload Me
End Sub

Private Sub cmd校对病区ALL_Click()
    Dim i As Integer
    
    For i = 0 To lst校对病区.ListCount - 1
        lst校对病区.Selected(i) = True
    Next
    lst校对病区.SetFocus
End Sub

Private Sub cmd校对病区Clear_Click()
    Dim i As Integer
    
    For i = 0 To lst校对病区.ListCount - 1
        lst校对病区.Selected(i) = False
    Next
    lst校对病区.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        '表格的checkbox按回车，不转移焦点
        If Not Me.ActiveControl Is vsfDrugStore Then
            Call ZLCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If cmd校对病区ALL.Enabled And cmd校对病区ALL.Visible Then Call cmd校对病区ALL_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If cmd校对病区Clear.Enabled And cmd校对病区Clear.Visible Then Call cmd校对病区Clear_Click
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPar As String, i As Long
    Dim objControl As Control
    Dim bln下达设置 As Boolean, bln发送设置 As Boolean
    Dim objctl As Object, arrTmp() As String
    Dim strDSIDs As String, strDefault As String, lngBackColor As Long, bytLockEdit As Byte
    Dim intType1 As Integer, intType2 As Integer, lngRow As Long
    Dim ctl As Control
    Dim strExcute As String
    
    On Error GoTo errH
    
    gblnOK = False
            
    If mint场合 <> 1 Then
        fra后续处理.Enabled = False
        fra超期收回.Enabled = False
        fraBat.Enabled = False
        fra校对病区.Enabled = False
        cmd校对病区ALL.Enabled = False
        cmd校对病区Clear.Enabled = False
        
        For Each ctl In Me.Controls
            If ctl.Container Is fra后续处理 Then
                ctl.Enabled = False
            ElseIf ctl.Container Is fra超期收回 Then
                ctl.Enabled = False
            ElseIf ctl.Container Is fraBat Then
                ctl.Enabled = False
            ElseIf ctl.Container Is fra校对病区 Then
                ctl.Enabled = False
            End If
        Next
        
                
        If mint场合 = 2 Then    '医技站
            fra入院诊断.Visible = False
            chk出院诊断.Visible = False
            chkStopNurseGrade.Visible = False
            chk停嘱审核.Visible = False
            
            tabPar.TabVisible(1) = False
        End If
    End If
    
    bln下达设置 = InStr(GetInsidePrivs(p住院医嘱下达), "医嘱选项设置") > 0
    bln发送设置 = InStr(GetInsidePrivs(p住院医嘱发送), "医嘱选项设置") > 0
    
    If mint场合 <> 0 Then
        chk关闭医嘱.Enabled = False
    Else
        fra后续处理.Enabled = True
        chk关闭医嘱.Enabled = True
        chk关闭医嘱.value = Val(zlDatabase.GetPara("发送完成后关闭医嘱窗体", glngSys, p住院医嘱下达, "0", Array(chk关闭医嘱), bln发送设置))
        cmdBloodTip.Tag = zlDatabase.GetPara("输血申请注意事项", glngSys, p住院医嘱下达, , Array(cmdBloodTip), bln下达设置)
    End If
    
    '要求输入入院诊断
    strPar = zlDatabase.GetPara("要求输入入院诊断", glngSys, p住院医嘱下达, , Array(chk入院诊断, lst入院诊断), bln下达设置)
    If Not chk入院诊断.Enabled Then lst入院诊断.Tag = "1" '固定标识为不可用
    If strPar <> "" Then
        chk入院诊断.value = 1
        Call chk入院诊断_Click
    End If
    strSQL = "Select 编码,名称 From 诊疗项目类别 Where 编码 Not IN('4','5','6','7','8','9') Union ALL Select '5','药品' From Dual Order by 编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    With lst入院诊断
        Do While Not rsTmp.EOF
            .AddItem rsTmp!编码 & "-" & rsTmp!名称
            .ItemData(.NewIndex) = Asc(rsTmp!编码)
            
            If strPar <> "" Then
                If InStr(strPar, Chr(.ItemData(.NewIndex))) > 0 Then
                    .Selected(.NewIndex) = True
                End If
            End If
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    '启用输液配置中心的医嘱期效
    strPar = zlDatabase.GetPara("启用输液配置中心的医嘱期效", glngSys, p住院医嘱下达, "2", Array(opt期效(0), opt期效(1), opt期效(2)), bln下达设置)
    opt期效(Val(strPar)).value = True
    
    '抗菌药物缺省用药目的
    strPar = zlDatabase.GetPara("抗菌药物缺省用药目的", glngSys, p住院医嘱下达, "1")
    If strPar = "3" Then strPar = "0"
    optPurMed(Val(strPar)).value = True
    
    '在输液配置中心发药的病人科室
    If gstr输液配置中心 <> "" Then
        strPar = zlDatabase.GetPara("在输液配置中心发药的病人科室", glngSys, p住院医嘱下达, "*", Array(lstDept, cmdAllSelect, cmdAllDel), bln下达设置)
        strSQL = "select distinct ID,编码,名称" & _
                    " from 部门表 D,部门性质说明 T" & _
                    " where D.ID=T.部门ID and t.工作性质='临床' And t.服务对象 in(2,3)" & _
                    "       and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                    " order by 编码"
        
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        i = -1
        Do While Not rsTmp.EOF
            lstDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
            lstDept.ItemData(lstDept.NewIndex) = rsTmp!ID
            If strPar = "*" Or InStr("," & strPar & ",", "," & rsTmp!ID & ",") > 0 Then
                lstDept.Selected(lstDept.NewIndex) = True
                If i = -1 Then i = lstDept.NewIndex
            End If
            rsTmp.MoveNext
        Loop
        If i <> -1 Then lstDept.ListIndex = i
        If lstDept.ListIndex = -1 And lstDept.ListCount > 0 Then lstDept.ListIndex = 0
        '配置中心不接收的静脉营养医嘱在病区配置
        chk静脉营养.value = Val(zlDatabase.GetPara("配置中心不接收的静脉营养医嘱在病区配置", glngSys, p住院医嘱下达, "0", Array(chk静脉营养), bln下达设置))
        chkWithSelf.value = Val(zlDatabase.GetPara("特殊性质药品允许发送到配制中心", glngSys, p住院医嘱下达, "0", Array(chkWithSelf), bln下达设置))
    Else
        lstDept.Enabled = False
        fra输液配置.Enabled = False
        cmdAllSelect.Enabled = False
        cmdAllDel.Enabled = False
        chk静脉营养.Enabled = False
        chkWithSelf.Enabled = False
        lstDept.AddItem "没有启用输液配置中心！"
        lstDept.ListIndex = -1
        fra期效.Enabled = False
        For i = 0 To 2
            opt期效(i).Enabled = False
        Next
    End If
    
    dtpEnd.value = "23:59:59"
    
    '自动校对的病区
    If mint场合 = 1 Then
        strPar = zlDatabase.GetPara("自动完成校对计价", glngSys, p住院医嘱发送, , Array(chk校对, lst校对病区, fra校对病区, cmd校对病区ALL, cmd校对病区Clear), bln发送设置)
        If strPar <> "" Then chk校对.value = 1
        Call chk校对_Click
        
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where A.ID=B.部门ID And B.服务对象 in(2,3) And B.工作性质='护理'" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by A.编码"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        i = -1
        Do While Not rsTmp.EOF
            lst校对病区.AddItem rsTmp!编码 & "-" & rsTmp!名称
            lst校对病区.ItemData(lst校对病区.NewIndex) = rsTmp!ID
            If strPar = "*" Or InStr("," & strPar & ",", "," & rsTmp!ID & ",") > 0 Then
                lst校对病区.Selected(lst校对病区.NewIndex) = True
                If i = -1 Then i = lst校对病区.NewIndex
            End If
            rsTmp.MoveNext
        Loop
        If i <> -1 Then lst校对病区.ListIndex = i
        If lst校对病区.ListIndex = -1 And lst校对病区.ListCount > 0 Then lst校对病区.ListIndex = 0
        
    
        
        '不收回的发药类型
        strPar = zlDatabase.GetPara("发药后不收回", glngSys, p住院医嘱发送, , Array(lst发药类型), bln发送设置)
        strSQL = "Select 编码, 名称 From 发药类型 Order by 编码"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        i = -1
        Do While Not rsTmp.EOF
            lst发药类型.AddItem rsTmp!名称
            If strPar <> "" Then
                If InStr("," & strPar & ",", "," & rsTmp!名称 & ",") > 0 Then
                    lst发药类型.Selected(lst发药类型.NewIndex) = True
                    If i = -1 Then i = lst发药类型.NewIndex
                End If
            End If
            rsTmp.MoveNext
        Loop
        If i <> -1 Then lst发药类型.ListIndex = i
        If lst发药类型.ListIndex = -1 And lst发药类型.ListCount > 0 Then lst发药类型.ListIndex = 0
        
        chkTurnCheck.value = Val(zlDatabase.GetPara("特殊医嘱发送前检查未生效医嘱", glngSys, p住院医嘱发送, 0, Array(chkTurnCheck), bln发送设置))
        chkLimit.value = Val(zlDatabase.GetPara("药嘱发送限制结束时间", glngSys, p住院医嘱发送, 0, Array(chkLimit), bln发送设置))
        strPar = zlDatabase.GetPara("长嘱口服药发送结束时间", glngSys, p住院医嘱发送, , Array(chkInfo(chk长嘱口服药发送结束时间)), bln发送设置)
        If InStr(strPar, "|") = 0 Then
            chkInfo(chk长嘱口服药发送结束时间).value = 0
        Else
            chkInfo(chk长嘱口服药发送结束时间).value = Val(Split(strPar, "|")(0))
            If chkInfo(chk长嘱口服药发送结束时间).value = 1 Then
                dtpEnd.value = Format(Split(strPar, "|")(1), "HH:MM:SS")
                dtpEnd.Enabled = True
            End If
        End If
    Else
        chkInfo(chk长嘱口服药发送结束时间).Enabled = False
    End If
    
    chk一次性.value = Val(zlDatabase.GetPara("临嘱缺省一次性", glngSys, p住院医嘱下达, , Array(chk一次性), bln下达设置))
    chk天数.value = Val(zlDatabase.GetPara("医嘱执行天数", glngSys, p住院医嘱下达, , Array(chk天数), bln下达设置))
    chkST.value = Val(zlDatabase.GetPara("自动处理皮试", glngSys, p住院医嘱下达, , Array(chkST), bln下达设置))
    optSTCheck(Val(zlDatabase.GetPara("根据皮试结果限制医嘱发送类型", glngSys, p住院医嘱下达, , Array(lblSTCheck, optSTCheck(0), optSTCheck(1)), bln下达设置))).value = True
    Call chkST_Click
    chk先输单量.value = Val(zlDatabase.GetPara("临嘱先输入单量", glngSys, p住院医嘱下达, , Array(chk先输单量), bln下达设置))
    chk出院诊断.value = Val(zlDatabase.GetPara("要求输入出院诊断", glngSys, p住院医嘱下达, , Array(chk出院诊断), bln下达设置))
    chkStopNurseGrade.value = Val(zlDatabase.GetPara("单独停止护理等级", glngSys, p住院医嘱下达, 1, Array(chkStopNurseGrade), bln下达设置))
    chkAdviceSort.value = Val(zlDatabase.GetPara("医嘱自动排序", glngSys, p住院医嘱下达, 0, Array(chkAdviceSort, cmdAdviceSortSet), bln下达设置))
    chk停嘱审核.value = Val(zlDatabase.GetPara("实习医生停止医嘱需要审核", glngSys, p住院医嘱下达, 0, Array(chk停嘱审核), bln下达设置))
    Call chkAdviceSort_Click
    chk待入住病人医嘱下达.value = Val(zlDatabase.GetPara("允许给待入住病人下达医嘱", glngSys, p住院医嘱下达, 1, Array(chk待入住病人医嘱下达), bln下达设置))
    chk术后医嘱.value = Val(zlDatabase.GetPara("手术完成后下达术后医嘱", glngSys, p住院医嘱下达, , Array(chk术后医嘱), bln下达设置))
    
    strExcute = zlDatabase.GetPara("本科执行自动完成", glngSys, p住院医嘱发送, , Array(chk执行, chkLong, chkShort), bln发送设置)
    chkLong.value = Val(Mid(strExcute, 1, 1))
    chkShort.value = Val(Mid(strExcute, 2, 1))
    
    chk打印.value = Val(zlDatabase.GetPara("自动进入医嘱打印", glngSys, p住院医嘱发送, , Array(chk打印), bln发送设置))
    chk医技.value = Val(zlDatabase.GetPara("医技医嘱后续处理", glngSys, p住院医嘱发送, , Array(chk医技), bln发送设置))
    chk皮试.value = Val(zlDatabase.GetPara("皮试验证身份", glngSys, p住院医嘱发送, , Array(chk皮试), bln发送设置))
    
    i = Val(zlDatabase.GetPara("超期收回产生负数费用", glngSys, p住院医嘱发送, , Array(optRoll(0), optRoll(1)), bln发送设置))
    If i = 1 Then
        optRoll(0).value = True
    Else
        optRoll(1).value = True
    End If
    chk超期审核.value = Val(zlDatabase.GetPara("超期收回费用本科自动审核", glngSys, p住院医嘱发送, , Array(chk超期审核), bln发送设置))
    Call optRoll_Click(IIF(optRoll(0).value, 0, 1))
    
    chk校对签名.value = Val(zlDatabase.GetPara("校对医嘱电子签名", glngSys, p住院医嘱发送, , Array(chk校对签名), bln发送设置))
    chk医保审批.value = Val(zlDatabase.GetPara("检查医保审批", glngSys, p住院医嘱发送, 1, Array(chk医保审批), bln发送设置))
    
    chkAutoVerify.value = Val(zlDatabase.GetPara("发送前自动校对", glngSys, p住院医嘱发送, , Array(chkAutoVerify), bln发送设置))
    chkAutoRoll.value = Val(zlDatabase.GetPara("停止后自动超期收回", glngSys, p住院医嘱发送, , Array(chkAutoRoll), bln发送设置))
    
    '批量操作
    chkBat(0).value = Val(zlDatabase.GetPara("批量医嘱校对", glngSys, p住院医嘱发送, , Array(chkBat(0)), bln发送设置))
    chkBat(1).value = Val(zlDatabase.GetPara("批量医嘱启停", glngSys, p住院医嘱发送, , Array(chkBat(1)), bln发送设置))
    
    '医嘱单打印模式
    If Val(zlDatabase.GetPara("医嘱单打印模式", glngSys, p住院医嘱下达, , Array(optPrintType(0), optPrintType(1)), bln下达设置)) <> 0 Then
        optPrintType(1).value = True
    Else
        optPrintType(0).value = True
    End If
    '输血申请单打印模式
    If Val(zlDatabase.GetPara("输血申请单打印模式", glngSys, p住院医嘱发送, , Array(optBloodPrintType(0), optBloodPrintType(1)), bln下达设置)) <> 1 Then
        optBloodPrintType(0).value = True
    Else
        optBloodPrintType(1).value = True
    End If
    
    '药房与发料部门
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称,B.工作性质 " & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " AND B.部门ID=A.ID And B.服务对象 IN(2,3) and B.工作性质 in('中药房','西药房','成药房','发料部门')" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by 工作性质,编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    With vsfDrugStore
        .Rows = .FixedRows
        .Editable = flexEDKbdMouse
        .MergeCol(.ColIndex("类别")) = True
        .MergeCells = flexMergeFixedOnly
        
        rsTmp.Filter = "工作性质<>'发料部门'"
        If Not rsTmp.EOF Then
            .Rows = .FixedRows + rsTmp.RecordCount
            lngRow = .FixedRows
            arrTmp = Split("西药房,成药房,中药房", ",")
            For i = 0 To UBound(arrTmp)
                rsTmp.Filter = "工作性质='" & arrTmp(i) & "'"
                strDefault = zlDatabase.GetPara("住院缺省" & arrTmp(i), glngSys, p住院医嘱下达, , , , intType1)
                strDSIDs = "," & zlDatabase.GetPara("住院可用" & arrTmp(i), glngSys, p住院医嘱下达, , , , intType2) & ","
                Do While Not rsTmp.EOF
                    .TextMatrix(lngRow, .ColIndex("类别")) = arrTmp(i)
                    .TextMatrix(lngRow, .ColIndex("药房")) = rsTmp!名称
                    .RowData(lngRow) = Val(rsTmp!ID)
                    
                    If Val(rsTmp!ID) = Val(strDefault) Then
                        .TextMatrix(lngRow, .ColIndex("缺省")) = "√"
                        .TextMatrix(lngRow, .ColIndex("可用")) = -1   'true
                    Else
                        .TextMatrix(lngRow, .ColIndex("缺省")) = ""
                        .TextMatrix(lngRow, .ColIndex("可用")) = IIF(InStr(strDSIDs, "," & rsTmp!ID & ",") > 0, -1, 0)
                    End If
                    
                    '缺省单元格
                    'intType-'返回参数类型：1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
                    bytLockEdit = 0
                    If InStr(1, ",1,3,15,", "," & intType1 & ",") > 0 Then
                        lngBackColor = IIF(bln下达设置, VsPubBackColor, &H8000000F)      '授权限控制
                        bytLockEdit = IIF(bln下达设置, 0, 1)
                    ElseIf intType1 = 5 Then
                        lngBackColor = VsPubBackColor       '公共模块,但不授权限控制
                    Else
                        lngBackColor = &H80000005     '正常编辑
                    End If
                    .Cell(flexcpBackColor, lngRow, .ColIndex("缺省")) = lngBackColor
                    .Cell(flexcpData, lngRow, .ColIndex("缺省")) = bytLockEdit
                     
                    '可用单元格
                    bytLockEdit = 0
                    If InStr(1, ",1,3,15,", "," & intType2 & ",") > 0 Then
                        lngBackColor = IIF(bln下达设置, VsPubBackColor, &H8000000F)      '授权限控制
                        bytLockEdit = IIF(bln下达设置, 0, 1)
                    ElseIf intType2 = 5 Then
                        lngBackColor = VsPubBackColor       '公共模块,但不授权限控制
                    Else
                        lngBackColor = &H80000005     '正常编辑
                    End If
                    .Cell(flexcpBackColor, lngRow, .ColIndex("可用")) = lngBackColor
                    .Cell(flexcpData, lngRow, .ColIndex("可用")) = bytLockEdit
                    
                    lngRow = lngRow + 1
                    rsTmp.MoveNext
                Loop
                If lngRow < .Rows - 1 Then  '划分隔线
                    .Select lngRow, .FixedCols, lngRow, .Cols - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
            Next
        End If
    End With
    
    cbo卫材.AddItem "人工选择"
    rsTmp.Filter = "工作性质='发料部门'"
    Do While Not rsTmp.EOF
        cbo卫材.AddItem rsTmp!名称
        cbo卫材.ItemData(cbo卫材.ListCount - 1) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    strPar = zlDatabase.GetPara("住院缺省发料部门", glngSys, p住院医嘱下达, , Array(lbl卫材, cbo卫材), bln下达设置)
    zlControl.CboLocate cbo卫材, strPar, True
        
    
    '整体未启用签名时，不允许设置
    If gintCA = 0 Or Mid(gstrESign, 2, 1) <> "1" Then
        chk校对签名.value = 0
        chk校对签名.Enabled = False
    End If
    
    '单据打印:0-不打印,1-手工打印,2-自动打印
    optPrint(Val(zlDatabase.GetPara("住院发送单据打印", glngSys, p住院医嘱发送, "2", Array(optPrint(0), optPrint(1), optPrint(2)), bln发送设置))).value = True
    
    '医嘱处理范围
    optBaby(Val(zlDatabase.GetPara("医嘱处理范围", glngSys, p住院医嘱发送, "0", Array(optBaby(0), optBaby(1), optBaby(2)), bln发送设置))).value = True
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    
    cmdCancel.Left = Me.ScaleLeft + Me.ScaleWidth - cmdCancel.Width - 200
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mMainPrivs = ""
End Sub

Private Sub lstDept_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    If KeyAscii = vbKeyBack Then Exit Sub
    
    For i = 0 To lstDept.ListCount - 1
        If ZLCommFun.SpellCode(Mid(lstDept.List(i), InStr(lstDept.List(i), "-") + 1)) Like UCase(Chr(KeyAscii)) & "*" Then
            lstDept.ListIndex = i: Exit For
        End If
    Next
End Sub

Private Sub vsfDrugStore_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfDrugStore.ColIndex("可用") Then
        Call Set可用药房(Row, True)
    ElseIf Col = vsfDrugStore.ColIndex("可用") Then
        Call Set缺省药房
    End If
    Cancel = True
End Sub

Private Sub vsfDrugStore_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDrugStore
        Select Case Col
        Case .ColIndex("可用")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case .ColIndex("缺省")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case Else
            Cancel = True
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_DblClick()
    With vsfDrugStore
        If .MouseCol = .ColIndex("缺省") Then
            Call Set缺省药房
        ElseIf .MouseCol = .ColIndex("药房") Then
            Call Set可用药房(.Row, True)
        ElseIf .MouseCol = .ColIndex("可用") And .MouseRow = .FixedRows - 1 Then
            Dim i As Long
            For i = .FixedRows To .Rows - 1
                Call Set可用药房(i)
            Next
        End If
    End With
End Sub
Private Sub vsfDrugStore_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If vsfDrugStore.Col = vsfDrugStore.ColIndex("缺省") Then
            Call Set缺省药房
        End If
    End If
End Sub

Private Sub Set缺省药房()
'功能：设置当前行的缺省药房，同时处理相同类型的其他行的缺省药房
    Dim i As Long
    
    With vsfDrugStore
        If Val("" & .Cell(flexcpData, .Row, .ColIndex("缺省"))) = 0 Then  '该参数允许修改的情况下
            If .TextMatrix(.Row, .ColIndex("缺省")) = "√" Then
                .TextMatrix(.Row, .ColIndex("缺省")) = ""
            Else
                '当没有有权限修改可用时且可用为0（false)时不允许设置缺省
                If Not (Val(.TextMatrix(.Row, .ColIndex("可用"))) = 0 And Val("" & .Cell(flexcpData, .Row, .ColIndex("可用"))) = 1) Then
                    '同类别的其他行取消缺省
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(.Row, .ColIndex("类别")) = .TextMatrix(i, .ColIndex("类别")) Then
                            If .TextMatrix(i, .ColIndex("缺省")) = "√" Then .TextMatrix(i, .ColIndex("缺省")) = ""
                        End If
                    Next
                    .TextMatrix(.Row, .ColIndex("可用")) = -1    '自动设置为可用
                    .TextMatrix(.Row, .ColIndex("缺省")) = "√"
                Else
                    MsgBox "设置当前药房为缺省时，会同时将当前药房设置为可用，" & vbNewLine & "你没有修改可用药房的权限。", vbInformation, gstrSysName
                End If
            End If
        Else
            MsgBox "你没有修改缺省药房的权限。", vbInformation, gstrSysName
        End If
    End With
End Sub

Private Sub Set可用药房(ByVal lngRow As Long, Optional ByVal blnAsk As Boolean = False)
'功能：设置当前行的可用药房，同时处理当前行的缺省药房

    With vsfDrugStore
        If Val("" & .Cell(flexcpData, lngRow, .ColIndex("可用"))) = 0 Then   '该参数允许修改的情况下
            If Val(.TextMatrix(lngRow, .ColIndex("可用"))) = -1 Then
                '当前科室勾选可用
                If Not (Val("" & .Cell(flexcpData, lngRow, .ColIndex("缺省"))) = 1 And .TextMatrix(lngRow, .ColIndex("缺省")) = "√") Then
                    .TextMatrix(lngRow, .ColIndex("可用")) = 0
                    .TextMatrix(lngRow, .ColIndex("缺省")) = ""
                Else
                    If blnAsk Then
                        MsgBox "取消当前药房可用时，会同时取消当前药房缺省，" & vbNewLine & "你没有修改缺省药房的权限。", vbInformation, gstrSysName
                    End If
                End If
            Else
                .TextMatrix(lngRow, .ColIndex("可用")) = -1    '自动设置为可用
            End If
        Else
            If blnAsk Then
                MsgBox "你没有修改可用药房的权限。", vbInformation, gstrSysName
            End If
        End If
    End With
End Sub




