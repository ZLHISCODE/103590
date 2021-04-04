VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmParPacs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "影像参数设置"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11250
   Icon            =   "frmParPacs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11250
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picFunc 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      FillColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   7995
      Left            =   0
      ScaleHeight     =   7995
      ScaleWidth      =   2415
      TabIndex        =   183
      Top             =   0
      Width           =   2415
      Begin VB.PictureBox picTPL 
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   0
         ScaleHeight     =   6135
         ScaleWidth      =   2250
         TabIndex        =   185
         Top             =   0
         Width           =   2250
         Begin XtremeSuiteControls.TaskPanel tplFunc 
            Height          =   5250
            Left            =   0
            TabIndex        =   186
            Top             =   720
            Width           =   2205
            _Version        =   589884
            _ExtentX        =   3889
            _ExtentY        =   9260
            _StockProps     =   64
            Behaviour       =   1
            ItemLayout      =   2
            HotTrackStyle   =   3
         End
         Begin XtremeCommandBars.ImageManager imgFunc 
            Left            =   1800
            Top             =   360
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
            Icons           =   "frmParPacs.frx":6852
         End
         Begin XtremeSuiteControls.ShortcutCaption sccFunc 
            Height          =   300
            Left            =   0
            TabIndex        =   187
            Top             =   0
            Width           =   2200
            _Version        =   589884
            _ExtentX        =   3881
            _ExtentY        =   529
            _StockProps     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            Alignment       =   1
         End
      End
      Begin VB.PictureBox picVbar 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         FillColor       =   &H8000000A&
         Height          =   5820
         Left            =   2280
         MousePointer    =   9  'Size W E
         ScaleHeight     =   5820
         ScaleWidth      =   45
         TabIndex        =   184
         Top             =   120
         Width           =   45
      End
      Begin XtremeSuiteControls.ShortcutBar scbFunc 
         Height          =   6765
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   2400
         _Version        =   589884
         _ExtentX        =   4233
         _ExtentY        =   11933
         _StockProps     =   64
      End
      Begin XtremeCommandBars.ImageManager imgType 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         Icons           =   "frmParPacs.frx":ADF2
      End
   End
   Begin VB.PictureBox PicBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   590
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   11250
      TabIndex        =   166
      Top             =   7995
      Width           =   11250
      Begin VB.CommandButton cmdApply 
         Caption         =   "应用(&A)"
         Height          =   350
         Left            =   10200
         TabIndex        =   309
         Top             =   120
         Width           =   1100
      End
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   1
         Left            =   4700
         TabIndex        =   189
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   0
         Left            =   2400
         TabIndex        =   182
         Top             =   120
         Width           =   1200
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   60
         TabIndex        =   180
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   9000
         TabIndex        =   179
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   7845
         TabIndex        =   178
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   6000
         TabIndex        =   190
         Top             =   165
         Width           =   2055
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "科室查找(&F)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   188
         Top             =   165
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "参数查找(&S)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   181
         Top             =   168
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab stabDesign 
      Height          =   7995
      Left            =   2400
      TabIndex        =   191
      TabStop         =   0   'False
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   14102
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   12
      TabHeight       =   520
      TabCaption(0)   =   "影像流程设置"
      TabPicture(0)   =   "frmParPacs.frx":14466
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "picPar(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "影像医技设置"
      TabPicture(1)   =   "frmParPacs.frx":14482
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "picPar(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "影像采集设置"
      TabPicture(2)   =   "frmParPacs.frx":1449E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picPar(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "影像病理设置"
      TabPicture(3)   =   "frmParPacs.frx":144BA
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picPar(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "病理归档设置"
      TabPicture(4)   =   "frmParPacs.frx":144D6
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picPar(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "病理借还设置"
      TabPicture(5)   =   "frmParPacs.frx":144F2
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "picPar(5)"
      Tab(5).ControlCount=   1
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7575
         Index           =   0
         Left            =   -75000
         ScaleHeight     =   7575
         ScaleWidth      =   8895
         TabIndex        =   197
         Top             =   360
         Width           =   8895
         Begin VB.ComboBox cmbDept 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   120
            Width           =   2055
         End
         Begin VB.TextBox txtLab 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   210
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   198
            Text            =   "影像科室"
            Top             =   165
            Width           =   735
         End
         Begin TabDlg.SSTab stabWorkFlow 
            Height          =   7095
            Left            =   120
            TabIndex        =   199
            Top             =   480
            Width           =   8715
            _ExtentX        =   15372
            _ExtentY        =   12515
            _Version        =   393216
            Style           =   1
            Tabs            =   7
            TabsPerRow      =   7
            TabHeight       =   520
            TabCaption(0)   =   "工作流设置"
            TabPicture(0)   =   "frmParPacs.frx":1450E
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fra(2)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "fra(3)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "fra(4)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "fra(0)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "fra(27)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "chkPreView"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).ControlCount=   6
            TabCaption(1)   =   "执行间设置"
            TabPicture(1)   =   "frmParPacs.frx":1452A
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "fra(13)"
            Tab(1).Control(1)=   "cmdAdd"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "cmdDel"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "cmdSave"
            Tab(1).Control(4)=   "cmdRestore"
            Tab(1).ControlCount=   5
            TabCaption(2)   =   "登记录入设置"
            TabPicture(2)   =   "frmParPacs.frx":14546
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "fra(14)"
            Tab(2).Control(1)=   "fra(15)"
            Tab(2).ControlCount=   2
            TabCaption(3)   =   "分组排队设置"
            TabPicture(3)   =   "frmParPacs.frx":14562
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "fra(17)"
            Tab(3).Control(1)=   "fra(16)"
            Tab(3).Control(2)=   "chkUseQueue"
            Tab(3).ControlCount=   3
            TabCaption(4)   =   "报告编辑器设置"
            TabPicture(4)   =   "frmParPacs.frx":1457E
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "fra(23)"
            Tab(4).Control(1)=   "fra(22)"
            Tab(4).Control(2)=   "fra(21)"
            Tab(4).Control(3)=   "fra(20)"
            Tab(4).Control(4)=   "fra(19)"
            Tab(4).Control(5)=   "fra(24)"
            Tab(4).Control(6)=   "fra(5)"
            Tab(4).ControlCount=   7
            TabCaption(5)   =   "检查列表设置"
            TabPicture(5)   =   "frmParPacs.frx":1459A
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "cmdDefault"
            Tab(5).Control(1)=   "fra(28)"
            Tab(5).ControlCount=   2
            TabCaption(6)   =   "检查号设置"
            TabPicture(6)   =   "frmParPacs.frx":145B6
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "fra(1)"
            Tab(6).ControlCount=   1
            Begin VB.CheckBox chkPreView 
               Caption         =   "启用缩略图预览"
               Height          =   375
               Left            =   840
               TabIndex        =   380
               Top             =   4960
               Width           =   1575
            End
            Begin VB.Frame fra 
               Height          =   1485
               Index           =   5
               Left            =   -74280
               TabIndex        =   355
               Top             =   1080
               Width           =   7215
               Begin VB.Frame fra 
                  Caption         =   "录入时机"
                  Height          =   1150
                  Index           =   6
                  Left            =   4920
                  TabIndex        =   372
                  Top             =   240
                  Width           =   2055
                  Begin VB.OptionButton optResultInput 
                     Caption         =   "报告打印前"
                     Height          =   240
                     Index           =   2
                     Left            =   210
                     TabIndex        =   375
                     Top             =   810
                     Width           =   1290
                  End
                  Begin VB.OptionButton optResultInput 
                     Caption         =   "审核签名后"
                     Height          =   240
                     Index           =   1
                     Left            =   210
                     TabIndex        =   374
                     Top             =   525
                     Width           =   1230
                  End
                  Begin VB.OptionButton optResultInput 
                     Caption         =   "诊断签名后"
                     Height          =   240
                     Index           =   0
                     Left            =   210
                     TabIndex        =   373
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   1215
                  End
               End
               Begin VB.TextBox txtImageLevel 
                  Height          =   270
                  Left            =   3690
                  TabIndex        =   364
                  Text            =   "甲,乙"
                  ToolTipText     =   "用于评定影像质量的登记，最多四个等级"
                  Top             =   990
                  Width           =   1035
               End
               Begin VB.TextBox txtReportLevel 
                  Height          =   270
                  Left            =   3690
                  TabIndex        =   363
                  Text            =   "甲,乙"
                  Top             =   600
                  Width           =   1035
               End
               Begin VB.CheckBox chkImageLevel 
                  Caption         =   "影像质量等级"
                  Height          =   180
                  Left            =   2280
                  TabIndex        =   362
                  Top             =   1035
                  Width           =   1410
               End
               Begin VB.CheckBox chkReportLevel 
                  Caption         =   "报告质量等级"
                  Height          =   180
                  Left            =   2280
                  TabIndex        =   361
                  Top             =   657
                  Width           =   1410
               End
               Begin VB.CheckBox chkConformDetermine 
                  Caption         =   "符合情况判断"
                  Height          =   180
                  Left            =   2280
                  TabIndex        =   360
                  ToolTipText     =   "激活符合情况功能和菜单"
                  Top             =   280
                  Width           =   1455
               End
               Begin VB.Frame fra 
                  Height          =   1125
                  Index           =   9
                  Left            =   120
                  TabIndex        =   356
                  Top             =   270
                  Width           =   2055
                  Begin VB.CheckBox chkDefaultPosi 
                     Caption         =   "诊断结果默认阳性"
                     Height          =   300
                     Left            =   120
                     TabIndex        =   359
                     ToolTipText     =   "弹出阴阳性选择窗口，默认选择阳性。"
                     Top             =   300
                     Width           =   1815
                  End
                  Begin VB.CheckBox chkReportAfterResult 
                     Caption         =   "无诊断内容为阴性"
                     Height          =   180
                     Left            =   120
                     TabIndex        =   358
                     ToolTipText     =   "书写报告时，没有录入诊断，则默认记录为阴性。"
                     Top             =   720
                     Width           =   1740
                  End
                  Begin VB.CheckBox chkIgnorePosi 
                     Caption         =   "忽略结果的阴阳性"
                     Height          =   180
                     Left            =   120
                     TabIndex        =   357
                     ToolTipText     =   "不记录和处理阴阳性。"
                     Top             =   0
                     Width           =   1800
                  End
               End
            End
            Begin VB.Frame fra 
               Height          =   1695
               Index           =   27
               Left            =   720
               TabIndex        =   354
               Top             =   5040
               Width           =   4215
               Begin VB.OptionButton optClickPreview 
                  Caption         =   "鼠标单击时预览图像"
                  Height          =   375
                  Left            =   240
                  TabIndex        =   379
                  Top             =   1200
                  Width           =   1935
               End
               Begin VB.OptionButton optMovePreview 
                  Caption         =   "鼠标移动时预览图像"
                  Height          =   375
                  Left            =   240
                  TabIndex        =   378
                  Top             =   360
                  Width           =   2055
               End
               Begin VB.TextBox txtDelayTime 
                  Height          =   270
                  Left            =   2880
                  MaxLength       =   2
                  TabIndex        =   377
                  ToolTipText     =   "0表示不自动关闭"
                  Top             =   795
                  Width           =   495
               End
               Begin VB.Label lblDelayTime 
                  Caption         =   "移动预览时自动关闭延时时间       秒"
                  Height          =   180
                  Left            =   480
                  TabIndex        =   376
                  Top             =   840
                  Width           =   3240
               End
            End
            Begin VB.Frame fra 
               Height          =   6375
               Index           =   1
               Left            =   -74880
               TabIndex        =   314
               Top             =   480
               Width           =   8385
               Begin VB.CheckBox chkCheckMaxNo 
                  Caption         =   "提取实际最大号码"
                  Height          =   300
                  Left            =   240
                  TabIndex        =   348
                  ToolTipText     =   "以实际最大号码为基础顺序编号。如果选择“前缀”，“分隔符”，“年月日”，或者数据库中有字符型检查号，则禁止“提取实际最大号码”。"
                  Top             =   5880
                  Width           =   1935
               End
               Begin VB.CheckBox chkChangeNO 
                  Caption         =   "允许手工调整检查号"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   347
                  ToolTipText     =   "允许根据实际需要，手动修改检查号。"
                  Top             =   5040
                  Width           =   1935
               End
               Begin VB.CheckBox chkCanOverWrite 
                  Caption         =   "允许检查号重复"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   346
                  ToolTipText     =   "允许登记病人的检查号出现重复，当选择“患者检查号保持不变”时，需要允许检查号重复。"
                  Top             =   5460
                  Width           =   1935
               End
               Begin VB.Frame fra 
                  Caption         =   "检查号一致性"
                  Height          =   4290
                  Index           =   12
                  Left            =   120
                  TabIndex        =   335
                  Top             =   360
                  Width           =   4000
                  Begin VB.OptionButton OptCode 
                     Caption         =   "每次检查用新检查号"
                     Height          =   180
                     Index           =   0
                     Left            =   120
                     TabIndex        =   345
                     ToolTipText     =   "报到时产生新的检查号。"
                     Top             =   360
                     Value           =   -1  'True
                     Width           =   1920
                  End
                  Begin VB.OptionButton OptCode 
                     Caption         =   "患者检查号保持不变"
                     Height          =   180
                     Index           =   1
                     Left            =   120
                     TabIndex        =   344
                     ToolTipText     =   "同一个患者，报到时保持检查号不变。"
                     Top             =   2520
                     Width           =   1935
                  End
                  Begin VB.Frame fra 
                     Height          =   1400
                     Index           =   7
                     Left            =   480
                     TabIndex        =   340
                     Top             =   2760
                     Width           =   3300
                     Begin VB.OptionButton OptUnicode 
                        Caption         =   "本检查类别统一"
                        Height          =   240
                        Index           =   0
                        Left            =   240
                        TabIndex        =   343
                        ToolTipText     =   "检查类别相同，保持检查号不变。"
                        Top             =   300
                        Value           =   -1  'True
                        Width           =   1590
                     End
                     Begin VB.OptionButton OptUnicode 
                        Caption         =   "本科室统一"
                        Height          =   210
                        Index           =   1
                        Left            =   240
                        TabIndex        =   342
                        ToolTipText     =   "科室相同，保持检查号不变。"
                        Top             =   705
                        Width           =   1290
                     End
                     Begin VB.OptionButton optUsePatientID 
                        Caption         =   "全院统一（使用病人ID）"
                        Height          =   210
                        Left            =   240
                        TabIndex        =   341
                        ToolTipText     =   "使用病人ID作为检查号。"
                        Top             =   1080
                        Width           =   2370
                     End
                  End
                  Begin VB.Frame Frame2 
                     Height          =   1400
                     Left            =   480
                     TabIndex        =   336
                     Top             =   720
                     Width           =   3300
                     Begin VB.OptionButton OptBuildcode 
                        Caption         =   "相同检查类别自动递增"
                        Height          =   210
                        Index           =   0
                        Left            =   240
                        TabIndex        =   339
                        ToolTipText     =   "检查号以检查类别为基础，自动递增。"
                        Top             =   300
                        Value           =   -1  'True
                        Width           =   2130
                     End
                     Begin VB.OptionButton OptBuildcode 
                        Caption         =   "本科室内自动递增"
                        Height          =   210
                        Index           =   1
                        Left            =   240
                        TabIndex        =   338
                        ToolTipText     =   "检查号以科室为基础，自动递增。"
                        Top             =   690
                        Width           =   1740
                     End
                     Begin VB.OptionButton optUseAdviceID 
                        Caption         =   "使用医嘱ID"
                        Height          =   210
                        Left            =   240
                        TabIndex        =   337
                        ToolTipText     =   "使用医嘱ID作为检查号。"
                        Top             =   1080
                        Width           =   1740
                     End
                  End
               End
               Begin VB.Frame Frame3 
                  Caption         =   "检查编号设置"
                  Height          =   5895
                  Left            =   4200
                  TabIndex        =   315
                  Top             =   360
                  Width           =   4000
                  Begin VB.ComboBox cboDelimeter 
                     BeginProperty Font 
                        Name            =   "宋体"
                        Size            =   10.5
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Index           =   2
                     Left            =   1280
                     Style           =   2  'Dropdown List
                     TabIndex        =   350
                     Top             =   3694
                     Width           =   2500
                  End
                  Begin VB.ComboBox cboDelimeter 
                     BeginProperty Font 
                        Name            =   "宋体"
                        Size            =   10.5
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Index           =   1
                     Left            =   1280
                     Style           =   2  'Dropdown List
                     TabIndex        =   349
                     Top             =   1942
                     Width           =   2500
                  End
                  Begin VB.CheckBox chkPreText 
                     Caption         =   "前缀"
                     Height          =   375
                     Left            =   240
                     TabIndex        =   334
                     ToolTipText     =   "检查号增加固定前缀。"
                     Top             =   360
                     Width           =   1215
                  End
                  Begin VB.Frame frmPreText 
                     Height          =   1100
                     Left            =   480
                     TabIndex        =   330
                     Top             =   720
                     Width           =   3300
                     Begin VB.OptionButton optPreText 
                        Caption         =   "影像类别"
                        Height          =   255
                        Index           =   0
                        Left            =   240
                        TabIndex        =   333
                        ToolTipText     =   "使用检查的影像类别作为前缀。"
                        Top             =   240
                        Width           =   1455
                     End
                     Begin VB.OptionButton optPreText 
                        Caption         =   "自由文本"
                        Height          =   255
                        Index           =   1
                        Left            =   240
                        TabIndex        =   332
                        ToolTipText     =   "使用自由文本作为前缀。"
                        Top             =   600
                        Value           =   -1  'True
                        Width           =   1215
                     End
                     Begin VB.TextBox txtPreText 
                        Height          =   375
                        Left            =   1440
                        MaxLength       =   10
                        TabIndex        =   331
                        ToolTipText     =   "前缀可以设置10个字符"
                        Top             =   540
                        Width           =   1600
                     End
                  End
                  Begin VB.CheckBox chkDelimiter 
                     Caption         =   "分隔符1"
                     Height          =   375
                     Index           =   1
                     Left            =   240
                     TabIndex        =   329
                     ToolTipText     =   "前缀之后的分隔符。"
                     Top             =   1920
                     Width           =   975
                  End
                  Begin VB.CheckBox chkDelimiter 
                     Caption         =   "分隔符2"
                     Height          =   375
                     Index           =   2
                     Left            =   240
                     TabIndex        =   328
                     ToolTipText     =   "年月日之后的分隔符。"
                     Top             =   3672
                     Width           =   975
                  End
                  Begin VB.CheckBox chkYear 
                     Caption         =   "年"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   327
                     ToolTipText     =   "在检查号之前增加当前年。"
                     Top             =   2448
                     Width           =   735
                  End
                  Begin VB.Frame frmYear 
                     Height          =   500
                     Left            =   1280
                     TabIndex        =   324
                     Top             =   2325
                     Width           =   2500
                     Begin VB.OptionButton optYear 
                        Caption         =   "四位"
                        Height          =   350
                        Index           =   0
                        Left            =   240
                        TabIndex        =   326
                        ToolTipText     =   "四位年份，比如“2008”。"
                        Top             =   120
                        Value           =   -1  'True
                        Width           =   735
                     End
                     Begin VB.OptionButton optYear 
                        Caption         =   "两位"
                        Height          =   350
                        Index           =   1
                        Left            =   1320
                        TabIndex        =   325
                        ToolTipText     =   "两位年份，比如“08”。"
                        Top             =   120
                        Width           =   735
                     End
                  End
                  Begin VB.CheckBox chkMonth 
                     Caption         =   "月"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   323
                     ToolTipText     =   "在检查号之前增加当前月。"
                     Top             =   2856
                     Width           =   735
                  End
                  Begin VB.CheckBox chkDay 
                     Caption         =   "日"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   322
                     ToolTipText     =   "在检查号之前增加当前日。"
                     Top             =   3264
                     Width           =   615
                  End
                  Begin VB.CheckBox chkNumber 
                     Caption         =   "顺序号"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   321
                     ToolTipText     =   "顺序号是默认必须要选择的"
                     Top             =   4200
                     Value           =   1  'Checked
                     Width           =   975
                  End
                  Begin VB.Frame Frame6 
                     Height          =   1335
                     Left            =   480
                     TabIndex        =   316
                     Top             =   4440
                     Width           =   3300
                     Begin VB.TextBox txtStartNum 
                        Height          =   375
                        Left            =   1440
                        MaxLength       =   4
                        TabIndex        =   319
                        ToolTipText     =   "检查号编号的起始号码，小于4位。"
                        Top             =   300
                        Width           =   1600
                     End
                     Begin VB.CheckBox chkFixedLen 
                        Caption         =   "固定位数"
                        Height          =   255
                        Left            =   240
                        TabIndex        =   318
                        ToolTipText     =   "检查号按照固定位数编号，前面补零。"
                        Top             =   840
                        Width           =   1095
                     End
                     Begin VB.TextBox txtFixedLen 
                        Height          =   375
                        Left            =   1440
                        MaxLength       =   2
                        TabIndex        =   317
                        ToolTipText     =   "固定位数小于18位"
                        Top             =   780
                        Width           =   1600
                     End
                     Begin VB.Label lblStartNum 
                        Caption         =   "起始号码"
                        Height          =   255
                        Left            =   240
                        TabIndex        =   320
                        ToolTipText     =   "检查号编号的起始号码。"
                        Top             =   360
                        Width           =   975
                     End
                  End
               End
            End
            Begin VB.Frame fra 
               Height          =   2715
               Index           =   0
               Left            =   720
               TabIndex        =   200
               Top             =   480
               Width           =   7335
               Begin VB.CheckBox chkSetFocusWithReport 
                  Caption         =   "检查切换时定位报告编辑"
                  Height          =   180
                  Left            =   2160
                  TabIndex        =   301
                  ToolTipText     =   "切换至报告页面时是否定位报告编辑"
                  Top             =   2400
                  Width           =   2415
               End
               Begin VB.CheckBox chkCompletePrint 
                  Caption         =   "终审后直接打印"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   297
                  ToolTipText     =   "终审签名后直接打印报告。"
                  Top             =   2100
                  Width           =   1680
               End
               Begin VB.CheckBox chkFinallyCompleteCommit 
                  Caption         =   "终审后直接完成"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   296
                  ToolTipText     =   "报告审核后，该检查自动完成，仅适用于报告文档编辑器。"
                  Top             =   1173
                  Width           =   1815
               End
               Begin VB.Frame Frame11 
                  Caption         =   "医生站查看报告"
                  Height          =   615
                  Left            =   4680
                  TabIndex        =   294
                  ToolTipText     =   "仅适用于报告文档编辑器。"
                  Top             =   1800
                  Width           =   2415
                  Begin VB.ComboBox cboViewReport 
                     Height          =   300
                     ItemData        =   "frmParPacs.frx":145D2
                     Left            =   240
                     List            =   "frmParPacs.frx":145DC
                     Style           =   2  'Dropdown List
                     TabIndex        =   295
                     ToolTipText     =   "仅适用于报告文档编辑器。"
                     Top             =   240
                     Width           =   1935
                  End
               End
               Begin VB.CheckBox chkAddons 
                  Caption         =   "显示附加主述"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   289
                  ToolTipText     =   "在登记报到窗口显示附加主述一项"
                  Top             =   2360
                  Width           =   1935
               End
               Begin VB.CheckBox chkReagent 
                  Caption         =   "显示造影剂"
                  Height          =   180
                  Left            =   2160
                  TabIndex        =   288
                  ToolTipText     =   "在登记报到窗口显示造影剂一项，病理工作站不显示"
                  Top             =   2040
                  Width           =   1935
               End
               Begin VB.TextBox txtRefreshInterval 
                  Enabled         =   0   'False
                  Height          =   270
                  Left            =   6480
                  TabIndex        =   12
                  Text            =   "0"
                  Top             =   778
                  Width           =   390
               End
               Begin VB.TextBox TxtLike 
                  Enabled         =   0   'False
                  Height          =   270
                  Left            =   6600
                  MaxLength       =   2
                  TabIndex        =   8
                  ToolTipText     =   "0天则无时间限制,模糊查找所有病人"
                  Top             =   470
                  Width           =   270
               End
               Begin VB.CheckBox chkAutoSendWorkList 
                  Caption         =   "报到时自动发送WorkList"
                  Height          =   180
                  Left            =   2160
                  TabIndex        =   19
                  Top             =   1740
                  Value           =   1  'Checked
                  Width           =   2532
               End
               Begin VB.TextBox txtViewHistoryImageDays 
                  Height          =   270
                  Left            =   6600
                  MaxLength       =   2
                  TabIndex        =   20
                  Text            =   "1"
                  Top             =   1395
                  Width           =   465
               End
               Begin VB.CheckBox chkCanViewImage 
                  Caption         =   "采图后医生站即可观片"
                  Height          =   180
                  Left            =   2160
                  TabIndex        =   16
                  ToolTipText     =   "采集图像后，在没有检查完成的情况下，医生站也可进行观片。"
                  Top             =   1443
                  Width           =   2160
               End
               Begin VB.CheckBox ChkFinishCommit 
                  Caption         =   "无报告完成后直接完成"
                  Height          =   180
                  Left            =   2160
                  TabIndex        =   14
                  ToolTipText     =   "点击无报告完成后，该检查自动完成。"
                  Top             =   1146
                  Width           =   2160
               End
               Begin VB.CheckBox chkPrintCommit 
                  Caption         =   "打印后直接完成"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   5
                  ToolTipText     =   "打印报告后，该检查自动完成。"
                  Top             =   561
                  Width           =   1815
               End
               Begin VB.CheckBox ChkCompleteCommit 
                  Caption         =   "审核后直接完成"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   9
                  ToolTipText     =   "报告审核后，该检查自动完成。"
                  Top             =   867
                  Width           =   1935
               End
               Begin VB.CheckBox chkSample 
                  Caption         =   "申请登记后直接报到"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   15
                  ToolTipText     =   "登记与报到同时进行。"
                  Top             =   1785
                  Width           =   1935
               End
               Begin VB.TextBox Txt默认天数 
                  Height          =   270
                  Left            =   6120
                  MaxLength       =   2
                  TabIndex        =   18
                  Text            =   "2"
                  Top             =   1086
                  Width           =   945
               End
               Begin VB.CheckBox chkReportAfterImging 
                  Caption         =   "有图像才能写报告"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   2
                  ToolTipText     =   "必须采集图像后才能编写影像报告。"
                  Top             =   255
                  Width           =   2040
               End
               Begin VB.CheckBox chkPrintNeedComplete 
                  Caption         =   "平诊检查需审核才能打报告"
                  Height          =   180
                  Left            =   2160
                  TabIndex        =   10
                  ToolTipText     =   "平诊检查必须经过审核后才能打印报告。"
                  Top             =   849
                  Width           =   2505
               End
               Begin VB.CheckBox chkTechReportSame 
                  Caption         =   "只能填写自己检查的报告"
                  Height          =   180
                  Left            =   2160
                  TabIndex        =   6
                  ToolTipText     =   "只有自己采集图像的检查，才能书写报告。"
                  Top             =   552
                  Width           =   2295
               End
               Begin VB.CheckBox chkWriteCapDoctor 
                  Caption         =   "采集图像者为检查技师"
                  Height          =   180
                  Left            =   4680
                  TabIndex        =   4
                  ToolTipText     =   "采集图像之后，自动将当前用户记录成检查技师。"
                  Top             =   240
                  Width           =   2400
               End
               Begin VB.CheckBox chkLocalizerBackward 
                  Caption         =   "定位片后置"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   13
                  ToolTipText     =   "将定位片放到最后一个序列显示。"
                  Top             =   1479
                  Width           =   1320
               End
               Begin VB.CheckBox chkRefreshInterval 
                  Caption         =   "病人自动刷新间隔      秒"
                  Height          =   180
                  Left            =   4680
                  TabIndex        =   11
                  ToolTipText     =   "病人检查列表会间隔N秒自动刷新。"
                  Top             =   847
                  Width           =   2500
               End
               Begin VB.CheckBox chkAllPatientIsOutside 
                  Caption         =   "所有登记病人标记为外来"
                  Height          =   180
                  Left            =   2160
                  TabIndex        =   3
                  ToolTipText     =   "凡在该工作站中登记的病人均标记为外来病人。"
                  Top             =   255
                  Width           =   2295
               End
               Begin VB.CheckBox ChkLike 
                  Caption         =   "登记时姓名模糊查找    天"
                  Height          =   195
                  Left            =   4680
                  TabIndex        =   7
                  ToolTipText     =   "登记时支持对姓名进行模糊查找，可以查找到N天内的信息。"
                  Top             =   536
                  Width           =   2500
               End
               Begin VB.Label lab 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "自动打开历史图像天数"
                  Height          =   180
                  Index           =   1
                  Left            =   4680
                  TabIndex        =   201
                  ToolTipText     =   "如果当前检查没有图像，则自动打开指定时间段内的历史图像"
                  Top             =   1440
                  Width           =   1800
               End
               Begin VB.Label lab 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "默认记录查询天数"
                  Height          =   180
                  Index           =   0
                  Left            =   4680
                  TabIndex        =   17
                  ToolTipText     =   "检查列表中默认显示对应天数内的检查记录。"
                  Top             =   1143
                  Width           =   1440
               End
            End
            Begin VB.Frame fra 
               Caption         =   "报告文档编辑器设置"
               Height          =   4335
               Index           =   24
               Left            =   -74280
               TabIndex        =   213
               Top             =   2640
               Width           =   7245
               Begin VB.Frame fra 
                  Caption         =   "历史报告查看编辑器"
                  Height          =   615
                  Index           =   25
                  Left            =   240
                  TabIndex        =   214
                  Top             =   360
                  Width           =   6855
                  Begin VB.OptionButton optHistoryReportEditor 
                     Caption         =   "PACS报告编辑器"
                     Height          =   255
                     Index           =   1
                     Left            =   4080
                     TabIndex        =   216
                     Top             =   240
                     Width           =   1695
                  End
                  Begin VB.OptionButton optHistoryReportEditor 
                     Caption         =   "电子病历编辑器"
                     Height          =   255
                     Index           =   0
                     Left            =   360
                     TabIndex        =   215
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   1695
                  End
               End
            End
            Begin VB.CheckBox chkUseQueue 
               Caption         =   "启用排队叫号"
               Height          =   180
               Left            =   -74160
               TabIndex        =   286
               ToolTipText     =   "激活排队叫号功能，仅限于影像采集站和影像医技站。"
               Top             =   5400
               Width           =   1455
            End
            Begin VB.Frame fra 
               Height          =   5805
               Index           =   13
               Left            =   -74280
               TabIndex        =   243
               Top             =   480
               Width           =   7305
               Begin VB.TextBox txtNoPrefix 
                  Height          =   300
                  Left            =   6075
                  MaxLength       =   20
                  TabIndex        =   37
                  Top             =   5340
                  Width           =   1050
               End
               Begin VB.ComboBox cboDevice 
                  Height          =   300
                  Left            =   3240
                  Style           =   2  'Dropdown List
                  TabIndex        =   36
                  Top             =   5340
                  Width           =   1830
               End
               Begin VB.TextBox txtName 
                  Height          =   300
                  Left            =   840
                  MaxLength       =   20
                  TabIndex        =   35
                  Top             =   5340
                  Width           =   1635
               End
               Begin MSComctlLib.ImageList img16 
                  Left            =   4320
                  Top             =   600
                  _ExtentX        =   1005
                  _ExtentY        =   1005
                  BackColor       =   -2147483643
                  ImageWidth      =   16
                  ImageHeight     =   16
                  MaskColor       =   12632256
                  _Version        =   393216
                  BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                     NumListImages   =   1
                     BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "frmParPacs.frx":14600
                        Key             =   "Room"
                     EndProperty
                  EndProperty
               End
               Begin MSComctlLib.ListView lvwRoom 
                  Height          =   4695
                  Left            =   120
                  TabIndex        =   34
                  Top             =   480
                  Width           =   7065
                  _ExtentX        =   12462
                  _ExtentY        =   8281
                  View            =   3
                  Arrange         =   1
                  LabelEdit       =   1
                  Sorted          =   -1  'True
                  LabelWrap       =   -1  'True
                  HideSelection   =   0   'False
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  _Version        =   393217
                  Icons           =   "img16"
                  SmallIcons      =   "img16"
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  Appearance      =   1
                  NumItems        =   0
               End
               Begin VB.Label lab 
                  AutoSize        =   -1  'True
                  Caption         =   "号码前缀"
                  Height          =   180
                  Index           =   6
                  Left            =   5250
                  TabIndex        =   247
                  Top             =   5400
                  Width           =   720
               End
               Begin VB.Label lab 
                  Caption         =   "设备(&D)"
                  Height          =   180
                  Index           =   5
                  Left            =   2565
                  TabIndex        =   246
                  Top             =   5400
                  Width           =   630
               End
               Begin VB.Label lab 
                  AutoSize        =   -1  'True
                  Caption         =   "名称(&N)"
                  Height          =   180
                  Index           =   4
                  Left            =   150
                  TabIndex        =   245
                  Top             =   5400
                  Width           =   630
               End
               Begin VB.Label lab 
                  Caption         =   "设置本科室的执行间后，才能有效进行执行的安排。"
                  Height          =   210
                  Index           =   3
                  Left            =   150
                  TabIndex        =   244
                  Top             =   210
                  Width           =   4140
                  WordWrap        =   -1  'True
               End
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "新增(&A)"
               Height          =   345
               Left            =   -71760
               Picture         =   "frmParPacs.frx":14B9A
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   6510
               Width           =   1100
            End
            Begin VB.CommandButton cmdDel 
               Caption         =   "删除(&D)"
               Height          =   345
               Left            =   -70560
               Picture         =   "frmParPacs.frx":14CE4
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   6510
               Width           =   1100
            End
            Begin VB.CommandButton cmdSave 
               Caption         =   "保存(&S)"
               Height          =   345
               Left            =   -68160
               TabIndex        =   41
               Top             =   6510
               Width           =   1100
            End
            Begin VB.CommandButton cmdRestore 
               Caption         =   "恢复(&R)"
               Height          =   345
               Left            =   -69360
               TabIndex        =   40
               Top             =   6510
               Width           =   1100
            End
            Begin VB.Frame fra 
               Caption         =   "光标跳过项目选择"
               Height          =   2010
               Index           =   14
               Left            =   -74280
               TabIndex        =   242
               Top             =   480
               Width           =   7300
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "检查技师二"
                  Height          =   180
                  Index           =   24
                  Left            =   6060
                  TabIndex        =   65
                  Top             =   1320
                  Width           =   1220
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "附加主述"
                  Height          =   180
                  Index           =   23
                  Left            =   180
                  TabIndex        =   66
                  Top             =   1590
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "检查技师"
                  Height          =   180
                  Index           =   22
                  Left            =   4800
                  TabIndex        =   64
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "造影剂"
                  Height          =   180
                  Index           =   21
                  Left            =   3525
                  TabIndex        =   63
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "出生日期"
                  Height          =   180
                  Index           =   20
                  Left            =   3525
                  TabIndex        =   45
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "检查时间"
                  Height          =   180
                  Index           =   19
                  Left            =   2415
                  TabIndex        =   62
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "申请时间"
                  Height          =   180
                  Index           =   18
                  Left            =   1305
                  TabIndex        =   61
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "紧急"
                  Height          =   180
                  Index           =   15
                  Left            =   4800
                  TabIndex        =   58
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   $"frmParPacs.frx":14E2E
                  Height          =   180
                  Index           =   17
                  Left            =   180
                  TabIndex        =   60
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "检查设备"
                  Height          =   180
                  Index           =   16
                  Left            =   6060
                  TabIndex        =   59
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "执行间"
                  Height          =   180
                  Index           =   14
                  Left            =   3525
                  TabIndex        =   57
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "地址"
                  Height          =   180
                  Index           =   13
                  Left            =   2415
                  TabIndex        =   56
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "邮编"
                  Height          =   180
                  Index           =   12
                  Left            =   1305
                  TabIndex        =   55
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "电话"
                  Height          =   180
                  Index           =   11
                  Left            =   180
                  TabIndex        =   54
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "婚姻"
                  Height          =   180
                  Index           =   10
                  Left            =   3525
                  TabIndex        =   51
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "职业"
                  Height          =   180
                  Index           =   9
                  Left            =   2415
                  TabIndex        =   50
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "民族"
                  Height          =   180
                  Index           =   8
                  Left            =   6060
                  TabIndex        =   53
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "身份证号"
                  Height          =   180
                  Index           =   7
                  Left            =   4800
                  TabIndex        =   52
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "付款方式"
                  Height          =   180
                  Index           =   6
                  Left            =   4800
                  TabIndex        =   46
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "费别"
                  Height          =   180
                  Index           =   5
                  Left            =   6060
                  TabIndex        =   47
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "体重"
                  Height          =   180
                  Index           =   4
                  Left            =   1305
                  TabIndex        =   49
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "身高"
                  Height          =   180
                  Index           =   3
                  Left            =   180
                  TabIndex        =   48
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "年龄"
                  Height          =   180
                  Index           =   2
                  Left            =   2415
                  TabIndex        =   44
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "性别"
                  Height          =   180
                  Index           =   1
                  Left            =   1305
                  TabIndex        =   43
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "英文名"
                  Height          =   180
                  Index           =   0
                  Left            =   180
                  TabIndex        =   42
                  Top             =   360
                  Width           =   1020
               End
            End
            Begin VB.Frame fra 
               Caption         =   "登记必录项目选择"
               Height          =   2010
               Index           =   15
               Left            =   -74280
               TabIndex        =   241
               Top             =   2760
               Width           =   7300
               Begin VB.CheckBox ChkInput 
                  Caption         =   "英文名"
                  Height          =   180
                  Index           =   0
                  Left            =   180
                  TabIndex        =   67
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "性别"
                  Height          =   180
                  Index           =   1
                  Left            =   1305
                  TabIndex        =   68
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "年龄"
                  Height          =   180
                  Index           =   2
                  Left            =   2415
                  TabIndex        =   69
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "身高"
                  Height          =   180
                  Index           =   3
                  Left            =   180
                  TabIndex        =   73
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "体重"
                  Height          =   180
                  Index           =   4
                  Left            =   1305
                  TabIndex        =   74
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "费别"
                  Height          =   180
                  Index           =   5
                  Left            =   6060
                  TabIndex        =   72
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "付款方式"
                  Height          =   180
                  Index           =   6
                  Left            =   4800
                  TabIndex        =   71
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "身份证号"
                  Height          =   180
                  Index           =   7
                  Left            =   4800
                  TabIndex        =   77
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "民族"
                  Height          =   180
                  Index           =   8
                  Left            =   6060
                  TabIndex        =   78
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "职业"
                  Height          =   180
                  Index           =   9
                  Left            =   2415
                  TabIndex        =   75
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "婚姻"
                  Height          =   180
                  Index           =   10
                  Left            =   3525
                  TabIndex        =   76
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "电话"
                  Height          =   180
                  Index           =   11
                  Left            =   180
                  TabIndex        =   79
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "邮编"
                  Height          =   180
                  Index           =   12
                  Left            =   1305
                  TabIndex        =   80
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "地址"
                  Height          =   180
                  Index           =   13
                  Left            =   2415
                  TabIndex        =   81
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "执行间"
                  Height          =   180
                  Index           =   14
                  Left            =   3525
                  TabIndex        =   82
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "检查设备"
                  Height          =   180
                  Index           =   16
                  Left            =   6060
                  TabIndex        =   84
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   $"frmParPacs.frx":14E3A
                  Height          =   180
                  Index           =   17
                  Left            =   180
                  TabIndex        =   85
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "紧急"
                  Height          =   180
                  Index           =   15
                  Left            =   4800
                  TabIndex        =   83
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "申请时间"
                  Height          =   180
                  Index           =   18
                  Left            =   1305
                  TabIndex        =   86
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "检查时间"
                  Height          =   180
                  Index           =   19
                  Left            =   2415
                  TabIndex        =   87
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "出生日期"
                  Height          =   180
                  Index           =   20
                  Left            =   3525
                  TabIndex        =   70
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "造影剂"
                  Height          =   180
                  Index           =   21
                  Left            =   3525
                  TabIndex        =   88
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "检查技师"
                  Height          =   180
                  Index           =   22
                  Left            =   4800
                  TabIndex        =   89
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "附加主述"
                  Height          =   180
                  Index           =   23
                  Left            =   180
                  TabIndex        =   91
                  Top             =   1590
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "检查技师二"
                  Height          =   180
                  Index           =   24
                  Left            =   6060
                  TabIndex        =   90
                  Top             =   1320
                  Width           =   1220
               End
            End
            Begin VB.Frame fra 
               Caption         =   "列表颜色配置"
               Height          =   5415
               Index           =   28
               Left            =   -74280
               TabIndex        =   222
               Top             =   480
               Width           =   7305
               Begin VB.Frame fra 
                  Caption         =   "颜色显示类型"
                  Height          =   615
                  Index           =   30
                  Left            =   3960
                  TabIndex        =   224
                  ToolTipText     =   "检查列表数据行颜色类型，为前景色时处理列表的前景色，反之处理背景色。"
                  Top             =   4680
                  Width           =   2055
                  Begin VB.OptionButton optListColorMark 
                     Caption         =   "前景色"
                     Height          =   255
                     Index           =   0
                     Left            =   120
                     TabIndex        =   141
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   855
                  End
                  Begin VB.OptionButton optListColorMark 
                     Caption         =   "背景色"
                     Height          =   255
                     Index           =   1
                     Left            =   1080
                     TabIndex        =   142
                     Top             =   240
                     Width           =   855
                  End
               End
               Begin VB.Frame fra 
                  Height          =   615
                  Index           =   29
                  Left            =   720
                  TabIndex        =   223
                  Top             =   4680
                  Width           =   2895
                  Begin VB.CheckBox chkNameColColorCfg 
                     Caption         =   "姓名分颜色显示"
                     Height          =   180
                     Left            =   120
                     TabIndex        =   139
                     ToolTipText     =   "姓名颜色根据病人类型显示。"
                     Top             =   0
                     Width           =   1800
                  End
                  Begin VB.CheckBox chkOrdinaryNameColColorCfg 
                     Caption         =   "启用缺省病人类型颜色"
                     Height          =   255
                     Left            =   600
                     TabIndex        =   140
                     Top             =   240
                     Width           =   2175
                  End
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "…"
                  Height          =   255
                  Index           =   10
                  Left            =   2655
                  TabIndex        =   135
                  Top             =   3600
                  Width           =   255
               End
               Begin VB.TextBox txtAudit 
                  Height          =   270
                  Left            =   4560
                  MaxLength       =   4
                  TabIndex        =   132
                  Text            =   "0"
                  Top             =   2400
                  Width           =   495
               End
               Begin VB.TextBox txtStudy 
                  Height          =   270
                  Left            =   4560
                  MaxLength       =   4
                  TabIndex        =   128
                  Text            =   "0"
                  Top             =   1440
                  Width           =   495
               End
               Begin VB.TextBox txtReport 
                  Height          =   270
                  Left            =   4560
                  MaxLength       =   4
                  TabIndex        =   130
                  Text            =   "0"
                  Top             =   1920
                  Width           =   495
               End
               Begin VB.TextBox txtCheckIn 
                  Height          =   270
                  Left            =   4560
                  MaxLength       =   4
                  TabIndex        =   126
                  Text            =   "0"
                  Top             =   960
                  Width           =   495
               End
               Begin VB.TextBox txtEnreg 
                  Height          =   270
                  Left            =   4560
                  MaxLength       =   4
                  TabIndex        =   124
                  Text            =   "0"
                  Top             =   480
                  Width           =   495
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "…"
                  Height          =   255
                  Index           =   9
                  Left            =   5295
                  TabIndex        =   134
                  Top             =   3120
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "…"
                  Height          =   255
                  Index           =   8
                  Left            =   2650
                  TabIndex        =   123
                  Top             =   480
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "…"
                  Height          =   255
                  Index           =   7
                  Left            =   2650
                  TabIndex        =   133
                  Top             =   3120
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "…"
                  Height          =   255
                  Index           =   6
                  Left            =   2650
                  TabIndex        =   131
                  Top             =   2400
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "…"
                  Height          =   255
                  Index           =   5
                  Left            =   5310
                  TabIndex        =   138
                  Top             =   4080
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "…"
                  Height          =   255
                  Index           =   4
                  Left            =   2650
                  TabIndex        =   129
                  Top             =   1920
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "…"
                  Height          =   255
                  Index           =   3
                  Left            =   2655
                  TabIndex        =   137
                  Top             =   4080
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "…"
                  Height          =   255
                  Index           =   2
                  Left            =   5295
                  TabIndex        =   136
                  Top             =   3600
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "…"
                  Height          =   255
                  Index           =   0
                  Left            =   2650
                  TabIndex        =   127
                  Top             =   1440
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "…"
                  Height          =   255
                  Index           =   1
                  Left            =   2650
                  TabIndex        =   125
                  Top             =   960
                  Width           =   255
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   10
                  Left            =   1560
                  Top             =   3600
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "已驳回："
                  Height          =   255
                  Index           =   19
                  Left            =   720
                  TabIndex        =   240
                  Top             =   3600
                  Width           =   735
               End
               Begin VB.Label lab 
                  BackColor       =   &H00000000&
                  BackStyle       =   0  'Transparent
                  Caption         =   "状态持续超出        分提醒"
                  Height          =   255
                  Index           =   25
                  Left            =   3360
                  TabIndex        =   239
                  Top             =   2430
                  Width           =   2415
               End
               Begin VB.Label lab 
                  BackColor       =   &H00000000&
                  BackStyle       =   0  'Transparent
                  Caption         =   "状态持续超出        分提醒"
                  Height          =   255
                  Index           =   23
                  Left            =   3360
                  TabIndex        =   238
                  Top             =   1470
                  Width           =   2415
               End
               Begin VB.Label lab 
                  BackColor       =   &H00000000&
                  BackStyle       =   0  'Transparent
                  Caption         =   "状态持续超出        分提醒"
                  Height          =   255
                  Index           =   24
                  Left            =   3360
                  TabIndex        =   237
                  Top             =   1950
                  Width           =   2415
               End
               Begin VB.Label lab 
                  BackColor       =   &H00000000&
                  BackStyle       =   0  'Transparent
                  Caption         =   "状态持续超出        分提醒"
                  Height          =   255
                  Index           =   22
                  Left            =   3360
                  TabIndex        =   236
                  Top             =   990
                  Width           =   2415
               End
               Begin VB.Label lab 
                  BackColor       =   &H00000000&
                  BackStyle       =   0  'Transparent
                  Caption         =   "状态持续超出        分提醒"
                  Height          =   255
                  Index           =   21
                  Left            =   3360
                  TabIndex        =   235
                  Top             =   510
                  Width           =   2415
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   9
                  Left            =   4200
                  Top             =   3120
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "已拒绝："
                  Height          =   255
                  Index           =   26
                  Left            =   3360
                  TabIndex        =   234
                  Top             =   3120
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  BackColor       =   &H00FFFFFF&
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   8
                  Left            =   1560
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "已登记："
                  Height          =   255
                  Index           =   13
                  Left            =   720
                  TabIndex        =   233
                  Top             =   480
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   7
                  Left            =   1560
                  Top             =   3120
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "已完成："
                  Height          =   255
                  Index           =   18
                  Left            =   720
                  TabIndex        =   232
                  Top             =   3120
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   6
                  Left            =   1560
                  Top             =   2400
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "已审核："
                  Height          =   255
                  Index           =   17
                  Left            =   720
                  TabIndex        =   231
                  Top             =   2400
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   5
                  Left            =   4215
                  Top             =   4080
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "审核中："
                  Height          =   255
                  Index           =   28
                  Left            =   3375
                  TabIndex        =   230
                  Top             =   4080
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   4
                  Left            =   1560
                  Top             =   1920
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "已报告："
                  Height          =   255
                  Index           =   16
                  Left            =   720
                  TabIndex        =   229
                  Top             =   1920
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   3
                  Left            =   1560
                  Top             =   4080
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "报告中："
                  Height          =   255
                  Index           =   20
                  Left            =   720
                  TabIndex        =   228
                  Top             =   4080
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   2
                  Left            =   4200
                  Top             =   3600
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "处理中："
                  Height          =   255
                  Index           =   27
                  Left            =   3360
                  TabIndex        =   227
                  Top             =   3600
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   0
                  Left            =   1560
                  Top             =   1440
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "已检查："
                  Height          =   255
                  Index           =   15
                  Left            =   720
                  TabIndex        =   226
                  Top             =   1440
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   1
                  Left            =   1560
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "已报到："
                  Height          =   255
                  Index           =   14
                  Left            =   720
                  TabIndex        =   225
                  Top             =   960
                  Width           =   735
               End
            End
            Begin VB.CommandButton cmdDefault 
               Caption         =   "恢复默认(&D)"
               Height          =   375
               Left            =   -69000
               TabIndex        =   143
               Top             =   6240
               Width           =   1335
            End
            Begin VB.Frame fra 
               Caption         =   "报告编辑器"
               Height          =   615
               Index           =   19
               Left            =   -74280
               TabIndex        =   221
               Top             =   480
               Width           =   7245
               Begin VB.OptionButton optReportEditor 
                  Caption         =   "PACS智能报告编辑器"
                  Height          =   255
                  Index           =   2
                  Left            =   4560
                  TabIndex        =   110
                  Top             =   240
                  Width           =   2052
               End
               Begin VB.OptionButton optReportEditor 
                  Caption         =   "电子病历编辑器"
                  Height          =   255
                  Index           =   0
                  Left            =   360
                  TabIndex        =   108
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.OptionButton optReportEditor 
                  Caption         =   "PACS报告编辑器"
                  Height          =   255
                  Index           =   1
                  Left            =   2400
                  TabIndex        =   109
                  Top             =   240
                  Width           =   1575
               End
            End
            Begin VB.Frame fra 
               Caption         =   "报告设置"
               Height          =   3255
               Index           =   20
               Left            =   -74280
               TabIndex        =   220
               Top             =   2640
               Width           =   7245
               Begin VB.Frame fra 
                  Caption         =   "报告文本段名称"
                  Height          =   1335
                  Index           =   26
                  Left            =   240
                  TabIndex        =   365
                  Top             =   1800
                  Width           =   3615
                  Begin VB.TextBox txtCheckView 
                     Height          =   270
                     Left            =   1560
                     TabIndex        =   368
                     Top             =   225
                     Width           =   1335
                  End
                  Begin VB.TextBox txtResult 
                     Height          =   270
                     Left            =   1560
                     TabIndex        =   367
                     Top             =   600
                     Width           =   1335
                  End
                  Begin VB.TextBox txtAdvice 
                     Height          =   270
                     Left            =   1560
                     TabIndex        =   366
                     Top             =   960
                     Width           =   1335
                  End
                  Begin VB.Label lab 
                     Caption         =   "检查所见："
                     Height          =   255
                     Index           =   10
                     Left            =   360
                     TabIndex        =   371
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.Label lab 
                     Caption         =   "诊断意见："
                     Height          =   255
                     Index           =   11
                     Left            =   360
                     TabIndex        =   370
                     Top             =   608
                     Width           =   975
                  End
                  Begin VB.Label lab 
                     Caption         =   "建    议："
                     Height          =   255
                     Index           =   12
                     Left            =   360
                     TabIndex        =   369
                     Top             =   975
                     Width           =   975
                  End
               End
               Begin VB.Frame Frame7 
                  Caption         =   "打印格式选择方式"
                  Height          =   1275
                  Left            =   4200
                  TabIndex        =   298
                  Top             =   1800
                  Width           =   2745
                  Begin VB.CheckBox chkPrintFormat 
                     Caption         =   "单选报告格式"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   310
                     Top             =   960
                     Width           =   1455
                  End
                  Begin VB.OptionButton optPrintFormat 
                     Caption         =   "记录最后一次打印格式"
                     Height          =   255
                     Index           =   0
                     Left            =   240
                     TabIndex        =   300
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   2175
                  End
                  Begin VB.OptionButton optPrintFormat 
                     Caption         =   "始终保持默认格式"
                     Height          =   255
                     Index           =   1
                     Left            =   240
                     TabIndex        =   299
                     Top             =   600
                     Width           =   1815
                  End
               End
               Begin VB.CheckBox chkUntreadPrinted 
                  Caption         =   "审核打印后允许回退"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   115
                  Top             =   720
                  Width           =   2055
               End
               Begin VB.CheckBox chkSpecialContent 
                  Caption         =   "显示专科报告内容："
                  Height          =   180
                  Left            =   240
                  TabIndex        =   116
                  Top             =   1080
                  Width           =   2055
               End
               Begin VB.ComboBox cboSpecialContent 
                  Enabled         =   0   'False
                  Height          =   300
                  Left            =   240
                  TabIndex        =   117
                  Text            =   "Combo1"
                  Top             =   1360
                  Width           =   6735
               End
               Begin VB.CheckBox chkExitAfterPrint 
                  Caption         =   "打印后退出"
                  Height          =   180
                  Left            =   2280
                  TabIndex        =   114
                  Top             =   720
                  Width           =   1335
               End
               Begin VB.CheckBox chkShowVideoCapture 
                  Caption         =   "显示视频采集区域"
                  Height          =   180
                  Left            =   2280
                  TabIndex        =   113
                  Top             =   360
                  Width           =   2055
               End
               Begin VB.TextBox txtMinImageCount 
                  Height          =   270
                  Left            =   6240
                  MaxLength       =   2
                  TabIndex        =   112
                  Text            =   "8"
                  Top             =   315
                  Width           =   615
               End
               Begin VB.CheckBox chkShowImage 
                  Caption         =   "显示报告图像区域                            报告缩略图显示数量："
                  Height          =   180
                  Left            =   240
                  TabIndex        =   111
                  Top             =   360
                  Width           =   6135
               End
            End
            Begin VB.Frame fra 
               Caption         =   "报告词句双击后"
               Height          =   855
               Index           =   21
               Left            =   -74280
               TabIndex        =   219
               Top             =   6060
               Width           =   2415
               Begin VB.OptionButton optWordDblClick 
                  Caption         =   "直接写入报告"
                  Height          =   255
                  Index           =   0
                  Left            =   360
                  TabIndex        =   118
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.OptionButton optWordDblClick 
                  Caption         =   "打开词句编辑窗口"
                  Height          =   255
                  Index           =   1
                  Left            =   360
                  TabIndex        =   119
                  Top             =   480
                  Width           =   1750
               End
            End
            Begin VB.Frame fra 
               Caption         =   "缩略图双击后"
               Height          =   855
               Index           =   22
               Left            =   -71880
               TabIndex        =   218
               Top             =   6060
               Width           =   2415
               Begin VB.OptionButton optImageDblClick 
                  Caption         =   "打开图片编辑窗口"
                  Height          =   255
                  Index           =   1
                  Left            =   360
                  TabIndex        =   121
                  Top             =   480
                  Width           =   1750
               End
               Begin VB.OptionButton optImageDblClick 
                  Caption         =   "直接写入报告"
                  Height          =   255
                  Index           =   0
                  Left            =   360
                  TabIndex        =   120
                  Top             =   240
                  Width           =   1575
               End
            End
            Begin VB.Frame fra 
               Caption         =   "词句模板显示"
               Height          =   855
               Index           =   23
               Left            =   -69480
               TabIndex        =   217
               Top             =   6060
               Width           =   2450
               Begin VB.OptionButton optShowWord 
                  Caption         =   "双击标题"
                  Height          =   180
                  Index           =   1
                  Left            =   360
                  TabIndex        =   150
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.OptionButton optShowWord 
                  Caption         =   "直接显示"
                  Height          =   180
                  Index           =   0
                  Left            =   360
                  TabIndex        =   122
                  Top             =   240
                  Width           =   1095
               End
            End
            Begin VB.Frame fra 
               Caption         =   "分组设置"
               Height          =   4815
               Index           =   16
               Left            =   -74280
               TabIndex        =   212
               Top             =   480
               Width           =   7335
               Begin VB.CheckBox chkSelectRoom 
                  Caption         =   "报到时分配默认执行间"
                  Height          =   210
                  Left            =   3840
                  TabIndex        =   98
                  Top             =   4485
                  Width           =   2220
               End
               Begin VB.CommandButton cmdAddGroup 
                  Caption         =   "新增分组(&A)"
                  Height          =   375
                  Left            =   120
                  Picture         =   "frmParPacs.frx":14E46
                  TabIndex        =   95
                  TabStop         =   0   'False
                  Top             =   4380
                  Width           =   1170
               End
               Begin VB.CommandButton cmdDelGroup 
                  Caption         =   "删除分组(&D)"
                  Height          =   375
                  Left            =   1320
                  Picture         =   "frmParPacs.frx":14F90
                  TabIndex        =   96
                  TabStop         =   0   'False
                  Top             =   4380
                  Width           =   1170
               End
               Begin VB.CommandButton cmdStudyAcc 
                  Caption         =   "关联项目(&R)"
                  Height          =   375
                  Left            =   6000
                  Picture         =   "frmParPacs.frx":150DA
                  TabIndex        =   99
                  TabStop         =   0   'False
                  Top             =   4380
                  Width           =   1155
               End
               Begin VB.CommandButton cmdModify 
                  Caption         =   "修改分组(&M)"
                  Height          =   375
                  Left            =   2520
                  Picture         =   "frmParPacs.frx":15224
                  TabIndex        =   97
                  TabStop         =   0   'False
                  Top             =   4380
                  Width           =   1170
               End
               Begin zl9BaseItem.ucFlexGrid ufgStudyProCfg 
                  Height          =   1905
                  Left            =   3405
                  TabIndex        =   94
                  Top             =   2430
                  Width           =   3735
                  _ExtentX        =   6588
                  _ExtentY        =   3360
                  DefaultCols     =   ""
                  ColNames        =   "|关联检查项目>名称,w2100,read|项目编码>编码,w1100,read|"
                  KeyName         =   "≡"
                  HeadCheckValue  =   1
                  IsCopyAdoMode   =   0   'False
                  IsEjectConfig   =   -1  'True
                  HeadFontCharset =   134
                  HeadFontWeight  =   400
                  HeadColor       =   0
                  DataFontCharset =   134
                  DataFontWeight  =   400
                  DataColor       =   -2147483640
                  GridLineColor   =   14737632
               End
               Begin zl9BaseItem.ucFlexGrid ufgRoomCfg 
                  Height          =   2175
                  Left            =   3405
                  TabIndex        =   93
                  Top             =   240
                  Width           =   3735
                  _ExtentX        =   6588
                  _ExtentY        =   3836
                  DefaultCols     =   ""
                  ColNames        =   "|ID,hide|执行间,w1400,read|号码前缀,w1400,read|"
                  KeyName         =   "≡"
                  HeadCheckValue  =   1
                  IsCopyAdoMode   =   0   'False
                  IsEjectConfig   =   -1  'True
                  HeadFontCharset =   134
                  HeadFontWeight  =   400
                  HeadColor       =   0
                  DataFontCharset =   134
                  DataFontWeight  =   400
                  DataColor       =   -2147483640
                  GridLineColor   =   14737632
               End
               Begin zl9BaseItem.ucFlexGrid ufgGroupCfg 
                  Height          =   4095
                  Left            =   120
                  TabIndex        =   92
                  Top             =   240
                  Width           =   3255
                  _ExtentX        =   5741
                  _ExtentY        =   7223
                  DefaultCols     =   ""
                  ColNames        =   "|ID,hide,key|组名,w1400,read|分组前缀,w1500,read|"
                  KeyName         =   "ID"
                  IsCopyAdoMode   =   0   'False
                  IsEjectConfig   =   -1  'True
                  HeadFontCharset =   134
                  HeadFontWeight  =   400
                  DataFontCharset =   134
                  DataFontWeight  =   400
               End
            End
            Begin VB.Frame fra 
               Height          =   1515
               Index           =   17
               Left            =   -74280
               TabIndex        =   207
               Top             =   5400
               Width           =   7335
               Begin VB.CheckBox chkAutoInQueue 
                  Caption         =   "报到后自动排队"
                  Height          =   180
                  Left            =   3480
                  TabIndex        =   106
                  Top             =   1125
                  Value           =   1  'Checked
                  Width           =   1575
               End
               Begin VB.CheckBox chkUseQueueMsg 
                  Caption         =   "启用排队消息处理"
                  Height          =   180
                  Left            =   5160
                  TabIndex        =   107
                  Top             =   1125
                  Value           =   1  'Checked
                  Width           =   1815
               End
               Begin VB.ComboBox cbxPrintQueueNoWay 
                  Height          =   300
                  ItemData        =   "frmParPacs.frx":1536E
                  Left            =   1635
                  List            =   "frmParPacs.frx":1537B
                  Style           =   2  'Dropdown List
                  TabIndex        =   105
                  Top             =   1080
                  Width           =   1740
               End
               Begin VB.Frame fra 
                  Caption         =   "未指定执行间的排队方式"
                  Height          =   810
                  Index           =   18
                  Left            =   4680
                  TabIndex        =   208
                  Top             =   240
                  Width           =   2265
                  Begin VB.OptionButton optNumberRule 
                     Caption         =   "按检查科室排队"
                     Height          =   180
                     Index           =   0
                     Left            =   105
                     TabIndex        =   103
                     ToolTipText     =   "对于分配了执行间的检查，排队号码将按执行间连续生成，对未分配执行的检查，排队号码将按科室连续生成。"
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   1755
                  End
                  Begin VB.OptionButton optNumberRule 
                     Caption         =   "按检查分组排队"
                     Height          =   180
                     Index           =   1
                     Left            =   105
                     TabIndex        =   104
                     ToolTipText     =   "对于分配了执行间的检查，排队号码将按执行间连续生成，对未分配执行的检查，排队号码将根据检查所属分组连续生成。"
                     Top             =   480
                     Width           =   1665
                  End
               End
               Begin VB.CheckBox chkSynStudyList 
                  Caption         =   "同步定位检查列表"
                  Height          =   180
                  Left            =   2760
                  TabIndex        =   101
                  ToolTipText     =   "点击排队列表或呼叫列表数据后，同步定位到检查列表"
                  Top             =   330
                  Width           =   1815
               End
               Begin VB.TextBox txtQueueReport 
                  Height          =   315
                  Left            =   1635
                  TabIndex        =   102
                  Top             =   690
                  Width           =   2820
               End
               Begin VB.TextBox txtValidDays 
                  Height          =   315
                  Left            =   1635
                  MaxLength       =   2
                  TabIndex        =   100
                  Text            =   "1"
                  Top             =   285
                  Width           =   555
               End
               Begin VB.Label lab 
                  Caption         =   "排号单打印方式："
                  Height          =   255
                  Index           =   9
                  Left            =   240
                  TabIndex        =   211
                  Top             =   1110
                  Width           =   1455
               End
               Begin VB.Label lab 
                  Caption         =   "排号单报表编号："
                  Height          =   225
                  Index           =   8
                  Left            =   240
                  TabIndex        =   210
                  ToolTipText     =   "排队打号时对应的自定义报表编号。"
                  Top             =   735
                  Width           =   1455
               End
               Begin VB.Label lab 
                  Caption         =   "数据有效天数：       天"
                  Height          =   210
                  Index           =   7
                  Left            =   420
                  TabIndex        =   209
                  Top             =   360
                  Width           =   2235
               End
            End
            Begin VB.Frame fra 
               Caption         =   "拼音名"
               Height          =   1695
               Index           =   4
               Left            =   5280
               TabIndex        =   205
               Top             =   5040
               Width           =   2775
               Begin VB.OptionButton optCapital 
                  Caption         =   "大写"
                  Height          =   255
                  Index           =   0
                  Left            =   240
                  TabIndex        =   29
                  ToolTipText     =   "选择后拼音名显示全为大写字母。"
                  Top             =   260
                  Width           =   735
               End
               Begin VB.OptionButton optCapital 
                  Caption         =   "小写"
                  Height          =   255
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   30
                  ToolTipText     =   "选择后拼音名显示全为小写字母。"
                  Top             =   240
                  Width           =   735
               End
               Begin VB.OptionButton optCapital 
                  Caption         =   "首字母大写"
                  Height          =   255
                  Index           =   2
                  Left            =   240
                  TabIndex        =   31
                  ToolTipText     =   "选择后拼音名首字母大写。"
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.Frame fra 
                  Caption         =   "间隔"
                  Height          =   540
                  Index           =   11
                  Left            =   240
                  TabIndex        =   206
                  Top             =   960
                  Width           =   1695
                  Begin VB.OptionButton optSplitter 
                     Caption         =   "无"
                     Height          =   255
                     Index           =   1
                     Left            =   960
                     TabIndex        =   33
                     ToolTipText     =   "拼音名之间无间隔。"
                     Top             =   200
                     Width           =   495
                  End
                  Begin VB.OptionButton optSplitter 
                     Caption         =   "空格"
                     Height          =   255
                     Index           =   0
                     Left            =   120
                     TabIndex        =   32
                     ToolTipText     =   "拼音名之间使用空格为间隔符。"
                     Top             =   200
                     Width           =   735
                  End
               End
            End
            Begin VB.Frame fra 
               Caption         =   "先检查后报到，图像匹配"
               Height          =   1545
               Index           =   3
               Left            =   5280
               TabIndex        =   204
               Top             =   3360
               Width           =   2775
               Begin VB.OptionButton optMatch 
                  Caption         =   "门诊/住院号"
                  Height          =   195
                  Index           =   1
                  Left            =   240
                  TabIndex        =   28
                  ToolTipText     =   "报到时通过门诊/住院号和图像信息进行匹配，仅用于影像医技站。"
                  Top             =   1000
                  Width           =   1335
               End
               Begin VB.OptionButton optMatch 
                  Caption         =   "检查号"
                  Height          =   195
                  Index           =   0
                  Left            =   240
                  TabIndex        =   26
                  ToolTipText     =   "报到时通过检查号和图像信息进行匹配，仅用于影像医技站。"
                  Top             =   360
                  Width           =   855
               End
               Begin VB.OptionButton optMatch 
                  Caption         =   "医嘱ID"
                  Height          =   195
                  Index           =   2
                  Left            =   240
                  TabIndex        =   27
                  ToolTipText     =   "报到时通过医嘱ID和图像信息进行匹配，仅用于影像医技站。"
                  Top             =   680
                  Width           =   855
               End
            End
            Begin VB.Frame fra 
               Caption         =   "功能设置"
               Height          =   1665
               Index           =   2
               Left            =   720
               TabIndex        =   202
               Top             =   3240
               Width           =   4215
               Begin VB.Frame Frame1 
                  Caption         =   "快速过滤设置"
                  Height          =   780
                  Left            =   1920
                  TabIndex        =   302
                  Top             =   840
                  Width           =   2055
                  Begin VB.CheckBox chkNameQueryTimeLimit 
                     Caption         =   "姓名查询时间限制"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   304
                     ToolTipText     =   "按姓名查询时，是否有查询时间限制"
                     Top             =   480
                     Width           =   1850
                  End
                  Begin VB.CheckBox chkNameFuzzySearch 
                     Caption         =   "姓名默认模糊查询"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   303
                     ToolTipText     =   "按姓名查询时使用模糊查询，没有勾选时则只有输入*后才进行模糊查询"
                     Top             =   240
                     Width           =   1850
                  End
               End
               Begin VB.CheckBox chkSwitchUser 
                  Caption         =   "启用切换用户"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   22
                  ToolTipText     =   "激活切换用户功能，可以进行用户切换。"
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.Frame fra 
                  Height          =   660
                  Index           =   10
                  Left            =   1920
                  TabIndex        =   203
                  ToolTipText     =   "选择采集图像和扫描申请单所使用的存储设备。"
                  Top             =   180
                  Width           =   2055
                  Begin VB.ComboBox cboSaveDevice 
                     Height          =   300
                     Left            =   120
                     Style           =   2  'Dropdown List
                     TabIndex        =   25
                     Top             =   240
                     Width           =   1605
                  End
                  Begin VB.CheckBox chkPetitionCapture 
                     Caption         =   "启用申请单扫描"
                     Height          =   180
                     Left            =   120
                     TabIndex        =   24
                     ToolTipText     =   "启用申请单扫描功能"
                     Top             =   0
                     Value           =   1  'Checked
                     Width           =   1575
                  End
               End
               Begin VB.CheckBox chkUseReferencePatient 
                  Caption         =   "启用关联病人"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   23
                  ToolTipText     =   "支持多个检查关联到同一个病人信息。"
                  Top             =   1080
                  Width           =   1455
               End
               Begin VB.CheckBox chkChangeUser 
                  Caption         =   "启用交换用户"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   21
                  ToolTipText     =   "激活交换用户功能，可以交换检查医生和报告医生，仅限于影像采集站。"
                  Top             =   360
                  Width           =   1455
               End
            End
         End
         Begin MSComDlg.CommonDialog dlgColor 
            Left            =   4920
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6975
         Index           =   5
         Left            =   -75000
         ScaleHeight     =   6975
         ScaleWidth      =   7815
         TabIndex        =   196
         Top             =   360
         Width           =   7815
         Begin VB.Frame fra 
            Height          =   1695
            Index           =   40
            Left            =   120
            TabIndex        =   282
            Top             =   120
            Width           =   4455
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   3
               Left            =   2160
               TabIndex        =   176
               Top             =   720
               Width           =   2175
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   21
               Left            =   2160
               TabIndex        =   175
               Text            =   "100"
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox chk 
               Caption         =   "借阅确认后自动打印借阅回执单"
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   177
               Top             =   1200
               Width           =   2895
            End
            Begin VB.Label lab 
               Caption         =   "天"
               Height          =   255
               Index           =   56
               Left            =   3000
               TabIndex        =   285
               Top             =   280
               Width           =   255
            End
            Begin VB.Label lab 
               Caption         =   "借阅回执对应报表名称："
               Height          =   255
               Index           =   57
               Left            =   120
               TabIndex        =   284
               Top             =   760
               Width           =   2055
            End
            Begin VB.Label lab 
               Caption         =   "借阅记录默认查询天数："
               Height          =   255
               Index           =   55
               Left            =   120
               TabIndex        =   283
               Top             =   285
               Width           =   2055
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6975
         Index           =   4
         Left            =   -75000
         ScaleHeight     =   6975
         ScaleWidth      =   7815
         TabIndex        =   195
         Top             =   360
         Width           =   7815
         Begin VB.Frame fra 
            Height          =   1215
            Index           =   39
            Left            =   120
            TabIndex        =   278
            Top             =   120
            Width           =   4455
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   20
               Left            =   2160
               TabIndex        =   173
               Text            =   "30"
               Top             =   240
               Width           =   735
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   2
               Left            =   2160
               TabIndex        =   174
               Top             =   720
               Width           =   2175
            End
            Begin VB.Label lab 
               Caption         =   "档案记录默认查询天数："
               Height          =   255
               Index           =   52
               Left            =   120
               TabIndex        =   281
               Top             =   285
               Width           =   2055
            End
            Begin VB.Label lab 
               Caption         =   "档案标签对应报表名称："
               Height          =   255
               Index           =   54
               Left            =   120
               TabIndex        =   280
               Top             =   760
               Width           =   2055
            End
            Begin VB.Label lab 
               Caption         =   "天"
               Height          =   255
               Index           =   53
               Left            =   3000
               TabIndex        =   279
               Top             =   280
               Width           =   255
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6975
         Index           =   3
         Left            =   -75000
         ScaleHeight     =   6975
         ScaleWidth      =   7815
         TabIndex        =   194
         Top             =   360
         Width           =   7815
         Begin VB.CheckBox chk 
            Caption         =   "录入外院信息"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   351
            Top             =   4440
            Width           =   1455
         End
         Begin VB.Frame Frame4 
            Caption         =   "Frame4"
            Height          =   735
            Left            =   120
            TabIndex        =   313
            Top             =   4440
            Width           =   4815
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   22
               Left            =   1680
               TabIndex        =   353
               Top             =   300
               Width           =   2895
            End
            Begin VB.Label lab 
               Caption         =   "送检单位配置："
               Height          =   255
               Index           =   61
               Left            =   240
               TabIndex        =   352
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.ListBox lst 
            Height          =   2790
            Index           =   0
            Left            =   5280
            Style           =   1  'Checkbox
            TabIndex        =   311
            Top             =   600
            Width           =   2055
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   5
            ItemData        =   "frmParPacs.frx":153AB
            Left            =   1410
            List            =   "frmParPacs.frx":153B8
            Style           =   2  'Dropdown List
            TabIndex        =   307
            Top             =   3960
            Width           =   2205
         End
         Begin VB.Frame fra 
            Caption         =   "当所选类型检查完成时自动弹出质量窗口"
            Height          =   1095
            Index           =   38
            Left            =   120
            TabIndex        =   277
            Top             =   2640
            Width           =   4815
            Begin VB.CheckBox chk 
               Caption         =   "常规"
               Height          =   375
               Index           =   2
               Left            =   480
               TabIndex        =   167
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox chk 
               Caption         =   "冰冻"
               Height          =   375
               Index           =   3
               Left            =   1800
               TabIndex        =   168
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox chk 
               Caption         =   "细胞"
               Height          =   375
               Index           =   4
               Left            =   3240
               TabIndex        =   169
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox chk 
               Caption         =   "会诊"
               Height          =   375
               Index           =   5
               Left            =   480
               TabIndex        =   170
               Top             =   600
               Width           =   735
            End
            Begin VB.CheckBox chk 
               Caption         =   "尸检"
               Height          =   375
               Index           =   6
               Left            =   1800
               TabIndex        =   171
               Top             =   600
               Width           =   735
            End
            Begin VB.CheckBox chk 
               Caption         =   "快片"
               Height          =   375
               Index           =   7
               Left            =   3240
               TabIndex        =   172
               Top             =   600
               Width           =   735
            End
         End
         Begin VB.Frame fra 
            Caption         =   "词句模板设置"
            Height          =   2415
            Index           =   37
            Left            =   120
            TabIndex        =   269
            Top             =   120
            Width           =   4815
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   15
               Left            =   1560
               TabIndex        =   162
               Top             =   495
               Width           =   3015
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   16
               Left            =   1560
               TabIndex        =   163
               Top             =   855
               Width           =   3015
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   17
               Left            =   1560
               TabIndex        =   164
               Top             =   1215
               Width           =   3015
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   18
               Left            =   1560
               TabIndex        =   165
               Top             =   1575
               Width           =   3015
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   19
               Left            =   1560
               TabIndex        =   270
               Top             =   1935
               Width           =   3015
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "对应词句分类"
               Height          =   180
               Index           =   51
               Left            =   2400
               TabIndex        =   276
               Top             =   240
               Width           =   1080
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "巨检描述模板："
               Height          =   180
               Index           =   46
               Left            =   240
               TabIndex        =   275
               Top             =   540
               Width           =   1260
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "常规报告模板："
               Height          =   180
               Index           =   47
               Left            =   240
               TabIndex        =   274
               Top             =   900
               Width           =   1260
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "免疫报告模板："
               Height          =   180
               Index           =   48
               Left            =   240
               TabIndex        =   273
               Top             =   1260
               Width           =   1260
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "分子报告模板："
               Height          =   180
               Index           =   50
               Left            =   240
               TabIndex        =   272
               Top             =   1980
               Width           =   1260
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "特染报告模板："
               Height          =   180
               Index           =   49
               Left            =   240
               TabIndex        =   271
               Top             =   1620
               Width           =   1260
            End
         End
         Begin VB.Label lab 
            AutoSize        =   -1  'True
            Caption         =   "病理取材显示信息"
            Height          =   180
            Index           =   60
            Left            =   5280
            TabIndex        =   312
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lab 
            Caption         =   "费用执行模式："
            Height          =   270
            Index           =   59
            Left            =   120
            TabIndex        =   308
            Top             =   4005
            Width           =   1305
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6975
         Index           =   2
         Left            =   -75000
         ScaleHeight     =   6975
         ScaleWidth      =   7815
         TabIndex        =   193
         Top             =   360
         Width           =   7815
         Begin VB.Frame fra 
            Height          =   1095
            Index           =   36
            Left            =   120
            TabIndex        =   290
            Top             =   120
            Width           =   4095
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   1
               ItemData        =   "frmParPacs.frx":153E0
               Left            =   1410
               List            =   "frmParPacs.frx":153ED
               Style           =   2  'Dropdown List
               TabIndex        =   292
               Top             =   600
               Width           =   2205
            End
            Begin VB.CheckBox chk 
               Caption         =   "录入外院信息"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   291
               Top             =   240
               Width           =   1590
            End
            Begin VB.Label lab 
               Caption         =   "费用执行模式："
               Height          =   270
               Index           =   45
               Left            =   120
               TabIndex        =   293
               Top             =   645
               Width           =   1305
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6975
         Index           =   1
         Left            =   0
         ScaleHeight     =   6975
         ScaleWidth      =   7815
         TabIndex        =   192
         Top             =   360
         Width           =   7815
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   4
            ItemData        =   "frmParPacs.frx":15415
            Left            =   1530
            List            =   "frmParPacs.frx":1541F
            Style           =   2  'Dropdown List
            TabIndex        =   305
            Top             =   6000
            Width           =   2205
         End
         Begin VB.CheckBox chk 
            Caption         =   "录入外院信息"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   144
            Top             =   120
            Width           =   1590
         End
         Begin VB.Frame fra 
            Caption         =   "XWPACS观片"
            Height          =   5295
            Index           =   31
            Left            =   240
            TabIndex        =   248
            Top             =   600
            Width           =   7335
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   23
               Left            =   1590
               TabIndex        =   381
               Top             =   3840
               Width           =   5640
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   11
               Left            =   1590
               TabIndex        =   158
               Text            =   "zlhis"
               Top             =   4320
               Width           =   5640
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   12
               Left            =   5430
               TabIndex        =   287
               Text            =   "DCMSHARE"
               Top             =   3840
               Visible         =   0   'False
               Width           =   1785
            End
            Begin VB.Frame fra 
               Caption         =   "删除图像用户"
               Height          =   1215
               Index           =   32
               Left            =   120
               TabIndex        =   266
               Top             =   480
               Width           =   2295
               Begin VB.TextBox txt 
                  Height          =   270
                  IMEMode         =   3  'DISABLE
                  Index           =   1
                  Left            =   720
                  PasswordChar    =   "*"
                  TabIndex        =   146
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.TextBox txt 
                  Height          =   270
                  Index           =   0
                  Left            =   720
                  TabIndex        =   145
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.Label lab 
                  Caption         =   "密码"
                  Height          =   255
                  Index           =   33
                  Left            =   120
                  TabIndex        =   268
                  Top             =   780
                  Width           =   615
               End
               Begin VB.Label lab 
                  Caption         =   "用户名"
                  Height          =   255
                  Index           =   32
                  Left            =   120
                  TabIndex        =   267
                  Top             =   420
                  Width           =   615
               End
            End
            Begin VB.Frame fra 
               Caption         =   "发送图像用户"
               Height          =   1215
               Index           =   33
               Left            =   2520
               TabIndex        =   263
               Top             =   480
               Width           =   2295
               Begin VB.TextBox txt 
                  Height          =   270
                  Index           =   2
                  Left            =   720
                  TabIndex        =   147
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.TextBox txt 
                  Height          =   270
                  IMEMode         =   3  'DISABLE
                  Index           =   3
                  Left            =   720
                  PasswordChar    =   "*"
                  TabIndex        =   148
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.Label lab 
                  Caption         =   "用户名"
                  Height          =   255
                  Index           =   34
                  Left            =   120
                  TabIndex        =   265
                  Top             =   420
                  Width           =   615
               End
               Begin VB.Label lab 
                  Caption         =   "密码"
                  Height          =   255
                  Index           =   35
                  Left            =   120
                  TabIndex        =   264
                  Top             =   780
                  Width           =   615
               End
            End
            Begin VB.Frame fra 
               Caption         =   "光盘刻录用户"
               Height          =   1215
               Index           =   34
               Left            =   4920
               TabIndex        =   260
               Top             =   480
               Width           =   2295
               Begin VB.TextBox txt 
                  Height          =   270
                  IMEMode         =   3  'DISABLE
                  Index           =   5
                  Left            =   720
                  PasswordChar    =   "*"
                  TabIndex        =   151
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.TextBox txt 
                  Height          =   270
                  Index           =   4
                  Left            =   720
                  TabIndex        =   149
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.Label lab 
                  Caption         =   "密码"
                  Height          =   255
                  Index           =   37
                  Left            =   120
                  TabIndex        =   262
                  Top             =   780
                  Width           =   615
               End
               Begin VB.Label lab 
                  Caption         =   "用户名"
                  Height          =   255
                  Index           =   36
                  Left            =   120
                  TabIndex        =   261
                  Top             =   420
                  Width           =   615
               End
            End
            Begin VB.Frame fra 
               Caption         =   "XWPACS 数据库服务器"
               Height          =   855
               Index           =   35
               Left            =   120
               TabIndex        =   249
               Top             =   1800
               Width           =   7095
               Begin VB.CommandButton cmdTestCon 
                  Caption         =   "测"
                  Height          =   375
                  Left            =   6600
                  TabIndex        =   155
                  Top             =   300
                  Width           =   375
               End
               Begin VB.TextBox txt 
                  Height          =   270
                  Index           =   6
                  Left            =   840
                  TabIndex        =   152
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.TextBox txt 
                  Height          =   270
                  Index           =   7
                  Left            =   3000
                  TabIndex        =   153
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.TextBox txt 
                  Height          =   270
                  IMEMode         =   3  'DISABLE
                  Index           =   8
                  Left            =   5040
                  PasswordChar    =   "*"
                  TabIndex        =   154
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.Label lab 
                  Caption         =   "服务名"
                  Height          =   255
                  Index           =   29
                  Left            =   240
                  TabIndex        =   252
                  Top             =   420
                  Width           =   855
               End
               Begin VB.Label lab 
                  Caption         =   "用户名"
                  Height          =   255
                  Index           =   30
                  Left            =   2400
                  TabIndex        =   251
                  Top             =   420
                  Width           =   855
               End
               Begin VB.Label lab 
                  Caption         =   "密码"
                  Height          =   255
                  Index           =   31
                  Left            =   4560
                  TabIndex        =   250
                  Top             =   420
                  Width           =   495
               End
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   13
               Left            =   1320
               TabIndex        =   159
               Text            =   "1"
               Top             =   4800
               Width           =   372
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   14
               Left            =   3135
               TabIndex        =   160
               Text            =   "2"
               Top             =   4800
               Width           =   372
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   9
               Left            =   1590
               TabIndex        =   156
               Text            =   "http://127.0.0.1:8080/TakeImage.aspx?colid0=22&colvalue0=[@STU_NO]"
               Top             =   2880
               Width           =   5640
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   0
               ItemData        =   "frmParPacs.frx":1543B
               Left            =   5400
               List            =   "frmParPacs.frx":15445
               Style           =   2  'Dropdown List
               TabIndex        =   161
               Top             =   4800
               Width           =   1812
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   10
               Left            =   1590
               TabIndex        =   157
               Text            =   "http://127.0.0.1:8080/KeyImage.aspx?colid0=22&colvalue0=[@STU_NO]"
               Top             =   3360
               Width           =   5640
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               Caption         =   "检查列表观片地址"
               Height          =   180
               Index           =   2
               Left            =   120
               TabIndex        =   382
               Top             =   3885
               Width           =   1560
            End
            Begin VB.Label lab 
               Caption         =   "历史图像共享目录"
               Height          =   255
               Index           =   41
               Left            =   3855
               TabIndex        =   259
               Top             =   3900
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               Caption         =   "接口包拥有者"
               Height          =   240
               Index           =   40
               Left            =   120
               TabIndex        =   258
               Top             =   4380
               Width           =   1440
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               Caption         =   "检查方案号"
               Height          =   180
               Index           =   42
               Left            =   240
               TabIndex        =   257
               Top             =   4845
               Width           =   900
            End
            Begin VB.Label lab 
               Caption         =   "序列方案号"
               Height          =   180
               Index           =   43
               Left            =   2160
               TabIndex        =   256
               Top             =   4845
               Width           =   975
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               Caption         =   "WEB观片地址"
               Height          =   240
               Index           =   38
               Left            =   255
               TabIndex        =   255
               Top             =   2925
               Width           =   1320
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               Caption         =   "3D观片类型"
               Height          =   240
               Index           =   44
               Left            =   4380
               TabIndex        =   254
               Top             =   4845
               Width           =   960
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               Caption         =   "关键图像地址"
               Height          =   240
               Index           =   39
               Left            =   135
               TabIndex        =   253
               Top             =   3405
               Width           =   1440
            End
         End
         Begin VB.Label lab 
            Caption         =   "费用执行模式："
            Height          =   270
            Index           =   58
            Left            =   240
            TabIndex        =   306
            Top             =   6045
            Width           =   1305
         End
      End
   End
End
Attribute VB_Name = "frmParPacs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type

Private Enum TNeedType
    tNeedName = 0
    tNeedNo = 1
    tNeedAll = 2
End Enum


Private Const Report_Form_frmReportES  As String = "内镜基本信息"
Private Const Report_Form_frmReportPathology As String = "病理妇科液基薄层信息"
Private Const Report_Form_frmReportUS As String = "B超心脏测量信息"
Private Const Report_Form_frmReportCustom As String = "自定义专科报告"


Private mrsPar As ADODB.Recordset '参数与控件对应记录集（同一个参数可能对应一组多个控件）
Private marrFunc(1) As String
Private mlngPreFind As Long

Private mrsDeptParas As ADODB.Recordset '本科参数表缓存

Private mstrPrivs As String         '本模块的权限
Private mlng科室ID As Long          'IN:当前执行科室ID
Private mlngCur科室ID As Long       '当前科室ID
Private mstrCur科室 As String       '当前科室 编码-名称
Private mstrCanUse科室 As String    '当前可用科室  ID_编码-名称
Private mblnOk As Boolean

Private UserInfo As TYPE_USER_INFO

Private Enum constLst
    lst_PatholInfo = 0
End Enum

Private Enum constTxtLocate
    txt_Par = 0
    txt_Dept = 1
End Enum

Private Enum constChk
    chk_录入外院放射检查信息 = 0
    chk_录入外院超声检查信息 = 1
    chk_常规病理检查质量窗口 = 2
    chk_冰冻病理检查质量窗口 = 3
    chk_细胞病理检查质量窗口 = 4
    chk_会诊病理检查质量窗口 = 5
    chk_尸检病理检查质量窗口 = 6
    chk_快片病理检查质量窗口 = 7
    chk_借阅后自动打印回执单 = 8
    chk_录入外院信息 = 9
End Enum

Private Enum constCbo
    cbo_3D观片类型 = 0
    cbo_采集费用执行模式 = 1    '采集站费用执行模式0-报到时执行，1-检查时执行，2-报告时执行
    cbo_档案标签报表名称 = 2
    cbo_借阅回执报表名称 = 3
    cbo_医技费用执行模式 = 4    '采集站费用执行模式0-报到时执行，1-报告时执行
    cbo_病理费用执行模式 = 5    '病理站费用执行模式0-报到时执行，1-检查时执行，2-报告时执行
End Enum


Private Enum constTxt
    txt_图像删除用户名称 = 0
    txt_图像删除用户密码 = 1
    txt_图像发送用户名称 = 2
    txt_图像发送用户密码 = 3
    txt_图像刻录用户名称 = 4
    txt_图像刻录用户密码 = 5
    txt_服务器IP = 6
    txt_服务器用户名称 = 7
    txt_服务器用户密码 = 8
    txt_WEB观片地址 = 9
    txt_关键图像地址 = 10
    txt_包拥有者 = 11
    'txt_共享目录 = 12
    txt_检查方案号 = 13
    txt_序列方案号 = 14
    txt_巨检描述模板 = 15
    txt_常规描述模板 = 16
    txt_免疫描述模板 = 17
    txt_特染描述模板 = 18
    txt_分子描述模板 = 19
    txt_档案默认查询天数 = 20
    txt_借还默认查询天数 = 21
    txt_病理外院设置 = 22
    txt_检查列表观片地址 = 23
End Enum
'
'Private Enum constListBox
'    lst_住院检查入院诊断 = 0
'End Enum
Private Const SERVICE_START = &H10
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const SC_MANAGER_CONNECT As Long = &H1
Private Const SC_MANAGER_CREATE_SERVICE As Long = &H2
Private Const SC_MANAGER_ENUMERATE_SERVICE As Long = &H4
Private Const SC_MANAGER_LOCK As Long = &H8
Private Const SC_MANAGER_QUERY_LOCK_STATUS As Long = &H10
Private Const SC_MANAGER_MODIFY_BOOT_CONFIG As Long = &H20
Private Const SC_MANAGER_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)

Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long




Private Sub cbo_Change(Index As Integer)
    If Not Me.Visible Then Exit Sub
    
    If Index = 1 Then
        Call SetParChange(cbo, Index, mrsPar)
    Else
        
        Call SetParChange(cbo, Index, mrsPar, True, cbo(Index).Text)
    End If
End Sub

Private Sub ConfigAppNoState()
'------------------------------------------------
'功能：设置检查号相关控件的可用性
'参数：无
'返回：
'------------------------------------------------
    Dim blnUseHisNo As Boolean  '使用HIS的病人ID或医嘱ID
    Dim blnCanOverWrite As Boolean  '“允许检查号重复”
    Dim blnCheckMaxNo As Boolean    '“提取实际最大号码”
    Dim blnChangeNo As Boolean      '“允许手工调整检查号”
    
    blnCanOverWrite = True
    blnCheckMaxNo = True
    blnChangeNo = True
    
    '设置检查号选项的一些逻辑关系
    '（1）患者检查号保持不变，同时需要勾选且灰掉“允许检查号重复”
    '（2）选择了“前缀”，“分隔符”，“年月日”后，禁止且灰掉“提取实际最大号码”
    '（3）检查号保持不变，“本科室类别统一”，保存时自动保存“按照科室递增”；“影像类别统一”，保存时自动保存“按照影像类别递增”
    '（4）选择“医嘱ID”，“病人ID”，禁止且灰掉“允许手工调整检查号”，禁止且灰掉“提取实际最大号码”，允许且灰掉“检查号重复”，禁止配置检查号编码
    '（5）顺序号永远被选中
    '（6）选择了“本科室内自动递增”则，禁止选择“影像类别”前缀。
    
    If OptCode(1).value = True Then
        chkCanOverWrite.value = 1
        blnCanOverWrite = False
    End If
    
    If chkPreText.value = 1 Or chkDelimiter(1).value = 1 Or chkDelimiter(2).value = 1 Or chkYear.value = 1 _
        Or chkMonth.value = 1 Or chkDay.value = 1 Then
        chkCheckMaxNo.value = 0
        blnCheckMaxNo = False
    End If
    
    If (optUseAdviceID.value = True And OptCode(0).value = True) Or (optUsePatientID.value = True And OptCode(1).value = True) Then
        chkChangeNO.value = 0
        blnChangeNo = False
        chkCanOverWrite.value = 1
        blnCanOverWrite = False
        chkCheckMaxNo.value = 0
        blnCheckMaxNo = False
        blnUseHisNo = True
    Else
        blnUseHisNo = False
    End If
    
    chkNumber.value = 1
    
    If OptBuildcode(1).value = True And optPreText(0).value = True Then
        optPreText(0).value = False
        Call MsgBox("检查号按照本科室内自动递增，禁止使用影像类别前缀。", vbOKOnly)
    End If
    
    chkPreText.Enabled = Not blnUseHisNo
    chkDelimiter(1).Enabled = Not blnUseHisNo
    chkDelimiter(2).Enabled = Not blnUseHisNo
    chkYear.Enabled = Not blnUseHisNo
    chkMonth.Enabled = Not blnUseHisNo
    chkDay.Enabled = Not blnUseHisNo
    chkChangeNO.Enabled = blnChangeNo
    chkCanOverWrite.Enabled = blnCanOverWrite
    chkCheckMaxNo.Enabled = blnCheckMaxNo
    chkNumber.Enabled = Not blnUseHisNo
    
    '设置按钮的可用性
    '检查号一致性
    OptBuildcode(0).Enabled = OptCode(0).value
    OptBuildcode(1).Enabled = OptCode(0).value
    optUseAdviceID.Enabled = OptCode(0).value
    OptUnicode(0).Enabled = OptCode(1).value
    OptUnicode(1).Enabled = OptCode(1).value
    optUsePatientID.Enabled = OptCode(1).value
    
    '前缀
    optPreText(0).Enabled = chkPreText.value And chkPreText.Enabled And OptBuildcode(1).value = False
    optPreText(1).Enabled = chkPreText.value And chkPreText.Enabled
    txtPreText.Enabled = (optPreText(1).value And optPreText(1).Enabled)
    
    '分隔符
    cboDelimeter(1).Enabled = chkDelimiter(1).value And chkDelimiter(1).Enabled
    cboDelimeter(2).Enabled = chkDelimiter(2).value And chkDelimiter(2).Enabled
    
    '年
    optYear(0).Enabled = chkYear.value And chkYear.Enabled
    optYear(1).Enabled = chkYear.value And chkYear.Enabled
    
    '顺序号--固定位数
    chkFixedLen.Enabled = chkNumber.value And chkNumber.Enabled
    txtFixedLen.Enabled = chkFixedLen.value And chkFixedLen.Enabled
    txtStartNum.Enabled = chkNumber.value And chkNumber.Enabled
    lblStartNum.Enabled = chkNumber.value And chkNumber.Enabled
End Sub

Private Sub chkCanOverWrite_Click()
    Call ConfigAppNoState
End Sub

Private Sub chkCheckMaxNo_Click()
    Call ConfigAppNoState
End Sub



Private Sub chkDay_Click()
    Call ConfigAppNoState
End Sub

Private Sub chkDelimiter_Click(Index As Integer)
    Call ConfigAppNoState
End Sub

Private Sub chkFixedLen_Click()
    Call ConfigAppNoState
End Sub

Private Sub chkMonth_Click()
    Call ConfigAppNoState
End Sub

Private Sub chkNumber_Click()
    Call ConfigAppNoState
End Sub

Private Sub chkPreText_Click()
    Call ConfigAppNoState
End Sub

Private Sub chkPreView_Click()
    If chkPreView.value = 1 Then
        optMovePreview.Enabled = True
        lblDelayTime.Enabled = True
        txtDelayTime.Enabled = True
        optClickPreview.Enabled = True
    Else
        optMovePreview.Enabled = False
        lblDelayTime.Enabled = False
        txtDelayTime.Enabled = False
        optClickPreview.Enabled = False
    End If
End Sub


Private Sub chkYear_Click()
    Call ConfigAppNoState
End Sub

Private Sub lst_ItemCheck(Index As Integer, Item As Integer)
    Dim blnValue As Boolean
    Dim strValue As String
    Dim i As Long
    Dim blNoPick As Boolean '所有选项都没有勾选

    If Not Me.Visible Then Exit Sub

    Select Case Index
    Case lst_PatholInfo
        blNoPick = True
        strValue = ""
        With lst(lst_PatholInfo)
            For i = 0 To .ListCount - 1
                If strValue <> "" Then strValue = strValue & ","
                
                If .Selected(i) Then
                    blNoPick = False
                    strValue = strValue & "1"
                Else
                    strValue = strValue & "0"
                End If
            Next
            
            '所有选项都没有勾选等同于都勾选
            If blNoPick Then
                strValue = "1,1,1,1,1,1,1,1,1,1"
            End If
        End With
        blnValue = True
    End Select
    Call SetParChange(lst, lst_PatholInfo, mrsPar, blnValue, strValue)
End Sub

Private Sub chkImageLevel_Click()
    txtImageLevel.Enabled = chkImageLevel.value = 1
End Sub

Private Sub ChkCompleteCommit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ChkCompleteCommit.value = 1 Then chkFinallyCompleteCommit.value = 0
End Sub

Private Sub chkFinallyCompleteCommit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkFinallyCompleteCommit.value = 1 Then ChkCompleteCommit.value = 0
End Sub

Private Sub ChkLike_Click()
    TxtLike.Enabled = IIF(ChkLike.value, True, False)
End Sub

Private Sub chkNameColColorCfg_Click()
    If chkNameColColorCfg.value = 1 Then
        chkOrdinaryNameColColorCfg.Enabled = True
    Else
        chkOrdinaryNameColColorCfg.value = 0
        chkOrdinaryNameColColorCfg.Enabled = False
    End If
End Sub

Private Sub chkPetitionCapture_Click()
    cboSaveDevice.Enabled = IIF(chkPetitionCapture.value, True, False)
End Sub

Private Sub chkRefreshInterval_Click()
    txtRefreshInterval.Enabled = IIF(chkRefreshInterval.value, True, False)
End Sub

Private Sub chkReportAfterResult_Click()
    If chkReportAfterResult.value = vbChecked Then
        chkIgnorePosi.Enabled = False
        chkIgnorePosi.value = vbUnchecked
    Else
        chkIgnorePosi.Enabled = True
    End If
End Sub

Private Sub chkReportLevel_Click()
    txtReportLevel.Enabled = chkReportLevel.value = 1
End Sub


Private Sub chkSpecialContent_Click()
    If chkSpecialContent.value = 1 Then
        cboSpecialContent.Enabled = True
    Else
        cboSpecialContent.Enabled = False
    End If
End Sub

Private Sub chkUseQueue_Click()
On Error GoTo ErrHandle
    optNumberRule(0).Enabled = chkUseQueue.value
    optNumberRule(1).Enabled = chkUseQueue.value
        
    'ufgGroupCfg.Enabled = chkUseQueue.value
    'ufgRoomCfg.Enabled = chkUseQueue.value
    'ufgStudyProCfg.Enabled = chkUseQueue.value
    
    'cmdAdd.Enabled = chkUseQueue.value
    'cmdDel.Enabled = chkUseQueue.value
    'cmdModify.Enabled = chkUseQueue.value
    'cmdStudyAcc.Enabled = chkUseQueue.value
    chkSynStudyList.Enabled = chkUseQueue.value
    
    txtValidDays.Enabled = chkUseQueue.value
    txtQueueReport.Enabled = chkUseQueue.value
    cbxPrintQueueNoWay.Enabled = chkUseQueue.value
    chkAutoInQueue.Enabled = chkUseQueue.value
    chkUseQueueMsg.Enabled = chkUseQueue.value
    
    lab(7).Enabled = chkUseQueue.value
    lab(8).Enabled = chkUseQueue.value
    lab(9).Enabled = chkUseQueue.value
    fra(18).Enabled = chkUseQueue.value
    
    'framGroup.Enabled = chkUseQueue.value
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmbDept_Click()
    mlng科室ID = cmbDept.ItemData(cmbDept.ListIndex)
    
    'If stabWorkFlow.Tabs = IIF(InStr(GetPrivFunc(glngSys, 1160), "基本") > 0, 6, 5) Then '判断tab数量，目的是为了确保在装载完tab之后才触发其中的语句
        
        Call Load科室参数
        
    'End If
End Sub

Private Sub cmdAdd_Click()
    Me.lab(4).Tag = "": Me.txtName.Text = "": Me.txtName.Enabled = True
    Me.cmdDel.Enabled = True: Me.cmdSave.Enabled = True: Me.cmdRestore.Enabled = True: cboDevice.Enabled = True: If cboDevice.ListCount > 0 Then cboDevice.ListIndex = 0
    Me.txtName.SetFocus
End Sub

Private Sub cmdAddGroup_Click()
'新增分组信息
On Error GoTo ErrHandle
    Dim lngGroupId As Long
    Dim strGroupName As String
    Dim strPrefix As String
    Dim objFrmAdd As frmTechnicGroup
    Dim lngRow As Long
    
    '调用分组添加窗口
    Set objFrmAdd = New frmTechnicGroup
    If objFrmAdd.ShowGroupCfg(Me, mlng科室ID, lngGroupId, strGroupName, strPrefix) Then
        lngRow = ufgGroupCfg.NewRow
    
        ufgGroupCfg.Text(lngRow, "ID") = lngGroupId
        ufgGroupCfg.Text(lngRow, "组名") = strGroupName
        ufgGroupCfg.Text(lngRow, "分组前缀") = strPrefix
        
        '载入分组执行间
        Call subLoadTechniRoom(lngGroupId)
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdApply_Click()
    '密码框加密处理
    If ValidateData() = False Then Exit Sub
    
    Call Save科室参数
    
    If SavePar(mrsPar, Me) = False Then Exit Sub
End Sub

Private Sub cmdColor_Click(Index As Integer)
    dlgColor.Color = shpColor(Index).FillColor
    dlgColor.ShowColor
    shpColor(Index).FillColor = dlgColor.Color
End Sub

Private Sub cmdDefault_Click()
    Call subLoadListDefColorConfig
End Sub

Private Sub cmdDel_Click()
    Dim strSQL As String
    
    If MsgBox("真的删除执行间“" & Trim(Me.txtName.Text) & "”吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
    Err = 0: On Error GoTo ErrHand
    
        strSQL = "zl_医技执行房间_Delete(" & Val(mlng科室ID) & ",'" & Trim(Me.txtName.Text) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Call subLoadRoomConfig
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdDelGroup_Click()
On Error GoTo ErrHandle
    Dim strSQL As String
    Dim lngGroupId As Long
    Dim lngMsgResult As Long
    
    If Not ufgGroupCfg.IsSelectionRow Then
        MsgBox "请选择需要删除的分组数据。", vbOKOnly, "提示"
        Exit Sub
    End If
    
    lngMsgResult = MsgBox("是否确认删除该分组数据? 删除后分组将不可恢复。", vbYesNo, "提示")
    If lngMsgResult = vbNo Then Exit Sub
    
    
    lngGroupId = ufgGroupCfg.KeyValue(ufgGroupCfg.SelectionRow)
    
    strSQL = "zl_影像执行分组_Del(" & lngGroupId & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "删除执行分组")
    
    Call ufgRoomCfg.ClearListData
    Call ufgGroupCfg.DelCurRow(False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdModify_Click()
'修改分组信息
On Error GoTo ErrHandle
    Dim lngGroupId As Long
    Dim strGroupName As String
    Dim strPrefix As String
    Dim objFrmUpdate As frmTechnicGroup
    
    If Not ufgGroupCfg.IsSelectionRow Then
        MsgBox "请选择需要修改的分组数据。", vbOKOnly, "提示"
        Exit Sub
    End If
    
    lngGroupId = ufgGroupCfg.KeyValue(ufgGroupCfg.SelectionRow)
    strGroupName = ufgGroupCfg.Text(ufgGroupCfg.SelectionRow, "组名")
    strPrefix = ufgGroupCfg.Text(ufgGroupCfg.SelectionRow, "分组前缀")
    
    '调用分组更新窗口
    Set objFrmUpdate = New frmTechnicGroup
    If objFrmUpdate.ShowGroupCfg(Me, mlng科室ID, lngGroupId, strGroupName, strPrefix) Then
        ufgGroupCfg.CurText("组名") = strGroupName
        ufgGroupCfg.CurText("分组前缀") = strPrefix
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdRestore_Click()
    Call subLoadRoomConfig
End Sub

Private Sub cmdSave_Click()
    Dim blnExist As Boolean, i As Integer
    Dim strSQL As String

    If Trim(Me.txtName.Text) = "" Then
        MsgBox "名称必须输入", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtName.Text), vbFromUnicode)) > Me.txtName.MaxLength Then
        MsgBox "名称超过" & Me.txtName.MaxLength & "的长度限制", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    
    For i = 1 To lvwRoom.ListItems.Count
        If txtName.Text = lvwRoom.ListItems(i).Text Then blnExist = True: Exit For '已经存在
    Next
    '-----------------------------------------
    Err = 0: On Error GoTo ErrHand
    If Me.lab(4).Tag = "" And Not blnExist Then
        strSQL = "zl_医技执行房间_Insert(" & Val(mlng科室ID) & ",'" & Trim(Me.txtName.Text) & "','" & zlStr.NeedCode(cboDevice.Text) & "','" & txtNoPrefix.Text & "')"
    Else
        strSQL = "zl_医技执行房间_Update(" & Val(mlng科室ID) & ",'" & Trim(Me.lab(4).Tag) & "','" & Trim(Me.txtName.Text) & "','" & zlStr.NeedCode(cboDevice.Text) & "','" & txtNoPrefix.Text & "')"
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    MsgBox "执行间保存成功！", vbInformation, gstrSysName
    
    Call subLoadRoomConfig
    
    txtName.SetFocus
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdStudyAcc_Click()
'影像检查项目关联设置
On Error GoTo ErrHandle
    Dim lngGroupId As Long
    Dim objStudyAssocia As frmTechnicStudy
    
    If Not ufgGroupCfg.IsSelectionRow Then
        MsgBox "请选择需要进行关联的分组数据。", vbOKOnly, "提示"
        Exit Sub
    End If
    
    lngGroupId = ufgGroupCfg.KeyValue(ufgGroupCfg.SelectionRow)
    
    Set objStudyAssocia = New frmTechnicStudy
    If objStudyAssocia.ShowStudyAssociation(mlng科室ID, lngGroupId, Me) Then
        Call ufgStudyProCfg.ClearListData
        Call subLoadStudyProAssociation(lngGroupId)
    End If
    
Exit Sub
ErrHandle:
If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdTestCon_Click()
On Error GoTo ErrHandle
    Call XWTestDBConnection(txt(6).Text, txt(7).Text, txt(8).Text)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub XWTestDBConnection(ByVal strServerName As String, ByVal strUser As String, ByVal strPwd As String)
'功能： 测试新网SQLServer数据库连接
'参数：
'返回：成功返回空字符
'--------------------------------------------
    Dim cnTest As New ADODB.Connection

    If strServerName = "" Then
        MsgBox "未找到数据库服务器配置信息，请设置。"
        Exit Sub
    End If
    
    On Error Resume Next
    Err = 0
    
    If cnTest.State = adStateOpen Then cnTest.Close
    
    Set cnTest = OraDataOpen(strServerName, strUser, strPwd)
    
    If Err <> 0 Or cnTest Is Nothing Then
        '数据库连接错误
        MsgBox "数据库连接失败。" & vbCrLf & vbCrLf & "错误代码是：" & Err.Number & "；错误描述是： " & Err.Description
        Exit Sub
    End If
    
    MsgBox "数据库连接成功。"
    Exit Sub
Err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As ADODB.Connection
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    Dim cnOra As New ADODB.Connection
    
    On Error Resume Next
    Err = 0
    
    DoEvents
    
    With cnOra
        If .State = adStateOpen Then .Close
        
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        
        If Err <> 0 Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "由于用户已经被禁用，无法登录。", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            'OraDataOpen = Nothing
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo ErrHand
    
    Set OraDataOpen = cnOra
    
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    Err = 0
End Function

Private Sub Form_Activate()
    If Me.Tag = "初始成功" Then
        Call scbFunc_SelectedChanged(scbFunc.Selected)
        Me.Tag = ""
    End If
End Sub

Private Sub Form_Load()
    Dim strCategory As String
    
    mblnOk = False
    mlng科室ID = 0
    mlngCur科室ID = 0
    mstrCur科室 = ""
    mstrCanUse科室 = ""

    Call GetUserInfo
    
    mstrPrivs = gstrPrivs
    
    strCategory = "参数设置" ',基础设置
    
    '图标编号,TaskPanelItem的ID(同时也是参数容器Picture控件数组号),TaskPanelItem的标题;......
    marrFunc(0) = ""
    marrFunc(0) = marrFunc(0) & "102,0,影像流程设置"
    marrFunc(0) = marrFunc(0) & ";101,1,影像医技设置"
    marrFunc(0) = marrFunc(0) & ";103,2,影像采集设置"
    marrFunc(0) = marrFunc(0) & ";100,3,影像病理设置"
    marrFunc(0) = marrFunc(0) & ";105,4,病理归档设置"
    marrFunc(0) = marrFunc(0) & ";106,5,病理借还设置"
    
    'marrFunc(1) = "102,2,科室药房设置"

    '1.初始化快捷面板的一级分类列表,缺省选中第一个
    Call InitSCBItem(scbFunc, strCategory, picTPL.hwnd)
    Call scbFunc.Icons.AddIcons(imgType.Icons)
      
    '2.初始化任务面板的二级分类列表,缺省选中第一个
    Call InitTPLItem(sccFunc, tplFunc, scbFunc.Selected.Caption, marrFunc(0))
    Call tplFunc.Icons.AddIcons(imgFunc.Icons)
    
    '判断是否具备排队叫号权限
    If (InStr(GetPrivFunc(glngSys, 1160), "基本") <= 0) Then
        stabWorkFlow.TabVisible(3) = False
    End If
    
    Call InitData
    Call ShowErrParasMsg(Me, mrsPar)
    Me.Tag = "初始成功"
End Sub


Private Sub lvwRoom_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.txtName.Text = Item.Text
    Me.lab(4).Tag = Me.txtName.Text
    Me.txtNoPrefix.Text = Item.SubItems(2)
        
    SeekIndex cboDevice, Item.SubItems(1), True, , tNeedNo
End Sub

Private Sub OptBuildcode_Click(Index As Integer)
    Call ConfigAppNoState
End Sub

Private Sub OptCode_Click(Index As Integer)
    Call ConfigAppNoState
End Sub

Private Sub optMovePreview_Click()
    If optMovePreview.value = True Then
        txtDelayTime.Enabled = True
        lblDelayTime.Enabled = True
    End If
End Sub

Private Sub optPreText_Click(Index As Integer)
    Call ConfigAppNoState
End Sub

Private Sub optReportEditor_Click(Index As Integer)
    Dim hService As Long
    Dim hSCManager As Long

On Error GoTo ErrHandle

    fra(24).Visible = Index = 2
    
    Exit Sub
ErrHandle:
    
End Sub

Private Sub OptUnicode_Click(Index As Integer)
    Call ConfigAppNoState
End Sub

Private Sub optUseAdviceID_Click()
    Call ConfigAppNoState
End Sub

Private Sub optUsePatientID_Click()
    Call ConfigAppNoState
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Dim i As Long
    
    For i = 0 To picPar.UBound
        picPar(i).Visible = (i = Item.ID)
    Next
    
    'PACS参数设置中不存在按科室查询
    'lblLocate(txt_Dept).Visible = (Item.ID = GetFuncID("业务流程控制", marrFunc))
    'txtLocate(txt_Dept).Visible = lblLocate(txt_Dept).Visible
    
    If txtLocate(txt_Dept).Visible Then
        lblPrompt.Left = txtLocate(txt_Dept).Left + txtLocate(txt_Dept).Width + 60
    Else
        lblPrompt.Left = txtLocate(txt_Par).Left + txtLocate(txt_Par).Width + 60
    End If
    
    lblPrompt.Width = cmdOk.Left - lblPrompt.Left - 120
    
    mlngPreFind = 1
    
    tplFunc.Tag = Item.ID   '用于获取当前选中的TaskPanelItem
End Sub

Private Sub Form_Resize()
    Dim i As Long
    Dim objPic As PictureBox
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If picVbar.Left < 1500 Then picVbar.Left = 1500
    If picVbar.Left > Me.ScaleWidth - 3000 Then picVbar.Left = Me.ScaleWidth - 3000
    picVbar.Top = 0
    
    picFunc.Width = picVbar.Left + picVbar.Width
    
    For Each objPic In picPar
        objPic.Top = 0
        objPic.Left = 0
        objPic.Width = Me.ScaleWidth - objPic.Left
        objPic.Height = Me.ScaleHeight - PicBottom.ScaleHeight
    Next
End Sub


Private Sub scbFunc_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub picBottom_Resize()
    cmdApply.Left = PicBottom.ScaleWidth - cmdApply.Width - 120
    cmdCancel.Left = cmdApply.Left - cmdCancel.Width - 120
    cmdOk.Left = cmdCancel.Left - cmdOk.Width - 120
End Sub


Private Sub picFunc_Resize()
    scbFunc.Top = picFunc.ScaleTop
    scbFunc.Left = picFunc.ScaleLeft + 45
    scbFunc.Width = picFunc.ScaleWidth - picVbar.Width - 45
    scbFunc.Height = picFunc.ScaleHeight
    
    picVbar.Height = picFunc.ScaleHeight
End Sub

Private Sub picTPL_Resize()
    sccFunc.Left = picTPL.ScaleLeft
    sccFunc.Width = picTPL.ScaleWidth
    
    tplFunc.Left = picTPL.ScaleLeft
    tplFunc.Top = sccFunc.Top + sccFunc.Height
    tplFunc.Height = picTPL.ScaleHeight - sccFunc.Height
    tplFunc.Width = picTPL.ScaleWidth
End Sub


Private Sub picVbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picVbar.Left = IIF(picVbar.Left + X < 2000, 2000, picVbar.Left + X)
        Call Form_Resize
    End If
End Sub

Private Sub scbFunc_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    If Me.Visible Then
        Call InitTPLItem(sccFunc, tplFunc, Item.Caption, marrFunc(Item.ID - 1)) 'ID是从1开始的（因为同时为图标序号）,数组是从0开始
        Call tplFunc_ItemClick(tplFunc.Groups(1).Items(1))
    End If
End Sub


Public Sub LocateFuncItem(ByVal lngFunc As Long)
'功能：根据ID选中一级和二级分类
    Dim i As Long, j As Long, lngId As Long
    Dim arrTmp As Variant
    Dim n As Long
    
    For i = 0 To UBound(marrFunc)
        arrTmp = Split(marrFunc(i), ";")
        For j = 0 To UBound(arrTmp)
            lngId = Split(arrTmp(j), ",")(1)
            If lngFunc = lngId Then
                tplFunc.Tag = lngId
                Set scbFunc.Selected = scbFunc(i)
                
                For n = 1 To tplFunc.Groups(1).Items.Count
                    tplFunc.Groups(1).Items(n).Selected = tplFunc.Groups(1).Items(n).ID = lngId
                Next
            End If
        Next
    Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not mblnOk Then
        mrsPar.Filter = "(修改状态=1 ANd ErrType =Null) OR  (修改状态=1 And ErrType=" & PET_值超限 & ")"
        If mrsPar.RecordCount > 0 Then
            If MsgBox("你已修改部分参数，如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
    Set mrsPar = Nothing
End Sub

Private Sub InitData()
'功能：初始化界面控件,读取并加载数据
    
    '1.初始化变量
    mlngPreFind = 1
    
    Call InitSystemPara
    
    
    '2.初始化界面控件
    Call InitEnv
    
    
    '3.加载系统参数
    Call LoadPar
    
    
    '载入工作流程参数
'    Call Load科室参数
End Sub

Private Sub LoadPar()
'功能：读取并加载参数到界面控件
    Dim strValue As String, strTmp As String
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim arrObj As Variant  '数组对象：模块1,参数号1,控件对象1,模块2,参数号2,控件对象2,......
    

    Set rsTmp = GetPar(mrsPar, p影像观片设置 & _
                            "," & p影像医技设置 & _
                            "," & p影像采集设置 & _
                            "," & p影像病理设置 & _
                            "," & p病理归档设置 & _
                            "," & p病理借还设置)

     '1.设置CheckBox类参数
    strTmp = p影像医技设置 & ":录入外院信息:" & chk_录入外院放射检查信息 & _
            "," & p影像采集设置 & ":录入外院信息:" & chk_录入外院超声检查信息 & _
            "," & p影像病理设置 & ":常规质量窗口:" & chk_常规病理检查质量窗口 & _
            "," & p影像病理设置 & ":冰冻质量窗口:" & chk_冰冻病理检查质量窗口 & _
            "," & p影像病理设置 & ":细胞质量窗口:" & chk_细胞病理检查质量窗口 & _
            "," & p影像病理设置 & ":会诊质量窗口:" & chk_会诊病理检查质量窗口 & _
            "," & p影像病理设置 & ":尸检质量窗口:" & chk_尸检病理检查质量窗口 & _
            "," & p影像病理设置 & ":快速石蜡质量窗口:" & chk_快片病理检查质量窗口 & _
            "," & p影像病理设置 & ":录入外院信息:" & chk_录入外院信息 & _
            "," & p病理借还设置 & ":借阅确认后自动打印回执:" & chk_借阅后自动打印回执单

    Call SetParToControl(strTmp, mrsPar, chk)
    
    
    rsTmp.Filter = "模块=" & p影像病理设置 & " and 参数名='录入外院信息'"
    If rsTmp.RecordCount > 0 Then
        If Val(NVL(rsTmp!参数值)) = 1 Then
            txt(txt_病理外院设置).Enabled = True
        Else
            txt(txt_病理外院设置).Enabled = False
        End If
    End If
    

    '2.设置ComboBox类参数
    strTmp = p影像观片设置 & ":XW3D观片类型:" & cbo_3D观片类型 & _
            "," & p病理归档设置 & ":档案标签报表名称:" & cbo_档案标签报表名称 & _
            "," & p病理借还设置 & ":借阅回执报表名称:" & cbo_借阅回执报表名称
    Call SetParToControl(strTmp, mrsPar, cbo, 3)
    
    strTmp = p影像采集设置 & ":采集费用执行模式:" & cbo_采集费用执行模式
    Call SetParToControl(strTmp, mrsPar, cbo)
    
    strTmp = p影像医技设置 & ":医技费用执行模式:" & cbo_医技费用执行模式
    Call SetParToControl(strTmp, mrsPar, cbo)
    
    strTmp = p影像病理设置 & ":病理费用执行模式:" & cbo_病理费用执行模式
    Call SetParToControl(strTmp, mrsPar, cbo)
    
    
    
'
'    '3.设置UpDown类参数
'    strTmp = "0:5:" & ud_补录医嘱识别间隔
'    Call SetParToControl(strTmp, mrsPar, ud)     'mrsPar存储的控件名是txtUD
'
    '4.设置TextBox类参数
'
    '"," & p影像观片设置 & ":13:" & txt_共享目录 & _'
    
    strTmp = p影像观片设置 & ":XW删除图像用户名:" & txt_图像删除用户名称 & _
            "," & p影像观片设置 & ":XW删除图像密码:" & txt_图像删除用户密码 & _
            "," & p影像观片设置 & ":XW发送图像用户名:" & txt_图像发送用户名称 & _
            "," & p影像观片设置 & ":XW发送图像密码:" & txt_图像发送用户密码 & _
            "," & p影像观片设置 & ":XW光盘刻录用户名:" & txt_图像刻录用户名称 & _
            "," & p影像观片设置 & ":XW光盘刻录密码:" & txt_图像刻录用户密码 & _
            "," & p影像观片设置 & ":XW数据库服务器IP:" & txt_服务器IP & _
            "," & p影像观片设置 & ":XW数据库服务器用户名:" & txt_服务器用户名称 & _
            "," & p影像观片设置 & ":XW数据库服务器密码:" & txt_服务器用户密码 & _
            "," & p影像观片设置 & ":XWWEB观片地址:" & txt_WEB观片地址 & _
            "," & p影像观片设置 & ":XW关键图像地址:" & txt_关键图像地址 & _
            "," & p影像观片设置 & ":XWOracle拥有者:" & txt_包拥有者 & _
            "," & p影像观片设置 & ":XW检查方案号:" & txt_检查方案号 & _
            "," & p影像观片设置 & ":XW序列方案号:" & txt_序列方案号 & _
            "," & p影像病理设置 & ":巨检描述模板:" & txt_巨检描述模板 & _
            "," & p影像病理设置 & ":常规报告模板:" & txt_常规描述模板 & _
            "," & p影像病理设置 & ":免疫报告模板:" & txt_免疫描述模板 & _
            "," & p影像病理设置 & ":特染报告模板:" & txt_特染描述模板 & _
            "," & p影像病理设置 & ":分子报告模板:" & txt_分子描述模板 & _
            "," & p影像病理设置 & ":外院单位结构分类:" & txt_病理外院设置 & _
            "," & p病理归档设置 & ":档案默认查询天数:" & txt_档案默认查询天数 & _
            "," & p病理借还设置 & ":借阅默认查询天数:" & txt_借还默认查询天数 & _
            "," & p影像观片设置 & ":XWWeb检查列表观片地址:" & txt_检查列表观片地址


            
            
    Call SetParToControl(strTmp, mrsPar, txt)
    
    '特殊实现
    rsTmp.Filter = "参数名='档案标签报表名称'"
    If rsTmp.RecordCount > 0 Then
        cbo(2).Text = "" & NVL(rsTmp!参数值)
        Call SetParRelation(cbo, cbo_档案标签报表名称, mrsPar, CStr(NVL(rsTmp!参数名)), p病理归档设置)
    End If
    
    rsTmp.Filter = "参数名='借阅回执报表名称'"
    If rsTmp.RecordCount > 0 Then
        cbo(3).Text = "" & NVL(rsTmp!参数值)
        Call SetParRelation(cbo, cbo_借阅回执报表名称, mrsPar, CStr(NVL(rsTmp!参数名)), p病理借还设置)
    End If
    
    rsTmp.Filter = "参数名='外院单位结构分类'"
    If rsTmp.RecordCount > 0 Then
        txt(txt_病理外院设置).Text = "" & NVL(rsTmp!参数值)
        Call SetParRelation(txt, txt_病理外院设置, mrsPar, CStr(NVL(rsTmp!参数名)), p影像病理设置)
    End If
    
    
    rsTmp.Filter = "参数名='取材内容设置'"
    If rsTmp.RecordCount > 0 Then
        strValue = "" & rsTmp!参数值
        With lst(lst_PatholInfo)
            For i = 0 To .ListCount - 1
                If Val(Split(strValue, ",")(i)) = 1 Then
                    .Selected(i) = True
                End If
            Next
        End With
        Call SetParRelation(lst, lst_PatholInfo, mrsPar, CStr(NVL(rsTmp!参数名)), p影像病理设置)
    End If
    
    
            
            
        
    '密码框解密处理
    txt(1).Text = GetDecryptionPassW(txt(1).Text)
    txt(3).Text = GetDecryptionPassW(txt(3).Text)
    txt(5).Text = GetDecryptionPassW(txt(5).Text)
    txt(8).Text = GetDecryptionPassW(txt(8).Text)
'    rsTmp.Filter = "参数=2"
'    While Not rsTmp.EOF
'
'        strValue = "" & rsTmp!参数值
'        Select Case rsTmp!参数名
'            Case "档案标签报表名称"
'                cbo(2).Text = strValue
'                'Call SetParRelation(cbo, cbo_档案标签报表名称, mrsPar)
'            Case "借阅回执报表名称"
'                cbo(3).Text = strValue
'                'Call SetParRelation(cbo, cbo_借阅回执报表名称, mrsPar)
'        End Select
'
'        rsTmp.MoveNext
'    Wend
'
'
    '5.设置ListBox类参数
'    strTmp = p住院医嘱下达 & ":4:" & lst_住院检查入院诊断
'    Call SetParToControl(strTmp, mrsPar, lst)
'
'    '6.设置OptionButton类参数
'    arrObj = Array(p门诊医嘱下达, 45, opt抗菌目的门诊, _
'                    p住院医嘱下达, 51, opt抗菌目的住院)
'    Call SetParToControl("", mrsPar, arrObj)
'
'
'    '7.其他系统参数
'    rsTmp.Filter = "模块=0"
'    Do Until rsTmp.EOF
'        strValue = "" & rsTmp!参数值
'        Select Case rsTmp!参数号
'        Case 70
'            ud(ud_过敏登记有效天数).value = IIF(Val(strValue) = 0, 1, Val(strValue))
'
'            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "") '已有CheckBox控件，所以需要再产生一条记录
'            Call SetParRelation(txtUD, ud_过敏登记有效天数, mrsPar)
'
'        Case 233
'            Call Load不写超量科室(strValue)
'            Call SetParRelation(vsUnWriteDept, 0, mrsPar, rsTmp!参数号)
'        End Select
'
'        rsTmp.MoveNext
'    Loop
'
'    '8.其他模块参数设置
'    rsTmp.Filter = "模块=" & p门诊医嘱下达
'    Do Until rsTmp.EOF
'        strValue = "" & rsTmp!参数值
'        Select Case rsTmp!参数号
'
'        End Select
'        rsTmp.MoveNext
'    Loop
'
End Sub

Private Sub InitEnv()
''功能：初始化界面控件，加载基础数据
On Error GoTo ErrHandle
    
    Call subInitTechincRoom
    
    Call LoadCheckNoDelimeter   '必须要放在subInitDepartInfo前面，确保先初始化分隔符内容，再从数据库读取参数
    
    Call subInitDepartInfo
    Call LoadPathol
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    '密码框加密处理
    If ValidateData() = False Then Exit Sub
    
    Call Save科室参数
    
    If SavePar(mrsPar, Me) = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub

Private Function ValidateData() As Boolean
'功能：验证数据的有效性
    Dim intTxtLen As Integer
    
    If txtImageLevel.Enabled Then
        '将中文状态下的 逗号替换成英文状态
        txtImageLevel.Text = Replace(txtImageLevel.Text, "，", ",")
        
        intTxtLen = Len(txtImageLevel.Text) - Len(Replace(txtImageLevel.Text, ",", ""))
        
        If intTxtLen > 3 Or intTxtLen < 1 Then
            MsgBox "影像等级最少为2种，最多为4种，请重新填写。", vbOKOnly, "提示信息"
            txtImageLevel.Text = NVL(GetDeptPara(mlng科室ID, "影像质量等级", "甲,乙"))
            txtImageLevel.SetFocus
            Exit Function
        End If
    End If
    
    
    If txtReportLevel.Enabled Then
        '将中文状态下的 逗号替换成英文状态
        txtReportLevel.Text = Replace(txtReportLevel.Text, "，", ",")
        
        intTxtLen = Len(txtReportLevel.Text) - Len(Replace(txtReportLevel.Text, ",", ""))
        
        If intTxtLen > 3 Or intTxtLen < 1 Then
            MsgBox "报告等级最少为2种，最多为4种，请重新填写。", vbOKOnly, "提示信息"
            txtReportLevel.Text = NVL(GetDeptPara(mlng科室ID, "报告质量等级", "甲,乙"))
            txtReportLevel.SetFocus
            Exit Function
        End If
    End If
    
    ValidateData = True
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub txtAudit_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtCheckIn_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtDelayTime_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtEnreg_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtFixedLen_Change()
    If Val(txtFixedLen.Text) > 18 Then
        MsgBox "固定位数最多为18位，请重新填写。", vbOKOnly, "提示信息"
        txtFixedLen.Text = ""
    End If
End Sub

Private Sub txtFixedLen_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub


Private Sub TxtLike_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtLocate_Change(Index As Integer)
    If Index = txt_Dept Then
        mlngPreFind = 1
    ElseIf Index = txt_Par Then
        txtLocate(Index).Tag = ""
    End If
End Sub

Private Sub txtLocate_GotFocus(Index As Integer)
    txtLocate(Index).SelStart = 0
    txtLocate(Index).SelLength = Len(txtLocate(Index).Text)
End Sub

Private Sub txtLocate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim strFind As String

        If Trim(txtLocate(Index).Text) = "" Then Exit Sub
        strFind = UCase(Trim(txtLocate(Index).Text))

        Select Case Index
        Case txt_Par
            Call LocatePar(txtLocate(Index), Me)
        End Select
    End If
End Sub

Private Sub lst_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(lst, Index, mrsPar)
End Sub

Private Sub chk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(chk, Index, mrsPar)
End Sub

Private Sub txt_Change(Index As Integer)
On Error GoTo ErrHandle
    If Me.Visible = False Then Exit Sub
    
    If Index = 1 Or Index = 3 Or Index = 5 Or Index = 8 Then
        Call SetParChange(txt, Index, mrsPar, True, GetEncryptionPassW(txt(Index).Text))
    Else
        Call SetParChange(txt, Index, mrsPar)
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = Asc(gstrParSplit1) Or KeyAscii = Asc(gstrParSplit2) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, Index, mrsPar)
End Sub


Private Sub cbo_GotFocus(Index As Integer)
    Call SetParTip(cbo, Index, mrsPar)
End Sub



Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub




Private Sub cbo_Click(Index As Integer)
    Dim blnValue As Boolean, strValue As String

    If Not Me.Visible Then Exit Sub
    
    Call SetParChange(cbo, Index, mrsPar)
    
'    Select Case Index
'    Case cbo_采集费用执行模式   '按索引进行保存
'        Call SetParChange(cbo, Index, mrsPar)
'    Case Else       '按文本内容进行保存
'        Call SetParChange(cbo, Index, mrsPar)
'    End Select
'
'    If Me.Visible Then
'        Call SetParChange(cbo, Index, mrsPar, blnValue, strValue)
'    End If
    
End Sub

Private Sub chk_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(chk, Index, mrsPar)   'Call SetParChange(chk, Index, mrsPar, blnValue, strValue)
    End If
    
    If Index = chk_录入外院信息 Then
        txt(txt_病理外院设置).Enabled = chk(chk_录入外院信息).value
    End If
End Sub


Private Sub txtName_GotFocus()
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 100
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Or KeyAscii = 95 Then KeyAscii = 0
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtPreText_Change()
    Dim lngSet As Long
    
    '判断是否包含小写字符，如果包含，则提示用户，并自动改成大写
    If txtPreText.Text <> UCase(txtPreText.Text) Then
        lngSet = txtPreText.SelStart
        txtPreText.Text = UCase(txtPreText.Text)
        txtPreText.SelStart = lngSet
    End If
End Sub

Private Sub txtRefreshInterval_Change()
    If Val(txtRefreshInterval.Text) > 600 Then
        txtRefreshInterval.Text = 600
    End If
End Sub

Private Sub txtRefreshInterval_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub


Private Sub txtRefreshInterval_LostFocus()
    If Val(txtRefreshInterval.Text) < 10 Then
        txtRefreshInterval.Text = 10
    End If
End Sub

Private Sub txtReport_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtStartNum_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtStudy_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtValidDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtViewHistoryImageDays_Change()
    If Val(txtViewHistoryImageDays.Text) > 15 Then
        txtViewHistoryImageDays.Text = 15
    End If
End Sub

Private Sub txtViewHistoryImageDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtViewHistoryImageDays_LostFocus()
    If Val(txtViewHistoryImageDays.Text) <= 0 Then
        txtViewHistoryImageDays.Text = 1
    End If
End Sub

Private Sub Txt默认天数_Change()
    If Val(Txt默认天数.Text) > 15 Then
        Txt默认天数.Text = 15
    End If
End Sub

Private Sub Txt默认天数_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub Txt默认天数_LostFocus()
    If Val(Txt默认天数.Text) <= 0 Then
        Txt默认天数.Text = 1
    End If
End Sub


Private Sub ufgGroupCfg_OnSelChange()
On Error GoTo ErrHandle
    Dim lngGroupId As Long
    lngGroupId = Val(ufgGroupCfg.CurKeyValue)
    
    '载入医技执行房间
    Call subLoadTechniRoom(lngGroupId)
    
    '载入分组检查项目关联
    Call subLoadStudyProAssociation(lngGroupId)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgRoomCfg_OnDblClick()
'双击执行间时，进行分组修改处理
On Error GoTo ErrHandle
    Call cmdModify_Click
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgStudyProCfg_OnDblClick()
'双击影像检查项目时，进行关联配置处理
On Error GoTo ErrHandle
    Call cmdStudyAcc_Click
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub Load科室参数()

    Call subLoadWorkFlowConfig      '读取科室流程参数
    Call subLoadRoomConfig          '读取执行间参数
    Call subLoadInputConfig         '读取录入项配置参数
    
    If stabWorkFlow.TabVisible(3) = True Then
        Call subLoadQueueGroupConfig    '读取队列分组配置参数
    End If
    
    Call subLoadSpecifyReportItemName
    Call subLoadReportConfig        '读取报告编辑器配置参数
    Call subLoadListColorConfig     '读取列表颜色配置参数
End Sub


Private Sub Save科室参数()
'保存医嘱内容定义
    On Error GoTo ErrHandle
        Call subSaveWorkFlowConfig
        Call subSaveInputConfig
        
        If stabWorkFlow.TabVisible(3) = True Then
            Call subSaveQueueGroupConfig
        End If
        
        Call subSaveReportConfig
        Call subSaveListColorConfig
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


'************************************************************************************************************************************
'************************************************************************************************************************************
    Private Sub subSaveListColorConfig()
        Dim i As Integer, strInput As String
        Dim strSQL As String
        
        If mlng科室ID < 0 Then Exit Sub
        
          
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '已登记','" & shpColor(8).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '已报到','" & shpColor(1).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '处理中','" & shpColor(2).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '已检查','" & shpColor(0).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '报告中','" & shpColor(3).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '已报告','" & shpColor(4).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '已审核','" & shpColor(6).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '已完成','" & shpColor(7).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '审核中','" & shpColor(5).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '已拒绝','" & shpColor(9).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '已驳回','" & shpColor(10).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '登记后提醒','" & Val(txtEnreg.Text) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '报到后提醒','" & Val(txtCheckIn.Text) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '检查后提醒','" & Val(txtStudy.Text) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '报告后提醒','" & Val(txtReport.Text) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '审核后提醒','" & Val(txtAudit.Text) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '姓名颜色区分','" & chkNameColColorCfg.value & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '缺省类型病人姓名颜色区分','" & chkOrdinaryNameColColorCfg.value & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '颜色显示类型','" & IIF(optListColorMark(0).value = True, 0, 1) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End Sub
    
    Private Sub subSaveReportConfig()
        Dim intMatch As Integer
        Dim strSQL As String
        
        On Error GoTo ErrHand
        
        If mlng科室ID < 0 Then Exit Sub
        
        
        If optReportEditor(0).value = True Then         '电子病历编辑器
            intMatch = 0
        ElseIf optReportEditor(1).value = True Then     'PACS报告编辑器
            intMatch = 1
        ElseIf optReportEditor(2).value = True Then     '报告文档编辑器
            intMatch = 2
        End If
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '报告编辑器','" & intMatch & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '显示报告图像','" & chkShowImage.value & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '报告缩略图数量','" & txtMinImageCount.Text & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '显示视频采集','" & chkShowVideoCapture.value & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '打印后退出','" & chkExitAfterPrint.value & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '显示专科报告','" & chkSpecialContent.value & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '专科报告页','" & cboSpecialContent.Text & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        If optWordDblClick(0).value = True Then         '报告词句双击后直接写入报告
            intMatch = 0
        ElseIf optWordDblClick(1).value = True Then     '报告词句双击后打开编辑窗口
            intMatch = 1
        End If
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '报告词句双击操作','" & intMatch & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        If optImageDblClick(0).value = True Then         '缩略图双击后直接写入报告
            intMatch = 0
        ElseIf optImageDblClick(1).value = True Then     '缩略图双击后打开图像编辑窗口
            intMatch = 1
        End If
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '缩略图双击操作','" & intMatch & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '检查所见名称','" & txtCheckView.Text & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '诊断意见名称','" & txtResult.Text & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '建议名称','" & txtAdvice.Text & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        If optShowWord(0).value = True Then         '直接显示词句示范
            intMatch = 0
        ElseIf optShowWord(1).value = True Then     '双击标题后显示词句示范
            intMatch = 1
        End If
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '显示词句示范','" & intMatch & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '审核打印后允许回退','" & chkUntreadPrinted.value & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        If optReportEditor(2) Then
            strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '查看历史报告','" & IIF(optHistoryReportEditor(0).value, 0, 1) & "')"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '打印格式选择方式','" & IIF(optPrintFormat(0).value, 0, 1) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '单选报告格式','" & IIF(chkPrintFormat.value, 1, 0) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        Exit Sub
ErrHand:
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End Sub

    Public Sub subSaveQueueGroupConfig()
    '保存配置参数
        If mlng科室ID < 0 Then Exit Sub
    
        SetDeptPara mlng科室ID, "启动排队叫号", chkUseQueue.value
        SetDeptPara mlng科室ID, "排队叫号编码规则", IIF(optNumberRule(0).value, 0, 1)
        SetDeptPara mlng科室ID, "排队数据保存天数", Val(txtValidDays.Text)
        SetDeptPara mlng科室ID, "排队单报表编号", txtQueueReport.Text
        SetDeptPara mlng科室ID, "同步定位检查列表", chkSynStudyList.value
        SetDeptPara mlng科室ID, "报到时分配默认执行间", chkSelectRoom.value
        SetDeptPara mlng科室ID, "排队单打印方式", cbxPrintQueueNoWay.ListIndex
        SetDeptPara mlng科室ID, "启用排队消息处理", chkUseQueueMsg.value
        SetDeptPara mlng科室ID, "报到后自动排队", chkAutoInQueue.value
    End Sub
    
    
    Private Function SetDeptPara(ByVal lngDeptID As Long, ByVal varPara As String, ByVal strValue As String) As Boolean
    '功能：设置指定的参数值
    '参数：lngDept=科室ID
    '      varPara=参数名
    '      strValue=参数名值
    '返回：设置是否成功
        Dim strSQL As String
        
        On Error GoTo errH
            
        strSQL = "ZL_影像流程参数_UPDATE(" & lngDeptID & ",'" & varPara & "','" & strValue & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "SetPara")
        
        '设置成功后清除缓存
        Set mrsDeptParas = Nothing
        
        SetDeptPara = True
        Exit Function
errH:
        If ErrCenter() = 1 Then Resume
    End Function


    Private Sub subSaveInputConfig()
        Call subSaveInputItem(0)
        Call subSaveInputItem(1)
    End Sub
    
    
    Private Sub subSaveInputItem(intType As Integer)
        Dim i As Integer, strInput As String
        Dim strSQL As String
        
        strInput = ""
        If intType = 0 Then
            For i = 0 To ChkMouseMove.UBound
                If ChkMouseMove(i).value = 1 Then strInput = strInput & "|" & ChkMouseMove(i).Caption
            Next
        Else
            For i = 0 To ChkInput.UBound
                If ChkInput(i).value = 1 Then strInput = strInput & "|" & ChkInput(i).Caption
            Next
        End If
        
        strSQL = "ZL_影像流程参数_UPDATE( " & mlng科室ID & ", '" & IIF(intType = 0, "输入控制", "必录控制") & "','" & strInput & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End Sub
    
    
    Private Sub subSaveWorkFlowConfig()
        Dim strTemp As String
        Dim lngTemp As Long
        
        On Error GoTo ErrHand
    
        SetDeptPara mlng科室ID, "启用申请单扫描", chkPetitionCapture.value        '启用申请单扫描 参数保存
        
        SetDeptPara mlng科室ID, "符合情况判定", chkConformDetermine.value         '符合情况判定 参数保存
'        SetDeptPara mlng科室ID, "危急情况判断", chkCriticalValues.value           '危急情况判断 参数保存
        
        SetDeptPara mlng科室ID, "忽略结果阴阳性", chkIgnorePosi.value
        SetDeptPara mlng科室ID, "无影像诊断为阴性", chkReportAfterResult.value
        SetDeptPara mlng科室ID, "诊断结果默认阳性", chkDefaultPosi.value   '诊断结果默认阳性 参数保存
        
        SetDeptPara mlng科室ID, "影像质量判定", chkImageLevel.value           '影像质量判定 参数保存
        SetDeptPara mlng科室ID, "影像质量等级", txtImageLevel.Text            '图像质量等级 参数保存
        SetDeptPara mlng科室ID, "报告质量判定", chkReportLevel.value           '报告质量判定 参数保存
        SetDeptPara mlng科室ID, "报告质量等级", txtReportLevel.Text           '报告质量等级 参数保存
        
        SetDeptPara mlng科室ID, "诊断结果提示类型", IIF(optResultInput(0).value = True, 0, IIF(optResultInput(1).value = True, 1, 2))
        
        SetDeptPara mlng科室ID, "无报告完成后直接完成", ChkFinishCommit.value
        SetDeptPara mlng科室ID, "采图后医生站即可观片", chkCanViewImage.value     '采图后医生站即可观片
        SetDeptPara mlng科室ID, "有图像才能写报告", chkReportAfterImging.value
        
        '检查号设置
        SetDeptPara mlng科室ID, "患者检查号保持不变", IIF(OptCode(1).value, 1, 0)
        SetDeptPara mlng科室ID, "检查号保持不变类别", IIF(OptUnicode(1).value, 1, 0)
        SetDeptPara mlng科室ID, "手工调整检查号", chkChangeNO.value
        SetDeptPara mlng科室ID, "允许检查号重复", chkCanOverWrite.value
        SetDeptPara mlng科室ID, "提取实际最大号码", chkCheckMaxNo.value
        
        SetDeptPara mlng科室ID, "使用患者号", IIF(optUsePatientID.value And optUsePatientID.Enabled, 1, 0)
        SetDeptPara mlng科室ID, "使用医嘱号", IIF(optUseAdviceID.value And optUseAdviceID.Enabled, 1, 0)
        
        If OptCode(0).value = True Then
            SetDeptPara mlng科室ID, "检查号生成方式", IIF(OptBuildcode(1).value, 1, 0)
        Else
            SetDeptPara mlng科室ID, "检查号生成方式", IIF(OptUnicode(1).value, 1, 0)
        End If
        
        If chkPreText.value = 1 Then
            If optPreText(0).value = True Then
                strTemp = 1
            Else
                strTemp = txtPreText.Text
            End If
        Else
            strTemp = ""
        End If
        SetDeptPara mlng科室ID, "检查号前缀", strTemp
        SetDeptPara mlng科室ID, "检查号分隔符1", IIF(chkDelimiter(1).value = 1, Left(cboDelimeter(1).Text, 1), "")
        SetDeptPara mlng科室ID, "检查号分隔符2", IIF(chkDelimiter(2).value = 1, Left(cboDelimeter(2).Text, 1), "")
        SetDeptPara mlng科室ID, "检查号年", IIF(chkYear.value = 1, IIF(optYear(0).value = True, 1, 2), 0)
        SetDeptPara mlng科室ID, "检查号月", chkMonth.value
        SetDeptPara mlng科室ID, "检查号日", chkDay.value
        SetDeptPara mlng科室ID, "检查号起始数", IIF(Val(txtStartNum.Text) = 0, 1, Val(txtStartNum.Text))
        SetDeptPara mlng科室ID, "检查号固定位数", IIF(chkFixedLen.value = 1, Val(txtFixedLen.Text), 0)
        
        SetDeptPara mlng科室ID, "定位片后置", chkLocalizerBackward.value
        SetDeptPara mlng科室ID, "允许交换用户", chkChangeUser.value
        SetDeptPara mlng科室ID, "允许切换用户", chkSwitchUser.value
        SetDeptPara mlng科室ID, "只能填写自己检查的报告", chkTechReportSame.value
        SetDeptPara mlng科室ID, "采集图像者为检查技师", chkWriteCapDoctor.value
        SetDeptPara mlng科室ID, "审核后直接完成", ChkCompleteCommit.value
        SetDeptPara mlng科室ID, "终审后直接完成", chkFinallyCompleteCommit.value
        SetDeptPara mlng科室ID, "打印后直接完成", chkPrintCommit.value
        SetDeptPara mlng科室ID, "终审后直接打印", chkCompletePrint.value
        SetDeptPara mlng科室ID, "登记后直接检查", chkSample.value
        SetDeptPara mlng科室ID, "匹配数据库项目", IIF(optMatch(0).value, 0, IIF(optMatch(1), 1, 2))
        
        SetDeptPara mlng科室ID, "登记时姓名模糊查找天数", IIF(ChkLike.value = 1, Abs(Val(TxtLike.Text)), 0)
        SetDeptPara mlng科室ID, "所有登记病人标记为外来", chkAllPatientIsOutside
        
        If Val(Txt默认天数.Text) > 15 Or Val(Txt默认天数.Text) <= 0 Then
            Txt默认天数.Text = 2
        End If
        SetDeptPara mlng科室ID, "默认过滤天数", Val(Txt默认天数.Text)
        
        If Val(txtViewHistoryImageDays.Text) > 15 Or Val(txtViewHistoryImageDays.Text) <= 0 Then
            txtViewHistoryImageDays.Text = 1
        End If
        SetDeptPara mlng科室ID, "自动打开历史图像天数", Val(txtViewHistoryImageDays.Text)
        
        
        SetDeptPara mlng科室ID, "启动关联病人", chkUseReferencePatient.value
        SetDeptPara mlng科室ID, "平诊需审核才能打报告", chkPrintNeedComplete.value
        
        SetDeptPara mlng科室ID, "拼音名大小写", IIF(optCapital(0).value, 0, IIF(optCapital(1), 1, 2))
        SetDeptPara mlng科室ID, "拼音名分隔符", IIF(optSplitter(0).value, 0, 1)
        
        If cboSaveDevice.Text <> "" Then
            SetDeptPara mlng科室ID, "申请单存储设备号", Split(cboSaveDevice.Text, "-")(0)
        Else
            SetDeptPara mlng科室ID, "申请单存储设备号", ""
        End If
        
        If Abs(Val(txtRefreshInterval.Text)) = 0 Or Abs(Val(txtRefreshInterval.Text)) > 65 Then
            txtRefreshInterval.Text = 10
        End If
        SetDeptPara mlng科室ID, "自动刷新间隔", IIF(chkRefreshInterval.value = 1, Abs(Val(txtRefreshInterval.Text)), 0)
        SetDeptPara mlng科室ID, "报道时自动发送WorkList", chkAutoSendWorkList.value
        SetDeptPara mlng科室ID, "显示附加主述", chkAddons.value
        SetDeptPara mlng科室ID, "显示造影剂", chkReagent.value
        SetDeptPara mlng科室ID, "医生站查看报告", cboViewReport.ListIndex
        SetDeptPara mlng科室ID, "检查切换时定位报告编辑", chkSetFocusWithReport.value
        SetDeptPara mlng科室ID, "姓名默认模糊查询", chkNameFuzzySearch.value
        SetDeptPara mlng科室ID, "姓名查询时间限制", chkNameQueryTimeLimit.value
        
        If chkPreView.value = 1 Then
            If optMovePreview.value Then
                lngTemp = 1
            ElseIf optClickPreview.value Then
                lngTemp = 2
            End If
        Else
            lngTemp = 0
        End If
        
        SetDeptPara mlng科室ID, "缩略图预览方式", lngTemp
        SetDeptPara mlng科室ID, "移动预览延时", Val(txtDelayTime.Text)
         
        Exit Sub
ErrHand:
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End Sub
'************************************************************************************************************************************
'************************************************************************************************************************************
    
    Private Sub subLoadSpecifyReportItemName()
        '装载专科报告名称
        Call cboSpecialContent.Clear
        Call cboSpecialContent.AddItem(Report_Form_frmReportES)
        Call cboSpecialContent.AddItem(Report_Form_frmReportPathology)
        Call cboSpecialContent.AddItem(Report_Form_frmReportUS)
        Call cboSpecialContent.AddItem(Report_Form_frmReportCustom)
    End Sub
    
    
    Private Sub subLoadListDefColorConfig()
    '载入列表默认颜色配置
        shpColor(10).FillColor = ColorConstants.vbYellow
        shpColor(9).FillColor = ColorConstants.vbRed
        shpColor(7).FillColor = ColorConstants.vbGreen
        
        shpColor(0).FillColor = ColorConstants.vbWhite
        shpColor(1).FillColor = ColorConstants.vbWhite
        shpColor(2).FillColor = ColorConstants.vbWhite
        shpColor(3).FillColor = ColorConstants.vbWhite
        shpColor(4).FillColor = ColorConstants.vbWhite
        shpColor(5).FillColor = ColorConstants.vbWhite
        shpColor(6).FillColor = ColorConstants.vbWhite
        shpColor(8).FillColor = ColorConstants.vbWhite
        
        txtEnreg.Text = "0"
        txtCheckIn.Text = "0"
        txtStudy.Text = "0"
        txtReport.Text = "0"
        txtAudit.Text = "0"
    End Sub
    
    Private Sub subLoadListColorConfig()
        Dim strSQL As String
        Dim rsTemp As ADODB.Recordset
        Dim lngTemp As Long
                 
        On Error GoTo Err
        
        
        Call subLoadListDefColorConfig
        
        strSQL = "select ID ,科室ID,参数名,参数值 from 影像流程参数 where 科室ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng科室ID)
        
        While Not rsTemp.EOF
            Select Case rsTemp!参数名
                Case "已登记"
                    shpColor(8).FillColor = Val(NVL(rsTemp!参数值))
                Case "已报到"
                    shpColor(1).FillColor = Val(NVL(rsTemp!参数值))
                Case "处理中"
                    shpColor(2).FillColor = Val(NVL(rsTemp!参数值))
                Case "已检查"
                    shpColor(0).FillColor = Val(NVL(rsTemp!参数值))
                Case "报告中"
                    shpColor(3).FillColor = Val(NVL(rsTemp!参数值))
                Case "已报告"
                    shpColor(4).FillColor = Val(NVL(rsTemp!参数值))
                Case "已审核"
                    shpColor(6).FillColor = Val(NVL(rsTemp!参数值))
                Case "已完成"
                    shpColor(7).FillColor = Val(NVL(rsTemp!参数值))
                Case "审核中"
                    shpColor(5).FillColor = Val(NVL(rsTemp!参数值))
                Case "已拒绝"
                    shpColor(9).FillColor = Val(NVL(rsTemp!参数值))
                Case "已驳回"
                    shpColor(10).FillColor = Val(NVL(rsTemp!参数值))
                Case "登记后提醒"
                    txtEnreg.Text = Val(NVL(rsTemp!参数值))
                Case "报到后提醒"
                    txtCheckIn.Text = Val(NVL(rsTemp!参数值))
                Case "检查后提醒"
                    txtStudy.Text = Val(NVL(rsTemp!参数值))
                Case "报告后提醒"
                    txtReport.Text = Val(NVL(rsTemp!参数值))
                Case "审核后提醒"
                    txtAudit.Text = Val(NVL(rsTemp!参数值))
                Case "颜色显示类型"
                    If Val(NVL(rsTemp!参数值)) = 0 Then
                        optListColorMark(0).value = True
                    Else
                        optListColorMark(1).value = True
                    End If
            End Select
            rsTemp.MoveNext
        Wend
        
        chkNameColColorCfg.value = Val(GetDeptPara(mlng科室ID, "姓名颜色区分", 0))
        If chkNameColColorCfg.value = 0 Then
            chkOrdinaryNameColColorCfg.value = 0
            chkOrdinaryNameColColorCfg.Enabled = False
        Else
            chkOrdinaryNameColColorCfg.Enabled = True
            chkOrdinaryNameColColorCfg.value = Val(GetDeptPara(mlng科室ID, "缺省类型病人姓名颜色区分", 0))
        End If
    
        
        Exit Sub
Err:
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End Sub
    

    Public Sub subLoadReportConfig()
        Dim strSQL As String
        Dim rsTemp As ADODB.Recordset
        Dim lngTemp As Long
        
        optReportEditor(0).value = True '默认使用电子病历编辑器编辑报告
        chkShowImage.value = 0          '默认不显示图像区域
        chkShowVideoCapture.value = 0   '默认不显示视频采集区域
        
        chkSpecialContent.value = 0     '默认不显示专科报告
        cboSpecialContent.Enabled = False
        chkExitAfterPrint.value = 0     '默认打印后不退出
        optWordDblClick(0).value = True '默认双击词句后直接写入报告
        optImageDblClick(0).value = True '默认报告缩略图双击后直接写入报告
        txtCheckView.Text = "检查所见"  '默认为检查所见
        txtResult.Text = "诊断意见"     '默认为诊断意见
        txtAdvice.Text = "建议"         '默认为建议
        optShowWord(0).value = True     '默认为直接显示词句模板
        chkUntreadPrinted.value = 0     '默认为审核打印后不允许回退
         
        On Error GoTo Err
        strSQL = "select ID ,科室ID,参数名,参数值 from 影像流程参数 where 科室ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng科室ID)
        
        While Not rsTemp.EOF
            Select Case rsTemp!参数名
                Case "报告编辑器"
                    If NVL(rsTemp!参数值, 0) = 0 Then
                        optReportEditor(0).value = True
                    ElseIf NVL(rsTemp!参数值, 0) = 1 Then
                        optReportEditor(1).value = True
                    Else
                        optReportEditor(2).value = True
                    End If
                Case "查看历史报告"
                    If NVL(rsTemp!参数值, 0) = 0 Then
                        optHistoryReportEditor(0).value = True
                    Else
                        optHistoryReportEditor(1).value = True
                    End If
                Case "显示报告图像"
                    chkShowImage.value = NVL(rsTemp!参数值, 0)
                Case "报告缩略图数量"
                    txtMinImageCount.Text = NVL(rsTemp!参数值, "8")
                Case "显示视频采集"
                    chkShowVideoCapture.value = NVL(rsTemp!参数值, 0)
                Case "打印后退出"
                    chkExitAfterPrint.value = NVL(rsTemp!参数值, 0)
                Case "显示专科报告"
                    chkSpecialContent.value = NVL(rsTemp!参数值, 0)
                    cboSpecialContent.Enabled = IIF(chkSpecialContent.value = 1, True, False)
                Case "专科报告页"
                    cboSpecialContent.Text = NVL(rsTemp!参数值)
                Case "报告词句双击操作"
                    If NVL(rsTemp!参数值, 0) = 0 Then
                        optWordDblClick(0).value = True
                    Else
                        optWordDblClick(1).value = True
                    End If
                Case "缩略图双击操作"
                    If NVL(rsTemp!参数值, 0) = 0 Then
                        optImageDblClick(0).value = True
                    Else
                        optImageDblClick(1).value = True
                    End If
                Case "检查所见名称"
                    txtCheckView.Text = NVL(rsTemp!参数值, "检查所见")
                Case "诊断意见名称"
                    txtResult.Text = NVL(rsTemp!参数值, "诊断意见")
                Case "建议名称"
                    txtAdvice.Text = NVL(rsTemp!参数值, "建议")
                Case "显示词句示范"
                    If NVL(rsTemp!参数值, 0) = 0 Then
                        optShowWord(0).value = True
                    Else
                        optShowWord(1).value = True
                    End If
                Case "审核打印后允许回退"
                    chkUntreadPrinted.value = NVL(rsTemp!参数值, 0)
                Case "打印格式选择方式"
                If NVL(rsTemp!参数值, 0) = 0 Then
                    optPrintFormat(0).value = True
                Else
                    optPrintFormat(1).value = True
                End If
                Case "单选报告格式"
                    chkPrintFormat.value = IIF(NVL(rsTemp!参数值, 0), 1, 0)
            End Select
            rsTemp.MoveNext
        Wend
        
        If optReportEditor(2).value Then
            fra(24).Visible = True
        Else
            fra(24).Visible = False
        End If
        
        Exit Sub
Err:
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End Sub

    Public Sub subLoadQueueGroupConfig()
    '刷新配置参数
        Dim strSQL As String
        Dim rsTemp As ADODB.Recordset
        Dim lngIndex As Long
    
        On Error GoTo Err
    
        lngIndex = Val(GetDeptPara(mlng科室ID, "排队叫号编码规则", 0))
        txtValidDays.Text = GetDeptPara(mlng科室ID, "排队数据保存天数", 1)
        txtQueueReport.Text = GetDeptPara(mlng科室ID, "排队单报表编号", "")
        chkSynStudyList.value = Val(GetDeptPara(mlng科室ID, "同步定位检查列表", 0))
        chkSelectRoom.value = Val(GetDeptPara(mlng科室ID, "报到时分配默认执行间", 0))
        chkUseQueueMsg.value = Val(GetDeptPara(mlng科室ID, "启用排队消息处理", 1))
        chkAutoInQueue.value = Val(GetDeptPara(mlng科室ID, "报到后自动排队", 1))
        
        '0-不打印，1-自动打印，2-提示打印
        cbxPrintQueueNoWay.ListIndex = Val(GetDeptPara(mlng科室ID, "排队单打印方式", 0))
        
        chkUseQueue.value = Val(GetDeptPara(mlng科室ID, "启动排队叫号", 0))
        
        Call subLoadGroupInf
    
        optNumberRule(lngIndex).value = True
    
        Call chkUseQueue_Click
    
        Exit Sub
Err:
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End Sub

    Private Sub subLoadGroupInf()
    '载入医技分组信息
        Dim strSQL As String
        Dim rsData As ADODB.Recordset
        
        strSQL = "select Id, 组名,分组前缀 from 影像执行分组 where 科室ID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询分组信息", mlng科室ID)
        
        Call ufgGroupCfg.ClearListData
        If rsData.RecordCount <= 0 Then Exit Sub
        
        rsData.Sort = "组名 asc"
        
        Set ufgGroupCfg.AdoData = rsData
        Call ufgGroupCfg.BindData
    End Sub
    
    Private Sub subLoadTechniRoom(ByVal lngGroupId As Long)
    '载入分组所含的医技执行房间
        Dim strSQL As String
        Dim rsData As ADODB.Recordset
        
        strSQL = "select 执行间, 号码前缀 from 医技执行房间 where 分组Id=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询医技执行房间", lngGroupId)
        
        Call ufgRoomCfg.ClearListData
        If rsData.RecordCount <= 0 Then Exit Sub
        
        rsData.Sort = "执行间 asc"
        
        Set ufgRoomCfg.AdoData = rsData
        Call ufgRoomCfg.BindData
    End Sub
    
    Private Sub subLoadStudyProAssociation(ByVal lngGroupId As Long)
    '载入检查项目关联
        Dim strSQL As String
        Dim rsData As ADODB.Recordset
        
        strSQL = "select 名称,编码 from 诊疗项目目录 a, 影像分组关联 b where a.id=b.诊疗项目Id and b.分组Id=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询影像分组关联检查项目", lngGroupId)
        
        Call ufgStudyProCfg.ClearListData
        If rsData.RecordCount <= 0 Then Exit Sub
        
        rsData.Sort = "名称"
        
        Set ufgStudyProCfg.AdoData = rsData
        Call ufgStudyProCfg.BindData
    End Sub

    Private Sub subLoadInputConfig()
        Call subLoadInputItem(0)
        Call subLoadInputItem(1)
    End Sub
    
    Private Sub subLoadInputItem(intType As Integer)
    '载入录入配置
    'intType 0-输入控制，1-必录控制
        Dim i As Integer, strInput As String, j As Integer
        Dim strSQL As String
        Dim rsTemp As ADODB.Recordset
        
        
        If intType = 0 Then
            '初始化关闭移动选择框
            For i = 0 To ChkMouseMove.UBound
                ChkMouseMove(i).value = 0
            Next
        Else
            '初始必录选择框
            For i = 0 To ChkInput.UBound
                ChkInput(i).value = 0
            Next
        End If
        
        strSQL = "select ID ,科室ID,参数值 from 影像流程参数 where 科室ID = [1] and 参数名 = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng科室ID, CStr(IIF(intType = 0, "输入控制", "必录控制")))
        
        If Not rsTemp.EOF Then
            strInput = NVL(rsTemp!参数值)
            For i = 0 To UBound(Split(strInput, "|"))
                If intType = 0 Then
                    For j = 0 To ChkMouseMove.UBound
                        If ChkMouseMove(j).Caption = Split(strInput, "|")(i) Then ChkMouseMove(j).value = 1: Exit For
                    Next
                
                Else
                    For j = 0 To ChkInput.UBound
                        If ChkInput(j).Caption = Split(strInput, "|")(i) Then ChkInput(j).value = 1: Exit For
                    Next
                End If
            Next
        End If
    End Sub


    Private Sub SeekIndex(objCbo As Object, ByVal strText As String, Optional blnEvent As Boolean, Optional blnPreserve As Boolean = False, Optional intIsSearchNo As TNeedType = tNeedName)
    '功能：在ComboBox中查找并定位
    '参数：blnEvent=定位时是否触发Click事件
          'blnPreserve--如果找不到匹配项目，则保持原有项目
          'intIsSearchNo -- 0:通过编码定位,1:通过名字定位,2:用过编码加名字定位
    '说明：未能定位时,设置ListIndex=-1
    '       Cbo.SeekIndex功能比较简单，设置index后会触发事件，不适合使用
        Dim i As Long
    
        For i = 0 To objCbo.ListCount - 1
            If IIF(Abs(intIsSearchNo) = tNeedAll, objCbo.List(i), IIF(Abs(intIsSearchNo) = tNeedNo, zlStr.NeedCode(objCbo.List(i)), zlStr.NeedName(objCbo.List(i)))) = strText Then
                If blnEvent Then
                    objCbo.ListIndex = i
                Else
                    Call zlControl.CboSetIndex(objCbo.hwnd, i)
                End If
                Exit Sub
            End If
        Next
        
        If blnPreserve = True Then
            If blnEvent = False Then
                Call zlControl.CboSetIndex(objCbo.hwnd, objCbo.ListIndex)
            End If
        Else
            If blnEvent Then
                objCbo.ListIndex = -1
            Else
                Call zlControl.CboSetIndex(objCbo.hwnd, -1)
            End If
        End If
        
    End Sub


    Private Sub subLoadRoomConfig()
        Dim ObjItem As ListItem
        Dim rsTemp As New ADODB.Recordset
        Dim strSQL As String
        
        On Error GoTo ErrHand
        
        strSQL = "Select 执行间,检查设备,号码前缀 From 医技执行房间 where 科室id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(Val(mlng科室ID)))
        Me.lvwRoom.ListItems.Clear
        With rsTemp
            Do While Not .EOF
                Set ObjItem = Me.lvwRoom.ListItems.Add(, , !执行间, 1, 1)
                
                ObjItem.SubItems(1) = NVL(!检查设备)
                ObjItem.SubItems(2) = NVL(!号码前缀)
                .MoveNext
            Loop
        End With
        
        Err = 0: On Error Resume Next
        If Me.lvwRoom.ListItems.Count > 0 Then
            Me.lvwRoom.ListItems(1).Selected = True
            Me.lvwRoom.SelectedItem.EnsureVisible
        End If
        
        Err = 0: On Error GoTo 0
        If Me.lvwRoom.ListItems.Count > 0 Then
            Call lvwRoom_ItemClick(Me.lvwRoom.SelectedItem)
            Me.txtName.Enabled = True: cboDevice.Enabled = True
            Me.cmdDel.Enabled = True: Me.cmdSave.Enabled = True: Me.cmdRestore.Enabled = True
        Else
            Me.lab(4).Tag = "": Me.txtName.Text = "": If cboDevice.ListCount > 0 Then cboDevice.ListIndex = 0
            Me.txtName.Enabled = False: cboDevice.Enabled = False
            Me.cmdDel.Enabled = False: Me.cmdSave.Enabled = False: Me.cmdRestore.Enabled = False
        End If
        Exit Sub
ErrHand:
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End Sub


    Private Sub subInitTechincRoom()
        Dim rsTemp As New ADODB.Recordset
        Dim strSQL As String
        
        Me.lvwRoom.ListItems.Clear
        With Me.lvwRoom.ColumnHeaders
            .Clear
            .Add , "名称", "名称", 3000
            .Add , "检查设备", "检查设备", 3000
            .Add , "号码前缀", "号码前缀", 2000
        End With
        
        strSQL = "Select 设备号,设备名 From 影像设备目录 Where 状态=1 and 类型=4"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
        cboDevice.Clear
        Do Until rsTemp.EOF
            cboDevice.AddItem rsTemp!设备号 & "-" & rsTemp!设备名
            rsTemp.MoveNext
        Loop
    End Sub
    
    
    Private Sub subLoadWorkFlowConfig()
        Dim rsTemp As ADODB.Recordset
        Dim lngTemp As Long
        Dim strTemp As String
            
        '初始化默认值,应该有一个统一的地方设置默认值，包括配置显示和最终读取
        chkIgnorePosi.value = 0     '忽略结果阴阳性
        chkReportAfterResult.value = 0 '无影像诊断为阴性
        ChkFinishCommit.value = 0   '无报告完成后直接完成
        chkReportAfterImging.value = 0  '无图像不可编辑报告
        chkLocalizerBackward.value = 0  '定位片后置
        chkChangeUser.value = 0         '允许交换用户
        chkSwitchUser.value = 0         '允许切换用户
        chkTechReportSame.value = 0     '只能填写自己检查的报告
        chkWriteCapDoctor.value = 0     '采集图像者为检查技师
        ChkCompleteCommit.value = 0     '审核后直接完成
        chkFinallyCompleteCommit.value = 0  '终审后直接完成
        optMatch(0).value = True        '匹配数据库项目
        
        ChkLike.value = 0               '启用登记时姓名模糊查找
        TxtLike.Text = 0                '登记时姓名模糊查找天数
        Txt默认天数.Text = 2            '默认过滤天数
        txtViewHistoryImageDays.Text = 1 '默认自动打开历史图像天数
        chkRefreshInterval.value = 0    '启用病人列表自动刷新
        txtRefreshInterval.Text = 0     '默认病人列表自动刷新间隔为0秒，不刷新
        cboSaveDevice.Clear                 '存储设备
        chkPrintCommit.value = 0        '打印后直接完成
        chkCompletePrint.value = 0      '终审后直接打印
        chkUseReferencePatient.value = 0  '默认不启用关联病人
        optCapital(0).value = True      '默认拼音使用大写
        optCapital(1).value = True      '默认拼音间隔用空格
        chkCheckMaxNo.value = 1         '默认提取实际最大号码
        chkDefaultPosi.value = 0        '诊断结果默认阳性为未勾选
        chkConformDetermine.value = 1       '符合情况判定默认为选中
        txtImageLevel.Text = "甲,乙"     '默认影像质量等级
        txtReportLevel.Text = "甲,乙"    '默认报告质量等级
        chkPetitionCapture.value = 1     '默认勾选启用申请单扫描
        chkAddons.value = 1              '在登记窗口显示附加主述
        chkReagent.value = 1             '在登记窗口显示造影剂

        If cboViewReport.ListCount > 0 Then cboViewReport.ListIndex = 0
        
        On Error GoTo Err
        
        lngTemp = Val(GetDeptPara(mlng科室ID, "诊断结果提示类型", 0))
        optResultInput(lngTemp).value = True
        
        chkIgnorePosi.value = Val(GetDeptPara(mlng科室ID, "忽略结果阴阳性", 0)) '第一次使用时需要重新读取
        chkDefaultPosi.value = Val(GetDeptPara(mlng科室ID, "诊断结果默认阳性", 0))  '读取默认阳性参数
        chkReportAfterResult.value = Val(GetDeptPara(mlng科室ID, "无影像诊断为阴性", 0))
        
        chkConformDetermine.value = Val(GetDeptPara(mlng科室ID, "符合情况判定", 0))    '读取符合情况判定
        
        chkImageLevel.value = Val(GetDeptPara(mlng科室ID, "影像质量判定", 0))   '读取影像质量判定
        txtImageLevel.Text = NVL(GetDeptPara(mlng科室ID, "影像质量等级", "甲,乙"))  '读取影像质量等级
        txtImageLevel.Enabled = chkImageLevel.value = 1
        
        chkReportLevel.value = Val(GetDeptPara(mlng科室ID, "报告质量判定", 0)) '读取报告质量判定
        txtReportLevel.Text = NVL(GetDeptPara(mlng科室ID, "报告质量等级", "甲,乙"))  '读取报告质量等级
        txtReportLevel.Enabled = chkReportLevel.value = 1
        
        chkPetitionCapture.value = Val(GetDeptPara(mlng科室ID, "启用申请单扫描", 1))    '读取启用申请单扫描参数
    
        ChkFinishCommit.value = Val(GetDeptPara(mlng科室ID, "无报告完成后直接完成", 0))
        chkCanViewImage.value = Val(GetDeptPara(mlng科室ID, "采图后医生站即可观片", 0))
        chkReportAfterImging.value = Val(GetDeptPara(mlng科室ID, "有图像才能写报告", 0))
        chkCanOverWrite.value = Val(GetDeptPara(mlng科室ID, "允许检查号重复", 0))
        chkCheckMaxNo.value = Val(GetDeptPara(mlng科室ID, "提取实际最大号码", 1))
        chkChangeNO.value = Val(GetDeptPara(mlng科室ID, "手工调整检查号", 0))
        chkLocalizerBackward.value = Val(GetDeptPara(mlng科室ID, "定位片后置", 0))
        chkChangeUser.value = Val(GetDeptPara(mlng科室ID, "允许交换用户", 0))
        chkSwitchUser.value = Val(GetDeptPara(mlng科室ID, "允许切换用户", 0))
        chkTechReportSame.value = Val(GetDeptPara(mlng科室ID, "只能填写自己检查的报告", 0))
        chkWriteCapDoctor.value = Val(GetDeptPara(mlng科室ID, "采集图像者为检查技师", 0))
        ChkCompleteCommit.value = Val(GetDeptPara(mlng科室ID, "审核后直接完成", 0))
        chkFinallyCompleteCommit.value = Val(GetDeptPara(mlng科室ID, "终审后直接完成", 0))
        chkPrintCommit.value = Val(GetDeptPara(mlng科室ID, "打印后直接完成", 0))
        chkCompletePrint.value = Val(GetDeptPara(mlng科室ID, "终审后直接打印", 0))
        
        TxtLike.Text = Val(GetDeptPara(mlng科室ID, "登记时姓名模糊查找天数", 0))
        chkSample.value = Val(GetDeptPara(mlng科室ID, "登记后直接检查", 0))
        ChkLike.value = IIF(Val(TxtLike.Text) <> 0, 1, 0)
        chkAllPatientIsOutside.value = Val(GetDeptPara(mlng科室ID, "所有登记病人标记为外来", 0))
        
        Txt默认天数.Text = Val(GetDeptPara(mlng科室ID, "默认过滤天数", 2))
        
        If Val(Txt默认天数.Text) > 15 Or Val(Txt默认天数.Text) <= 0 Then
            Txt默认天数.Text = 2
        End If
        
        txtViewHistoryImageDays.Text = Val(GetDeptPara(mlng科室ID, "自动打开历史图像天数", 1))
        If Val(txtViewHistoryImageDays.Text) > 15 Or Val(txtViewHistoryImageDays.Text) <= 0 Then
            txtViewHistoryImageDays.Text = 1
        End If
        
        txtRefreshInterval.Text = Val(GetDeptPara(mlng科室ID, "自动刷新间隔", 0))
        chkRefreshInterval.value = IIF(Val(txtRefreshInterval.Text) <> 0, 1, 0)
        optMatch(Val(GetDeptPara(mlng科室ID, "匹配数据库项目", 0))).value = True
        
        chkAutoSendWorkList.value = Val(GetDeptPara(mlng科室ID, "报道时自动发送WorkList", "1"))
        chkAddons.value = Val(GetDeptPara(mlng科室ID, "显示附加主述", "1"))
        chkReagent.value = Val(GetDeptPara(mlng科室ID, "显示造影剂", "1"))
        chkSetFocusWithReport.value = Val(GetDeptPara(mlng科室ID, "检查切换时定位报告编辑", "1"))
        chkNameFuzzySearch.value = Val(GetDeptPara(mlng科室ID, "姓名默认模糊查询", "1"))
        chkNameQueryTimeLimit.value = Val(GetDeptPara(mlng科室ID, "姓名查询时间限制", "1"))
        
        chkPreView.value = IIF(Val(GetDeptPara(mlng科室ID, "缩略图预览方式", "0")) > 0, 1, 0)
        
        If chkPreView.value = 1 Then
            optMovePreview.Enabled = True
            lblDelayTime.Enabled = True
            txtDelayTime.Enabled = True
            optClickPreview.Enabled = True
        Else
            optMovePreview.Enabled = False
            lblDelayTime.Enabled = False
            txtDelayTime.Enabled = False
            optClickPreview.Enabled = False
        End If
        
        optMovePreview.value = Val(GetDeptPara(mlng科室ID, "缩略图预览方式", "0")) = 1
        optClickPreview.value = Val(GetDeptPara(mlng科室ID, "缩略图预览方式", "0")) = 2
        txtDelayTime.Text = Val(GetDeptPara(mlng科室ID, "移动预览延时", "2"))
        
        If Val(GetDeptPara(mlng科室ID, "医生站查看报告", "1")) = 0 Then
            cboViewReport.ListIndex = 0
        Else
            cboViewReport.ListIndex = 1
        End If
        
        OptCode(Val(GetDeptPara(mlng科室ID, "患者检查号保持不变", 0))).value = True
        OptUnicode(Val(GetDeptPara(mlng科室ID, "检查号保持不变类别", 0))).value = True
        optUsePatientID.value = Val(GetDeptPara(mlng科室ID, "使用患者号", 0))
        OptBuildcode(Val(GetDeptPara(mlng科室ID, "检查号生成方式", 0))).value = True
        optUseAdviceID.value = Val(GetDeptPara(mlng科室ID, "使用医嘱号", 0))
        
        '检查号编号设置
        strTemp = GetDeptPara(mlng科室ID, "检查号前缀", "")
        If strTemp = "" Then
            '不使用前缀
            chkPreText.value = 0
        Else
            '使用前缀
            chkPreText.value = 1
            If strTemp = "1" Then
                optPreText(0).value = 1
                txtPreText.Text = ""
            Else
                optPreText(1).value = 1
                txtPreText.Text = strTemp
            End If
        End If
        
        strTemp = GetDeptPara(mlng科室ID, "检查号分隔符1", "")
        strTemp = Left(strTemp, 1) '只取一个字符
        Call setCheckNoDelimeter(1, strTemp)
        
        strTemp = GetDeptPara(mlng科室ID, "检查号分隔符2", "")
        strTemp = Left(strTemp, 1) '只取一个字符
        Call setCheckNoDelimeter(2, strTemp)
        
        lngTemp = Val(GetDeptPara(mlng科室ID, "检查号年", 0))
        chkYear.value = IIF(lngTemp = 0, 0, 1)
        optYear(0).value = IIF((lngTemp = 1 Or lngTemp = 0), 1, 0)
        optYear(1).value = IIF(lngTemp = 2, 1, 0)
        
        chkMonth.value = IIF(Val(GetDeptPara(mlng科室ID, "检查号月", 0)) = 1, 1, 0)
        chkDay.value = IIF(Val(GetDeptPara(mlng科室ID, "检查号日", 0)) = 1, 1, 0)
        
        txtStartNum.Text = Val(GetDeptPara(mlng科室ID, "检查号起始数", 1))
        lngTemp = Val(GetDeptPara(mlng科室ID, "检查号固定位数", 0))
        chkFixedLen.value = IIF(lngTemp = 0, 0, 1)
        txtFixedLen.Text = IIF(lngTemp = 0, "", lngTemp)
        
        '设置检查号配置的参数相关可用性
        Call ConfigAppNoState
        
        chkUseReferencePatient.value = Val(GetDeptPara(mlng科室ID, "启动关联病人", 0))
        chkPrintNeedComplete.value = Val(GetDeptPara(mlng科室ID, "平诊需审核才能打报告", 0))
        
        '拼音名设置
        optCapital(Val(GetDeptPara(mlng科室ID, "拼音名大小写", 0))).value = True
        optSplitter(Val(GetDeptPara(mlng科室ID, "拼音名分隔符", 0))).value = True
        
        Call LoadScanDevice
        
        Exit Sub
Err:
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End Sub

    
    Private Function GetDeptPara(ByVal lngDeptID As Long, ByVal varPara As String, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
    '功能：读取指定的参数值
    '参数：lngDept=科室ID
    '      varPara=参数名
    '      strDefault=当数据库中没有该参数时使用的缺省值(注意不是为空时)
    '      blnNotCache=是否不从缓存中读取
    '返回：参数值，字符串形式
        Dim rsTmp As ADODB.Recordset
        Dim strSQL As String, blnNew As Boolean
        
        On Error GoTo errH
        
        If blnNotCache Then
            Set rsTmp = New ADODB.Recordset
            strSQL = "Select 参数值 from 影像流程参数 where 科室ID = [1] and 参数名=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取参数", lngDeptID, varPara)
            
            If Not rsTmp.EOF Then
                GetDeptPara = NVL(rsTmp!参数值)
            Else
                GetDeptPara = strDefault
            End If
        Else
            '第一次加载参数缓存
            If mrsDeptParas Is Nothing Then
                blnNew = True
            ElseIf mrsDeptParas.State = 0 Then
                blnNew = True
            End If
            If blnNew Then
                strSQL = "Select 参数值,参数名,科室ID from 影像流程参数"
                Set mrsDeptParas = New ADODB.Recordset
                Set mrsDeptParas = zlDatabase.OpenSQLRecord(strSQL, "读取参数")
            End If
            
            '根据缓存读取参数值
            mrsDeptParas.Filter = "参数名='" & CStr(varPara) & "' AND 科室ID=" & lngDeptID
            If Not mrsDeptParas.EOF Then
                GetDeptPara = NVL(mrsDeptParas!参数值)
            Else
                GetDeptPara = strDefault
            End If
        End If
        Exit Function
errH:
        If ErrCenter() = 1 Then Resume
    End Function

    Private Sub LoadScanDevice()
        Dim strSQL As String
        Dim rsTemp As ADODB.Recordset
        
        strSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型=1 and NVL(状态,0)=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsTemp.EOF Then
            MsgBox "未定义申请单存储设备，请到影像设备目录中设置！", vbInformation, gstrSysName
            Exit Sub
        Else
            cboSaveDevice.AddItem ""
            
            Do While Not rsTemp.EOF
                cboSaveDevice.AddItem rsTemp!设备号 & "-" & NVL(rsTemp!设备名)
                
                If GetDeptPara(mlng科室ID, "申请单存储设备号", "") = rsTemp!设备号 Then
                    cboSaveDevice.ListIndex = cboSaveDevice.NewIndex
                End If
                
                rsTemp.MoveNext
            Loop
        End If
    End Sub

    Private Function GetUserInfo() As Boolean
    '功能：获取登陆用户信息
        Dim rsTmp As New ADODB.Recordset
        Dim strSQL As String
        
        Set rsTmp = zlDatabase.GetUserInfo
        
        UserInfo.用户名 = gstrDbUser
        UserInfo.姓名 = gstrDbUser
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.编号 = rsTmp!编号
            UserInfo.部门ID = IIF(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
            UserInfo.简码 = IIF(IsNull(rsTmp!简码), "", rsTmp!简码)
            UserInfo.姓名 = IIF(IsNull(rsTmp!姓名), "", rsTmp!姓名)
            UserInfo.用户名 = IIF(IsNull(rsTmp!用户名), "", rsTmp!用户名)
            GetUserInfo = True
        End If
        Exit Function
errH:
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End Function
    

    Private Function GetUser科室IDs(Optional ByVal bln病区 As Boolean) As String
    '功能：获取操作员所属的科室(本身所在科室+所属病区包含的科室),可能有多个
    '参数：是否取所属病区下的科室
        Dim rsTmp As New ADODB.Recordset
        Dim strSQL As String, i As Long
        
        strSQL = "Select 部门ID From 部门人员 Where 人员ID=[1]"
        If bln病区 Then
            strSQL = strSQL & " Union" & _
                " Select Distinct B.科室ID From 部门人员 A,床位状况记录 B" & _
                " Where A.部门ID=B.病区ID And A.人员ID=[1]"
        End If
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID)
        For i = 1 To rsTmp.RecordCount
            GetUser科室IDs = GetUser科室IDs & "," & rsTmp!部门ID
            rsTmp.MoveNext
        Next
        GetUser科室IDs = Mid(GetUser科室IDs, 2)
        Exit Function
errH:
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End Function


    Private Function subInitDepartInfo()
    '载入检查科室
        Dim rsTmp As New ADODB.Recordset
        Dim strSQL As String, i As Long
        Dim str科室IDs As String, str来源 As String
        Dim strDepartment() As String
        Dim intCurDept As Integer
        
        On Error GoTo errH
        
        If InStr(mstrPrivs, "所有科室") > 0 Then
            strSQL = _
                " Select Distinct A.ID,A.编码,A.名称" & _
                " From 部门表 A,部门性质说明 B " & _
                " Where B.部门ID = A.ID " & _
                " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
                " And B.工作性质 IN('检查')  Order by A.编码"
        Else
            strSQL = _
                " Select Distinct A.ID,A.编码,A.名称" & _
                " From 部门表 A,部门性质说明 B,部门人员 C " & _
                " Where B.部门ID = A.ID And A.ID=C.部门ID And C.人员ID=" & UserInfo.ID & _
                " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
                " And B.工作性质 IN('检查')  Order by A.编码"
        End If
         
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        
        If rsTmp.EOF Then
            MsgBox "没有发现医技科室信息,请先到部门管理中设置。", vbInformation, gstrSysName
            Exit Function
        Else
            str科室IDs = GetUser科室IDs
            Do Until rsTmp.EOF
                mstrCanUse科室 = mstrCanUse科室 & "|" & rsTmp!ID & "_" & rsTmp!编码 & "-" & rsTmp!名称
                If rsTmp!ID = UserInfo.部门ID Then mlngCur科室ID = rsTmp!ID: mstrCur科室 = rsTmp!编码 & "-" & rsTmp!名称 '提取默认科室
                If InStr("," & str科室IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur科室ID = 0 Then mlngCur科室ID = rsTmp!ID: mstrCur科室 = rsTmp!编码 & "-" & rsTmp!名称 '没有默认科室,取所属检查科室第一个
                rsTmp.MoveNext
            Loop
            
            str科室IDs = GetUser科室IDs
            Do Until rsTmp.EOF
                mstrCanUse科室 = mstrCanUse科室 & "|" & rsTmp!ID & "_" & rsTmp!编码 & "-" & rsTmp!名称
                If rsTmp!ID = UserInfo.部门ID Then mlngCur科室ID = rsTmp!ID: mstrCur科室 = rsTmp!编码 & "-" & rsTmp!名称 '提取默认科室
                If InStr("," & str科室IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur科室ID = 0 Then mlngCur科室ID = rsTmp!ID: mstrCur科室 = rsTmp!编码 & "-" & rsTmp!名称 '没有默认科室,取所属检查科室第一个
                rsTmp.MoveNext
            Loop
            mstrCanUse科室 = Mid(mstrCanUse科室, 2)
            If InStr(mstrPrivs, "所有科室") > 0 And mlngCur科室ID = 0 Then
                mlngCur科室ID = Split(Split(mstrCanUse科室, "|")(0), "_")(0)
                mstrCur科室 = Split(Split(mstrCanUse科室, "|")(0), "_")(1)
            End If
            
            If mlngCur科室ID = 0 And InStr(mstrPrivs, "所有科室") <= 0 Then '没有所有科室操作权限,而且操作者科室不属于检查类科室
                MsgBox "没有发现你所属科室,不能使用医技工作站。", vbInformation, gstrSysName
                Exit Function
            End If
            
            '填充cmbDept
            cmbDept.Clear
            intCurDept = -1
            strDepartment = Split(mstrCanUse科室, "|")
            For i = 0 To UBound(strDepartment)
                cmbDept.AddItem Split(strDepartment(i), "_")(1)
                cmbDept.ItemData(cmbDept.ListCount - 1) = Split(strDepartment(i), "_")(0)
                If Split(strDepartment(i), "_")(0) = mlngCur科室ID Then
                    intCurDept = i
                End If
            Next i
            If intCurDept <> -1 Then
                cmbDept.ListIndex = intCurDept
            Else
                cmbDept.ListIndex = 0
            End If
            mlng科室ID = cmbDept.ItemData(cmbDept.ListIndex)
            
            subInitDepartInfo = True
        End If
        
        
        Exit Function
errH:
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End Function

Private Function GetEncryptionPassW(ByVal strPswd As String) As String
'获取加密密码
    Dim strEncryptionPassW As String
    
    If strPswd = "" Then Exit Function
    
    strEncryptionPassW = EncryptionPassW(Trim(strPswd))
    strEncryptionPassW = Mid(strEncryptionPassW, 1, 1) & "※" & Mid(strEncryptionPassW, 2)
    strEncryptionPassW = "★" & strEncryptionPassW & "★"
    strEncryptionPassW = Replace(strEncryptionPassW, "'", "''")
    
    GetEncryptionPassW = strEncryptionPassW
End Function

Private Function GetDecryptionPassW(ByVal strPswd As String) As String
'获取解密密码
    Dim strDecryptionPassW As String
    
    GetDecryptionPassW = strPswd
    
    If Len(strPswd) >= 3 Then
        If Mid(strPswd, 1, 1) & Mid(strPswd, 3, 1) & Mid(strPswd, Len(strPswd), 1) = "★※★" Then
            strDecryptionPassW = Mid(strPswd, 2)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, Len(strDecryptionPassW) - 1)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, 1) & Mid(strDecryptionPassW, 3)
            strDecryptionPassW = DecryptionPassW(strDecryptionPassW)
            
            GetDecryptionPassW = strDecryptionPassW
        End If
    End If
End Function

Private Function GetRandom(ByVal lngBase As Long) As String
    Dim lngNum As Long
    
    Randomize 99
    
    lngNum = Fix(Rnd * lngBase)
    
    If lngNum <= 0 Then lngNum = 1
    
    GetRandom = Chr(lngNum)
End Function

'获取加密密码
Private Function EncryptionPassW(ByVal strPassW As String) As String
    Dim i As Integer
    Dim lngAsc  As Long
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim strRandom As String
    Dim strBase As String
        
    i = 0
    
    lngPassWLength = Len(strPassW)
    
    strBase = GetRandom(30)
    strRandom = GetRandom(30)
    
    ReDim intASC(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
     
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassW, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strBase) Xor Asc(strRandom)
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop
    
    EncryptionPassW = strBase & Join(strTemp, "") & strRandom '加密后的字串
End Function

'获取解密密码
Private Function DecryptionPassW(ByVal strPassW As String) As String
    Dim i As Integer
    Dim lngAsc  As Integer
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim lngBase As Long
    Dim strRandom As String
    Dim strPassSouce As String

    i = 0
    
    strPassSouce = Mid(strPassW, 2, Len(strPassW) - 2)
    lngPassWLength = Len(strPassSouce)
    lngBase = Asc(Mid(strPassW, 1, 1))
    
    strRandom = Right(strPassW, 1)
    
    ReDim intASC(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
    
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassSouce, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strRandom) Xor lngBase
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop

    DecryptionPassW = Join(strTemp, "") '解密后的字串
End Function

Private Sub optImageDblClick_Click(Index As Integer)
    If Index = 1 Then
        If chkPreText.value = 1 And optClickPreview.value Then
            MsgBox "您已经启用‘鼠标单击时预览图像’，与缩略图双击后打开图像编辑的功能重合，建议缩略图双击后的功能选择直接写入报告", vbOKOnly, "提示信息"
        End If
    End If
End Sub

'Private Sub optBigImgAction_Click(Index As Integer)
'    If Index = 3 Then optImageDblClick(0).value = True
'End Sub

Private Sub optClickPreview_Click()
    If optClickPreview.value Then optImageDblClick(0).value = True
    
    If optMovePreview.value = False Then
        txtDelayTime.Enabled = False
        lblDelayTime.Enabled = False
    End If
End Sub

Private Sub LoadPathol()
    With lst(lst_PatholInfo)
        .AddItem "标本名称", 0
        .AddItem "取材位置", 1
        .AddItem "形状", 2
        .AddItem "蜡块数", 3
        .AddItem "制片数", 4
        .AddItem "主取医师", 5
        .AddItem "取材时间", 6
        .AddItem "性质", 7
        .AddItem "颜色", 8
        .AddItem "标本量", 9
    End With
End Sub

'初始化检查号分隔符可以输入的字符，支持键盘中的所有单字节符号，除了双引号和单引号之外。
Private Sub LoadCheckNoDelimeter()
    Dim i As Integer
    
    For i = 1 To 2
        cboDelimeter(i).Clear
        cboDelimeter(i).AddItem "~"
        cboDelimeter(i).AddItem "`"
        cboDelimeter(i).AddItem "!"
        cboDelimeter(i).AddItem "@"
        cboDelimeter(i).AddItem "#"
        cboDelimeter(i).AddItem "$"
        cboDelimeter(i).AddItem "%"
        cboDelimeter(i).AddItem "^"
        cboDelimeter(i).AddItem "&"
        cboDelimeter(i).AddItem "*"
        cboDelimeter(i).AddItem "("
        cboDelimeter(i).AddItem ")"
        cboDelimeter(i).AddItem "_"
        cboDelimeter(i).AddItem "-"
        cboDelimeter(i).AddItem "+"
        cboDelimeter(i).AddItem "="
        cboDelimeter(i).AddItem "{"
        cboDelimeter(i).AddItem "}"
        cboDelimeter(i).AddItem "["
        cboDelimeter(i).AddItem "]"
        cboDelimeter(i).AddItem "\"
        cboDelimeter(i).AddItem "|"
        cboDelimeter(i).AddItem ";"
        cboDelimeter(i).AddItem ":"
        cboDelimeter(i).AddItem ","
        cboDelimeter(i).AddItem "<"
        cboDelimeter(i).AddItem ">"
        cboDelimeter(i).AddItem "."
        cboDelimeter(i).AddItem "/"
        cboDelimeter(i).AddItem "?"
    Next i
    
End Sub

'设置检查号分隔符的内容
Private Sub setCheckNoDelimeter(lngIndex As Long, strText As String)
    On Error GoTo Err
    
    cboDelimeter(lngIndex).Text = strText
    chkDelimiter(lngIndex).value = 1
    
    Exit Sub
Err:
    '如果赋值失败，则取消分隔符的选择
    cboDelimeter(lngIndex).ListIndex = -1
    chkDelimiter(lngIndex).value = 0
End Sub
