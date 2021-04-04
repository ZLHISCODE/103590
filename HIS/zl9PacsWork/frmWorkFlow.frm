VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmWorkFlow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "工作流设置"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10785
   Icon            =   "frmWorkFlow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame framWorkFlow 
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   10575
      Begin VB.Frame frmResultInput 
         Height          =   435
         Left            =   1185
         TabIndex        =   45
         Top             =   6060
         Width           =   5490
         Begin VB.OptionButton optResultInput 
            Caption         =   "报告打印前"
            Height          =   240
            Index           =   2
            Left            =   4050
            TabIndex        =   56
            Top             =   150
            Width           =   1290
         End
         Begin VB.OptionButton optResultInput 
            Caption         =   "审核签名后"
            Height          =   240
            Index           =   1
            Left            =   2625
            TabIndex        =   55
            Top             =   150
            Width           =   1290
         End
         Begin VB.OptionButton optResultInput 
            Caption         =   "诊断签名后"
            Height          =   240
            Index           =   0
            Left            =   1290
            TabIndex        =   54
            Top             =   150
            Value           =   -1  'True
            Width           =   1290
         End
         Begin VB.Label lblImageQuality 
            Caption         =   "录入时机："
            Height          =   180
            Left            =   210
            TabIndex        =   46
            Top             =   165
            Width           =   1035
         End
      End
      Begin VB.Frame Frame13 
         Height          =   1650
         Left            =   120
         TabIndex        =   47
         Top             =   4680
         Width           =   7695
         Begin VB.TextBox txtImageLevel 
            Height          =   270
            Left            =   4290
            TabIndex        =   68
            Text            =   "甲,乙"
            ToolTipText     =   "用于评定影像质量的登记，最多四个等级"
            Top             =   960
            Width           =   1035
         End
         Begin VB.TextBox txtReportLevel 
            Height          =   270
            Left            =   4290
            TabIndex        =   67
            Text            =   "甲,乙"
            Top             =   360
            Width           =   1035
         End
         Begin VB.CheckBox chkImageLevel 
            Caption         =   "影像质量等级"
            Height          =   180
            Left            =   2880
            TabIndex        =   66
            Top             =   990
            Width           =   1410
         End
         Begin VB.CheckBox chkReportLevel 
            Caption         =   "报告质量等级"
            Height          =   180
            Left            =   2880
            TabIndex        =   65
            Top             =   375
            Width           =   1410
         End
         Begin VB.CheckBox chkConformDetermine 
            Caption         =   "符合情况判断"
            Height          =   180
            Left            =   5700
            TabIndex        =   53
            ToolTipText     =   "激活符合情况功能和菜单"
            Top             =   990
            Width           =   1455
         End
         Begin VB.CheckBox chkCriticalValues 
            Caption         =   "危急情况判断"
            Height          =   180
            Left            =   5700
            TabIndex        =   52
            ToolTipText     =   "激活危急情况功能和菜单"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Frame Frame5 
            Height          =   1005
            Left            =   120
            TabIndex        =   48
            Top             =   270
            Width           =   2655
            Begin VB.CheckBox chkDefaultPosi 
               Caption         =   "诊断结果默认阳性"
               Height          =   300
               Left            =   240
               TabIndex        =   51
               ToolTipText     =   "弹出阴阳性选择窗口，默认选择阳性。"
               Top             =   300
               Width           =   1815
            End
            Begin VB.CheckBox chkReportAfterResult 
               Caption         =   "无诊断内容为阴性"
               Height          =   180
               Left            =   240
               TabIndex        =   50
               ToolTipText     =   "书写报告时，没有录入诊断，则默认记录为阴性。"
               Top             =   720
               Width           =   1740
            End
            Begin VB.CheckBox chkIgnorePosi 
               Caption         =   "忽略结果的阴阳性"
               Height          =   180
               Left            =   240
               TabIndex        =   49
               ToolTipText     =   "不记录和处理阴阳性。"
               Top             =   0
               Width           =   1920
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "拼音名"
         Height          =   2070
         Left            =   8040
         TabIndex        =   36
         Top             =   4280
         Width           =   2415
         Begin VB.OptionButton optCapital 
            Caption         =   "大写"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   42
            ToolTipText     =   "选择后拼音名显示全为大写字母。"
            Top             =   260
            Width           =   735
         End
         Begin VB.OptionButton optCapital 
            Caption         =   "小写"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   41
            ToolTipText     =   "选择后拼音名显示全为小写字母。"
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton optCapital 
            Caption         =   "首字母大写"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   40
            ToolTipText     =   "选择后拼音名首字母大写。"
            Top             =   960
            Width           =   1215
         End
         Begin VB.Frame Frame9 
            Caption         =   "间隔"
            Height          =   540
            Left            =   120
            TabIndex        =   37
            Top             =   1365
            Width           =   2175
            Begin VB.OptionButton optSplitter 
               Caption         =   "无"
               Height          =   255
               Index           =   1
               Left            =   1200
               TabIndex        =   39
               ToolTipText     =   "拼音名之间无间隔。"
               Top             =   200
               Width           =   495
            End
            Begin VB.OptionButton optSplitter 
               Caption         =   "空格"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   38
               ToolTipText     =   "拼音名之间使用空格为间隔符。"
               Top             =   200
               Width           =   735
            End
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "检查号设置"
         Height          =   2100
         Left            =   120
         TabIndex        =   29
         Top             =   2600
         Width           =   7695
         Begin VB.CheckBox chkAutoInc 
            Caption         =   "自动递增检查号"
            Height          =   180
            Left            =   5160
            TabIndex        =   57
            Top             =   375
            Width           =   1635
         End
         Begin VB.Frame Frame11 
            Height          =   1695
            Left            =   5040
            TabIndex        =   72
            Top             =   240
            Width           =   2415
            Begin VB.OptionButton OptBuildcode 
               Caption         =   "相同检查类别自动递增"
               Height          =   210
               Index           =   0
               Left            =   120
               TabIndex        =   74
               ToolTipText     =   "检查号以检查类别为基础，自动递增。"
               Top             =   600
               Value           =   -1  'True
               Width           =   2130
            End
            Begin VB.OptionButton OptBuildcode 
               Caption         =   "本科室内自动递增"
               Height          =   210
               Index           =   1
               Left            =   120
               TabIndex        =   73
               ToolTipText     =   "检查号以科室为基础，自动递增。"
               Top             =   1170
               Width           =   1740
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "检查号一致性"
            Height          =   1770
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   2295
            Begin VB.Frame Frame10 
               Height          =   735
               Left            =   375
               TabIndex        =   61
               Top             =   930
               Width           =   1695
               Begin VB.OptionButton OptUnicode 
                  Caption         =   "本科室统一"
                  Height          =   210
                  Index           =   1
                  Left            =   75
                  TabIndex        =   63
                  ToolTipText     =   "科室相同，保持检查号不变。"
                  Top             =   390
                  Width           =   1290
               End
               Begin VB.OptionButton OptUnicode 
                  Caption         =   "本检查类别统一"
                  Height          =   210
                  Index           =   0
                  Left            =   75
                  TabIndex        =   62
                  ToolTipText     =   "检查类别相同，保持检查号不变。"
                  Top             =   165
                  Width           =   1590
               End
            End
            Begin VB.OptionButton OptCode 
               Caption         =   "患者检查号保持不变"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   35
               ToolTipText     =   "同一个患者，报到时保持检查号不变。"
               Top             =   660
               Width           =   1935
            End
            Begin VB.OptionButton OptCode 
               Caption         =   "每次检查用新检查号"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   34
               ToolTipText     =   "报到时产生新的检查号。"
               Top             =   345
               Value           =   -1  'True
               Width           =   1920
            End
         End
         Begin VB.CheckBox chkCanOverWrite 
            Caption         =   "允许检查号重复"
            Height          =   300
            Left            =   2760
            TabIndex        =   32
            ToolTipText     =   "允许登记病人的检查号出现重复。"
            Top             =   960
            Width           =   1935
         End
         Begin VB.CheckBox chkChangeNO 
            Caption         =   "允许手工调整检查号"
            Height          =   180
            Left            =   2760
            TabIndex        =   31
            ToolTipText     =   "允许根据实际需要手动修改检查号。"
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox chkCheckMaxNo 
            Caption         =   "提取实际最大号码"
            Height          =   300
            Left            =   2760
            TabIndex        =   30
            ToolTipText     =   "以实际最大号码为基础顺序编号；不勾选，则以当前设置的最大号码顺序编号。"
            Top             =   1680
            Width           =   1935
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "先检查后报到，图像匹配"
         Height          =   1605
         Left            =   8025
         TabIndex        =   25
         Top             =   2600
         Width           =   2415
         Begin VB.OptionButton optMatch 
            Caption         =   "门诊/住院号"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   28
            ToolTipText     =   "报到时通过门诊/住院号和图像信息进行匹配，仅用于影像医技站。"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.OptionButton optMatch 
            Caption         =   "检查号"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   27
            ToolTipText     =   "报到时通过检查号和图像信息进行匹配，仅用于影像医技站。"
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optMatch 
            Caption         =   "医嘱ID"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   26
            ToolTipText     =   "报到时通过医嘱ID和图像信息进行匹配，仅用于影像医技站。"
            Top             =   780
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "功能设置"
         Height          =   2500
         Left            =   8040
         TabIndex        =   22
         Top             =   0
         Width           =   2415
         Begin VB.CheckBox chkSwitchUser 
            Caption         =   "启用切换用户"
            Height          =   180
            Left            =   240
            TabIndex        =   64
            ToolTipText     =   "激活切换用户功能，可以进行用户切换，仅限于影像病理站。"
            Top             =   697
            Width           =   1455
         End
         Begin VB.Frame Frame2 
            Height          =   780
            Left            =   120
            TabIndex        =   58
            ToolTipText     =   "选择采集图像和扫描申请单所使用的存储设备。"
            Top             =   1560
            Width           =   2175
            Begin VB.ComboBox cboSaveDevice 
               Height          =   300
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   315
               Width           =   1965
            End
            Begin VB.CheckBox chkPetitionCapture 
               Caption         =   "启用申请单扫描"
               Height          =   180
               Left            =   135
               TabIndex        =   59
               ToolTipText     =   "报告审核后，该检查自动完成。"
               Top             =   0
               Value           =   1  'Checked
               Width           =   1575
            End
         End
         Begin VB.CheckBox chkUseReferencePatient 
            Caption         =   "启用关联病人"
            Height          =   180
            Left            =   240
            TabIndex        =   24
            ToolTipText     =   "支持多个检查关联到同一个病人信息。"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CheckBox chkChangeUser 
            Caption         =   "启用交换用户"
            Height          =   180
            Left            =   240
            TabIndex        =   23
            ToolTipText     =   "激活交换用户功能，可以交换检查医生和报告医生，仅限于影像采集站。"
            Top             =   315
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "工作流设置"
         Height          =   2505
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   7695
         Begin VB.CheckBox chkReagent 
            Caption         =   "显示造影剂"
            Height          =   180
            Left            =   2280
            TabIndex        =   78
            ToolTipText     =   "在登记报到窗口显示造影剂一项，病理工作站不显示"
            Top             =   2200
            Width           =   1935
         End
         Begin VB.CheckBox chkAddons 
            Caption         =   "显示附加主述"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            ToolTipText     =   "在登记报到窗口显示附加主述一项"
            Top             =   2160
            Width           =   1935
         End
         Begin VB.CheckBox chkAutoSendWorkList 
            Caption         =   "报到时自动发送WorkList"
            Height          =   180
            Left            =   2280
            TabIndex        =   76
            Top             =   1920
            Value           =   1  'Checked
            Width           =   2532
         End
         Begin VB.CheckBox chkCompletePrint 
            Caption         =   "终审后直接打印"
            Height          =   180
            Left            =   240
            TabIndex        =   75
            ToolTipText     =   "终审签名后直接打印报告，仅适用于新版报告文档编辑器。"
            Top             =   1920
            Width           =   1680
         End
         Begin VB.TextBox txtViewHistoryImageDays 
            Height          =   270
            Left            =   6960
            MaxLength       =   2
            TabIndex        =   70
            Text            =   "1"
            Top             =   1878
            Width           =   465
         End
         Begin VB.CheckBox chkCanViewImage 
            Caption         =   "采图后医生站即可观片"
            Height          =   180
            Left            =   2280
            TabIndex        =   69
            ToolTipText     =   "采集图像后，在没有检查完成的情况下，医生站也可进行观片。"
            Top             =   1611
            Width           =   2160
         End
         Begin VB.TextBox txtRefreshInterval 
            Height          =   270
            Left            =   4080
            TabIndex        =   44
            Text            =   "0"
            Top             =   960
            Width           =   390
         End
         Begin VB.TextBox TxtLike 
            Enabled         =   0   'False
            Height          =   270
            Left            =   4200
            TabIndex        =   43
            ToolTipText     =   "0天则无时间限制,模糊查找所有病人"
            Top             =   600
            Width           =   270
         End
         Begin VB.CheckBox ChkFinishCommit 
            Caption         =   "无报告完成后直接完成"
            Height          =   180
            Left            =   2280
            TabIndex        =   20
            ToolTipText     =   "点击无报告完成后，该检查自动完成。"
            Top             =   1302
            Width           =   2160
         End
         Begin VB.CheckBox chkPrintCommit 
            Caption         =   "打印后直接完成"
            Height          =   180
            Left            =   240
            TabIndex        =   19
            ToolTipText     =   "打印报告后，该检查自动完成。"
            Top             =   684
            Width           =   1815
         End
         Begin VB.CheckBox ChkCompleteCommit 
            Caption         =   "审核后直接完成"
            Height          =   180
            Left            =   240
            TabIndex        =   18
            ToolTipText     =   "报告审核后，该检查自动完成。"
            Top             =   993
            Width           =   1935
         End
         Begin VB.CheckBox chkSample 
            Caption         =   "申请登记后直接报到"
            Height          =   180
            Left            =   240
            TabIndex        =   17
            ToolTipText     =   "登记与报到同时进行。"
            Top             =   1611
            Width           =   1935
         End
         Begin VB.TextBox Txt默认天数 
            Height          =   270
            Left            =   6720
            MaxLength       =   2
            TabIndex        =   16
            Text            =   "2"
            Top             =   1491
            Width           =   705
         End
         Begin VB.CheckBox chkReportAfterImging 
            Caption         =   "有图像才能写报告"
            Height          =   180
            Left            =   240
            TabIndex        =   15
            ToolTipText     =   "必须采集图像后才能编写影像报告。"
            Top             =   375
            Width           =   2040
         End
         Begin VB.CheckBox chkPrintNeedComplete 
            Caption         =   "平诊检查需审核才能打报告"
            Height          =   180
            Left            =   5040
            TabIndex        =   14
            ToolTipText     =   "平诊检查必须经过审核后才能打印报告。"
            Top             =   762
            Width           =   2505
         End
         Begin VB.CheckBox chkTechReportSame 
            Caption         =   "只能填写自己检查的报告"
            Height          =   180
            Left            =   5040
            TabIndex        =   13
            ToolTipText     =   "只有自己采集图像的检查，才能书写报告。"
            Top             =   375
            Width           =   2295
         End
         Begin VB.CheckBox chkWriteCapDoctor 
            Caption         =   "采集图像者为检查技师"
            Height          =   180
            Left            =   5040
            TabIndex        =   12
            ToolTipText     =   "采集图像之后，自动将当前用户记录成检查技师。"
            Top             =   1149
            Width           =   2400
         End
         Begin VB.CheckBox chkLocalizerBackward 
            Caption         =   "定位片后置"
            Height          =   180
            Left            =   240
            TabIndex        =   11
            ToolTipText     =   "将定位片放到最后一个序列显示。"
            Top             =   1302
            Width           =   1320
         End
         Begin VB.CheckBox chkRefreshInterval 
            Caption         =   "病人自动刷新间隔      秒"
            Height          =   180
            Left            =   2280
            TabIndex        =   10
            ToolTipText     =   "病人检查列表会间隔N秒自动刷新。"
            Top             =   993
            Width           =   2500
         End
         Begin VB.CheckBox ChkLike 
            Caption         =   "登记时姓名模糊查找    天"
            Height          =   195
            Left            =   2280
            TabIndex        =   9
            ToolTipText     =   "登记时支持对姓名进行模糊查找，可以查找到N天内的信息。"
            Top             =   684
            Width           =   2500
         End
         Begin VB.CheckBox chkAllPatientIsOutside 
            Caption         =   "所有登记病人标记为外来"
            Height          =   180
            Left            =   2280
            TabIndex        =   8
            ToolTipText     =   "凡在该工作站中登记的病人均标记为外来病人。"
            Top             =   375
            Width           =   2295
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "自动打开历史图像天数"
            Height          =   180
            Left            =   5040
            TabIndex        =   71
            ToolTipText     =   "如果当前检查没有图像，则自动打开指定时间段内的历史图像"
            Top             =   1923
            Width           =   1800
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "默认记录查询天数"
            Height          =   180
            Left            =   5040
            TabIndex        =   21
            ToolTipText     =   "检查列表中默认显示对应天数内的检查记录。"
            Top             =   1536
            Width           =   1440
         End
      End
   End
   Begin VB.ComboBox cmbDept 
      Height          =   300
      Left            =   1110
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   75
      Width           =   2055
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9465
      TabIndex        =   3
      Top             =   7640
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8265
      TabIndex        =   2
      Top             =   7640
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7650
      Width           =   1100
   End
   Begin XtremeSuiteControls.TabControl TabWindow 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10695
      _Version        =   589884
      _ExtentX        =   18865
      _ExtentY        =   12303
      _StockProps     =   64
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "影像科室"
      Height          =   180
      Left            =   165
      TabIndex        =   5
      Top             =   135
      Width           =   735
   End
End
Attribute VB_Name = "frmWorkFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String         '本模块的权限
Public mlng科室ID As Long 'IN:当前执行科室ID
Private mlngCur科室ID As Long       '当前科室ID
Private mstrCur科室 As String      '当前科室 编码-名称
Private mstrCanUse科室 As String    '当前可用科室  ID_编码-名称
Private mobjfrmTabPass As New FrmReqInput     '光标经过控制
Private mobjfrmEnableCtr As New FrmReqInput  '必须输入项控制
Private mobjFrmReportSetup As New frmReportSetup '报告设置
Private mobjFrmStudyListCfg As New frmStudyListCfg '检查列表配置
Private mobjfrmTechnicGroupCfg As New frmTechnicQueueCfg '医技执行间分组配置


Private Sub chkAutoInc_Click()
On Error Resume Next
    If chkAutoInc.value = 0 Then
        OptBuildcode(0).Enabled = False
        OptBuildcode(1).Enabled = False
        
        chkChangeNO.value = 1
        chkChangeNO.Enabled = False
        
        chkCheckMaxNo.value = 0
        chkCheckMaxNo.Enabled = False
    Else
        OptBuildcode(0).Enabled = True
        OptBuildcode(1).Enabled = True
        
        chkChangeNO.Enabled = True
        chkCheckMaxNo.Enabled = True
    End If
err.Clear
End Sub


Private Sub chkImageLevel_Click()
    txtImageLevel.Enabled = chkImageLevel.value = 1
End Sub

Private Sub ChkLike_Click()
    TxtLike.Enabled = IIf(ChkLike.value, True, False)
End Sub

Private Sub chkPetitionCapture_Click()
    cboSaveDevice.Enabled = IIf(chkPetitionCapture.value, True, False)
End Sub

Private Sub chkRefreshInterval_Click()
    txtRefreshInterval.Enabled = IIf(chkRefreshInterval.value, True, False)
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

Private Sub cmbDept_Click()
    mlng科室ID = cmbDept.ItemData(cmbDept.ListIndex)
    If TabWindow.ItemCount = IIf(InStr(GetPrivFunc(glngSys, 1160), "基本") > 0, 7, 6) Then '判断tab数量，目的是为了确保在装载完tab之后才触发其中的语句
        '刷新工作流程参数界面
        Call frmWorkFlowRefresh
        '刷新执行间界面
        Call frmTechRoomRefresh
        '刷新输入设置界面
        Call frmReqInputRefresh(0)
        '必须项控制
        Call frmReqInputRefresh(1)
        '刷新报告设置
        Call frmReportRefresh
        '刷新颜色设置
        Call frmStudyListCfgRefresh
        '刷新排队叫号设置
        RefreshTechnicRoomGroupCfg
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub CmdOK_Click()

    Dim intTxtLen As Integer
    
    If txtImageLevel.Enabled Then
        '将中文状态下的 逗号替换成英文状态
        txtImageLevel.Text = Replace(txtImageLevel.Text, "，", ",")
        
        intTxtLen = Len(txtImageLevel.Text) - Len(Replace(txtImageLevel.Text, ",", ""))
        
        If intTxtLen > 3 Or intTxtLen < 1 Then
            MsgBoxD Me, "影像等级最少为2种，最多为4种，请重新填写。", vbOKOnly, "提示信息"
            txtImageLevel.Text = Nvl(GetDeptPara(mlng科室ID, "影像质量等级", "甲,乙"))
            txtImageLevel.SetFocus
            Exit Sub
        End If
    End If
    
    
    If txtReportLevel.Enabled Then
        '将中文状态下的 逗号替换成英文状态
        txtReportLevel.Text = Replace(txtReportLevel.Text, "，", ",")
        
        intTxtLen = Len(txtReportLevel.Text) - Len(Replace(txtReportLevel.Text, ",", ""))
        
        If intTxtLen > 3 Or intTxtLen < 1 Then
            MsgBoxD Me, "报告等级最少为2种，最多为4种，请重新填写。", vbOKOnly, "提示信息"
            txtReportLevel.Text = Nvl(GetDeptPara(mlng科室ID, "报告质量等级", "甲,乙"))
            txtReportLevel.SetFocus
            Exit Sub
        End If
    End If
    

    Call SaveWorkFlow
    Call mobjfrmTabPass.zlSave
    Call mobjfrmEnableCtr.zlSave
    Call mobjFrmReportSetup.zlSave
    Call mobjFrmStudyListCfg.zlSave
    Call mobjfrmTechnicGroupCfg.zlSave
    
    Unload Me
End Sub

Private Sub Form_Load()
    '初始化模块级变量
    mstrPrivs = gstrPrivs
    mlng科室ID = 0
    mlngCur科室ID = 0
    mstrCur科室 = ""
    mstrCanUse科室 = ""
    
    mobjfrmTabPass.mintType = 0
    mobjfrmEnableCtr.mintType = 1
    
    '没有对应的科室，则退出
    If InitDepts = False Then
        Unload Me
        Exit Sub
    End If
    
    '装载子窗口
    Call InitFaceScheme
    
    '初始化子窗口
    '刷新工作流程参数界面
    Call frmWorkFlowRefresh
    '刷新执行间界面
    Call frmTechRoomRefresh
    '刷新输入设置界面
    Call frmReqInputRefresh(0)
    '必须项控制
    Call frmReqInputRefresh(1)
    '刷新报告设置
    Call frmReportRefresh
    '刷新检查列表配置
    Call frmStudyListCfgRefresh
    '刷新排队叫号设置
    Call RefreshTechnicRoomGroupCfg
End Sub

Private Sub Form_Resize()
    TabWindow.Left = 1
    TabWindow.Top = 480
    TabWindow.Width = Me.ScaleWidth
    TabWindow.Height = Me.ScaleHeight - 480
End Sub

Private Sub InitFaceScheme()
    Dim Item As TabControlItem
    
    mobjfrmTabPass.mlngDeptId = mlng科室ID
    mobjfrmEnableCtr.mlngDeptId = mlng科室ID
    frmTechnicRoom.mlngdept = mlng科室ID
    
    With TabWindow
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameBorder
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .InsertItem 1, "工作流设置", framWorkFlow.hWnd, 0
        .InsertItem 2, "执行间设置", frmTechnicRoom.hWnd, 0
        
        '有1160的权限时才允许进行配置
        If InStr(GetPrivFunc(glngSys, 1160), "基本") > 0 Then
            .InsertItem 3, "分组排队设置", mobjfrmTechnicGroupCfg.hWnd, 0
        End If
        
        .InsertItem 4, "输入经过控制", mobjfrmTabPass.hWnd, 0
        .InsertItem 5, "输入必录控制", mobjfrmEnableCtr.hWnd, 0
        .InsertItem 6, "PACS报告设置", mobjFrmReportSetup.hWnd, 0
        .InsertItem 7, "检查列表设置", mobjFrmStudyListCfg.hWnd, 0
        
        framWorkFlow.BorderStyle = 0
        .Item(0).Selected = True
    End With
    framWorkFlow.Width = Me.ScaleWidth
    framWorkFlow.Height = Me.ScaleHeight
    frmTechnicRoom.Width = Me.ScaleWidth
    frmTechnicRoom.Height = Me.ScaleHeight
    mobjfrmTabPass.Width = Me.ScaleWidth
    mobjfrmTabPass.Height = Me.ScaleHeight
    mobjfrmEnableCtr.Width = Me.ScaleWidth
    mobjfrmEnableCtr.Height = Me.ScaleHeight
    mobjFrmReportSetup.Width = Me.ScaleWidth
    mobjFrmReportSetup.Height = Me.ScaleHeight
    mobjFrmStudyListCfg.Width = Me.ScaleWidth
    mobjFrmStudyListCfg.Height = Me.ScaleHeight
    mobjfrmTechnicGroupCfg.Width = Me.ScaleWidth
    mobjfrmTechnicGroupCfg.Height = Me.ScaleHeight
End Sub

Private Sub frmTechRoomRefresh()
    '刷新执行间页面
    frmTechnicRoom.mlngdept = mlng科室ID
    frmTechnicRoom.zlRoomRef
End Sub

Private Sub frmReqInputRefresh(ByVal intType As Integer)
    If intType = 0 Then
        mobjfrmTabPass.mlngDeptId = mlng科室ID
        mobjfrmTabPass.zlRefresh
    ElseIf intType = 1 Then
        mobjfrmEnableCtr.mlngDeptId = mlng科室ID
        mobjfrmEnableCtr.zlRefresh
    End If
End Sub

Private Sub frmStudyListCfgRefresh()
    Call mobjFrmStudyListCfg.zlRefresh(mlng科室ID)
End Sub


Private Sub RefreshTechnicRoomGroupCfg()
'刷新执行间分组配置
    Call mobjfrmTechnicGroupCfg.zlRefresh(mlng科室ID)
End Sub


Private Sub frmWorkFlowRefresh()
    Dim rsTemp As ADODB.Recordset
    Dim lngHintType As Long
        
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
    chkCriticalValues.value = 1      '危急情况判断默认为选中
    txtImageLevel.Text = "甲,乙"     '默认影像质量等级
    txtReportLevel.Text = "甲,乙"    '默认报告质量等级
    chkPetitionCapture.value = 1     '默认勾选启用申请单扫描
    chkAddons.value = 1              '在登记窗口显示附加主述
    chkReagent.value = 1             '在登记窗口显示造影剂
    
    On Error GoTo err
    
    lngHintType = Val(GetDeptPara(mlng科室ID, "诊断结果提示类型", 0))
    optResultInput(lngHintType).value = True
    
    chkIgnorePosi.value = Val(GetDeptPara(mlng科室ID, "忽略结果阴阳性", 0)) '第一次使用时需要重新读取
    chkDefaultPosi.value = Val(GetDeptPara(mlng科室ID, "诊断结果默认阳性", 0))  '读取默认阳性参数
    chkReportAfterResult.value = Val(GetDeptPara(mlng科室ID, "无影像诊断为阴性", 0))
    
    chkCriticalValues.value = Val(GetDeptPara(mlng科室ID, "危急情况判断", 0))    '读取危急情况判断
    chkConformDetermine.value = Val(GetDeptPara(mlng科室ID, "符合情况判定", 0))    '读取符合情况判定
    
    chkImageLevel.value = Val(GetDeptPara(mlng科室ID, "影像质量判定", 0))   '读取影像质量判定
    txtImageLevel.Text = Nvl(GetDeptPara(mlng科室ID, "影像质量等级", "甲,乙"))  '读取影像质量等级
    txtImageLevel.Enabled = chkImageLevel.value = 1
    
    chkReportLevel.value = Val(GetDeptPara(mlng科室ID, "报告质量判定", 0)) '读取报告质量判定
    txtReportLevel.Text = Nvl(GetDeptPara(mlng科室ID, "报告质量等级", "甲,乙"))  '读取报告质量等级
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
    chkPrintCommit.value = Val(GetDeptPara(mlng科室ID, "打印后直接完成", 0))
    chkCompletePrint.value = Val(GetDeptPara(mlng科室ID, "终审后直接打印", 0))
    
    TxtLike.Text = Val(GetDeptPara(mlng科室ID, "登记时姓名模糊查找天数", 0))
    chkSample.value = Val(GetDeptPara(mlng科室ID, "登记后直接检查", 0))
    ChkLike.value = IIf(Val(TxtLike.Text) <> 0, 1, 0)
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
    chkRefreshInterval.value = IIf(Val(txtRefreshInterval.Text) <> 0, 1, 0)
    optMatch(Val(GetDeptPara(mlng科室ID, "匹配数据库项目", 0))).value = True
    
    OptBuildcode(Val(GetDeptPara(mlng科室ID, "检查号生成方式", 0))).value = True
    chkAutoInc.value = Val(GetDeptPara(mlng科室ID, "自动递增检查号"))
    chkAutoSendWorkList.value = Val(GetDeptPara(mlng科室ID, "报道时自动发送WorkList", "1"))
    chkAddons.value = Val(GetDeptPara(mlng科室ID, "显示附加主述", "1"))
    chkReagent.value = Val(GetDeptPara(mlng科室ID, "显示造影剂", "1"))
    
    If chkAutoInc.value = 0 Then
        OptBuildcode(0).Enabled = False
        OptBuildcode(1).Enabled = False
        
        chkChangeNO.value = 1
        chkChangeNO.Enabled = False
        
        chkCheckMaxNo.value = 0
        chkCheckMaxNo.Enabled = False
    Else
        OptBuildcode(0).Enabled = True
        OptBuildcode(1).Enabled = True
        
        chkChangeNO.Enabled = True
        chkCheckMaxNo.Enabled = True
    End If
    
    OptCode(Val(GetDeptPara(mlng科室ID, "患者检查号保持不变", 0))).value = True
    If OptCode(1).value = True Then
        OptUnicode(0).Enabled = True
        OptUnicode(1).Enabled = True
        OptUnicode(Val(GetDeptPara(mlng科室ID, "检查号保持不变类别", 0))).value = True
    Else
        OptUnicode(0).Enabled = False: OptUnicode(0).value = False
        OptUnicode(1).Enabled = False: OptUnicode(1).value = False
    End If
    
    chkUseReferencePatient.value = Val(GetDeptPara(mlng科室ID, "启动关联病人", 0))
    chkPrintNeedComplete.value = Val(GetDeptPara(mlng科室ID, "平诊需审核才能打报告", 0))
    
    '拼音名设置
    optCapital(Val(GetDeptPara(mlng科室ID, "拼音名大小写", 0))).value = True
    optSplitter(Val(GetDeptPara(mlng科室ID, "拼音名分隔符", 0))).value = True
    
    
    gstrSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型=1 and NVL(状态,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTemp.EOF Then
        MsgBoxD Me, "未定义影像存储设备，请到影像设备目录中设置！", vbInformation, gstrSysName
        Exit Sub
    Else
        cboSaveDevice.AddItem ""
        
        Do While Not rsTemp.EOF
            cboSaveDevice.AddItem rsTemp!设备号 & "-" & Nvl(rsTemp!设备名)
            
            If GetDeptPara(mlng科室ID, "存储设备号", "") = rsTemp!设备号 Then
                cboSaveDevice.ListIndex = cboSaveDevice.NewIndex
            End If
            
            rsTemp.MoveNext
        Loop
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub SaveWorkFlow()
    On Error GoTo errHand

    SetDeptPara mlng科室ID, "启用申请单扫描", chkPetitionCapture.value        '启用申请单扫描 参数保存
    
    SetDeptPara mlng科室ID, "符合情况判定", chkConformDetermine.value         '符合情况判定 参数保存
    SetDeptPara mlng科室ID, "危急情况判断", chkCriticalValues.value           '危急情况判断 参数保存
    
    SetDeptPara mlng科室ID, "忽略结果阴阳性", chkIgnorePosi.value
    SetDeptPara mlng科室ID, "无影像诊断为阴性", chkReportAfterResult.value
    SetDeptPara mlng科室ID, "诊断结果默认阳性", chkDefaultPosi.value   '诊断结果默认阳性 参数保存
    
    SetDeptPara mlng科室ID, "影像质量判定", chkImageLevel.value           '影像质量判定 参数保存
    SetDeptPara mlng科室ID, "影像质量等级", txtImageLevel.Text            '图像质量等级 参数保存
    SetDeptPara mlng科室ID, "报告质量判定", chkReportLevel.value          '报告质量判定 参数保存
    SetDeptPara mlng科室ID, "报告质量等级", txtReportLevel.Text           '报告质量等级 参数保存
    
    SetDeptPara mlng科室ID, "诊断结果提示类型", IIf(optResultInput(0).value = True, 0, IIf(optResultInput(1).value = True, 1, 2))
    
    SetDeptPara mlng科室ID, "无报告完成后直接完成", ChkFinishCommit.value
    SetDeptPara mlng科室ID, "采图后医生站即可观片", chkCanViewImage.value     '采图后医生站即可观片
    SetDeptPara mlng科室ID, "有图像才能写报告", chkReportAfterImging.value
    SetDeptPara mlng科室ID, "患者检查号保持不变", IIf(OptCode(1).value, 1, 0)
    SetDeptPara mlng科室ID, "检查号保持不变类别", IIf(OptUnicode(1).value, 1, 0)
    SetDeptPara mlng科室ID, "检查号生成方式", IIf(OptBuildcode(1).value, 1, 0)
    SetDeptPara mlng科室ID, "自动递增检查号", chkAutoInc.value
    SetDeptPara mlng科室ID, "手工调整检查号", chkChangeNO.value
    SetDeptPara mlng科室ID, "允许检查号重复", chkCanOverWrite.value
    SetDeptPara mlng科室ID, "提取实际最大号码", chkCheckMaxNo.value
    SetDeptPara mlng科室ID, "定位片后置", chkLocalizerBackward.value
    SetDeptPara mlng科室ID, "允许交换用户", chkChangeUser.value
    SetDeptPara mlng科室ID, "允许切换用户", chkSwitchUser.value
    SetDeptPara mlng科室ID, "只能填写自己检查的报告", chkTechReportSame.value
    SetDeptPara mlng科室ID, "采集图像者为检查技师", chkWriteCapDoctor.value
    SetDeptPara mlng科室ID, "审核后直接完成", ChkCompleteCommit.value
    SetDeptPara mlng科室ID, "打印后直接完成", chkPrintCommit.value
    SetDeptPara mlng科室ID, "终审后直接打印", chkCompletePrint.value
    SetDeptPara mlng科室ID, "登记后直接检查", chkSample.value
    SetDeptPara mlng科室ID, "匹配数据库项目", IIf(optMatch(0).value, 0, IIf(optMatch(1), 1, 2))
    
    SetDeptPara mlng科室ID, "登记时姓名模糊查找天数", IIf(ChkLike.value = 1, Abs(Val(TxtLike.Text)), 0)
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
    
    SetDeptPara mlng科室ID, "拼音名大小写", IIf(optCapital(0).value, 0, IIf(optCapital(1), 1, 2))
    SetDeptPara mlng科室ID, "拼音名分隔符", IIf(optSplitter(0).value, 0, 1)
    
    If cboSaveDevice.Text <> "" Then
        SetDeptPara mlng科室ID, "存储设备号", Split(cboSaveDevice.Text, "-")(0)
    Else
        SetDeptPara mlng科室ID, "存储设备号", ""
    End If
    
    If Abs(Val(txtRefreshInterval.Text)) = 0 Or Abs(Val(txtRefreshInterval.Text)) > 65 Then
        txtRefreshInterval.Text = 10
    End If
    SetDeptPara mlng科室ID, "自动刷新间隔", IIf(chkRefreshInterval.value = 1, Abs(Val(txtRefreshInterval.Text)), 0)
    SetDeptPara mlng科室ID, "报道时自动发送WorkList", chkAutoSendWorkList.value
    SetDeptPara mlng科室ID, "显示附加主述", chkAddons.value
    SetDeptPara mlng科室ID, "显示造影剂", chkReagent.value
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub


Private Function InitDepts() As Boolean
'功能：初始化科室
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
        MsgBoxD Me, "没有发现医技科室信息,请先到部门管理中设置。", vbInformation, gstrSysName
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
            MsgBoxD Me, "没有发现你所属科室,不能使用医技工作站。", vbInformation, gstrSysName
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
        InitDepts = True
    End If
    
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Unload frmTechnicRoom
    Unload mobjfrmEnableCtr
    Unload mobjfrmTabPass
    Unload mobjFrmReportSetup
    Unload mobjFrmStudyListCfg
    Unload mobjfrmTechnicGroupCfg
End Sub


Private Sub OptCode_Click(Index As Integer)
    OptUnicode(0).Enabled = Index = 1
    OptUnicode(1).Enabled = Index = 1
End Sub

Private Sub frmReportRefresh()
    mobjFrmReportSetup.zlRefresh (mlng科室ID)
End Sub


Private Sub txtViewHistoryImageDays_Change()
    If Val(txtViewHistoryImageDays.Text) > 15 Or Val(txtViewHistoryImageDays.Text) <= 0 Then
        MsgBoxD Me, "自动打开历史图像天数最少为1天，最多为15天，请重新填写。", vbOKOnly, "提示信息"
    End If
End Sub

Private Sub Txt默认天数_Change()
    If Val(Txt默认天数.Text) > 15 Or Val(Txt默认天数.Text) <= 0 Then
        MsgBoxD Me, "默认天数最少为1天，最多为15天，请重新填写。", vbOKOnly, "提示信息"
    End If
End Sub
