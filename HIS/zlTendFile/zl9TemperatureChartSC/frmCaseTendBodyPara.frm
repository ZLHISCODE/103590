VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCaseTendBodyPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "护理选项"
   ClientHeight    =   7125
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8205
   Icon            =   "frmCaseTendBodyPara.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chk 
      Caption         =   "婴儿体温单首日天数从0开始"
      Height          =   180
      Index           =   16
      Left            =   4995
      TabIndex        =   55
      Top             =   5805
      Width           =   2880
   End
   Begin VB.CheckBox chk 
      Caption         =   "灌肠后大便以分子分母显示"
      Height          =   180
      Index           =   15
      Left            =   4995
      TabIndex        =   54
      Top             =   5505
      Width           =   2880
   End
   Begin VB.CheckBox chk 
      Caption         =   "只在当前页中显示跨页数据（不勾两页均显示）"
      Height          =   180
      Index           =   13
      Left            =   120
      TabIndex        =   50
      Top             =   6120
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "体温单文件的开始时间"
      Height          =   645
      Left            =   5385
      TabIndex        =   56
      Top             =   3825
      Width           =   2790
      Begin VB.OptionButton opt体温单开始时间 
         Caption         =   "入科时间"
         Height          =   195
         Index           =   1
         Left            =   1605
         TabIndex        =   58
         Top             =   300
         Width           =   1125
      End
      Begin VB.OptionButton opt体温单开始时间 
         Caption         =   "入院时间"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   57
         Top             =   300
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.CheckBox chk 
      Caption         =   "体温自动标志按顺序在当天排列"
      Height          =   180
      Index           =   12
      Left            =   5000
      TabIndex        =   53
      Top             =   5205
      Width           =   2895
   End
   Begin VB.CheckBox chk 
      Caption         =   "汇总、波动项目显示当天数据（不勾显示昨天）"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   49
      Top             =   5805
      Width           =   4215
   End
   Begin VB.CheckBox chk 
      Caption         =   "体温单输出时打印医院名称"
      Height          =   180
      Index           =   1
      Left            =   5000
      TabIndex        =   52
      Top             =   4905
      Width           =   2895
   End
   Begin VB.Frame FraSplit 
      Height          =   45
      Left            =   120
      TabIndex        =   62
      Top             =   6540
      Width           =   8055
   End
   Begin VB.Frame fra 
      Caption         =   "护理小结标志"
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox cboNodule 
         Height          =   300
         ItemData        =   "frmCaseTendBodyPara.frx":000C
         Left            =   1320
         List            =   "frmCaseTendBodyPara.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2100
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "小结缺省格式"
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.CheckBox chk 
      Caption         =   "护理文件页码按文件格式顺序编号"
      Height          =   180
      Index           =   11
      Left            =   120
      TabIndex        =   46
      Top             =   4905
      Width           =   3135
   End
   Begin VB.CheckBox chk 
      Caption         =   "护理文件放大模式（不勾此项显示标准大小）"
      Height          =   180
      Index           =   10
      Left            =   120
      TabIndex        =   47
      Top             =   5205
      Width           =   4095
   End
   Begin VB.CheckBox chk 
      Caption         =   "住院病人同一时间需要记录多份护理文件"
      Height          =   180
      Index           =   9
      Left            =   120
      TabIndex        =   45
      Top             =   4605
      Width           =   4095
   End
   Begin VB.CheckBox chk 
      Caption         =   "婴儿体温单显示出院信息"
      Height          =   180
      Index           =   5
      Left            =   5000
      TabIndex        =   51
      Top             =   4605
      Width           =   2535
   End
   Begin VB.CheckBox chk 
      Caption         =   "呼吸项表格数据打印输出时上下显示（无数据继承)"
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   48
      Top             =   5505
      Width           =   4400
   End
   Begin VB.Frame fra 
      Caption         =   "体温自动标志"
      Height          =   3615
      Index           =   15
      Left            =   5400
      TabIndex        =   61
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   7
         ItemData        =   "frmCaseTendBodyPara.frx":0010
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   2835
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   6
         ItemData        =   "frmCaseTendBodyPara.frx":0014
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   2466
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   0
         ItemData        =   "frmCaseTendBodyPara.frx":0018
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":001A
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   270
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   1
         ItemData        =   "frmCaseTendBodyPara.frx":001C
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   636
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   2
         ItemData        =   "frmCaseTendBodyPara.frx":0020
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1002
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   3
         ItemData        =   "frmCaseTendBodyPara.frx":0024
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":0026
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1368
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   4
         ItemData        =   "frmCaseTendBodyPara.frx":0028
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1734
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   5
         ItemData        =   "frmCaseTendBodyPara.frx":002C
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   2100
         Width           =   2100
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生"
         Height          =   180
         Index           =   2
         Left            =   135
         TabIndex        =   42
         Top             =   2895
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "分娩"
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   40
         Top             =   2526
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院"
         Height          =   180
         Index           =   44
         Left            =   135
         TabIndex        =   28
         Top             =   315
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入科"
         Height          =   180
         Index           =   45
         Left            =   135
         TabIndex        =   30
         Top             =   675
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转科"
         Height          =   180
         Index           =   46
         Left            =   135
         TabIndex        =   32
         Top             =   1062
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "换床"
         Height          =   180
         Index           =   48
         Left            =   135
         TabIndex        =   34
         Top             =   1428
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "手术"
         Height          =   180
         Index           =   49
         Left            =   135
         TabIndex        =   36
         Top             =   1794
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院"
         Height          =   180
         Index           =   50
         Left            =   135
         TabIndex        =   38
         Top             =   2160
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5760
      TabIndex        =   59
      Top             =   6690
      Width           =   1100
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6960
      TabIndex        =   60
      Top             =   6690
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   3750
      Index           =   0
      Left            =   120
      TabIndex        =   44
      Top             =   720
      Width           =   5175
      Begin VB.CheckBox chk 
         Caption         =   "体温单打印时，不输出心率列(仅在心率单独使用有效)"
         Height          =   315
         Index           =   14
         Left            =   210
         TabIndex        =   27
         Top             =   3345
         Width           =   4770
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   3
         Left            =   2900
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "18"
         Top             =   885
         Width           =   350
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   4
         Left            =   4250
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "6"
         Top             =   885
         Width           =   350
      End
      Begin VB.ComboBox cboSplit 
         Height          =   300
         ItemData        =   "frmCaseTendBodyPara.frx":0030
         Left            =   2400
         List            =   "frmCaseTendBodyPara.frx":0032
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1860
         Width           =   900
      End
      Begin VB.CheckBox chk 
         Caption         =   "体温单输出时，显示皮试结果"
         Height          =   315
         Index           =   8
         Left            =   210
         TabIndex        =   26
         Top             =   3060
         Width           =   2790
      End
      Begin VB.CheckBox chk 
         Caption         =   "体温单以单格显示(不勾此项以双格显示)"
         Height          =   315
         Index           =   7
         Left            =   210
         TabIndex        =   25
         Top             =   2760
         Width           =   3630
      End
      Begin VB.CheckBox chk 
         Caption         =   "体温单上显示病人的诊断信息"
         Height          =   315
         Index           =   3
         Left            =   210
         TabIndex        =   24
         Top             =   2475
         Width           =   2790
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1665
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   19
         Text            =   "1"
         Top             =   1515
         Width           =   420
      End
      Begin VB.CheckBox chk 
         Caption         =   "未记说明显示在体温单的下面（不勾此项时显示在上面）"
         Height          =   315
         Index           =   2
         Left            =   210
         TabIndex        =   23
         Top             =   2175
         Width           =   4800
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2580
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "0"
         Top             =   1185
         Width           =   375
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "0"
         Top             =   885
         Width           =   370
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   1155
         TabIndex        =   4
         Text            =   "14"
         Top             =   270
         Width           =   255
      End
      Begin VB.CheckBox chk 
         Caption         =   "手术后标注天数内再次手术时,停止前一次手术标注"
         Height          =   375
         Index           =   0
         Left            =   195
         TabIndex        =   5
         Top             =   525
         Width           =   4500
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   6
         Left            =   2050
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   885
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txt(6)"
         BuddyDispid     =   196622
         BuddyIndex      =   6
         OrigLeft        =   2190
         OrigTop         =   870
         OrigRight       =   2430
         OrigBottom      =   1170
         Max             =   4
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   1
         Left            =   2970
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1185
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txt(1)"
         BuddyDispid     =   196622
         BuddyIndex      =   1
         OrigLeft        =   2190
         OrigTop         =   870
         OrigRight       =   2430
         OrigBottom      =   1170
         Max             =   30
         Min             =   2
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   0
         Left            =   2100
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1515
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txt(2)"
         BuddyDispid     =   196622
         BuddyIndex      =   2
         OrigLeft        =   2190
         OrigTop         =   870
         OrigRight       =   2430
         OrigBottom      =   1170
         Max             =   30
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   3
         Left            =   4580
         TabIndex        =   14
         Top             =   885
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   8
         BuddyControl    =   "txt(4)"
         BuddyDispid     =   196622
         BuddyIndex      =   4
         OrigLeft        =   4580
         OrigTop         =   885
         OrigRight       =   4835
         OrigBottom      =   1170
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   2
         Left            =   3230
         TabIndex        =   11
         Top             =   885
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   18
         BuddyControl    =   "txt(3)"
         BuddyDispid     =   196622
         BuddyIndex      =   3
         OrigLeft        =   3230
         OrigTop         =   885
         OrigRight       =   3485
         OrigBottom      =   1170
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "夜班从       点至"
         Height          =   180
         Index           =   7
         Left            =   2350
         TabIndex        =   9
         Top             =   945
         Width           =   1530
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "次日       点"
         Height          =   180
         Index           =   8
         Left            =   3885
         TabIndex        =   12
         Top             =   945
         Width           =   1170
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体温自动标志与时间之间以           相连"
         Height          =   180
         Index           =   6
         Left            =   210
         TabIndex        =   21
         Top             =   1920
         Width           =   3510
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "允许录入超过当前        天的护理记录数据"
         Height          =   180
         Index           =   4
         Left            =   210
         TabIndex        =   18
         Top             =   1560
         Width           =   3600
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体温表输出时，表格数据固定        行"
         Height          =   180
         Index           =   3
         Left            =   210
         TabIndex        =   15
         Top             =   1230
         Width           =   3240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体温开始记录时间"
         Height          =   180
         Index           =   31
         Left            =   210
         TabIndex        =   6
         Top             =   945
         Width           =   1440
      End
      Begin VB.Line Line1 
         X1              =   1125
         X2              =   1410
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "手术后标注    天"
         Height          =   180
         Index           =   0
         Left            =   195
         TabIndex        =   3
         Top             =   270
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmCaseTendBodyPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mfrmMain As Object
Private mblnOK As Boolean
Private mstrPrivs As String

Public Function ShowPara(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    Dim intLoop As Integer
    Dim strTmp As String
    Dim strSQL As String, strPar As String
    Dim curDate As Date, intDay As Integer
    Dim intStart As Integer
    
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    
    '初始体温单标记
    '------------------------------------------------------------------------------------------------------------------
    cboBody(0).Clear
    cboBody(0).AddItem "0-不显示"
    cboBody(0).AddItem "1-显示说明"
    cboBody(0).AddItem "2-显示说明和时间"
    
    cboBody(1).Clear
    cboBody(1).AddItem "0-不显示"
    cboBody(1).AddItem "1-显示说明"
    cboBody(1).AddItem "2-显示说明和时间"
    
    cboBody(2).Clear
    cboBody(2).AddItem "0-不显示"
    cboBody(2).AddItem "1-显示说明"
    cboBody(2).AddItem "2-显示说明和时间"
    cboBody(2).AddItem "3-显示说明和科室"
    cboBody(2).AddItem "4-显示说明,科室,时间"
    
    cboBody(3).Clear
    cboBody(3).AddItem "0-不显示"
    cboBody(3).AddItem "1-显示说明"
    cboBody(3).AddItem "2-显示说明和时间"
    
    cboBody(4).Clear
    cboBody(4).AddItem "0-不显示"
    cboBody(4).AddItem "1-显示说明"
    cboBody(4).AddItem "2-显示说明和时间"
    
    cboBody(5).Clear
    cboBody(5).AddItem "0-不显示"
    cboBody(5).AddItem "1-显示说明"
    cboBody(5).AddItem "2-显示说明和时间"
    
    cboBody(6).Clear
    cboBody(6).AddItem "0-不显示"
    cboBody(6).AddItem "1-显示说明"
    cboBody(6).AddItem "2-显示说明和时间"
    
    cboBody(7).Clear
    cboBody(7).AddItem "0-不显示"
    cboBody(7).AddItem "1-显示说明"
    cboBody(7).AddItem "2-显示说明和时间"
    
    cboNodule.Clear
    cboNodule.AddItem "0-不处理"
    cboNodule.AddItem "1-上下红线"
    cboNodule.AddItem "2-下双红线"
    cboNodule.AddItem "3-上红线"
    
    cboSplit.Clear
    cboSplit.AddItem "——"
    cboSplit.AddItem "于"
    
    intStart = zldatabase.GetPara("体温单文件开始时间", glngSys, 1255, 1, Array(opt体温单开始时间(0), opt体温单开始时间(1)), InStr(mstrPrivs, "护理选项设置") > 0)
    opt体温单开始时间(intStart).Value = True
    txt(6).Text = zldatabase.GetPara("体温开始时间", glngSys, 1255, 4, Array(txt(6), ud(6), lbl(31)), InStr(mstrPrivs, "护理选项设置") > 0)
    txt(1).Text = zldatabase.GetPara("体温表格行数", glngSys, 1255, 8, Array(txt(1), ud(1), lbl(3)), InStr(mstrPrivs, "护理选项设置") > 0)
    strTmp = zldatabase.GetPara("体温单标记", glngSys, 1255, "1;1;1;1;1;1;1:1", Array(cboBody(0), cboBody(1), cboBody(2), cboBody(3), cboBody(4), cboBody(5), cboBody(6), cboBody(7)), InStr(mstrPrivs, "护理选项设置") > 0)
    
    For intLoop = 0 To 7
        If UBound(Split(strTmp, ";")) >= intLoop Then
            cboBody(intLoop).ListIndex = Val(Split(strTmp, ";")(intLoop))
        Else
            cboBody(intLoop).ListIndex = 0
        End If
    Next
    strTmp = zldatabase.GetPara("小结缺省格式", glngSys, 1255, "0", Array(cboNodule, lbl(5)), InStr(mstrPrivs, "护理选项设置") > 0)
    
    If Val(strTmp) >= 0 And Val(strTmp) <= 2 Then
        cboNodule.ListIndex = Val(strTmp)
    Else
        cboNodule.ListIndex = 0
    End If
    
    strTmp = zldatabase.GetPara("体温标志分隔符", glngSys, 1255, "0", Array(cboSplit, lbl(6)), InStr(mstrPrivs, "护理选项设置") > 0)
    
    If Val(strTmp) >= 0 And Val(strTmp) <= 1 Then
        cboSplit.ListIndex = Val(strTmp)
    Else
        cboSplit.ListIndex = 0
    End If
    
    '体温夜班标志
    strTmp = zldatabase.GetPara("体温时间夜班标志", glngSys, 1255, "18;6", Array(lbl(7), txt(3), ud(2), lbl(8), txt(4), ud(3)), InStr(mstrPrivs, "护理选项设置") > 0)
    If UBound(Split(strTmp, ";")) >= 1 Then
        txt(3).Text = Abs(Val(Split(strTmp, ";")(0)))
        txt(4).Text = Abs(Val(Split(strTmp, ";")(1)))
    Else
         txt(3).Text = Abs(Val(strTmp))
    End If
    
    txt(0).Text = Val(zldatabase.GetPara("手术后标注天数", glngSys, 1255, "10", Array(txt(0), lbl(0)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(0).Value = Val(zldatabase.GetPara("再次手术停止前次标注", glngSys, 1255, "0", Array(chk(0)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(1).Value = Val(zldatabase.GetPara("打印医院名称", glngSys, 1255, "1", Array(chk(1)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(2).Value = Val(zldatabase.GetPara("未记说明显示位置", glngSys, 1255, "0", Array(chk(2)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(3).Value = Val(zldatabase.GetPara("体温单显示诊断", glngSys, 1255, "1", Array(chk(3)), InStr(mstrPrivs, "护理选项设置") > 0))
    txt(2).Text = Val(zldatabase.GetPara("超期录入护理数据天数", glngSys, 1255, "1", Array(txt(2), lbl(4)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(4).Value = Val(zldatabase.GetPara("呼吸表格输出", glngSys, 1255, "0", Array(chk(4)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(5).Value = Val(zldatabase.GetPara("婴儿体温单显示出院信息", glngSys, 1255, "1", Array(chk(5)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(6).Value = Val(zldatabase.GetPara("汇总波动显示当天数据", glngSys, 1255, "1", Array(chk(6)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(7).Value = Val(zldatabase.GetPara("体温单显示格式", glngSys, 1255, "0", Array(chk(7)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(8).Value = Val(zldatabase.GetPara("体温单显示皮试结果", glngSys, 1255, "0", Array(chk(8))))
    chk(9).Value = Val(zldatabase.GetPara("对应多份护理文件", glngSys, 1255, "0", Array(chk(9)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(10).Value = Val(zldatabase.GetPara("护理文件显示模式", glngSys, 1255, "0", Array(chk(10)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(11).Value = Val(zldatabase.GetPara("护理文件页码规则", glngSys, 1255, "0", Array(chk(11)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(12).Value = Val(zldatabase.GetPara("体温标志按顺序当天排列", glngSys, 1255, "0", Array(chk(12)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(13).Value = Val(zldatabase.GetPara("跨页数据只显示在第一页", glngSys, 1255, "0", Array(chk(13)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(14).Value = Val(zldatabase.GetPara("体温单不打印心率列", glngSys, 1255, "0", Array(chk(14)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(15).Value = Val(zldatabase.GetPara("灌肠后大便显示格式", glngSys, 1255, "0", Array(chk(15)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(16).Value = Val(zldatabase.GetPara("婴儿体温单首日天数显示0", glngSys, 1255, "0", Array(chk(16)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(11).Enabled = Not CheckPrintDate
    
    Me.Show 1, mfrmMain
    ShowPara = mblnOK
    
End Function

Private Function CheckPrintDate() As Boolean
'---------------------------------------------------------
'功能:'检查病人是否存在打印数据,如果存在就不允许设置护理文件页码规则
'---------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    CheckPrintDate = False
    
    strSQL = "Select 1 From  病人护理打印  Where 打印页号 is not null and   Rownum<2"
    Call zldatabase.OpenRecordset(rsTemp, strSQL, "检查是否存在打印数据")
    If rsTemp.RecordCount > 0 Then
        CheckPrintDate = True
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cboBody_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intStart As Integer
    Dim strTmp As String
    
    If opt体温单开始时间(0).Value Then
        intStart = 0
    Else
        intStart = 1
    End If
    
    strTmp = cboBody(0).ListIndex & ";" & cboBody(1).ListIndex & ";" & cboBody(2).ListIndex & ";" & cboBody(3).ListIndex & ";" & cboBody(4).ListIndex & ";" & cboBody(5).ListIndex & ";" & cboBody(6).ListIndex & ";" & cboBody(7).ListIndex
    Call zldatabase.SetPara("体温开始时间", Val(txt(6).Text), glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("体温表格行数", Val(txt(1).Text), glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("体温单标记", strTmp, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("手术后标注天数", Val(txt(0).Text), glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("超期录入护理数据天数", Val(txt(2).Text), glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("再次手术停止前次标注", chk(0).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("未记说明显示位置", chk(2).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("体温单显示诊断", chk(3).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("打印医院名称", chk(1).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("呼吸表格输出", chk(4).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("婴儿体温单显示出院信息", chk(5).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("汇总波动显示当天数据", chk(6).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("体温单显示格式", chk(7).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("体温单显示皮试结果", chk(8).Value, glngSys, 1255)
    Call zldatabase.SetPara("对应多份护理文件", chk(9).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("护理文件显示模式", chk(10).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("护理文件页码规则", chk(11).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("小结缺省格式", Val(cboNodule.ListIndex), glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("体温标志分隔符", Val(cboSplit.ListIndex), glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("体温标志按顺序当天排列", chk(12).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("体温单文件开始时间", intStart, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("体温时间夜班标志", txt(3).Text & ";" & txt(4).Text, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("跨页数据只显示在第一页", chk(13).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("体温单不打印心率列", chk(14).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("灌肠后大便显示格式", chk(15).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zldatabase.SetPara("婴儿体温单首日天数显示0", chk(16).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    
    mblnOK = True
    
    Unload Me
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt(Index))
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

