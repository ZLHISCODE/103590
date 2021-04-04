VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm部门发药参数设置 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   5775
   ClientLeft      =   8805
   ClientTop       =   3960
   ClientWidth     =   6735
   Icon            =   "Frm部门发药参数设置.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6735
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4200
      TabIndex        =   0
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5400
      TabIndex        =   1
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   1100
   End
   Begin TabDlg.SSTab tabShow 
      Height          =   5010
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   8837
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "基础(&1)"
      TabPicture(0)   =   "Frm部门发药参数设置.frx":1CFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblNote(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lbl发药药房"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblNote(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Lbl操作模式"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Cbo记帐人"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Cbo发药药房"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Cbo操作模式"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Chk按科室汇总显示"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkDetailPage"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fra发药数据检查"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "辅助(&2)"
      TabPicture(1)   =   "Frm部门发药参数设置.frx":1D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboName"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frm高危药品发放"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cbo退药清单"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cbo发药清单"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fra设备定义"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblName"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lbl退药清单"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lbl发药清单"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "提醒(&3)"
      TabPicture(2)   =   "Frm部门发药参数设置.frx":1D32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "领药部门(&4)"
      TabPicture(3)   =   "Frm部门发药参数设置.frx":1D4E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Lvw来源科室"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "包药机(&5)"
      TabPicture(4)   =   "Frm部门发药参数设置.frx":1D6A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame5"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "打印设置(&6)"
      TabPicture(5)   =   "Frm部门发药参数设置.frx":1D86
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmd打印设置"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "cbo票据设置"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "lbl票据"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      Begin VB.Frame fra发药数据检查 
         Caption         =   " 发药数据检查"
         Height          =   1335
         Left            =   240
         TabIndex        =   56
         Top             =   3360
         Width           =   5895
         Begin VB.CheckBox chk检查销帐申请 
            Caption         =   "检查和提醒销帐申请数据"
            Height          =   180
            Left            =   240
            TabIndex        =   58
            Top             =   720
            Width           =   3825
         End
         Begin VB.CheckBox chk检查存储库房 
            Caption         =   "检查储存库房"
            Height          =   180
            Left            =   240
            TabIndex        =   57
            Top             =   360
            Width           =   1785
         End
      End
      Begin VB.CheckBox chkDetailPage 
         Caption         =   "保持上一次窗体关闭时的页签"
         Height          =   180
         Left            =   240
         TabIndex        =   55
         Top             =   2880
         Width           =   2745
      End
      Begin VB.CheckBox Chk按科室汇总显示 
         Caption         =   "按科室汇总显示"
         Height          =   180
         Left            =   240
         TabIndex        =   54
         Top             =   2520
         Width           =   1785
      End
      Begin VB.CommandButton cmd打印设置 
         Caption         =   "打印设置(&P)"
         Height          =   345
         Left            =   -74760
         TabIndex        =   52
         Top             =   1050
         Width           =   3315
      End
      Begin VB.ComboBox cbo票据设置 
         Height          =   300
         Left            =   -74010
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   600
         Width           =   2565
      End
      Begin VB.Frame Frame4 
         Caption         =   " 传送控制  "
         Height          =   615
         Left            =   -74880
         TabIndex        =   49
         Top             =   360
         Width           =   4935
         Begin VB.CheckBox chkStopTrans 
            Caption         =   "暂停向药品包装机传送发药数据"
            Height          =   255
            Left            =   360
            TabIndex        =   50
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 传送数据内容控制  "
         Height          =   3855
         Left            =   -74880
         TabIndex        =   43
         Top             =   1080
         Width           =   4935
         Begin VB.Frame Frame6 
            Caption         =   " 单据类型  "
            Height          =   615
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   4695
            Begin VB.CheckBox chkType 
               Caption         =   "长嘱"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   48
               Top             =   240
               Value           =   1  'Checked
               Width           =   975
            End
            Begin VB.CheckBox chkType 
               Caption         =   "临嘱"
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   47
               Top             =   240
               Value           =   1  'Checked
               Width           =   975
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   " 剂型选择"
            Height          =   2775
            Left            =   120
            TabIndex        =   44
            Top             =   960
            Width           =   4695
            Begin MSComctlLib.ListView Lvw药品剂型 
               Height          =   2385
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   4425
               _ExtentX        =   7805
               _ExtentY        =   4207
               View            =   2
               Arrange         =   1
               LabelEdit       =   1
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               Icons           =   "imgLvwSel"
               SmallIcons      =   "imgLvwSel"
               ColHdrIcons     =   "imgLvwSel"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "名称"
                  Object.Width           =   3528
               EndProperty
            End
         End
      End
      Begin VB.ComboBox cboName 
         ForeColor       =   &H80000012&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   -74040
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Frame frm高危药品发放 
         Caption         =   " 选择高危药品单独发放的类别"
         Height          =   580
         Left            =   -74880
         TabIndex        =   36
         Top             =   1845
         Width           =   6135
         Begin VB.CheckBox chk高危 
            Caption         =   "C类"
            Height          =   375
            Index           =   2
            Left            =   2040
            TabIndex        =   39
            Top             =   180
            Width           =   615
         End
         Begin VB.CheckBox chk高危 
            Caption         =   "B类"
            Height          =   375
            Index           =   1
            Left            =   1140
            TabIndex        =   38
            Top             =   180
            Width           =   615
         End
         Begin VB.CheckBox chk高危 
            Caption         =   "A类"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   37
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.ComboBox cbo退药清单 
         Height          =   300
         Left            =   -74040
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   945
         Width           =   2655
      End
      Begin VB.ComboBox cbo发药清单 
         Height          =   300
         Left            =   -74040
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   555
         Width           =   2655
      End
      Begin VB.Frame fra设备定义 
         Caption         =   "  智能卡及其他设备定义 "
         Height          =   1095
         Left            =   -71280
         TabIndex        =   28
         Top             =   510
         Width           =   2415
         Begin VB.CommandButton cmdDeviceSetup 
            Caption         =   "设备配置(&S)"
            Height          =   350
            Left            =   480
            TabIndex        =   29
            Top             =   360
            Width           =   1500
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 查询明细记录条数，超过时提醒"
         Height          =   1095
         Left            =   -74760
         TabIndex        =   24
         Top             =   2040
         Width           =   5295
         Begin VB.TextBox txtMaxRecordCount 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1440
            TabIndex        =   25
            Text            =   "3000"
            Top             =   420
            Width           =   645
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "条"
            Height          =   180
            Left            =   2160
            TabIndex        =   27
            Top             =   480
            Width           =   180
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "查询明细记录"
            Height          =   180
            Left            =   240
            TabIndex        =   26
            Top             =   480
            Width           =   1080
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " 设置查询发药、退药科室时的时间范围，超过时提醒"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   17
         Top             =   480
         Width           =   5295
         Begin VB.TextBox txtTimeArea_Sended 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   21
            Text            =   "3"
            Top             =   840
            Width           =   405
         End
         Begin VB.TextBox txtTimeArea_Send 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   18
            Text            =   "7"
            Top             =   360
            Width           =   405
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "天"
            Height          =   180
            Left            =   1920
            TabIndex        =   23
            Top             =   900
            Width           =   180
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "查询退药天数"
            Height          =   180
            Left            =   240
            TabIndex        =   22
            Top             =   900
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "天"
            Height          =   180
            Left            =   1920
            TabIndex        =   20
            Top             =   420
            Width           =   180
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "查询发药天数"
            Height          =   180
            Left            =   240
            TabIndex        =   19
            Top             =   420
            Width           =   1080
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 选择在发药时自动标记为不处理的药品类型"
         Height          =   2340
         Left            =   -74880
         TabIndex        =   12
         Top             =   2640
         Width           =   6135
         Begin MSComctlLib.ListView lvw价值分类 
            Height          =   1755
            Left            =   2160
            TabIndex        =   16
            Top             =   480
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   3096
            View            =   2
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lvw毒理分类 
            Height          =   1750
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   3096
            View            =   2
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lvw高危分类 
            Height          =   1755
            Left            =   4200
            TabIndex        =   34
            Top             =   480
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   3096
            View            =   2
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "高危药品等级分类"
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   4200
            TabIndex        =   35
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "药品价值分类"
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   2160
            TabIndex        =   15
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "药品毒理分类"
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1080
         End
      End
      Begin VB.ComboBox Cbo操作模式 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox Cbo发药药房 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   690
         Width           =   1815
      End
      Begin VB.ComboBox Cbo记帐人 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2025
         Width           =   1815
      End
      Begin MSComctlLib.ListView Lvw来源科室 
         Height          =   4605
         Left            =   -74880
         TabIndex        =   40
         Top             =   360
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   8123
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label lbl票据 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "票据(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74730
         TabIndex        =   53
         Top             =   660
         Width           =   630
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "药名显示"
         Height          =   180
         Left            =   -74880
         TabIndex        =   42
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label lbl退药清单 
         Caption         =   "退药清单"
         Height          =   195
         Left            =   -74880
         TabIndex        =   33
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label lbl发药清单 
         Caption         =   "发药清单"
         Height          =   195
         Left            =   -74880
         TabIndex        =   31
         Top             =   615
         Width           =   735
      End
      Begin VB.Label Lbl操作模式 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单据类型"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   180
         TabIndex        =   11
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label LblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "你所属的药房"
         ForeColor       =   &H00000080&
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Lbl发药药房 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "发药药房"
         Height          =   180
         Left            =   180
         TabIndex        =   9
         Top             =   750
         Width           =   720
      End
      Begin VB.Label LblNote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "你所要操作的是处方单、记帐表亦或两者兼有"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   1200
         Width           =   4710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "记 帐 人"
         Height          =   180
         Left            =   180
         TabIndex        =   7
         Top             =   2085
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm部门发药参数设置.frx":1DA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm部门发药参数设置.frx":20BC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Frm部门发药参数设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strPrivs As String
Private mblnSetPara As Boolean                          '是否具有参数设置权限
Private BlnStart As Boolean
Private intDays As Integer
Private lng药房ID As Long
Private Lng操作模式 As Long
Private Lng汇总显示 As Long
Private Lng自动打印 As Long
Private Lng缺药检查 As Long
Private Lng领药人签名 As Long
Private Lng退药人签名 As Long
Private str毒理分类 As String
Private str价值分类 As String
Private RecDrugStore As New ADODB.Recordset             '药房
Private mstrSourceDep As String                         '来源科室
Private mLng打印退药清单 As Long                        '退药清单
Public blnStartPacker As Boolean                       '是否启用药品分包机接口
Private Sub Get记帐人()
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    
    '设置记帐人
    On Error GoTo errHandle
    strsql = "Select Distinct A.姓名" & _
             " From 人员表 A,部门人员 B,部门性质说明 C,人员性质说明 D " & _
             " Where A.Id=B.人员id And B.部门id=C.部门Id And D.人员id=A.Id And D.人员性质 = '药房发药人' " & _
             " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) "
        
    If Cbo发药药房.ListIndex <> -1 Then
        strsql = strsql & " AND B.部门id=[1] "
    End If
    
    strsql = strsql & " ORDER BY A.姓名 "

    Set rs = zldatabase.OpenSQLRecord(strsql, Me.Caption, Cbo发药药房.ItemData(Cbo发药药房.ListIndex))
    
    Cbo记帐人.Clear
    Cbo记帐人.AddItem "所有记帐人"
    Do While Not rs.EOF
        Cbo记帐人.AddItem rs!姓名
        rs.MoveNext
    Loop
    
    rs.Close
    
    Cbo记帐人.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub Cbo发药药房_Click()
    Call Get记帐人
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call FS.DeviceSetup(Me, 100, 1342)
    
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    Dim n As Integer
    Dim str单据 As String
    Dim str剂型 As String
    Dim i As Integer
    Dim str高危分类 As String
    Dim str高危发放 As String
    
    str毒理分类 = ""
    For n = 1 To lvw毒理分类.ListItems.count
        If lvw毒理分类.ListItems(n).Checked = True Then
            str毒理分类 = IIf(str毒理分类 = "", lvw毒理分类.ListItems(n).Text, str毒理分类 & "," & lvw毒理分类.ListItems(n).Text)
        End If
    Next
    
    str价值分类 = ""
    For n = 1 To lvw价值分类.ListItems.count
        If lvw价值分类.ListItems(n).Checked = True Then
            str价值分类 = IIf(str价值分类 = "", lvw价值分类.ListItems(n).Text, str价值分类 & "," & lvw价值分类.ListItems(n).Text)
        End If
    Next
    
    For n = 1 To lvw高危分类.ListItems.count
        If lvw高危分类.ListItems(n).Checked = True Then
            str高危分类 = IIf(str高危分类 = "", n, str高危分类 & "," & n)
        End If
    Next
    
    If chk高危(0).Value = 1 Then str高危发放 = IIf(str高危发放 = "", 1, str高危发放 & "," & 1)
    If chk高危(1).Value = 1 Then str高危发放 = IIf(str高危发放 = "", 2, str高危发放 & "," & 2)
    If chk高危(2).Value = 1 Then str高危发放 = IIf(str高危发放 = "", 3, str高危发放 & "," & 3)
    
    '保存私有参数
    zldatabase.SetPara "按科室汇总显示汇总清单", Chk按科室汇总显示.Value, glngSys, 1342
    zldatabase.SetPara "操作模式", Cbo操作模式.ListIndex, glngSys, 1342
    zldatabase.SetPara "记帐人", Cbo记帐人.Text, glngSys, 1342
    zldatabase.SetPara "毒理分类", str毒理分类, glngSys, 1342
    zldatabase.SetPara "价值分类", str价值分类, glngSys, 1342
    zldatabase.SetPara "高危分类", str高危分类, glngSys, 1342
    zldatabase.SetPara "高危药品发放", str高危发放, glngSys, 1342
    zldatabase.SetPara "发药药房", Cbo发药药房.ItemData(Cbo发药药房.ListIndex), glngSys, 1342
    zldatabase.SetPara "自动打印", Me.cbo发药清单.ListIndex, glngSys, 1342
    zldatabase.SetPara "查询发药天数", Val(txtTimeArea_Send.Text), glngSys, 1342
    zldatabase.SetPara "查询退药天数", Val(txtTimeArea_Sended.Text), glngSys, 1342
    zldatabase.SetPara "查询明细记录数", Val(txtMaxRecordCount.Text), glngSys, 1342
    zldatabase.SetPara "打印退药清单", Me.cbo退药清单.ListIndex, glngSys, 1342
    zldatabase.SetPara "发药时检查存储库房", Me.chk检查存储库房.Value, glngSys, 1342
    zldatabase.SetPara "发药时检查销帐申请数据", Me.chk检查销帐申请.Value, glngSys, 1342
    
    '来源科室
    mstrSourceDep = ""
    With Me.Lvw来源科室
        For i = 1 To .ListItems.count
            If .ListItems(i).Checked Then
                If mstrSourceDep = "" Then
                    mstrSourceDep = Mid(.ListItems(i).Key, 2)
                Else
                    mstrSourceDep = mstrSourceDep & "," & Mid(.ListItems(i).Key, 2)
                End If
            End If
        Next
    End With
    zldatabase.SetPara "来源科室", mstrSourceDep, glngSys, 1342
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品部门发药管理", "药品名称显示方式", Me.cboName.ListIndex)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品部门发药管理", "保持上一次窗体关闭时的页签", Me.chkDetailPage.Value)
    
    '保存包装机设置
    If blnStartPacker = True Then
        SaveSetting "ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "部门发药管理\包装机设置", "暂停传送", chkStopTrans.Value
        
        str单据 = ""
        str单据 = str单据 & chkType(0).Value
        str单据 = str单据 & chkType(1).Value

        SaveSetting "ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "部门发药管理\包装机设置", "单据类型", str单据
        
        
        If Lvw药品剂型.ListItems(1).Checked Then
             str剂型 = "所有"
        Else
            For n = 1 To Lvw药品剂型.ListItems.count
                If Lvw药品剂型.ListItems(n).Checked Then
                    str剂型 = IIf(str剂型 = "", "", str剂型 & ",") & Lvw药品剂型.ListItems(n).Text
                End If
            Next
        End If
        
        SaveSetting "ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "部门发药管理\包装机设置", "选择剂型", str剂型
    End If
    
'    Frm部门发药管理.BlnSetPara = True
    frm部门发药管理New.BlnSetPara = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If BlnStart = False Then
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strsql As String
    Dim rsTmp As New ADODB.Recordset
    Dim intTrans As Integer
    Dim str单据 As String
    Dim str剂型 As String
    Dim n As Integer
    
    BlnStart = False
    On Error GoTo errHandle
    
    If zlStr.IsHavePrivs(strPrivs, "所有药房") Then
        strsql = "(Select Distinct 部门ID From 部门性质说明 Where 工作性质 Like '%药房' And 服务对象 IN (2,3))"
    Else
        strsql = "(Select distinct A.部门ID From 部门人员 A,部门性质说明 B " & _
                 " Where A.人员ID=[1] And A.部门ID=B.部门ID And B.工作性质 Like '%药房' And B.服务对象 IN (2,3))"
    End If
    gstrSQL = " Select ID,编码||'-'||名称 药房 From 部门表 Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And ID In " & strsql & _
             " And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) " & _
             " Order by 编码||'-'||名称"
    Set RecDrugStore = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId)
    
    With RecDrugStore
        If .EOF Then
            MsgBox "请初始化药房！（部门管理）", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Cbo发药药房.Clear
        Do While Not .EOF
            Cbo发药药房.AddItem !药房
            Cbo发药药房.ItemData(Cbo发药药房.NewIndex) = !Id
            .MoveNext
        Loop
        Cbo发药药房.ListIndex = 0
    End With
    
    With Cbo操作模式
        .Clear
        .AddItem "0-包含所有单据"
        .AddItem "1-仅包含记帐单"
        .AddItem "2-仅包含记帐表"
        .ListIndex = 0
    End With
        
    With cbo发药清单
        .Clear
        .AddItem "0_发药后不打印"
        .AddItem "1-发药后自动打印"
        .AddItem "2_发药后提示是否打印"
        .ListIndex = 0
    End With
    
    With cbo退药清单
        .Clear
        .AddItem "0_退药后不打印"
        .AddItem "1-退药后自动打印"
        .AddItem "2_退药后提示是否打印"
        .ListIndex = 0
    End With
    
    With Me.cboName
        .Clear
        .AddItem "0-显示药品编码与名称"
        .AddItem "1-仅显示药品编码"
        .AddItem "2-仅显示药品名称"
        .ListIndex = 0
    End With
    
    With cbo票据设置
        .Clear
        .AddItem "1-汇总发药清单"
        .AddItem "2-退药清单"
        .ListIndex = 0
    End With
    
    Call Get记帐人
    
    '毒理分类
    gstrSQL = "Select 名称 From 药品毒理分类 Order By 编码 "
    Call zldatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption & "-取毒理分类")
    
    With rsTmp
        Do While Not .EOF
            lvw毒理分类.ListItems.Add , "_" & lvw毒理分类.ListItems.count + 1, !名称
            .MoveNext
        Loop
    End With
    
    '价值分类
    gstrSQL = "Select 名称 From 药品价值分类 Order By 编码 "
    Call zldatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption & "-取价值分类")
    
    With rsTmp
        Do While Not .EOF
            lvw价值分类.ListItems.Add , "_" & lvw价值分类.ListItems.count + 1, !名称
            .MoveNext
        Loop
    End With
    
    '高危药品分类
    With lvw高危分类
        .ListItems.Clear
        .ListItems.Add , "_" & .ListItems.count + 1, "A类"
        .ListItems.Add , "_" & .ListItems.count + 1, "B类"
        .ListItems.Add , "_" & .ListItems.count + 1, "C类"
    End With
    
    '恢复设置
    WriteCons

    '来源科室
    Call SetSourceDep
    
    '包装机接口相关设置
    Call Load药品剂型(Cbo发药药房.ItemData(Cbo发药药房.ListIndex))
    
    tabShow.TabVisible(4) = blnStartPacker
    
    If blnStartPacker = True Then
        intTrans = Val(GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "部门发药管理\包装机设置", "暂停传送", "0"))
        chkStopTrans.Value = IIf(intTrans = 1, 1, 0)
        
        str单据 = GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "部门发药管理\包装机设置", "单据类型", "11")
        chkType(0).Value = Val(Mid(str单据, 1, 1))
        chkType(1).Value = Val(Mid(str单据, 2, 1))
        
        str剂型 = GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "部门发药管理\包装机设置", "选择剂型", "所有")
        
        For n = 1 To Lvw药品剂型.ListItems.count
            Lvw药品剂型.ListItems(n).Checked = False
            If str剂型 = "所有" Then
                Lvw药品剂型.ListItems(n).Checked = True
            Else
                If InStr(1, "," & str剂型 & ",", "," & Lvw药品剂型.ListItems(n).Text & ",") > 0 Then
                    Lvw药品剂型.ListItems(n).Checked = True
                End If
            End If
        Next
    End If
    
    BlnStart = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load药品剂型(ByVal lng药房ID As Long)
    Dim rsData As ADODB.Recordset
    
    Set rsData = DeptSendWork_Get剂型(lng药房ID)
    
    With Lvw药品剂型
        .ListItems.Clear
        .ListItems.Add , "_" & .ListItems.count + 1, "所有药品剂型", 1, 1
        .ListItems(.ListItems.count).Checked = True
        Do While Not rsData.EOF
            .ListItems.Add , "_" & .ListItems.count + 1, Mid(rsData!剂型, InStr(1, rsData!剂型, "-") + 1), 1, 1
            .ListItems(.ListItems.count).Checked = True
            rsData.MoveNext
        Loop
    End With
End Sub
Private Function WriteCons()
    Dim IntLocate As Integer
    Dim str记帐人 As String
    Dim n As Integer
    Dim i As Integer
    Dim int自动刷新 As Integer
    Dim strArr
    Dim int查询发药天数 As Integer
    Dim int查询退药天数 As Integer
    Dim lng最大记录数 As Long
    Dim int审核出院销账申请 As Integer
    Dim str高危分类 As String
    Dim str高危发放 As String
    Dim int汇总销账 As Integer
    Dim int检查存储库房 As Integer
    Dim int检查销帐申请 As Integer
    
    mblnSetPara = zlStr.IsHavePrivs(strPrivs, "参数设置")
    
    '取公共及私有参数
    Lng操作模式 = Val(zldatabase.GetPara("操作模式", glngSys, 1342, 0, Array(Cbo操作模式), mblnSetPara))
    Lng汇总显示 = Val(zldatabase.GetPara("按科室汇总显示汇总清单", glngSys, 1342, 0, Array(Chk按科室汇总显示), mblnSetPara))
    str记帐人 = zldatabase.GetPara("记帐人", glngSys, 1342, "所有记帐人", Array(Label1, Cbo记帐人), mblnSetPara)
    str毒理分类 = zldatabase.GetPara("毒理分类", glngSys, 1342, "", Array(Label2, lvw毒理分类), mblnSetPara)
    str价值分类 = zldatabase.GetPara("价值分类", glngSys, 1342, "", Array(Label3, lvw价值分类), mblnSetPara)
    str高危分类 = zldatabase.GetPara("高危分类", glngSys, 1342, "", Array(Label11, lvw高危分类), mblnSetPara)
    str高危发放 = zldatabase.GetPara("高危药品发放", glngSys, 1342, "", Array(frm高危药品发放), mblnSetPara)
    lng药房ID = Val(zldatabase.GetPara("发药药房", glngSys, 1342, 0, Array(Lbl发药药房, Cbo发药药房), mblnSetPara))
    Lng自动打印 = Val(zldatabase.GetPara("自动打印", glngSys, 1342, 0, Array(Me.lbl发药清单, Me.cbo发药清单), mblnSetPara))
    int查询发药天数 = Val(zldatabase.GetPara("查询发药天数", glngSys, 1342, 7, Array(txtTimeArea_Send), mblnSetPara))
    int查询退药天数 = Val(zldatabase.GetPara("查询退药天数", glngSys, 1342, 3, Array(txtTimeArea_Sended), mblnSetPara))
    lng最大记录数 = Val(zldatabase.GetPara("查询明细记录数", glngSys, 1342, 3000, Array(txtMaxRecordCount), mblnSetPara))
    int汇总销账 = Val(zldatabase.GetPara("发药时汇总退药销帐记录", glngSys, 1342, 0))
    int检查存储库房 = Val(zldatabase.GetPara("发药时检查存储库房", glngSys, 1342, 0))
    int检查销帐申请 = Val(zldatabase.GetPara("发药时检查销帐申请数据", glngSys, 1342, 0))
    
    mstrSourceDep = zldatabase.GetPara("来源科室", glngSys, 1342, "", Array(Lvw来源科室), mblnSetPara)
    mLng打印退药清单 = Val(zldatabase.GetPara("打印退药清单", glngSys, 1342, 0, Array(lbl退药清单, Me.cbo退药清单), mblnSetPara))
    
    '根据参数值设置
    If lng药房ID <> 0 Then                                  '定位药房
        '不存在该药房则提示
        For IntLocate = 0 To Me.Cbo发药药房.ListCount - 1
            If Me.Cbo发药药房.ItemData(IntLocate) = lng药房ID Then
                Me.Cbo发药药房.ListIndex = IntLocate
                Exit For
            End If
        Next
        If IntLocate > (Cbo发药药房.ListCount - 1) Then
            MsgBox "请重新设置药房（原来设置的药房已失效）！", vbInformation, gstrSysName
            If Cbo发药药房.ListCount >= 1 Then Cbo发药药房.ListIndex = 0
        End If
    End If
    Me.Cbo操作模式.ListIndex = Lng操作模式
    Me.cbo发药清单.ListIndex = Lng自动打印
    Me.Chk按科室汇总显示.Value = Lng汇总显示
    Me.cbo退药清单.ListIndex = mLng打印退药清单
    Me.cboName.ListIndex = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品部门发药管理", "药品名称显示方式", 0))
    Me.chkDetailPage.Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品部门发药管理", "保持上一次窗体关闭时的页签", 0))
    Me.chk检查存储库房.Value = int检查存储库房
    Me.chk检查销帐申请.Value = int检查销帐申请
    
    If int汇总销账 = 1 Then
        Me.Chk按科室汇总显示.Value = 1
        Me.Chk按科室汇总显示.Enabled = False
    End If

    For n = 0 To Cbo记帐人.ListCount - 1
        If Cbo记帐人.List(n) = str记帐人 Then
            Cbo记帐人.ListIndex = n
            Exit For
        End If
    Next
    
    If str毒理分类 <> "" Then
        For n = 1 To lvw毒理分类.ListItems.count
            If InStr("," & str毒理分类 & ",", "," & lvw毒理分类.ListItems(n).Text & ",") > 0 Then
                lvw毒理分类.ListItems(n).Checked = True
            End If
        Next
    End If
    
    If str价值分类 <> "" Then
        For n = 1 To lvw价值分类.ListItems.count
            If InStr("," & str价值分类 & ",", "," & lvw价值分类.ListItems(n).Text & ",") > 0 Then
                lvw价值分类.ListItems(n).Checked = True
            End If
        Next
    End If
    
    If str高危分类 <> "" Then
        For n = 1 To lvw高危分类.ListItems.count
            If InStr("," & str高危分类 & ",", "," & n & ",") > 0 Then
                lvw高危分类.ListItems(n).Checked = True
            End If
        Next
    End If
    
    If str高危发放 <> "" Then
        If InStr(1, str高危发放, "1") Then chk高危(0).Value = 1
        If InStr(1, str高危发放, "2") Then chk高危(1).Value = 1
        If InStr(1, str高危发放, "3") Then chk高危(2).Value = 1
    End If
    
    If int查询发药天数 <= 0 Or int查询发药天数 > 99 Then
        int查询发药天数 = 7
    End If
    txtTimeArea_Send.Text = int查询发药天数
        
    If int查询退药天数 <= 0 Or int查询退药天数 > 99 Then
        int查询退药天数 = 3
    End If
    txtTimeArea_Sended.Text = int查询退药天数
    
    If lng最大记录数 <= 0 Then
        lng最大记录数 = 3000
    End If
    txtMaxRecordCount.Text = lng最大记录数
    
End Function

Private Sub Lvw药品剂型_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With Lvw药品剂型
        For n = 1 To .ListItems.count
            .ListItems(n).Selected = False
        Next
        Item.Selected = True
        If Item.Text = "所有药品剂型" Then
            If Item.Checked Then
                blnAllChecked = True
            End If
                
            For n = 1 To .ListItems.count
                .ListItems(n).Checked = blnAllChecked
            Next
        Else
            If Item.Checked = False Then
                .ListItems(1).Checked = False
            End If
        End If
    End With
End Sub


Private Sub tabShow_Click(PreviousTab As Integer)
    Select Case tabShow.Tab
    Case 0
        If Cbo发药药房.Enabled = True Then Cbo发药药房.SetFocus
    End Select
End Sub

Private Sub cmd打印设置_Click()
    Dim strBill As String
    
    Select Case cbo票据设置.ListIndex
    Case 0
        '汇总发药单
        strBill = "ZL1_BILL_1342"
    Case 1
        '退药清单
        strBill = "ZL1_BILL_1342_1"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub
Private Sub txtMaxRecordCount_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMaxRecordCount_Validate(Cancel As Boolean)
    If Val(txtMaxRecordCount.Text) <= 0 Then
        txtMaxRecordCount.Text = 3000
    End If
End Sub


Private Sub txtTimeArea_Send_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTimeArea_Send_Validate(Cancel As Boolean)
    If Val(txtTimeArea_Send.Text) <= 0 Then
        txtTimeArea_Send.Text = 7
    End If
End Sub


Private Sub txtTimeArea_Sended_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTimeArea_Sended_Validate(Cancel As Boolean)
    If Val(txtTimeArea_Sended.Text) <= 0 Then
        txtTimeArea_Sended.Text = 3
    End If
End Sub


Private Sub SetSourceDep()
    Dim rs As New ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select distinct A.编码 || '-' || A.名称 科室, A.Id " & _
            " From 部门表 A,部门性质说明 B" & _
            " Where A.Id =B.部门id and B.工作性质 in ('检查','检验','治疗','手术','营养', '临床','护理') And B.服务对象 In (2,3)  And " & _
            " (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By A.编码 || '-' || A.名称"

    Call SQLTest(App.Title, Me.Caption, gstrSQL)
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, "SetSourceDep")
    Call SQLTest

    With rs
        If .EOF Then
            MsgBox "没有设置该类部门！（部门管理）", vbInformation, gstrSysName
            Exit Sub
        End If
        Lvw来源科室.ListItems.Clear
        Do While Not .EOF
            Lvw来源科室.ListItems.Add , "_" & !Id, !科室, 1, 1
            If mstrSourceDep <> "" Then
                If InStr("," & mstrSourceDep & ",", "," & CStr(!Id) & ",") > 0 Then
                    Lvw来源科室.ListItems("_" & !Id).Checked = True
                End If
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



