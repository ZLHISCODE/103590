VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendWavePrintSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "体温单设置"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   Icon            =   "frmTendWavePrintSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPrint 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   90
      ScaleHeight     =   3405
      ScaleWidth      =   6885
      TabIndex        =   2
      Top             =   450
      Width           =   6885
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2010
         Left            =   3960
         ScaleHeight     =   463.459
         ScaleMode       =   0  'User
         ScaleWidth      =   460
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   795
         Width           =   1995
         Begin VB.PictureBox picPaper 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   1485
            Left            =   405
            ScaleHeight     =   1455
            ScaleMode       =   0  'User
            ScaleWidth      =   1140
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   270
            Width           =   1170
         End
         Begin VB.PictureBox picShadow 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1485
            Left            =   450
            ScaleHeight     =   1485
            ScaleWidth      =   1170
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   315
            Width           =   1170
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "纸张"
         Height          =   1065
         Left            =   120
         TabIndex        =   5
         Top             =   675
         Width           =   3825
         Begin VB.TextBox txtHeight 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   2415
            MaxLength       =   3
            TabIndex        =   13
            Top             =   630
            Width           =   480
         End
         Begin VB.TextBox txtWidth 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   720
            MaxLength       =   3
            TabIndex        =   9
            Top             =   630
            Width           =   480
         End
         Begin VB.ComboBox cboPage 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   225
            Width           =   2955
         End
         Begin MSComCtl2.UpDown UDHeight 
            Height          =   285
            Left            =   2895
            TabIndex        =   14
            Top             =   630
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtHeight"
            BuddyDispid     =   196614
            OrigLeft        =   2985
            OrigTop         =   630
            OrigRight       =   3225
            OrigBottom      =   930
            Max             =   460
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UDWidth 
            Height          =   285
            Left            =   1200
            TabIndex        =   10
            Top             =   630
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtWidth"
            BuddyDispid     =   196615
            OrigLeft        =   1200
            OrigTop         =   645
            OrigRight       =   1440
            OrigBottom      =   945
            Max             =   460
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            Height          =   180
            Left            =   3210
            TabIndex        =   15
            Top             =   690
            Width           =   180
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            Height          =   180
            Left            =   1515
            TabIndex        =   11
            Top             =   690
            Width           =   180
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "高度"
            Height          =   180
            Left            =   2010
            TabIndex        =   12
            Top             =   690
            Width           =   360
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "宽度"
            Height          =   180
            Left            =   300
            TabIndex        =   8
            Top             =   690
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "大小"
            Height          =   180
            Left            =   285
            TabIndex        =   6
            Top             =   300
            Width           =   360
         End
      End
      Begin VB.Frame fraOrient 
         Caption         =   "纸向"
         Height          =   1065
         Left            =   2520
         TabIndex        =   29
         Top             =   1755
         Width           =   1425
         Begin VB.OptionButton opt横向 
            Caption         =   "横向"
            Height          =   285
            Left            =   675
            TabIndex        =   31
            Top             =   600
            Width           =   660
         End
         Begin VB.OptionButton opt纵向 
            Caption         =   "纵向"
            Height          =   285
            Left            =   675
            TabIndex        =   30
            Top             =   315
            Value           =   -1  'True
            Width           =   660
         End
         Begin VB.Image img横向 
            Height          =   480
            Left            =   120
            Picture         =   "frmTendWavePrintSet.frx":1CCA
            Top             =   330
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image img纵向 
            Height          =   480
            Left            =   120
            Picture         =   "frmTendWavePrintSet.frx":2594
            Top             =   330
            Width           =   480
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "边距(mm)"
         Height          =   1065
         Left            =   120
         TabIndex        =   16
         Top             =   1755
         Width           =   2385
         Begin VB.TextBox txt左 
            Height          =   300
            Left            =   360
            MaxLength       =   3
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "25"
            Top             =   615
            Width           =   525
         End
         Begin VB.TextBox txt上 
            Height          =   300
            Left            =   360
            MaxLength       =   3
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "25"
            Top             =   270
            Width           =   540
         End
         Begin VB.TextBox txt下 
            Height          =   300
            Left            =   1455
            MaxLength       =   3
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "25"
            Top             =   270
            Width           =   540
         End
         Begin VB.TextBox txt右 
            Height          =   300
            Left            =   1455
            MaxLength       =   3
            TabIndex        =   27
            TabStop         =   0   'False
            Text            =   "25"
            Top             =   600
            Width           =   540
         End
         Begin MSComCtl2.UpDown UD下 
            Height          =   315
            Left            =   2010
            TabIndex        =   22
            Top             =   270
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Value           =   25
            BuddyControl    =   "txt下"
            BuddyDispid     =   196632
            OrigLeft        =   3750
            OrigTop         =   255
            OrigRight       =   3990
            OrigBottom      =   525
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UD上 
            Height          =   315
            Left            =   915
            TabIndex        =   19
            Top             =   270
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Value           =   25
            BuddyControl    =   "txt上"
            BuddyDispid     =   196631
            OrigLeft        =   2385
            OrigTop         =   240
            OrigRight       =   2625
            OrigBottom      =   540
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UD左 
            Height          =   315
            Left            =   915
            TabIndex        =   25
            Top             =   615
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Value           =   25
            BuddyControl    =   "txt左"
            BuddyDispid     =   196630
            OrigLeft        =   1080
            OrigTop         =   240
            OrigRight       =   1320
            OrigBottom      =   540
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UD右 
            Height          =   300
            Left            =   2010
            TabIndex        =   28
            Top             =   600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Value           =   25
            BuddyControl    =   "txt右"
            BuddyDispid     =   196633
            OrigLeft        =   1080
            OrigTop         =   240
            OrigRight       =   1320
            OrigBottom      =   540
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "左"
            Height          =   180
            Left            =   150
            TabIndex        =   23
            Top             =   675
            Width           =   180
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "上"
            Height          =   180
            Left            =   150
            TabIndex        =   17
            Top             =   330
            Width           =   180
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "下"
            Height          =   180
            Left            =   1245
            TabIndex        =   20
            Top             =   330
            Width           =   180
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "右"
            Height          =   180
            Left            =   1245
            TabIndex        =   26
            Top             =   660
            Width           =   180
         End
      End
      Begin VB.ComboBox cboPrinter 
         Height          =   300
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   315
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmTendWavePrintSet.frx":2E5E
         Top             =   75
         Width           =   480
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "体温单：打印机设置"
         Height          =   180
         Left            =   720
         TabIndex        =   3
         Top             =   315
         Width           =   1620
      End
      Begin VB.Label lblPaperHint 
         AutoSize        =   -1  'True
         Caption         =   "注意:  如果实际打印机和当前打印机不符，可能导致纸张设置失效！"
         Height          =   180
         Left            =   135
         TabIndex        =   35
         Top             =   2985
         Width           =   5490
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4005
      Index           =   0
      Left            =   210
      ScaleHeight     =   4005
      ScaleWidth      =   6645
      TabIndex        =   0
      Top             =   525
      Width           =   6645
      Begin XtremeSuiteControls.TabControl tbcStyle 
         Height          =   3930
         Left            =   600
         TabIndex        =   1
         Top             =   210
         Width           =   5460
         _Version        =   589884
         _ExtentX        =   9631
         _ExtentY        =   6932
         _StockProps     =   64
      End
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   7590
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   47
      Top             =   4500
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendWavePrintSet.frx":3728
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9551
            Text            =   "可以根据医院实际情况，设置调整体温单打印和页眉页脚。"
            TextSave        =   "可以根据医院实际情况，设置调整体温单打印和页眉页脚。"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "编辑"
            TextSave        =   "编辑"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picFoot 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   435
      ScaleHeight     =   3405
      ScaleWidth      =   6885
      TabIndex        =   36
      Top             =   435
      Width           =   6885
      Begin VB.ComboBox cboFont 
         Height          =   300
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1560
         Width           =   1905
      End
      Begin VB.ComboBox cboFSize 
         Height          =   300
         Left            =   3810
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1560
         Width           =   750
      End
      Begin VB.CheckBox chkB 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4590
         Picture         =   "frmTendWavePrintSet.frx":3FBC
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "粗体(Alt+B)"
         Top             =   1530
         Width           =   345
      End
      Begin VB.CheckBox chkU 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4920
         Picture         =   "frmTendWavePrintSet.frx":A80E
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "下划线(Alt+U)"
         Top             =   1530
         Width           =   345
      End
      Begin VB.CheckBox chkI 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5250
         Picture         =   "frmTendWavePrintSet.frx":11060
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "斜体(Alt+I)"
         Top             =   1530
         Width           =   345
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "插图(&I)"
         Height          =   350
         Left            =   150
         TabIndex        =   38
         Top             =   1530
         Width           =   1200
      End
      Begin VB.CommandButton cmd同步 
         Caption         =   "同步(&G)"
         Height          =   350
         Left            =   5730
         TabIndex        =   45
         ToolTipText     =   "所有护理文件的页眉页脚与当前文件的页眉页脚格式一致"
         Top             =   1530
         Width           =   1100
      End
      Begin RichTextLib.RichTextBox rtbHead 
         Height          =   1425
         Left            =   30
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   30
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   2514
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmTendWavePrintSet.frx":178B2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtbFoot 
         Height          =   1425
         Left            =   30
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1950
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   2514
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmTendWavePrintSet.frx":1794F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl字体 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "字体"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1515
         TabIndex        =   39
         Top             =   1620
         Width           =   360
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmTendWavePrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'页眉页脚相关
'######################################################################################################
Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
'矩形
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'包含用于格式化指定设备的相关信息
Private Type FORMATRANGE
    hDC As Long             '渲染设备
    hdcTarget As Long       '目标设备
    rc As RECT              '渲染区域，单位：缇。
    rcPage As RECT          '渲染设备的整体区域，单位：缇。
    chrg As CHARRANGE       '用于格式化的文本范围。
End Type

Private Type PageInfo
    PageNumber As Long      '页码
    Start As Long           '字符起始位置
    End As Long             '字符终止位置
    ActualHeight As Long    '本页实际打印高度
End Type
Private AllPages() As PageInfo   '页信息
Private Const WM_PASTE = &H302&              '粘贴
Private Const WM_USER = &H400                '通常用 WM_USER + X 来自定义消息
Private Const EM_FORMATRANGE = (WM_USER + 57)    '为某一设备格式化指定范围的文本。
Private Const EM_SETTARGETDEVICE = (WM_USER + 72) '设置用于所见即所得的目标设备和行宽。
Private Const EM_HIDESELECTION = (WM_USER + 63)  '显示/隐藏文本。
Private Const PHYSICALOFFSETX = 112  '对于打印设备而言，表示从物理页的左边缘到可打印区域的左边缘的距离，采用设备单位。
Private Const PHYSICALOFFSETY = 113  '对于打印设备而言，表示从物理页的上边缘到可打印区域的上边缘的距离，采用设备单位。
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long '获取中英文混合字符串长度
'######################################################################################################

Private mdblW As Double  '左边不可打印比例
Private mdblH As Double  '上边不可打印比例

'打印参数变量
Private mintPage As Integer '纸张
Private mlngWidth As Long '自定义纸张宽度,Twip
Private mlngHeight As Long '自定义纸张高度'Twip
Private mintOrient As Integer   '纸向
Private mlngLeft As Long '左边距'mm
Private mlngRight As Long '右边距'mm
Private mlngTop As Long '上边距'mm
Private mlngBottom As Long '下边距'mm
Private mblnRTBFoot As Boolean
'事件控制
Private mblnChange As Boolean  '控制打印设置
Private mblnChanged As Boolean '记录数据是否发生变化
Private rtbThis As Object
Public mbytMode As Byte
Public mlngFileID As Long  '病历文件列表的ID

'--修改说明：50182,刘鹏飞,2012-08-24,新增体温单设置页眉页脚功能

Private Property Let DataChanged(vData As Boolean)
    
    mblnChanged = vData
    If mblnChanged Then
        stbThis.Panels(3).Enabled = True
    Else
        stbThis.Panels(3).Enabled = False
    End If
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnChanged
End Property

Public Function ShowMe(ByVal frmParent As Object, ByVal lngFileID As Long) As Boolean
    mlngFileID = lngFileID
    gblnOK = False
    DataChanged = False
    If frmParent Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmParent
    End If
    ShowMe = gblnOK
End Function

Private Sub cboPage_Click()
    Dim blnOK As Boolean
    Dim dblRight As Double
    Dim dblDown As Double
    
    '纸张
    Select Case cboPage.ItemData(cboPage.ListIndex)
    Case 256
        '强行设置自定义纸张可用,不检查
        mintPage = 256
    Case Else
        Printer.PaperSize = cboPage.ItemData(cboPage.ListIndex)
        mintPage = Printer.PaperSize
    End Select
        
    opt纵向.Enabled = True
    opt横向.Enabled = True
    Err = 0
    On Error Resume Next
    opt横向.Tag = Printer.Orientation
    Printer.Orientation = 1
    If Printer.Orientation <> 1 Then opt纵向.Enabled = False
    Printer.Orientation = 2
    If Printer.Orientation <> 2 Then opt横向.Enabled = False
    
    If opt横向.Enabled = False Then
        opt纵向.Value = True
        img纵向.Visible = True
        img横向.Visible = False
    End If
    If Printer.Orientation <> mintOrient Then Printer.Orientation = mintOrient
    mintOrient = Printer.Orientation
    '最后实际设置纸张大小(纸向影响之后)
    Select Case mintPage
    Case 256
        '自定义纸张认为全部可以打印
        mdblW = 0
        mdblH = 0
        
'        If cboPage.Text = "B5, 182 x 257 毫米" Then
'            mlngWidth = 182 * conRatemmToTwip
'            mlngHeight = 257 * conRatemmToTwip
'        End If
        If Val(opt横向.Tag) <> mintOrient Then
            Call SetCustonPager(Me.hwnd, mlngWidth, mlngHeight)
            mlngWidth = Printer.Width
            mlngHeight = Printer.Height
        End If
        
        txtWidth.Enabled = True
        txtHeight.Enabled = True
        UDWidth.Enabled = True
        UDHeight.Enabled = True
    Case Else
        '取该打印机支持该幅面的真实尺寸
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
        
        '不可打印区域比例
        mdblW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
        mdblH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
        
        txtWidth.Enabled = False
        txtHeight.Enabled = False
        UDWidth.Enabled = False
        UDHeight.Enabled = False
    
    End Select
        
    '显示纸张尺寸
    mblnChange = False
    txtWidth.Tag = mlngWidth
    txtWidth.Text = CLng(mlngWidth / conRatemmToTwip)
    txtHeight.Tag = mlngHeight
    txtHeight.Text = CLng(mlngHeight / conRatemmToTwip)
    mblnChange = True
    
    '显示可用边距
    '最小在可打印区域之内
    '最大不超过宽高的1/4
'    If cboPage.Text = "B5, 182 x 257 毫米" Then
'        UD左.Min = 0
'        UD左.Max = 5
'    Else
    UD左.Min = mlngWidth / conRatemmToTwip * mdblW
    UD左.Max = mlngWidth / conRatemmToTwip / 4
'    End If
    UD右.Min = UD左.Min
    UD右.Max = UD左.Max
    
    UD上.Min = mlngHeight / conRatemmToTwip * mdblH
    UD上.Max = mlngHeight / conRatemmToTwip / 4
    UD下.Min = UD上.Min
    UD下.Max = UD上.Max
    
    If mlngLeft >= UD左.Min And mlngLeft <= UD左.Max Then
        UD左.Value = mlngLeft
    Else
        UD左.Value = UD左.Min
    End If
    If mlngRight >= UD右.Min And mlngRight <= UD右.Max Then
        UD右.Value = mlngRight
    Else
        UD右.Value = UD右.Min
    End If
    If mlngTop >= UD上.Min And mlngTop <= UD上.Max Then
        UD上.Value = mlngTop
    Else
        UD上.Value = UD上.Min
    End If
    If mlngBottom >= UD下.Min And mlngBottom <= UD下.Max Then
        UD下.Value = mlngBottom
    Else
        UD下.Value = UD下.Min
    End If
    
    mlngLeft = UD左.Value
    mlngRight = UD右.Value
    mlngTop = UD上.Value
    mlngBottom = UD下.Value
    
    '显示纸向
    mblnChange = False
    If mintOrient = 1 Then
        opt纵向.Value = True: opt纵向_Click
    Else
        opt横向.Value = True: opt横向_Click
    End If
    mblnChange = True
    
    '显示预览纸张
    Call ShowPaper
    
    DataChanged = True
End Sub

Private Sub LoadPage()
    Dim i As Integer
    Dim strPrinter As String
    
    '初始打印机列表
    strPrinter = GetSetting("ZLSOFT", "公共模块\zl9PrintMode\Default", "DeviceName", Printer.DeviceName)
    With cboPrinter
        .Clear
        For i = 0 To Printers.Count - 1
            .AddItem Printers(i).DeviceName
            .ItemData(.ListCount - 1) = i '打印机索引
            
            '读取存储的打印机为当前打印机,并初始化可用页面
            If strPrinter = Printers(i).DeviceName Then .ListIndex = .NewIndex
        Next
        
        '缺省初始化为当前打印机
        If .ListIndex = -1 Then
            For i = 0 To .ListCount - 1
                '读取系统当前的打印机为当前打印机,并初始化可用页面
                If .List(i) = Printer.DeviceName Then .ListIndex = i: Exit For
            Next
        End If
        .Visible = False
        .Enabled = False
    End With
End Sub

Private Sub cboPrinter_Click()
    
    Dim i As Integer, j As Integer
    Dim lngCount As Long, strTmp As String
    Dim strPaperSize As String * 300
    Dim strPrinter As String
    
    Set Printer = Printers(cboPrinter.ItemData(cboPrinter.ListIndex))
    mintPage = Printer.PaperSize
     '如果支持,则保持原有纸张
     If mintPage <> 256 Then
         On Error Resume Next
         Printer.PaperSize = mintPage
         On Error GoTo 0
         mintPage = Printer.PaperSize
         mintOrient = Printer.Orientation
     End If
     
     '特殊处理，对于体温单只支持A4及B5大小的纸张
     cboPage.Clear
     '------------------------------------------------------------------------------------------
     '纸张大小
     lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERS, strPaperSize, 0)
     For i = 1 To lngCount
         j = Asc(Mid(strPaperSize, i * 2, 1)) * 256# + Asc(Mid(strPaperSize, i * 2 - 1, 1))
         
         If mbytMode = 1 Then
             If j = 9 Or j = 13 Then
                 cboPage.AddItem GetPaperName(j)
                 cboPage.ItemData(cboPage.ListCount - 1) = j
                 If j = mintPage Then cboPage.ListIndex = cboPage.NewIndex
             End If
         Else
             If j >= 1 And j <= 41 Then '只列出标准支持的纸张
                 cboPage.AddItem GetPaperName(j)
                 cboPage.ItemData(cboPage.ListCount - 1) = j
                 If j = mintPage Then cboPage.ListIndex = cboPage.NewIndex
             End If
         End If
         
     Next
    
     '------------------------------------------------------------------------------------------
     '自定义纸张处理
     i = 256
     cboPage.AddItem GetPaperName(i)
     cboPage.ItemData(cboPage.ListCount - 1) = i
     If mintPage = 256 Then cboPage.ListIndex = cboPage.NewIndex
     If cboPage.ListIndex = -1 And cboPage.ListCount > 0 Then cboPage.ListIndex = 0
End Sub


Private Function SaveData() As Boolean
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strPaper As String
    Dim blnTrans As Boolean
    
    On Error GoTo errHand
    
    If Not IsNumeric(txtWidth.Text) Then
        MsgBox "请确定报表的纸张宽度！", vbExclamation, App.Title
        txtWidth.SetFocus: Exit Function
    End If
    If CInt(txtWidth.Text) > UDWidth.Max Then
        MsgBox "报表的纸张宽度不能超过" & UDWidth.Max & "毫米！", vbExclamation, App.Title
        txtWidth.SetFocus: Exit Function
    End If
    
    If Not IsNumeric(txtHeight.Text) Then
        MsgBox "请确定报表的纸张高度！", vbExclamation, App.Title
        txtWidth.SetFocus: Exit Function
    End If
    If CInt(txtHeight.Text) > UDHeight.Max Then
        MsgBox "报表的纸张高度不能超过" & UDHeight.Max & "毫米！", vbExclamation, App.Title
        txtHeight.SetFocus: Exit Function
    End If
    
    If Not PageHeadTest Then Exit Function
    
    strSQL = "Select 编号,名称,报表,页眉,页脚 From 病历页面格式 Where 种类 = 3 And 编号 = (Select 页面 From 病历文件列表 Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取文件打印设置", mlngFileID)
    If rsTemp.EOF Then
        MsgBox "病人体温表不存在病历页面格式，无法进行打印设置。请检查！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '自定义纸张始终纵向保存高度和宽度
    If mintPage = 256 Then
        Call SetCustonPager(Me.hwnd, mlngWidth, mlngHeight)
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
    Else
        Printer.PaperSize = mintPage
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
    End If
    
    strPaper = mintPage & ";" & mintOrient & ";" & mlngHeight & ";" & mlngWidth & ";" & CLng(Me.ScaleY(mlngLeft, vbMillimeters, vbTwips)) & ";" & _
        CLng(Me.ScaleY(mlngRight, vbMillimeters, vbTwips)) & ";" & CLng(Me.ScaleY(mlngTop, vbMillimeters, vbTwips)) & ";" & _
        CLng(Me.ScaleY(mlngBottom, vbMillimeters, vbTwips))
    '保存打印数据
    strSQL = "Zl_病历页面格式_Update(3" & ",'"
    '种类_In In 病历页面格式.种类%Type,
    '编号_In In 病历页面格式.编号%Type,
    strSQL = strSQL & NVL(rsTemp!编号) & "','"
    '名称_In In 病历页面格式.名称%Type,
    strSQL = strSQL & NVL(rsTemp!名称) & "','"
    '报表_In In 病历页面格式.报表%Type,
    strSQL = strSQL & NVL(rsTemp!报表) & "','"
    '格式_In In 病历页面格式.格式%Type,
    strSQL = strSQL & strPaper & "','"
    '页眉_In In 病历页面格式.页眉%Type,
    strSQL = strSQL & NVL(rsTemp!页眉) & "','"
    '页脚_In In 病历页面格式.页脚%Type
    strSQL = strSQL & NVL(rsTemp!页脚) & "')"
    
    
    gcnOracle.BeginTrans
    blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSQL, "Zl_病历页面格式_Update")
    If Not SavePageHead(picFoot.Tag) Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    If Not SavePageFoot(picFoot.Tag) Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    gcnOracle.CommitTrans
    blnTrans = False
    
    gblnOK = True
    SaveData = True
    cmd同步.Enabled = True
    Exit Function
errHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_SaveExit
        If SaveData Then
            DataChanged = False
            Unload Me
        End If
        
    Case conMenu_Edit_Transf_Save
        
        If SaveData Then
           DataChanged = False
        End If
        
    Case conMenu_Edit_Transf_Cancle
                
        Call zlRefreshData
        DataChanged = False
    Case conMenu_File_Exit
        
        gblnOK = False
        Unload Me
        
    Case conMenu_Help_Help
        
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '客户区域的大小

    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    
    With picPane(0)
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Top = lngTop
        .Height = lngBottom - lngTop
    End With
    
    tbcStyle.Move 15, 15, picPane(0).Width - 30, picPane(0).Height - 30
    
    rtbHead.Width = picFoot.Width - 60
    rtbFoot.Width = rtbHead.Width
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_SaveExit
        
        Control.Enabled = DataChanged
        
    Case conMenu_Edit_Transf_Save
        
        Control.Enabled = DataChanged
        
    Case conMenu_Edit_Transf_Cancle
                
        Control.Enabled = DataChanged
        
    End Select
End Sub

Private Sub cmdOpen_Click()
    Dim picTemp As StdPicture
    
    With Me.dlgThis
        .DialogTitle = "标志图选择"
        .Filename = ""
        .Filter = "图像(*.jpg;*.bmp;*.ico;*.gif)|*.jpg;*.bmp;*.ico;*.gif"
        .CancelError = False
        On Error Resume Next
        .ShowOpen
        If Err.Number <> 0 Then
            Err.Clear
            Exit Sub
        End If
    End With
    Set picTemp = Nothing
    Set picTemp = LoadPicture(Me.dlgThis.Filename)
    If picTemp Is Nothing Then MsgBox "不是有效的图片文件！", vbExclamation, Me.Caption: Exit Sub
    
    Clipboard.Clear
    Clipboard.SetData picTemp
    
    Call GetrtbObject
    SendMessageLong rtbThis.hwnd, WM_PASTE, 0, 0
    DataChanged = True
End Sub

Private Sub cboFont_Click()
    Call GetrtbObject
    If rtbThis.SelFontName <> cboFont.List(cboFont.ListIndex) Then
        rtbThis.SelFontName = cboFont.List(cboFont.ListIndex)
        DataChanged = True
    End If
End Sub

Private Sub cboFSize_Click()
    Dim sngNum As Single
    Call GetrtbObject
    sngNum = GetFontSizeNumber(cboFSize.List(cboFSize.ListIndex))
    If rtbThis.SelFontSize <> sngNum Then
        rtbThis.SelFontSize = sngNum
        DataChanged = True
    End If
End Sub

Private Sub chkB_Click()
    Call GetrtbObject
    If chkB.Value = vbChecked Then
        rtbThis.SelBold = True
    Else
        rtbThis.SelBold = False
    End If
    DataChanged = True
End Sub

Private Sub chkI_Click()
    Call GetrtbObject
    If chkI.Value = vbChecked Then
        rtbThis.SelItalic = True
    Else
        rtbThis.SelItalic = False
    End If
    DataChanged = True
End Sub

Private Sub chkU_Click()
    Call GetrtbObject
    If chkU.Value = vbChecked Then
        rtbThis.SelUnderline = True
    Else
        rtbThis.SelUnderline = False
    End If
    DataChanged = True
End Sub

Private Sub cmd同步_Click()
    Dim strZIPHead As String, strZIPFoot As String
    Dim rsTemp As New ADODB.Recordset
    Dim blnTrans As Boolean
    On Error GoTo errHand
    '将当前格式应用到所有护理文件
    
    gstrSQL = " Select 种类||'-'||编号 AS KEY From 病历文件列表 Where 种类=3 and ID<>[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理文件", mlngFileID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "当前只有一份护理文件，不需要执行同步功能！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("请再次确认：" & vbCrLf & "        执行该功能后，所有护理文件的页眉页脚格式将统一与当前文件设置保存一致！", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '获取当前设置的页眉页脚
    strZIPHead = ReadPageHeadFile(picFoot.Tag)
    strZIPFoot = ReadPageFootFile(picFoot.Tag)
    
    gcnOracle.BeginTrans
    blnTrans = True
    '循环写入数据库
    With rsTemp
        Do While Not .EOF
            If Not SavePageHead(!Key, strZIPHead) Then
                gcnOracle.RollbackTrans
                Exit Sub
            End If
            If Not SavePageFoot(!Key, strZIPFoot) Then
                gcnOracle.RollbackTrans
                Exit Sub
            End If
            .MoveNext
        Loop
    End With
    gcnOracle.CommitTrans
    blnTrans = False
    
    '删除临时文件
    gobjFSO.DeleteFile strZIPHead, True
    gobjFSO.DeleteFile strZIPFoot, True
    
    MsgBox "同步成功！", vbInformation, gstrSysName
    Exit Sub
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errHand
    
    Me.picPrint.BackColor = Me.BackColor
    Me.picFoot.BackColor = Me.BackColor
    
    If Not ExistsPrinter Then
        MsgBox "系统中没有安装任何打印机,请先安装打印机！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    Call RestoreWinState(Me, App.ProductName)

    With Me.tbcStyle
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameBorder
        End With
        .InsertItem 0, "打印设置", Me.picPrint.hwnd, 0
        .InsertItem 1, "页面格式", Me.picFoot.hwnd, 0
        .Item(0).Selected = True
    End With
    Call InitMenuBar  '加载菜单
    If Not zlRefreshData Then Unload Me
    DataChanged = False
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function zlRefreshData()
    Dim i As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, strPaper As String
    Dim blnHead As Boolean, blnFoot As Boolean
    
    On Error Resume Next
    Printer.Orientation = 1
    mintOrient = 1
    mintPage = 256
    mlngLeft = 20: mlngRight = 20: mlngTop = 20: mlngBottom = 20
    Err.Clear: On Error GoTo errHand
    '刷新数据信息
    gblnOK = False
    mblnChange = True
    Call LoadPage
    Call PrepareFont
    mblnChange = False
    '从病历页面格式中提取打印设置数据
    strSQL = "Select  种类||'-'||编号 AS KEY,格式 From 病历页面格式 Where 种类 = 3 And 编号 = (Select 页面 From 病历文件列表 Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取文件打印设置", mlngFileID)
    If Not rsTemp.EOF Then
        strPaper = "" & rsTemp!格式
        blnHead = ReadPageHead(rtbHead, rsTemp!Key)
        blnFoot = ReadPageFoot(rtbFoot, rsTemp!Key)
        cmd同步.Enabled = blnHead Or blnFoot
        picFoot.Tag = rsTemp!Key
    End If
    
    If UBound(Split(strPaper, ";")) >= 4 Then mlngLeft = Round(Me.ScaleY(Val(Split(strPaper, ";")(4)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 5 Then mlngRight = Round(Me.ScaleY(Val(Split(strPaper, ";")(5)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 6 Then mlngTop = Round(Me.ScaleX(Val(Split(strPaper, ";")(6)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 7 Then mlngBottom = Round(Me.ScaleX(Val(Split(strPaper, ";")(7)), vbTwips, vbMillimeters), 2)

    If UBound(Split(strPaper, ";")) >= 0 Then
        For i = 0 To Me.cboPage.ListCount - 1
            If Me.cboPage.ItemData(i) = Val(Split(strPaper, ";")(0)) Then Me.cboPage.ListIndex = i: Exit For
        Next
        If cboPage.ListIndex = -1 And cboPage.ListCount > 0 Then cboPage.ListIndex = cboPage.ListCount - 1
        mblnChange = False
        If Me.cboPage.ListIndex >= 0 Then
            mintPage = cboPage.ItemData(i)
            If UBound(Split(strPaper, ";")) >= 2 Then mlngHeight = Val(Split(strPaper, ";")(2))
            If UBound(Split(strPaper, ";")) >= 3 Then mlngWidth = Val(Split(strPaper, ";")(3))
            Me.txtHeight.Text = CLng(mlngHeight / conRatemmToTwip)
            Me.txtWidth.Text = CLng(mlngWidth / conRatemmToTwip)
        End If
    End If
    
    If UBound(Split(strPaper, ";")) >= 1 Then
        mintOrient = Val(Split(strPaper, ";")(1))
        If Val(Split(strPaper, ";")(1)) = 2 Then
            Me.opt横向.Value = True
        Else
            Me.opt纵向.Value = True
        End If
    End If
        
    txt左.Text = mlngLeft
    txt右.Text = mlngRight
    txt上.Text = mlngTop
    txt下.Text = mlngBottom
    
    On Error Resume Next
    If mintOrient = Printer.Orientation And mintPage = 256 Then
        If mintOrient = 1 Then
            Printer.Orientation = 2
        Else
            Printer.Orientation = 1
        End If
    End If
    Err.Clear: On Error GoTo errHand
    Call cboPage_Click: mblnChange = True
    
    zlRefreshData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function PrepareFont()
    Dim sFont As String, i As Integer
    
    For i = 0 To Screen.FontCount - 1
       sFont = Screen.Fonts(i)
       cboFont.AddItem sFont
       If sFont = "宋体" Then cboFont.ListIndex = i
    Next i
    With cboFSize
        .AddItem "初号"
        .AddItem "小初"
        .AddItem "一号"
        .AddItem "小一"
        .AddItem "二号"
        .AddItem "小二"
        .AddItem "三号"
        .AddItem "小三"
        .AddItem "四号"
        .AddItem "小四"
        .AddItem "五号"
        .AddItem "小五"
        .AddItem "六号"
        .AddItem "小六"
        .AddItem "七号"
        .AddItem "八号"
        .AddItem 5
        .AddItem 5.5
        .AddItem 6.5
        .AddItem 7.5
        .AddItem 8
        .AddItem 9
        .AddItem 10
        .AddItem 10.5
        .AddItem 11
        .AddItem 12
        .AddItem 14
        .AddItem 16
        .AddItem 18
        .AddItem 20
        .AddItem 22
        .AddItem 24
        .AddItem 26
        .AddItem 28
        .AddItem 36
        .AddItem 48
        .AddItem 72
        .ListIndex = 10
    End With
End Function

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim rs As ADODB.Recordset
    Dim objExtendedBar As CommandBar
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.ActiveMenuBar.Visible = False
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
    
     '快键绑定
    With cbsThis.KeyBindings

        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F2, conMenu_Edit_Transf_Save
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义
    Set cbrToolBar = cbsThis.Add("标准", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SaveExit, "保存并退出"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "保存"): cbrControl.ToolTipText = "保存已更改的数据(Ctrl+S,F2)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "恢复"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "恢复到上次保存时的数据状态"
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "帮助(F1)"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): cbrControl.ToolTipText = "退出当前的设计窗体(Esc)"

    End With
        
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.STYLE = xtpButtonIconAndCaption
        End If
    Next
    
     '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_ESCAPE, conMenu_File_Exit
        
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save
    End With
    
    InitMenuBar = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    
    If DataChanged Then
        Cancel = (MsgBox("更改后的设计必须保存后才生效，是否放弃保存？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
    
    If Cancel Then Exit Sub
    
    DataChanged = False
    
    Call SaveWinState(Me, App.ProductName)
    
    Set rtbThis = Nothing
End Sub

Private Sub opt横向_Click()
    Dim lngL As Long, lngR As Long
    Dim lngT As Long, lngB As Long
    
    If opt横向.Value Then
        img纵向.Visible = False
        img横向.Visible = True
        
        If mintOrient = 1 Then
            lngL = mlngLeft
            lngR = mlngRight
            lngT = mlngTop
            lngB = mlngBottom
            
            mlngLeft = lngB
            mlngRight = lngT
            mlngTop = lngL
            mlngBottom = lngR
            If mintPage = 256 Then
                Call SetCustonPager(Me.hwnd, mlngWidth, mlngHeight)
                mlngWidth = Printer.Width
                mlngHeight = Printer.Height
            End If
        End If
        
        mintOrient = 2
        
        If mblnChange Then Call cboPage_Click
        
        DataChanged = True
    End If
End Sub

Private Sub opt纵向_Click()
    Dim lngL As Long, lngR As Long
    Dim lngT As Long, lngB As Long
    
    If opt纵向.Value Then
        img纵向.Visible = True
        img横向.Visible = False
        
        If mintOrient = 2 Then
            lngL = mlngLeft
            lngR = mlngRight
            lngT = mlngTop
            lngB = mlngBottom
              
            mlngLeft = lngT
            mlngRight = lngB
            mlngTop = lngR
            mlngBottom = lngL
            
            If mintPage = 256 Then
                Call SetCustonPager(Me.hwnd, mlngWidth, mlngHeight)
                mlngWidth = Printer.Width
                mlngHeight = Printer.Height
            End If
        End If
        
        mintOrient = 1
        
        If mblnChange Then Call cboPage_Click
        
        DataChanged = True
    End If
End Sub

Private Sub rtbFoot_Change()
    DataChanged = True
End Sub

Private Sub rtbFoot_GotFocus()
    mblnRTBFoot = True
End Sub

Private Sub rtbHead_Change()
    DataChanged = True
End Sub

Private Sub rtbHead_GotFocus()
    mblnRTBFoot = False
End Sub

Private Sub tbcStyle_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim intPage As Integer
    If Item.Caption = "页面格式" Then
        On Error Resume Next
        intPage = cboPage.ItemData(cboPage.ListIndex)
        Printer.PaperSize = intPage
        Printer.Orientation = IIf(opt纵向.Value, 1, 2)
        If intPage = 256 Then
            If Printer.Orientation = 1 Then
                mlngWidth = CLng(Val(txtWidth.Text) * conRatemmToTwip)
                mlngHeight = CLng(Val(txtHeight.Text) * conRatemmToTwip)
            Else
                mlngHeight = CLng(Val(txtWidth.Text) * conRatemmToTwip)
                mlngWidth = CLng(Val(txtHeight.Text) * conRatemmToTwip)
            End If
            Call SetCustonPager(Me.hwnd, mlngWidth, mlngHeight)
            mlngWidth = Printer.Width
            mlngHeight = Printer.Height
        Else
            mlngWidth = Printer.Width
            mlngHeight = Printer.Height
        End If
        Call SendMessage(rtbHead.hwnd, EM_SETTARGETDEVICE, Me.hDC, ByVal CLng(Printer.ScaleWidth))
        SendMessageLong rtbHead.hwnd, EM_HIDESELECTION, 0, 0
        Call SendMessage(rtbFoot.hwnd, EM_SETTARGETDEVICE, Me.hDC, ByVal CLng(Printer.ScaleWidth))
        SendMessageLong rtbFoot.hwnd, EM_HIDESELECTION, 0, 0
    
        rtbHead.Width = picFoot.Width - 60
        rtbFoot.Width = rtbHead.Width
        rtbHead.SetFocus
    End If
End Sub

Private Sub txtHeight_Change()
    If Not mblnChange Then Exit Sub
    If IsNumeric(txtHeight.Text) Then
        txtHeight.Tag = CLng(txtHeight.Text * conRatemmToTwip)
        mlngHeight = CLng(txtHeight.Text * conRatemmToTwip)
        
        If mintPage = 256 Then cboPage.ListIndex = cboPage.ListCount - 1
    End If
    Call ShowPaper
    
    DataChanged = True
End Sub

Private Sub txtWidth_Change()
    If Not mblnChange Then Exit Sub
    If IsNumeric(txtWidth.Text) Then
        txtWidth.Tag = CLng(txtWidth.Text * conRatemmToTwip)
        mlngWidth = CLng(txtWidth.Text * conRatemmToTwip)
        
        If mintPage = 256 Then cboPage.ListIndex = cboPage.ListCount - 1
    End If
    Call ShowPaper
    
    DataChanged = True
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: VBA.Beep
End Sub

Private Sub txtHeight_GotFocus()
    txtHeight.SelStart = 0: txtHeight.SelLength = Len(txtHeight.Text)
End Sub

Private Sub txtWidth_GotFocus()
    txtWidth.SelStart = 0: txtWidth.SelLength = Len(txtWidth.Text)
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: VBA.Beep
End Sub

Private Sub txt上_Change()
    DataChanged = True
End Sub

Private Sub txt上_GotFocus()
    zlControl.TxtSelAll txt上
End Sub

Private Sub txt上_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt上_Validate(Cancel As Boolean)
    If Not mblnChange Then Exit Sub
    If IsNumeric(txt上.Text) Then
        If txt上.Text >= UD上.Min And txt上.Text <= UD上.Max Then
            UD上.Value = txt上.Text
        Else
            UD上.Value = UD上.Min
        End If
    End If
End Sub

Private Sub txt下_Change()
    DataChanged = True
End Sub

Private Sub txt下_Validate(Cancel As Boolean)
    If Not mblnChange Then Exit Sub
    If IsNumeric(txt下.Text) Then
        If txt下.Text >= UD下.Min And txt下.Text <= UD下.Max Then
            UD下.Value = txt下.Text
        Else
            UD下.Value = UD下.Min
        End If
    End If
End Sub

Private Sub txt下_GotFocus()
    zlControl.TxtSelAll txt下
End Sub

Private Sub txt下_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt右_Change()
    DataChanged = True
End Sub

Private Sub txt右_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt右_Validate(Cancel As Boolean)
    If Not mblnChange Then Exit Sub
    If IsNumeric(txt右.Text) Then
        If txt右.Text >= UD右.Min And txt右.Text <= UD右.Max Then
            UD右.Value = txt右.Text
        Else
            UD右.Value = UD右.Min
        End If
    End If
End Sub

Private Sub txt右_GotFocus()
    zlControl.TxtSelAll txt右
End Sub

Private Sub txt左_Change()
    DataChanged = True
End Sub

Private Sub txt左_Validate(Cancel As Boolean)
    If Not mblnChange Then Exit Sub
    If IsNumeric(txt左.Text) Then
        If txt左.Text >= UD左.Min And txt左.Text <= UD左.Max Then
            UD左.Value = txt左.Text
        Else
            UD左.Value = UD左.Min
        End If
    End If
End Sub

Private Sub txt左_GotFocus()
    zlControl.TxtSelAll txt左
End Sub

Private Sub txt左_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub UD上_Change()
    mlngTop = UD上.Value
    Call ShowPaper
End Sub

Private Sub UD下_Change()
    mlngBottom = UD下.Value
    Call ShowPaper
End Sub

Private Sub UD右_Change()
    mlngRight = UD右.Value
    Call ShowPaper
End Sub

Private Sub UD左_Change()
    mlngLeft = UD左.Value
    Call ShowPaper
End Sub

Private Sub ShowPaper()
'功能：显示设置的纸张的预览
    On Error Resume Next
    
    picPaper.Cls
    
    picPaper.Width = mlngWidth / conRatemmToTwip
    picPaper.Height = mlngHeight / conRatemmToTwip
    picPaper.Left = (picBack.ScaleWidth - picPaper.Width) / 2
    picPaper.Top = (picBack.ScaleHeight - picPaper.Height) / 2
    
    picShadow.Width = picPaper.Width
    picShadow.Height = picPaper.Height
    
    picShadow.Left = picPaper.Left + 5
    picShadow.Top = picPaper.Top + 5
    
    picPaper.ScaleWidth = mlngWidth
    picPaper.ScaleHeight = mlngHeight
    
    picPaper.Line (0, mlngTop * conRatemmToTwip)-(picPaper.ScaleWidth, mlngTop * conRatemmToTwip), &H808080
    picPaper.Line (0, picPaper.ScaleHeight - (mlngBottom + 2) * conRatemmToTwip)-(picPaper.ScaleWidth, picPaper.ScaleHeight - (mlngBottom + 2) * conRatemmToTwip), &H808080
    
    picPaper.Line (mlngLeft * conRatemmToTwip, 0)-(mlngLeft * conRatemmToTwip, picPaper.ScaleHeight), &H808080
    picPaper.Line (picPaper.ScaleWidth - (mlngRight + 2) * conRatemmToTwip, 0)-(picPaper.ScaleWidth - (mlngRight + 2) * conRatemmToTwip, picPaper.ScaleHeight), &H808080
    
    Me.Refresh
End Sub

Private Sub GetrtbObject()
    If mblnRTBFoot Then
        Set rtbThis = rtbFoot
    Else
        Set rtbThis = rtbHead
    End If
End Sub


Private Function ReadPageHead(objHead As RichTextBox, ByVal strKey As String) As Boolean
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(12, strKey, App.Path & "\Head_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Head_S.RTF")
        objHead.LoadFile strFile, rtfRTF           '读取文件
        gobjFSO.DeleteFile strFile, True      '删除临时文件
        ReadPageHead = True
    Else
        objHead.Text = ""
    End If
End Function

Private Function ReadPageFoot(objFoot As RichTextBox, ByVal strKey As String) As Boolean
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(13, strKey, App.Path & "\Foot_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Foot_S.RTF")
        objFoot.LoadFile strFile, rtfRTF           '读取文件
        gobjFSO.DeleteFile strFile, True      '删除临时文件
        ReadPageFoot = True
    Else
        objFoot.Text = ""
    End If
End Function

Private Function ReadPageHeadFile(ByVal strKey As String) As String
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strZip As String
    strZip = zlBlobRead(12, strKey, App.Path & "\Head_L.zip")
    If gobjFSO.FileExists(strZip) Then
        ReadPageHeadFile = strZip
    End If
End Function

Private Function ReadPageFootFile(ByVal strKey As String) As String
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strZip As String
    strZip = zlBlobRead(13, strKey, App.Path & "\Foot_L.zip")
    If gobjFSO.FileExists(strZip) Then
        ReadPageFootFile = strZip
    End If
End Function

'################################################################################################################
'## 功能：  在压缩文件相同目录释放产生解压文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
Private Function UnzipTendPage(ByVal strZipFile As String, ByVal strTarFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String
    Dim mclsUnzip As New cUnzip
    
    On Error GoTo errHand
    
    If Not gobjFSO.FileExists(strZipFile) Then UnzipTendPage = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    strZipPath = gobjFSO.GetSpecialFolder(2)
    strZipPathTmp = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer)
    Call gobjFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp ' & "\TMP.RTF"
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    If gobjFSO.FolderExists(strZipFileTmp) Then
        
        strZipFileName = gobjFSO.GetFile(strZipFileTmp & "\" & strTarFile)
        Call gobjFSO.CopyFile(strZipFileName, "C:\" & strTarFile)
        
        On Error Resume Next
        gobjFSO.DeleteFolder strZipPathTmp, True
        gobjFSO.DeleteFile strZipFile, True
        
        UnzipTendPage = "C:\" & strTarFile
    Else
        UnzipTendPage = ""
    End If
    
    Exit Function
    
errHand:
    Call SaveErrLog
End Function

Private Function SavePageHead(ByVal strKey As String, Optional ByVal strZipFile As String = "") As Boolean
    'blnBuild=False:产生文件并压缩;True:已产生压缩文件
    Dim strFile As String, strZip As String
    If strZipFile = "" Then
        strFile = App.Path & "\Head_S.rtf"
        If gobjFSO.FileExists(strFile) = True Then gobjFSO.DeleteFile strFile, True
        rtbHead.SaveFile strFile
        strZip = zlFileZip(strFile)
    Else
        strZip = strZipFile
    End If
    SavePageHead = zlBlobSave(12, strKey, strZip)
    If strZipFile = "" Then
        gobjFSO.DeleteFile strFile, True
        gobjFSO.DeleteFile strZip, True
    End If
End Function

Private Function SavePageFoot(ByVal strKey As String, Optional ByVal strZipFile As String = "") As Boolean
    'blnBuild=False:产生文件并压缩;True:已产生压缩文件
    Dim strFile As String, strZip As String
    If strZipFile = "" Then
        strFile = App.Path & "\Foot_S.rtf"
        If gobjFSO.FileExists(strFile) = True Then gobjFSO.DeleteFile strFile, True
        rtbFoot.SaveFile strFile
        strZip = zlFileZip(strFile)
    Else
        strZip = strZipFile
    End If
    SavePageFoot = zlBlobSave(13, strKey, strZip)
    If strZipFile = "" Then
        gobjFSO.DeleteFile strFile, True
        gobjFSO.DeleteFile strZip, True
    End If
End Function

Private Function PageHeadTest() As Boolean
    '超过上边距返回假
    Dim fr As FORMATRANGE           '格式化的文本范围
    Dim rcDrawTo As RECT            '目标文字区域
    Dim rcPage As RECT              '目标页面区域
    Dim gTargetDC As Long
    Dim lngOffsetLeft As Long
    Dim lngOffsetTop As Long
'    Dim lngOffsetWidth As Long
'    Dim lngOffsetHeight As Long
    Dim lngNextPos As Long, lngLen As Long, lngTMP As Long, lngPageCount As Long
    
    lngLen = lstrlen(rtbHead.Text)
    'printer.Duplex = vbPRDPHorizontal
    'printer.ScaleMode = vbTwips
    lngOffsetLeft = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    lngOffsetTop = Printer.ScaleY(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)
'    lngOffsetWidth = Printer.ScaleWidth
'    lngOffsetHeight = Printer.ScaleHeight
    
    gTargetDC = hDC
    With rcPage
        .Left = 0
        .Top = 0
        .Right = Printer.Width
        .Bottom = Printer.Height
    End With
    With rcDrawTo
        .Left = lngOffsetLeft
        .Top = lngOffsetTop
        .Right = Printer.Width - lngOffsetLeft
        .Bottom = Printer.ScaleX(txt上.Text, vbMillimeters, vbTwips)
    End With
    With fr
        .hDC = Printer.hDC
        .hdcTarget = gTargetDC
        .rc = rcDrawTo
        .rcPage = rcPage
        .chrg.cpMin = 0
        .chrg.cpMax = -1
    End With
    
    Do
        lngNextPos = SendMessage(rtbHead.hwnd, EM_FORMATRANGE, 0, fr)
        
        lngPageCount = lngPageCount + 1             ' 页数＋1
        '记录分页信息
        ReDim Preserve AllPages(1 To lngPageCount) As PageInfo
        AllPages(lngPageCount).PageNumber = lngPageCount
        AllPages(lngPageCount).ActualHeight = fr.rc.Bottom - fr.rc.Top          '实际打印高度
        AllPages(lngPageCount).Start = lngTMP
        AllPages(lngPageCount).End = lngNextPos
        
        fr.chrg.cpMin = lngNextPos
        If lngNextPos <= lngTMP Or lngNextPos >= lngLen Then Exit Do      ' 完成所有页面的分页
        lngTMP = lngNextPos
    Loop
    Call SendMessage(rtbHead.hwnd, EM_FORMATRANGE, 0, ByVal CLng(0))
    
    If fr.rc.Bottom > rcDrawTo.Bottom Or lngPageCount > 1 Then
        MsgBox "设计的页眉内容超过上边距！", vbInformation, gstrSysName
        Exit Function
    End If
    PageHeadTest = True
End Function
