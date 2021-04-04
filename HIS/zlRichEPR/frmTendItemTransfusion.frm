VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTendItemTransfusion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "移动护士站基础设置"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   Icon            =   "frmTendItemTransfusion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   150
      Top             =   420
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5475
      Left            =   90
      TabIndex        =   47
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9657
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "基础设置"
      TabPicture(0)   =   "frmTendItemTransfusion.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl液体名称"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cbo液体名称"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cbo液体量"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lvw病区"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "体温薄定义"
      TabPicture(1)   =   "frmTendItemTransfusion.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstColumnUsed"
      Tab(1).Control(1)=   "cmdColumn(1)"
      Tab(1).Control(2)=   "cmdColumn(0)"
      Tab(1).Control(3)=   "lstColumnItems"
      Tab(1).Control(4)=   "cmdMove(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdMove(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblColumnItems(1)"
      Tab(1).Control(7)=   "lblColumnItems(0)"
      Tab(1).Control(8)=   "Label1(2)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "项目分类"
      TabPicture(2)   =   "frmTendItemTransfusion.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmd修改"
      Tab(2).Control(1)=   "cmd删除"
      Tab(2).Control(2)=   "cmd增加"
      Tab(2).Control(3)=   "txt分类名"
      Tab(2).Control(4)=   "lstClass"
      Tab(2).Control(5)=   "lstItems"
      Tab(2).Control(6)=   "Label3"
      Tab(2).Control(7)=   "Label1(3)"
      Tab(2).Control(8)=   "lblColumnItems(3)"
      Tab(2).Control(9)=   "lblColumnItems(2)"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "公告栏设置"
      TabPicture(3)   =   "frmTendItemTransfusion.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pic病区"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmd模板"
      Tab(3).Control(2)=   "Frame1"
      Tab(3).Control(3)=   "picDraw"
      Tab(3).Control(4)=   "SSTab1"
      Tab(3).ControlCount=   5
      Begin MSComctlLib.ListView lvw病区 
         Height          =   2085
         Left            =   450
         TabIndex        =   54
         Top             =   3240
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   3678
         View            =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Left            =   240
         TabIndex        =   53
         Top             =   2610
         Width           =   6525
      End
      Begin VB.PictureBox pic病区 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -74550
         ScaleHeight     =   315
         ScaleWidth      =   4035
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   390
         Width           =   4035
         Begin VB.ComboBox cbo病区 
            Height          =   300
            Left            =   465
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   0
            Width           =   3555
         End
         Begin VB.Label lbl病区 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "病区"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   30
            TabIndex        =   26
            Top             =   60
            Width           =   360
         End
      End
      Begin VB.CommandButton cmd模板 
         Caption         =   "依模板产生当前病区公告"
         Height          =   315
         Left            =   -70380
         TabIndex        =   28
         Top             =   390
         Width           =   2325
      End
      Begin VB.Frame Frame1 
         Caption         =   "要素设置"
         Height          =   4635
         Left            =   -70380
         TabIndex        =   29
         Top             =   780
         Width           =   2325
         Begin VB.CommandButton cmd公告栏_保存 
            Caption         =   "新增"
            Height          =   350
            Left            =   180
            TabIndex        =   43
            Top             =   4170
            Width           =   945
         End
         Begin VB.CommandButton cmd公告栏_删除 
            Caption         =   "删除"
            Height          =   350
            Left            =   1230
            TabIndex        =   44
            Top             =   4170
            Width           =   945
         End
         Begin VB.TextBox txt绑定诊疗项目 
            Appearance      =   0  'Flat
            Height          =   1275
            Left            =   180
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   42
            Top             =   2820
            Width           =   2025
         End
         Begin VB.CommandButton cmd绑定诊疗项目 
            Caption         =   "绑定诊疗项目"
            Height          =   345
            Left            =   180
            TabIndex        =   41
            Top             =   2460
            Width           =   2025
         End
         Begin VB.CheckBox chk无数据时隐藏 
            Caption         =   "无数据时隐藏该项"
            Height          =   225
            Left            =   180
            TabIndex        =   40
            Top             =   2160
            Width           =   1905
         End
         Begin VB.ComboBox cbo位置 
            Height          =   300
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1770
            Width           =   1575
         End
         Begin VB.TextBox txt行号 
            Height          =   300
            Left            =   600
            TabIndex        =   37
            Top             =   1410
            Width           =   1575
         End
         Begin VB.ComboBox cbo名称 
            Height          =   300
            Left            =   600
            TabIndex        =   33
            Text            =   "Combo1"
            Top             =   690
            Width           =   1575
         End
         Begin VB.TextBox txt别名 
            Height          =   300
            Left            =   600
            TabIndex        =   35
            Top             =   1050
            Width           =   1575
         End
         Begin VB.ComboBox cbo分类 
            Height          =   300
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   330
            Width           =   1575
         End
         Begin VB.Label lbl位置 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "位置"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   195
            TabIndex        =   38
            Top             =   1830
            Width           =   360
         End
         Begin VB.Label lbl行号 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "行号"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   195
            TabIndex        =   36
            Top             =   1470
            Width           =   360
         End
         Begin VB.Label lbl别名 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "别名"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   195
            TabIndex        =   34
            Top             =   1110
            Width           =   360
         End
         Begin VB.Label lbl名称 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "名称"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   195
            TabIndex        =   32
            Top             =   750
            Width           =   360
         End
         Begin VB.Label lbl分类 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "分类"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   195
            TabIndex        =   30
            Top             =   390
            Width           =   360
         End
      End
      Begin VB.PictureBox picDraw 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   -74610
         Picture         =   "frmTendItemTransfusion.frx":007C
         ScaleHeight     =   4665
         ScaleWidth      =   4065
         TabIndex        =   25
         Top             =   720
         Width           =   4095
         Begin VB.Shape Shape 
            BorderColor     =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   7
            Left            =   -30
            Top             =   120
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   6
            Left            =   -30
            Top             =   270
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   5
            Left            =   330
            Top             =   270
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   4
            Left            =   720
            Top             =   270
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   3
            Left            =   720
            Top             =   120
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   2
            Left            =   720
            Top             =   -30
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   1
            Left            =   330
            Top             =   -30
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   105
            Index           =   0
            Left            =   -30
            Top             =   -30
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Label lbl要素内容 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "要素内容"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   0
            Left            =   840
            TabIndex        =   50
            Top             =   45
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label lbl要素名 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "要素名"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   60
            TabIndex        =   49
            Top             =   60
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5100
         Left            =   -74970
         TabIndex        =   24
         Top             =   330
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   8996
         _Version        =   393216
         TabOrientation  =   2
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "病区概况"
         TabPicture(0)   =   "frmTendItemTransfusion.frx":46526
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "工作备忘"
         TabPicture(1)   =   "frmTendItemTransfusion.frx":46542
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "注意事项"
         TabPicture(2)   =   "frmTendItemTransfusion.frx":4655E
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
      End
      Begin VB.CommandButton cmd修改 
         Caption         =   "修改"
         Height          =   315
         Left            =   -72930
         TabIndex        =   48
         Top             =   4410
         Width           =   585
      End
      Begin VB.CommandButton cmd删除 
         Caption         =   "删除"
         Height          =   350
         Left            =   -73410
         TabIndex        =   21
         Top             =   4800
         Width           =   1065
      End
      Begin VB.CommandButton cmd增加 
         Caption         =   "增加"
         Height          =   350
         Left            =   -74460
         TabIndex        =   20
         Top             =   4800
         Width           =   1065
      End
      Begin VB.TextBox txt分类名 
         Height          =   315
         Left            =   -74040
         TabIndex        =   19
         Top             =   4410
         Width           =   1095
      End
      Begin VB.ListBox lstClass 
         Height          =   3120
         Left            =   -74430
         TabIndex        =   17
         Top             =   1215
         Width           =   2100
      End
      Begin VB.ListBox lstItems 
         Height          =   4020
         Left            =   -71850
         MultiSelect     =   2  'Extended
         TabIndex        =   23
         Top             =   1200
         Width           =   2100
      End
      Begin VB.ListBox lstColumnUsed 
         Height          =   4020
         Left            =   -71070
         TabIndex        =   10
         Top             =   1200
         Width           =   2100
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "删除(&E)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   -72180
         TabIndex        =   12
         Top             =   2445
         Width           =   975
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "选用(&S)"
         Height          =   300
         Index           =   0
         Left            =   -72180
         TabIndex        =   11
         Top             =   2145
         Width           =   975
      End
      Begin VB.ListBox lstColumnItems 
         Height          =   4020
         Left            =   -74430
         TabIndex        =   9
         Top             =   1215
         Width           =   2100
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "上移(&U)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   -72180
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3270
         Width           =   975
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "下移(&D)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   -72180
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3570
         Width           =   975
      End
      Begin VB.ComboBox cbo液体量 
         Height          =   300
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2010
         Width           =   2535
      End
      Begin VB.ComboBox cbo液体名称 
         Height          =   300
         Left            =   630
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2040
         Width           =   2715
      End
      Begin VB.Label Label1 
         Caption         =   "    医院如使用了新版护士工作站，请勾选使用了新版护士站的病区"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   52
         Top             =   3000
         Width           =   6135
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "分类"
         Height          =   180
         Left            =   -74415
         TabIndex        =   18
         Top             =   4470
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "    根据医院内部查体温时所使用的体温薄进行项目设置。"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   -74790
         TabIndex        =   15
         Top             =   630
         Width           =   6135
      End
      Begin VB.Label lblColumnItems 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "护理项目分类"
         Height          =   180
         Index           =   3
         Left            =   -74400
         TabIndex        =   16
         Top             =   990
         Width           =   2040
      End
      Begin VB.Label lblColumnItems 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "护理项目列表"
         Height          =   180
         Index           =   2
         Left            =   -71820
         TabIndex        =   22
         Top             =   990
         Width           =   2070
      End
      Begin VB.Label lblColumnItems 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "体温薄项目"
         Height          =   180
         Index           =   1
         Left            =   -71040
         TabIndex        =   8
         Top             =   990
         Width           =   2070
      End
      Begin VB.Label lblColumnItems 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "可选护理记录项目"
         Height          =   180
         Index           =   0
         Left            =   -74400
         TabIndex        =   7
         Top             =   990
         Width           =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "    根据医院内部查体温时所使用的体温薄进行项目设置。"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   -74790
         TabIndex        =   6
         Top             =   630
         Width           =   6135
      End
      Begin VB.Label Label1 
         Caption         =   "    移动护士工作站在执行输液类医嘱的时候，如果医生下达了记出入量医嘱，程序会自动产生入量草稿数据。"
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   0
         Top             =   600
         Width           =   6135
      End
      Begin VB.Label lbl液体名称 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1)请设置液体名称关联的护理项目"
         ForeColor       =   &H00808080&
         Height          =   180
         Left            =   630
         TabIndex        =   2
         Top             =   1740
         Width           =   2700
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "2)请设置液体量关联的汇总项目"
         ForeColor       =   &H00808080&
         Height          =   180
         Left            =   3720
         TabIndex        =   4
         Top             =   1740
         Width           =   2520
      End
      Begin VB.Label Label1 
         Caption         =   "    只设置液体量表时产生的一条汇总后的入量草稿数据；如果同时设置了液体名称，则产生明细的入量草稿数据。"
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   1
         Left            =   330
         TabIndex        =   1
         Top             =   1080
         Width           =   6135
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5430
      TabIndex        =   46
      Top             =   5700
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4140
      TabIndex        =   45
      Top             =   5700
      Width           =   1100
   End
End
Attribute VB_Name = "frmTendItemTransfusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mblnALLOW As Boolean        '是否有编辑的权限
Private mblnStart As Boolean
Private mblnClick As Boolean
Private mblnEdit As Boolean         '是否进行了编辑
Private Enum emnuPage
    病区概况
    工作备忘
End Enum
Private mrsBoard As New ADODB.Recordset

Private Sub cbo病区_Click()
    Frame1.Enabled = (cbo病区.ListCount > 0) And InStr(1, mstrPrivs, "编辑") <> 0
    Timer1.Enabled = True
End Sub

Private Sub cbo名称_Change()
    If mblnClick Then Exit Sub
    cmd公告栏_保存.Caption = "新增"
    Call SetShape
    Frame1.Tag = ""
End Sub

Private Sub cbo液体量_Click()
    If Not mblnStart Then Exit Sub
    mblnEdit = True
End Sub

Private Sub cbo液体名称_Click()
    If Not mblnStart Then Exit Sub
    mblnEdit = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim blnClear As Boolean, blnTrans As Boolean
    Dim intRow As Integer, intCount As Integer
    Dim sqlText As String
    
    On Error GoTo errHand
    
    If cbo液体名称.ItemData(cbo液体名称.ListIndex) > 0 And cbo液体量.ItemData(cbo液体量.ListIndex) = 0 Then
        MsgBox "请设置液体量对应的护理记录项目！", vbInformation, gstrSysName
        cbo液体量.SetFocus
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    blnTrans = True
    
    '保存入量项目
    blnClear = True
    If cbo液体名称.ItemData(cbo液体名称.ListIndex) > 0 Then
        gstrSQL = "ZL_护理记录项目_Transfusion(" & cbo液体名称.ItemData(cbo液体名称.ListIndex) & ",'11'," & IIf(blnClear, "1", "0") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存液体名称")
        blnClear = False
    End If
    If cbo液体量.ItemData(cbo液体量.ListIndex) > 0 Then
        gstrSQL = "ZL_护理记录项目_Transfusion(" & cbo液体量.ItemData(cbo液体量.ListIndex) & ",'12'," & IIf(blnClear, "1", "0") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存液体量")
        blnClear = False
    End If
    
    '保存体温薄定义
    intCount = lstColumnUsed.ListCount
    For intRow = 1 To intCount
        sqlText = sqlText & "," & Me.lstColumnUsed.ItemData(intRow - 1)
    Next
    sqlText = Mid(sqlText, 2)
    Call zlDatabase.SetPara("体温薄项目", sqlText, 100)
    
    '保存新版试点病区
    sqlText = ""
    intCount = lvw病区.ListItems.Count
    For intRow = 1 To intCount
        If lvw病区.ListItems(intRow).Checked Then
            sqlText = sqlText & "," & Mid(lvw病区.ListItems(intRow).Key, 2)
        End If
    Next
    If sqlText <> "" Then
        sqlText = Mid(sqlText, 2)
        Call zlDatabase.SetPara("移动护士站新版病区列表", sqlText, 100)
    End If
    
    gcnOracle.CommitTrans
    blnTrans = False
    mblnEdit = False
    
    Unload Me
    Exit Sub
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd添加分类_Click()
    
End Sub

Private Sub cmd绑定诊疗项目_Click()
    Dim strIDs As String, strNames As String
    Call frmClinicSelect.ShowMe(Me, strIDs, strNames)
    txt绑定诊疗项目.Tag = strIDs
    txt绑定诊疗项目.Text = strNames
End Sub

Private Sub cmd公告栏_保存_Click()
    Dim lngID As Long
    Dim intPos As Integer
    Dim intCount As Integer
    Dim blnTrans As Boolean
    Dim strIDs As String, strItems As String
    
    If Trim(txt别名.Text) = "" Then
        MsgBox "别名不能为空！", vbInformation, gstrSysName
        txt别名.SetFocus
        Exit Sub
    End If
    If Trim(txt行号.Text) = "" Then
        MsgBox "行号不能为空！", vbInformation, gstrSysName
        txt行号.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txt行号.Text) Then
        MsgBox "行号中不能含有非法字符！", vbInformation, gstrSysName
        txt行号.SetFocus
        Exit Sub
    End If
    If Val(txt行号.Text) < 0 Or Val(txt行号.Text) > 13 Then
        MsgBox "行号不能小于零或大于13！", vbInformation, gstrSysName
        txt行号.SetFocus
        Exit Sub
    End If
    
    Me.Caption = "移动护士站基础设置" & "(正在保存数据,请稍候......)"
    gcnOracle.BeginTrans
    blnTrans = True
    
    If txt绑定诊疗项目.Tag <> "" Then
        strIDs = txt绑定诊疗项目.Tag
    End If
    lngID = Val(Frame1.Tag)
    If lngID = 0 Then lngID = zlDatabase.GetNextId("公告栏样式")
    gstrSQL = "ZL_公告栏样式_APPENDITEM(" & lngID & "," & Me.cbo病区.ItemData(Me.cbo病区.ListIndex) & "," & Me.cbo分类.ListIndex + 1 & "," & _
        "'" & Me.cbo名称.Text & "','" & Me.txt别名.Text & "'," & Me.txt行号.Text & "," & Me.cbo位置.ListIndex + 1 & "," & _
        IIf(Me.cbo名称.ListIndex = -1, 0, 1) & "," & chk无数据时隐藏.Value & ")"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '保存格式:<ITEMLIST><ITEM><XH/><MC/></ITEM></ITEMLIST>
    intCount = 0
    Do While strIDs <> ""
        If Len(strIDs) > 3800 Then
            '向左搜寻逗号
            intPos = GetSplit(Mid(strIDs, 1, 3800))
            strItems = Mid(strIDs, 1, intPos)
            strIDs = Mid(strIDs, intPos + 1)
        Else
            strItems = strIDs
            strIDs = ""
        End If
        
        gstrSQL = "ZL_公告栏样式_UPDATEZLXM(" & lngID & ",'" & strItems & "'," & IIf(intCount = 0, "1", "0") & ")"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        intCount = intCount + 1
    Loop
    Me.Caption = "移动护士站基础设置"
    gcnOracle.CommitTrans
    blnTrans = False
    
    Timer1.Enabled = True
    Exit Sub
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Me.Caption = "移动护士站基础设置"
End Sub

Private Function GetSplit(ByVal strInput As String) As Integer
    Dim intPos As Integer
    '向左搜寻逗号,返回逗号的位置
    
    intPos = 3800
    Do While True
        If Mid(strInput, intPos, 1) = "," Then
            intPos = intPos - 1
            GetSplit = intPos
            Exit Function
        End If
        intPos = intPos - 1
    Loop
End Function

Private Sub cmd公告栏_删除_Click()
    On Error GoTo errHand
    
    If Val(Frame1.Tag) = 0 Then Exit Sub
    If MsgBox("你确定要删除要素：" & cbo名称.Text & "？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    gcnOracle.Execute "ZL_公告栏样式_DELETEITEM(" & Val(Frame1.Tag) & ")", , adCmdStoredProc
    
    Timer1.Enabled = True
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd模板_Click()
    If MsgBox("将删除当前病区现有内容后依据公告栏模板产生，你确定吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call zlDatabase.ExecuteProcedure("ZL_公告栏样式_BUILD(" & Me.cbo病区.ItemData(Me.cbo病区.ListIndex) & ")", "模板产生当前病区公告")
    Timer1.Enabled = True
End Sub

Private Sub cmd删除_Click()
    Dim intIndex As Integer
    Dim strXH As String
    On Error GoTo errHand
    
    strXH = GetSelItems()
    If strXH = "" Then
        MsgBox "至少要选择一个项目！", vbInformation, gstrSysName
        lstItems.SetFocus
        Exit Sub
    End If
    
    Call zlDatabase.ExecuteProcedure("ZL_护理记录项目_MOBILE('','" & strXH & "')", "删除分类数据")
    
    intIndex = lstClass.ListIndex
    If intIndex = -1 Then Exit Sub
    lstClass.RemoveItem intIndex
    If intIndex < lstClass.ListCount Then
        lstClass.ListIndex = intIndex
    Else
        If lstClass.ListCount >= 1 Then
            lstClass.ListIndex = intIndex - 1
        End If
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd修改_Click()
    Dim strXH As String
    On Error GoTo errHand
    
    If Trim(txt分类名.Text) = "" Then
        MsgBox "分类名不能为空！", vbInformation, gstrSysName
        txt分类名.SetFocus
        Exit Sub
    End If
    strXH = GetSelItems
    Call zlDatabase.ExecuteProcedure("ZL_护理记录项目_MOBILE('" & txt分类名.Text & "','" & strXH & "')", "更新分类名")
    lstClass.List(lstClass.ListIndex) = txt分类名.Text
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd增加_Click()
    Dim strXH As String
    On Error GoTo errHand
    
    If Trim(txt分类名.Text) = "" Then
        MsgBox "分类名不能为空！", vbInformation, gstrSysName
        txt分类名.SetFocus
        Exit Sub
    End If
    strXH = GetSelItems(True)
    If strXH = "" Then
        MsgBox "至少要选择一个项目！", vbInformation, gstrSysName
        lstItems.SetFocus
        Exit Sub
    End If
    
    Call zlDatabase.ExecuteProcedure("ZL_护理记录项目_MOBILE('" & txt分类名.Text & "','" & strXH & "')", "更新分类数据")
    lstClass.AddItem txt分类名.Text
    lstClass.ListIndex = 0
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errHand
    Dim str体温薄 As String, str试点 As String
    Dim rsTemp As New ADODB.Recordset
    
    mblnClick = False
    mblnStart = False
    mblnEdit = False
    mstrPrivs = gstrPrivs
    mblnALLOW = (InStr(1, gstrPrivs, "编辑") > 0)
    
    '1)入量项目
    '提取文本型项目
    gstrSQL = "Select 项目名称,项目序号,项目类型 From 护理记录项目 Where (项目类型=1 And 项目长度>10) or 项目表示=4 Order by 项目序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取汇总项目")
    rsTemp.Filter = "项目类型=1"
    cbo液体名称.AddItem "未设置"
    Call zlControl.CboAddData(cbo液体名称, rsTemp, False)
    rsTemp.Filter = "项目类型<>1"
    cbo液体量.AddItem "未设置"
    Call zlControl.CboAddData(cbo液体量, rsTemp, False)
    rsTemp.Filter = 0
    '提取设置的数据
    gstrSQL = " Select 项目序号,操作类型 From 护理记录项目 Where 操作类型 IN ('11','12')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取汇总项目")
    
    rsTemp.Filter = "操作类型='11'"
    If rsTemp.RecordCount <> 0 Then
        Call zlControl.CboLocate(cbo液体名称, rsTemp!项目序号, True)
    Else
        cbo液体名称.ListIndex = 0
    End If
    rsTemp.Filter = "操作类型='12'"
    If rsTemp.RecordCount <> 0 Then
        Call zlControl.CboLocate(cbo液体量, rsTemp!项目序号, True)
    Else
        cbo液体量.ListIndex = 0
    End If
    rsTemp.Filter = 0
    
    '提取新版试点病区
    str试点 = "," & zlDatabase.GetPara("移动护士站新版病区列表", 100) & ","
    gstrSQL = "" & _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B" & _
            " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by A.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病区列表")
    With lvw病区
        .ListItems.Clear
        Do While Not rsTemp.EOF
            .ListItems.Add , "K" & rsTemp!ID, "[" & rsTemp!编码 & "]" & rsTemp!名称
            If InStr(1, str试点, "," & rsTemp!ID & ",") <> 0 Then .ListItems("K" & rsTemp!ID).Checked = True
            rsTemp.MoveNext
        Loop
    End With
    
    '2)体温薄
    str体温薄 = zlDatabase.GetPara("体温薄项目", 100)
    gstrSQL = " Select A.项目序号,A.项目名称 From 护理记录项目 A" & _
              " Where A.应用方式<>0 AND 项目序号 NOT IN (SELECT * FROM TABLE(F_NUM2LIST([1])))  " & _
              " Order by 项目序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str体温薄)
    With rsTemp
        Me.lstColumnItems.Clear
        Do While Not .EOF
            Me.lstColumnItems.AddItem !项目名称
            Me.lstColumnItems.ItemData(Me.lstColumnItems.NewIndex) = !项目序号
            .MoveNext
        Loop
    End With
    '提取已选择的项目清单
    gstrSQL = " Select B.项目序号,B.项目名称 From 护理记录项目 B " & _
              " Where B.项目序号 IN (SELECT * FROM TABLE(F_NUM2LIST([1])))" & _
              " Order by B.项目序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str体温薄)
    With rsTemp
        Me.lstColumnUsed.Clear
        Do While Not .EOF
            Me.lstColumnUsed.AddItem !项目名称
            Me.lstColumnUsed.ItemData(Me.lstColumnUsed.NewIndex) = !项目序号
            .MoveNext
        Loop
    End With
    cmdColumn(0).Enabled = (lstColumnItems.ListCount <> 0)
    cmdColumn(1).Enabled = (lstColumnUsed.ListCount <> 0)
    
    '3)护理项目分类
    gstrSQL = " Select distinct 移动分组 From 护理记录项目 B where 移动分组 Is Not NULL"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.lstClass.Clear
        Me.lstClass.AddItem "未设定"
        Me.lstClass.AddItem "全部"
        Do While Not .EOF
            Me.lstClass.AddItem !移动分组
            .MoveNext
        Loop
        If lstClass.ListCount > 0 Then Me.lstClass.ListIndex = 0
    End With
    
    '4)公告栏
    Call InitUnits
    With cbo名称
        .Clear
        .AddItem "病区原有人数"
        .ItemData(.NewIndex) = 1
        .AddItem "病区现有人数"
        .ItemData(.NewIndex) = 2
        .AddItem "一级护理列表"
        .ItemData(.NewIndex) = 3
        .AddItem "特级护理列表"
        .ItemData(.NewIndex) = 4
        .AddItem "病危列表"
        .ItemData(.NewIndex) = 5
        .AddItem "入院列表"
        .ItemData(.NewIndex) = 6
        .AddItem "出院列表"
        .ItemData(.NewIndex) = 7
        .AddItem "预出院列表"
        .ItemData(.NewIndex) = 8
        .AddItem "手术列表"
        .ItemData(.NewIndex) = 9
        .AddItem "预手术列表"
        .ItemData(.NewIndex) = 10
        .AddItem "转床列表"
        .ItemData(.NewIndex) = 11
    End With
    
    cbo分类.Clear
    cbo分类.AddItem "病区概况"
    cbo分类.AddItem "工作备忘"
    cbo分类.AddItem "注意事项"
    cbo分类.ListIndex = 0
    
    cbo位置.Clear
    cbo位置.AddItem "左"
    cbo位置.AddItem "右"
    cbo位置.ListIndex = 0
    
    Call SetEnabled
    mblnStart = True
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdColumn_Click(Index As Integer)
    Dim intIndex As Integer
    Dim objlst As ListBox
    If Index = 0 Then
        If Me.lstColumnItems.ListIndex < 0 Then Exit Sub
        intIndex = Me.lstColumnItems.ListIndex
        Me.lstColumnUsed.AddItem Me.lstColumnItems.Text
        Me.lstColumnUsed.ItemData(Me.lstColumnUsed.NewIndex) = Me.lstColumnItems.ItemData(Me.lstColumnItems.ListIndex)
        Me.lstColumnItems.RemoveItem Me.lstColumnItems.ListIndex
        Set objlst = lstColumnItems
    Else
        If Me.lstColumnUsed.ListIndex < 0 Then Exit Sub
        intIndex = Me.lstColumnUsed.ListIndex
        Me.lstColumnItems.AddItem Me.lstColumnUsed.Text
        Me.lstColumnItems.ItemData(Me.lstColumnItems.NewIndex) = Me.lstColumnUsed.ItemData(Me.lstColumnUsed.ListIndex)
        Me.lstColumnUsed.RemoveItem Me.lstColumnUsed.ListIndex
        Set objlst = lstColumnUsed
    End If
    If objlst.ListCount >= intIndex + 1 Then
        objlst.ListIndex = intIndex
    Else
        objlst.ListIndex = objlst.ListCount - 1
    End If
    
    cmdColumn(0).Enabled = (lstColumnItems.ListCount <> 0) And mblnALLOW
    cmdColumn(1).Enabled = (lstColumnUsed.ListCount <> 0) And mblnALLOW
    
    Call SetMoveState
    
    If Not mblnStart Then Exit Sub
    mblnEdit = True
End Sub

Private Sub cmdMove_Click(Index As Integer)
    Dim arrData
    Dim strCopy As String
    Dim lngDo As Long, lngMAX As Long
    Dim lngSelIndex As Long, lngTarIndex As Long
    
    '当前索引
    lngSelIndex = lstColumnUsed.ListIndex
    '目标索引
    lngTarIndex = lngSelIndex + IIf(Index = 0, -1, 1)
    lngMAX = lstColumnUsed.ListCount - 1
    For lngDo = 0 To lngMAX
        If lngDo = lngTarIndex Then
            strCopy = strCopy & "|" & lstColumnUsed.List(lngSelIndex) & "," & lstColumnUsed.ItemData(lngSelIndex)
        ElseIf lngDo = lngSelIndex Then
            strCopy = strCopy & "|" & lstColumnUsed.List(lngTarIndex) & "," & lstColumnUsed.ItemData(lngTarIndex)
        Else
            strCopy = strCopy & "|" & lstColumnUsed.List(lngDo) & "," & lstColumnUsed.ItemData(lngDo)
        End If
    Next
    strCopy = Mid(strCopy, 2)
    Debug.Print strCopy
    
    lstColumnUsed.Clear
    arrData = Split(strCopy, "|")
    For lngDo = 0 To lngMAX
        lstColumnUsed.AddItem Split(arrData(lngDo), ",")(0)
        lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = Val(Split(arrData(lngDo), ",")(1))
    Next
    lstColumnUsed.ListIndex = lngTarIndex
    
    Call SetMoveState
    If Not mblnStart Then Exit Sub
    mblnEdit = True
End Sub

Private Sub lbl要素名_Click(Index As Integer)
    Dim intDo As Integer, intCount As Integer
    
    mblnClick = True
    intCount = lbl要素名.Count - 1
    For intDo = 1 To intCount
        lbl要素名(intDo).BackStyle = 0
    Next
    Frame1.Tag = lbl要素名(Index).Tag
    cmd公告栏_保存.Caption = "修改"
    lbl要素名(Index).BackStyle = 1
    Call SetShape(Index)
    
    '定位该要素，显示相应的属性
    mrsBoard.Filter = "ID=" & Val(lbl要素名(Index).Tag)
    If mrsBoard.RecordCount = 0 Then Exit Sub
    
    cbo分类.ListIndex = mrsBoard!分类 - 1
    If Not zlControl.CboLocate(cbo名称, mrsBoard!名称) Then cbo名称.Text = mrsBoard!名称
    txt别名.Text = IIf(IsNull(mrsBoard!别名), "", mrsBoard!别名)
    txt行号.Text = mrsBoard!行号
    cbo位置.ListIndex = mrsBoard!位置 - 1
    chk无数据时隐藏.Value = mrsBoard!是否隐藏
    
    txt绑定诊疗项目.Text = Get诊疗项目NAME(Val(lbl要素名(Index).Tag))
    txt绑定诊疗项目.Tag = Get诊疗项目ID(Val(lbl要素名(Index).Tag))
    mblnClick = False
End Sub

Private Function Get诊疗项目ID(ByVal lngID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = "" & _
        " SELECT a.XH" & _
        " FROM 公告栏样式 p," & _
        " XMLTable('/ITEMLIST/ITEM/XH' PASSING p.诊疗项目" & _
        " COLUMNS XH VARCHAR2(256) PATH '/XH') a" & _
        " Where p.id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取项目内容", lngID)
    With rsTemp
        Do While Not rsTemp.EOF
            Get诊疗项目ID = Get诊疗项目ID & "," & rsTemp!xh
            rsTemp.MoveNext
        Loop
    End With
    
    If Get诊疗项目ID <> "" Then Get诊疗项目ID = Mid(Get诊疗项目ID, 2)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get诊疗项目NAME(ByVal lngID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = "" & _
        " SELECT a.MC" & _
        " FROM 公告栏样式 p," & _
        " XMLTable('/ITEMLIST/ITEM/MC' PASSING p.诊疗项目" & _
        " COLUMNS MC VARCHAR2(256) PATH '/MC') a" & _
        " Where p.id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取项目内容", lngID)
    With rsTemp
        Do While Not rsTemp.EOF
            Get诊疗项目NAME = Get诊疗项目NAME & "," & rsTemp!MC
            rsTemp.MoveNext
        Loop
    End With
    
    If Get诊疗项目NAME <> "" Then Get诊疗项目NAME = Mid(Get诊疗项目NAME, 2)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub lstClass_Click()
    Dim strCond As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    cmd修改.Enabled = False
    cmd删除.Enabled = False
    txt分类名.Text = lstClass.Text
    gstrSQL = " Select 项目序号,项目名称 From 护理记录项目"
    If lstClass.Text = "全部" Then
        
    ElseIf lstClass.Text = "未设定" Then
        strCond = " Where 移动分组 Is NULL"
    Else
        cmd修改.Enabled = True And mblnALLOW
        cmd删除.Enabled = True And mblnALLOW
        strCond = " Where 移动分组 =[1]"
    End If
    gstrSQL = gstrSQL & strCond & " Order by 项目序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取分组数据", lstClass.Text)
    
    With rsTemp
        lstItems.Clear
        Do While Not .EOF
            lstItems.AddItem !项目名称
            lstItems.ItemData(lstItems.NewIndex) = !项目序号
            .MoveNext
        Loop
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub lstColumnItems_DblClick()
    If lstColumnItems.ListCount = 0 Then Exit Sub
        If Not mblnALLOW Then Exit Sub
    Call cmdColumn_Click(0)
End Sub

Private Sub lstColumnUsed_Click()
    Call SetMoveState
End Sub

Private Sub lstColumnUsed_DblClick()
    If lstColumnUsed.ListCount = 0 Then Exit Sub
        If Not mblnALLOW Then Exit Sub
    Call cmdColumn_Click(1)
End Sub

Private Sub SetMoveState()
    cmdMove(0).Enabled = False
    cmdMove(1).Enabled = False
    
    If lstColumnUsed.ListIndex < 0 Then Exit Sub
    If lstColumnUsed.SelCount < 0 Then Exit Sub
    cmdMove(0).Enabled = (lstColumnUsed.ListIndex > 0) And mblnALLOW
    cmdMove(1).Enabled = (lstColumnUsed.ListIndex < lstColumnUsed.ListCount - 1) And mblnALLOW
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnEdit Then
        If MsgBox("确认要退出吗？你所做的修改还未保存！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1
    End If
End Sub

Private Function GetSelItems(Optional ByVal blnSel As Boolean = False) As String
    Dim i As Integer, j As Integer
    
    j = lstItems.ListCount
    For i = 1 To j
        If blnSel Then
            If lstItems.Selected(i - 1) Then
                GetSelItems = GetSelItems & "," & lstItems.ItemData(i - 1)
            End If
        Else
            GetSelItems = GetSelItems & "," & lstItems.ItemData(i - 1)
        End If
    Next
    If GetSelItems <> "" Then GetSelItems = Mid(GetSelItems, 2)
End Function

Private Sub SetEnabled()
    Me.cbo液体量.Enabled = mblnALLOW
    Me.cbo液体名称.Enabled = mblnALLOW
    lvw病区.Enabled = mblnALLOW
    cmdColumn(0).Enabled = mblnALLOW
    cmdColumn(1).Enabled = mblnALLOW
    cmdMove(0).Enabled = mblnALLOW
    cmdMove(1).Enabled = mblnALLOW
    cmd增加.Enabled = mblnALLOW
    cmd修改.Enabled = mblnALLOW
    cmd删除.Enabled = mblnALLOW
    Frame1.Enabled = (cbo病区.ListCount > 0) And InStr(1, mstrPrivs, "编辑") <> 0
    cmd公告栏_保存.Enabled = InStr(1, mstrPrivs, "编辑") <> 0
    cmd公告栏_删除.Enabled = cmd公告栏_保存.Enabled
End Sub

Private Function GetUser病区IDs() As String
'功能：获取操作员所属的病区(直接属于病区或所在科室所属的病区),可能有多个
    Static rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, blnNew As Boolean
        
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    If blnNew Then
        strSQL = _
            "Select Distinct 病区ID From (" & _
            " Select A.部门ID as 病区ID" & _
            " From 部门性质说明 A,部门人员 B" & _
            " Where A.部门ID=B.部门ID And B.人员ID=[1]" & _
            " And A.服务对象 in(1,2,3) And A.工作性质='护理'" & _
            " Union" & _
            " Select A.病区ID From 病区科室对应 A,部门人员 B" & _
            " Where A.科室ID=B.部门ID And B.人员ID=[1])"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", glngUserId)
    ElseIf rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
    End If
    For i = 1 To rsTmp.RecordCount
        GetUser病区IDs = GetUser病区IDs & "," & rsTmp!病区ID
        rsTmp.MoveNext
    Next
    
    GetUser病区IDs = Mid(GetUser病区IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitUnits() As Boolean
'功能：初始化住院护理病区
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strUnits As String, i As Long

    On Error GoTo errH
    strUnits = GetUser病区IDs
    
    '包含门观察室
    If InStr(mstrPrivs, "所有病区") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by A.编码"
    Else
        '求有权病区：直接所在病区+所在科室所属病区
        strSQL = _
            " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.编码,C.名称,Nvl(B.缺省,0) as 缺省" & _
            " From 病区科室对应 A,部门人员 B,部门表 C" & _
            " Where A.病区ID=C.ID And B.部门ID=A.科室ID And B.人员ID=[1]" & _
            " And Exists(Select 1 From 部门性质说明 Where 工作性质='临床' And 部门ID=A.科室ID)" & _
            " And Not Exists(Select 1 From 部门性质说明 Where 工作性质='护理' And 部门ID=A.科室ID)" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = "Select ID,编码,名称,Max(缺省) as 缺省 From (" & strSQL & ") Group by ID,编码,名称 Order by 编码"
    End If

    cbo病区.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, glngUserId)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo病区.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cbo病区.ItemData(cbo病区.NewIndex) = rsTmp!ID
            If InStr(mstrPrivs, "所有病区") > 0 Then
                If rsTmp!ID = glngDeptId Then  '直接所属优先
                    Call zlControl.CboSetIndex(cbo病区.hwnd, cbo病区.NewIndex)
                End If
                If InStr("," & strUnits & ",", "," & rsTmp!ID & ",") > 0 And cbo病区.ListIndex = -1 Then
                    Call zlControl.CboSetIndex(cbo病区.hwnd, cbo病区.NewIndex)
                End If
            Else '所属缺省病区包含的可能有多个
                If rsTmp!缺省 = 1 And cbo病区.ListIndex = -1 Then
                    Call zlControl.CboSetIndex(cbo病区.hwnd, cbo病区.NewIndex)
                End If
            End If
            rsTmp.MoveNext
        Next
    End If
    If cbo病区.ListIndex = -1 And cbo病区.ListCount > 0 Then
        Call zlControl.CboSetIndex(cbo病区.hwnd, 0)
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub RefreshBoard()
    Dim lng病区ID As Long
    Dim intDel As Integer, intCount As Integer
    On Error GoTo errHand
    '刷新公告栏
    
    '先删除所有控件
    intCount = lbl要素名.Count - 1
    For intDel = 1 To intCount
        Unload lbl要素名(intDel)
        Unload lbl要素内容(intDel)
    Next
    '赋初值
    cbo名称.Text = ""
    txt别名.Text = ""
    txt行号.Text = ""
    cbo位置.ListIndex = 0
    chk无数据时隐藏.Value = 0
    txt绑定诊疗项目.Text = ""
    txt绑定诊疗项目.Tag = ""
    
    
    '提取数据
    Frame1.Tag = ""
    lng病区ID = Me.cbo病区.ItemData(Me.cbo病区.ListIndex)
    gstrSQL = " Select ID,分类,名称,别名,行号,位置,是否固定,是否隐藏,内容" & _
              " From 公告栏样式 " & _
              " Where 病区ID=[1] " & _
              " Order by 行号,位置"
    Set mrsBoard = zlDatabase.OpenSQLRecord(gstrSQL, "提取病区公告", lng病区ID)
    
    '依次加载控件
    With mrsBoard
        .Filter = "分类=" & SSTab1.Tab + 1
        Do While Not .EOF
            Load lbl要素名(.AbsolutePosition)
            lbl要素名(.AbsolutePosition).Tag = !ID
            lbl要素名(.AbsolutePosition).Caption = !别名
            lbl要素名(.AbsolutePosition).Top = lbl要素名(0).Top + (!行号 - 1) * 360
            lbl要素名(.AbsolutePosition).Left = IIf(!位置 = 1, 60, 2580)
            lbl要素名(.AbsolutePosition).Visible = True
            
            Load lbl要素内容(.AbsolutePosition)
            lbl要素内容(.AbsolutePosition).Caption = IIf(IsNull(!内容), "", !内容)
            lbl要素内容(.AbsolutePosition).Top = lbl要素内容(0).Top + (!行号 - 1) * 360
            lbl要素内容(.AbsolutePosition).Left = lbl要素名(.AbsolutePosition).Left + lbl要素名(.AbsolutePosition).Width + 60
            lbl要素内容(.AbsolutePosition).AutoSize = False
            lbl要素内容(.AbsolutePosition).WordWrap = False
            lbl要素内容(.AbsolutePosition).Height = 240
            lbl要素内容(.AbsolutePosition).Visible = True
            
            .MoveNext
        Loop
        .Filter = 0
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call RefreshBoard
    Call SetShape
End Sub

Private Sub SetShape(Optional ByVal intIndex As Integer = 0)
    Dim blnShow As Boolean
    blnShow = (intIndex > 0)
    
    If blnShow Then
        Shape(0).Left = lbl要素名(intIndex).Left - Shape(0).Width
        Shape(0).Top = lbl要素名(intIndex).Top - Shape(0).Height
        Shape(1).Left = lbl要素名(intIndex).Left + (lbl要素名(intIndex).Width - Shape(0).Width) / 2
        Shape(1).Top = Shape(0).Top
        Shape(2).Left = lbl要素名(intIndex).Left + lbl要素名(intIndex).Width
        Shape(2).Top = Shape(0).Top
        Shape(3).Left = Shape(2).Left
        Shape(3).Top = lbl要素名(intIndex).Top + (lbl要素名(intIndex).Height - Shape(3).Height) / 2
        Shape(4).Left = Shape(2).Left
        Shape(4).Top = lbl要素名(intIndex).Top + lbl要素名(intIndex).Height
        Shape(5).Left = Shape(1).Left
        Shape(5).Top = Shape(4).Top
        Shape(6).Left = Shape(0).Left
        Shape(6).Top = Shape(4).Top
        Shape(7).Left = Shape(0).Left
        Shape(7).Top = Shape(3).Top
    End If
    
    Shape(0).Visible = blnShow
    Shape(1).Visible = blnShow
    Shape(2).Visible = blnShow
    Shape(3).Visible = blnShow
    Shape(4).Visible = blnShow
    Shape(5).Visible = blnShow
    Shape(6).Visible = blnShow
    Shape(7).Visible = blnShow
End Sub




