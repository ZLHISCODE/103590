VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmPACStation 
   AutoRedraw      =   -1  'True
   Caption         =   "影像工作站"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10995
   Icon            =   "frmPacStation.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame fraLR_s 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   3330
      MousePointer    =   9  'Size W E
      TabIndex        =   13
      Top             =   750
      Width           =   30
   End
   Begin VB.PictureBox picKind 
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   2955
      TabIndex        =   18
      Top             =   720
      Width           =   3015
      Begin MSComctlLib.ListView lvwPati 
         Height          =   3180
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   5609
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "img16"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   18
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "来源"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "单据号"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "姓名"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "内容"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "状态"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "科室"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "病人标识号"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "费别"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "检查时间"
            Object.Width           =   2081
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "执行间"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "检查标识"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "标本部位"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "报告人"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "审核人"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "检查号"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "急"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "打印"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "开嘱时间"
            Object.Width           =   2081
         EndProperty
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "已完成的检查(&3)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   2
         Left            =   240
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   660
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "正进行的检查(&2)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   330
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "待执行的检查(&1)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   0
         Left            =   240
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Width           =   2295
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   3600
      ScaleHeight     =   5535
      ScaleWidth      =   6735
      TabIndex        =   15
      Top             =   1200
      Width           =   6735
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   735
      Top             =   2190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":066C
            Key             =   "未执行"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":0C06
            Key             =   "正执行"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":11A0
            Key             =   "已执行"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":173A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":735C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":7676
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":7AD1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   1244
      BandCount       =   2
      _CBWidth        =   10995
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrMain"
      MinWidth1       =   4995
      MinHeight1      =   645
      NewRow1         =   0   'False
      Caption2        =   "医技科室"
      Child2          =   "cboDept"
      MinWidth2       =   1995
      MinHeight2      =   300
      Width2          =   3495
      NewRow2         =   0   'False
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   8445
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   195
         Width           =   2460
      End
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   645
         Left            =   165
         TabIndex        =   12
         Top             =   30
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   34
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "预览"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "打印"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "记录"
               Key             =   "记录"
               Description     =   "记录"
               Object.ToolTipText     =   "记录执行情况"
               Object.Tag             =   "记录"
               ImageKey        =   "记录"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "调整"
               Key             =   "调整"
               Description     =   "调整"
               Object.ToolTipText     =   "调整执行情况"
               Object.Tag             =   "调整"
               ImageKey        =   "调整"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "完成"
               Key             =   "完成"
               Description     =   "完成"
               Object.ToolTipText     =   "确认执行完成"
               Object.Tag             =   "完成"
               ImageKey        =   "完成"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Exec_"
               Description     =   "执行"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "主费"
               Key             =   "主费"
               Description     =   "申请"
               Object.ToolTipText     =   "调整主费用"
               Object.Tag             =   "主费"
               ImageKey        =   "主费"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "补费"
               Key             =   "补费"
               Description     =   "申请"
               Object.ToolTipText     =   "补充费用"
               Object.Tag             =   "补费"
               ImageKey        =   "补费"
               Style           =   5
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "改费"
               Key             =   "改费"
               Description     =   "申请"
               Object.ToolTipText     =   "修改费用"
               Object.Tag             =   "改费"
               ImageKey        =   "改费"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删费"
               Key             =   "删费"
               Description     =   "申请"
               Object.ToolTipText     =   "删除费用"
               Object.Tag             =   "删费"
               ImageKey        =   "删费"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Money_"
               Description     =   "申请"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "新开"
               Key             =   "新开"
               Description     =   "医嘱"
               Object.ToolTipText     =   "新开医嘱"
               Object.Tag             =   "新开"
               ImageKey        =   "新嘱"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "修改"
               Description     =   "医嘱"
               Object.ToolTipText     =   "修改医嘱"
               Object.Tag             =   "修改"
               ImageKey        =   "修改"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "删除"
               Description     =   "医嘱"
               Object.ToolTipText     =   "删除医嘱"
               Object.Tag             =   "删除"
               ImageKey        =   "删除"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "作废"
               Key             =   "作废"
               Description     =   "医嘱"
               Object.ToolTipText     =   "作废医嘱"
               Object.Tag             =   "作废"
               ImageKey        =   "作废"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Advice_"
               Description     =   "医嘱"
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "病历"
               Key             =   "病历"
               Description     =   "病历"
               Object.ToolTipText     =   "书写一份新的病历文件"
               Object.Tag             =   "病历"
               ImageKey        =   "新嘱"
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "病历修改"
               Description     =   "病历"
               Object.ToolTipText     =   "修改或查阅病历文件"
               Object.Tag             =   "修改"
               ImageKey        =   "修改"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "删病历"
               Description     =   "病历"
               Object.ToolTipText     =   "删除当前病历文件"
               Object.Tag             =   "删除"
               ImageKey        =   "删除"
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "File_"
               Description     =   "病历"
               Style           =   3
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "观片"
               Key             =   "观片"
               Description     =   "影像"
               Object.ToolTipText     =   "在观片工作站中处理当前选择的影像序列"
               Object.Tag             =   "观片"
               ImageKey        =   "观片"
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "采集"
               Key             =   "采集"
               Object.ToolTipText     =   "采集视频图像(B超、胃镜等)"
               Object.Tag             =   "采集"
               ImageKey        =   "Capture"
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "显示"
               Key             =   "显示"
               Description     =   "影像"
               Object.ToolTipText     =   "显示当前序列影像"
               Object.Tag             =   "显示"
               ImageKey        =   "ViewPic"
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "全选"
               Key             =   "全选"
               Description     =   "影像"
               Object.ToolTipText     =   "选择所有影像序列"
               Object.Tag             =   "全选"
               ImageKey        =   "SelectAll"
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "全清"
               Key             =   "全清"
               Description     =   "影像"
               Object.ToolTipText     =   "清除所有影像序列的选择标志"
               Object.Tag             =   "全清"
               ImageKey        =   "ClearAll"
            EndProperty
            BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "View_"
               Description     =   "影像"
               Style           =   3
            EndProperty
            BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "报告"
               Key             =   "报告"
               Object.ToolTipText     =   "填写检查报告"
               Object.Tag             =   "报告"
               ImageKey        =   "Report"
            EndProperty
            BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "审核"
               Key             =   "审核"
               Object.ToolTipText     =   "审核报告"
               Object.Tag             =   "审核"
               ImageKey        =   "Auditing"
            EndProperty
            BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "驳回"
               Key             =   "驳回"
               Object.ToolTipText     =   "驳回当前检查报告"
               Object.Tag             =   "驳回"
               ImageKey        =   "Rollback"
            EndProperty
            BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_Rep"
               Style           =   3
            EndProperty
            BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "过滤"
               Object.ToolTipText     =   "检查记录过滤查询"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "帮助"
            EndProperty
            BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "退出"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   7575
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":7BB3
            Key             =   "预览"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":7DCD
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":7FE7
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":8201
            Key             =   "退出"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":841B
            Key             =   "记录"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":8B15
            Key             =   "调整"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":920F
            Key             =   "完成"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":9909
            Key             =   "主费"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":A003
            Key             =   "补费"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":A6FD
            Key             =   "改费"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":ADF7
            Key             =   "删费"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":B4F1
            Key             =   "新嘱"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":BBEB
            Key             =   "修改"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":C2E5
            Key             =   "删除"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":C9DF
            Key             =   "作废"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":D0D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":D61A
            Key             =   "观片"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":DB5B
            Key             =   "ViewPic"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":E2D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":E4EF
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":E709
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":E923
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":EB3D
            Key             =   "Capture"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":10847
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":10A61
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":10C7B
            Key             =   "Rollback"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   8175
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":10E95
            Key             =   "预览"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":110AF
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":112C9
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":114E3
            Key             =   "退出"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":116FD
            Key             =   "记录"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":11DF7
            Key             =   "调整"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":124F1
            Key             =   "完成"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":12BEB
            Key             =   "主费"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":132E5
            Key             =   "补费"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":139DF
            Key             =   "改费"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":140D9
            Key             =   "删费"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":147D3
            Key             =   "新嘱"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":14ECD
            Key             =   "修改"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":155C7
            Key             =   "删除"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":15CC1
            Key             =   "作废"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":163BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":16B35
            Key             =   "观片"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":172AF
            Key             =   "ViewPic"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":174C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":176E3
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":178FD
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":17B17
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":17D31
            Key             =   "Capture"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":19A3B
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":19C55
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":19E6F
            Key             =   "Rollback"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraState 
      Height          =   1290
      Left            =   75
      TabIndex        =   14
      Top             =   5580
      Width           =   3165
      Begin VB.CheckBox chkFilter 
         Caption         =   "按当前病人筛选"
         Height          =   195
         Left            =   255
         TabIndex        =   24
         Top             =   1005
         Width           =   1860
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Index           =   2
         ItemData        =   "frmPacStation.frx":1A089
         Left            =   1080
         List            =   "frmPacStation.frx":1A09C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Index           =   1
         ItemData        =   "frmPacStation.frx":1A0C2
         Left            =   1080
         List            =   "frmPacStation.frx":1A0CF
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Index           =   0
         ItemData        =   "frmPacStation.frx":1A0E7
         Left            =   1080
         List            =   "frmPacStation.frx":1A0F4
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdSeek 
         Height          =   360
         Left            =   2400
         Picture         =   "frmPacStation.frx":1A10E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "病人定位(F3)"
         Top             =   195
         Width           =   375
      End
      Begin VB.TextBox txt标识号 
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   225
         Width           =   1275
      End
      Begin VB.CheckBox chk状态 
         Caption         =   "包含尚未安排报到的检查"
         Height          =   195
         Index           =   0
         Left            =   2400
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.CheckBox chk状态 
         Caption         =   "包含已经执行完成的检查"
         Height          =   195
         Index           =   2
         Left            =   2400
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "状态(&U)"
         Height          =   200
         Left            =   240
         TabIndex        =   3
         Top             =   650
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标识号(&N)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   285
         Width           =   810
      End
   End
   Begin MSComctlLib.TabStrip TabFile 
      Height          =   330
      Left            =   3480
      TabIndex        =   16
      Top             =   780
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   582
      TabWidthStyle   =   2
      TabFixedWidth   =   1939
      TabFixedHeight  =   441
      HotTracking     =   -1  'True
      ImageList       =   "iLsTree"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "申请(&A)"
            Key             =   "申请"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "影像(&B)"
            Key             =   "影像"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "医嘱(&C)"
            Key             =   "医嘱"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "病历(&D)"
            Key             =   "病历"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pgbLoad 
      Height          =   195
      Left            =   1680
      TabIndex        =   17
      Top             =   7005
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   6945
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPacStation.frx":1A258
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14314
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1920
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Menu mnufile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSetup 
         Caption         =   "参数设置(&S)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileRoom 
         Caption         =   "执行间设置(&R)"
      End
      Begin VB.Menu mnufileImageDevice 
         Caption         =   "影像设备设置(&I)"
      End
      Begin VB.Menu mnufileSendImage 
         Caption         =   "发送图像(&T)"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "退出(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuExec 
      Caption         =   "检查(&E)"
      Begin VB.Menu mnuExecFunc 
         Caption         =   "检查申请(&R)"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "取消申请(&Q)"
         Index           =   1
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "取消安排(&E)"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "开始检查(&A)"
         Index           =   4
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "影像采集(&V)"
         Index           =   5
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "胶片扫描(&S)"
         Index           =   6
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "取消检查(&D)"
         Index           =   7
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "关联影像(&S)"
         Index           =   9
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "取消关联(&G)"
         Index           =   10
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "获取设备影像(&I)"
         Index           =   11
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "删除影像(&P)"
         Index           =   12
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "确认检查完成(&F)"
         Index           =   14
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "取消检查完成(&C)"
         Index           =   15
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "填写报告(&W)"
         Index           =   17
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "审核完成(&U)"
         Index           =   18
      End
   End
   Begin VB.Menu mnuRep 
      Caption         =   "报告(&R)"
      Begin VB.Menu mnuImageView 
         Caption         =   "影像处理(&K)"
         Index           =   0
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuImageView 
         Caption         =   "影像对比(&B)"
         Index           =   1
      End
      Begin VB.Menu mnuImageView 
         Caption         =   "选择所有序列(&A)"
         Index           =   2
      End
      Begin VB.Menu mnuImageView 
         Caption         =   "清除选择标志(&M)"
         Index           =   3
      End
      Begin VB.Menu mnuImageView 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "填写报告(&W)"
         Index           =   0
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "重新填写报告(&R)"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "审核完成(&C)"
         Index           =   3
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "报告驳回(&H)"
         Index           =   4
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "报告打印(&P)"
         Index           =   6
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "报告预览(&V)"
         Index           =   7
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "胶片打印(&R)"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "报告格式(&F)"
         Index           =   9
      End
   End
   Begin VB.Menu mnuMoney 
      Caption         =   "费用(&M)"
      Begin VB.Menu mnuMoneyFunc 
         Caption         =   "生成主费用(&N)"
         Index           =   0
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuMoneyFunc 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuMoneyFunc 
         Caption         =   "补充附加费用(&A)"
         Index           =   2
         Begin VB.Menu mnuMoneyAdd 
            Caption         =   "收费单据(&1)"
            Index           =   0
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuMoneyAdd 
            Caption         =   "记帐单据(&2)"
            Index           =   1
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuMoneyAdd 
            Caption         =   "零费耗用登记(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuMoneyFunc 
         Caption         =   "修改附加费用(&M)"
         Index           =   3
      End
      Begin VB.Menu mnuMoneyFunc 
         Caption         =   "删除附加费用(&D)"
         Index           =   4
      End
   End
   Begin VB.Menu mnuReq 
      Caption         =   "申请(&S)"
      Begin VB.Menu mnuReqFunc 
         Caption         =   "新增申请单(&S)"
         Index           =   0
         Begin VB.Menu ReqList 
            Caption         =   "无可用单据"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "修改申请单(&G)"
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "删除申请单(&R)"
         Index           =   2
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "打印通知单(&P)"
         Index           =   4
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "查阅报告(&V)"
         Index           =   6
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "预览报告(&Y)"
         Index           =   7
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "打印报告(&D)"
         Index           =   8
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "影像对比(&B)"
         Index           =   10
      End
   End
   Begin VB.Menu mnuPFile 
      Caption         =   "病历(&L)"
      Begin VB.Menu mnuPFileFunc 
         Caption         =   "新增病历(&A)"
         Index           =   0
         Begin VB.Menu FileList 
            Caption         =   "无病历文件"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuPFileFunc 
         Caption         =   "修改病历(&M)"
         Index           =   1
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuPFileFunc 
         Caption         =   "删除病历(&D)"
         Index           =   2
      End
   End
   Begin VB.Menu mnuAdvice 
      Caption         =   "医嘱(&Y)"
      Begin VB.Menu mnuAdviceFunc 
         Caption         =   "新开医嘱(&A)"
         Index           =   0
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuAdviceFunc 
         Caption         =   "修改医嘱(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuAdviceFunc 
         Caption         =   "删除医嘱(&D)"
         Index           =   2
      End
      Begin VB.Menu mnuAdviceFunc 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuAdviceFunc 
         Caption         =   "医嘱停止(&S)"
         Index           =   4
      End
      Begin VB.Menu mnuAdviceFunc 
         Caption         =   "医嘱作废(&R)"
         Index           =   5
      End
      Begin VB.Menu mnuAdviceFunc 
         Caption         =   "发送医嘱(&S)"
         Index           =   6
      End
      Begin VB.Menu mnuAdviceFunc 
         Caption         =   "作废医嘱(&R)"
         Index           =   7
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "工具(&T)"
      Begin VB.Menu mnuToolItemRef 
         Caption         =   "诊疗参考(&I)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuToolDiagRef 
         Caption         =   "诊断参考(&D)"
      End
      Begin VB.Menu mnuTool_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolReport 
         Caption         =   "检查工作报表(&1)"
         Index           =   0
      End
      Begin VB.Menu mnuReport 
         Caption         =   "其他报表"
         Begin VB.Menu mnuReportItem 
            Caption         =   "无"
            Enabled         =   0   'False
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolItem 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewToolItem 
            Caption         =   "科室选择(&D)"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuViewTool_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewInfo 
         Caption         =   "病人信息(&I)"
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewCharge 
         Caption         =   "只显示已经收费的病人(&P)"
      End
      Begin VB.Menu mnuViewAdviceSelf 
         Caption         =   "只显示本次下达的医嘱(&O)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewFileSelf 
         Caption         =   "只显示本次书写的病历(&C)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewHistory 
         Caption         =   "显示病人历史病历(&H)"
      End
      Begin VB.Menu mnuViewAdviceAppend 
         Caption         =   "显示发送明细(&D)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewPic 
         Caption         =   "显示当前序列图像(&V)"
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "数据过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "定位方式"
         Begin VB.Menu mnuViewFindItem 
            Caption         =   "标识号(&1)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFindItem 
            Caption         =   "就诊卡(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuViewFindItem 
            Caption         =   "姓名(&3)"
            Index           =   2
         End
         Begin VB.Menu mnuViewFindItem 
            Caption         =   "单据号(&4)"
            Index           =   3
         End
         Begin VB.Menu mnuViewFindItem 
            Caption         =   "检查号(&5)"
            Index           =   4
         End
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmPACStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Private mstrFilter As String
Private mstrPrivs As String
Private mlngPreDept As Long
Private mstrPrePati As String
Private TabIndex As Integer
Private Const COLOR_LOST = &HFFEBD7
Private Const COLOR_FOCUS = &HFFCC99

Private mstrRoom As String, blnIfOnlyShow As Boolean '当前执行间

Private mfrmActive As Object '当前活动窗口
Private WithEvents mfrmRepEdit As Form
Attribute mfrmRepEdit.VB_VarHelpID = -1
Private aForms(4) As Object
Private objImgCapture As Object
Private mBeforeDays As Integer '默认查询的天数
Private mDispImgs As Integer '缩略图显示数
Private mblnEmergencyPrint As Boolean   'True=紧急审核时，不审核只做已打印标记；False=紧急审核时正常进行审核
'过滤条件变量
Private mdatFBegin As Date
Private mdatFEnd As Date
Private mDatType As Integer '1=按检查时间、2=按发送时间
Private mstrFNO As String
Private mlngF科室ID As Long
Private mstrF来源 As String
Private mdblF标识号 As Double
Private mstrF就诊卡 As String
Private mstrF姓名 As String
Private mdblFChkNO As Double
Private mblnViewImage As Boolean '报告时是否观片
Private mblnSample As Boolean '登记后是否直接核收
Private mstr标本部位 As String  '检查标本部位
Private mstrPatiName As String '当前筛选病人

Private Sub cboState_Click(Index As Integer)
    If Me.Tag <> "" Then Me.Tag = "": Exit Sub
    
    Call LoadPatiList
End Sub

Private Sub chkFilter_Click()
    If Me.lvwPati.SelectedItem Is Nothing Then
        Me.chkFilter.Value = 0
    End If
    If Me.chkFilter.Value = 1 Then mstrPatiName = Me.lvwPati.SelectedItem.SubItems(2)
    Call LoadPatiList
End Sub

Private Sub chk状态_Click(Index As Integer)
    If Me.Tag <> "" Then Me.Tag = "": Exit Sub
    
    Call LoadPatiList
End Sub

Private Sub cmdSeek_Click()
    Call Form_KeyDown(vbKeyF3, 0)
End Sub

Private Sub cmdKind_Click(Index As Integer)
    '装数据并调整界面
    If Val(lvwPati.Tag) <> Index Then
        Me.lvwPati.Tag = Index
        Call picKind_Resize
        Call LoadPatiList
        '定位到想查找的病人检查
        If txt标识号.Text <> "" Then Call SeekNextPati(True)
    End If
    If Me.lvwPati.Visible Then
        Me.lvwPati.SetFocus
    End If
    ShowCheck Index
End Sub

Private Sub ShowCheck(ByVal Index As Integer)
    Dim intCount As Integer
    On Error Resume Next
'    With chk状态
'        For intCount = .LBound To .UBound
'            .Item(intCount).Visible = False
'        Next
'        .Item(Index).Visible = True
'    End With
    With cboState
        For intCount = .LBound To .UBound
            .Item(intCount).Visible = False
        Next
        .Item(Index).Visible = True
    End With
End Sub

Private Sub FileList_Click(Index As Integer)
    mfrmActive.zlMenuClick FileList(Index)
End Sub

Private Sub Form_Activate()
    If Me.Tag = "Loading" Then
        Me.Tag = ""
        TabFile.Tabs(TabIndex).Selected = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnFirst As Boolean
    If KeyCode = vbKeyF3 Then
        If txt标识号.Text = "" Then
            txt标识号.SetFocus
        Else
            Call txt标识号_Validate(False)
            Call zlControl.TxtSelAll(txt标识号)
            Call SeekNextPati(txt标识号.Tag <> txt标识号.Text)
        End If
    ElseIf KeyCode = vbKeyF4 Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub SeekNextPati(ByVal blnFirst As Boolean)
    Dim intB As Integer, blnDo As Boolean
    Dim strItem As String, strFind As String, i As Long
    
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    
    intB = 1
    If Not blnFirst Then
        intB = lvwPati.SelectedItem.Index + 1
        If intB > lvwPati.ListItems.Count Then intB = 1
    End If
    Do While True
        For i = intB To lvwPati.ListItems.Count
            blnDo = False
            If txt标识号.Text <> "" Then
                strItem = Split(Label1.Caption, "(")(0)
                With lvwPati.ListItems(i)
                    If strItem = "标识号" Then
                        strFind = .SubItems(6)
                    ElseIf strItem = "就诊卡" Then
                        strFind = .ListSubItems(10).Tag
                    ElseIf strItem = "姓名" Then
                        strFind = .SubItems(2)
                        If strFind Like txt标识号.Text & "*" Then blnDo = True
                        If zlCommFun.SpellCode(strFind) Like UCase(txt标识号.Text) & "*" Then blnDo = True
                    ElseIf strItem = "单据号" Then
                        strFind = .SubItems(1)
                    ElseIf strItem = "检查号" Then
                        strFind = .SubItems(14)
                    End If
                    If strFind = txt标识号.Text Then blnDo = True
                End With
            End If
            If blnDo Then
                txt标识号.Tag = txt标识号.Text
                If lvwPati.SelectedItem.Key <> lvwPati.ListItems(i).Key Then
                    lvwPati.ListItems(i).Selected = True
                    Call lvwPati_ItemClick(lvwPati.SelectedItem)
                    lvwPati.SelectedItem.EnsureVisible
                End If
                Exit Sub
            End If
        Next
        If Not blnFirst And intB > 1 Then
            intB = 1
        Else
            Exit Sub
        End If
    Loop
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Label1_Click()
    Me.PopupMenu Me.mnuViewFind
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub

Private Sub lvwPati_DblClick()
    If Not lvwPati.SelectedItem Is Nothing And mnuRepFunc(0).Visible Then mnuRepFunc_Click 0
End Sub

Private Sub lvwPati_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Key = mstrPrePati Then Exit Sub
    mstrPrePati = Item.Key
    Call tabFile_Click
End Sub

Private Sub lvwPati_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If lvwPati.Tag = "2" And mnuRep.Enabled Then
            If InStr(mstrPrivs, "填写报告") > 0 Or InStr(mstrPrivs, "报告审核") > 0 Then Me.PopupMenu mnuRep
        ElseIf mnuExec.Enabled Then
            If InStr(mstrPrivs, "影像检查") > 0 Then Me.PopupMenu mnuExec
        End If
    End If
End Sub

Private Sub mfrmRepEdit_Unload(Cancel As Integer)
    Dim objPacsCore As Object
    
    Call LoadPatiList
    Set mfrmRepEdit = Nothing
    '关闭观片站窗口
    If mblnViewImage Then
        Set objPacsCore = CreateObject("zl9PacsCore.clsViewer")
        objPacsCore.Closefrom
    End If
End Sub

Private Sub mnufileImageDevice_Click()
    frmPACSImageDeviceSetup.Show vbModal, Me
End Sub

Private Sub mnufileSendImage_Click()
    frmPacsSendImage.ShowMe Me
End Sub

Private Sub mnuImageView_Click(Index As Integer)
    mfrmActive.zlMenuClick mnuImageView(Index)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim lng医嘱ID As Long, lng发送号 As Long
    
    If Me.lvwPati.SelectedItem Is Nothing Then Exit Sub
    With lvwPati.SelectedItem
        lng医嘱ID = Val(Split(Mid(.Key, 2), "_")(0))
        lng发送号 = Val(Split(Mid(.Key, 2), "_")(1))
    End With
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
         "医嘱ID=" & lng医嘱ID, "发送号=" & lng发送号)
End Sub

Private Sub mnuViewFindItem_Click(Index As Integer)
    Dim strItem As String, i As Long
    
    For i = 0 To mnuViewFindItem.UBound
        mnuViewFindItem(i).Checked = i = Index
    Next
    strItem = Split(mnuViewFindItem(Index).Caption, "(")(0)
    Label1.Caption = strItem & "(&D)"
    If strItem = "就诊卡" And gblnCardHide Then
        txt标识号.PasswordChar = "*"
    Else
        txt标识号.PasswordChar = ""
    End If
    txt标识号.Text = "": txt标识号.Tag = ""
    If Visible Then txt标识号.SetFocus
End Sub

Private Sub mnuViewInfo_Click()
    Dim lng病人id As Long
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    
    lng病人id = Val(Split(lvwPati.SelectedItem.Tag, "_")(0))
    Call frmDegreeCard.ShowInfo(Me, lng病人id)
End Sub

Private Sub mnuViewPic_Click()
'    mnuViewPic.Checked = Not mnuViewPic.Checked
    mfrmActive.zlMenuClick mnuViewPic
End Sub

Private Sub picKind_Resize()
    Dim intCount As Integer
    On Error Resume Next
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        Me.cmdKind(intCount).Left = Me.picKind.ScaleLeft + 15
        Me.cmdKind(intCount).Width = Me.picKind.ScaleWidth
        Me.cmdKind(intCount).Height = 300
        If intCount <= Val(lvwPati.Tag) Then
            Me.cmdKind(intCount).Top = Me.picKind.ScaleTop + 285 * intCount
            Me.lvwPati.Top = Me.picKind.ScaleTop + 285 * (intCount + 1)
        Else
            Me.cmdKind(intCount).Top = Me.picKind.ScaleHeight - 285 * (Me.cmdKind.UBound - intCount + 1)
        End If
    Next
    Me.lvwPati.Left = Me.picKind.ScaleLeft + 15
    Me.lvwPati.Width = Me.picKind.ScaleWidth
    Me.lvwPati.Height = Me.picKind.ScaleHeight - 285 * (Me.cmdKind.UBound + 1) - 15
End Sub

Private Function Get执行内容(ByVal lng发送号 As Long, ByVal lng医嘱ID As Long, ByVal lng相关ID As Long, ByVal str类别 As String) As String
'功能：根据指定的医嘱ID,返回医嘱内容供显示
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim bln给药途径 As Boolean, strTmp As String
    
    On Error GoTo errH
    
    '读取医嘱内容
    If str类别 <> "E" Or lng相关ID <> 0 Then
        '配方煎法,手术麻醉或其它医嘱,直接显示医嘱内容
        strSQL = "Select 医嘱内容 From 病人医嘱记录 Where ID= " & IIf(str类别 = "E", "[1]", "[2]")
        
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng相关ID, lng医嘱ID)
            
        If Not rsTmp.EOF Then strTmp = Nvl(rsTmp!医嘱内容)
    Else
        strSQL = "Select A.ID,A.相关ID,A.诊疗类别,A.医嘱内容,A.执行频次,A.执行时间方案,B.名称" & _
            " From 病人医嘱记录 A,诊疗项目目录 B" & _
            " Where Not (A.诊疗类别='E' And 相关ID is Not NULL) And A.诊疗项目ID=B.ID" & _
            " And (A.相关ID= [1] Or A.ID= [1] )" & _
            " Order by A.序号"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
        rsTmp.Filter = "相关ID=" & lng医嘱ID
        If Not rsTmp.EOF Then bln给药途径 = InStr(",5,6,", rsTmp!诊疗类别) > 0
        
        If Not bln给药途径 Then
            '一般治疗项目或中药用法
            rsTmp.Filter = ""
            strSQL = "Select 医嘱内容 From 病人医嘱记录 Where ID=[1] "
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
            If Not rsTmp.EOF Then strTmp = Nvl(rsTmp!医嘱内容)
        Else
            '给药途径
            For i = 1 To rsTmp.RecordCount
                strTmp = strTmp & "," & rsTmp!医嘱内容
                rsTmp.MoveNext
            Next
            rsTmp.Filter = "ID=" & lng医嘱ID
            strTmp = rsTmp!名称 & "," & rsTmp!执行频次 & "(" & rsTmp!执行时间方案 & "):" & Mid(strTmp, 2)
        End If
    End If
    
    '读取发送数次
    strSQL = "Select A.发送数次,C.计算单位" & _
        " From 病人医嘱发送 A,病人医嘱记录 B,诊疗项目目录 C" & _
        " Where A.医嘱ID= [1] And A.发送号= [2] " & _
        " And A.医嘱ID=B.ID And B.诊疗项目ID=C.ID"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
    If IsNull(rsTmp!发送数次) Then
        Get执行内容 = "　执行内容:" & strTmp
    Else
        Get执行内容 = "　发送数次:" & FormatEx(rsTmp!发送数次, 5) & " " & Nvl(rsTmp!计算单位) & ",执行内容:" & strTmp
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mnuAdviceFunc_Click(Index As Integer)
    If mfrmActive Is Nothing Then Exit Sub
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    If Val(lvwPati.SelectedItem.ListSubItems(3).Tag) = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Sub
    End If
    If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
        MsgBox "当前检查已转入备份，不能执行本操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mfrmActive.zlMenuClick mnuAdviceFunc(Index)
End Sub

Private Sub mnuExecFunc_Click(Index As Integer)
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim strSQL As String, rsTmp As New ADODB.Recordset, rsSel As ADODB.Recordset
    Dim iCurrItemIndex As Integer
    
    Dim strImageType As String, strCheckUID As String
    Dim iReturn As Integer, blnCancel As Boolean
'    Dim inte As New clsFtp
    Dim strImageDeviceNumber As String                              '设备号
    Dim strFilter As String                     '取消关联时的序列选择过滤字符串
    
    On Error GoTo DBError
    If Me.lvwPati.SelectedItem Is Nothing And Index > 0 Then Exit Sub
    If Not Me.lvwPati.SelectedItem Is Nothing Then
        With lvwPati.SelectedItem
            lng医嘱ID = Val(Split(Mid(.Key, 2), "_")(0))
            lng发送号 = Val(Split(Mid(.Key, 2), "_")(1))
        End With
    End If
    Select Case Index
        Case 0 '申请预约
'             If RequestRegister(Me, Me.cboDept.ItemData(Me.cboDept.ListIndex)) Then
             If frmPACSReqEdit.ShowMe_Request(Me, Me.cboDept.ItemData(Me.cboDept.ListIndex), blnCheck:=mblnSample) Then
                If mblnSample Then
                    lvwPati.Tag = 1: picKind_Resize
                    Call LoadPatiList("_" & lng医嘱ID & "_" & lng发送号)
                    ShowCheck 1
                Else
                    lvwPati.Tag = 0: picKind_Resize
                    Call LoadPatiList("_" & lng医嘱ID & "_" & lng发送号)
                    ShowCheck 0
                End If
             End If
        Case 1 '取消申请
            '正在执行或已执行不允许拒绝
            If Val(lvwPati.SelectedItem.ListSubItems(3).Tag) = 3 Then
                MsgBox "该申请项目当前正在执行，不能取消。", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(lvwPati.SelectedItem.ListSubItems(3).Tag) = 1 Then
                MsgBox "该申请项目当前已经执行，不能取消。", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "当前检查已转入备份，不能执行本操作！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If MsgBox("确认要取消当前申请吗？" & Chr(10) & Chr(13) & "申请取消后，其对应的医嘱将拒绝执行！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            strSQL = "ZL_病人医嘱执行_拒绝执行(" & lng医嘱ID & "," & lng发送号 & ")"
            ExecuteProc strSQL, Me.Caption
            '处理新的当前申请
            iCurrItemIndex = lvwPati.SelectedItem.Index
            If iCurrItemIndex < lvwPati.ListItems.Count Then
                lvwPati.ListItems(iCurrItemIndex + 1).Selected = True
            ElseIf lvwPati.ListItems.Count > 1 Then
                lvwPati.ListItems(lvwPati.ListItems.Count - 1).Selected = True
            End If
            
            Call LoadPatiList
            If Not lvwPati.SelectedItem Is Nothing Then
                lvwPati.SelectedItem.EnsureVisible
            End If
        Case 4 '开始检查
            '判断执行状态
            If InStr("1", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) > 0 Then
                MsgBox "该检查不能再重新开始。", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "当前检查已转入备份，不能执行本操作！", vbInformation, gstrSysName
                Exit Sub
            End If
'            If Me.lvwPati.SelectedItem.ListSubItems(3).Tag = 3 And _
'                Val(Me.lvwPati.SelectedItem.ListSubItems(9).Tag) > 1 Then
'                MsgBox "该检查正在进行，不能再重新开始。", vbInformation, gstrSysName
'                Exit Sub
'            End If
            
            iReturn = frmPACSReg.ShowMe(Me, lng医嘱ID, lng发送号)
            Select Case iReturn
                Case 1
                    lvwPati.Tag = 1: picKind_Resize
                    Call LoadPatiList("_" & lng医嘱ID & "_" & lng发送号)
                    ShowCheck 1
                Case 2
                    Call LoadPatiList("_" & lng医嘱ID & "_" & lng发送号)
            End Select
        Case 5 '采集
            '判断执行状态
            If Me.lvwPati.SelectedItem.ListSubItems(3).Tag <> "3" Or _
                Me.lvwPati.SelectedItem.ListSubItems(9).Tag <> "2" Then
                MsgBox "当前未进行该检查！", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "当前检查已转入备份，不能执行本操作！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            strSQL = "Select 影像类别,检查UID From 影像检查记录 Where 医嘱ID= [1] And 发送号= [2] "
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
            If rsTmp.EOF Then
                MsgBox "该检查未正常开始，请取消后重新开始。", vbInformation, gstrSysName
                Exit Sub
            End If
            strImageType = Nvl(rsTmp(0)): strCheckUID = Nvl(rsTmp(1))
            
            With lvwPati.SelectedItem
                objImgCapture.ImageCapture mstrPrivs, lng医嘱ID, lng发送号, Me, .SubItems(1), CInt(.ListSubItems(5).Tag), _
                 CLng(Split(.ListSubItems(8).Tag, "|")(0)), CLng(Split(.ListSubItems(8).Tag, "|")(1)), "", _
                 strImageType, strCheckUID
            End With
            
            
'            With lvwPati.SelectedItem
'                EditReport Me, .SubItems(1), CInt(.ListSubItems(5).Tag), _
'                    CLng(Split(.ListSubItems(8).Tag, "|")(0)), CLng(Split(.ListSubItems(8).Tag, "|")(1)), "", _
'                    Val(Me.lvwPati.SelectedItem.ListSubItems(9).Tag) = 6, False, tmpObject, , _
'                    Not (InStr("1,3", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) = 0 Or _
'                InStr("3,4,5,6", Me.lvwPati.SelectedItem.ListSubItems(9).Tag) = 0), , lng医嘱ID, Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1"
'                Set mfrmRepEdit = tmpObject
'            End With
            
            Call tabFile_Click
            Call LoadPatiList("_" & lng医嘱ID & "_" & lng发送号)
            On Error GoTo DBError
        Case 6 '扫描
            '判断执行状态
            If Me.lvwPati.SelectedItem.ListSubItems(3).Tag <> "3" Or _
                Me.lvwPati.SelectedItem.ListSubItems(9).Tag <> "2" Then
                MsgBox "当前未进行该检查！", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "当前检查已转入备份，不能执行本操作！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            strSQL = "Select 影像类别,检查UID From 影像检查记录 Where 医嘱ID= [1] And 发送号= [2] "
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
            If rsTmp.EOF Then
                MsgBox "该检查未正常开始，请取消后重新开始。", vbInformation, gstrSysName
                Exit Sub
            End If
            strImageType = Nvl(rsTmp(0)): strCheckUID = Nvl(rsTmp(1))
            On Error Resume Next
            objImgCapture.ImageScan lng医嘱ID, lng发送号, strImageType, strCheckUID
            
            Call tabFile_Click
            On Error GoTo DBError
        Case 7 '取消检查
            '判断执行状态
            If Me.lvwPati.SelectedItem.ListSubItems(3).Tag <> "3" Or _
                Me.lvwPati.SelectedItem.ListSubItems(9).Tag <> "2" Then
                MsgBox "当前未进行该检查！", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "当前检查已转入备份，不能执行本操作！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If MsgBox("取消本次检查将删除相应的检查序列及其图像，是否继续？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            '删除影像文件和目录
            RemoveCheckImages lng医嘱ID, lng发送号
            strSQL = "ZL_影像检查_CANCEL(" & lng医嘱ID & "," & lng发送号 & ")"
            ExecuteProc strSQL, Me.Caption
            
            lvwPati.Tag = 0: picKind_Resize
            Call LoadPatiList("_" & lng医嘱ID & "_" & lng发送号)
            ShowCheck 0
        Case 9 '关联图像
'            If InStr("0,3", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) = 0 Or _
'                Val(Me.lvwPati.SelectedItem.ListSubItems(9).Tag) > 2 Then
'                MsgBox "当前检查已完成！", vbInformation, gstrSysName
'                Exit Sub
'            End If
            If InStr("0,3", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) = 0 Or _
                InStr("1,2", Me.lvwPati.SelectedItem.ListSubItems(9).Tag) = 0 Then
                MsgBox "当前未进行该检查！", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "当前检查已转入备份，不能执行本操作！", vbInformation, gstrSysName
                Exit Sub
            End If
'            strSQL = "Select Count(*) From 影像检查记录 A,影像检查序列 B" & _
'                " Where A.检查UID=B.检查UID And A.医嘱ID=" & lng医嘱ID & " And 发送号=" & lng发送号
'            OpenRecord rsTmp, strSQL, Me.Caption
'            If rsTmp(0) > 0 Then
'                MsgBox "当前检查项目已有检查图像，不能提取其他检查记录！", vbInformation, gstrSysName
'                Exit Sub
'            End If
            strSQL = "Select A.检查UID As ID,Nvl(A.检查设备,' ') As 检查设备,Nvl(A.检查号,0) As 检查号,Nvl(A.接收日期,Sysdate) As 检查时间," & _
                "Nvl(A.姓名,' ') As 姓名,Nvl(A.英文名,' ') As 英文名,Nvl(A.性别,' ') As 性别,Nvl(A.年龄,' ') As 年龄," & _
                "Nvl(A.出生日期,Sysdate) As 出生日期," & _
                "Nvl(A.身高,0) As 身高,Nvl(A.体重,0) As 体重" & _
                " From 影像临时记录 a,病人医嘱记录 b,影像检查记录 c,病人信息 d" & _
                " Where c.医嘱ID=b.ID And b.病人ID=d.病人ID" & _
                " And (a.检查号=c.检查号 Or a.检查号=c.医嘱ID Or a.检查号=Decode(b.病人来源,2,d.住院号,d.门诊号))" & _
                " And c.医嘱ID=" & lng医嘱ID & " And c.发送号=" & lng发送号
                
            Set rsTmp = OpenSQLRecord(strSQL, "自动对应项目", lng医嘱ID, lng发送号)
'''            If rsTmp.State <> adStateClosed Then rsTmp.Close
'''            Set rsSel = zlDatabase.ShowSelect(Me, strSQL, 0, "检查影像", blnNoneWin:=False, Cancel:=blnCancel, blnSearch:=True)
'''            If rsSel Is Nothing And Not blnCancel Then

            '没有符合筛选条件的记录，查询全部
'                strSQL = "Select A.检查UID As ID,Nvl(A.检查设备,' ') As 检查设备,Nvl(A.检查号,0) As 检查号,Nvl(A.接收日期,Sysdate) As 检查时间," & _
'                    "Nvl(A.姓名,' ') As 姓名,Nvl(A.英文名,' ') As 英文名,Nvl(A.性别,' ') As 性别,Nvl(A.年龄,' ') As 年龄," & _
'                    "Nvl(A.出生日期,Sysdate) As 出生日期," & _
'                    "Nvl(A.身高,0) As 身高,Nvl(A.体重,0) As 体重" & _
'                    " From 影像临时记录 a,病人医嘱记录 b,影像检查项目 c" & _
'                    " Where a.影像类别=c.影像类别 And b.诊疗项目id=c.诊疗项目id And b.id= " & lng医嘱ID

                strSQL = "Select A.检查UID As ID,Nvl(A.检查号,0) As 检查号,Nvl(A.姓名,' ') As 姓名,Nvl(A.检查设备,' ') As 检查设备,Nvl(A.接收日期,Sysdate) As 检查时间," & _
                    "Nvl(A.英文名,' ') As 英文名,Nvl(A.性别,' ') As 性别,Nvl(A.年龄,' ') As 年龄," & _
                    "Nvl(A.出生日期,Sysdate) As 出生日期," & _
                    "Nvl(A.身高,0) As 身高,Nvl(A.体重,0) As 体重" & _
                    " From 影像临时记录 a,影像检查记录 b" & _
                    " Where a.影像类别=b.影像类别 And b.医嘱id= " & lng医嘱ID & " And b.发送号=" & lng发送号 & " order by A.检查号"
                    
                Set rsSel = zlDatabase.ShowSelect(Me, strSQL, 0, "检查影像", False, IIf(rsTmp.EOF, "", rsTmp!检查号), "", , , False, , _
                                                    picKind.Top + lvwPati.Top * 3 + lvwPati.SelectedItem.Height, , _
                                                    blnCancel, , True)
                                                    
'''                If rsTmp.State <> adStateClosed Then rsTmp.Close
'''                Set rsSel = zlDatabase.ShowSelect(Me, strSQL, 0, "检查影像", blnNoneWin:=False, blnSearch:=True)
'''            End If

            If Not rsSel Is Nothing Then
                If MsgBox("是否确认选择的影像是当前检查的？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
                '移动Ftp上的影像文件
                strSQL = "Select 检查UID From 影像检查记录 Where 医嘱ID=[1] And 发送号=[2]"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
                If Not rsTmp.EOF Then
                    If Len(Trim(Nvl(rsTmp(0)))) > 0 Then Call MergeImageFiles(rsSel("ID"), rsTmp(0))
                End If
                
                strSQL = "ZL_影像检查_SET(" & lng医嘱ID & "," & lng发送号 & ",'" & _
                    rsSel("ID") & "')"
                ExecuteProc strSQL, Me.Caption
                
                lvwPati.Tag = 1: picKind_Resize
                Call LoadPatiList("_" & lng医嘱ID & "_" & lng发送号)
                ShowCheck 1
            End If
        Case 10   '取消关联
            '取消关联的最后结果是，每次取消关联后，图象全部按照序列被拆散成N条临时记录
            
            If InStr("0,3", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) = 0 Or _
                InStr("1,2", Me.lvwPati.SelectedItem.ListSubItems(9).Tag) = 0 Then
                MsgBox "当前未进行该检查！", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "当前检查已转入备份，不能执行本操作！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            '显示序列选择窗口
            strSQL = "select 0 as 选择,B.序列UID as ID ,B.序列号,B.序列描述,SUM(1) AS 图像数 from 影像检查记录 A ," & _
                    "影像检查序列 B, 影像检查图象 C Where a.检查UID = B.检查UID And B.序列UID = C.序列UID" & _
                    " And a.医嘱ID = " & lng医嘱ID & " and A.发送号= " & lng发送号 & " group by B.序列UID,B.序列号,B.序列描述"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
            
            frmSelectMuli.ShowSelect Me, rsTmp, "ID,3000,0,1;序列号,800,0,1;序列描述,2000,0,1;图像数,800,0,1", 0, 0, 7000, 5000
            
            If frmSelectMuli.mblnOK = True Then
                strFilter = frmSelectMuli.strFilter
                rsTmp.Filter = strFilter
                '如果有选中序列，则处理每一个序列的取消
                While Not rsTmp.EOF
                    subCancelSeriesRelate lng医嘱ID, lng发送号, rsTmp!ID
                    rsTmp.MoveNext
                Wend
                
                '重新装载病人记录
                lvwPati.Tag = 1: picKind_Resize
                Call LoadPatiList("_" & lng医嘱ID & "_" & lng发送号)
                ShowCheck 1
            End If
        Case 11  '从设备直接提取图像
            strImageDeviceNumber = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPACSImageDeviceSetup", "默认影像设备", "")
            
            '没有默认设备时处理
            If strImageDeviceNumber = "" Then
                If MsgBox("没有设置默认影像检查设备！是否现在设置？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                Else
                    frmPACSImageDeviceSetup.Show vbModal, Me
                    Exit Sub
                End If
            End If
            
            strSQL = "select 设备号 , 设备名, IP地址,端口号,本地AE,设备AE from 影像设备目录 where 设备号 = [1] "
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Mid(strImageDeviceNumber, 2))
            
            '当默认设备被删除后重新设置
            If rsTmp.EOF = True Then
                MsgBox "默认设备已被删除，请重新设置！", vbInformation, gstrSysName
                frmPACSImageDeviceSetup.Show vbModal, Me
                Exit Sub
            End If
                
            frmPACSGetDeviceImage.ShowMe Me, rsTmp("IP地址"), rsTmp("端口号"), rsTmp("设备名"), Nvl(rsTmp("本地AE")), Nvl(rsTmp("设备AE")), lng医嘱ID
            Call LoadPatiList("_" & lng医嘱ID & "_" & lng发送号)
        
        Case 12 '删除检查图像
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "当前检查已转入备份，不能执行本操作！", vbInformation, gstrSysName
                Exit Sub
            End If
            strSQL = "select 检查UID from 影像检查记录 where 医嘱ID = " & lng医嘱ID & " and  发送号 = [1]"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng发送号)
            If rsTmp.RecordCount = 0 Then Exit Sub
            If MsgBox("是否确认要删除该检查的所有影像？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            '删除影像文件和目录
            RemoveCheckImages lng医嘱ID, lng发送号
            strSQL = "ZL_影像检查_PhotoDelete(" & lng医嘱ID & "," & lng发送号 & ")"
            ExecuteProc strSQL, Me.Caption
            
            Call LoadPatiList("_" & lng医嘱ID & "_" & lng发送号)
        Case 14 '检查完成
            '判断执行状态
            If Me.lvwPati.SelectedItem.ListSubItems(3).Tag <> "3" Or _
                Me.lvwPati.SelectedItem.ListSubItems(9).Tag <> "2" Then
                MsgBox "当前未进行该检查！", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "当前检查已转入备份，不能执行本操作！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If MsgBox("确认该项检查已完成吗？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            strSQL = "ZL_影像检查_STATE(" & lng医嘱ID & "," & lng发送号 & ",3)"
            ExecuteProc strSQL, Me.Caption
            
            lvwPati.Tag = 2: picKind_Resize
            Call LoadPatiList("_" & lng医嘱ID & "_" & lng发送号)
            ShowCheck 2
        Case 15 '取消检查完成
            '判断执行状态
            If Me.lvwPati.SelectedItem.ListSubItems(3).Tag <> "3" Then
                MsgBox "当前未进行该检查！", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(Me.lvwPati.SelectedItem.ListSubItems(9).Tag) < 3 Then
                MsgBox "当前检查未完成！", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "当前检查已转入备份，不能执行本操作！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If MsgBox("确认继续进行该项检查吗？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            strSQL = "ZL_影像检查_STATE(" & lng医嘱ID & "," & lng发送号 & ",2)"
            ExecuteProc strSQL, Me.Caption
            
            lvwPati.Tag = 1: picKind_Resize
            Call LoadPatiList("_" & lng医嘱ID & "_" & lng发送号)
            ShowCheck 1
        Case 17 '填写报告
            mnuRepFunc_Click 0
        Case 18 '审核完成
            mnuRepFunc_Click 3
    End Select
    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RemoveCheckImages(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long)
    '删除指定医嘱的检查影像
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    
    Dim Inte As New clsFtp
    Dim strDeviceNO As String
    On Error GoTo ProcError
    '先删除图像
    strSQL = "select a.IP地址, a.FTP目录, a.用户名, a.密码, a.医嘱ID, a.发送号, a.检查UID, a.位置, a.接收日期 ,a.设备号 ,c.图像UID" & _
             " from (select IP地址, FTP目录, 用户名, 密码, 医嘱ID, 发送号, 检查UID, 位置一 as 位置, 接收日期, a.设备号 " & _
             "       from 影像设备目录 a, 影像检查记录 b " & _
             "       Where a.设备号 = B.位置一 " & _
             "       Union All " & _
             "       select IP地址, FTP目录, 用户名, 密码, 医嘱ID, 发送号, 检查UID, 位置二 as 位置, 接收日期, a.设备号" & _
             "       from 影像设备目录 a, 影像检查记录 b " & _
             "       Where a.设备号 = B.位置二 " & _
             "       Union All " & _
             "       select IP地址, FTP目录, 用户名, 密码, 医嘱ID, 发送号, 检查UID, 位置三 as 位置, 接收日期, a.设备号 " & _
             "       from 影像设备目录 a, 影像检查记录 b " & _
             "       Where a.设备号 = B.位置三 " & _
             "       ) a , 影像检查序列 b , 影像检查图象 c " & _
             " Where a.检查uid = B.检查uid " & _
             " and b.序列uid = c.序列uid " & _
             " and a.医嘱ID = [1] And 发送号 = [2] "
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
    Do Until rsTmp.EOF
'        inte.strIPAddress = rsTmp("IP地址")
'        inte.strUser = IIf(IsNull(rsTmp("用户名")), "", rsTmp("用户名"))
'        inte.strPsw = IIf(IsNull(rsTmp("密码")), "", rsTmp("密码"))
        If strDeviceNO <> rsTmp("设备号") Then
            strDeviceNO = rsTmp("设备号")
            Inte.FuncFtpConnect rsTmp("IP地址"), rsTmp("用户名"), rsTmp("密码")
        End If
        Inte.FuncDelFile IIf(IsNull(rsTmp("FTP目录")), "", rsTmp("FTP目录") & "/") & Format(rsTmp("接收日期"), "YYYYMMDD") & "/" & rsTmp("检查UID"), rsTmp("图像UID")
        rsTmp.MoveNext
    Loop
    strDeviceNO = ""
    Inte.FuncFtpDisConnect
    '删除目录
    strSQL = "select IP地址,FTP目录,用户名,密码,医嘱ID,发送号,检查UID,设备号,位置,接收日期 from " & _
             "      (select IP地址,FTP目录,用户名,密码,医嘱ID,发送号,检查UID,a.设备号,位置一 as 位置,接收日期 from 影像设备目录 a , 影像检查记录 b " & _
             "      Where a.设备号 = B.位置一 " & _
             "      Union All " & _
             "      select IP地址,FTP目录,用户名,密码,医嘱ID,发送号,检查UID,a.设备号,位置二 as 位置,接收日期 from 影像设备目录 a , 影像检查记录 b " & _
             "      Where a.设备号 = B.位置二 " & _
             "      Union All " & _
             "      select IP地址,FTP目录,用户名,密码,医嘱ID,发送号,检查UID,a.设备号,位置三 as 位置,接收日期 from 影像设备目录 a , 影像检查记录 b " & _
             "      where a.设备号 = b.位置三 ) a " & _
             " Where a.医嘱ID = [1] And 发送号 = [2] "
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
    Do Until rsTmp.EOF
'        Inte.strIPAddress = rsTmp("IP地址")
'        Inte.strUser = IIf(IsNull(rsTmp("用户名")), "", rsTmp("用户名"))
'        Inte.strPsw = IIf(IsNull(rsTmp("密码")), "", rsTmp("密码"))
        If strDeviceNO <> rsTmp("设备号") Then
            strDeviceNO = rsTmp("设备号")
            Inte.FuncFtpConnect rsTmp("IP地址"), rsTmp("用户名"), rsTmp("密码")
        End If
        Inte.FuncFtpDelDir IIf(IsNull(rsTmp("FTP目录")), "", rsTmp("FTP目录")), Format(rsTmp("接收日期"), "YYYYMMDD") & "/" & rsTmp("检查UID")
        rsTmp.MoveNext
    Loop
    Inte.FuncFtpDisConnect
    Exit Sub
ProcError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuFileRoom_Click()
    Dim strTmp As String, blnTmp As Boolean
    
    On Error Resume Next
    strTmp = mstrRoom: blnTmp = blnIfOnlyShow
    If frmPACSRoom.ShowMe(Me, strTmp, blnTmp, cboDept.ItemData(cboDept.ListIndex)) Then
        mstrRoom = strTmp: blnIfOnlyShow = blnTmp
        Call LoadPatiList
    End If
End Sub

Private Sub mnuMoneyAdd_Click(Index As Integer)
    If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
        MsgBox "当前检查已转入备份，不能执行本操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    mfrmActive.zlMenuClick mnuMoneyAdd(Index)
End Sub

Private Sub mnuMoneyFunc_Click(Index As Integer)
    If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
        MsgBox "当前检查已转入备份，不能执行本操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    mfrmActive.zlMenuClick mnuMoneyFunc(Index)
End Sub

Private Sub mnuPFileFunc_Click(Index As Integer)
    mfrmActive.zlMenuClick mnuPFileFunc(Index)
End Sub

Private Sub mnuRepFunc_Click(Index As Integer)
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strRptName As String
    Dim tmpObject As Object
    Dim iMsgReturn As Integer
    Dim strAudiName As String '审阅人
    Dim blnEmerge As Boolean '急诊
    Dim strIfAuditing As String, strIfRollback As String
    Dim strIfPrint As String
    
    On Error GoTo DBError
    If Me.lvwPati.SelectedItem Is Nothing Then Exit Sub
    
    With lvwPati.SelectedItem
        lng医嘱ID = Val(Split(Mid(.Key, 2), "_")(0))
        lng发送号 = Val(Split(Mid(.Key, 2), "_")(1))
    End With
    
    Select Case Index
        Case 0 '填写报告
            '判断执行状态
'            If InStr("1,3", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) = 0 Or _
'                InStr("3,4,5,6", Me.lvwPati.SelectedItem.ListSubItems(9).Tag) = 0 Then
'                MsgBox "该项检查现在不能填写报告！", vbInformation, gstrSysName
'                Exit Sub
'            End If
            '刷新本条记录
            strSQL = "Select 执行状态,Nvl(执行过程,0) As 执行过程 From 病人医嘱发送 Where 医嘱ID=[1] And 发送号=[2]"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
            If Not rsTmp.EOF Then
                Me.lvwPati.SelectedItem.ListSubItems(3).Tag = rsTmp("执行状态")
                Me.lvwPati.SelectedItem.ListSubItems(9).Tag = rsTmp("执行过程")
            End If
            If Not mfrmRepEdit Is Nothing Then
                Unload mfrmRepEdit
'                MsgBox "正在填写报告。要编辑其他报告，请关闭当前的报告填写窗口！", vbInformation, gstrSysName
'                Call ShowWindow(mfrmRepEdit.Hwnd, SW_RESTORE)
'                Call BringWindowToTop(mfrmRepEdit.Hwnd)
'                Exit Sub
            End If
            '判断是否允许审核
            strIfAuditing = IIf((InStr(mstrPrivs, "报告审核") <> 0), "1", "0")
'            If InStr(mstrPrivs, "报告审核") = 0 And Len(Me.lvwPati.SelectedItem.SubItems(15)) = 0 Then
'                '判断是否急诊
'                blnEmerge = False
'                strSQL = "Select 急诊 From 病人挂号记录 Where NO=[1]"
'                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Split(Me.lvwPati.SelectedItem.Tag, "_")(2))
'                If Not rsTmp.EOF Then blnEmerge = (Nvl(rsTmp(0), 0) = 1)
'
'                If Not blnEmerge Then strIfAuditing = "0"
'            End If
            If Me.lvwPati.SelectedItem.ListSubItems(3).Tag <> "3" Then strIfAuditing = "0"
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then strIfAuditing = "0"
            '判断是否允许驳回
            strIfRollback = IIf((InStr(mstrPrivs, "报告驳回") <> 0), "1", "0")
'            If InStr(mstrPrivs, "报告驳回") = 0 And Me.lvwPati.SelectedItem.SubItems(12) <> UserInfo.姓名 Then
'                strIfRollback = "0"
'            End If
            If InStr("1,3", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) = 0 Then strIfRollback = "0"
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then strIfRollback = "0"
            '判断是否允许打印
            strIfPrint = "0"
            If InStr(mstrPrivs, "报告审核") <> 0 Or Me.lvwPati.SelectedItem.SubItems(12) = "" Or Me.lvwPati.SelectedItem.SubItems(12) = UserInfo.姓名 Then
                strIfPrint = "1"
            End If
            
            With lvwPati.SelectedItem
                EditReport Me, .SubItems(1), CInt(.ListSubItems(5).Tag), _
                    CLng(Split(.ListSubItems(8).Tag, "|")(0)), CLng(Split(.ListSubItems(8).Tag, "|")(1)), "", _
                    Val(Me.lvwPati.SelectedItem.ListSubItems(9).Tag) = 6 Or InStr(mstrPrivs, "填写报告") = 0, False, tmpObject, , _
                    Not (InStr("1,3", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) = 0 Or _
                    InStr("3,4,5,6", Me.lvwPati.SelectedItem.ListSubItems(9).Tag) = 0), True, lng医嘱ID, _
                    Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1", strIfAuditing & strIfRollback & strIfPrint
                Set mfrmRepEdit = tmpObject
            End With
            DoEvents
            '打开观片站
            If mblnViewImage Then
                Me.TabFile.Tabs("影像").Selected = True
                mfrmActive.zlMenuClick Me.mnuImageView(2) '选择所有序列
                mfrmActive.zlMenuClick Me.mnuImageView(0)
            End If
        Case 3 '报告完成
           
            If InStr(mstrPrivs, "报告审核") = 0 And Me.lvwPati.SelectedItem.SubItems(12) <> UserInfo.姓名 Then
                    MsgBox "紧急审核只能审核自己填写的报告！", vbInformation, gstrSysName
                    Exit Sub
            End If
            
            blnEmerge = (Me.lvwPati.SelectedItem.SubItems(15) = "√")
            If InStr(mstrPrivs, "报告审核") = 0 And Len(Me.lvwPati.SelectedItem.SubItems(15)) = 0 Then
                '没有紧急标志的,增加判断是否急诊
                strSQL = "Select 急诊 From 病人挂号记录 Where NO=[1]"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Split(Me.lvwPati.SelectedItem.Tag, "_")(2))
                If Not rsTmp.EOF Then blnEmerge = (Nvl(rsTmp(0), 0) = 1)
                
                If Not blnEmerge Then
                    MsgBox "你只能审核急诊检查！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            
            
            
            If Me.lvwPati.SelectedItem.ListSubItems(3).Tag <> "3" Then
                MsgBox "该项检查报告不能完成审核！", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "当前检查已转入备份，不能执行本操作！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            '没有报告，则直接指定阴阳性
            If CLng(Split(lvwPati.SelectedItem.ListSubItems(8).Tag, "|")(1)) = 0 Then
                If Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "忽略结果阴阳性", 0)) = 0 Then
                    iMsgReturn = MsgBox("请确认检查结果是否为阳性？" & vbCrLf & "选择取消则放弃审核。", vbYesNoCancel + vbQuestion + vbDefaultButton1, gstrSysName)
                    If iMsgReturn = vbCancel Then Exit Sub
                    iMsgReturn = IIf(iMsgReturn = vbYes, 1, 0)
                Else
                    iMsgReturn = 0
                End If
            Else
                iMsgReturn = -1
            End If
            If InStr(mstrPrivs, "报告审核") = 0 And mblnEmergencyPrint = True And blnEmerge = True Then
                '紧急审核且处理成打印
                strSQL = "Select 检查UID From 影像检查记录 Where 医嘱ID= [1] And 发送号= [2] "
                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
                If Not rsTmp.EOF Then
                    strSQL = "ZL_影像检查记录_FilmState('" & rsTmp(0) & "',1)"
                    ExecuteProc strSQL, Me.Caption
                End If
                Call LoadPatiList
            Else
                '正常处理审核
                Call ExeFinish(lng医嘱ID, lng发送号, False, iMsgReturn)
            
                If lvwPati.Tag <> "2" Then
                    lvwPati.Tag = 2: picKind_Resize
                    Call LoadPatiList("_" & lng医嘱ID & "_" & lng发送号)
                    ShowCheck 2
                Else
                    Call LoadPatiList
                End If
            End If
            
        Case 4 '报告驳回
            
'            If InStr(mstrPrivs, "报告驳回") = 0 And strAudiName <> UserInfo.姓名 Then
            If InStr(mstrPrivs, "报告驳回") = 0 And _
                (Me.lvwPati.SelectedItem.SubItems(12) <> UserInfo.姓名 Or (Me.lvwPati.SelectedItem.SubItems(13) <> UserInfo.姓名 And Me.lvwPati.SelectedItem.SubItems(13) <> "")) Then
                MsgBox "你只能驳回自己的报告！", vbInformation, gstrSysName
                Exit Sub
            End If
            If InStr("1,3", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) = 0 Then
                MsgBox "该项检查报告还未填写，无需驳回！", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "当前检查已转入备份，不能执行本操作！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If MsgBox("确认要驳回该项检查报告吗？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            If InStr(mstrPrivs, "报告驳回") = 0 And mblnEmergencyPrint = True Then
                '紧急审核权限,同时使用紧急审核打印的方法进行审核,驳回时,只驳回打印状态
                strSQL = "Select 检查UID From 影像检查记录 Where 医嘱ID= [1] And 发送号= [2] "
                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
                If Not rsTmp.EOF Then
                    strSQL = "ZL_影像检查记录_FilmState('" & rsTmp(0) & "',0)"
                    ExecuteProc strSQL, Me.Caption
                End If
            Else
                If Me.lvwPati.SelectedItem.ListSubItems(9).Tag <> "6" Then
                    strSQL = "ZL_影像检查_STATE(" & lng医嘱ID & "," & lng发送号 & ",5)"
                    ExecuteProc strSQL, Me.Caption
                Else
                    Call ExeFinish(lng医嘱ID, lng发送号, True)
                End If
            End If
            Call LoadPatiList
        Case 6 '报告打印
            If InStr(mstrPrivs, "报告审核") = 0 And Me.lvwPati.SelectedItem.SubItems(12) <> UserInfo.姓名 Then
                MsgBox "你只能打印自己填写的报告！", vbInformation, gstrSysName
                Exit Sub
            End If
            Me.MousePointer = vbHourglass
            PrintDiagReport lng医嘱ID, lng发送号, Me, , Me.picBuffer
            Me.MousePointer = vbDefault
        Case 7 '报告预览
            If InStr(mstrPrivs, "报告审核") = 0 And Me.lvwPati.SelectedItem.SubItems(12) <> UserInfo.姓名 Then
                MsgBox "你只能打印自己填写的报告！", vbInformation, gstrSysName
                Exit Sub
            End If
            Me.MousePointer = vbHourglass
            Me.lvwPati.Enabled = False
            PrintDiagReport lng医嘱ID, lng发送号, Me, 1, Me.picBuffer
            Me.lvwPati.Enabled = True
            Me.MousePointer = vbDefault
        Case 9 '报告格式
            On Error Resume Next
            gintReportFormat = Val(InputBox("请输入报告格式编号：", "输入报告格式", 1))
            If gintReportFormat = 0 Then gintReportFormat = 1
            On Error GoTo 0
    End Select
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExeFinish(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, ByVal blnCancel As Boolean, Optional ByVal iCheckResult As Integer = -1)
'参数：iCHeckResult 结果阴阳性
'       -1：忽略，保留以前结果
'       0：结果阴性
'       1：结果阳性
    Dim strSQL As String
    
    gcnOracle.BeginTrans
    On Error GoTo DBError
    If blnCancel Then
        strSQL = "ZL_病人医嘱执行_Cancel(" & lngAdviceID & "," & lngSendNO & ")"
        ExecuteProc strSQL, Me.Caption
        strSQL = "ZL_影像检查_STATE(" & lngAdviceID & "," & lngSendNO & ",5)"
        ExecuteProc strSQL, Me.Caption
    Else
        If iCheckResult = -1 Then
            strSQL = "ZL_病人医嘱执行_Finish(" & lngAdviceID & "," & lngSendNO & ")"
        Else
            strSQL = "ZL_病人医嘱执行_Finish(" & lngAdviceID & "," & lngSendNO & "," & iCheckResult & ")"
        End If
        ExecuteProc strSQL, Me.Caption
        strSQL = "ZL_影像检查_STATE(" & lngAdviceID & "," & lngSendNO & ",6)"
        ExecuteProc strSQL, Me.Caption
    End If
    gcnOracle.CommitTrans
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    Err.Raise Err.Number, "影像医技完成"
End Sub

Private Sub mnuReqFunc_Click(Index As Integer)
    mfrmActive.zlMenuClick mnuReqFunc(Index)
End Sub

Private Sub mnuToolReport_Click(Index As Integer)
    Select Case Index
        Case 0 '医技工作报表
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1206_1", Me, _
                "执行科室=" & NeedName(cboDept.Text) & "|" & cboDept.ItemData(cboDept.ListIndex))
    End Select
End Sub

Private Sub mnuViewAdviceAppend_Click()
'功能：显示或隐藏医嘱附加表格
'接口：Function zlMenuClick(objMenu as Menu) as Boolean
    
    '调用当前功能窗口接口
    If Not mfrmActive Is Nothing Then
        Call mfrmActive.zlMenuClick(mnuViewAdviceAppend)
    End If
End Sub

Private Sub mnuViewAdviceSelf_Click()
    mnuViewAdviceSelf.Checked = Not mnuViewAdviceSelf.Checked
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuViewCharge_Click()
    mnuViewCharge.Checked = Not mnuViewCharge.Checked
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuViewFileSelf_Click()
    mnuViewFileSelf.Checked = Not mnuViewFileSelf.Checked
    If mnuViewFileSelf.Checked Then mnuViewHistory.Checked = False
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuViewFilter_Click()
    frmPACSFilter.mBeforeDays = mBeforeDays
    frmPACSFilter.Show 1, Me
    mBeforeDays = frmPACSFilter.mBeforeDays
    If frmPACSFilter.mblnOK Then
        '重置过滤变量
        With frmPACSFilter
            '发送时间
            mdatFBegin = Format(.dtpBegin.Value, "yyyy-MM-dd HH:mm:00")
            If Format(.dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(.dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
                mdatFEnd = CDate(0) '表示取当前时间
            Else
                mdatFEnd = Format(.dtpEnd.Value, "yyyy-MM-dd HH:mm:59")
            End If
            mDatType = .FindType
            
            '单据号
            If .txtNO.Text <> "" Then
                mstrFNO = .txtNO.Text
            Else
                mstrFNO = ""
            End If
            
            '检查标本部位
            If Trim(.cboPart.Text) <> "" Then
                mstr标本部位 = .cboPart.Text
            End If
            
            '病人科室
            If .cboDept.ListIndex <> 0 Then
                mlngF科室ID = .cboDept.ItemData(.cboDept.ListIndex)
            Else
                mlngF科室ID = 0
            End If
            
            '病人来源
            If Not (.chk来源(0).Value = 1 And .chk来源(1).Value = 1) Then
                If .chk来源(0).Value = 1 Then
                    mstrF来源 = "1,3,4"
                ElseIf .chk来源(1).Value = 1 Then
                    mstrF来源 = "2,3"
                End If
            Else
                mstrF来源 = ""
            End If
            
            '病人标识
            If .txt标识号.Text <> "" Then
                mdblF标识号 = Val(.txt标识号.Text)
            Else
                mdblF标识号 = 0
            End If
            If .txt就诊卡.Text <> "" Then
                mstrF就诊卡 = .txt就诊卡.Text
            Else
                mstrF就诊卡 = ""
            End If
            If .txt姓名.Text <> "" Then
                mstrF姓名 = .txt姓名.Text
            Else
                mstrF姓名 = ""
            End If
            If .txtChkNO.Text <> "" Then
                mdblFChkNO = Val(.txtChkNO.Text)
            Else
                mdblFChkNO = 0
            End If
        End With
        Call mnuViewRefresh_Click
        
        Me.chkFilter.Value = 0
    End If
End Sub

Private Sub mnuViewHistory_Click()
    mnuViewHistory.Checked = Not mnuViewHistory.Checked
    If mnuViewHistory.Checked Then mnuViewFileSelf.Checked = False
    Call mnuViewRefresh_Click
End Sub

Private Sub picFile_Resize()
'功能：处理窗体的Resize
    If Not mfrmActive Is Nothing Then
        SetWindowPos mfrmActive.Hwnd, 0, 0, 0, picFile.ScaleWidth / Screen.TwipsPerPixelX, picFile.ScaleHeight / Screen.TwipsPerPixelY, SWP_NOREPOSITION Or SWP_FRAMECHANGED
    End If
End Sub

Private Sub ReqList_Click(Index As Integer)
    mfrmActive.zlMenuClick ReqList(Index)
End Sub

Private Sub tabFile_Click()
'功能：根据选项分别调用相应窗体
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim lng病人id As Long, int病人来源 As Integer
    Dim lng主页ID As Long, str挂号单 As String
    Dim int计费状态 As Integer, int记录性质 As Integer
    Dim iNum As Integer
    Dim lngPatientID As Long, strCheckID As String
    Dim i As Integer
    Dim strMsg As String
        
    '1.显示医嘱内容
'    lblAdvice.Caption = Get执行内容(lng发送号, lng医嘱ID, Val(Item.ListSubItems(1).Tag), Item.ListSubItems(2).Tag)
    On Error Resume Next
    If Not mfrmActive Is Nothing Then 'And TabIndex <> TabFile.SelectedItem.Index Then
        mfrmActive.Hide
        Set mfrmActive = Nothing
    End If
    TabIndex = TabFile.SelectedItem.Index
    
    Me.mnuPFile.Visible = False
    Me.mnuReq.Visible = False
    Me.mnuAdvice.Visible = False
    Me.mnuMoney.Visible = False
    
    Me.mnuViewAdviceSelf.Visible = False
    Me.mnuViewFileSelf.Visible = False
    Me.mnuViewHistory.Visible = False
    Me.mnuViewAdviceAppend.Visible = False
    Me.mnuViewPic.Visible = False
    
    Me.mnuImageView(0).Visible = False
    Me.mnuImageView(1).Visible = False
    Me.mnuImageView(2).Visible = False
    Me.mnuImageView(3).Visible = False
    
    Select Case TabFile.SelectedItem.Key
        Case "病历"
            Me.mnuPFile.Visible = True
            Me.mnuReq.Visible = True
            Me.mnuViewFileSelf.Visible = True
            Me.mnuViewHistory.Visible = True
            If aForms(4) Is Nothing Then Set aForms(4) = New frmPACSRec
            Set mfrmActive = aForms(4)
            
            Set mfrmActive.mfrmParent = Me
            mfrmActive.mstrPrivs = mstrPrivs
        Case "医嘱"
            Me.mnuAdvice.Visible = True
            Me.mnuViewAdviceSelf.Visible = True
            Me.mnuViewAdviceAppend.Visible = True
            
            If Me.lvwPati.SelectedItem Is Nothing Then
                If aForms(1) Is Nothing Then Set aForms(1) = InDoctorAdvice
                Set mfrmActive = aForms(1)
            ElseIf lvwPati.SelectedItem.Text = "门诊" Then
                If aForms(2) Is Nothing Then Set aForms(2) = OutDoctorAdvice
                Set mfrmActive = aForms(2)
            Else
                If aForms(1) Is Nothing Then Set aForms(1) = InDoctorAdvice
                Set mfrmActive = aForms(1)
            End If
            
            Set mfrmActive.mfrmParent = Me
            mfrmActive.mstrPrivs = mstrPrivs
        Case "申请"
            Me.mnuMoney.Visible = True
            If aForms(0) Is Nothing Then Set aForms(0) = New frmPACSReq
            Set mfrmActive = aForms(0)
        Case "影像"
            Me.mnuImageView(0).Visible = True
            Me.mnuImageView(1).Visible = True
            Me.mnuImageView(2).Visible = True
            Me.mnuImageView(3).Visible = True
            Me.mnuViewPic.Visible = True
            If aForms(3) Is Nothing Then Set aForms(3) = New frmPACSImg
            Set mfrmActive = aForms(3)
    End Select
    
    '工具栏处理
    For iNum = 1 To Me.tbrMain.Buttons.Count
        If Len(Me.tbrMain.Buttons(iNum).Description) > 0 And _
            Me.tbrMain.Buttons(iNum).Description <> TabFile.SelectedItem.Key Then
            Me.tbrMain.Buttons(iNum).Visible = False
        Else
            Me.tbrMain.Buttons(iNum).Visible = True
        End If
    Next
    '根据授权设置主界面权限(Visible),其它权限(Enabled)在子窗口中处理
    Call SetFuncPrivs
    If mfrmActive Is Nothing Then Exit Sub
    
    SetWindowLong mfrmActive.Hwnd, GWL_STYLE, WS_CHILD
    mfrmActive.Show , Me
    SetParent mfrmActive.Hwnd, picFile.Hwnd
    mfrmActive.ZOrder 0
    
    picFile_Resize
    
    If Me.lvwPati.SelectedItem Is Nothing Then
        lng医嘱ID = 0
        lng发送号 = 0
        lng病人id = 0
        lng主页ID = 0
        str挂号单 = ""
        int病人来源 = 2
    Else
        With lvwPati.SelectedItem
            lng医嘱ID = Val(Split(Mid(.Key, 2), "_")(0))
            lng发送号 = Val(Split(Mid(.Key, 2), "_")(1))
            lng病人id = Val(Split(.Tag, "_")(0))
            lng主页ID = Val(Split(.Tag, "_")(1))
            str挂号单 = Split(.Tag, "_")(2)
            int病人来源 = IIf(.Text = "门诊", 1, 2)
        End With
    End If
    '菜单及工具栏处理
    ShowMenu
    If Me.Visible Then
        Select Case TabFile.SelectedItem.Key
            Case "病历"
                ShowAddFileMenu int病人来源 '显示病历菜单
                
                Me.MousePointer = vbHourglass
                mfrmActive.zlRefresh lng病人id, IIf(int病人来源 = 1, str挂号单, lng主页ID), lng医嘱ID, Not Me.mnuViewFileSelf.Checked, Me.mnuViewHistory.Checked, _
                    Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1"
                Me.MousePointer = vbDefault
            Case "医嘱"
                Me.MousePointer = vbHourglass
                If int病人来源 = 2 Then
                    mfrmActive.zlRefresh lng病人id, lng主页ID, 0, 0, False, lng医嘱ID, Not Me.mnuViewAdviceSelf.Checked
                Else
                    mfrmActive.zlRefresh lng病人id, str挂号单, 1, 0, lng医嘱ID, Not Me.mnuViewAdviceSelf.Checked
                End If
                Me.MousePointer = vbDefault
                Me.stbThis.Panels(2).Text = ""
            Case "申请"
                strMsg = Me.stbThis.Panels(2).Text
                BeginShowProgress "正在读取："
                Me.MousePointer = vbHourglass
                mfrmActive.zlRefresh Me, lng医嘱ID, lng发送号, mstrPrivs, pgbLoad, Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1"
                Me.MousePointer = vbDefault
                Me.pgbLoad.Visible = False
                Me.stbThis.Panels(2).Text = strMsg
            Case "影像"
                strMsg = Me.stbThis.Panels(2).Text
                BeginShowProgress "正在读取："
                Me.MousePointer = vbHourglass
                If mfrmActive.zlRefresh(Me, lng医嘱ID, lng发送号, mstrPrivs, pgbLoad, mnuViewPic.Checked, Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1", mDispImgs) = True Then
                    mnuExecFunc(10).Enabled = True
                    mnuExecFunc(12).Enabled = True
                Else
                    mnuExecFunc(10).Enabled = False
                    mnuExecFunc(12).Enabled = False
                End If
                Me.MousePointer = vbDefault
                Me.pgbLoad.Visible = False
                Me.stbThis.Panels(2).Text = strMsg
        End Select
    End If
    
    lvwPati.SetFocus
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
'        Case "完成"
'            mnuExecFinish_Click
        Case "退出"
            mnuFileQuit_Click
        Case "采集"
            mnuExecFunc_Click 5
        Case "打印"
            mnuFilePrint_Click
        Case "预览"
            mnuFilePreview_Click
        Case "报告"
            mnuRepFunc_Click 0
        Case "审核"
            mnuRepFunc_Click 3
        Case "驳回"
            mnuRepFunc_Click 4
        Case "帮助"
            mnuHelpTitle_Click
        Case "过滤"
            mnuViewFilter_Click
        Case Else
            mfrmActive.zlButtonClick Button
    End Select
End Sub

Private Sub mnuToolDiagRef_Click()
'功能：调用诊断参考
    Call ShowDiagHelp(0, Me)
End Sub

Private Sub mnuToolItemRef_Click()
'功能: 调用诊疗参考
    Dim lng诊疗项目ID As Long

    If Me.TabFile.SelectedItem.Key = "医嘱" Then
        mfrmActive.zlItemRef
    Else
        If Not lvwPati.SelectedItem Is Nothing Then lng诊疗项目ID = Val(lvwPati.SelectedItem.ListSubItems(7).Tag)
        Call ShowClinicHelp(0, Me, lng诊疗项目ID)
    End If
End Sub

Private Sub mnuFilePrintSet_Click()
'功能：打印设置
    Call zlPrintSet
End Sub

Private Sub mnuFileExcel_Click()
'功能：输出到Excel
    Call OutputList(3)
End Sub

Private Sub mnuFilePreview_Click()
'功能：打印预览
    Call OutputList(2)
End Sub

Private Sub mnuFilePrint_Click()
'功能：打印
    Call OutputList(1)
End Sub

Private Sub SetFuncPrivs()
'功能：根据授权设置主界面权限(Visible)
    Dim i As Integer
    On Error Resume Next
    If InStr(mstrPrivs, "直接申请") = 0 Then
        Me.mnuExecFunc(0).Visible = False
        Me.mnuExecFunc(1).Visible = False
        Me.mnuExecFunc(3).Visible = False
    End If
    If InStr(mstrPrivs, "影像检查") = 0 Then
        Me.mnuExec.Visible = False
    End If
    If InStr(mstrPrivs, "填写报告") = 0 Then
'        Me.mnuRepFunc(0).Visible = False
        Me.mnuRepFunc(1).Visible = False
        Me.mnuRepFunc(2).Visible = False
        
        Me.mnuExecFunc(16).Visible = False
        Me.mnuExecFunc(17).Visible = False
        
        Me.tbrMain.Buttons("报告").Visible = False
        '没有填写报告权限也不能打印报告
        Me.mnuRepFunc(5).Visible = False
        Me.mnuRepFunc(6).Visible = (InStr(mstrPrivs, "报告审核") > 0) '有审核权限可以打印
        Me.mnuRepFunc(7).Visible = False
    End If
    If InStr(mstrPrivs, "报告审核") = 0 And InStr(mstrPrivs, "紧急审核") = 0 Then
        Me.mnuRepFunc(3).Visible = False
        Me.tbrMain.Buttons("审核").Visible = False
        
        Me.mnuExecFunc(18).Visible = False
    End If
    If InStr(mstrPrivs, "报告驳回") = 0 And InStr(mstrPrivs, "紧急审核") = 0 Then
        Me.mnuRepFunc(4).Visible = False
        Me.tbrMain.Buttons("驳回").Visible = False
    End If
    If Not Me.mnuRepFunc(3).Visible And Not Me.mnuRepFunc(4).Visible Then
        Me.mnuRepFunc(5).Visible = False
    End If
    If InStr(mstrPrivs, "影像处理") = 0 Then
        Me.mnuImageView(0).Visible = False
        Me.mnuImageView(1).Visible = False
        Me.mnuImageView(2).Visible = False
        Me.mnuImageView(3).Visible = False
        
        Me.mnuReqFunc(9).Visible = False
        Me.mnuReqFunc(10).Visible = False
        Me.mnuReqFunc(11).Visible = False
        
        
        Me.tbrMain.Buttons("观片").Visible = False
        Me.tbrMain.Buttons("全选").Visible = False
        Me.tbrMain.Buttons("全清").Visible = False
    End If
    If InStr(mstrPrivs, "填写报告") = 0 And InStr(mstrPrivs, "报告审核") = 0 And InStr(mstrPrivs, "报告驳回") = 0 And InStr(mstrPrivs, "紧急审核") = 0 Then
        Me.mnuImageView(3).Visible = False
        For i = 0 To mnuRepFunc.Count - 1
            mnuRepFunc(i).Visible = False
        Next
    
        Me.tbrMain.Buttons("Split_Rep").Visible = False
    End If
    If InStr(mstrPrivs, "影像处理") = 0 And InStr(mstrPrivs, "填写报告") = 0 _
        And InStr(mstrPrivs, "报告审核") = 0 And InStr(mstrPrivs, "报告驳回") = 0 Then Me.mnuRep.Visible = False
'    If InStr(mstrPrivs, "影像处理") = 0 Or (InStr(mstrPrivs, "填写报告") = 0 And InStr(mstrPrivs, "报告审核") = 0 And InStr(mstrPrivs, "报告驳回") = 0) Then
'        If InStr(mstrPrivs, "填写报告") = 0 And InStr(mstrPrivs, "报告审核") = 0 And InStr(mstrPrivs, "报告驳回") = 0 Then Me.mnuRep.Visible = False
'        Me.mnuImageView(0).Visible = False
'        Me.mnuImageView(1).Visible = False
'        Me.mnuViewPic.Visible = False
'
'        Me.tbrMain.Buttons("观片").Visible = False
'        Me.tbrMain.Buttons("显示").Visible = False
'        Me.tbrMain.Buttons("View_").Visible = False
'    End If
    If InStr(mstrPrivs, "补充费用") = 0 Then
        Me.mnuMoney.Visible = False
    
        Me.tbrMain.Buttons("主费").Visible = False
        Me.tbrMain.Buttons("补费").Visible = False
        Me.tbrMain.Buttons("改费").Visible = False
        Me.tbrMain.Buttons("删费").Visible = False
        Me.tbrMain.Buttons("Money_").Visible = False
    End If
    If InStr(mstrPrivs, "医嘱下达") = 0 Then
        Me.mnuAdvice.Visible = False
    
        Me.tbrMain.Buttons("新开").Visible = False
        Me.tbrMain.Buttons("修改").Visible = False
        Me.tbrMain.Buttons("删除").Visible = False
        Me.tbrMain.Buttons("作废").Visible = False
        Me.tbrMain.Buttons("Advice_").Visible = False
    End If
    If InStr(mstrPrivs, "病历书写") = 0 Then
        Me.mnuPFile.Visible = False
    
        Me.tbrMain.Buttons("病历").Visible = False
        Me.tbrMain.Buttons("病历修改").Visible = False
        Me.tbrMain.Buttons("删病历").Visible = False
        Me.tbrMain.Buttons("File_").Visible = False
    End If
    
    If InStr(mstrPrivs, "填写申请") = 0 Then
        'Me.mnuReq.Visible = False
        Me.mnuReqFunc(0).Visible = False
        Me.mnuReqFunc(1).Visible = False
        Me.mnuReqFunc(2).Visible = False
        Me.mnuReqFunc(3).Visible = False
    End If
    
    '控制“申请”菜单下面的“打印预览”和“报告打印”菜单项
    If InStr(mstrPrivs, "报告审核") = 0 Then
        Me.mnuReqFunc(7).Visible = False
        Me.mnuReqFunc(8).Visible = False
    End If
    
    If InStr(mstrPrivs, "视频采集") = 0 Then
        Me.mnuExecFunc(5).Visible = False
        Me.tbrMain.Buttons("采集").Visible = False
    End If
    If InStr(mstrPrivs, "开始检查") = 0 Then
        Me.mnuExecFunc(4).Visible = False
    End If
    If InStr(mstrPrivs, "取消检查") = 0 Then
        Me.mnuExecFunc(7).Visible = False
    End If
    If InStr(mstrPrivs, "清除检查图像") = 0 Then
        Me.mnuExecFunc(12).Visible = False
    End If
    
    '去掉分隔线
    If InStr(mstrPrivs, "开始检查") = 0 And InStr(mstrPrivs, "视频采集") = 0 And InStr(mstrPrivs, "取消检查") = 0 Then
        Me.mnuExecFunc(8).Visible = False
    End If
    '文件发送
    If InStr(mstrPrivs, "文件发送") = 0 Then
        mnufileSendImage.Visible = False
    End If
    
End Sub

Private Sub mnuHelpTitle_Click()
'功能：调用帮助主题
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim blnTmp As Boolean
    Dim i As Integer
    Dim ret As Long
    
    Call RestoreWinState(Me, App.ProductName)
    Me.lvwPati.ColumnHeaders(16).Position = 1
    
    InitLocalPars
    mBeforeDays = 2
    
    '增加发布到模块的报表菜单
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    '查找方式
    i = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "FindItem", 0))
    mnuViewFindItem(i).Checked = True
    Call mnuViewFindItem_Click(CInt(i))
    
    Me.Tag = "Loading"
    cboState(0).ListIndex = 0
    Me.Tag = "Loading"
    cboState(1).ListIndex = 0
    Me.Tag = "Loading"
    cboState(2).ListIndex = 0
    Me.Tag = "Loading"
    
    '过滤条件变量
    mdatFBegin = CDate(Format(zlDatabase.Currentdate - mBeforeDays, "yyyy-mm-dd 00:00"))
    mdatFEnd = CDate(0)
    mstrFNO = ""
    mlngF科室ID = 0
    mstrF来源 = ""
    mdblF标识号 = 0
    mstrF就诊卡 = ""
    mstrF姓名 = ""
    mDatType = 1
    
    '初始化报告格式为1
    gintReportFormat = 1
    
    '权限处理
    mstrPrivs = gstrPrivs
    Call SetFuncPrivs
    
    mlngPreDept = -1
    mstrPrePati = ""
    mstrFilter = ""
        
    Call InitSysPar '初始化系统参数
    
    AddFileList '构造病历菜单
    LoadBillList
    
    lvwPati.ListItems.Add , , "Temp", , 1
    lvwPati.ListItems.Clear
    
    '初始化医技科室
    If Not InitDepts Then Unload Me: Exit Sub
    If cboDept.ListIndex = -1 Then
        If InStr(mstrPrivs, "所有科室") > 0 Then
            MsgBox "没有发现医技科室信息,请先到部门管理中设置。", vbInformation, gstrSysName
        Else
            MsgBox "没有发现你所属科室,不能使用医技工作站。", vbInformation, gstrSysName
        End If
        Unload Me: Exit Sub
    End If

    TabIndex = 1
    TabFile.Tabs(TabIndex).Selected = True

    Set objImgCapture = CreateObject("zl9ImgCapture.clsImgCapture")
    objImgCapture.InitImgCapture gcnOracle

    '定义热键
    '记录原来的window程序地址
    If App.LogMode <> 0 Then
        preWinProc = GetWindowLong(Me.Hwnd, GWL_WNDPROC)
        '用自定义程序代替原来的window程序
        ret = SetWindowLong(Me.Hwnd, GWL_WNDPROC, AddressOf Wndproc)
    End If
    idHotKey = 1
    Modifiers = MOD_CONTROL     'Ctrl 键
    uVirtKey = vbKey1  '1键
    ret = RegisterHotKey(Me.Hwnd, idHotKey, Modifiers, uVirtKey)

End Sub

Private Sub InitLocalPars()
    mnuViewAdviceSelf.Checked = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "只显示本科下达的医嘱", 1)) <> 0
    mnuViewFileSelf.Checked = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "只显示本次书写的病历", 1)) <> 0
    mnuViewCharge.Checked = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "只显示已收费的病人", 0)) <> 0
'    mnuViewPic.Checked = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "显示当前序列图像", 0)) <> 0
    
    mstrRoom = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "当前执行间")
    blnIfOnlyShow = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "只处理当前执行间项目", False)
    mDispImgs = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "显示图像数", 20)
    mblnEmergencyPrint = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "紧急审核时打印", 0)
    
    mblnViewImage = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "报告时观片", 0))
    mblnSample = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "登记直接检查", 0))
'    mBeforeDays = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "申请查询的天数", 3)
'    If mBeforeDays <= 0 Then mBeforeDays = 3
End Sub

Public Sub mnuViewRefresh_Click()
    Call LoadPatiList
End Sub

Private Sub cboDept_Click()
    If cboDept.ListIndex = mlngPreDept Then Exit Sub
    mlngPreDept = cboDept.ListIndex
    
    Call LoadPatiList
End Sub

Private Sub cbr_Resize()
    Call Form_Resize
End Sub

Private Sub fraLR_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngMinWidth As Long
    If Button <> 1 Then Exit Sub
    
    lngMinWidth = Me.cmdSeek.Left + Me.cmdSeek.Width + Me.fraState.Left + 150
    fraLR_s.BackColor = RGB(0, 0, 0)
    On Error Resume Next
    If fraLR_s.Left + x < lngMinWidth Then
        fraLR_s.Left = lngMinWidth
    ElseIf Me.ScaleWidth - fraLR_s.Left - x < 2000 Then
        fraLR_s.Left = Me.ScaleWidth - 2000
    Else
        fraLR_s.Left = fraLR_s.Left + x
    End If
End Sub

Private Sub fraLR_s_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub

    fraLR_s.BackColor = Me.BackColor
    Form_Resize
End Sub

Private Sub mnuhelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuFileSetup_Click()
    frmTechnicSetup.Show 1, Me
    If frmTechnicSetup.mblnOK Then
'        Call LoadBillDetail(vsMoney.Row)
        InitLocalPars
    
        Call LoadPatiList
        '定位到想查找的病人检查
        If txt标识号.Text <> "" Then Call SeekNextPati(True)
        If Me.lvwPati.Visible Then
            Me.lvwPati.SetFocus
        End If
    End If
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolItem_Click(Index As Integer)
    Dim blnEnabled As Boolean, blnVisible As Boolean, i As Integer
    
    mnuViewToolItem(Index).Checked = Not mnuViewToolItem(Index).Checked
    cbr.Bands(Index + 1).Visible = Not cbr.Bands(Index + 1).Visible

    blnEnabled = False: blnVisible = False
    For i = 1 To cbr.Bands.Count
        '只有有一个ToolBar可见,则"显示文本"菜单可见
        If TypeName(cbr.Bands(i).Child) = "Toolbar" Then
            If cbr.Bands(i).Visible Then
                blnEnabled = True
            End If
        End If
        '只要有一个Band可见,则CoolBar可见
        If cbr.Bands(i).Visible Then
            blnVisible = True
        End If
    Next
    mnuViewToolText.Enabled = blnEnabled
    cbr.Visible = blnVisible
    
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer, j As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To cbr.Bands.Count
        If TypeName(cbr.Bands(i).Child) = "Toolbar" Then
            For j = 1 To cbr.Bands(i).Child.Buttons.Count
                cbr.Bands(i).Child.Buttons(j).Caption = IIf(mnuViewToolText.Checked, cbr.Bands(i).Child.Buttons(j).Tag, "")
            Next
            If Not mnuViewToolText.Checked Then
                cbr.Bands(i).Child.TextAlignment = tbrTextAlignBottom
            End If
            cbr.Bands(i).MinHeight = cbr.Bands(i).Child.ButtonHeight
            cbr.Bands(i).Child.Refresh
        End If
    Next
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage Hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo Hwnd
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long, i As Long
    Dim lngMinWidth As Long

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    lngMinWidth = Me.cmdSeek.Left + Me.cmdSeek.Width + Me.fraState.Left + 150
    If Me.fraLR_s.Left > Me.ScaleWidth Then Me.fraLR_s.Left = Me.ScaleWidth - 2000
    If Me.fraLR_s.Left < lngMinWidth Then Me.fraLR_s.Left = lngMinWidth
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    picKind.Left = 0
    picKind.Top = cbrH
    picKind.Height = Me.ScaleHeight - cbrH - staH - Me.fraState.Height + 45
    picKind.Width = fraLR_s.Left
    
    Me.fraState.Left = 0
    Me.fraState.Top = Me.ScaleHeight - staH - fraState.Height
    Me.fraState.Width = Me.picKind.Width
    
    fraLR_s.Top = picKind.Top
    fraLR_s.Height = Me.ScaleHeight - staH - cbrH
    
    With TabFile
        .Left = fraLR_s.Left + fraLR_s.Width: .Top = cbrH
        .Width = Me.ScaleWidth - .Left
    End With
    With picFile
        .Left = fraLR_s.Left + fraLR_s.Width: .Top = TabFile.Top + TabFile.Height '- 140
        .Width = Me.ScaleWidth - .Left: .Height = Me.ScaleHeight - staH - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    Dim objPacsCore As Object
    
    On Error Resume Next
    For i = 0 To 4
        If Not aForms(i) Is Nothing Then
            Unload aForms(i)
            Set aForms(i) = Nothing
        End If
    Next
    If Not mfrmActive Is Nothing Then
        Unload mfrmActive
        Set mfrmActive = Nothing
    End If
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "只显示已收费的病人", IIf(mnuViewCharge.Checked, 1, 0)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "只显示本科下达的医嘱", IIf(mnuViewAdviceSelf.Checked, 1, 0)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "只显示本次书写的病历", IIf(mnuViewFileSelf.Checked, 1, 0)
'    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "显示当前序列图像", IIf(mnuViewPic.Checked, 1, 0)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "包含未报到检查", Me.chk状态(0).Value = 1
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "包含已执行检查", Me.chk状态(2).Value = 1
'    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "申请查询的天数", mBeforeDays
    '查找方式
    For i = 0 To mnuViewFindItem.UBound
        If mnuViewFindItem(i).Checked Then
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "FindItem", i
        End If
    Next
    
    Call SaveWinState(Me, App.ProductName)
    
    '关闭观片站窗口
    Set objPacsCore = CreateObject("zl9PacsCore.clsViewer")
    objPacsCore.Closefrom
    
    '关闭采集站
    objImgCapture.UnladImgCapture
    Set objImgCapture = Nothing
End Sub

Private Function InitDepts() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str科室IDs As String, str来源 As String
    
    On Error GoTo errH
    
    '包含门诊/住院医技科室
    str来源 = "3"
    If InStr(mstrPrivs, "门诊病人") > 0 And InStr(mstrPrivs, "住院病人") > 0 Then
        str来源 = "1,2,3"
    ElseIf InStr(mstrPrivs, "门诊病人") > 0 Then
        str来源 = "1,3"
    ElseIf InStr(mstrPrivs, "住院病人") > 0 Then
        str来源 = "2,3"
    End If
    If InStr(mstrPrivs, "所有科室") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where B.部门ID = A.ID " & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And instr([1],','||B.服务对象||',')> 0 And B.工作性质 IN('检查')" & _
            " Order by A.编码"
    Else
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B,部门人员 C " & _
            " Where B.部门ID = A.ID And A.ID=C.部门ID And C.人员ID=" & UserInfo.ID & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And instr([1],','||B.服务对象||',')>0  And B.工作性质 IN('检查')" & _
            " Order by A.编码"
    End If
   
    cboDept.Clear
    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, "," & str来源 & ",")
    
    str科室IDs = GetUser科室IDs
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        
        If rsTmp!ID = UserInfo.部门ID Then cboDept.ListIndex = cboDept.NewIndex '直接所属优先
        If InStr("," & str科室IDs & ",", "," & rsTmp!ID & ",") > 0 And cboDept.ListIndex = -1 Then cboDept.ListIndex = cboDept.NewIndex
        
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then cboDept.ListIndex = 0
    InitDepts = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub tbrMain_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    If Button.Key = "补费" Then
'        PopupButtonMenu tbrMain, Button, mnuMoneyNew
    End If
End Sub

Private Sub tbrMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub txt标识号_Change()
    If txt标识号.Text = "" Then txt标识号.Tag = ""
End Sub

Private Sub txt标识号_GotFocus()
    Call zlControl.TxtSelAll(txt标识号)
End Sub

Private Sub txt标识号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call Form_KeyDown(vbKeyF3, 0)
    Else
        Select Case Split(Label1.Caption, "(")(0)
            Case "标识号"
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "就诊卡"
                Dim blnCard As Boolean
    
                '去掉磁卡的其他的特殊字符
                If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
                
                blnCard = InputIsCard(Me.txt标识号, KeyAscii)
                
                '刷卡完成或确认输入
                If blnCard And Len(Me.txt标识号.Text) = gbytCardLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txt标识号.Text <> "" Then
                    If KeyAscii <> 13 Then
                        Me.txt标识号.Text = Me.txt标识号.Text & Chr(KeyAscii)
                        Me.txt标识号.SelStart = Len(Me.txt标识号.Text)
                    End If
                    KeyAscii = 0
                    Me.txt标识号.Text = UCase(Me.txt标识号)
                    Me.txt标识号.SetFocus
                End If
            Case "单据号"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Not (txt标识号.Text = "" Or txt标识号.SelLength = Len(txt标识号.Text)) _
                    And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case "姓名"
            
        End Select
    End If
End Sub

Private Sub txt标识号_Validate(Cancel As Boolean)
    If Split(Label1.Caption, "(")(0) = "单据号" Then
        If IsNumeric(txt标识号.Text) Then
            txt标识号.Text = GetFullNO(txt标识号.Text, 0)
        End If
    End If
End Sub

Private Function LoadPatiList(Optional ByVal strKey As String = "") As Boolean
'功能：读取当前医技科室的执行医嘱(病人)清单
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSQLBak As String, i As Long, j As Long
    Dim objItem As ListItem, strPre As String
    Dim blnDo As Boolean, lngCount As Long
    Dim str来源 As String
    Dim strFilter As String
    Dim blnMoved As Boolean
    
    If Not lvwPati.SelectedItem Is Nothing And Len(strKey) = 0 Then
        strPre = lvwPati.SelectedItem.Key
    Else
        strPre = strKey
    End If
    blnMoved = MovedByDate(IIf(mdatFBegin = CDate(0), CDate(zlDatabase.Currentdate) - mBeforeDays, mdatFBegin))
    
    '清除界面数据
    mstrPrePati = ""
    lvwPati.ListItems.Clear
'    lblAdvice.Caption = ""
    
    On Error GoTo errH
        
    '病人来源权限:(1-门诊,2-住院,3-外来,4-体检)
    If InStr(mstrPrivs, "门诊病人") > 0 And InStr(mstrPrivs, "住院病人") > 0 Then
        str来源 = "1,2,3,4"
    ElseIf InStr(mstrPrivs, "门诊病人") > 0 Then
        str来源 = "1,4"
    ElseIf InStr(mstrPrivs, "住院病人") > 0 Then
        str来源 = "2"
    Else
        str来源 = "3"
    End If
        
    If Me.chkFilter.Value = 1 Then
        strFilter = " And D.姓名=[17] "
    Else
        '发送时间
        If mdatFEnd <> CDate(0) Then
            strFilter = " And " & IIf(Val(lvwPati.Tag) = 0, "A.发送时间", IIf(mDatType = 2, "A.发送时间", "A.首次时间")) & " Between [1] and [2] "
        Else '缺省查询条件
            strFilter = " And " & IIf(Val(lvwPati.Tag) = 0, "A.发送时间", IIf(mDatType = 2, "A.发送时间", "A.首次时间")) & " Between [1] and Sysdate "
        End If
        '单据号
        If mstrFNO <> "" Then
            strFilter = strFilter & " And A.NO= [3] "
        End If
        
        '病人科室
        If mlngF科室ID <> 0 Then
            strFilter = strFilter & " And B.病人科室ID+0= [4] "
        End If
        
        '病人来源
        If mstrF来源 <> "" Then
            strFilter = strFilter & " And instr([5],','||B.病人来源||',') > 0 "
        End If
        
        '病人标识
        
        If mdblF标识号 <> 0 Then
            strFilter = strFilter & " And Decode(B.病人来源,1,D.门诊号,2,D.住院号,NULL)= [6] "
        End If
        
        If mstrF就诊卡 <> "" Then
            strFilter = strFilter & " And D.就诊卡号 = [7] "
        End If
        
        If mstrF姓名 <> "" Then
    '        strFilter = strFilter & " And D.姓名 = [8] "
            strFilter = strFilter & " And Instr(D.姓名 , [8])>0 "
        End If
        
        If mstr标本部位 <> "" Then
            strFilter = strFilter & " And b.标本部位 = [16]"
        End If
        
        If mdblFChkNO <> 0 Then
            strFilter = strFilter & " And H.检查号=[13] "
        End If
    End If
    
    '中药煎法,用法各自独立显示一条
    '附加手术,检查部位执行科室及时间与主项目相同,不显示
    '手术麻醉执行科室为单独，需要显示
    '特殊医嘱不显示(虽然执行科室一般不会为医技科室)
'        " And Not (B.诊疗类别 IN('F','D') And B.相关ID is Not NULL)" & _
'        " And Not(B.诊疗类别='Z' And Nvl(C.操作类型,'0')<>'0')" & strWhere &
'        " And X.记录状态(+)<>2 And X.医嘱序号(+)=A.医嘱ID And X.序号(+)=1 And C.类别='D'" &
    If Len(Trim(frmPACSFilter.cboContent.Text)) = 0 Or Me.chkFilter.Value = 1 Then
        strSQL = _
            " Select Distinct /*多个收费项目*/ X.记录性质 as 费用性质,X.记录状态 as 费用状态," & _
            " A.医嘱ID,A.发送号,B.相关ID,B.序号,B.诊疗类别,B.诊疗项目ID," & _
            " A.首次时间 As 检查时间,A.发送时间 As 开嘱时间,A.NO," & _
            " A.记录性质,A.执行状态,A.计费状态,B.病人ID,B.主页ID,B.挂号单,B.病人科室ID,E.名称 as 科室,D.姓名," & _
            " Decode(B.病人来源,1,D.门诊号,2,D.住院号,4,D.门诊号,NULL) as 标识号,Nvl(D.费别,'普通') As 费别," & _
            " Decode(B.病人来源,1,'门诊',2,'住院',3,'外来',4,'体检') as 来源,C.名称 as 内容,A.执行间," & _
            " Nvl(Z.病历文件ID,0) As 单据ID,Nvl(A.报告ID,0) As 报告ID,Nvl(A.执行过程,0) As 执行过程," & _
            " B.医嘱内容,G.书写人,Decode(A.执行状态,1,Nvl(G.完成人,G.审阅人),NULL) As 审阅人,D.就诊卡号,0 As 转出,Nvl(H.检查号,'') As 检查号,Nvl(H.检查UID,'') As 检查UID,Nvl(A.结果阳性,0) As 阳性,B.紧急标志,Nvl(H.是否打印,0) As 是否打印" & _
            " From 病人医嘱发送 A,病人医嘱记录 B,诊疗项目目录 C,病人信息 D,部门表 E,病人病历记录 G,影像检查记录 H,病人费用记录 X,影像检查项目 Y,诊疗单据应用 Z" & _
            " Where A.医嘱ID=B.ID And A.报告ID=G.ID(+) And B.诊疗项目ID=C.ID And B.病人ID=D.病人ID" & _
            " And B.病人科室ID=E.ID And C.ID=Y.诊疗项目ID And A.医嘱ID=H.医嘱ID(+) And A.发送号=H.发送号(+)" & _
            " And C.ID=Z.诊疗项目ID(+) And (Z.应用场合=B.病人来源 Or B.病人来源=3 Or Z.诊疗项目ID Is Null)" & _
            " And instr([10],','||B.病人来源||',')> 0 And A.执行部门ID+0= [11] " & _
            " And B.相关ID is NULL " & _
            strFilter
    Else
        strSQL = _
            " Select Distinct /*多个收费项目*/ X.记录性质 as 费用性质,X.记录状态 as 费用状态," & _
            " A.医嘱ID,A.发送号,B.相关ID,B.序号,B.诊疗类别,B.诊疗项目ID," & _
            " A.首次时间 As 检查时间,A.发送时间 As 开嘱时间,A.NO," & _
            " A.记录性质,A.执行状态,A.计费状态,B.病人ID,B.主页ID,B.挂号单,B.病人科室ID,E.名称 as 科室,D.姓名," & _
            " Decode(B.病人来源,1,D.门诊号,2,D.住院号,4,D.门诊号,NULL) as 标识号,Nvl(D.费别,'普通') As 费别," & _
            " Decode(B.病人来源,1,'门诊',2,'住院',3,'外来',4,'体检') as 来源,C.名称 as 内容,A.执行间," & _
            " Nvl(Z.病历文件ID,0) As 单据ID,Nvl(A.报告ID,0) As 报告ID,Nvl(A.执行过程,0) As 执行过程," & _
            " B.医嘱内容,G.书写人,Decode(A.执行状态,1,Nvl(G.完成人,G.审阅人),NULL) As 审阅人,D.就诊卡号,0 As 转出,Nvl(H.检查号,'') As 检查号,Nvl(H.检查UID,'') As 检查UID,Nvl(A.结果阳性,0) As 阳性,B.紧急标志,Nvl(H.是否打印,0) As 是否打印" & _
            " From 病人医嘱发送 A,病人医嘱记录 B,诊疗项目目录 C,病人信息 D,部门表 E,病人病历记录 G,影像检查记录 H,病人费用记录 X,影像检查项目 Y,诊疗单据应用 Z," & _
            " 病人病历内容 I, 病人病历文本段 J" & _
            " Where A.医嘱ID=B.ID And A.报告ID=G.ID(+) And B.诊疗项目ID=C.ID And B.病人ID=D.病人ID" & _
            " And B.病人科室ID=E.ID And C.ID=Y.诊疗项目ID And A.医嘱ID=H.医嘱ID(+) And A.发送号=H.发送号(+)" & _
            " And C.ID=Z.诊疗项目ID(+)" & _
            " And G.ID = I.病历记录id And I.ID = J.病历id And I.标题文本 = [14] AND Instr(J.内容,[15])>0 And (Z.应用场合=B.病人来源 Or B.病人来源=3 Or Z.诊疗项目ID Is Null)" & _
            " And instr([10],','||B.病人来源||',')> 0 And A.执行部门ID+0= [11] " & _
            " And B.相关ID is NULL " & _
            strFilter
    End If
    Select Case Val(lvwPati.Tag)
        Case 0
'            strSQL = strSQL & " And (A.执行状态=0 Or (A.执行状态=3 And " & IIf(chk状态(0).Value, _
'                "Nvl(A.执行过程,0)<2", "A.执行过程=1") & "))"
            strSQL = strSQL & " And ((A.执行状态=3 Or A.执行状态=0) And " & Decode(cboState(0).ListIndex, _
                 1, "Nvl(A.执行过程,0)=0)", 2, "Nvl(A.执行过程,0)=1)", "Nvl(A.执行过程,0)<2)")
        Case 1
'            strSQL = strSQL & " And A.执行状态 =3 And A.执行过程=2"
            strSQL = strSQL & " And A.执行状态 =3 And A.执行过程=2" & Decode(cboState(1).ListIndex, _
                 1, " And Nvl(A.报告ID,0)=0", 2, " And Nvl(A.报告ID,0)>0", "")
        Case 2
'            strSQL = strSQL & " And ((A.执行状态 =3 And A.执行过程>2) Or " & IIf(chk状态(2).Value, _
'                "A.执行状态=1", "1=2") & ")"
            strSQL = strSQL & " And " & Decode(cboState(2).ListIndex, _
                1, "A.执行状态 =3 And A.执行过程 =3", _
                2, "A.执行状态 =3 And A.执行过程 =4", _
                3, "A.执行状态 =3 And A.执行过程 =5", _
                4, "A.执行状态 =1", _
                "((A.执行状态 =3 And A.执行过程>2) Or A.执行状态=1)")
    End Select
'    strSQL = strSQL & " And A.NO=X.NO(+) And A.记录性质=Decode(X.记录性质(+),0,1,X.记录性质(+))" & _
'        " And X.记录状态(+)<>2 And X.医嘱序号(+)=A.医嘱ID And X.序号(+)=1" & _
'        IIf(blnIfOnlyShow, " And A.执行间= [12] ", "")
    strSQL = strSQL & " And A.NO=X.NO(+) And A.记录性质=Decode(X.记录性质(+),0,1,X.记录性质(+))" & _
        " And X.记录状态(+)<>2 And X.序号(+)=1" & _
        IIf(blnIfOnlyShow, " And A.执行间= [12] ", "")
    '如果有数据转出则还要检索后备表
    If blnMoved Then
        strSQLBak = strSQL
        strSQLBak = Replace(strSQLBak, "0 As 转出", "1 As 转出")
        strSQLBak = Replace(strSQLBak, "病人医嘱记录", "H病人医嘱记录")
        strSQLBak = Replace(strSQLBak, "病人医嘱发送", "H病人医嘱发送")
        strSQLBak = Replace(strSQLBak, "病人病历记录", "H病人病历记录")
        strSQLBak = Replace(strSQLBak, "病人费用记录", "H病人费用记录")
        strSQLBak = Replace(strSQLBak, "病人病历内容", "H病人病历内容")
        strSQLBak = Replace(strSQLBak, "病人病历文本段", "H病人病历文本段")
        strSQL = strSQL & " Union ALL " & strSQLBak
    End If
    strSQL = strSQL & " Order by 检查时间 Desc,开嘱时间,病人ID,序号"
    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, CDate(Format(mdatFBegin, "yyyy-MM-dd HH:mm:00")), CDate(Format(mdatFEnd, "yyyy-MM-dd HH:mm:59")), _
    mstrFNO, mlngF科室ID, "," & mstrF来源 & ",", mdblF标识号, mstrF就诊卡, mstrF姓名, mBeforeDays, "," & str来源 & ",", _
    cboDept.ItemData(cboDept.ListIndex), mstrRoom, mdblFChkNO, frmPACSFilter.cboItem.Text, frmPACSFilter.cboContent.Text, mstr标本部位, mstrPatiName)
    
    lngCount = 0
    For i = 1 To rsTmp.RecordCount
        '是否只显示已收费的病人
        '1.只管主费用,不判断附加费用.也不管主费用退费
        '2.记帐单据当作已收费显示
        '3.无需计费或尚未生成主费用的也显示
        blnDo = True
        If mnuViewCharge.Checked Then
            If Nvl(rsTmp!费用性质, 0) = 1 And Nvl(rsTmp!费用状态, 0) <> 1 Then
                blnDo = False '0-未收费或未审核,3-已退费或已销帐;
            End If
        End If
        
        If blnDo Then
            Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!医嘱ID & "_" & rsTmp!发送号, Nvl(rsTmp!来源), , IIf(Len(rsTmp!检查UID) > 0, 5, 0))
            objItem.SubItems(1) = Nvl(rsTmp!NO)
            objItem.SubItems(2) = LTrim(RTrim(Nvl(rsTmp!姓名)))
            objItem.SubItems(3) = Nvl(rsTmp!内容)
            objItem.SubItems(4) = IIf(rsTmp!执行状态 = 1, "执行完", _
                IIf(Nvl(rsTmp!执行过程, 0) = 2 And Nvl(rsTmp!报告ID, 0) > 0, "报告", _
                Decode(Nvl(rsTmp!执行过程, 0), 0, "未报到", 1, "已报到", 2, "检查中", 3, "检查完", 4, "报告", 5, "驳回", 6, "报告完")))
            objItem.SubItems(5) = Nvl(rsTmp!科室)
            objItem.SubItems(6) = Nvl(rsTmp!标识号)
            objItem.SubItems(7) = Nvl(rsTmp!费别)
            objItem.SubItems(8) = Format(rsTmp!检查时间, "yy-MM-dd HH:mm")
            objItem.SubItems(9) = Nvl(rsTmp!执行间)
            objItem.SubItems(10) = rsTmp!医嘱ID '& "_" & rsTmp!发送号
'            objItem.SmallIcon = IIf(objItem.SubItems(4) = "拒执行", "未执行", objItem.SubItems(4))
            objItem.SubItems(11) = GetPart(Nvl(rsTmp!医嘱内容))
            objItem.SubItems(12) = Nvl(rsTmp!书写人)
            objItem.SubItems(13) = Nvl(rsTmp!审阅人)
            objItem.SubItems(14) = Nvl(rsTmp!检查号)
            objItem.SubItems(15) = IIf(Nvl(rsTmp!紧急标志, 0) = 1, "√", "")
            objItem.SubItems(16) = IIf(Nvl(rsTmp!是否打印, 0) = 0, "", "√")
            objItem.SubItems(17) = Format(rsTmp!开嘱时间, "yy-MM-dd HH:mm")
            
            Select Case objItem.SubItems(4)
            Case "已报到", "检查中", "检查完"
                objItem.ForeColor = 0 '黑色
            Case "未报到", "执行完" '灰色
                objItem.ForeColor = &H808080
            Case "驳回" '棕色
                objItem.ForeColor = &H40C0&
            Case "报告", "报告完" '兰色
                objItem.ForeColor = &HC00000
            End Select
            For j = 1 To Me.lvwPati.ColumnHeaders.Count - 1
                objItem.ListSubItems(j).ForeColor = objItem.ForeColor
            Next
            
            '存放附加数据
            objItem.Tag = rsTmp!病人ID & "_" & Nvl(rsTmp!主页ID, 0) & "_" & Nvl(rsTmp!挂号单)
            
            objItem.ListSubItems(1).Tag = Nvl(rsTmp!相关ID, 0)
            objItem.ListSubItems(2).Tag = rsTmp!诊疗类别
            objItem.ListSubItems(3).Tag = Nvl(rsTmp!执行状态, 0)
            objItem.ListSubItems(4).Tag = Nvl(rsTmp!计费状态, 0)
            objItem.ListSubItems(5).Tag = Nvl(rsTmp!记录性质, 1)
            objItem.ListSubItems(6).Tag = Nvl(rsTmp!病人科室ID, 0)
            objItem.ListSubItems(7).Tag = Nvl(rsTmp!诊疗项目ID, 0)
            objItem.ListSubItems(8).Tag = Nvl(rsTmp!单据ID, 0) & "|" & Nvl(rsTmp!报告ID, 0)
            objItem.ListSubItems(9).Tag = Nvl(rsTmp!执行过程, 0)
            objItem.ListSubItems(10).Tag = Nvl(rsTmp!就诊卡号)
            objItem.ListSubItems(11).Tag = Nvl(rsTmp!转出, 0)
            
            objItem.ListSubItems(2).ReportIcon = IIf(rsTmp!阳性 = 1, 6, 7)
            
            If objItem.Key = strPre Then objItem.Selected = True
            
            lngCount = lngCount + 1
        End If
        rsTmp.MoveNext
    Next
    
    If Not lvwPati.SelectedItem Is Nothing Then
        Call lvwPati_ItemClick(lvwPati.SelectedItem)
        lvwPati.SelectedItem.EnsureVisible
        If (lvwPati.SelectedItem.Index <> lvwPati.ListItems.Count) Then
            lvwPati.ListItems(lvwPati.SelectedItem.Index + 1).EnsureVisible
        End If
    Else
        Call tabFile_Click
    End If
    
    Me.stbThis.Panels(2).Text = IIf(lngCount = 0, "没有", "共有：" & lngCount & " 项") & Decode(lvwPati.Tag, _
        "0", "待执行的检查", _
        "1", "正在执行的检查", _
        "2", "已完成的检查")
        
    mstr标本部位 = ""
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowMenu()
    Dim blnEnable As Boolean, i As Integer
    Dim int病人来源 As Integer
    
    On Error Resume Next
    blnEnable = False
    
    mnuExecFunc(0).Enabled = True
    If Not lvwPati.SelectedItem Is Nothing Then
        Me.mnuExec.Enabled = True
        Me.mnuRep.Enabled = True
        Me.mnuPFile.Enabled = True
        Me.mnuReq.Enabled = True
        Me.mnuAdvice.Enabled = True
        Me.mnuMoney.Enabled = True
        With lvwPati.SelectedItem
'            mnuExecFunc(4).Enabled = Not (InStr("1", .ListSubItems(3).Tag) > 0 Or _
'                 (.ListSubItems(3).Tag = 3 And Val(.ListSubItems(9).Tag) > 1))
            mnuExecFunc(1).Enabled = InStr("0", .ListSubItems(3).Tag) > 0
            
            mnuExecFunc(4).Enabled = Not (InStr("1", .ListSubItems(3).Tag) > 0)
            
            mnuExecFunc(4).Caption = IIf(.ListSubItems(3).Tag = 3 And Val(.ListSubItems(9).Tag) > 1, "修改信息(&A)", "开始检查(&A)")
            
            mnuExecFunc(5).Enabled = Not (.ListSubItems(3).Tag <> "3" Or .ListSubItems(9).Tag <> "2")
            tbrMain.Buttons("采集").Enabled = mnuExecFunc(5).Enabled
            
            mnuExecFunc(6).Enabled = Not (.ListSubItems(3).Tag <> "3" Or .ListSubItems(9).Tag <> "2")
            
            mnuExecFunc(7).Enabled = Not (.ListSubItems(3).Tag <> "3" Or .ListSubItems(9).Tag <> "2")
            
            mnuExecFunc(9).Enabled = Not (InStr("0,3", .ListSubItems(3).Tag) = 0 Or InStr("1,2", .ListSubItems(9).Tag) = 0)
            
            mnuExecFunc(10).Enabled = Not (InStr("0,3", .ListSubItems(3).Tag) = 0 Or InStr("1,2", .ListSubItems(9).Tag) = 0)
            
            mnuExecFunc(11).Enabled = Not (InStr("0,3", .ListSubItems(3).Tag) = 0 Or InStr("1,2", .ListSubItems(9).Tag) = 0)
            
            mnuExecFunc(12).Enabled = Not (InStr("0,3", .ListSubItems(3).Tag) = 0 Or InStr("1,2", .ListSubItems(9).Tag) = 0)
            
            mnuExecFunc(14).Enabled = Not (.ListSubItems(3).Tag <> "3" Or .ListSubItems(9).Tag <> "2")
            
            mnuExecFunc(15).Enabled = Not (.ListSubItems(3).Tag <> "3" Or Val(.ListSubItems(9).Tag) < 3)
            
            mnuExecFunc(17).Enabled = mnuRepFunc(0).Enabled
                
'            mnuRepFunc(0).Enabled = Not (InStr("1,3", .ListSubItems(3).Tag) = 0 Or _
'                InStr("3,4,5,6", .ListSubItems(9).Tag) = 0)
'            mnuRepFunc(3).Enabled = Not (.ListSubItems(3).Tag <> "3" Or _
'                InStr("4", .ListSubItems(9).Tag) = 0)
            mnuRepFunc(3).Enabled = Not (.ListSubItems(3).Tag <> "3")
            mnuExecFunc(18).Enabled = mnuRepFunc(3).Enabled
'            mnuRepFunc(4).Enabled = Not (InStr("1,3", .ListSubItems(3).Tag) = 0 Or _
'                InStr("4,6", .ListSubItems(9).Tag) = 0)
            mnuRepFunc(4).Enabled = Not InStr("1,3", .ListSubItems(3).Tag) = 0
        
            If Val(Split(.Tag, "_")(1)) = 0 And Len(Split(.Tag, "_")(2)) = 0 Then Me.mnuAdvice.Visible = False
            int病人来源 = IIf(.Text = "住院", 2, 1)
            If TabFile.SelectedItem.Key = "医嘱" Then
                mnuAdviceFunc(4).Visible = int病人来源 = 2
                mnuAdviceFunc(5).Visible = int病人来源 = 2
                mnuAdviceFunc(6).Visible = Not (int病人来源 = 2)
                mnuAdviceFunc(7).Visible = Not (int病人来源 = 2)
            Else
                mnuAdvice.Visible = False
            End If
        End With
        Me.mnuViewInfo.Enabled = True
    Else
        Me.mnuExec.Enabled = True
        For i = 1 To mnuExecFunc.Count - 1
            mnuExecFunc(i).Enabled = False
        Next
        Me.mnuRep.Enabled = False
        Me.mnuPFile.Enabled = False
        Me.mnuReq.Enabled = False
        Me.mnuAdvice.Enabled = False
        Me.mnuMoney.Enabled = False
        Me.mnuViewInfo.Enabled = False
    
        tbrMain.Buttons("采集").Enabled = False
    End If
    
    With Me.tbrMain.Buttons
        For i = 1 To .Count
            Select Case .Item(i).Description
                Case "申请"
                    .Item(i).Enabled = mnuMoney.Enabled
                Case "医嘱"
                    .Item(i).Enabled = mnuAdvice.Enabled
                    .Item(i).Visible = mnuAdvice.Visible
                Case "病历"
                    .Item(i).Enabled = mnuPFile.Enabled
                Case "影像"
                    .Item(i).Enabled = mnuRep.Enabled
            End Select
        Next
    End With

    Me.tbrMain.Buttons("报告").Enabled = mnuRepFunc(0).Enabled
    Me.tbrMain.Buttons("审核").Enabled = mnuRepFunc(3).Enabled
    Me.tbrMain.Buttons("驳回").Enabled = mnuRepFunc(4).Enabled
End Sub

Private Sub OutputList(bytStyle As Byte)
'功能: 输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrintLvw

    On Error Resume Next
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    
    Select Case Me.TabFile.SelectedItem.Key
        Case "申请"
            mfrmActive.zlPrint bytStyle
            Exit Sub
        Case "医嘱"
            Select Case bytStyle
                Case 1
                    mfrmActive.zlPrint
                Case 2
                    mfrmActive.zlPreview
                Case 3
                    mfrmActive.zlExcel
            End Select
            Exit Sub
    End Select

    Set objOut.Body.objData = Me.lvwPati
    objOut.Title.Text = Decode(Val(lvwPati.Tag), 0, "待执行", 1, "正进行", 2, "已完成") & _
        "检查清单"
    If bytStyle = 1 Then
        bytStyle = zlPrintAsk(objOut)
        If bytStyle <> 0 Then zlPrintOrViewLvw objOut, bytStyle
    Else
        zlPrintOrViewLvw objOut, bytStyle
    End If
End Sub

Private Sub BeginShowProgress(ByVal strCaption As String)
    With pgbLoad
        .Left = stbThis.Panels(2).Left + Me.TextWidth(strCaption) + 200
        .Top = stbThis.Top + (stbThis.Height - .Height) / 2
        .Width = stbThis.Panels(2).Width + stbThis.Panels(2).Left - .Left
        .Value = 0
        
        stbThis.Panels(2).Text = strCaption
        .Visible = Me.stbThis.Visible: Me.Refresh
    End With
End Sub

'构造可增加的病历文件菜单
Private Sub AddFileList()
    Dim rsFileList As ADODB.Recordset
    Dim i As Integer, iNum As Integer
    
    '清除文件清单
    iNum = FileList.Count
    FileList(0).Visible = True
    For i = 1 To iNum - 1
        Unload FileList(i)
    Next
    
    Set rsFileList = GetPatientFileList(UserInfo.部门ID, 0)
    If Not rsFileList Is Nothing Then
        i = 1
        Do While Not rsFileList.EOF
            Load FileList(FileList.Count)
            With FileList(FileList.Count - 1)
                .Caption = "&" & i & " " & rsFileList("名称")
                .Tag = "O" & rsFileList("ID")
                .Enabled = True
                .Visible = True
            End With
            
            i = i + 1
            rsFileList.MoveNext
        Loop
        
        On Error Resume Next
        FileList(0).Visible = False
    End If
    Set rsFileList = GetPatientFileList(UserInfo.部门ID, 1)
    If Not rsFileList Is Nothing Then
        i = 1
        Do While Not rsFileList.EOF
            Load FileList(FileList.Count)
            With FileList(FileList.Count - 1)
                .Caption = "&" & i & " " & rsFileList("名称")
                .Tag = "I" & rsFileList("ID")
                .Enabled = True
                .Visible = True
            End With
            
            i = i + 1
            rsFileList.MoveNext
        Loop
        
        On Error Resume Next
        FileList(0).Visible = False
    End If
End Sub

Private Sub ShowAddFileMenu(ByVal SrcType As Integer)
    Dim i As Integer, iNum As Integer
    Dim blnOutVisible As Boolean
    
    If SrcType = 1 Then blnOutVisible = True
    
    '清除文件清单
    iNum = FileList.Count
    For i = 1 To iNum - 1
        FileList(i).Visible = True
    Next
    For i = 1 To iNum - 1
        If FileList(i).Tag Like "O*" Then
            FileList(i).Visible = blnOutVisible
        Else
            FileList(i).Visible = Not blnOutVisible
        End If
    Next
End Sub

Private Function LoadBillList() As Boolean
'功能：读取当前可用的辅诊单据清单
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objMenu As Menu
    
    On Error GoTo errH
    
    '清除现有单据清单
    For i = Me.ReqList.UBound To 0 Step -1
        Me.ReqList(i).Tag = ""
        If i = 0 Then
            ReqList(i).Caption = "<无可用单据>"
        Else
            Unload ReqList(i)
        End If
    Next
    
    '加载可用单据
    strSQL = "Select Distinct A.ID,A.编号,A.名称,A.说明" & _
        " From 病历文件目录 A,病历文件组成 B" & _
        " Where A.种类=5 And A.前提 IN(2,3)" & _
        " And A.ID=B.病历文件ID And B.填写时机 IN(1,2)" & _
        " Order by A.编号"
    Call OpenRecord(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        If i <> 1 Then Load ReqList(ReqList.UBound + 1)
        Set objMenu = ReqList(ReqList.UBound)
        objMenu.Caption = rsTmp!名称
        If i <= 10 Then
            objMenu.Caption = objMenu.Caption & "(&" & i - 1 & ")"
        ElseIf i <= 36 Then
            objMenu.Caption = objMenu.Caption & "(&" & Chr(i - 11 + Asc("A")) & ")"
        End If
        objMenu.Tag = rsTmp!ID: objMenu.Enabled = True
        rsTmp.MoveNext
    Next
    LoadBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPart(ByVal strAdvice As String) As String
'功能：根据医嘱内容获取标本部位
    Dim iPos As Integer, iPos1 As Integer
    iPos = 0
    Do While True
        iPos1 = InStr(iPos + 1, strAdvice, "(")
        If iPos1 = 0 Then Exit Do
        iPos = iPos1
    Loop
    If iPos > 0 Then
        GetPart = Mid(strAdvice, iPos + 1, Len(strAdvice) - iPos - 1)
    Else
        GetPart = ""
    End If
End Function
Public Function GetReprotFrm() As Form
    Set GetReprotFrm = mfrmRepEdit
End Function

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.Hwnd)
End Sub


Private Sub subCancelSeriesRelate(lngAdviceNo As Long, lngSendNO As Long, strSeriesNo As String)
'-----------------------------------------------------------------------------
'功能:取消序列图象的关联
'修改人:黄捷
'修改日期:2007-1-30
'-----------------------------------------------------------------------------
    
    Dim mcnFTP As New clsFtp
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strCachePath As String
    Dim strCacheFileName As String
    Dim objFile As New Scripting.FileSystemObject
    Dim imgs As New DicomImages
    Dim img As New DicomImage
    Dim strNewStudyUID As String    '新生成的检查UID
    Dim strOldStudyUID As String    '图象里面原来的检查UID
    Dim strDBStudyUID As String     '数据库中保存的检查UID，跟图象存储路径相关
    Dim strMoveFiles As String  '存储需要移动的图象文件名，使用“|”分隔
    Dim blnNoImage As Boolean   '1没有图象，直接读取数据库信息。0有图象，使用图象信息
    
    '图像中的病人基本信息
    Dim strModality As String
    Dim strPatientID As String
    Dim strPatientName As String
    Dim strSex As String
    Dim strAge As String
    Dim strDateOfBirth As String
    Dim strManufacturer As String
    Dim strReceiveDateTime As String
    
    
    
    '查找序列中第一个图像的 病人ID，英文名，性别，年龄，出生日期，检查UID，检查设备，接收时间
    strCachePath = App.Path & "\TmpImage\"
    strSQL = "Select A.图像号,D.用户名 As User1,D.密码 As Pwd1,a.图像UID, " & _
        "D.IP地址 As Host1,c.检查uid," & _
        "'/'||D.Ftp目录||'/' As Root1,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/'||A.图像UID As URL1,d.设备号 as 设备号1, " & _
        "E.用户名 As User2,E.密码 As Pwd2," & _
        "E.IP地址 As Host2," & _
        "'/'||E.Ftp目录||'/' As Root2,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/'||A.图像UID As URL2 , e.设备号 as 设备号2 " & _
        "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) " & _
        "And A.序列UID= [1] Order By A.图像号"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, strSeriesNo)
    
    If Not rsTmp.EOF Then   '序列中存在图象
        strDBStudyUID = Nvl(rsTmp("检查uid"))
        '下载图象
        If rsTmp("设备号1") <> "" Then
            mcnFTP.FuncFtpConnect rsTmp("Host1"), rsTmp("User1"), rsTmp("Pwd1")
            strCacheFileName = strCachePath & objFile.GetFileName(rsTmp("URL1"))
            mcnFTP.FuncDownloadFile objFile.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL1")), strCacheFileName, objFile.GetFileName(rsTmp("URL1"))
            mcnFTP.FuncFtpDisConnect
        ElseIf rsTmp("设备号2") <> "" Then
            mcnFTP.FuncFtpConnect rsTmp("Host2"), rsTmp("User2"), rsTmp("Pwd2")
            strCacheFileName = strCachePath & objFile.GetFileName(rsTmp("URL2"))
            mcnFTP.FuncDownloadFile objFile.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL2")), strCacheFileName, objFile.GetFileName(rsTmp("URL2"))
            mcnFTP.FuncFtpDisConnect
        End If
        '读取图象
        If Dir(strCacheFileName) <> vbNullString Then
            Set img = imgs.ReadFile(strCacheFileName)
            '-----------是否使用变量将图象基本信息读取出来？
            strOldStudyUID = img.StudyUID
            strModality = GetImageAttribute(img.Attributes, ATTR_影像类别)
            strPatientID = img.PatientID
            strPatientName = img.Name
            strSex = img.Sex
            If IsDate(img.DateOfBirthAsDate) Then
                strAge = CStr(Year(Date) - Year(img.DateOfBirthAsDate))
                strDateOfBirth = Format(img.DateOfBirthAsDate, "YYYY-MM-DD")
            Else
                strAge = "": strDateOfBirth = ""
            End If
            strManufacturer = GetImageAttribute(img.Attributes, ATTR_检查设备)
            strReceiveDateTime = GetImageAttribute(img.Attributes, ATTR_检查日期) & " " & _
                        Format(GetImageAttribute(img.Attributes, ATTR_检查时间), "HH:MM")
            '删除临时图象
            Set img = Nothing
            imgs.Remove (1)
            objFile.DeleteFile strCacheFileName
        Else
            '如果第一个图象下载不正确，读取数据库信息
            blnNoImage = True
        End If
    Else
        '序列中没有图象，只接使用本序列在数据库中的值
        blnNoImage = True
    End If
    
    '对于没有图象信息可读取的情况，直接读取数据库中的信息
    If blnNoImage = True Then
        strSQL = "select a.影像类别,a.检查号,a.姓名,a.英文名,a.性别,a.年龄,a.出生日期,a.检查uid," & _
                " a.检查设备,a.接收日期 from 影像检查记录 a where a.医嘱id =[1] and a.发送号 =[2]"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngAdviceNo, lngSendNO)
        If Not rsTmp.EOF Then
            strOldStudyUID = Nvl(rsTmp("检查uid"))
            strDBStudyUID = Nvl(rsTmp("检查uid"))
            strModality = Nvl(rsTmp("影像类别"))
            strPatientID = Nvl(rsTmp("检查号"))
            strPatientName = Nvl(rsTmp("英文名"))
            strSex = Nvl(rsTmp("性别"))
            strAge = Nvl(rsTmp("年龄"))
            strDateOfBirth = Nvl(rsTmp("出生日期"), "1899-12-30")
            strManufacturer = Nvl(rsTmp("检查设备"))
            strReceiveDateTime = Nvl(rsTmp("接收日期"))
        End If
    End If
    '组织图象文件名称串
    strSQL = "select 图像UID from 影像检查序列 a,影像检查图象 b where a.序列UID =[1] and a.序列UID = b.序列UID"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, strSeriesNo)
    If Not rsTmp.EOF Then
        strMoveFiles = rsTmp(0)
        rsTmp.MoveNext
        While Not rsTmp.EOF
            strMoveFiles = strMoveFiles & "|" & rsTmp(0)
            rsTmp.MoveNext
        Wend
    End If
    
    '如果检查UID跟数据库中现存的检查UID相同，则创建新的检查UID，且修改图像FTP路径
    strNewStudyUID = funGetStudyUID(strOldStudyUID)
    If strNewStudyUID <> strDBStudyUID Then
        Call MergeImageFiles(strDBStudyUID, strNewStudyUID, Format(strReceiveDateTime, "YYYY-MM-DD"), strMoveFiles)
    End If
    
    '修改数据库，正常记录转成临时记录
        strSQL = "ZL_影像检查_PhotoCancel(" & lngAdviceNo & "," & lngSendNO & ",'" & strNewStudyUID & "','" & _
                  strSeriesNo & "','" & strModality & "'," & Val(strPatientID) & ",'" & _
                  strPatientName & "','" & strSex & "','" & strAge & "'," & _
                  IIf(Len(strDateOfBirth) = 0, "null", "to_date('" & strDateOfBirth & "','YYYY-MM-DD')") & _
                  ",'" & strManufacturer & "',to_date('" & strReceiveDateTime & "','YYYY-MM-DD HH24:MI:SS'))"
                  
        ExecuteProc strSQL, Me.Caption
End Sub

Private Function funGetStudyUID(ByVal strOldStudyUID As String) As String
'-----------------------------------------------------------------------------
'功能:查询数据库，判断当前图像的检查UID是否已经存在于正常表和临时表中，
'     如果存在，则在检查UID后面增加后缀，不存在则直接返回输入的检查UID
'修改人:黄捷
'修改日期:2007-1-27
'-----------------------------------------------------------------------------
    '
    Dim rsMatch As New ADODB.Recordset
    
    funGetStudyUID = strOldStudyUID
    gstrSQL = "select 检查UID from 影像检查记录 where 检查UID = [1]" & _
              " Union All Select 检查UID from 影像临时记录 where 检查UID = [1]"
    Set rsMatch = OpenSQLRecord(gstrSQL, "PACS图像保存", strOldStudyUID)
    If Not rsMatch.EOF Then
        '创建一个新的检查UID
        gstrSQL = "Select 影像检查UID序号_ID.Nextval From Dual"
        Set rsMatch = OpenSQLRecord(gstrSQL, "PACS图像保存")
        If Len(strOldStudyUID) <= 55 Then
            funGetStudyUID = strOldStudyUID & ".A" & rsMatch(0)
        Else
            funGetStudyUID = Left(strOldStudyUID, 55) & ".A" & rsMatch(0)
        End If
    End If
End Function


Public Function GetImageAttribute(objAttr As DicomAttributes, ByVal AttrName As String) As Variant
'-----------------------------------------------------------------------------
'功能:提取DICOM属性集中的指定属性值
'修改人:黄捷
'修改日期:2007-2-6
'-----------------------------------------------------------------------------
    Dim AttrTag() As String
    
    GetImageAttribute = ""
    AttrTag = Split(AttrName, ":")
    If objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).Exists Then
        GetImageAttribute = Nvl(objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).Value)
    End If
End Function
