VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageExecute 
   AutoRedraw      =   -1  'True
   Caption         =   "执行登记管理"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11490
   Icon            =   "frmManageExecute.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   5835
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageExecute.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15187
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
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   1376
      _CBWidth        =   11490
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   5910
      NewRow1         =   0   'False
      Child2          =   "picCondition"
      MinWidth2       =   3105
      MinHeight2      =   495
      Width2          =   3105
      NewRow2         =   0   'False
      Caption3        =   "科室"
      Child3          =   "cboUnit"
      MinWidth3       =   1605
      MinHeight3      =   300
      Width3          =   1605
      NewRow3         =   0   'False
      Begin VB.PictureBox picCondition 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   6045
         ScaleHeight     =   495
         ScaleWidth      =   3105
         TabIndex        =   5
         Top             =   135
         Width           =   3105
         Begin VB.CheckBox chkAuto 
            Caption         =   "主从项一起选"
            Height          =   375
            Left            =   0
            TabIndex        =   7
            Top             =   48
            Width           =   855
         End
         Begin VB.TextBox txtValue 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   6
            ToolTipText     =   "定位F4"
            Top             =   55
            Width           =   1425
         End
         Begin VB.Label lblKind 
            Caption         =   "↓单据号"
            Height          =   225
            Left            =   885
            TabIndex        =   8
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   9795
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1605
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   4
         Top             =   30
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Description     =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "登记"
               Key             =   "Log"
               Description     =   "登记"
               Object.ToolTipText     =   "执行登记"
               Object.Tag             =   "登记"
               ImageKey        =   "Log"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "取消"
               Key             =   "Cancel"
               Description     =   "取消"
               Object.ToolTipText     =   "取消登记"
               Object.Tag             =   "取消"
               ImageKey        =   "Cancel"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit_"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "View"
               Description     =   "查看"
               Object.ToolTipText     =   "查看登记"
               Object.Tag             =   "查看"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "View_"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "全选"
               Key             =   "SelAll"
               Description     =   "全选"
               Object.ToolTipText     =   "全部选择"
               Object.Tag             =   "全选"
               ImageKey        =   "SelAll"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "全清"
               Key             =   "Clear"
               Description     =   "全清"
               Object.ToolTipText     =   "全部清除"
               Object.Tag             =   "全清"
               ImageKey        =   "Clear"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Clear_"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "按设置条件重新筛选记录"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Go"
               Description     =   "定位"
               Object.ToolTipText     =   "定位到满足条件的记录上"
               Object.Tag             =   "定位"
               ImageKey        =   "Go"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   4995
      Left            =   75
      TabIndex        =   0
      Top             =   795
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   8811
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      MergeCells      =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmManageExecute.frx":0C3E
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   5205
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":0F58
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":1172
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":138C
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":15A6
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":1D20
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":1F3A
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":2154
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":236E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":2588
            Key             =   "Log"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":2C82
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":337C
            Key             =   "SelAll"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":3596
            Key             =   "Clear"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   4620
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":37B0
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":39CA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":3BE4
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":3DFE
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":4578
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":4792
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":49AC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":4BC6
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":4DE0
            Key             =   "Log"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":54DA
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":5BD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":5DEE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   3105
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExecute.frx":6008
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   3690
      Top             =   90
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
            Picture         =   "frmManageExecute.frx":68E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFile_PreView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSetup 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditLog 
         Caption         =   "执行登记(&A)"
      End
      Begin VB.Menu mnuEditCancel 
         Caption         =   "取消执行(&C)"
      End
      Begin VB.Menu mnuEditView 
         Caption         =   "查看登记(&V)"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelAll 
         Caption         =   "全部选择(&S)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "全部清除(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPrint 
         Caption         =   "打印票据"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "报表(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolUnit 
            Caption         =   "执行科室(&U)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_Tlb_1 
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
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowHead 
         Caption         =   "单据头(&H)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "定位(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewreFlash 
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
   Begin VB.Menu mnuIDKind 
      Caption         =   "身份类别"
      Visible         =   0   'False
      Begin VB.Menu mnuIDKinds 
         Caption         =   "单据号"
         Index           =   0
      End
      Begin VB.Menu mnuIDKinds 
         Caption         =   "门诊号"
         Index           =   1
      End
      Begin VB.Menu mnuIDKinds 
         Caption         =   "住院号"
         Index           =   2
      End
      Begin VB.Menu mnuIDKinds 
         Caption         =   "姓名"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmManageExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mrsList As ADODB.Recordset  '单据列表
Private mstrFilter As String, mstrPreNO As String
Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mlngDeptID As Long
Private Type Type_SQLCondition
    Default As Boolean          '是否是缺省进入，此时没有条件值,缺省值在mstrFilter中
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    State As Byte
    Operator As String
    ID As Double
    Patient As String
End Type
Private SQLCondition As Type_SQLCondition

Private Const COL_科室 = 14
Private Const COL_状态 = 15

Private mstrPrivs As String     '保存当前模块的授权功能
Private mlngModul As Long
Private mblnNOMoved As Boolean '记录当前选择的单据是否是在后备数据表中

Private mrsWarn As ADODB.Recordset

Private Sub cboUnit_Click()
    
    If cboUnit.ItemData(cboUnit.ListIndex) = mlngDeptID Then Exit Sub
    mlngDeptID = cboUnit.ItemData(cboUnit.ListIndex)
        
    If Visible Then Call ShowBills(mstrFilter)
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub chkAuto_Click()
    zlDatabase.SetPara "主从项目同时选择", chkAuto.Value, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub Form_Activate()
    Call InitLocPar(mlngModul)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub lblKind_Click()
    PopupMenu mnuIDKind, 2
End Sub

Private Sub mnuIDKinds_Click(Index As Integer)
    Dim i As Long
    
    For i = 0 To mnuIDKinds.UBound
        mnuIDKinds(i).Checked = i = Index
    Next
    
    lblKind.Caption = "↓" & Choose(Index + 1, "单据号", "门诊号", "住院号", "姓名")
End Sub


Private Sub txtvalue_KeyPress(KeyAscii As Integer)
    If mnuIDKinds(1).Checked Or mnuIDKinds(2).Checked Then
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub


Private Sub txtvalue_GotFocus()
    zlControl.TxtSelAll txtValue
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtValue.Text <> "" Then
            If mnuIDKinds(0).Checked And IsNumeric(txtValue.Text) Then txtValue.Text = GetFullNO(txtValue.Text, 0)
        End If
        txtValue.Text = UCase(Trim(txtValue.Text))
        
        With frmExecuteFilter
            If mnuIDKinds(0).Checked Then
                .txtNOBegin.Text = txtValue.Text
                .txtNoEnd.Text = ""
                .txt标识号.Text = ""
                .txt姓名.Text = ""
            ElseIf mnuIDKinds(1).Checked Or mnuIDKinds(2).Checked Then
                .txtNOBegin.Text = ""
                .txtNoEnd.Text = ""
                .txt标识号.Text = txtValue.Text
                .txt姓名.Text = ""
            ElseIf mnuIDKinds(3).Checked Then
                .txtNOBegin.Text = ""
                .txtNoEnd.Text = ""
                .txt标识号.Text = ""
                .txt姓名.Text = txtValue.Text
            End If
            .MakeFilter
        End With
        
        Call FindBills
        
        zlControl.TxtSelAll txtValue
    ElseIf KeyCode = vbKeyF4 Then
        Dim i As Integer
        
        For i = 0 To mnuIDKinds.Count - 1
            If mnuIDKinds(i).Checked = True Then Exit For
        Next
        If i >= mnuIDKinds.Count - 1 Then
            i = 0
        Else
            i = i + 1
        End If
        Call mnuIDKinds_Click(i)
    End If
End Sub
Private Function Is门诊费用(ByVal lngRow As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：当前记录是否为门诊费用
    '入参：lngRow:指定行的数据
    '出参：
    '返回：是门诊费用,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-03-08 15:45:31
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim int门诊标志 As Integer, int记录性质 As Integer, int主页ID As Integer
    int门诊标志 = GetColValue(lngRow, "门诊标志")
    int记录性质 = GetColValue(lngRow, "记录性质")
    int主页ID = GetColValue(lngRow, "主页ID")
    '1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检
    If int门诊标志 = 1 Or int门诊标志 = 4 Or int记录性质 = 1 Then Is门诊费用 = True: Exit Function
    If int门诊标志 = 2 Then
        If Val(int主页ID) = 0 Then
            Is门诊费用 = True: Exit Function
        End If
    End If
End Function
Private Sub mnuEditCancel_Click()
    Dim strNO As String, int性质 As Integer, int序号 As Integer
    Dim blnDo As Boolean, i As Long, blnTrans As Boolean
    Dim cllData As Collection, j As Long, blnFind As Boolean
    Dim varTemp As Variant, cllPro As Collection
    Dim strSQL As String
    If cboUnit.ListIndex = -1 Then Exit Sub
    'arrSQL = Array()
    Set cllData = New Collection
    For i = 1 To mshList.Rows - 1
        If Val(mshList.TextMatrix(i, 0)) <> -1 Then
            strNO = GetColValue(i, "单据号")
            If strNO <> "" Then
                If mshList.TextMatrix(i, 1) <> "" And mshList.TextMatrix(i, GetColNum("执行人")) <> "" Then
                    '进入此条件表示存在打勾的行
                    int性质 = GetColValue(i, "记录性质")
                    blnDo = True
                     '当前选择的单据列表可能不止一个,所以不能取之前确定的是否在后备表的标记,需要现判断
                    '是否已转入后备数据表中
                    If frmExecuteFilter.mblnDateMoved Then
                        If zlDatabase.NOMoved(IIf(Is门诊费用(i), "门诊费用记录", "住院费用记录"), strNO, , int性质, Me.Caption) Then
                            If Not ReturnMovedExes(strNO, int性质, Me.Caption) Then blnDo = False
                            'mblnNOMoved = False  '此句不能要,否则影响不选的情况
                        End If
                    End If
                
                    If blnDo Then
                        If InStr(mstrPrivs, ";取消他人登记;") = 0 And mshList.TextMatrix(i, GetColNum("执行人")) <> UserInfo.姓名 Then
                            mshList.TopRow = i
                            mshList_LeaveCell
                            mshList.Row = i
                            mshList_EnterCell
                            MsgBox strNO & " 中项目 """ & mshList.TextMatrix(i, GetColNum("项目")) & """ 的执行人为其他人，你没有权限取消登记！", vbInformation, gstrSysName
                            blnDo = False
                        End If
                    
                        '费用SQL
                        int序号 = Val(mshList.TextMatrix(i, 0))
                        
                        blnFind = False
                        For j = 1 To cllData.Count
                            varTemp = cllData(j)
                            If varTemp(0) = strNO And Val(varTemp(1)) = int性质 Then
                                cllData.Remove j
                                cllData.Add Array(strNO, int性质, varTemp(2) & "," & int序号, IIf(Is门诊费用(i), 1, 2))
                                blnFind = True: Exit For
                            End If
                        Next
                        If blnFind = False Then
                             cllData.Add Array(strNO, int性质, "" & int序号, IIf(Is门诊费用(i), 1, 2))
                        End If
                        
'                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
'                        arrSQL(UBound(arrSQL)) = "zl_病人费用记录_UNExecute('" & strNO & "'," & int性质 & "," & int序号 & "," & IIf(Is门诊费用(i), 1, 2) & ")"
                    End If
                    
                    blnDo = True  '表示当前列表中存在打勾的项
                End If
            End If
        End If
    Next
    Set cllPro = New Collection
    For j = 1 To cllData.Count
        'NO,性质,序号(多个),门诊标志
        varTemp = cllData(j)
        strSQL = "zl_病人费用记录_UNExecute('" & varTemp(0) & "'," & Val(varTemp(1)) & ",'" & varTemp(2) & "'," & Val(varTemp(3)) & ")"
        Call zlAddArray(cllPro, strSQL)
    Next
    
    If blnDo Then  '表示存在打勾的行
        If cllPro.Count = 0 Then Exit Sub '如果存在打勾的,但又全部在后备表中则退出
        If MsgBox("确实要将选择的记录全部取消登记吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        '没有选择,则只处理当前行
        If MsgBox("确实要将当前记录取消登记吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        strNO = GetColValue(mshList.Row, "单据号")
        If mshList.Row = 0 Or strNO = "" Then Exit Sub
        
        int序号 = Val(mshList.TextMatrix(mshList.Row, 0))
        If int序号 = -1 Then Exit Sub
        int性质 = GetColValue(mshList.Row, "记录性质")
        
        '如果没有打勾,只有当前行的情况,是否已转入后备数据表中
        If mblnNOMoved Then
            If Not ReturnMovedExes(strNO, int性质, Me.Caption) Then Exit Sub
            mblnNOMoved = False  '此时已转入在线数据表
        End If
        
        If InStr(mstrPrivs, ";取消他人登记;") = 0 And mshList.TextMatrix(mshList.Row, GetColNum("执行人")) <> UserInfo.姓名 Then
            MsgBox "当前项目 """ & mshList.TextMatrix(mshList.Row, GetColNum("项目")) & """ 的执行人为其他人，你没有权限取消登记！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strSQL = "zl_病人费用记录_UNExecute('" & strNO & "'," & int性质 & "," & int序号 & "," & IIf(Is门诊费用(mshList.Row), 1, 2) & ")"
        Call zlAddArray(cllPro, strSQL)
    End If
            
    Screen.MousePointer = 11
    On Error GoTo errH
    zlExecuteProcedureArrAy cllPro, Me.Caption
    On Error GoTo 0
    Screen.MousePointer = 0
    mnuViewReFlash_Click
    Exit Sub
errH:
    Screen.MousePointer = 0
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditClear_Click()
    Dim i As Long, j As Long
    j = GetColNum("单据号")
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, j) <> "" And Val(mshList.TextMatrix(i, 0)) <> -1 Then
            mshList.TextMatrix(i, 1) = ""
        End If
    Next
End Sub

Private Sub mnuEditLog_Click()
    Dim strNO As String, bytFlag As Byte, int序号 As Integer
    Dim strOper As String, strLog As String, lng病人ID As Long, lng主页ID As Long
    Dim arrSQL() As Variant, i As Long
    Dim arrPar() As Variant, blnPrint As Boolean
    Dim strInfo As String, blnDo As Boolean, blnTrans As Boolean, blnCheck As Boolean
    
    If cboUnit.ListIndex = -1 Then Exit Sub
    
    arrSQL() = Array()
    arrPar() = Array()
    
    'a.批量执行处理
    For i = 1 To mshList.Rows - 1
        If Val(mshList.TextMatrix(i, 0)) <> -1 Then
            strNO = GetColValue(i, "单据号")
            
            If strNO <> "" And mshList.TextMatrix(i, 1) <> "" Then
                '进入此条件表示存在打勾的行
                
                bytFlag = Val(GetColValue(i, "记录性质"))
            
                blnDo = True
                '当前选择的单据列表可能不止一个,所以不能取之前确定的是否在后备表的标记,需要现判断
                '是否已转入后备数据表中
                '筛选时的时间在最后一次转出之前
                If frmExecuteFilter.mblnDateMoved Then
                    If zlDatabase.NOMoved(IIf(Is门诊费用(i), "门诊费用记录", "住院费用记录"), strNO, , bytFlag, Me.Caption) Then
                        If Not ReturnMovedExes(strNO, bytFlag, Me.Caption) Then blnDo = False
                        'mblnNOMoved = False '此句不能要,否则影响不选的情况
                    End If
                End If
                
                If blnDo Then
                    int序号 = Val(mshList.TextMatrix(i, 0))
                    '对于已执行的，获取第一个执行人和登记内容作为批量执行参考值
                    If strOper = "" Then strOper = mshList.TextMatrix(mshList.Row, GetColNum("执行人"))
                    
                    If strLog = "" Then strLog = GetItemLog(IIf(Is门诊费用(i), 1, 2), strNO, bytFlag, int序号) '前面已处理为在线表,此处不必传mblnNOMoved
                    
                    If gbln执行后审核 And GetColValue(i, "记录性质") = 2 And GetColValue(i, "记录状态") = 0 Then
                        If AuditingWarn(mstrPrivs, mrsWarn, strNO, int序号) Then
                            If lng病人ID <> Val(GetColValue(i, "病人ID")) Then
                                lng病人ID = Val(GetColValue(i, "病人ID"))
                                lng主页ID = Val(GetColValue(i, "主页ID"))
                                blnCheck = PatiCanBilling(lng病人ID, lng主页ID, mstrPrivs)
                            End If
                        Else
                            blnCheck = False
                        End If
                    Else
                        blnCheck = True
                    End If
                            
                    If blnCheck Then
                        '1.票据打印参数
                        ReDim Preserve arrPar(UBound(arrPar) + 1)
                        arrPar(UBound(arrPar)) = strNO & "," & bytFlag & "," & int序号
                                        
                        '2.费用SQL
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "zl_病人费用记录_Execute('" & strNO & "'," & bytFlag & "," & int序号 & "," & IIf(Is门诊费用(i), 1, 2) & ","
                    End If
                End If
                
                blnDo = True  '表示当前列表中存在打勾的行
            End If
        End If
    Next
    
    If blnDo Then  '表示存在打勾的行
        If UBound(arrSQL) < 0 And blnDo Then Exit Sub   '如果存在打勾的,但又全部在后备表中则退出
    
        If MsgBox("确实要对选择的记录全部进行登记吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '不存在打勾的行，则只处理当前行
    Else
        strNO = GetColValue(mshList.Row, "单据号")
        If mshList.Row = 0 Or strNO = "" Then Exit Sub
        
        int序号 = Val(mshList.TextMatrix(mshList.Row, 0))
        If int序号 = -1 Then Exit Sub
        bytFlag = GetColValue(mshList.Row, "记录性质")
        strOper = mshList.TextMatrix(mshList.Row, GetColNum("执行人"))
        strLog = GetItemLog(IIf(Is门诊费用(mshList.Row), 1, 2), strNO, bytFlag, int序号)
        
        '如果没有打勾,只有当前行的情况,是否已转入后备数据表中
        If mblnNOMoved Then
            If Not ReturnMovedExes(strNO, bytFlag, Me.Caption) Then Exit Sub
            mblnNOMoved = False  '此时已转入在线数据表
        End If
        
        blnCheck = True
        If gbln执行后审核 And GetColValue(mshList.Row, "记录性质") = 2 And GetColValue(mshList.Row, "记录状态") = 0 Then
            If AuditingWarn(mstrPrivs, mrsWarn, strNO, int序号) Then
                lng病人ID = Val(GetColValue(mshList.Row, "病人ID"))
                lng主页ID = Val(GetColValue(mshList.Row, "主页ID"))
                blnCheck = PatiCanBilling(lng病人ID, lng主页ID, mstrPrivs)
            Else
                blnCheck = False
            End If
        End If
        
        If blnCheck Then
            '票据打印参数
            ReDim arrPar(0)
            arrPar(0) = strNO & "," & bytFlag & "," & int序号
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_病人费用记录_Execute('" & strNO & "'," & bytFlag & "," & int序号 & "," & IIf(Is门诊费用(mshList.Row), 1, 2) & ","
        Else
            Exit Sub
        End If
    End If
    
    
    On Error Resume Next
    frmExeEdit.mlngDeptID = cboUnit.ItemData(cboUnit.ListIndex)
    frmExeEdit.mstrOper = strOper
    frmExeEdit.mstrLog = strLog
    frmExeEdit.Show 1, Me
    If gblnOK Then
        For i = 0 To UBound(arrSQL)
            With frmExeEdit
                arrSQL(i) = arrSQL(i) & "'" & .mstrLog & "','" & .mstrOper & "',To_Date('" & Format(.mvDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            End With
        Next
                
        Screen.MousePointer = 11
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
            Next
        gcnOracle.CommitTrans: blnTrans = False
        On Error GoTo 0
        Call mshList_EnterCell
        Screen.MousePointer = 0
        
        '打印票据
        blnPrint = False
        If gbytExe打印方式 = 1 Then
            blnPrint = True
        ElseIf gbytExe打印方式 = 2 Then
            If UBound(arrPar) > 0 Then
                strInfo = "执行登记完毕,要打印刚才选择的所有执行登记单吗？"
            Else
                strInfo = "执行登记完毕,要打印执行登记单吗？"
            End If
            blnPrint = MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
        End If
        
        If blnPrint Then
            Screen.MousePointer = 11
            For i = 0 To UBound(arrPar)
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1142", Me, _
                    "NO=" & Split(arrPar(i), ",")(0), _
                    "记录性质=" & Split(arrPar(i), ",")(1), _
                    "序号=" & Split(arrPar(i), ",")(2), 2)
            Next
            Screen.MousePointer = 0
        End If
        
        '刷新
        mnuViewReFlash_Click
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditPrint_Click()
    Dim strNO As String, int性质 As Integer, int序号 As Integer
    Dim arrPar() As Variant, i As Long
    Dim blnDo As Boolean
    
    If cboUnit.ListIndex = -1 Then Exit Sub
    
    arrPar() = Array()
    For i = 1 To mshList.Rows - 1
        If Val(mshList.TextMatrix(i, 0)) <> -1 Then
            strNO = GetColValue(i, "单据号")
            If strNO <> "" And mshList.TextMatrix(i, 1) <> "" Then
                int序号 = Val(mshList.TextMatrix(i, 0))
                int性质 = GetColValue(i, "记录性质")
                
                blnDo = True
            
                '当前选择的单据列表可能不止一个,所以不能取之前确定的是否在后备表的标记,需要现判断
                '是否已转入后备数据表中
                If frmExecuteFilter.mblnDateMoved Then
                    If zlDatabase.NOMoved(IIf(Is门诊费用(i), "门诊费用记录", "住院费用记录"), strNO, , int性质, Me.Caption) Then
                        If Not ReturnMovedExes(strNO, int性质, Me.Caption) Then blnDo = False
                        mblnNOMoved = False
                    End If
                End If
                
                If blnDo Then
                    '票据打印参数
                    ReDim Preserve arrPar(UBound(arrPar) + 1)
                    arrPar(UBound(arrPar)) = strNO & "," & int性质 & "," & int序号
                End If
            End If
        End If
    Next
    
    If UBound(arrPar) >= 0 Then
        If MsgBox("确实要对选择的记录全部进行打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        strNO = GetColValue(mshList.Row, "单据号")
        If mshList.Row = 0 Or strNO = "" Then Exit Sub
        int序号 = Val(mshList.TextMatrix(mshList.Row, 0))
        If int序号 = -1 Then Exit Sub
        int性质 = GetColValue(mshList.Row, "记录性质")
        
        '票据打印参数
        ReDim arrPar(0)
        arrPar(0) = strNO & "," & int性质 & "," & int序号
    End If
    
    If Not ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1142", Me) Then Exit Sub
    
    Screen.MousePointer = 11
    For i = 0 To UBound(arrPar)
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1142", Me, _
            "NO=" & Split(arrPar(i), ",")(0), _
            "记录性质=" & Split(arrPar(i), ",")(1), _
            "序号=" & Split(arrPar(i), ",")(2), 2)
    Next
    Screen.MousePointer = 0
    
    mnuViewReFlash_Click
    'Call mshList_EnterCell '刷新中已执行
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditSelAll_Click()
    Dim i As Long, j As Long
    j = GetColNum("单据号")
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, j) <> "" And Val(mshList.TextMatrix(i, 0)) <> -1 Then
            mshList.TextMatrix(i, 1) = "√"
        End If
    Next
End Sub

Private Sub mnuEditView_Click()
    Dim str单据号 As String, bytFlag As Byte, int序号 As Integer
    Dim strOper As String, strLog As String, strDate As String
    
    If cboUnit.ListIndex = -1 Then Exit Sub
    
    str单据号 = GetColValue(mshList.Row, "单据号")
    If mshList.Row = 0 Or str单据号 = "" Then Exit Sub
    
    int序号 = Val(mshList.TextMatrix(mshList.Row, 0))
    If int序号 = -1 Then Exit Sub
    
    bytFlag = GetColValue(mshList.Row, "记录性质")
    strDate = GetColValue(mshList.Row, "执行时间")
    
    strOper = mshList.TextMatrix(mshList.Row, GetColNum("执行人"))
    strLog = GetItemLog(IIf(Is门诊费用(mshList.Row), 1, 2), str单据号, bytFlag, int序号, mblnNOMoved)
    
    frmExeEdit.mblnView = True
    frmExeEdit.mlngDeptID = cboUnit.ItemData(cboUnit.ListIndex)
    frmExeEdit.mstrOper = strOper
    frmExeEdit.mstrLog = strLog
    frmExeEdit.mstrDate = strDate
    
    frmExeEdit.Show 1, Me
End Sub

Private Sub mnuFileSetup_Click()
    frmExecuteSet.mlngModul = mlngModul
    frmExecuteSet.mstrPrivs = mstrPrivs
    frmExecuteSet.Show 1, Me
    If frmExecuteSet.mblnOK Then
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNO As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNO = "" Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "病人科室=" & mlngDeptID)
    Else
        With mshList
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "病人科室=" & mlngDeptID, "NO=" & strNO, _
                "病人ID=" & .TextMatrix(.Row, GetColNum("病人ID")), _
                "主页ID=" & .TextMatrix(.Row, GetColNum("主页ID")), _
                "住院号=" & .TextMatrix(.Row, GetColNum("住院号")), _
                "开单人=" & .TextMatrix(.Row, GetColNum("开单人")))
        End With
    End If
End Sub

Private Sub mnuViewFilter_Click()
    
    If frmExecuteFilter.mlngDept <> mlngDeptID Then
        frmExecuteFilter.mlngDept = mlngDeptID
        frmExecuteFilter.LoadOper
    End If
    
    frmExecuteFilter.Show 1, Me
    If gblnOK Then Call FindBills
End Sub

Private Sub FindBills()
    With frmExecuteFilter
        mstrFilter = .mstrFilter
        
        SQLCondition.Default = False
        SQLCondition.DateB = .dtpBegin.Value
        SQLCondition.DateE = .dtpEnd.Value
        SQLCondition.NOB = .txtNOBegin.Text
        SQLCondition.NOE = .txtNoEnd.Text
        If .cbo状态.Text <> "所有状态" Then SQLCondition.State = Val(.cbo状态.Text) - 1
        SQLCondition.Operator = zlStr.NeedName(.cbo执行人.Text)
        SQLCondition.ID = Val(.txt标识号.Text)
        SQLCondition.Patient = gstrLike & UCase(.txt姓名.Text) & "%"
    End With
    
    mnuViewReFlash_Click
End Sub

Private Sub mnuViewShowHead_Click()
    mnuViewShowHead.Checked = Not mnuViewShowHead.Checked
    Call SetBillHead(False)
    Call SetHeader
End Sub

Private Sub mshList_DblClick()
Dim i As Integer
Dim bln正序 As Boolean
Dim strNO As String

'算法:向后再向前快速选择或反选同一单据的明细行,同一单据相同主项的明细行

    If mshList.MouseRow = 0 Then Exit Sub
    If mshList.MouseCol = GetColNum("选择") Then
        If Val(mshList.TextMatrix(mshList.Row, 0)) <> -1 Then                  '单据行的值为-1
            '1.如果在明细行双击
            strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))  '以该单据第一行明细的单据号为准
            If strNO <> "" Then  '无数据时才会空
                '如果是主项,则选择或反选它的所有儿子,如果是从项,则选择或反选它的父亲及所有兄弟
                '先正序
                If chkAuto.Value = 1 Then
                    If mshList.Row <> mshList.Rows - 1 Then
                        For i = mshList.Row + 1 To mshList.Rows - 1       '当前行在倒序时再选择
                            If mshList.TextMatrix(i, GetColNum("单据号")) <> strNO Then Exit For
                            If mshList.TextMatrix(i, GetColNum("父号")) = mshList.TextMatrix(mshList.Row, GetColNum("父号")) Then
                                If mshList.TextMatrix(i, 1) = "" Then
                                    mshList.TextMatrix(i, 1) = "√"
                                Else
                                    mshList.TextMatrix(i, 1) = ""
                                End If
                            End If
                        Next
                    End If
                    '再倒序
                    For i = mshList.Row To 0 Step -1
                        If mshList.TextMatrix(i, GetColNum("单据号")) <> strNO Then Exit For
                        If mshList.TextMatrix(i, GetColNum("父号")) = mshList.TextMatrix(mshList.Row, GetColNum("父号")) Then
                            If mshList.TextMatrix(i, 1) = "" Then
                                mshList.TextMatrix(i, 1) = "√"
                            Else
                                mshList.TextMatrix(i, 1) = ""
                            End If
                        End If
                    Next
                Else
                    If mshList.TextMatrix(mshList.Row, 1) = "" Then
                        mshList.TextMatrix(mshList.Row, 1) = "√"
                    Else
                        mshList.TextMatrix(mshList.Row, 1) = ""
                    End If
                End If
            End If
        Else
            '2.如果是在单据行双击,则选择或反选该单据的所有明细行
            '先假设正序,向后搜索,单据行在最后一行,一定是倒序
            strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
            strNO = Mid(strNO, InStr(1, strNO, ")") + 1, InStr(1, strNO, " ") - InStr(1, strNO, ")") - 1)
            
            If mshList.Row <> mshList.Rows - 1 Then
                For i = mshList.Row + 1 To mshList.Rows - 1
                    If mshList.TextMatrix(i, GetColNum("单据号")) <> strNO Then Exit For
                    If mshList.TextMatrix(i, 1) = "" Then
                        mshList.TextMatrix(i, 1) = "√"
                    Else
                        mshList.TextMatrix(i, 1) = ""
                    End If
                    bln正序 = True
                Next
            End If
            If Not bln正序 Then
                For i = mshList.Row - 1 To 0 Step -1 '如果是第一行,则前面肯定会执行
                    If mshList.TextMatrix(i, GetColNum("单据号")) <> strNO Then Exit For
                    If mshList.TextMatrix(i, 1) = "" Then
                        mshList.TextMatrix(i, 1) = "√"
                    Else
                        mshList.TextMatrix(i, 1) = ""
                    End If
                Next
            End If
        End If
    ElseIf mnuEditView.Enabled Then
        Call mnuEditView_Click
    ElseIf mnuEditLog.Visible And mnuEditLog.Enabled Then
        Call mnuEditLog_Click
    End If
End Sub

Private Sub mshList_EnterCell()
    Dim strNO As String, int序号 As Integer, i As Long
    Dim intRows As Integer, bln As Boolean
    Dim lng开单ID As Long, lng执行ID As Long
    Dim bln执行 As Boolean, lng科室ID As Long
    Dim bytFlag As Byte
    
    strNO = GetColValue(mshList.Row, "单据号")
    If mshList.Row = 0 Or strNO = "" Then Exit Sub
    
    mlngGo = mshList.Row
    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
    
    
    bln = mshList.Redraw
    mshList.Redraw = False
    '设置背景色
    For i = 0 To mshList.Cols - 1
        mshList.Col = i
        mshList.CellBackColor = mshList.BackColorSel
        mshList.CellForeColor = mshList.ForeColorSel
    Next
    mshList.Col = 0
    
    '设置顶行
    intRows = (mshList.Height - mshList.RowHeight(0) - 60) \ 250
    If mshList.TopRow > mshList.Row Then
        mshList.TopRow = mshList.Row
    ElseIf mshList.Row - mshList.TopRow >= intRows Then
        mshList.TopRow = mshList.Row - intRows + 1
    End If
    
    mshList.Redraw = bln
    
    int序号 = Val(mshList.TextMatrix(mshList.Row, 0))
    bln执行 = (GetColValue(mshList.Row, "状态") = "已执行")
    
    
    mnuEditLog.Enabled = int序号 <> -1 And Not bln执行
    tbr.Buttons("Log").Enabled = mnuEditLog.Enabled
    
    mnuEditCancel.Enabled = int序号 <> -1 And bln执行
    tbr.Buttons("Cancel").Enabled = mnuEditCancel.Enabled
    
    mnuEditView.Enabled = mnuEditCancel.Enabled
    tbr.Buttons("View").Enabled = mnuEditCancel.Enabled
    
    mnuEditPrint.Enabled = mnuEditLog.Enabled
    
    If int序号 = -1 Then
        stbThis.Panels(2) = stbThis.Tag
        mblnNOMoved = False
    Else
        bytFlag = GetColValue(mshList.Row, "记录性质")
        If frmExecuteFilter.mblnDateMoved Then
            mblnNOMoved = zlDatabase.NOMoved(IIf(Is门诊费用(mshList.Row), "门诊费用记录", "住院费用记录"), strNO, , bytFlag, Me.Caption)
        Else
            mblnNOMoved = False
        End If
        
        stbThis.Panels(2) = "执行情况:" & GetItemLog(IIf(Is门诊费用(mshList.Row), 1, 2), strNO, bytFlag, int序号, mblnNOMoved)
    End If
End Sub

Private Sub mshList_GotFocus()
    mshList.BackColorSel = &H8000000D
    Call mshList_EnterCell
End Sub

Private Sub mshList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mnuEditView.Enabled Then
            mnuEditView_Click
        ElseIf mnuEditLog.Enabled And mnuEditLog.Visible Then
            mnuEditLog_Click
        End If
    ElseIf KeyAscii = Asc(" ") Then
        If Val(mshList.TextMatrix(mshList.Row, 0)) <> -1 Then
            If mshList.TextMatrix(mshList.Row, GetColNum("单据号")) <> "" Then
                If mshList.TextMatrix(mshList.Row, 1) = "" Then
                    mshList.TextMatrix(mshList.Row, 1) = "√"
                Else
                    mshList.TextMatrix(mshList.Row, 1) = ""
                End If
            End If
        End If
    End If
End Sub

Private Sub mshList_LeaveCell()
    Dim i As Long
    Dim bln As Boolean
    
    '设置背景色
    bln = mshList.Redraw
    mshList.Redraw = False
    For i = 0 To mshList.Cols - 1
        mshList.Col = i
        If Val(mshList.TextMatrix(mshList.Row, 0)) = -1 Then
            mshList.CellBackColor = &HEBFFFF '&HE6FFFF '&HE0E0E0
        Else
            mshList.CellBackColor = mshList.BackColor
        End If
        mshList.CellForeColor = mshList.ForeColor
    Next
    mshList.Redraw = bln
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Button = 2 Then PopupMenu mnuEdit, 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF4
            If Not ActiveControl Is txtValue Then Call txtValue.SetFocus
        Case vbKeyF3
            '始终从当前行开始
            If mnuViewGo.Enabled Then Call SeekBill(False)
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewReFlash_Click()
    Call ShowBills(mstrFilter)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Long
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).minHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub mnuViewToolUnit_Click()
    mnuViewToolUnit.Checked = Not mnuViewToolUnit.Checked
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = False
    cbr.Bands(2).Visible = Not cbr.Bands(2).Visible
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = True
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Bands(1).Visible = Not cbr.Bands(1).Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "Go" '定位
            mnuViewGo_Click
        Case "Filter" '过滤
            mnuViewFilter_Click
        Case "Log"
            mnuEditLog_Click
        Case "Cancel"
            mnuEditCancel_Click
        Case "View"
            mnuEditView_Click
        Case "SelAll"
            mnuEditSelAll_Click
        Case "Clear"
            mnuEditClear_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub mnuFile_Excel_Click()
    Call OutputList(3)
End Sub

Private Sub mnuFile_PreView_Click()
    Call OutputList(2)
End Sub

Private Sub mnuFile_Print_Click()
    Call OutputList(1)
End Sub

Private Sub mnuFile_PrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub OutputList(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    intRow = mshList.Row
    
    '表头
    objOut.Title.Text = "医技单据清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    With frmExecuteFilter
        objRow.Add "时间：" & Format(.dtpBegin.Value, .dtpBegin.CustomFormat) & " 至 " & Format(.dtpEnd.Value, .dtpEnd.CustomFormat)
        objOut.UnderAppRows.Add objRow
    End With
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    mshList.Redraw = False
    mshList_LeaveCell
    Set objOut.Body = mshList
    
    '输出
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mshList.Row = intRow
    mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
    mshList_EnterCell
    mshList.Redraw = True
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub SetMenu(blnUsed As Boolean)
'功能：根据有无记录设置菜单可用状态
    mnuFile_Print.Enabled = blnUsed
    mnuFile_PreView.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    
    mnuEditLog.Enabled = blnUsed
    tbr.Buttons("Log").Enabled = blnUsed
    
    mnuEditCancel.Enabled = blnUsed
    tbr.Buttons("Cancel").Enabled = blnUsed
        
    mnuEditView.Enabled = blnUsed
    tbr.Buttons("View").Enabled = blnUsed
        
    mnuEditPrint.Enabled = blnUsed
    
    mnuViewGo.Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    Call RestoreWinState(Me, App.ProductName)
    mnuViewShowHead.Checked = zlDatabase.GetPara("显示单据头", glngSys, mlngModul, "1") = "1"
    
    mlngCurRow = 1: mlngTopRow = 1
    
    '权限设置
    If InStr(mstrPrivs, ";执行登记;") = 0 And InStr(mstrPrivs, ";取消登记;") = 0 Then
        mnuEditLog.Visible = False
        mnuEditCancel.Visible = False
        
        tbr.Buttons("Log").Visible = False
        tbr.Buttons("Cancel").Visible = False
        tbr.Buttons("Edit_").Visible = False
    ElseIf InStr(mstrPrivs, ";执行登记;") = 0 Then
        mnuEditLog.Visible = False
        tbr.Buttons("Log").Visible = False
    ElseIf InStr(mstrPrivs, ";取消登记;") = 0 Then
        mnuEditCancel.Visible = False
        tbr.Buttons("Cancel").Visible = False
    End If
    
    '主从项目同时选择
    chkAuto.Value = IIf(zlDatabase.GetPara("主从项目同时选择", glngSys, mlngModul, "0") = "1", 1, 0)
    mnuIDKinds(0).Checked = True '单据号
    
    If gbln执行后审核 Then Set mrsWarn = GetUnitWarn
    
    '科室
    If Not InitUnits Then Unload Me: Exit Sub
    If cboUnit.ListIndex = -1 Then
        MsgBox "没有发现你所属科室,且你不具有所有科室权限,不能使用医技科室记帐！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    Call SetHeader
    Call SetMenu(False)
    stbThis.Panels(2).Text = "请刷新清单或重新设置过滤条件"
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    mshList.MousePointer = 0
    
    '靠齐控件宽度和高度
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    mshList.Left = 0
    mshList.Top = cbrH
    mshList.Width = Me.ScaleWidth
    mshList.Height = Me.ScaleHeight - cbrH - staH
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrFilter = ""
    mlngDeptID = 0
    mstrPreNO = ""
    
    Unload frmExecuteFilter
    Unload frmExecuteGo
    Call SaveWinState(Me, App.ProductName)
    zlDatabase.SetPara "显示单据头", IIf(mnuViewShowHead.Checked, 1, 0), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub mnuViewGo_Click()
    frmExecuteGo.Show 1, Me
    mstrPreNO = ""
    If gblnOK Then Call SeekBill(frmExecuteGo.optHead)
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long, bln As Boolean, intRows As Integer
    Dim blnFill As Boolean, j As Long
    Dim strCurNO As String
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "正在定位满足条件的单据,按ESC终止 ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
        DoEvents

        strCurNO = GetColValue(i, "单据号")
        
        If Val(mshList.TextMatrix(i, 0)) <> -1 Then
            '比较条件
            blnFill = True
            With frmExecuteGo
                If .txtNO.Text <> "" Then
                    blnFill = blnFill And strCurNO = .txtNO.Text
                End If
                If .txt标识号.Text <> "" Then
                    blnFill = blnFill And GetColValue(i, frmExecuteGo.lbl标识号.Caption) = .txt标识号.Text
                End If
                If .txt病人ID.Text <> "" Then
                    blnFill = blnFill And GetColValue(i, "病人ID") = .txt病人ID.Text
                End If
                If .txt姓名.Text <> "" Then
                    blnFill = blnFill And UCase(GetColValue(i, "姓名")) Like "*" & UCase(.txt姓名.Text) & "*"
                End If
            End With
            blnFill = blnFill And (strCurNO <> mstrPreNO)
            
            '满足则退出
            If blnFill Then
                mstrPreNO = strCurNO
                
                mlngGo = i + 1
    
                'LeaveCell设置
                bln = mshList.Redraw
                mshList.Redraw = False
                For j = 0 To mshList.Cols - 1
                    mshList.Col = j
                    If Val(mshList.TextMatrix(mshList.Row, 0)) = -1 Then
                        mshList.CellBackColor = &HEBFFFF '&HE6FFFF '&HE0E0E0
                    Else
                        mshList.CellBackColor = mshList.BackColor
                    End If
                    mshList.CellForeColor = mshList.ForeColor
                Next
                '''''''''''''''''''''
    
                mshList.Row = i
                mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
    
                'EnterCell设置
                For j = 0 To mshList.Cols - 1
                    mshList.Col = j
                    mshList.CellBackColor = mshList.BackColorSel
                    mshList.CellForeColor = mshList.ForeColorSel
                Next
                intRows = (mshList.Height - mshList.RowHeight(0) - 60) \ 250
                If mshList.TopRow > mshList.Row Then
                    mshList.TopRow = mshList.Row
                ElseIf mshList.Row - mshList.TopRow >= intRows Then
                    mshList.TopRow = mshList.Row - intRows + 2
                End If
                mshList.Redraw = bln
                ''''''''''''''''''''
                
                stbThis.Panels(2).Text = "找到一张单据"
                Screen.MousePointer = 0: Exit Sub
            End If
        End If
        
        '按ESC取消
        If mblnGo = False Then
            stbThis.Panels(2).Text = "用户取消定位操作"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1: mstrPreNO = ""
    stbThis.Panels(2).Text = "已定位到清单尾部"
    Screen.MousePointer = 0
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Long
    For i = 0 To mshList.Cols - 1
        If mshList.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshList.MouseRow = 0 Then
        mshList.MousePointer = 99
    Else
        mshList.MousePointer = 0
    End If
End Sub

Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshList.MouseCol
    If lngCol = GetColNum("选择") Then Exit Sub
    
    If Button = 1 And mshList.MousePointer = 99 Then
                 
        If mshList.TextMatrix(0, lngCol) = "" Then Exit Sub
        If GetColValue(mshList.Row, "单据号") = "" Then Exit Sub
        If mrsList Is Nothing Then Exit Sub
        
        Set mshList.DataSource = Nothing
        
        If mshList.TextMatrix(0, lngCol) = "住院号" Or mshList.TextMatrix(0, lngCol) = "门诊号" Then
            mrsList.Sort = "单据号 Desc,标识号" & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        Else
            mrsList.Sort = "单据号 Desc," & mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        End If
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        Call ShowBills(, True)
    End If
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    strHead = "序号,7,0|选择,4,450|单据号,1,0|姓名,1,0|病人ID,1,0|标识号,1,0|床号,1,0|" & _
        "科室,1,1000|开单人,1,0|类别,4,0|项目,1,3000" & IIf(gTy_System_Para.byt药品名称显示 = 2, "|商品名,1,1600", "") & "|规格,1,1000|数量,1,850|单价,7,850|应收金额,7,850|实收金额,7,850|" & _
        "状态,1,650|执行人,1,700|执行时间,4,2000|操作员,1,700|登记时间,4,2000|记录性质,1,0|门诊标志,1,0|父号,1,0|住院号,1,0|主页ID,1,0|记录状态,1,0"
    With mshList
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        '刘兴洪:27990 2010-02-22 17:34:32
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "商品名" Then
                If gTy_System_Para.byt药品名称显示 = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 1600
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        
        mshList.ColWidth(0) = 0 '序号
        mshList.ColWidth(2) = 0 'NO
        mshList.ColWidth(3) = 0 '姓名
        mshList.ColWidth(4) = 0 '病人ID
        mshList.ColWidth(5) = 0 '标识号
        mshList.ColWidth(6) = 0 '床号
        mshList.ColWidth(8) = 0 '开单人
        mshList.ColWidth(9) = 0 '类别
        mshList.ColWidth(mshList.Cols - 3) = 0 '记录性质
        mshList.ColWidth(mshList.Cols - 2) = 0 '门诊标志
        mshList.ColWidth(mshList.Cols - 1) = 0 '父号
        .RowHeight(0) = 320
        
        '恢复上次行
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        
        .Col = 0: .ColSel = .Cols - 1
                
        Call mshList_EnterCell
    End With
End Sub

Private Sub ShowBills(Optional ByVal strIF As String, Optional blnSort As Boolean)
'功能:按条件读取单据列表(过滤功能)
'参数:strIF=以"AND"开始的条件串
'     blnSort=不重新读取数据,仅重新显示已排序的内容
    Dim i As Long, Curdate As Date
    Dim strDept As String, strSQL As String, strType As String
    Dim bytType As Byte '0-门诊,1-住院,2-门诊和住院
    On Error GoTo errH
        
    If Not blnSort Then
        Call zlCommFun.ShowFlash("正在读取单据列表,请稍候 ...", Me)
        DoEvents
        Me.Refresh
        LockWindowUpdate Me.hWnd
        
        '取缺省条件
        If strIF = "" Then
            strIF = " And 登记时间 Between Trunc(sysdate-3) And Trunc(sysdate+1)-1/24/60/60"
            strIF = strIF & " And Nvl(执行状态,0)=0"
        Else
            strIF = strIF & " And Nvl(执行状态,0)<>9"   '问题:44510
        End If
        
        strIF = strIF & " And 执行部门ID+0=[9]"
        If gstrExe类别 = "" Then
            strIF = strIF & " And 收费类别 Not IN('1','5','6','7','J')"
        Else
            strIF = strIF & " And 收费类别 IN(" & gstrExe类别 & ")"
        End If
        
        '000:门诊;住院;体检
        If gstrExe来源 <> "000" Then '将"111"改为000,因为全选时,下面的项目不一定选中 :30493
            If Mid(gstrExe来源, 1, 1) = 1 Then strType = " 门诊标志=1 And 记录状态 in(" & IIf(gbytExe门诊单据类型 = 2, "0,1", gbytExe门诊单据类型) & ")"
            If Mid(gstrExe来源, 2, 1) = 1 Then strType = IIf(strType <> "", strType & " Or", "") & " 门诊标志=2 And 记录状态 in(" & IIf(gbytExe住院单据类型 = 2, "0,1", gbytExe住院单据类型) & ")"
            If Mid(gstrExe来源, 3, 1) = 1 Then strType = IIf(strType <> "", strType & " Or", "") & " 门诊标志=4 And 记录状态 in(" & IIf(gbytExe体检单据类型 = 2, "0,1", gbytExe体检单据类型) & ")"
            If strType <> "" Then strIF = strIF & " And (" & strType & ")"
            If Mid(gstrExe来源, 2, 1) = 0 Then  '无住院,肯定就是门诊的了
                bytType = 0
            ElseIf Mid(gstrExe来源, 1, 1) = 1 Or Mid(gstrExe来源, 3, 1) = 1 Then
                '肯定包含门诊和住院
                bytType = 2
            Else    '只存在住院情况
                bytType = 1
            End If
           bytType = IIf(Mid(gstrExe来源, 2, 1) = 0, 0, 2)
        Else
            strIF = strIF & " And 记录状态 in(0,1)"
            bytType = 2
        End If
        
        
        strIF = strIF & IIf(Not gblnExe医嘱, "  And 医嘱序号 is NULL", "")
        
        If (gstrExe类别 = "" Or InStr(gstrExe类别, "'4'") > 0) And gbln执行后发料 = False Then
            strIF = strIF & " And (收费类别 <> '4' or 收费类别 = '4' And Not Exists(Select 1 From 材料特性 C Where A.收费细目id = C.材料id And C.跟踪在用 = 1))"
        End If
        '77838,冉俊明,2014-9-16,医保病人部分退费后提出不出来剩余未执行部分的记录
        strIF = " Where 记录性质>0 And (记录性质<10 Or 记录性质=11) And 记录性质<>3 " & strIF
        
        Dim strTable As String
 
        strTable = "" & _
        "   Select  A.价格父号, A.序号, A.从属父号, A.NO, A.姓名, A.病人id,A.标识号, A.门诊标志, A.床号, A.开单部门id, A.开单人, A.收费类别, A.收费细目id, " & _
        "           A.数次, A.计算单位, A.标准单价, A.应收金额, A.实收金额, A.执行状态, A.执行人, A.执行时间, A.操作员姓名, A.划价人, A.登记时间, A.记录性质, " & _
        "           A.多病人单, A.主页id, 记录状态 " & _
        "  From 住院费用记录 A " & _
           strIF
        If frmExecuteFilter.mblnDateMoved Then     '筛选时的时间在最后一次转出之前
            strTable = strTable & " Union ALL " & Replace(strTable, "住院费用记录", "H住院费用记录")
        End If
        Select Case bytType
        Case 0  '门诊
            strTable = Replace(Replace(Replace(Replace(strTable, "住院费用记录", "门诊费用记录"), "A.床号", "'' as 床号"), "A.多病人单", " 0 as 多病人单"), "A.主页id", "0 as 主页id")
        Case 1
        Case Else
            strTable = strTable & " Union ALL " & Replace(Replace(Replace(Replace(strTable, "住院费用记录", "门诊费用记录"), "A.床号", "'' as 床号"), "A.多病人单", " 0 as 多病人单"), "A.主页id", "0 as 主页id")
        End Select
          
        strSQL = _
        " Select A.序号,NULL as 选择,A.单据号,A.姓名,A.病人ID,A.标识号,A.床号,D.名称 as 科室,A.开单人," & _
        "       B.名称 as 类别,Nvl(E.名称,C.名称) as 项目," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 as 商品名,", "") & "C.规格,A.数量,A.单价,A.应收金额,A.实收金额,A.状态,A.执行人,A.执行时间," & _
        "       A.操作员,A.登记时间,A.记录性质,A.门诊标志,A.父号,A.住院号,A.主页ID,A.记录状态 " & _
        " From ( Select Nvl(A.价格父号,A.序号) as 序号,Nvl(A.从属父号,A.序号) 父号,A.NO as 单据号,A.姓名,A.病人ID,A.标识号,Decode(A.门诊标志,2,A.床号,NULL) as 床号," & _
        "               A.开单部门ID,A.开单人,A.收费类别,A.收费细目ID,Avg(A.数次)||A.计算单位 as 数量,To_Char(Sum(A.标准单价),'9999990.000') as 单价," & _
        "               To_Char(Sum(A.应收金额),'9999999" & gstrDec & "') as 应收金额,To_Char(Sum(A.实收金额),'9999999" & gstrDec & "') as 实收金额," & _
        "               Decode(A.执行状态,1,'已执行','未执行') as 状态,A.执行人,To_Char(A.执行时间,'YYYY-MM-DD HH24:MI:SS') as 执行时间," & _
        "               Nvl(A.操作员姓名,A.划价人) as 操作员,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间,A.记录性质,A.门诊标志,To_Number(Decode(Nvl(A.多病人单,0),1,NULL,A.标识号)) as 住院号,A.主页ID,A.记录状态" & _
        "        From (" & strTable & ") A " & _
        "        Group by Nvl(A.价格父号,A.序号),A.NO,A.标识号,A.病人ID,A.姓名,Decode(A.门诊标志,2,A.床号,NULL),A.开单部门ID,A.开单人," & _
        "                 A.收费类别,A.收费细目ID,A.计算单位,A.执行状态,A.执行人,A.执行时间,Nvl(A.操作员姓名,A.划价人),A.登记时间,A.记录性质,A.门诊标志,To_Number(Decode(Nvl(A.多病人单,0),1,NULL,A.标识号)),A.主页ID,Nvl(A.从属父号,A.序号),A.记录状态" & _
        "       ) A,收费项目类别 B,收费项目目录 C,部门表 D,收费项目别名 E" & _
                IIf(gTy_System_Para.byt药品名称显示 = 2, ",收费项目别名 E1", "") & _
        " Where A.收费类别 = B.编码 And A.收费细目ID = C.ID And A.开单部门ID=D.ID" & _
        "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                IIf(gTy_System_Para.byt药品名称显示 = 2, "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3", "") & _
        " Union All" & _
        " Select -1 as 序号,NULL as 选择,A.NO as 单据号,Decode(A.门诊标志,1,'(门诊)',4,'(体检)','(住院)')||A.NO||" & _
        " '  姓名：'||A.姓名||'  标识号：'||A.标识号||Decode(A.门诊标志,2,'  床号：'||A.床号,NULL)||'  金额：'||LTrim(To_Char(Sum(A.实收金额),'9999999" & gstrDec & "')) as 姓名," & _
        " -NULL as 病人ID,-NULL as 标识号,NULL as 床号,NULL as 科室,NULL as 开单人,NULL as 类别,NULL as 项目," & IIf(gTy_System_Para.byt药品名称显示 = 2, "NULL as 商品名,", "") & _
        " NULL as 规格,NULL as 数量,NULL as 单价,NULL as 应收金额,NULL as 实收金额,NULL as 状态,NULL,NULL," & _
        " NULL,NULL,A.记录性质,A.门诊标志,-Null as 父号,-Null as 住院号,-Null as 主页ID,-Null as 记录状态 From (" & strTable & ") A" & _
        " Group by A.NO,A.记录性质,A.姓名,A.标识号,A.床号,A.门诊标志" & _
        " Order by 单据号 Desc,门诊标志,记录性质,序号"
        With SQLCondition
            Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .DateB, .DateE, .NOB, .NOE, .State, .Operator, .ID, .Patient, cboUnit.ItemData(cboUnit.ListIndex))
        End With
    End If
    
    mshList.Redraw = False
    mshList.ClearStructure
    mshList.Clear
    mshList.Rows = 2
    
    If mrsList.EOF Then
        stbThis.Panels(2).Text = "当前设置没有过滤出任何单据"
        Call SetMenu(False)
    Else
        Set mshList.DataSource = mrsList
        Call SetMenu(True)
    End If
    Call SetBillHead
    Call SetHeader
    
    'Call mshList_EnterCell   'SetHeader中已执行
    mshList.Redraw = True
    LockWindowUpdate 0
    If Not blnSort Then Call zlCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    LockWindowUpdate 0
End Sub

Private Sub SetBillHead(Optional blnMerge As Boolean = True)
    Dim i As Long, j As Long
    Dim lngRow As Long, lngRows As Long

    lngRow = mshList.Row
    
    Screen.MousePointer = 11
    Me.Refresh
    mshList.Redraw = False
    
    For i = 1 To mshList.Rows - 1
        If Val(mshList.TextMatrix(i, 0)) = -1 Then
            lngRows = lngRows + 1
            
            If blnMerge Then
                mshList.Row = i
                For j = 1 To mshList.Cols - 1
                    mshList.Col = j
                    mshList.CellBackColor = &HEBFFFF
                    mshList.CellAlignment = 1
                    If j <> 3 Then
                        mshList.TextMatrix(i, j) = mshList.TextMatrix(i, 3)
                    End If
                Next
                mshList.MergeRow(i) = True
            End If
            
            mshList.RowHeight(i) = IIf(mnuViewShowHead.Checked, 250, 0)
        ElseIf blnMerge Then
            mshList.MergeRow(i) = False
        End If
    Next
    
    If mshList.RowHeight(lngRow) = 0 Then
        For i = lngRow To mshList.Rows - 1
            If mshList.RowHeight(i) > 0 Then
                Call mshList_LeaveCell
                lngRow = i: Exit For
            End If
        Next
    End If
    mshList.Row = lngRow
    
    'Call mshList_EnterCell   '后面总会有setheader中执行
    
    mshList.Redraw = True
    Screen.MousePointer = 0
    
    stbThis.Panels(2) = "共 " & lngRows & " 张单据"
    stbThis.Tag = stbThis.Panels(2)
End Sub

Private Function GetColValue(ByVal intRow As Integer, strItem As String) As String
'功能：获取指定列的值,因为某些列是合并显示,所以要单独处理
    Dim i As Long, strTmp As String
    If Val(mshList.TextMatrix(intRow, 0)) = -1 Then
        GetColValue = mshList.TextMatrix(IIf(mshList.Row < intRow, intRow + 1, intRow - 1), GetColNum(strItem))
    Else
        GetColValue = mshList.TextMatrix(intRow, GetColNum(strItem))
    End If
End Function

Private Function InitUnits() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    
    '包含门诊/住院医技科室
    If InStr(mstrPrivs, ";所有科室;") > 0 Then
        gstrSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where B.部门ID = A.ID " & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " And B.服务对象 IN(1,2,3) And B.工作性质 IN('检查','检验','手术','治疗','营养')" & _
            " Order by A.编码"
    Else
        gstrSQL = _
            " Select Distinct A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C " & _
            " Where B.部门ID = A.ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " And B.服务对象 IN(1,2,3) And B.工作性质 IN('检查','检验','手术','治疗','营养')" & _
            " Order by A.编码"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.ID)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If cboUnit.ListIndex = -1 Then
                If InStr(mstrPrivs, ";所有科室;") > 0 Then
                    If UserInfo.部门ID = rsTmp!ID Then cboUnit.ListIndex = cboUnit.NewIndex
                Else
                    If rsTmp!缺省 = 1 Then cboUnit.ListIndex = cboUnit.NewIndex
                End If
            End If
            rsTmp.MoveNext
        Next
        If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    ElseIf InStr(mstrPrivs, ";所有科室;") > 0 Then
        MsgBox "没有可用的医技科室,请先到部门管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetFirstRow(ByVal lngRow As Long, lngCol As Long, strValue As String) As Long
'功能：在当前单据中，以当前行为准，获取指定列中值为strValue的行号
    Dim lngRowB As Long, lngRowE As Long
    Dim i As Long
    
    If mshList.TextMatrix(lngRow, lngCol) = strValue Then GetFirstRow = lngRow
    
    If Val(mshList.TextMatrix(lngRow, 0)) = -1 Then
        lngRowB = lngRow + 1
    Else
        lngRowB = 2
        For i = lngRow To 1 Step -1
            If Val(mshList.TextMatrix(i, 0)) = -1 Then
                lngRowB = i + 1: Exit For
            End If
        Next
    End If
    
    lngRowE = mshList.Rows - 1
    For i = IIf(Val(mshList.TextMatrix(lngRow, 0)) = -1, lngRow + 1, lngRow) To mshList.Rows - 1
        If Val(mshList.TextMatrix(i, 0)) = -1 Then
            lngRowE = i - 1: Exit For
        End If
    Next
    
    For i = lngRowB To lngRowE
        If mshList.TextMatrix(i, lngCol) = strValue Then
            GetFirstRow = i: Exit For
        End If
    Next
End Function

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub


