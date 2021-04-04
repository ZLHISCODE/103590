VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmVItemLists 
   BackColor       =   &H8000000C&
   Caption         =   "诊治所见项管理"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9735
   Icon            =   "frmVItemLists.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picHBar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   2790
      MousePointer    =   7  'Size N S
      ScaleHeight     =   30
      ScaleWidth      =   6075
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5310
      Width           =   6075
   End
   Begin VB.PictureBox picVBar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6660
      Left            =   2520
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6660
      ScaleWidth      =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   810
      Width           =   30
   End
   Begin VB.PictureBox picClass 
      Height          =   6270
      Left            =   0
      ScaleHeight     =   6210
      ScaleWidth      =   2340
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   2400
      Begin VB.CommandButton cmdKind 
         Caption         =   "基础项目(&1)"
         Height          =   350
         Index           =   0
         Left            =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   15
         Width           =   2295
      End
      Begin MSComctlLib.TreeView tvwClass 
         Height          =   5580
         Left            =   105
         TabIndex        =   5
         Tag             =   "1000"
         Top             =   405
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   9843
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imgList"
         Appearance      =   0
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1935
      Top             =   6825
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":058A
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":0B24
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":10BE
            Key             =   "item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   3210
      Left            =   2775
      TabIndex        =   1
      Top             =   1260
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   5662
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6900
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmVItemLists.frx":1658
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12091
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
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
   Begin ComCtl3.CoolBar clbThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9735
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbThis"
      MinWidth1       =   24000
      MinHeight1      =   720
      Width1          =   8730
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   24000
         _ExtentX        =   42333
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
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Description     =   "预览"
               Object.ToolTipText     =   "预览当前表"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印当前表"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "分类"
               Key             =   "Class"
               Description     =   "分类"
               Object.ToolTipText     =   "调整药品分类"
               Object.Tag             =   "分类"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "Add"
               Description     =   "增加"
               Object.ToolTipText     =   "增加新的项目"
               Object.Tag             =   "增加"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modify"
               Description     =   "修改"
               Object.ToolTipText     =   "修改当前项目"
               Object.Tag             =   "修改"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Description     =   "删除"
               Object.ToolTipText     =   "删除当前项目"
               Object.Tag             =   "删除"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "Find"
               Description     =   "查找"
               Object.ToolTipText     =   "查找诊断条目"
               Object.Tag             =   "查找"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split4"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   7680
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":1EEA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":2104
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":231E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":2538
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":2752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":296C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":2B86
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":2DA0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":2FC0
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   6915
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":31E0
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":3400
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":3620
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":383A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":3A54
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":3C6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":3E88
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":40A2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":42C2
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   8490
      Top             =   6135
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdInfo 
      Height          =   1695
      Left            =   2835
      TabIndex        =   8
      Top             =   5325
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   7
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483628
      GridColorFixed  =   16777215
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      MergeCells      =   1
      BorderStyle     =   0
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblScale 
      AutoSize        =   -1  'True
      Caption         =   "比例尺寸"
      Height          =   180
      Left            =   8115
      TabIndex        =   9
      Top             =   5775
      Visible         =   0   'False
      Width           =   1185
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
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
      Begin VB.Menu mnuFileSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuClass 
      Caption         =   "分类(&K)"
      Begin VB.Menu mnuClassAdd 
         Caption         =   "新增(&I)"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuClassMod 
         Caption         =   "修改(&U)"
      End
      Begin VB.Menu mnuClassDel 
         Caption         =   "删除(&E)"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "项目(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "新增(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
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
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolbarStand 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolbarText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStates 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web上的中联(&W)"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&E)..."
         End
      End
      Begin VB.Menu mnuHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmVItemLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long
Public mstrPrivs As String       '用户具有本程序的具体权限

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer, intRow As Integer, intCol As Integer
Dim strTemp As String

Private Const conTab执行科室 As Integer = 0
Private Const conTab收费对照 As Integer = 1
Private Const conTab检验指标 As Integer = 2
Private Const conTab检查部位 As Integer = 3
Private Const conTab用法用量 As Integer = 4
Private Const conTab配伍禁忌 As Integer = 5
Private Const conTab配方组成 As Integer = 6
Private Const conTab成套方案 As Integer = 7
Private Const conTab应用参考 As Integer = 8

Private Sub cmdKind_Click(Index As Integer)
    Dim intCount As Integer
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        If intCount <= Index Then
            Me.cmdKind(intCount).Tag = 0
        Else
            Me.cmdKind(intCount).Tag = 1
        End If
    Next
    
    '装数据并调整界面
    If Me.lvwItems.Visible Then
        Call picClass_Resize
        Me.tvwClass.SetFocus
    End If
    If Val(tvwClass.Tag) <> Index Then
        Me.tvwClass.Tag = Index
        Call zlRefClasses
    End If
End Sub

Private Sub clbThis_Resize()
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Call Form_Resize
End Sub

Private Sub Form_Activate()
    Me.lvwItems.Visible = True
End Sub

Private Sub Form_Load()
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    On Error GoTo ErrHand
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "_中文名", "中文名", 1800
        .Add , "_编码", "编码", 1000
        .Add , "_英文名", "英文名", 1500
        .Add , "_类型", "类型", 800
        .Add , "_长度", "长度", 600
        .Add , "_小数", "小数", 600
        .Add , "_单位", "单位", 800
        .Add , "_必填", "必填", 600
    End With
    With Me.lvwItems
        .ColumnHeaders("_编码").Position = 1
        .SortKey = .ColumnHeaders("_编码").Index - 1: .SortOrder = lvwAscending
    End With
    
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    If GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", "1") = "1" Then
        strTemp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\分割", "横向", "0")
        If strTemp <> "0" Then
            Me.picVBar.Left = CLng(strTemp)
        End If
        strTemp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\分割", "纵向", "0")
        If strTemp <> "0" Then
            Me.picHBar.Top = CLng(strTemp)
        End If
    End If
    
    '可直接通过菜单进行的权限控制
    If InStr(1, mstrPrivs, "增删改") = 0 Then
        Me.mnuClass.Visible = False
        Me.mnuEdit.Visible = False
        Me.tlbThis.Buttons("Class").Visible = False
        Me.tlbThis.Buttons("Split2").Visible = False
        Me.tlbThis.Buttons("Add").Visible = False
        Me.tlbThis.Buttons("Modify").Visible = False
        Me.tlbThis.Buttons("Delete").Visible = False
        Me.tlbThis.Buttons("Split3").Visible = False
    End If
    With Me.hgdInfo
        .Rows = 7: .Cols = 2: .FixedRows = 0: .FixedCols = 0
        .ColWidth(0) = 1000: .ColAlignment(0) = 6
        .ColWidth(1) = .Width - .ColWidth(0) - Me.SysInfo.ScrollBarSize: .ColAlignment(1) = 1
        .TextMatrix(0, 0) = "[临床意义]"
        .TextMatrix(1, 0) = "[性别限制]"
        .TextMatrix(2, 0) = "[ 表示法 ]"
        .TextMatrix(3, 0) = "[ 数值域 ]"
'        .TextMatrix(4, 0) = "[ 初始值 ]"
'        .TextMatrix(5, 0) = "[文字表述]"
        .TextMatrix(4, 0) = "[必填项目]"
    End With
    
    '调入已经设置的诊治所见性质
    gstrSql = "select 编码,名称 from 诊治所见性质 Where 编码<>3 order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
    With rsTemp
        Do While Not .EOF
            If .AbsolutePosition > Me.cmdKind.Count Then
                Load Me.cmdKind(.AbsolutePosition - 1)
            End If
            Me.cmdKind(.AbsolutePosition - 1).Caption = !名称 & "(&" & .AbsolutePosition & ")"
            Me.cmdKind(.AbsolutePosition - 1).Left = 0
            Me.cmdKind(.AbsolutePosition - 1).ZOrder 0
            Me.cmdKind(.AbsolutePosition - 1).Visible = True
            .MoveNext
        Loop
    End With
    Call cmdKind_Click(0)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.clbThis.Visible, Me.clbThis.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    Err = 0: On Error Resume Next
    With Me.picVBar
        .Top = lngTools
        .Height = Me.ScaleHeight - picClass.Top - lngStatus
        If .Left < 2000 Then .Left = 2000
        If .Left > Me.ScaleWidth - 4000 Then .Left = Me.ScaleWidth - 4000
    End With
    With Me.picHBar
        .Left = Me.picVBar.Left + Me.picVBar.Width
        .Width = Me.ScaleWidth - .Left
        If .Top < 3000 Then .Top = 3000
        If .Top > Me.ScaleHeight - lngStatus - 1000 Then .Top = Me.ScaleHeight - lngStatus - 1000
    End With
    With Me.picClass
        .Left = Me.ScaleLeft
        .Top = lngTools
        .Height = Me.ScaleHeight - picClass.Top - lngStatus
        .Width = Me.picVBar.Left - Me.picClass.Left
    End With
    
    With Me.lvwItems
        .Left = Me.picVBar.Left + Me.picVBar.Width
        .Top = lngTools
        .Height = Me.picHBar.Top - .Top
        .Width = Me.ScaleWidth - .Left
    End With
    With Me.hgdInfo
        .Left = Me.picVBar.Left + Me.picVBar.Width + 15
        .Top = Me.picHBar.Top + Me.picHBar.Height
        .Height = Me.ScaleHeight - lngStatus - .Top
        .Width = Me.ScaleWidth - .Left - 15
        .ColWidth(1) = .Width - .ColWidth(0) - Me.SysInfo.ScrollBarSize - 15
    End With
    Call zlGrdRowHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\分割", "横向", Me.picVBar.Left)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\分割", "纵向", Me.picHBar.Top)
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call frmVItemEdit.ShowMe(Me, 2, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
End Sub

Private Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Err = 0: On Error GoTo ErrHand
    '------------------------------------------------
    gstrSql = "select 临床意义,表示法,性别域,数值域,初始值,文字表述,空值文字,必填 from 诊治所见项目 where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Mid(Item.Key, 2))
    
    With rsTemp
        If .RecordCount > 0 Then
            Me.hgdInfo.TextMatrix(0, 1) = IIf(IsNull(!临床意义), "", !临床意义)
            Select Case IIf(IsNull(!性别域), 0, !性别域)
            Case 0
                Me.hgdInfo.TextMatrix(1, 1) = "无性别限制"
            Case 1
                Me.hgdInfo.TextMatrix(1, 1) = "限男性使用"
            Case 2
                Me.hgdInfo.TextMatrix(1, 1) = "限女性使用"
            End Select
            Select Case IIf(IsNull(!表示法), 0, !表示法)
            Case 0
                Me.hgdInfo.TextMatrix(2, 1) = "文本框"
            Case 1
                Me.hgdInfo.TextMatrix(2, 1) = "上下按钮"
            Case 2
                Me.hgdInfo.TextMatrix(2, 1) = "下拉选择"
            Case 3
                Me.hgdInfo.TextMatrix(2, 1) = "复选按钮"
            Case 4
                Me.hgdInfo.TextMatrix(2, 1) = "单选按钮"
            End Select
            Me.hgdInfo.TextMatrix(3, 1) = IIf(IsNull(!数值域), "", !数值域)
'            Me.hgdInfo.TextMatrix(4, 1) = IIf(IsNull(!初始值), "", !初始值)
'            Select Case IIf(IsNull(!文字表述), 0, !文字表述)
'            Case 0
'                Me.hgdInfo.TextMatrix(5, 1) = "项目名+项目值+单位"
'            Case 1
'                Me.hgdInfo.TextMatrix(5, 1) = "项目值+单位+项目名"
'            Case 2
'                Me.hgdInfo.TextMatrix(5, 1) = "项目值+单位"
'            End Select
'            strTemp = IIf(IsNull(!空值文字), "", !空值文字)
'            If Trim(strTemp) <> "" Then
'                Me.hgdInfo.TextMatrix(5, 1) = Me.hgdInfo.TextMatrix(5, 1) & "，空值时表述为“" & strTemp & "”"
'            End If
            Me.hgdInfo.TextMatrix(4, 1) = IIf(!必填 = 0, "否", "是")
        Else
            Call zlClearDetail
        End If
    End With
    Call zlGrdRowHeight
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_DblClick
End Sub

Private Sub lvwItems_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And Me.mnuEdit.Visible Then Call PopupMenu(Me.mnuEdit, 2)
End Sub

Private Sub mnuClassAdd_Click()
    With frmVItemClass
        strTemp = Me.cmdKind(Val(Me.tvwClass.Tag)).Caption
        .lblKind.Tag = IIf(Val(Me.tvwClass.Tag) < 2, Val(Me.tvwClass.Tag), Val(Me.tvwClass.Tag) + 1) + 1
        If .lblKind.Tag = 5 Then .lblKind.Tag = 6
        
        .lblKind.Caption = Mid(strTemp, 1, Len(strTemp) - 4)
        If Me.tvwClass.SelectedItem Is Nothing Then
            .txtParent.Tag = 0
        Else
            .txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        End If
        .Tag = "增加"
        .Show 1, Me
    End With
    If Me.tvwClass.SelectedItem Is Nothing Then
        Call zlRefClasses
    Else
        Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
    End If
End Sub

Private Sub mnuClassDel_Click()
    Err = 0: On Error GoTo ErrHand
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("真的删除该分类“" & Me.tvwClass.SelectedItem.Text & "”吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSql = "ZL_所见分类_DELETE(" & Mid(Me.tvwClass.SelectedItem.Key, 2) & ")"
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    Dim strParentKey As String
    If Me.tvwClass.SelectedItem.Next Is Nothing Then
        If Me.tvwClass.SelectedItem.Parent Is Nothing Then
            Call zlRefClasses
        Else
            strParentKey = Me.tvwClass.SelectedItem.Parent.Key
            Call Me.tvwClass.Nodes.Remove(Me.tvwClass.SelectedItem.Key)
            If Me.tvwClass.Nodes(strParentKey).Children = 0 Then
                Call zlRefClasses(Mid(Me.tvwClass.Nodes(strParentKey).Key, 2))
            Else
                Call zlRefClasses(Mid(Me.tvwClass.Nodes(strParentKey).Child.Key, 2))
            End If
        End If
    Else
        Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Next.Key, 2))
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuClassMod_Click()
    With frmVItemClass
        strTemp = Me.cmdKind(Val(Me.tvwClass.Tag)).Caption
        .lblKind.Tag = IIf(Val(Me.tvwClass.Tag) < 2, Val(Me.tvwClass.Tag), Val(Me.tvwClass.Tag) + 1) + 1
        If .lblKind.Tag = 5 Then .lblKind.Tag = 6
        .lblKind.Caption = Mid(strTemp, 1, Len(strTemp) - 4)
        If Me.tvwClass.SelectedItem.Parent Is Nothing Then
            .txtParent.Tag = 0
            .txtParent.Text = "(无)"
            .txtUpCode.Text = ""
            .txtCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Text, "]")(0), 2)
            .txtCode.MaxLength = Len(.txtCode.Text)
            .txtCode.Tag = .txtCode.MaxLength
        Else
            .txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Parent.Key, 2)
            .txtParent.Text = Me.tvwClass.SelectedItem.Parent.Text
            .txtUpCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Parent.Text, "]")(0), 2)
            .txtCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Text, "]")(0), Len(.txtUpCode.Text) + 2)
            .txtCode.MaxLength = Len(.txtCode.Text)
            .txtCode.Tag = .txtCode.MaxLength
        End If
        .txtName = Split(Me.tvwClass.SelectedItem.Text, "]")(1)
        .txtSymbol = Me.tvwClass.SelectedItem.Tag
        .Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        .Show 1, Me
    End With
    Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
End Sub

Private Sub mnuEditAdd_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "尚未设置分类,不能增删项目！", vbExclamation, gstrSysName: Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then
        Call frmVItemEdit.ShowMe(Me, 0, Val(Mid(Me.tvwClass.SelectedItem.Key, 2)))
    Else
        Call frmVItemEdit.ShowMe(Me, 0, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
    End If
    Call zlRefRecords
End Sub

Private Sub mnuEditDelete_Click()
    On Error GoTo ErrHand
    With Me.lvwItems
        If .SelectedItem Is Nothing Then Exit Sub
        If InStr(mnuEditDelete.Caption, "删除") > 0 Then  '删除
            If MsgBox("真的删除“" & .SelectedItem.Text & "”吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSql = "ZL_所见项目_DELETE(" & Mid(.SelectedItem.Key, 2) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            Call .ListItems.Remove(.SelectedItem.Key)
            If .SelectedItem Is Nothing Then
                Call zlClearDetail
            Else
                Call lvwItems_ItemClick(.SelectedItem)
            End If
        Else                                    '标记为非必填项
            Call EarMarkMustItem(Mid(.SelectedItem.Key, 2), False)
            Call zlRefRecords
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditModify_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "尚未设置分类,不能增删项目！", vbExclamation, gstrSysName: Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    If InStr(mnuEditModify.Caption, "修改") > 0 Then '修改
        Call frmVItemEdit.ShowMe(Me, 1, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
    Else                                        '标记为必填项
        Call EarMarkMustItem(Mid(lvwItems.SelectedItem.Key, 2), True)
    End If
    
    Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
End Sub

Private Sub mnuFileExcel_Click()
    Call zlRptPrint(3)
End Sub

Private Sub mnuFilePreview_Click()
    Call zlRptPrint(0)
End Sub

Private Sub mnuFilePrint_Click()
    Call zlRptPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuhelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    '默认参数：分类=分类id，项目=项目id
    Dim lng分类id As Long
    Dim lng项目id As Long
    
    If Not Me.tvwClass.SelectedItem Is Nothing Then
        lng分类id = Mid(Me.tvwClass.SelectedItem.Key, 2)
    End If
    
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        lng项目id = Mid(lvwItems.SelectedItem.Key, 2)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "分类=" & IIf(lng分类id = 0, "", lng分类id), _
        "项目=" & IIf(lng项目id = 0, "", lng项目id))
End Sub

Private Sub mnuViewFind_Click()
    With frmVItemFind
        .Show , Me
    End With
End Sub

Private Sub mnuViewRefresh_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Call zlRefRecords
End Sub

Private Sub mnuViewStates_Click()
    Me.mnuViewStates.Checked = Not Me.mnuViewStates.Checked
    Me.stbThis.Visible = Me.mnuViewStates.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolbarStand_Click()
    Me.mnuViewToolbarStand.Checked = Not Me.mnuViewToolbarStand.Checked
    Me.clbThis.Visible = Me.mnuViewToolbarStand.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolBarText_Click()
    Dim i As Integer
    Me.mnuViewToolbarText.Checked = Not Me.mnuViewToolbarText.Checked
    If Me.mnuViewToolbarText.Checked Then
        For i = 1 To Me.tlbThis.Buttons.Count
            Me.tlbThis.Buttons(i).Caption = Me.tlbThis.Buttons(i).Tag
        Next
    Else
        For i = 1 To Me.tlbThis.Buttons.Count
            Me.tlbThis.Buttons(i).Caption = ""
        Next
    End If
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Form_Resize
End Sub

Private Sub picClass_Resize()
    Dim intCount As Integer
    Err = 0: On Error Resume Next
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        Me.cmdKind(intCount).Left = Me.picClass.ScaleLeft + 15
        Me.cmdKind(intCount).Width = Me.picClass.ScaleWidth
        Me.cmdKind(intCount).Height = 300
        If Val(Me.cmdKind(intCount).Tag) = 0 Then
            Me.cmdKind(intCount).Top = Me.picClass.ScaleTop + 285 * intCount
            Me.tvwClass.Top = Me.picClass.ScaleTop + 285 * (intCount + 1)
        Else
            Me.cmdKind(intCount).Top = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound - intCount + 1)
        End If
    Next
    Me.tvwClass.Left = Me.picClass.ScaleLeft + 15
    Me.tvwClass.Width = Me.picClass.ScaleWidth
    Me.tvwClass.Height = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound + 1) - 15
End Sub

Private Sub picHBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.picHBar.Top = Me.picHBar.Top + y
End Sub

Private Sub picHBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Call Form_Resize
End Sub

Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.picVBar.Left = Me.picVBar.Left + x
End Sub

Private Sub picVBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Call Form_Resize
End Sub

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        Call mnuFilePreview_Click
    Case "Print"
        Call mnuFilePrint_Click
    Case "Class"
        If Me.mnuClass.Visible Then Call PopupMenu(Me.mnuClass, 2)
    Case "Add"
        Call mnuEditAdd_Click
    Case "Modify"
        Call mnuEditModify_Click
    Case "Delete"
        Call mnuEditDelete_Click
    Case "Find"
        Call mnuViewFind_Click
    Case "Help"
        Call mnuHelpHelp_Click
    Case "Exit"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tlbThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu Me.mnuViewToolbar, 2
End Sub

Private Sub tvwClass_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And Me.mnuClass.Visible Then Call PopupMenu(Me.mnuClass, 2)
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    If Me.lvwItems.Tag = Node.Key Then Exit Sub
    Me.lvwItems.Tag = Node.Key
    Call zlRefRecords
End Sub

Private Sub zlRefClasses(Optional lngNode As Long)
    '---------------------------------------------
    '填写诊疗分类项目(此处为药品分类)并按照不同类型调整界面
    '---------------------------------------------
    Dim lngTmp As Long
    '权限控制
    
    '调整显示界面
    If Val(Me.tvwClass.Tag) = 0 Then '固定项目
        Me.mnuClass.Enabled = False: Me.mnuClassAdd.Enabled = False: Me.mnuClassMod.Enabled = False: Me.mnuClassDel.Enabled = False
        'Me.mnuEdit.Enabled = False: Me.mnuEditAdd.Enabled = False: Me.mnuEditModify.Enabled = False: Me.mnuEditDelete.Enabled = False
        '将项目菜单改为必填项目/非必填项
        Me.mnuEditAdd.Visible = False: Me.mnuEditModify.Caption = "必填(&M)": Me.mnuEditDelete.Caption = "可选(&O)"
        Me.tlbThis.Buttons("Class").Enabled = False
        Me.tlbThis.Buttons("Add").Visible = False: Me.tlbThis.Buttons("Modify").Caption = "必填": Me.tlbThis.Buttons("Delete").Caption = "可选"
    Else                            '可更改项目
        Me.mnuClass.Enabled = True: Me.mnuClassAdd.Enabled = True: Me.mnuClassMod.Enabled = True: Me.mnuClassDel.Enabled = True
        Me.mnuEditAdd.Visible = True: Me.mnuEditModify.Caption = "修改(&M)": Me.mnuEditDelete.Caption = "删除(&D)"
        Me.tlbThis.Buttons("Class").Enabled = True
        Me.tlbThis.Buttons("Add").Visible = True: Me.tlbThis.Buttons("Modify").Caption = "修改": Me.tlbThis.Buttons("Delete").Caption = "删除"
    End If
    
    Me.lvwItems.ListItems.Clear
    '填写分类
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select ID,上级ID,编码,名称,简码" & _
            " From 诊治所见分类" & _
            " Where 性质 = [1] " & _
            " start with 上级ID is null" & _
            " connect by prior ID=上级ID"
    lngTmp = 1 + IIf(Val(Me.tvwClass.Tag) < 2, Val(Me.tvwClass.Tag), Val(Me.tvwClass.Tag) + 1)
    lngTmp = IIf(lngTmp = 5, 6, lngTmp)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngTmp)
    
    With rsTemp
        Me.tvwClass.Visible = False
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            End If
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!简码), "", !简码)
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        Me.tvwClass.Visible = True
    End With
    If Me.tvwClass.Nodes.Count > 0 Then
        If lngNode <> 0 Then
            Me.tvwClass.Nodes("_" & lngNode).Selected = True
        Else
            Me.tvwClass.Nodes(1).Selected = True
        End If
        Call zlRefRecords
    Else
        Call zlClearDetail
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlRefRecords(Optional lngItem As Long)
    '---------------------------------------------
    '填写项目列表
    '---------------------------------------------
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select  I.ID,I.编码,I.中文名,I.英文名,I.类型,I.长度,I.小数,I.小数,I.单位,I.必填" & _
            " from 诊治所见项目 I" & _
            " where I.分类ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Mid(Me.tvwClass.SelectedItem.Key, 2))
    
    With rsTemp
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !中文名)
            objItem.Icon = "item": objItem.SmallIcon = "item"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_英文名").Index - 1) = IIf(IsNull(!英文名), "", !英文名)
            Select Case IIf(IsNull(!类型), 0, !类型)
            Case 0
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1) = "数值"
            Case 1
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1) = "文字"
            Case 2
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1) = "日期"
            Case 3
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1) = "逻辑"
            End Select
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_长度").Index - 1) = IIf(IsNull(!长度), "", !长度)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_小数").Index - 1) = IIf(IsNull(!小数), "", !小数)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_必填").Index - 1) = IIf(!必填 = 0, "否", "是")
            If !ID = lngItem Then
                objItem.Selected = True
            End If
            .MoveNext
        Loop
    
    End With
    If Me.lvwItems.ListItems.Count > 0 Then
        If Me.lvwItems.SelectedItem Is Nothing Then Me.lvwItems.ListItems(1).Selected = True
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
        Err = 0: On Error Resume Next
        DoEvents: Me.lvwItems.SelectedItem.EnsureVisible
        Me.stbThis.Panels(2).Text = "该分类共有" & Me.lvwItems.ListItems.Count & "个项目"
    Else
        Call zlClearDetail
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlClearDetail()
    '---------------------------------------------
    '清理调整详细信息显示区域
    '---------------------------------------------
    With Me.hgdInfo
        .TextMatrix(0, 1) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(2, 1) = ""
        .TextMatrix(3, 1) = ""
        .TextMatrix(4, 1) = ""
        .TextMatrix(5, 1) = ""
        .TextMatrix(6, 1) = ""
    End With
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:记录表打印
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrintLvw
    On Error Resume Next
    Set objPrint.Body.objData = Me.lvwItems
    strTemp = Me.cmdKind(Val(Me.tvwClass.Tag)).Caption
    objPrint.Title.Text = Mid(strTemp, 1, Len(strTemp) - 4) & "所见项目清单"
    
    objPrint.UnderAppItems.Add "分类：" & Me.tvwClass.SelectedItem.Text
    objPrint.BelowAppItems.Add "打印时间：" & Now
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrViewLvw objPrint, bytMode
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Private Sub zlGrdRowHeight()
    '---------------------------------------------
    '根据调整内容调整内容网格的行高度，以保证内容的正常显示
    '---------------------------------------------
    Dim intRow As Integer, lngColWidth As Long
    With Me.hgdInfo
        For intRow = .FixedRows To .Rows - 1
            lngColWidth = .ColWidth(1)
            Me.lblScale.Width = lngColWidth - 90
            Me.lblScale.Caption = .TextMatrix(intRow, 1)
            .RowHeight(intRow) = Me.lblScale.Height + 75
        Next
    End With
End Sub

Public Sub zlLocateItem(lngClassId As Long, lngItemID As Long)
    '---------------------------------------------
    '定位到指定的项目，在查找时使用
    '---------------------------------------------
    Set Me.tvwClass.SelectedItem = Me.tvwClass.Nodes("_" & lngClassId)
    Me.tvwClass.Nodes("_" & lngClassId).Selected = True
    Me.tvwClass.SelectedItem.EnsureVisible
    Call zlRefRecords
    Set Me.lvwItems.SelectedItem = Me.lvwItems.ListItems("_" & lngItemID)
    Me.lvwItems.SelectedItem.EnsureVisible
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub
Private Sub EarMarkMustItem(ByVal lngItemID As Long, ByVal ItemMust As Boolean)
Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    gstrSql = "Select 分类id, 编码, 中文名, 英文名, 类型, 长度, 小数, 单位, 临床意义, 表示法, 性别域, 数值域, 初始值, 文字表述, 空值文字,动态域" & vbNewLine & _
                "From 诊治所见项目" & vbNewLine & _
                "Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    With rsTemp
        If .EOF Then Exit Sub
        gstrSql = Nvl(!分类id, 0) & ",'" & !编码 & "','" & !中文名 & "','" & !英文名 & "'," & Nvl(!类型, 0) & _
                "," & Nvl(!长度, 0) & "," & Nvl(!小数, 0) & ",'" & !单位 & "','" & !临床意义 & "'," & Nvl(!表示法, 0) & _
                "," & Nvl(!性别域, 0) & ",'" & !数值域 & "','" & !初始值 & "'," & Nvl(!文字表述, 1) & ",'" & !空值文字 & "'," & IIf(ItemMust, 1, 0) & "," & Nvl(!动态域, 0)
    End With
    gstrSql = "ZL_所见项目_UPDATE(" & lngItemID & "," & gstrSql & ")"
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

