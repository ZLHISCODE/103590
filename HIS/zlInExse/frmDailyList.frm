VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDailyList 
   Caption         =   "一日费用清单"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   -495
   ClientWidth     =   8880
   Icon            =   "frmDailyList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboPage 
      Height          =   300
      Left            =   6675
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   855
      Width           =   1410
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3855
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "日期：2000年10月20日"
      Top             =   885
      Width           =   2475
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4785
      Left            =   3870
      TabIndex        =   4
      Top             =   1170
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   8440
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   280
      BackColorBkg    =   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ListView lvwPati 
      Height          =   5160
      Left            =   15
      TabIndex        =   0
      Top             =   795
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   9102
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483628
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "姓名"
         Text            =   "姓名"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "病员号"
         Text            =   "住院号"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "床号"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "费别"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "性别"
         Text            =   "性别"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "入院日期"
         Text            =   "入院日期"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "出院日期"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "病区"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "病人类型"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.StatusBar sta状态 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6210
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDailyList.frx":030A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8281
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "病人颜色"
            TextSave        =   "病人颜色"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
            Object.ToolTipText     =   "当前数字键状态"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "当前大写键状态"
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   8880
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   3705
      NewRow1         =   0   'False
      Caption2        =   "病人病区"
      Child2          =   "cbo病区"
      MinHeight2      =   300
      Width2          =   1995
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   165
         TabIndex        =   6
         Top             =   30
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgTbrStard"
         HotImageList    =   "imgTbrHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "打印"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "重置"
               Key             =   "重置"
               Object.ToolTipText     =   "重置条件"
               Object.Tag             =   "重置"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "查找"
               Object.ToolTipText     =   "查找"
               Object.Tag             =   "查找"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "字体"
               Key             =   "字体"
               Object.ToolTipText     =   "字体"
               Object.Tag             =   "字体"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Description     =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cbo病区 
         Height          =   300
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   4110
      End
   End
   Begin MSComctlLib.ImageList imgTbrHot 
      Left            =   2145
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":0B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":0DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":0FD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":11F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":140C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":1626
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":1840
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":1A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":1C78
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":1E94
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":258E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTbrStard 
      Left            =   1410
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":27A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":29C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":2BE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":2DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":3016
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":3230
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":344A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":3666
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":3882
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":3A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":4198
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic分隔 
      Height          =   5340
      Left            =   3510
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5340
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   720
      Width           =   45
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   3255
      Top             =   5385
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
            Picture         =   "frmDailyList.frx":43B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":46CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   3210
      Top             =   4710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":4FA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyList.frx":52C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
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
      Begin VB.Menu mnuFileOpen 
         Caption         =   "条件重置(&J)"
      End
      Begin VB.Menu mnuFileLine 
         Caption         =   "-"
      End
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
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewQuitFee 
         Caption         =   "显示退费(&Q)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewZero 
         Caption         =   "显示零费用(&Z)"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuViewFindLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewFindNext 
         Caption         =   "查找下一个(&N)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "字体(&F)"
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "小字体"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "中字体"
            Index           =   1
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "大字体"
            Index           =   2
         End
      End
      Begin VB.Menu mnuViewP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewALLSele 
         Caption         =   "全选(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuViewALLClear 
         Caption         =   "全清(&C)"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web上的中联"
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
      Begin VB.Menu mnuHelp_Line 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp_about 
         Caption         =   "关于(&A)…"
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuPopDisp 
         Caption         =   "详细资料(&D)"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmDailyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mstrPrivs As String
Private mlngModul As Long

Private mrsPati As New ADODB.Recordset
Private mintBedLen As Integer
Private mblnPrint As Boolean '控清单是否打印
Private mdtMin As Date, mdtMax As Date
Private mbln非医保病人 As Boolean
Private mbln医保病人 As Boolean
Private mbln在院病人 As Boolean
Private mbln出院病人 As Boolean
Private mstr费用时间 As String
Private mbyt病人病区模式 As Byte '0-有费用的病区(缺省)、1-病人当前病区

Private Sub cboPage_Change()
    Refresh费用清单 lvwPati.SelectedItem, Val(cboPage.ItemData(cboPage.ListIndex))
End Sub

Private Sub cboPage_Click()
    Refresh费用清单 lvwPati.SelectedItem, Val(cboPage.ItemData(cboPage.ListIndex))
End Sub

Private Sub cbo病区_Click()
    If cbo病区.ListIndex = -1 Then Exit Sub
    ReFresh病人信息
'    If lvwPati.ListItems.Count > 0 Then
'        lvwPati_ItemClick lvwPati.ListItems(1)
'    End If
End Sub

Private Sub cbrThis_HeightChanged(ByVal NewHeight As Single)
     Form_Resize
End Sub

Private Sub lvwALLCLear(ByVal blnCheck As Boolean)
    Dim itm As ListItem
    For Each itm In lvwPati.ListItems
        itm.Checked = blnCheck
    Next
End Sub

Private Sub Form_Activate()
    If cbo病区.ListCount = 0 Then
        MsgBox "没有可司职的病区(未初始或权限不具备)", vbExclamation, "提示"
        Unload Me: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim lngTmp As Long
    Dim strStartTime  As String
    Dim strEndTime As String
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnLimitUnit As Boolean
    Dim strUnitIDs As String
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    
    RestoreWinState Me, App.ProductName
    
    mnuViewQuitFee.Checked = zlDatabase.GetPara("显示退费", glngSys, mlngModul) = "1"
    mnuViewZero.Checked = zlDatabase.GetPara("显示零费用", glngSys, mlngModul) = "1"
    mstr费用时间 = IIf(zlDatabase.GetPara("费用时间", glngSys, mlngModul) = "1", "发生时间", "登记时间") '注册表值为1表示按发生时间
    
    If InStr(mstrPrivs, ";参数设置;") = 0 Then
        mnuViewQuitFee.Enabled = False
        mnuViewZero.Enabled = False
    End If
    
    
    mbyt病人病区模式 = IIf(zlDatabase.GetPara("病人病区模式", glngSys, mlngModul, "0") = "1", 1, 0)
    
    strEndTime = zlDatabase.GetPara("结束时间", glngSys, mlngModul, "23:59:59")
    lngTmp = Val(zlDatabase.GetPara("结束间隔", glngSys, mlngModul, 0))
    If lngTmp > 7 Then lngTmp = 7
    mdtMax = CDate(Format(zlDatabase.Currentdate() - lngTmp, "yyyy-MM-dd") & " " & strEndTime)
    
    strStartTime = zlDatabase.GetPara("开始时间", glngSys, mlngModul, "00:00:00")
    lngTmp = Val(zlDatabase.GetPara("开始间隔", glngSys, mlngModul, 0))
    If lngTmp > 7 Then lngTmp = 7
    mdtMin = CDate(Format(mdtMax - lngTmp, "yyyy-MM-dd") & " " & strStartTime)
    
    mbln非医保病人 = zlDatabase.GetPara("非医保病人", glngSys, mlngModul, "1") = "1"
    mbln医保病人 = zlDatabase.GetPara("医保病人", glngSys, mlngModul, "1") = "1"
    
        
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs, "ZL" & glngSys \ 100 & "_INSIDE_1141")
    
    If InStr(";" & mstrPrivs, ";出院病人查询;") = 0 Then
        mbln在院病人 = True
        mbln出院病人 = False
    Else
        mbln在院病人 = zlDatabase.GetPara("在院病人", glngSys, mlngModul, "1") = "1"
        mbln出院病人 = zlDatabase.GetPara("出院病人", glngSys, mlngModul, "1") = "1"
    End If
    
    mblnPrint = True
    If InStr(";" & mstrPrivs, ";清单打印;") = 0 Then '判断清单打印权限
        mblnPrint = False
    End If
    
    
    txtDate.Text = "日期：" & Format(mdtMin, "yyyy年MM月DD日 hh:mm:ss") & "～" & Format(mdtMax, "yyyy年MM月DD日 hh:mm:ss")
    txtDate.Tag = txtDate.Text
    
    
    cbo病区.Clear
    If InStr(";" & mstrPrivs, ";所有病区;") > 0 Then cbo病区.AddItem "所有病区"
    Set rsTmp = GetUnit(InStr(mstrPrivs, ";所有病区;") = 0, "1,2,3", "护理")
    With rsTmp
        Do While Not .EOF
            cbo病区.AddItem !名称
            cbo病区.ItemData(cbo病区.NewIndex) = !ID
            If !ID = UserInfo.部门ID Then cbo病区.ListIndex = cbo病区.NewIndex
            .MoveNext
        Loop
        If cbo病区.ListIndex = -1 And cbo病区.ListCount > 0 Then cbo病区.ListIndex = 0
    End With
End Sub

Private Sub Form_Resize()
    Dim intHeightTbr As Integer, intHeightStb
   
    If WindowState = 1 Then Exit Sub
    On Error Resume Next
    intHeightTbr = IIf(cbrThis.Visible, cbrThis.Height, 0)
    intHeightStb = IIf(sta状态.Visible, sta状态.Height, 0)
    
    pic分隔.Top = 0
    pic分隔.Height = ScaleHeight
    If pic分隔.Left < 1000 Then pic分隔.Left = 1000
    If ScaleWidth - pic分隔.Left < 1000 Then pic分隔.Left = ScaleWidth - 1000
    
    With lvwPati
        .Top = ScaleTop + intHeightTbr
        .Height = ScaleHeight - intHeightStb - .Top
        .Left = ScaleLeft
        .Width = pic分隔.Left - ScaleLeft
    End With
    
    With cboPage
        .Left = pic分隔.Left + pic分隔.Width
        .Top = ScaleTop + intHeightTbr + 15
    End With
    
    With txtDate
        .Top = ScaleTop + intHeightTbr + 45
        .Left = cboPage.Left + cboPage.Width + 120
        .Width = ScaleWidth - .Left
    End With
    
    With grdList
        .Top = txtDate.Top + txtDate.Height
        .Height = ScaleHeight - intHeightStb - .Top
        .Left = pic分隔.Left + pic分隔.Width
        .Width = ScaleWidth - .Left
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmDailyListAsk
        
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwPati.Sorted = True
    With lvwPati
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With
    lvwPati.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub lvwPati_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lvwPati.SelectedItem = Item
    Load住院次数 Val(Mid(Item.Key, 2, InStr(Mid(Item.Key, 2), "_") - 1)), Val(Mid(Item.Key, InStr(Mid(Item.Key, 2), "_") + 2))
    Refresh费用清单 lvwPati.SelectedItem, Val(Mid(Item.Key, InStr(Mid(Item.Key, 2), "_") + 2))
End Sub

Private Sub lvwPati_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If lvwPati.ListItems.Count = 0 Then Exit Sub
            PopupMenu mnuPop, 2
    End If
End Sub

Private Sub mnuExcel_Click()
    'GrdPrint 1
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    Screen.MousePointer = 11
    Call PrintContent(3, Split(lvwPati.SelectedItem.Key, "_")(1))
    Screen.MousePointer = 0
End Sub

Private Sub GrdMuchPrint(ByVal Item As ListItem)
    '---------------------------------------------------
    '功能：    根据屏幕组织表上附加项目，打印预览
    '参数：blnIsPreview false表示预览
    '返回：
    '---------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim blnMuch As Boolean
    Dim i As Long
    Dim old病人id As Long
    Dim blnNext As Boolean
    objPrint.Title.Text = GetUnitName & "一日清单"
    Set objRow = New zlTabAppRow
    objRow.Add "住院号：" & Item.ListSubItems(1).Text & _
        "      姓  名：" & Item.Text & _
        "        性别：" & Item.ListSubItems(3).Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "入院日：" & Item.ListSubItems(5).Text & _
        "  出院日：" & Item.ListSubItems(6).Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objRow.Add ""
    objRow.Add txtDate
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印时间:" & Format(zlDatabase.Currentdate, "yyyy年MM月DD日 HH:MM")
    objPrint.BelowAppRows.Add objRow
    Set objPrint.Body = grdList
     objPrint.PageFooter = 2
    zlPrintOrView1Grd objPrint, 1
    Set objPrint = Nothing
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    frmDailyListAsk.mlngModul = mlngModul
    frmDailyListAsk.mstrPrivs = mstrPrivs
    frmDailyListAsk.Show 1, Me
    If Not frmDailyListAsk.mblnAskOk Then
        Unload frmDailyListAsk
        Exit Sub
    End If
    
    If frmDailyListAsk.mblnDateMoved Then
        MsgBox "当前选择的时间范围内的费用可能位于离线数据表,以下功能将被禁用:" & vbCrLf & _
            "打印预览、打印、输出到Excel、打印所选病人." & vbCrLf & vbCrLf & _
            "如要进行这些操作,请尽可能选择离现在最近的日期或与系统管理员联系!", vbInformation, gstrSysName
        Me.mnuFilePrint.Enabled = False
        Me.mnuFilePrintView.Enabled = False
        Me.mnuExcel.Enabled = False
        tbrThis.Buttons(1).Enabled = False
        tbrThis.Buttons(2).Enabled = False
    Else
        Me.mnuFilePrint.Enabled = True
        Me.mnuFilePrintView.Enabled = True
        Me.mnuExcel.Enabled = True
        tbrThis.Buttons(1).Enabled = True
        tbrThis.Buttons(2).Enabled = True
    End If
    
    mstr费用时间 = IIf(zlDatabase.GetPara("费用时间", glngSys, mlngModul) = "1", "发生时间", "登记时间") '注册表值为1表示按发生时间
    
    With frmDailyListAsk
        mdtMin = .dtpBegin
        mdtMax = .dtpEnd
        mbln非医保病人 = .chkPatiType(0).Value = 1
        mbln医保病人 = .chkPatiType(1).Value = 1
        mbln在院病人 = .chkInOut(0).Value = 1
        mbln出院病人 = .chkInOut(1).Value = 1
        txtDate.Text = "日期：" & Format(mdtMin, "yyyy年MM月DD日 hh:mm:ss") & "～" & Format(mdtMax, "yyyy年MM月DD日 hh:mm:ss")
        txtDate.Tag = txtDate.Text
        mbyt病人病区模式 = IIf(.optUnit(0).Value = True, 0, 1)
    End With
    
    Call ReFresh病人信息
    
    If lvwPati.SelectedItem Is Nothing Then zlCommFun.StopFlash: Exit Sub
    If lvwPati.ListItems.Count = 0 Then zlCommFun.StopFlash: Exit Sub
    
    Call Refresh费用清单(lvwPati.SelectedItem)
End Sub

Private Sub mnuFilePrint_Click()
    Dim Item As ListItem
    Dim newPatiId As Long
    Dim blnNOSelect As Boolean
    Dim intPreIdx As Integer, lngCount As Long
    Dim lng主页ID As Long
    
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    
    intPreIdx = lvwPati.SelectedItem.Index
    For Each Item In lvwPati.ListItems
        If Item.Checked Then lngCount = lngCount + 1
    Next
    
    If lngCount > 0 Then
        If MsgBox("你确定要打印所选病人的一日费用清单!", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    Screen.MousePointer = 11
    blnNOSelect = True
    grdList.Redraw = False
    For Each Item In lvwPati.ListItems
        If Item.Checked Or (lngCount = 0 And Item Is lvwPati.SelectedItem) Then
            blnNOSelect = False
            
            Item.Selected = True
            Item.EnsureVisible
            Me.Refresh
            lng主页ID = 0
            If (lngCount = 1 And lvwPati.SelectedItem.Key = Item.Key) Or (lngCount = 0 And Item Is lvwPati.SelectedItem) Then
                If cboPage.ListIndex >= 0 Then
                    lng主页ID = cboPage.ItemData(cboPage.ListIndex)
                End If
            End If
            Call PrintContent(2, Split(Item.Key, "_")(1), lng主页ID)
        End If
    Next
    If blnNOSelect Then MsgBox "没有选择要打印清单的病人！", vbInformation, gstrSysName
    grdList.Redraw = True
    
    lvwPati.ListItems(intPreIdx).Selected = True
    lvwPati.SelectedItem.EnsureVisible
    
    Screen.MousePointer = 0
End Sub

Private Sub mnuFilePrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1141", Me)
End Sub

Private Sub mnuFilePrintView_Click()
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    
    Screen.MousePointer = 11
    Call PrintContent(1, Split(lvwPati.SelectedItem.Key, "_")(1), Val(cboPage.ItemData(cboPage.ListIndex)))
    Screen.MousePointer = 0
End Sub

Private Sub PrintContent(ByVal bytMode As Byte, ByVal str病人ID As String, Optional lng主页ID As Long = 0)
    ReportOpen gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1141", Me, "病人ID=" & str病人ID, _
        "开始时间=" & Format(mdtMin, "yyyy-MM-dd HH:mm:ss"), _
        "结束时间=" & Format(mdtMax, "yyyy-MM-dd HH:mm:ss"), _
        "显示退费=" & IIf(mnuViewQuitFee.Checked, "1", "0"), _
        "显示零费用=" & IIf(mnuViewZero.Checked, "1", "0"), _
        "病人病区=" & cbo病区.ItemData(cbo病区.ListIndex), _
        "主页ID=" & lng主页ID, _
        "费用时间=" & mstr费用时间, bytMode
End Sub

Private Sub mnuHelp_About_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub
Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuPopDisp_Click()
   If mnuPopDisp.Checked = False Then
        mnuPopDisp.Checked = True
        lvwPati.View = lvwReport
    Else
        mnuPopDisp.Checked = False
        lvwPati.View = lvwIcon
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim lng病人ID As Long, lng主页ID As Long, str住院号 As String, lng病区ID As Long
    
    If cbo病区.ListIndex <> -1 Then
        lng病区ID = cbo病区.ItemData(cbo病区.ListIndex)
    End If
    
    If Not lvwPati.SelectedItem Is Nothing Then
        lng病人ID = Val(Mid(lvwPati.SelectedItem.Key, 2, InStr(Mid(lvwPati.SelectedItem.Key, 2), "_") - 1))
        lng主页ID = Val(Mid(lvwPati.SelectedItem.Key, InStr(Mid(lvwPati.SelectedItem.Key, 2), "_") + 2))
        str住院号 = lvwPati.SelectedItem.SubItems(1)
        
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, "病区=" & lng病区ID, _
            "开始时间=" & Format(mdtMin, "yyyy-MM-dd HH:mm:ss "), _
            "结束时间=" & Format(mdtMax, "yyyy-MM-dd HH:mm:ss "), "住院号=" & str住院号)
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "病区=" & lng病区ID)
    End If
End Sub

Private Sub mnuViewALLClear_Click()
    lvwALLCLear False
End Sub

Private Sub mnuViewALLSele_Click()
    lvwALLCLear True
End Sub

Private Sub mnuViewFind_Click()
    Dim strBed As String
    
    Load frmPatiFeeFind
    
    With frmPatiFeeFind
        .Show 1, Me
        If Not gblnOK Then Unload frmPatiFeeFind: Exit Sub
                
        strBed = .txtBed.Text
        If mintBedLen - Len(strBed) > 0 Then
            strBed = String(mintBedLen - Len(strBed), " ") & strBed
        End If
        
        mrsPati.Filter = 0
        mrsPati.Filter = "住院号=" & Val(.txt住院号) & _
            IIf(Trim(.txtBed) = "", "", " Or 床号='" & strBed & "'") & _
            IIf(Trim(.txt姓名) = "", "", " Or 姓名 Like '" & gstrLike & Trim(.txt姓名) & "%'")
    End With
    Unload frmPatiFeeFind
    
    If mrsPati.RecordCount = 0 Then
        MsgBox "无此信息的病人!", vbInformation, gstrSysName
        mrsPati.Filter = 0: Exit Sub
    End If
    mrsPati.MoveFirst
    lvwPati.ListItems("_" & mrsPati!病人ID & "_" & mrsPati!主页ID).Selected = True
    lvwPati.SelectedItem.EnsureVisible
    Call lvwPati_ItemClick(lvwPati.SelectedItem)
End Sub

Private Sub mnuViewFindNext_Click()
    On Error Resume Next
    If mrsPati Is Nothing Then Exit Sub
    If mrsPati.RecordCount = 0 Or mrsPati.RecordCount = 1 Then Exit Sub
    If mrsPati.EOF Then
        mrsPati.MoveFirst
    Else
        mrsPati.MoveNext
        If mrsPati.EOF Then
            mrsPati.MoveFirst
        End If
    End If
    lvwPati.ListItems("_" & mrsPati!病人ID & "_" & mrsPati!主页ID).Selected = True
    lvwPati.ListItems("_" & mrsPati!病人ID & "_" & mrsPati!主页ID).EnsureVisible
    Call lvwPati_ItemClick(lvwPati.SelectedItem)
End Sub

Private Sub mnuViewFontSize_Click(Index As Integer)
    Dim i As Long
    For i = 0 To 2
        mnuViewFontSize(i).Checked = False
    Next
        mnuViewFontSize(Index).Checked = True
    Select Case Index
    Case 0
        lvwPati.Font.Size = 9
        grdList.Font.Size = 9
        grdList.FontFixed = 9
    Case 1
        lvwPati.Font.Size = 11
        grdList.Font.Size = 11
        grdList.FontFixed = 11
    Case 2
        lvwPati.Font.Size = 12
        grdList.Font.Size = 12
        grdList.FontFixed = 12
    End Select
    Form_Resize
End Sub

Private Sub mnuViewQuitFee_Click()
    mnuViewQuitFee.Checked = Not mnuViewQuitFee.Checked
    zlDatabase.SetPara "显示退费", IIf(mnuViewQuitFee.Checked, 1, 0), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    lvwPati_ItemClick lvwPati.SelectedItem
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    sta状态.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub
Private Sub mnuViewToolbarStAnd_Click()
    Dim intCount As Integer
    mnuViewToolbarStand.Checked = Not mnuViewToolbarStand.Checked
    mnuViewToolbarText.Enabled = mnuViewToolbarStand.Checked
    cbrThis.Visible = mnuViewToolbarStand.Checked
    If mnuViewToolbarText.Checked Then
        For intCount = 1 To tbrThis.Buttons.Count
            tbrThis.Buttons(intCount).Caption = tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To tbrThis.Buttons.Count
            tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    cbrThis.Bands(1).minHeight = tbrThis.Height
    cbrThis.Refresh
    Form_Resize
End Sub

Private Sub mnuViewToolbarText_Click()
    Dim intCount As Integer
    mnuViewToolbarText.Checked = Not mnuViewToolbarText.Checked
    If mnuViewToolbarText.Checked Then
        For intCount = 1 To tbrThis.Buttons.Count
            tbrThis.Buttons(intCount).Caption = tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To tbrThis.Buttons.Count
            tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    cbrThis.Bands(1).minHeight = tbrThis.Height
    cbrThis.Refresh
    Form_Resize
End Sub

Private Sub mnuViewZero_Click()
    mnuViewZero.Checked = Not mnuViewZero.Checked
    zlDatabase.SetPara "显示零费用", IIf(mnuViewZero.Checked, 1, 0), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    lvwPati_ItemClick lvwPati.SelectedItem
End Sub

Private Sub pic分隔_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        pic分隔.Left = pic分隔.Left + X
        Form_Resize
        Me.Refresh
    End If
End Sub

Private Sub sta状态_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Text = "病人颜色" Then Call zlDatabase.ShowPatiColorTip(Me)
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    With Button
        Select Case .Key
        Case "预览"
            mnuFilePrintView_Click
        Case "打印"
            mnuFilePrint_Click
        Case "重置"
            mnuFileOpen_Click
        Case "查找"
            mnuViewFind_Click
        Case "字体"
             PopupMenu mnuViewFont
        Case "帮助"
            mnuHelpTitle_Click
        Case "退出"
           mnuFileExit_Click
        End Select
    End With
  
End Sub

Private Sub Load住院次数(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
    Dim strSql As String, rsPage As ADODB.Recordset
    On Error GoTo errH
    strSql = "Select Distinct 主页ID From 病案主页 Where 病人ID = [1] And 病人性质 = 0 Order By 主页ID Desc"
    Set rsPage = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID)
    cboPage.Clear
    cboPage.AddItem "所有住院"
    cboPage.ItemData(cboPage.NewIndex) = 0
    Do While Not rsPage.EOF
        cboPage.AddItem "第" & Val(NVL(rsPage!主页ID)) & "次住院"
        cboPage.ItemData(cboPage.NewIndex) = Val(NVL(rsPage!主页ID))
        If lng主页ID = Val(NVL(rsPage!主页ID)) Then cboPage.ListIndex = cboPage.NewIndex
        rsPage.MoveNext
    Loop
    If cboPage.ListIndex < 0 Then cboPage.ListIndex = 0
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub Refresh费用清单(Item As ListItem, Optional ByVal lngPageID As Long = 0)
    Dim rsTmp As ADODB.Recordset
    Dim arrFields As Variant, strSql As String
    Dim lngRow As Long, lngCol As Integer
    Dim strTmp As String, i As Long
    
    Dim lng病人ID As Long, lng主页ID As Long, lng病区ID As Long, lngInsure As Long
    
    On Error GoTo errHandle
    
    lngInsure = Val("" & Item.Tag)
    lng病人ID = Val(Mid(Item.Key, 2, InStr(Mid(Item.Key, 2), "_") - 1))

    lng主页ID = lngPageID

    If mbyt病人病区模式 = 0 And cbo病区.ListIndex <> -1 Then '费用发生的病区
        lng病区ID = cbo病区.ItemData(cbo病区.ListIndex)
    End If
    
    '不隐藏退费:求出汇总到收费细目行,包括每次退费的数量,金额
    strSql = _
    " Select Mod(记录性质,10) as 记录性质,NO,Nvl(价格父号,序号) as 序号,收费细目ID," & _
    "       计算单位,Avg(Nvl(付数,1)) as 付数,Avg(数次) as 数次," & _
    "       Sum(应收金额) as 应收金额,Sum(实收金额) as 实收金额,发生时间,费用类型 " & _
    " From " & IIf(frmDailyListAsk.mblnDateMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录") & _
    " Where 记录状态<>0 And 记帐费用=1 And 病人ID=[1] " & IIf(lng主页ID = 0, "", " And 主页ID=[2] ") & _
            IIf(lng病区ID = 0, "", " And 病人病区ID=[3]") & _
    "       And " & mstr费用时间 & " Between [4] And [5]" & _
    " Group by Mod(记录性质,10),NO,记录状态,Nvl(价格父号,序号),收费细目ID,计算单位,执行状态,发生时间,费用类型 "
    
    '隐藏退费:求出汇总到收费细目行的剩余数量,金额
    If Not mnuViewQuitFee.Checked Then
            strSql = _
            " Select 记录性质,NO,序号,收费细目ID,计算单位," & _
            " Sum(付数) as 付数,Sum(数次) as 数次," & _
            " Sum(应收金额) as 应收金额,Sum(实收金额) as 实收金额,发生时间,费用类型" & _
            " From (" & strSql & ")" & _
            " Group by 记录性质,NO,序号,收费细目ID,计算单位,发生时间,费用类型"
    End If
    
    '是否显示零费用
    If mnuViewZero.Checked Then
        strSql = strSql & " Having Nvl(Sum(应收金额),0)<>0"
    Else
        strSql = strSql & " Having Nvl(Sum(实收金额),0)<>0"
    End If
    
    strSql = _
        " Select To_Char(L.发生时间,'YYYY-MM-DD') as 日期,L.NO as 单据号," & _
        " Nvl(X.名称,I.名称)||' '||I.规格||'   '||LTrim(To_Char(L.数次,'9999990.00000'))||L.计算单位||Decode(I.类别,'7','×'||L.付数||'付',NULL) as 项目," & _
        " LTrim(To_Char(L.实收金额,'9999999" & gstrDec & "')) as 金额,NVL(L.费用类型,I.费用类型) as 费用类型,N.名称 医保大类" & _
        " From 收费项目目录 I,(" & strSql & ") L,收费项目别名 X,保险支付项目 M,保险支付大类 N" & _
        " Where I.ID=L.收费细目ID And I.ID=X.收费细目ID(+) And X.码类(+)=1 And X.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        " And I.ID=M.收费细目ID(+) And M.险类(+)=[6] And M.大类ID=N.ID(+)" & vbNewLine & _
        " Order by 日期,单据号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID, lng主页ID, lng病区ID, mdtMin, mdtMax, lngInsure)
    
    grdList.Redraw = False
    grdList.Clear
    grdList.Rows = 2
    If Not rsTmp.EOF Then
        Set grdList.Recordset = rsTmp
        For i = 0 To grdList.Cols - 1
            grdList.ColAlignmentFixed(i) = 4
            Select Case grdList.TextMatrix(0, i)
                Case "金额"
                    grdList.ColAlignment(i) = 7
                    strTmp = strTmp & "," & i
                Case Else
                    grdList.ColAlignment(i) = 1
            End Select
        Next
        If strTmp <> "" Then
            grdList.Rows = grdList.Rows + 1
            arrFields = Split(Mid(strTmp, 2), ",")
            For i = 0 To grdList.Rows - 1
                If i <> grdList.Rows - 1 Then
                     For lngCol = 0 To UBound(arrFields)
                         grdList.TextMatrix(grdList.Rows - 1, arrFields(lngCol)) = Val(grdList.TextMatrix(grdList.Rows - 1, arrFields(lngCol))) + Val(grdList.TextMatrix(i, arrFields(lngCol)))
                         grdList.TextMatrix(grdList.Rows - 1, 0) = "合计"
                     Next
                End If
                Call RefreshGridColWidth(grdList, i)
            Next
            For lngCol = 0 To UBound(arrFields)
                grdList.TextMatrix(grdList.Rows - 1, arrFields(lngCol)) = Format(Val(grdList.TextMatrix(grdList.Rows - 1, arrFields(lngCol))), "####" & gstrDec & ";-####" & gstrDec & "; ;")
            Next
        End If
        If lngInsure = 0 Then
            grdList.ColWidth(MshGetColNum(grdList, "医保大类")) = 0
        Else
            grdList.ColWidth(MshGetColNum(grdList, "医保大类")) = grdList.ColWidth(MshGetColNum(grdList, "费用类型"))
        End If
    End If
    grdList.Row = 1: grdList.Col = 0
    grdList.ColSel = grdList.Cols - 1
    
    grdList.Redraw = True
    
    strTmp = GetPatientDue(lng病人ID)
    If Val(strTmp) <> 0 Then
        txtDate.Text = txtDate.Tag & "" & "，应收款:" & Format(strTmp, "0.00")
    Else
        txtDate.Text = txtDate.Tag
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ReFresh病人信息()
    Dim objItem As ListItem
    Dim strSql As String, i As Integer, lng病区ID As Long
    
    Call zlCommFun.ShowFlash("正在统计数据,请稍候 ...", Me)
    DoEvents

    On Error GoTo errHandle
    
    grdList.Clear
    lng病区ID = cbo病区.ItemData(cbo病区.ListIndex)
    mintBedLen = GetMaxBedLen(lng病区ID)


    strSql = " Where " & mstr费用时间 & " Between [1] And [2] And 记录状态 IN(1,2,3) And 记帐费用=1 And 记录性质 In (2,3,5) "
        
    If mbyt病人病区模式 = 0 Then
        If lng病区ID > 0 Then strSql = strSql & " And 病人病区id+0=[3]"
    End If
            
    If frmDailyListAsk.mblnDateMoved Then
       strSql = "Select Distinct 病人id From (Select 病人id From 住院费用记录 " & strSql & _
                " Union All Select 病人id From H住院费用记录 " & strSql & ")"
    Else
       strSql = "Select Distinct 病人id From 住院费用记录 " & strSql
    End If
    
    strSql = "" & _
    "   Select /*+ rule*/ I.病人id,P.主页id,nvl(P.姓名,I.姓名) as 姓名,I.住院号,LPAD(P.出院病床," & mintBedLen & ",' ') as 床号," & _
    "           nvl(P.性别,I.性别) as 性别,P.入院日期,P.出院日期,P.病人性质,P.费别,X.名称 as 病区,P.病人类型,P.险类" & _
    "   From 部门表 X,病人信息 I,病案主页 P,(" & strSql & " ) L" & _
    "   Where I.病人id=P.病人id and P.主页id=I.主页id and P.病人id=L.病人id And P.当前病区ID=X.ID" & _
    "           And (X.站点=[4] or X.站点 is NULL)"
        
    If mbyt病人病区模式 = 1 Then
        If lng病区ID > 0 Then strSql = strSql & " And P.当前病区ID+0=[3]"
    End If
    
    
    '在院或出院病人
    If mbln在院病人 And mbln出院病人 Then
    ElseIf mbln在院病人 Then
        strSql = strSql & " And P.出院日期 is NULL"
    ElseIf mbln出院病人 Then
        strSql = strSql & " And P.出院日期 is Not NULL"
    End If
    
    '医保或普通病人
    If mbln非医保病人 And mbln医保病人 Then
    ElseIf mbln非医保病人 Then
        strSql = strSql & " And P.险类 is NULL"
    ElseIf mbln医保病人 Then
        strSql = strSql & " And P.险类 is Not NULL"
    End If
    
    If cbo病区.ItemData(cbo病区.ListIndex) = 0 Then
        strSql = strSql & " Order BY 病区,LPAD(床号,10,' ')"
    Else
        strSql = strSql & " Order BY LPAD(床号,10,' ')"
    End If
    
    
    Set mrsPati = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mdtMin, mdtMax, lng病区ID, gstrNodeNo)
    With mrsPati
         If .RecordCount <> 0 And mblnPrint Then
            If Not frmDailyListAsk.mblnDateMoved Then
             tbrThis.Buttons.Item(1).Enabled = True
             tbrThis.Buttons.Item(2).Enabled = True
             mnuFilePrint.Enabled = True
             mnuFilePrintView.Enabled = True
             mnuExcel.Enabled = True
            End If
        Else
             tbrThis.Buttons.Item(1).Enabled = False
             tbrThis.Buttons.Item(2).Enabled = False
             mnuFilePrint.Enabled = False
             mnuFilePrintView.Enabled = False
             mnuExcel.Enabled = False
         End If
        
        .Filter = 0
        lvwPati.ListItems.Clear
        Do While Not .EOF
            If IIf(IsNull(!病人性质), 0, !病人性质) = 0 Then
                Set objItem = lvwPati.ListItems.Add(, "_" & !病人ID & "_" & !主页ID, !姓名, 1, 1)
            Else
                Set objItem = lvwPati.ListItems.Add(, "_" & !病人ID & "_" & !主页ID, !姓名, 2, 2)
            End If
            objItem.ListSubItems.Add , , IIf(IsNull(!住院号), "", !住院号)
            objItem.ListSubItems.Add , , IIf(IsNull(!床号), "", !床号)
            objItem.ListSubItems.Add , , IIf(IsNull(!费别), "", !费别)
            objItem.ListSubItems.Add , , IIf(IsNull(!性别), "", !性别)
            objItem.ListSubItems.Add , , Format(!入院日期, "yyyy-MM-DD")
            objItem.ListSubItems.Add , , Format(IIf(IsNull(!出院日期), Empty, !出院日期), "yyyy-MM-DD")
            objItem.ListSubItems.Add , , IIf(IsNull(!病区), "", !病区)
            objItem.ListSubItems.Add , , IIf(IsNull(!病人类型), "", !病人类型)
            objItem.Tag = Val("" & !险类)
        
            objItem.ForeColor = zlDatabase.GetPatiColor(NVL(!病人类型))
            For i = 1 To objItem.ListSubItems.Count
                objItem.ListSubItems(i).ForeColor = zlDatabase.GetPatiColor(NVL(!病人类型))
            Next
            .MoveNext
        Loop
    End With
    sta状态.Panels(2).Text = "共" & lvwPati.ListItems.Count & "人"
    If mrsPati.RecordCount = 0 Then Call RefreshListStru
    If Not lvwPati.SelectedItem Is Nothing Then
         lvwPati.SelectedItem.Selected = False
         Set lvwPati.SelectedItem = Nothing:
         Call RefreshListStru
    End If
    Call zlCommFun.StopFlash
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub RefreshListStru()
    '--------------------------------------------------------------
    '功能：获取病人费用清单的表头结构
    '参数：
    '返回：
    '--------------------------------------------------------------
   
   Dim intRow As Long
   Dim intCol As Long
    '0  表示费用清单
   
   With grdList
        .Redraw = False
        For intRow = 0 To .Cols - 1
            .MergeCol(intRow) = False
        Next
        For intRow = 1 To .Rows - 1
            .RowData(intRow) = 0
            .MergeRow(intRow) = False
        Next
        .Clear
        .Rows = 2
        .FixedRows = 1
        .RowHeight(0) = TextHeight("刘") * 2
        .MergeCells = flexMergeRestrictRows
        .Cols = 11
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7
        .ColAlignment(8) = 1
        .ColAlignment(9) = 1
        .ColAlignment(10) = 1
        
        .ColWidth(0) = 1400
        .ColWidth(1) = 800
        .ColWidth(2) = 1600
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
        .ColWidth(9) = 0
        .ColWidth(10) = 600
        
        .MergeCol(0) = False
        .MergeCol(1) = False
        .MergeCol(2) = False
        .MergeCol(3) = False
        .MergeCol(4) = False
        .MergeCol(5) = False
        .MergeCol(6) = False
        .MergeCol(7) = False
        .MergeCol(8) = False
        
        .TextMatrix(0, 0) = "日期"
        .TextMatrix(0, 1) = "单据号"
        .TextMatrix(0, 2) = "摘要"
        .TextMatrix(0, 3) = "费用类型"
        .TextMatrix(0, 4) = "收据项目"
        .TextMatrix(0, 5) = "单价"
        .TextMatrix(0, 6) = "应收金额"
        .TextMatrix(0, 7) = "实收金额"
        .TextMatrix(0, 8) = "科室"
        .TextMatrix(0, 9) = "操作员"
        .TextMatrix(0, 10) = "操作员"
        
        For intCol = 0 To .Cols - 1
          .ColAlignmentFixed(intCol) = 4
        Next
        .Redraw = True
     End With
    Call RefreshGridColWidth(grdList, 0)
End Sub

Private Sub GrdPrint(blnIsPreview As Byte)
    '---------------------------------------------------
    '功能：    根据屏幕组织表上附加项目，打印预览
    '参数：
    '     blnIsPreview: 0表示预览 1表示输出到EXCEL 其它表示打印
    '返回：
    '---------------------------------------------------
    '0 表示费用明细 1表示逐日汇总清单 2表示预交明细清单,3结帐明细,4 未结费用
    
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    objPrint.Title.Text = GetUnitName & "病人一日清单"
    Set objRow = New zlTabAppRow
    objRow.Add "住院号：" & lvwPati.SelectedItem.ListSubItems(1).Text & _
        "      姓  名：" & lvwPati.SelectedItem.Text & _
        "        性别：" & lvwPati.SelectedItem.ListSubItems(3).Text & _
        ""
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "入院日：" & lvwPati.SelectedItem.ListSubItems(5).Text
    objRow.Add "  出院日：" & lvwPati.SelectedItem.ListSubItems(6).Text
    objPrint.UnderAppRows.Add objRow
     
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objRow.Add ""
    objRow.Add txtDate
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印时间:" & Format(zlDatabase.Currentdate, "yyyy年MM月DD日")
    objPrint.BelowAppRows.Add objRow

    Set objPrint.Body = grdList
    objPrint.PageFooter = 2
    If blnIsPreview = 0 Then
        zlPrintOrView1Grd objPrint, 2
    Else
        If blnIsPreview = 1 Then
            zlPrintOrView1Grd objPrint, 3
        Else
            Select Case zlPrintAsk(objPrint)
            Case 1
                 zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
            End Select
        End If
    End If
    Set objPrint = Nothing
    
End Sub

Private Sub RefreshGridColWidth(ByVal objGrid As Object, lngRow As Long)
    Dim lngWidth As Long, lngCol As Long
    
    For lngCol = 0 To objGrid.Cols - 1
        lngWidth = Me.TextWidth(objGrid.TextMatrix(lngRow, lngCol) & "字")
        If objGrid.ColWidth(lngCol) <> 0 Then
            If objGrid.ColWidth(lngCol) < lngWidth Or lngRow = 0 Then
                objGrid.ColWidth(lngCol) = lngWidth
            End If
        End If
    Next
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewToolbar
    End If
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

