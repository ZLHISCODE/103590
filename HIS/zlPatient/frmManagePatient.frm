VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManagePatient 
   Caption         =   "病人信息管理"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   Icon            =   "frmManagePatient.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboNodeList 
      Height          =   300
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   870
      Width           =   2100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPati 
      Height          =   5325
      Left            =   2880
      TabIndex        =   6
      Top             =   1065
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   9393
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorSel    =   12632256
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmManagePatient.frx":06EA
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.TabStrip TabPatiState 
      Height          =   5655
      Left            =   2865
      TabIndex        =   5
      Top             =   750
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   9975
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "所有病人"
            Key             =   "T_所有病人"
            Object.Tag             =   "所有病人"
            Object.ToolTipText     =   "所有病人"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "在院病人"
            Key             =   "T_在院病人"
            Object.Tag             =   "在院病人"
            Object.ToolTipText     =   "在院病人"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "出院病人"
            Key             =   "T_出院病人"
            Object.Tag             =   "出院病人"
            Object.ToolTipText     =   "出院病人"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "门诊病人"
            Key             =   "T_门诊病人"
            Object.Tag             =   "门诊病人"
            Object.ToolTipText     =   "门诊病人"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "留观病人"
            Key             =   "T_留观病人"
            Object.Tag             =   "留观病人"
            Object.ToolTipText     =   "留观病人"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6390
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManagePatient.frx":0A04
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10901
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "病人类型"
            TextSave        =   "病人类型"
            Key             =   "PatiColor"
            Object.Tag             =   "PatiColor"
            Object.ToolTipText     =   "病人类型说明"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5595
      Left            =   2730
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5595
      ScaleWidth      =   45
      TabIndex        =   4
      Top             =   720
      Width           =   45
   End
   Begin MSComctlLib.TreeView tvwDist_s 
      Height          =   5175
      Left            =   -15
      TabIndex        =   3
      Top             =   1230
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   9128
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9975
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   9855
         _ExtentX        =   17383
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
            NumButtons      =   21
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
               Key             =   "Add"
               Description     =   "登记"
               Object.ToolTipText     =   "登记新病人信息"
               Object.Tag             =   "登记"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modi"
               Description     =   "修改"
               Object.ToolTipText     =   "修改当前选中病人信息"
               Object.Tag             =   "修改"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Del"
               Description     =   "删除"
               Object.ToolTipText     =   "删除当前选中病人信息"
               Object.Tag             =   "删除"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "合并"
               Key             =   "Merge"
               Description     =   "合并"
               Object.ToolTipText     =   "将当前选择病人的信息合并到另外一个病人中"
               Object.Tag             =   "合并"
               ImageKey        =   "Merge"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Merge_"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "卡片"
               Key             =   "View"
               Description     =   "卡片"
               Object.ToolTipText     =   "以卡片方式查阅当前病人信息"
               Object.Tag             =   "卡片"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "在当前病人清单中过滤满足条件的病人"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Go"
               Description     =   "定位"
               Object.ToolTipText     =   "定位到满点条件的病人上"
               Object.Tag             =   "定位"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "轧帐"
               Key             =   "轧帐"
               Object.ToolTipText     =   "收费轧帐"
               Object.Tag             =   "轧帐"
               ImageKey        =   "RollingCurtain"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SplitRollingCurtain"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "家属"
               Key             =   "Family"
               Description     =   "家属"
               Object.ToolTipText     =   "家属登记"
               Object.Tag             =   "家属"
               ImageKey        =   "Family"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "FamilyAdd"
                     Text            =   "家属登记"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "FamilyView"
                     Text            =   "家属信息"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "扩展"
               Key             =   "PlugIn"
               Object.ToolTipText     =   "扩展功能"
               Object.Tag             =   "扩展"
               ImageKey        =   "PlugIn"
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "-"
               Key             =   "FamilySplit"
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList imgGray 
      Left            =   645
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1296
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":14B0
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":16CA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":18E4
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1AFE
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1D18
            Key             =   "Merge"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":2412
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":2B0C
            Key             =   "Patis"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":3206
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":3420
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":363A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":3854
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":3A6E
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":D405
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":13C67
            Key             =   "PlugIn"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1260
      Top             =   240
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
            Picture         =   "frmManagePatient.frx":1A4C9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   120
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1A623
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1A83D
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1AA57
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1AC71
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1AE8B
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1B0A5
            Key             =   "Merge"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1B79F
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1BE99
            Key             =   "Patis"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1C593
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1C7AD
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1C9C7
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1CBE1
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1CDFB
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1D4F5
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":23D57
            Key             =   "PlugIn"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNode 
      AutoSize        =   -1  'True
      Caption         =   "站点"
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   930
      Width           =   360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFile_Preview 
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
      Begin VB.Menu mnuFilePrintMed 
         Caption         =   "打印病案(&M)"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRollingCurtain 
         Caption         =   "收费轧帐(&M)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuFileRollingCurtainSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileInsure 
         Caption         =   "保险类别(&I)"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuEdit_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Quit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit_Add 
         Caption         =   "登记(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Modi 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPatiInfo 
         Caption         =   "基本信息调整(&J)"
      End
      Begin VB.Menu mnuEdit_Split1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDelCard 
         Caption         =   "取消卡号绑定(&C)"
      End
      Begin VB.Menu mnuEditBlackList 
         Caption         =   "特殊病人(&T)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEdit_ToInPati 
         Caption         =   "转为住院病人(&I)"
      End
      Begin VB.Menu mnuEdit_Merge 
         Caption         =   "病人合并(&G)"
      End
      Begin VB.Menu mnuEdit_Surety 
         Caption         =   "担保信息(&B)"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuEdit_Merge_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Stop 
         Caption         =   "停用病人(&S)"
      End
      Begin VB.Menu mnuEdit_Restore 
         Caption         =   "取消停用(&R)"
      End
      Begin VB.Menu mnuEdit_Restore_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_QueryPass 
         Caption         =   "设置查询密码(&P)"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "身份卡片(&V)"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditMzReCalc 
         Caption         =   "按费别重算门诊费用(&F)"
      End
      Begin VB.Menu mnuEdit_Family 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_FamilyAdd 
         Caption         =   "家属登记"
      End
      Begin VB.Menu mnuEdit_FamilyView 
         Caption         =   "家属信息"
      End
      Begin VB.Menu mnuEdit_PlugIn 
         Caption         =   "扩展(&E)"
         Begin VB.Menu mnuEdit_PlugItem 
            Caption         =   "功能"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "查询(&Q)"
      Begin VB.Menu mnuQuery_ChangeLog 
         Caption         =   "病人信息变动日志(&C)"
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
         Begin VB.Menu mnuViewToolDist 
            Caption         =   "病人分布(&D)"
            Checked         =   -1  'True
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
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "定位(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStop 
         Caption         =   "显示停用病人(&P)"
      End
      Begin VB.Menu mnuViewPatiMode 
         Caption         =   "显示病人方式(&M)"
         Begin VB.Menu mnuViewByDept 
            Caption         =   "按病区显示(&U)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewByDept 
            Caption         =   "按科室显示(&D)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuView_3 
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
End
Attribute VB_Name = "frmManagePatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mrsPati As ADODB.Recordset
Private mblnMax As Boolean, mblnUnLoad As Boolean
Private mblnDown As Boolean, mblnGo As Boolean
Private mstrFilter As String, mstrFilterInfo As String, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mstrPrivs As String
Private mlngModul As Long
Private mlngCardType As Long  '缺省医疗卡类别
Private mbln是否取消绑定 As Boolean '设置是否可以执行取消卡号绑定操作（只有第三方卡才可以取消绑定卡操作）
Private mblnInitGrid As Boolean '是否完成表格初始化

Private mstrUserUnitIDs As String

Private Type Type_SQLCondition
    Default As Boolean          '是否是缺省进入，此时没有条件值,缺省值在mstrFilter中
    登记时间B As Date
    登记时间E As Date
    出生时间B As Date
    出生时间E As Date
    入院时间B As Date
    入院时间E As Date
    出院时间B As Date
    出院时间E As Date
    住院号 As String
    性别 As String
    费别 As String
    区域 As String
    医疗付款方式 As String
    Patient As String
End Type
Private SQLCondition As Type_SQLCondition
Private mstrPrivs_RollingCurtain As String  '收费轧帐管理权限

Private Sub cboNodeList_Click()
    Call InitUnits
    Call ShowPatis(mstrFilter, , gblnMyStyle, mstrFilterInfo)
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub mnuEdit_FamilyAdd_Click()
'功能:病人家属设置
    If Not CreatePublicPatient Then Exit Sub
    Call gobjPublicPatient.MakePatiFamily(Me, 0, 2, mlngModul) '编辑
End Sub

Private Sub mnuEdit_FamilyView_Click()
    Dim lng病人ID As Long
    
    If glngSys Like "8??" Then
        If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("客户ID"))) Then
            MsgBox "没有客户信息可以查看家属信息！", vbExclamation, gstrSysName: Exit Sub
        End If
    Else
        If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID"))) Then
            MsgBox "没有病人信息可以查看家属信息！", vbExclamation, gstrSysName: Exit Sub
        End If
    End If
    
    If glngSys Like "8??" Then
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("客户ID")))
    Else
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")))
    End If
    
    If Not CreatePublicPatient Then Exit Sub
    Call gobjPublicPatient.MakePatiFamily(Me, lng病人ID, 1, mlngModul) '查看
End Sub

Private Sub mnuEdit_PlugItem_Click(Index As Integer)
    Call ExcPlugInFun(mnuEdit_PlugItem(Index).Tag)
End Sub

Private Sub mnuEdit_QueryPass_Click()
    Dim strFirstPassWord As String, strSecPassWord As String, rsTemp As ADODB.Recordset, lng病人ID As Long
    Dim strPassWord As String
    Dim strPassInput As String
    
    On Error GoTo errH
    If glngSys Like "8??" Then
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("客户ID")))
    Else
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")))
    End If

    If InStr(mstrPrivs, "强制更改查询密码") <= 0 Then '有"强制更改查询密码"权限无需校验旧密码
        If frmInput.InputVal(Me, "原查询密码", "请输入原查询密码,如果无原密码直接确定。", strFirstPassWord, 3, 10, True, False, False, "*") Then
                strPassInput = zlCommFun.zlStringEncode(strFirstPassWord)
                
                '校验原密码
                gstrSQL = "select 查询密码 from 病人信息 where 病人ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取密码", lng病人ID)
                If Not rsTemp.EOF Then
                    If strPassInput <> Nvl(rsTemp!查询密码) Then
                        MsgBox "原查询密码输入错误，禁止修改，请检查！", vbExclamation, gstrSysName: Exit Sub
                    End If
                End If
        Else
            Exit Sub
        End If
    End If
    
    '输入新密码
    strFirstPassWord = "": strSecPassWord = ""
    If frmInput.InputVal(Me, "新查询密码", "请输入查询新密码，密码长度0～10位。" & vbCrLf & "该密码用在通过《病人自助查询》进行费用查询！", strFirstPassWord, 3, 10, True, False, False, "*") Then
        strPassWord = zlCommFun.zlStringEncode(strFirstPassWord)
        '再次确认新密码
        If frmInput.InputVal(Me, "确认新密码", "请再次输入新密码,以确认新密码" & vbCrLf & "该密码用在通过《病人自助查询》进行费用查询！", strSecPassWord, 3, 10, True, False, False, "*") Then
            strPassInput = zlCommFun.zlStringEncode(strSecPassWord)
            If strPassWord <> strPassInput Then
                '两次输入不一至
                MsgBox "前后两次输入的新密码不一至，此次密码设置未生效，请检查！", vbExclamation, gstrSysName: Exit Sub
            Else
                gstrSQL = "ZL_病人信息_UpdatePass(" & lng病人ID & ",'" & strPassWord & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "")
                MsgBox "密码修改成功！", vbInformation, gstrSysName
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Restore_Click()
    Dim intRow As Long, lng病人ID As Long
    Dim strSQL As String, i As Long
    Dim blnTrans As Boolean
    
    intRow = mshPati.Row
    
    If glngSys Like "8??" Then
        lng病人ID = Val(mshPati.TextMatrix(intRow, GetColNum("客户ID")))
        If lng病人ID = 0 Then
            MsgBox "没有客户信息可以取消停用！", vbExclamation, gstrSysName: Exit Sub
        End If
    Else
        lng病人ID = Val(mshPati.TextMatrix(intRow, GetColNum("病人ID")))
        If lng病人ID = 0 Then
            MsgBox "没有病人信息可以取消停用！", vbExclamation, gstrSysName: Exit Sub
        End If
    End If
    If MsgBox("确实要取消停用""" & mshPati.TextMatrix(intRow, GetColNum("姓名")) & """的信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    strSQL = "zl_病人信息_Restore(" & lng病人ID & ")"
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
'    Call SQLTest(App.ProductName, Me.Caption, strSQL)
'    gcnOracle.Execute strSQL, , adCmdStoredProc
'    Call SQLTest
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    '行直接处理
    mshPati.TextMatrix(intRow, GetColNum("停用时间")) = ""
    mshPati.Redraw = False
    For i = 0 To mshPati.Cols - 1
        mshPati.Col = i
        mshPati.CellForeColor = Me.ForeColor
    Next
    mshPati.Redraw = True
    mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    Call mshPati_EnterCell
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Stop_Click()
    Dim intRow As Long, lng病人ID As Long, int就诊次数 As Integer
    Dim strSQL As String, i As Long
    Dim blnTrans As Boolean
    
    intRow = mshPati.Row
    
    If glngSys Like "8??" Then
        lng病人ID = Val(mshPati.TextMatrix(intRow, GetColNum("客户ID")))
        If lng病人ID = 0 Then
            MsgBox "没有客户信息可以停用！", vbExclamation, gstrSysName: Exit Sub
        End If
    Else
        lng病人ID = Val(mshPati.TextMatrix(intRow, GetColNum("病人ID")))
        If lng病人ID = 0 Then
            MsgBox "没有病人信息可以停用！", vbExclamation, gstrSysName: Exit Sub
        End If
    End If
    int就诊次数 = GetColNum("就诊次数")
    If int就诊次数 <> -1 Then
        int就诊次数 = Val(mshPati.TextMatrix(intRow, int就诊次数))
        If int就诊次数 > 0 Then
            MsgBox """" & mshPati.TextMatrix(intRow, GetColNum("姓名")) & """已经在院就诊过 " & int就诊次数 & " 次，不允许停用。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If MsgBox("确实要停用""" & mshPati.TextMatrix(intRow, GetColNum("姓名")) & """的信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    strSQL = "zl_病人信息_Stop(" & lng病人ID & ")"
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
'    Call SQLTest(App.ProductName, Me.Caption, strSQL)
'    gcnOracle.Execute strSQL, , adCmdStoredProc
'    Call SQLTest
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    '行直接处理
    mshPati.TextMatrix(intRow, GetColNum("停用时间")) = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If Not mnuViewStop.Checked Then
        If mshPati.Rows > 2 Then
            mshPati.RemoveItem intRow
        Else
            With mshPati
                For i = 0 To .Cols - 1
                    .TextMatrix(intRow, i) = ""
                Next
            End With
        End If
        
        If intRow <= mshPati.Rows - 1 Then
            mshPati.Row = intRow
        Else
            mshPati.Row = mshPati.Rows - 1
        End If
    Else
        mshPati.Redraw = False
        For i = 0 To mshPati.Cols - 1
            mshPati.Col = i
            mshPati.CellForeColor = &HC0&
        Next
        mshPati.Redraw = True
    End If
    mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    Call mshPati_EnterCell
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_ToInPati_Click()
    Dim lng病人ID As Long, lng主页ID As Long
    Dim str住院号 As String, str姓名 As String
    Dim strSQL As String, strNote As String
    Dim rsTemp As New ADODB.Recordset
        
    lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID"))) '药店系统不用在院病人
    lng主页ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("就诊次数")))
    str住院号 = mshPati.TextMatrix(mshPati.Row, GetColNum("住院号"))
    
    If lng病人ID = 0 Then
        MsgBox "没有病人可以转为住院病人。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    strSQL = "Select Nvl(状态,0) 状态 From 病案主页 Where 病人ID=[1] And 主页ID=[2] And 病人性质=2"
    On Error GoTo errH
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    
    If rsTemp!状态 = 1 Then
        MsgBox "病人当前尚未入科,不能转为住院病人。请先将病人入科后再试。", vbInformation, gstrSysName
        Exit Sub
    ElseIf rsTemp!状态 = 2 Then
        MsgBox "病人当前正在转科,不能转为住院病人。请先将病人转科或取消转科后再试。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("确实要将该住院留观病人转为住院病人吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '没有住院号则分配一个
    If str住院号 = "" Then
        str住院号 = zlDatabase.GetNextNo(2)
        strNote = "在留观病人 " & str姓名 & " 转为住院病人之前，请先为该病人确定一个住院号。"
        If Not frmInput.InputVal(Me, "住院号", strNote, str住院号, 1, 10, False) Then Exit Sub
    End If
    
    
    strSQL = "ZL_病人变动记录_转住院(" & lng病人ID & "," & lng主页ID & "," & str住院号 & ")"
'    Call SQLTest(App.ProductName, Me.Caption, strSQL)
'    gcnOracle.Execute strSQL, , adCmdStoredProc
'    Call SQLTest
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Call mnuViewReFlash_Click
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditBlackList_Click()
    frmBlackList.mstrPrivs = mstrPrivs
    frmBlackList.Show 1, Me
End Sub

Private Sub mnuEditDelCard_Click()
    Dim strSQL As String, lng病人ID As Long
    
    lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")))
    
    If CheckBindCard(lng病人ID) = False Then
        '刘兴洪:24537
        MsgBox "该病人的卡号不是绑定卡,请到医疗卡发放管理界面进行退卡操作!", vbInformation, gstrSysName
        Exit Sub
    Else
        If MsgBox("你确定要取消当前病人的卡号绑定吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
     'Zl_医疗卡变动_Insert
       strSQL = "Zl_医疗卡变动_Insert("
      '      变动类型_In   Number,
      '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
      strSQL = strSQL & "" & 14 & ","
      '      病人id_In     住院费用记录.病人id%Type,
      strSQL = strSQL & "" & lng病人ID & ","
      '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
      strSQL = strSQL & "" & mlngCardType & ","
      '      原卡号_In     病人医疗卡信息.卡号%Type,
      strSQL = strSQL & "NULL,"
      '      医疗卡号_In   病人医疗卡信息.卡号%Type,
      strSQL = strSQL & "'" & mshPati.TextMatrix(mshPati.Row, GetColNum("就诊卡号")) & "',"
      '      变动原因_In   病人医疗卡变动.变动原因%Type,
      strSQL = strSQL & "'取消卡号绑定',"
      '      密码_In       病人信息.卡验证码%Type,
      strSQL = strSQL & "NULL,"
      '      操作员姓名_In 住院费用记录.操作员姓名%Type,
      strSQL = strSQL & "NULL,"
      '      变动时间_In   住院费用记录.登记时间%Type,
      'strSQL = strSQL & "to_date('" & Format(curDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
      '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
      strSQL = strSQL & "NULL,"
      '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
      strSQL = strSQL & "NULL)"
    
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    mshPati.TextMatrix(mshPati.Row, GetColNum("就诊卡")) = ""
    mshPati.TextMatrix(mshPati.Row, GetColNum("就诊卡号")) = ""
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckBindCard(ByVal lng病人ID As Long) As Boolean
'功能：检查病人是否有就诊卡记录
'问题号:52133
    Dim rsTmp As ADODB.Recordset, strSQL As String
    'by lesfeng 2009-12-30 大表拆分  病人费用记录 --〉住院费用记录 这里只对住院 记录性质 = 5属于就诊卡费用,而就诊卡保存在住院费用记录中
    strSQL = "Select Count(*) As 是否存在 From 病人医疗卡变动 Where 病人ID=[1] And 卡类别ID=[2] And 卡号=[3] And 变动类别=11"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, mlngCardType, Trim(mshPati.TextMatrix(mshPati.Row, GetColNum("就诊卡号"))))
    If rsTmp Is Nothing Then CheckBindCard = False: Exit Function
    If rsTmp.RecordCount = 0 Then CheckBindCard = False: Exit Function
    CheckBindCard = rsTmp!是否存在 > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mnuEditMzReCalc_Click()
    Dim lng病人ID As Long
    Dim str姓名 As String, strSQL As String
    
    '按费别重算门诊记帐费用
    '问题:41034
    If mshPati.Row <= 0 Then Exit Sub
    lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")))
    If lng病人ID = 0 Then
        MsgBox "请选择需要费用重算的病人！", vbExclamation, gstrSysName: Exit Sub
    End If
    str姓名 = mshPati.TextMatrix(mshPati.Row, GetColNum("姓名"))
    If MsgBox("你确定要将[" & str姓名 & "]的未结的门诊记帐费用按当前费别重算吗?" & vbCrLf & vbCrLf & _
        "本操作将按病人当前费别对应的优惠比率对未结费用重新进行打折计算!", vbInformation + vbYesNo + vbDefaultButton1, App.ProductName) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo errH
    strSQL = "Zl_病人未结门诊费用_Recalc(" & lng病人ID & ")"
    zlDatabase.ExecuteProcedure strSQL, App.ProductName
    MsgBox "费用重算成功!", vbOKOnly + vbInformation, gstrSysName
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub mnuEditPatiInfo_Click()
    Dim lng病人ID As Long, lng就诊ID As Long
    Dim strInfo As String
    Dim blnOK As Boolean
    
    If CreatePublicPatient = False Then Exit Sub
    '65802:刘鹏飞
    If glngSys Like "8??" Then
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("客户ID")))
    Else
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")))
    End If
    If lng病人ID <> 0 Then
        Select Case TabPatiState.SelectedItem.Key
            Case "T_在院病人", "T_出院病人", "T_留观病人"
                lng就诊ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("主页ID")))
            Case Else
                lng就诊ID = 0
        End Select
    Else
        lng就诊ID = 0
    End If
    '病人信息调整
    blnOK = gobjPublicPatient.ModiPatiBaseInfo(Me, "病人信息管理", lng病人ID, lng就诊ID, 2)
    If blnOK = True And lng病人ID <> 0 Then Call mnuViewReFlash_Click
End Sub

Private Sub mnuFileLocalSet_Click()
    Call frmLocalSet.zlSetPara(Me, mstrPrivs, mlngModul)
End Sub

Private Sub mnuFileInsure_Click()
    gclsInsure.InsureSupport
End Sub

Private Sub mnuFilePrintMed_Click()
    Dim lng病人ID As Long
    
    If glngSys Like "8??" Then
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("客户ID")))
    Else
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")))
    End If
    If lng病人ID = 0 Then Exit Sub
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1101", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1101", Me, "病人ID=" & lng病人ID, 2)
    End If
End Sub

Private Sub mnuFileRollingCurtain_Click()
    Call zlExecuteChargeRollingCurtain(Me)
End Sub

Private Sub mnuQuery_ChangeLog_Click()
    Dim lng病人ID As Long
    
    If glngSys Like "8??" Then
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("客户ID")))
    Else
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")))
    End If
    
    Call frmPatiInfoChangeLog.ShowMe(Me, mstrPrivs, lng病人ID)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim str病人ID As String
    
    str病人ID = mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID"))
    If str病人ID <> "" Then
        With mshPati
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "病人ID=" & str病人ID)
        End With
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
    End If
End Sub

Private Sub mnuViewByDept_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewByDept.Count - 1
        mnuViewByDept(i).Checked = (i = Index)
    Next
    Call InitUnits
    Call ShowPatis(mstrFilter, , gblnMyStyle, mstrFilterInfo)
End Sub

Private Sub mnuViewFilter_Click()
    frmPatiFilter.mbytType = Val(mshPati.Tag)
    frmPatiFilter.Show 1, Me
    If gblnOK Then
        With frmPatiFilter
            mstrFilter = .mstrFilter
            mstrFilterInfo = .mstrFilterInfo
            SQLCondition.登记时间B = .dtp登记B
            SQLCondition.登记时间E = .dtp登记E
            SQLCondition.出生时间B = .dtp出生B
            SQLCondition.出生时间E = .dtp出生E
            
            SQLCondition.入院时间B = .dtp入院B
            SQLCondition.入院时间E = .dtp入院E
            SQLCondition.出院时间B = .dtp出院B
            SQLCondition.出院时间E = .dtp出院E
            
            SQLCondition.住院号 = Trim(.txt住院号.Text)
            SQLCondition.性别 = zlCommFun.GetNeedName(.cbo性别.Text)
            SQLCondition.费别 = zlCommFun.GetNeedName(.cbo费别.Text)
            SQLCondition.区域 = zlCommFun.GetNeedName(.txt区域.Text)
            SQLCondition.医疗付款方式 = zlCommFun.GetNeedName(.cboPayPlan.Text)
            
            '59340:刘鹏飞,2013-04-23,姓名匹配添加gstrLike
            If .PatiIdentify.GetCurCard.名称 = "姓名" And .mlngPatiId = 0 And (.chk登记.Value = 1 Or .chk入院.Value = 1 Or .chk出院.Value = 1) Then     '姓名
                SQLCondition.Patient = gstrLike & Trim(.PatiIdentify.Text) & "%"
            Else
                SQLCondition.Patient = IIf(.mlngPatiId <> 0, .mlngPatiId, .PatiIdentify.Text)
            End If
        End With
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuViewGo_Click()
    frmPatiFind.mbytType = Val(mshPati.Tag)
    frmPatiFind.Show 1, Me
    If gblnOK Then Call SeekPati(frmPatiFind.optHead)
End Sub

Private Sub mnuViewStop_Click()
    mnuViewStop.Checked = Not mnuViewStop.Checked
    Call ShowPatis(mstrFilter, , gblnMyStyle, mstrFilterInfo)
End Sub

Private Sub mnuEdit_Surety_Click()
    Dim lng病人ID As Long, lngRow As Long
    Dim bln在院病人 As Boolean
    
    lngRow = mshPati.Row
    If glngSys Like "8??" Then
        lng病人ID = Val(mshPati.TextMatrix(lngRow, GetColNum("客户ID")))
    Else
        lng病人ID = Val(mshPati.TextMatrix(lngRow, GetColNum("病人ID")))
    End If
    
    If GetColNum("科室") <> -1 Then
        bln在院病人 = Trim(mshPati.TextMatrix(lngRow, GetColNum("科室"))) <> ""
    End If
    
    If lng病人ID <> 0 Then
        frmSurety.mlng病人ID = lng病人ID
        frmSurety.mbln在院病人 = bln在院病人
        frmSurety.mstrPrivs = mstrPrivs
        frmSurety.Show 1, Me
    End If
End Sub

Private Sub mnuViewToolDist_Click()
    mnuViewToolDist.Checked = Not mnuViewToolDist.Checked
    tvwDist_s.Visible = mnuViewToolDist.Checked
    pic.Visible = tvwDist_s.Visible
    Call Form_Resize
    Me.Refresh
End Sub

Private Sub mshPati_DblClick()
    If glngSys Like "8??" Then
        If mshPati.MouseRow = 0 Or mshPati.TextMatrix(mshPati.MouseRow, GetColNum("客户ID")) = "" Then Exit Sub
    Else
        If mshPati.MouseRow = 0 Or mshPati.TextMatrix(mshPati.MouseRow, GetColNum("病人ID")) = "" Then Exit Sub
    End If
    mnuEdit_View_Click
End Sub

Private Sub mshPati_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuEdit, 2
    ElseIf Button = 1 Then
        mblnDown = True
    End If
End Sub

Private Sub Form_Activate()
    If mblnUnLoad Then
        Unload Me
    Else
        Call InitLocPar(mlngModul)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            '始终从当前行开始
            If mnuViewGo.Enabled Then Call SeekPati(False)
        Case vbKeyReturn
            If mnuEdit_View.Enabled Then mnuEdit_View_Click
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer, Curdate As Date, lngTmp As Long
    Dim blnHavePrivs As Boolean
    mstrPrivs_RollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    mblnInitGrid = False
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    lngTmp = Val(zlDatabase.GetPara("显示病人方式", glngSys, mlngModul, 0))
    For i = 0 To mnuViewByDept.UBound
        mnuViewByDept(i).Checked = (i = lngTmp)
    Next
   
    '初始化站点列表
    
    
    '恢复个性病人清单类型
    mshPati.Tag = zlDatabase.GetPara("病人类型", glngSys, mlngModul, 1) 'mshPati.Tag中保存着真实病人类别：0-所有,1-在院,2-出院,3-门诊,4-留观
    
    TabPatiState.Tabs(Val(mshPati.Tag) + 1).Selected = True
    TabPatiState.Tag = TabPatiState.SelectedItem.Key
    mnuViewToolDist.Enabled = TabPatiState.SelectedItem.Key = "T_在院病人" Or TabPatiState.SelectedItem.Key = "T_出院病人"
    If mnuViewToolDist.Enabled Then InitUnits
    Call InitFace
    
    If glngSys Like "8??" Then
        Me.Caption = "客户信息管理"
        mnuEditBlackList.Visible = False
        mnuEdit_Merge.Caption = "客户合并(&G)"
        mnuEdit_Stop.Caption = "停用客户(&S)"
        For i = 1 To tbr.Buttons.Count
            tbr.Buttons(i).ToolTipText = Replace(tbr.Buttons(i).ToolTipText, "病人", "客户")
        Next
        
        mshPati.Tag = 3
        TabPatiState.Tabs.Remove 5
        TabPatiState.Tabs.Remove 3
        TabPatiState.Tabs.Remove 2
        TabPatiState.Tabs.Remove 1
    End If
    
    RestoreWinState Me, App.ProductName
    
    mblnUnLoad = False
    
    '权限设置
    If InStr(mstrPrivs, ";修改;") = 0 Then
        mnuEdit_Modi.Visible = False
        tbr.Buttons("Modi").Visible = False
    End If
    
    If InStr(mstrPrivs, ";增加;") = 0 Then
         mnuEdit_Add.Visible = False
         tbr.Buttons("Add").Visible = False
    End If
    
    If InStr(mstrPrivs, ";删除;") = 0 Then
        mnuEdit_Del.Visible = False
        tbr.Buttons("Del").Visible = False
    End If
    
    If InStr(mstrPrivs, ";启停;") = 0 Then
        mnuEdit_Stop.Visible = False
        mnuEdit_Restore.Visible = False
        mnuEdit_Restore_.Visible = False
    End If
    
    If InStr(mstrPrivs, ";增加;") = 0 And InStr(mstrPrivs, ";修改;") = 0 And InStr(mstrPrivs, ";删除;") = 0 Then
        mnuEdit_.Visible = False
        tbr.Buttons("Edit_").Visible = False
    End If
    
    If Not (InStr(mstrPrivs, ";增加;") = 0 And InStr(mstrPrivs, ";修改;") = 0 And InStr(mstrPrivs, ";删除;") = 0 And InStr(mstrPrivs, ";启停;") = 0) Then
        If gstr磁卡ID <> "" Then
            Call UpdateShareID(mlngModul, gstr磁卡ID, 5)
        End If
    End If
    
    If InStr(mstrPrivs, "身份合并") = 0 Then
        mnuEdit_Merge.Visible = False
        mnuEdit_Merge_.Visible = False
        tbr.Buttons("Merge").Visible = False
        tbr.Buttons("Merge_").Visible = False
    End If
    
    If InStr(mstrPrivs, "住院留观转住院") = 0 Then
        mnuEdit_ToInPati.Visible = False
    End If
    If InStr(mstrPrivs, "取消卡号绑定") = 0 Then
        mnuEditDelCard.Visible = False
    End If
    '收费轧帐管理
    blnHavePrivs = InStr(mstrPrivs_RollingCurtain, ";轧帐;") > 0
    mnuFileRollingCurtain.Visible = blnHavePrivs
    mnuFileRollingCurtainSplit.Visible = blnHavePrivs
    tbr.Buttons("轧帐").Visible = blnHavePrivs
    tbr.Buttons("SplitRollingCurtain").Visible = blnHavePrivs
    
    '问题:41034
    mnuEditMzReCalc.Visible = InStr(1, mstrPrivs, ";重算门诊费用;") > 0
    mnuEditSplit.Visible = InStr(1, mstrPrivs, ";重算门诊费用;") > 0
    
    mstrUserUnitIDs = GetUserUnits '所在科室所属病区+本身所在病区
    '进入时缺省不显示内容
    Call SetHeader(gblnMyStyle)
    Call mshPati_EnterCell
    
    '在院或出院病人才显示科室分布
    If TabPatiState.SelectedItem.Key = "T_在院病人" Or TabPatiState.SelectedItem.Key = "T_出院病人" Then
        tvwDist_s.Visible = mnuViewToolDist.Enabled
    Else
        tvwDist_s.Visible = False
    End If
    
    With frmPatiFilter
        .mbytType = Val(mshPati.Tag)
        Call .MakeFilter
        
        mstrFilter = .mstrFilter
        mstrFilterInfo = .mstrFilterInfo
        SQLCondition.登记时间B = .dtp登记B
        SQLCondition.登记时间E = .dtp登记E
        SQLCondition.出生时间B = .dtp出生B
        SQLCondition.出生时间E = .dtp出生E
        
        SQLCondition.入院时间B = .dtp入院B
        SQLCondition.入院时间E = .dtp入院E
        SQLCondition.出院时间B = .dtp出院B
        SQLCondition.出院时间E = .dtp出院E
        
        SQLCondition.住院号 = Trim(.txt住院号.Text)
        SQLCondition.性别 = zlCommFun.GetNeedName(.cbo性别.Text)
        SQLCondition.费别 = zlCommFun.GetNeedName(.cbo费别.Text)
        SQLCondition.区域 = zlCommFun.GetNeedName(.txt区域.Text)
        '59340:刘鹏飞,2013-04-23,姓名匹配添加gstrLike
        If .PatiIdentify.GetCurCard.名称 = "姓名" And .mlngPatiId = 0 And (.chk登记.Value = 1 Or .chk入院.Value = 1 Or .chk出院.Value = 1) Then     '姓名
            SQLCondition.Patient = gstrLike & Trim(.PatiIdentify.Text) & "%"
        Else
            SQLCondition.Patient = IIf(.mlngPatiId <> 0, .mlngPatiId, .PatiIdentify.Text)
        End If
    End With
    
    '初始化病人信息公共部件
    Call CreatePublicPatient
    '扩展功能
    Call LoadPlugInMnu

End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '工具条占用高度
    Dim staH As Long '状态栏占用高度
    Dim DisW As Long '病人分布表宽度
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    mshPati.MousePointer = 0
    
    mshPati.Redraw = False
    
    If mblnMax Then
        tvwDist_s.Width = 2500
        mblnMax = False
    End If
    If Me.WindowState = 2 Then mblnMax = True
    
    '靠齐控件宽度和高度
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    DisW = IIf(tvwDist_s.Visible, tvwDist_s.Width + pic.Width, 0)
    
    pic.Visible = tvwDist_s.Visible
    
    With tvwDist_s
        .Left = Me.ScaleLeft
        .Top = Me.ScaleTop + cbrH + IIf(cboNodeList.Visible, cboNodeList.Height + 100, 0)
        .Height = Me.ScaleHeight - staH - cbrH - IIf(cboNodeList.Visible, cboNodeList.Height + 100, 0)
    End With
    With pic
        .Left = tvwDist_s.Left + tvwDist_s.Width
        '.Top = tvwDist_s.Top
        .Top = IIf(cboNodeList.Visible, Me.ScaleTop + cbrH, tvwDist_s.Top)
        '.Height = tvwDist_s.Height
        .Height = IIf(cboNodeList.Visible, Me.ScaleHeight - staH - cbrH, tvwDist_s.Height)
    End With
    
    With TabPatiState
        .Left = DisW
        '.Top = tvwDist_s.Top
        .Top = IIf(cboNodeList.Visible, Me.ScaleTop + cbrH, tvwDist_s.Top)
        .Width = Me.ScaleWidth - DisW
        '.Height = tvwDist_s.Height
        .Height = IIf(cboNodeList.Visible, Me.ScaleHeight - staH - cbrH, tvwDist_s.Height)
    End With
    With mshPati
        .Left = TabPatiState.ClientLeft
        .Top = TabPatiState.ClientTop
        .Height = TabPatiState.ClientHeight
        .Width = TabPatiState.ClientWidth
    End With
    cboNodeList.Width = tvwDist_s.Width - 600
    mshPati.Redraw = True
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngTmp As Long, i As Long
    
    If Not gobjPublicPatient Is Nothing Then
        Set gobjPublicPatient = Nothing
    End If
    
    mstrFilter = ""
    mstrFilterInfo = ""
    zlDatabase.SetPara "病人类型", Val(mshPati.Tag), glngSys, mlngModul
    SaveWinState Me, App.ProductName
    
    '显示病人方式
    lngTmp = 0
    For i = 0 To mnuViewByDept.UBound
        If mnuViewByDept(i).Checked Then
            lngTmp = i
            Exit For
        End If
    Next
    zlDatabase.SetPara "显示病人方式", lngTmp, glngSys, mlngModul
    
    Unload frmPatiFind
    Unload frmPatiFilter
End Sub

Private Sub mnuEdit_Del_Click()
    Dim strSQL As String, intRow As Long, i As Long
    Dim strSQL1 As String
    Dim blnTrans As Boolean
    
    intRow = mshPati.Row
    
    If glngSys Like "8??" Then
        If Not IsNumeric(mshPati.TextMatrix(intRow, GetColNum("客户ID"))) Then
            MsgBox "没有客户信息可以删除！", vbExclamation, gstrSysName: Exit Sub
        End If
        If MsgBox("该操作将删除和客户""" & mshPati.TextMatrix(intRow, GetColNum("姓名")) & """相关的所有信息，并且不可恢复，要删除吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        strSQL1 = "Zl_病人照片_Delete(" & mshPati.TextMatrix(intRow, GetColNum("客户ID")) & ")"
        strSQL = "zl_病人信息_DELETE(" & mshPati.TextMatrix(intRow, GetColNum("客户ID")) & ")"
    Else
        If Not IsNumeric(mshPati.TextMatrix(intRow, GetColNum("病人ID"))) Then
            MsgBox "没有病人信息可以删除！", vbExclamation, gstrSysName: Exit Sub
        End If
        If MsgBox("该操作将删除和病人""" & mshPati.TextMatrix(intRow, GetColNum("姓名")) & """相关的所有信息，并且不可恢复，要删除吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        strSQL1 = "Zl_病人照片_Delete(" & mshPati.TextMatrix(intRow, GetColNum("病人ID")) & ")"
        strSQL = "zl_病人信息_DELETE(" & mshPati.TextMatrix(intRow, GetColNum("病人ID")) & ")"
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
'    Call SQLTest(App.ProductName, Me.Caption, strSQL)
'    gcnOracle.Execute strSQL, , adCmdStoredProc
'    Call SQLTest
    zlDatabase.ExecuteProcedure strSQL1, Me.Caption
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    '行直接处理
    If mshPati.Rows > 2 Then
        mshPati.RemoveItem intRow
    Else
        With mshPati
            For i = 0 To .Cols - 1
                .TextMatrix(intRow, i) = ""
            Next
        End With
    End If
    
    If intRow <= mshPati.Rows - 1 Then
        mshPati.Row = intRow
    Else
        mshPati.Row = mshPati.Rows - 1
    End If
    mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    Call mshPati_EnterCell
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Merge_Click()
    Dim lng病人ID As Long
    
    If glngSys Like "8??" Then
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("客户ID")))
        If lng病人ID = 0 Then
            MsgBox "没有客户信息可供合并！", vbExclamation, gstrSysName: Exit Sub
        End If
    Else
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")))
        If lng病人ID = 0 Then
            MsgBox "没有病人信息可供合并！", vbExclamation, gstrSysName: Exit Sub
        End If
    End If
    
    If ExistFeeInsurePatient(lng病人ID) Then
        MsgBox "该医保病人存在未结费用,请先结清后再合并！", vbExclamation, gstrSysName: Exit Sub
    End If
    
    On Error Resume Next
    
    frmMergePatient.mstrPrivs = mstrPrivs
    frmMergePatient.mlng病人ID = lng病人ID
    frmMergePatient.Show 1, Me
    
    If gblnOK Then
        If MsgBox("当前操作已更改清单内容,要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEdit_Modi_Click()
    If glngSys Like "8??" Then
        If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("客户ID"))) Then
            MsgBox "没有客户信息可以修改！", vbExclamation, gstrSysName: Exit Sub
        End If
    Else
        If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID"))) Then
            MsgBox "没有病人信息可以修改！", vbExclamation, gstrSysName: Exit Sub
        End If
    End If
    
    On Error Resume Next
    Err.Clear
    
    If glngSys Like "8??" Then
        frmPatient.mlng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("客户ID")))
    Else
        frmPatient.mlng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")))
    End If
    frmPatient.mlngModul = mlngModul
    frmPatient.mstrPrivs = mstrPrivs
    frmPatient.mbytInState = 1
    frmPatient.mbytView = Val(mshPati.Tag)
    frmPatient.Show 1, Me
    If gblnOK Then
        If MsgBox("当前操作已更改清单内容,要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEdit_Add_Click()
    On Error Resume Next
    Err.Clear
    
    frmPatient.mlngModul = mlngModul
    frmPatient.mstrPrivs = mstrPrivs
    frmPatient.mbytInState = 0
    frmPatient.mbytView = Val(mshPati.Tag)
    frmPatient.Show 1, Me
    If gblnOK Then
        If MsgBox("当前操作已更改清单内容,要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuEdit_View_Click()
    Dim lng病人ID  As Long
    Dim lng主页ID  As Long
    
    On Error Resume Next
    If glngSys Like "8??" Then
        If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("客户ID"))) Then
            MsgBox "没有客户信息可以查看！", vbExclamation, gstrSysName: Exit Sub
        End If
        frmPatient.mlngModul = mlngModul
        frmPatient.mstrPrivs = mstrPrivs
        frmPatient.mbytInState = 2
        frmPatient.mbytView = Val(mshPati.Tag)
        frmPatient.mlng病人ID = CLng(mshPati.TextMatrix(mshPati.Row, GetColNum("客户ID")))
        frmPatient.Show 1, Me
        mshPati.Refresh
    Else
        If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID"))) Then
            MsgBox "没有病人信息可以查看！", vbExclamation, gstrSysName: Exit Sub
        End If
        lng病人ID = CLng(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")))
        lng主页ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("主页ID"))) 'CLNG传空串 报错 13-类型不匹配
        If CreatePublicPatient Then
            Call gobjPublicPatient.ReadPatiDegreeCard(Me, lng病人ID, lng主页ID)
        End If
        
        mshPati.Refresh
    End If
End Sub

Private Sub mnuFile_Quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewReFlash_Click()
    Call ShowPatis(mstrFilter, , gblnMyStyle, mstrFilterInfo)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Visible = Not cbr.Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If pic.Left + X < 1000 Or TabPatiState.Width - X < 2000 Then Exit Sub
        pic.Left = pic.Left + X
        tvwDist_s.Width = tvwDist_s.Width + X
        TabPatiState.Left = TabPatiState.Left + X
        TabPatiState.Width = TabPatiState.Width - X
        mshPati.Left = TabPatiState.ClientLeft
        mshPati.Width = TabPatiState.ClientWidth
        cboNodeList.Width = tvwDist_s.Width - 600
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PatiColor" Then
        zlDatabase.ShowPatiColorTip Me
    End If
End Sub

Private Sub TabPatiState_Click()
    If TabPatiState.Tag = TabPatiState.SelectedItem.Key Then Exit Sub
    '35632:刘鹏飞,2013-07-29
    If mblnInitGrid = True Then SaveFlexState mshPati, App.ProductName & "\" & Me.Name
    mshPati.Tag = TabPatiState.SelectedItem.Index - 1 '存储病人类别：0-所有,1-在院,2-出院,3-门诊,4-留观
    cboNodeList.Visible = TabPatiState.SelectedItem.Index <> 1 And TabPatiState.SelectedItem.Index <> 4 And TabPatiState.SelectedItem.Index <> 5 And cboNodeList.ListCount > 0
    mnuViewToolDist.Enabled = TabPatiState.SelectedItem.Key = "T_在院病人" Or TabPatiState.SelectedItem.Key = "T_出院病人"
    mnuViewPatiMode.Enabled = mnuViewToolDist.Enabled
    If mnuViewToolDist.Enabled Then InitUnits
    
    Unload frmPatiFilter
    Unload frmPatiFind
    If TabPatiState.Tag <> "" Then '窗体加载时不查数据
        With frmPatiFilter
            .mbytType = Val(mshPati.Tag)
            Call .MakeFilter
        
        '切换病人类型时条件恢复为空(使用缺省条件)
            mstrFilter = .mstrFilter
            mstrFilterInfo = .mstrFilterInfo
            SQLCondition.登记时间B = .dtp登记B
            SQLCondition.登记时间E = .dtp登记E
            SQLCondition.出生时间B = .dtp出生B
            SQLCondition.出生时间E = .dtp出生E
            
            SQLCondition.入院时间B = .dtp入院B
            SQLCondition.入院时间E = .dtp入院E
            SQLCondition.出院时间B = .dtp出院B
            SQLCondition.出院时间E = .dtp出院E
            
            SQLCondition.住院号 = Trim(.txt住院号.Text)
            SQLCondition.性别 = zlCommFun.GetNeedName(.cbo性别.Text)
            SQLCondition.费别 = zlCommFun.GetNeedName(.cbo费别.Text)
            SQLCondition.区域 = zlCommFun.GetNeedName(.txt区域.Text)
            '59340:刘鹏飞,2013-04-23,姓名匹配添加gstrLike
            If .PatiIdentify.GetCurCard.名称 = "姓名" And .mlngPatiId = 0 And (.chk登记.Value = 1 Or .chk入院.Value = 1 Or .chk出院.Value = 1) Then      '姓名
                SQLCondition.Patient = gstrLike & Trim(.PatiIdentify.Text) & "%"
            Else
                SQLCondition.Patient = IIf(.mlngPatiId <> 0, .mlngPatiId, .PatiIdentify.Text)
            End If
        End With
        mshPati.Clear: mshPati.Rows = 2
        Call SetHeader(gblnMyStyle)
        Select Case TabPatiState.SelectedItem.Key
            Case "T_所有病人"  '所有病人
                tvwDist_s.Visible = False
            Case "T_在院病人" '在院病人
                tvwDist_s.Tag = tvwDist_s.SelectedItem.Key
                tvwDist_s.Visible = mnuViewToolDist.Enabled
            Case "T_出院病人"  '出院病人
                tvwDist_s.Tag = tvwDist_s.SelectedItem.Key
                tvwDist_s.Visible = mnuViewToolDist.Enabled
            Case "T_门诊病人" '门诊病人
                tvwDist_s.Visible = False
            Case "T_留观病人"    '留观病人
                tvwDist_s.Visible = False
        End Select
        Call Form_Resize
        
        Dim blnAutoRefresh As Boolean
        '54701:刘鹏飞,2012-10-19
        blnAutoRefresh = (Val(zlDatabase.GetPara("自动刷新数据", glngSys, mlngModul, 0)) = 1)
        If blnAutoRefresh = True Then
            Call ShowPatis(mstrFilter, , gblnMyStyle, mstrFilterInfo) '强行设置并恢复列宽
        Else
            tvwDist_s.Tag = ""
        End If
    End If
    TabPatiState.Tag = TabPatiState.SelectedItem.Key
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_Quit_Click
        Case "Go"
            mnuViewGo_Click
        Case "Filter"
            mnuViewFilter_Click
        Case "Modi"
            mnuEdit_Modi_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "Add"
            mnuEdit_Add_Click
        Case "Merge"
            mnuEdit_Merge_Click
        Case "View"
            mnuEdit_View_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "轧帐"
            mnuFileRollingCurtain_Click
        Case "Family"
            mnuEdit_FamilyAdd_Click
        Case "PlugIn"
            PopupMenu mnuEdit_PlugIn, vbPopupMenuRightButton
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "FamilyAdd"
       mnuEdit_FamilyAdd_Click
    Case "FamilyView"
       mnuEdit_FamilyView_Click
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub tvwDist_s_NodeClick(ByVal Node As MSComctlLib.Node)
    '相同点击不再处理
    If tvwDist_s.Tag = Node.Key Then Exit Sub
    tvwDist_s.Tag = Node.Key
    
    Call ShowPatis(mstrFilter, , gblnMyStyle, mstrFilterInfo)
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
    
    intRow = mshPati.Row
    
    '表头
    If glngSys Like "8??" Then
        objOut.Title.Text = "客户清单"
    Else
        objOut.Title.Text = "病人清单"
    End If
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    If Not glngSys Like "8??" Then
        Select Case TabPatiState.SelectedItem.Key
            Case "T_所有病人"
                objRow.Add "分类：所有病人"
            Case "T_在院病人"
                objRow.Add "分类：在院病人"
                objRow.Add "部门：" & tvwDist_s.SelectedItem.Text
            Case "T_出院病人"
                objRow.Add "分类：出院病人"
            Case "T_门诊病人"
                objRow.Add "分类：门诊病人"
            Case "T_留观病人"
                objRow.Add "分类：留观病人"
        End Select
        objOut.UnderAppRows.Add objRow
    End If
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    mshPati.Redraw = False
    Set objOut.Body = mshPati
    
    '输出
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mshPati.Row = intRow
    mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    mshPati.Redraw = True
End Sub

Private Sub mshPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mnuEdit_Del.Enabled Then mnuEdit_Del_Click
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Private Sub InitFace()
    Dim rsTmp As New ADODB.Recordset
    Dim objNode As Node, i As Integer
    Dim strPreKey  As String, strSQL As String, strUnitIDs As String
    Dim blnLimitIn As Boolean, blnByDept As Boolean, blnLimitUnit As Boolean
    
    On Error GoTo errHandle
    
    blnLimitIn = TabPatiState.SelectedItem.Key = "T_在院病人"
    blnByDept = mnuViewByDept(1).Checked
    blnLimitUnit = InStr(mstrPrivs, "所有病区") = 0
    If blnLimitUnit Then strUnitIDs = GetUserUnits
    
    '加载站点选项
    strSQL = "SELECT DISTINCT a.站点, c.名称" & vbNewLine & _
            " FROM 部门表 a, 部门性质说明 b, Zlnodelist c" & vbNewLine & _
            " WHERE a.Id = b.部门id AND a.站点 = c.编号 AND (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') OR a.撤档时间 IS NULL) AND" & vbNewLine & _
            "      b.工作性质 = [1] " & vbNewLine & _
            IIf(blnLimitIn, " And ID In (Select Distinct " & IIf(blnByDept, "科室id", "病区id") & " From 床位状况记录 Where 病人id Is Not Null)", "") & vbNewLine & _
            IIf(blnLimitUnit, " And A.ID In (" & strUnitIDs & ")", "") & _
            " ORDER BY a.站点"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(blnByDept, "临床", "护理"))
    cboNodeList.Clear
    If rsTmp.RecordCount > 0 Then
        While Not rsTmp.EOF
            cboNodeList.AddItem rsTmp!站点 & "-" & rsTmp!名称
            cboNodeList.ItemData(rsTmp.AbsolutePosition - 1) = rsTmp!站点
            rsTmp.MoveNext
        Wend
        Call cbo.Locate(cboNodeList, gstrNodeNo, True)
    Else
        lblNode.Visible = False
        cboNodeList.Visible = False
        Form_Resize
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitUnits() As Boolean
'功能：初始化病人病区科室分布列表
'说明：以病区-科室分层,所有病区、科室在当前在院病人之中获得
    Dim rsTmp As New ADODB.Recordset
    Dim objNode As Node, i As Integer
    Dim strPreKey  As String, strSQL As String, strUnitIDs As String
    Dim blnLimitIn As Boolean, blnByDept As Boolean, blnLimitUnit As Boolean
    Dim strNodeNo
    
    On Error GoTo errH
    
    blnLimitIn = TabPatiState.SelectedItem.Key = "T_在院病人"
    blnByDept = mnuViewByDept(1).Checked
    blnLimitUnit = InStr(mstrPrivs, "所有病区") = 0
    If blnLimitUnit Then strUnitIDs = GetUserUnits
    If cboNodeList.ListIndex <> -1 Then
        strNodeNo = Mid(cboNodeList.Text, 1, InStr(cboNodeList.Text, "-") - 1)
    Else
        strNodeNo = 0
    End If
             
    strPreKey = ""
    If Not tvwDist_s.SelectedItem Is Nothing Then strPreKey = tvwDist_s.SelectedItem.Key
    
    tvwDist_s.Nodes.Clear
    Set objNode = tvwDist_s.Nodes.Add(, , "Root", IIf(blnByDept, "所有科室", "所有病区"), 1)
    objNode.Expanded = True
    If objNode.Key = strPreKey Then objNode.Selected = True
    'by lesfeng 2010-03-08 性能优化
    strSQL = "Select A.ID, A.编码, A.名称" & vbNewLine & _
            "From 部门表 A, 部门性质说明 B" & vbNewLine & _
            "Where A.ID = B.部门id And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & vbNewLine & _
            "      B.工作性质 =[1] " & _
            IIf(blnLimitIn, " And ID In (Select Distinct " & IIf(blnByDept, "科室id", "病区id") & " From 床位状况记录 Where 病人id Is Not Null)", "") & vbNewLine & _
            IIf(blnLimitUnit, " And A.ID In (" & strUnitIDs & ")", "") & _
            " And (A.站点=[2] Or A.站点 is Null)" & _
            "Order By A.编码"
'    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(blnByDept, "临床", "护理"), strNodeNo)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            Set objNode = tvwDist_s.Nodes.Add("Root", 4, "D" & rsTmp!ID, "[" & rsTmp!编码 & "]" & rsTmp!名称, 1)
            
            If rsTmp!ID = UserInfo.部门ID Then objNode.Selected = True
            If objNode.Key = strPreKey Then objNode.Selected = True
            
            objNode.Expanded = True
            rsTmp.MoveNext
        Next
    End If
    If tvwDist_s.SelectedItem Is Nothing Then
        tvwDist_s.Nodes(IIf(tvwDist_s.Nodes.Count > 1, 2, 1)).Selected = True
    End If
    
    InitUnits = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowPatis(Optional ByVal strIF As String, Optional blnSort As Boolean, Optional blnSet As Boolean, Optional ByVal strIFInfo As String)
'功能：根据当前菜单浏览要求(自动生成条件),读取病人信息
'参数：strIF=" And ...."形式的过滤条件
    Dim strSQL As String, strInfo As String, strCard As String
    Dim i As Double, j As Double, lngFamily As Long, lngDeptID As Long
    Dim lngCol床号 As Long, lngCol停用 As Long, lngPreRow As Long
    Dim blnByDept As Boolean
    Dim str卡号SQL As String, strFileds As String
    Dim rsTemp As Recordset
    
    On Error GoTo errH
    
    If Not blnSort Then
        blnByDept = mnuViewByDept(1).Checked
                
        '第一次(每次切换)设置缺省条件
        If strIF = "" Then
            Select Case TabPatiState.SelectedItem.Key
                Case "T_所有病人", "T_门诊病人", "T_留观病人" '所有病人,门诊病人或留观病人(当天)
                    strIF = " And A.登记时间 Between trunc(Sysdate) And Sysdate"
                Case "T_在院病人" '在院病人
                    'strIF = " And P.入院日期 Between trunc(Sysdate) And Sysdate"
                Case "T_出院病人" '出院病人(当天)
                    strIF = " And P.出院日期 Between trunc(Sysdate) And Sysdate"
            End Select
        End If
        If strIFInfo = "" Then
            Select Case TabPatiState.SelectedItem.Key
                Case "T_所有病人", "T_门诊病人", "T_留观病人" '所有病人,门诊病人或留观病人(当天)
                    strIFInfo = " And A.登记时间 Between trunc(Sysdate) And Sysdate"
            End Select
        End If
        
        If Not mnuViewStop.Checked Then strIF = strIF & " And A.停用时间 is NULL"
        
        '就诊卡号显示
        '55849:刘鹏飞,2012-11-21,将原有Decode判断的方式改为固定提取字段,
        '因为Decode第一个变量使用常量从指标中提取字段数据，可能导致导致查不出结果，或者返回的记录集访问出现E-FAIL错误，估计是ADO和Oracle兼容性的Bug，在特定的Decode和子表查询同时使用时会出现，但没有明确的规律。
        'strCard = "Decode(" & IIf(gblnShowCard, 1, 0) & ",1,H.就诊卡号,LPAD('*',Length(H.就诊卡号),'*')) as 就诊卡,H.就诊卡号 as 就诊卡号,"
        If gblnShowCard = True Then
            strCard = "A.就诊卡号 as 就诊卡,A.就诊卡号 as 就诊卡号,"
        Else
            strCard = "LPAD('*',Length(A.就诊卡号),'*') as 就诊卡,A.就诊卡号 as 就诊卡号,"
        End If
        
        Select Case TabPatiState.SelectedItem.Key
            Case "T_在院病人", "T_出院病人" '在院病人或出院病人
                lngDeptID = Val(Mid(tvwDist_s.SelectedItem.Key, 2)) '所有病区或部门时,为root得0
                If lngDeptID <> 0 Then
                    If TabPatiState.SelectedItem.Key = "T_在院病人" Then
                        If blnByDept Then
                            strIF = strIF & " And R.科室ID=[1]"
                        Else
                            strIF = strIF & " And R.病区ID=[1]"
                        End If
                    Else
                        If blnByDept Then
                            strIF = strIF & " And P.出院科室ID=[1]"
                        Else
                            strIF = strIF & " And P.当前病区ID=[1]"
                        End If
                    End If
                ElseIf InStr(mstrPrivs, "所有病区") = 0 Then
                    If blnByDept Then
                        strIF = strIF & " And (A.当前科室id Is NULL Or A.当前科室id In(Select 科室ID From 病区科室对应 Where Instr(','||[2]||',',','||病区ID||',')>0))"
                    Else
                        strIF = strIF & " And (A.当前病区ID Is NULL Or Instr(','||[2]||',',','||A.当前病区ID||',')>0)"
                    End If
                End If
            Case "T_所有病人"
                If InStr(mstrPrivs, "所有病区") = 0 Then
                    If blnByDept Then
                        strIF = strIF & " And (A.当前科室id Is NULL Or A.当前科室id In(Select 科室ID From 病区科室对应 Where Instr(','||[2]||',',','||病区ID||',')>0))"
                    Else
                        strIF = strIF & " And (A.当前病区ID Is NULL Or Instr(','||[2]||',',','||A.当前病区ID||',')>0)"
                    End If
                End If
        End Select
        
        '问题号:51223
        
        '问题号:52133
        '获取缺省医疗卡类别
        mlngCardType = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModul, , , True))
        '问题号:53807
        If mlngCardType = 0 Then '当没有设置缺省发卡类型时,缺省缺就诊卡
            strSQL = "Select ID From 医疗卡类别 A Where A.名称='就诊卡'"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取医疗卡类别ID")
            If rsTemp.EOF = False Then mlngCardType = rsTemp!ID
            Set rsTemp = Nothing
        End If
        strSQL = "" & _
        "   Select 是否自制 From 医疗卡类别 Where ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取医疗卡类别", mlngCardType)
        If rsTemp Is Nothing Then mbln是否取消绑定 = False
        If rsTemp.RecordCount = 0 Then
            mbln是否取消绑定 = False
        Else
            mbln是否取消绑定 = rsTemp!是否自制 = 0
        End If
        '病人卡号
        str卡号SQL = "(Select f_List2str(Cast(COLLECT(G.卡号) as t_Strlist))" & _
            " From 病人医疗卡信息 G, 医疗卡类别 H" & _
            " Where G.病人ID = A.病人ID And G.卡类别ID = H.ID and G.状态 = 0 And H.ID=[16]) 就诊卡号,"
         
        '住院留观转住院
        If TabPatiState.SelectedItem.Key = "T_在院病人" Or TabPatiState.SelectedItem.Key = "T_留观病人" Then '在院病人
            mnuEdit_ToInPati.Visible = InStr(mstrPrivs, "住院留观转住院") > 0
        Else
            mnuEdit_ToInPati.Visible = False
        End If
        mnuViewStop.Enabled = TabPatiState.SelectedItem.Key = "T_所有病人" Or TabPatiState.SelectedItem.Key = "T_门诊病人"
        '问题号:521333
        Select Case TabPatiState.SelectedItem.Key
            Case "T_所有病人"  '所有病人
                strFileds = "A.病人ID,A.门诊号,A.住院号," & strCard & "A.姓名,A.性别,A.年龄,A.门诊费别,A.医疗付款方式," & _
                    " A.医保号,A.险类,A.病区,A.科室,A.床号,A.入院时间,A.住院目的,A.出院时间,A.住院次数,A.出生日期," & _
                    " A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份,A.身份证号,A.手机号,A.家庭地址,A.工作单位,A.登记时间, A.登记人," & _
                    " A.停用时间,A.病人性质,A.病人类型,A.主页ID,A.就诊次数"
                    
                strSQL = "Select A.病人ID,A.门诊号,A.住院号," & str卡号SQL & "A.姓名,A.性别,A.年龄,A.费别 as 门诊费别,Nvl(P.医疗付款方式,A.医疗付款方式) 医疗付款方式," & _
                    " Nvl(A.医保号,E.信息值) as 医保号,X.名称 as 险类," & _
                    " B.名称 as 病区,C.名称 as 科室,A.当前床号 as 床号,To_Char(A.入院时间,'YYYY-MM-DD') as 入院时间,P.住院目的," & _
                    " To_Char(A.出院时间,'YYYY-MM-DD') as 出院时间,A.住院次数,To_Char(A.出生日期,'YYYY-MM-DD HH24:MI') as 出生日期," & _
                    " A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份,A.身份证号,A.手机号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间, p.登记人," & _
                    " To_Char(A.停用时间,'YYYY-MM-DD') as 停用时间,0 as 病人性质,Nvl(P.病人类型,Decode(P.险类,Null,'普通病人','医保病人')) 病人类型,P.主页ID,A.主页ID 就诊次数" & _
                    " From 病案主页 P,病人信息 A,病案主页从表 E,部门表 B,部门表 C,保险类别 X" & _
                    " Where A.当前病区ID=B.ID(+) And A.当前科室ID=C.ID(+) And A.险类=X.序号(+)" & _
                    " And A.病人ID=P.病人ID(+) And A.主页ID=P.主页ID(+) " & strIF & _
                    " And A.病人ID=E.病人ID(+) And Nvl(A.主页ID,0)=E.主页ID(+) And E.信息名(+)='医保号'" & _
                    " Order by A.登记时间 Desc"
                strSQL = "Select " & strFileds & " From (" & strSQL & ") A"
                strInfo = "正在读取所有病人清单,请稍候 ..."
                tvwDist_s.Visible = False
            Case "T_在院病人" '在院病人
                strFileds = "A.病人ID,A.住院号," & strCard & "A.姓名,A.性别,A.年龄,A.住院费别,A.医疗付款方式," & _
                    " A.医保号,A.险类,A.病区,A.科室,A.床号,A.入院时间,A.住院目的," & _
                    " A.住院次数,A.出生日期,A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份," & _
                    " A.身份证号,A.手机号,A.家庭地址,A.工作单位,A.登记时间, A.登记人, " & _
                    " A.停用时间,A.病人性质,A.病人类型,A.主页ID,A.就诊次数"
                '58842,刘鹏飞,2013-02-25,在院病人读取(从在院病人中读取)
                strSQL = "Select A.病人ID,A.住院号," & str卡号SQL & "NVL(P.姓名,A.姓名) 姓名,NVL(P.性别,A.性别) 性别,NVL(P.年龄,A.年龄) 年龄,P.费别 as 住院费别,P.医疗付款方式," & _
                    " E.信息值 as 医保号,X.名称 as 险类," & _
                    " B.名称 as 病区,C.名称 as 科室,Decode(P.状态,1,P.出院病床,Nvl(P.出院病床,'家庭')) as 床号,To_Char(A.入院时间,'YYYY-MM-DD') as 入院时间,P.住院目的," & _
                    " A.住院次数,To_Char(A.出生日期,'YYYY-MM-DD HH24:MI') as 出生日期,A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份," & _
                    " A.身份证号,A.手机号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间, p.登记人, " & _
                    " To_Char(A.停用时间,'YYYY-MM-DD') as 停用时间,Nvl(P.病人性质,0) as 病人性质,Nvl(P.病人类型,Decode(P.险类,Null,'普通病人','医保病人')) 病人类型,P.主页ID,A.主页ID 就诊次数" & _
                    " From 病案主页 P,病人信息 A,病案主页从表 E,部门表 B,部门表 C,保险类别 X,在院病人 R" & _
                    " Where P.当前病区ID=B.ID(+) And P.出院科室ID=C.ID And A.险类=X.序号(+)" & _
                    " And A.病人ID=P.病人ID And A.主页ID=P.主页ID And Nvl(P.主页ID,0)<>0 " & strIF & _
                    " And E.信息名(+)='医保号' And P.病人ID=E.病人ID(+) And P.主页ID=E.主页ID(+) And R.病人ID=A.病人ID" & _
                    " Order by A.入院时间 Desc,A.住院号 Desc"
                strSQL = "Select " & strFileds & " From (" & strSQL & ") A"
                strInfo = "正在读取在院病人清单,请稍候 ..."
                tvwDist_s.Tag = tvwDist_s.SelectedItem.Key
                tvwDist_s.Visible = mnuViewToolDist.Enabled
            Case "T_出院病人"  '出院病人
                '问题28813 by lesfeng 2010-04-07 A.住院号 A.入院时间 A.出院时间 A.住院次数
                strFileds = "A.病人ID,A.住院号," & strCard & "A.姓名,A.性别,A.年龄,A.住院费别,A.医疗付款方式," & _
                    " A.医保号,A.险类,A.入院时间,A.住院目的,A.出院时间," & _
                    " A.住院次数,A.出生日期,A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份," & _
                    " A.身份证号,A.手机号,A.家庭地址,A.工作单位,A.登记时间, A.登记人, " & _
                    " A.停用时间,A.病人性质,A.病人类型,A.主页ID,A.就诊次数"
                    
                strSQL = "Select A.病人ID,P.住院号," & str卡号SQL & "NVL(P.姓名,A.姓名) 姓名,NVL(P.性别,A.性别) 性别,NVL(P.年龄,A.年龄) 年龄,P.费别 as 住院费别,P.医疗付款方式," & _
                    " E.信息值 as 医保号,X.名称 as 险类," & _
                    " To_Char(P.入院日期,'YYYY-MM-DD') as 入院时间,P.住院目的,To_Char(P.出院日期,'YYYY-MM-DD') as 出院时间," & _
                    " P.主页ID as 住院次数,To_Char(A.出生日期,'YYYY-MM-DD HH24:MI') as 出生日期,A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份," & _
                    " A.身份证号,A.手机号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间, p.登记人, " & _
                    " To_Char(A.停用时间,'YYYY-MM-DD') as 停用时间,Nvl(P.病人性质,0) as 病人性质,Nvl(P.病人类型,Decode(P.险类,Null,'普通病人','医保病人')) 病人类型,P.主页ID,A.主页ID 就诊次数" & _
                    " From 病案主页 P,病人信息 A,病案主页从表 E,保险类别 X" & _
                    " Where A.病人ID=P.病人ID And Nvl(P.主页ID,0)<>0" & _
                    " And P.出院日期 is Not NULL And A.险类=X.序号(+)" & strIF & _
                    " And P.病人ID=E.病人ID(+) And NVL(P.主页ID,0)=E.主页ID(+) And E.信息名(+)='医保号'" & _
                    " Order by A.出院时间 Desc,A.住院号"
                strSQL = "Select " & strFileds & " From (" & strSQL & ") A"
                strInfo = "正在读取出院病人清单,请稍候 ..."
                tvwDist_s.Tag = tvwDist_s.SelectedItem.Key
                tvwDist_s.Visible = mnuViewToolDist.Enabled
            Case "T_门诊病人" '门诊病人
                strFileds = "A.病人ID,A.门诊号," & strCard & "A.姓名,A.性别,A.年龄," & _
                     IIf(glngSys Like "8??", "A.会员等级", "A.门诊费别") & ",A.医疗付款方式," & _
                    " A.医保号,A.险类,A.出生日期,A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份," & _
                    " A.身份证号,A.手机号,A.家庭地址,A.工作单位,A.登记时间,A.停用时间,A.病人性质,A.病人类型," & _
                    " A.住院次数,A.主页ID,A.就诊次数"
                    
                strSQL = "Select A.病人ID,A.门诊号," & str卡号SQL & "A.姓名,A.性别,A.年龄," & _
                    " A.费别 as " & IIf(glngSys Like "8??", "会员等级", "门诊费别") & ",A.医疗付款方式," & _
                    " A.医保号,X.名称 as 险类," & _
                    " To_Char(A.出生日期,'YYYY-MM-DD HH24:MI') as 出生日期,A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份," & _
                    " A.身份证号,A.手机号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间," & _
                    " To_Char(A.停用时间,'YYYY-MM-DD') as 停用时间,0 as 病人性质,Decode(A.险类,Null,'普通病人','医保病人') 病人类型," & _
                    " NULL 住院次数,NULL 主页ID,NULL 就诊次数" & _
                    " From 病人信息 A,保险类别 X" & _
                    " Where A.当前病区ID is NULL And A.当前科室ID Is NULL" & _
                    " And A.主页ID IS NULL And A.险类=X.序号(+)" & strIF & _
                    " Order by A.登记时间 ,A.门诊号 Desc"
                strSQL = "Select " & strFileds & " From (" & strSQL & ") A"
                If glngSys Like "8??" Then
                    strInfo = "正在读取客户清单,请稍候 ..."
                Else
                    strInfo = "正在读取门诊病人清单,请稍候 ..."
                End If
                tvwDist_s.Visible = False
            Case "T_留观病人"    '留观病人
                strFileds = "A.病人ID,A.性质,A.门诊号,A.住院号,A.住院次数," & strCard & "A.姓名,A.性别,A.年龄," & _
                     IIf(glngSys Like "8??", "A.会员等级", "A.门诊费别") & ",A.医疗付款方式," & _
                    " A.医保号,A.险类,A.出生日期,A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份," & _
                    " A.身份证号,A.手机号,A.家庭地址,A.工作单位,A.登记时间, A.登记人, " & _
                    " A.停用时间,A.病人性质,A.病人类型,A.主页ID,A.就诊次数"
                    
                strSQL = "Select Distinct A.病人ID,Decode(P.病人性质,1,'门诊留观','住院留观') as 性质,A.门诊号,A.住院号,NULL as 住院次数," & str卡号SQL & "NVL(P.姓名,A.姓名) 姓名,NVL(P.性别,A.性别) 性别,NVL(P.年龄,A.年龄) 年龄," & _
                    " A.费别 as " & IIf(glngSys Like "8??", "会员等级", "门诊费别") & ",A.医疗付款方式," & _
                    " A.医保号,X.名称 as 险类," & _
                    " To_Char(A.出生日期,'YYYY-MM-DD HH24:MI') as 出生日期,A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份," & _
                    " A.身份证号,A.手机号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间, p.登记人, " & _
                    " To_Char(A.停用时间,'YYYY-MM-DD') as 停用时间,Nvl(P.病人性质,0) as 病人性质,Nvl(P.病人类型,Decode(P.险类,Null,'普通病人','医保病人')) 病人类型,P.主页ID,Decode(P.病人性质,2,A.主页ID,NULL) 就诊次数" & _
                    " From 病案主页 P,病人信息 A,保险类别 X" & _
                    " Where A.病人ID=P.病人ID And A.主页ID=P.主页ID And P.病人性质<>0 And P.住院号 Is Null" & _
                    " And A.险类=X.序号(+)" & strIF & _
                    " Order by 性质,登记时间 Desc"
                strSQL = "Select " & strFileds & " From (" & strSQL & ") A"
                strInfo = "正在读取留观病人清单,请稍候 ..."
                tvwDist_s.Visible = False
        End Select
        
        Call Form_Resize
        
        Call zlCommFun.ShowFlash(strInfo, Me)
        DoEvents
        Me.Refresh
        
        With SQLCondition
            Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngDeptID, mstrUserUnitIDs, .登记时间B, .登记时间E, .出生时间B, .出生时间E, _
                .入院时间B, .入院时间E, .出院时间B, .出院时间E, .住院号, .性别, .区域, .费别, .Patient, mlngCardType, .医疗付款方式)
        End With
    End If
    
    '35632:刘鹏飞,2013-07-29
    If mblnInitGrid = True Then SaveFlexState mshPati, App.ProductName & "\" & Me.Name
    
    mshPati.Clear
    mshPati.Rows = 2
    
    If mrsPati.EOF Then
        Call SetHeader(blnSet)
        If glngSys Like "8??" Then
            stbThis.Panels(2).Text = "当前设置没有过滤出任何客户"
        Else
            stbThis.Panels(2).Text = "当前设置没有过滤出任何病人"
        End If
    Else
        Set mshPati.DataSource = mrsPati
        Call SetHeader(blnSet)
        
        lngFamily = 0
        lngCol床号 = GetColNum("床号")
        lngCol停用 = GetColNum("停用时间")
        lngPreRow = mshPati.Row
        
        mshPati.Redraw = False
        For i = 1 To mshPati.Rows - 1
            If TabPatiState.SelectedItem.Key = "T_在院病人" Then '在院病人统计家庭病床人数
                If mshPati.TextMatrix(i, lngCol床号) = "家庭" Then
                    lngFamily = lngFamily + 1
                End If
            End If
            If mshPati.TextMatrix(i, lngCol停用) <> "" Then '停用病人红色显示
                mshPati.Row = i
                For j = 0 To mshPati.Cols - 1
                    mshPati.Col = j
                    mshPati.CellForeColor = &HC0&
                Next
            End If
        Next
        mshPati.Row = lngPreRow: mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
        mshPati.Redraw = True
        
        If glngSys Like "8??" Then
            stbThis.Panels(2) = "共 " & mrsPati.RecordCount & " 个客户"
        Else
            If TabPatiState.SelectedItem.Key = "T_在院病人" Then
                stbThis.Panels(2) = "共 " & mrsPati.RecordCount & " 个病人,其中家庭病床 " & lngFamily & " 人"
            Else
                stbThis.Panels(2) = "共 " & mrsPati.RecordCount & " 个病人"
            End If
        End If
    End If
    Call mshPati_EnterCell
    
    If Not blnSort Then Call zlCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetHeader(Optional blnSet As Boolean)
    Dim strHead As String
    Dim i As Integer
    
    mblnInitGrid = False
    
    Select Case TabPatiState.SelectedItem.Key
        Case "T_所有病人" '所有病人
            strHead = "病人ID,1,750|门诊号,1,750|住院号,1,750|就诊卡,1,850|就诊卡号,1,0|姓名,1,800|性别,1,500|年龄,1,800|门诊费别,1,850|医疗付款方式,1,1400|" & _
                "医保号,1,1200|险类,1,1500|病区,1,850|科室,1,850|床号,1,500|入院时间,1,1000|住院目的,1,800|出院时间,1,1000|住院次数,4,850|出生日期,1,1000|" & _
                "国籍,1,500|民族,1,800|区域,1,600|学历,1,500|职业,1,1000|身份,1,750|身份证号,1,2000|手机号,1,1100|家庭地址,1,2000|工作单位,1,2000|登记时间,1,1000|登记人,1,800|" & _
                "停用时间,1,0|病人性质,1,0|病人类型,1,1000|主页ID,1,0|就诊次数,1,0"
        Case "T_在院病人" '在院病人
            strHead = "病人ID,1,750|住院号,1,750|就诊卡,1,850|就诊卡号,1,0|姓名,1,800|性别,1,500|年龄,1,800|住院费别,1,850|医疗付款方式,1,1400|" & _
                "医保号,1,1200|险类,1,1500|病区,1,850|科室,1,850|床号,1,500|入院时间,1,1000|住院目的,1,800|住院次数,4,850|出生日期,1,1000|" & _
                "国籍,1,500|民族,1,800|区域,1,600|学历,1,500|职业,1,1000|身份,1,750|身份证号,1,2000|手机号,1,1100|家庭地址,1,2000|工作单位,1,2000|登记时间,1,1000|登记人,1,800|" & _
                "停用时间,1,0|病人性质,1,0|病人类型,1,1000|主页ID,1,0|就诊次数,1,0"
        Case "T_出院病人" '出院病人
            strHead = "病人ID,1,750|住院号,1,750|就诊卡,1,850|就诊卡号,1,0|姓名,1,800|性别,1,500|年龄,1,800|住院费别,1,850|医疗付款方式,1,1400|" & _
                "医保号,1,1200|险类,1,1500|入院时间,1,1000|住院目的,1,800|出院时间,1,1000|住院次数,4,850|出生日期,1,1000|国籍,1,500|民族,1,800|区域,1,600|" & _
                "学历,1,500|职业,1,1000|身份,1,750|身份证号,1,2000|手机号,1,1100|家庭地址,1,2000|工作单位,1,2000|登记时间,1,1000|登记人,1,800|停用时间,1,0|病人性质,1,0|病人类型,1,1000|主页ID,1,0|就诊次数,1,0"
        Case "T_门诊病人" '门诊病人
            If glngSys Like "8??" Then
                strHead = "客户ID,1,750|客户号,1,0|会员卡,1,850|姓名,1,800|性别,1,500|年龄,1,800|会员等级,1,850|医疗付款方式,1,1400|" & _
                    "医保号,1,1200|险类,1,1500|出生日期,1,1000|国籍,1,500|民族,1,800|区域,1,600|学历,1,500|职业,1,1000|身份,1,750|身份证号,1,2000|手机号,1,1100|" & _
                    "家庭地址,1,2000|工作单位,1,2000|登记时间,1,1000|停用时间,1,0|病人性质,1,0|病人类型,1,1000|住院次数,1,0|主页ID,1,0|就诊次数,1,0"
            Else
                strHead = "病人ID,1,750|门诊号,1,750|就诊卡,1,850|就诊卡号,1,0|姓名,1,800|性别,1,500|年龄,1,800|门诊费别,1,850|医疗付款方式,1,1400|" & _
                    "医保号,1,1200|险类,1,1500|出生日期,1,1000|国籍,1,500|民族,1,800|区域,1,600|学历,1,500|职业,1,1000|身份,1,750|身份证号,1,2000|手机号,1,1100|" & _
                    "家庭地址,1,2000|工作单位,1,2000|登记时间,1,1000|停用时间,1,0|病人性质,1,0|病人类型,1,1000|住院次数,1,0|主页ID,1,0|就诊次数,1,0"
            End If
        Case "T_留观病人" '留观病人
            strHead = "病人ID,1,750|性质,1,1000|门诊号,1,750|住院号,1,750|住院次数,1,750|就诊卡,1,850|就诊卡号,1,0|姓名,1,800|性别,1,500|年龄,1,800|门诊费别,1,850|医疗付款方式,1,1400|" & _
                    "医保号,1,1200|险类,1,1500|出生日期,1,1000|国籍,1,500|民族,1,800|区域,1,600|学历,1,500|职业,1,1000|身份,1,750|身份证号,1,2000|手机号,1,1100|" & _
                    "家庭地址,1,2000|工作单位,1,2000|登记时间,1,1000|登记人,1,800|停用时间,1,0|病人性质,1,0|病人类型,1,1000|主页ID,1,0|就诊次数,1,0"
    End Select
    
    With mshPati
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Or blnSet Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Or blnSet Then Call RestoreFlexState(mshPati, App.ProductName & "\" & Me.Name)
        
        If glngSys Like "8??" Then .ColWidth(1) = 0
        
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
        Call mshPati_EnterCell

        .Redraw = True
    End With
    mblnInitGrid = True
End Sub

Private Sub mshPati_EnterCell()
    mshPati.ForeColorSel = mshPati.CellForeColor
    Call SetMenuEnabled
    mlngGo = mshPati.Row
    mlngCurRow = mshPati.Row: mlngTopRow = mshPati.TopRow
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To mshPati.Cols - 1
        If mshPati.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
    GetColNum = -1
End Function

Private Sub mshPati_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshPati.MouseRow = 0 Then
        mshPati.MousePointer = 99
    Else
        mshPati.MousePointer = 0
    End If
End Sub

Private Sub mshPati_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshPati.MouseCol
    
    If Button = 1 And mshPati.MousePointer = 99 And mblnDown Then '双击最大化时会执行
        mblnDown = False
        
        If mshPati.TextMatrix(0, lngCol) = "" Then Exit Sub
        If glngSys Like "8??" Then
            If mshPati.TextMatrix(1, GetColNum("客户ID")) = "" Then Exit Sub
        Else
            If mshPati.TextMatrix(1, GetColNum("病人ID")) = "" Then Exit Sub
        End If
        
        Set mshPati.DataSource = Nothing
        
        If glngSys Like "8??" Then
            Select Case mshPati.TextMatrix(0, lngCol)
                Case "客户ID"
                    mrsPati.Sort = "病人ID" & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
                Case "会员卡"
                    mrsPati.Sort = "就诊卡" & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
                Case Else
                    mrsPati.Sort = mshPati.TextMatrix(0, lngCol) & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
            End Select
        Else
            mrsPati.Sort = mshPati.TextMatrix(0, lngCol) & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
        End If
        mshPati.ColData(lngCol) = (mshPati.ColData(lngCol) + 1) Mod 2
        
        Call ShowPatis(, True, gblnMyStyle)
    End If
End Sub

Private Sub SetMenuEnabled()
'功能：根据当前记录情况设置菜单可用状态
    Dim lng病人ID As Long, byt病人性质 As Byte, str停用时间 As String, lng就诊次数 As Long, lng主页ID As Long
    Dim strCard As String
    Dim blnPrivs As Boolean
    
    If glngSys Like "8??" Then
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("客户ID")))
    Else
        lng病人ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")))
    End If
    lng就诊次数 = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("就诊次数")))
    lng主页ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("主页ID")))
    strCard = Trim(mshPati.TextMatrix(mshPati.Row, GetColNum("就诊卡")))
    
    byt病人性质 = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("病人性质")))
    str停用时间 = mshPati.TextMatrix(mshPati.Row, GetColNum("停用时间"))
        
    mnuEdit_Stop.Enabled = lng病人ID <> 0 And str停用时间 = "" And mnuViewStop.Enabled And lng就诊次数 = lng主页ID '停用
    mnuEdit_Restore.Enabled = lng病人ID <> 0 And str停用时间 <> "" And mnuViewStop.Enabled And lng就诊次数 = lng主页ID '取消停用
    mnuEdit_ToInPati.Enabled = lng病人ID <> 0 And byt病人性质 = 2 And lng就诊次数 = lng主页ID                        '转为住院病人
    '----
    mnuFile_Print.Enabled = lng病人ID <> 0 And lng就诊次数 = lng主页ID                                               '打印
    mnuFile_Preview.Enabled = lng病人ID <> 0 And lng就诊次数 = lng主页ID                                             '预览
    mnuFile_Excel.Enabled = lng病人ID <> 0 And lng就诊次数 = lng主页ID                                               'excel导出
    tbr.Buttons("Print").Enabled = lng病人ID <> 0 And lng就诊次数 = lng主页ID                                        '打印
    tbr.Buttons("Preview").Enabled = lng病人ID <> 0 And lng就诊次数 = lng主页ID                                      '预览
    
    mnuEdit_Modi.Enabled = lng病人ID <> 0 And str停用时间 = "" And lng就诊次数 = lng主页ID                           '修改
    If mnuEdit_Modi.Enabled And TabPatiState.SelectedItem.Key = "T_出院病人" Then                                    '出院
        mnuEdit_Modi.Enabled = InStr(mstrPrivs, "修改出院病人") > 0
    End If
    mnuEdit_Del.Enabled = lng病人ID <> 0 And str停用时间 = "" And lng就诊次数 = lng主页ID                            '删除
    mnuEdit_Merge.Enabled = lng病人ID <> 0 And str停用时间 = "" And lng就诊次数 = lng主页ID                          '合并
    mnuEdit_View.Enabled = lng病人ID <> 0                                                                            '身份卡片
    mnuEditDelCard.Enabled = lng病人ID <> 0 And strCard <> "" And mbln是否取消绑定  '问题号:52133                    '取消卡号绑定
    
    tbr.Buttons("Modi").Enabled = lng病人ID <> 0 And str停用时间 = "" And lng就诊次数 = lng主页ID                    '修改
    If tbr.Buttons("Modi").Enabled And TabPatiState.SelectedItem.Key = "T_出院病人" Then                                    '出院
        tbr.Buttons("Modi").Enabled = InStr(mstrPrivs, "修改出院病人") > 0
    End If
    tbr.Buttons("Del").Enabled = lng病人ID <> 0 And str停用时间 = "" And lng就诊次数 = lng主页ID                     '删除
    tbr.Buttons("Merge").Enabled = lng病人ID <> 0 And str停用时间 = "" And lng就诊次数 = lng主页ID                   '合并
    tbr.Buttons("View").Enabled = lng病人ID <> 0                                           '身份卡片
    
    mnuViewGo.Enabled = lng病人ID <> 0 And lng就诊次数 = lng主页ID                                                   '定位
    tbr.Buttons("Go").Enabled = lng病人ID <> 0 And lng就诊次数 = lng主页ID                                           '定位
    mnuEditMzReCalc.Enabled = lng病人ID <> 0
    mnuEdit_Surety.Enabled = lng病人ID <> 0 And lng就诊次数 = lng主页ID                                              '担保
    mnuEdit_QueryPass.Enabled = lng病人ID <> 0 And lng就诊次数 = lng主页ID
    
    blnPrivs = InStr(";" & GetPrivFunc(glngSys, 9003) & ";", ";病人家属;") > 0
    mnuEdit_Family.Visible = blnPrivs
    mnuEdit_FamilyView.Visible = blnPrivs
    mnuEdit_FamilyAdd.Visible = blnPrivs
    mnuEdit_FamilyView.Enabled = lng病人ID <> 0
    tbr.Buttons("FamilySplit").Visible = blnPrivs
    tbr.Buttons("Family").Visible = blnPrivs
    tbr.Buttons("Family").ButtonMenus.Item("FamilyView").Enabled = lng病人ID <> 0
    '基本信息调整
    mnuEditPatiInfo.Visible = InStr(1, ";" & GetPrivFunc(glngSys, 9003) & ";", ";基本信息调整;")
    If lng病人ID <> 0 Then
        mnuEditPatiInfo.Enabled = str停用时间 = "" And mnuEditPatiInfo.Visible
    Else
        mnuEditPatiInfo.Enabled = mnuEditPatiInfo.Visible
    End If
    
End Sub

Private Sub SeekPati(blnHead As Boolean)
    Dim i As Long
    Dim blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    If glngSys Like "8??" Then
        stbThis.Panels(2).Text = "正在定位满足条件的客户,按ESC终止 ..."
    Else
        stbThis.Panels(2).Text = "正在定位满足条件的病人,按ESC终止 ..."
    End If
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshPati.Rows - 1
        DoEvents
        
        '比较条件
        blnFill = True
        With frmPatiFind
            If .txt病人ID.Text <> "" Then
                If glngSys Like "8??" Then
                    blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("客户ID")) = .txt病人ID.Text
                Else
                    blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("病人ID")) = .txt病人ID.Text
                End If
            End If
            If .txt就诊卡.Text <> "" Then
                If glngSys Like "8??" Then
                    blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("会员卡")) = .txt就诊卡.Text
                Else
                    blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("就诊卡")) = .txt就诊卡.Text
                End If
            End If
            If .txt门诊号.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("门诊号")) = .txt门诊号.Text
            End If
            If .txt住院号.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("住院号")) = .txt住院号.Text
            End If
            If .txt床号.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("床号")) = .txt床号.Text
            End If
            If .txt姓名.Text <> "" Then
                blnFill = blnFill And UCase(mshPati.TextMatrix(i, GetColNum("姓名"))) Like "*" & UCase(.txt姓名.Text) & "*"
            End If
            If .txt身份证.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("身份证号")) = .txt身份证.Text
            End If
        End With
        
        '满足则退出
        If blnFill Then
            mlngGo = i + 1
            mshPati.Row = i: mshPati.TopRow = i
            mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
            stbThis.Panels(2).Text = "找到一条记录"
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '按ESC取消
        If mblnGo = False Then
            stbThis.Panels(2).Text = "用户取消定位操作"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    stbThis.Panels(2).Text = "已定位到清单尾部"
    Screen.MousePointer = 0
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub LoadPlugInMnu()
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Integer
    Dim blnHave As Boolean
    
    If CreatePlugInOK(glngModul) Then
        blnHave = True
    End If
    
    If glngSys Like "8??" Then blnHave = False
    
    mnuEdit_PlugIn.Visible = blnHave
    tbr.Buttons("PlugIn").Visible = blnHave
    
    If blnHave Then
        On Error Resume Next
        strTmp = gobjPlugIn.GetFuncNames(glngSys, glngModul)
        Call zlPlugInErrH(Err, "GetFuncNames")
        Err.Clear: On Error GoTo 0
        
        If strTmp = "" Then Exit Sub
        strTmp = Replace(strTmp, "Auto:", "")
        arrTmp = Split(strTmp, ",")
        For i = 0 To UBound(arrTmp)
            If i <> 0 Then
                Load mnuEdit_PlugItem(i)
            End If
            mnuEdit_PlugItem(i).Caption = CStr(arrTmp(i))
            mnuEdit_PlugItem(i).Tag = CStr(arrTmp(i))
            
            If i <= 9 Then
                mnuEdit_PlugItem(i).Caption = CStr(arrTmp(i)) & "(&" & IIf(i = 9, 0, i + 1) & ")"
            End If
        Next
    End If
End Sub

Private Sub ExcPlugInFun(ByVal strFunName As String)
    Dim lngPatiId As Long
    Dim lngPageID As Long
    
    If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID"))) Then
        MsgBox "未选中任何病人，不能执行此操作！", vbExclamation, gstrSysName: Exit Sub
    End If
        
    If CreatePlugInOK(glngModul) Then
        lngPatiId = CLng(mshPati.TextMatrix(mshPati.Row, GetColNum("病人ID")))
        lngPageID = CLng(Val(mshPati.TextMatrix(mshPati.Row, GetColNum("主页ID"))))
        On Error Resume Next
        Call gobjPlugIn.ExecuteFunc(glngSys, glngModul, strFunName, lngPatiId, lngPageID, 0)
        Call zlPlugInErrH(Err, "ExecuteFunc")
        Err.Clear: On Error GoTo 0
    End If
End Sub
