VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmKind 
   Caption         =   "体检类型设置"
   ClientHeight    =   6720
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10455
   Icon            =   "frmKind.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6360
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmKind.frx":1CFA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13361
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
   Begin MSComctlLib.TreeView tvw 
      Height          =   1770
      Left            =   330
      TabIndex        =   4
      Top             =   945
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3122
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1035
      Top             =   4695
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
            Picture         =   "frmKind.frx":258E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":29E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1485
      Left            =   3135
      TabIndex        =   3
      Top             =   870
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   2619
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   405
      Top             =   4695
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
            Picture         =   "frmKind.frx":2CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":314C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   10455
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "分类"
               Key             =   "分类"
               Object.ToolTipText     =   "分类"
               Object.Tag             =   "分类"
               ImageKey        =   "Class"
               Style           =   5
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "增加"
               Object.ToolTipText     =   "增加"
               Object.Tag             =   "增加"
               ImageKey        =   "Add"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "删除"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "项目"
               Key             =   "项目"
               Object.ToolTipText     =   "项目"
               Object.Tag             =   "项目"
               ImageKey        =   "Item"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "列表"
               Key             =   "列表"
               Object.ToolTipText     =   "列表"
               Object.Tag             =   "列表"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Large"
                     Text            =   "大图标(&G)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Text            =   "小图标(&M)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Text            =   "列表(&L)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Text            =   "详细资料(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   8325
      Top             =   3210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":3466
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":3686
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":38A6
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":3AC2
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":3CDE
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":3EF8
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":4118
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":4338
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":4558
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":4778
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   7515
      Top             =   3210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":4998
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":4BB8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":4DD8
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":4FF4
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":5210
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":5562
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":5782
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":59A2
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":5BC2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKind.frx":5DE2
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid vsfPrint 
      Height          =   780
      Left            =   5940
      TabIndex        =   5
      Top             =   1305
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   1376
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   270
      MergeCells      =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1740
      Left            =   3135
      TabIndex        =   6
      Top             =   2535
      Width           =   3030
      _cx             =   5345
      _cy             =   3069
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Image imgX_S 
      Height          =   45
      Left            =   4860
      MousePointer    =   7  'Size N S
      Top             =   2385
      Width           =   5115
   End
   Begin VB.Image imgY_S 
      Height          =   4395
      Left            =   2670
      MousePointer    =   9  'Size W E
      Top             =   840
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileOutExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditClass 
         Caption         =   "类型分类(&C)"
         Begin VB.Menu mnuEditClassAdd 
            Caption         =   "增加分类(&A)"
         End
         Begin VB.Menu mnuEditClassModify 
            Caption         =   "修改分类(&M)"
         End
         Begin VB.Menu mnuEditClassDelete 
            Caption         =   "删除分类(&D)"
         End
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAdd 
         Caption         =   "增加类型(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改类型(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除类型(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelect 
         Caption         =   "体检项目(&S)"
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
      Begin VB.Menu mnuViewIcon 
         Caption         =   "大图标(&G)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "详细资料(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web上的中联"
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
      Begin VB.Menu h1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmKind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mstrVsf As String                               '表格列标题
Private mstrKey As String                               '保存以前的选择
Private Const mstrLvw As String = "名称,2400,0,1;编码,900,0,0;简码,900,0,0;基本价格,1200,1,0;体检价格,1200,1,0;折扣,900,1,0;适用,900,0,0,;说明,1500,0,0;所属分类,1500,0,0"
Private mlngLoop As Long
Private WithEvents mobjPopMenu As clsPopMenu                '自定义弹出菜单对象
Attribute mobjPopMenu.VB_VarHelpID = -1
Private mbytPopMenu As Byte

Private Enum mCol
    项目名称 = 0
    项目类别
    检查部位
    采集方式
    检验标本
    基本价格
    体检价格
    折扣
End Enum


'（２）自定义过程或函数************************************************************************************************
Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据，发生在窗体的Load事件
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
    
    mstrKey = ""
    lvw.Tag = "可变化的"
        
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", "0")) = 1 Then
        '使用个性化设置
        mstrVsf = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "表格标题", mstrVsf)
                        
    End If
    
    If lvw.ListItems.Count = 0 Then zlControl.LvwSelectColumns lvw, mstrLvw, True
                
    mstrVsf = "项目名称,3000,1,1,1,;项目类别,900,1,1,1,;检查部位,1200,1,1,1,;采集方式,1200,1,1,1,;检验标本,900,1,1,1,;基本价格,1080,7,1,1,;体检价格,1080,7,1,1,;折扣,1080,7,1,1,"
    Call CreateVsf(vsf, mstrVsf)
    
    vsf.ColFormat(mCol.基本价格) = "0.00"
    vsf.ColFormat(mCol.体检价格) = "0.00"
    vsf.ColFormat(mCol.折扣) = "0.000"
    InitLoad = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ApplyPrivilege(ByVal strPrivilege As String)
    '------------------------------------------------------------------------------------------------------------------
    '功能： 应用权限处理
    '参数： strPrivilege                    权限
    '------------------------------------------------------------------------------------------------------------------
    
    '调试语句
    'strPrivilege = "基本;增删改"
    
    '不具有“增删改”权限时
    If InStr(strPrivilege, "增删改") = 0 Then
        mnuEdit.Visible = False
    End If
    
    tbrThis.Buttons("分类").Visible = mnuEdit.Visible
    
    tbrThis.Buttons("增加").Visible = mnuEdit.Visible
    tbrThis.Buttons("修改").Visible = mnuEdit.Visible
    tbrThis.Buttons("删除").Visible = mnuEdit.Visible
    
    tbrThis.Buttons("Split_2").Visible = mnuEdit.Visible
    tbrThis.Buttons("Split_3").Visible = mnuEdit.Visible
    
    tbrThis.Buttons("项目").Visible = mnuEdit.Visible
    tbrThis.Buttons("Split_4").Visible = mnuEdit.Visible
End Sub

Private Sub AdjustEnableState()
    '------------------------------------------------------------------------------------------------------------------
    '功能： 调整各功能菜单的可用状态
    '------------------------------------------------------------------------------------------------------------------
    
    mnuFilePrint.Enabled = True
    mnuFilePrintView.Enabled = True
    mnuFileOutExcel.Enabled = True
    
    mnuEditClassAdd.Enabled = True
    mnuEditClassModify.Enabled = True
    mnuEditClassDelete.Enabled = True
    
    mnuEditClass.Enabled = True
    
    mnuEditAdd.Enabled = True
    mnuEditModify.Enabled = True
    mnuEditDelete.Enabled = True
    
    mnuEditSelect.Enabled = True
    
    If lvw.ListItems.Count = 0 Then
                
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
        
        mnuEditSelect.Enabled = False
    End If
    
    If Val(vsf.RowData(1)) = 0 Then
        mnuFilePrint.Enabled = False
        mnuFilePrintView.Enabled = False
        mnuFileOutExcel.Enabled = False
    End If
    
    If tvw.SelectedItem.Key = "K0" Then
        mnuEditClassModify.Enabled = False
        mnuEditClassDelete.Enabled = False
    End If
    
    tbrThis.Buttons("预览").Enabled = mnuFilePrintView.Enabled
    tbrThis.Buttons("打印").Enabled = mnuFilePrint.Enabled
    
    tbrThis.Buttons("分类").Enabled = mnuEditClassModify.Enabled Or mnuEditClassAdd.Enabled
    tbrThis.Buttons("增加").Enabled = mnuEditAdd.Enabled
    tbrThis.Buttons("修改").Enabled = mnuEditModify.Enabled
    tbrThis.Buttons("删除").Enabled = mnuEditDelete.Enabled
    tbrThis.Buttons("项目").Enabled = mnuEditSelect.Enabled
    
End Sub

Private Sub RefreshStateInfo()
    '------------------------------------------------------------------------------------------------------------------
    '功能：刷新状态栏显示信息
    '------------------------------------------------------------------------------------------------------------------
    If lvw.SelectedItem Is Nothing Then
        stbThis.Panels(2).Text = "共有 " & lvw.ListItems.Count & " 个体检类型！"
    Else
        If vsf.Rows = 2 And vsf.RowData(1) = 0 Then
            stbThis.Panels(2).Text = "共有 " & lvw.ListItems.Count & " 个体检类型！"
        Else
            stbThis.Panels(2).Text = "共有 " & lvw.ListItems.Count & " 个体检类型，“" & lvw.SelectedItem.Text & "”下有 " & vsf.Rows - 1 & " 个体检项目！"
        End If
    End If
    
End Sub

Public Function GetItem(ByRef lngKey As Long, ByVal intFoot As Integer) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：供编辑数据窗体调用，接口函数
    '------------------------------------------------------------------------------------------------------------------
    Dim lngIndex As Long
    Dim objItem As ListItem
    
    On Error GoTo errHand
    
    Set objItem = lvw.ListItems("K" & lngKey)
    If Not (objItem Is Nothing) Then
        
        lngIndex = objItem.Index
        lngIndex = lngIndex + intFoot
        
        Set objItem = Nothing
        Set objItem = lvw.ListItems(lngIndex)
        
        If Not (objItem Is Nothing) Then lngKey = Val(Mid(objItem.Key, 2))
            
        GetItem = True
    Else
        GetItem = False
    End If
    
    Exit Function
    
errHand:
    
End Function

Public Function EditRefresh(ByVal strMenuItem As String, ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：供编辑数据窗体调用，接口函数
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
    
    On Error GoTo errHand

    Select Case strMenuItem
    Case "体检类型分类"
        
        Call ClearData("体检类型分类;体检类型;体检项目")
        
        Call RefreshData("体检类型分类")
        
        On Error Resume Next
        tvw.Nodes("K" & lngKey).Selected = True
        tvw.Nodes("K" & lngKey).EnsureVisible
        On Error GoTo 0
        
        Call RefreshData("体检类型")
        Call RefreshData("体检项目")
        
        
    Case "体检类型"
    
        Call ClearData("体检类型;体检项目")
        Call RefreshData("体检类型")
        
        '恢复体检类型
        Call zlControl.LvwRestoreItem(lvw, "K" & lngKey)
            
        Call RefreshData("体检项目")
                
    Case "体检项目"
        If lvw.SelectedItem.Key = "K" & lngKey Then
            mstrKey = ""
            Call lvw_ItemClick(lvw.SelectedItem)
        End If
    End Select
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    
    strMenuItem = ";" & strMenuItem & ";"
    
    If InStr(strMenuItem, ";体检类型分类;") > 0 Then
        tvw.Nodes.Clear
    End If
    
    If InStr(strMenuItem, ";体检类型;") > 0 Then
        lvw.ListItems.Clear
    End If
    
    If InStr(strMenuItem, ";体检类型;") > 0 Then
        lvw.ListItems.Clear
    End If
    If InStr(strMenuItem, ";体检项目;") > 0 Then
        Call ResetVsf(vsf)
    End If
        
End Function

Private Function RefreshData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：刷新/装载数据
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim rs As New ADODB.Recordset
    Dim objNode As Node
    Dim rsPrice As New ADODB.Recordset
    
    On Error GoTo errHand
    
    Select Case strMenuItem
    Case "体检类型分类"
        
        gstrSQL = GetPublicSQL(SQL.体检类型分类)
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        If rs.BOF = False Then Call FillTreeData(tvw, rs)
        
    Case "体检类型"
        
        If Val(Mid(tvw.SelectedItem.Key, 2)) = 0 Then
            
            gstrSQL = "select a.序号 AS ID," & _
                            "a.编码," & _
                            "a.名称," & _
                            "a.简码," & _
                            "DECODE(a.适用范围,0,'所有',1,'个人',2,'团体') AS 适用," & _
                            "Trim(To_Char(c.基本价格,'99999999.00')) As 基本价格," & _
                            "Trim(To_Char(c.体检价格,'99999999.00')) As 体检价格," & _
                            "Trim(To_Char(Decode(c.基本价格,Null,0,0,0,10*c.体检价格/c.基本价格),'99999999.000')) As 折扣," & _
                            "a.说明," & _
                            "b.名称 AS 所属分类," & _
                            "1 as 图标 " & _
                    "from 体检类型 a," & _
                         "体检类型 b," & _
                         "(Select a.序号,Sum(b.现价*a.数次) As 基本价格,Sum(b.现价*a.数次*Nvl(a.折扣,1)) As 体检价格 " & _
                         "From 体检类型计价 a," & _
                              "收费价目 b " & _
                         "Where b.收费细目id=a.收费细目id " & _
                               "and b.执行日期<=SYSDATE " & _
                               "and (b.终止日期 IS NULL OR b.终止日期>SYSDATE) " & _
                         "group by a.序号 " & _
                         ") c " & _
                    "where a.末级 = 1 AND a.上级序号 = b.序号(+) " & _
                          "and a.序号=c.序号(+)"

        Else
            
            gstrSQL = "select a.序号 AS ID," & _
                            "a.编码," & _
                            "a.名称," & _
                            "a.简码," & _
                            "DECODE(a.适用范围,0,'所有',1,'个人',2,'团体') AS 适用," & _
                            "Trim(To_Char(c.基本价格,'99999999.00')) As 基本价格," & _
                            "Trim(To_Char(c.体检价格,'99999999.00')) As 体检价格," & _
                            "Trim(To_Char(Decode(c.基本价格,Null,0,0,0,10*c.体检价格/c.基本价格),'99999999.00')) As 折扣," & _
                            "a.说明," & _
                            "b.名称 AS 所属分类," & _
                            "1 as 图标 " & _
                    "from 体检类型 a," & _
                         "体检类型 b," & _
                         "(Select a.序号,Sum(b.现价*a.数次) As 基本价格,Sum(b.现价*a.数次*Nvl(a.折扣,1)) As 体检价格 " & _
                         "From 体检类型计价 a," & _
                              "收费价目 b " & _
                         "Where b.收费细目id=a.收费细目id " & _
                               "and b.执行日期<=SYSDATE " & _
                               "and (b.终止日期 IS NULL OR b.终止日期>SYSDATE) " & _
                         "group by a.序号 " & _
                         ") c " & _
                    "where a.末级 = 1 AND a.上级序号 = b.序号(+) " & _
                          "and a.序号=c.序号(+) and a.上级序号=[1]"
            
        End If
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(tvw.SelectedItem.Key, 2)))
        If rs.BOF = False Then Call FillLvw(lvw, rs)
                
    Case "体检项目"
        
        If lvw.SelectedItem Is Nothing Then Exit Function
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        gstrSQL = GetPublicSQL(SQL.体检类型项目, CStr(lngKey))
    
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        
        If rs.BOF = False Then
            Do While Not rs.EOF
                
                If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                    vsf.Rows = vsf.Rows + 1
                End If
                
                vsf.RowData(vsf.Rows - 1) = zlCommFun.NVL(rs("ID"), 0)
                vsf.TextMatrix(vsf.Rows - 1, mCol.项目名称) = zlCommFun.NVL(rs("项目名称"))
                vsf.TextMatrix(vsf.Rows - 1, mCol.项目类别) = zlCommFun.NVL(rs("项目类别"))
                vsf.TextMatrix(vsf.Rows - 1, mCol.检查部位) = zlCommFun.NVL(rs("检查部位"))
                vsf.TextMatrix(vsf.Rows - 1, mCol.采集方式) = zlCommFun.NVL(rs("采集方式"))
                vsf.TextMatrix(vsf.Rows - 1, mCol.检验标本) = zlCommFun.NVL(rs("检验标本"))
                vsf.TextMatrix(vsf.Rows - 1, mCol.基本价格) = zlCommFun.NVL(rs("基本价格"))
                vsf.TextMatrix(vsf.Rows - 1, mCol.体检价格) = zlCommFun.NVL(rs("体检价格"))
                vsf.TextMatrix(vsf.Rows - 1, mCol.折扣) = zlCommFun.NVL(rs("折扣"))
                                            
                rs.MoveNext
            Loop
        End If
    End Select
    
    RefreshData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function MenuClick(ByVal strMenuItem As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：数据编辑/处理
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
                
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    Select Case strMenuItem
    Case "增加分类"
        
        If tvw.SelectedItem Is Nothing Then Exit Function
        
        If Not frmKindClassEdit.ShowEdit(Me, 0, Val(Mid(tvw.SelectedItem.Key, 2))) Then Exit Function
        
        
    Case "修改分类"
        
        If tvw.SelectedItem Is Nothing Then Exit Function
        If tvw.SelectedItem.Key = "K0" Then Exit Function
                        
        If Not frmKindClassEdit.ShowEdit(Me, Val(Mid(tvw.SelectedItem.Key, 2)), Val(Mid(tvw.SelectedItem.Parent.Key, 2))) Then Exit Function
        
        
    Case "删除分类"
        
        If tvw.SelectedItem Is Nothing Then Exit Function
        If tvw.SelectedItem.Key = "K0" Then Exit Function
        
        If MsgBox("你真的要删除“" & tvw.SelectedItem.Text & "”分类？" & vbCrLf & "删除分类同时也删除对应的体检类型和体检项目。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        lngKey = Val(Mid(tvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        strSQL(ReDimArray(strSQL)) = "ZL_体检类型_DELETE(" & lngKey & ")"
        
    Case "增加类型"
        If tvw.SelectedItem Is Nothing Then Exit Function
        
        If Not frmKindEdit.ShowEdit(Me, 0, Val(Mid(tvw.SelectedItem.Key, 2))) Then Exit Function
        
    Case "修改类型"
        If tvw.SelectedItem Is Nothing Then Exit Function
        If lvw.SelectedItem Is Nothing Then Exit Function
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        If Not frmKindEdit.ShowEdit(Me, lngKey, Val(Mid(tvw.SelectedItem.Key, 2))) Then Exit Function
        
    Case "删除类型"
        If tvw.SelectedItem Is Nothing Then Exit Function
        If lvw.SelectedItem Is Nothing Then Exit Function
        
        If MsgBox("你真的要删除“" & lvw.SelectedItem.Text & "”及对应的体检项目？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        strSQL(ReDimArray(strSQL)) = "ZL_体检类型_DELETE(" & lngKey & ")"
        
    Case "体检项目"
        If lvw.SelectedItem Is Nothing Then Exit Function
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        frmKindCustom.ShowEdit Me, lngKey
        
    End Select
    
    blnTran = True
    gcnOracle.BeginTrans
    
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTran = False
        
    Select Case strMenuItem
    Case "删除分类"
        
        If Not (tvw.SelectedItem Is Nothing) Then tvw.Nodes.Remove tvw.SelectedItem.Index
        
        Call ClearData("体检类型;体检项目")
        Call RefreshData("体检类型")
        If Not (lvw.SelectedItem Is Nothing) Then Call RefreshData("体检项目")
                
        
    Case "删除类型"
    
        '删除行
        lngLoop = lvw.SelectedItem.Index
        lvw.ListItems.Remove lngLoop
        Call NextLvwPos(lvw, lngLoop)
        
    End Select
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
    MenuClick = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
End Function

Private Sub PrintData(ByVal bytMode As Byte)
    '------------------------------------------------------------------------------------------------------------------
    '功能： 打印数据
    '参数： bytMode                         打印方式（1-打印；2-预览；3-输出到Excel）
    '------------------------------------------------------------------------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    If UserInfo.姓名 = "" Then Call GetUserInfo

    objPrint.Title.Text = "体检项目清单"
    Call CopyGrid(vsf, vsfPrint)
    
    Set objRow = New zlTabAppRow
    objRow.Add "类型：" & lvw.SelectedItem.Text
    objRow.Add ""
    
    objPrint.UnderAppRows.Add objRow
    
    Set objPrint.Body = vsfPrint

    If bytMode = 1 Then bytMode = zlPrintAsk(objPrint)

    If bytMode >= 1 And bytMode <= 3 Then Call zlPrintOrView1Grd(objPrint, bytMode)
        
End Sub





'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    DoEvents
    
    Call mnuViewIcon_Click(lvw.View)
    Call mnuViewRefresh_Click
    
    mblnStartUp = False
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    
    Call RestoreWinState(Me, App.ProductName)
    Call InitLoad
    Call ApplyPrivilege(gstrPrivs)
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    '处理特殊情况
    
    If imgX_S.Top > Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1000 Then
        imgX_S.Top = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1000
    End If
    
    If imgY_S.Left > Me.ScaleWidth - 1000 Then
        imgY_S.Left = Me.ScaleWidth - 1000
    End If
    
    With tvw
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = imgY_S.Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With imgY_S
        .Top = tvw.Top
        .Height = tvw.Height
    End With
    
    With lvw
        .Left = imgY_S.Left + imgY_S.Width
        .Top = tvw.Top
        .Width = Me.ScaleWidth - .Left
        .Height = imgX_S.Top - .Top
    End With
    
    With imgX_S
        .Left = lvw.Left
        .Width = lvw.Width
    End With
    
    With vsf
        .Left = lvw.Left
        .Top = imgX_S.Top + imgX_S.Height
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top
    End With

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Cancel = mblnStartUp
    If Cancel Then Exit Sub
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "表格标题", mstrVsf)
                
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgX_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    imgX_S.Top = imgX_S.Top + Y
    
    If imgX_S.Top < 1500 Then imgX_S.Top = 1500
    If Me.Height - imgX_S.Top - imgX_S.Height < 1000 Then imgX_S.Top = Me.Height - imgX_S.Height - 1000
    
            
    Form_Resize
End Sub

Private Sub imgY_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    imgY_S.Left = imgY_S.Left + X
    
    If imgY_S.Left < 1500 Then imgY_S.Left = 1500
    If Me.Width - imgY_S.Left - imgY_S.Width < 1000 Then imgY_S.Left = Me.Width - imgY_S.Width - 1000

    Form_Resize
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvw, ColumnHeader.Index)
End Sub

Private Sub lvw_DblClick()
    If mnuEdit.Visible And mnuEditModify.Visible And mnuEditModify.Enabled Then Call mnuEditModify_Click
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lngKey As Long
    
    If mstrKey = Item.Key Then Exit Sub
    mstrKey = Item.Key
    
    '保存
    lngKey = Val(vsf.RowData(vsf.Row))
    
    Call ClearData("体检项目")
    Call RefreshData("体检项目")
    
    '恢复
    Call AdjustEnableState
    Call RefreshStateInfo
    
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call lvw_DblClick
End Sub

Private Sub lvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    
    mbytPopMenu = 2
    Set mobjPopMenu = New clsPopMenu
    mobjPopMenu.ShowPopupMenuByCursor

End Sub

Private Sub mnuEditClassAdd_Click()
    Call MenuClick("增加分类")
End Sub

Private Sub mnuEditClassDelete_Click()
    Call MenuClick("删除分类")
End Sub

Private Sub mnuEditClassModify_Click()
    Call MenuClick("修改分类")
End Sub

Private Sub mnuEditDelete_Click()
    Call MenuClick("删除类型")
End Sub

Private Sub mnuEditModify_Click()
    Call MenuClick("修改类型")
End Sub

Private Sub mnuEditAdd_Click()
    Call MenuClick("增加类型")
End Sub

Private Sub mnuEditSelect_Click()
    Call MenuClick("体检项目")
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOutExcel_Click()
    Call PrintData(3)
End Sub

Private Sub mnuFilePrint_Click()
    Call PrintData(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFilePrintView_Click()
    Call PrintData(2)
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    mnuViewIcon(0).Checked = False
    mnuViewIcon(1).Checked = False
    mnuViewIcon(2).Checked = False
    mnuViewIcon(3).Checked = False
    
    mnuViewIcon(Index).Checked = True
    
    lvw.View = Index
End Sub

Private Sub mnuViewRefresh_Click()
    Dim strKey As String
    Dim strKeyClass As String
        
    '保存体检类型分类、体检类型
    If Not (tvw.SelectedItem Is Nothing) Then strKeyClass = tvw.SelectedItem.Key
    strKey = zlControl.LvwSaveItem(lvw)
            
    Call ClearData("体检项目;体检类型;体检类型分类")
    
    Call RefreshData("体检类型分类")
    
    '恢复刷新前选择的体检类型分类
    
    tvw.Nodes(1).Selected = True
    tvw.Nodes(1).Expanded = True
    
    On Error Resume Next
    tvw.Nodes(strKeyClass).Selected = True
    tvw.Nodes(strKeyClass).EnsureVisible
    On Error GoTo 0
    
    If Not (tvw.SelectedItem Is Nothing) Then
        Call RefreshData("体检类型")
        
        '恢复刷新前选择的体检类型
        Call zlControl.LvwRestoreItem(lvw, strKey)
        
        If Not (lvw.SelectedItem Is Nothing) Then Call RefreshData("体检项目")
    End If
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub mobjPopMenu_MenuBeforeShow(Cancel As Boolean)
    
    Select Case mbytPopMenu
    Case 1
        
        If mnuEdit.Visible = False Then Exit Sub
        
        If mnuEditClassAdd.Visible Then mobjPopMenu.Add 1, mnuEditClassAdd.Caption, , , mnuEditClassAdd.Enabled
        If mnuEditClassModify.Visible Then mobjPopMenu.Add 2, mnuEditClassModify.Caption, , , mnuEditClassModify.Enabled
        If mnuEditClassDelete.Visible Then mobjPopMenu.Add 3, mnuEditClassDelete.Caption, , , mnuEditClassDelete.Enabled
        
    Case 2
        
        If mnuEdit.Visible = False Then Exit Sub
        
        If mnuEditAdd.Visible Then mobjPopMenu.Add 1, mnuEditAdd.Caption, , , mnuEditAdd.Enabled
        If mnuEditModify.Visible Then mobjPopMenu.Add 2, mnuEditModify.Caption, , , mnuEditModify.Enabled
        If mnuEditDelete.Visible Then mobjPopMenu.Add 3, mnuEditDelete.Caption, , , mnuEditDelete.Enabled
        
        mobjPopMenu.Add 4, "-", , 2, True
        
        If mnuEditSelect.Visible Then mobjPopMenu.Add 5, mnuEditSelect.Caption, , , mnuEditSelect.Enabled
    
    End Select
    
End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case mbytPopMenu
    Case 1
        Select Case Key
        Case 1
            Call mnuEditClassAdd_Click
        Case 2
            Call mnuEditClassModify_Click
        Case 3
            Call mnuEditClassDelete_Click
        End Select
    Case 2
        Select Case Key
        Case 1
            Call mnuEditAdd_Click
        Case 2
            Call mnuEditModify_Click
        Case 3
            Call mnuEditDelete_Click
        Case 5
            Call mnuEditSelect_Click
        End Select
    End Select
End Sub


Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(tbrThis.hWnd, objPoint)
    
    Select Case Button.Key
    Case "预览"
        Call mnuFilePrintView_Click
    Case "打印"
        
        Call mnuFilePrint_Click
        
    Case "分类"
                
        mbytPopMenu = 1
        Set mobjPopMenu = New clsPopMenu
        Call mobjPopMenu.ShowPopupMenu(objPoint.X * 15 + Button.Left - 15, objPoint.Y * 15 + Button.Top + Button.Height + 15)
        
    Case "增加"
        Call mnuEditAdd_Click
    Case "修改"
        Call mnuEditModify_Click
    Case "删除"
        Call mnuEditDelete_Click
    Case "项目"
        Call mnuEditSelect_Click
    Case "列表"
        Call mnuViewIcon_Click(IIf(lvw.View = 3, 0, lvw.View + 1))
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Call tbrThis_ButtonClick(Button)
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "Large"
        Call mnuViewIcon_Click(0)
    Case "Small"
        Call mnuViewIcon_Click(1)
    Case "List"
        Call mnuViewIcon_Click(2)
    Case "Detail"
        Call mnuViewIcon_Click(3)
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub tvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
    
        mbytPopMenu = 1
        Set mobjPopMenu = New clsPopMenu
        mobjPopMenu.ShowPopupMenuByCursor
        
    End If
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Call ClearData("体检类型;体检项目")
    
    Call RefreshData("体检类型")
    
    If Not (lvw.SelectedItem Is Nothing) Then Call RefreshData("体检项目")
    
    Call AdjustEnableState
    
End Sub

Private Sub vsf_GotFocus()
    vsf.BackColorSel = COLOR.焦点
End Sub

Private Sub vsf_LostFocus()
    vsf.BackColorSel = COLOR.非焦点
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

