VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmChargeSortGrade 
   Caption         =   "费别等级管理"
   ClientHeight    =   6000
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10005
   Icon            =   "frmChargeSortGrade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab sstabItem 
      Height          =   3255
      Left            =   5040
      TabIndex        =   5
      Top             =   1080
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   5741
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "  收入项目(&0)  "
      TabPicture(0)   =   "frmChargeSortGrade.frx":0582
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fgdDetail"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "  收费项目(&1)  "
      TabPicture(1)   =   "frmChargeSortGrade.frx":059E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgdItem"
      Tab(1).ControlCount=   1
      Begin VSFlex8Ctl.VSFlexGrid fgdDetail 
         Height          =   1935
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   3735
         _cx             =   6588
         _cy             =   3413
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   -2147483631
         GridColorFixed  =   -2147483630
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
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
      Begin VSFlex8Ctl.VSFlexGrid fgdItem 
         Height          =   1935
         Left            =   -74640
         TabIndex        =   7
         Top             =   600
         Width           =   3735
         _cx             =   6588
         _cy             =   3413
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   -2147483631
         GridColorFixed  =   -2147483630
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
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
   End
   Begin VB.PictureBox picSplitV 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   4770
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   45
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   45
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   10005
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   615
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgToolsStard"
         HotImageList    =   "imgToolsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "打印预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "New"
               Object.ToolTipText     =   "增加床位等级"
               Object.Tag             =   "增加"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modify"
               Object.ToolTipText     =   "修改床位等级"
               Object.Tag             =   "修改"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.ToolTipText     =   "删除床位等级"
               Object.Tag             =   "删除"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "View"
               Object.ToolTipText     =   "查看方式"
               Object.Tag             =   "查看"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  大图标"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  小图标"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  列表"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  详细资料"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgToolsHot 
      Left            =   8280
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":05BA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":07D4
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":09EE
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":0C08
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":0E22
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":103C
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":125C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":1476
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolsStard 
      Left            =   9000
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":1690
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":18AA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":1AC4
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":1CDE
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":1EF8
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":2112
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":2332
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":254C
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   3090
      Top             =   5250
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
            Picture         =   "frmChargeSortGrade.frx":2766
            Key             =   "Key"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":2A80
            Key             =   "KeyD"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2310
      Top             =   5310
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
            Picture         =   "frmChargeSortGrade.frx":2D9A
            Key             =   "Key"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeSortGrade.frx":2EF4
            Key             =   "KeyD"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain_S 
      Height          =   4755
      Left            =   30
      TabIndex        =   0
      Top             =   810
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8387
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5640
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   635
      SimpleText      =   $"frmChargeSortGrade.frx":304E
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmChargeSortGrade.frx":3095
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12568
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
      Begin VB.Menu mnusplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "增加(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "单项收费编辑(&S)"
      End
      Begin VB.Menu mnuEditUnion 
         Caption         =   "统一实收比率(&U)"
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
         Begin VB.Menu mnuViewToolspilt1 
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
      Begin VB.Menu mnuviewsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "大图标(&G)"
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
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewColumn 
         Caption         =   "选择列(&C)"
      End
      Begin VB.Menu mnuViewSplit4 
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
      Begin VB.Menu mnuHelpSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
   Begin VB.Menu mnuShort2 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "增加(&A)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "修改(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "删除(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuShortsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "大图标(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "详细资料(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmChargeSortGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrsDetail As New ADODB.Recordset
Dim mrsItem As New ADODB.Recordset
Dim msngStartX As Single
Dim mblnLoad As Boolean
Dim mintColumn As Integer
Dim mblnItem As Boolean
Private Const mstrLvw As String = "名称,1200,0,1;编码,600,0,2;简码,600,0,0;有效期开始时间,1500,0,0;有效期结束时间,1500,0,0;适用科室,900,0,0;属性,1000,0,0;仅限初诊,900,0,0;说明,2000,0,0"
Private mlngMode As Long
Private mstrPrivs As String                              '权限串
Private mstrCharge As String
Private Sub SetItemMenu()
    Dim blnEnabled As Boolean
    If Me.ActiveControl Is fgdItem Then
        blnEnabled = (fgdItem.Rows > 1 And fgdItem.TextMatrix(1, 0) <> "")
        Toolbar1.Buttons("New").Enabled = True
        Toolbar1.Buttons("Modify").Enabled = blnEnabled
        Toolbar1.Buttons("Delete").Enabled = blnEnabled
        mnuEditAdd.Enabled = True
        mnuEditDelete.Enabled = blnEnabled
        mnuEditModify.Enabled = blnEnabled
        
        mnuEditItem.Visible = False
        mnuEditUnion.Visible = False
        mnuEditSplit.Visible = False
    End If
    
    If Me.ActiveControl Is fgdDetail Then
        Toolbar1.Buttons("New").Enabled = False
        Toolbar1.Buttons("Modify").Enabled = False
        Toolbar1.Buttons("Delete").Enabled = False
        mnuEditAdd.Enabled = False
        mnuEditDelete.Enabled = False
        mnuEditModify.Enabled = False
        
        mnuEditItem.Visible = True
        mnuEditUnion.Visible = True
        mnuEditSplit.Visible = True
    End If
End Sub

Private Sub fgdDetail_DblClick()
    If mnuEditModify.Visible = False Then Exit Sub
    If mnuEditItem.Enabled = False Then Exit Sub
    mnuEditItem_Click
End Sub


Private Sub fgdDetail_EnterCell()
    Call SetItemMenu
End Sub


Private Sub fgdDetail_LostFocus()
    Call SetMenu
End Sub

Private Sub fgdItem_DblClick()
    If mnuEditModify.Visible = False Then Exit Sub
    If mnuEditItem.Enabled = False Then Exit Sub
    mnuEditItem_Click
End Sub


Private Sub fgdItem_EnterCell()
    Call SetItemMenu
End Sub


Private Sub fgdItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If fgdItem.Rows > 1 Then
            PopupMenu mnuEdit, 2
        End If
    End If
End Sub


Private Sub Form_Activate()
    Dim rsItem As New ADODB.Recordset
    
    Call Form_Resize '为了正确计算coolbar的高度
    
    On Error GoTo ErrHandle
    If mblnLoad = True Then
        gstrSQL = "select ID,名称,编码 from 收入项目 where 末级=1 and rownum<2"
        Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

        If rsItem.RecordCount = 0 Then
            MsgBox "没找到收入项目，不能运行费别等级管理。" & vbCrLf & "请在《收入项目管理》中设置收入项目。", vbExclamation, "提示"
            Unload Me
            Exit Sub
        End If
    
        Call FillList
    End If
    mblnLoad = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim Item As ListItem
    
    'ByZT20030722
    If glngSys Like "8??" Then
        Caption = "会员等级管理"
    End If
    
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    Call 权限控制
    
    '允许进行列删除的ListView须做标记
    lvwMain_S.Tag = "可变化的"
    '-----------
    RestoreWinState Me, App.ProductName
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    '如果ListView的还未被设置，比如第一次使用，那就调用缺省的初始化
    If lvwMain_S.ColumnHeaders.Count = 0 Then
        zlControl.LvwSelectColumns lvwMain_S, mstrLvw, True
    End If
    '根据LvwMain显示设置对应菜单
     mnuViewIcon_Click lvwMain_S.View
    mblnLoad = True
    
    '进行部分初始化
    With fgdDetail
        .Cols = 5
        .ColWidth(0) = 0
        .ColWidth(1) = 1700
        .ColWidth(2) = 2200
        .ColWidth(3) = 1050
        .ColWidth(4) = 2000
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .TextMatrix(0, 1) = "收入项目"
        .TextMatrix(0, 2) = "应收金额(元)"
        .TextMatrix(0, 3) = "实收比率(%)"
        .TextMatrix(0, 4) = "计算方法"
    End With
    
    With fgdItem
        .Cols = 5
        .ColWidth(0) = 0
        .ColWidth(1) = 2500
        .ColWidth(2) = 3000
        .ColWidth(3) = 1050
        .ColWidth(4) = 2000
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .TextMatrix(0, 1) = "收费项目"
        .TextMatrix(0, 2) = "应收金额(元)"
        .TextMatrix(0, 3) = "实收比率(%)"
        .TextMatrix(0, 4) = "计算方法"
    End With
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    SizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwMain_S_DblClick()
    If mblnItem Then
        If mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
    End If
End Sub

Private Sub lvwMain_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Tag <> "" Then stbThis.Panels(2).Text = "说明：" & Item.Tag
    mblnItem = True
    mstrCharge = lvwMain_S.SelectedItem.Text
    If sstabItem.Tab = 0 Then
        Call FillDetail
    Else
        Call FillItem
    End If
End Sub
Private Sub lvwMain_S_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
    End If
End Sub

Private Sub lvwMain_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub lvwMain_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Button = 2 Then
        mnuShortMenu2(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu2(3).Enabled = mnuEditDelete.Enabled
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuShort2, vbPopupMenuRightButton
    End If
End Sub

Private Sub lvwMain_S_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwMain_S.SortOrder = IIF(lvwMain_S.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain_S.SortKey = mintColumn
        lvwMain_S.SortOrder = lvwAscending
    End If
End Sub

Private Sub mnuEditAdd_Click()
    If sstabItem.Tab = 0 Then
        Call frmChargeSortEdit.编辑费别("")
    Else
        If frmChargeSortItemEdit.ShowMe(Me, 1, lvwMain_S.SelectedItem.Text, 0, "") = True Then
            Call FillItem
        End If
    End If
End Sub

Private Sub mnuEditDelete_Click()
    On Error GoTo ErrHandle
    If sstabItem.Tab = 0 Then
        If lvwMain_S.ListItems.Count = 0 Then Exit Sub
        If Not lvwMain_S.SelectedItem.Selected Then Exit Sub
        If MsgBox("是否删除费别：" & lvwMain_S.SelectedItem.Text, vbQuestion Or vbDefaultButton2 Or vbYesNo, gstrSysName) = vbNo Then Exit Sub
        Err = 0
        On Error Resume Next
        With lvwMain_S.SelectedItem
            gstrSQL = "zl_费别_delete('" & Mid(.Key, 2) & "')"
        End With
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        If Err <> 0 Then
            MsgBox "删除失败，可能该费别已经使用", vbExclamation, gstrSysName
            Err.Clear
            Exit Sub
        End If
        
        Dim intIndex As Integer
        With lvwMain_S
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
            End If
        End With
        
        If sstabItem.Tab = 0 Then
            Call FillDetail
        Else
            Call FillItem
        End If
    Else
        If MsgBox("是否删除收费项目：[" & fgdItem.TextMatrix(fgdItem.Row, 1) & "]的费别设置？", vbQuestion Or vbDefaultButton2 Or vbYesNo, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "zl_费别明细_update('" & lvwMain_S.SelectedItem.Text & "'," & Val(fgdItem.TextMatrix(fgdItem.Row, 0)) & ",Null,0,1,Null)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        Call FillItem
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub mnuEditItem_Click()
'    With mrsDetail
'        .Filter = "收入项目id=" & fgdDetail.TextMatrix(fgdDetail.Row, 0)
'        If .RecordCount <> 0 Then
'            frmChargeSortItemEdit.mstrGrade = .Fields("费别").Value
'            frmChargeSortItemEdit.mlngItemId = .Fields("收入项目id").Value
'            frmChargeSortItemEdit.txtStage.Text = .RecordCount
'            frmChargeSortItemEdit.UdStage.Value = .RecordCount
'            frmChargeSortItemEdit.cboMeasure.ListIndex = Val(.Fields("计算方法").Value)     '调用Click事件设置相关控件
'            frmChargeSortItemEdit.lblItem.Caption = fgdDetail.TextMatrix(fgdDetail.Row, 1) & "分段数："
'            Do While Not .EOF
'                frmChargeSortItemEdit.lblNo(.AbsolutePosition - 1).Visible = True
'                frmChargeSortItemEdit.lblNo(.AbsolutePosition - 1).Caption = .AbsolutePosition
'                frmChargeSortItemEdit.txtMoney(.AbsolutePosition - 1).Visible = True
'                frmChargeSortItemEdit.txtMoney(.AbsolutePosition - 1).Text = Format(.Fields("应收段首值").Value, "###########0.00;-##########0.00;0.00;0.00")
'                frmChargeSortItemEdit.txtTax(.AbsolutePosition - 1).Visible = True
'                frmChargeSortItemEdit.txtTax(.AbsolutePosition - 1).Text = Format(.Fields("实收比率").Value, "###0.000;-##0.000;0.000;0.000")
'                .MoveNext
'            Loop
'            frmChargeSortItemEdit.mblnChange = False
'            frmChargeSortItemEdit.Show 1, Me
'            .Filter = adFilterNone
'            If frmChargeSortItemEdit.mblnOK = True Then Call FillDetail
'        Else
'            .Filter = adFilterNone
'        End If
'
'    End With
    If sstabItem.Tab = 0 Then
        If frmChargeSortItemEdit.ShowMe(Me, 0, lvwMain_S.SelectedItem.Text, Val(fgdDetail.TextMatrix(fgdDetail.Row, 0)), fgdDetail.TextMatrix(fgdDetail.Row, 1)) = True Then
            Call FillDetail
        End If
    Else
        If frmChargeSortItemEdit.ShowMe(Me, 1, lvwMain_S.SelectedItem.Text, Val(fgdItem.TextMatrix(fgdItem.Row, 0)), fgdItem.TextMatrix(fgdItem.Row, 1)) = True Then
            Call FillItem
        End If
    End If
    Call SetItemMenu
End Sub

Private Sub mnuEditModify_Click()
    If sstabItem.Tab = 0 Then
        If lvwMain_S.ListItems.Count = 0 Then Exit Sub
        If Not lvwMain_S.SelectedItem.Selected Then Exit Sub
        
        Call frmChargeSortEdit.编辑费别(lvwMain_S.SelectedItem.Text)
    Else
        If frmChargeSortItemEdit.ShowMe(Me, 1, lvwMain_S.SelectedItem.Text, Val(fgdItem.TextMatrix(fgdItem.Row, 0)), fgdItem.TextMatrix(fgdItem.Row, 1)) = True Then
            Call FillItem
        End If
    End If
End Sub

Private Sub mnuEditUnion_Click()
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    
    If frmChargeSortRate.UnifyPercentage(lvwMain_S.SelectedItem.Text, Val(fgdDetail.TextMatrix(fgdDetail.Row, 3))) = True Then
        Call FillDetail
    End If
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnufilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintset_Click()
    zlPrintSet
End Sub

Private Sub subPrint(ByVal intMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If lvwMain_S.ListItems.Count = 0 Then Exit Sub
    objPrint.Title = IIF(sstabItem.Tab = 0, "费别表(收入项目)", "费别表(收费项目)")
    
    If sstabItem.Tab = 0 Then
        Set objPrint.Body = fgdDetail
    Else
        Set objPrint.Body = fgdItem
    End If
    
    objRow.Add ""
    objRow.Add "费别等级：" & lvwMain_S.SelectedItem.Text & "    "
    objPrint.UnderAppRows.Add objRow
    
    If intMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, intMode
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuhelpTitle_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub


Private Sub mnuReportItem_Click(Index As Integer)
    '默认参数：编码=费别编码
    Dim str编码 As String
        
    If Not lvwMain_S.SelectedItem Is Nothing Then
        str编码 = Mid(lvwMain_S.SelectedItem.Key, 2)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "编码=" & str编码)
End Sub

Private Sub mnuShortMenu2_Click(Index As Integer)
    Select Case Index
        Case 1
            mnuEditAdd_Click
        Case 2
            mnuEditModify_Click
        Case 3
            mnuEditDelete_Click
    End Select
        
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuViewColumn_Click()
    If zlControl.LvwSelectColumns(lvwMain_S, mstrLvw) = True Then
        '列有变化就要重新刷新
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    Me.cbrThis.Visible = mnuViewToolButton.Checked
    SizeControls
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
        Toolbar1.Buttons("View").ButtonMenus(i + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(i + 1).Text, "●", "  ")
    Next
    mnuViewIcon(Index).Checked = True
    Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text, "  ", "●")
    lvwMain_S.View = Index
End Sub

Private Sub mnuViewRefresh_Click()
    Call FillList
End Sub

Private Sub mnuViewStatus_Click()
    Me.mnuViewStatus.Checked = Not Me.mnuViewStatus.Checked
    Me.stbThis.Visible = Me.mnuViewStatus.Checked
    SizeControls
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    If mnuViewToolText.Checked Then
        For intCount = 1 To Me.Toolbar1.Buttons.Count
            Me.Toolbar1.Buttons(intCount).Caption = Me.Toolbar1.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.Toolbar1.Buttons.Count
            Me.Toolbar1.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.Toolbar1.Height
    Me.cbrThis.Refresh
    SizeControls
End Sub

Private Sub picSplitV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
    End If
End Sub

Private Sub picSplitV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplitV.Left + X - msngStartX
        If sngTemp > 1000 And Me.ScaleWidth - (sngTemp + picSplitV.Width) > 1000 Then
            picSplitV.Left = sngTemp
            lvwMain_S.Width = picSplitV.Left - lvwMain_S.Left
            sstabItem.Left = picSplitV.Left + picSplitV.Width
            sstabItem.Width = Me.ScaleWidth - sstabItem.Left
            
            fgdDetail.Top = 480
            fgdDetail.Width = sstabItem.Width - 240
            fgdItem.Top = 480
            fgdItem.Width = sstabItem.Width - 240
        End If
        lvwMain_S.SetFocus
    End If
End Sub


Private Sub SizeControls()
'功能:当改变窗口大小时,对各个控件的位置进行重新排列
    
    Dim sngTop As Single, sngBottom As Single
    
    On Error Resume Next
    
    sngTop = IIF(Me.cbrThis.Visible, Me.cbrThis.Height, 0)
    sngBottom = IIF(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    lvwMain_S.Top = sngTop
    picSplitV.Top = sngTop
    sstabItem.Top = sngTop
    
    fgdDetail.Top = 480
    fgdItem.Top = 480
    
    lvwMain_S.Height = Me.ScaleHeight - sngBottom - lvwMain_S.Top
    picSplitV.Height = Me.ScaleHeight - sngTop - sngBottom
    sstabItem.Height = Me.ScaleHeight - sngTop - sngBottom
    sstabItem.Width = Me.ScaleWidth - sstabItem.Left
    
    fgdDetail.Height = sstabItem.Height - fgdDetail.Top - 100
    fgdItem.Height = sstabItem.Height - fgdItem.Top - 100
    
    lvwMain_S.Left = Me.ScaleLeft
    picSplitV.Left = lvwMain_S.Left + lvwMain_S.Width
    sstabItem.Left = picSplitV.Left + picSplitV.Width
    
    fgdDetail.Left = 80
    fgdItem.Left = 80
    
    fgdDetail.Width = sstabItem.Width - 240
    fgdItem.Width = sstabItem.Width - 240

End Sub

Private Sub sstabItem_Click(PreviousTab As Integer)
    fgdDetail.Visible = True
    fgdItem.Visible = True
    
    If sstabItem.Tab = 0 Then
        fgdItem.Visible = False
        Call FillDetail
    Else
        fgdDetail.Visible = False
        Call FillItem
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    lvwMain_S.View = ButtonMenu.Index - 1
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Preview"
            mnufilePreview_Click
        Case "Print"
            mnuFilePrint_Click
        Case "New"
            mnuEditAdd_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Help"
            mnuhelpTitle_Click
        Case "Exit"
            mnuFileExit_Click
        Case "View"
            mnuViewIcon(lvwMain_S.View).Checked = False
            If lvwMain_S.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwMain_S.View = 0
            Else
                mnuViewIcon(lvwMain_S.View + 1).Checked = True
                lvwMain_S.View = lvwMain_S.View + 1
            End If
    End Select
End Sub

Public Sub FillList()
    Dim rsItem As New ADODB.Recordset
    Dim lst As ListItem
    Dim strKey As String, strIcon As String
    
    '刷新ListView
    On Error GoTo ErrHandle
    If Not lvwMain_S.SelectedItem Is Nothing Then
        '保留原有键值
        strKey = lvwMain_S.SelectedItem.Key
    End If
    With rsItem
        gstrSQL = "select 编码,名称,简码,缺省标志 as 缺省项,说明" & _
                   ",to_char(有效开始,'yyyy-MM-dd') as 有效期开始时间,to_char(有效结束,'yyyy-MM-dd') as 有效期结束时间" & _
                   ",decode(适用科室,2,'指定','全部') as 适用科室,decode(属性,2,'动态性项目','身份唯一项目') as 属性,decode(仅限初诊,1,'是','否') 仅限初诊 from 费别"
        Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

        Dim lngCol  As Long
        Dim varValue As Variant
        lvwMain_S.ListItems.Clear
        Do Until rsItem.EOF
            strIcon = IIF(rsItem("缺省项") = 1, "KeyD", "Key")
            
            Set lst = lvwMain_S.ListItems.Add(, "C" & rsItem("编码"), rsItem("名称"), strIcon, strIcon)
        
            '根据ListView的列名从数据库取数
            For lngCol = 2 To lvwMain_S.ColumnHeaders.Count
                varValue = rsItem(lvwMain_S.ColumnHeaders(lngCol).Text).Value
                lst.SubItems(lngCol - 1) = IIF(IsNull(varValue), "", varValue)
            Next
            lst.Tag = IIF(IsNull(rsItem("说明")), "", rsItem("说明"))
            rsItem.MoveNext
        Loop
        If rsItem.RecordCount > 0 Then
            On Error Resume Next
            Set lst = lvwMain_S.ListItems(strKey)
            If Err <> 0 Then
                Err.Clear
                Set lst = lvwMain_S.ListItems(1)
                lst.Selected = True
                lst.EnsureVisible
            Else
                lst.Selected = True
                lst.EnsureVisible
            End If
        End If
    End With
    
    mstrCharge = lvwMain_S.SelectedItem.Text
    If sstabItem.Tab = 0 Then
        Call FillDetail
    Else
        Call FillItem
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FillDetail()
    Dim i As Integer
    If lvwMain_S.SelectedItem Is Nothing Then
        fgdDetail.Rows = 2
        fgdDetail.RowData(1) = 0
        For i = 0 To fgdDetail.Cols - 1
            fgdDetail.TextMatrix(1, i) = ""
        Next
        Call SetMenu
        Exit Sub
    End If
     
    On Error GoTo ErrHandle
'    gstrSQL = "zl_费别_NEW('" & mstrCharge & "')"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    If lvwMain_S.ListItems.Count > 0 Then
        mstrCharge = lvwMain_S.SelectedItem.Text
    End If
    
    gstrSQL = "select a.费别,a.收入项目ID,b.名称 as 收入项目,a.段号,应收段首值,应收段尾值,实收比率,Decode(计算方法,1,'1-成本价加收比例计算','0-分段比例计算') as 计算方法" & _
            " from 费别明细 A,收入项目 B" & _
            " Where a.收入项目ID = B.id" & _
            "       and 费别=[1] " & _
            " Order by b.编码,应收段首值"
    Set mrsDetail = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrCharge)
        
    With fgdDetail
        .Clear
        .redraw = False
        .Rows = IIF(mrsDetail.RecordCount > 0, mrsDetail.RecordCount + 1, 2)
        .TextMatrix(0, 1) = "收入项目"
        .TextMatrix(0, 2) = "应收金额(元)"
        .TextMatrix(0, 3) = "实收比率(%)"
        .TextMatrix(0, 4) = "计算方法"
        .MergeCol(1) = True
        .MergeCol(2) = False
        .MergeCol(3) = False
        .MergeCol(4) = False
        
        Do While Not mrsDetail.EOF()
            .RowData(mrsDetail.AbsolutePosition) = mrsDetail.Fields("段号").Value
            .TextMatrix(mrsDetail.AbsolutePosition, 0) = mrsDetail.Fields("收入项目ID").Value
            .TextMatrix(mrsDetail.AbsolutePosition, 1) = mrsDetail.Fields("收入项目").Value
            .TextMatrix(mrsDetail.AbsolutePosition, 2) = Format(mrsDetail.Fields("应收段首值").Value, "##########0.00;-#########0.00;0.00;0.00") & _
                    " ～ " & Format(mrsDetail.Fields("应收段尾值").Value, "##########0.00;-#########0.00;0.00;0.00")
            .TextMatrix(mrsDetail.AbsolutePosition, 3) = Format(mrsDetail.Fields("实收比率").Value, "###0.00;-##0.00;0.00;0.00")
            .TextMatrix(mrsDetail.AbsolutePosition, 4) = mrsDetail.Fields("计算方法").Value
            .Row = mrsDetail.AbsolutePosition
            .Col = 3
            .CellBackColor = &H80000005
            .Col = 1
            mrsDetail.MoveNext
        Loop
        .Row = 1
        .redraw = True
    End With
    Call SetMenu
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FillItem()
    Dim i As Integer
    If lvwMain_S.SelectedItem Is Nothing Then
        fgdItem.Rows = 2
        fgdItem.RowData(1) = 0
        For i = 0 To fgdItem.Cols - 1
            fgdItem.TextMatrix(1, i) = ""
        Next
        Call SetMenu
        Exit Sub
    End If
     
    On Error GoTo ErrHandle
    gstrSQL = "select a.费别,a.收费细目id,B.名称 || Decode(B.类别, '5', '(' || B.编码 || ')', '6', '(' || B.编码 || ')', '7', '(' || B.编码 || ')',  '('||C.名称||')') As 收费项目,a.段号,应收段首值,应收段尾值,实收比率,Decode(计算方法,1,'1-成本价加收比例计算','0-分段比例计算') as 计算方法" & _
            " from 费别明细 A,收费项目目录 B, 收费项目类别 C " & _
            " Where a.收费细目id = B.id And B.类别 = C.编码 " & _
            "       and 费别=[1] " & _
            " Order by C.编码, B.编码, A.应收段首值 "
    Set mrsItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrCharge)
        
    With fgdItem
        .Clear
        .redraw = False
        .Rows = IIF(mrsItem.RecordCount > 0, mrsItem.RecordCount + 1, 2)
        .TextMatrix(0, 1) = "收费项目"
        .TextMatrix(0, 2) = "应收金额(元)"
        .TextMatrix(0, 3) = "实收比率(%)"
        .TextMatrix(0, 4) = "计算方法"
        .MergeCol(1) = True
        .MergeCol(2) = False
        .MergeCol(3) = False
        .MergeCol(4) = False
        
        Do While Not mrsItem.EOF()
            .RowData(mrsItem.AbsolutePosition) = mrsItem.Fields("段号").Value
            .TextMatrix(mrsItem.AbsolutePosition, 0) = mrsItem.Fields("收费细目ID").Value
            .TextMatrix(mrsItem.AbsolutePosition, 1) = mrsItem.Fields("收费项目").Value
            .TextMatrix(mrsItem.AbsolutePosition, 2) = Format(mrsItem.Fields("应收段首值").Value, "##########0.00;-#########0.00;0.00;0.00") & _
                    " ～ " & Format(mrsItem.Fields("应收段尾值").Value, "##########0.00;-#########0.00;0.00;0.00")
            .TextMatrix(mrsItem.AbsolutePosition, 3) = Format(mrsItem.Fields("实收比率").Value, "###0.00;-##0.00;0.00;0.00")
            .TextMatrix(mrsItem.AbsolutePosition, 4) = mrsItem.Fields("计算方法").Value
            .Row = mrsItem.AbsolutePosition
            .Col = 3
            .CellBackColor = &H80000005
            .Col = 1
            mrsItem.MoveNext
        Loop
        .Row = 1
        .redraw = True
    End With
    Call SetMenu
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub 权限控制()
'功能:由于有的用户权限不够,故使一些菜单项或按钮不可见
    If InStr(mstrPrivs, "增删改") = 0 Then
        mnuEdit.Visible = False
        mnuEditModify.Visible = False
        mnuShortMenu2(1).Visible = False
        mnuShortMenu2(2).Visible = False
        mnuShortMenu2(3).Visible = False
        mnuShortsplit1.Visible = -False
        Toolbar1.Buttons("Split").Visible = False
        Toolbar1.Buttons("New").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
    End If
End Sub

Private Sub SetMenu()
'功能:设置打印和预鉴按钮的有效值
    Dim blnEnabled As Boolean
    
    mnuEditItem.Visible = True
    mnuEditUnion.Visible = True
    mnuEditSplit.Visible = True
    
    If sstabItem.Tab = 0 Then
        blnEnabled = Not (lvwMain_S.SelectedItem Is Nothing)
        Toolbar1.Buttons("New").Enabled = True
        Toolbar1.Buttons("Modify").Enabled = blnEnabled
        Toolbar1.Buttons("Delete").Enabled = blnEnabled
        mnuEditAdd.Enabled = True
        mnuEditDelete.Enabled = blnEnabled
        mnuEditModify.Enabled = blnEnabled
        
        blnEnabled = lvwMain_S.ListItems.Count > 0
        Toolbar1.Buttons("Print").Enabled = blnEnabled
        Toolbar1.Buttons("Preview").Enabled = blnEnabled
        mnuFilePreview.Enabled = blnEnabled
        mnuFilePrint.Enabled = blnEnabled
        mnuFileExcel.Enabled = blnEnabled
    End If
    
    If sstabItem.Tab = 1 Then
        blnEnabled = (fgdItem.Rows > 1 And fgdItem.TextMatrix(1, 0) <> "")
        Toolbar1.Buttons("New").Enabled = True
        Toolbar1.Buttons("Modify").Enabled = blnEnabled
        Toolbar1.Buttons("Delete").Enabled = blnEnabled
        mnuEditAdd.Enabled = True
        mnuEditDelete.Enabled = blnEnabled
        mnuEditModify.Enabled = blnEnabled
        
        mnuEditItem.Visible = False
        mnuEditUnion.Visible = False
        mnuEditSplit.Visible = False
    End If
    
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

