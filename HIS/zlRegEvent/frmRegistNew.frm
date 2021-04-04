VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmRegistNew 
   AutoRedraw      =   -1  'True
   Caption         =   "门诊挂号管理"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9120
   Icon            =   "frmRegistNew.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9120
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   9000
         _ExtentX        =   15875
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
            NumButtons      =   20
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Description     =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "挂号"
               Key             =   "Add"
               Description     =   "挂号"
               Object.ToolTipText     =   "进入挂号窗口"
               Object.Tag             =   "挂号"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退号"
               Key             =   "Del"
               Description     =   "退号"
               Object.ToolTipText     =   "对当前选中单据退号"
               Object.Tag             =   "退号"
               ImageIndex      =   4
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "DelBook"
                     Object.Tag             =   "退病历费"
                     Text            =   "退病历费"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "DelExtra"
                     Object.Tag             =   "退附加费"
                     Text            =   "退附加费"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Fun_1"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预约"
               Key             =   "预约"
               Description     =   "预约"
               Object.ToolTipText     =   "预约挂号"
               Object.Tag             =   "预约"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "接收"
               Key             =   "接收"
               Description     =   "接收"
               Object.ToolTipText     =   "接收预约"
               Object.Tag             =   "接收"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "取消"
               Key             =   "取消"
               Description     =   "取消"
               Object.ToolTipText     =   "取消预约"
               Object.Tag             =   "取消"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Fun_2"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查阅"
               Key             =   "View"
               Description     =   "查阅"
               Object.ToolTipText     =   "查阅当前单据的内容"
               Object.Tag             =   "查阅"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "重新读满足条件的记录"
               Object.Tag             =   "过滤"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Go"
               Description     =   "定位"
               Object.ToolTipText     =   "定位到当前列表内满足条件的记录上"
               Object.Tag             =   "定位"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "轧帐"
               Key             =   "轧帐"
               Object.ToolTipText     =   "收费轧帐"
               Object.Tag             =   "轧帐"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SplitRollingCurtain"
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "扩展"
               Key             =   "Extra"
               ImageIndex      =   14
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ExtraItem"
                     Object.Tag             =   "功能"
                     Text            =   "功能"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   12
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5295
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRegistNew.frx":014A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11007
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
   Begin MSComctlLib.ImageList imgGray 
      Left            =   7455
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":09DE
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":0BF8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":0E12
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":150C
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":1C06
            Key             =   "预约"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":2300
            Key             =   "接收"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":29FA
            Key             =   "取消"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":30F4
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":37EE
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":3A08
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":3C22
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":3E3C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":4056
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":D9ED
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   6870
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":E167
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":E381
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":E59B
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":EC95
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":F38F
            Key             =   "预约"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":FA89
            Key             =   "接收"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":10183
            Key             =   "取消"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":1087D
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":10F77
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":11191
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":113AB
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":115C5
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":117DF
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":11ED9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsThis 
      Height          =   4455
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   9015
      _cx             =   15901
      _cy             =   7858
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
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
      ExplorerBar     =   7
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
   Begin MSComctlLib.TabStrip tbsType 
      Height          =   720
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   1270
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      TabFixedWidth   =   2290
      TabFixedHeight  =   564
      HotTracking     =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "挂号清单(&1)"
            Key             =   "挂号"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "预约清单(&2)"
            Key             =   "预约"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "预约待接收(&3)"
            Key             =   "接收"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
      Begin VB.Menu mnuFileMoneyEnum 
         Caption         =   "现金点钞(&E)"
         Shortcut        =   {F9}
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
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit_Add 
         Caption         =   "病人挂号(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Del_ 
         Caption         =   "病人退号"
         Begin VB.Menu mnuEdit_Del 
            Caption         =   "病人退号(&D)"
            Shortcut        =   {DEL}
         End
         Begin VB.Menu mnuEdit_DelBook 
            Caption         =   "退病历费(&B)"
         End
         Begin VB.Menu mnuEdit_DelExtra 
            Caption         =   "退附加费(&E)"
         End
      End
      Begin VB.Menu mnuEdit_21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_BindPatNum 
         Caption         =   "绑定门诊号"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_BatchChangeNum 
         Caption         =   "批量换号"
      End
      Begin VB.Menu mnuEdit_22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Bespeak 
         Caption         =   "预约挂号(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEdit_Incept 
         Caption         =   "接收预约(&I)"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuEdit_CancelAuditing 
         Caption         =   "退号审核(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuEdit_Cancel 
         Caption         =   "取消预约(&C)"
      End
      Begin VB.Menu mnuEdit_Clear 
         Caption         =   "清除预约(&R)"
      End
      Begin VB.Menu mnuEdit_Defer 
         Caption         =   "预约延期(&F)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "查阅单据(&V)"
      End
      Begin VB.Menu mnuEdit_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Print 
         Caption         =   "重打票据(&P)"
      End
      Begin VB.Menu mnuEdit_Print_Supplemental 
         Caption         =   "补打票据(&S)"
      End
      Begin VB.Menu mnuEdit_Print_Slip 
         Caption         =   "打印挂号凭条(&I)"
      End
      Begin VB.Menu mnuEdit_Print_Case 
         Caption         =   "打印病历标签(&Q)"
      End
      Begin VB.Menu mnuEdit_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Extra 
         Caption         =   "扩展"
         Begin VB.Menu mnuEdit_ExtraItem 
            Caption         =   "功能"
            Index           =   0
         End
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
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "定位(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefeshOption 
         Caption         =   "刷新方式(&O)"
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后不要刷新数据(&1)"
            Index           =   0
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后提示是否刷新(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后自动刷新数据(&3)"
            Checked         =   -1  'True
            Index           =   2
         End
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewReFlash 
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
Attribute VB_Name = "frmRegistNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private Type Type_SQLCondition
    Default As Boolean          '是否是缺省进入，此时没有条件值,缺省值在mstrFilter中
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    FactB As String
    FactE As String
    DeptID As Long
    Patientid As Long
    Doctor As String
    Operator As String
    FeeType As String   '费别
    ItemType As String '号类
    PatiName As String
End Type
Private SQLCondition As Type_SQLCondition
Private mrsList As ADODB.Recordset  '单据列表
Private mstrVsType As String
Private mstrFilter As String
Private mbytCancel As Byte
Private mstr附加费 As String, mstr附加项目ID As String
Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mstrPrivs As String
Private mlngModul As Long
Private mblnNOMoved As Boolean
Private mstrColWidth As String
Private WithEvents mobjMsgModule As clsMipModule
Attribute mobjMsgModule.VB_VarHelpID = -1

'界面的一个处理流程类型
Private Enum AcitonType
    t_普通
    t_时段
End Enum
'模块参数
Private Type Ty_ModulePara
    lngN天取消预约          As Long    '预约N天内不能取消预约
    bln退号审核             As Boolean '在N天内取消预约 是否需要通过审核
    blnReuseRegNo           As Boolean '已退序号允许挂号
End Type
Private mTy_Para     As Ty_ModulePara
Private mactionType  As AcitonType
'退卡相关处理
Private mstrPassWord As String
Private mcolCardPayMode As Collection
Private mstrPrivs_RollingCurtain As String  '收费轧帐管理权限

Private Sub Form_Activate()
    Call InitLocPar(mlngModul)
    If InStr(mstrPrivs, ";LED与语音;") = 0 Then gblnLED = False
End Sub

Private Sub mnuEdit_BatchChangeNum_Click()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:批量换号功能
    '编制:王吉
    '日期:2011-08-24 10:42:19
    '问题号:45507
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim datNow As Date, strMsgResult As String
    Err.Clear
    datNow = zlDatabase.Currentdate
    If Format(datNow, "yyyy-mm-dd hh:mm:ss") < Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") Then
        strMsgResult = zlCommFun.ShowMsgbox(gstrSysName, "系统将在" & Format(gdatRegistTime, "yyyy-mm-dd") & _
                                            "启用新版出诊表排班模式挂号,请根据预约时间选择模式:" & vbCrLf & _
                                            "计划排班预约(旧):" & Format(datNow, "yyyy-mm-dd hh:mm:ss") & "至" & Format(gdatRegistTime - 1, "yyyy-mm-dd hh:mm:ss") & vbCrLf & _
                                            "出诊表排班预约(新):" & Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") & "及以后", "计划排班预约,出诊表排班预约,取消", Me, vbQuestion)
        If strMsgResult = "" Or strMsgResult = "取消" Then Exit Sub
        If strMsgResult = "计划排班预约" Then
            frmBatchChangeNum.Show 1
        End If
        If strMsgResult = "出诊表排班预约" Then
            frmBatchChangeNumNew.Show 1
        End If
    Else
        frmBatchChangeNumNew.Show 1
    End If
End Sub

Private Sub mnuEdit_Bespeak_Click()
    On Error Resume Next
    Dim datNow As Date, strMsgResult As String
    Err.Clear
    datNow = zlDatabase.Currentdate
    If Format(datNow, "yyyy-mm-dd hh:mm:ss") < Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") Then
        strMsgResult = zlCommFun.ShowMsgbox(gstrSysName, "系统将在" & Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") & _
                                            "启用新版出诊表排班模式挂号,请根据预约时间选择模式:" & vbCrLf & _
                                            "计划排班预约(旧):" & Format(datNow, "yyyy-mm-dd hh:mm:ss") & "至" & Format(gdatRegistTime - 1, "yyyy-mm-dd hh:mm:ss") & vbCrLf & _
                                            "出诊表排班预约(新):" & Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") & "及以后", "计划排班预约,出诊表排班预约,取消", Me, vbQuestion)
        If strMsgResult = "" Or strMsgResult = "取消" Then Exit Sub
        If strMsgResult = "计划排班预约" Then
            If gbln精简界面 Then
                frmRegistEditSimple.mlngModul = mlngModul
                frmRegistEditSimple.mstrPrivs = mstrPrivs
                frmRegistEditSimple.mbytMode = 1
                frmRegistEditSimple.mbytInState = 0
                Set frmRegistEditSimple.mobjMsgModule = mobjMsgModule '消息处理模块
                frmRegistEditSimple.Show IIf(gfrmMain Is Nothing, 0, 1), Me
            Else
                frmRegistEdit.mlngModul = mlngModul
                frmRegistEdit.mstrPrivs = mstrPrivs
                frmRegistEdit.mbytMode = 1
                frmRegistEdit.mbytInState = 0
                Set frmRegistEdit.mobjMsgModule = mobjMsgModule '消息处理模块
                frmRegistEdit.Show IIf(gfrmMain Is Nothing, 0, 1), Me
            End If
        End If
        If strMsgResult = "出诊表排班预约" Then
            frmRegistEditNew.mlngModul = mlngModul
            frmRegistEditNew.mstrPrivs = mstrPrivs
            frmRegistEditNew.mbytMode = 1
            frmRegistEditNew.mbytInState = 0
            Set frmRegistEditNew.mobjMsgModule = mobjMsgModule '消息处理模块
            frmRegistEditNew.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        End If
    Else
        frmRegistEditNew.mlngModul = mlngModul
        frmRegistEditNew.mstrPrivs = mstrPrivs
        frmRegistEditNew.mbytMode = 1
        frmRegistEditNew.mbytInState = 0
        Set frmRegistEditNew.mobjMsgModule = mobjMsgModule '消息处理模块
        frmRegistEditNew.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End If
    If gblnOk And tbsType.SelectedItem.Key <> "挂号" Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
    If Me.Enabled And Me.Visible Then Me.SetFocus
End Sub

Private Sub mnuEdit_BindPatNum_Click()
  ' 打开门诊号 绑定窗体 进行门诊号的绑定
  frmBindPatientNo.Show 1, Me
End Sub

Private Sub mnuEdit_Del_Click()
    Call DeleteRegist
End Sub

Private Sub mnuEdit_DelBook_Click()
    Dim strNO As String
    Dim str挂号时间 As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))

    If strNO = "" Then
        MsgBox "当前没有记录可以退病历费！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If CheckBillExistReplenishData(strNO) Then
        MsgBox "选择的挂号记录进行了医保补充结算，不允许进行退病历费操作！", vbInformation, gstrSysName
        Exit Sub
    End If

    If InStr(1, mstrPrivs, ";强制退号;") = 0 Then
        '判断当前人员对单据是否有操作权限,时间限制,无需检查挂号单有效天数
        If Not BillOperCheck(1, vsThis.TextMatrix(vsThis.Row, getColNum("挂号员")), _
                              CDate(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("挂号时间"))), "退号") Then Exit Sub
    End If

    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 4, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If

    If frmRegistEditNew.CancelBill(Me, strNO, mlngModul, mstrPrivs, False, 1) = False Then Exit Sub
    
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
    If Me.Enabled And Me.Visible Then Me.SetFocus
End Sub

Private Sub LoadPlugInMnu()
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Integer
    Dim blnHave As Boolean, blnTool As Boolean
    Dim strTemp As String
    Dim intToolCounter As Integer
    
    If CreatePlugInOK(mlngModul) Then
        blnHave = True
    End If
    
    mnuEdit_Extra.Visible = blnHave
    tbr.Buttons("Extra").Visible = blnHave
    
    If blnHave Then
        blnTool = False
        On Error Resume Next
        strTmp = gobjPlugIn.GetFuncNames(glngSys, mlngModul, 3)
        Call zlPlugInErrH(Err, "GetFuncNames")
        Err.Clear: On Error GoTo 0
        
        If strTmp = "" Then
            mnuEdit_Extra.Visible = False
            tbr.Buttons("Extra").Visible = False
            Exit Sub
        End If
        strTmp = Replace(strTmp, "Auto:", "")
        arrTmp = Split(strTmp, ",")
        intToolCounter = 0
        For i = 0 To UBound(arrTmp)
            If i <> 0 Then
                Load mnuEdit_ExtraItem(i)
            End If
            mnuEdit_ExtraItem(i).Caption = Replace(CStr(arrTmp(i)), "InTool:", "")
            mnuEdit_ExtraItem(i).Tag = Replace(CStr(arrTmp(i)), "InTool:", "")
            
            If InStr(CStr(arrTmp(i)), "InTool:") > 0 Then
                strTemp = Split(CStr(arrTmp(i)), ":")(1)
                blnTool = True
                If intToolCounter <> 0 Then
                    tbr.Buttons("Extra").ButtonMenus.Add tbr.Buttons("Extra").ButtonMenus.Count + 1, strTemp, strTemp
                    intToolCounter = intToolCounter + 1
                End If
                tbr.Buttons("Extra").ButtonMenus(tbr.Buttons("Extra").ButtonMenus.Count).Text = strTemp
                tbr.Buttons("Extra").ButtonMenus(tbr.Buttons("Extra").ButtonMenus.Count).Tag = strTemp
            End If
        Next
        tbr.Buttons("Extra").Visible = blnTool
    End If
End Sub

Private Sub mnuEdit_ExtraItem_Click(index As Integer)
    Call ExcPlugInFun(mnuEdit_ExtraItem(index).Tag)
End Sub

Private Sub ExcPlugInFun(ByVal strFunName As String)
    Dim lngPatiID As Long
    Dim strNO As String
    
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
    
    If strNO = "" Or strNO = "单据号" Then
        MsgBox "未选中任何单据，不能执行此操作！", vbExclamation, gstrSysName: Exit Sub
    End If
        
    If CreatePlugInOK(mlngModul) Then
        lngPatiID = Val(Me.vsThis.TextMatrix(vsThis.Row, getColNum("病人ID")))
        On Error Resume Next
        Call gobjPlugIn.ExecuteFunc(glngSys, mlngModul, strFunName, lngPatiID, strNO, 0, "", 3)
        Call zlPlugInErrH(Err, "ExecuteFunc")
        Err.Clear: On Error GoTo 0
    End If
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    With ButtonMenu
        Select Case .Key
            Case "DelBook"
                mnuEdit_DelBook_Click
            Case "DelExtra"
                mnuEdit_DelExtra_Click
            Case Else
                Call ExcPlugInFun(.Tag)
        End Select
    End With
End Sub


Private Sub mnuEdit_DelExtra_Click()
    Dim strNO As String, str挂号时间 As String
    '退附加费
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))

    If strNO = "" Then
        MsgBox "当前没有记录可以退" & mstr附加费 & "！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If CheckBillExistReplenishData(strNO) Then
        MsgBox "选择的挂号记录进行了医保补充结算，不允许进行退" & mstr附加费 & "操作！", vbInformation, gstrSysName
        Exit Sub
    End If

    If InStr(1, mstrPrivs, ";强制退号;") = 0 Then
        '判断当前人员对单据是否有操作权限,时间限制,无需检查挂号单有效天数
        If Not BillOperCheck(1, vsThis.TextMatrix(vsThis.Row, getColNum("挂号员")), _
                              CDate(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("挂号时间"))), "退号") Then Exit Sub
    End If

    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 4, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If

    If frmRegistEditNew.CancelBill(Me, strNO, mlngModul, mstrPrivs, False, 2) = False Then Exit Sub
    
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
    If Me.Enabled And Me.Visible Then Me.SetFocus
End Sub

Private Sub CancelOldRegist()
    Dim strSQL As String, strNO As String
    Dim Datsys As Date
    Dim datTmp As Date
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
    If strNO = "" Then
        MsgBox "当前没有挂号预约可以取消。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If Not BillOperCheck(1, vsThis.TextMatrix(vsThis.Row, getColNum("挂号员")), _
        CDate(vsThis.TextMatrix(vsThis.Row, getColNum("登记时间"))), "取消预约") Then Exit Sub
    If mbytCancel <> 1 Then
        If vsThis.TextMatrix(vsThis.Row, getColNum("记录状态")) <> "1" Then
            MsgBox "当前挂号预约已经取消。", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    If CheckRegistAppointment(strNO) = False Then
        MsgBox "当前挂号预约已经接收。", vbExclamation, gstrSysName
        Exit Sub
    End If
    If tbsType.SelectedItem.Key <> "挂号" And mTy_Para.bln退号审核 And mTy_Para.lngN天取消预约 > 0 Then
        '退号审核 限制收费预约和预约接收的退号
        If vsThis.TextMatrix(vsThis.Row, getColNum("预约时间")) <> "" Then
            If vsThis.TextMatrix(vsThis.Row, getColNum("退号审核人")) = "" Then
                '是否预约判断放到里面 外面影响性能
                 Datsys = zlDatabase.Currentdate
                 datTmp = DateAdd("d", -1 * mTy_Para.lngN天取消预约, CDate(vsThis.TextMatrix(vsThis.Row, getColNum("预约时间"))))
                   '预约时间-K >datSys
                   If Format(Datsys, "yyyy-MM-dd hh:mm:ss") > Format(datTmp, "yyyy-MM-dd hh:mm:ss") Then
                           MsgBox "单据号为" & strNO & "的收费预约单据没有经过退号审核!不能进行退号!", vbInformation, Me.Caption
                           Exit Sub
                   End If
            End If
        End If
        
    End If
    If gbln精简界面 Then
        frmRegistEditSimple.mlngModul = mlngModul
        frmRegistEditSimple.mstrPrivs = mstrPrivs
        frmRegistEditSimple.mstrNoIn = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
        frmRegistEditSimple.mblnNOMoved = mblnNOMoved
        frmRegistEditSimple.mbytMode = 3
        frmRegistEditSimple.mbytInState = 1
        Set frmRegistEditSimple.mobjMsgModule = mobjMsgModule '消息处理模块
        frmRegistEditSimple.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else
        frmRegistEdit.mlngModul = mlngModul
        frmRegistEdit.mstrPrivs = mstrPrivs
        frmRegistEdit.mstrNoIn = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
        frmRegistEdit.mblnNOMoved = mblnNOMoved
        frmRegistEdit.mbytMode = 3
        frmRegistEdit.mbytInState = 1
        Set frmRegistEdit.mobjMsgModule = mobjMsgModule '消息处理模块
        frmRegistEdit.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End If
    If mnuViewRefeshOptionItem(1).Checked And gblnOk Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked And gblnOk Then
        mnuViewReFlash_Click
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub mnuEdit_Cancel_Click()
    Dim strSQL As String, strNO As String
    Dim Datsys As Date
    Dim datTmp As Date
    
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
    If strNO = "" Then
        MsgBox "当前没有挂号预约可以取消。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If IsNewModeRegist(strNO) = False Then
        Call CancelOldRegist
        Exit Sub
    End If
    
    If Not BillOperCheck(1, vsThis.TextMatrix(vsThis.Row, getColNum("挂号员")), _
        CDate(vsThis.TextMatrix(vsThis.Row, getColNum("登记时间"))), "取消预约") Then Exit Sub
    If mbytCancel <> 1 Then
        If vsThis.TextMatrix(vsThis.Row, getColNum("记录状态")) <> "1" Then
            MsgBox "当前挂号预约已经取消。", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    If CheckRegistAppointment(strNO) = False Then
        MsgBox "当前挂号预约已经接收。", vbExclamation, gstrSysName
        Exit Sub
    End If
    If tbsType.SelectedItem.Key <> "挂号" And mTy_Para.bln退号审核 And mTy_Para.lngN天取消预约 > 0 Then
        '退号审核 限制收费预约和预约接收的退号
        If vsThis.TextMatrix(vsThis.Row, getColNum("预约时间")) <> "" Then
            If vsThis.TextMatrix(vsThis.Row, getColNum("退号审核人")) = "" Then
                '是否预约判断放到里面 外面影响性能
                 Datsys = zlDatabase.Currentdate
                 datTmp = DateAdd("d", -1 * mTy_Para.lngN天取消预约, CDate(vsThis.TextMatrix(vsThis.Row, getColNum("预约时间"))))
                   '预约时间-K >datSys
                   If Format(Datsys, "yyyy-MM-dd hh:mm:ss") > Format(datTmp, "yyyy-MM-dd hh:mm:ss") Then
                           MsgBox "单据号为" & strNO & "的收费预约单据没有经过退号审核!不能进行退号!", vbInformation, Me.Caption
                           Exit Sub
                   End If
            End If
        End If
        
    End If
    
    frmRegistEditNew.mlngModul = mlngModul
    frmRegistEditNew.mstrPrivs = mstrPrivs
    frmRegistEditNew.mstrNoIn = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
    frmRegistEditNew.mblnNOMoved = mblnNOMoved
    frmRegistEditNew.mbytMode = 3
    frmRegistEditNew.mbytInState = 1
    Set frmRegistEditNew.mobjMsgModule = mobjMsgModule '消息处理模块
    frmRegistEditNew.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If mnuViewRefeshOptionItem(1).Checked And gblnOk Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked And gblnOk Then
        mnuViewReFlash_Click
    End If
    If Me.Enabled And Me.Visible Then Me.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

 

Private Sub mnuEdit_CancelAuditing_Click()
    '退号审核
    Dim strSQL As String, strNO As String
    If vsThis.Rows <= 1 Then Exit Sub
    If InStr(1, mstrPrivs, ";退号审核;") = 0 Then
        MsgBox "你没有对预约号进行退号审核的权限。", vbExclamation, gstrSysName
        Exit Sub
    End If
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
    If strNO = "" Then
        MsgBox "当前没有挂号预约可以进行退号审核。", vbExclamation, gstrSysName
        Exit Sub
    End If

    If mbytCancel <> 1 Then
        If vsThis.TextMatrix(vsThis.Row, getColNum("记录状态")) <> "1" Then
            MsgBox "当前挂号预约已经取消。", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    Select Case tbsType.SelectedItem.Key
    Case "预约", "接收":
        If Not BillOperCheck(1, vsThis.TextMatrix(vsThis.Row, getColNum("挂号员")), _
            CDate(vsThis.TextMatrix(vsThis.Row, getColNum("登记时间"))), "取消预约") Then Exit Sub
        
    Case "挂号":
        If vsThis.TextMatrix(vsThis.Row, getColNum("预约时间")) = "" Then Exit Sub
        If vsThis.TextMatrix(vsThis.Row, getColNum("退号审核人")) <> "" Then Exit Sub
    Case Else:
        Exit Sub
    End Select
    
   
     If MsgBox("确实要将单据[" & strNO & "]进行取消退号审核吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
           '  Zl_病人预约挂号_Cancelauditing
   strSQL = "Zl_病人预约挂号_Cancelauditing("
           '  No_In       病人挂号记录.NO%Type,
   strSQL = strSQL & "'" & strNO & "',"
           '  操作员_In   病人挂号记录.退号审核人%Type,
   strSQL = strSQL & "'" & UserInfo.姓名 & "',"
           '  审核时间_In 病人挂号记录.退号审核时间%Type
    strSQL = strSQL & "Sysdate)"
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    mnuViewReFlash_Click
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Clear_Click()
    If MsgBox("该操作将清除最近 " & gint预约天数 & " 天内登记，但预约时间已过期的预约记录，要继续吗？" & _
        vbCrLf & vbCrLf & "说明：为保证有效清除这些过期的预约记录，你需要定期执行该功能。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure("zl_病人预约挂号_Clear", Me.Caption)
    On Error GoTo 0
    
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("清除操作已执行完毕。清单内容可能已更改，要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        MsgBox "清除操作已执行完毕。", vbInformation, gstrSysName
        mnuViewReFlash_Click
    Else
        MsgBox "清除操作已执行完毕。", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Defer_Click()
    Dim datNow As Date, strMsgResult As String
    datNow = zlDatabase.Currentdate
    If Format(datNow, "yyyy-mm-dd hh:mm:ss") < Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") Then
        strMsgResult = zlCommFun.ShowMsgbox(gstrSysName, "系统将在" & Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") & _
                                            "启用新版出诊表排班模式挂号,请根据预约时间选择模式:" & vbCrLf & _
                                            "计划排班预约(旧):" & Format(datNow, "yyyy-mm-dd hh:mm:ss") & "至" & Format(gdatRegistTime - 1, "yyyy-mm-dd hh:mm:ss") & vbCrLf & _
                                            "出诊表排班预约(新):" & Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") & "及以后", "计划排班预约,出诊表排班预约,取消", Me, vbQuestion)
        If strMsgResult = "" Or strMsgResult = "取消" Then Exit Sub
        If strMsgResult = "计划排班预约" Then
            frmBookingDefer.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        End If
        If strMsgResult = "出诊表排班预约" Then
            frmBookingDeferNew.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        End If
    Else
        frmBookingDeferNew.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End If
End Sub

Private Sub InceptOldRegist()
    Dim strNO As String
    Dim datTime As Date
    Dim datThis As Date
    
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
    If strNO = "" Then
        MsgBox "当前没有接收的预约挂号。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If CheckRegistAppointment(strNO) = False Then
        MsgBox "当前单据已经被接收。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    datTime = CDate(vsThis.TextMatrix(vsThis.Row, getColNum("预约时间")))
    datThis = zlDatabase.Currentdate
    If Format(datTime, "YYYY-MM-DD") > Format(datThis, "YYYY-MM-DD") Then
        If MsgBox("当前接收的记录不是当天的预约记录，是否继续接收？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    If gbln精简界面 Then
        frmRegistEditSimple.mlngModul = mlngModul
        frmRegistEditSimple.mstrPrivs = mstrPrivs
        frmRegistEditSimple.mbytMode = 2
        frmRegistEditSimple.mbytInState = 0
        frmRegistEditSimple.mstrNoIn = strNO
        Set frmRegistEditSimple.mobjMsgModule = mobjMsgModule '消息处理模块
        frmRegistEditSimple.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else
        frmRegistEdit.mlngModul = mlngModul
        frmRegistEdit.mstrPrivs = mstrPrivs
        frmRegistEdit.mbytMode = 2
        frmRegistEdit.mbytInState = 0
        frmRegistEdit.mstrNoIn = strNO
        Set frmRegistEdit.mobjMsgModule = mobjMsgModule '消息处理模块
        frmRegistEdit.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End If
    If gblnOk And tbsType.SelectedItem.Key <> "挂号" Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEdit_Incept_Click()
    Dim strNO As String
    Dim datTime As Date
    Dim datThis As Date
    
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
    If strNO = "" Then
        MsgBox "当前没有接收的预约挂号。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If IsNewModeRegist(strNO) = False Then
        Call InceptOldRegist
        Exit Sub
    End If
    
    If CheckRegistAppointment(strNO) = False Then
        MsgBox "当前单据已经被接收。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    datTime = CDate(vsThis.TextMatrix(vsThis.Row, getColNum("预约时间")))
    datThis = zlDatabase.Currentdate
    If Format(datTime, "YYYY-MM-DD") > Format(datThis, "YYYY-MM-DD") Then
        If MsgBox("当前接收的记录不是当天的预约记录，是否继续接收？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    frmRegistEditNew.mlngModul = mlngModul
    frmRegistEditNew.mstrPrivs = mstrPrivs
    frmRegistEditNew.mbytMode = 2
    frmRegistEditNew.mbytInState = 0
    frmRegistEditNew.mstrNoIn = strNO
    Set frmRegistEditNew.mobjMsgModule = mobjMsgModule '消息处理模块
    frmRegistEditNew.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOk And tbsType.SelectedItem.Key <> "挂号" Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
    If Me.Enabled And Me.Visible Then Me.SetFocus
End Sub
 

Private Sub mnuEdit_Print_Case_Click()
    Dim lng病人ID As Long
    lng病人ID = Val(Me.vsThis.TextMatrix(vsThis.Row, getColNum("病人ID")))
    If lng病人ID <> 0 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_2", Me, "病人ID=" & lng病人ID, 2)
    Else
        MsgBox "该挂号单相关的病人没有建立病人档案!", vbInformation
    End If
End Sub

Private Sub mnuEdit_Print_Slip_Click()
    Dim strNO As String
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
    If strNO <> "" Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me, "NO=" & strNO, 2)
        '72704:谢荣,2014-07-23,写入凭条打印记录
        gstrSQL = "Zl_凭条打印记录_Update(4,'" & strNO & "',1,'" & UserInfo.姓名 & "')"
        zlDatabase.ExecuteProcedure gstrSQL, ""
    Else
        MsgBox "当前没有挂号或接收记录！", vbExclamation, gstrSysName
    End If
End Sub

Private Sub mnuFileInsure_Click()
    gclsInsure.InsureSupport
End Sub

Private Sub mnuFileLocalSet_Click()
    frmLocalPara.mlngModul = mlngModul
    frmLocalPara.mstrPrivs = mstrPrivs
    frmLocalPara.Show 1, Me
    If gblnOk Then InitPara
End Sub

Private Sub mnuFileMoneyEnum_Click()
    Call frmMoneyEnum.ShowMe(Me)
End Sub
 

Private Sub mnuFileRollingCurtain_Click()
    Call zlExecuteChargeRollingCurtain(Me)
End Sub

Private Sub mnuReportItem_Click(index As Integer)
    Dim strNO As String
    
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
    If strNO <> "" Then
        With vsThis
            Call ReportOpen(gcnOracle, Split(mnuReportItem(index).Tag, ",")(0), Split(mnuReportItem(index).Tag, ",")(1), Me, _
                "NO=" & strNO, "票据号=" & .TextMatrix(.Row, getColNum("首张票据")), _
                "号别=" & .TextMatrix(.Row, getColNum("号别")), "医生=" & .TextMatrix(.Row, getColNum("医生")), _
                "门诊号=" & .TextMatrix(.Row, getColNum("门诊号")))
        End With
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(index).Tag, ",")(0), Split(mnuReportItem(index).Tag, ",")(1), Me)
    End If
End Sub

Private Sub mnuViewFilter_Click()
    '定位病人范围
    frmRegistFilter.mlngModule = mlngModul
    frmRegistFilter.bytType = tbsType.SelectedItem.index - 1
    
    '列表显示方式,按照发生时间显示 还是按照登记时间显示？
    'frmRegistFilter.mblnFilterType = mTy_Para.bln挂号列表过滤
    frmRegistFilter.Show 1, Me
    If gblnOk Then
        mstrFilter = frmRegistFilter.mstrFilter
        mbytCancel = IIf(frmRegistFilter.optRegistRecord(0).Value = True, 1, IIf(frmRegistFilter.optRegistRecord(1).Value = True, 2, 3))
        
        With SQLCondition
            .DateB = frmRegistFilter.dtpBegin.Value
            .DateE = frmRegistFilter.dtpEnd.Value
            .NOB = frmRegistFilter.txtNOBegin.Text
            .NOE = frmRegistFilter.txtNOEnd.Text
            .FactB = frmRegistFilter.txtFactBegin.Text
            .FactE = frmRegistFilter.txtFactEnd.Text
            If frmRegistFilter.cbo科室.ListIndex > 0 Then .DeptID = frmRegistFilter.cbo科室.ItemData(frmRegistFilter.cbo科室.ListIndex)
            .PatiName = gstrLike & frmRegistFilter.txtPatient.Text & "%"
            .Patientid = frmRegistFilter.mlngPrePatient
            .Doctor = frmRegistFilter.txt医生.Text & "%"
            If frmRegistFilter.cbo操作员.ListIndex > 0 Then .Operator = NeedName(frmRegistFilter.cbo操作员.Text)
            If frmRegistFilter.cbo费别.ListIndex > 0 Then .FeeType = NeedName(frmRegistFilter.cbo费别.Text)
            If frmRegistFilter.cbo号类.ListIndex > 0 Then .ItemType = frmRegistFilter.cbo号类.Text
        End With
        
        mnuViewReFlash_Click
    End If
End Sub

Private Sub vsThis_AfterMoveColumn(ByVal Col As Long, Position As Long)
     zl_vsGrid_Para_Save mlngModul, vsThis, Me.Caption, Me.tbsType.SelectedItem.Key, False, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub vsThis_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngWidth As Long
    Dim i As Integer
    If mstrColWidth <> "" Then
        lngWidth = vsThis.ColWidth(Col)
        For i = 0 To UBound(Split(mstrColWidth, "|"))
            vsThis.ColWidth(i) = Split(mstrColWidth, "|")(i)
        Next i
        vsThis.ColWidth(Col) = lngWidth
    End If
    zl_vsGrid_Para_Save mlngModul, vsThis, Me.Caption, Me.tbsType.SelectedItem.Key, False, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub vsThis_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Integer
    mstrColWidth = ""
    For i = 0 To vsThis.Cols - 1
        mstrColWidth = mstrColWidth & "|" & vsThis.ColWidth(i)
    Next i
    If mstrColWidth <> "" Then mstrColWidth = Mid(mstrColWidth, 2)
End Sub

Private Sub vsThis_DblClick()
    Dim lngCols As Long
    Dim lngRow As Long
    If vsThis.MouseRow <= 0 Then Exit Sub
    If vsThis.Row <= 0 Then Exit Sub
     lngCols = getColNum("记录状态")
     lngRow = vsThis.Row
    
    If mnuEdit_View.Enabled Then mnuEdit_View_Click
End Sub

Private Sub vsThis_EnterCell()
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strNO As String
    If vsThis.MouseRow <= 0 Then Exit Sub
    If Mid(stbThis.Panels(2).Text, 1, 2) = "摘要" Then stbThis.Panels(2).Text = ""
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
    If vsThis.Row = 0 Or strNO = "" Then Exit Sub
    
    mlngGo = vsThis.Row
    mlngCurRow = vsThis.Row: mlngTopRow = vsThis.TopRow
    
    If tbsType.SelectedItem.Key <> "挂号" Then
        stbThis.Panels(2).Text = "摘要:" & vsThis.TextMatrix(vsThis.Row, getColNum("摘要"))
    End If
    Call SetMenuEnable
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsThis_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mnuEdit_Del.Enabled And mnuEdit_Del.Visible Then
        Call mnuEdit_Del_Click
        Exit Sub
    End If
    If (tbsType.SelectedItem.Key = "预约" Or tbsType.SelectedItem.Key = "接收") And KeyCode = vbKeyDelete And mnuEdit_Cancel.Enabled = True And mnuEdit_Cancel.Visible Then Call mnuEdit_Cancel_Click
End Sub

Private Sub vsThis_RowColChange()
    Call SetMenuEnable
End Sub
Private Sub SetMenuEnable()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置菜单的Enable属性
    '编制:刘兴洪
    '日期:2013-11-05 16:03:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng记录状态 As Long, lng病人ID As Long
    Dim bln记帐 As Boolean, bln未退完 As Boolean
    Dim blnEnabled As Boolean, bln附加 As Boolean, bln病历 As Boolean
    Dim strStatus As String, rsStatus As ADODB.Recordset
    
    With vsThis
      If .Row <= 0 Then
            blnEnabled = False
            mnuEdit_Cancel.Enabled = blnEnabled
            tbr.Buttons("取消").Enabled = blnEnabled
            mnuEdit_Del.Enabled = blnEnabled
            mnuEdit_DelBook.Enabled = blnEnabled
            mnuEdit_DelExtra.Enabled = blnEnabled
            tbr.Buttons("Del").Enabled = blnEnabled
             mnuEdit_CancelAuditing.Enabled = blnEnabled
             Exit Sub
      End If
      
      lng记录状态 = 0
      If getColNum("记录状态") <> -1 Then
            lng记录状态 = Val(.TextMatrix(.Row, getColNum("记录状态")))
      End If
      
      If getColNum("记录状态") <> -1 Then
        lng病人ID = Val(.TextMatrix(.Row, getColNum("病人ID")))
      End If
      
      If getColNum("单据号") <> -1 Then
            strStatus = "Select Sum(附加) As 附加, Sum(病历) As 病历, Sum(未退完) As 未退完" & vbNewLine & _
                        "From (Select Decode(Nvl(Max(a.Id), 0), 0, 0, 1) As 附加, 0 As 病历, 0 As 未退完" & vbNewLine & _
                        "       From 门诊费用记录 A" & vbNewLine & _
                        "       Where a.No(+) = [1] And a.记录性质(+) = 4 And a.记录状态(+) = 1 And" & vbNewLine & _
                        "             Instr(',' || [2] || ',' , ',' || a.收费细目id(+) || ',') > 0" & vbNewLine & _
                        "       Union All" & vbNewLine & _
                        "       Select 0 As 附加, Decode(Nvl(Max(b.Id), 0), 0, 0, 1) As 病历, 0 As 未退完" & vbNewLine & _
                        "       From 门诊费用记录 B" & vbNewLine & _
                        "       Where b.No(+) = [1] And b.记录性质(+) = 4 And b.记录状态(+) = 1 And b.附加标志(+) = 1" & vbNewLine & _
                        "       Union All" & vbNewLine & _
                        "       Select 0 As 附加, 0 As 病历, Decode(Nvl(Max(c.Id), 0), 0, 0, 1) As 未退完" & vbNewLine & _
                        "       From 门诊费用记录 C" & vbNewLine & _
                        "       Where c.No(+) = [1] And c.记录性质(+) = 4 And c.记录状态(+) = 1)"
            Set rsStatus = zlDatabase.OpenSQLRecord(strStatus, Me.Caption, .TextMatrix(.Row, getColNum("单据号")), mstr附加项目ID)
      End If
      
      mnuEdit_CancelAuditing.Enabled = lng记录状态 = 1 And .TextMatrix(.Row, getColNum("退号审核人")) = ""
      
      Select Case tbsType.SelectedItem.Key
      Case "挂号"
            If getColNum("记帐费用") <> -1 Then
                bln记帐 = Val(.TextMatrix(.Row, getColNum("记帐费用"))) = 1
            Else
                bln记帐 = False
            End If
            
            If Not rsStatus.EOF Then
                bln未退完 = Val(Nvl(rsStatus!未退完)) = 1
                bln附加 = Val(Nvl(rsStatus!附加)) = 1
                bln病历 = Val(Nvl(rsStatus!病历)) = 1
            Else
                bln未退完 = False
                bln附加 = False
                bln病历 = False
            End If
            
            '补打和重打
            mnuEdit_Print.Enabled = bln未退完 And Not bln记帐
            blnEnabled = Trim(vsThis.TextMatrix(vsThis.Row, getColNum("首张票据"))) = ""
            mnuEdit_Print_Supplemental.Enabled = bln未退完 And Not bln记帐 And blnEnabled
            
            mnuEdit_Cancel.Enabled = False
            tbr.Buttons("取消").Enabled = False
            
            blnEnabled = lng记录状态 = 1
            mnuEdit_Del.Enabled = blnEnabled
            tbr.Buttons("Del").Enabled = blnEnabled
            
            blnEnabled = mnuEdit_CancelAuditing.Enabled And .TextMatrix(.Row, getColNum("预约时间")) <> ""
            mnuEdit_CancelAuditing.Enabled = blnEnabled
            
            If lng记录状态 <> 2 Then
                mnuEdit_DelBook.Enabled = .TextMatrix(.Row, getColNum("病历")) <> ""
                tbr.Buttons("Del").ButtonMenus("DelBook").Enabled = .TextMatrix(.Row, getColNum("病历")) <> ""
            Else
                mnuEdit_DelBook.Enabled = bln病历
                tbr.Buttons("Del").ButtonMenus("DelBook").Enabled = bln病历
            End If
            
            mnuEdit_DelExtra.Enabled = bln附加
            tbr.Buttons("Del").ButtonMenus("DelExtra").Enabled = bln附加
      Case Else
            mnuEdit_Cancel.Enabled = lng记录状态 = 1
            tbr.Buttons("取消").Enabled = mnuEdit_Cancel.Enabled
            mnuEdit_Del.Enabled = False
            tbr.Buttons("Del").Enabled = False
            
            mnuEdit_Print.Enabled = False
            mnuEdit_Print_Supplemental.Enabled = False
      End Select
   End With
 End Sub

Private Sub vsThis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuEdit, 2
    Else
        Call vsThis_EnterCell
    End If
End Sub

Private Function getColNum(strColName As String) As Long
    Dim i As Long
    For i = 0 To vsThis.Cols - 1
        If vsThis.TextMatrix(0, i) = strColName Then
            getColNum = i
            Exit Function
        End If
    Next
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            '始终从当前行开始
            If mnuViewGo.Enabled Then Call SeekBill(False)
        Case vbKeyReturn
            If mnuEdit_View.Enabled Then mnuEdit_View_Click
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub
Private Function IsCheckCancelValied(ByVal lng挂号结帐ID As Long, ByVal lng卡费结帐ID As Long, _
    ByVal cllBillBalance As Collection, ByVal dbl金额 As Double, Optional ByVal strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费时的数据有效性
    '返回:有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-12 18:43:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strName As String, bln消费卡 As Boolean, lng卡类别ID As Long
    Dim str验证卡号  As String, strSQL As String, rsTmp As ADODB.Recordset
    Dim strXmlIn As String, bln退款验卡 As Boolean, str刷卡密码 As String
    strName = IIf(glngSys \ 100 = 8, "会员卡", "医疗卡")
    If cllBillBalance Is Nothing Then IsCheckCancelValied = True: Exit Function
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO,结帐ID
    bln消费卡 = Val(cllBillBalance(1)(2)) = 1
    lng卡类别ID = cllBillBalance(1)(0)
    If lng卡类别ID = 0 Then IsCheckCancelValied = True: Exit Function
    '4.3.3.2.6   zlReturnCheck:帐户回退交易前的检查
    'zlPaymentCheck帐户扣款交易检查
    '参数名  参数类型    入/出   备注
    'frmMain Object  In  调用的主窗体
    'lngModule   Long    In  模块号
    'lngCardTypeID   Long    In  卡类别ID:医疗卡类别.ID
    'strCardNo   String  IN  卡号
    'strBalanceIDs:格式:收费类型( 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款)|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    'dblMoney    Double  IN  退款金额
    'strSwapNo   String  In  交易流水号(退款时检查)
    'strSwapMemo String  In  交易说明(退款时传入)
    '    Boolean 函数返回    True:调用成功,False:调用失败
    '说明:
    '在调用扣款前，由于存在Oracle事务问题，因此，再调用回退交易前，先进行数据的合法性检查,以便控制死锁情况。
    
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO,结帐ID
    'mcolBillBalance.Add Array(Val(Nvl(rsTmp!卡类别ID)), Trim(Nvl(rsTmp!卡号)), IIf(Val(Nvl(rsTmp!结算卡序号)) <> 0, 1, 0), Trim(Nvl(rsTmp!交易流水号)), Trim(Nvl(rsTmp!交易说明))), strNO
    Dim str卡号 As String, str交易流水号 As String, str交易说明 As String, str结算信息 As String
    Dim strXMLExpend As String
    str卡号 = cllBillBalance(1)(1)
    str交易流水号 = cllBillBalance(1)(3)
    str交易说明 = cllBillBalance(1)(4)
    If lng卡费结帐ID <> 0 Then str结算信息 = str结算信息 & "||5|" & lng卡费结帐ID
    If lng挂号结帐ID <> 0 Then str结算信息 = str结算信息 & "||4|" & lng挂号结帐ID
    If str结算信息 <> "" Then str结算信息 = Mid(str结算信息, 3)
    
    If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, lng卡类别ID, bln消费卡, str卡号, str结算信息, dbl金额, str交易流水号, str交易说明, strXMLExpend) = False Then
        Exit Function
    End If
    
    strSQL = "Select 是否退款验卡 From 医疗卡类别 Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng卡类别ID)
    If rsTmp.EOF Then
        bln退款验卡 = False
    Else
        bln退款验卡 = Val(Nvl(rsTmp!是否退款验卡)) = 1
    End If
    
    strSQL = "Select 姓名,性别,年龄 From 病人挂号记录 Where NO = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    
    If bln退款验卡 Then
        strXmlIn = "<IN><CZLX>2</CZLX></IN>"
        If rsTmp.EOF Then
            If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, lng卡类别ID, bln消费卡, _
                "", "", "", dbl金额, str卡号, str刷卡密码, _
                False, True, False, True, Nothing, False, True, strXmlIn) = False Then Exit Function
        Else
            If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, lng卡类别ID, bln消费卡, _
                Nvl(rsTmp!姓名), Nvl(rsTmp!性别), Nvl(rsTmp!年龄), dbl金额, str卡号, str刷卡密码, _
                False, True, False, True, Nothing, False, True, strXmlIn) = False Then Exit Function
        End If
    End If
    
    IsCheckCancelValied = True
End Function

Private Function IsCheckCancel退预交(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取消卡绑定时时检查病人是否有预交款未退
    '返回:有效,返回true,否则返回False
    '编制:刘尔旋
    '日期:2014-04-24
    '问题号:62568
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsBill As Recordset, rsCard As Recordset
    
    strSQL = "Select Count(1) As 医疗卡数 From 病人医疗卡信息 Where 状态=0 And 病人ID=[1]"
    Set rsCard = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    
    strSQL = _
            "Select 预交余额,费用余额 From 病人余额 Where 性质=1 And 类型=1 And 病人ID=[1]"
    Set rsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    
    If Format(Nvl(rsBill!预交余额, 0) - Nvl(rsBill!费用余额, 0), "0.00") > 0 Then
        If Val(Nvl(rsCard!医疗卡数)) = 1 Then
            MsgBox "该病人尚有预交余额，请去医疗卡发放管理界面对该卡进行取消绑定操作!", vbInformation, gstrSysName
            IsCheckCancel退预交 = False
            Exit Function
        End If
    End If
    IsCheckCancel退预交 = True
End Function

Private Function CheckRegistAppointment(ByVal strNO As String) As Boolean
    '检查预约记录是否被接收
    'True-预约记录未接收;False-预约记录已被接收
    Dim strSQL As String, rsTmp As ADODB.Recordset
    strSQL = "Select 1 From 病人挂号记录 Where NO = [1] And 接收时间 Is Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If Not rsTmp.EOF Then
        CheckRegistAppointment = True
    Else
        CheckRegistAppointment = False
    End If
End Function

Private Sub DelOldRegist()
    Dim strSQL As String, strNO As String, str划价NO As String, strCardNo As String
    Dim intInsure As Integer, lng结帐ID As Long, lngCard结帐ID As Long, msgBoxResult As String
    Dim str挂号时间 As String, strSQLCard As String, strMessage As String
    Dim str门诊号 As String, strAdvance As String, str个人帐户 As String
    Dim blnEnableDel As Boolean, blnTrans As Boolean, strSQLBound As String
    Dim bytTogetherDo As Byte, bln退费重打 As Boolean, blnPromptClear As Boolean
    Dim rsTmp As ADODB.Recordset, rsOneCard1 As ADODB.Recordset, rsOneCard2 As ADODB.Recordset
    Dim objICCard As Object
    Dim cllPro As Collection, cllBillBalance As Collection, dblThreeMoney As Double
    Dim cllUpdate As Collection, cllThreeIns As Collection, strErrMsg As String
    Dim byt退费方式 As Byte  '0-全退 1-只退挂号费 2-只退病历
    Dim bln病历费 As Boolean    '是否包含病历费
    Dim blnCardReprint As Boolean    '不退卡重打
    Dim Datsys As Date
    Dim datTmp As Date
    Dim lngPatientID As Long
    Dim dblAdvanceMoney As Double    '预交结算费用
    Dim strInvoice As String, lng病人ID As Long, lng领用ID As Long
    Dim blnVirtualPrint As Boolean
    Dim bln记帐 As Boolean, int险类 As Integer, bln结帐 As Boolean
    
    Set cllPro = New Collection
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))

    If strNO = "" Then
        MsgBox "当前没有记录可以退号！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If CheckBillExistReplenishData(strNO) Then
        MsgBox "选择的挂号记录进行了医保补充结算，不允许进行退号操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errH
    str挂号时间 = vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("挂号时间"))
    str门诊号 = vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("门诊号"))
    bln病历费 = Trim(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("病历"))) <> ""
    lngPatientID = Val(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("病人ID")))
    
    int险类 = Val(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("险类")))
    bln记帐 = Val(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("记帐费用"))) = 1
    
    
    If InStr(1, mstrPrivs, ";强制退号;") = 0 Then
        '判断当前人员对单据是否有操作权限,时间限制,无需检查挂号单有效天数
        If Not BillOperCheck(1, vsThis.TextMatrix(vsThis.Row, getColNum("挂号员")), _
                             CDate(str挂号时间), "退号") Then Exit Sub
    End If

    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 4, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If

    If InStr(1, mstrPrivs, ";强制退号;") = 0 Then   '问题:
        '检查挂号单是否已执行
        If InStr(";" & mstrPrivs & ";", ";下医嘱后退号;") > 0 Then
            blnEnableDel = True
        End If
        If CheckExecuted(strNO, blnEnableDel) Then
            MsgBox "挂号单" & strNO & "已经被医生接诊或下过医嘱,不能退号！", vbInformation, gstrSysName
            Exit Sub
        End If
        '医生站挂的号-收费判断
        If CheckPriceHaveFee(strNO, str划价NO) Then Exit Sub
        '是否发生过费用,但未退费
        If InStr(1, mstrPrivs, ";收费后退号;") = 0 Then
            If ExistFee(strNO) Then
                MsgBox strNO & "挂号单的病人已经产生了费用,须先退费才能退号.", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    If str划价NO = "" Then
        '退号,获取划价单
        strSQL = "Select NO,记录状态 From 门诊费用记录 " & _
                " Where 记录性质=1 And No = (Select 收费单 From 病人挂号记录 Where NO=[1] And 记录性质=1 and 记录状态=1 and  Rownum<2 )" & _
                " And 记录状态 IN(0,1,3) And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO, "%" & strNO & "%")
        If Not rsTmp.EOF Then
            If Nvl(rsTmp!记录状态, 0) = 0 Then
                str划价NO = Nvl(rsTmp!NO)
            End If
        End If
    End If
    
    '退号总显示详细信息
    If gbln精简界面 Then
        If frmRegistEditSimple.CancelBill(Me, strNO, mlngModul, mstrPrivs) = False Then Exit Sub
    Else
        If frmRegistEdit.CancelBill(Me, strNO, mlngModul, mstrPrivs) = False Then Exit Sub
    End If
    GoTo ReFlash    '刷新


    '去掉了医保连接匹配检查
    lng结帐ID = GetBill结帐ID(strNO, 4, lng病人ID, bln记帐)
   If zlCheckIsAllowBackSN(strNO, bln记帐, bln结帐) = False Then Exit Sub

    '医保退号检查
    If bln记帐 Then
        intInsure = int险类
    Else
        intInsure = ExistInsure(strNO)
    End If
    
    Dim blnStartFactUseType  As Boolean, strUseType As String
    If gblnSharedInvoice And bln记帐 = False Then
        '挂号用门诊票据:42703
        blnStartFactUseType = zlStartFactUseType("1")
        If blnStartFactUseType Then
            strUseType = zl_GetInvoiceUserType(lng病人ID, 0, intInsure)
        End If
    End If
    
    If intInsure > 0 And bln记帐 = False Then
        Set rsTmp = Get结算方式("挂号", "3")
        If rsTmp.RecordCount > 0 Then str个人帐户 = rsTmp!名称
        strAdvance = IIf(str个人帐户 <> "", str个人帐户, "个人帐户")
        If gclsInsure.GetCapability(support门诊结算作废, , intInsure, strAdvance) Then
            strAdvance = ""     '向过程传入不允许退的结算方式,空表示全部允许
        End If
        '67143
        blnVirtualPrint = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure)
        If blnVirtualPrint Then
            If zlGetInvoiceGroupUseID(lng领用ID, , , strUseType) = False Then Exit Sub
            strInvoice = GetNextBill(lng领用ID)
        End If
    End If
    If bln记帐 = False Then
        
        Call zlReadRegThreeBalance(strNO, cllBillBalance)
        
        If Not cllBillBalance Is Nothing Then
            '存在三方账户支付的,需要弹出界面来显示
            If gbln精简界面 Then
                If frmRegistEditSimple.CancelBill(Me, strNO, mlngModul, mstrPrivs) = False Then Exit Sub
            Else
                If frmRegistEdit.CancelBill(Me, strNO, mlngModul, mstrPrivs) = False Then Exit Sub
            End If
            GoTo ReFlash    '刷新
        End If
    End If
    
    blnPromptClear = True
    strSQLCard = ExistCardFee(strNO, lngCard结帐ID)
    If strSQLCard <> "" Then
        '针对第三方代发卡的,则需要弹出提示,是退卡还是取消绑定!
        strSQL = "Select c.卡类别id, c.卡号, c.病人id, d.是否自制" & vbNewLine & _
             "From 门诊费用记录 A, 病人医疗卡变动 B, 病人医疗卡信息 C, 医疗卡类别 D" & vbNewLine & _
             "Where a.记录性质 = 4 And a.No = [1] And a.记录状态 = 1 And b.病人id = a.病人id And b.变动时间 = a.登记时间 And" & vbNewLine & _
             "      b.卡类别id = c.卡类别id And b.卡号 = c.卡号 And c.病人id = a.病人id And c.状态 = 0 And c.卡类别id=d.id And Rownum = 1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTmp.RecordCount <> 0 Then
            If Val(Nvl(rsTmp!是否自制)) = 0 Then
                msgBoxResult = zlCommFun.ShowMsgbox(gstrSysName, "该病人挂号时有代发卡,退号时退卡还是取消绑定?", "退卡,取消绑定,取消", Me, vbQuestion)
                If msgBoxResult = "" Or msgBoxResult = "取消" Then
                    blnPromptClear = False
                    strSQLCard = ""
                    lngCard结帐ID = 0
                    bln退费重打 = gbln退费重打
                    blnCardReprint = gbln退费重打
                ElseIf msgBoxResult = "退卡" Then
                    strSQLCard = "zl_医疗卡记录_DELETE('" & strSQLCard & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                ElseIf msgBoxResult = "取消绑定" Then
                    If IsCheckCancel退预交(rsTmp!病人ID) = True Then
                        strSQLBound = "Zl_医疗卡变动_Insert(14," & Nvl(rsTmp!病人ID) & "," & Nvl(rsTmp!卡类别ID) & ",Null,'" & Nvl(rsTmp!卡号) & "','退号取消绑定'," & _
                                  "Null,'" & UserInfo.姓名 & "',Sysdate)"
                    End If
                    blnPromptClear = False
                    strSQLCard = ""
                    lngCard结帐ID = 0
                    bln退费重打 = gbln退费重打
                    blnCardReprint = gbln退费重打
                End If
            Else
                If MsgBox("该病人挂号时发过卡,退号同时退卡吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    strSQLCard = "zl_医疗卡记录_DELETE('" & strSQLCard & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                Else
                    blnPromptClear = False
                    strSQLCard = ""
                    lngCard结帐ID = 0
                    bln退费重打 = gbln退费重打
                    blnCardReprint = gbln退费重打
                End If
            End If
        Else
            blnPromptClear = False
            strSQLCard = ""
            lngCard结帐ID = 0
            bln退费重打 = gbln退费重打
        End If
    End If
    
    If bln记帐 = False Then
        dblThreeMoney = zlGetRegThreeMoney(lng结帐ID, lngCard结帐ID, cllBillBalance)
    End If
    bytTogetherDo = 0
    '如果挂号单的登记日期-病人信息的登记日期在挂号单有效天数之内,则提示是否删除门诊号   txt时间
    If str门诊号 <> "" And blnPromptClear Then
        If Check挂号时建档(strNO, str挂号时间) Then
            Select Case gbyt清除门诊信息  '35176
            Case 0  '不清除
            Case 1  '清除
                bytTogetherDo = 1
            Case 2  '提示清除
                If MsgBox("退号后要清除与该病人相关的门诊号信息吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    bytTogetherDo = 1
                End If
            End Select
        End If
    End If

    If tbsType.SelectedItem.Key = "挂号" And mTy_Para.bln退号审核 And mTy_Para.lngN天取消预约 > 0 Then
        '退号审核 限制收费预约和预约接收的退号
        If vsThis.TextMatrix(vsThis.Row, getColNum("预约时间")) <> "" Then
            If vsThis.TextMatrix(vsThis.Row, getColNum("退号审核人")) = "" Then
                '是否预约判断放到里面 外面影响性能
                Datsys = zlDatabase.Currentdate
                datTmp = DateAdd("d", -1 * mTy_Para.lngN天取消预约, CDate(vsThis.TextMatrix(vsThis.Row, getColNum("预约时间"))))
                '预约时间-K >datSys
                If Format(Datsys, "yyyy-MM-dd hh:mm:ss") > Format(datTmp, "yyyy-MM-dd hh:mm:ss") Then
                    MsgBox "单据号为" & strNO & "的收费预约单据没有经过退号审核,不能进行退号!", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        End If

    End If
  
    
    If dblThreeMoney <> 0 And bln病历费 And bln记帐 = False Then
        '三方接口,同时买了病历,退号同时也必须退病历,因为接口基本上都要求全退,不支持部分退款
        If MsgBox("单据号为" & strNO & "的单据,挂号的同时购买了病例,同时采用了三方接口扣费,退号时需同时退病历,是否继续?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
        '如果存在三方接口,必须全退
        byt退费方式 = 0    '全退,因为涉及到接口,接口一般要求全退,不能部分退

    ElseIf bln病历费 Then
        If MsgBox("单据号为" & strNO & "的单据,在挂号的同时购买了病例,是否同时退病例?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            byt退费方式 = 0    '全退
        Else
            byt退费方式 = 1    '只退挂号费
            bln退费重打 = gbln退费重打
        End If
    End If

    If MsgBox(strMessage & "确实要将单据[" & strNO & "]退号吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub

    If intInsure = 0 And bln记帐 = False Then
        Set rsOneCard1 = GetOneCardBalance(lng结帐ID)
        If rsOneCard1.RecordCount > 0 Then
            On Error Resume Next
            Set objICCard = CreateObject("zlICCard.clsICCard")
            On Error GoTo 0
            If objICCard Is Nothing Then
                MsgBox "一卡通接口创建失败,不能进行退费!请检查接口文件.", vbInformation, gstrSysName
                Exit Sub
            End If
            strCardNo = objICCard.Read_Card(Me)
            If strCardNo = "" Then Exit Sub
            If strCardNo <> rsOneCard1!单位帐号 Then
                MsgBox "当前卡号与扣款卡号不一致!不能进行退费.", vbInformation, gstrSysName
                Exit Sub
            End If

            If lngCard结帐ID <> 0 Then
                Set rsOneCard2 = GetOneCardBalance(lngCard结帐ID)
            End If
        End If
        '检查三方结算
        If IsCheckCancelValied(lng结帐ID, lngCard结帐ID, cllBillBalance, dblThreeMoney, strNO) = False Then Exit Sub
    End If

    If str划价NO <> "" And bln记帐 = False Then
        strSQL = "zl_门诊划价记录_Delete('" & str划价NO & "')"
        zlAddArray cllPro, strSQL
    End If
    strSQL = "zl_病人挂号记录_DELETE( "
    '单据号_In       门诊费用记录.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '操作员编号_In   门诊费用记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '操作员姓名_In   门诊费用记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '摘要_In         门诊费用记录.摘要%Type := Null, --预约取消时 填写 存放预约取消原因
    strSQL = strSQL & "NULL,"
    '删除门诊号_In   Number := 0,
    strSQL = strSQL & "" & bytTogetherDo & ","
    '非原样退结算_In Varchar2 := Null, --医保不允许的退费结算方式,空表示全部允许
    strSQL = strSQL & "'" & strAdvance & "',"
    '退费类型_In     In Number := 0, --0-全退 1-退挂号费 2-退病历费
    strSQL = strSQL & "" & 0 & ","
    '退指定结算_In   病人预交记录.结算方式%Type := Null,
    strSQL = strSQL & "NULL,"
    '退号重用_In   Number := 1
    strSQL = strSQL & IIf(mTy_Para.blnReuseRegNo, 1, 0) & ")"
    
    zlAddArray cllPro, strSQL
    If strSQLCard <> "" Then zlAddArray cllPro, strSQLCard
    If strSQLBound <> "" Then zlAddArray cllPro, strSQLBound
    If gbyt预存款退费验卡 <> 0 And bln记帐 = False Then
        dblAdvanceMoney = zlGetRegAdvanceMoney(lng结帐ID, lngCard结帐ID)
        If dblAdvanceMoney <> 0 Then
            If Not zlDatabase.PatiIdentify(Me, glngSys, lngPatientID, dblAdvanceMoney, mlngModul, 1, , , True, _
                , , (gbyt预存款退费验卡 = 2)) Then Exit Sub
        End If
    End If
    
    Err = 0: On Error GoTo Errhand:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If intInsure > 0 Then
        Dim strAdvanceTemp As String
        If bln记帐 Then
            strAdvanceTemp = "1|" & strNO
        End If
        If Not gclsInsure.RegistDelSwap(lng结帐ID, intInsure, strAdvanceTemp) Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
    ElseIf Not rsOneCard1 Is Nothing And bln记帐 = False Then
        If rsOneCard1.RecordCount > 0 Then
            If Not objICCard.ReturnSwap(Nvl(rsOneCard1!单位帐号), rsOneCard1!医院编码, "" & Nvl(rsOneCard1!结算号码), rsOneCard1!金额) Then
                gcnOracle.RollbackTrans
                MsgBox "一卡通退费交易调用失败！退费操作失败！", vbExclamation, gstrSysName
                Exit Sub
            End If
            If Not rsOneCard2 Is Nothing Then
                If rsOneCard2.RecordCount > 0 Then
                    If Not objICCard.ReturnSwap(rsOneCard2!单位帐号, rsOneCard2!医院编码, "" & rsOneCard2!结算号码, rsOneCard2!金额) Then
                        gcnOracle.RollbackTrans
                        MsgBox "一卡通退卡费交易调用失败！退费操作失败！", vbExclamation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    If bln记帐 = False Then
        '三方交易
        '退费
        If CallBackBalanceInterface(cllBillBalance, lng结帐ID, lngCard结帐ID, dblThreeMoney, cllUpdate, cllThreeIns, strErrMsg) = False Then
            gcnOracle.RollbackTrans
            If strErrMsg <> "" Then
                MsgBox strErrMsg, vbExclamation + vbOKOnly, gstrSysName
            Else
                MsgBox "调用第三方接口交易失败,此次退费操作失败!", vbExclamation + vbOKOnly, gstrSysName
            End If
            Exit Sub
        End If
    
        If Not cllBillBalance Is Nothing And Not cllUpdate Is Nothing Then
            zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
        End If
    End If
    gcnOracle.CommitTrans
    If Not cllThreeIns Is Nothing And bln记帐 = False Then
        Err = 0: On Error GoTo ErrOthers:
        zlExecuteProcedureArrAy cllThreeIns, Me.Caption
    End If
    '继续执行
ResumeExecute:
    '问题:31634
    Err = 0: On Error GoTo NotCommit:
    If intInsure > 0 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistDelSwap, True, intInsure)
    blnTrans = False
    If gblnBillPrint Then
        Err = 0: On Error Resume Next
        Call gobjBillPrint.zlEraseBill_Reg("'" & strNO & "'")
        If Err <> 0 Then
            Err = 0
        End If
        On Error GoTo errH
    End If
    '70262:刘尔旋,2014-03-04,退号走票的问题
'    If bln退费重打 And bln记帐 = False And (byt退费方式 <> 0 Or blnCardReprint) Then Call RePrintBill(Me, strNO, lng结帐ID, intInsure, blnVirtualPrint, , strUseType)
    If strAdvance <> "" And bln记帐 = False Then
        MsgBox "医保不支持[" & strAdvance & "]回退,退为现金." & vbCrLf & vbCrLf & "退款共计:" & Format(GetCashMoney(strNO), "0.00") & " 元.", vbInformation, gstrSysName
    End If
ReFlash:
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
ErrOthers:
    If ErrCenter = 1 Then gcnOracle.RollbackTrans: Resume
    gcnOracle.CommitTrans
    GoTo ResumeExecute:
    Exit Sub
    '问题:31634
NotCommit:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    If intInsure > 0 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistDelSwap, False, intInsure)
End Sub

Private Sub DeleteRegist()
    Dim strSQL As String, strNO As String, str划价NO As String, strCardNo As String
    Dim intInsure As Integer, lng结帐ID As Long, lngCard结帐ID As Long, msgBoxResult As String
    Dim str挂号时间 As String, strSQLCard As String, strMessage As String
    Dim str门诊号 As String, strAdvance As String, str个人帐户 As String
    Dim blnEnableDel As Boolean, blnTrans As Boolean, strSQLBound As String
    Dim bytTogetherDo As Byte, bln退费重打 As Boolean, blnPromptClear As Boolean
    Dim rsTmp As ADODB.Recordset, rsOneCard1 As ADODB.Recordset, rsOneCard2 As ADODB.Recordset
    Dim objICCard As Object
    Dim cllPro As Collection, cllBillBalance As Collection, dblThreeMoney As Double
    Dim cllUpdate As Collection, cllThreeIns As Collection, strErrMsg As String
    Dim byt退费方式 As Byte  '0-全退 1-只退挂号费 2-只退病历
    Dim bln病历费 As Boolean    '是否包含病历费
    Dim blnCardReprint As Boolean    '不退卡重打
    Dim Datsys As Date
    Dim datTmp As Date
    Dim lngPatientID As Long
    Dim dblAdvanceMoney As Double    '预交结算费用
    Dim strInvoice As String, lng病人ID As Long, lng领用ID As Long
    Dim blnVirtualPrint As Boolean
    Dim bln记帐 As Boolean, int险类 As Integer, bln结帐 As Boolean
    
    Set cllPro = New Collection
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))

    If strNO = "" Then
        MsgBox "当前没有记录可以退号！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If IsNewModeRegist(strNO) = False Then
        Call DelOldRegist
        Exit Sub
    End If
    
    If CheckBillExistReplenishData(strNO) Then
        MsgBox "选择的挂号记录进行了医保补充结算，不允许进行退号操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errH
    str挂号时间 = vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("挂号时间"))
    str门诊号 = vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("门诊号"))
    bln病历费 = Trim(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("病历"))) <> ""
    lngPatientID = Val(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("病人ID")))
    
    int险类 = Val(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("险类")))
    bln记帐 = Val(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("记帐费用"))) = 1
    
    
    If InStr(1, mstrPrivs, ";强制退号;") = 0 Then
        '判断当前人员对单据是否有操作权限,时间限制,无需检查挂号单有效天数
        If Not BillOperCheck(1, vsThis.TextMatrix(vsThis.Row, getColNum("挂号员")), _
                             CDate(str挂号时间), "退号") Then Exit Sub
    End If

    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 4, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If

    If InStr(1, mstrPrivs, ";强制退号;") = 0 Then   '问题:
        '检查挂号单是否已执行
        If InStr(";" & mstrPrivs & ";", ";下医嘱后退号;") > 0 Then
            blnEnableDel = True
        End If
        If CheckExecuted(strNO, blnEnableDel) Then
            MsgBox "挂号单" & strNO & "已经被医生接诊或下过医嘱,不能退号！", vbInformation, gstrSysName
            Exit Sub
        End If
        '医生站挂的号-收费判断
        If CheckPriceHaveFee(strNO, str划价NO) Then Exit Sub
        '是否发生过费用,但未退费
        If InStr(1, mstrPrivs, ";收费后退号;") = 0 Then
            If ExistFee(strNO) Then
                MsgBox strNO & "挂号单的病人已经产生了费用,须先退费才能退号.", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    If frmRegistEditNew.CancelBill(Me, strNO, mlngModul, mstrPrivs) = False Then Exit Sub
    GoTo ReFlash    '刷新


    '去掉了医保连接匹配检查
    lng结帐ID = GetBill结帐ID(strNO, 4, lng病人ID, bln记帐)
   If zlCheckIsAllowBackSN(strNO, bln记帐, bln结帐) = False Then Exit Sub

    '医保退号检查
    If bln记帐 Then
        intInsure = int险类
    Else
        intInsure = ExistInsure(strNO)
    End If
    
    Dim blnStartFactUseType  As Boolean, strUseType As String
    If gblnSharedInvoice And bln记帐 = False Then
        '挂号用门诊票据:42703
        blnStartFactUseType = zlStartFactUseType("1")
        If blnStartFactUseType Then
            strUseType = zl_GetInvoiceUserType(lng病人ID, 0, intInsure)
        End If
    End If
    
    If intInsure > 0 And bln记帐 = False Then
        Set rsTmp = Get结算方式("挂号", "3")
        If rsTmp.RecordCount > 0 Then str个人帐户 = rsTmp!名称
        strAdvance = IIf(str个人帐户 <> "", str个人帐户, "个人帐户")
        If gclsInsure.GetCapability(support门诊结算作废, , intInsure, strAdvance) Then
            strAdvance = ""     '向过程传入不允许退的结算方式,空表示全部允许
        End If
        '67143
        blnVirtualPrint = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure)
        If blnVirtualPrint Then
            If zlGetInvoiceGroupUseID(lng领用ID, , , strUseType) = False Then Exit Sub
            strInvoice = GetNextBill(lng领用ID)
        End If
    End If
    
    If bln记帐 = False Then
        Call zlReadRegThreeBalance(strNO, cllBillBalance)
        If Not cllBillBalance Is Nothing Then
            '存在三方账户支付的,需要弹出界面来显示
            If frmRegistEditNew.CancelBill(Me, strNO, mlngModul, mstrPrivs) = False Then Exit Sub
            GoTo ReFlash    '刷新
        End If
    End If
    
    blnPromptClear = True
    strSQLCard = ExistCardFee(strNO, lngCard结帐ID)
    If strSQLCard <> "" Then
        '针对第三方代发卡的,则需要弹出提示,是退卡还是取消绑定!
        strSQL = "Select c.卡类别id, c.卡号, c.病人id, d.是否自制" & vbNewLine & _
             "From 门诊费用记录 A, 病人医疗卡变动 B, 病人医疗卡信息 C, 医疗卡类别 D" & vbNewLine & _
             "Where a.记录性质 = 4 And a.No = [1] And a.记录状态 = 1 And b.病人id = a.病人id And b.变动时间 = a.登记时间 And" & vbNewLine & _
             "      b.卡类别id = c.卡类别id And b.卡号 = c.卡号 And c.病人id = a.病人id And c.状态 = 0 And c.卡类别id=d.id And Rownum = 1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTmp.RecordCount <> 0 Then
            If Val(Nvl(rsTmp!是否自制)) = 0 Then
                msgBoxResult = zlCommFun.ShowMsgbox(gstrSysName, "该病人挂号时有代发卡,退号时退卡还是取消绑定?", "退卡,取消绑定,取消", Me, vbQuestion)
                If msgBoxResult = "" Or msgBoxResult = "取消" Then
                    blnPromptClear = False
                    strSQLCard = ""
                    lngCard结帐ID = 0
                    bln退费重打 = gbln退费重打
                    blnCardReprint = gbln退费重打
                ElseIf msgBoxResult = "退卡" Then
                    strSQLCard = "zl_医疗卡记录_DELETE('" & strSQLCard & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                ElseIf msgBoxResult = "取消绑定" Then
                    If IsCheckCancel退预交(rsTmp!病人ID) = True Then
                        strSQLBound = "Zl_医疗卡变动_Insert(14," & Nvl(rsTmp!病人ID) & "," & Nvl(rsTmp!卡类别ID) & ",Null,'" & Nvl(rsTmp!卡号) & "','退号取消绑定'," & _
                                  "Null,'" & UserInfo.姓名 & "',Sysdate)"
                    End If
                    blnPromptClear = False
                    strSQLCard = ""
                    lngCard结帐ID = 0
                    bln退费重打 = gbln退费重打
                    blnCardReprint = gbln退费重打
                End If
            Else
                If MsgBox("该病人挂号时发过卡,退号同时退卡吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    strSQLCard = "zl_医疗卡记录_DELETE('" & strSQLCard & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                Else
                    blnPromptClear = False
                    strSQLCard = ""
                    lngCard结帐ID = 0
                    bln退费重打 = gbln退费重打
                    blnCardReprint = gbln退费重打
                End If
            End If
        Else
            blnPromptClear = False
            strSQLCard = ""
            lngCard结帐ID = 0
            bln退费重打 = gbln退费重打
        End If
    End If
    
    If bln记帐 = False Then
        dblThreeMoney = zlGetRegThreeMoney(lng结帐ID, lngCard结帐ID, cllBillBalance)
    End If
    bytTogetherDo = 0
    '如果挂号单的登记日期-病人信息的登记日期在挂号单有效天数之内,则提示是否删除门诊号   txt时间
    If str门诊号 <> "" And blnPromptClear Then
        If Check挂号时建档(strNO, str挂号时间) Then
            Select Case gbyt清除门诊信息  '35176
            Case 0  '不清除
            Case 1  '清除
                bytTogetherDo = 1
            Case 2  '提示清除
                If MsgBox("退号后要清除与该病人相关的门诊号信息吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    bytTogetherDo = 1
                End If
            End Select
        End If
    End If

    If tbsType.SelectedItem.Key = "挂号" And mTy_Para.bln退号审核 And mTy_Para.lngN天取消预约 > 0 Then
        '退号审核 限制收费预约和预约接收的退号
        If vsThis.TextMatrix(vsThis.Row, getColNum("预约时间")) <> "" Then
            If vsThis.TextMatrix(vsThis.Row, getColNum("退号审核人")) = "" Then
                '是否预约判断放到里面 外面影响性能
                Datsys = zlDatabase.Currentdate
                datTmp = DateAdd("d", -1 * mTy_Para.lngN天取消预约, CDate(vsThis.TextMatrix(vsThis.Row, getColNum("预约时间"))))
                '预约时间-K >datSys
                If Format(Datsys, "yyyy-MM-dd hh:mm:ss") > Format(datTmp, "yyyy-MM-dd hh:mm:ss") Then
                    MsgBox "单据号为" & strNO & "的收费预约单据没有经过退号审核,不能进行退号!", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        End If

    End If
  
    
    If dblThreeMoney <> 0 And bln病历费 And bln记帐 = False Then
        '三方接口,同时买了病历,退号同时也必须退病历,因为接口基本上都要求全退,不支持部分退款
        If MsgBox("单据号为" & strNO & "的单据,挂号的同时购买了病例,同时采用了三方接口扣费,退号时需同时退病历,是否继续?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
        '如果存在三方接口,必须全退
        byt退费方式 = 0    '全退,因为涉及到接口,接口一般要求全退,不能部分退

    ElseIf bln病历费 Then
        If MsgBox("单据号为" & strNO & "的单据,在挂号的同时购买了病例,是否同时退病例?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            byt退费方式 = 0    '全退
        Else
            byt退费方式 = 1    '只退挂号费
            bln退费重打 = gbln退费重打
        End If
    End If

    If MsgBox(strMessage & "确实要将单据[" & strNO & "]退号吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub

    If intInsure = 0 And bln记帐 = False Then
        Set rsOneCard1 = GetOneCardBalance(lng结帐ID)
        If rsOneCard1.RecordCount > 0 Then
            On Error Resume Next
            Set objICCard = CreateObject("zlICCard.clsICCard")
            On Error GoTo 0
            If objICCard Is Nothing Then
                MsgBox "一卡通接口创建失败,不能进行退费!请检查接口文件.", vbInformation, gstrSysName
                Exit Sub
            End If
            strCardNo = objICCard.Read_Card(Me)
            If strCardNo = "" Then Exit Sub
            If strCardNo <> rsOneCard1!单位帐号 Then
                MsgBox "当前卡号与扣款卡号不一致!不能进行退费.", vbInformation, gstrSysName
                Exit Sub
            End If

            If lngCard结帐ID <> 0 Then
                Set rsOneCard2 = GetOneCardBalance(lngCard结帐ID)
            End If
        End If
        '检查三方结算
        If IsCheckCancelValied(lng结帐ID, lngCard结帐ID, cllBillBalance, dblThreeMoney, strNO) = False Then Exit Sub
    End If

    If str划价NO <> "" And bln记帐 = False Then
        strSQL = "zl_门诊划价记录_Delete('" & str划价NO & "')"
        zlAddArray cllPro, strSQL
    End If
    strSQL = "zl_病人挂号记录_出诊_DELETE( "
    '单据号_In       门诊费用记录.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '操作员编号_In   门诊费用记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '操作员姓名_In   门诊费用记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '摘要_In         门诊费用记录.摘要%Type := Null, --预约取消时 填写 存放预约取消原因
    strSQL = strSQL & "NULL,"
    '删除门诊号_In   Number := 0,
    strSQL = strSQL & "" & bytTogetherDo & ","
    '非原样退结算_In Varchar2 := Null, --医保不允许的退费结算方式,空表示全部允许
    strSQL = strSQL & "'" & strAdvance & "',"
    '退费类型_In     In Number := 0, --0-全退 1-退挂号费 2-退病历费
    strSQL = strSQL & "" & 0 & ","
    '退指定结算_In   病人预交记录.结算方式%Type := Null,
    strSQL = strSQL & "NULL,"
    '退号重用_In   Number := 1
    strSQL = strSQL & IIf(mTy_Para.blnReuseRegNo, 1, 0) & ")"
    
    zlAddArray cllPro, strSQL
    If strSQLCard <> "" Then zlAddArray cllPro, strSQLCard
    If strSQLBound <> "" Then zlAddArray cllPro, strSQLBound
    If gbyt预存款退费验卡 <> 0 And bln记帐 = False Then
        dblAdvanceMoney = zlGetRegAdvanceMoney(lng结帐ID, lngCard结帐ID)
        If dblAdvanceMoney <> 0 Then
            If Not zlDatabase.PatiIdentify(Me, glngSys, lngPatientID, dblAdvanceMoney, mlngModul, 1, , , True, _
                , , (gbyt预存款退费验卡 = 2)) Then Exit Sub
        End If
    End If
    
    Err = 0: On Error GoTo Errhand:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If intInsure > 0 Then
        Dim strAdvanceTemp As String
        If bln记帐 Then
            strAdvanceTemp = "1|" & strNO
        End If
        If Not gclsInsure.RegistDelSwap(lng结帐ID, intInsure, strAdvanceTemp) Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
    ElseIf Not rsOneCard1 Is Nothing And bln记帐 = False Then
        If rsOneCard1.RecordCount > 0 Then
            If Not objICCard.ReturnSwap(Nvl(rsOneCard1!单位帐号), rsOneCard1!医院编码, "" & Nvl(rsOneCard1!结算号码), rsOneCard1!金额) Then
                gcnOracle.RollbackTrans
                MsgBox "一卡通退费交易调用失败！退费操作失败！", vbExclamation, gstrSysName
                Exit Sub
            End If
            If Not rsOneCard2 Is Nothing Then
                If rsOneCard2.RecordCount > 0 Then
                    If Not objICCard.ReturnSwap(rsOneCard2!单位帐号, rsOneCard2!医院编码, "" & rsOneCard2!结算号码, rsOneCard2!金额) Then
                        gcnOracle.RollbackTrans
                        MsgBox "一卡通退卡费交易调用失败！退费操作失败！", vbExclamation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    If bln记帐 = False Then
        '三方交易
        '退费
        If CallBackBalanceInterface(cllBillBalance, lng结帐ID, lngCard结帐ID, dblThreeMoney, cllUpdate, cllThreeIns, strErrMsg) = False Then
            gcnOracle.RollbackTrans
            If strErrMsg <> "" Then
                MsgBox strErrMsg, vbExclamation + vbOKOnly, gstrSysName
            Else
                MsgBox "调用第三方接口交易失败,此次退费操作失败!", vbExclamation + vbOKOnly, gstrSysName
            End If
            Exit Sub
        End If
    
        If Not cllBillBalance Is Nothing And Not cllUpdate Is Nothing Then
            zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
        End If
    End If
    gcnOracle.CommitTrans
    If Not cllThreeIns Is Nothing And bln记帐 = False Then
        Err = 0: On Error GoTo ErrOthers:
        zlExecuteProcedureArrAy cllThreeIns, Me.Caption
    End If
    '继续执行
ResumeExecute:
    '问题:31634
    Err = 0: On Error GoTo NotCommit:
    If intInsure > 0 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistDelSwap, True, intInsure)
    blnTrans = False
    '70262:刘尔旋,2014-03-04,退号走票的问题
'    If bln退费重打 And bln记帐 = False And (byt退费方式 <> 0 Or blnCardReprint) Then Call RePrintBill(Me, strNO, lng结帐ID, intInsure, blnVirtualPrint, , strUseType)
    If strAdvance <> "" And bln记帐 = False Then
        MsgBox "医保不支持[" & strAdvance & "]回退,退为现金." & vbCrLf & vbCrLf & "退款共计:" & Format(GetCashMoney(strNO), "0.00") & " 元.", vbInformation, gstrSysName
    End If
ReFlash:
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
ErrOthers:
    If ErrCenter = 1 Then gcnOracle.RollbackTrans: Resume
    gcnOracle.CommitTrans
    GoTo ResumeExecute:
    Exit Sub
    '问题:31634
NotCommit:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    If intInsure > 0 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistDelSwap, False, intInsure)
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub mnuEdit_Add_Click()
    On Error Resume Next
    Dim datNow As Date
    Err.Clear
    
    datNow = zlDatabase.Currentdate
    If Format(datNow, "yyyy-mm-dd hh:mm:ss") < Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") Then
        If gbln精简界面 Then
            frmRegistEditSimple.mlngModul = mlngModul
            frmRegistEditSimple.mstrPrivs = mstrPrivs
            frmRegistEditSimple.mbytMode = 0
            frmRegistEditSimple.mbytInState = 0
            Set frmRegistEditSimple.mobjMsgModule = mobjMsgModule '消息处理模块
            frmRegistEditSimple.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            frmRegistEdit.mlngModul = mlngModul
            frmRegistEdit.mstrPrivs = mstrPrivs
            frmRegistEdit.mbytMode = 0
            frmRegistEdit.mbytInState = 0
            Set frmRegistEdit.mobjMsgModule = mobjMsgModule '消息处理模块
            frmRegistEdit.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        End If
    Else
        frmRegistEditNew.mlngModul = mlngModul
        frmRegistEditNew.mstrPrivs = mstrPrivs
        frmRegistEditNew.mbytMode = 0
        frmRegistEditNew.mbytInState = 0
        Set frmRegistEditNew.mobjMsgModule = mobjMsgModule '消息处理模块
        frmRegistEditNew.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End If
    If gblnOk And tbsType.SelectedItem.Key = "挂号" Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
    If Me.Enabled And Me.Visible Then Me.SetFocus
End Sub

Private Sub ViewOldRegist()
    If vsThis.TextMatrix(vsThis.Row, getColNum("单据号")) = "" Then
        MsgBox "当前没有记录可以查阅！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    Err.Clear
    Dim blnCancel As Boolean
    Dim bytInState As Byte
    Dim bytViewState As Byte
    Dim strNO As String, rsTemp As ADODB.Recordset, strSQL As String
    bytInState = 1
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
    If tbsType.SelectedItem.Key = "预约" Then
        blnCancel = vsThis.TextMatrix(vsThis.Row, getColNum("记录状态")) <> "1"
        bytInState = IIf(vsThis.TextMatrix(vsThis.Row, getColNum("记录状态")) = "1", 1, Val(vsThis.TextMatrix(vsThis.Row, getColNum("记录状态"))))
    Else
        blnCancel = vsThis.TextMatrix(vsThis.Row, getColNum("记录状态")) = "2"
    End If
    If tbsType.SelectedItem.Key = "挂号" Then
        bytViewState = vsThis.TextMatrix(vsThis.Row, getColNum("记录状态"))
        '保险补充结算的退费单时，检查是否为异常退费单据
        If bytViewState = 2 Then
            If CheckBillExistReplenishData(strNO) Then
                strSQL = "Select 1 From 门诊费用记录 Where 记录性质 = 4 And 记录状态 = 2 And Nvl(费用状态, 0) = 1 And NO = [1] And Rownum < 2"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否为退费异常单据", strNO)
                If Not rsTemp.EOF Then
                    MsgBox "当前选择挂号单正处于保险补充结算退费异常状态，暂不允许查看！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        If mbytCancel <> 1 And bytViewState = 1 Then
            bytInState = IIf(vsThis.TextMatrix(vsThis.Row, getColNum("记录状态")) = "1", 1, Val(vsThis.TextMatrix(vsThis.Row, getColNum("记录状态"))))
            blnCancel = bytInState <> 1
        End If
    End If
    If gbln精简界面 Then
        frmRegistEditSimple.mlngModul = mlngModul
        frmRegistEditSimple.mstrPrivs = mstrPrivs
        frmRegistEditSimple.mstrNoIn = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
        frmRegistEditSimple.mblnNOMoved = mblnNOMoved
        frmRegistEditSimple.mbytMode = tbsType.SelectedItem.index - 1
        frmRegistEditSimple.mbytInState = bytInState
        frmRegistEditSimple.mblnViewCancel = blnCancel
        frmRegistEditSimple.mblnViewOriginal = vsThis.TextMatrix(vsThis.Row, getColNum("记录状态")) = "3"
        Set frmRegistEditSimple.mobjMsgModule = mobjMsgModule '消息处理模块
        frmRegistEditSimple.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else
        frmRegistEdit.mlngModul = mlngModul
        frmRegistEdit.mstrPrivs = mstrPrivs
        frmRegistEdit.mstrNoIn = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
        frmRegistEdit.mblnNOMoved = mblnNOMoved
        frmRegistEdit.mbytMode = tbsType.SelectedItem.index - 1
        frmRegistEdit.mbytInState = bytInState
        frmRegistEdit.mblnViewCancel = blnCancel
        frmRegistEdit.mblnViewOriginal = vsThis.TextMatrix(vsThis.Row, getColNum("记录状态")) = "3"
        Set frmRegistEdit.mobjMsgModule = mobjMsgModule '消息处理模块
        frmRegistEdit.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End If
End Sub

Private Sub mnuEdit_View_Click()
    If vsThis.TextMatrix(vsThis.Row, getColNum("单据号")) = "" Then
        MsgBox "当前没有记录可以查阅！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    Err.Clear
    Dim blnCancel As Boolean
    Dim bytInState As Byte
    Dim bytViewState As Byte
    Dim strNO As String, rsTemp As ADODB.Recordset, strSQL As String
    bytInState = 1
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
    
    If IsNewModeRegist(strNO) = False Then
        Call ViewOldRegist
        Exit Sub
    End If
    
    If tbsType.SelectedItem.Key = "预约" Then
        blnCancel = vsThis.TextMatrix(vsThis.Row, getColNum("记录状态")) <> "1"
        bytInState = IIf(vsThis.TextMatrix(vsThis.Row, getColNum("记录状态")) = "1", 1, Val(vsThis.TextMatrix(vsThis.Row, getColNum("记录状态"))))
    Else
        blnCancel = vsThis.TextMatrix(vsThis.Row, getColNum("记录状态")) = "2"
    End If
    If tbsType.SelectedItem.Key = "挂号" Then
        bytViewState = vsThis.TextMatrix(vsThis.Row, getColNum("记录状态"))
        '保险补充结算的退费单时，检查是否为异常退费单据
        If bytViewState = 2 Then
            If CheckBillExistReplenishData(strNO) Then
                strSQL = "Select 1 From 门诊费用记录 Where 记录性质 = 4 And 记录状态 = 2 And Nvl(费用状态, 0) = 1 And NO = [1] And Rownum < 2"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否为退费异常单据", strNO)
                If Not rsTemp.EOF Then
                    MsgBox "当前选择挂号单正处于保险补充结算退费异常状态，暂不允许查看！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        If mbytCancel <> 1 And bytViewState = 1 Then
            bytInState = IIf(vsThis.TextMatrix(vsThis.Row, getColNum("记录状态")) = "1", 1, Val(vsThis.TextMatrix(vsThis.Row, getColNum("记录状态"))))
            blnCancel = bytInState <> 1
        End If
    End If
    frmRegistEditNew.mlngModul = mlngModul
    frmRegistEditNew.mstrPrivs = mstrPrivs
    frmRegistEditNew.mstrNoIn = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
    frmRegistEditNew.mblnNOMoved = mblnNOMoved
    frmRegistEditNew.mbytMode = tbsType.SelectedItem.index - 1
    frmRegistEditNew.mbytInState = bytInState
    frmRegistEditNew.mblnViewCancel = blnCancel
    frmRegistEditNew.mblnViewOriginal = vsThis.TextMatrix(vsThis.Row, getColNum("记录状态")) = "3"
    Set frmRegistEditNew.mobjMsgModule = mobjMsgModule '消息处理模块
    frmRegistEditNew.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If Me.Enabled And Me.Visible Then Me.SetFocus
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewReFlash_Click()
    Call ShowBills
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

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "Go" '定位
            mnuViewGo_Click
        Case "Filter" '过滤
            mnuViewFilter_Click
        Case "View"
            mnuEdit_View_Click
        Case "Add"
            mnuEdit_Add_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "预约"
            mnuEdit_Bespeak_Click
        Case "接收"
            mnuEdit_Incept_Click
        Case "取消"
            mnuEdit_Cancel_Click
        Case "轧帐"
            mnuFileRollingCurtain_Click
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
    
    intRow = vsThis.Row
    
    '表头
    objOut.Title.Text = "门诊挂号单据清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    With frmRegistFilter
        If IsNull(.dtpEnd.Value) Then
            objRow.Add "时间：" & Format(.dtpBegin.Value, "yyyy-MM-dd")
        Else
            objRow.Add "时间：" & Format(.dtpBegin.Value, "yyyy-MM-dd HH:MM") & " 至 " & Format(.dtpEnd.Value, "yyyy-MM-dd HH:MM")
        End If
        objRow.Add "性质：" & IIf(.optRegistRecord(1).Value = True, "退款记录", "收款记录")
        objOut.UnderAppRows.Add objRow
    End With
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    vsThis.Redraw = False
    Set objOut.Body = vsThis
    
    '输出
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    vsThis.Row = intRow
    vsThis.Col = 0: vsThis.ColSel = vsThis.Cols - 1
    vsThis.Redraw = True
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage Hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo Hwnd
End Sub

Private Function IsNewModeRegist(ByVal strNO As String) As Boolean
'功能：判断挂号单是否为出诊表排班模式挂号单
'参数：strNo = 挂号单单据号
    Dim strSQL As String, rsTemp As ADODB.Recordset
    strSQL = "Select 1 From 病人挂号记录 Where NO = [1] And 出诊记录Id Is Null And 发生时间 < [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, gdatRegistTime)
    If rsTemp.EOF Then
        IsNewModeRegist = True
    Else
        IsNewModeRegist = False
    End If
End Function

Private Sub SetMenu(blnUsed As Boolean)
'功能：根据有无记录设置菜单可用状态
'参数：blnUsed=表明当前清单中有无数据
    
    mnuFile_Print.Enabled = blnUsed
    mnuFile_Preview.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    
    mnuEdit_Del.Enabled = blnUsed And tbsType.SelectedItem.Key = "挂号"
    '重打
    mnuEdit_Print.Enabled = blnUsed And tbsType.SelectedItem.Key = "挂号"
    '补打
    mnuEdit_Print_Supplemental.Enabled = blnUsed And tbsType.SelectedItem.Key = "挂号"
    mnuEdit_Print_Slip.Enabled = blnUsed And tbsType.SelectedItem.Key = "挂号"
    mnuEdit_Print_Case.Enabled = blnUsed And tbsType.SelectedItem.Key = "挂号"
    tbr.Buttons("Del").Enabled = blnUsed And tbsType.SelectedItem.Key = "挂号"
    
    mnuEdit_Incept.Enabled = blnUsed And tbsType.SelectedItem.Key = "接收"
    tbr.Buttons("接收").Enabled = blnUsed And tbsType.SelectedItem.Key = "接收"
    
    mnuEdit_Cancel.Enabled = blnUsed And tbsType.SelectedItem.Key <> "挂号"
    tbr.Buttons("取消").Enabled = blnUsed And tbsType.SelectedItem.Key <> "挂号"


    mnuEdit_View.Enabled = blnUsed
    tbr.Buttons("View").Enabled = blnUsed
    mnuViewGo.Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
    '
    mnuEdit_CancelAuditing.Enabled = blnUsed And mnuEdit_CancelAuditing.Visible
    '问题号:45507
    mnuEdit_BatchChangeNum.Enabled = InStr(1, mstrPrivs, ";门诊批量换号;") > 0
    
    Call vsThis_RowColChange
End Sub

Private Sub Form_Load()
    Dim i As Integer, blnHavePrivs As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    mstrPrivs_RollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    
    '自动作业
    strSQL = "zl1_Auto_Buildingregisterplan"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    mstrVsType = tbsType.SelectedItem.Key
    
    mstr附加费 = ""
    mstr附加项目ID = ""
    strSQL = "Select zl_Fun_RegCustomName As 附加费 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        mstr附加费 = Split(Nvl(rsTmp!附加费) & "|", "|")(0)
        mstr附加项目ID = Split(Nvl(rsTmp!附加费) & "|", "|")(1)
    End If
    
    If mstr附加费 <> "" And mstr附加项目ID <> "" Then
        mnuEdit_DelExtra.Caption = "退" & mstr附加费 & "(&E)"
        tbr.Buttons("Del").ButtonMenus("DelExtra").Text = "退" & mstr附加费
        mnuEdit_DelExtra.Visible = True
        tbr.Buttons("Del").ButtonMenus("DelExtra").Visible = True
    Else
        mnuEdit_DelExtra.Visible = False
        tbr.Buttons("Del").ButtonMenus("DelExtra").Visible = False
    End If
    
    Call Form_Resize '避免Bh调用不触发事件Form_Resize
    Call tbsType_Click
    Call RestoreWinState(Me, App.ProductName)
      
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If i = Val(zlDatabase.GetPara("刷新方式", glngSys, mlngModul, 2)) Then
            mnuViewRefeshOptionItem(i).Checked = True
        Else
            mnuViewRefeshOptionItem(i).Checked = False
        End If
    Next
    
    '权限设置
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL" & glngSys \ 100 & "_INSIDE_1111_1")
    
    If InStr(mstrPrivs, ";LED与语音;") = 0 Then gblnLED = False
    
    '正常挂号
    If InStr(";" & mstrPrivs & ";", ";挂收费号;") = 0 And InStr(";" & mstrPrivs & ";", ";挂免费号;") = 0 Then
        mnuEdit_Add.Visible = False
        tbr.Buttons("Add").Visible = False
        mnuEdit_Print.Visible = False
    End If
    '52328
    mnuEdit_Print_Supplemental.Visible = (InStr(mstrPrivs, ";挂免费号;") > 0 Or InStr(mstrPrivs, ";挂收费号;") > 0) And InStr(mstrPrivs, ";补打票据;") > 0
    If InStr(";" & mstrPrivs & ";", ";重打票据;") = 0 Then
        mnuEdit_Print.Visible = False
    End If
    mnuEdit_Print_Slip.Visible = InStr(mstrPrivs, ";挂号凭条打印;") > 0
    If InStr(";" & mstrPrivs & ";", ";退号;") = 0 Then
        mnuEdit_Del.Visible = False
        tbr.Buttons("Del").Visible = False
    End If
    If InStr(";" & mstrPrivs & ";", ";挂免费号;") = 0 And InStr(mstrPrivs, ";挂收费号;") = 0 _
        And InStr(";" & mstrPrivs & ";", ";退号;") = 0 Then
        mnuEdit_1.Visible = False
        tbr.Buttons("Fun_1").Visible = False
    End If
    If InStr(";" & mstrPrivs & ";", ";退号审核;") = 0 Then
        mnuEdit_CancelAuditing.Visible = False
    End If
    If InStr(";" & mstrPrivs & ";", ";门诊号绑定;") = 0 Then
    '门诊号绑定
        mnuEdit_BindPatNum.Enabled = False
    End If
    
    '收费轧帐管理
    blnHavePrivs = InStr(mstrPrivs_RollingCurtain, ";轧帐;") > 0
    mnuFileRollingCurtain.Visible = blnHavePrivs
    mnuFileRollingCurtainSplit.Visible = blnHavePrivs
    tbr.Buttons("轧帐").Visible = blnHavePrivs
    tbr.Buttons("SplitRollingCurtain").Visible = blnHavePrivs
    
    '预约挂号
    If InStr(mstrPrivs, ";预约挂号;") = 0 Then
        mnuEdit_Bespeak.Visible = False
        tbr.Buttons("预约").Visible = False
        frmBookingDefer.Visible = False
    End If
    If InStr(mstrPrivs, ";接收预约;") = 0 Then
        mnuEdit_Incept.Visible = False
        tbr.Buttons("接收").Visible = False
    End If
    If InStr(mstrPrivs, ";取消预约;") = 0 Then
        mnuEdit_Cancel.Visible = False
        mnuEdit_Clear.Visible = False
        tbr.Buttons("取消").Visible = False
    End If
    If InStr(mstrPrivs, "'预约挂号;") = 0 _
        And InStr(mstrPrivs, ";接收预约;") = 0 _
        And InStr(mstrPrivs, ";取消预约;") = 0 Then
        mnuEdit_2.Visible = False
        tbr.Buttons("Fun_2").Visible = False
    End If
            
    Call SetHeader
    Call SetMenu(False)
    mbytCancel = 1: mstrFilter = ""
    stbThis.Panels(2).Text = "请刷新清单或重新设置过滤条件"
    '根据操作权限定位缺省清单
    If InStr(";" & mstrPrivs & ";", ";挂免费号;") > 0 Or InStr(mstrPrivs, ";挂收费号;") > 0 Then
        '缺省就是该页
    ElseIf InStr(mstrPrivs, ";预约挂号;") > 0 Then
        tbsType.Tabs("预约").Selected = True
    ElseIf InStr(mstrPrivs, ";接收预约;") > 0 Then
        tbsType.Tabs("接收").Selected = True
    End If
    
    On Error GoTo errH
    InitActionType
    Call InitPara
    If mactionType = t_时段 Then
        mnuEdit_Defer.Enabled = False
    End If
    
    '创建第三方票据打印部件
    On Error Resume Next
    gblnBillPrint = False
    Set gobjBillPrint = CreateObject("zlBillPrint.clsBillPrint")
    If Not gobjBillPrint Is Nothing Then
        gblnBillPrint = gobjBillPrint.zlInitialize(gcnOracle, glngSys, mlngModul, UserInfo.编号, UserInfo.姓名)
    End If
    On Error GoTo errH
    
    '初始化消息处理对象模块
    Call InitMsgModule
    
    '创建并检测税控打印对象
    If gobjTax Is Nothing Then
        On Error Resume Next
        Set gobjTax = CreateObject("zl9TaxBill.clsTaxBill")
        If Err.Number = 0 And Not gobjTax Is Nothing Then
            gblnTax = gobjTax.zlTaxUseable(1)
        End If
        On Error GoTo 0
    End If
    
    Call LoadPlugInMnu
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '工具条占用高度
    Dim staH As Long '状态栏占用高度
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    vsThis.MousePointer = 0
    
    '靠齐控件宽度和高度
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    With tbsType
        .Top = Me.ScaleTop + cbrH + 15
        .Left = Me.ScaleLeft + 30
        .Width = Me.ScaleWidth - 60
    End With
    
    With vsThis
        .Top = Me.ScaleTop + cbrH + 350
        .Height = Me.ScaleHeight - cbrH - staH - 350
        .Left = Me.ScaleLeft
        .Width = Me.ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    Unload frmRegistFilter
    Unload frmRegistFind
    Call SaveWinState(Me, App.ProductName)
    
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "刷新方式", i, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
            Exit For
        End If
    Next
    
    If Not gobjBillPrint Is Nothing Then
        Call gobjBillPrint.zlTerminate
        Set gobjBillPrint = Nothing
    End If
    
    '撤卸消息对象模块
    Call UnloadMsgModule
End Sub

Private Sub mnuViewGo_Click()
    If tbsType.SelectedItem.Key = "挂号" Then
        frmRegistFind.txtFact.Enabled = True
        frmRegistFind.txtFact.BackColor = Me.vsThis.BackColor
    Else
        frmRegistFind.txtFact.Text = ""
        frmRegistFind.txtFact.Enabled = False
        frmRegistFind.txtFact.BackColor = Me.BackColor
    End If
    If mbytCancel <> 2 Then
        frmRegistFind.lbl操作员.Caption = "挂号员"
    Else
        frmRegistFind.lbl操作员.Caption = "退号员"
    End If
    frmRegistFind.Show 1, Me
    If gblnOk Then Call SeekBill(frmRegistFind.optHead)
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long
    Dim blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "正在定位满足条件的单据,按ESC终止 ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To vsThis.Rows - 1
        DoEvents
        
        '比较条件
        blnFill = True
        With frmRegistFind
            If .txtNO.Text <> "" Then
                blnFill = blnFill And vsThis.TextMatrix(i, getColNum("单据号")) = .txtNO.Text
            End If
            If .txtFact.Text <> "" Then
                blnFill = blnFill And vsThis.TextMatrix(i, getColNum("首张票据")) = .txtFact.Text
            End If
            If .cbo操作员.ListIndex > 0 Then
                If mbytCancel <> 2 Then
                    blnFill = blnFill And vsThis.TextMatrix(i, getColNum("挂号员")) = NeedName(.cbo操作员.Text)
                Else
                    blnFill = blnFill And vsThis.TextMatrix(i, getColNum("退号员")) = NeedName(.cbo操作员.Text)
                End If
            End If
            If .txt姓名.Text <> "" Then
                blnFill = blnFill And UCase(vsThis.TextMatrix(i, getColNum("姓名"))) Like "*" & UCase(.txt姓名.Text) & "*"
            End If
            If IsNumeric(.txt门诊号.Text) Then
                blnFill = blnFill And Val(vsThis.TextMatrix(i, getColNum("门诊号"))) = Val(.txt门诊号.Text)
            End If
        End With
        
        '满足则退出
        If blnFill Then
            mlngGo = i + 1
            vsThis.Row = i: vsThis.TopRow = i
            vsThis.Col = 0: vsThis.ColSel = vsThis.Cols - 1
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

Private Sub InitPara()
    '获取挂号相关参数
     With mTy_Para
            .bln退号审核 = Val(zlDatabase.GetPara("退号审核", glngSys, mlngModul, 0)) = 1
            .lngN天取消预约 = Val(zlDatabase.GetPara("N天内不能取消预约号", glngSys, mlngModul, 0))
            .blnReuseRegNo = Val(zlDatabase.GetPara("已退序号允许挂号", glngSys, mlngModul, 1)) = 1
     End With
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Integer
    '问题号:48911
    If tbsType.SelectedItem.Key = "挂号" Then
            '问题号:51672
            strHead = "医保,4,500|单据号,1,850|首张票据,1,850|号别,1,1600|号序,1,650|科室,1,1000|医生,1,700|病历,4,500|门诊号,1,900|姓名,1,650|就诊卡号,1,1000|手机号,1,1100|费别,1,1000|总金额,7,800|挂号费,7,800|病历费,7,800|挂号时间,4,1800|登记时间,4,1800|退号时间,4," & IIf(mbytCancel = 1, 0, 1800) & "|" & IIf(mbytCancel = 2, "退号员", "挂号员") & ",1,650|收费员,1,650|收费单,1,850|摘要,1,1800|预约时间,1,1800|社区,4,500|病人ID,1,0|记录状态,1,0|退号审核人,1,650|退号审核时间,4,1800|预约操作员,1,650|险类,1,0|记帐费用,1,0"
    Else
            strHead = "停用安排,1,800|单据号,1,850|预约时间,1,1550|号别,1,1600|号序,1,650|科室,1,1000|医生,1,700|病历,4,500|门诊号,1,900|姓名,1,650|身份证号,1,2000|联系电话,1,2650|手机号,1,2650|费别,1,1000|金额,7,800|摘要,1,2000|登记时间,4,1800|挂号员,1,650,|记录状态,2,0|退号审核人,1,650|退号审核时间,4,1800|预约操作员,1,650|病人ID,1,0|险类,1,0|记帐费用,1,0"
    End If
     For i = 0 To Me.vsThis.Cols - 2
        With vsThis
            .FixedAlignment(i) = flexAlignCenterCenter
        End With
    
    Next
    vsThis.FixedAlignment(vsThis.Cols - 2) = flexAlignLeftCenter
    vsThis.ColAlignment(vsThis.Cols - 2) = flexAlignLeftCenter
    With vsThis
        .Redraw = False
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .ColKey(i) = Split(Split(strHead, "|")(i), ",")(0)
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
           ' .ColAlignmentFixed(i) = 4
        Next
      '   .ColHidden(getColNum("记录状态")) = True
        .RowHeight(0) = 320
        .ExtendLastCol = True
        
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
         Call vsThis_EnterCell
         zl_vsGrid_Para_Restore mlngModul, vsThis, Me.Caption, Me.tbsType.SelectedItem.Key, False, InStr(1, mstrPrivs, ";参数设置;") > 0
        .Redraw = True
    End With
End Sub

Private Sub ShowBills(Optional blnSort As Boolean)
'功能:按条件读取单据列表(过滤功能)
    Dim strSQL          As String
    Dim strTime         As String
    Dim i               As Long
    Dim strFilter       As String
    Dim str门诊费用记录  As String
    Dim strDate As String
    Dim strTmp          As String
    Dim blnChange       As Boolean
    Dim strPlanFilter   As String
    blnChange = tbsType.SelectedItem.Key <> mstrVsType
    mstrVsType = tbsType.SelectedItem.Key
    On Error GoTo errH
    
    If Not blnSort Then
        Call zlCommFun.ShowFlash("正在读取挂号数据,请稍候 ...", Me)
        DoEvents
        Me.Refresh
        
        SQLCondition.Default = (mstrFilter = "")
        strFilter = mstrFilter
        
        '问题号:48911
        If tbsType.SelectedItem.Key = "挂号" Then
            '已挂或已接收的号:登记时间在指定范围内的,
            If mstrFilter = "" Then
                '缺省显示当前操作员今天内挂的号
                mbytCancel = 1
                strFilter = " And A.登记时间 Between Trunc(Sysdate) And Trunc(Sysdate)+1-1/24/60/60 And A.操作员姓名||''=[1]"
            End If
            '问题号:49528
            '问题号:51672
              strSQL = "  " & _
                "       Select D.手机号,a.No As 单据号, f.实际票号 As 首张票据, Decode(a.号别, Null, Null, '[' || a.号别 || ']') || Max(Decode(f.序号,1,c3.名称,Null)) As 号别, a.号序, e.名称 As 科室," & vbNewLine & _
                "              a.执行人 As 医生, Decode(Max(Nvl(f.附加标志, 0)), 1, '√', Null) As 病历, a.门诊号, a.姓名, d.就诊卡号, f.费别, " & vbNewLine & _
                "              To_Char(Decode(a.记录状态,2,-1*Sum(Nvl(f.实收金额, 0)),Sum(Nvl(f.实收金额, 0))), '99999999999999990.00') As 金额,To_Char(Decode(a.记录状态,2,-1*Sum(Decode(Sign(Nvl(f.附加标志, 0)), 1, 0, 1) * Nvl(f.实收金额, 0)),Sum(Decode(Sign(Nvl(f.附加标志, 0)), 1, 0, 1) * Nvl(f.实收金额, 0))), '99999999999999990.00') As 挂号费," & _
                "               To_Char(Decode(a.记录状态,2,-1*Sum(Decode(Sign(Nvl(f.附加标志, 0)), 1, 1, 0) * Nvl(f.实收金额, 0)),Sum(Decode(Sign(Nvl(f.附加标志, 0)), 1, 1, 0) * Nvl(f.实收金额, 0))), '99999999999999990.00') As 病例费, a.发生时间 as 挂号时间,a.登记时间,Decode(A.记录状态,2,A.登记时间,Null) As 退号时间 ," & IIf(mbytCancel = 2, "a.操作员姓名 as 退号员", "a.操作员姓名 as 挂号员") & ", a.操作员姓名 As 收费员, a.收费单,a.摘要, " & vbNewLine & _
                "              Decode(a.预约, 1, a.发生时间, Null) As 预约时间, Decode(a.社区, 1, '√', Null) As 社区, a.病人id, a.记录状态, a.退号审核人,A.退号审核时间,a.预约操作员 ," & _
                "              Max(A.险类) as 险类,Max(F.记帐费用) as 记帐费用" & vbNewLine & _
                "       From 病人挂号记录 A, 病人信息 D, 临床出诊记录 B, 临床出诊号源 B1, 挂号安排 B2, 部门表 E, 门诊费用记录 F, 收费项目目录 C, 收费项目目录 C1, 收费项目目录 C2, 收费项目目录 C3 " & vbNewLine & _
                "       Where a.病人id = d.病人id(+) And a.收费单 Is Null And a.出诊记录ID = b.ID(+) And a.号别 = b1.号码(+) And a.号别 = b2.号码(+) And b2.项目id = c2.id(+) And f.收费细目ID=C3.ID(+) And b1.项目id = c1.id(+)  And a.执行部门id = e.Id(+) And b.项目id = c.Id(+) And a.记录性质 = 1 " & " And (e.站点='" & gstrNodeNo & "' Or e.站点 is Null) And " & vbNewLine & _
                "             a.记录状态 = f.记录状态 And a.No = f.No And F.记录性质=4 " & IIf(mbytCancel = 1, " And a.记录状态 = 1 ", IIf(mbytCancel = 2, " And a.记录状态 = 2 ", "")) & strFilter & vbNewLine & _
                "       Group By a.No, f.实际票号, a.号别, a.号序, e.名称, a.执行人, a.门诊号, a.姓名, d.就诊卡号,d.手机号, f.费别, a.发生时间,a.登记时间,Decode(A.记录状态,2,A.登记时间,Null), a.操作员姓名, a.摘要," & _
                "              Decode(a.预约, 1, a.发生时间, Null), a.病人id, a.记录状态, a.社区,a.退号审核人,a.退号审核时间,a.预约操作员,a.操作员姓名, a.收费单" & vbNewLine

              strSQL = strSQL & " Union All " & _
                "       Select D.手机号,a.No As 单据号, f.实际票号 As 首张票据, Decode(a.号别, Null, Null, '[' || a.号别 || ']') || Max(Decode(f.序号,1,c3.名称,Null)) As 号别, a.号序, e.名称 As 科室," & vbNewLine & _
                "              a.执行人 As 医生, Decode(Max(Nvl(f.附加标志, 0)), 1, '√', Null) As 病历, a.门诊号, a.姓名, d.就诊卡号, f.费别, " & vbNewLine & _
                "              To_Char(Sum(Nvl(f.实收金额, 0)), '99999999999999990.00') As 金额,To_Char(Sum(Decode(Sign(Nvl(f.附加标志, 0)), 1, 0, 1) * Nvl(f.实收金额, 0)), '99999999999999990.00') As 挂号费," & _
                "               To_Char(Sum(Decode(Sign(Nvl(f.附加标志, 0)), 1, 1, 0) * Nvl(f.实收金额, 0)), '99999999999999990.00') As 病例费, a.发生时间 as 挂号时间,a.登记时间,Decode(A.记录状态,2,A.登记时间,Null) As 退号时间 ," & IIf(mbytCancel = 2, "a.操作员姓名 as 退号员", "a.操作员姓名 as 挂号员") & ", Decode(Max(f.记录状态),0,Null,Max(f.操作员姓名)) As 收费员, a.收费单,a.摘要, " & vbNewLine & _
                "              Decode(a.预约, 1, a.发生时间, Null) As 预约时间, Decode(a.社区, 1, '√', Null) As 社区, a.病人id, a.记录状态, a.退号审核人,A.退号审核时间,a.预约操作员 ," & _
                "              Max(A.险类) as 险类,Max(F.记帐费用) as 记帐费用" & vbNewLine & _
                "       From 病人挂号记录 A, 病人信息 D, 临床出诊记录 B, 临床出诊号源 B1, 挂号安排 B2, 部门表 E, 门诊费用记录 F, 收费项目目录 C, 收费项目目录 C1, 收费项目目录 C2, 收费项目目录 C3 " & vbNewLine & _
                "       Where a.病人id = d.病人id(+) And a.出诊记录ID = b.ID(+) And a.号别 = b1.号码(+) And a.号别 = b2.号码(+) And b2.项目id = c2.id(+) And f.收费细目ID=C3.ID(+) And b1.项目id = c1.id(+)  And a.执行部门id = e.Id(+) And b.项目id = c.Id(+) And a.记录性质 = 1 " & " And (e.站点='" & gstrNodeNo & "' Or e.站点 is Null) And " & vbNewLine & _
                "             a.收费单 Is Not Null And a.收费单 = f.No And F.记录性质=1 And f.记录状态 <> 2 " & IIf(mbytCancel = 1, " And a.记录状态 = 1 ", IIf(mbytCancel = 2, " And a.记录状态 = 2 ", "")) & strFilter & vbNewLine & _
                "       Group By a.No, f.实际票号, a.号别, a.号序, e.名称, a.执行人, a.门诊号, a.姓名, d.就诊卡号,D.手机号, f.费别, a.发生时间,a.登记时间,Decode(A.记录状态,2,A.登记时间,Null), a.操作员姓名, a.摘要," & _
                "              Decode(a.预约, 1, a.发生时间, Null), a.病人id, a.记录状态, a.社区,a.退号审核人,a.退号审核时间,a.预约操作员 , a.收费单" & vbNewLine

                If frmRegistFilter.mblnDateMoved Then
                      strSQL = strSQL & " Union All " & Replace(strSQL, "门诊费用记录", "H门诊费用记录")
                End If
                
                strSQL = _
                "       Select Decode(Nvl(险类,0),0,'','√') As 医保,单据号, 首张票据, 号别, 号序, 科室,医生, 病历, 门诊号, 姓名, 就诊卡号,手机号, 费别, " & vbNewLine & _
                "              金额,挂号费,病例费,挂号时间,登记时间,退号时间 ," & IIf(mbytCancel = 2, "退号员", "挂号员") & ", 收费员, 收费单, 摘要, " & vbNewLine & _
                "              预约时间, 社区, 病人id, 记录状态, 退号审核人,退号审核时间,预约操作员, 险类,记帐费用" & vbNewLine & _
                "       From (" & strSQL & ")" & _
                "       Order By 单据号 Desc,挂号时间 Desc"
          
        ElseIf tbsType.SelectedItem.Key = "预约" Then
            '已预约的号:在最大可能有效范围内(登记时间>=当前时间-允许预约天数),预约时间在指定范围内的
            If mstrFilter = "" Then
                '缺省显示当前操作员挂的预约时间未失效的单据
                strFilter = " And A.发生时间 Between Trunc(Sysdate) And Trunc(Sysdate)+1-1/24/60/60 + zl_Fun_GetAppointmentDays + Decode(Nvl(B1.预约天数," & gint预约天数 & "),0,15,Nvl(B1.预约天数," & gint预约天数 & "))" & _
                    " And A.操作员姓名||''=[1]"
            End If
            strFilter = Replace(strFilter, "发生时间+0", "发生时间")
              strPlanFilter = Replace(strFilter, "And (F.费别 = [11] or F.费别 is Null)", "")
               strSQL = "" & _
                " Select decode(A.停用,1,'已停','')  As 停用安排, A.单据号,A.预约时间,A.号别,To_Char(A.号序,'99999') 号序,D.名称 As 科室,A.医生,A.病历,A.门诊号,A.姓名,A.身份证号,A.联系电话,A.手机号,A.费别,A.金额,A.摘要,A.登记时间,A.挂号员,A.记录状态 ,a.退号审核人,a.退号审核时间,a.预约操作员 ,A.病人ID" & vbNewLine & _
                " From (  Select Max(M.停用) as 停用,A.NO as 单据号,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI') as 预约时间," & vbNewLine & _
                "                   Decode(A.号别,NULL,NULL,'['||A.号别||']') || Nvl(C.名称,Nvl(C1.名称,C2.名称)) as 号别,A.号序,A.执行人 as 医生,a.记录状态,a.执行部门Id," & vbNewLine & _
                "                   Decode(Max(Decode(F.附加标志,1,1,0)),1,'√',NULL) as 病历," & vbNewLine & _
                "                   A.门诊号,A.姓名,F.费别 as 费别,D.身份证号,D.家庭电话 as 联系电话,D.手机号," & vbNewLine & _
                "                   To_Char(Sum(decode(f.记录状态,2,-1,1)*nvl(f.实收金额,0)), '9999990.00') as 金额," & vbNewLine & _
                "                   A.摘要,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间,A.操作员姓名 as 挂号员,a.退号审核人,a.退号审核时间,a.预约操作员 ,A.病人ID,0 as 险类,0 as 记帐费用" & vbNewLine & _
                "           From 病人挂号记录 A, 病人信息 D,临床出诊记录 B, 临床出诊号源 B1, 挂号安排 B2, 收费项目目录 C, 收费项目目录 C1, 收费项目目录 C2 , 门诊费用记录 F, " & vbNewLine & _
                "               (   Select A.ID,Max(1) as 停用 From  病人挂号记录 A, 病人信息 D,临床出诊记录 B,收费项目目录 C,临床出诊号源 B1 " & vbNewLine & _
                "                    Where   A.出诊记录ID=B.ID And B.号源ID=B1.ID And B.项目ID=C.ID(+) And a.病人id = d.病人id(+) " & vbNewLine & _
                "                               And A.记录性质=2  " & IIf(mbytCancel = 1, " And A.记录状态=1", "") & strPlanFilter & vbNewLine & _
                "                               And A.登记时间>=Sysdate - zl_Fun_GetAppointmentDays - Decode(Nvl(B1.预约天数," & gint预约天数 & "),0,15,Nvl(B1.预约天数," & gint预约天数 & "))" & vbNewLine & _
                "                               And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & vbNewLine & _
                "                               And A.发生时间 between B.停诊开始时间 and B.停诊终止时间 " & vbNewLine & _
                "                     Group by A.ID ) M" & vbNewLine & _
                "           Where A.出诊记录ID=B.ID(+) And a.号别=b1.号码(+) And b1.项目id=c1.id(+) And a.号别=b2.号码(+) And b2.项目id=c2.id(+) And a.病人id = d.病人id(+) And b.项目ID=c.ID(+) " & vbNewLine & _
                "                 And A.NO=F.NO(+)  And A.ID=M.ID(+)  And A.记录性质(+)=2 And F.记录性质(+)=4 and a.记录状态=decode(F.记录状态,0,1,F.记录状态)  " & IIf(mbytCancel = 1, " And a.记录状态=1", "") & strFilter & vbNewLine & _
                "                   And A.登记时间>=Sysdate - zl_Fun_GetAppointmentDays - Decode(Nvl(B1.预约天数," & gint预约天数 & "),0,15,Nvl(B1.预约天数," & gint预约天数 & "))" & " And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & vbNewLine & _
                "           Group by A.NO,A.发生时间,A.号别,a.号序,Nvl(C.名称,Nvl(C1.名称,C2.名称)),a.记录状态,a.执行部门Id," & vbNewLine & _
                "                       A.执行人,A.门诊号,A.姓名,D.身份证号,D.家庭电话,D.手机号,A.摘要,A.登记时间,A.操作员姓名,f.费别,a.退号审核人,a.退号审核时间,a.预约操作员,A.病人ID" & vbNewLine & _
                "            " & vbNewLine & _
                "   ) A, 部门表 D" & vbNewLine & _
                "   Where A.执行部门ID=D.ID " & _
                "   Order by A.登记时间 Desc "
         
            
        ElseIf tbsType.SelectedItem.Key = "接收" Then
            '应接收的号:在最大可能有效范围内,预约在今天的,当前时间在预约号时间段范围之内的
            If mstrFilter = "" Then
                '缺省显示当前操作员当前应该接收的号
                strFilter = " And A.操作员姓名||''=[1]"
            End If
            strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")

            strTime = ""

            '取现在的星期数对应安排的时间段
            strSQL = "Decode(To_Char(SysDate,'D'),'1',B.周日,'2',B.周一,'3',B.周二,'4',B.周三,'5',B.周四,'6',B.周五,'7',B.周六,NULL)"
            
            '号别不能变,变了求不出来
            strSQL = " " & _
            "      Select Decode(a.停用, 1, '已停', '') As 停用安排, a.单据号, a.预约时间, a.号别, To_Char(a.号序, '99999') 号序, d.名称 As 科室, a.医生, a.病历, a.门诊号," & vbNewLine & _
            "             a.姓名 , a.身份证号, a.联系电话,a.手机号, a.费别, a.金额, a.摘要, a.登记时间, a.挂号员,A.记录状态 ,A.退号审核人,a.退号审核时间,a.预约操作员 ,A.病人ID" & vbNewLine & _
            "      From (Select Max(m.停用) As 停用, a.No As 单据号," & vbNewLine & _
            "                   To_Char(a.发生时间, 'YYYY-MM-DD HH24:MI') As 预约时间, Decode(a.号别, Null, Null, '[' || a.号别 || ']') || Nvl(c.名称,Nvl(c1.名称,c2.名称)) As 号别," & vbNewLine & _
            "                   a.号序, a.执行人 As 医生, Decode(Max(Decode(F.附加标志, 1, 1, 0)), 1, '√', Null) As 病历, a.门诊号, a.姓名, d.身份证号," & vbNewLine & _
            "                   d.家庭电话 As 联系电话,D.手机号, F.费别 As 费别, To_Char(Sum(F.实收金额), '9999990.00') As 金额, a.摘要, " & vbNewLine & _
            "                   To_Char(a.登记时间, 'YYYY-MM-DD HH24:MI:SS') As 登记时间, a.操作员姓名 As 挂号员,a.退号审核人,a.退号审核时间,a.记录状态 ,a.预约操作员,A.病人ID,0 as 险类,0 as 记帐费用" & vbNewLine & _
            "            From 病人挂号记录 A, 病人信息 D, 临床出诊记录 B, 临床出诊号源 B1, 挂号安排 B2, 收费项目目录 C, 收费项目目录 C1, 收费项目目录 C2, 门诊费用记录 F," & vbNewLine & _
            "                 (Select A.ID, Max(1) As 停用" & vbNewLine & _
            "                  From 病人挂号记录 A, 病人信息 D, 临床出诊记录 B, 收费项目目录 C " & vbNewLine & _
            "                  Where a.出诊记录ID = b.ID(+) And b.项目id = c.Id(+) And a.病人id = d.病人id(+) And a.记录性质 = 2 And a.记录状态 = 1 And " & vbNewLine & _
             IIf(SQLCondition.Default = False, "  a.发生时间 Between [1] And [2] ", "                        a.发生时间 Between Trunc(Sysdate) And Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60  ") & vbNewLine & _
            "                        And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & vbNewLine & vbNewLine & _
            "                         And a.发生时间 Between B.停诊开始时间 And B.停诊终止时间" & vbNewLine & _
            "                  Group By A.ID) M " & vbNewLine & _
            "           Where a.出诊记录ID = b.ID(+) And a.号别=b1.号码(+) And b1.项目id=c1.id(+) And a.号别=b2.号码(+) And b2.项目id=c2.id(+) And b.项目id = c.Id(+) And a.病人id = d.病人id(+) And a.记录性质 = 2 And a.记录状态 = 1 And a.No = F.No(+) And" & vbNewLine & _
            "                 A.Id = m.ID(+) And F.记录性质=4   " & strFilter & vbNewLine & _
            IIf(strTime = "", "", "                 And " & strSQL & " IN(" & strTime & ")") & vbNewLine & _
            IIf(SQLCondition.Default = False, " and  a.发生时间 Between [1] And [2] ", "                 And a.发生时间 Between Trunc(Sysdate) And Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60 ") & vbNewLine & _
            "                 And (c.站点 = '" & gstrNodeNo & "' Or c.站点 Is Null)" & vbNewLine & _
            "           Group By a.No, a.发生时间, F.费别,a.号别, Nvl(c.名称,Nvl(c1.名称,c2.名称)), a.号序, a.执行人, a.门诊号, a.姓名, d.身份证号, d.家庭电话,d.手机号, a.摘要, a.登记时间, a.操作员姓名,a.退号审核人,a.退号审核时间,a.记录状态,a.预约操作员 ,A.病人ID" & vbNewLine & _
            "           Order By a.发生时间 Desc) A, 病人挂号记录 B, 部门表 D " & vbNewLine & _
            "     Where a.单据号 = b.No And b.记录性质 = 2 And b.记录状态 = 1 And b.执行部门id = d.ID "
        End If
        
        If SQLCondition.Default Then
                Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名)
        Else
            With SQLCondition
                Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .DateB, .DateE, .NOB, .NOE, .Operator, .Patientid, .Doctor, .FactB, .FactE, .DeptID, .FeeType, .ItemType, .PatiName)
            End With
        End If
    End If
    
    vsThis.Clear
    vsThis.Rows = 2
    
    If mrsList.EOF Then
        Call SetHeader
        stbThis.Panels(2).Text = "当前条件下没有任何数据"
        Call SetMenu(False)
    Else
        Set vsThis.DataSource = mrsList
        Call SetHeader
        stbThis.Panels(2) = "共 " & mrsList.RecordCount & " 条数据"
        If tbsType.SelectedItem.Key = "挂号" Then
            stbThis.Panels(2) = stbThis.Panels(2) & ",金额合计:" & Format(GetBillSum, "0.00") & "元(含划价单" & Format(GetHJSum, "0.00") & "元)"
        End If
        Call SetMenu(True)
    End If
    
    If tbsType.SelectedItem.Key = "预约" Then
        Call vsThis_RowColChange
    End If
    Call SetRowColor
    Call SetMenuEnable  '设置菜单
    
    If Not blnSort Then Call zlCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetHJSum() As Currency
    Dim i As Long, lngCol As Long
    If tbsType.SelectedItem.Key = "挂号" Then
        lngCol = getColNum("总金额")
    Else
        lngCol = getColNum("金额")
    End If
    For i = 1 To vsThis.Rows - 1
        If vsThis.TextMatrix(i, getColNum("收费单")) <> "" Then
            GetHJSum = GetHJSum + Val(vsThis.TextMatrix(i, lngCol))
        End If
    Next
End Function

Private Sub vsThis_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strVal As String
    If tbsType.SelectedItem.Key = "接收" Then
        With Me.vsThis
             strVal = .TextMatrix(NewRow, getColNum("停用安排"))
             If strVal = "" Then
                .ForeColorSel = -2147483634
             Else
                .ForeColorSel = vbRed
             End If
        End With
    Else
        vsThis.ForeColorSel = -2147483634
    End If
End Sub

Private Sub vsThis_AfterSort(ByVal Col As Long, Order As Integer)
    Dim strVal As String
    If tbsType.SelectedItem.Key = "接收" Then
        With Me.vsThis
             strVal = .TextMatrix(.Row, getColNum("停用安排"))
             If strVal = "" Then
                .ForeColorSel = -2147483634
             Else
                .ForeColorSel = vbRed
             End If
        End With
    Else
        vsThis.ForeColorSel = -2147483634
    End If
End Sub

Private Function SetRowColor()
    '--------------------------------
    '设置颜色
    '--------------------------------
    Dim X As Long, i As Long, strVal As String
    For X = 1 To Me.vsThis.Rows - 1
          With Me.vsThis
               strVal = .TextMatrix(X, getColNum("记录状态"))
               If strVal = "2" Then
                  .Cell(flexcpForeColor, X, 0, X, .Cols - 1) = &HFF&
               ElseIf strVal = "3" Then
                    .Cell(flexcpForeColor, X, 0, X, .Cols - 1) = &HFF0000
               Else
                    .Cell(flexcpForeColor, X, 0, X, .Cols - 1) = &H80000008
               End If
               
          End With
     Next
    
    If tbsType.SelectedItem.Key = "接收" Then
        For X = 1 To Me.vsThis.Rows - 1
              With Me.vsThis
                   strVal = .TextMatrix(X, getColNum("停用安排"))
                   If strVal <> "" Then
                        .Cell(flexcpForeColor, X, 0, X, .Cols - 1) = vbRed
                   End If
                   If X = 1 Then
                    If strVal <> "" Then
                         .ForeColorSel = vbRed
                    Else
                         .ForeColorSel = -2147483634
                    End If
                   End If
              End With
         Next
    End If
End Function

Private Function GetBillSum() As Currency
    Dim i As Long, lngCol As Long
    If tbsType.SelectedItem.Key = "挂号" Then
        lngCol = getColNum("总金额")
    Else
        lngCol = getColNum("金额")
    End If
    If mbytCancel = 3 Then
        For i = 1 To vsThis.Rows - 1
             If Val(vsThis.TextMatrix(i, vsThis.ColIndex("记录状态"))) <> 2 Then GetBillSum = GetBillSum + Val(vsThis.TextMatrix(i, lngCol))
        Next
    Else
        For i = 1 To vsThis.Rows - 1
            GetBillSum = GetBillSum + Val(vsThis.TextMatrix(i, lngCol))
        Next
    End If
End Function

Private Sub mnuEdit_Print_Click()
    Call PrintBill(0)
End Sub

Private Sub mnuEdit_Print_Supplemental_Click()
    Call PrintBill(1)
End Sub

Private Sub PrintBill(BytMode As Byte)
'功能：当前收款记录重新打印一张票据
'bytMode=0-重打,1-补打
    Dim strNO As String, str挂号时间 As String
    Dim lng结帐ID As Long, lng病人ID As Long, intInsure As Integer
    Dim blnVirtualPrint As Boolean, lngShareUseID As Long
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("单据号"))
    
    If strNO = "" Then
        MsgBox "当前没有记录可以重打票据！", vbExclamation, gstrSysName
        Exit Sub
    End If
    If CheckBillExistReplenishData(strNO) Then
        MsgBox "选择的挂号记录进行了医保补充结算，不允许重打补打操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    str挂号时间 = vsThis.TextMatrix(vsThis.Row, getColNum("挂号时间"))
    
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 4, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
        
    If BytMode = 0 Then
        If Not BillOperCheck(1, vsThis.TextMatrix(vsThis.Row, getColNum("挂号员")), _
            CDate(str挂号时间), "重打") Then Exit Sub
    Else
        If Trim(vsThis.TextMatrix(vsThis.Row, getColNum("首张票据"))) <> "" Then
            MsgBox "当前单据已打印过票据,不能进行补打！", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    
    lng结帐ID = GetBill结帐ID(strNO, 4, lng病人ID)
    intInsure = ExistInsure(strNO)
    If intInsure <> 0 Then
        blnVirtualPrint = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure)
    End If
        
    Dim blnStartFactUseType  As Boolean, strUseType As String
    
    If gblnSharedInvoice Then
        '挂号用门诊票据:42703
        blnStartFactUseType = zlStartFactUseType("1")
        If blnStartFactUseType Then
            strUseType = zl_GetInvoiceUserType(lng病人ID, 0, intInsure)
        End If
    End If
    If Not RePrintBill(Me, IIf(BytMode = 0, 3, 4), strNO, lng结帐ID, intInsure, blnVirtualPrint, strUseType, True) Then Exit Sub
                      
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If

End Sub

 
 

Private Sub mnuViewRefeshOptionItem_Click(index As Integer)
    Dim i As Integer
    For i = 0 To mnuViewRefeshOptionItem.UBound
        mnuViewRefeshOptionItem(i).Checked = i = index
    Next
End Sub

Private Sub tbsType_Click()
    If Visible Then
        Call SaveFlexState(vsThis, App.ProductName & "\" & Me.Name)
    End If
    vsThis.ForeColorSel = -2147483634
    If Val(vsThis.Tag) = tbsType.SelectedItem.index Then Exit Sub
    vsThis.Tag = tbsType.SelectedItem.index
    Call SetHeader
    
    If Visible Then
        Call RestoreFlexState(vsThis, App.ProductName & "\" & Me.Name)
    End If
    
    If Visible Or tbsType.SelectedItem.Key <> "挂号" Then
        '切换清单时恢复缺省值
        Unload frmRegistFilter
        mbytCancel = 1: mstrFilter = ""
        vsThis.Clear 1
        vsThis.Rows = 1
        SetMenu False '问题: 50358
    End If
    
    If Visible Then vsThis.SetFocus
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.Hwnd)
End Sub
 

Private Function CallBackBalanceInterface(ByVal cllBalance As Collection, _
    ByVal lng挂号结帐ID As Long, ByVal lng卡费结帐ID As Long, _
    ByVal dblMoney As Double, _
    ByRef cllUpdate As Collection, _
    ByRef cllThreeSwap As Collection, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用回退接口
    '入参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-13 10:33:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str卡号 As String, strSwapGlideNO As String, strSwapMemo As String, str结算信息 As String, strSwapExtendInfor As String
    Dim varData As Variant, varTemp As Variant, i As Long, cllPro As Collection
    Dim bln消费卡 As Boolean, lng卡类别ID As Long, strSQL As String, rsTemp As ADODB.Recordset
    Dim lng挂号冲销ID As Long, lng退卡冲销ID As Long, strTemp As String
    
    Err = 0: On Error GoTo Errhand:
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO
    'cllBalance.Add Array(Val(Nvl(rsTmp!卡类别ID)), Trim(Nvl(rsTmp!卡号)), IIf(Val(Nvl(rsTmp!结算卡序号)) <> 0, 1, 0), Trim(Nvl(rsTmp!交易流水号)), Trim(Nvl(rsTmp!交易说明))), strNO
    If cllBalance Is Nothing Then CallBackBalanceInterface = True: Exit Function
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO,结帐ID
    bln消费卡 = Val(cllBalance(1)(2)) = 1
    lng卡类别ID = cllBalance(1)(0)
    
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO
    If lng卡类别ID = 0 Then CallBackBalanceInterface = True: Exit Function
    
    str卡号 = cllBalance(1)(1)
    strSwapGlideNO = cllBalance(1)(3)
    strSwapMemo = cllBalance(1)(4)
    If lng卡费结帐ID <> 0 Then str结算信息 = str结算信息 & "||5|" & lng卡费结帐ID
    If lng挂号结帐ID <> 0 Then str结算信息 = str结算信息 & "||4|" & lng挂号结帐ID
    If str结算信息 <> "" Then str结算信息 = Mid(str结算信息, 3)
    
    
    If lng卡费结帐ID <> 0 Then
        strSQL = " Select 结帐ID,记帐费用 From 住院费用记录  Where 记录性质=5 And NO =(Select Max(NO) From 住院费用记录 where 结帐ID=[1] and  记录性质=5  )  and 记录状态=2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng卡费结帐ID)
        If rsTemp.EOF Then
            strErrMsg = "未找到退卡信息，不能继续": Exit Function
        End If
        lng退卡冲销ID = Val(Nvl(rsTemp!结帐ID))
    End If
    
    If lng挂号结帐ID <> 0 Then
        strSQL = "Select 结帐ID From 门诊费用记录  Where 记录性质=4 And NO =(Select Max(NO) From 门诊费用记录 where 结帐ID=[1] and  记录性质=4  )  and 记录状态=2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng挂号结帐ID)
        If rsTemp.EOF Then
            strErrMsg = "未找到退号信息，不能继续": Exit Function
        End If
        lng挂号冲销ID = Val(Nvl(rsTemp!结帐ID))
    End If
    
    '81489,冉俊明,2015-1-22,退费传入冲销ID
    If lng退卡冲销ID <> 0 Then strSwapExtendInfor = strSwapExtendInfor & "||5|" & lng退卡冲销ID
    If lng挂号冲销ID <> 0 Then strSwapExtendInfor = strSwapExtendInfor & "||4|" & lng挂号冲销ID
    If strSwapExtendInfor <> "" Then strSwapExtendInfor = Mid(strSwapExtendInfor, 3)
    strTemp = strSwapExtendInfor
    
    'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, _
        ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
        ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款回退交易
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID:医疗卡类别.ID
    '       strCardNo-卡号
    '       strBalanceIDs-本次支付所涉及的结算ID(这是原结帐ID):
    '                           格式:收费类型(|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(扣款时的交易流水号)
    '       strSwapMemo-交易说明(扣款时的交易说明)
    '       strSwapExtendInfor-传入，本次退费的冲销ID：
    '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       strSwapExtendInfor-传出，交易的扩展信息
    '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
    If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModul, lng卡类别ID, bln消费卡, str卡号, str结算信息, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    If lng退卡冲销ID <> 0 Then
        '问题号:58536
        If Not bln消费卡 Then
            Call zlAddUpdateSwapSQL(False, lng退卡冲销ID, lng卡类别ID, bln消费卡, str卡号, strSwapGlideNO, strSwapMemo, cllUpdate)
        End If
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, lng退卡冲销ID, lng卡类别ID, bln消费卡, str卡号, strSwapExtendInfor, cllThreeSwap)
        End If
    End If
    If lng挂号冲销ID <> 0 Then
        Call zlAddUpdateSwapSQL(False, lng挂号冲销ID, lng卡类别ID, bln消费卡, str卡号, strSwapGlideNO, strSwapMemo, cllUpdate)
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, lng挂号冲销ID, lng卡类别ID, bln消费卡, str卡号, strSwapExtendInfor, cllThreeSwap)
        End If
    End If
    CallBackBalanceInterface = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function


Private Sub InitActionType()
    '-------------------------
    '获取 是否采用了分时段的处理方式
    '判断依据为 挂号安排列表是否有数据
    '-------------------------
    Dim strSQL       As String
    Dim rsTmp        As ADODB.Recordset
    strSQL = "Select 1  dt From  临床出诊记录 Where 是否分时段=1 And Rownum < 2"
    On Error GoTo Hd
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    mactionType = t_普通
    If rsTmp.RecordCount <> 0 Then mactionType = t_时段
    Set rsTmp = Nothing
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub
 
Private Function zlGetInvoiceGroupUseID(ByRef lng领用ID As Long, _
    Optional intNum As Integer = 1, _
    Optional strInvoiceNO As String = "", Optional strUseType As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的领用ID
    '入参:lng领用ID-领用id
    '       intNum-页数
    '       strInvoiceNO-输入的发票号
    '       strUserType-使用类别
    '出参:lng领用ID-领用ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-11-19 16:32:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    lng领用ID = GetInvoiceGroupID(IIf(gblnSharedInvoice, 1, 4), intNum, lng领用ID, glng挂号ID, strInvoiceNO, strUseType)
    If lng领用ID <= 0 Then
        Select Case lng领用ID
            Case 0 '操作失败
            Case -1
                If Trim(strUseType) = "" Then
                    MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "你没有自用和共用的『" & strUseType & "』收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(strUseType) = "" Then
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "本地的共用票据的『" & strUseType & "』收费票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
            Case -3
                MsgBox "当前票据号码不在可用领用批次的有效票据号范围内,请重新输入！", vbInformation, gstrSysName
                Exit Function
        End Select
    End If
    zlGetInvoiceGroupUseID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitMsgModule()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化消息模块
    '编制:刘兴洪
    '日期:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobjMsgModule = New clsMipModule
    Call mobjMsgModule.InitMessage(glngSys, mlngModul, mstrPrivs)
    Call AddMipModule(mobjMsgModule)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub UnloadMsgModule()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:拆卸消息模块
    '编制:刘兴洪
    '日期:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    
    If mobjMsgModule Is Nothing Then Exit Sub
    Call mobjMsgModule.CloseMessage
    Call DelMipModule(mobjMsgModule)
    Set mobjMsgModule = Nothing
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub



