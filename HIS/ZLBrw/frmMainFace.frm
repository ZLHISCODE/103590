VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBrower 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   Caption         =   "标题"
   ClientHeight    =   5355
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   9000
   Icon            =   "frmMainFace.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   OLEDropMode     =   1  'Manual
   Picture         =   "frmMainFace.frx":1CFA
   ScaleHeight     =   5355
   ScaleWidth      =   9000
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock winSock 
      Left            =   3240
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrThis 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6720
      Top             =   840
   End
   Begin VB.Timer tmrUpdateConnect 
      Enabled         =   0   'False
      Left            =   5400
      Top             =   840
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2925
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":309E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTry 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   7410
      ScaleHeight     =   570
      ScaleWidth      =   1485
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   1485
      Begin VB.Label lblTry 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "试用版"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   90
         TabIndex        =   8
         Top             =   60
         Width           =   1305
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   4
         Height          =   525
         Left            =   15
         Top             =   30
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lvwFunc 
      Height          =   3600
      Left            =   3750
      TabIndex        =   4
      Top             =   1260
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   6350
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "功能"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "说明"
         Object.Width           =   35278
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMenu 
      Height          =   3555
      Left            =   30
      TabIndex        =   3
      Top             =   1290
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   6271
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   88
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImgList"
      Appearance      =   1
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9000
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   7740
      NewRow1         =   0   'False
      Child2          =   "TbrUsual"
      MinHeight2      =   330
      Width2          =   525
      NewRow2         =   0   'False
      Visible2        =   0   'False
      Begin MSComctlLib.Toolbar TbrUsual 
         Height          =   330
         Left            =   7935
         TabIndex        =   9
         Top             =   225
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImgUsualBlack"
         _Version        =   393216
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   165
         TabIndex        =   6
         Top             =   30
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "imgToolsStard"
         HotImageList    =   "imgToolsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "预览功能表"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Object.ToolTipText     =   "打印功能表"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "printbar"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "后退"
               Key             =   "Back"
               Object.ToolTipText     =   "后退"
               Object.Tag             =   "后退"
               ImageKey        =   "Back"
               Style           =   5
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "前进"
               Key             =   "Forward"
               Object.ToolTipText     =   "前进"
               Object.Tag             =   "前进"
               ImageKey        =   "Forward"
               Style           =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "向上"
               Key             =   "UpGrade"
               Object.ToolTipText     =   "向上一级"
               Object.Tag             =   "向上"
               ImageKey        =   "UpGrade"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "GotoBar"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "工具"
               Key             =   "工具"
               Object.ToolTipText     =   "外接工具"
               Object.Tag             =   "工具"
               ImageKey        =   "Tool"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "工具1"
                     Object.Tag             =   "工具1"
                     Text            =   "工具1"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageKey        =   "Font"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "字体"
               Key             =   "FontSize"
               Object.ToolTipText     =   "字体"
               Object.Tag             =   "字体"
               ImageKey        =   "Font"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ViewBar"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助主题"
               Object.ToolTipText     =   "帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出应用"
               Object.Tag             =   "退出"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   4995
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   1764
            Picture         =   "frmMainFace.frx":9900
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9075
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList imgToolsHot 
      Left            =   3165
      Top             =   960
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
            Picture         =   "frmMainFace.frx":A194
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":A3AE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":A5C8
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":A7E2
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":A9FC
            Key             =   "UpGrade"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":AC16
            Key             =   "Font"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":AE30
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":B04A
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":B264
            Key             =   "Tool2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":C2F6
            Key             =   "Tool"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolsStard 
      Left            =   2625
      Top             =   960
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
            Picture         =   "frmMainFace.frx":D388
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":D5A2
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":D7BC
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":D9D6
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":DBF0
            Key             =   "UpGrade"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":DE0A
            Key             =   "Font"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":E024
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":E23E
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":E458
            Key             =   "Tool2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainFace.frx":F4EA
            Key             =   "Tool"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   2880
      Top             =   3825
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dgdList 
      Height          =   1050
      Left            =   285
      TabIndex        =   2
      Top             =   4845
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1852
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picVLine 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5940
      Left            =   4215
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5940
      ScaleWidth      =   30
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   30
   End
   Begin VB.Timer TimePass 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5790
      Top             =   210
   End
   Begin MSComctlLib.ImageList ImgUsualBlack 
      Left            =   2670
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgUsualColor 
      Left            =   3240
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
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
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileReLogin 
         Caption         =   "注销(&L)"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuGoto 
      Caption         =   "转到(&G)"
      Begin VB.Menu mnuGotoBack 
         Caption         =   "后退(&B)"
         Enabled         =   0   'False
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuGotoForward 
         Caption         =   "前进(&F)"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuGotoUp 
         Caption         =   "向上一级(&U)"
         Enabled         =   0   'False
         Shortcut        =   ^U
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
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "字体(&F)"
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "小字体(&S)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "中字体(&M)"
            Index           =   1
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "大字体(&L)"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "工具(&T)"
      Begin VB.Menu MnuToolTester 
         Caption         =   "使用SQL速度测试工具(&U)"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu MnuToolIndividuation 
         Caption         =   "使用个性化设置(&I)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolNotify 
         Caption         =   "消息通知(&N)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolShowDisReport 
         Caption         =   "显示停用报表(&P)"
      End
      Begin VB.Menu mnuToolSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolDictonary 
         Caption         =   "字典管理工具(&D)"
      End
      Begin VB.Menu mnuToolMessage 
         Caption         =   "消息收发管理(&M)"
      End
      Begin VB.Menu mnuToolNotice 
         Caption         =   "提醒消息查阅(&R)"
      End
      Begin VB.Menu mnuTooleSelect 
         Caption         =   "系统选项(&S)"
      End
      Begin VB.Menu mnuToolExcel 
         Caption         =   "启动&EXCEL报表"
      End
      Begin VB.Menu mnuToolSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolHistory 
         Caption         =   "清除历史记录(&H)"
      End
      Begin VB.Menu mnuToolOutTool 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolOutToolSet 
         Caption         =   "外接工具设置(&O)"
      End
      Begin VB.Menu mnuToolOutToolExecute 
         Caption         =   "工具(&1)"
         Index           =   0
      End
   End
   Begin VB.Menu mnuRepair 
      Caption         =   "修复(&R)"
      Begin VB.Menu mnuRepairIndividuationClear 
         Caption         =   "清除本机界面异常(&C)"
      End
      Begin VB.Menu mnuRepairComponent 
         Caption         =   "检测安装部件(&T)"
      End
      Begin VB.Menu mnuRepairClientUpdate 
         Caption         =   "客户端修复(&U)"
      End
   End
   Begin VB.Menu History 
      Caption         =   "历史(&O)"
      Visible         =   0   'False
      Begin VB.Menu HistoryItem 
         Caption         =   "隐藏(&D)"
         Index           =   0
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
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "弹出"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuOpen 
         Caption         =   "打开(&O)"
      End
      Begin VB.Menu mnuPopuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopuFontSize 
         Caption         =   "小字体(&S)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuPopuFontSize 
         Caption         =   "中字体(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuPopuFontSize 
         Caption         =   "大字体(&L)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmBrower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCurTime As Date       '当前预升级时间检查点.
Private mblnFirst As Boolean
Private mblnVisible As Boolean
Private mstrTitle As String  '产品标题
Private mblnHide As Boolean '是否显示本窗体
Private mblnRemote As Boolean '是否开启远程
Private Const M_INT_DIRECTORY As Integer = 99                 '分类的缺省图标
Private Const M_INT_MODUL As Integer = 100                    '模块的缺省图标
Private Const M_INT_RPTDISABLED As Integer = 242              '禁用报表图标
Private mobjPreNode As MSComctlLib.Node '左边列表上一次选中的项
Private mblnMouseDown As Boolean
Private mpCenture As POINTAPI
Private mclsAppTool As New zl9AppTool.clsAppTool
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1

Public Property Get frmHide() As Boolean
'检索属性值时使用，位于赋值语句的右边。
' X.编号
    frmHide = mblnHide
End Property

Public Property Get ObjLogin() As Object
'检索属性值时使用，位于赋值语句的右边。
' X.编号
    Set ObjLogin = gobjRelogin
End Property

Public Property Get mobjEmr() As Object
'检索属性值时使用，位于赋值语句的右边。
' X.编号
    Set mobjEmr = gobjRelogin.Emr
End Property

Private Sub cbrThis_Resize()
    Call Form_Resize
End Sub

Private Sub Form_Activate()
    Dim objNode As Node
    Dim StrClickNode As String, strSQL As String
    Dim strCode As String
    Dim lngInstanceNo As Long
    
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
            
    Me.tvwMenu.Nodes.Clear
    Me.lvwFunc.ListItems.Clear
    Me.tbrThis.Buttons("Back").ButtonMenus.Clear
    Me.tbrThis.Buttons("Forward").ButtonMenus.Clear
    Me.lvwFunc.GridLines = True
    With grsMenus
        StrClickNode = 0
        Do While Not .EOF
            
            On Error Resume Next
            If .Fields("模块").Value = 0 Then
                If .Fields("上级") = 0 Then
                    Set objNode = Me.tvwMenu.Nodes.Add(, , "_" & .Fields("编号").Value, .Fields("标题").Value, "K_" & IIf(!图标 = 0, M_INT_DIRECTORY, !图标))
                Else
                    Set objNode = Me.tvwMenu.Nodes.Add("_" & .Fields("上级").Value, 4, "_" & .Fields("编号").Value, .Fields("标题").Value, "K_" & IIf(!图标 = 0, M_INT_DIRECTORY, !图标))
                End If
            Else
                If StrClickNode = 0 Then StrClickNode = .Fields("上级").Value
            End If
            .MoveNext
        Loop
    End With
    If StrClickNode <> 0 Then
        tvwMenu_NodeClick Me.tvwMenu.Nodes("_" & StrClickNode)
    End If
    If Me.tvwMenu.Nodes.Count < 2 Then
        mblnVisible = False
    Else
        mblnVisible = True
    End If
    
    Call LoadUsual
    Call LoadHistory
    '刘兴宏:加载外部工具
    '2007/08/22
    Call LoadOutTools
    
    Call Form_Resize
    
    '此段必须在创建同义词后(因消息通知属于ZlAppTool部件,执行其函数--GetUserInfo时出错)
    MnuToolIndividuation.Checked = IIf(Val(zlDatabase.GetPara("使用个性化风格")) = 0, False, True)
    mnuToolNotify.Checked = IIf(Val(zlDatabase.GetPara("接收邮件消息")) = 0, True, False)
    MnuToolTester.Checked = IIf(GetSetting("ZLSOFT", "公共全局", "SQLTest", 0) = 0, False, True)
    mnuToolShowDisReport.Checked = IIf(Val(zlDatabase.GetPara("显示停用报表")) = 0, False, True)
    mnuToolNotify_Click
    
    stbThis.Panels(2).Text = ""
    stbThis.Panels(3).Text = IIf(gstrNodeName = "-", "", "院区：" & gstrNodeName)
    stbThis.Panels(4).Text = gobjRelogin.DBUser & IIf(gobjRelogin.ServerName = "", "", "@" & gobjRelogin.ServerName) & IIf(zlDatabase.CheckRAC(lngInstanceNo), "(RAC:" & lngInstanceNo & ")", "")
    If stbThis.Panels(5).Tag = "" Then stbThis.Panels(5).Tag = Sys.IP
    stbThis.Panels(5).Text = stbThis.Panels(5).Tag
    stbThis.Panels(6).Text = gstrUserName
    stbThis.Panels(7).Text = gstrDeptName
    Call SetMainForm(Me)
    
    '如果只有一可用模块,则打开
    On Error Resume Next
    With grsMenus
        .Filter = "模块<>0 And 报表=0"
        If Not .EOF Then
            If .RecordCount = 1 Then
                Call AddHistory(!系统 & "," & !模块)
                Call LoadHistory
                .Filter = "模块<>0 And 报表=0"
                Call ExecuteFunc(.Fields("系统").Value, IIf(IsNull(.Fields("部件").Value), "", .Fields("部件").Value), .Fields("模块").Value)
            End If
        End If
        .Filter = 0
    End With
    
    On Error GoTo ErrHand
    
    '启动消息服务平台客户端收发服务
    '------------------------------------------------------------------------------------------------------------------
    If ConnectMip(Me.hwnd) = True Then
        Set mclsMipModule = New zl9ComLib.clsMipModule
        Call mclsMipModule.InitMessage(0, 0, "")
        Call AddMipModule(mclsMipModule)
    End If
    '------------------------------------------------------------------------------------------------------------------
    
    '启动自动提醒服务
    mclsAppTool.CodeMan 0, 5, gcnOracle, Me, gstrDbUser
    If mblnHide Then Me.Hide '是外部调用，隐藏主窗体,by 陈东
    Exit Sub
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    If mblnHide Then Me.Hide '是外部调用，隐藏主窗体,by 陈东
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Static StrPass As String                                '输入密码(Open zlReport.ReportMan )
    Dim objItem As ListItem, BlnExist As Boolean
    
    TimePass.Enabled = False
    If KeyCode = vbKeyF12 And Shift = 7 Then
        StrPass = ""
        Exit Sub
    End If
    
    If KeyCode <> vbKeyReturn Then
        If InStr(1, "1234567890 ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyCode))) <> 0 Then StrPass = StrPass & UCase(Chr(KeyCode))
        
        If StrPass = "OPEN ZLREPORT REPORTMAN" Then
            If OwnerUser(gstrDbUser) Then
                StrPass = ""
                If FindWindow(vbNullString, "报表管理") <> 0 Then Exit Sub
                If MsgBox("您确定要运行自定义报表工具吗？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                Call ExecuteFunc(0, "ZL9REPORT", 99999901)
            End If
        End If
    End If
    TimePass.Enabled = True
End Sub

Private Sub Form_Load()
    Dim lngSize As Long, strTmp As String
    Dim strTag As String, strTitle As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim IntCount As Integer
    
    mblnFirst = True
    gblnHideBtn = True
    On Error Resume Next
    gintCurTheme = GetCurTheme
    strTitle = zlRegInfo("产品标题")
    strTag = ""
    If strTitle <> "" Then
        If InStr(strTitle, "-") > 0 Then
            If Split(strTitle, "-")(1) = "Ultimate" Then
                strTag = "旗舰版"
            ElseIf Split(strTitle, "-")(1) = "Professional" Then
                strTag = "专业版"
            End If
        End If
    End If
    strTitle = Split(strTitle, "-")(0)
    mstrTitle = strTitle & IIf(strTag = "", "", "(" & strTag & ")")
    Me.Caption = gstrUserName & "-(锁屏：Ctrl+Alt+L)"
    Call CheckTools
    Call LoadInitIcon
    RestoreWinState Me
    Call ApplyOEM_Picture(Me, "Icon")
    If zlRegInfo("授权性质") <> "1" Then
        picTry.Visible = True
    End If
    
    IntCount = Val(zlDatabase.GetPara("zlBrwFontSize"))
    Me.mnuViewFontSize(0).Checked = False
    Me.mnuViewFontSize(1).Checked = False
    Me.mnuViewFontSize(2).Checked = False
    Me.mnuViewFontSize(IntCount).Checked = True
    Select Case IntCount
    Case 0
        lngSize = 9
        lvwFunc.ColumnHeaders(1).Width = 2000
    Case 1
        lngSize = 11
        lvwFunc.ColumnHeaders(1).Width = 2400
    Case 2
        lvwFunc.ColumnHeaders(1).Width = 2500
        lngSize = 12
    End Select
    Me.tvwMenu.Font.Size = lngSize
    Me.lvwFunc.Font.Size = lngSize

    Me.WindowState = 2
    
    '设置基准菜单
    菜单基准.功能菜单 = 90000001
    菜单基准.窗口菜单 = 99990001
    菜单基准.其它功能菜单 = 99999901
    菜单基准.分隔菜单 = 99999999
    
    Call CheckWinVersion
    
    '传递数据库活动连接给打印部件
    IniPrintMode gcnOracle, gstrDbUser
    
    strSQL = "Select Nvl(Max(参数值), 0) 参数值 From Zltools.Zloptions Where 参数号 = 24"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断是否允许关闭锁定的导航台")
    gblnShutDown = rsTemp!参数值 = 1
    
    '首先判断会话中是否有消息服务器名称
    'select 参数值 from zloptions where 参数号 =17
    strSQL = "select 参数值 from zloptions where 参数号 =17"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断轮询服务器是否开启")
    If rsTemp.RecordCount = 1 Then
        If NVL(rsTemp!参数值) <> "" Then
            '开了轮询服务器,关闭TIME
            tmrUpdateConnect.Enabled = False
        Else
            '没开轮询服务器,使用TIME进行 预升级检查
            tmrUpdateConnect.Enabled = True
            tmrUpdateConnect.Interval = 30000
            mCurTime = Now
        End If
    Else
        '没开轮询服务器,使用TIME进行 预升级检查
        tmrUpdateConnect.Enabled = True
        tmrUpdateConnect.Interval = 30000
        mCurTime = Now
    End If

    '外部调用的处理,by 陈东
    mblnHide = False
    If gstrCommand <> "" Then Call DoCommand
    
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, 0, 0)
            If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
                MsgBox "zlPlugIn 外挂部件执行 Initialize 时出错：" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.LogInAfter
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            MsgBox "zlPlugIn 外挂部件执行 LogInAfter 时出错：" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        Err.Clear: On Error GoTo 0
    End If
    '获取自动锁屏，并转换为微妙
    glngLockTime = Val(zlDatabase.GetPara("自动锁屏")) * 60 * 1000
    tmrThis.Enabled = Not OS.IsDesinMode And glngLockTime > 0
    '监控键盘消息
    If Not OS.IsDesinMode Then
        glngHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf MyKBHook, 0, App.ThreadID)
    End If
    
    '初始化监听
    InitWinsock
End Sub

Private Sub Form_Resize()
    Dim intTop As Integer, intButton As Integer
    If Me.WindowState = 1 Then Exit Sub
    intTop = IIf(Me.cbrThis.Visible, Me.cbrThis.Height, 0)
    intButton = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    On Error Resume Next
    If mblnVisible Then
        Me.tvwMenu.Visible = True
        If mblnFirst Then Me.picVLine.Left = Me.ScaleWidth / 4
        Me.picVLine.Visible = True
    Else
        picVLine.Visible = False
        tvwMenu.Visible = False
    End If
    Me.picVLine.Top = Me.ScaleTop
    Me.picVLine.Height = Me.ScaleHeight
    If Me.picVLine.Left < 100 Then Me.picVLine.Left = 100
    If Me.picVLine.Left > Me.ScaleWidth - 100 Then Me.picVLine.Left = Me.ScaleWidth - 100
    
    Me.tvwMenu.Left = Me.ScaleLeft
    Me.tvwMenu.Width = Me.picVLine.Left - Me.tvwMenu.Left
    Me.tvwMenu.Top = Me.ScaleTop + intTop
    Me.tvwMenu.Height = Me.ScaleHeight - Me.tvwMenu.Top - intButton
    
    If mblnVisible Then
        Me.lvwFunc.Left = Me.picVLine.Left + Me.picVLine.Width
    Else
        Me.lvwFunc.Left = Me.ScaleLeft
    End If
    Me.lvwFunc.Width = Me.ScaleWidth - Me.lvwFunc.Left
    Me.lvwFunc.Top = Me.ScaleTop + intTop
    Me.lvwFunc.Height = Me.ScaleHeight - Me.lvwFunc.Top - intButton
    Me.lvwFunc.ColumnHeaders(2).Width = Me.lvwFunc.Width - Me.lvwFunc.ColumnHeaders(1).Width
    
    picTry.Left = Me.ScaleWidth - picTry.Width - 120
    Me.Refresh
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim FrmThis As Form, ClsClose As Object, LngErr As Long
    Dim objInsure As Object
    Dim IntCount As Integer
    Dim blnCloaseWin As Boolean
    
    blnCloaseWin = Val(zlDatabase.GetPara("关闭Windows")) <> 0
    '取消键盘消息监听
    If glngHook <> 0 Then
        Call UnhookWindowsHookEx(glngHook)
    End If
    If Not mobjPreNode Is Nothing Then Set mobjPreNode = Nothing
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, 0, 0)
            If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
                MsgBox "zlPlugIn 外挂部件执行 Initialize 时出错：" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.LogOutBefore
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            MsgBox "zlPlugIn 外挂部件执行 LogOutBefore 时出错：" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        Err.Clear: On Error GoTo 0
    End If
    
    On Error Resume Next
    '关闭本部件窗体
    For Each FrmThis In Forms
        If FrmThis.Caption <> Me.Caption Then Unload FrmThis
    Next
    
    '关闭所有部件的窗体
    Err = 0
    LngErr = UBound(gstrObj)
    If Err = 0 Then
        For IntCount = 0 To UBound(gstrObj)
            Set ClsClose = gobjCls(IntCount)
            ClsClose.CloseWindows
        Next
    End If
    Err = 0
    
    '关闭应用工具包部件的窗体
    mclsAppTool.CloseWindows
    '关闭公共部件的窗体
    CloseWindows
    
    Err = 0
    Set objInsure = GetObject("", "zl9Insure.clsInsure")
    
    Call objInsure.Releaseme

    Err = 0
    LngErr = UBound(gstrObj)
    If Err = 0 Then
        For IntCount = 0 To UBound(gstrObj)
            Set gobjCls(IntCount) = Nothing
        Next
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    Call DisConnectMip
    If Not (mclsMipModule Is Nothing) Then
        mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    '------------------------------------------------------------------------------------------------------------------
    Call gobjRelogin.Dispose '需要先卸载对象
    Set gobjRelogin = Nothing
    SaveSetting "ZLSOFT", "公共全局", "SQLTest", 0
    
    Call ShutDown(blnCloaseWin)
    If Not gcnOracle Is Nothing Then
        If gcnOracle.State = 1 Then gcnOracle.Close
        Set gcnOracle = Nothing
    End If
    ReDim Preserve gobjCls(0)
    ReDim Preserve gstrObj(0)
End Sub

Private Sub HistoryItem_Click(Index As Integer)
    Dim str系统 As String, str序号 As String
    str系统 = Split(HistoryItem(Index).Tag, ",")(0)
    str序号 = Split(HistoryItem(Index).Tag, ",")(1)
    With grsMenus
        .Filter = "系统=" & str系统 & " And 模块=" & str序号
        If .RecordCount <> 0 Then
            Call AddHistory(!系统 & "," & !模块)
            Call LoadHistory
            .Filter = "系统=" & str系统 & " And 模块=" & str序号
            Call ExecuteFunc(.Fields("系统").Value, IIf(IsNull(.Fields("部件").Value), "", .Fields("部件").Value), .Fields("模块").Value)
        End If
        .Filter = 0
    End With
End Sub

Private Sub lvwFunc_DblClick()
    If Me.lvwFunc.SelectedItem Is Nothing Then Exit Sub
    With grsMenus
        .Filter = "编号=" & Mid(Me.lvwFunc.SelectedItem.Key, 2)
        If .RecordCount = 0 Then .Filter = 0: Exit Sub
        If .Fields("模块").Value <> 0 Then
            Call AddHistory(!系统 & "," & !模块)
            Call LoadHistory
            .Filter = "编号=" & Mid(Me.lvwFunc.SelectedItem.Key, 2)
            Call ExecuteFunc(.Fields("系统").Value, IIf(IsNull(.Fields("部件").Value), "", .Fields("部件").Value), .Fields("模块").Value)
        Else
            tvwMenu_NodeClick Me.tvwMenu.Nodes("_" & .Fields("编号").Value)
        End If
        .Filter = 0
    End With
End Sub

Private Sub lvwFunc_DragDrop(Source As Control, x As Single, y As Single)
    If TypeOf Source Is Toolbar Then
        Call AddOrDelUsual(Split(Source.Tag, "|")(0), True)
    End If
End Sub

Private Sub lvwFunc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objItem As ListItem
    Dim i As Long
    mblnMouseDown = False
    If Not lvwFunc.HitTest(x, y) Is Nothing Then
        Set objItem = lvwFunc.HitTest(x, y)
        grsMenus.Filter = "编号=" & Mid(objItem.Key, 2)
        If grsMenus.RecordCount = 0 Then grsMenus.Filter = 0: Exit Sub
        If grsMenus.Fields("模块").Value <> 0 Then
            If Not HaveFavorite(grsMenus!系统 & "", grsMenus!模块 & "") Then
                objItem.Selected = True
                mblnMouseDown = Button = 1
            End If
        End If
    End If
End Sub

Private Sub lvwFunc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mblnMouseDown And Button = 1 Then
        lvwFunc.DragIcon = lvwFunc.SelectedItem.CreateDragImage
        lvwFunc.Drag 1
    Else
        Set lvwFunc.DragIcon = Nothing
        lvwFunc.Drag 0
        mblnMouseDown = False
    End If
End Sub

Private Sub lvwFunc_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMouseDown = False
    Set lvwFunc.DragIcon = Nothing
    lvwFunc.Drag 0
End Sub

Private Sub lvwFunc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lvwFunc_DblClick
End Sub

Private Sub mclsMipModule_ConnectStateChanged(ByVal IsConnected As Boolean)
    '连接状态已经变化
    If IsConnected Then
        tmrUpdateConnect.Enabled = False
    Else
        tmrUpdateConnect.Enabled = True
        tmrUpdateConnect.Interval = 30000
        mCurTime = Now
    End If
End Sub

Private Sub mclsMipModule_OpenModule(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strPara As String)
    Call RunModual(lngSystem, lngModule, strPara)
End Sub

Private Sub mclsMipModule_OpenReport(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strPara As String)
    Call RunModual(lngSystem, lngModule, strPara, True)
End Sub

Private Sub mclsMipModule_ReceiveMessage(ByVal strMessageItemKey As String, ByVal strMessageConent As String)
    
    Select Case UCase(strMessageItemKey)
    '--------------------------------------------------------------------------------------------------------------
    Case "ZLHIS_PUB_005"            '产品升级通知
        Call gobjRelogin.UpdateClient
    End Select

End Sub

Private Sub mnuFileExcel_Click()
    MenuPrint grsMenus, 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    MenuPrint grsMenus, 2
End Sub

Private Sub mnuFilePrint_Click()
    MenuPrint grsMenus, 0
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub mnuFileReLogin_Click()
    If MsgBox("你确定要注销吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call ReLogin
End Sub

Private Sub mnuGotoBack_Click()
    If Me.tbrThis.Buttons("Back").ButtonMenus.Count >= 2 Then
        tbrThis_ButtonMenuClick Me.tbrThis.Buttons("Back").ButtonMenus(2)
    End If
End Sub

Private Sub mnuGotoForward_Click()
    If Me.tbrThis.Buttons("Forward").ButtonMenus.Count >= 1 Then
        tbrThis_ButtonMenuClick Me.tbrThis.Buttons("Forward").ButtonMenus(1)
    End If
End Sub

Private Sub mnuGotoUp_Click()
    tvwMenu_NodeClick Me.tvwMenu.SelectedItem.Parent
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    Shell "hh.exe  zl9start.chm", vbNormalFocus
End Sub

Private Sub mnuHelpWebForum_Click()
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuPopuFontSize_Click(Index As Integer)
    mnuViewFontSize_Click (Index)
End Sub

Private Sub mnuRepairClientUpdate_Click()
    If MsgBox("本操作将重新检测本机部件环境，对本机部件环境进行修复，对修复后的所有部件进行重新注册。你确认要进行客户端修复吗？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Call gobjRelogin.UpdateClient(True)
    End If
End Sub

Private Sub mnuRepairComponent_Click()
    '--清空注册表[本机部件]--
    SaveSetting "ZLSOFT", "注册信息", "本机部件", ""
    MsgBox "部件检测完毕，所有改动在重新登录后生效！", vbInformation, gstrSysName
End Sub

Private Sub mnuRepairIndividuationClear_Click()
    Dim strSQL As String, rsTmp As Recordset
    Dim strAnalyseComputer As String
    
    If MsgBox("本操作将清除ZLHIS相关的注册表参数，以及数据库中存储的本人、本机参数，产品相关功能将按参数缺省值运行，你确定要继续吗？", vbYesNo + vbDefaultButton2 + vbQuestion, "清除本机界面异常") = vbYes Then
        strSQL = "Select Distinct 部件 From zlPrograms Where 部件 Is Not Null"
        On Error GoTo ErrHand
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "清除本机界面异常")
        Do While Not rsTmp.EOF
            Call DelWinState(Me, rsTmp!部件 & "")
            rsTmp.MoveNext
        Loop
        strAnalyseComputer = OS.ComputerName
        strSQL = "Zl_zluserparas_Clear('" & gstrDbUser & "','" & strAnalyseComputer & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, gstrSysName)
        MsgBox "清除成功，请关闭程序重新进入，确认是否解决界面异常问题。", vbInformation, "清除本机界面异常"
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuToolDictonary_Click()
    mclsAppTool.CodeMan 0, 1, gcnOracle, Me, gstrDbUser
End Sub

Private Sub mnuToolExcel_Click()
    Dim ObjExcel As Object, strHaveSys As String
    
    If gstrUserName = "" Then
        MsgBox "请为操作员设置对应的用户后再使用本功能！", vbInformation, gstrSysName
        Exit Sub
    End If
    strHaveSys = gobjRelogin.Systems
    On Error Resume Next
    Err = 0
    Set ObjExcel = CreateObject("Zl9Excel.ClsExcel")
    If Err <> 0 Then
        MsgBox "无法创建EXCEL部件，您将不能使用EXCEL报表！", vbInformation, gstrSysName
        Exit Sub
    End If
    Call ObjExcel.CodeMan(0, 0, gcnOracle, Me, gstrDbUser)
    Call ObjExcel.SetHaveSys(strHaveSys)
    Call ObjExcel.ExcelReportMain
    Set ObjExcel = Nothing
End Sub

Private Sub mnuToolHistory_Click()
    Call zlDatabase.SetPara("最近使用模块", "")
    Call LoadHistory
End Sub

Private Sub MnuToolIndividuation_Click()
    MnuToolIndividuation.Checked = MnuToolIndividuation.Checked Xor True
    Call zlDatabase.SetPara("使用个性化风格", IIf(MnuToolIndividuation.Checked, "1", "0"))
    SaveSetting "ZLSOFT", "私有全局\" & gstrDbUser, "使用个性化风格", IIf(MnuToolIndividuation.Checked, "1", "0")
End Sub

Private Sub mnuToolMessage_Click()
    mclsAppTool.CodeMan 0, 2, gcnOracle, Me, gstrDbUser
End Sub

Private Sub mnuTooleSelect_Click()
    mclsAppTool.CodeMan 0, 3, gcnOracle, Me, gstrDbUser, gstrMenuSys
    '获取自动锁屏，并转换为微妙
    glngLockTime = Val(zlDatabase.GetPara("自动锁屏")) * 60 * 1000
    tmrThis.Enabled = Not OS.IsDesinMode And glngLockTime > 0
    If Val(zlDatabase.GetPara("允许远程控制")) <> winSock.LocalPort Then
        Call InitWinsock
    End If
    If mclsAppTool.IsRestart Then
        mclsAppTool.IsRestart = False
        Call ReLogin
    Else
        Call ShutUsual
        Call LoadUsual
    End If
End Sub

Private Sub mnuToolNotice_Click()
    mclsAppTool.CodeMan 0, 6, gcnOracle, Me, gstrDbUser
End Sub

Private Sub mnuToolNotify_Click()
    mnuToolNotify.Checked = Not mnuToolNotify.Checked
    Call zlDatabase.SetPara("接收邮件消息", IIf(mnuToolNotify.Checked, "1", "0"))
    mclsAppTool.CodeMan 0, 4, gcnOracle, Me, gstrDbUser, IIf(mnuToolNotify.Checked = True, "Open", "Close")
End Sub

Private Sub mnuToolOutToolExecute_Click(Index As Integer)
    '刘兴宏:2007/08/22
    '增加对外部工具的执行
    Call ExeCuteToolFile(mnuToolOutToolExecute(Index).Tag)
End Sub
Private Sub ExeCuteToolFile(ByVal strFile As String)
    '-----------------------------------------------------------------------------------
    '功能:执行工具文件
    '参数:strFile-文件名
    '编制:刘兴宏
    '日期:2007/08/22
    '-----------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Err = 0: On Error GoTo ErrHand:
    If objFile.FileExists(strFile) = False Then
        MsgBox "工具文件:" & strFile & vbCrLf & "不存在,可能已被删除,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    Shell strFile, vbNormalFocus
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub mnuToolOutToolSet_Click()
    Dim blnApply As Boolean
    '刘兴宏:2007/08/22
    '增加外部工具的设置
    Call frm工具设置.ShowEdit(Me, blnApply)
    If blnApply = False Then Exit Sub
    Call LoadOutTools
End Sub
Private Function LoadOutTools() As Boolean
    '-----------------------------------------------------------------------------------
    '功能:加载外部工具
    '参数:
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/08/22
    '-----------------------------------------------------------------------------------
    Dim i As Long
    Dim strReg As String, arrTemp As Variant, ArrTool As Variant
    Dim objButton As ButtonMenu
    Err = 0: On Error Resume Next
    '先清除外部工具菜单
    For i = 1 To mnuToolOutToolExecute.UBound
        Unload mnuToolOutToolExecute(i)
    Next
    
    '再清除工具栏
    Do While True
        If tbrThis.Buttons("工具").ButtonMenus.Count = 0 Then Exit Do
        tbrThis.Buttons("工具").ButtonMenus.Remove tbrThis.Buttons("工具").ButtonMenus.Count
    Loop
    tbrThis.Buttons("工具").Style = tbrDefault
    mnuToolOutToolExecute(0).Visible = False
    '加载工具菜单
    strReg = GetSetting("ZLSOFT", "公共全局\TOOLS", "TOOLFILES", "")
    If strReg = "" Then Exit Function
    ArrTool = Split(strReg, "|")
    For i = 0 To UBound(ArrTool)
        arrTemp = Split(ArrTool(i) & ",", ",")
        If arrTemp(0) <> "" And arrTemp(1) <> "" Then
            If i = 0 Then
                With mnuToolOutToolExecute(0)
                    .Caption = arrTemp(0) & "(&1)"
                    .Tag = arrTemp(1)
                    .Visible = True
                End With
            Else
                Load mnuToolOutToolExecute(i)
                With mnuToolOutToolExecute(i)
                    .Caption = arrTemp(0) & IIf(i + 1 > 9, "", "(&" & i + 1 & ")")
                    .Tag = arrTemp(1)
                    .Visible = True
                End With
            End If
            With tbrThis.Buttons("工具").ButtonMenus
                Set objButton = .Add(, "K" & i, arrTemp(0))
                objButton.Tag = arrTemp(1)
            End With
            tbrThis.Buttons("工具").Style = tbrDropdown
        End If
    Next
    LoadOutTools = True
End Function

Private Sub mnuToolShowDisReport_Click()
    mnuToolShowDisReport.Checked = Not mnuToolShowDisReport.Checked
    Call zlDatabase.SetPara("显示停用报表", IIf(mnuToolShowDisReport.Checked, 1, 0))
    '动态调整报表模块显示
    Call tvwMenu_NodeClick(tvwMenu.SelectedItem)
End Sub

Private Sub MnuToolTester_Click()
    '使用SQL速度测试工具(&U)
    MnuToolTester.Checked = MnuToolTester.Checked Xor True
    SaveSetting "ZLSOFT", "公共全局", "SQLTest", IIf(MnuToolTester.Checked, 1, 0)
End Sub

Private Sub mnuViewFontSize_Click(Index As Integer)
    Dim lngSize As Long, IntCount As Integer
    
    For IntCount = 0 To 2
        Me.mnuViewFontSize(IntCount).Checked = (IntCount = Index)
        Me.mnuPopuFontSize(IntCount).Checked = (IntCount = Index)
    Next
    Call zlDatabase.SetPara("zlBrwFontSize", Index)
    Select Case Index
    Case 0
        lngSize = 9
        lvwFunc.ColumnHeaders(1).Width = 2000
    Case 1
        lngSize = 11
        lvwFunc.ColumnHeaders(1).Width = 2400
    Case 2
        lngSize = 12
        lvwFunc.ColumnHeaders(1).Width = 2500
    End Select
    Me.tvwMenu.Font.Size = lngSize
    Me.lvwFunc.Font.Size = lngSize
End Sub

Private Sub mnuViewStatusBar_Click()
    Me.mnuViewStatusBar.Checked = Not Me.mnuViewStatusBar.Checked
    Me.stbThis.Visible = Me.mnuViewStatusBar.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolbarStand_Click()
    Dim IntCount As Integer
    Me.mnuViewToolbarStand.Checked = Not Me.mnuViewToolbarStand.Checked
    Me.cbrThis.Visible = Me.mnuViewToolbarStand.Checked
    If Me.mnuViewToolbarText.Checked Then
        For IntCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(IntCount).Caption = Me.tbrThis.Buttons(IntCount).Tag
        Next
    Else
        For IntCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(IntCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrThis.Refresh
    Form_Resize
End Sub

Private Sub mnuViewToolbarText_Click()
    Dim IntCount As Integer
    Me.mnuViewToolbarText.Checked = Not Me.mnuViewToolbarText.Checked
    If Me.mnuViewToolbarText.Checked Then
        For IntCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(IntCount).Caption = Me.tbrThis.Buttons(IntCount).Tag
        Next
    Else
        For IntCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(IntCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrThis.Refresh
    Form_Resize
End Sub

Private Sub picVline_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.picVLine.Left = Me.picVLine.Left + x
        Form_Resize
    End If
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        mnuFilePreview_Click
    Case "Print"
        mnuFilePrint_Click
    Case "Back"
        mnuGotoBack_Click
    Case "Forward"
        mnuGotoForward_Click
    Case "UpGrade"
        mnuGotoUp_Click
    Case "工具"
        '刘兴宏:2007/08/22
        '问题:加入外部工具
        Call mnuToolOutToolSet_Click
    Case "FontSize"
        PopupMenu Me.mnuViewFont, vbPopupMenuLeftAlign + vbPopupMenuRightButton
    Case "Help"
        mnuHelpTitle_Click
    Case "Quit"
        mnuFileExit_Click
    End Select
    
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim objMenu As ButtonMenu
    If ButtonMenu.Parent.Key = "工具" Then
        '刘兴宏:2007/08/22
        '问题:加入外部工具
        Call ExeCuteToolFile(ButtonMenu.Tag)
        Exit Sub
    End If
    
    If ButtonMenu.Parent.Key = "Back" Then
        For Each objMenu In Me.tbrThis.Buttons("Back").ButtonMenus
            If ButtonMenu.Key = objMenu.Key Then
                ButtonMenu.Visible = False
                Exit For
            End If
            Me.tbrThis.Buttons("Forward").ButtonMenus.Add 1, objMenu.Key, objMenu.Text
        Next
        For Each objMenu In Me.tbrThis.Buttons("Forward").ButtonMenus
            Me.tbrThis.Buttons("Back").ButtonMenus.Remove 1
        Next
    
    Else
        Me.tbrThis.Buttons("Back").ButtonMenus(1).Visible = True
        For Each objMenu In Me.tbrThis.Buttons("Forward").ButtonMenus
            Me.tbrThis.Buttons("Back").ButtonMenus.Add 1, objMenu.Key, objMenu.Text
            If ButtonMenu.Key = objMenu.Key Then
                Exit For
            End If
        Next
        Me.tbrThis.Buttons("Back").ButtonMenus(1).Visible = False
    
        Err = 0
        On Error Resume Next
        For Each objMenu In Me.tbrThis.Buttons("Back").ButtonMenus
            Me.tbrThis.Buttons("Forward").ButtonMenus.Remove 1
        Next
    End If
    
    Me.tbrThis.Buttons("Back").Enabled = (Me.tbrThis.Buttons("Back").ButtonMenus.Count > 1)
    Me.mnuGotoBack.Enabled = (Me.tbrThis.Buttons("Back").ButtonMenus.Count > 1)
    
    Me.tbrThis.Buttons("Forward").Enabled = (Me.tbrThis.Buttons("Forward").ButtonMenus.Count > 0)
    Me.mnuGotoForward.Enabled = (Me.tbrThis.Buttons("Forward").ButtonMenus.Count > 0)
    tvwMenu_NodeClick Me.tvwMenu.Nodes(ButtonMenu.Key)
End Sub

Private Sub tbrThis_DragDrop(Source As Control, x As Single, y As Single)
    '添加常用菜单
    If TypeOf Source Is ListView Then
        Call AddOrDelUsual(Mid(Source.SelectedItem.Key, 2))
        lvwFunc.Drag 0
    End If
End Sub

Private Sub tbrThis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuView, 2
End Sub

Private Sub TbrUsual_ButtonClick(ByVal Button As MSComctlLib.Button)
    With grsMenus
        TbrUsual.Tag = Button.Tag & "|" & Button.Key
        .Filter = "系统=" & Split(Button.Tag, ",")(0) & " And 模块=" & Split(Button.Tag, ",")(1)
        If .RecordCount <> 0 Then
            Call AddHistory(!系统 & "," & !模块)
            Call LoadHistory
            .Filter = "系统=" & Split(Button.Tag, ",")(0) & " And 模块=" & Split(Button.Tag, ",")(1)
            Call ExecuteFunc(.Fields("系统").Value, IIf(IsNull(.Fields("部件").Value), "", .Fields("部件").Value), .Fields("模块").Value)
        End If
        .Filter = 0
    End With
End Sub

Private Sub TbrUsual_DragDrop(Source As Control, x As Single, y As Single)
    '添加常用菜单
    If TypeOf Source Is ListView Then
        Call AddOrDelUsual(Mid(Source.SelectedItem.Key, 2))
    End If
End Sub

Private Sub TbrUsual_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngIdx As Long
    mblnMouseDown = False
    lngIdx = x \ TbrUsual.ButtonWidth + 1
    If lngIdx <= TbrUsual.Buttons.Count Then
        TbrUsual.Tag = TbrUsual.Buttons(lngIdx).Tag & "|" & TbrUsual.Buttons(lngIdx).Key
        TbrUsual.Buttons(lngIdx).Value = tbrPressed
        mblnMouseDown = Button = 1
        '获取中心点
        If mblnMouseDown Then
            mpCenture.x = TbrUsual.Buttons(lngIdx).Left + 0.5 * TbrUsual.ButtonWidth
            mpCenture.y = TbrUsual.Buttons(lngIdx).Top + 0.5 * TbrUsual.ButtonHeight
        End If
    End If
End Sub

Private Sub TbrUsual_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim p As POINTAPI
    Dim strKey As String
    If mblnMouseDown And Button = 1 Then
        If MouseMove(mpCenture.x, mpCenture.y, x, y) Then
            strKey = Split(TbrUsual.Tag, "|")(1)
            p = zlControl.GetCursorPosition
            TbrUsual.DragIcon = ImgUsualBlack.ListImages(strKey).Picture
            TbrUsual.Drag 1
            Call SetCursorPos(p.x, p.y)
        End If
    End If
End Sub

Private Sub TbrUsual_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngIdx As Long
    
    lngIdx = x \ TbrUsual.ButtonWidth + 1
    If lngIdx <= TbrUsual.Buttons.Count Then
        If Button = 2 Then TbrUsual.Buttons(lngIdx).Value = tbrUnpressed
    End If
    If mblnMouseDown Then Set TbrUsual.DragIcon = Nothing
    mblnMouseDown = False
End Sub

Private Function MouseMove(ByVal XCenter As Long, YCenter As Long, x As Single, y As Single) As Boolean
'功能：有时鼠标点击会发生移动
    Dim lngXDif As Long
    Dim lngYDif As Long
    lngXDif = x - XCenter
    lngYDif = y - YCenter
    If Abs(lngXDif) * 2 > TbrUsual.ButtonWidth Or Abs(lngYDif) * 2 > TbrUsual.ButtonHeight Then
        MouseMove = True
    End If
End Function

Private Sub TimePass_Timer()
    Call Form_KeyDown(vbKeyF12, 7)  '清除静态变量
End Sub

Private Sub tmrThis_Timer()
    If Not gblnLock Then
        If TimeToLock Then
            '隐藏界面
            Call LockProg(True)
        End If
    End If
End Sub

Private Sub tvwMenu_Collapse(ByVal Node As MSComctlLib.Node)
    tvwMenu_NodeClick Node
End Sub

Private Sub tvwMenu_DragDrop(Source As Control, x As Single, y As Single)
    If TypeOf Source Is Toolbar Then
        Call AddOrDelUsual(Split(Source.Tag, "|")(0), True)
    End If
End Sub

Private Sub tvwMenu_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim objItem As ListItem
    Dim objMenu As ButtonMenu
    If Not mobjPreNode Is Nothing Then mobjPreNode.Bold = False
    Node.Selected = True
    Set mobjPreNode = Node
    Node.Bold = True
    Me.Caption = gstrUserName & "-" & mstrTitle & "-(锁屏：Ctrl+Alt+L)"
    '后退记录
    Err = 0
    On Error Resume Next
    With Me.tbrThis.Buttons("Back")
        Set objMenu = .ButtonMenus.Add(1, Me.tvwMenu.SelectedItem.Key, Me.tvwMenu.SelectedItem.Text)
        If Err = 0 Then
            Me.tbrThis.Buttons("Forward").ButtonMenus.Clear
            Me.tbrThis.Buttons("Forward").Enabled = False
            Me.mnuGotoForward.Enabled = False
            
            objMenu.Visible = False
            .ButtonMenus(2).Visible = True
            If .ButtonMenus.Count > 8 Then .ButtonMenus.Remove .ButtonMenus.Count
            If .ButtonMenus.Count > 1 Then
                .Enabled = True
                Me.mnuGotoBack.Enabled = True
            End If
        End If
    End With
    
    '向上处理
    If Node.Parent Is Nothing Then
        Me.tbrThis.Buttons("UpGrade").Enabled = False
        Me.mnuGotoUp.Enabled = False
    Else
        Me.tbrThis.Buttons("UpGrade").Enabled = True
        Me.mnuGotoUp.Enabled = True
    End If
    
    Me.lvwFunc.ListItems.Clear
    With grsMenus
        .Filter = "上级=" & Mid(Node.Key, 2)
        .MoveFirst
        Do While Not .EOF
            '模块 = 0表示它是一个分组
            If .Fields("模块").Value = 0 Then
                Set objItem = Me.lvwFunc.ListItems.Add(, "_" & .Fields("编号").Value, .Fields("标题").Value, "K_" & IIf(!图标 = 0, M_INT_DIRECTORY, !图标), "K_" & IIf(!图标 = 0, M_INT_DIRECTORY, !图标))
                objItem.SubItems(1) = .Fields("说明").Value
            Else
                If !报表 = 1 And Val(!是否停用) = 1 Then
                    If mnuToolShowDisReport.Checked Then
                        Set objItem = Me.lvwFunc.ListItems.Add(, "_" & .Fields("编号").Value, .Fields("标题").Value, "K_" & IIf(!图标 = 0, M_INT_MODUL, M_INT_RPTDISABLED), "K_" & IIf(!图标 = 0, M_INT_MODUL, M_INT_RPTDISABLED))
                        objItem.SubItems(1) = .Fields("说明").Value
                    End If
                Else
                    Set objItem = Me.lvwFunc.ListItems.Add(, "_" & .Fields("编号").Value, .Fields("标题").Value, "K_" & IIf(!图标 = 0, M_INT_MODUL, !图标), "K_" & IIf(!图标 = 0, M_INT_MODUL, !图标))
                    objItem.SubItems(1) = .Fields("说明").Value
                End If
            End If
            .MoveNext
        Loop
        .Filter = 0
    End With
    If Me.lvwFunc.ListItems.Count > 0 Then Me.lvwFunc.ListItems(1).Selected = True
    
End Sub

Public Sub MenuPrint(rsMenuList As ADODB.Recordset, intOutMode As Byte)
    '---------------------------------------------------
    '功能：    根据屏幕打印预览
    '参数：    输出方式
    '返回：
    '---------------------------------------------------
    Dim objCol As Column, intCol As Integer
    Dim objPrint As New zlPrintDbGrd

    With dgdList
        For intCol = 1 To .Columns.Count - 1
            .Columns.Remove 0
        Next
        Set objCol = .Columns(0)
        objCol.Caption = "序号"
        objCol.DataField = "ID"
        objCol.Alignment = dbgCenter
        objCol.Width = 500

        Set objCol = .Columns.Add(.Columns.Count)
        objCol.Caption = "功能"
        objCol.DataField = "标题"
        objCol.Alignment = dbgLeft
        objCol.Width = 1400

        Set objCol = .Columns.Add(.Columns.Count)
        objCol.Caption = "说明"
        objCol.DataField = "说明"
        objCol.Alignment = dbgLeft
        objCol.Width = 8000

        .HoldFields
    End With
    rsMenuList.Filter = 0
    Set dgdList.DataSource = rsMenuList

    '----------------------------------------------------

    If rsMenuList.EOF Or rsMenuList.BOF Then Exit Sub
    If InStr(1, Caption, "-") = 0 Then
        objPrint.Title.Text = Caption & "功能清单"
    Else
        objPrint.Title.Text = Mid(Caption, 1, InStr(1, Caption, "-") - 1) & "功能清单"
    End If

    Set objPrint.BodyGrid = dgdList
    Set objPrint.DataSource = rsMenuList

    If intOutMode = 0 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrViewDBGrd objPrint, 1
        Case 2
            zlPrintOrViewDBGrd objPrint, 2
        Case 3
            zlPrintOrViewDBGrd objPrint, 3
        Case Else
        End Select
    Else
        zlPrintOrViewDBGrd objPrint, intOutMode
    End If

    Set dgdList.DataSource = Nothing
End Sub

Public Sub Show任务(ByVal ChildObj As Object, Optional ByVal strCode As String = "", Optional ByVal StrCaption As String = "")
    '
End Sub

Public Sub Shut任务(ByVal ObjFrm As Object)
    '
End Sub

Private Function LoadInitIcon()
    Dim intIcon As Integer
    Dim strIcon As String
    
    strIcon = ","
    With ImgList
        .ListImages.Clear
        .ImageHeight = 32
        .ImageWidth = 32
    End With
    
    With grsMenus
        Do While Not .EOF
            intIcon = IIf(!图标 = 0, M_INT_DIRECTORY, !图标)
            If InStr(1, strIcon, "," & intIcon & ",") = 0 Then
                strIcon = strIcon & intIcon & ","
                ImgList.ListImages.Add , "K_" & intIcon, mclsAppTool.GetIcon(intIcon)
            End If
            .MoveNext
        Loop
        '禁用报表图标
        ImgList.ListImages.Add , "K_" & M_INT_RPTDISABLED, mclsAppTool.GetIcon(M_INT_RPTDISABLED)
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    Set Me.tvwMenu.ImageList = ImgList
    Set Me.lvwFunc.Icons = ImgList
    Set Me.lvwFunc.SmallIcons = ImgList
End Function

Private Sub LoadHistory()
    Dim strValue As String
    Dim str系统 As String, str序号 As String
    Dim arr系统 As Variant, arr序号 As Variant
    Dim int系统_Cur As Integer, int序号_Cur As Integer
    Dim int系统_Max As Integer, int序号_Max As Integer
    
    '将历史记录装入菜单
    Call ClearHistoryMenu
    
    strValue = zlDatabase.GetPara("最近使用模块")
    If UBound(Split(strValue, "|")) < 1 Then Exit Sub
    str系统 = Trim(Split(strValue, "|")(0))
    str序号 = Trim(Split(strValue, "|")(1))
    If str系统 = "" Or str序号 = "" Then Exit Sub
    
    arr系统 = Split(str系统, ",")
    arr序号 = Split(str序号, ",")
    int系统_Max = UBound(arr系统)
    int序号_Max = UBound(arr序号)
    If int系统_Max > 8 Then int系统_Max = 8 '最多八个历史记录
    
    For int系统_Cur = 0 To int系统_Max
        int序号_Cur = int系统_Cur
        If int序号_Cur > int序号_Max Then Exit For
        
        With grsMenus
            .Filter = "系统=" & IIf(arr系统(int系统_Cur) = "", 0, arr系统(int系统_Cur)) & " And 模块=" & arr序号(int序号_Cur)
            If .RecordCount <> 0 Then
                '设置缺省值
                Load HistoryItem(HistoryItem.Count)
                With HistoryItem(HistoryItem.Count - 1)
                    .Caption = grsMenus!标题
                    .Visible = True
                    .Tag = grsMenus!系统 & "," & grsMenus!模块
                End With
            End If
            .Filter = 0
        End With
    Next
    HistoryItem(0).Visible = False
    History.Visible = True
End Sub

Private Sub ClearHistoryMenu()
    Dim MenuItem As Menu
    On Error Resume Next
    
    '删除历史记录菜单
    For Each MenuItem In HistoryItem
        If MenuItem.Index <> 0 Then
            Unload MenuItem
        Else
            MenuItem.Visible = True
        End If
    Next
    History.Visible = False
End Sub

Private Sub LoadUsual()
    Dim str系统 As String, str序号 As String, str图标 As String, str标题 As String
    Dim arr系统 As Variant, arr序号 As Variant, arr图标 As Variant, arr标题 As Variant
    Dim int系统_Cur As Integer, int序号_Cur As Integer, int图标_Cur As Integer, int标题_Cur As Integer
    Dim int系统_Max As Integer, int序号_Max As Integer, int图标_Max As Integer, int标题_Max As Integer
    Dim objButton As Button, strValue As String
    
    '增加常用功能
    strValue = zlDatabase.GetPara("常用功能模块")
    If UBound(Split(strValue, "|")) < 3 Then Exit Sub
    str系统 = Trim(Split(strValue, "|")(0))
    str序号 = Trim(Split(strValue, "|")(1))
    str图标 = Trim(Split(strValue, "|")(2))
    str标题 = Trim(Split(strValue, "|")(3))
    If str系统 = "" Or str序号 = "" Then Exit Sub
    
    arr系统 = Split(str系统, ",")
    arr序号 = Split(str序号, ",")
    arr图标 = Split(str图标, ",")
    arr标题 = Split(str标题, ",")
    int系统_Max = UBound(arr系统)
    int序号_Max = UBound(arr序号)
    int图标_Max = UBound(arr图标)
    int标题_Max = UBound(arr标题)
    
    '增加图标
    For int系统_Cur = 0 To int系统_Max
        int序号_Cur = int系统_Cur
        int图标_Cur = int系统_Cur
        int标题_Cur = int系统_Cur
        If int序号_Cur > int序号_Max Then Exit For
        
        ImgUsualBlack.ImageHeight = 32
        ImgUsualBlack.ImageWidth = 32
        With grsMenus
            .Filter = "系统=" & arr系统(int系统_Cur) & " And 模块=" & arr序号(int序号_Cur)
            If .RecordCount <> 0 Then
                '设置缺省值
                If int图标_Cur <= int图标_Max Then
                    str图标 = arr图标(int图标_Cur)
                Else
                    str图标 = !图标
                End If
                ImgUsualBlack.ListImages.Add , "K" & int系统_Cur, GetPicDisp(str图标)
            End If
            .Filter = 0
        End With
    Next
    
    '增加按钮
    If ImgUsualBlack.ListImages.Count = 0 Then Exit Sub
    TbrUsual.Buttons.Clear
    Set TbrUsual.ImageList = ImgUsualBlack
    For int系统_Cur = 0 To int系统_Max
        int序号_Cur = int系统_Cur
        int标题_Cur = int系统_Cur
        If int序号_Cur > int序号_Max Then Exit For
        
        With grsMenus
            .Filter = "系统=" & arr系统(int系统_Cur) & " And 模块=" & arr序号(int序号_Cur)
            If .RecordCount <> 0 Then
                '设置缺省值
                str系统 = !系统
                str序号 = !模块
                If int标题_Cur <= int标题_Max Then
                    str标题 = arr标题(int标题_Cur)
                Else
                    str标题 = !标题
                End If
                Set objButton = TbrUsual.Buttons.Add()
                objButton.Caption = ""
                objButton.ToolTipText = str标题
                objButton.Tag = str系统 & "," & str序号
                objButton.Image = "K" & int系统_Cur
                objButton.Key = "K" & int系统_Cur
                objButton.Visible = True
            End If
            .Filter = 0
        End With
    Next
    DoEvents
    cbrThis.Bands(2).MinHeight = TbrUsual.Height
    Set cbrThis.Bands(2).Child = TbrUsual
    cbrThis.Bands(2).Visible = True
    DoEvents
End Sub

Private Sub ShutUsual()
    Dim intButton As Integer
    '删除所有常用功能
    
    Set TbrUsual.ImageList = Nothing
    For intButton = 1 To TbrUsual.Buttons.Count
        TbrUsual.Buttons.Remove (1)
    Next
    ImgUsualBlack.ListImages.Clear
    cbrThis.Bands(2).Visible = False
    Call Form_Resize
End Sub

Private Sub CheckTools()
    Dim blnSplit As Boolean         '是否显示分隔条
    '消息收发与EXCEL报表的权限控制：
    '1、如果授权码中含有此功能
    '2、如果该用户拥有此权限
    '3、显示这两个功能
    '其他工具模块仅判断该用户是否拥有此权限
    
    '工具对应说明
    '打印、预览、输出EXCEL  ,10,'导航功能清单','基本'
    'mnuToolDictonary       ,11,'字典管理工具','基本'
    'mnuToolMessage         ,12,'消息收发工具','基本,发送消息'
    'mnuTooleSelect         ,13,'系统选项设置','基本'
    'mnuToolExcel           ,14,'EXCEL报表工具','基本,报表增删,报表计算,所有系统'
    
    Dim intGrant As Integer
    
    '导航功能清单
    mnuFilePrint.Visible = False
    mnuFilePreview.Visible = False
    mnuFileExcel.Visible = False
    tbrThis.Buttons("Print").Visible = False
    tbrThis.Buttons("Preview").Visible = False
    tbrThis.Buttons("printbar").Visible = False
    'Excel报表工具
    mnuToolExcel.Visible = False
    '消息收发工具
    mnuToolMessage.Visible = False
    mnuToolNotify.Visible = False
    '系统选项设置
    mnuTooleSelect.Visible = False
    '字典管理工具
    mnuToolDictonary.Visible = False
    '当然,分隔条一定是要禁止的,只要存在其中一个功能（字典管理、消息收发、EXCEL报表或系统选项），就需要显示分隔条
    blnSplit = False
    
    intGrant = zlRegTool
    If ((intGrant And 4) = 4) Then
        If InStr(1, GetPrivFunc(0, 工具清单.消息收发工具), "基本") <> 0 Then
            mnuToolMessage.Visible = True
            mnuToolNotify.Visible = True
            blnSplit = True
        Else
            Call zlDatabase.SetPara("接收邮件消息", "0")
        End If
    End If
    If ((intGrant And 8) = 8) Then
        If InStr(1, GetPrivFunc(0, 工具清单.EXCEL报表工具), "基本") Then
            mnuToolExcel.Visible = True
            blnSplit = True
        End If
    End If
    '----------------------------------------------------------------------------------------------
    If InStr(1, GetPrivFunc(0, 工具清单.导航功能清单), "基本") Then
        mnuFilePrint.Visible = True
        mnuFilePreview.Visible = True
        mnuFileExcel.Visible = True
        tbrThis.Buttons("Print").Visible = True
        tbrThis.Buttons("Preview").Visible = True
        tbrThis.Buttons("printbar").Visible = False
    End If
    If InStr(1, GetPrivFunc(0, 工具清单.系统选项设置), "基本") Then
        mnuTooleSelect.Visible = True
        blnSplit = True
    End If
    If InStr(1, GetPrivFunc(0, 工具清单.字典管理工具), "基本") Then
        mnuToolDictonary.Visible = True
        blnSplit = True
    End If
    mnuToolSplit2.Visible = blnSplit
End Sub

Public Sub RunModual(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strPara As String, Optional ByVal blnReport As Boolean)
    '------------------------------------------------------------------------------------------------------
    '功能:调用执行报表,此功能是为自动提醒调用而写,by 陈福容
    '参数:lngSys 系统编号;lngModual 模块号
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHand
    
    With grsMenus
        If blnReport Then
            .Filter = "系统=" & lngSys & " AND 模块=" & lngModual & " And 报表=1"
        Else
            .Filter = "系统=" & lngSys & " AND 模块=" & lngModual
        End If
        If .RecordCount = 0 Then .Filter = 0: Exit Sub
        If .Fields("模块").Value <> 0 Then
            Call ExecuteFunc(.Fields("系统").Value, IIf(IsNull(.Fields("部件").Value), "", .Fields("部件").Value), .Fields("模块").Value, strPara)
        End If
        .Filter = 0
    End With
    
ErrHand:
    
End Sub

Public Function GetCommand() As String
    '功能:用于业务部件获取命令行参数,by 陈东
    '参数:无
    GetCommand = gstrCommand
End Function

Private Sub DoCommand()
    '功能：外部调用导航台时，根据传入参数启动业务部件。,by 陈东
    '参数：无
    Dim i As Integer, lngModual As Long
    Dim varCmd As Variant
    On Error GoTo errH
    varCmd = Split(gstrCommand, " ")
    For i = LBound(varCmd) To UBound(varCmd)
        If UCase(varCmd(i)) Like "PROGRAM=*" Then
            lngModual = Val(Split(varCmd(i), "=")(1))
            grsMenus.Filter = "模块=" & lngModual
            If Not grsMenus.EOF Then
                mblnHide = True
                Call RunModual(grsMenus!系统, lngModual, "")
            End If
            grsMenus.Filter = 0
        End If
    Next
    Exit Sub
errH:
    
End Sub

Public Sub UnloadForm()
    '功能：外部调用导航台启动业务部件后，业务部件在退出时，要调用此函数关闭导航台。by 陈东
    '参数：无
    Unload Me
End Sub

Private Sub tmrUpdateConnect_Timer()
    '预升级处理
    If DateAdd("n", -30, Now) >= mCurTime Then '30分钟检查一次
        tmrUpdateConnect.Enabled = False
        Call gobjRelogin.UpdateClient
        mCurTime = Now
        tmrUpdateConnect.Enabled = True
    End If
End Sub

Private Function HaveFavorite(ByVal strSys As String, ByVal strModule As String) As Boolean
'功能：判断指定系统指定模块是否已经添加到常用菜单
    Dim strValue As String, arrTmp As Variant
    Dim arrSys As Variant, arrModule As Variant, arrImage As Variant, arrCaption As Variant
    Dim lngMax As Long, lngCount As Long
    
    strValue = zlDatabase.GetPara("常用功能模块")
    arrTmp = Split(strValue, "|")
    If UBound(arrTmp) < 3 Then HaveFavorite = False: Exit Function
    arrSys = Split(arrTmp(0), ",")
    arrModule = Split(arrTmp(1), ",")
    arrImage = Split(arrTmp(2), ",")
    arrCaption = Split(arrTmp(3), ",")
    lngMax = IIf(UBound(arrSys) > UBound(arrModule), UBound(arrSys), UBound(arrModule))
    lngMax = IIf(lngMax > UBound(arrImage), lngMax, UBound(arrImage))
    lngMax = IIf(lngMax > UBound(arrCaption), lngMax, UBound(arrCaption))
    If lngMax = -1 Then HaveFavorite = False: Exit Function
    On Error Resume Next
    If lngMax >= 9 Then HaveFavorite = True: Exit Function
    For lngCount = 0 To lngMax
        If arrSys(lngCount) = strSys And arrModule(lngCount) = strModule Then
            HaveFavorite = True: Exit Function
        End If
    Next
    Err.Clear: On Error GoTo 0
End Function

Private Sub AddOrDelUsual(ByVal str编号 As String, Optional ByVal blnDel As Boolean)
'功能：将指定的菜单加入或删除常用菜单栏
'参数：str编号=添加常用时，为菜单编号，删除常用时，为菜单对应的[系统,模块]
'          blnDel=Ture-删除常用;False-添加常用
    Dim arrTmp As Variant, strValue As String
    Dim arrSys As Variant, arrModule As Variant, arrImage As Variant, arrCaption As Variant
    Dim strSys As String, strModules As String, strImages As String, strCaptions As String
    Dim lngMax As Long, lngCount As Long
    
    strValue = zlDatabase.GetPara("常用功能模块")
    arrTmp = Split(strValue, "|")
    If blnDel Then
        If UBound(arrTmp) < 3 Then
            ReDim arrTmp(3)
        Else
            arrSys = Split(arrTmp(0), ","): arrModule = Split(arrTmp(1), ",")
            arrImage = Split(arrTmp(2), ","): arrCaption = Split(arrTmp(3), ",")
            lngMax = IIf(UBound(arrSys) > UBound(arrModule), UBound(arrSys), UBound(arrModule))
            lngMax = IIf(lngMax > UBound(arrImage), lngMax, UBound(arrImage))
            lngMax = IIf(lngMax > UBound(arrCaption), lngMax, UBound(arrCaption))
            If lngMax <> -1 Then
                On Error Resume Next
                For lngCount = 0 To lngMax
                    If arrSys(lngCount) & "," & arrModule(lngCount) <> str编号 Then
                        strSys = strSys & "," & arrSys(lngCount)
                        strModules = strModules & "," & arrModule(lngCount)
                        strImages = strImages & "," & arrImage(lngCount)
                        strCaptions = strCaptions & "," & arrCaption(lngCount)
                    End If
                Next
                Err.Clear: On Error GoTo 0
            End If
            arrTmp(0) = Mid(strSys, 2): arrTmp(1) = Mid(strModules, 2)
            arrTmp(2) = Mid(strImages, 2): arrTmp(3) = Mid(strCaptions, 2)
        End If
    Else
        With grsMenus
            .Filter = "编号=" & str编号
            If .RecordCount = 0 Then Exit Sub
            If UBound(arrTmp) < 3 Then
                ReDim arrTmp(3)
                arrTmp(0) = !系统 & ""
                arrTmp(1) = !模块 & ""
                arrTmp(2) = !图标 & ""
                arrTmp(3) = !标题 & ""
            Else
                arrTmp(0) = arrTmp(0) & IIf(arrTmp(0) <> "", ",", "") & !系统 & ""
                arrTmp(1) = arrTmp(1) & IIf(arrTmp(1) <> "", ",", "") & !模块 & ""
                arrTmp(2) = arrTmp(2) & IIf(arrTmp(2) <> "", ",", "") & !图标 & ""
                arrTmp(3) = arrTmp(3) & IIf(arrTmp(3) <> "", ",", "") & !标题 & ""
            End If
            .Filter = 0
        End With
    End If
    strValue = arrTmp(0) & "|" & arrTmp(1) & "|" & arrTmp(2) & "|" & arrTmp(3)
    Call zlDatabase.SetPara("常用功能模块", strValue)
    Call ShutUsual
    Call LoadUsual
End Sub

Public Function CloseChildWindows(ByVal frmMain As Object) As Boolean
     '功能:关闭所有子窗口
    Dim FrmThis     As Form, ClsClose As Object, IntCount As Integer, LngErr As Long
    Dim objInsure   As Object
    Dim blnOK       As Boolean
    
    On Error Resume Next
    blnOK = True
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, 0, 0)
            If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
                blnOK = False
                MsgBox "zlPlugIn 外挂部件执行 Initialize 时出错：" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.LogOutBefore
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            blnOK = False
            MsgBox "zlPlugIn 外挂部件执行 LogOutBefore 时出错：" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        Err.Clear: On Error GoTo 0
    End If
    On Error Resume Next
    For Each FrmThis In Forms
        If FrmThis.Caption <> frmMain.Caption Then Unload FrmThis
    Next
    '关闭所有部件的窗体
    If Err.Number <> 0 Then Err.Clear
    LngErr = UBound(gstrObj)
    If Err.Number = 0 Then
        For IntCount = 0 To LngErr
            Set ClsClose = gobjCls(IntCount)
            blnOK = blnOK And ClsClose.CloseWindows
            Set gobjCls(IntCount) = Nothing
        Next
    End If
    '关闭应用工具包部件的窗体
    blnOK = blnOK And mclsAppTool.CloseWindows
    '关闭公共部件的窗体
    blnOK = blnOK And CloseWindows
    Set objInsure = GetObject("", "zl9Insure.clsInsure")
    Call objInsure.Releaseme
    If Err.Number <> 0 Then Err.Clear
    CloseChildWindows = blnOK
End Function

Public Function GetPicDisp(Optional ByVal intIcon As Long = 0, Optional ByVal Bln模块 As Boolean = True) As IPictureDisp
    '编制人:朱玉宝
    '编制日期:2000-12-12
    '得到图片对象

    On Error Resume Next
    If intIcon = 0 Then intIcon = IIf(Bln模块, -5, -4)
    Select Case intIcon
    Case -1
        Set GetPicDisp = LoadResPicture("HELP", 1)
    Case -2
        Set GetPicDisp = LoadResPicture("RELOGIN", 1)
    Case -3
        Set GetPicDisp = LoadResPicture("EXIT", 1)
    Case -4
        Set GetPicDisp = LoadResPicture("DIRECTORY", 1)
    Case -5
        Set GetPicDisp = LoadResPicture("MODUL", 1)
    Case Else
        Set GetPicDisp = mclsAppTool.GetIcon(intIcon)
    End Select
End Function



Private Sub InitWinsock()
'功能:获取参数,初始化服务器
    Dim lngPort As Long
            
    On Error Resume Next
    
    lngPort = Val(zlDatabase.GetPara("允许远程控制"))
    mblnRemote = Not lngPort = -1
    winSock.Tag = "1"
    With winSock
        If mblnRemote Then
            .LocalPort = IIf(Val(lngPort) = 0, "1001", Val(lngPort))
            .Listen
        Else
            If .State <> sckClosed Then .Close
        End If
    End With
    winSock.Tag = ""
End Sub

Private Sub winSock_Close()
    If winSock.Tag = "" Then
        If winSock.State <> sckClosed And mblnRemote Then winSock.Close: winSock.Listen  '重新监听
    End If
End Sub

Private Sub winSock_ConnectionRequest(ByVal requestID As Long)
    If winSock.State <> sckClosed Then winSock.Close
    winSock.Accept requestID
End Sub

Private Sub winSock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim strMsg  As String
    
    winSock.GetData strData
    
    On Error GoTo errH
    If strData = "请求远程" Then
                RunCommand "REG ADD HKLM\SYSTEM\CurrentControlSet\Control\Terminal"" ""Server /v fDenyTSConnections /t REG_DWORD /d 0 /f"
                winSock.SendData "YES"
    End If

    Exit Sub
errH:
    MsgBox Err.Description
End Sub

Private Sub winSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    winSock.Close: winSock.Listen
    If winSock.Tag = "" Then
        Select Case Number
            Case 10053
                MsgBox "由于长时间没有操作，连接自动中断。", vbInformation, gstrSysName
            Case Else
                MsgBox Number & Description, vbInformation, gstrSysName
         End Select
    Else
        winSock.Tag = ""
    End If
End Sub
