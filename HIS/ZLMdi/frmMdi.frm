VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMdi 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "#"
   ClientHeight    =   6570
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   9990
   Icon            =   "frmMdi.frx":0000
   KeyPreview      =   -1  'True
   Moveable        =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6570
   ScaleWidth      =   9990
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock winSock 
      Left            =   5040
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrUpdateConnect 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ImageList ImgUsualBlack 
      Left            =   30
      Top             =   1170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgUsualColor 
      Left            =   600
      Top             =   1170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox PicBackBitmap 
      AutoRedraw      =   -1  'True
      Height          =   585
      Left            =   60
      Picture         =   "frmMdi.frx":1CFA
      ScaleHeight     =   525
      ScaleWidth      =   1605
      TabIndex        =   6
      Top             =   1740
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Timer TimePass 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ImageList ImgBlack 
      Left            =   30
      Top             =   630
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
            Picture         =   "frmMdi.frx":DEB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":E0CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":E2E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":E600
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":E91A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":EC34
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":F32E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":F548
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":F762
            Key             =   "Tool"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgColor 
      Left            =   600
      Top             =   630
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
            Picture         =   "frmMdi.frx":107F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":10A0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":10C28
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":1107A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":114CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":1191E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":12018
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":12232
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":1244C
            Key             =   "Tool"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9990
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinWidth1       =   5295
      MinHeight1      =   720
      Width1          =   1425
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "TbrUsual"
      MinWidth2       =   1200
      MinHeight2      =   330
      Width2          =   675
      NewRow2         =   0   'False
      Visible2        =   0   'False
      Begin MSComctlLib.Toolbar TbrUsual 
         Height          =   330
         Left            =   8700
         TabIndex        =   7
         Top             =   225
         Width           =   1200
         _ExtentX        =   2117
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
         TabIndex        =   4
         Top             =   30
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   1270
         ButtonWidth     =   1455
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImgBlack"
         HotImageList    =   "ImgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "printbar"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "字典管理"
               Key             =   "Dictionary"
               Object.ToolTipText     =   "字典管理"
               Object.Tag             =   "字典管理"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "消息收发"
               Key             =   "Message"
               Object.ToolTipText     =   "消息收发"
               Object.Tag             =   "消息收发"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "系统选项"
               Key             =   "Choose"
               Object.ToolTipText     =   "界面选择"
               Object.Tag             =   "界面选择"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "bar"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "检测部件"
               Key             =   "Check"
               Object.ToolTipText     =   "检测部件"
               Object.Tag             =   "检测部件"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "外接工具"
               Key             =   "工具"
               Object.ToolTipText     =   "外接工具设置"
               Object.Tag             =   "工具"
               ImageKey        =   "Tool"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   8
            EndProperty
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   4770
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LvwList 
      Height          =   5475
      Left            =   30
      TabIndex        =   2
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9657
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   8421504
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TvwMenu 
      Height          =   2745
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4842
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSDataGridLib.DataGrid dgdList 
      Height          =   1050
      Left            =   285
      TabIndex        =   0
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   6216
      Width           =   9984
      _ExtentX        =   17621
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2223
            MinWidth        =   1764
            Picture         =   "frmMdi.frx":134DE
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11774
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
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Image ImgTry 
      Height          =   675
      Left            =   4020
      Top             =   3060
      Width           =   1125
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      NegotiatePosition=   1  'Left
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
         Caption         =   "注销(&R)"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuOper 
      Caption         =   "操作(&O)"
      Begin VB.Menu mnuOperDefault 
         Caption         =   "Default"
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "工具(&T)"
      Begin VB.Menu mnuOrderMenu 
         Caption         =   "横向排列功能菜单(&L)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolTester 
         Caption         =   "使用SQL速度测试工具(&U)"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu MnuToolIndividuation 
         Caption         =   "使用个性化设置(&I)"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuToolNotify 
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
      Begin VB.Menu HistoryItem 
         Caption         =   "空"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "窗口(&W)"
      Begin VB.Menu mnuWindowList 
         Caption         =   "重排窗口(&L)"
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
   Begin VB.Menu MnuRightMenu 
      Caption         =   "右键菜单"
      Visible         =   0   'False
      Begin VB.Menu MnuRightAbout 
         Caption         =   "关于(&A)"
      End
      Begin VB.Menu MnuRightBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRightTester 
         Caption         =   "使用SQL速度测试工具(&U)"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRightIndividuation 
         Caption         =   "使用个性化设置(&I)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRightNotify 
         Caption         =   "消息通知(&N)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRightShowDisReport 
         Caption         =   "显示停用报表(&P)"
      End
      Begin VB.Menu MnuRightBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRightDictonary 
         Caption         =   "字典管理工具(&D)"
      End
      Begin VB.Menu mnuRightMessage 
         Caption         =   "消息收发管理(&M)"
      End
      Begin VB.Menu mnuRightNotice 
         Caption         =   "提醒消息查阅(&T)"
      End
      Begin VB.Menu MnuRightStyle 
         Caption         =   "系统选项(&S)"
      End
      Begin VB.Menu MnuRightExcel 
         Caption         =   "启动&EXCEL报表"
      End
      Begin VB.Menu MnuRightBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRightIndividuationClear 
         Caption         =   "清除本机界面异常(&C)"
      End
      Begin VB.Menu MnuRightComponent 
         Caption         =   "检测安装部件(&C)"
      End
      Begin VB.Menu mnuRightClientUpdate 
         Caption         =   "客户端修复(&U)"
      End
      Begin VB.Menu MnuRightHistory 
         Caption         =   "清除历史记录(&H)"
      End
      Begin VB.Menu MnuRightBar5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRightSetColor 
         Caption         =   "设置字体颜色(&O)"
      End
      Begin VB.Menu MnuRightBackBmp 
         Caption         =   "选择背景图片(&B)"
      End
      Begin VB.Menu MnuRightBar6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRightReLogin 
         Caption         =   "注销(&R)"
      End
      Begin VB.Menu MnuRightExit 
         Caption         =   "退出(&X)"
      End
   End
End
Attribute VB_Name = "frmMdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCurTime As Date                        '当前预升级时间检查点.
Private mblnFirst As Boolean
Private mlngMainMenu As Long                    '本窗体的菜单体系的句柄
Private mblnRemote As Boolean '是否开启远程
'----附加说明----
'    请修改人员认真遵守
'1.功能菜单的菜单ID从(90000001|10001)开始
'2.窗口菜单的菜单ID从(99990001|30001)开始(窗口菜单下的人为增加的菜单,用来显示当前已打开模块的动态菜单)
'3.其它功能的菜单的菜单ID从(99999901|65001)开始
'4.只会存在一个分隔条菜单,即窗口菜单下(99999999|65535)
Private mblnHide As Boolean '是否显示本窗体
Private mclsAppTool As New zl9AppTool.clsAppTool
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1

Public Property Get frmHide() As Boolean
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
    Set mobjEmr = gobjRelogin.EMR
End Property

Private Sub cbrThis_Resize()
    Call Form_Resize
End Sub

Private Sub Form_Activate()
    Dim strSQL As String
    Dim lngInstanceNo As Long
    
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
    
    mnuOrderMenu.Checked = Val(zlDatabase.GetPara("zlMdiMenuArray")) <> 0
    mnuToolShowDisReport.Checked = IIf(Val(zlDatabase.GetPara("显示停用报表")) = 0, False, True)
    mnuRightShowDisReport.Checked = mnuToolShowDisReport.Checked
    If Not mnuOrderMenu.Checked Then
        Call LoadMenuPortrait
    Else
        Call LoadMenuLandscape
    End If
    
    '此段必须在创建同义词后(因消息通知属于ZlAppTool部件,执行其函数--GetUserInfo时出错)
    MnuToolIndividuation.Checked = IIf(Val(zlDatabase.GetPara("使用个性化风格")) = 0, False, True)
    MnuToolNotify.Checked = IIf(Val(zlDatabase.GetPara("接收邮件消息")) = 0, True, False)
    mnuToolTester.Checked = IIf(GetSetting("ZLSOFT", "公共全局", "SQLTest", 0) = 0, False, True)
    mnuRightTester.Checked = mnuToolTester.Checked
    mnuRightIndividuation.Checked = MnuToolIndividuation.Checked
    MnuToolNotify_Click
    mnuRightNotify.Checked = MnuToolNotify.Checked
    
    Me.stbThis.Panels(2).Text = ""
    stbThis.Panels(3).Text = IIf(gstrNodeName = "-", "", "院区：" & gstrNodeName)
    Me.stbThis.Panels(4).Text = gobjRelogin.DBUser & IIf(gobjRelogin.ServerName = "", "", "@" & gobjRelogin.ServerName) & IIf(zlDatabase.CheckRAC(lngInstanceNo), "(RAC:" & lngInstanceNo & ")", "")
    Me.stbThis.Panels(5).Text = gstrUserName
    Me.stbThis.Panels(6).Text = gstrDeptName
    Call SetMainForm(Me)                                '初始化公共部件的主窗体
    Call InitEvn
    Call LoadUsual
    Call LoadHistory
    
    '刘兴宏:加载外部工具
    '2007/08/22
    Call LoadOutTools
    
    Call Form_Resize
    
    '如果只有一可用模块,则打开
    On Error Resume Next
    With grsMenus
        .Filter = "模块<>0 And 报表=0"
        If Not .EOF Then
            If .RecordCount = 1 Then
                Dim LngFind As Long, lngMenu As Long
                For LngFind = 0 To CollMenu.Count - 1
                    If CollMenu("K_" & LngFind)(Menu_Modul) = !模块 Then
                        lngMenu = CollMenu("K_" & LngFind)(Menu_ID)
                        Exit For
                    End If
                Next
                
                If lngMenu <> 0 Then Call MenuProc(Me.hwnd, WM_COMMAND, lngMenu, 0)
            End If
        End If
        .Filter = 0
    End With
    
    Call CheckWinVersion
    
    '启动消息服务平台客户端收发服务
    '------------------------------------------------------------------------------------------------------------------
    If ConnectMip(Me.hwnd) = True Then
        Set mclsMipModule = New zl9ComLib.clsMipModule
        Call mclsMipModule.InitMessage(0, 0, "")
        Call AddMipModule(mclsMipModule)
    End If
    
    '启动自动提醒服务
    mclsAppTool.CodeMan 0, 5, gcnOracle, Me, gstrDbUser
    If mblnHide Then Me.Hide '是外部调用，隐藏主窗体,by 陈东
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Static StrPass As String                                '输入密码(Open zlReport.ReportMan )
    Dim LngFind As Long, BlnExist As Boolean, LngUpperMenu As Long, lngMenu As Long

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
                Call ExecuteFunc(0, "ZL9REPORT", 菜单基准.其它功能菜单)
                SetParent FindWindow(vbNullString, "报表管理"), Me.hwnd
            End If
        End If
    End If
    TimePass.Enabled = True
End Sub

Private Sub Form_Load()
    Dim intGrant As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTitle As String, strTag As String
    mblnFirst = True
    
    On Error Resume Next
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
    strTitle = Split(strTitle, "-")(0) & IIf(strTag = "", "", "(" & strTag & ")")
    
    '判断是否有权限使用消息收发功能
    Me.Caption = gstrUserName & "-" & strTitle
    Call CheckTools
    RestoreWinState Me
    Call ApplyOEM_Picture(Me, "Icon")
    
    '获取窗体最小值
    gLngMinH = Screen.Height - 400
    gLngMinW = Screen.Width
    gLngMaxH = gLngMinH
    gLngMaxW = gLngMinW
    
    Dim LngHdl As Long
'    '取系统窗体句柄
    Me.Width = gLngMinW
    Me.Height = gLngMinH
    
    Call InitEvn
    
    '检测系统
    LngHdl = GetSubMenu(GetMenu(Me.hwnd), 0)
    Call InsertMenu(LngHdl, MF_BYPOSITION, MF_STRING, 99999999, "测试菜单(&T)")
    Me.Tag = GetMenuItemID(LngHdl, GetMenuItemCount(LngHdl) - 1)
    Call DeleteMenu(LngHdl, GetMenuItemCount(LngHdl) - 1, MF_BYPOSITION)
    Call SetMenu(Me.hwnd, GetMenu(Me.hwnd))
    Call DrawMenuBar(Me.hwnd)
    
    If Me.Tag <> 99999999 Then
        菜单基准.功能菜单 = 10001
        菜单基准.窗口菜单 = 30001
        菜单基准.其它功能菜单 = 65001
        菜单基准.分隔菜单 = 65535
    Else
        菜单基准.功能菜单 = 90000001
        菜单基准.窗口菜单 = 99990001
        菜单基准.其它功能菜单 = 99999901
        菜单基准.分隔菜单 = 99999999
    End If
    
    '传递数据库活动连接给打印部件
    IniPrintMode gcnOracle, gstrDbUser
    
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

    '初始化监听
    InitWinsock
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.Top = 0
    Me.Left = 0
    
    With LvwList
        .Top = cbrThis.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - stbThis.Height - .Top
    End With
    With PicBackBitmap
        .Top = cbrThis.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - stbThis.Height - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnCloaseWin As Boolean
    
    blnCloaseWin = Val(zlDatabase.GetPara("关闭Windows")) <> 0
    Set CollOpenWindowHdl = New Collection
    Set CollMenu = New Collection
    
    On Error Resume Next
    '恢复窗体原函数的地址
    Call SetWindowLong(Me.hwnd, GWL_WNDPROC, LngAddFunc)
    '清理外挂医保，以及业务窗体
    Call CloseChildWindows(Me)
    '清理消息对象
    Call DisConnectMip
    If Not (mclsMipModule Is Nothing) Then
        mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    Call gobjRelogin.Dispose '需要先卸载对象
    Set gobjRelogin = Nothing
    SaveSetting "ZLSOFT", "公共全局", "SQLTest", 0
    '清除缓存的参数值
    zlDatabase.ClearParaCache
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
    Debug.Print str系统 & ";" & str序号
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

Private Sub LvwList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu MnuRightMenu, 2
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

    If intOutMode = 1 Then
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

Private Sub LoadMenuPortrait()
    Dim objNode As Node
    Dim LngMenuID As Long                           '菜单ID号
    Dim LngLoop As Long                             '循环变量
    Dim LngInsertMenu As Long                       '弹出菜单句柄
    Dim LngUpperMenu As Long                        '上级菜单句柄
    Dim StrHotKey As String                         '快捷键
    On Error Resume Next
    '纵向排列功能菜单
    '--菜单ID由90000001开始--
    
    LngMenuID = 菜单基准.功能菜单
    Set CollMenu = New Collection                   '保存添加菜单的相关信息
    TvwMenu.Nodes.Clear
    mlngMainMenu = GetMenu(Me.hwnd)
    mlngMainMenu = GetSubMenu(mlngMainMenu, 1)        '获取操作子菜单
    
    With grsMenus
        Do While Not .EOF
            Err = 0
            If .Fields("上级") = 0 Then
                Set objNode = Me.TvwMenu.Nodes.Add(, , "_" & !编号, !短标题)
            Else
                Set objNode = Me.TvwMenu.Nodes.Add("_" & !上级, 4, "_" & !编号, !短标题)
            End If
            
            If Err = 0 Then
                    
                '找其上级菜单句柄
                LngUpperMenu = mlngMainMenu
                If Val(!上级) <> 0 Then
                    For LngLoop = 0 To CollMenu.Count - 1
                        If CollMenu("K_" & LngLoop)(Menu_Code) = !上级 Then
                            LngUpperMenu = CollMenu("K_" & LngLoop)(Menu_Hdl)
                            Exit For
                        End If
                    Next
                End If
                
                StrHotKey = UCase(IIf(IsNull(!快键), "", !快键))
                StrHotKey = !短标题 & IIf(StrHotKey = "", "", "(&" & StrHotKey & ")")
                '添加菜单项(如果模块值为零,则为菜单项;否则添加弹出菜单)
                If !模块 = 0 Then
                    LngInsertMenu = CreatePopupMenu()
                    CollMenu.Add Array(LngInsertMenu, .Fields("编号").Value, .Fields("模块").Value, IIf(IsNull(!部件), "", .Fields("部件").Value), LngUpperMenu, StrHotKey, IIf(!模块 = 0, 0, LngMenuID), .Fields("系统").Value), "K_" & CollMenu.Count
                Else
                    If !报表 = 1 And Val(!是否停用) = 1 Then
                        If mnuToolShowDisReport.Checked Then
                            StrHotKey = StrHotKey & "(停用)"
                            LngInsertMenu = InsertMenu(LngUpperMenu, MF_BYPOSITION, MF_STRING, LngMenuID, StrHotKey)
                            CollMenu.Add Array(LngInsertMenu, .Fields("编号").Value, .Fields("模块").Value, IIf(IsNull(!部件), "", .Fields("部件").Value), LngUpperMenu, StrHotKey, IIf(!模块 = 0, 0, LngMenuID), .Fields("系统").Value), "K_" & CollMenu.Count
                            LngMenuID = LngMenuID + 1
                        End If
                    Else
                        LngInsertMenu = InsertMenu(LngUpperMenu, MF_BYPOSITION, MF_STRING, LngMenuID, StrHotKey)
                        CollMenu.Add Array(LngInsertMenu, .Fields("编号").Value, .Fields("模块").Value, IIf(IsNull(!部件), "", .Fields("部件").Value), LngUpperMenu, StrHotKey, IIf(!模块 = 0, 0, LngMenuID), .Fields("系统").Value), "K_" & CollMenu.Count
                        LngMenuID = LngMenuID + 1
                    End If
                End If
            End If
            .MoveNext
        Loop
        .MoveFirst
    End With
    
    '绑定所有弹出菜单到mnuOper菜单下,做为其下级菜单
    Dim IntMenuLocate As Integer
    IntMenuLocate = 1
    For LngLoop = 0 To CollMenu.Count - 1
        If CollMenu("K_" & LngLoop)(Menu_Modul) = 0 Then
            StrHotKey = CollMenu("K_" & LngLoop)(Menu_Caption)              '短标题及快捷键
            Call InsertMenu(CollMenu("K_" & LngLoop)(Menu_UpperHdl), IntMenuLocate, MF_BYPOSITION + MF_POPUP, CollMenu("K_" & LngLoop)(Menu_Hdl), StrHotKey)
            IntMenuLocate = IntMenuLocate + 1
        End If
    Next
    
    '删除缺省菜单
    Call DeleteMenu(mlngMainMenu, 0, MF_BYPOSITION)
    
    '刷新菜单
    Call SetMenu(Me.hwnd, GetMenu(Me.hwnd))
    Call DrawMenuBar(Me.hwnd)

    '设置窗体函数的地址
    LngAddFunc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf MenuProc)
    '恢复主菜单句柄
    mlngMainMenu = GetMenu(Me.hwnd)
End Sub

Private Sub LoadMenuLandscape()
    Dim objNode As Node
    Dim LngMenuID As Long                           '菜单ID号
    Dim LngLoop As Long                             '循环变量
    Dim LngInsertMenu As Long                       '弹出菜单句柄
    Dim LngUpperMenu As Long                        '上级菜单句柄
    Dim StrHotKey As String                         '快捷键
    On Error Resume Next
    '横向排列功能菜单
    '--菜单ID由90000001开始--
    
    LngMenuID = 菜单基准.功能菜单
    Set CollMenu = New Collection                   '保存添加菜单的相关信息
    TvwMenu.Nodes.Clear
    mlngMainMenu = GetMenu(Me.hwnd)
    '删除"操作"菜单
    Call DeleteMenu(mlngMainMenu, 1, MF_BYPOSITION)
    
    '直接对主菜单进行增加操作,以实现横向排列功能菜单的方式
    With grsMenus
        Do While Not .EOF
            Err = 0
            If .Fields("上级") = 0 Then
                Set objNode = Me.TvwMenu.Nodes.Add(, , "_" & !编号, !短标题)
            Else
                Set objNode = Me.TvwMenu.Nodes.Add("_" & !上级, 4, "_" & !编号, !短标题)
            End If
            
            If Err = 0 Then
                    
                '找其上级菜单句柄
                LngUpperMenu = mlngMainMenu
                If Val(!上级) <> 0 Then
                    For LngLoop = 0 To CollMenu.Count - 1
                        If CollMenu("K_" & LngLoop)(Menu_Code) = !上级 Then
                            LngUpperMenu = CollMenu("K_" & LngLoop)(Menu_Hdl)
                            Exit For
                        End If
                    Next
                End If
                
                StrHotKey = UCase(IIf(IsNull(!快键), "", !快键))
                StrHotKey = !短标题 & IIf(StrHotKey = "", "", "(&" & StrHotKey & ")")
                '添加菜单项(如果模块值为零,则为菜单项;否则添加弹出菜单)
                If !模块 = 0 Then
                    LngInsertMenu = CreatePopupMenu()
                    CollMenu.Add Array(LngInsertMenu, .Fields("编号").Value, .Fields("模块").Value, IIf(IsNull(!部件), "", .Fields("部件").Value), LngUpperMenu, StrHotKey, IIf(!模块 = 0, 0, LngMenuID), .Fields("系统").Value), "K_" & CollMenu.Count
                Else
                    If !报表 = 1 And Val(!是否停用) = 1 Then
                        If mnuToolShowDisReport.Checked Then
                            StrHotKey = StrHotKey & "(停用)"
                            LngInsertMenu = InsertMenu(LngUpperMenu, MF_BYPOSITION, MF_STRING, LngMenuID, StrHotKey)
                            CollMenu.Add Array(LngInsertMenu, .Fields("编号").Value, .Fields("模块").Value, IIf(IsNull(!部件), "", .Fields("部件").Value), LngUpperMenu, StrHotKey, IIf(!模块 = 0, 0, LngMenuID), .Fields("系统").Value), "K_" & CollMenu.Count
                            LngMenuID = LngMenuID + 1
                        End If
                    Else
                        LngInsertMenu = InsertMenu(LngUpperMenu, MF_BYPOSITION, MF_STRING, LngMenuID, StrHotKey)
                        CollMenu.Add Array(LngInsertMenu, .Fields("编号").Value, .Fields("模块").Value, IIf(IsNull(!部件), "", .Fields("部件").Value), LngUpperMenu, StrHotKey, IIf(!模块 = 0, 0, LngMenuID), .Fields("系统").Value), "K_" & CollMenu.Count
                        LngMenuID = LngMenuID + 1
                    End If
                End If
            End If
            .MoveNext
        Loop
        .MoveFirst
    End With
    
    '绑定所有弹出菜单到mnuOper菜单下,做为其下级菜单
    Dim IntMenuLocate As Integer
    IntMenuLocate = 1
    For LngLoop = 0 To CollMenu.Count - 1
        If CollMenu("K_" & LngLoop)(Menu_Modul) = 0 Then
            StrHotKey = CollMenu("K_" & LngLoop)(Menu_Caption)              '短标题及快捷键
            Call InsertMenu(CollMenu("K_" & LngLoop)(Menu_UpperHdl), IntMenuLocate, MF_BYPOSITION + MF_POPUP, CollMenu("K_" & LngLoop)(Menu_Hdl), StrHotKey)
            IntMenuLocate = IntMenuLocate + 1
        End If
    Next
    
    '刷新菜单
    Call SetMenu(Me.hwnd, mlngMainMenu)
    Call DrawMenuBar(Me.hwnd)
    
    '设置窗体函数的地址
    LngAddFunc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf MenuProc)
    '恢复主菜单句柄
    mlngMainMenu = GetMenu(Me.hwnd)
End Sub

Public Function Show任务(ByVal ChildObj As Object)
    Dim LngWin As Long, BlnIn As Boolean, LngCount As Long
    Dim LngInsertMenu As Long, StrCaption As String
    Dim LngMenuCount As Integer, ClientRect As RECT, ClientPT As POINTAPI
    If grsMenus.State = 0 Then Exit Function
    If grsMenus.EOF Then Exit Function
    
    With grsMenus
        .MoveFirst
        .Find "标题='" & ChildObj.Caption & "'"
        If .EOF Then
            .MoveFirst

            '如果属于工具
            If Trim(ChildObj.Caption) = "" Then Exit Function
            If InStr(1, "自定义报表管理,字典管理工具,消息收发管理", ChildObj.Caption) <> 0 Then GoTo Normal
            Exit Function
        End If
    End With
    
Normal:                                                     '正常进入
    SetParent ChildObj.hwnd, Me.hwnd
    StrCaption = ChildObj.Caption
    '恢复子窗体的高度(减去主窗体的标题栏高度及菜单高度)
    ClientPT.x = 0
    ClientPT.y = 0
    Call ClientToScreen(Me.hwnd, ClientPT)
    ChildObj.Top = ChildObj.Top - (ClientPT.y * 30)
    
    '把窗体句柄加入集合
    BlnIn = False
    For LngWin = 0 To CollOpenWindowHdl.Count - 1
        If ChildObj.hwnd = CollOpenWindowHdl("K_" & LngWin)(0) Then
            BlnIn = True
            Exit For
        End If
    Next
    LngCount = CollOpenWindowHdl.Count
    
    If BlnIn = False Then
        CollOpenWindowHdl.Add Array(ChildObj.hwnd, ChildObj.Caption, 菜单基准.窗口菜单 + CollOpenWindowHdl.Count), "K_" & LngCount
        
        grsMenus.Filter = "上级 =0"
        LngMenuCount = IIf(mnuOrderMenu.Checked, grsMenus.RecordCount, 1) + IIf(History.Visible, 3, 2)
        grsMenus.Filter = 0
        
        LngInsertMenu = GetSubMenu(mlngMainMenu, LngMenuCount)           '获取窗口子菜单
        
        '加入菜单
        If CollOpenWindowHdl.Count = 1 Then
            '加入分隔菜单项
            Call InsertMenu(LngInsertMenu, MF_BYPOSITION, MF_SEPARATOR, 菜单基准.分隔菜单, "")
        End If
        '加入窗口菜单项
        Call InsertMenu(LngInsertMenu, MF_BYPOSITION, MF_STRING, 菜单基准.窗口菜单 + CollOpenWindowHdl.Count - 1, StrCaption)
    End If
    
End Function

Public Sub Shut任务(ByVal ObjFrm As Object)
    Dim LngDeleteMenu As Long, LngMenuCount As Integer
    Dim IntChange As Integer, IntDelete As Integer
    On Error Resume Next
        
    With grsMenus
        .Filter = "上级 =0"
        LngMenuCount = IIf(mnuOrderMenu.Checked, .RecordCount, 1) + IIf(History.Visible, 3, 2)
        .Filter = 0
        
        .MoveFirst
        .Find "标题='" & ObjFrm.Caption & "'"
        If .EOF Then
            .MoveFirst

            '如果属于工具
            If Trim(ObjFrm.Caption) = "" Then Exit Sub
            If InStr(1, "自定义报表管理,字典管理工具,消息收发管理", ObjFrm.Caption) = 0 Then Exit Sub
        End If
    End With
    
    LngDeleteMenu = GetSubMenu(mlngMainMenu, LngMenuCount)           '获取窗口子菜单
    
    '--清除集合--
    For IntChange = 0 To CollOpenWindowHdl.Count - 1
        If CollOpenWindowHdl("K_" & IntChange)(1) = ObjFrm.Caption Then IntDelete = IntChange: Exit For
    Next
    
    If IntChange > CollOpenWindowHdl.Count Then Exit Sub
    '依次修改后继
    For IntChange = IntChange To CollOpenWindowHdl.Count - 1
        CollOpenWindowHdl.Remove "K_" & IntChange
        CollOpenWindowHdl.Add CollOpenWindowHdl("K_" & IntChange + 1), "K_" & IntChange
    Next
    CollOpenWindowHdl.Remove "K_" & CollOpenWindowHdl.Count
    
    '删除对应菜单
    Call DeleteMenu(LngDeleteMenu, IntDelete + 2, MF_BYPOSITION)
    
    '--清除对应菜单--
    If CollOpenWindowHdl.Count = 0 Then
        '清除分隔菜单项
        Call DeleteMenu(LngDeleteMenu, 1, MF_BYPOSITION)
    End If
End Sub

Private Sub InitEvn()
    Dim StrPicPath As String, BlnShow As Boolean
    Dim LngColor As Long
    
    StrPicPath = zlDatabase.GetPara("zlMdiBackPic")
    
    If Trim(StrPicPath) <> "" Then
        '用户选择图片,测试是否正常
        On Error Resume Next
        Err = 0
        BlnShow = False
        
        ImgTry.Picture = LoadPicture(StrPicPath)
        If Err <> 0 Then
            MsgBox "显示背景图片时，发生错误！（恢复为缺省图片）", vbInformation, gstrSysName
        Else
            BlnShow = True
        End If
        If BlnShow Then PicBackBitmap.Picture = LoadPicture(StrPicPath)
    Else
        BlnShow = True
    End If
    
    Call PicBackBitmap.PaintPicture(PicBackBitmap.Picture, 0, 0, PicBackBitmap.Width, PicBackBitmap.Height, _
                    0, 0, PicBackBitmap.Picture.Width * 0.57, PicBackBitmap.Picture.Height * 0.57)
    LvwList.Picture = PicBackBitmap.Image
    '恢复原来设置的图片
    'ImgTry.Picture = LoadResPicture(101, 0) '菜单标识
    '取字体色
    LngColor = Val(zlDatabase.GetPara("zlMdiFontColor"))
    If LngColor <> -1 Then
        LvwList.ForeColor = LngColor
    End If
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
    MenuPrint grsMenus, 1
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub mnuFileReLogin_Click()
    If MsgBox("你确定要注销吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call ReLogin
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


Private Sub mnuOrderMenu_Click()
    Dim IntMenus As Integer, IntLastOrder As Integer, LngAddMenu As Long
    Dim FrmThis As Form, lngErr As Long, ClsClose As Object, intMenuCount As Integer
    
    grsMenus.Filter = 0
    
    IntLastOrder = IIf(mnuOrderMenu.Checked, 1, 0)
    mnuOrderMenu.Checked = Not mnuOrderMenu.Checked
    Call zlDatabase.SetPara("zlMdiMenuArray", IIf(mnuOrderMenu.Checked, 1, 0))
    If IntLastOrder = IIf(mnuOrderMenu.Checked, 1, 0) Then Exit Sub
    
    '恢复窗体原函数的地址
    Call SetWindowLong(Me.hwnd, GWL_WNDPROC, LngAddFunc)
    
    '循环删除现有菜单,再增加
    If IntLastOrder = 1 Then
        intMenuCount = GetMenuItemCount(mlngMainMenu)
        intMenuCount = intMenuCount - 6
        For IntMenus = 1 To intMenuCount
            Call DeleteMenu(mlngMainMenu, 1, MF_BYPOSITION)
        Next
    End If
    
    '如果改为纵向排列方式,则须另增加一个操作菜单
    If mnuOrderMenu.Checked = False Then
        LngAddMenu = CreatePopupMenu()
        Call InsertMenu(mlngMainMenu, 1, MF_STRING + MF_POPUP + MF_BYPOSITION, LngAddMenu, "操作(&O)")
        mlngMainMenu = LngAddMenu
        LngAddMenu = CreateMenu()
        Call InsertMenu(mlngMainMenu, 0, MF_STRING, LngAddMenu, "Default")
    End If
    Call SetMenu(Me.hwnd, GetMenu(Me.hwnd))
    Call DrawMenuBar(Me.hwnd)
    
    '增加菜单
    If Not mnuOrderMenu.Checked Then
        Call LoadMenuPortrait
    Else
        Call LoadMenuLandscape
    End If
    Call LoadHistory
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

Private Sub MnuRightAbout_Click()
    mnuHelpAbout_Click
End Sub

Private Sub mnuRightClientUpdate_Click()
    Call mnuRepairClientUpdate_Click
End Sub

Private Sub MnuRightComponent_Click()
    mnuRepairComponent_Click
End Sub

Private Sub mnuRightDictonary_Click()
    mnuToolDictonary_Click
End Sub

Private Sub MnuRightExcel_Click()
    Call mnuToolExcel_Click
End Sub

Private Sub MnuRightExit_Click()
    mnuFileExit_Click
End Sub

Private Sub MnuRightHistory_Click()
    Call mnuToolHistory_Click
End Sub

Private Sub mnuRightIndividuation_Click()
    MnuToolIndividuation_Click
End Sub

Private Sub mnuRightIndividuationClear_Click()
    mnuRepairIndividuationClear_Click
End Sub

Private Sub mnuRightMessage_Click()
    mnuToolMessage_Click
End Sub

Private Sub mnuRightNotice_Click()
    Call mnuToolNotice_Click
End Sub

Private Sub mnuRightNotify_Click()
    MnuToolNotify_Click
End Sub

Private Sub MnuRightReLogin_Click()
    mnuFileReLogin_Click
End Sub

Private Sub mnuRightShowDisReport_Click()
    Call mnuToolShowDisReport_Click
End Sub

Private Sub MnuRightStyle_Click()
    mnuTooleSelect_Click
End Sub

Private Sub mnuRightTester_Click()
    mnuToolTester_Click
End Sub

Private Sub mnurightBackBmp_Click()
    Dim BlnShow As Boolean              '能否正常显示
    Dim StrPicPath As String            '背景图片路径
    '--供用户选择背景图片--
    On Error GoTo ErrHand
    With Dialog
        .CancelError = True
        .Filter = "背景图片 (*.bmp;*.jpg)|*.bmp;*.jpg"
        .ShowOpen
        
        '用户选择图片,测试是否正常
        On Error Resume Next
        Err = 0
        BlnShow = False
        
        StrPicPath = .FileName
        ImgTry.Picture = LoadPicture(StrPicPath)
        If Err <> 0 Then
            MsgBox "您所选择的图片文件不正常显示！", vbInformation, gstrSysName
        Else
            BlnShow = True
        End If
    End With
    
    PicBackBitmap.Picture = LoadPicture(StrPicPath)
    Call PicBackBitmap.PaintPicture(PicBackBitmap.Picture, 0, 0, PicBackBitmap.Width, PicBackBitmap.Height, _
                    0, 0, PicBackBitmap.Picture.Width * 0.57, PicBackBitmap.Picture.Height * 0.57)
    LvwList.Picture = PicBackBitmap.Image
    '保存图片位置供下次提取
    Call zlDatabase.SetPara("zlMdiBackPic", StrPicPath)
    '恢复原来设置的图片
    'ImgTry.Picture = LoadResPicture(101, 0) '菜单标识
ErrHand:
End Sub

Private Sub mnurightSetColor_Click()
    '--供用户选择字体颜色--
    On Error GoTo ErrHand
    With Dialog
        .CancelError = True
        .ShowColor
        
        LvwList.ForeColor = .Color
        
        '保存字体色供下次提取
        Call zlDatabase.SetPara("zlMdiFontColor", .Color)
    End With
ErrHand:
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
    mnuRightIndividuation.Checked = MnuToolIndividuation.Checked
    Call zlDatabase.SetPara("使用个性化风格", IIf(MnuToolIndividuation.Checked, "1", "0"))
    SaveSetting "ZLSOFT", "私有全局\" & gstrDbUser, "使用个性化风格", IIf(MnuToolIndividuation.Checked, "1", "0")
End Sub

Private Sub MnuToolIndividuationClear_Click()
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

Private Sub mnuToolMessage_Click()
    mclsAppTool.CodeMan 0, 2, gcnOracle, Me, gstrDbUser
End Sub

Private Sub mnuTooleSelect_Click()
    mclsAppTool.CodeMan 0, 3, gcnOracle, Me, gstrDbUser, gstrMenuSys
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

Private Sub MnuToolNotify_Click()
    MnuToolNotify.Checked = Not MnuToolNotify.Checked
    mnuRightNotify.Checked = MnuToolNotify.Checked
    Call zlDatabase.SetPara("接收邮件消息", IIf(MnuToolNotify.Checked, "1", "0"))
    mclsAppTool.CodeMan 0, 4, gcnOracle, Me, gstrDbUser, IIf(MnuToolNotify.Checked = True, "Open", "Close")
End Sub

Private Sub mnuToolShowDisReport_Click()
    Dim IntMenus As Integer, intMenuCount As Integer
    Dim LngAddMenu As Long
    
    mnuToolShowDisReport.Checked = Not mnuToolShowDisReport.Checked
    mnuRightShowDisReport.Checked = mnuToolShowDisReport.Checked
    Call zlDatabase.SetPara("显示停用报表", IIf(mnuToolShowDisReport.Checked, 1, 0))

    grsMenus.Filter = 0
    
    '恢复窗体原函数的地址
    Call SetWindowLong(Me.hwnd, GWL_WNDPROC, LngAddFunc)
    
    '循环删除现有菜单,再增加
    intMenuCount = GetMenuItemCount(mlngMainMenu)
    If mnuOrderMenu.Checked Then
        intMenuCount = intMenuCount - 7
    Else
        intMenuCount = intMenuCount - 6
    End If
    For IntMenus = 1 To intMenuCount
        Call DeleteMenu(mlngMainMenu, 1, MF_BYPOSITION)
    Next
    
    '如果改为纵向排列方式,则须另增加一个操作菜单
    If Not mnuOrderMenu.Checked Then
        LngAddMenu = CreatePopupMenu()
        Call InsertMenu(mlngMainMenu, 1, MF_STRING + MF_POPUP + MF_BYPOSITION, LngAddMenu, "操作(&O)")
        mlngMainMenu = LngAddMenu
        LngAddMenu = CreateMenu()
        Call InsertMenu(mlngMainMenu, 0, MF_STRING, LngAddMenu, "Default")
    End If
    Call SetMenu(Me.hwnd, GetMenu(Me.hwnd))
    Call DrawMenuBar(Me.hwnd)
    
    '增加菜单
    If Not mnuOrderMenu.Checked Then
        Call LoadMenuPortrait
    Else
        Call LoadMenuLandscape
    End If
End Sub
 
Private Sub mnuToolTester_Click()
    mnuToolTester.Checked = mnuToolTester.Checked Xor True
    mnuRightTester.Checked = mnuToolTester.Checked
    SaveSetting "ZLSOFT", "公共全局", "SQLTest", IIf(mnuToolTester.Checked, 1, 0)
End Sub

Private Sub mnuWindowList_Click()
    Dim RectThis As RECT
    
    Call GetClientRect(Me.hwnd, RectThis)
    Call CascadeWindows(Me.hwnd, 0, RectThis, 0, 0)
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        mnuFilePreview_Click
    Case "Print"
        mnuFilePrint_Click
    Case "Dictionary"
        mnuToolDictonary_Click
    Case "Message"
        mnuToolMessage_Click
    Case "Choose"
        mnuTooleSelect_Click
    Case "Check"
        mnuRepairComponent_Click
    Case "工具"
        '刘兴宏:2007/08/22
        '问题:加入外部工具
        Call mnuToolOutToolSet_Click
    Case "Help"
        mnuHelpTitle_Click
    Case "Exit"
        mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Key = "工具" Then
        '刘兴宏:2007/08/22
        '问题:加入外部工具
        Call ExeCuteToolFile(ButtonMenu.Tag)
        Exit Sub
    End If
End Sub

Private Sub TbrUsual_ButtonClick(ByVal Button As MSComctlLib.Button)
    With grsMenus
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

Private Sub TimePass_Timer()
    Call Form_KeyDown(vbKeyF12, 7)  '清除静态变量
End Sub

Public Sub LoadHistory()
    Dim str系统 As String, str序号 As String
    Dim arr系统 As Variant, arr序号 As Variant
    Dim int系统_Cur As Integer, int序号_Cur As Integer
    Dim int系统_Max As Integer, int序号_Max As Integer
    Dim strValue As String
    
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
                    .Enabled = True
                    .Tag = grsMenus!系统 & "," & grsMenus!模块
                End With
            End If
            .Filter = 0
        End With
    Next
    If HistoryItem.UBound > 0 Then
        HistoryItem(0).Visible = False
    End If
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
    'mnuToolUp              ,15,'本地参数上传' ,'基本'
    
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
    MnuRightExcel.Visible = False
    '消息收发工具
    mnuToolMessage.Visible = False
    MnuToolNotify.Visible = False
    mnuRightMessage.Visible = False
    mnuRightNotify.Visible = False
    tbrThis.Buttons("Message").Visible = False
    '系统选项设置
    mnuTooleSelect.Visible = False
    MnuRightStyle.Visible = False
    tbrThis.Buttons("Choose").Visible = False
    '字典管理工具
    mnuToolDictonary.Visible = False
    mnuRightDictonary.Visible = False
    tbrThis.Buttons("Dictionary").Visible = False
    '当然,分隔条一定是要禁止的,只要存在其中一个功能（字典管理、消息收发、EXCEL报表或系统选项），就需要显示分隔条
    blnSplit = False
    
    intGrant = zlRegTool '(GetUnitInfo("注册码"))
    If ((intGrant And 4) = 4) Then
        If InStr(1, GetPrivFunc(0, 工具清单.消息收发工具), "基本") <> 0 Then
            mnuToolMessage.Visible = True
            MnuToolNotify.Visible = True
            mnuRightMessage.Visible = True
            mnuRightNotify.Visible = True
            tbrThis.Buttons("Message").Visible = True
            blnSplit = True
        Else
            Call zlDatabase.SetPara("接收邮件消息", "0")
        End If
    End If
    If ((intGrant And 8) = 8) Then
        If InStr(1, GetPrivFunc(0, 工具清单.EXCEL报表工具), "基本") Then
            mnuToolExcel.Visible = True
            MnuRightExcel.Visible = True
            blnSplit = True
        End If
    End If

    If InStr(1, GetPrivFunc(0, 工具清单.导航功能清单), "基本") Then
        mnuFilePrint.Visible = True
        mnuFilePreview.Visible = True
        mnuFileExcel.Visible = True
        tbrThis.Buttons("Print").Visible = True
        tbrThis.Buttons("Preview").Visible = True
        tbrThis.Buttons("printbar").Visible = True
    End If
    If InStr(1, GetPrivFunc(0, 工具清单.系统选项设置), "基本") Then
        mnuTooleSelect.Visible = True
        MnuRightStyle.Visible = True
        tbrThis.Buttons("Choose").Visible = True
        blnSplit = True
    End If
    If InStr(1, GetPrivFunc(0, 工具清单.字典管理工具), "基本") Then
        mnuToolDictonary.Visible = True
        mnuRightDictonary.Visible = True
        tbrThis.Buttons("Dictonary").Visible = True
        blnSplit = True
    End If
    mnuToolSplit2.Visible = blnSplit
    MnuRightBar3.Visible = blnSplit
    
    '如果没有"消息收发、系统选项及字典管理"
    tbrThis.Buttons("bar").Visible = (mnuToolDictonary.Visible Or mnuTooleSelect.Visible Or mnuToolMessage.Visible)
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
                Call RunModual(grsMenus!系统, lngModual, "")
                mblnHide = True
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

Public Function CloseChildWindows(ByVal frmMain As Object) As Boolean
     '功能:关闭所有子窗口
    Dim FrmThis     As Form, ClsClose As Object, IntCount As Integer, lngErr As Long
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
    lngErr = UBound(gstrObj)
    If Err.Number = 0 Then
        For IntCount = 0 To lngErr
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


