VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMedicalStationSendMail 
   Caption         =   "发送报告邮件"
   ClientHeight    =   5760
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   11220
   Icon            =   "frmMedicalStationSendMail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   11220
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   11220
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   25
         Top             =   30
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   1270
         ButtonWidth     =   1429
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&S.发送"
               Key             =   "发送"
               Object.ToolTipText     =   "发送(Alt+S)"
               Object.Tag             =   "&S.发送"
               ImageKey        =   "SendMail"
               Style           =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&O.输出"
               Key             =   "输出"
               Object.ToolTipText     =   "输出为Html格式文件"
               Object.Tag             =   "&O.输出"
               ImageKey        =   "Html"
               Style           =   5
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&A.全选"
               Key             =   "全选"
               Object.ToolTipText     =   "全选(Alt+A)"
               Object.Tag             =   "&A.全选"
               ImageKey        =   "SelectAll"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&C.全清"
               Key             =   "全清"
               Object.ToolTipText     =   "全清(Alt+C)"
               Object.Tag             =   "&C.全清"
               ImageKey        =   "ClearAll"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助(Alt+H)"
               Object.Tag             =   "&H.帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出(Alt+X)"
               Object.Tag             =   "&X.退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fra2 
      Height          =   3825
      Left            =   3315
      TabIndex        =   14
      Top             =   810
      Width           =   7875
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1260
         Left            =   1065
         TabIndex        =   20
         Top             =   2475
         Width           =   6735
         _cx             =   11880
         _cy             =   2222
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
         Begin VB.Line lnX 
            Index           =   0
            Visible         =   0   'False
            X1              =   -4635
            X2              =   -2850
            Y1              =   -1695
            Y2              =   -1695
         End
         Begin VB.Line lnY 
            Index           =   0
            Visible         =   0   'False
            X1              =   270
            X2              =   270
            Y1              =   420
            Y2              =   1635
         End
      End
      Begin VB.TextBox txt 
         Height          =   1890
         Index           =   7
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   525
         Width           =   6705
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   8
         Left            =   1065
         TabIndex        =   16
         Top             =   165
         Width           =   3750
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&M.邮件内容"
         Height          =   180
         Index           =   8
         Left            =   60
         TabIndex        =   17
         Top             =   555
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&P.体检人员"
         Height          =   180
         Index           =   9
         Left            =   105
         TabIndex        =   19
         Top             =   2505
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&E.团体邮件"
         Height          =   180
         Index           =   10
         Left            =   60
         TabIndex        =   15
         Top             =   225
         Width           =   900
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   630
      Left            =   3105
      TabIndex        =   30
      Top             =   4620
      Width           =   3165
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   9
         Left            =   1980
         TabIndex        =   22
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton cmdMenu 
         Height          =   270
         Left            =   675
         Picture         =   "frmMedicalStationSendMail.frx":076A
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   285
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "查找"
         Height          =   180
         Index           =   7
         Left            =   180
         TabIndex        =   32
         Tag             =   "姓名"
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&6.姓名"
         Height          =   180
         Index           =   11
         Left            =   1020
         TabIndex        =   21
         Tag             =   "姓名"
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.TextBox txtInfo 
      Height          =   2295
      Left            =   11520
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "frmMedicalStationSendMail.frx":09F0
      Top             =   1605
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSWinsockLib.Winsock sckMail 
      Left            =   3960
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtHead 
      Height          =   2295
      Left            =   11790
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "frmMedicalStationSendMail.frx":11E8
      Top             =   2895
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   23
      Top             =   5400
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalStationSendMail.frx":1964
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14711
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
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   7950
      Top             =   1050
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
            Picture         =   "frmMedicalStationSendMail.frx":21F8
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":2972
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":30EC
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":3306
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":3526
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":3746
            Key             =   "PrintSet"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":3960
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":3B7A
            Key             =   "SendMail"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":42F4
            Key             =   "Html"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   8625
      Top             =   1050
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
            Picture         =   "frmMedicalStationSendMail.frx":450E
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":4C88
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":5402
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":561C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":583C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":5A5C
            Key             =   "PrintSet"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":5C76
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":5E90
            Key             =   "SendMail"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationSendMail.frx":660A
            Key             =   "Html"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Height          =   4425
      Left            =   45
      TabIndex        =   0
      Top             =   870
      Width           =   2820
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   2145
         TabIndex        =   13
         Text            =   "30"
         Top             =   3915
         Width           =   525
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   2145
         TabIndex        =   11
         Text            =   "5"
         Top             =   3540
         Width           =   525
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Text            =   "25"
         Top             =   1050
         Width           =   2580
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   2
         Top             =   435
         Width           =   2580
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "&6.保存发送者密码"
         Height          =   255
         Left            =   90
         TabIndex        =   9
         Top             =   3240
         Width           =   1845
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2865
         Width           =   2580
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   2235
         Width           =   2580
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1635
         Width           =   2580
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&8.等待服务应答间隔(秒)"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   3960
         Width           =   1980
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&7.连续发送邮件间隔(秒)"
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   1980
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&2.端口号"
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   28
         Top             =   795
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.邮件服务器"
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   1
         Top             =   195
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&5.密  码"
         Height          =   180
         Index           =   3
         Left            =   105
         TabIndex        =   7
         Top             =   2625
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&4.用户名"
         Height          =   180
         Index           =   2
         Left            =   105
         TabIndex        =   5
         Top             =   2010
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&3.发送人地址"
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   3
         Top             =   1425
         Width           =   1080
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileMail 
         Caption         =   "发送个人报告(&M)"
      End
      Begin VB.Menu mnuFileMailGroup 
         Caption         =   "发送团体报告(&E)"
      End
      Begin VB.Menu mnuFileOut 
         Caption         =   "输出个人报告(&O)"
      End
      Begin VB.Menu mnuFileOutGroup 
         Caption         =   "输出团体报告(&U)"
      End
      Begin VB.Menu mnuFile_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSelectAll 
         Caption         =   "全选(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileClearAll 
         Caption         =   "全清(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
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
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&T)"
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
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmMedicalStationSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean
Private mblnMaining As Boolean
Private mlng病人id As Long
Private mblnDataMoved As Boolean
Private Enum mCol
    选择 = 0
    姓名
    门诊号
    性别
    出生日期
    婚姻状况
    电子邮件
    状态
End Enum

Public WithEvents mobjPopMenu As clsPopMenu                '自定义弹出菜单对象
Attribute mobjPopMenu.VB_VarHelpID = -1
Public mbytPopMenu As Byte

'（２）自定义过程或函数************************************************************************************************
Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
    
    mnuFileMail.Enabled = True
        
    If vData = False Then mnuFileMail.Enabled = False
        
    tbrThis.Buttons("发送").Enabled = mnuFileMail.Enabled
    
End Property

Private Function CreateTmpFile(Optional ByVal strFile As String) As String
    '------------------------------------------------------------------------------------------------------------------
    '
    '功能:
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim strFileTemp As String
    Dim lngTemp As Long
    
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    
    strFileTemp = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
    
    strFileTemp = strFileTemp & strFile
    
    CreateTmpFile = strFileTemp
    
End Function

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    On Error Resume Next



    On Error GoTo 0

    Call InitData

    EditChanged = True


End Function

Public Function ShowEdit(ByVal frmMain As Object, ByRef lngKey As Long, Optional lng病人id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:  frmMain         调用窗体对象
    '       lngKey          预约登记id
    '返回:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False

    mlngKey = lngKey
        
    Set mfrmMain = frmMain
    mlng病人id = lng病人id
    
    If InitData = False Then Exit Function
    If ReadData(mlngKey, lng病人id) = False Then Exit Function
    
    EditChanged = (Val(vsf.RowData(1)) > 0)

    Me.Show 1, frmMain
    
    ShowEdit = mblnOK

End Function

Private Function ReadData(ByVal lngKey As Long, ByVal lng病人id As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  读取数据
    '参数:  lngKey      体检类型序号
    '返回:  True        读取成功
    '       False       读取失败
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand
    
    gstrSQL = "SELECT 1 AS 选择,A.病人id AS ID,A.姓名,B.门诊号,b.健康号,b.就诊卡号,a.体检编号,b.身份证号,B.性别,B.婚姻状况,TO_CHAR(B.出生日期,'yyyy-mm-dd') AS 出生日期,A.电子邮件,'' as 状态 " & _
                "FROM 体检人员档案 A,病人信息 B " & _
                "WHERE A.体检报到=1 AND A.体检状态 IN (4,5) AND A.病人id=B.病人id and A.登记id=[1] "
    If lng病人id > 0 Then gstrSQL = gstrSQL & " AND B.病人id=[2] "
    
    gstrSQL = gstrSQL & " Order By B.门诊号"
    
    mblnDataMoved = DataMove(lngKey)
    If mblnDataMoved Then
        gstrSQL = Replace(gstrSQL, "体检人员档案", "H体检人员档案")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, lng病人id)
    If rs.BOF = False Then
        Call FillGrid(vsf, rs)
        Call AppendRows(vsf, lnX, lnY)
    End If
    
    gstrSQL = "SELECT A.电子邮件 FROM 合约单位 A,体检登记记录 B WHERE A.ID=B.合约单位id AND B.ID=[1]"
    
    If mblnDataMoved Then
        gstrSQL = Replace(gstrSQL, "体检登记记录", "H体检登记记录")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then txt(8).Text = zlCommFun.NVL(rs("电子邮件"))
        
    ReadData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume

End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化设置
    '返回:  True        初始化成功
    '       False       初始化失败
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
    
    strVsf = "选择,450,1,1,1,;姓名,1080,1,1,1,;门诊号,810,7,1,1,;健康号,810,7,1,1,;就诊卡号,0,1,1,1,;体检编号,990,1,1,1,;身份证号,1200,1,1,0,;性别,600,1,1,1,;出生日期,990,1,1,1,;婚姻状况,900,1,1,1,;电子邮件,1800,1,1,1,;状态,750,1,1,1,"
    
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(0) = flexDTBoolean
    vsf.Editable = True
    
    Call AppendRows(vsf, lnX, lnY)
    
    If mlng病人id > 0 Then
        
        mnuFileMailGroup.Visible = False
        mnuFileOutGroup.Visible = False
        
        txt(8).Visible = False
        lbl(10).Visible = False
    End If
    
    InitData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  校验数据的有效性
    '返回:  True        数据有效
    '       False       数据无效
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long


    ValidEdit = True

End Function


Private Sub chk_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub cmdMenu_Click()
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(cmdMenu.hWnd, objPoint)
    
    mbytPopMenu = 3
    Set mobjPopMenu = New clsPopMenu
    Call mobjPopMenu.ShowPopupMenu(objPoint.X * Screen.TwipsPerPixelX, objPoint.Y * Screen.TwipsPerPixelY - 255 * 8 - 300)
    
    txt(9).Text = ""
    LocationObj txt(9)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyA
            If tbrThis.Buttons("全选").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("全选"))
        Case vbKeyC
            If tbrThis.Buttons("全清").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("全清"))
        Case vbKeyS
            If tbrThis.Buttons("发送").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("发送"))
            
        Case vbKeyO
            If tbrThis.Buttons("输出").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("输出"))
            
        Case vbKeyH
            If tbrThis.Buttons("帮助").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("帮助"))
        Case vbKeyX
            If tbrThis.Buttons("退出").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("退出"))
        End Select
    ElseIf Shift = 0 Then
        If KeyCode = vbKeyEscape Then
            If tbrThis.Buttons("退出").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("退出"))
        End If
    End If
End Sub

'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub Form_Load()
    
    txt(0).Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "发送人", txt(0).Text)
    txt(1).Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "发送人地址", txt(1).Text)
    txt(2).Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "用户名", txt(2).Text)
    txt(3).Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "密码", txt(3).Text)
    
    txt(4).Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "邮件服务器", txt(4).Text)
    
    txt(5).Text = Val(GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "续发间隔", txt(5).Text))
    txt(6).Text = Val(GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "等待间隔", txt(6).Text))
    
    chk.Value = Val(GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "是否保存密码", chk.Value))
    txt(7).Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "邮件内容", txt(7).Text)
    
    Call RestoreWinState(Me, App.ProductName)
    
    
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", "0")) = 1 Then
        '使用个性化设置
      
        lbl(11).Caption = "&6." & (GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "查找信息", "姓名"))
        lbl(11).Tag = Mid(lbl(11).Caption, 4)
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    With fra
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) - 90
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With fra2
        .Left = fra.Left + fra.Width + 15
        .Top = fra.Top
        .Width = Me.ScaleWidth - .Left
        .Height = fra.Height - fraInfo.Height + 90
    End With
    
    With fraInfo
        .Left = fra2.Left
        .Top = fra2.Top + fra2.Height - 90
        .Width = fra2.Width
    End With
    
    txt(8).Width = fra2.Width - txt(8).Left - 60
    With txt(7)
        .Width = fra2.Width - txt(7).Left - 60
    End With
    
    If mlng病人id > 0 Then
        txt(7).Top = txt(8).Top
        lbl(8).Top = lbl(10).Top
        vsf.Top = txt(7).Top + txt(7).Height + 30
    End If
    
    lbl(9).Top = vsf.Top + 60
    With vsf
        .Width = fra2.Width - .Left - 60
        .Height = fra2.Height - .Top - 60
    End With
    
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    If mblnMaining Then
        Cancel = True
        Exit Sub
    End If
    
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "发送人", txt(0).Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "发送人地址", txt(1).Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "用户名", txt(2).Text)
    
    If chk.Value = 1 Then
        Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "密码", txt(3).Text)
    Else
        Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "密码", "")
    End If
    
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "邮件服务器", txt(4).Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "续发间隔", Val(txt(5).Text))
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "等待间隔", Val(txt(6).Text))
    
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "是否保存密码", chk.Value)
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "邮件内容", txt(7).Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "查找信息", lbl(11).Tag)

    Call SaveWinState(Me, App.ProductName)
    
End Sub

Private Sub mnuFileClearAll_Click()
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            vsf.TextMatrix(lngLoop, mCol.选择) = 0
        End If
    Next
    
    EditChanged = False
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileMail_Click()
    Dim objMail As clsMail
    Dim blnSuccess As Boolean
    Dim strMessage As String
    Dim lngLoop As Long
    
    Dim strFile As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    '检查
    If ValidData = False Then Exit Sub
    
    Set objMail = New clsMail
    Set objMail.WinSockObj = sckMail
    
    mblnMaining = True
    
    tbrThis.Buttons("发送").Enabled = False
    tbrThis.Buttons("全清").Enabled = False
    tbrThis.Buttons("全选").Enabled = False
    tbrThis.Buttons("帮助").Enabled = False
    tbrThis.Buttons("退出").Enabled = False
    
    vsf.Editable = flexEDNone
    mnuFile.Enabled = False
    mnuView.Enabled = False
    mnuHelp.Enabled = False
    
    vsf.Cell(flexcpText, 1, mCol.状态, vsf.Rows - 1, mCol.状态) = ""
    vsf.Cell(flexcpForeColor, 1, mCol.状态, vsf.Rows - 1, mCol.状态) = COLOR.黑色
    
    frmWait.OpenWait Me, "发送电子邮件"
    frmWait.WaitInfo = "正在连接邮件服务器..."
    
    objMail.ResponseInternal = Val(txt(6).Text)
    
    If objMail.OpenMailServer(txt(4).Text, txt(2).Text, txt(3).Text, Val(txt(0).Text)) Then
'    If objMail.OpenOutLookExMail() Then
        
        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(lngLoop)) > 0 And Abs(Val(vsf.TextMatrix(lngLoop, mCol.选择))) = 1 And Trim(vsf.TextMatrix(lngLoop, mCol.电子邮件)) <> "" Then
                
                frmWait.WaitInfo = "正在发送“" & vsf.TextMatrix(lngLoop, mCol.姓名) & "”的体检报告邮件..."
                
                txtInfo.Text = txtHead.Text
                Call GetReportMessageHtml(mlngKey, Val(vsf.RowData(lngLoop)))
           
'                strFile = CreateTmpFile("体检报告.htm")
'                Set objText = objFile.CreateTextFile(strFile, True)
'                objText.Write txtInfo.Text
'                objText.Close

                blnSuccess = objMail.SendHead(vsf.TextMatrix(lngLoop, mCol.电子邮件), txt(2).Text, txt(1).Text, "您的体检报告", vbMultipartAlternative)
                blnSuccess = objMail.SendMessage(txtInfo.Text, vbTextHtml)
                blnSuccess = objMail.SendOver
'                blnSuccess = objMail.SendOutLookExMail(vsf.TextMatrix(lngLoop, mCol.电子邮件), "您的体检报告", txt(7).Text, strFile)
                
                If blnSuccess Then
                    vsf.TextMatrix(lngLoop, mCol.状态) = "已发送"
                Else
                    vsf.TextMatrix(lngLoop, mCol.状态) = "失  败"
                    vsf.Cell(flexcpForeColor, lngLoop, mCol.状态) = COLOR.红色
                End If
                
                Sleep Val(txt(5).Text) * 1000
                
            End If
        Next
    End If
    
    frmWait.WaitInfo = "正在关闭邮件服务器..."
    
    Call objMail.CloseMailServer
'    Call objMail.CloseOutLookExMail
    
    tbrThis.Buttons("发送").Enabled = True
    tbrThis.Buttons("全清").Enabled = True
    tbrThis.Buttons("全选").Enabled = True
    tbrThis.Buttons("帮助").Enabled = True
    tbrThis.Buttons("退出").Enabled = True
    
    vsf.Editable = flexEDKbdMouse
    mnuFile.Enabled = True
    mnuView.Enabled = True
    mnuHelp.Enabled = True
    mblnMaining = False
    
    frmWait.CloseWait
    
End Sub

Private Function GetReportMessageHtml(ByVal lngKey As Long, ByVal lng病人id As Long) As String
    '------------------------------------------------------------------------------------------------------------------
    '
    '功能:生成人员体检报告Html格式,用于邮件发送，注意此格式是固定的
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim lngLoop1 As Long
    Dim lngLoop2 As Long
    Dim lngLoop3 As Long
    Dim strTmp1 As String
    Dim strTmp2 As String
    
    Dim strSQL As String
    
    On Error GoTo errHand
    
    strTmp1 = txt(7).Text
    strTmp1 = ReplaceAll(strTmp1, vbCrLf, "<br>")
    
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<BODY BGCOLOR=#FFFFFF>" & vbCrLf & _
        "<table x:str border=0 cellpadding=5 cellspacing=0 width=728 style='border-collapse:collapse;table-layout:fixed;width:548pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:150pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:150pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:120pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:40pt'>"

    txtInfo.Text = txtInfo.Text & vbCrLf & _
            "<tr><td colspan=4 class=xl39 style='font-weight:300'>" & strTmp1 & "<br></td></tr>"
    
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<tr><td colspan=4 class=xlTitle style='width:536pt'>" & GetUnitName & "体检报告单</td></tr>"
                        
    strTmp1 = ""
    
    strSQL = "SELECT A.体检号,C.门诊号,B.体检时间,C.姓名,B.体检病历id,B.复查时间,D.书写人,C.门诊号,E.名称 " & _
                "FROM 体检登记记录 A,体检人员档案 B,病人信息 C,病人病历记录 D,合约单位 E " & _
                "WHERE A.合约单位ID=E.ID(+) AND D.ID(+)=B.体检病历id AND C.病人id=B.病人id AND A.ID=B.登记id AND A.ID=[1] AND B.病人id=[2] "
    
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "体检登记记录", "H体检登记记录")
        strSQL = Replace(strSQL, "体检人员档案", "H体检人员档案")
        strSQL = Replace(strSQL, "病人病历记录", "H病人病历记录")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey, lng病人id)
    
    If rs.BOF Then Exit Function
    
    txtInfo.Text = txtInfo.Text & _
        "<tr><td class=xl39 style='font-weight:700'>受检单位：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs("名称")) & "</td></tr>" & _
        "<tr><td class=xl39 style='font-weight:700'>受检人员：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs("姓名")) & "</td></tr>" & _
        "<tr><td class=xl39 style='font-weight:700'>体检日期：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & Format(zlCommFun.NVL(rs("体检时间")), "YYYY-MM-DD") & "</td></tr>" & _
        "<tr><td class=xl39 style='font-weight:700'>门 诊 号：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs("门诊号")) & "</td></tr>"
        
        
    '总检
    '------------------------------------------------------------------------------------------------------------------
    strTmp1 = ""
    strTmp2 = ""
    
    strSQL = "SELECT * FROM 体检人员结论 WHERE 病历id in (select id from 病人病历内容 where 病历记录id=[1]) ORDER BY 记录性质,记录序号"
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "体检人员结论", "H体检人员结论")
    End If
    Set rs3 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(zlCommFun.NVL(rs("体检病历id"))))
    If rs3.BOF = False Then
        For lngLoop3 = 1 To rs3.RecordCount
            
            If zlCommFun.NVL(rs3("记录性质"), 0) = 0 Then strTmp1 = strTmp1 & zlCommFun.NVL(rs3("结论描述")) & vbCrLf
            If zlCommFun.NVL(rs3("记录性质"), 0) = 1 Then strTmp2 = zlCommFun.NVL(rs3("参考建议"))
            
            rs3.MoveNext
        Next
    End If
            
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<tr><td colspan=2 class=xl39 style='font-weight:700'>一、总检报告</td>" & vbCrLf & _
        "<td colspan=2 class=xl39 style='text-align:right'>总检医生：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs("书写人")) & "</td></tr>"
    
    strTmp1 = ReplaceAll(strTmp1, vbCrLf, "<br>")
    strTmp2 = ReplaceAll(strTmp2, vbCrLf, "<br>")
    
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<tr><td colspan=4 class=xl25 style='text-align:left'>结论：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & strTmp1 & "</td></tr>" & vbCrLf & _
        "<tr><td colspan=4 class=xl25 style='text-align:left'>建议：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & strTmp2 & "</td></tr>" & vbCrLf & _
        "<tr><td colspan=4 class=xl25 style='text-align:left'>复查：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & Format(zlCommFun.NVL(rs("复查时间")), "yyyy-MM-dd") & "</td></tr>"
                    
    '体检项目报告
    '------------------------------------------------------------------------------------------------------------------
    txtInfo.Text = txtInfo.Text & _
        "<tr><td colspan=4 class=xl39 style='font-weight:700'>二、项目报告</td></tr>"

    '1.科室
    strSQL = _
        "Select c.名称, c.Id" & vbNewLine & _
        "From 部门表 c," & vbNewLine & _
        "        (Select b.执行科室id, Max(Nvl(s.排列顺序, 0)) As 排列顺序" & vbNewLine & _
        "            From 体检项目医嘱 a, 体检项目清单 b, 体检项目排列 s" & vbNewLine & _
        "            Where b.登记id = [1] And a.病人id = [2] And a.清单id = b.Id And s.诊疗项目id(+) = b.诊疗项目id And s.排列性质(+) = 1" & vbNewLine & _
        "            Group By b.执行科室id) b" & vbNewLine & _
        "Where c.Id = b.执行科室id" & vbNewLine & _
        "Order By Decode(b.排列顺序, 0, 9999999)"
    
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "体检项目医嘱", "H体检项目医嘱")
        strSQL = Replace(strSQL, "体检项目清单", "H体检项目清单")
    End If
    
    Set rs1 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey, lng病人id)
    If rs1.BOF Then Exit Function
    
    For lngLoop1 = 1 To rs1.RecordCount
        
        '2.体检项目(填写了报告的)
        strSQL = "select C.名称,B.报告id,D.书写人 " & _
                        "from ( " & _
                             "SELECT * FROM 病人医嘱记录 WHERE 病人id=[2] AND 挂号单=[1] AND 执行科室id=[3] AND 病人来源=4 AND 医嘱状态<>4 AND 诊疗类别='D' AND 相关id IS NULL " & _
                             "Union All " & _
                             "SELECT * FROM 病人医嘱记录 WHERE 病人id=[2] AND 挂号单=[1] AND 执行科室id=[3] AND 病人来源=4 AND 医嘱状态<>4 AND 诊疗类别='C' AND 相关id>0 " & _
                             ") A, " & _
                             "病人医嘱发送 B, " & _
                             "诊疗项目目录 C, " & _
                             "病人病历记录 D,体检项目排列 S " & _
                        "Where A.ID = B.医嘱id " & _
                              "AND B.报告id>0 " & _
                              "AND C.ID=A.诊疗项目ID " & _
                              "AND D.ID=B.报告id And s.诊疗项目id(+)=a.诊疗项目ID And s.排列性质(+)=1 " & _
                        "Order By Nvl(s.排列顺序,9999999)"
        
        If mblnDataMoved Then
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            strSQL = Replace(strSQL, "病人病历记录", "H病人病历记录")
        End If
    
        Set rs2 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(zlCommFun.NVL(rs("体检号"))), lng病人id, Val(zlCommFun.NVL(rs1("ID"))))
        If rs2.BOF = False Then
                
            txtInfo.Text = txtInfo.Text & "<tr><td colspan=4 class=xl39 style='font-weight:700'>● " & zlCommFun.NVL(rs1("名称")) & "</td></tr>"
            
            txtInfo.Text = txtInfo.Text & "<tr>"
            
            For lngLoop2 = 1 To rs2.RecordCount
                
                txtInfo.Text = txtInfo.Text & "<td colspan=2 class=xl39 style='font-weight:600'>· " & zlCommFun.NVL(rs2("名称")) & "</td>"
                txtInfo.Text = txtInfo.Text & "<td colspan=2 class=xl39 style='text-align:right'>检查医生：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs2("书写人")) & "</td>"
                txtInfo.Text = txtInfo.Text & "</tr>"
                
                txtInfo.Text = txtInfo.Text & _
                            "<tr><td class=xl25>项目名称</td>" & vbCrLf & _
                            "<td class=xl25>检查结果</td>" & vbCrLf & _
                            "<td class=xl25>参考范围</td>" & vbCrLf & _
                            "<td class=xl25>提示</td></tr>"
                
                '具体检查项目及结果
                strSQL = _
                    "SELECT * FROM ( " & _
                        "SELECT " & _
                               "项目, " & _
                               "内容, " & _
                               "参考, " & _
                               "Decode(标志,Null,'', '正常', '', '异常', '(+)', '偏低', '↓', '偏高', '↑',标志) As 提示," & _
                               "排列序号, " & _
                               "元素内序号 " & _
                        "FROM ( " & _
                        "SELECT " & _
                               "项目, " & _
                               "内容, " & _
                               "DECODE(SIGN(INSTR(参考,'''')),1,SUBSTR(参考,1,INSTR(参考,'''')-1),'') AS 标志, " & _
                               "DECODE(SIGN(INSTR(参考,'''')),1,SUBSTR(参考,INSTR(参考,'''')+1,1000),'') AS 参考, " & _
                               "排列序号, " & _
                               "元素内序号 " & _
                        "FROM ( " & _
                        "SELECT " & _
                               "项目, " & _
                               "DECODE(SIGN(INSTR(内容,'''')),1,SUBSTR(内容,1,INSTR(内容,'''')-1),内容) AS 内容, " & _
                               "DECODE(SIGN(INSTR(内容,'''')),1,SUBSTR(内容,INSTR(内容,'''')+1,1000),'') AS 参考, " & _
                               "排列序号, " & _
                               "元素内序号 "
                strSQL = strSQL & _
                        "FROM ( " & _
                        "SELECT C.中文名 AS 项目,DECODE(A.所见内容,NULL,NULL,A.所见内容||' '||DECODE(C.单位,NULL,'',C.单位)) AS 内容,B.排列序号,A.控件号 AS 元素内序号 FROM 病人病历所见单 A,病人病历内容 B,诊治所见项目 C " & _
                        "Where A.病历ID = B.ID " & _
                              "AND B.病历记录ID=[1] " & _
                              "AND C.ID=A.所见项ID " & _
                        "))) " & _
                        "Union All " & _
                        "SELECT B.标题文本 AS 项目,A.内容,'' AS 参考,'' AS 提示,B.排列序号,0 AS 元素内序号 FROM 病人病历文本段 A,病人病历内容 B " & _
                        "Where A.病历ID = B.ID " & _
                                "And B.病历记录ID =[1] " & _
                              "AND 元素类型 IN (0,-5) " & _
                        ") ORDER BY 排列序号,元素内序号"
                        
                If mblnDataMoved Then
                    strSQL = Replace(strSQL, "病人病历文本段", "H病人病历文本段")
                    strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
                    strSQL = Replace(strSQL, "病人病历所见单", "H病人病历所见单")
                End If
                Set rs3 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(zlCommFun.NVL(rs2("报告id"))))
                If rs3.BOF = False Then
                    For lngLoop3 = 1 To rs3.RecordCount
                        txtInfo.Text = txtInfo.Text & vbCrLf & _
                                "<tr><td class=xl28>" & zlCommFun.NVL(rs3("项目")) & "</td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs3("内容")) & "</td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs3("参考")) & "</td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs3("提示")) & "</td></tr>"
                        rs3.MoveNext
                    Next
                Else
                    txtInfo.Text = txtInfo.Text & vbCrLf & _
                                "<tr><td class=xl28 style='mso-height-source:userset;height:15.0pt'></td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & "></td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & "></td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & "></td></tr>"
                End If
                                        
                strTmp1 = ""
                strTmp2 = ""
                
                strSQL = "SELECT * FROM 体检人员结论 WHERE 病历id in (select id from 病人病历内容 where 病历记录id=[1]) ORDER BY 记录性质,记录序号"
                
                If mblnDataMoved Then
                    strSQL = Replace(strSQL, "体检人员结论", "H体检人员结论")
                End If
                
                Set rs3 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(zlCommFun.NVL(rs2("报告id"))))
                If rs3.BOF = False Then
                    For lngLoop3 = 1 To rs3.RecordCount
                        
                        If zlCommFun.NVL(rs3("记录性质"), 0) = 0 Then strTmp1 = strTmp1 & zlCommFun.NVL(rs3("结论描述")) & vbCrLf
                        If zlCommFun.NVL(rs3("记录性质"), 0) = 1 Then strTmp2 = zlCommFun.NVL(rs3("参考建议"))
                        
                        rs3.MoveNext
                    Next
                End If
                
                txtInfo.Text = txtInfo.Text & vbCrLf & _
                    "<tr><td colspan=4 class=xl28 style='font-weight:600'>结论：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & strTmp1 & "</td></tr>" & vbCrLf & _
                    "<tr><td colspan=4 class=xl28 style='font-weight:600'>建议：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & strTmp2 & "</td></tr>"
                    
                txtInfo.Text = txtInfo.Text & vbCrLf & "<tr><td class=xl39 style='mso-height-source:userset;height:15.0pt'></td></tr>"
                
                rs2.MoveNext
            Next
        End If
        
        rs1.MoveNext
    Next
                
    '完结
    txtInfo.Text = txtInfo.Text & vbCrLf & "</tr></table></BODY></HTML>"
    
    GetReportMessageHtml = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetGroupReportMessageHtml(ByVal lngKey As Long) As String
    '------------------------------------------------------------------------------------------------------------------
    '
    '功能:生成团体体检报告Html格式,用于邮件发送
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim intCount As Integer
    Dim strSQL As String
    Dim strTmp1 As String
    
    strTmp1 = txt(7).Text
    strTmp1 = ReplaceAll(strTmp1, vbCrLf, "<br>")
    
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<BODY BGCOLOR=#FFFFFF>" & vbCrLf & _
        "<table x:str border=0 cellpadding=5 cellspacing=0 style='border-collapse:collapse;table-layout:fixed;width:400pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>" & vbCrLf & _
        "<col style='mso-width-source:userset;mso-width-alt:512;width:25pt'>"
        
    txtInfo.Text = txtInfo.Text & vbCrLf & _
            "<tr><td colspan=10 class=xl39 style='font-weight:300'>" & strTmp1 & "<br></td></tr>"
    
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<tr><td colspan=10 class=xlTitle>团体体检报告单</td></tr>"
                        
    strTmp1 = ""
    
    strSQL = "SELECT A.体检号,A.体检时间,B.名称 FROM 体检登记记录 A,合约单位 B WHERE B.ID=A.合约单位id AND A.ID=[1]"
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "体检登记记录", "H体检登记记录")
    End If
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF Then Exit Function
    
    txtInfo.Text = txtInfo.Text & _
        "<tr><td class=xl39 colspan=10 style='font-weight:700'>体检团体：<font class=font8>" & zlCommFun.NVL(rs("名称")) & "</td></tr>" & _
        "<tr><td class=xl39 colspan=10 style='font-weight:700'>体检日期：<font class=font8>" & Format(zlCommFun.NVL(rs("体检时间")), "YYYY-MM-DD") & "</td></tr>" & _
        "<tr><td class=xl39 colspan=10 style='font-weight:700'>体检单号：<font class=font8>" & zlCommFun.NVL(rs("体检号")) & "</td></tr>"
        
    '1.人数情况
            
    strSQL = _
        "SELECT " & _
            "DECODE(男性人数,0,NULL,男性人数) AS 男性人数, " & _
            "DECODE(女性人数,0,NULL,女性人数) AS 女性人数, " & _
            "DECODE(人数,0,NULL,人数) AS 人数, " & _
            "DECODE(已检男性人数,0,NULL,已检男性人数) AS 已检男性人数, " & _
            "DECODE(已检女性人数,0,NULL,已检女性人数) AS 已检女性人数, " & _
            "DECODE(已检人数,0,NULL,已检人数) AS 已检人数, " & _
            "DECODE(未检男性人数,0,NULL,未检男性人数) AS 未检男性人数, " & _
            "DECODE(未检女性人数,0,NULL,未检女性人数) AS 未检女性人数, " & _
            "DECODE(未检人数, 0, Null, 未检人数) As 未检人数 " & _
        "From " & _
        "( " & _
        "SELECT A.男性人数, " & _
               "A.女性人数, " & _
               "nvl(A.男性人数,0)+nvl(A.女性人数,0) AS 人数, " & _
               "A.已检男性人数, " & _
               "A.已检女性人数, " & _
               "nvl(A.已检男性人数,0)+nvl(A.已检女性人数,0) AS 已检人数, " & _
               "nvl(A.男性人数,0)-nvl(A.已检男性人数,0) AS 未检男性人数, " & _
               "nvl(A.女性人数,0)-nvl(A.已检女性人数,0) AS 未检女性人数, " & _
               "(nvl(A.男性人数,0)-nvl(A.已检男性人数,0))+(nvl(A.女性人数,0)-nvl(A.已检女性人数,0)) AS 未检人数 "
               
    strSQL = strSQL & _
        "From " & _
        "( " & _
        "select SUM(DECODE(sign(instr(B.性别,'女')-0),1,0,1)) AS 男性人数, " & _
               "SUM(DECODE(sign(instr(B.性别,'女')-0),1,1,0)) AS 女性人数, " & _
               "SUM(DECODE(sign(0 - NVL(B.体检病历ID,0)),-1, DECODE(SIGN(instr(B.性别,'女')-0),1,0,1),0)) AS 已检男性人数, " & _
               "SUM(DECODE(sign(0 - NVL(B.体检病历ID,0)),-1, DECODE(SIGN(instr(B.性别,'女')-0),1,1,0),0)) AS 已检女性人数 " & _
        "from 体检登记记录 A, " & _
             "体检人员档案 B " & _
        "Where A.ID = B.登记ID " & _
              "AND A.ID=[1] " & _
        ") A " & _
        ")"
        
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "体检登记记录", "H体检登记记录")
        strSQL = Replace(strSQL, "体检人员档案", "H体检人员档案")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF Then Exit Function
    
    intCount = intCount + 1
    txtInfo.Text = txtInfo.Text & "<tr><td colspan=10 class=xl39 style='font-weight:700'>" & intCount & ".人数情况</td></tr>"
        
    txtInfo.Text = txtInfo.Text & _
        "<tr>" & _
        "<td rowspan=2 class=xl25></td>" & _
        "<td colspan=3 class=xl25>人数</td>" & _
        "<td colspan=3 class=xl25>已检人数</td>" & _
        "<td colspan=3 class=xl25>未检人数</td>" & _
        "</tr>" & _
        "<tr>" & _
        "<td class=xl25>男性</td>" & _
        "<td class=xl25>女性</td>" & _
        "<td class=xl25>合计</td>" & _
        "<td class=xl25>男性</td>" & _
        "<td class=xl25>女性</td>" & _
        "<td class=xl25>合计</td>" & _
        "<td class=xl25>男性</td>" & _
        "<td class=xl25>女性</td>" & _
        "<td class=xl25>合计</td>" & _
        "</tr>"
    
    txtInfo.Text = txtInfo.Text & _
        "<tr>" & _
        "<td class=xl25>人数</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("男性人数")) & "</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("女性人数")) & "</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("人数")) & "</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("已检男性人数")) & "</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("已检女性人数")) & "</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("已检人数")) & "</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("未检男性人数")) & "</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("未检女性人数")) & "</td>" & _
        "<td class=xl29><font class=font8>" & zlCommFun.NVL(rs("未检人数")) & "</td>" & _
        "</tr>"
    
    '2.患病情况
    strSQL = _
        "Select 结论描述,count(病人id) As 人数,100*Count(病人id)/Decode(已检总人数,Null,1,0,1,已检总人数) As 比例 From " & _
        "( " & _
        "Select Distinct 结论描述,病人id From 体检人员结论 " & _
        "Where 记录性质 = 0 " & _
              "And 病历id in " & _
                  "( " & _
                   "Select A.ID From 病人病历内容 A,病历元素目录 B " & _
                   "Where A.元素编码=B.编码 AND upper(B.部件)='ZL9CISCORE.USRMEDICALSUM' " & _
                         "And A.病历记录id In " & _
                             "( " & _
                              "Select 体检病历id From 体检人员档案 Where 登记id=[1] " & _
                             ") " & _
                  ") " & _
        ") A, " & _
        "(Select Count(1) As 已检总人数 From 体检人员档案 Where 体检病历id>0 And 登记id=[1]) B " & _
        "Group by 结论描述,已检总人数"
        
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
        strSQL = Replace(strSQL, "体检人员档案", "H体检人员档案")
        strSQL = Replace(strSQL, "体检人员结论", "H体检人员结论")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF Then Exit Function
    
    intCount = intCount + 1
    txtInfo.Text = txtInfo.Text & "<tr><td colspan=10 class=xl39 style='font-weight:700'><br>" & intCount & ".患病情况</td></tr>"
        
    txtInfo.Text = txtInfo.Text & _
        "<tr>" & _
        "<td colspan=6 class=xl25>疾病名称</td>" & _
        "<td colspan=2 class=xl25>人数</td>" & _
        "<td colspan=2 class=xl25>比例</td>" & _
        "</tr>"
        
    Do While Not rs.EOF
        
        txtInfo.Text = txtInfo.Text & _
            "<tr>" & _
            "<td colspan=6 class=xl28><font class=font8>" & zlCommFun.NVL(rs("结论描述")) & "</td>" & _
            "<td colspan=2 class=xl29><font class=font8>" & zlCommFun.NVL(rs("人数")) & "</td>" & _
            "<td colspan=2 class=xl29><font class=font8>" & Format(zlCommFun.NVL(rs("比例")), "0.00") & "%</td>" & _
            "</tr>"
            
        rs.MoveNext
    Loop
    
    strSQL = _
        "Select Count(病人id) As 人数,100*Count(病人id)/Decode(已检总人数,Null,1,0,1,已检总人数) As 比例 From " & _
        "( " & _
        "Select Distinct 结论描述,病人id From 体检人员结论 " & _
        "Where 记录性质 = 0 " & _
              "And 病历id in " & _
                  "( " & _
                   "Select A.ID From 病人病历内容 A,病历元素目录 B " & _
                   "Where A.元素编码=B.编码 AND upper(B.部件)='ZL9CISCORE.USRMEDICALSUM' " & _
                         "And A.病历记录id In " & _
                             "( " & _
                              "Select 体检病历id From 体检人员档案 Where 登记id=[1] " & _
                             ") " & _
                  ") " & _
        ") A, " & _
        "(Select Count(1) As 已检总人数 From 体检人员档案 Where 体检病历id>0 And 登记id=[1]) B " & _
        "Group by 已检总人数"
        
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
        strSQL = Replace(strSQL, "体检人员档案", "H体检人员档案")
        strSQL = Replace(strSQL, "体检人员结论", "H体检人员结论")
    End If
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF Then Exit Function
    txtInfo.Text = txtInfo.Text & _
        "<tr>" & _
        "<td colspan=6 class=xl25>合计</td>" & _
        "<td colspan=2 class=xl29><font class=font8>" & zlCommFun.NVL(rs("人数")) & "</td>" & _
        "<td colspan=2 class=xl29><font class=font8>" & Format(zlCommFun.NVL(rs("比例")), "0.00") & "%</td>" & _
        "</tr>"
    
    '3.患病名单
    strSQL = _
        "Select Distinct A.结论描述,B.姓名 " & _
        "From 体检人员结论 A, " & _
             "病人信息 B " & _
        "Where A.记录性质 = 0 " & _
              "And 病历id in " & _
                  "( " & _
                   "Select A.ID From 病人病历内容 A,病历元素目录 B " & _
                   "Where A.元素编码=B.编码 AND upper(B.部件)='ZL9CISCORE.USRMEDICALSUM' " & _
                         "And A.病历记录id In " & _
                             "( " & _
                              "Select 体检病历id From 体检人员档案 Where 登记id=[1]" & _
                             ") " & _
                  ") " & _
               "And B.病人id=A.病人id " & _
        "Group By A.结论描述,B.姓名"
        
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
        strSQL = Replace(strSQL, "体检人员档案", "H体检人员档案")
        strSQL = Replace(strSQL, "体检人员结论", "H体检人员结论")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF Then Exit Function
    
    intCount = intCount + 1
    txtInfo.Text = txtInfo.Text & "<tr><td colspan=10 class=xl39 style='font-weight:700'><br>" & intCount & ".患病名单</td></tr>"
        
    txtInfo.Text = txtInfo.Text & _
        "<tr>" & _
        "<td colspan=6 class=xl25>疾病名称</td>" & _
        "<td colspan=4 class=xl25>姓名</td>" & _
        "</tr>"
    
    Dim strSvrName As String
    Dim strList As String
    
    Do While Not rs.EOF
        
        If strSvrName <> "" Then
            If strSvrName <> zlCommFun.NVL(rs("结论描述")) Then
                
                If strList <> "" Then strList = Mid(strList, 2)
                
                txtInfo.Text = txtInfo.Text & _
                    "<tr>" & _
                    "<td colspan=6 class=xl28><font class=font8>" & strSvrName & "</td>" & _
                    "<td colspan=4 class=xl28><font class=font8>" & strList & "</td>" & _
                    "</tr>"
                    
                strList = ""
            End If
        End If
        
        strList = strList & "、" & zlCommFun.NVL(rs("姓名"))
        strSvrName = zlCommFun.NVL(rs("结论描述"))
        
        rs.MoveNext
    Loop
                   
    If strSvrName <> "" Then
            
        If strList <> "" Then strList = Mid(strList, 2)
        
        txtInfo.Text = txtInfo.Text & _
            "<tr>" & _
            "<td colspan=6 class=xl28><font class=font8>" & strSvrName & "</td>" & _
            "<td colspan=4 class=xl28><font class=font8>" & strList & "</td>" & _
            "</tr>"
            
        strList = ""
    
    End If
                
    '完结
    txtInfo.Text = txtInfo.Text & vbCrLf & "</table></BODY></HTML>"
End Function


Private Function ValidData() As Boolean
    '检查
    If Trim(txt(4).Text) = "" Then
        MsgBox "必须确定邮件服务器！"
        LocationObj txt(4)
        Exit Function
    End If
    
    If Val(txt(0).Text) = 0 Then
        MsgBox "必须邮件端口号（一般为25）！"
        LocationObj txt(0)
        Exit Function
    End If
    
    If Trim(txt(1).Text) = "" Then
        MsgBox "必须确定发送人的电子邮件地址！"
        LocationObj txt(1)
        Exit Function
    End If
    
    
    If Trim(txt(2).Text) = "" Then
        MsgBox "必须确定用户名！"
        LocationObj txt(2)
        Exit Function
    End If
    
    If Trim(txt(8).Text) = "" And mlng病人id = 0 Then
        MsgBox "必须确定团体电子邮件地址！"
        LocationObj txt(8)
        Exit Function
    End If
    
    ValidData = True
    
End Function

Private Sub mnuFileMailGroup_Click()
    Dim objMail As clsMail
    Dim blnSuccess As Boolean
    Dim strMessage As String
    Dim lngLoop As Long
    
    Dim strFile As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    '检查
    If ValidData = False Then Exit Sub
    
    Set objMail = New clsMail
    Set objMail.WinSockObj = sckMail
    
    mblnMaining = True
    
    tbrThis.Buttons("发送").Enabled = False
    tbrThis.Buttons("全清").Enabled = False
    tbrThis.Buttons("全选").Enabled = False
    tbrThis.Buttons("帮助").Enabled = False
    tbrThis.Buttons("退出").Enabled = False
    
    vsf.Editable = flexEDNone
    mnuFile.Enabled = False
    mnuView.Enabled = False
    mnuHelp.Enabled = False
    
    vsf.Cell(flexcpText, 1, mCol.状态, vsf.Rows - 1, mCol.状态) = ""
    vsf.Cell(flexcpForeColor, 1, mCol.状态, vsf.Rows - 1, mCol.状态) = COLOR.黑色
    
    frmWait.OpenWait Me, "发送电子邮件"
    frmWait.WaitInfo = "正在连接邮件服务器..."
    
    objMail.ResponseInternal = Val(txt(6).Text)
    
    If objMail.OpenMailServer(txt(4).Text, txt(2).Text, txt(3).Text, Val(txt(0).Text)) Then
'    If objMail.OpenOutLookExMail() Then
        
        frmWait.WaitInfo = "正在发送团体报告邮件..."
        
        txtInfo.Text = txtHead.Text
        Call GetGroupReportMessageHtml(mlngKey)
        
'        strFile = CreateTmpFile("团体体检报告.htm")
'        Set objText = objFile.CreateTextFile(strFile, True)
'        objText.Write txtInfo.Text
                
        blnSuccess = objMail.SendHead(Trim(txt(8).Text), txt(2).Text, txt(1).Text, "团体体检报告", vbMultipartAlternative)
        blnSuccess = objMail.SendMessage(txt(7).Text, vbTextPlain)
        blnSuccess = objMail.SendMessage(txtInfo.Text, vbTextHtml)
        blnSuccess = objMail.SendOver
                
'        objText.Close
        
'        blnSuccess = objMail.SendOutLookExMail(txt(8).Text, "团体体检报告", txt(7).Text, strFile)
                
        If blnSuccess = False Then ShowSimpleMsg "团体报告发送失败！"
                
    End If
    
    frmWait.WaitInfo = "正在关闭邮件服务器..."
    
    Call objMail.CloseMailServer
'    Call objMail.CloseOutLookExMail
    
    tbrThis.Buttons("发送").Enabled = True
    tbrThis.Buttons("全清").Enabled = True
    tbrThis.Buttons("全选").Enabled = True
    tbrThis.Buttons("帮助").Enabled = True
    tbrThis.Buttons("退出").Enabled = True
    
    vsf.Editable = flexEDKbdMouse
    mnuFile.Enabled = True
    mnuView.Enabled = True
    mnuHelp.Enabled = True
    mblnMaining = False
    
    frmWait.CloseWait
    
End Sub

Private Sub mnuFileOut_Click()
    Dim objMail As clsMail
    Dim blnSuccess As Boolean
    Dim strMessage As String
    Dim lngLoop As Long
    
    Dim strFile As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strPath
    
    On Error GoTo errHand
    
    strPath = zlCommFun.OpenDir(Me.hWnd, "指定报告文件保存的目录")
    
    If Trim(strPath) = "" Then Exit Sub
    
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
        
    mblnMaining = True
    
    tbrThis.Buttons("发送").Enabled = False
    tbrThis.Buttons("全清").Enabled = False
    tbrThis.Buttons("全选").Enabled = False
    tbrThis.Buttons("帮助").Enabled = False
    tbrThis.Buttons("退出").Enabled = False
    
    vsf.Editable = flexEDNone
    mnuFile.Enabled = False
    mnuView.Enabled = False
    mnuHelp.Enabled = False
    
    vsf.Cell(flexcpText, 1, mCol.状态, vsf.Rows - 1, mCol.状态) = ""
    vsf.Cell(flexcpForeColor, 1, mCol.状态, vsf.Rows - 1, mCol.状态) = COLOR.黑色
    
    frmWait.OpenWait Me, "生成体检报告"
    frmWait.WaitInfo = "正在生成Html格式的体检报告..."
            
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 And Abs(Val(vsf.TextMatrix(lngLoop, mCol.选择))) = 1 Then
            
            frmWait.WaitInfo = "正在生成“" & vsf.TextMatrix(lngLoop, mCol.姓名) & "”的Html格式体检报告..."
            
            txtInfo.Text = txtHead.Text
            Call GetReportMessageHtml(mlngKey, Val(vsf.RowData(lngLoop)))
            
            strFile = strPath & "体检报告(" & vsf.TextMatrix(lngLoop, mCol.门诊号) & "_" & vsf.TextMatrix(lngLoop, mCol.姓名) & ").htm"
            Set objText = objFile.CreateTextFile(strFile, True)
            objText.Write txtInfo.Text
            objText.Close
                        
        End If
    Next
    
errHand:
    
    tbrThis.Buttons("发送").Enabled = True
    tbrThis.Buttons("全清").Enabled = True
    tbrThis.Buttons("全选").Enabled = True
    tbrThis.Buttons("帮助").Enabled = True
    tbrThis.Buttons("退出").Enabled = True
    
    vsf.Editable = flexEDKbdMouse
    mnuFile.Enabled = True
    mnuView.Enabled = True
    mnuHelp.Enabled = True
    mblnMaining = False
    
    On Error Resume Next
    
    frmWait.CloseWait
End Sub

Private Sub mnuFileOutGroup_Click()
    Dim objMail As clsMail
    Dim blnSuccess As Boolean
    Dim strMessage As String
    Dim lngLoop As Long
    
    Dim strFile As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream

    Dim strPath
    
    strPath = zlCommFun.OpenDir(Me.hWnd, "指定报告文件保存的目录")
    
    If Trim(strPath) = "" Then Exit Sub
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    mblnMaining = True
    
    tbrThis.Buttons("发送").Enabled = False
    tbrThis.Buttons("全清").Enabled = False
    tbrThis.Buttons("全选").Enabled = False
    tbrThis.Buttons("帮助").Enabled = False
    tbrThis.Buttons("退出").Enabled = False
    
    vsf.Editable = flexEDNone
    mnuFile.Enabled = False
    mnuView.Enabled = False
    mnuHelp.Enabled = False
    
    vsf.Cell(flexcpText, 1, mCol.状态, vsf.Rows - 1, mCol.状态) = ""
    vsf.Cell(flexcpForeColor, 1, mCol.状态, vsf.Rows - 1, mCol.状态) = COLOR.黑色
    
    frmWait.OpenWait Me, "生成体检报告"
    frmWait.WaitInfo = "正在生成Html格式的体检报告..."
    
    txtInfo.Text = txtHead.Text
    Call GetGroupReportMessageHtml(mlngKey)
    
'        strFile = CreateTmpFile("团体体检报告.htm")
    strFile = strPath & "团体体检报告" & mlngKey & ".htm"
    
    Set objText = objFile.CreateTextFile(strFile, True)
    objText.Write txtInfo.Text
    
    tbrThis.Buttons("发送").Enabled = True
    tbrThis.Buttons("全清").Enabled = True
    tbrThis.Buttons("全选").Enabled = True
    tbrThis.Buttons("帮助").Enabled = True
    tbrThis.Buttons("退出").Enabled = True
    
    vsf.Editable = flexEDKbdMouse
    mnuFile.Enabled = True
    mnuView.Enabled = True
    mnuHelp.Enabled = True
    mblnMaining = False
    
    frmWait.CloseWait
End Sub

Private Sub mnuFileSelectAll_Click()
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            vsf.TextMatrix(lngLoop, mCol.选择) = 1
            EditChanged = True
        End If
    Next
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
        If mnuFileMail.Visible Then mobjPopMenu.Add 1, mnuFileMail.Caption, , , mnuFileMail.Enabled
        If mnuFileMailGroup.Visible Then mobjPopMenu.Add 2, mnuFileMailGroup.Caption, , , mnuFileMailGroup.Enabled
    Case 2
        
        If mnuFileOut.Visible Then mobjPopMenu.Add 1, mnuFileOut.Caption, , , mnuFileOut.Enabled
        If mnuFileOutGroup.Visible Then mobjPopMenu.Add 2, mnuFileOutGroup.Caption, , , mnuFileOutGroup.Enabled
    Case 3
        
        mobjPopMenu.Add 1, "&1.姓名", , , True, , (lbl(11).Tag = "姓名")
        mobjPopMenu.Add 2, "&2.门诊号", , , True, , (lbl(11).Tag = "门诊号")
        mobjPopMenu.Add 3, "&3.健康号", , , True, , (lbl(11).Tag = "健康号")
        mobjPopMenu.Add 4, "&4.就诊卡号", , , True, , (lbl(11).Tag = "就诊卡号")
        mobjPopMenu.Add 5, "&5.姓名拼音", , , True, , (lbl(11).Tag = "姓名拼音")
        mobjPopMenu.Add 6, "&6.姓名五笔", , , True, , (lbl(11).Tag = "姓名五笔")
        mobjPopMenu.Add 7, "&7.身份证号", , , True, , (lbl(11).Tag = "身份证号")
        mobjPopMenu.Add 8, "&8.体检编号", , , True, , (lbl(11).Tag = "体检编号")
        
    End Select
    
End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case mbytPopMenu
    Case 1
        Select Case Key
        Case 1
            Call mnuFileMail_Click
        Case 2
            Call mnuFileMailGroup_Click
        End Select
    Case 2
        Select Case Key
        Case 1
            Call mnuFileOut_Click
        Case 2
            Call mnuFileOutGroup_Click
        End Select
    Case 3
    
        Caption = Mid(Caption, 4)
        
        lbl(11).Caption = "&6." & Left(Trim(Caption), Len(Trim(Caption)) - 1)
        lbl(11).Tag = Left(Trim(Caption), Len(Trim(Caption)) - 1)
        
    End Select
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(tbrThis.hWnd, objPoint)
    
    Select Case Button.Key
    Case "全选"
        Call mnuFileSelectAll_Click
    Case "全清"
        Call mnuFileClearAll_Click
    Case "发送"
    
        mbytPopMenu = 1
        Set mobjPopMenu = New clsPopMenu
        Call mobjPopMenu.ShowPopupMenu(objPoint.X * 15 + Button.Left - 15, objPoint.Y * 15 + Button.Top + Button.Height + 15)
        
    Case "输出"
        
        mbytPopMenu = 2
        Set mobjPopMenu = New clsPopMenu
        Call mobjPopMenu.ShowPopupMenu(objPoint.X * 15 + Button.Left - 15, objPoint.Y * 15 + Button.Top + Button.Height + 15)
        
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Call tbrThis_ButtonClick(Button)
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub txt_Change(Index As Integer)
    If Index = 2 Then txt(2).Tag = "Changed"
End Sub

Private Sub txt_GotFocus(Index As Integer)
    If Index <> 7 Then zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim strCol As String
    Dim lngCol As Long
    Dim lngRow As Long
    
    Dim blnCard As Boolean

    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
    strCol = Mid(lbl(11).Caption, 4)
    lngCol = GetCol(vsf, strCol)
            
    If strCol = "就诊卡号" And KeyAscii <> vbKeyReturn And Index = 9 Then
        '就诊卡号，自动识别

        blnCard = InputIsCard(txt(Index).Text, KeyAscii)

        If blnCard And Len(txt(Index).Text) = ParamInfo.就诊卡号码长度 - 1 And KeyAscii <> 8 And txt(Index).Text <> "" Then
            If KeyAscii <> 13 Then
                txt(Index).Text = txt(Index).Text & Chr(KeyAscii)
                txt(Index).SelStart = Len(txt(Index).Text)
            End If
            KeyAscii = vbKeyReturn
        End If
    End If
        
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        If Index = 9 And Trim(txt(Index).Text) <> "" Then
            
            strCol = Mid(lbl(11).Caption, 4)
            
            Select Case strCol
            Case "姓名拼音"
                lngCol = GetCol(vsf, "姓名")
            Case "姓名五笔"
                lngCol = GetCol(vsf, "姓名")
            Case Else
                lngCol = GetCol(vsf, strCol)
            End Select
'            lngCol = GetCol(vsf, strCol)

            If lngCol < 0 Then Exit Sub
            
            lngRow = 0
            If vsf.Row + 1 <= vsf.Rows - 1 Then
                For lngLoop = vsf.Row + 1 To vsf.Rows - 1
                
                    lngRow = 0
                    Select Case strCol
                    Case "门诊号"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "健康号"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "就诊卡号"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "身份证号"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名拼音"
                        If zlGetSymbol(UCase(vsf.TextMatrix(lngLoop, lngCol))) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名五笔"
                        If zlGetSymbol(UCase(vsf.TextMatrix(lngLoop, lngCol)), 1) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case Else
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    End Select
            
                    If lngRow > 0 Then Exit For
 
                Next
            End If
            
            If lngRow = 0 Then
                For lngLoop = 1 To vsf.Row

                    lngRow = 0
                    Select Case strCol
                    Case "门诊号"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "健康号"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "就诊卡号"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "身份证号"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名拼音"
                        If zlGetSymbol(UCase(vsf.TextMatrix(lngLoop, lngCol))) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名五笔"
                        If zlGetSymbol(UCase(vsf.TextMatrix(lngLoop, lngCol)), 1) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case Else
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    End Select
            
                    If lngRow > 0 Then Exit For
                    
                Next
            End If
            
            If lngRow <= 0 Then
                ShowSimpleMsg "没有找到符合要求的信息！"
                txt(Index).Text = ""
            Else
                vsf.ShowCell lngRow, vsf.Col
                vsf.Row = lngRow
            End If
            
            txt(Index).SetFocus
            zlControl.TxtSelAll txt(Index)
    
        Else
            zlCommFun.PressKey vbKeyTab
        End If
    End If
    
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngLoop As Long
    
    If Abs(Val(vsf.TextMatrix(Row, mCol.选择))) = 1 Then
        EditChanged = True
        Exit Sub
    End If
        
    For lngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(lngLoop, mCol.选择))) = 1 Then
            EditChanged = True
            Exit Sub
        End If
    Next
    
    If lngLoop = vsf.Rows Then EditChanged = False
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mCol.选择 Or Val(vsf.RowData(Row)) <= 0 Then
        Cancel = True
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

