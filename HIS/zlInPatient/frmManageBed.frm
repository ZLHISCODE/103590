VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmManageBed 
   AutoRedraw      =   -1  'True
   Caption         =   "病区床位管理"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8715
   Icon            =   "frmManageBed.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwBeds 
      Height          =   4560
      Left            =   15
      TabIndex        =   3
      Tag             =   "可变化的"
      Top             =   870
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   8043
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   8715
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   7635
      NewRow1         =   0   'False
      BandForeColor2  =   8388608
      Caption2        =   "当前病区"
      Child2          =   "cboUnit"
      MinWidth2       =   1995
      MinHeight2      =   300
      Width2          =   1215
      UseCoolbarColors2=   0   'False
      NewRow2         =   0   'False
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   6630
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1995
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   5460
         _ExtentX        =   9631
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
            NumButtons      =   14
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
               Caption         =   "增加"
               Key             =   "Add"
               Description     =   "增加"
               Object.ToolTipText     =   "增加病床"
               Object.Tag             =   "增加"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "调整"
               Key             =   "Modi"
               Description     =   "调整"
               Object.ToolTipText     =   "调整病床"
               Object.Tag             =   "调整"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "撤消"
               Key             =   "Del"
               Description     =   "撤消"
               Object.ToolTipText     =   "撤消病床"
               Object.Tag             =   "撤消"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修缮"
               Key             =   "Remedy"
               Description     =   "修缮"
               Object.ToolTipText     =   "将空床转为修缮床"
               Object.Tag             =   "修缮"
               ImageKey        =   "Remedy"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "空床"
               Key             =   "Empty"
               Description     =   "空床"
               Object.ToolTipText     =   "将修好的床转为空床"
               Object.Tag             =   "空床"
               ImageKey        =   "Empty"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "列表"
               Key             =   "View"
               Description     =   "列表"
               Object.ToolTipText     =   "床位列表显示方式"
               Object.Tag             =   "列表"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Icon"
                     Object.Tag             =   "大图标(&G)"
                     Text            =   "大图标(&G)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Object.Tag             =   "小图标(&M)"
                     Text            =   "小图标(&M)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Object.Tag             =   "列表(&L)"
                     Text            =   "列表(&L)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Object.Tag             =   "详细资料(&D)"
                     Text            =   "详细资料(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5520
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageBed.frx":030A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10292
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
   Begin MSComctlLib.ImageList imgColor 
      Left            =   60
      Top             =   450
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
            Picture         =   "frmManageBed.frx":0B9E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":0DB8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":0FD2
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":11EC
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":1406
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":1620
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":183A
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":1A54
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":1C6E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":1E88
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   645
      Top             =   450
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
            Picture         =   "frmManageBed.frx":20A2
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":22BC
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":24D6
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":26F0
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":290A
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":2B24
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":2D3E
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":2F58
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":3172
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":338C
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   2760
      Top             =   495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":35A6
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":38C0
            Key             =   "M_Empty"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":3BDA
            Key             =   "F_Empty"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":3EF4
            Key             =   "Holding"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":420E
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":4528
            Key             =   "MASK_加床"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":4842
            Key             =   "MASK_非编"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":4B5C
            Key             =   "MASK_共用"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":4E76
            Key             =   "MASK_共用_加床"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":5190
            Key             =   "MASK_共用_非编"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   3345
      Top             =   495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":54AA
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":57C4
            Key             =   "M_Empty"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":5ADE
            Key             =   "F_Empty"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":5DF8
            Key             =   "Holding"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":6112
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":642C
            Key             =   "MASK_加床"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":6586
            Key             =   "MASK_非编"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":66E0
            Key             =   "MASK_共用"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":683A
            Key             =   "MASK_共用_加床"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBed.frx":6994
            Key             =   "MASK_共用_非编"
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
      Begin VB.Menu mnuFile_quit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit_Add 
         Caption         =   "增加(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Modi 
         Caption         =   "调整(&M)"
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "撤消(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Remedy 
         Caption         =   "转修缮(&R)"
      End
      Begin VB.Menu mnuEdit_Empty 
         Caption         =   "转空床(&E)"
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
            Caption         =   "病区选择(&U)"
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
      Begin VB.Menu mnuView_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSelCol 
         Caption         =   "选择列(&C)"
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_ListView 
         Caption         =   "大图标(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuView_ListView 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuView_ListView 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuView_ListView 
         Caption         =   "详细资料(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_reFlash 
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
Attribute VB_Name = "frmManageBed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Private mblnUnload As Boolean
Private mintEmpty As Integer, intHolding, intRemedy As Integer
Private Const STR_HEAD = "床号,600,0,1;科室,1200,0,2;房间号,800,0,2;状态,600,0,2;性别分类,1000,0,2;等级,1000,0,2;床位编制,1000,0,2;姓名,1000,0,0;性别,600,0,0;年龄,600,0,0"
Private mstrPrivs As String

Private Sub cboUnit_Click()
    Call ReadBeds(cboUnit.ItemData(cboUnit.ListIndex))
    Call SetMenuState
    Me.Refresh
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    If mblnUnload Then Unload Me
End Sub

Private Sub Form_Load()
    
    Call RestoreWinState(Me, App.ProductName)
    If lvwBeds.ColumnHeaders.Count = 0 Then
        Call zlcontrol.LvwSelectColumns(lvwBeds, STR_HEAD, True)
    End If
    
    mstrPrivs = gstrPrivs
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    '根据保存列表方式设置菜单
     Call SetView(lvwBeds.View)
     
     Call MakeBedIcon
        
    '读取病区
    If Not InitUnits Then mblnUnload = True: Exit Sub
    If cboUnit.ListIndex = -1 Then
        MsgBox "你不具有所有病区的权限,并且不能确定你所属病区,不能使用床位管理！", vbExclamation, gstrSysName
        mblnUnload = True: Exit Sub
    End If
    
    If Not ReadBeds(cboUnit.ItemData(cboUnit.ListIndex)) Then
        mblnUnload = True: Exit Sub
    End If
    Call SetMenuState
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '工具条占用高度
    Dim staH As Long '状态栏占用高度
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    '靠齐控件宽度和高度
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(sta.Visible, sta.Height, 0)
    
    With lvwBeds
        .Left = Me.ScaleLeft
        .Top = Me.ScaleTop + cbrH
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - cbrH - staH
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnUnload = False
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwBeds_DblClick()
    mnuEdit_Modi_Click
End Sub

Private Sub lvwBeds_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call SetMenuState
End Sub

Private Sub lvwBeds_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Not lvwBeds.SelectedItem Is Nothing Then
        mnuEdit_Modi_Click
    End If
End Sub

Private Sub lvwBeds_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objItem As ListItem
    
    If Button = 2 Then
        PopupMenu mnuEdit, 2
    ElseIf lvwBeds.View <> lvwReport Then
        Set objItem = lvwBeds.HitTest(X, Y)
        If Not objItem Is Nothing Then
            With objItem
                sta.Panels(2) = "床号[" & Trim(.Text) & "]" & _
                    " 状态:" & .SubItems(lvwBeds.ColumnHeaders("_状态").Index - 1) & _
                    " 性别分类:" & .SubItems(lvwBeds.ColumnHeaders("_性别分类").Index - 1) & _
                    " 科室:" & .SubItems(lvwBeds.ColumnHeaders("_科室").Index - 1) & _
                    " 等级:" & .SubItems(lvwBeds.ColumnHeaders("_等级").Index - 1)
            End With
        Else
            sta.Panels(2) = "当前病区共 " & lvwBeds.ListItems.Count & " 张病床,其中病人占用 " & intHolding & " 张,空床 " & mintEmpty & " 张,正在修缮 " & intRemedy & " 张！"
        End If
    Else
        sta.Panels(2) = "当前病区共 " & lvwBeds.ListItems.Count & " 张病床,其中病人占用 " & intHolding & " 张,空床 " & mintEmpty & " 张,正在修缮 " & intRemedy & " 张！"
    End If
End Sub

Private Sub mnuEdit_Del_Click()
    Dim intIdx As Integer, strSQL As String
    
    If lvwBeds.SelectedItem Is Nothing Then
        MsgBox "请选择要撤消的病床！", vbExclamation, gstrSysName: Exit Sub
    End If
    If lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_状态").Index - 1) = "占用" Then
        MsgBox "该病床已被病人占用,现在不能撤消！", vbExclamation, gstrSysName: Exit Sub
    End If
    If MsgBox("确实要撤消病床" & Mid(lvwBeds.SelectedItem.Key, 2) & " 吗？", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    On Error GoTo errH
    intIdx = lvwBeds.SelectedItem.Index
    
    strSQL = "zl_床位状况记录_Delete('" & Mid(lvwBeds.SelectedItem.Key, 2) & "'," & cboUnit.ItemData(cboUnit.ListIndex) & ")"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    On Error GoTo 0
    
    lvwBeds.ListItems.Remove intIdx
    If lvwBeds.ListItems.Count <> 0 Then
        If intIdx <= lvwBeds.ListItems.Count Then
            lvwBeds.ListItems(intIdx).Selected = True
        Else
            lvwBeds.ListItems(lvwBeds.ListItems.Count).Selected = True
        End If
        lvwBeds.SelectedItem.EnsureVisible
    End If
    Call SetBedNOLen
    Call SetMenuState
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Empty_Click()
    Dim strSQL As String
    
    If lvwBeds.SelectedItem Is Nothing Then
        MsgBox "请选择已经修缮好的病床！", vbExclamation, gstrSysName: Exit Sub
    End If
    If lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_状态").Index - 1) <> "修缮" Then
        MsgBox "该病床没有进行修缮,不能执行该操作！", vbExclamation, gstrSysName: Exit Sub
    End If
    
    strSQL = "zl_床位状况记录_REUSE('" & Mid(lvwBeds.SelectedItem.Key, 2) & "'," & cboUnit.ItemData(cboUnit.ListIndex) & ")"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    On Error GoTo 0
    
    lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_状态").Index - 1) = "空床"
    If lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_性别分类").Index - 1) = "男床" Then
        lvwBeds.SelectedItem.Icon = "M_Empty"
        lvwBeds.SelectedItem.SmallIcon = "M_Empty"
    ElseIf lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_性别分类").Index - 1) = "女床" Then
        lvwBeds.SelectedItem.Icon = "F_Empty"
        lvwBeds.SelectedItem.SmallIcon = "F_Empty"
    Else
        lvwBeds.SelectedItem.Icon = "Empty"
        lvwBeds.SelectedItem.SmallIcon = "Empty"
    End If
    
    Call SetBedIcon(lvwBeds, lvwBeds.SelectedItem)
    
    Call SetMenuState
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Modi_Click()
    If lvwBeds.SelectedItem Is Nothing Then
        MsgBox "请选择要调整的病床！", vbExclamation, gstrSysName: Exit Sub
    End If
    If lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_状态").Index - 1) = "占用" Then
        MsgBox "该病床已被病人占用,现在不能进行调整！", vbExclamation, gstrSysName: Exit Sub
    End If
    If lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_状态").Index - 1) = "修缮" Then
        MsgBox "该病床正在修缮,现在不能进行调整！", vbExclamation, gstrSysName: Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    frmEditBed.mblnModi = True
    Set frmEditBed.mlvwBeds = lvwBeds
    Set frmEditBed.mobjSta = sta
    frmEditBed.mlngUnit = cboUnit.ItemData(cboUnit.ListIndex)
    frmEditBed.Show 1, Me
    
    If gblnOK Then Call SetMenuState
End Sub

Private Sub mnuEdit_Add_Click()
    On Error Resume Next
    Err.Clear
    
    frmEditBed.mblnModi = False
    frmEditBed.mlngUnit = cboUnit.ItemData(cboUnit.ListIndex)
    Set frmEditBed.mlvwBeds = lvwBeds
    Set frmEditBed.mobjSta = sta
    frmEditBed.Show 1, Me
End Sub

Private Sub mnuEdit_Remedy_Click()
    Dim strSQL As String
    
    If lvwBeds.SelectedItem Is Nothing Then
        MsgBox "请选择要修缮的病床！", vbExclamation, gstrSysName: Exit Sub
    End If
    If lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_状态").Index - 1) <> "空床" Then
        MsgBox "该病床不是空床,不能执行该操作！", vbExclamation, gstrSysName: Exit Sub
    End If
    
    strSQL = "zl_床位状况记录_STOP('" & Mid(lvwBeds.SelectedItem.Key, 2) & "'," & cboUnit.ItemData(cboUnit.ListIndex) & ")"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    lvwBeds.SelectedItem.Icon = "Remedy"
    lvwBeds.SelectedItem.SmallIcon = "Remedy"
    lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_状态").Index - 1) = "修缮"
    
    Call SetBedIcon(lvwBeds, lvwBeds.SelectedItem)
    
    Call SetMenuState
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim lngUnitID As Long, str床号 As Long
        
    If cboUnit.ListIndex <> -1 Then lngUnitID = cboUnit.ItemData(cboUnit.ListIndex)
    If Not lvwBeds.SelectedItem Is Nothing Then str床号 = Trim(lvwBeds.SelectedItem.Text)
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "病区=" & lngUnitID, _
        "床号=" & str床号)
End Sub

Private Sub mnuView_ListView_Click(Index As Integer)
    Call SetView(CByte(Index))
End Sub

Private Sub mnuView_reFlash_Click()
    Call ReadBeds(cboUnit.ItemData(cboUnit.ListIndex))
    Call SetMenuState
    Me.Refresh
End Sub

Private Sub mnuViewSelCol_Click()
    If zlcontrol.LvwSelectColumns(lvwBeds, STR_HEAD) Then
        mnuView_reFlash_Click
        Call SetView(3)
    End If
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    sta.Visible = Not sta.Visible
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
        Case "View"
            Call SetView((lvwBeds.View + 1) Mod 4)
        Case "Add"
            mnuEdit_Add_Click
        Case "Modi"
            mnuEdit_Modi_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "Empty"
            mnuEdit_Empty_Click
        Case "Remedy"
            mnuEdit_Remedy_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Function InitUnits() As Boolean
'功能：初始化住院科室
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, lngUnitID As Long, blnLimitUnit As Boolean
    Dim strUnitIDs As String
    
    On Error GoTo errH
        
    '包含门诊观察室
    blnLimitUnit = InStr(mstrPrivs, "所有病区") = 0
    '问题30922 by lesfeng 2010-06-18 b
    If blnLimitUnit Then strUnitIDs = UserInfo.ID
    'by lesfeng 2010-1-8 性能优化
    gstrSQL = _
        " Select A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B" & IIf(blnLimitUnit, ",部门人员 C ", "") & _
        " Where B.部门ID = A.ID" & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.服务对象 IN(1,2,3) And B.工作性质='护理'" & _
        IIf(blnLimitUnit, " And A.ID = C.部门ID And C.人员ID In ([1])", "") & _
        " And (A.站点=[2] Or A.站点 is Null)" & _
        " Order by A.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strUnitIDs), gstrNodeNo)
    
'    If blnLimitUnit Then strUnitIDs = GetUserUnits
'    'by lesfeng 2010-1-8 性能优化
'    gstrSQL = _
'        " Select A.ID,A.编码,A.名称" & _
'        " From 部门表 A,部门性质说明 B" & _
'        " Where B.部门ID = A.ID" & _
'        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
'        " And B.服务对象 IN(1,2,3) And B.工作性质='护理'" & _
'        IIf(blnLimitUnit, " And A.ID In ([1])", "") & _
'        " And (A.站点=[2] Or A.站点 is Null)" & _
'        " Order by A.编码"
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strUnitIDs), gstrNodeNo)
    '问题30922 by lesfeng 2010-06-18 e
    If Not rsTmp.EOF Then
        lngUnitID = UserInfo.部门ID
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = lngUnitID And cboUnit.ListIndex = -1 Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
        If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    ElseIf InStr(";" & mstrPrivs, "所有病区") > 0 Then
        MsgBox "没有设置病区,请你先到部门管理中设置工作性质为护理的部门！", vbExclamation, gstrSysName
        Exit Function
    Else
        MsgBox "你没有 [所有病区] 的权限,并且你所在部门不是病区！", vbExclamation, gstrSysName
        Exit Function
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ReadBeds(lngUnitID As Long) As Boolean
'功能：读取指定病区的床位列表
    Dim i As Integer, j As Integer
    Dim objItem As ListItem
    Dim intBedLen As Integer
    Dim mrsBeds As ADODB.Recordset
    
    On Error GoTo errH
    intBedLen = GetMaxBedLen(lngUnitID)
    gstrSQL = _
        " Select LPAD(A.床号,[1],' ') 床号,A.病区ID," & _
        " A.房间号,A.性别分类,A.床位编制,A.等级ID,A.状态,A.病人ID,A.共用," & _
        " Nvl(B.名称,Decode(A.共用,1,'<共用病床>',NULL)) as 科室," & _
        " A.科室ID,C.名称 as 等级,D.姓名,D.性别,D.年龄" & _
        " From 床位状况记录 A,部门表 B,收费项目目录 C,病人信息 D" & _
        " Where A.科室ID=B.ID(+) And A.等级ID=C.ID(+)" & _
        " And A.病人ID=D.病人ID(+) And A.病区ID=[2] " & _
        " Order by LPAD(A.床号,[1],' ')"
    Set mrsBeds = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, intBedLen, lngUnitID)
    
    lvwBeds.ListItems.Clear
    intHolding = 0: mintEmpty = 0: intRemedy = 0
    
    If Not mrsBeds.EOF Then
        For i = 1 To mrsBeds.RecordCount
            Select Case mrsBeds!状态
                Case "空床"
                    If mrsBeds!性别分类 = "男床" Then
                        Set objItem = lvwBeds.ListItems.Add(, "_" & Trim(mrsBeds!床号), mrsBeds!床号, "M_Empty", "M_Empty")
                    ElseIf mrsBeds!性别分类 = "女床" Then
                        Set objItem = lvwBeds.ListItems.Add(, "_" & Trim(mrsBeds!床号), mrsBeds!床号, "F_Empty", "F_Empty")
                    Else
                        Set objItem = lvwBeds.ListItems.Add(, "_" & Trim(mrsBeds!床号), mrsBeds!床号, "Empty", "Empty")
                    End If
                    mintEmpty = mintEmpty + 1
                Case "占用"
                    Set objItem = lvwBeds.ListItems.Add(, "_" & Trim(mrsBeds!床号), mrsBeds!床号, "Holding", "Holding")
                    intHolding = intHolding + 1
                Case "修缮"
                    Set objItem = lvwBeds.ListItems.Add(, "_" & Trim(mrsBeds!床号), mrsBeds!床号, "Remedy", "Remedy")
                    intRemedy = intRemedy + 1
                Case Else '当作修缮
                    Set objItem = lvwBeds.ListItems.Add(, "_" & Trim(mrsBeds!床号), mrsBeds!床号, "Remedy", "Remedy")
                    mintEmpty = mintEmpty + 1
            End Select
            For j = 2 To lvwBeds.ColumnHeaders.Count
                objItem.SubItems(j - 1) = IIf(IsNull(mrsBeds.Fields(lvwBeds.ColumnHeaders(j).Text).Value), "", mrsBeds.Fields(lvwBeds.ColumnHeaders(j).Text).Value)
            Next
            objItem.Tag = IIf(IsNull(mrsBeds!科室ID), 0, mrsBeds!科室ID)
            objItem.ListSubItems(1).Tag = IIf(IsNull(mrsBeds!共用), 0, mrsBeds!共用) '记录是否共用病床
            
            Call SetBedIcon(lvwBeds, objItem)
            
            mrsBeds.MoveNext
        Next
    End If
    Call SetBedNOLen
    ReadBeds = True
    sta.Panels(2) = "当前病区共 " & lvwBeds.ListItems.Count & " 张病床,其中病人占用 " & intHolding & " 张,空床 " & mintEmpty & " 张,正在修缮 " & intRemedy & " 张！"
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetView(bytStyle As Byte)
'功能：调整床位列表显示方式
'参数：bytstyle=0-大图标,1-小图标,2-列表,3-详细资料
    mnuView_ListView(0).Checked = False
    mnuView_ListView(1).Checked = False
    mnuView_ListView(2).Checked = False
    mnuView_ListView(3).Checked = False
    mnuView_ListView(bytStyle).Checked = True
    lvwBeds.View = bytStyle
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Icon"
            Call SetView(0)
        Case "Small"
            Call SetView(1)
        Case "List"
            Call SetView(2)
        Case "Detail"
            Call SetView(3)
    End Select
End Sub

Private Sub lvwBeds_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwBeds.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwBeds.SortOrder = lvwDescending
    Else
        lvwBeds.SortOrder = lvwAscending
    End If
    lvwBeds.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwBeds.SelectedItem Is Nothing Then lvwBeds.SelectedItem.EnsureVisible
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub mnuFile_Excel_Click()
    If lvwBeds.ListItems.Count > 100 Then
        If MsgBox("输出到Excel的数据过多,这将耗费许多时间,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
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
    Dim objOut As New zlPrintLvw
    Dim bytR As Byte
    
    On Error GoTo errH
    
    '表头
    objOut.Title.Text = "住院病床清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    objOut.UnderAppItems.Add "病区:" & NeedName(cboUnit.Text)
    objOut.BelowAppItems.Add "打印人：" & UserInfo.姓名
    objOut.BelowAppItems.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    
    '表体
    Set objOut.Body.objData = lvwBeds
    
    '输出
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrViewLvw objOut, bytR
    Else
        zlPrintOrViewLvw objOut, bytStyle
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Public Sub SetMenuState()
'功能：根据病区床位情况确定功能的使用状态
    If lvwBeds.SelectedItem Is Nothing Then
        mnuFile_Print.Enabled = False
        mnuFile_Preview.Enabled = False
        mnuFile_Excel.Enabled = False
        
        tbr.Buttons("Print").Enabled = False
        tbr.Buttons("Preview").Enabled = False
        
        mnuEdit_Modi.Enabled = False
        mnuEdit_Del.Enabled = False
        mnuEdit_Remedy.Enabled = False
        mnuEdit_Empty.Enabled = False
        
        tbr.Buttons("Modi").Enabled = False
        tbr.Buttons("Del").Enabled = False
        tbr.Buttons("Remedy").Enabled = False
        tbr.Buttons("Empty").Enabled = False
    Else
        mnuFile_Print.Enabled = True
        mnuFile_Preview.Enabled = True
        mnuFile_Excel.Enabled = True
        
        tbr.Buttons("Print").Enabled = True
        tbr.Buttons("Preview").Enabled = True
        
        Select Case lvwBeds.SelectedItem.SubItems(lvwBeds.ColumnHeaders("_状态").Index - 1)
            Case "占用"
                mnuEdit_Modi.Enabled = False
                mnuEdit_Del.Enabled = False
                mnuEdit_Remedy.Enabled = False
                mnuEdit_Empty.Enabled = False
                tbr.Buttons("Modi").Enabled = False
                tbr.Buttons("Del").Enabled = False
                tbr.Buttons("Remedy").Enabled = False
                tbr.Buttons("Empty").Enabled = False
            Case "空床"
                mnuEdit_Modi.Enabled = True
                mnuEdit_Del.Enabled = True
                mnuEdit_Remedy.Enabled = True
                mnuEdit_Empty.Enabled = False
                tbr.Buttons("Modi").Enabled = True
                tbr.Buttons("Del").Enabled = True
                tbr.Buttons("Remedy").Enabled = True
                tbr.Buttons("Empty").Enabled = False
            Case "修缮"
                mnuEdit_Modi.Enabled = False
                mnuEdit_Del.Enabled = False
                mnuEdit_Remedy.Enabled = False
                mnuEdit_Empty.Enabled = True
                tbr.Buttons("Modi").Enabled = False
                tbr.Buttons("Del").Enabled = False
                tbr.Buttons("Remedy").Enabled = False
                tbr.Buttons("Empty").Enabled = True
        End Select
    End If
End Sub

Public Sub SetBedNOLen()
    Dim bytLen As Byte, i As Integer
    
    If lvwBeds.ListItems.Count = 0 Then Exit Sub
    
    bytLen = GetMaxBedLen(cboUnit.ItemData(cboUnit.ListIndex))
    
    For i = 1 To lvwBeds.ListItems.Count
        lvwBeds.ListItems(i).Text = Space(bytLen - Len(CStr(Trim(lvwBeds.ListItems(i).Text)))) & Trim(lvwBeds.ListItems(i).Text)
    Next
End Sub

Private Sub MakeBedIcon()
    Dim i As Integer, k As Integer
    
    k = img32.ListImages.Count
    For i = 1 To img32.ListImages.Count
        If Not img32.ListImages(i).Key Like "MASK_*" Then
            img32.ListImages.Add , "加床_" & img32.ListImages(i).Key, img32.Overlay("MASK_加床", i)
            img32.ListImages.Add , "非编_" & img32.ListImages(i).Key, img32.Overlay("MASK_非编", i)
            img32.ListImages.Add , "共用_" & img32.ListImages(i).Key, img32.Overlay("MASK_共用", i)
            img32.ListImages.Add , "共用_加床_" & img32.ListImages(i).Key, img32.Overlay("MASK_共用_加床", i)
            img32.ListImages.Add , "共用_非编_" & img32.ListImages(i).Key, img32.Overlay("MASK_共用_非编", i)
        End If
    Next
    
    k = img16.ListImages.Count
    For i = 1 To img16.ListImages.Count
        If Not img16.ListImages(i).Key Like "MASK_*" Then
            img16.ListImages.Add , "加床_" & img16.ListImages(i).Key, img16.Overlay("MASK_加床", i)
            img16.ListImages.Add , "非编_" & img16.ListImages(i).Key, img16.Overlay("MASK_非编", i)
            img16.ListImages.Add , "共用_" & img16.ListImages(i).Key, img16.Overlay("MASK_共用", i)
            img16.ListImages.Add , "共用_加床_" & img16.ListImages(i).Key, img16.Overlay("MASK_共用_加床", i)
            img16.ListImages.Add , "共用_非编_" & img16.ListImages(i).Key, img16.Overlay("MASK_共用_非编", i)
        End If
    Next
End Sub

Private Sub SetBedIcon(objLvw As Object, objItem As ListItem)
    If objItem.SubItems(objLvw.ColumnHeaders("_床位编制").Index - 1) = "加床" Then
        objItem.Icon = "加床_" & objItem.Icon
        objItem.SmallIcon = "加床_" & objItem.SmallIcon
    ElseIf objItem.SubItems(objLvw.ColumnHeaders("_床位编制").Index - 1) = "非编" Then
        objItem.Icon = "非编_" & objItem.Icon
        objItem.SmallIcon = "非编_" & objItem.SmallIcon
    End If
    
    If Val(objItem.ListSubItems(1).Tag) <> 0 Then
        objItem.Icon = "共用_" & objItem.Icon
        objItem.SmallIcon = "共用_" & objItem.SmallIcon
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

