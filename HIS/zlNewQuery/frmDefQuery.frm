VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmDefQuery 
   Caption         =   "查询页面定义"
   ClientHeight    =   6165
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "frmDefQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   5805
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   635
      SimpleText      =   $"frmDefQuery.frx":212A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDefQuery.frx":2171
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10583
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
   Begin MSComctlLib.TreeView tvw 
      Height          =   2550
      Left            =   150
      TabIndex        =   11
      Top             =   1125
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   4498
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   6960
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":2A05
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":2C25
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":2E45
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":3065
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":327F
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":349F
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":36BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":3C19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":4173
            Key             =   "View"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":438F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":45A9
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":47C9
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   7560
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":49E9
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":4C09
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":4E29
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":5049
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":5263
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":5483
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":56A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":5BFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":6157
            Key             =   "View"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":6373
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":658D
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":67AD
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H8000000E&
      Height          =   1140
      Left            =   7605
      ScaleHeight     =   1080
      ScaleWidth      =   3855
      TabIndex        =   7
      Top             =   1380
      Width           =   3915
      Begin VB.HScrollBar hsb 
         Height          =   225
         Left            =   285
         TabIndex        =   10
         Top             =   4035
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.VScrollBar vsb 
         Height          =   510
         Left            =   60
         TabIndex        =   9
         Top             =   3750
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.PictureBox picDraw 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2265
         Left            =   225
         Picture         =   "frmDefQuery.frx":69CD
         ScaleHeight     =   2235
         ScaleWidth      =   2595
         TabIndex        =   8
         Top             =   150
         Width           =   2625
      End
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   1605
      Left            =   3150
      TabIndex        =   6
      Top             =   3780
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   2831
      View            =   3
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "标题文本"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "标题字体"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "标题位置"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "标题隐藏"
         Object.Width           =   1746
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "返回页首"
         Object.Width           =   1746
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "标题图标"
         Object.Width           =   1746
      EndProperty
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2850
      ScaleHeight     =   315
      ScaleWidth      =   6435
      TabIndex        =   4
      Top             =   720
      Width           =   6435
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   180
         Left            =   105
         TabIndex        =   5
         Top             =   60
         Width           =   630
      End
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   1905
      Top             =   4740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":7B373
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":7D07D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":7FDCF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   795
      Top             =   4650
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":821B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":83EBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":86C0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQuery.frx":88FEF
            Key             =   "folder"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   885
      Left            =   3390
      TabIndex        =   3
      Top             =   1245
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   1561
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "页面名称"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "编码"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "简码"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "页面风格"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "固定页面"
         Object.Width           =   1587
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   8880
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
         TabIndex        =   1
         Top             =   30
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
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
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "页面"
               Key             =   "页面"
               Object.ToolTipText     =   "页面"
               Object.Tag             =   "页面"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "项目"
               Key             =   "项目"
               Object.ToolTipText     =   "增加页面的组成项目"
               Object.Tag             =   "项目"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "删除"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "上移"
               Key             =   "上移"
               Object.ToolTipText     =   "项目顺序上移"
               Object.Tag             =   "上移"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "下移"
               Key             =   "下移"
               Object.ToolTipText     =   "项目顺序下移"
               Object.Tag             =   "下移"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "查看"
               Object.ToolTipText     =   "页面/项目查看方式"
               Object.Tag             =   "查看"
               ImageIndex      =   9
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "大图标"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "小图标"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "列表"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "详细资料"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "效果"
               Key             =   "效果"
               Object.ToolTipText     =   "页面的生成效果"
               Object.Tag             =   "效果"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   12
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image imgY 
      Height          =   60
      Left            =   2925
      MousePointer    =   7  'Size N S
      Top             =   2745
      Width           =   2955
   End
   Begin VB.Image imgX 
      Height          =   4230
      Left            =   2580
      MousePointer    =   9  'Size W E
      Top             =   840
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnusplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileUpatePage 
         Caption         =   "更新查询页面(&U)"
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
         Caption         =   "增加分类(&C)"
      End
      Begin VB.Menu mnuEditNew 
         Caption         =   "增加页面(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditNewItem 
         Caption         =   "增加项目(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&R)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditUp 
         Caption         =   "项目顺序上移(&U)"
      End
      Begin VB.Menu mnuEditDown 
         Caption         =   "项目顺序下移(&D)"
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditExpand 
         Caption         =   "加长下级编码(&E)"
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
      Begin VB.Menu mnuViewPreview 
         Caption         =   "页面效果(&V)"
         Shortcut        =   ^V
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
      Begin VB.Menu mnuHelpSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmDefQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnFist As Boolean
Private mintColumn As Integer
Private mstrKey As String

Public Sub RefreshClass(ByVal strKey As String)

'功能：编辑刷新分类

    Call mnuViewRefresh_Click
    
    On Error Resume Next
    
    tvw.Nodes("K" & strKey).Selected = True
    tvw.Nodes("K" & strKey).EnsureVisible
    
    mstrKey = ""
    If tvw.SelectedItem Is Nothing Then Exit Sub
    Call tvw_NodeClick(tvw.SelectedItem)
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    Call lvw_ItemClick(lvw.SelectedItem)
    
End Sub

Public Sub RefreshPage(ByVal strKey As String)

'功能：编辑刷新新面

    If tvw.SelectedItem Is Nothing Then Exit Sub
    
    mstrKey = ""
    Call tvw_NodeClick(tvw.SelectedItem)
    
    On Error Resume Next
    
    lvw.ListItems("K" & strKey).Selected = True
    lvw.ListItems("K" & strKey).EnsureVisible
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    Call lvw_ItemClick(lvw.SelectedItem)
    
End Sub

Public Sub RefreshItem(ByVal strKey As String)

'功能：编辑刷新新面

    If lvw.SelectedItem Is Nothing Then Exit Sub
        
    Call lvw_ItemClick(lvw.SelectedItem)
    
    On Error Resume Next
    
    lvwItem.ListItems("K" & strKey).Selected = True
    lvwItem.ListItems("K" & strKey).EnsureVisible
    
    
    
End Sub

Private Sub Form_Activate()
    If mblnFist = False Then Exit Sub
    mblnFist = False
    
    Call mnuViewRefresh_Click
    Call AdjustEnabled
End Sub

Private Sub Form_Load()
    mblnFist = True
    
    RestoreWinState Me, App.ProductName
    lblTitle.Caption = ""
    picBack.Visible = False
    
    Call ReadRegister
    Call ModulePrivs
    
End Sub

Private Sub Form_Resize()
    '根据窗体状态,调整窗体中各控件的显示位置
    Dim sglCbrH As Single
    Dim sglStbH As Single
    
    On Error Resume Next
    sglCbrH = IIf(cbrThis.Visible, cbrThis.Height, 0)
    sglStbH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    Call ResizeControl(tvw, 0, sglCbrH, imgX.Left, Me.ScaleHeight - sglStbH - sglCbrH)
    
    With lvw
        .Left = imgX.Left + imgX.Width
        .Top = tvw.Top
        .Width = Me.ScaleWidth - .Left
        .Height = imgY.Top - .Top
    End With
    With picTitle
        .Left = lvw.Left
        .Top = imgY.Top + imgY.Height
        .Width = lvw.Width
    End With
    
    Call ResizeControl(lvwItem, picTitle.Left, picTitle.Top + picTitle.Height + 15, lvw.Width, Me.ScaleHeight - sglStbH - picTitle.Top - picTitle.Height - 15)
    
    Call ResizeControl(imgX, imgX.Left, lvw.Top, imgX.Width, lvw.Height)
    Call ResizeControl(imgY, picTitle.Left, imgY.Top, picTitle.Width, imgY.Height)
    
    'Call ResizeControl(picDraw, 0, 0, picDraw.Width, picDraw.Height)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call WriteRegister
    SaveWinState Me, App.ProductName
End Sub

Private Sub imgX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    imgX.Left = imgX.Left + X
    If imgX.Left < 1000 Then imgX.Left = 1000
    If Me.Width - imgX.Left - imgX.Width < 1000 Then imgX.Left = Me.Width - imgX.Width - 1000
    
    Form_Resize
End Sub

Private Sub imgY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    imgY.Top = imgY.Top + Y
    If imgY.Top < 1800 Then imgY.Top = 1800
    If Me.Height - imgY.Top - imgY.Height < 2100 Then imgY.Top = Me.Height - imgY.Height - 2100
    
    Form_Resize
End Sub

Private Sub lvw_Click()
    Call AdjustEnabled
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then
        lvw.SortOrder = IIf(lvw.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvw.SortKey = mintColumn
        lvw.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvw_DblClick()
    If mnuEditModify.Visible And mnuEditModify.Enabled Then Call mnuEditModify_Click
End Sub

Private Sub lvw_GotFocus()
    Call SetView(lvw, lvw.View)
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim svrKey As String
    
    svrKey = SaveLvwItem(lvwItem)
    lblTitle.Caption = Item.Text & "的组成项目"
    Call LoadPageItemList(Val(Mid(Item.Key, 2)))
    
    Call RestoreLvwItem(lvwItem, svrKey)
    Call AdjustEnabled
    Call LoadStatus
End Sub

Private Sub lvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And mnuEdit.Visible Then
        mnuEditClass.Visible = False
        mnuEditModify.Visible = True
        mnuEditDelete.Visible = True
        
        mnuEditNew.Visible = True
        mnuEditNewItem.Visible = False
        
        mnuEditDown.Visible = False
        mnuEditUp.Visible = False
        mnuEditExpand.Visible = False
        
        mnuEdit_1.Visible = False
        mnuEdit_2.Visible = False
        
        Call lvw_Click
        
        Me.PopupMenu mnuEdit, 2
        
        Call DisEditMenu(True)
    End If
End Sub

Private Sub lvwItem_Click()
    Call AdjustEnabled
End Sub

Private Sub lvwItem_DblClick()
    If mnuEdit.Visible And mnuEditModify.Enabled Then Call mnuEditModify_Click
End Sub

Private Sub lvwItem_GotFocus()
    Call SetView(lvwItem, lvwItem.View)
End Sub

Private Sub lvwItem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call AdjustEnabled
End Sub

Private Sub lvwItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And mnuEdit.Visible Then
        Call lvwItem_Click
        
        mnuEditClass.Visible = False
        mnuEditModify.Visible = True
        mnuEditDelete.Visible = True
        
        mnuEditNew.Visible = False
        mnuEditNewItem.Visible = True
        
        mnuEditDown.Visible = False
        mnuEditUp.Visible = False
        mnuEditExpand.Visible = False
        
        mnuEdit_1.Visible = False
        mnuEdit_2.Visible = False
        
        
        Me.PopupMenu mnuEdit, 2
        
        Call DisEditMenu(True)
    End If
End Sub

Private Sub mnuEditClass_Click()
    Call EditFolder(1)
End Sub

Private Sub mnuEditDelete_Click()
    Dim vIndex As Long
    
            
    If Me.ActiveControl Is tvw Then
        '删除页面分类,并同时删除此分类下的所有页面目录
        
        If tvw.SelectedItem Is Nothing Then Exit Sub
        If tvw.SelectedItem.Key = "K0" Then Exit Sub
        
        If MsgBox("确认要删除页面分类[" & tvw.SelectedItem.Text & "]及页面？" & vbCrLf & "注意：如果下级有固定页面，必须先移走固定页面才能删除！", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Sub
        
        gstrSQL = "zl_咨询页面目录_delete(" & Val(Mid(tvw.SelectedItem.Key, 2)) & ")"
        
    ElseIf Me.ActiveControl Is lvw Then
        '删除页面,并同时删除页面下的所有组成项目
        
        If lvw.SelectedItem Is Nothing Then Exit Sub
        
        If MsgBox("确认要删除页面[" & lvw.SelectedItem.Text & "]及页面组成项目？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Sub
        
        gstrSQL = "select 1 from 咨询段落链接 where 链接页面=" & Val(Mid(lvw.SelectedItem.Key, 2))
        
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If gRs.BOF = False Then
            MsgBox "此页面已经被其它页面所链接，不能删除！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        gstrSQL = "zl_咨询页面目录_delete(" & Val(Mid(lvw.SelectedItem.Key, 2)) & ")"
    Else
        '删除页面的组成项目
        
        If lvwItem.SelectedItem Is Nothing Then Exit Sub
        If MsgBox("确认要删除页面[" & lvw.SelectedItem.Text & "]组成项目[" & lvwItem.SelectedItem.Text & "]？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Sub
        
        gstrSQL = "select 1 from 咨询段落链接 where 链接页面=" & Val(Mid(lvw.SelectedItem.Key, 2)) & " and 页内段号=" & Val(Mid(lvwItem.SelectedItem.Key, 2))
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If gRs.BOF = False Then
            MsgBox "此项目已经被其它页面所链接，不能删除！", vbInformation, gstrSysName
            Exit Sub
        End If
        gstrSQL = "zl_咨询段落目录_delete(" & Val(Mid(lvw.SelectedItem.Key, 2)) & "," & Val(Mid(lvwItem.SelectedItem.Key, 2)) & ")"
    End If
    
    On Error GoTo errHand
    
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    '确定并选中下一个节点或选项
    If Me.ActiveControl Is tvw Then
        
        vIndex = tvw.SelectedItem.Index
        tvw.Nodes.Remove tvw.SelectedItem.Index
        Call NextTvwPos(tvw, vIndex)
        If Not (tvw.SelectedItem Is Nothing) Then Call tvw_NodeClick(tvw.SelectedItem)
    ElseIf Me.ActiveControl Is lvw Then
        lvwItem.ListItems.Clear
        
        vIndex = lvw.SelectedItem.Index
        lvw.ListItems.Remove lvw.SelectedItem.Index
        Call NextLvwPos(lvw, vIndex)
        If Not (lvw.SelectedItem Is Nothing) Then Call lvw_ItemClick(lvw.SelectedItem)
    Else
        vIndex = lvw.SelectedItem.Index
        
        lvwItem.ListItems.Remove lvwItem.SelectedItem.Index
        Call NextLvwPos(lvwItem, vIndex)
    End If
    
    Call AdjustEnabled
    Call LoadStatus
    
    Exit Sub
errHand:
    
    If ErrCenter() = -1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditDown_Click()
    '将当前的项目向上移一行，同时更新数据库
    Dim lngPageNo As Long
    Dim svrAry(7) As String
    Dim intPre As Long
    Dim strSQL(3) As String
          
    If lvw.SelectedItem Is Nothing Then Exit Sub
    If lvwItem.SelectedItem Is Nothing Then Exit Sub
    
    lngPageNo = Val(Mid(lvw.SelectedItem.Key, 2))
    intPre = lvwItem.SelectedItem.Index + 1
    
    If intPre < lvwItem.ListItems.Count + 1 Then
        strSQL(0) = "zl_咨询段落目录_adjust(" & lngPageNo & "," & Val(Mid(lvwItem.SelectedItem.Key, 2)) & ",0)"
        strSQL(1) = "zl_咨询段落目录_adjust(" & lngPageNo & "," & Val(Mid(lvwItem.ListItems(intPre).Key, 2)) & "," & Val(Mid(lvwItem.SelectedItem.Key, 2)) & ")"
        strSQL(2) = "zl_咨询段落目录_adjust(" & lngPageNo & ",0," & Val(Mid(lvwItem.ListItems(intPre).Key, 2)) & ")"
        
        On Error GoTo errHand
        gcnOracle.BeginTrans
        
        Call zlDatabase.ExecuteProcedure(strSQL(0), Me.Caption)
        Call zlDatabase.ExecuteProcedure(strSQL(1), Me.Caption)
        Call zlDatabase.ExecuteProcedure(strSQL(2), Me.Caption)
        gcnOracle.CommitTrans
        On Error GoTo 0
        
        svrAry(0) = lvwItem.ListItems(intPre).Text
        svrAry(1) = lvwItem.ListItems(intPre).SubItems(1)
        svrAry(2) = lvwItem.ListItems(intPre).SubItems(2)
        svrAry(3) = lvwItem.ListItems(intPre).SubItems(3)
        svrAry(4) = lvwItem.ListItems(intPre).SubItems(4)
        svrAry(5) = lvwItem.ListItems(intPre).SubItems(5)
        svrAry(6) = lvwItem.ListItems(intPre).Tag
        
        lvwItem.ListItems(intPre).Text = lvwItem.SelectedItem.Text
        lvwItem.ListItems(intPre).SubItems(1) = lvwItem.SelectedItem.SubItems(1)
        lvwItem.ListItems(intPre).SubItems(2) = lvwItem.SelectedItem.SubItems(2)
        lvwItem.ListItems(intPre).SubItems(3) = lvwItem.SelectedItem.SubItems(3)
        lvwItem.ListItems(intPre).SubItems(4) = lvwItem.SelectedItem.SubItems(4)
        lvwItem.ListItems(intPre).SubItems(5) = lvwItem.SelectedItem.SubItems(5)
        lvwItem.ListItems(intPre).Tag = lvwItem.SelectedItem.Tag
        
        lvwItem.SelectedItem.Text = svrAry(0)
        lvwItem.SelectedItem.SubItems(1) = svrAry(1)
        lvwItem.SelectedItem.SubItems(2) = svrAry(2)
        lvwItem.SelectedItem.SubItems(3) = svrAry(3)
        lvwItem.SelectedItem.SubItems(4) = svrAry(4)
        lvwItem.SelectedItem.SubItems(5) = svrAry(5)
        lvwItem.SelectedItem.Tag = svrAry(6)
        
        lvwItem.SelectedItem.Tag = svrAry(5)
        
        lvwItem.ListItems(intPre).Selected = True
        Call AdjustEnabled
    End If
    Exit Sub
errHand:
    
    gcnOracle.RollbackTrans
    
    If ErrCenter() = -1 Then Resume
    

End Sub

Private Sub mnuEditExpand_Click()
    Dim strTemp As String
    Dim str父编码 As String
    Dim intChild As Integer '目前下级的编码长度
    Dim intNew As Integer '目前最长的

    On Error GoTo ErrHandle
    
    With tvw.SelectedItem
        If .Key = "K0" Then
            str父编码 = ""
            intNew = GetDownCodeLength("", "咨询页面目录")
            intChild = GetLocalCodeLength("", "咨询页面目录")
        Else
            str父编码 = Mid(.Text, 2, InStr(.Text, "】") - 2)
            intNew = GetDownCodeLength(Mid(.Key, 2), "咨询页面目录") ')
            intChild = GetLocalCodeLength(Mid(.Key, 2), "咨询页面目录")
        End If
        If intNew = 0 Or intChild = 0 Then Exit Sub
        If intNew = 10 Then
            MsgBox "不能再加长编码，某一个下级已经用足了长度。", vbExclamation, gstrSysName
            Exit Sub
        End If
        
        intNew = frm编码长度.GetLength(intChild, 10 - (intNew - intChild))
        If intNew = 0 Then Exit Sub
        strTemp = str父编码 & String(intNew - intChild, "0")
        If .Key = "K0" Then
            gstrSQL = "zl_咨询页面目录_EXPAND('" & strTemp & "'," & Len(str父编码) + 1 & ",0)"
        Else
            gstrSQL = "zl_咨询页面目录_EXPAND('" & strTemp & "'," & Len(str父编码) + 1 & "," & Val(Mid(.Key, 2)) & ")"
        End If
                                    
                            
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                
        Call FillTree
        
    End With
    
    Exit Sub
    
ErrHandle:
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
    
End Sub

Private Sub mnuEditModify_Click()
    If Me.ActiveControl Is lvw Then
        Call EditPage(2)
    ElseIf Me.ActiveControl Is lvwItem Then
        Call EditPageItem(2)
    Else
        Call EditFolder(2)
    End If
End Sub

Private Sub mnuEditNew_Click()
    Call EditPage(1)
End Sub

Private Sub mnuEditNewItem_Click()
    Call EditPageItem(1)
End Sub

Private Sub mnuEditUp_Click()
    '将当前的项目向上移一行，同时更新数据库
    Dim svrAry(7) As String
    Dim intPre As Long
    Dim strSQL(3) As String
    Dim lngPageNo As Long
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    If lvwItem.SelectedItem Is Nothing Then Exit Sub
    
    lngPageNo = Val(Mid(lvw.SelectedItem.Key, 2))
    intPre = lvwItem.SelectedItem.Index - 1
    
    If intPre > 0 Then
        strSQL(0) = "zl_咨询段落目录_adjust(" & lngPageNo & "," & Val(Mid(lvwItem.SelectedItem.Key, 2)) & ",0)"
        strSQL(1) = "zl_咨询段落目录_adjust(" & lngPageNo & "," & Val(Mid(lvwItem.ListItems(intPre).Key, 2)) & "," & Val(Mid(lvwItem.SelectedItem.Key, 2)) & ")"
        strSQL(2) = "zl_咨询段落目录_adjust(" & lngPageNo & ",0," & Val(Mid(lvwItem.ListItems(intPre).Key, 2)) & ")"
        
        On Error GoTo errHand
        gcnOracle.BeginTrans
                
        Call zlDatabase.ExecuteProcedure(strSQL(0), Me.Caption)
        Call zlDatabase.ExecuteProcedure(strSQL(1), Me.Caption)
        Call zlDatabase.ExecuteProcedure(strSQL(2), Me.Caption)
        
        gcnOracle.CommitTrans
        On Error GoTo 0
        
        svrAry(0) = lvwItem.ListItems(intPre).Text
        svrAry(1) = lvwItem.ListItems(intPre).SubItems(1)
        svrAry(2) = lvwItem.ListItems(intPre).SubItems(2)
        svrAry(3) = lvwItem.ListItems(intPre).SubItems(3)
        svrAry(4) = lvwItem.ListItems(intPre).SubItems(4)
        svrAry(5) = lvwItem.ListItems(intPre).SubItems(5)
        svrAry(6) = lvwItem.ListItems(intPre).Tag
        
        lvwItem.ListItems(intPre).Text = lvwItem.SelectedItem.Text
        lvwItem.ListItems(intPre).SubItems(1) = lvwItem.SelectedItem.SubItems(1)
        lvwItem.ListItems(intPre).SubItems(2) = lvwItem.SelectedItem.SubItems(2)
        lvwItem.ListItems(intPre).SubItems(3) = lvwItem.SelectedItem.SubItems(3)
        lvwItem.ListItems(intPre).SubItems(4) = lvwItem.SelectedItem.SubItems(4)
        lvwItem.ListItems(intPre).SubItems(5) = lvwItem.SelectedItem.SubItems(5)
        lvwItem.ListItems(intPre).Tag = lvwItem.SelectedItem.Tag
        
        lvwItem.SelectedItem.Text = svrAry(0)
        lvwItem.SelectedItem.SubItems(1) = svrAry(1)
        lvwItem.SelectedItem.SubItems(2) = svrAry(2)
        lvwItem.SelectedItem.SubItems(3) = svrAry(3)
        lvwItem.SelectedItem.SubItems(4) = svrAry(4)
        lvwItem.SelectedItem.SubItems(5) = svrAry(5)
        lvwItem.SelectedItem.Tag = svrAry(6)
        
        lvwItem.ListItems(intPre).Selected = True
        Call AdjustEnabled
    End If
    Exit Sub
errHand:
    
    gcnOracle.RollbackTrans
    
    If ErrCenter() = -1 Then Resume
    
End Sub

Private Sub mnuFileExcel_Click()
    Call PrintObject(3)
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Sub mnuFilePreView_Click()
    Call PrintObject(2)
End Sub

Private Sub mnuFilePrint_Click()
    Call PrintObject(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFileUpatePage_Click()
    Call gfrmMain.FrameDefault.RefreshPage
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTopic_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    If Me.ActiveControl Is lvw Then
        Call SetView(lvw, Index)
    Else
        Call SetView(lvwItem, Index)
    End If
End Sub

Private Sub mnuViewPreview_Click()
    '
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    Call CheckPicture
    Call frmPagePreview.ShowPreview(Me, Val(Mid(lvw.SelectedItem.Key, 2)))
End Sub

Private Sub mnuViewRefresh_Click()
    Dim svrKey As String
    
    svrKey = SaveLvwItem(lvw)
    
    Call FillTree
    If Not (tvw.SelectedItem Is Nothing) Then Call tvw_NodeClick(tvw.SelectedItem)

    Call RestoreLvwItem(lvw, svrKey)
    Call LoadStatus
    lvwItem.ListItems.Clear

    If Not (lvw.SelectedItem Is Nothing) Then Call lvw_ItemClick(lvw.SelectedItem)
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
    Dim i As Long
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(i).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub

'Private Sub picDraw_Paint()
'    DrawText picDraw, 3000, 2400, "测试"
'End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "预览"
        Call mnuFilePreView_Click
    Case "打印"
        Call mnuFilePrint_Click
    Case "页面"
        Call mnuEditNew_Click
    Case "项目"
        Call mnuEditNewItem_Click
    Case "修改"
        Call mnuEditModify_Click
    Case "删除"
        Call mnuEditDelete_Click
    Case "上移"
        Call mnuEditUp_Click
    Case "下移"
        Call mnuEditDown_Click
    Case "查看"
        If Me.ActiveControl Is lvw Then
            Call SetView(lvw, IIf(lvw.View = 3, 0, lvw.View + 1))
        Else
            Call SetView(lvwItem, IIf(lvwItem.View = 3, 0, lvwItem.View + 1))
        End If
    Case "效果"
        Call mnuViewPreview_Click
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub EditPage(ByVal Mode As Byte)
    Dim lngKey As Long
    
    If Mode = 2 And Not (lvw.SelectedItem Is Nothing) Then lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
    
    '主页面设置
    If Not (lvw.SelectedItem Is Nothing) Then
        If lvw.SelectedItem.Key = "K0" And Mode = 2 Then
            frmHomePage.Show 1, Me
            Exit Sub
        End If
    End If
    
    If frmDefQueryPage.ShowPageEdit(Me, lngKey, Val(Mid(tvw.SelectedItem.Key, 2))) Then
        
    End If
End Sub

Private Sub EditFolder(ByVal Mode As Byte)
    Dim lngKey As Long
    Dim lngUpKey As Long
    
    lngUpKey = Val(Mid(tvw.SelectedItem.Key, 2))
    If Mode = 2 And Not (tvw.SelectedItem Is Nothing) Then
        lngKey = Val(Mid(tvw.SelectedItem.Key, 2))
        If Not (tvw.SelectedItem.Parent Is Nothing) Then lngUpKey = Val(Mid(tvw.SelectedItem.Parent.Key, 2))
    End If
        
    If frmDefQueryClass.ShowEdit(Me, lngKey, lngUpKey) Then
        
    End If
End Sub

Private Sub EditPageItem(ByVal Mode As Byte)
    Dim lngKey As Long
    Dim lngPageNo As Long
        
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    lngPageNo = Val(Mid(lvw.SelectedItem.Key, 2))
    
    'If lvw.SelectedItem.Tag = "1" And lngPageNo >= 0 Then Exit Sub
        
    If Mode = 2 And Not (lvwItem.SelectedItem Is Nothing) Then lngKey = Val(Mid(lvwItem.SelectedItem.Key, 2))

    If frmDefQueryItem.ShowItemEdit(Me, lngPageNo, lngKey) Then
        
    End If
End Sub

Private Sub LoadPage(ByVal lngKey As Long)
    '加载页面数据
    Dim Itmx As ListItem
    
    On Error GoTo errHand
    
    lvw.ListItems.Clear
    lvwItem.ListItems.Clear
    lblTitle.Caption = ""
    
    If lngKey = 0 Then
        gstrSQL = "Select 页面序号,上级序号,编码,页面名称,简码,固定页面,页面风格,宣传标语,页面背景,背景音乐,命令参数,末级 from 咨询页面目录  where 末级=1 and (上级序号=0 OR 上级序号 is null) order by 固定页面 desc,页面序号"
    Else
        gstrSQL = "Select 页面序号,上级序号,编码,页面名称,简码,固定页面,页面风格,宣传标语,页面背景,背景音乐,命令参数,末级 from 咨询页面目录  where 末级=1 and 上级序号=[1] order by 固定页面 desc,页面序号"
    End If
    
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If gRs.BOF = False Then
        While Not gRs.EOF
            Set Itmx = lvw.ListItems.Add(, "K" & gRs!页面序号, IIf(IsNull(gRs!页面名称), "", gRs!页面名称), 1, 1)
            Itmx.Tag = IIf(IsNull(gRs!固定页面), 0, gRs!固定页面)
            Itmx.SubItems(1) = IIf(IsNull(gRs!编码), "", gRs!编码)
            Itmx.SubItems(2) = IIf(IsNull(gRs!简码), "", gRs!简码)
            Itmx.SubItems(4) = IIf(Itmx.Tag = "1", "√", "")
            If Val(Itmx.Tag) = 1 Then
                Itmx.SmallIcon = 2
                Itmx.Icon = 2
            End If
            Select Case IIf(IsNull(gRs!页面风格), 0, gRs!页面风格)
            Case 0
                Itmx.SubItems(3) = "标准"
            End Select
            
            gRs.MoveNext
        Wend
    End If
    
    If Not (lvw.SelectedItem Is Nothing) Then Call lvw_ItemClick(lvw.SelectedItem)
    
    Exit Sub
    
errHand:
    If ErrCenter() = -1 Then Exit Sub
    Call SaveErrLog
End Sub

Private Sub LoadPageItemList(ByVal PageNo As Long)
'加载页面的组成项目列表
    Dim Itmx As ListItem
    
    On Error GoTo errHand
    
    lvwItem.ListItems.Clear
    
'    If lvw.SelectedItem.Tag = "1" And PageNo > 0 Then Exit Sub
    gstrSQL = "select 段落序号,标题文本,标题图标,标题隐藏,标题位置,标题字体,返回页首 from 咨询段落目录 where 页面序号=[1] order by 段落序号"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then
        While Not gRs.EOF
            Set Itmx = lvwItem.ListItems.Add(, "K" & gRs!段落序号, IIf(IsNull(gRs!标题文本), "", gRs!标题文本), 3, 3)
            Itmx.SubItems(1) = IIf(IsNull(gRs!标题字体), "宋体;12;0;0;0", gRs!标题字体)
            Select Case IIf(IsNull(gRs!标题位置), 0, gRs!标题位置)
            Case 0
                Itmx.SubItems(2) = "左对齐"
            Case 1
                Itmx.SubItems(2) = "右对齐"
            Case 2
                Itmx.SubItems(2) = "居中"
            End Select
            Itmx.SubItems(3) = IIf(IIf(IsNull(gRs!标题隐藏), 0, gRs!标题隐藏) = 0, "", "√")
            Itmx.SubItems(4) = IIf(IIf(IsNull(gRs!返回页首), 0, gRs!返回页首) = 0, "", "√")
            Itmx.SubItems(5) = IIf(IIf(IsNull(gRs!标题图标), 0, gRs!标题图标) = 0, "", "√")
            gRs.MoveNext
        Wend
    End If
    Exit Sub
errHand:
    If ErrCenter() = -1 Then Exit Sub
    Call SaveErrLog
End Sub

Private Sub SetView(lvwObj As ListView, ByVal View As Byte)
    lvwObj.View = View
    
    mnuViewIcon(0).Checked = False
    mnuViewIcon(1).Checked = False
    mnuViewIcon(2).Checked = False
    mnuViewIcon(3).Checked = False
    
    mnuViewIcon(View).Checked = True
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Text
    Case "大图标"
        Call mnuViewIcon_Click(0)
    Case "小图标"
        Call mnuViewIcon_Click(1)
    Case "列表"
        Call mnuViewIcon_Click(2)
    Case "详细资料"
        Call mnuViewIcon_Click(3)
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu mnuViewTool, 2
End Sub

Private Sub PrintObject(ByVal intMode As Byte)
    '---------------------------------------------------
    '功能：    根据屏幕组织表上附加项目，打印预览
    '参数：
    '     intMode: 2表示预览 1打印 3输出到EXCEL
    '返回：
    '---------------------------------------------------
    
    Dim objPrint As New zlPrintLvw
    Dim objRow As New zlTabAppRow
        
    If lvw.SelectedItem Is Nothing Then Exit Sub
    If UserInfo.姓名 = "" Then Call GetUserInfo
    
    If Me.ActiveControl Is lvw Then
        objPrint.Title = "用户定义页面清单"
        Set objPrint.Body.objData = lvw
    Else
        If lvwItem.SelectedItem Is Nothing Then Exit Sub
        objPrint.Title = "页面:" & lvw.SelectedItem.Text & "的组成项目"
        Set objPrint.Body.objData = lvwItem
    End If
    objPrint.BelowAppItems.Add "打印人:" & UserInfo.姓名
    objPrint.BelowAppItems.Add "打印时间:" & Format(zlDatabase.Currentdate, "YYYY年MM月DD日")
    objPrint.Footer = "第[页码]页;;"

    If intMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
        End Select
    Else
        zlPrintOrViewLvw objPrint, intMode
    End If

End Sub

Private Sub ModulePrivs()
'功能:根据模块权限,处理功能项的隐藏或显示
'     权限有:增删改
    
'    mnuEdit.Visible = True
'
'    If InStr(gstrPrivs, "增删改") = 0 Then
'        mnuEdit.Visible = False
'
'        tbrThis.Buttons("页面").Visible = False
'        tbrThis.Buttons("项目").Visible = False
'        tbrThis.Buttons("修改").Visible = False
'        tbrThis.Buttons("删除").Visible = False
'        tbrThis.Buttons("Split_2").Visible = False
'
'        tbrThis.Buttons("上移").Visible = False
'        tbrThis.Buttons("下移").Visible = False
'        tbrThis.Buttons("Split_3").Visible = False
'    End If
End Sub

Private Sub AdjustEnabled()
'功能:调整功能菜单或按钮的可用状态
    mnuFilePreView.Enabled = True
    mnuFilePrint.Enabled = True
    mnuFileExcel.Enabled = True
    mnuEditDelete.Enabled = True
    mnuEditModify.Enabled = True
    mnuEditNew.Enabled = True
    mnuEditNewItem.Enabled = True
    mnuViewPreview.Enabled = True
    
    mnuEditUp.Enabled = True
    mnuEditDown.Enabled = True
    
    If Not (lvw.SelectedItem Is Nothing) Then
        If lvw.SelectedItem.Tag = "1" And Val(Mid(lvw.SelectedItem.Key, 2)) >= 0 Then mnuViewPreview.Enabled = False
    End If
                
    If Me.ActiveControl Is tvw Then
        If tvw.SelectedItem.Key = "K0" Then
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
        End If
        
        If lvw.SelectedItem Is Nothing Then
            mnuFilePreView.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFileExcel.Enabled = False
        End If
        
    ElseIf Me.ActiveControl Is lvw Then
        If lvw.SelectedItem Is Nothing Then
            mnuFilePreView.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFileExcel.Enabled = False
            
            mnuEditNewItem.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
        Else
            If lvw.SelectedItem.Tag = "1" And Val(Mid(lvw.SelectedItem.Key, 2)) >= 0 Then
                mnuEditNewItem.Enabled = False
                mnuEditDelete.Enabled = False
                mnuViewPreview.Enabled = False
            End If
        End If
    Else
        
        If lvwItem.ListItems.Count = 0 Then
            mnuFilePreView.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFileExcel.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
            
            mnuEditDown.Enabled = False
            mnuEditUp.Enabled = False
        End If
        
        If Not (lvw.SelectedItem Is Nothing) Then
            If lvw.SelectedItem.Tag = "1" And Val(Mid(lvw.SelectedItem.Key, 2)) >= 0 Then
                mnuEditDelete.Enabled = False
                mnuEditNewItem.Enabled = False
            End If
                
            If lvw.SelectedItem.Tag = "1" And Val(Mid(lvw.SelectedItem.Key, 2)) >= 0 Then
                mnuFilePreView.Enabled = False
                mnuFilePrint.Enabled = False
                mnuFileExcel.Enabled = False
                mnuEditModify.Enabled = (lvwItem.ListItems.Count > 0)
                                
                mnuEditDown.Enabled = False
                mnuEditUp.Enabled = False
            End If
        End If
        
    End If
    
    If Not (lvwItem.SelectedItem Is Nothing) Then
        If lvwItem.SelectedItem.Index - 1 <= 0 Then mnuEditUp.Enabled = False
        If lvwItem.SelectedItem.Index + 1 > lvwItem.ListItems.Count Then mnuEditDown.Enabled = False
    Else
        mnuEditUp.Enabled = False
        mnuEditDown.Enabled = False
    End If
    
    If lvw.SelectedItem Is Nothing Then
        mnuEditNewItem.Enabled = False
        mnuViewPreview.Enabled = False
    End If
                                                
    tbrThis.Buttons("预览").Enabled = mnuFilePreView.Enabled
    tbrThis.Buttons("打印").Enabled = mnuFilePrint.Enabled
    tbrThis.Buttons("页面").Enabled = mnuEditNew.Enabled
    tbrThis.Buttons("项目").Enabled = mnuEditNewItem.Enabled
    tbrThis.Buttons("修改").Enabled = mnuEditModify.Enabled
    tbrThis.Buttons("删除").Enabled = mnuEditDelete.Enabled
        
    tbrThis.Buttons("上移").Enabled = mnuEditUp.Enabled
    tbrThis.Buttons("下移").Enabled = mnuEditDown.Enabled
    
    tbrThis.Buttons("效果").Enabled = mnuViewPreview.Enabled
End Sub

Private Sub ReadRegister()
'功能:读取注册表信息
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Sub
    imgX.Left = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\界面\" & Me.Name, "imgX位置", 2385)
End Sub

Private Sub WriteRegister()
'功能:将信息写回注册表
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\界面\" & Me.Name, "imgX位置", imgX.Left
End Sub

Private Sub LoadStatus()
'功能:填写状态栏信息
    Dim vTmp As String
    
    If lvw.ListItems.Count > 0 Then
        If lvwItem.ListItems.Count > 0 Then
            vTmp = "当前共有" & lvw.ListItems.Count & "个自定义页面，页面" & lvw.SelectedItem.Text & "共有由" & lvwItem.ListItems.Count & "个项目组成！"
        Else
            vTmp = "当前共有" & lvw.ListItems.Count & "个自定义页面，页面" & lvw.SelectedItem.Text & "还没有一个项目！"
        End If
    Else
        vTmp = "当前没有用户定义页面！"
    End If
    stbThis.Panels(2).Text = vTmp
End Sub

Private Sub FillTree()
'功能:装入所有页面目录到Tvw控件中
'参数:
    Dim strTemp As String
    Dim strKey As String
    Dim rs目录 As New ADODB.Recordset
    
    
    mstrKey = ""
    
    rs目录.CursorLocation = adUseClient
    rs目录.CursorType = adOpenKeyset
    rs目录.LockType = adLockReadOnly
    
    If Not tvw.SelectedItem Is Nothing Then strKey = tvw.SelectedItem.Key
    
    gstrSQL = "select 页面序号,上级序号,编码,页面名称 from 咨询页面目录 where (末级=0 or 末级 is null) start with 上级序号 is null connect by prior 页面序号 =上级序号"
    Call zlDatabase.OpenRecordset(rs目录, gstrSQL, Me.Caption)
    
    tvw.Nodes.Clear
    tvw.Nodes.Add , , "K0", "所有页面分类", "folder", "folder"
    tvw.Nodes("K0").Sorted = True
    tvw.Nodes("K0").Expanded = True
    tvw.Nodes("K0").Selected = True
    
    Do Until rs目录.EOF
                        
        If IIf(IsNull(rs目录("上级序号")), 0, rs目录("上级序号")) = 0 Then
            tvw.Nodes.Add "K0", tvwChild, "K" & rs目录("页面序号"), "【" & rs目录("编码") & "】" & rs目录("页面名称"), "folder", "folder"
        Else
            tvw.Nodes.Add "K" & IIf(IsNull(rs目录("上级序号")), "0", rs目录("上级序号")), tvwChild, "K" & rs目录("页面序号"), "【" & rs目录("编码") & "】" & rs目录("页面名称"), "folder", "folder"
        End If
                
        tvw.Nodes("K" & rs目录("页面序号")).Sorted = True
        rs目录.MoveNext
    Loop
    
    Dim nod As Node
    
    On Error Resume Next
    
    Set nod = tvw.Nodes(strKey)
    If Err <> 0 Then
        Set nod = tvw.Nodes("ROOT")
        nod.Selected = True
        nod.Expanded = True
        tvw_NodeClick nod
    Else
        Err.Clear
        nod.Selected = True
        nod.Expanded = True
        nod.EnsureVisible
        tvw_NodeClick nod
    End If
End Sub

Private Sub tvw_Click()
    Call AdjustEnabled
End Sub

Private Sub tvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And mnuEdit.Visible Then
        Call tvw_Click
        
        mnuEditClass.Visible = True
        mnuEditModify.Visible = True
        mnuEditDelete.Visible = True
        
        mnuEditNew.Visible = False
        mnuEditNewItem.Visible = False
        
        mnuEditDown.Visible = False
        mnuEditUp.Visible = False
        mnuEditExpand.Visible = True
        
        mnuEdit_1.Visible = False
        mnuEdit_2.Visible = True
        
        Me.PopupMenu mnuEdit, 2
        
        Call DisEditMenu(True)
        
    End If
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    
    If mstrKey <> Node.Key Then
        mstrKey = Node.Key
        Call LoadPage(Val(Mid(Node.Key, 2)))
    End If
    
    Call AdjustEnabled
End Sub

Private Sub DisEditMenu(ByVal blnFlag As Boolean)
    mnuEditClass.Visible = blnFlag
    mnuEditModify.Visible = blnFlag
    mnuEditDelete.Visible = blnFlag
    
    mnuEditNew.Visible = blnFlag
    mnuEditNewItem.Visible = blnFlag
    
    mnuEditDown.Visible = blnFlag
    mnuEditUp.Visible = blnFlag
    mnuEditExpand.Visible = blnFlag
    
    mnuEdit_1.Visible = blnFlag
    mnuEdit_2.Visible = blnFlag
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

