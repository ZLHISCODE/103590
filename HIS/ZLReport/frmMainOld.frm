VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "报表管理"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImgGroup32 
      Left            =   4290
      Top             =   2190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0464
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgGroup16 
      Left            =   4290
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1226
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1380
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1634
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   4845
      Top             =   2190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":178E
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AA8
            Key             =   "Publish"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DC2
            Key             =   "Fixed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20DC
            Key             =   "PubFixed"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23F6
            Key             =   "Bill"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2710
            Key             =   "BillPublish"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   4860
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A2A
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B84
            Key             =   "Publish"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CDE
            Key             =   "Fixed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E38
            Key             =   "PubFixed"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F92
            Key             =   "Bill"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30EC
            Key             =   "BillPublish"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwReport 
      Height          =   5025
      Left            =   2355
      TabIndex        =   2
      Top             =   1590
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8864
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "编号"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "说明"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "修改时间"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "发布时间"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "最后执行时间"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "最后执行人"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "种类"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "类型"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "性能问题数据源"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "简码"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "其他数据连接"
         Object.Width           =   6223
      EndProperty
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11475
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      Key1            =   "cbr_Funcs"
      NewRow1         =   0   'False
      BandForeColor2  =   8388608
      Child2          =   "picSysFind"
      MinWidth2       =   2310
      MinHeight2      =   495
      Width2          =   2370
      UseCoolbarColors2=   0   'False
      Key2            =   "cbr_SysFind"
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin VB.PictureBox picSysFind 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   9075
         ScaleHeight     =   495
         ScaleWidth      =   2370
         TabIndex        =   8
         Top             =   135
         Width           =   2370
         Begin VB.CommandButton cmdNext 
            Caption         =   "下一个(F3)"
            Height          =   350
            Left            =   5450
            TabIndex        =   13
            Top             =   75
            Width           =   1065
         End
         Begin VB.ComboBox cboSys 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   100
            Width           =   2100
         End
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   3900
            MaxLength       =   50
            TabIndex        =   12
            Tag             =   "简码"
            ToolTipText     =   "支持按名称、编号、拼音简码查找"
            Top             =   100
            Width           =   1515
         End
         Begin VB.Label lblSys 
            AutoSize        =   -1  'True
            Caption         =   "系统(&S)"
            ForeColor       =   &H008B0000&
            Height          =   180
            Left            =   60
            TabIndex        =   9
            Top             =   165
            Width           =   630
         End
         Begin VB.Label lblFind 
            AutoSize        =   -1  'True
            Caption         =   "查找报表(&L)"
            ForeColor       =   &H008B0000&
            Height          =   180
            Left            =   2880
            TabIndex        =   11
            Top             =   165
            Width           =   990
         End
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   23
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "执行"
               Key             =   "Report"
               Description     =   "执行"
               Object.ToolTipText     =   "执行报表"
               Object.Tag             =   "执行"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "新增"
               Key             =   "GroupAdd"
               Description     =   "新增"
               Object.ToolTipText     =   "增加报表组"
               Object.Tag             =   "新增"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "GroupModify"
               Description     =   "修改"
               Object.ToolTipText     =   "修改报表"
               Object.Tag             =   "修改"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "GroupDel"
               Description     =   "删除"
               Object.ToolTipText     =   "删除报表组"
               Object.Tag             =   "删除"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Group_"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "新增"
               Key             =   "Add"
               Description     =   "新增"
               Object.ToolTipText     =   "新增自定义报表"
               Object.Tag             =   "新增"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modi"
               Description     =   "修改"
               Object.ToolTipText     =   "修改报表属性"
               Object.Tag             =   "修改"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Del"
               Description     =   "删除"
               Object.ToolTipText     =   "删除当前报表"
               Object.Tag             =   "删除"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Report_"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "设计"
               Key             =   "Design"
               Description     =   "设计"
               Object.ToolTipText     =   "设计报表"
               Object.Tag             =   "设计"
               ImageKey        =   "Design"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Design_"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "向导"
               Key             =   "Guide"
               Description     =   "向导"
               Object.ToolTipText     =   "报表向导"
               Object.Tag             =   "向导"
               ImageKey        =   "Guide"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Guide_"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "发布"
               Key             =   "Publish"
               Description     =   "发布"
               Object.ToolTipText     =   "发布报表"
               Object.Tag             =   "发布"
               ImageIndex      =   9
               Style           =   5
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "取消"
               Key             =   "unPub"
               Description     =   "取消"
               Object.ToolTipText     =   "取消发布"
               Object.Tag             =   "取消"
               ImageIndex      =   10
               Style           =   5
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Pub_"
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "View"
               Description     =   "查看"
               Object.ToolTipText     =   "列表查看方式"
               Object.Tag             =   "查看"
               ImageIndex      =   11
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   5
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Icon"
                     Object.Tag             =   "大图标(&I)"
                     Text            =   "大图标(&I)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Object.Tag             =   "小图标(&S)"
                     Text            =   "小图标(&S)"
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
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "RunLog"
                     Object.Tag             =   "报表日志"
                     Text            =   "报表日志"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   13
            EndProperty
         EndProperty
         Begin MSComctlLib.Toolbar tbrCheck 
            Height          =   720
            Left            =   2610
            TabIndex        =   7
            Top             =   0
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   1270
            ButtonWidth     =   1455
            ButtonHeight    =   1270
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            ImageList       =   "imgGray"
            HotImageList    =   "imgColor"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "性能检查"
                  Key             =   "Check"
                  Description     =   "性能检查"
                  Object.ToolTipText     =   "性能检查"
                  Object.Tag             =   "性能检查"
                  ImageIndex      =   15
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   75
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3246
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3460
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":367A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3894
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":40FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4316
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4530
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":474A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4964
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B7E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D98
            Key             =   "Guide"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4FB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":51CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   705
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5600
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":581A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E68
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6082
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":629C
            Key             =   "Design"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":64B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":66D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":68EA
            Key             =   "View"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B04
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D1E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6F38
            Key             =   "Guide"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7152
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":736C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   2265
      Top             =   1245
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwGroup 
      Height          =   5085
      Left            =   30
      TabIndex        =   4
      Top             =   1530
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   8969
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImgGroup32"
      SmallIcons      =   "ImgGroup16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "编号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "说明"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "发布时间"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "简码"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   6645
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":7586
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15161
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
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
   Begin VB.Label LblReport 
      Alignment       =   2  'Center
      BackColor       =   &H009B6737&
      Caption         =   "报表"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Top             =   1290
      Width           =   2055
   End
   Begin VB.Label lblGroup 
      Alignment       =   2  'Center
      BackColor       =   &H009B6737&
      Caption         =   "报表组"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   1290
      Width           =   2055
   End
   Begin VB.Image ImgSplit_S 
      Height          =   4980
      Left            =   2280
      MousePointer    =   9  'Size W E
      Top             =   1065
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFile_Report 
         Caption         =   "执行报表(&E)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Exp 
         Caption         =   "导出报表(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFile_Imp 
         Caption         =   "导入报表(&I)"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFile_ExpAll 
         Caption         =   "全部导出"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuFile_ImpAll 
         Caption         =   "全部导入"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuFile_PARA_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Para 
         Caption         =   "参数设置"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuFile_IO_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "退出(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit_Add 
         Caption         =   "新增报表(&W)"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuEdit_Modi 
         Caption         =   "修改报表(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "删除报表(&R)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_Report_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Group_Add 
         Caption         =   "新增报表组(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEdit_Group_Modify 
         Caption         =   "修改报表组(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuEdit_Group_Delete 
         Caption         =   "删除报表组(&O)"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuEdit_Group_Setup 
         Caption         =   "设置子报表(&S)"
      End
      Begin VB.Menu mnuEdit_Group_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Design 
         Caption         =   "设计报表(&D)"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEdit_Clear 
         Caption         =   "清除历史数据源(&C)"
      End
      Begin VB.Menu mnuEdit_Design_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Guide 
         Caption         =   "报表向导(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEdit_Guide_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Group_Publish 
         Caption         =   "报表组发布(&P)"
      End
      Begin VB.Menu mnuEdit_Group_unPub 
         Caption         =   "取消发布报表组(&T)"
      End
      Begin VB.Menu mnuEdit_Publish 
         Caption         =   "报表发布(&B)"
         Begin VB.Menu mnuEdit_Publish_Main 
            Caption         =   "到导航台菜单(&1)"
         End
         Begin VB.Menu mnuEdit_Publish_Module 
            Caption         =   "到模块内菜单(&2)"
         End
      End
      Begin VB.Menu mnuEdit_unPub 
         Caption         =   "取消发布(&U)"
         Begin VB.Menu mnuEdit_unPub_Main 
            Caption         =   "从导航台菜单(&1)"
         End
         Begin VB.Menu mnuEdit_unPub_Module 
            Caption         =   "从模块内菜单(&2)"
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&B)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_Tlb_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&L)"
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
      Begin VB.Menu mnuView_View 
         Caption         =   "大图标(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuView_View 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuView_View 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuView_View 
         Caption         =   "详细资料(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOnly 
         Caption         =   "仅显示独立项(&O)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu RunLog 
         Caption         =   "报表日志"
         Index           =   5
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "查找下一个"
         Shortcut        =   {F3}
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
   Begin VB.Menu mnuPopPublish 
      Caption         =   "发布"
      Visible         =   0   'False
      Begin VB.Menu mnuPopPublish_Group 
         Caption         =   "报表组到导航台菜单"
      End
      Begin VB.Menu mnuPopPublish_ReportMain 
         Caption         =   "报表到导航台菜单"
      End
      Begin VB.Menu mnuPopPublish_ReportModule 
         Caption         =   "报表到模块内菜单"
      End
   End
   Begin VB.Menu mnuPopUnpub 
      Caption         =   "取消"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUnpub_Group 
         Caption         =   "报表组从导航台菜单"
      End
      Begin VB.Menu mnuPopUnpub_ReportMain 
         Caption         =   "报表从导航台菜单"
      End
      Begin VB.Menu mnuPopUnpub_ReportModule 
         Caption         =   "报表从模块内菜单"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnItem As Boolean
Private mobjSelItem As ListItem
Private mblnMouseDown As Boolean
Private mblnGrant As Boolean '是否注册了报表增删功能
Private mblnModule As Boolean '是否允许发布到模块
Private mstrRepName As String
Private mstrPreGroup As String
Private mstrFindValue As String     '记录查询文本框的值
Private mrsFind As New ADODB.Recordset
Private Enum CurSelect
    CS_报表组 = 0
    CS_报表 = 1
End Enum
Private mcsActive As CurSelect '活动控件,0-报表组列表，1-报表列表
Private mfrmReportPara As frmReportPara
'SubItems索引
Private Enum ReportCol
    RC_名称 = 0 '不能索引
    RC_编号 = 1
    RC_说明 = 2
    RC_修改时间 = 3
    RC_发布时间 = 4
    RC_最后执行时间 = 5
    RC_最后执行人 = 6
    RC_种类 = 7
    RC_类型 = 8
    RC_性能问题数据源 = 9
    RC_简码 = 10
    RC_其他数据连接 = 11
End Enum
'SubItems索引
Private Enum GroupCol
    GC_名称 = 0 '不能索引
    GC_编号 = 1
    GC_说明 = 2
    GC_发布时间 = 3
    GC_简码 = 4
End Enum

Private Sub cboSys_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    mblnModule = False
    If cboSys.ListIndex <> -1 Then
        If InStr(",0,2,5,7,", cboSys.ItemData(cboSys.ListIndex) \ 100) > 0 Then
            '共享，人事，成本，帐务因定允许发布到模块
            mblnModule = True
        Else
            '其它系统仅10版本允许
            strSQL = "Select 版本号 From zlSystems Where 编号=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, cboSys.ItemData(cboSys.ListIndex))
            If Not rsTmp.EOF Then
                mblnModule = Val(Split(rsTmp!版本号, ".")(0)) >= 10
            End If
        End If
    End If
    On Error GoTo 0
    
    Call mnuView_reFlash_Click
    
    On Error Resume Next
    If lvwReport.Visible Then lvwReport.SetFocus
    On Error GoTo 0
    
    Call SetFuncEnabled(True)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbr_HeightChanged(ByVal NewHeight As Single)
    Call Form_Resize
End Sub

Private Sub GotoLVW(ByVal intCurPosition As Integer, lvwCur As ListView)
    Dim objItem As ListItem
    Dim i As Integer
    
    On Error Resume Next
    
    Set objItem = lvwCur.FindItem(mstrRepName, 0, intCurPosition, 1)
    If Not objItem Is Nothing Then
        Set lvwCur.SelectedItem = objItem
        lvwCur.ListItems(lvwCur.SelectedItem.Index).Selected = True
        lvwCur.SelectedItem.EnsureVisible
        If lvwCur.name = "lvwGroup" Then
            mstrPreGroup = ""
            Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
        End If
    End If
End Sub

Private Sub cmdNext_Click()
    Call txtFind_KeyPress(vbKeyReturn)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF And Shift = vbCtrlMask Then
        txtFind.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Static strPass As String
    Static vTime As Date
    Dim blnDo As Boolean
    If Me.ActiveControl Is txtFind Then
        strPass = ""
    Else
        blnDo = cboSys.ItemData(cboSys.ListIndex) > 0
        If blnDo Then blnDo = lvwGroup.SelectedItem.Key = "_-1"
        If blnDo Then blnDo = Not lvwReport.SelectedItem Is Nothing
        If blnDo Then blnDo = InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789", UCase(Chr(KeyAscii))) > 0
        If blnDo Then
            If DateDiff("s", vTime, Now) <= 2 Then
                strPass = strPass & Chr(KeyAscii)
            Else
                strPass = Chr(KeyAscii)
            End If
            KeyAscii = 0
            vTime = Now
            If UCase(strPass) = UCase("Publish Report") Then
                Call mnuEdit_Publish_Module_Click
                strPass = ""
            ElseIf UCase(strPass) = UCase("unPublish Report") Then
                Call mnuEdit_unPub_Module_Click
                strPass = ""
            End If
        Else
            strPass = ""
        End If
    End If
End Sub

Private Sub Form_Load()
    '读取自定义报表授权功能
    mblnModule = True
    mblnGrant = (zlRegTool() And 2) = 2
    If Not mblnGrant Then
        mnuEdit_Add.Visible = False
        mnuEdit_Del.Visible = False
        
        mnuEdit_Group_Add.Visible = False
        mnuEdit_Group_Delete.Visible = False
        
        mnuEdit_Guide.Visible = False
        mnuEdit_Guide_.Visible = False
        
        mnuEdit_Publish.Visible = False
        mnuEdit_unPub.Visible = False
                
        mnuEdit_Design_.Visible = False
        
        tbr.Buttons("Add").Visible = False
        tbr.Buttons("Del").Visible = False
        tbr.Buttons("GroupAdd").Visible = False
        tbr.Buttons("GroupDel").Visible = False
        tbr.Buttons("Guide").Visible = False
        tbr.Buttons("Guide_").Visible = False
        tbr.Buttons("Publish").Visible = False
        tbr.Buttons("unPub").Visible = False
        tbr.Buttons("Pub_").Visible = False
    End If
    
    lvwReport.ColumnHeaders(RC_编号 + 1).Position = 1
    mblnMouseDown = False
    RestoreWinState Me, App.ProductName
    tbrCheck.ZOrder
    mnuViewOnly.Checked = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name, "显示项", 1)

    Call ReadSystem
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '工具条占用高度s
    Dim staH As Long '状态栏占用高度
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    If cbr.Bands(2).MinWidth < 6615 Then cbr.Bands(2).MinWidth = 6615
    If Width < 8000 Then Width = 8000
    If Height < 5000 Then Height = 5000
    
    '靠齐控件宽度和高度
    cbrH = IIF(cbr.Visible, cbr.Height, 0)
    staH = IIF(sta.Visible, sta.Height, 0)
    With ImgSplit_S
        .Top = ScaleTop + cbrH
        .Height = ScaleHeight - cbrH - staH
    End With
    
    With lblGroup
        .Left = 0
        .Top = ImgSplit_S.Top + 30
        .Width = ImgSplit_S.Left
    End With
    
    With lvwGroup
        .Left = 0
        .Top = lblGroup.Top + lblGroup.Height + 30
        .Width = ImgSplit_S.Left
        .Height = ImgSplit_S.Top + ImgSplit_S.Height - .Top
    End With
    
    With LblReport
        .Left = ImgSplit_S.Left + ImgSplit_S.Width
        .Top = ImgSplit_S.Top + 30
        .Width = ScaleWidth - .Left
    End With
    
    With lvwReport
        .Left = ImgSplit_S.Left + ImgSplit_S.Width
        .Top = LblReport.Top + LblReport.Height + 30
        .Width = ScaleWidth - .Left
        .Height = ImgSplit_S.Top + ImgSplit_S.Height - .Top
    End With
End Sub

Private Sub Sub查看菜单(ByVal mnuLable As String)
    Dim i As Integer
    
    Select Case mnuLable
        Case "标准按钮(&B)"
            mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
            mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
            cbr.Visible = Not cbr.Visible
            Form_Resize
        Case "文本标签(&L)"
            mnuViewToolText.Checked = Not mnuViewToolText.Checked
            For i = 1 To tbr.Buttons.count
                If mnuViewToolText.Checked Then
                    tbr.Buttons(i).Caption = tbr.Buttons(i).Tag
                Else
                    tbr.Buttons(i).Caption = ""
                End If
            Next
            cbr.Bands(1).MinHeight = tbr.ButtonHeight
            Form_Resize
        Case "状态栏(&S)"
            mnuViewStatus.Checked = Not mnuViewStatus.Checked
            sta.Visible = Not sta.Visible
            Form_Resize
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub ImgSplit_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With ImgSplit_S
            If .Left + X < 1500 Or Me.ScaleWidth - .Left - X < 2000 Then Exit Sub
            .Move .Left + X
        End With
        Form_Resize
    End If
End Sub


Private Sub lvwGroup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuEdit, 2
End Sub

Private Sub lvwReport_DblClick()
    mcsActive = CS_报表
    lblFind.Caption = "查找报表(&F)"
    If mblnItem Then mnuEdit_Design_Click
End Sub

Private Sub lvwReport_GotFocus()
    mcsActive = CS_报表
    lblFind.Caption = "查找报表(&F)"
    Call SetFuncEnabled(True)
End Sub

Private Sub lvwReport_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mcsActive = CS_报表
    lblFind.Caption = "查找报表(&F)"
    Item.Selected = True '可以多选时如果不写这句话，SelectedItem 为Nothing
    Call SetFuncEnabled(True)
    
    If Item.SubItems(RC_发布时间) <> "" Then
        sta.Panels(2) = Item.Text & "位置:" & GetMenuPath(Val(Mid(Item.Key, 2)))
    Else
        sta.Panels(2) = "[" & Item.SubItems(RC_编号) & "]" & Item.Text & IIF(Item.SubItems(RC_说明) = "", "", ":" & Item.SubItems(RC_说明))
    End If
    
End Sub

Private Sub lvwReport_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    If KeyCode = vbKeyDelete Then
        mnuEdit_Del_Click
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To lvwReport.ListItems.count
            lvwReport.ListItems(i).Selected = True
        Next
    End If
End Sub

Private Sub lvwReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Not lvwReport.SelectedItem Is Nothing Then
        mnuEdit_Modi_Click
    End If
End Sub

Private Sub lvwReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnMouseDown = False
    If lvwReport.HitTest(X, Y) Is Nothing Then
        mblnItem = False
        If Button = 1 Then sta.Panels(2) = "共 " & lvwReport.ListItems.count & " 张报表"
    Else
        mblnItem = True
        mblnMouseDown = (Button = 1) And (cboSys.Text = "所有系统共享")
    End If
End Sub

Private Sub lvwReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnMouseDown = False
    Set lvwReport.DragIcon = Nothing
    lvwReport.Drag 0
    
    If Button = 2 Then
        If Not mblnItem Then
            PopupMenu mnuView, 2
        Else
            PopupMenu mnuEdit, 2
        End If
    End If
End Sub

Private Sub LvwGroup_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    mcsActive = CS_报表组
    lblFind.Caption = "查找分组(&F)"
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwGroup.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwGroup.SortOrder = lvwDescending
    Else
        lvwGroup.SortOrder = lvwAscending
    End If
    lvwGroup.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwGroup.SelectedItem Is Nothing Then lvwGroup.SelectedItem.EnsureVisible
End Sub

Private Sub LvwGroup_DblClick()
    mcsActive = CS_报表组
    lblFind.Caption = "查找分组(&F)"
    mnuEdit_Group_Modify_Click
End Sub

Private Sub LvwGroup_DragDrop(Source As Control, X As Single, Y As Single)
    Dim rsInsert As New ADODB.Recordset
    Dim rsGetGroups As New ADODB.Recordset
    Dim objLastSel As ListItem, objCurSel As ListItem
    Dim lngReportID As Long, intRptCount As Integer
    Dim blnInsert As Boolean, strSQL As String
    
    With lvwGroup
        If .SelectedItem.Key = "_-1" Then Exit Sub
                        
        strSQL = "Select 1 From zlRPTSubs A,zlReports B Where B.名称=[1] And A.报表ID=B.ID And A.组ID=[2]"
        Set rsGetGroups = OpenSQLRecord(strSQL, Me.Caption, Source.SelectedItem.Text, Val(Mid(lvwGroup.SelectedItem.Key, 2)))
        If Not rsGetGroups.EOF Then
            MsgBox "该报表组中已经包含相同名称的报表！", vbInformation, App.Title
            lvwGroup.ListItems("_-1").Selected = True: Exit Sub
        End If
        
        intRptCount = 1
        strSQL = "Select Count(*) Records From zlRPTSubs Where 组ID=[1]"
        Set rsGetGroups = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(lvwGroup.SelectedItem.Key, 2)))
        If Not rsGetGroups.EOF Then
            intRptCount = Nvl(rsGetGroups!Records, 0) + 1
        End If
    End With
    
    Set objLastSel = mobjSelItem
    Set objCurSel = lvwGroup.SelectedItem
    '删除当前组中拖动的子表,并插入到新组
    If objLastSel.Key <> "_-1" Then
        '修改序号
        gcnOracle.Execute "Update zlRPTSubs Set 序号=序号-1 Where 序号>(Select 序号 From zlRPTSubs Where 组ID=" & Mid(objLastSel.Key, 2) & " And 报表ID=" & Mid(Source.SelectedItem.Key, 2) & ") And 组ID=" & Mid(objLastSel.Key, 2)
        gcnOracle.Execute "Delete zlRPTSubs Where 组ID=" & Mid(objLastSel.Key, 2) & " And 报表ID=" & Mid(Source.SelectedItem.Key, 2)
    End If
    
    blnInsert = True
    strSQL = "Select Count(*) Records From zlRPTSubs Where 组ID=[1] And 报表ID=[2]"
    Set rsInsert = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(objCurSel.Key, 2)), Val(Mid(Source.SelectedItem.Key, 2)))
    If Not rsInsert.EOF Then
        blnInsert = Nvl(rsInsert!Records, 0) = 0
    End If

    If blnInsert Then gcnOracle.Execute "Insert Into zlRPTSubs(组ID,报表ID,序号,功能) Values(" & Mid(objCurSel.Key, 2) & "," & Mid(Source.SelectedItem.Key, 2) & "," & intRptCount & ",'" & Source.SelectedItem.Text & "')"
    If Not mobjSelItem Is Nothing Then Set lvwGroup.SelectedItem = mobjSelItem
    mstrPreGroup = ""
    Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
    
    '更新已发布报表组的权限
    If Val(objCurSel.Tag) <> 0 Then Call ReportGrantToNavigatorAgain(objCurSel)
    If objCurSel.Key <> objLastSel.Key Then
        If Val(objLastSel.Tag) <> 0 Then
            strSQL = "Select Count(*) Records from zlRPTSubs Where 组ID=[1]"
            Set rsGetGroups = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(objLastSel.Key, 2)))
            If Not rsGetGroups.EOF Then
                If rsGetGroups!Records = 0 Then
                    '取消该报表组的发布
                    Call ReportRevokeFromNavigator(True)
                Else
                    Call ReportGrantToNavigatorAgain(objLastSel)
                End If
            Else
                Call ReportRevokeFromNavigator(True)
            End If
        End If
    End If
End Sub

Private Sub LvwGroup_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Dim objTest As ListItem
    
    Set objTest = lvwGroup.HitTest(X, Y)
    If Not objTest Is Nothing Then
        If objTest.Key = "_-1" Then Exit Sub
        objTest.Selected = True
    End If
End Sub

Private Sub LvwGroup_GotFocus()
    Call SetFuncEnabled(False)
End Sub

Private Sub LvwGroup_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mcsActive = CS_报表组
    lblFind.Caption = "查找分组(&F)"
    Call SetFuncEnabled(False)
    
    Set mobjSelItem = Item
    
    If Item.Key <> mstrPreGroup Then
        If Not ReadReports(Mid(Item.Key, 2)) Then
            MsgBox "报表读取失败！", vbInformation, App.Title
            Exit Sub
        End If
        mstrPreGroup = Item.Key
    End If
    sta.Panels(2) = "共 " & lvwReport.ListItems.count & " 张报表"
    If Item.SubItems(GC_发布时间) <> "" Then
        sta.Panels(2) = sta.Panels(2) & "," & Item.Text & "位置:" & GetMenuPath(Val(Mid(Item.Key, 2)), True)
    End If
End Sub

Private Sub LvwGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then LvwGroup_DblClick
End Sub

Private Sub LvwGroup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strPath As String
    Static objItem As Object
    
    If Not objItem Is Nothing And Not lvwGroup.HitTest(X, Y) Is Nothing Then
        If objItem.Key = lvwGroup.HitTest(X, Y).Key Then Exit Sub
    End If
    
    Set objItem = lvwGroup.HitTest(X, Y)
    lvwGroup.ToolTipText = ""
End Sub

Private Sub mnuEdit_Add_Click()
    Dim strName As String, lngID As Long, str编码 As String, str说明 As String
    
    If cboSys.ItemData(cboSys.ListIndex) <> 0 Then cboSys.ListIndex = 0
    If frmReportEdit.ShowMe(Me, cboSys.ItemData(cboSys.ListIndex), False, Val(lvwGroup.SelectedItem.Tag), IIF(lvwGroup.SelectedItem.Key = "_-1", 0, Mid(lvwGroup.SelectedItem.Key, 2)), _
                                                lngID, strName, str编码, str说明) Then
        mstrPreGroup = ""
        Call AfterItemEdit(True, False, lngID, strName, str编码, str说明)
        Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
    Else
        Call CustomToolBarRefresh
    End If
End Sub

Private Sub mnuEdit_Clear_Click()
    frmClearHistory.Show 1, Me
End Sub

Private Sub mnuEdit_Del_Click()
    Dim rsCheck As New ADODB.Recordset
    Dim rsGetGroups As New ADODB.Recordset
    Dim intIdx As Integer, strSQL As String
    
    If lvwReport.SelectedItem Is Nothing Then
        MsgBox "当前没有报表可以删除！", vbInformation, App.Title: Exit Sub
    End If
    If cboSys.ItemData(cboSys.ListIndex) > 0 Then Exit Sub
    
    intIdx = lvwReport.SelectedItem.Index
    If lvwGroup.SelectedItem.Key = "_-1" Then
        '检查是否属于报表组，是否不允许删除
        strSQL = "Select ID 组ID,名称 From zlRPTGroups Where ID=(Select 组ID From zlRPTSubs Where 报表ID=[1])"
        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(lvwReport.SelectedItem.Key, 2)))
        If Not rsCheck.EOF Then
            If Not IsNull(rsCheck!组ID) Then
                MsgBox "请先把报表[" & lvwReport.SelectedItem & "]从报表组[" & rsCheck!名称 & "]中移除后再删除！", vbInformation, App.Title
                Exit Sub
            End If
        End If
        '检查是否已发布
        If Val(lvwReport.SelectedItem.Tag) <> 0 Then
            MsgBox "该报表已经发布,请先取消发布后再删除！", vbInformation, App.Title: Exit Sub
        End If
        strSQL = "Select 报表ID From zlRPTPuts Where 报表ID=[1]"
        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(lvwReport.SelectedItem.Key, 2)))
        If Not rsCheck.EOF Then
            MsgBox "该报表已经发布,请先取消发布后再删除！", vbInformation, App.Title: Exit Sub
        End If
        
        If MsgBox("确实要删除报表[" & lvwReport.SelectedItem & "]吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        On Error GoTo errH
        gcnOracle.BeginTrans
        gcnOracle.Execute "Delete From zlReports Where ID=" & Mid(lvwReport.SelectedItem.Key, 2)
        gcnOracle.CommitTrans
        On Error GoTo 0
    Else
        If MsgBox("你确定要从报表组[" & lvwGroup.SelectedItem & "]中移除报表[" & lvwReport.SelectedItem & "]吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        On Error GoTo errH
        gcnOracle.BeginTrans
        gcnOracle.Execute "Update zlRPTSubs Set 序号=序号-1 Where 序号>(Select 序号 From zlRPTSubs Where 报表ID=" & Mid(lvwReport.SelectedItem.Key, 2) & " And 组ID=" & Mid(lvwGroup.SelectedItem.Key, 2) & ") And 组ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
        gcnOracle.Execute "Delete From zlRPTSubs Where 报表ID=" & Mid(lvwReport.SelectedItem.Key, 2) & " And 组ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
        gcnOracle.CommitTrans
        On Error GoTo 0
    End If
    
    lvwReport.ListItems.Remove intIdx
    
    If lvwReport.ListItems.count <> 0 Then
        If intIdx <= lvwReport.ListItems.count Then
            lvwReport.ListItems(intIdx).Selected = True
        Else
            lvwReport.ListItems(lvwReport.ListItems.count).Selected = True
        End If
        lvwReport.SelectedItem.EnsureVisible
        Call lvwReport_ItemClick(lvwReport.SelectedItem)
    Else
        sta.Panels(2) = "共 0 张报表"
    End If
    
    '更新已发布报表组的权限
    If Val(lvwGroup.SelectedItem.Tag) <> 0 Then
        strSQL = "Select Count(*) Records from zlRPTSubs Where 组ID=[1]"
        Set rsGetGroups = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(lvwGroup.SelectedItem.Key, 2)))
        If Not rsGetGroups.EOF Then
            If rsGetGroups!Records = 0 Then
                '取消该报表组的发布
                Call ReportRevokeFromNavigator(True)
            Else
                Call ReportGrantToNavigatorAgain(lvwGroup.SelectedItem)
            End If
        Else
            Call ReportRevokeFromNavigator(True)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Design_Click()
    Dim lngIndex As Long, i As Long
    
    If lvwReport.SelectedItem Is Nothing Then
        MsgBox "当前没有报表可以设计！", vbInformation, App.Title: Exit Sub
    End If
    If CheckPass(CLng(Mid(lvwReport.SelectedItem.Key, 2))) = False Then
        MsgBox "报表数据错误，不能设计该报表！", vbInformation, App.Title: Exit Sub
    End If
    If Not CheckReportPriv(CLng(Mid(lvwReport.SelectedItem.Key, 2))) Then
        MsgBox "你没有权限查询该报表某些数据源中的对象,请在设计环境下修正！", vbInformation, App.Title
    End If
    
    glngSys = cboSys.ItemData(cboSys.ListIndex)
    frmDesign.lngRPTID = CLng(Mid(lvwReport.SelectedItem.Key, 2))
    
    On Error Resume Next
    frmDesign.Show 1, Me
    On Error GoTo 0
    
    If gblnModi Then
        '调整选项和选项的内容
        lngIndex = lvwReport.SelectedItem.Index
        Call ReadGroups
        If lngIndex > lvwReport.ListItems.count Then
            lngIndex = lvwReport.ListItems.count
        End If
        For i = 1 To lvwReport.ListItems.count
            If lvwReport.ListItems(i).Selected Then
                lvwReport.ListItems(i).Selected = False
            End If
        Next
        lvwReport.ListItems(lngIndex).Selected = True
    End If
End Sub

Private Sub mnuEdit_Group_Add_Click()
    Dim strName As String, lngID As Long, str编码 As String, str说明 As String
    If frmReportEdit.ShowMe(Me, cboSys.ItemData(cboSys.ListIndex), True, 0, lngID, , strName, str编码, str说明) Then
        Call AfterItemEdit(True, True, lngID, strName, str编码, str说明)
        Call mnuView_reFlash_Click
    End If
End Sub

Private Sub mnuEdit_Group_Delete_Click()
    If lvwGroup.SelectedItem Is Nothing Then
        MsgBox "当前没有报表组可以删除！", vbInformation, App.Title: Exit Sub
    End If
    If lvwGroup.SelectedItem.Key = "_-1" Then
        MsgBox "当前没有报表组可以删除！", vbInformation, App.Title: Exit Sub
    End If
    If lvwGroup.SelectedItem.Icon = 3 Then
        MsgBox "系统固有的报表组不能删除！", vbInformation, App.Title: Exit Sub
    End If
    If Val(lvwGroup.SelectedItem.Tag) <> 0 Then
        MsgBox "请取消该报表组的发布后，再试！", vbInformation, App.Title: Exit Sub
    End If
    If MsgBox("你确定要删除报表组[" & lvwGroup.SelectedItem & "]吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
    
    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    gcnOracle.Execute "Delete zlRPTSubs Where 组ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
    gcnOracle.Execute "Delete zlRPTGroups Where ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
    gcnOracle.CommitTrans
    
    mnuView_reFlash_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub mnuEdit_Group_Modify_Click()
    Dim lngSys As Long
    Dim strName As String, lngID As Long, str编码 As String, str说明 As String
    
    If lvwGroup.SelectedItem Is Nothing Then
        MsgBox "当前没有报表组可以修改！", vbInformation, App.Title: Exit Sub
    End If
    If lvwGroup.SelectedItem.Key = "_-1" Then Exit Sub
    lngSys = cboSys.ItemData(cboSys.ListIndex)
    lngID = CLng(Mid(lvwGroup.SelectedItem.Key, 2))
    str编码 = lvwGroup.SelectedItem.SubItems(GC_编号)
    strName = lvwGroup.SelectedItem.Text
    str说明 = lvwGroup.SelectedItem.SubItems(GC_说明)
    If frmReportEdit.ShowMe(Me, lngSys, True, Val(lvwGroup.SelectedItem.Tag), lngID, , strName, str编码, str说明) Then
        Call AfterItemEdit(False, True, lngID, strName, str编码, str说明)
        mnuView_reFlash_Click
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
    Unload frmReportEdit
End Sub

Private Sub mnuEdit_Group_Publish_Click()
    Call ReportGrantToNavigator
End Sub

Private Sub mnuEdit_Group_Setup_Click()
    Dim rsGetGroups As New ADODB.Recordset
    Dim strSQL As String
    
    '设置哪些报表属于该报表组
    If lvwGroup.SelectedItem Is Nothing Then Exit Sub
    If lvwGroup.SelectedItem.Key = "_-1" Then Exit Sub
    
    With frmSetGroup
        .LngGroupID = Mid(lvwGroup.SelectedItem.Key, 2)
        .strCaption = "设置报表组[" & lvwGroup.SelectedItem & "]的从属报表"
        .Show 1, Me
    End With
    
    '更新已发布报表组的权限
    If Val(lvwGroup.SelectedItem.Tag) <> 0 Then
        strSQL = "Select Count(*) Records From zlRPTSubs Where 组ID=[1]"
        Set rsGetGroups = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(lvwGroup.SelectedItem.Key, 2)))
        If Not rsGetGroups.EOF Then
            If rsGetGroups!Records = 0 Then
                '取消该报表组的发布
                Call ReportRevokeFromNavigator(True)
            Else
                Call ReportGrantToNavigatorAgain(lvwGroup.SelectedItem)
            End If
        Else
            Call ReportRevokeFromNavigator(True)
        End If
    End If
    
    mstrPreGroup = ""
    Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
End Sub

Private Sub mnuEdit_Group_unPub_Click()
    Call ReportRevokeFromNavigator
End Sub

Private Sub mnuEdit_Modi_Click()
    Dim lngSys As Long, lngRPTID As Long
    Dim strName As String, lngID As Long, str编码 As String, str说明 As String
    
    If lvwReport.SelectedItem Is Nothing Then
        MsgBox "当前没有报表可以修改！", vbInformation, App.Title: Exit Sub
    End If
    lngSys = cboSys.ItemData(cboSys.ListIndex)
    lngRPTID = CLng(Mid(lvwReport.SelectedItem.Key, 2))
    lngID = IIF(lvwGroup.SelectedItem.Key = "_-1", 0, Mid(lvwGroup.SelectedItem.Key, 2))
    str编码 = lvwReport.SelectedItem.SubItems(RC_编号)
    strName = lvwReport.SelectedItem.Text
    str说明 = lvwReport.SelectedItem.SubItems(RC_说明)
    If frmReportEdit.ShowMe(Me, lngSys, False, Val(lvwReport.SelectedItem.Tag), lngID, _
                                                lngRPTID, strName, str编码, str说明) Then
        Call AfterItemEdit(False, False, lngRPTID, strName, str编码, str说明)
        mstrPreGroup = ""
        Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
    Else
        Call CustomToolBarRefresh
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
    Unload frmReportEdit
End Sub

Private Sub mnuEdit_Publish_Module_Click()
    Call ReportGrantToModule
End Sub

Private Sub mnuEdit_unPub_Module_Click()
    Call ReportRevokeFromModule
End Sub

Private Sub mnuFile_ExpAll_Click()
    Dim rsReportInfo As New ADODB.Recordset, strSQL As String
    Dim strPath As String, strFile As String, strPathTmp As String
    Dim i As Long, j As Long, lngCount As Long, lngExp As Long
    Dim objFile As New FileSystemObject
    
    strPath = BrowseForFolder(Me.hwnd, "选择报表导出目录", strPath)
    If strPath <> "" Then
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Path", "Export", strPath
        strSQL = "Select A.Id, A.编号, A.名称, C.Id 组id, C.编号 组编号, C.名称 组名" & vbNewLine & _
                    "From zlReports A, zlRPTSubs B, zlRPTGroups C" & vbNewLine & _
                    "Where A.Id = B.报表id(+) And B.组id = C.Id(+)  And " & IIF(cboSys.ItemData(cboSys.ListIndex) = 0, " A.系统 Is Null ", " A.系统=[1] ") & vbNewLine & _
                    "Order By A.编号"
        Set rsReportInfo = OpenSQLRecord(strSQL, Me.Caption, cboSys.ItemData(cboSys.ListIndex))
        lngCount = rsReportInfo.RecordCount
        If MsgBox("本次共导出 " & cboSys.List(cboSys.ListIndex) & lngCount & " 张报表到 " & strPath & "，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        lngExp = 0
        For i = 1 To lvwGroup.ListItems.count
            '所有报表
            If Val(Mid(lvwGroup.ListItems(i).Key, 2)) = -1 Then
                rsReportInfo.Filter = "组id=Null"
                strPathTmp = strPath
            Else
                rsReportInfo.Filter = "组id=" & Val(Mid(lvwGroup.ListItems(i).Key, 2))
                strPathTmp = strPath & "\[" & lvwGroup.ListItems(i).SubItems(GC_编号) & "]" & lvwGroup.ListItems(i).Text
                If Not objFile.FolderExists(strPathTmp) Then
                    Call objFile.CreateFolder(strPathTmp)
                End If
            End If
            For j = 1 To rsReportInfo.RecordCount
                lngExp = lngExp + 1
                Call ShowFlash("正在导出:" & rsReportInfo!名称 & ".ZLR", lngExp / lngCount, Me, True)
                strFile = "[" & rsReportInfo!编号 & "]" & rsReportInfo!名称 & ".ZLR"
                If Not ExportReport(Val(rsReportInfo!ID & ""), strPathTmp & "\" & strFile) Then
                    Call ShowFlash
                    If MsgBox("导出报表时出现错误，要继续导出下一张报表吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
                End If
                rsReportInfo.MoveNext
            Next
        Next
        Call ShowFlash
    End If
End Sub

Private Sub mnuFile_ImpAll_Click()
    Dim strPath As String, objFSO As New FileSystemObject, objFile As File, objFolder As Folder
    Dim lngSys As Long, lngCurGroup As Long
    Dim rsFiles As ADODB.Recordset
    Dim arrTmp As Variant, strFile As String, i As Long
    Dim rsGroups As ADODB.Recordset, strName As String, strCode As String, strSQL As String
    Dim LngGroupID As Long, lngReportID As Long
    
    On Error GoTo errH
    strPath = BrowseForFolder(Me.hwnd, "选择需要导入报表所在目录", strPath)
    If strPath <> "" Then
        If MsgBox("是否导入""" & strPath & """文件夹及子文件夹下的所有报表？", vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
        lngSys = cboSys.ItemData(cboSys.ListIndex)
        lngCurGroup = IIF(lvwGroup.SelectedItem.Key = "_-1", 0, Val(Mid(lvwGroup.SelectedItem.Key, 2)))
        If Not lvwReport.SelectedItem Is Nothing Then lngReportID = CLng(Mid(lvwReport.SelectedItem.Key, 2))
        'FilePath=报表全路径；FileName=报表文件名；组ID=报表要导入的报表组ID
        '同名ID=与将要导入的报表同名的报表的报表ID，固定报表通过编码匹配，非固定通过名称匹配
        '导入类型=0-不导入，1-新增导入,2-覆盖导入;覆盖类型=0-整体覆盖，1-仅数据源覆盖
        'ErrType=0-无错误,1-多个相同报表一起新增，2-多个相同报表一起覆盖，3-系统报表只能覆盖，但是无同名报表。
        '                            4-内容存在问题,5-版本存在问题,6-名称编号存在问题
        'ImportResult=-1-已经成功导入但是报表对象检查未通过，0-不导入,1-导入成功,2-导入失败
        'ImportInfo=报表成功导入后返回的报表信息
        Set rsFiles = CopyNewRec(Nothing, , True, _
                                    Array("FilePath", adVarChar, 1000, Empty, "FileName", adVarChar, 200, Empty, "组ID", adBigInt, Empty, Empty, _
                                             "同名ID", adBigInt, Empty, Empty, "导入类型", adInteger, Empty, Empty, "覆盖类型", adInteger, Empty, Empty, _
                                             "ErrType", adInteger, Empty, Empty, "ImportResult", adInteger, Empty, Empty, "ImportInfo", adVarChar, 200, Empty))
        
        
        With rsFiles
            '搜集导入到所有报表中的的报表,即当前文件夹下的报表
            For Each objFile In objFSO.GetFolder(strPath).Files
                If UCase(objFile.name) Like "*.ZLR" Then
                    rsFiles.AddNew Array("FilePath", "FileName", "组ID", "同名ID", "导入类型", "覆盖类型", "ErrType", "ImportResult", "ImportInfo"), _
                                            Array(objFile.Path, objFile.name, 0, 0, 0, 0, 0, 0, "")
                End If
            Next
            '仅需要查找自定义报表的分组
            '固定报表由于编码唯一性，已经确定分组
            If lngSys = 0 Then
                strSQL = "Select ID,编号,名称 From zlRPTGroups  Where  系统 Is Null"
                Set rsGroups = CopyNewRec(OpenSQLRecord(strSQL, Me.Caption))
            End If
            '搜集当前文件下的子级文件夹
            For Each objFolder In objFSO.GetFolder(strPath).SubFolders
                strFile = ""
                For Each objFile In objFolder.Files
                    If UCase(objFile.name) Like "*.ZLR" Then
                        strFile = strFile & "|" & objFile.name
                    End If
                Next
                If strFile <> "" Then
                    arrTmp = Split(Mid(strFile, 2), "|")
                    LngGroupID = 0
                    '仅自定报表需要查找分组，固定报表会有系统号编码确定分组
                    If lngSys = 0 Then
                        Call SplitNameCode(objFolder.name, strName, strCode)
                        rsGroups.Filter = "编号='" & strCode & "'" '编号唯一性
                        If rsGroups.EOF Then rsGroups.Filter = "名称='" & strName & "'"  '可能子分类没有编码
                        If Not rsGroups.EOF Then
                            LngGroupID = Val(rsGroups!ID & "")
                        Else '生成经常性的报表组
                            '将编码名称规范化，并生成新的编码名称
                            LngGroupID = GetNextID("zlRPTGroups")
                            If TLen(strName) > 30 Then strName = ConvertSBC(MidB(strName, 1, 30))
                            If strCode <> "" Then
                                If TLen(strCode) > 20 Then strCode = ConvertSBC(MidB(strCode, 1, 20))
                                If CheckExist("zlRPTGroups", "编号", strCode) Then
                                    strCode = GetNextNO(True)
                                End If
                            Else
                                strCode = GetNextNO(True)
                            End If
                            strSQL = "Insert Into zlRPTGroups(ID,编号,名称,说明) Values(" & LngGroupID & ",'" & strCode & "','" & strName & "',Null)"
                            On Error Resume Next
                            gcnOracle.Execute strSQL
                            If Err.Number <> 0 Then
                                LngGroupID = 0 '生成报表组失败，则自动将该分组下的报表导入到所遇分类
                            Else '生成分组成功，加入到组信息缓存中
                                rsGroups.AddNew Array("ID", "编号", "名称"), Array(LngGroupID, strCode, strName)
                            End If
                            On Error GoTo errH
                        End If
                    End If
                    For i = LBound(arrTmp) To UBound(arrTmp)
                        rsFiles.AddNew Array("FilePath", "FileName", "组ID", "同名ID", "导入类型", "覆盖类型", "ErrType", "ImportResult", "ImportInfo"), _
                                                Array(objFolder.Path & "\" & arrTmp(i), arrTmp(i), LngGroupID, 0, 0, 0, 0, 0, "")
                    
                    Next
                End If
            Next
            .Filter = "": .Sort = "组ID"
            If .RecordCount = 0 Then
                MsgBox "当前路径下未找到任何可导入的报表", vbInformation, App.Title
                Exit Sub
            End If
            Call ImportReportBeach(lngSys, lngCurGroup, lngReportID, rsFiles, True)
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SplitNameCode(ByVal strInput As String, ByRef strName As String, ByRef strCode As String)
'功能:分割编码名称
'参数：strInput=输入的字符串，如果格式为[编码]名称,则自动分割，否则默认为只获取到名称
'返回：strName=名称
'           strCode=编码
    Dim arrTmp As Variant
    Dim strTmp As Variant
    If InStr(strInput, "\") > 0 Then
        strTmp = strReverse(strInput)
        strInput = strReverse(Mid(strTmp, 1, InStr(strTmp, "\") - 1))
    End If
    
    If strInput Like "[[]?*[]]?*" Then '符合规范的文件名
        arrTmp = Split(strInput, "]")
        strName = arrTmp(1)
        strCode = Mid(arrTmp(0), 2)
    Else
        strName = strInput
        strCode = ""
    End If
End Sub

Private Sub mnuFile_Para_Click()
    '打开参数设置
    If mfrmReportPara Is Nothing Then
        Set mfrmReportPara = New frmReportPara
    End If
    If mfrmReportPara.ShowMe(Me) Then
        '更新参数
        Call InitPar
    End If
End Sub

Private Sub mnuFile_Report_Click()
    Dim objCheck As ListItem
    
    If Me.ActiveControl.name = "lvwReport" Or lvwGroup.SelectedItem.Key = "_-1" Then
        If lvwReport.SelectedItem Is Nothing Then MsgBox "当前没有可执行的报表！", vbInformation, App.Title: Exit Sub
        If Not CheckReportPriv(CLng(Mid(lvwReport.SelectedItem.Key, 2))) Then
            MsgBox "你没有权限查询该报表某些数据源中的对象！", vbInformation, App.Title: Exit Sub
        End If
    Else
        For Each objCheck In lvwReport.ListItems
            If Not CheckReportPriv(CLng(Mid(objCheck.Key, 2))) Then
                MsgBox "你没有权限查询报表[" & objCheck.Text & "]中某些数据源的对象！", vbInformation, App.Title: Exit Sub
            End If
        Next
    End If
    
    If Not (Me.ActiveControl.name = "lvwReport" Or lvwGroup.SelectedItem.Key = "_-1") Then
        '执行报表组
        Set gobjReport = Nothing
        glngGroup = CLng(Mid(lvwGroup.SelectedItem.Key, 2))
    Else
        '执行报表
        If CheckPass(CLng(Mid(lvwReport.SelectedItem.Key, 2))) = False Then
            MsgBox "报表数据错误，不能执行该报表！", vbInformation, App.Title: Exit Sub
        End If
        
        glngGroup = 0
        Set gobjReport = Nothing
        Set gobjReport = ReadReport(CLng(Mid(lvwReport.SelectedItem.Key, 2)))
    End If
    
    glngSys = cboSys.ItemData(cboSys.ListIndex)
    garrPars = Array() '使用缺省参数
    If Not ShowReport(Me) Then MsgBox "报表打开失败！", vbInformation, App.Title
End Sub

Private Sub mnuFile_Quit_Click()
    Unload Me
End Sub

Private Sub mnuFindNext_Click()
    Call txtFind_KeyPress(vbKeyReturn)
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me)
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelpRpt(Me.hwnd, "main", 0)
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Private Sub mnuPopPublish_Group_Click()
    Call mnuEdit_Group_Publish_Click
End Sub

Private Sub mnuPopPublish_ReportMain_Click()
    Call mnuEdit_Publish_Main_Click
End Sub

Private Sub mnuPopPublish_ReportModule_Click()
    Call mnuEdit_Publish_Module_Click
End Sub

Private Sub mnuPopUnpub_Group_Click()
    Call mnuEdit_Group_unPub_Click
End Sub

Private Sub mnuPopUnpub_ReportMain_Click()
    Call mnuEdit_unPub_Main_Click
End Sub

Private Sub mnuPopUnpub_ReportModule_Click()
    Call mnuEdit_unPub_Module_Click
End Sub

Private Sub mnuView_reFlash_Click()
    Call ReadGroups
End Sub

Private Sub mnuViewOnly_Click()
    mnuViewOnly.Checked = mnuViewOnly.Checked Xor True
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name, "显示项", mnuViewOnly.Checked
    If lvwGroup.SelectedItem.Key = "_-1" Then
        mstrPreGroup = ""
        Call LvwGroup_ItemClick(lvwGroup.ListItems("_-1"))
    End If
End Sub

Private Sub mnuViewStatus_Click()
    Sub查看菜单 mnuViewStatus.Caption
End Sub

Private Sub mnuViewToolButton_Click()
    Sub查看菜单 mnuViewToolButton.Caption
End Sub

Private Sub mnuViewToolText_Click()
    Sub查看菜单 mnuViewToolText.Caption
End Sub

Private Sub picSysFind_Resize()
    txtFind.Top = (picSysFind.Height - txtFind.Height) / 2
    cboSys.Top = txtFind.Top
    lblSys.Top = (picSysFind.Height - lblSys.Height) / 2
    lblFind.Top = lblSys.Top
End Sub

Private Sub RunLog_Click(Index As Integer)
    If Not lvwReport.SelectedItem Is Nothing Then
        Call ShowRunLog
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_Quit_Click
        Case "View"
            If Me.ActiveControl.name = "lvwGroup" Then
                Call SetView((lvwGroup.View + 1) Mod 4)
            Else
                Call SetView((lvwReport.View + 1) Mod 4)
            End If
        Case "Add"
            mnuEdit_Add_Click
        Case "Modi"
            mnuEdit_Modi_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "GroupAdd"
            mnuEdit_Group_Add_Click
        Case "GroupModify"
            mnuEdit_Group_Modify_Click
        Case "GroupDel"
            mnuEdit_Group_Delete_Click
        Case "Design"
            mnuEdit_Design_Click
        Case "Report"
            mnuFile_Report_Click
        Case "Publish"
            PopupButtonMenu tbr, Button, mnuPopPublish
        Case "unPub"
            PopupButtonMenu tbr, Button, mnuPopUnpub
        Case "Guide"
            mnuEdit_Guide_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbr_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Publish"
            PopupButtonMenu tbr, Button, mnuPopPublish
        Case "unPub"
            PopupButtonMenu tbr, Button, mnuPopUnpub
    End Select
End Sub

Private Sub mnuView_View_Click(Index As Integer)
    Call SetView(CByte(Index))
End Sub

Private Sub SetView(bytStyle As Byte)
'功能：调整床位列表显示方式
'参数：bytstyle=0-大图标,1-小图标,2-列表,3-详细资料
    mnuView_View(0).Checked = False
    mnuView_View(1).Checked = False
    mnuView_View(2).Checked = False
    mnuView_View(3).Checked = False
    mnuView_View(bytStyle).Checked = True
    
    On Error Resume Next
    If Me.ActiveControl.name = "lvwGroup" Then
        lvwGroup.View = bytStyle
    Else
        lvwReport.View = bytStyle
    End If
End Sub

Private Sub ShowRunLog()
    Dim lngReportKey As Long
    Dim strReportName As String
    lngReportKey = Val(Mid(lvwReport.SelectedItem.Key, 2))
    '查看报表运行日志记录
    If lngReportKey > 0 Then
        Call frmReportRunLog.ShowMe(Me, lngReportKey, "报表[" & lvwReport.SelectedItem.Text & "]的运行日志")
    End If
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
        Case "RunLog"
            If Not lvwReport.SelectedItem Is Nothing Then
                Call ShowRunLog
            End If
    End Select
End Sub

Private Function ReadReports(ByVal lngKey As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As Object, strKey As String
    Dim strSQL As String, i As Integer
    
    On Error GoTo errH
    
    If Not lvwReport.SelectedItem Is Nothing Then
        strKey = lvwReport.SelectedItem.Key
    End If
    lvwReport.ListItems.Clear
    
    LockWindowUpdate lvwReport.hwnd
    
    If lngKey = -1 Then     '所有报表
        If mnuViewOnly.Checked Then '仅显示独立项
            strSQL = _
                "Select Distinct A.ID,A.编号,A.名称,A.说明,A.程序ID,A.修改时间,A.发布时间,A.系统,Nvl(A.票据,0) 票据,A.最后执行时间, " & vbCr & _
                "     A.执行人员 最后执行人, zlSpellCode(A.名称) 简码, b.其他数据连接 " & vbCr & _
                "From zlReports A, " & vbCr & _
                "     (Select B1.报表id, f_list2str(Cast(Collect(B2.名称) As t_Strlist)) 其他数据连接 " & vbCr & _
                "      From zlRPTDatas B1, zlConnections B2 " & vbCr & _
                "      Where b1.数据连接编号 = b2.编号 And Not Exists(Select 1 From zlRPTSubs Where 报表id = b1.报表id) " & vbCr & _
                "      Group By b1.报表id) B " & vbCr & _
                "Where a.id = b.报表id(+) " & vbCr & _
                IIF(cboSys.ItemData(cboSys.ListIndex) = 0, " and A.系统 Is Null ", " and A.系统=[1] ") & vbCr & _
                "     And Not Exists(Select 1 From zlRPTSubs Where 报表id = a.Id) " & vbCr & _
                "Order by A.编号"
        Else
            strSQL = _
                "Select Distinct A.ID,A.编号,A.名称,A.说明,A.程序ID,A.修改时间,A.发布时间,A.系统,Nvl(A.票据,0) 票据,A.最后执行时间, " & vbCr & _
                "     A.执行人员 最后执行人, zlSpellCode(A.名称) 简码, b.其他数据连接  " & vbCr & _
                "From zlReports A, " & vbCr & _
                "     (Select B1.报表id, f_list2str(Cast(Collect(B2.名称) As t_Strlist)) 其他数据连接 " & vbCr & _
                "      From zlRPTDatas B1, zlConnections B2 " & vbCr & _
                "      Where b1.数据连接编号 = b2.编号 " & vbCr & _
                "      Group By b1.报表id) B " & vbCr & _
                "Where a.id = b.报表id(+) " & vbCr & _
                IIF(cboSys.ItemData(cboSys.ListIndex) = 0, " and A.系统 Is Null ", " and A.系统=[1] ") & vbCr & _
                "Order by A.编号"
        End If
    Else
        strSQL = _
            "Select Distinct A.ID,A.编号,A.名称,A.说明,A.程序ID,A.修改时间,A.发布时间,A.系统,Nvl(A.票据,0) 票据,A.最后执行时间, " & vbCr & _
            "     A.执行人员 最后执行人, zlSpellCode(A.名称) 简码, b.其他数据连接 " & vbCr & _
            "From zlReports A, " & vbCr & _
            "     (Select B1.报表id, f_list2str(Cast(Collect(B2.名称) As t_Strlist)) 其他数据连接 " & vbCr & _
            "      From zlRPTDatas B1, zlConnections B2 " & vbCr & _
            "      Where b1.数据连接编号 = b2.编号 And Exists(Select 1 From zlRPTSubs Where 报表id = b1.报表id And 组ID=[2]) " & vbCr & _
            "      Group By b1.报表id) B, zlRPTSubs C " & vbCr & _
            "Where a.id = b.报表id(+) And a.Id = c.报表Id And c.组ID=[2] " & vbCr & _
            "Order by A.编号"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, cboSys.ItemData(cboSys.ListIndex), lngKey)
    For i = 1 To rsTmp.RecordCount
        If Not IsNull(rsTmp!系统) Then '固定安装报表
            If IsNull(rsTmp!发布时间) Then
                Set objItem = lvwReport.ListItems.Add(, "_" & rsTmp!ID, rsTmp!名称, "Fixed", "Fixed")
            Else
                Set objItem = lvwReport.ListItems.Add(, "_" & rsTmp!ID, rsTmp!名称, "PubFixed", "PubFixed")
            End If
            objItem.Tag = Val(Nvl(rsTmp!程序ID, 0))
        Else
            If Not IsNull(rsTmp!发布时间) Then '已发布
                If Nvl(rsTmp!票据, 0) = 1 Then
                    Set objItem = lvwReport.ListItems.Add(, "_" & rsTmp!ID, rsTmp!名称, "BillPublish", "BillPublish")
                Else
                    Set objItem = lvwReport.ListItems.Add(, "_" & rsTmp!ID, rsTmp!名称, "Publish", "Publish")
                End If
            Else
                If Nvl(rsTmp!票据, 0) = 1 Then
                    Set objItem = lvwReport.ListItems.Add(, "_" & rsTmp!ID, rsTmp!名称, "Bill", "Bill")
                Else
                    Set objItem = lvwReport.ListItems.Add(, "_" & rsTmp!ID, rsTmp!名称, "Report", "Report")
                End If
            End If
            objItem.Tag = Val(Nvl(rsTmp!程序ID, 0))
        End If
        objItem.SubItems(RC_编号) = rsTmp!编号
        objItem.SubItems(RC_说明) = Nvl(rsTmp!说明)
        objItem.SubItems(RC_修改时间) = Format(rsTmp!修改时间, "yyyy-MM-dd")
        objItem.SubItems(RC_发布时间) = Format(Nvl(rsTmp!发布时间), "yyyy-MM-dd")
        objItem.SubItems(RC_最后执行时间) = Format(Nvl(rsTmp!最后执行时间), "yyyy-MM-dd hh:mm")
        objItem.SubItems(RC_最后执行人) = Nvl(rsTmp!最后执行人)
        objItem.SubItems(RC_种类) = IIF(Nvl(rsTmp!票据, 0) = 1, "票据", "报表")
        objItem.SubItems(RC_类型) = IIF(IsNull(rsTmp!系统), "自制", "系统")
        objItem.SubItems(RC_简码) = rsTmp!简码 & ""
        objItem.SubItems(RC_其他数据连接) = mdlPublic.Nvl(rsTmp!其他数据连接)
        If objItem.Key = strKey Then objItem.Selected = True
        rsTmp.MoveNext
    Next
    
    If Not lvwReport.SelectedItem Is Nothing Then
        lvwReport.SelectedItem.EnsureVisible
    End If
    
    'If rsTmp.RecordCount > 0 Then Call AutoSizeCol(lvw)
    LockWindowUpdate 0
    
    ReadReports = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub lvwReport_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    mcsActive = CS_报表
    lblFind.Caption = "查找报表(&F)"
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwReport.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwReport.SortOrder = lvwDescending
    Else
        lvwReport.SortOrder = lvwAscending
    End If
    lvwReport.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwReport.SelectedItem Is Nothing Then lvwReport.SelectedItem.EnsureVisible
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuView, 2
End Sub

Private Function GetNewProgID() As Long
'功能：获取下一个可用的自定义报表程序号,用于发布
'说明：程序号从100000开始,并自动补缺
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Decode(Sign(Max(序号)-99999),1,Max(序号),99999) as ID From zlPrograms"
    Call OpenRecord(rsTmp, strSQL, Me.Caption)
    GetNewProgID = IIF(IsNull(rsTmp!ID), 100000, rsTmp!ID + 1)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetMainTreeMenu(Optional ByVal lngProgID As Long) As ADODB.Recordset
'功能：获取发布到导航台报表树形菜单体系
'参数：lngProgID=是否只显示指定程序ID的报表
'说明：菜单体系中包含自定义报表发布的菜单项(如果有),标志为"FLAG=999"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngSys As Long
    
    On Error GoTo errH
    
    lngSys = cboSys.ItemData(cboSys.ListIndex)
    If lngSys = 0 Then
        '只显示用户发布部份报表
        strSQL = _
            "Select Distinct * From (" & _
            " Select 编号 as SCOL,0 as Flag,-编号 as ID,-NULL as 上级ID,'['||编号||']'||名称 as 标题,-NULL as 模块 From zlSystems Union ALL" & _
            " Select 99999 as SCOL,Level as FLAG,ID,Nvl(上级ID,-系统) as 上级ID,标题,模块 From zlMenus Where 组别='缺省' And 模块 is NULL" & _
            " Start With 上级ID is NULL And 组别='缺省' Connect by Prior ID=上级ID And 组别='缺省'" & _
            " Union ALL" & _
            " Select 99999 as SCOL,999 as FLAG,A.ID,A.上级ID,A.标题,A.模块" & _
            " From zlMenus A,zlPrograms B,zlRPTGroups C" & _
            " Where A.模块=B.序号 And A.组别='缺省' And C.程序ID=A.模块 " & _
            " And Upper(B.部件)='ZL9REPORT'" & IIF(lngProgID = 0, "", " And B.序号=[1]") & _
            " And A.系统 is NULL And B.系统 is Null And C.系统 is Null" & _
            " Union ALL" & _
            " Select 99999 as SCOL,888 as FLAG,A.ID,A.上级ID,A.标题,A.模块" & _
            " From zlMenus A,zlPrograms B,zlReports C" & _
            " Where A.模块=B.序号 And A.组别='缺省' And C.程序ID=A.模块 " & _
            " And Upper(B.部件)='ZL9REPORT'" & IIF(lngProgID = 0, "", " And B.序号=[1]") & _
            " And A.系统 is NULL And B.系统 is Null And C.系统 is Null" & _
            " ) Order by SCOL,FLAG,ID"
    Else
        '只显示固定部份报表(已授权部份)
        strSQL = _
            "Select Distinct * From (" & _
            " Select 编号 as SCOL,0 as Flag,-编号 as ID,-NULL as 上级ID,'['||编号||']'||名称 as 标题,-NULL as 模块 From zlSystems Union ALL" & _
            " Select 99999 as SCOL,Level as FLAG,ID,Nvl(上级ID,-系统) as 上级ID,标题,模块 From zlMenus Where 组别='缺省' And 模块 is NULL" & _
            " Start With 上级ID is NULL And 组别='缺省' Connect by Prior ID=上级ID And 组别='缺省'" & _
            " Union ALL" & _
            " Select 99999 as SCOL,999 as FLAG,A.ID,A.上级ID,A.标题,A.模块" & _
            " From zlMenus A,zlPrograms B,zlRPTGroups C,(Select 系统,序号 From zlRegFunc Group By 系统,序号) D" & _
            " Where A.模块=B.序号 And A.组别='缺省' And C.程序ID=A.模块 " & _
            " And Upper(B.部件)='ZL9REPORT'" & IIF(lngProgID = 0, "", " And B.序号=[1]") & _
            " And A.系统=B.系统 And A.系统=C.系统 And Trunc(B.系统/100)=D.系统 And B.序号=D.序号" & _
            " Union ALL" & _
            " Select 99999 as SCOL,888 as FLAG,A.ID,A.上级ID,A.标题,A.模块" & _
            " From zlMenus A,zlPrograms B,zlReports C,(Select 系统,序号 From zlRegFunc Group By 系统,序号) D" & _
            " Where A.模块=B.序号 And A.组别='缺省' And C.程序ID=A.模块 " & _
            " And Upper(B.部件)='ZL9REPORT'" & IIF(lngProgID = 0, "", " And B.序号=[1]") & _
            " And A.系统=B.系统 And A.系统=C.系统 And Trunc(B.系统/100)=D.系统 And B.序号=D.序号" & _
            " ) Order by SCOL,FLAG,ID"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngProgID)
    Set GetMainTreeMenu = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetModuleTreeMenu(ByVal lngRPTID As Long) As ADODB.Recordset
'功能：获取发布到模块的报表树形菜单体系
'参数：lngRPTID=要发布或取消发布的报表ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '按菜单显示模块的方式
    '-------------------------------------------------------------------------------------------------------------
    '系统 + 中间菜单 + 模块菜单(授权模块) + 发布报表(发布到授权模块下)
    '注意同一模块可能重复位于不同菜单,都显示(包括下面的报表)
    '排开独立的自定义报表模块(部件='zl9Report')
    
    '排开无有效模块的菜单部份(比较慢)
    strSQL = _
        " Select Distinct Id From zlMenus Where 组别='缺省'" & _
        " Start With (系统,模块) In(Select 系统,序号 From zlPrograms Where Upper(部件)<>Upper('zl9Report'))" & _
        " Connect By Prior 上级ID=Id"
    
    strSQL = _
        " Select '1' as Sort1,To_Char(编号) as Sort2," & _
        "   'S'||编号 as ID,Null as 上级ID,编号 as 系统,-Null as 程序ID,Null as 功能,'['||编号||']'||名称 as 标题" & _
        " From zlSystems" & _
        " Union ALL " & _
        " Select '2' as Sort1,To_Char(Level) as Sort2," & _
        "   'T'||ID as ID,Decode(上级ID,NULL,'S'||系统,'T'||上级ID) as 上级ID,系统,-Null as 程序ID,Null as 功能,标题" & _
        " From zlMenus Where 组别='缺省' And 模块 is Null" & _
        " Start With 上级ID is NULL And 组别='缺省' Connect by Prior ID=上级ID And 组别='缺省'" & _
        " Union ALL " & _
        " Select '3' as Sort1,To_Char(B.序号) as Sort2," & _
        "   'M'||B.序号||'_'||A.ID as ID,'T'||A.上级ID as 上级ID,B.系统,B.序号 as 程序ID,Null as 功能,B.标题" & _
        " From zlMenus A,zlPrograms B,(Select 系统,序号 From zlRegFunc Group By 系统,序号) C" & _
        " Where A.组别='缺省' And A.系统=B.系统 And A.模块=B.序号 And Upper(B.部件)<>Upper('zl9Report')" & _
        " And Trunc(B.系统/100)=C.系统 And B.序号=C.序号" & _
        " Union All " & _
        " Select '4' as Sort1,C.编号 as Sort2," & _
        "   'R'||Rownum as ID,'M'||B.程序ID||'_'||X.ID as 上级ID,B.系统,B.程序ID,B.功能,'['||C.编号||']'||C.名称 as 标题" & _
        " From zlMenus X,zlPrograms A,zlRPTPuts B,zlReports C,(Select 系统,序号 From zlRegFunc Group By 系统,序号) D" & _
        " Where X.组别='缺省' And X.系统=A.系统 And X.模块=A.序号" & _
        "   And A.系统=B.系统 And A.序号=B.程序ID And Upper(A.部件)<>Upper('zl9Report')" & _
        "   And Trunc(A.系统/100)=D.系统 And A.序号=D.序号" & _
        "   And B.报表ID=C.ID And C.ID=[1]" & _
        " Order by Sort1,Sort2"
    
    '只显示模块的方式
    '-------------------------------------------------------------------------------------------------------------
    strSQL = _
        " Select '1' as Sort1,To_Char(编号) as Sort2," & _
        "   'S'||编号 as ID,Null as 上级ID,编号 as 系统,-Null as 程序ID,Null as 功能,'['||编号||']'||名称 as 标题" & _
        " From zlSystems" & _
        " Union ALL " & _
        " Select '3' as Sort1,To_Char(B.序号) as Sort2," & _
        "   'M'||B.序号||'_'||B.系统 as ID,'S'||B.系统 as 上级ID,B.系统,B.序号 as 程序ID,Null as 功能,'['||B.序号||']'||B.标题" & _
        " From zlPrograms B,(Select 系统,序号 From zlRegFunc Group By 系统,序号) C" & _
        " Where Upper(B.部件)<>Upper('zl9Report') And Trunc(B.系统/100)=C.系统 And B.序号=C.序号" & _
        " Union All " & _
        " Select '4' as Sort1,C.编号 as Sort2," & _
        "   'R'||Rownum as ID,'M'||B.程序ID||'_'||B.系统 as 上级ID,B.系统,B.程序ID,B.功能,'['||C.编号||']'||C.名称 as 标题" & _
        " From zlPrograms A,zlRPTPuts B,zlReports C,(Select 系统,序号 From zlRegFunc Group By 系统,序号) D" & _
        " Where A.系统=B.系统 And A.序号=B.程序ID And Upper(A.部件)<>Upper('zl9Report')" & _
        "   And Trunc(A.系统/100)=D.系统 And A.序号=D.序号" & _
        "   And B.报表ID=C.ID And C.ID=[1]" & _
        " Order by Sort1,Sort2"
    
    '固定允许发布到人事、成本、帐务系统的模块，其它系统仅10版本允许
    strSQL = "Select A.* From (" & strSQL & ") A,zlSystems B" & _
        " Where A.系统=B.编号 And (To_Number(Substr(B.版本号,1,Instr(B.版本号,'.')-1))>=10 Or Trunc(编号/100) IN(2,5,7))" & _
        " Order by Sort1,Sort2"
    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
    Set GetModuleTreeMenu = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mnuEdit_Publish_Main_Click()
    Call ReportGrantToNavigator
End Sub

Private Sub mnuEdit_unPub_Main_Click()
    Call ReportRevokeFromNavigator
End Sub

Private Sub mnuEdit_Guide_Click()
    Dim objReport As Report, objItem As Object
    Dim lngNext As Long, lngSys As Long, strSQL As String
    Dim i As Integer
    
    Set objReport = New Report
    With objReport
        .进纸 = 15 '缺省为自动选择
        '缺省使用当前打印机
        If Printers.count > 0 Then .打印机 = Printer.DeviceName
        '缺省为A4幅面,为纵向
        .Fmts.Add 1, "格式1", INIT_WIDTH, INIT_HEIGHT, 9, 1, False, 0, "_1"
    End With
    
    frmGuide.blnNew = True
    Set frmGuide.objReport = objReport
    Set frmGuide.mobjFmt = objReport.Fmts(1)
    frmGuide.Show 1, Me
    
    If gblnOK Then
        If cboSys.ListIndex <> 0 Then cboSys.ListIndex = 0
        Me.Refresh
        With frmGuide
            Set objReport.Items = .objGuide.Items
            Set objReport.Datas = .objGuide.Datas
            Set objReport.Fmts = .objGuide.Fmts
            
            '增加报表
            'lngSys = Split(GetSysNO, ",")(0)
            lngNext = GetNextID("zlReports")
            strSQL = "Insert Into zlReports(ID,编号,名称,说明,系统,密码) Values(" & _
                lngNext & ",'" & .txtNO.Text & "','" & .txtTitle.Text & "','" & _
                .txtNote.Text & "'," & IIF(lngSys = 0, "NULL", lngSys) & "," & AdjustStr(GetPass(.txtNO, .txtTitle)) & ")"
        
            On Error GoTo errH
            gcnOracle.BeginTrans
            gcnOracle.Execute strSQL
            gcnOracle.CommitTrans
            On Error GoTo 0
            
            '报表内容
            If Not SaveReport(lngNext, objReport, sta.Panels(2)) Then
                On Error GoTo errH
                gcnOracle.BeginTrans
                gcnOracle.Execute "Delete From zlReports Where ID=" & lngNext
                gcnOracle.CommitTrans
                On Error GoTo 0
                MsgBox "在生成报表时遇到意外错误,请重试该操作！", vbInformation, App.Title
                Unload frmGuide: Exit Sub
            End If
        
            '界面内容
            Set objItem = lvwReport.ListItems.Add(, "_" & lngNext, .txtTitle.Text, "Report", "Report")
            objItem.Tag = 0
            objItem.SubItems(RC_编号) = .txtNO.Text
            objItem.SubItems(RC_说明) = .txtNote.Text
            objItem.SubItems(RC_修改时间) = Format(Currentdate, "yyyy-MM-dd")
            
            '更新选项
            For i = 1 To lvwReport.ListItems.count
                lvwReport.ListItems(i).Selected = (i >= lvwReport.ListItems.count)
            Next
            
            '更新状态
            lvwReport.SelectedItem.EnsureVisible
            lvwReport_ItemClick lvwReport.SelectedItem
        End With
        Unload frmGuide
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
    Unload frmGuide
End Sub

Private Sub lvwReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strPath As String
    Static objItem As Object
    
    If mblnMouseDown And Button = 1 Then
        lvwReport.DragIcon = lvwReport.SelectedItem.CreateDragImage
        lvwReport.Drag 1
    Else
        Set lvwReport.DragIcon = Nothing
        lvwReport.Drag 0
        mblnMouseDown = False
    End If
    If Not objItem Is Nothing And Not lvwReport.HitTest(X, Y) Is Nothing Then
        If objItem.Key = lvwReport.HitTest(X, Y).Key Then Exit Sub
    End If
    
    Set objItem = lvwReport.HitTest(X, Y)
    If Not objItem Is Nothing Then
        lvwReport.ToolTipText = objItem.SubItems(RC_说明)
    End If
End Sub

Private Sub mnuFile_Exp_Click()
    Dim strMsg As String
    Dim strPath As String
    Dim strFile As String
    Dim i As Long, lngSelCount As Long
    
    If lvwReport.SelectedItem Is Nothing Then
        MsgBox "当前没有报表可以导出！", vbInformation, App.Title: Exit Sub
    Else
        'SelectedItem只代表第一个选中得行，因此需循环遍历，查看是否多选
        lngSelCount = 0
        For i = 1 To lvwReport.ListItems.count
            If lvwReport.ListItems(i).Selected Then lngSelCount = lngSelCount + 1
        Next
        If lngSelCount = 1 Then
            strMsg = frmMsgBox.ShowMsgBox(App.Title, "请选择报表导出方式。^导出当前清单中的所有报表时，文件自动按""[编号]名称""命名；^如果导出目录中存在相同名称的报表文件，文件内容将被覆盖。", "所有报表(&Y),!当前报表(&N),?取消(&C)", Me)
             If strMsg = "" Then Exit Sub
        End If
    End If
    
    strPath = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Path", "Export", GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Path", "Import", App.Path))
    If strMsg = "当前报表" Then
        cdg.DialogTitle = "导出报表文件"
        cdg.Filter = "自定义报表文件|*.ZLR"
        cdg.Flags = &H200000 Or &H4 Or &H2 Or &H800 Or &H4000
        cdg.InitDir = strPath
        
        strFile = "[" & lvwReport.SelectedItem.SubItems(RC_编号) & "]" & lvwReport.SelectedItem.Text & ".ZLR"  '缺省以报表名称作文件名
        strFile = Replace(strFile, "\", "﹨")
        strFile = Replace(strFile, "/", "∕")
        strFile = Replace(strFile, ":", "：")
        strFile = Replace(strFile, "*", "﹡")
        strFile = Replace(strFile, "?", "？")
        strFile = Replace(strFile, """", "")
        strFile = Replace(strFile, "<", "〈")
        strFile = Replace(strFile, ">", "〉")
        strFile = Replace(strFile, "|", "∣")
        cdg.FileName = strFile
        cdg.CancelError = True
        
        On Error Resume Next
        
        cdg.ShowSave
        If Err.Number = 0 Then
            Err.Clear
            On Error GoTo 0
            Me.Refresh
            SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Path", "Export", Left(cdg.FileName, Len(cdg.FileName) - Len(cdg.FileTitle))
            Call ExportReport(CLng(Mid(lvwReport.SelectedItem.Key, 2)), cdg.FileName)
            VBA.Beep
        End If
    ElseIf strMsg = "所有报表" Or lngSelCount > 1 Then
        strFile = BrowseForFolder(Me.hwnd, "选择报表导出目录", strPath)
        If strFile <> "" Then
            strPath = strFile
            SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Path", "Export", strPath
            lngSelCount = IIF(strMsg = "", lngSelCount, lvwReport.ListItems.count)
            If MsgBox("本次共导出 " & lngSelCount & " 张报表到 " & strPath & "，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
            
            For i = 1 To lvwReport.ListItems.count
                If lvwReport.ListItems(i).Selected Or strMsg <> "" Then
                    Call ShowFlash("正在导出:" & lvwReport.ListItems(i).Text & ".ZLR", i / lngSelCount, Me, True)
                    
                    strFile = "[" & lvwReport.ListItems(i).SubItems(RC_编号) & "]" & lvwReport.ListItems(i).Text & ".ZLR"
                    If Not ExportReport(CLng(Mid(lvwReport.ListItems(i).Key, 2)), strPath & "\" & strFile) Then
                        Call ShowFlash
                        If MsgBox("导出报表时出现错误，要继续导出下一张报表吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
                    End If
                End If
            Next
            Call ShowFlash
        End If
    End If
End Sub

Private Sub mnuFile_Imp_Click()
    Dim arrFile As Variant, strFile As String, i As Long
    Dim lngSys As Long, LngGroupID As Long, lngReportID As Long
    Dim rsFiles As ADODB.Recordset
    
    On Error GoTo errH
    cdg.DialogTitle = "选择导入报表"
    cdg.Filter = "自定义报表文件|*.ZLR"
    cdg.Flags = &H200 Or &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
    cdg.InitDir = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Path", "Import", GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Path", "Export", App.Path))
    cdg.FileName = ""
    cdg.MaxFileSize = 32767
    cdg.CancelError = True
    On Error Resume Next
    cdg.ShowOpen
    If Err.Number = 0 Then
        On Error GoTo errH
        Me.Refresh
        If cdg.FileTitle = "" Then
            '选择多个文件导入
            SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Path", "Import", Left(cdg.FileName, InStr(cdg.FileName, Chr(0)) - 1)
            arrFile = Split(cdg.FileName, Chr(0))
            For i = 1 To UBound(arrFile)
                strFile = strFile & "|" & arrFile(0) & "\" & arrFile(i)
            Next
            strFile = Mid(strFile, 2)
        Else
            '选择单个文件导入
            SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Path", "Import", Left(cdg.FileName, Len(cdg.FileName) - Len(cdg.FileTitle))
            strFile = cdg.FileName
        End If
        If strFile = "" Then Exit Sub
        arrFile = Split(strFile, "|")
        lngSys = cboSys.ItemData(cboSys.ListIndex)
        LngGroupID = IIF(lvwGroup.SelectedItem.Key = "_-1", 0, Val(Mid(lvwGroup.SelectedItem.Key, 2)))
        If Not lvwReport.SelectedItem Is Nothing Then lngReportID = CLng(Mid(lvwReport.SelectedItem.Key, 2))
        'FilePath=报表全路径；FileName=报表文件名；组ID=报表要导入的报表组ID
        '同名ID=与将要导入的报表同名的报表的报表ID，固定报表通过编码匹配，非固定通过名称匹配
        '导入类型=0-不导入，1-新增导入,2-覆盖导入;覆盖类型=0-整体覆盖，1-仅数据源覆盖
        'ErrType=0-无错误,1-多个相同报表一起新增，2-多个相同报表一起覆盖，3-系统报表只能覆盖，但是无同名报表。
        '                            4-内容存在问题,5-版本存在问题,6-名称编号存在问题
        'ImportResult=-1-已经成功导入但是报表对象检查未通过，0-不导入,1-导入成功,2-导入失败
        'ImportInfo=报表成功导入后返回的报表信息
        Set rsFiles = CopyNewRec(Nothing, , True, _
                                    Array("FilePath", adVarChar, 1000, Empty, "FileName", adVarChar, 200, Empty, "组ID", adBigInt, Empty, Empty, _
                                             "同名ID", adBigInt, Empty, Empty, "导入类型", adInteger, Empty, Empty, "覆盖类型", adInteger, Empty, Empty, _
                                             "ErrType", adInteger, Empty, Empty, "ImportResult", adInteger, Empty, Empty, "ImportInfo", adVarChar, 200, Empty))
        For i = LBound(arrFile) To UBound(arrFile)
            rsFiles.AddNew Array("FilePath", "FileName", "组ID", "同名ID", "导入类型", "覆盖类型", "ErrType", "ImportResult", "ImportInfo"), _
                                    Array(arrFile(i), gobjFile.GetFileName(arrFile(i)), 0, 0, 0, 0, 0, 0, "")
        
        Next
        Call ImportReportBeach(lngSys, LngGroupID, lngReportID, rsFiles)
    Else
        Err.Clear
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ImportReportBeach(ByVal lngSys As Long, ByVal lngGroup As Long, ByVal lngCurPRTID As Long, ByVal rsFiles As ADODB.Recordset, Optional ByVal blnALLImp As Boolean) As Boolean
'功能：批量导入报表，可以导入1个至多个
'参数：
'          lngSys=当前选择的系统
'          lngGroup=当前选择的记录集
'          rsFiles=需要导入的报表文件
'          lngCurPRTID=当前选择的报表ID
'          blnALLImp=是否是全部倒入，非固定报表全部导入时，也需要读取所有报表
'返回：是否成功导入

    Dim rsReports As New ADODB.Recordset, strSQL As String
    Dim arrTmp As Variant, strInfo As String
    Dim strFilter As String
    Dim intErrType As Integer, intImpType As Integer, lngImpGroup As Long, lngRPTID As Long
    Dim strMsg As String, strOption As String, strReturn As String
    Dim i As Long, lngCount As Long
    Dim blnSingle  As Boolean, strFileName As String
    Dim strCurRPT As String, strSameRPT As String
    
    On Error GoTo errH
    '固定报表，以及非显示独立项下的非固定报表的所有报表分组时，需要读取所有报表
    If lngSys <> 0 Or Not mnuViewOnly.Checked And lngGroup = 0 And lngSys = 0 Or blnALLImp Then
        '查询所有的报表
        strSQL = "Select A.ID,A.编号,A.名称,A.说明,Nvl(B.组id,0) 组id" & vbNewLine & _
                        "From zlReports A,zlRPTSubs B" & vbNewLine & _
                        "Where " & IIF(lngSys = 0, " A.系统 Is Null", "A.系统=[1]") & vbNewLine & _
                        "And  A. ID=B.报表ID(+)" & vbNewLine & _
                        "Order by A.编号"
    Else '非固定报表读取
        If lngGroup <> 0 Then
            strSQL = "Select Id, 编号, 名称,[2] 组id" & vbNewLine & _
                            "From Zlreports" & vbNewLine & _
                            "Where Id In (Select 报表id From Zlrptsubs Where 组id = [2])" & vbNewLine & _
                            "Order By 编号"
        Else
            strSQL = "Select ID,编号,名称,0 组id" & vbNewLine & _
                            "From zlReports" & vbNewLine & _
                            "Where " & IIF(lngSys = 0, " 系统 Is Null", "系统=[1]") & vbNewLine & _
                            "And ID Not In (Select 报表ID From zlRPTSubs)" & vbNewLine & _
                            "Order by 编号"
        End If
    End If
    Set rsReports = CopyNewRec(OpenSQLRecord(strSQL, Me.Caption, lngSys, lngGroup))
    If lngCurPRTID <> 0 Then
        rsReports.Filter = "ID=" & lngCurPRTID
        If rsReports.EOF Then
            MsgBox "当前选中报表已经被删除，请刷新后继续！", vbInformation, App.Title
            Exit Function
        Else
            strCurRPT = "[" & rsReports!编号 & "]" & rsReports!名称
        End If
    End If
    With rsFiles
        '不同子文件导入到同一分组时的同名文件检查
        '具体情况如下：[GROUP_001]住院工作报表ASD，住院工作报表，[GROUP_001]住院工作报表
        '                       这三个子文件的报表可以导入到[GROUP_001]住院工作报分组中
        '不同文件名的报表，可能是同一个报表。
        '检查导入文件，以及确定导入类型，倒入分组以及覆盖的报表ID等
        .Filter = "": .Sort = "FilePath Desc"
        blnSingle = rsFiles.RecordCount = 1 '是否单个报表导入
        If blnSingle Then strFileName = rsFiles!FileName
        Do While Not .EOF
            intErrType = 0: intImpType = 0: lngImpGroup = 0: lngRPTID = 0
            arrTmp = Split(GetReportInfo(!FilePath & ""), ";") '获取文件信息
            If UBound(arrTmp) <> 2 Then
                intErrType = 4 '文件检查
            ElseIf Val(arrTmp(2)) <> 9 Then
                intErrType = 5  '版本检查
                If blnSingle Then strFileName = strFileName & "(原始名称：[" & arrTmp(0) & "]" & arrTmp(1) & ")"
            Else
                If blnSingle Then strFileName = strFileName & "(原始名称：[" & arrTmp(0) & "]" & arrTmp(1) & ")"
                If lngSys = 0 Then '非系统报表要求分组的报表中不能存在相同报表
                    '非固定报表全部导入已经确定报表要导入的分组
                    rsReports.Filter = "名称='" & arrTmp(1) & "' And 编号='" & arrTmp(0) & "' And ID>0 " & IIF(blnALLImp, " And 组ID=" & !组ID, "")
                    If rsReports.EOF Then rsReports.Filter = "名称='" & arrTmp(1) & "'  And ID>0 " & IIF(blnALLImp, " And 组ID=" & !组ID, "")
                Else '系统报表通过编号直接查找
                    rsReports.Filter = "名称='" & arrTmp(1) & "' And 编号='" & arrTmp(0) & "' And ID>0"
                    If rsReports.EOF Then rsReports.Filter = "编号='" & arrTmp(0) & "' And ID>0"
                End If
                '确定报表导入的分组，如果存在的同名的，优先查找没有分组的报表
                rsReports.Sort = "ID Desc,组ID"
                If Not rsReports.EOF Then
                    lngRPTID = rsReports!ID: lngImpGroup = rsReports!组ID
                    If lngRPTID = 0 Then
                        intErrType = 1 '该报表已经被标记新增
                    ElseIf lngRPTID < 0 Then
                        intErrType = 2 '该报表已经被标记覆盖
                    Else
                        intImpType = 2
                        '编号名称不匹配
                        If (CStr(arrTmp(0)) <> rsReports!编号 & "" Or CStr(arrTmp(1)) <> rsReports!名称) Then intErrType = 6
                        rsReports.Update "Id", lngRPTID * -1 '标记已经覆盖
                        If blnSingle Then strSameRPT = "[" & rsReports!编号 & "]" & rsReports!名称
                    End If
                Else
                    If lngSys <> 0 Then
                        intErrType = 3  '系统固定报表必须覆盖同名报表
                    Else
                        intImpType = 1  '非系统报表没有同名，则新增报表
                        If lngSys = 0 And blnALLImp Then lngImpGroup = !组ID '非固定报表导入取原来的分组
                        '该报表是新增报表，则加入缓存，防止多次增加
                        rsReports.AddNew Array("Id", "编号", "名称", "组iD"), Array(lngRPTID, arrTmp(0), arrTmp(1), !组ID)
                    End If
                End If
            End If
            If lngSys = 0 And blnALLImp Then lngImpGroup = !组ID '非固定报表导入取原来的分组
            .Update Array("组ID", "同名ID", "导入类型", "ErrType"), Array(lngImpGroup, lngRPTID, intImpType, intErrType)
            .MoveNext
        Loop
        If blnSingle Then
            .Filter = ""
            Select Case !ErrType
                Case 4
                    MsgBox "报表""" & strFileName & """由于内容存在问题而无法导入！", vbInformation, App.Title
                    Exit Function
                Case 5
                    MsgBox "报表""" & strFileName & """由于版本不对而无法导入！", vbInformation, App.Title
                    Exit Function
                Case 3
                    If lngCurPRTID <> 0 Then '更新状态，默认覆盖当前的报表
                        .Update Array("组ID", "同名ID", "导入类型", "ErrType"), Array(lngGroup, lngCurPRTID, 2, 6)
                    Else
                        MsgBox "请选择你要覆盖的报表后继续！", vbInformation, App.Title
                        Exit Function
                    End If
            End Select
            Select Case !导入类型
                Case 1
                    strReturn = frmMsgBox.ShowMsgBox(App.Title, "是否新增导入报表""" & strFileName & """！", "新增导入(&N),!?取消(&C)", Me)
                Case 2
                    If lngSys = 0 And lngGroup = 0 Then '所有系统共享的为分组的报表,此时可以存在新增报表选项
                        If lngCurPRTID = !同名ID Then
                            strMsg = IIF(!ErrType = 6, "报表""" & strFileName & """编号或名称" & vbNewLine & "与要覆盖的当前选择报表""" & strCurRPT & """不相符，请选择确认！", _
                                        "报表""" & strFileName & """编号和名称" & vbNewLine & "与当前选择报表""" & strCurRPT & """都相符，请选择确认！") & vbNewLine & "^^注意：如果要覆盖报表，请先对要覆盖报表进行备份。"
                            strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "覆盖当前(&S),新增导入(&N),!?取消(&C)", Me)
                        ElseIf lngCurPRTID = 0 Then
                            strMsg = IIF(!ErrType = 6, "报表""" & strFileName & """存在部分匹配的报表""" & strSameRPT & """," & vbNewLine & "但是二者编号或名称不相符，请选择确认！", _
                                        "报表""" & strFileName & """存在编码与名称均相符的报表""" & strSameRPT & """，请选择确认！") & vbNewLine & "^^注意：如果要覆盖报表，请先对要覆盖报表进行备份。"
                            strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "覆盖匹配(&O),新增导入(&N),!?取消(&C)", Me)
                        Else
                            strMsg = IIF(!ErrType = 6, "报表""" & strFileName & """的编号或名称" & vbNewLine & "与部分匹配报表""" & strSameRPT & """" & vbNewLine & "以及当前选择报表""" & strCurRPT & """均不相符，请选择确认！", _
                                        "报表""" & strFileName & """编号或名称" & vbNewLine & "与当前选择报""" & strCurRPT & """不相符，" & vbNewLine & "但是存在编码与名称均相符的报表""" & strSameRPT & """，请选择确认！") & vbNewLine & "^^注意：如果要覆盖报表，请先对要覆盖报表进行备份。"
                            strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "覆盖当前(&S),覆盖匹配(&O),新增导入(&N),!?取消(&C)", Me)
                        End If
                    Else
                       If lngCurPRTID = !同名ID Then
                            strMsg = IIF(!ErrType = 6, "报表""" & strFileName & """编号或名称" & vbNewLine & "与要覆盖的当前选择报表""" & strCurRPT & """不相符，请选择确认！", _
                                        "报表""" & strFileName & """编号和名称" & vbNewLine & "与当前选择报表""" & strCurRPT & """都相符，请选择确认！") & vbNewLine & "^^注意：如果要覆盖报表，请先对要覆盖报表进行备份。"
                            strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "覆盖当前(&S),!?取消(&C)", Me)
                        ElseIf lngCurPRTID = 0 Then
                            strMsg = IIF(!ErrType = 6, "报表""" & strFileName & """存在部分匹配的报表""" & strSameRPT & """," & vbNewLine & "但是二者编号或名称不相符，请选择确认！", _
                                        "报表""" & strFileName & """存在" & vbNewLine & "编码与名称均相符的报表""" & strSameRPT & """，请选择确认！") & vbNewLine & "^^注意：如果要覆盖报表，请先对要覆盖报表进行备份。"
                            strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "覆盖匹配(&O),!?取消(&C)", Me)
                        Else
                            strMsg = IIF(!ErrType = 6, "报表""" & strFileName & """的编号或名称" & vbNewLine & "与部分匹配报表""" & strSameRPT & """" & vbNewLine & " 以及当前选择报表""" & strCurRPT & """均不相符，请选择确认！", _
                                        "报表""" & strFileName & """编号或名称" & vbNewLine & "与当前选择报""" & strCurRPT & """不相符，" & vbNewLine & "但是存在编码与名称均相符的报表""" & strSameRPT & """，请选择确认！") & vbNewLine & "^^注意：如果要覆盖报表，请先对要覆盖报表进行备份。"
                            strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "覆盖当前(&S),覆盖匹配(&O),!?取消(&C)", Me)
                        End If
                    End If
            End Select
            If strReturn = "" Then
                Exit Function
            ElseIf strReturn = "新增导入" Then
                .Update Array("组ID", "同名ID", "导入类型", "ErrType"), Array(lngGroup, 0, 1, 0)
            Else
                If strReturn = "覆盖当前" Then
                    .Update Array("组ID", "同名ID", "导入类型", "ErrType"), Array(lngGroup, lngCurPRTID, 2, 0)
                Else
                    .Update Array("导入类型", "ErrType"), Array(2, 0)
                End If
                strMsg = frmMsgBox.ShowMsgBox(App.Title, "是否只导入数据源？" & vbNewLine & "只导入数据源可以保持现有报表的格式，更详细的情况请咨询系统管理员！", "仅数据源(&D),!?整体导入(&F)", Me)
                If strMsg = "仅数据源" Then
                    .Update "覆盖类型", 1
                End If
            End If
        Else
            If MsgBox("当前导入多张报表，系统将自动寻找编码或名称匹配的报表进行覆盖。请确认是否继续！", vbInformation + vbYesNo, App.Title) = vbNo Then
                Exit Function
            End If
            '不能导入的类型信息生成
            .Filter = "ErrType>0 And ErrType<6": .Sort = "ErrType": intImpType = 0
            Do While Not .EOF
                If intImpType <> Val(!ErrType & "") Then
                    If intImpType <> 0 Then
                        strMsg = strMsg & vbNewLine
                    End If
                    intImpType = Val(!ErrType & ""): lngCount = 0
                    Select Case intImpType
                        Case 1
                            strMsg = strMsg & vbNewLine & "以下报表由于存在相同内容的报表而无法新增导入："
                        Case 2
                            strMsg = strMsg & vbNewLine & "以下报表由于存在相同内容的报表而无法覆盖导入："
                        Case 3
                            strMsg = strMsg & vbNewLine & "以下报表由于没有可以覆盖的报表而无法导入："
                        Case 4
                            strMsg = strMsg & vbNewLine & "以下报表由于内容存在问题而无法导入："
                        Case 5
                            strMsg = strMsg & vbNewLine & "以下报表由于版本不对而无法导入："
                    End Select
                End If
                If lngCount < 4 Then
                    strMsg = strMsg & vbNewLine & !FileName
                ElseIf lngCount = 4 Then
                    strMsg = strMsg & vbNewLine & "... ..."
                End If
                lngCount = lngCount + 1: .MoveNext
                If .EOF Then strMsg = strMsg & vbNewLine
            Loop
            .Filter = "导入类型<>0"
            If .RecordCount = 0 Then '没有导入报表
                MsgBox "没有可以导入的报表！" & Mid(strMsg, 1, Len(strMsg) - 2) & "。", vbInformation, App.Title
                Exit Function
            End If
            '文件名以及编码不匹配提示
            .Filter = "ErrType=6"
            If Not .EOF Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "编号或名称与覆盖的报表不相符，请选择确认："
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
                .Filter = "ErrType=0" '不存在可以直接导入的，则提示是否继续
                If .RecordCount = 0 Then
                    strReturn = frmMsgBox.ShowMsgBox(App.Title, Mid(strMsg, 1, Len(strMsg) - Len(vbNewLine)), "整体覆盖(&A),数据源覆盖(&D),!?取消(&C)", Me)
                    If strReturn = "" Then Exit Function
                End If
            End If
            .Filter = "导入类型=2 And ErrType=0": .Sort = "ErrType" '存在覆盖报表，则提示选择整体覆盖，还是数据源覆盖
            If Not .EOF Then
                strMsg = strMsg & vbNewLine & "以下报表将会覆盖原有报表，请选择确认："
                strOption = "整体覆盖(&A),数据源覆盖(&D),!?取消(&C)"
                lngCount = 0
            End If

            Do While Not .EOF
                If lngCount < 4 Then
                    strMsg = strMsg & vbNewLine & !FileName
                ElseIf lngCount = 4 Then
                    strMsg = strMsg & vbNewLine & "... ..."
                End If
                lngCount = lngCount + 1: .MoveNext
                If .EOF Then strMsg = strMsg & vbNewLine
            Loop
            .Filter = "导入类型=1" '新增导入
            If .RecordCount <> 0 And strReturn = "" And strOption = "" Then '所有报表新增
                strReturn = frmMsgBox.ShowMsgBox(App.Title, Mid(strMsg, Len(vbNewLine) + 1) & "请确认是否导入？", "导入(&N),!?取消(&C)", Me)
                If strReturn = "" Then Exit Function
            End If
            '选择覆盖类型
            If strReturn = "" And strOption <> "" Then '存在覆盖,且不存在ErrType=6的类型
                strReturn = frmMsgBox.ShowMsgBox(App.Title, Mid(strMsg, Len(vbNewLine) + 1, Len(strMsg) - Len(vbNewLine) * 2), strOption, Me)
                If strReturn = "" Then Exit Function
            End If
        End If
        If strReturn = "数据源覆盖" Then
            .Filter = "导入类型=2"
            Do While Not .EOF
                .Update "覆盖类型", 1
                .MoveNext
            Loop
        End If
        Screen.MousePointer = 11
        .Filter = "导入类型<>0": .Sort = "导入类型"
        lngCount = .RecordCount
        Do While Not .EOF
            If Not blnSingle Then
                Call ShowFlash("正在导入:" & !FileName, i / lngCount, Me, True)
            Else
                Call ShowFlash("正在导入:" & !FileName, , Me, True)
            End If
            Me.Refresh
            DoEvents
            strInfo = ImportReport(!FilePath & "", Val(!同名ID & ""), Val(!覆盖类型 & "") = 1, Val(!组ID & ""))
            .Update Array("ImportResult", "ImportInfo"), Array(IIF(strInfo <> "", 1, 2), strInfo)
            '报表对象权限检查
            If strInfo <> "" Then
                arrTmp = Split(strInfo, "|")
                If Not CheckReportPriv(CLng(arrTmp(0))) Then
                    .Update Array("ImportResult", "同名ID"), Array(-1, Val(arrTmp(0)))
                Else
                    .Update "同名ID", Val(arrTmp(0))
                End If
            End If
            i = i + 1
            .MoveNext
        Loop
        Call ShowFlash
        If Not blnSingle Then
            lngGroup = Val(Mid(mstrPreGroup, 2))
        Else
            .Filter = ""
            lngGroup = Val(!组ID & "")
        End If
        '刷新界面，重新加载数据
        On Error Resume Next
        mstrPreGroup = ""
        '非固定报表全部导入需要刷新分组
        If lngSys = 0 And blnALLImp Then
            '记录报表组ID
            Call ReadGroups
        End If
        '重新定位当前分组
        For i = 1 To lvwGroup.ListItems.count
            If lvwGroup.ListItems(i).Key = "_" & IIF(lngGroup = 0, -1, lngGroup) Then
                lvwGroup.ListItems(i).Selected = True
            Else
                lvwGroup.ListItems(i).Selected = False
            End If
        Next
        lvwGroup.SelectedItem.EnsureVisible: lvwGroup.Refresh
        Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
        '清空选择的报表
        For i = 1 To lvwReport.ListItems.count
            lvwReport.ListItems(i).Selected = False
        Next
        '导入报表选择
        .Filter = "组ID= " & lngGroup
        .Sort = "同名ID"
        Do While Not .EOF
            lvwReport.ListItems("_" & !同名ID).Selected = True
            .MoveNext
        Loop
        lvwReport.SelectedItem.EnsureVisible: lvwReport.Refresh
        Call cbr.Refresh
        Err.Clear: On Error GoTo errH
        '导入情况提示
        strMsg = ""
        If Not blnSingle Then
            .Filter = "ImportResult=1 Or ImportResult=-1"
            If .RecordCount = 0 Then
                strMsg = "所有报表均为导入成功。"
            Else
                strMsg = "成功导入了 " & .RecordCount & " 张报表。"
            End If
            .Filter = "ImportResult=2"
            If .RecordCount <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "以下报表的报表文件内容可能已被非法修改："
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
            .Filter = "ImportResult=-1 And 导入类型=1"
            If .RecordCount <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "你没有权限查询以下导入报表中全部或部份数据对象："
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
            .Filter = "ImportResult=-1 And 导入类型=2"
            If .RecordCount <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "你没有权限查询以下导入报表中全部或部份数据对象,在使用该报表之前,请手工对报表内容进行调整："
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
            .Filter = "ImportResult=1 And 导入类型=2"
            If .RecordCount <> 0 And lngSys <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "以下报表成功覆盖相应报表,你可能需要重新授权才能正常使用这些报表："
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
            .Filter = "ImportResult=2"
            If .RecordCount <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "以下报表导入失败："
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
        Else
            .Filter = ""
            Select Case !ImportResult
                Case -1
                    strMsg = "你没有权限查询报表""" & strFileName & """中全部或部份数据对象" & IIF(!导入类型 = 2, "。你可能需要手工对报表内容进行调整并重新授权才能正常使用该报表！", "！")
                Case 1
                    strMsg = "报表""" & strFileName & """导入成功" & IIF(!导入类型 = 2, "。你可能需要重新授权才能正常使用该报表！", "！")
                Case 2
                    strMsg = "报表""" & strFileName & """" & IIF(!导入类型 = 2, "覆盖失败。报表文件内容可能已被非法修改！", "新增导入失败！")
            End Select
        End If
        MsgBox strMsg, vbInformation, App.Title
        Screen.MousePointer = 0
    End With
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    Call ShowFlash
    Call SaveErrLog
End Function

Private Sub ReadSystem()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    On Error GoTo errH
    
    cboSys.Clear
    cboSys.AddItem "所有系统共享"
    
    strSQL = "Select 编号,名称 From zlSystems Order by 编号"
    Call OpenRecord(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        cboSys.AddItem Lpad(rsTmp!编号, 4) & "-" & rsTmp!名称
        cboSys.ItemData(cboSys.NewIndex) = rsTmp!编号
        rsTmp.MoveNext
    Next
    cboSys.ListIndex = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ReadGroups() As Boolean
    Dim rsReportGroup As New ADODB.Recordset
    Dim strSQL As String, ItemThis As ListItem
    '装入所有报表组
    
    strSQL = "Select ID,编号,名称,说明,系统,程序ID,发布时间 , zlSpellCode(名称) 简码 From zlRPTGroups " & _
        " Where " & IIF(cboSys.ItemData(cboSys.ListIndex) = 0, " 系统 Is Null", " 系统=[1]")
    Set rsReportGroup = OpenSQLRecord(strSQL, Me.Caption, cboSys.ItemData(cboSys.ListIndex))
    
    LockWindowUpdate lvwGroup.hwnd
    lvwGroup.ListItems.Clear
    lvwGroup.ListItems.Add , "_-1", "所有报表", 5, 5
    
    With rsReportGroup
        Do While Not .EOF
            If Not IsNull(!系统) Then
                If Not IsNull(!发布时间) Then     '固定报表(不允许发布或取消发布)
                    Set ItemThis = lvwGroup.ListItems.Add(, "_" & !ID, !名称, 4, 4)
                Else
                    Set ItemThis = lvwGroup.ListItems.Add(, "_" & !ID, !名称, 3, 3)
                End If
            Else
                If Not IsNull(!发布时间) Then     '非固定报表(允许发布及取消发布)
                    Set ItemThis = lvwGroup.ListItems.Add(, "_" & !ID, !名称, 2, 2)
                Else
                    Set ItemThis = lvwGroup.ListItems.Add(, "_" & !ID, !名称, 1, 1)
                End If
            End If
            ItemThis.SubItems(GC_编号) = !编号
            ItemThis.SubItems(GC_说明) = Nvl(!说明)
            ItemThis.SubItems(GC_发布时间) = Format(Nvl(!发布时间), "yyyy-MM-dd")
            ItemThis.SubItems(GC_简码) = !简码
            ItemThis.Tag = Val(Nvl(!程序ID, 0))
            .MoveNext
        Loop
    End With
    
    lvwGroup.ListItems("_-1").Selected = True
    lvwGroup.SelectedItem.Selected = True
    
    'Call AutoSizeCol(lvwGroup)
    LockWindowUpdate 0
    
    mstrPreGroup = ""
    Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
    ReadGroups = True
End Function

Private Sub ReportGrantToNavigator()
'功能：发布当前报表(组)到导航台,可能不是第一次
    Dim rsTmp As ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    Dim rsSubRPT As New ADODB.Recordset
    Dim objNode As Object, i As Integer, j As Integer, k As Integer
    Dim strObject As String, strOwner As String, strName As String
    Dim lngRPTID As Long, lngProgID As Long, lngMenu As Long
    Dim strTmp As String, lngNewMenu As Long, lngSys As Long
    Dim strSQL As String
    
    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_报表 Then
        '选择所有报表时
        If lvwReport.SelectedItem Is Nothing Then MsgBox "当前没有报表可以发布！", vbInformation, App.Title: Exit Sub
        If lvwReport.SelectedItem.Tag <> 0 Then
            If lvwReport.SelectedItem.Icon = "Fixed" Or lvwReport.SelectedItem.Icon = "PubFixed" Then
                MsgBox "该报表为系统固有的报表,操作不能继续！", vbInformation, App.Title: Exit Sub
            End If
        End If
        If CheckPass(CLng(Mid(lvwReport.SelectedItem.Key, 2))) = False Then
            MsgBox "报表数据错误，不能发布该报表！", vbInformation, App.Title: Exit Sub
        End If
        If Not CheckReportPriv(CLng(Mid(lvwReport.SelectedItem.Key, 2))) Then
            MsgBox "你没有权限查询该报表某些数据源中的对象,操作不能继续！", vbInformation, App.Title
            Exit Sub
        End If
        lngRPTID = CLng(Mid(lvwReport.SelectedItem.Key, 2))
    Else
        '具体某个报表组
        If lvwGroup.SelectedItem Is Nothing Then MsgBox "当前没有报表组可以发布！", vbInformation, App.Title: Exit Sub
        If lvwGroup.SelectedItem.Tag <> 0 Then
            If lvwGroup.SelectedItem.Icon = 3 Or lvwGroup.SelectedItem.Icon = 4 Then
                MsgBox "该报表组为系统固有的报表,操作不能继续！", vbInformation, App.Title: Exit Sub
            End If
        End If
        
        If Me.lvwReport.ListItems.count = 0 Then
            MsgBox "该报表组中不包含任何报表，不能发布！", vbInformation, App.Title
            Exit Sub
        Else
            For i = 1 To lvwReport.ListItems.count - 1
                For j = i + 1 To lvwReport.ListItems.count
                    If lvwReport.ListItems(i).Text = lvwReport.ListItems(j).Text Then
                        MsgBox "该报表组中包含相同名称的报表：""" & lvwReport.ListItems(i).Text & """，不能发布！", vbInformation, App.Title
                        Exit Sub
                    End If
                Next
            Next
        End If
        
        strSQL = "Select ID,名称 From zlReports Where ID in (Select 报表ID From zlRPTSubs Where 组ID=[1])"
        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(lvwGroup.SelectedItem.Key, 2)))
        Do While Not rsCheck.EOF
            If Not CheckReportPriv(rsCheck!ID) Then
                MsgBox "你没有权限查询报表[" & rsCheck!名称 & "]中某些数据源的对象！", vbInformation, App.Title: Exit Sub
            End If
            rsCheck.MoveNext
        Loop
        lngRPTID = CLng(Mid(lvwGroup.SelectedItem.Key, 2))
    End If
    
    '1.选择一个菜单位置
    Set rsTmp = GetMainTreeMenu
    If rsTmp Is Nothing Then MsgBox "读取菜单体系时遇到意外错误,报表发布中断！", vbInformation, App.Title: Exit Sub
    
    Load frmSelTree
    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_报表 Then
        frmSelTree.Caption = "发布报表到导航台 - 菜单位置选择"
    Else
        frmSelTree.Caption = "发布报表组到导航台 - 菜单位置选择"
    End If
    With frmSelTree.tvw
        .Nodes.Clear
        For i = 1 To rsTmp.RecordCount
            If rsTmp!Flag = 0 Then
                Set objNode = .Nodes.Add(, , "_" & rsTmp!ID, rsTmp!标题, "Root")
                objNode.Tag = "请选择本系统下一个具体的菜单位置！"
            Else
                If InStr(1, "888,999", rsTmp!Flag) = 0 Then
                    Set objNode = .Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!ID, rsTmp!标题, "Path")
                Else
                    Set objNode = .Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!ID, rsTmp!标题, IIF(rsTmp!Flag = 999, "GroupNode", "ReportNode"))
                    objNode.ForeColor = vbBlue
                    objNode.Tag = "这是已发布的报表,选择一个菜单位置！"
                    
                    '不能发布到相同位置
                    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_报表 Then
                        If objNode.Text = lvwReport.SelectedItem.Text Then
                            objNode.Parent.Tag = "同一个报表或组不能发布到相同的位置,请选择其他菜单位置！"
                        End If
                    Else
                        If objNode.Text = lvwGroup.SelectedItem.Text Then
                            objNode.Parent.Tag = "同一个报表或组不能发布到相同的位置,请选择其他菜单位置！"
                        End If
                    End If
                End If
            End If
            objNode.Expanded = True
            rsTmp.MoveNext
        Next
        If .Nodes.count > 0 Then .Nodes(1).Selected = True
    End With
    frmSelTree.Show 1, Me
    If Not gblnOK Then Exit Sub
    lngMenu = CLng(Mid(frmSelTree.tvw.SelectedItem.Key, 2)) '要加菜单的上级ID
    Unload frmSelTree
    
    lngNewMenu = GetNextID("zlMenus")
    
    '2.填写程序、权限
    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_报表 Then
        If lvwReport.SelectedItem.Tag <> 0 Then
            '不用再处理
            lngProgID = lvwReport.SelectedItem.Tag
        Else
            lngProgID = GetNewProgID
            
            '分析该报表的数据源访问对象
            strSQL = "Select 对象 From zlRPTDatas Where 对象 is Not NULL And 报表ID=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
            If Not rsTmp.EOF Then
                For i = 1 To rsTmp.RecordCount
                    For j = 0 To UBound(Split(rsTmp!对象, ","))
                        If InStr(strObject & ",", "," & Split(rsTmp!对象, ",")(j) & ",") = 0 Then
                            strObject = strObject & "," & Split(rsTmp!对象, ",")(j)
                        End If
                    Next
                    rsTmp.MoveNext
                Next
            End If
            
            '分析该报表的参数数据源访问对象
            strSQL = "Select B.对象 From zlRPTDatas A,zlRPTPars B Where A.ID=B.源ID And B.对象 is Not NULL And A.报表ID=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
            If Not rsTmp.EOF Then
                For i = 1 To rsTmp.RecordCount
                    For j = 0 To UBound(Split(rsTmp!对象, "|"))
                        strTmp = Split(rsTmp!对象, "|")(j)
                        For k = 0 To UBound(Split(strTmp, ","))
                            If InStr(strObject & ",", "," & Split(strTmp, ",")(k) & ",") = 0 Then
                                strObject = strObject & "," & Split(strTmp, ",")(k)
                            End If
                        Next
                    Next
                    rsTmp.MoveNext
                Next
            End If
            
            If strObject <> "" Then strObject = Mid(strObject, 2)
        End If
    Else
        If lvwGroup.SelectedItem.Tag <> 0 Then
            '不用再处理
            lngProgID = lvwGroup.SelectedItem.Tag
        Else
            lngProgID = GetNewProgID
        End If
    End If
    
    lngSys = cboSys.ItemData(cboSys.ListIndex)
    
    On Error GoTo errH
    
    gcnOracle.BeginTrans
    
    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_报表 Then
        If lvwReport.SelectedItem.Tag = 0 Then
            gcnOracle.Execute "Update zlReports Set 功能='基本',程序ID=" & lngProgID & ",发布时间=Sysdate Where ID=" & lngRPTID
            gcnOracle.Execute "Insert Into zlPrograms(序号,标题,说明,系统,部件)" & _
                " Values(" & lngProgID & ",'" & lvwReport.SelectedItem.Text & "','" & lvwReport.SelectedItem.SubItems(RC_说明) & "'," & _
                IIF(lngSys = 0, "NULL", lngSys) & ",'zl9Report')"
            gcnOracle.Execute "Insert Into zlProgFuncs(系统,序号,功能) Values(" & IIF(lngSys = 0, "NULL", lngSys) & "," & lngProgID & ",'基本')"
            If strObject <> "" Then '该表格有可能不访问数据库
                For i = 0 To UBound(Split(strObject, ","))
                    strOwner = Left(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") - 1)
                    If strOwner <> "SYS" And strOwner <> "ZLTOOLS" And strOwner <> "SYSTEM" Then
                        strName = Mid(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") + 1)
                        gcnOracle.Execute "Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(" & _
                        IIF(lngSys = 0, "NULL", lngSys) & "," & lngProgID & ",'基本','" & strName & "','" & strOwner & "','SELECT')"
                    End If
                Next
            End If
        Else
            gcnOracle.Execute "Update zlReports Set 发布时间=Sysdate Where ID=" & lngRPTID
        End If
    Else
        If lvwGroup.SelectedItem.Tag = 0 Then
            '更新报表组中各子报表的功能为各子报表的名称
            gcnOracle.Execute "Update zlRPTSubs A Set 功能=(Select 名称 From zlReports Where ID=A.报表ID) Where 组ID=" & lngRPTID
            gcnOracle.Execute "Update zlRPTGroups Set 程序ID=" & lngProgID & ",发布时间=Sysdate Where ID=" & lngRPTID
            gcnOracle.Execute "Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(" & lngProgID & "," & _
                " '" & lvwGroup.SelectedItem.Text & "','" & lvwGroup.SelectedItem.SubItems(GC_说明) & "'," & _
                IIF(lngSys = 0, "NULL", lngSys) & ",'zl9Report')"
            gcnOracle.Execute "Insert Into zlProgFuncs(系统,序号,功能,说明)" & _
                " Select " & IIF(lngSys = 0, "NULL", lngSys) & "," & lngProgID & ",名称,说明 From zlReports" & _
                " Where ID In (Select 报表ID From zlRPTSubs Where 组ID=" & Mid(lvwGroup.SelectedItem.Key, 2) & ")"
            
            strSQL = "Select A.报表ID,B.名称 From zlRPTSubs A,zlReports B Where A.组ID=[1] And A.报表ID=B.ID"
            Set rsSubRPT = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
            '循环读取各子报表的权限
            Do While Not rsSubRPT.EOF
                '分析该子报表的数据源访问对象
                strObject = ""
                strSQL = "Select 对象 From zlRPTDatas Where 对象 is Not NULL And 报表ID=[1]"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Val(rsSubRPT!报表id))
                If Not rsTmp.EOF Then
                    For i = 1 To rsTmp.RecordCount
                        For j = 0 To UBound(Split(rsTmp!对象, ","))
                            If InStr(strObject & ",", "," & Split(rsTmp!对象, ",")(j) & ",") = 0 Then
                                strObject = strObject & "," & Split(rsTmp!对象, ",")(j)
                            End If
                        Next
                        rsTmp.MoveNext
                    Next
                End If
                
                '分析该子报表的参数数据源访问对象
                strSQL = "Select B.对象 From zlRPTDatas A,zlRPTPars B Where A.ID=B.源ID And B.对象 is Not NULL And A.报表ID=[1]"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Val(rsSubRPT!报表id))
                If Not rsTmp.EOF Then
                    For i = 1 To rsTmp.RecordCount
                        For j = 0 To UBound(Split(rsTmp!对象, "|"))
                            strTmp = Split(rsTmp!对象, "|")(j)
                            For k = 0 To UBound(Split(strTmp, ","))
                                If InStr(strObject & ",", "," & Split(strTmp, ",")(k) & ",") = 0 Then
                                    strObject = strObject & "," & Split(strTmp, ",")(k)
                                End If
                            Next
                        Next
                        rsTmp.MoveNext
                    Next
                End If
                
                If strObject <> "" Then '该表格有可能不访问数据库
                    strObject = Mid(strObject, 2)
                    For i = 0 To UBound(Split(strObject, ","))
                        strOwner = Left(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") - 1)
                        If strOwner <> "SYS" And strOwner <> "ZLTOOLS" And strOwner <> "SYSTEM" Then
                            strName = Mid(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") + 1)
                            gcnOracle.Execute "Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(" & _
                            IIF(lngSys = 0, "NULL", lngSys) & "," & lngProgID & ",'" & rsSubRPT!名称 & "','" & strName & "','" & strOwner & "','SELECT')"
                        End If
                    Next
                End If
                
                rsSubRPT.MoveNext
            Loop
        Else
            gcnOracle.Execute "Update zlRPTGroups Set 发布时间=Sysdate Where ID=" & lngRPTID
        End If
    End If
    
    '3.填写菜单
    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_报表 Then
        gcnOracle.Execute "Insert Into zlMenus(组别,ID,上级ID,标题,快键,说明,系统,模块,短标题,图标)" & _
            " Values('缺省'," & lngNewMenu & "," & lngMenu & ",'" & lvwReport.SelectedItem.Text & "',NULL," & _
            "'" & lvwReport.SelectedItem.SubItems(RC_说明) & "'," & IIF(lngSys = 0, "NULL", lngSys) & "," & _
            lngProgID & ",'" & lvwReport.SelectedItem.Text & "',105)"
    Else
        gcnOracle.Execute "Insert Into zlMenus(组别,ID,上级ID,标题,快键,说明,系统,模块,短标题,图标)" & _
            " Values('缺省'," & lngNewMenu & "," & lngMenu & ",'" & lvwGroup.SelectedItem.Text & "',NULL," & _
            " '" & lvwGroup.SelectedItem.SubItems(GC_说明) & "'," & IIF(lngSys = 0, "NULL", lngSys) & "," & _
            lngProgID & ",'" & lvwGroup.SelectedItem.Text & "',105)"
    End If
    
    gcnOracle.CommitTrans
    
    Set grsReport = Nothing '清除缓存
    
    '4.更新界面
    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_报表 Then
        If lngSys = 0 Then
            If lvwReport.SelectedItem.SubItems(RC_种类) = "票据" Then
                lvwReport.SelectedItem.Icon = "BillPublish"
                lvwReport.SelectedItem.SmallIcon = "BillPublish"
            Else
                lvwReport.SelectedItem.Icon = "Publish"
                lvwReport.SelectedItem.SmallIcon = "Publish"
            End If
        Else
            lvwReport.SelectedItem.Icon = "PubFixed"
            lvwReport.SelectedItem.SmallIcon = "PubFixed"
        End If
        lvwReport.SelectedItem.Tag = lngProgID
        lvwReport.SelectedItem.SubItems(RC_发布时间) = Format(Currentdate, "yyyy-MM-dd")
        Call lvwReport_ItemClick(lvwReport.SelectedItem)
    Else
        lvwGroup.SelectedItem.Icon = IIF(lngSys = 0, 2, 4)
        lvwGroup.SelectedItem.SmallIcon = IIF(lngSys = 0, 2, 4)
        lvwGroup.SelectedItem.Tag = lngProgID
        lvwGroup.SelectedItem.SubItems(GC_发布时间) = Format(Currentdate, "yyyy-MM-dd")
        Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub ReportRevokeFromNavigator(Optional ByVal blnRevokeByProgram As Boolean = False)
'功能：取消当前报表(组)在导航台上的一个发布
'1:如果发布位置大于1,则让使用者选择取消发布的一个位置,删除zlMenus对应位置内容,完成
'2:如果只有一个发布位置,则将zlReport中的程序ID=NULL,删除zlPrograms中的发布模块,完成
    Dim rsTmp As ADODB.Recordset
    Dim objNode As Node, lngSys As Long
    Dim lngProgID As Long, lngMenu As Long, i As Integer
    
    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_报表 Then
        If lvwReport.SelectedItem Is Nothing Then MsgBox "当前没有报表可以取消发布！", vbInformation, App.Title: Exit Sub
        If lvwReport.SelectedItem.Tag = 0 Then MsgBox "当前报表没有发布到导航台菜单！", vbInformation, App.Title: Exit Sub
        If lvwReport.SelectedItem.Icon = "Fixed" Or lvwReport.SelectedItem.Icon = "PubFixed" Then
            MsgBox "该报表为系统固有的报表,操作不能继续！", vbInformation, App.Title: Exit Sub
        End If
    
        lngProgID = CLng(lvwReport.SelectedItem.Tag)
    Else
        If lvwGroup.SelectedItem Is Nothing Then MsgBox "当前没有报表组可以取消发布！", vbInformation, App.Title: Exit Sub
        If lvwGroup.SelectedItem.Tag = 0 Then MsgBox "当前报表组没有发布！", vbInformation, App.Title: Exit Sub
        If lvwGroup.SelectedItem.Icon = 3 Or lvwGroup.SelectedItem.Icon = 4 Then
            MsgBox "该报表为系统固有的报表组,操作不能继续！", vbInformation, App.Title: Exit Sub
        End If
        
        lngProgID = CLng(lvwGroup.SelectedItem.Tag)
    End If
    
    '1.分析当前发布位置
    Set rsTmp = GetMainTreeMenu(lngProgID)
    If rsTmp Is Nothing Then MsgBox "读取菜单体系时遇到意外错误,取消发布中断！", vbInformation, App.Title: Exit Sub
    
    rsTmp.Filter = "模块=" & lngProgID
    lngSys = cboSys.ItemData(cboSys.ListIndex)
    
    If rsTmp.EOF Then
        MsgBox "当前报表的发布处于不正常状态,这可能是数据不正确引起的！", vbInformation, App.Title
        On Error GoTo errH
        
        gcnOracle.BeginTrans
        If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_报表 Then
            gcnOracle.Execute "Update zlReports Set 功能=NULL,程序ID=NULL,发布时间=NULL Where ID=" & Mid(lvwReport.SelectedItem.Key, 2)
        Else
            gcnOracle.Execute "Update zlRPTGroups Set 程序ID=NULL,发布时间=NULL Where ID=" & lngProgID
            gcnOracle.Execute " Update zlRPTSubs A Set 功能=Null Where 组ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
        End If
        gcnOracle.Execute "Delete From zlMenus Where 模块=" & lngProgID & " And Nvl(系统,0)=" & lngSys
        gcnOracle.Execute "Delete From zlProgPrivs Where 序号=" & lngProgID & " And Nvl(系统,0)=" & lngSys
        gcnOracle.Execute "Delete From zlProgFuncs Where 序号=" & lngProgID & " And Nvl(系统,0)=" & lngSys
        gcnOracle.Execute "Delete From zlPrograms Where 序号=" & lngProgID & " And Nvl(系统,0)=" & lngSys
        gcnOracle.Execute "Delete From zlRoleGrant Where 序号=" & lngProgID & " And Nvl(系统,0)=" & lngSys
        
        gcnOracle.CommitTrans
        
        Set grsReport = Nothing '清除缓存
        
        If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_报表 Then
            If lvwReport.SelectedItem.SubItems(RC_种类) = "票据" Then
                lvwReport.SelectedItem.Icon = "Bill"
                lvwReport.SelectedItem.SmallIcon = "Bill"
            Else
                lvwReport.SelectedItem.Icon = "Report"
                lvwReport.SelectedItem.SmallIcon = "Report"
            End If
            lvwReport.SelectedItem.Tag = 0
            lvwReport.SelectedItem.SubItems(RC_发布时间) = ""
        Else
            lvwGroup.SelectedItem.Icon = 1
            lvwGroup.SelectedItem.SmallIcon = 1
            lvwGroup.SelectedItem.Tag = 0
            lvwGroup.SelectedItem.SubItems(GC_发布时间) = ""
        End If
    ElseIf rsTmp.RecordCount = 1 Then
        '只剩一个发布位置
        If Not blnRevokeByProgram Then
            If MsgBox("如果把该报表从导航台菜单中取消发布，其他用户不能再使用该报表。要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        End If
        On Error GoTo errH
        
        gcnOracle.BeginTrans
        
        If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_报表 Then
            gcnOracle.Execute "Update zlReports Set 功能=NULL,程序ID=NULL,发布时间=NULL Where ID=" & Mid(lvwReport.SelectedItem.Key, 2)
        Else
            gcnOracle.Execute "Update zlRPTGroups Set 程序ID=NULL,发布时间=NULL Where ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
            gcnOracle.Execute " Update zlRPTSubs A Set 功能=Null Where 组ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
        End If
        gcnOracle.Execute "Delete From zlMenus Where 模块=" & lngProgID & " And Nvl(系统,0)=" & lngSys
        gcnOracle.Execute "Delete From zlProgPrivs Where 序号=" & lngProgID & " And Nvl(系统,0)=" & lngSys
        gcnOracle.Execute "Delete From zlProgFuncs Where 序号=" & lngProgID & " And Nvl(系统,0)=" & lngSys
        gcnOracle.Execute "Delete From zlPrograms Where 序号=" & lngProgID & " And Nvl(系统,0)=" & lngSys
        gcnOracle.Execute "Delete From zlRoleGrant Where 序号=" & lngProgID & " And Nvl(系统,0)=" & lngSys
        
        gcnOracle.CommitTrans
        
        Set grsReport = Nothing '清除缓存
        
        If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_报表 Then
            If lvwReport.SelectedItem.SubItems(RC_种类) = "票据" Then
                lvwReport.SelectedItem.Icon = "Bill"
                lvwReport.SelectedItem.SmallIcon = "Bill"
            Else
                lvwReport.SelectedItem.Icon = "Report"
                lvwReport.SelectedItem.SmallIcon = "Report"
            End If
            lvwReport.SelectedItem.Tag = 0
            lvwReport.SelectedItem.SubItems(RC_发布时间) = ""
        Else
            lvwGroup.SelectedItem.Icon = 1
            lvwGroup.SelectedItem.SmallIcon = 1
            lvwGroup.SelectedItem.Tag = 0
            lvwGroup.SelectedItem.SubItems(GC_发布时间) = ""
        End If
    Else
        '还有多个发布位置,选择性取消
        rsTmp.Filter = 0
        
        Load frmSelTree
        frmSelTree.Caption = "取消发布 - 导航台菜单位置"
        With frmSelTree.tvw
            .Nodes.Clear
            For i = 1 To rsTmp.RecordCount
                If rsTmp!Flag = 0 Then
                    Set objNode = .Nodes.Add(, , "_" & rsTmp!ID, rsTmp!标题, "Root")
                    objNode.Tag = "请在本系统下选择一个要取消发布的报表或组！"
                Else
                    If rsTmp!Flag <> 999 And rsTmp!Flag <> 888 Then
                        Set objNode = .Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!ID, rsTmp!标题, "Path")
                        objNode.Tag = "请在菜单上选择一个要取消发布的报表或组！"
                    Else
                        Set objNode = .Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!ID, rsTmp!标题, IIF(rsTmp!Flag = 999, "GroupNode", "ReportNode"))
                        objNode.ForeColor = vbBlue
                        If .SelectedItem Is Nothing Then
                            objNode.Selected = True
                        ElseIf .SelectedItem.Index = 1 Then
                            objNode.Selected = True
                        End If
                    End If
                End If
                objNode.Expanded = True
                
                '标记有报表(组)的路径
                If rsTmp!Flag = 999 Or rsTmp!Flag = 888 Then
                    objNode.SelectedImage = objNode.Image
                    Do While Not objNode.Parent Is Nothing
                        Set objNode = objNode.Parent
                        objNode.SelectedImage = objNode.Image
                    Loop
                End If
                
                rsTmp.MoveNext
            Next
            
            '删除没有报表(组)的路径
            For i = .Nodes.count To 1 Step -1
                If .Nodes(i).SelectedImage = "" Then
                    .Nodes.Remove i
                End If
            Next
        End With
        frmSelTree.Show 1, Me
        If Not gblnOK Then Exit Sub
        lngMenu = CLng(Mid(frmSelTree.tvw.SelectedItem.Key, 2)) '报表菜单ID
        Unload frmSelTree
        
        On Error GoTo errH
        
        gcnOracle.BeginTrans
        If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_报表 Then
            gcnOracle.Execute "Update zlReports Set 发布时间=Sysdate Where ID=" & Mid(lvwReport.SelectedItem.Key, 2)
        Else
            gcnOracle.Execute "Update zlRPTGroups Set 发布时间=Sysdate Where ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
            gcnOracle.Execute "Update zlRPTSubs A Set 功能=Null Where 组ID=" & Mid(lvwGroup.SelectedItem.Key, 2)
        End If
        '只需删除菜单内容
        gcnOracle.Execute "Delete From zlMenus Where ID=" & lngMenu & " And Nvl(系统,0)=" & lngSys
        gcnOracle.CommitTrans
        
        Set grsReport = Nothing '清除缓存
        
        If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_报表 Then
            lvwReport.SelectedItem.SubItems(RC_发布时间) = Format(Currentdate, "yyyy-MM-dd")
        Else
            lvwGroup.SelectedItem.SubItems(GC_发布时间) = Format(Currentdate, "yyyy-MM-dd")
        End If
    End If
    
    If lvwGroup.SelectedItem.Key = "_-1" Or mcsActive = CS_报表 Then
        Call lvwReport_ItemClick(lvwReport.SelectedItem)
    Else
        Call LvwGroup_ItemClick(lvwGroup.SelectedItem)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Public Sub ReportGrantToNavigatorAgain(ByVal objItem As ListItem)
'功能：根据报表组中子表的增删情况，重新更新指定“报表组”的发布授权情况
    Dim rsTmp As New ADODB.Recordset
    Dim rsSubRPT As New ADODB.Recordset
    Dim strTmp As String, strOwner As String, strName As String
    Dim strSQL As String, i As Integer, j As Integer, k As Integer
    Dim strObject As String, lngSys As Long, lngProgID As Long
    
    lngProgID = Val(objItem.Tag)
    If lngProgID = 0 Then Exit Sub
    lngSys = cboSys.ItemData(cboSys.ListIndex)
    
    On Error GoTo errH
    
    gcnOracle.BeginTrans
    gcnOracle.Execute "Update zlRPTGroups Set 程序ID=" & lngProgID & ",发布时间=Sysdate Where ID=" & Mid(objItem.Key, 2)
    gcnOracle.Execute "Delete zlProgFuncs Where 序号=" & lngProgID & " And Nvl(系统,0)=" & lngSys
    gcnOracle.Execute "Delete zlProgPrivs Where 序号=" & lngProgID & " And Nvl(系统,0)=" & lngSys
    gcnOracle.Execute "Insert Into zlProgFuncs(系统,序号,功能,说明) Select " & IIF(lngSys = 0, "NULL", lngSys) & "," & _
        lngProgID & ",名称,说明 From zlReports Where ID In (Select 报表ID From zlRPTSubs Where 组ID=" & Mid(objItem.Key, 2) & ")"
            
    strSQL = "Select A.报表ID,B.名称 From zlRPTSubs A,zlReports B Where A.组ID=[1] And A.报表ID=B.ID"
    Set rsSubRPT = OpenSQLRecord(strSQL, Me.Caption, Val(Mid(objItem.Key, 2)))
    '循环读取各子报表的权限
    Do While Not rsSubRPT.EOF
        '分析该子报表的数据源访问对象
        strObject = ""
        strSQL = "Select 对象 From zlRPTDatas Where 对象 is Not NULL And 报表ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Val(rsSubRPT!报表id))
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                For j = 0 To UBound(Split(rsTmp!对象, ","))
                    If InStr(strObject & ",", "," & Split(rsTmp!对象, ",")(j) & ",") = 0 Then
                        strObject = strObject & "," & Split(rsTmp!对象, ",")(j)
                    End If
                Next
                rsTmp.MoveNext
            Next
        End If
        
        '分析该子报表的参数数据源访问对象
        strSQL = "Select B.对象 From zlRPTDatas A,zlRPTPars B Where A.ID=B.源ID And B.对象 is Not NULL And A.报表ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Val(rsSubRPT!报表id))
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                For j = 0 To UBound(Split(rsTmp!对象, "|"))
                    strTmp = Split(rsTmp!对象, "|")(j)
                    For k = 0 To UBound(Split(strTmp, ","))
                        If InStr(strObject & ",", "," & Split(strTmp, ",")(k) & ",") = 0 Then
                            strObject = strObject & "," & Split(strTmp, ",")(k)
                        End If
                    Next
                Next
                rsTmp.MoveNext
            Next
        End If
        
        If strObject <> "" Then '该表格有可能不访问数据库
            strObject = Mid(strObject, 2)
            For i = 0 To UBound(Split(strObject, ","))
                strOwner = Left(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") - 1)
                If strOwner <> "SYS" And strOwner <> "ZLTOOLS" And strOwner <> "SYSTEM" Then
                    strName = Mid(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") + 1)
                    gcnOracle.Execute "Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(" & _
                        IIF(lngSys = 0, "NULL", lngSys) & "," & lngProgID & ",'" & rsSubRPT!名称 & "'," & _
                        "'" & strName & "','" & strOwner & "','SELECT')"
                End If
            Next
        End If
        rsSubRPT.MoveNext
    Loop
    gcnOracle.CommitTrans
    
    Set grsReport = Nothing '清除缓存
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub ReportGrantToModule()
'功能：发布当前报表到模块,可能不是第一次
    Dim rsTmp As ADODB.Recordset
    Dim objNode As Node, strSQL As String
    Dim strObject As String, strOwner As String, strName As String
    Dim lngRPTID As Long, lngSys As Long, lngProgID As Long
    Dim i As Integer, j As Integer, k As Integer
    Dim strFunc As String, blnTran As Boolean
    
    '当前有具体报表组选择时，不支持报表组发布到模块
    If Val(Mid(lvwGroup.SelectedItem.Key, 2)) <> -1 And mcsActive = CS_报表组 Then Exit Sub
    
    If lvwReport.SelectedItem Is Nothing Then
        MsgBox "当前没有报表可以发布！", vbInformation, App.Title: Exit Sub
    End If
    If CheckPass(CLng(Mid(lvwReport.SelectedItem.Key, 2))) = False Then
        MsgBox "报表数据错误，不能发布该报表！", vbInformation, App.Title: Exit Sub
    End If
    If Not CheckReportPriv(Val(Mid(lvwReport.SelectedItem.Key, 2))) Then
        MsgBox "你没有权限查询该报表某些数据源中的对象,操作不能继续！", vbInformation, App.Title
        Exit Sub
    End If
    lngRPTID = CLng(Mid(lvwReport.SelectedItem.Key, 2))
    
    On Error GoTo errH
    
    '1.选择一个菜单模块位置
    '----------------------------------------------------
    Set rsTmp = GetModuleTreeMenu(lngRPTID)
    If rsTmp Is Nothing Then
        MsgBox "读取模块菜单体系时遇到意外错误,报表发布中断！", vbInformation, App.Title: Exit Sub
    End If
    Load frmSelTree
    frmSelTree.Caption = "发布报表到模块 - 模块位置选择"
    With frmSelTree.tvw
        .Nodes.Clear
        For i = 1 To rsTmp.RecordCount
            If IsNull(rsTmp!上级ID) Then
                Set objNode = .Nodes.Add(, , "_" & rsTmp!ID, rsTmp!标题)
            Else
                Set objNode = .Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!ID, rsTmp!标题)
            End If
            If Left(rsTmp!ID, 1) = "S" Then 'System
                objNode.Image = "Root"
                objNode.Tag = "请选择本系统中菜单下的模块位置。"
            ElseIf Left(rsTmp!ID, 1) = "T" Then 'MenuTree
                objNode.Image = "Path"
                objNode.Tag = "请选择本系统中菜单下的模块位置。"
            ElseIf Left(rsTmp!ID, 1) = "M" Then 'Module
                objNode.Image = "App"
            ElseIf Left(rsTmp!ID, 1) = "R" Then 'Report
                objNode.Image = "ReportNode"
                objNode.ForeColor = vbBlue
                objNode.Tag = "这是已发布的报表,选择其他菜单下的模块位置。"
                objNode.Parent.Tag = "报表不能重复发布到同一个模块,请选择其他模块。"
            End If
            objNode.Expanded = True
            
            '标记有下级模块的菜单(用SQL较慢)
            If Left(rsTmp!ID, 1) = "M" Then
                If objNode.Parent.SelectedImage = "" Then
                    Do While Not objNode.Parent Is Nothing
                        Set objNode = objNode.Parent
                        objNode.SelectedImage = objNode.Image
                    Loop
                End If
            End If
            
            rsTmp.MoveNext
        Next
        
        '删除无下级模块的空菜单
        For i = .Nodes.count To 1 Step -1
            If .Nodes(i).SelectedImage = "" And Mid(.Nodes(i).Key, 2, 1) = "T" Then
                .Nodes.Remove i
            End If
        Next
        
        If .Nodes.count > 0 Then .Nodes(1).Selected = True
    End With
    frmSelTree.Show 1, Me
    If Not gblnOK Then Exit Sub
    rsTmp.Filter = "ID='" & Mid(frmSelTree.tvw.SelectedItem.Key, 2) & "'"
    If rsTmp.EOF Then Exit Sub
    lngSys = rsTmp!系统: lngProgID = rsTmp!程序ID
    strFunc = lvwReport.SelectedItem.Text
    Unload frmSelTree
        
    '数据重复检查
    strSQL = _
        " Select 功能 From zlRPTPuts Where 报表ID=[1] And 系统=[2] And 程序ID=[3]" & _
        " Union ALL " & _
        " Select 功能 From zlProgFuncs Where 系统=[2] And 序号=[3] And 功能=[4]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID, lngSys, lngProgID, strFunc)
    If Not rsTmp.EOF Then
        MsgBox "报表发布位置或发布功能重复，数据库中的数据可能不正确。", vbInformation, App.Title
        Exit Sub
    End If
    
    '2.授权权限分析
    '----------------------------------------------------
    '分析该报表的数据源访问对象
    strSQL = "Select 对象 From zlRPTDatas Where 对象 is Not NULL And 报表ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            For j = 0 To UBound(Split(rsTmp!对象, ","))
                If InStr(strObject & ",", "," & Split(rsTmp!对象, ",")(j) & ",") = 0 Then
                    strObject = strObject & "," & Split(rsTmp!对象, ",")(j)
                End If
            Next
            rsTmp.MoveNext
        Next
    End If
    
    '分析该报表的参数数据源访问对象
    strSQL = "Select B.对象 From zlRPTDatas A,zlRPTPars B Where A.ID=B.源ID And B.对象 is Not NULL And A.报表ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            For j = 0 To UBound(Split(rsTmp!对象, "|"))
                strName = Split(rsTmp!对象, "|")(j)
                For k = 0 To UBound(Split(strName, ","))
                    If InStr(strObject & ",", "," & Split(strName, ",")(k) & ",") = 0 Then
                        strObject = strObject & "," & Split(strName, ",")(k)
                    End If
                Next
            Next
            rsTmp.MoveNext
        Next
    End If
    If strObject <> "" Then strObject = Mid(strObject, 2)
        
    '3.填写程序、权限
    '----------------------------------------------------
    gcnOracle.BeginTrans: blnTran = True
    
    gcnOracle.Execute "Update zlReports Set 发布时间=Sysdate Where ID=" & lngRPTID
    gcnOracle.Execute "Insert Into zlRPTPuts(报表ID,系统,程序ID,功能) Values(" & _
        lngRPTID & "," & lngSys & "," & lngProgID & ",'" & strFunc & "')"
    gcnOracle.Execute "Insert Into zlProgFuncs(系统,序号,功能,说明) Values(" & _
        lngSys & "," & lngProgID & ",'" & strFunc & "','" & lvwReport.SelectedItem.SubItems(RC_说明) & "')"
    If strObject <> "" Then '该表格有可能不访问数据库
        For i = 0 To UBound(Split(strObject, ","))
            strOwner = Left(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") - 1)
            If strOwner <> "SYS" And strOwner <> "ZLTOOLS" And strOwner <> "SYSTEM" Then
                strName = Mid(Split(strObject, ",")(i), InStr(Split(strObject, ",")(i), ".") + 1)
                gcnOracle.Execute "Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(" & _
                lngSys & "," & lngProgID & ",'" & strFunc & "','" & strName & "','" & strOwner & "','SELECT')"
            End If
        Next
    End If
    
    gcnOracle.CommitTrans: blnTran = False
    
    Set grsReport = Nothing '清除缓存
    
    '4.更新界面
    If cboSys.ItemData(cboSys.ListIndex) = 0 Then
        If lvwReport.SelectedItem.SubItems(RC_种类) = "票据" Then
            lvwReport.SelectedItem.Icon = "BillPublish"
            lvwReport.SelectedItem.SmallIcon = "BillPublish"
        Else
            lvwReport.SelectedItem.Icon = "Publish"
            lvwReport.SelectedItem.SmallIcon = "Publish"
        End If
    Else
        lvwReport.SelectedItem.Icon = "PubFixed"
        lvwReport.SelectedItem.SmallIcon = "PubFixed"
    End If
    lvwReport.SelectedItem.SubItems(RC_发布时间) = Format(Currentdate, "yyyy-MM-dd")
    Call lvwReport_ItemClick(lvwReport.SelectedItem)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub ReportRevokeFromModule()
'功能：取消当前报表在模块上的一个发布
'1:如果发布位置大于1,则让使用者选择取消发布的一个位置
'2:如果只有一个发布位置,则直接提示处理
    Dim rsTmp As ADODB.Recordset, strFunc As String
    Dim objNode As Node, blnTran As Boolean
    Dim lngRPTID As Long, lngSys As Long, lngProgID As Long
    Dim strSQL As String, i As Integer
    
    If lvwReport.SelectedItem Is Nothing Then
        MsgBox "当前没有报表可以取消发布！", vbInformation, App.Title: Exit Sub
    End If
    lngRPTID = CLng(Mid(lvwReport.SelectedItem.Key, 2))
        
    On Error GoTo errH
    
    '1.分析当前发布位置
    strSQL = "Select 系统,程序ID,功能 From zlRPTPuts Where 报表ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngRPTID)
    If rsTmp.EOF Then
        MsgBox "当前报表没有发布到模块中。", vbInformation, App.Title: Exit Sub
    ElseIf rsTmp.RecordCount = 1 Then
        '只剩一个发布位置
        If MsgBox("如果把报表从该模块菜单中取消发布，其他用户不能再使用该报表。要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        
        lngSys = rsTmp!系统: lngProgID = rsTmp!程序ID: strFunc = rsTmp!功能
        
        gcnOracle.BeginTrans: blnTran = True

        gcnOracle.Execute "Update zlReports Set 发布时间=NULL Where 程序ID Is Null And ID=" & lngRPTID
        gcnOracle.Execute "Delete From zlRPTPuts Where 报表ID=" & lngRPTID & " And 系统=" & lngSys & " And 程序ID=" & lngProgID
        gcnOracle.Execute "Delete From zlProgPrivs Where 系统=" & lngSys & " And 序号=" & lngProgID & " And 功能='" & strFunc & "'"
        gcnOracle.Execute "Delete From zlProgFuncs Where 系统=" & lngSys & " And 序号=" & lngProgID & " And 功能='" & strFunc & "'"
        gcnOracle.Execute "Delete From zlRoleGrant Where 系统=" & lngSys & " And 序号=" & lngProgID & " And 功能='" & strFunc & "'"
        
        gcnOracle.CommitTrans: blnTran = False
        
        Set grsReport = Nothing '清除缓存
        
        If Val(lvwReport.SelectedItem.Tag) = 0 Then
            If cboSys.ItemData(cboSys.ListIndex) = 0 Then
                If lvwReport.SelectedItem.SubItems(RC_种类) = "票据" Then
                    lvwReport.SelectedItem.Icon = "Bill"
                    lvwReport.SelectedItem.SmallIcon = "Bill"
                Else
                    lvwReport.SelectedItem.Icon = "Report"
                    lvwReport.SelectedItem.SmallIcon = "Report"
                End If
            Else
                lvwReport.SelectedItem.Icon = "Fixed"
                lvwReport.SelectedItem.SmallIcon = "Fixed"
            End If
            lvwReport.SelectedItem.SubItems(RC_发布时间) = ""
        End If
    Else
        '还有多个发布位置,选择性取消
        Set rsTmp = GetModuleTreeMenu(lngRPTID)
        If rsTmp Is Nothing Then
            MsgBox "读取模块菜单体系时遇到意外错误,报表发布中断！", vbInformation, App.Title: Exit Sub
        End If
        Load frmSelTree
        frmSelTree.Caption = "取消发布 - 模块菜单位置"
        With frmSelTree.tvw
            .Nodes.Clear
            For i = 1 To rsTmp.RecordCount
                If IsNull(rsTmp!上级ID) Then
                    Set objNode = .Nodes.Add(, , "_" & rsTmp!ID, rsTmp!标题)
                Else
                    Set objNode = .Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!ID, rsTmp!标题)
                End If
                If Left(rsTmp!ID, 1) = "S" Then 'System
                    objNode.Image = "Root"
                    objNode.Tag = "请选择要取消发布的报表。"
                ElseIf Left(rsTmp!ID, 1) = "T" Then 'MenuTree
                    objNode.Image = "Path"
                    objNode.Tag = "请选择要取消发布的报表。"
                ElseIf Left(rsTmp!ID, 1) = "M" Then 'Module
                    objNode.Image = "App"
                    objNode.Tag = "请选择要取消发布的报表。"
                ElseIf Left(rsTmp!ID, 1) = "R" Then 'Report
                    objNode.Image = "ReportNode"
                    objNode.ForeColor = vbBlue
                End If
                objNode.Expanded = True
                
                '标记有发布报表的上级
                If Left(rsTmp!ID, 1) = "R" Then
                    objNode.SelectedImage = objNode.Image
                    If objNode.Parent.SelectedImage = "" Then
                        Do While Not objNode.Parent Is Nothing
                            Set objNode = objNode.Parent
                            objNode.SelectedImage = objNode.Image
                        Loop
                    End If
                End If
                
                rsTmp.MoveNext
            Next
            
            '删除无发布报表的路径
            For i = .Nodes.count To 1 Step -1
                If .Nodes(i).SelectedImage = "" Then
                    .Nodes.Remove i
                End If
            Next
            
            If .Nodes.count > 0 Then .Nodes(1).Selected = True
        End With
        frmSelTree.Show 1, Me
        If Not gblnOK Then Exit Sub
        rsTmp.Filter = "ID='" & Mid(frmSelTree.tvw.SelectedItem.Key, 2) & "'"
        If rsTmp.EOF Then Exit Sub
        lngSys = rsTmp!系统: lngProgID = rsTmp!程序ID: strFunc = rsTmp!功能
        Unload frmSelTree
        
        gcnOracle.BeginTrans: blnTran = True

        gcnOracle.Execute "Delete From zlRPTPuts Where 报表ID=" & lngRPTID & " And 系统=" & lngSys & " And 程序ID=" & lngProgID
        gcnOracle.Execute "Delete From zlProgPrivs Where 系统=" & lngSys & " And 序号=" & lngProgID & " And 功能='" & strFunc & "'"
        gcnOracle.Execute "Delete From zlProgFuncs Where 系统=" & lngSys & " And 序号=" & lngProgID & " And 功能='" & strFunc & "'"
        gcnOracle.Execute "Delete From zlRoleGrant Where 系统=" & lngSys & " And 序号=" & lngProgID & " And 功能='" & strFunc & "'"
        
        gcnOracle.CommitTrans: blnTran = False
        
        Set grsReport = Nothing '清除缓存
    End If
    
    Call lvwReport_ItemClick(lvwReport.SelectedItem)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub SetFuncEnabled(ByVal blnReport As Boolean)
'功能：设置菜单及按钮可用状态
'参数：blnReport=当前活动列表是否报表列表
    Dim blnFree As Boolean
    
    On Error Resume Next
    
    '指定系统,只能执行修改名称、执行及设计的功能
    blnFree = cboSys.ItemData(cboSys.ListIndex) = 0
        
    '设置功能可见性
    '------------------------------------------------------------
    mnuEdit_Group_Add.Visible = Not blnReport And mblnGrant
    mnuEdit_Group_Delete.Visible = Not blnReport And mblnGrant
    mnuEdit_Group_Modify.Visible = Not blnReport
    mnuEdit_Group_Setup.Visible = Not blnReport
    mnuEdit_Group_.Visible = Not blnReport
    tbr.Buttons("GroupAdd").Visible = Not blnReport And mblnGrant
    tbr.Buttons("GroupDel").Visible = Not blnReport And mblnGrant
    tbr.Buttons("GroupModify").Visible = Not blnReport
    tbr.Buttons("Group_").Visible = Not blnReport
    
    mnuEdit_Add.Visible = blnReport And mblnGrant
    mnuEdit_Del.Visible = blnReport And mblnGrant
    mnuEdit_Modi.Visible = blnReport
    mnuEdit_Report_.Visible = blnReport
    tbr.Buttons("Add").Visible = blnReport And mblnGrant
    tbr.Buttons("Del").Visible = blnReport And mblnGrant
    tbr.Buttons("Modi").Visible = blnReport
    tbr.Buttons("Report_").Visible = blnReport
    
    '设置功能可用性
    '------------------------------------------------------------
    If Me.ActiveControl Is lvwReport Or lvwGroup.SelectedItem.Key = "_-1" Then
        mnuFile_Report.Enabled = Not lvwReport.SelectedItem Is Nothing
    Else
        mnuFile_Report.Enabled = True
    End If
    tbr.Buttons("Report").Enabled = mnuFile_Report.Enabled
    
    If Not blnFree Then
        mnuFile_Imp.Enabled = True
        mnuFile_ImpAll.Enabled = True
    Else
        mnuFile_Imp.Enabled = mblnGrant
        mnuFile_ImpAll.Enabled = mblnGrant
    End If
    mnuFile_Exp.Enabled = Not lvwReport.SelectedItem Is Nothing
    mnuEdit_Add.Enabled = blnFree
    mnuEdit_Modi.Enabled = Not lvwReport.SelectedItem Is Nothing
    mnuEdit_Del.Enabled = blnFree And Not lvwReport.SelectedItem Is Nothing
    tbr.Buttons("Add").Enabled = mnuEdit_Add.Enabled
    tbr.Buttons("Modi").Enabled = mnuEdit_Modi.Enabled
    tbr.Buttons("Del").Enabled = mnuEdit_Del.Enabled
    
    mnuEdit_Group_Add.Enabled = blnFree
    mnuEdit_Group_Modify.Enabled = lvwGroup.SelectedItem.Key <> "_-1"
    mnuEdit_Group_Delete.Enabled = blnFree And lvwGroup.SelectedItem.Key <> "_-1"
    tbr.Buttons("GroupAdd").Enabled = mnuEdit_Group_Add.Enabled
    tbr.Buttons("GroupModify").Enabled = mnuEdit_Group_Modify.Enabled
    tbr.Buttons("GroupDel").Enabled = mnuEdit_Group_Delete.Enabled
    
    mnuEdit_Group_Setup.Enabled = blnFree And lvwGroup.SelectedItem.Key <> "_-1"
    
    mnuEdit_Design.Enabled = Not lvwReport.SelectedItem Is Nothing
    tbr.Buttons("Design").Enabled = mnuEdit_Design.Enabled
    
    mnuEdit_Guide.Enabled = blnFree
    tbr.Buttons("Guide").Enabled = mnuEdit_Guide.Enabled
        
    '仅共享报表可以发布,报表组不能发布到模块
    mnuEdit_Publish.Enabled = blnFree
    mnuEdit_unPub.Enabled = blnFree
    '判断发布报表组和发布报表选项的显示状态
    If mcsActive = CS_报表 Then
        If lvwGroup.SelectedItem.Key = "_-1" Then
            mnuEdit_Group_Publish.Visible = False
            mnuEdit_Group_unPub.Visible = False
        End If
        mnuEdit_Publish.Visible = True
        mnuEdit_unPub.Visible = True
    Else
        mnuEdit_Group_Publish.Visible = True And blnFree
        mnuEdit_Group_unPub.Visible = True And blnFree
        mnuEdit_Publish.Visible = False
        mnuEdit_unPub.Visible = False
    End If
    mnuPopPublish_Group.Visible = mnuEdit_Group_Publish.Visible
    mnuPopUnpub_Group.Visible = mnuEdit_Group_unPub.Visible
    mnuPopUnpub_ReportMain.Visible = mnuEdit_unPub_Module.Enabled
    mnuPopUnpub_ReportModule.Visible = mnuEdit_unPub_Module.Enabled
    mnuPopPublish_ReportMain.Visible = mnuEdit_Publish_Module.Enabled
    mnuPopPublish_ReportModule.Visible = mnuEdit_Publish_Module.Enabled
    If lvwGroup.SelectedItem.Key <> "_-1" Or mcsActive = CS_报表 Then
        '发布具体报表组到导航台
        mnuEdit_Publish_Main.Enabled = blnFree
        mnuEdit_unPub_Main.Enabled = blnFree
        '发布具体报表组到模块
        mnuEdit_Publish_Module.Enabled = IIF(mcsActive = CS_报表组, False, True)
        mnuEdit_unPub_Module = IIF(mcsActive = CS_报表组, False, True)
    ElseIf lvwGroup.SelectedItem.Key = "_-1" And mcsActive = CS_报表组 Then
        mnuEdit_Group_Publish.Visible = False
        mnuEdit_Group_unPub.Visible = False
        mnuEdit_Publish.Visible = False
        mnuEdit_unPub.Visible = False
    Else
        '发布具体报表到导航台
        mnuEdit_Publish_Main.Enabled = blnFree And Not lvwReport.SelectedItem Is Nothing
        mnuEdit_unPub_Main.Enabled = blnFree And Not lvwReport.SelectedItem Is Nothing
        
        '发布具体报表组到模块
        mnuEdit_Publish_Module.Enabled = mblnModule And blnFree And Not lvwReport.SelectedItem Is Nothing
        mnuEdit_unPub_Module = mblnModule And blnFree And Not lvwReport.SelectedItem Is Nothing
    End If
    If Not mnuEdit_Publish_Main.Enabled And Not mnuEdit_Publish_Module.Enabled Then
        mnuEdit_Publish.Enabled = False
    End If
    If Not mnuEdit_unPub_Main.Enabled And Not mnuEdit_unPub_Module.Enabled Then
        mnuEdit_unPub.Enabled = False
    End If
    tbr.Buttons("Publish").Enabled = (mnuEdit_Group_Publish.Visible Or mnuEdit_Publish.Visible) And blnFree
    tbr.Buttons("unPub").Enabled = (mnuEdit_Group_unPub.Visible Or mnuEdit_Publish.Visible) And blnFree
    '标题及视图
    '------------------------------------------------------------
    If blnReport Or lvwGroup.SelectedItem.Key = "_-1" Then
        mnuFile_Report.Caption = "执行报表(&E)"
        tbr.Buttons("Report").ToolTipText = "执行报表"
        tbr.Buttons("Report").Image = 1
    Else
        mnuFile_Report.Caption = "执行报表组(&E)"
        tbr.Buttons("Report").ToolTipText = "执行报表组"
        tbr.Buttons("Report").Image = 15
    End If
    
    mnuEdit_Del.Caption = IIF(lvwGroup.SelectedItem.Key <> "_-1", "移除", "删除")
    Me.tbr.Buttons("Del").Caption = mnuEdit_Del.Caption
    Me.tbr.Buttons("Del").Tag = mnuEdit_Del.Caption
    Me.tbr.Buttons("Del").ToolTipText = mnuEdit_Del.Caption & "当前报表"
    mnuEdit_Del.Caption = mnuEdit_Del.Caption & "报表(&D)"
    
    mnuView_View(0).Checked = False
    mnuView_View(1).Checked = False
    mnuView_View(2).Checked = False
    mnuView_View(3).Checked = False
    If blnReport Then
        mnuView_View(lvwReport.View).Checked = True
    Else
        mnuView_View(lvwGroup.View).Checked = True
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

Private Sub tbrCheck_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "Check" Then
        Call CheckSQLPlanEx
    End If
End Sub

Private Sub CheckSQLPlanEx()
'功能：检查当前列表中的报表执行计划是否存在性能问题
    Dim i As Long, objReport As Report, objData As RPTData
    Dim strSQLCheck As String, strErr As String, strFields As String
    Dim strMsg As String, objPar As RPTPar, strSQL As String
    Dim lngCount As Long
    
    If MsgBox("当前目录一共" & lvwReport.ListItems.count & "张报表，即将对这些报表(及参数)数据源中的SQL解析执行计划，" & _
         "然后检查执行计划是否存在以下情况：" & vbCrLf & _
         "    1.大表或中型表的全表扫描;" & vbCrLf & _
         "    2.大表或中型表的索引全扫描或跳跃式索引扫描;" & vbCrLf & _
         "    3.大表上引用基础表（非大表）的外键索引（例：病人医嘱记录_IX_诊疗项目ID）;" & vbCrLf & _
         "    其中大表是指zlBakTables ZlBigTables中定义的表;" & vbCrLf & _
         "    中型表是指收集统计信息后记录行数缺省在3千到1百万之间的表 (在设计界面的执行计划查看中可重新定义);" & vbCrLf & vbCrLf & _
         "此过程可能会花费几分钟的时间，你确定要继续吗？" _
         , vbQuestion + vbOKCancel + vbDefaultButton1, "性能检查") = vbCancel Then Exit Sub
    
    If lvwReport.ColumnHeaders(RC_性能问题数据源 + 1).Width = 0 Then lvwReport.ColumnHeaders(RC_性能问题数据源 + 1).Width = 3440

    For i = 1 To lvwReport.ListItems.count
        Set objReport = ReadReport(Val(Mid(lvwReport.ListItems(i).Key, 2)), , True)
        strMsg = ""
        For Each objData In objReport.Datas
            With objData
                '先检查数据源的SQL
                strSQLCheck = ""
                strFields = ""
                strSQL = RemoveNote(.SQL)
                strSQL = TrimChar(strSQL)
                strSQL = Replace(strSQL, "[系统]", cboSys.ItemData(cboSys.ListIndex))
                If GetParCount(strSQL) = 0 Then
                    strFields = CheckSQL(strSQL, strErr, , strSQLCheck, , objReport.Datas, .数据连接编号)
                Else
                    strFields = CheckSQL(strSQL, strErr, ReplaceParSysNo(.Pars, cboSys.ItemData(cboSys.ListIndex)) _
                        , strSQLCheck, , objReport.Datas, .数据连接编号)
                End If
                If strFields <> "" Then
                    If strSQLCheck <> "" Then
                        If CheckSQLPlan(strSQLCheck, , .数据连接编号) = True Then
                            strMsg = strMsg & "," & .名称
                        End If
                    End If
                End If
                '再检查参数明细和分类SQL
                For Each objPar In .Pars
                    '排除已经检查过的
                    If objPar.分类SQL <> "" And InStr(strMsg, "(" & objPar.名称 & ")[分类]") = 0 Then
                        strSQLCheck = ""
                        strFields = ""
                        strSQL = RemoveNote(objPar.分类SQL)
                        strSQL = TrimChar(strSQL)
                        strSQL = Replace(strSQL, "[系统]", cboSys.ItemData(cboSys.ListIndex))
                        Call CheckParsRela(strSQL, objReport.Datas, objPar.名称, True)
                        strFields = CheckSQL(strSQL, strErr, , strSQLCheck, , objReport.Datas, .数据连接编号)
                        If strFields <> "" Then
                            If strSQLCheck <> "" Then
                                If CheckSQLPlan(strSQLCheck, , .数据连接编号) = True Then
                                    strMsg = strMsg & "," & .名称 & "(" & objPar.名称 & ")[分类]"
                                End If
                            End If
                        End If
                    End If
                    
                    If objPar.明细SQL <> "" And InStr(strMsg, "(" & objPar.名称 & ")[明细]") = 0 Then
                        strSQLCheck = ""
                        strFields = ""
                        strSQL = RemoveNote(objPar.明细SQL)
                        strSQL = TrimChar(strSQL)
                        strSQL = Replace(strSQL, "[系统]", cboSys.ItemData(cboSys.ListIndex))
                        Call CheckParsRela(strSQL, objReport.Datas, objPar.名称, True)
                        strFields = CheckSQL(strSQL, strErr, , strSQLCheck, , , .数据连接编号)
                        If strFields <> "" Then
                            If strSQLCheck <> "" Then
                                If CheckSQLPlan(strSQLCheck, , .数据连接编号) = True Then
                                    strMsg = strMsg & "," & .名称 & "(" & objPar.名称 & ")[明细]"
                                End If
                            End If
                        End If
                    End If
                Next
            End With
        Next
        strMsg = Mid(strMsg, 2)
        lvwReport.ListItems(i).SubItems(RC_性能问题数据源) = strMsg
        If strMsg <> "" Then lngCount = lngCount + 1
        ShowFlash "正在检查报表数据源SQL存在的性能问题,请稍候 ...", i / lvwReport.ListItems.count
    Next
    ShowFlash
    If lngCount > 0 Then
        MsgBox "一共对" & lvwReport.ListItems.count & "张报表进行了性能检查，其中" & lngCount & "张报表(及参数)的数据源可能存在性能问题，详见""性能问题数据源""列的信息。" & vbCrLf & vbCrLf & _
            "请在报表设计界面查看详细的执行计划，并进行SQL性能优化。", vbInformation, "性能检查结果"
    End If
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0: txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim lngKey As Long, intActive As Integer
    Dim strSQL As String
    If KeyAscii = vbKeyReturn Then
        If txtFind.Text = "" Then Exit Sub
        If mstrFindValue <> txtFind.Text And txtFind.Text <> "" Then
            mstrFindValue = txtFind.Text
            Call LocateItem(mstrFindValue, True)
        Else
            Call LocateItem(mstrFindValue, False)
        End If
    End If
End Sub

Private Sub LocateItem(ByVal strInput As String, Optional ByVal blnClearSel As Boolean)
'功能：定位匹配项目
'参数：strInput=输入内容
'           blnClearSel=清空以前选中的项目
    Dim i As Long, lngStart As Long
    Dim lng编码 As Long, lng简码 As Long
    Dim lvwTmp As ListView
    Dim strTmp As String
    Dim blnFind As Long
    Dim lngOldSel As Long
    
    strInput = UCase(strInput)
     If mcsActive = CS_报表 Then
        lblFind.Caption = "查找报表(&F)"
        Set lvwTmp = lvwReport
        lng编码 = RC_编号: lng简码 = RC_简码
    Else
        lblFind.Caption = "查找分组(&F)"
        Set lvwTmp = lvwGroup
        lng编码 = GC_编号: lng简码 = GC_简码
    End If
    With lvwTmp
        If Not .SelectedItem Is Nothing And Not blnClearSel Then lngStart = .SelectedItem.Index + 1
        lngOldSel = .SelectedItem.Index
        For i = 1 To .ListItems.count
            .ListItems(i).Selected = False
        Next
        Set .SelectedItem = Nothing
        .SetFocus
        For i = IIF(lngStart = 0, 1, lngStart) To lvwTmp.ListItems.count
            
            strTmp = UCase(lvwTmp.ListItems(i).Text & "|" & .ListItems(i).SubItems(lng编码) & "|" & .ListItems(i).SubItems(lng简码))
            If strTmp Like "*" & strInput & "*" Then
                Set .SelectedItem = .ListItems(i)
                .ListItems(.SelectedItem.Index).Selected = True
                .SelectedItem.EnsureVisible
                If mcsActive = CS_报表组 Then
                    mstrPreGroup = ""
                    Call LvwGroup_ItemClick(.SelectedItem)
                End If
                blnFind = True: Exit For
            End If
        Next
        If blnFind Then
            Exit Sub
        ElseIf i >= .ListItems.count Then
            '恢复原始分组或报表，保持始终选择
            Set .SelectedItem = .ListItems(lngOldSel)
            .ListItems(.SelectedItem.Index).Selected = True
            .SelectedItem.EnsureVisible
            If mcsActive = CS_报表组 Then
                mstrPreGroup = ""
                Call LvwGroup_ItemClick(.SelectedItem)
            End If
            If lngStart <> 0 Then
                If MsgBox(" 已经定位完所有找到的信息，是否重新查找？", vbInformation + vbYesNo, App.Title) = vbYes Then
                    Call LocateItem(strInput, True)
                Else
                    txtFind.SetFocus
                End If
                Exit Sub
            Else
                MsgBox " 没有找到符合条件的信息！", vbInformation, App.Title
                txtFind.SetFocus
                Exit Sub
            End If
        End If
    End With
End Sub

Private Sub AfterItemEdit(ByVal blnAdd As Boolean, ByVal blnGroup As Boolean, ByVal lngID As Long, _
                                            ByVal str名称 As String, ByVal str编码 As String, ByVal str说明 As String)
'功能：修改，新增报表或报表组后界面处理
'参数:blnAdd=True-新增，False-修改
'       blnGroup=True-对组进行编辑，False-对报表进行编辑
'       lngID=组ID或报表ID
'       str名称=组或报表名称
'       str编码=组或报表编码
'       str说明=组或报表说明
    Dim objItem As ListItem
    If blnGroup Then
        If blnAdd Then
            Set objItem = lvwGroup.ListItems.Add(, "_" & lngID, str名称, 2, 2)
            objItem.Tag = 0
        Else
            Set objItem = lvwGroup.SelectedItem
            objItem.Text = str名称
        End If
        objItem.SubItems(GC_编号) = str编码
        objItem.SubItems(GC_说明) = str说明
        objItem.SubItems(GC_发布时间) = Format(Currentdate, "yyyy-MM-dd")
        objItem.Selected = True
        lvwGroup.SelectedItem.EnsureVisible
    Else
        If blnAdd Then
            Set objItem = lvwReport.ListItems.Add(, "_" & lngID, str名称, "Report", "Report")
            objItem.Tag = 0
        Else
            Set objItem = lvwReport.SelectedItem
            objItem.Text = str名称
        End If
        objItem.SubItems(RC_编号) = str编码
        objItem.SubItems(RC_说明) = str说明
        objItem.SubItems(RC_发布时间) = Format(Currentdate, "yyyy-MM-dd")
        objItem.Selected = True
        lvwReport.SelectedItem.EnsureVisible
    End If
End Sub

Private Sub CustomToolBarRefresh()
    Dim i As Integer
    
    For i = 1 To cbr.Bands.count
        cbr.Bands(i).Visible = False
        cbr.Bands(i).Visible = True
    Next
End Sub
