VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frm供应商管理 
   Caption         =   "供应商管理"
   ClientHeight    =   8175
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13065
   Icon            =   "frm供应商管理.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList ils32 
      Left            =   3180
      Top             =   3135
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
            Picture         =   "frm供应商管理.frx":08CA
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":0D22
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":117A
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":15CE
            Key             =   "No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":1A26
            Key             =   "Write"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3300
      Top             =   3795
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
            Picture         =   "frm供应商管理.frx":1E7E
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":22D6
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":272E
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":2B82
            Key             =   "No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":2FDA
            Key             =   "Write"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   5985
      Left            =   2805
      TabIndex        =   4
      Top             =   720
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   10557
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ColHdrIcons     =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   19
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "编码"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "简码"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "类型"
         Object.Tag             =   "类型"
         Text            =   "类型"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "许可证号"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "许可证效期"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "执照号"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "执照效期"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "地址"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "电话"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "税务登记号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "开户银行"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "帐号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "联系人"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "信用期"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "信用额"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Key             =   "站点号"
         Object.Tag             =   "站点号"
         Text            =   "院区号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   17
         Key             =   "建档时间"
         Object.Tag             =   "建档时间"
         Text            =   "建档时间"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Key             =   "撤档时间"
         Object.Tag             =   "撤档时间"
         Text            =   "撤档时间"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwList 
      Height          =   5985
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   10557
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7815
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm供应商管理.frx":3432
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17965
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
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   4800
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":3CC6
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":3EE6
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":4106
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":4322
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":4542
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":4762
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":497E
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":4B9A
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":4DB4
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":4F0E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":512A
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":534A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":5564
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":577E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":5998
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":5BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":5DCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":5FE6
            Key             =   "View"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   5520
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":6200
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":6420
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":6640
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":685C
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":6A7C
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":6C9C
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":6EB8
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":70D4
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":72EE
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":7448
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":7668
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":7888
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":7AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":7CBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":7ED6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":80F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":830A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商管理.frx":8524
            Key             =   "View"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Height          =   780
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   1376
      BandCount       =   2
      BandBorders     =   0   'False
      _CBWidth        =   13065
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbThis"
      MinHeight1      =   720
      Width1          =   11040
      NewRow1         =   0   'False
      MinHeight2      =   0
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   165
         TabIndex        =   3
         Top             =   30
         Width           =   12810
         _ExtentX        =   22595
         _ExtentY        =   1270
         ButtonWidth     =   1455
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsCold"
         HotImageList    =   "ilsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   17
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "PrintView"
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
               Key             =   "PrintSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加分类"
               Key             =   "Add"
               Description     =   "增加"
               Object.ToolTipText     =   "增加"
               Object.Tag             =   "增加"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改分类"
               Key             =   "Modify"
               Description     =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除分类"
               Key             =   "Delete"
               Description     =   "删除"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "启用"
               Key             =   "Restore"
               Description     =   "启用"
               Object.ToolTipText     =   "启用"
               Object.Tag             =   "启用"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "停用"
               Key             =   "Stop"
               Description     =   "停用"
               Object.ToolTipText     =   "停用"
               Object.Tag             =   "停用"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "StateSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "View"
               Object.ToolTipText     =   "查看方式"
               Object.Tag             =   "查看"
               ImageKey        =   "View"
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
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "filtrate"
               Description     =   "过滤"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "Refresh"
               Description     =   "刷新"
               Object.ToolTipText     =   "刷新"
               Object.Tag             =   "刷新"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助主题"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frm供应商管理.frx":873E
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   11400
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "简码"
            Top             =   210
            Width           =   1425
         End
         Begin VB.PictureBox picFind 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   10800
            ScaleHeight     =   285.714
            ScaleMode       =   0  'User
            ScaleWidth      =   495
            TabIndex        =   7
            Top             =   210
            Width           =   495
            Begin VB.Label lbl查找 
               Caption         =   "查找"
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   74
               Width           =   495
            End
         End
      End
   End
   Begin MSComctlLib.ListView lvwTemp 
      Height          =   5985
      Left            =   2805
      TabIndex        =   6
      Top             =   750
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   10557
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ColHdrIcons     =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "编码"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "简码"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "许可证号"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "许可证效期"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "执照号"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "执照效期"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "地址"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "电话"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "税务登记号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "开户银行"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "帐号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "联系人"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "信用期"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "信用额"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label lblHsc 
      Height          =   5985
      Left            =   2745
      MousePointer    =   9  'Size W E
      TabIndex        =   5
      Top             =   750
      Width           =   60
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditAddP 
         Caption         =   "增加分类(&P)"
      End
      Begin VB.Menu mnuEditUpdateP 
         Caption         =   "修改分类(&U)"
      End
      Begin VB.Menu mnuEditDeleteP 
         Caption         =   "删除分类(&D)"
      End
      Begin VB.Menu mnuEditLine0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAdd 
         Caption         =   "增加项目(&A)"
      End
      Begin VB.Menu mnuEditUpdate 
         Caption         =   "修改项目(&X)"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "删除项目(&B)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "停用(&S)"
      End
      Begin VB.Menu mnuEditRestore 
         Caption         =   "启用(&R)"
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
         End
         Begin VB.Menu mnuViewLine1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
      End
      Begin VB.Menu mnuViewLine2 
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
         Index           =   3
      End
      Begin VB.Menu mnuViewLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewHide 
         Caption         =   "显示停用项目(&H)"
      End
      Begin VB.Menu mnuViewLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFiltrate 
         Caption         =   "过滤(&I)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewASP 
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
            Caption         =   "发送反馈(&K)"
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)"
      End
   End
   Begin VB.Menu mnuFast 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuFastAdd 
         Caption         =   "增加项目(&A)"
      End
      Begin VB.Menu mnuFastModify 
         Caption         =   "修改项目(&E)"
      End
      Begin VB.Menu mnuFastDelete 
         Caption         =   "删除项目(&D)"
      End
      Begin VB.Menu mnuFastLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFastRestore 
         Caption         =   "启用(&R)"
      End
      Begin VB.Menu mnuFastStop 
         Caption         =   "停用(&S)"
      End
      Begin VB.Menu mnuFastLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFastIcon 
         Caption         =   "大图标(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuFastIcon 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuFastIcon 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuFastIcon 
         Caption         =   "详细资料(&T)"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frm供应商管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msngDownX As Single, msngDownY As Single, mSaveKey As String, mFocus As Integer, mstrFilt As String, mintFilt As Integer
Private mrstFind As New ADODB.Recordset, mFirstID As String, mLastID As String, mintColumn As Integer
Private mcllFilter As Collection
Private mblnFirst As Boolean
Private mstrPrivs As String
Dim mstr默认权限 As String
Private Declare Function SetParent Lib "user32 " (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private mrsFind As ADODB.Recordset
Private mstrFindValue As String

Private Sub Form_Activate()
    Call Form_Resize
    If mblnFirst = False Then Exit Sub
    mSaveKey = ""
    mblnFirst = False
    
    Call InitFilter
    '权限设置
    Call 权限控制
     
    '加载树数据
    Call FullType
    '加载明细数据
    If tvwList.SelectedItem Is Nothing Then
        tvwList.Nodes("Root").Selected = True
        tvwList.Nodes("Root").Expanded = True
    End If
    tvwList_NodeClick tvwList.SelectedItem
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim i As Integer
    mstrPrivs = gstrPrivs
    mblnFirst = True
    
    RestoreWinState Me, App.ProductName
    
    mnuViewIcon(lvwList.View).Checked = True
    mnuFastIcon(lvwList.View).Checked = True
    lvwList.Sorted = False
    
    Call InitFilter
    mstr默认权限 = GetDefault类型
    
    Err = 0
    On Error Resume Next
    mstrFilt = ""
    For i = 1 To Len(mstr默认权限)
        If Mid(mstr默认权限, i, 1) = 1 Then
            mstrFilt = mstrFilt & " or substr(类型," & i & ",1)=1"
        End If
    Next
    If mstrFilt <> "" Then
        mstrFilt = "  ( " & Mid(mstrFilt, 4) & " )"
    End If
   
    '2006-04-25:刘兴宏,统一增加报表发布到模块的功能
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
End Sub

Private Sub mnuEditDel_Click()
    Call mnuEditDelete_Click
End Sub

Private Sub mnuEditDeleteP_Click()
    Call mnuEditDelete_Click
End Sub

Private Sub mnuEditUpdate_Click()
    Call mnuEditModify_Click
End Sub

Private Sub mnuEditUpdateP_Click()
    Call mnuEditModify_Click
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim lng分类id As Long
    Dim lng供应商ID As Long
    'Dim byt包含停用 As Byte
    
    'byt包含停用 = IIf(mnuViewHide.Checked, 1, 0)
    If Not tvwList.SelectedItem Is Nothing Then
        lng分类id = Val(Mid(Me.tvwList.SelectedItem.Key, 2))
    End If
    
    If Not lvwList.SelectedItem Is Nothing Then
        lng供应商ID = Val(Mid(lvwList.SelectedItem.Key, 2))
    End If
    
    '2006-04-25:刘兴宏:增加自定义报表发布到模块的功能
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "分类=" & lng分类id, "供应商=" & lng供应商ID)
    
End Sub

Private Sub InitFilter()
    Set mcllFilter = New Collection
    mcllFilter.Add Array("", ""), "编码"
    mcllFilter.Add "", "名称"
    mcllFilter.Add Array("0", "0"), "信用期"
    mcllFilter.Add Array("0", "0"), "信用额"
End Sub

Private Sub FullType()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:装入供应商分类
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim TmpNode As Node
    
    gstrSQL = "" & _
        "   Select ID,上级ID,编码,名称 " & _
        "   From 供应商  " & _
        "   Where 末级 <> 1 " & _
        "   Start with 上级ID is null connect by prior ID =上级ID"
    
    Err = 0
    
    On Error GoTo ErrHand:
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    tvwList.Nodes.Clear
    Set TmpNode = tvwList.Nodes.Add(, , "Root", "所有供应商", 1, 1)
    
    TmpNode.Sorted = True
    TmpNode.Expanded = True
    TmpNode.Selected = True
    
    Do While Not rsTemp.EOF
        If IsNull(rsTemp!上级ID) Then
            Set TmpNode = tvwList.Nodes.Add("Root", 4, "K" & rsTemp!ID, "[" & rsTemp!编码 & "]" & rsTemp!名称, 5, 5)
        Else
            Set TmpNode = tvwList.Nodes.Add("K" & rsTemp!上级ID, 4, "K" & rsTemp!ID, "[" & rsTemp!编码 & "]" & rsTemp!名称, 5, 5)
        End If
        TmpNode.Sorted = True
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    SetParent txtFind.hwnd, tlbThis.hwnd
    SetParent picFind.hwnd, tlbThis.hwnd
    txtFind.Left = Me.Width - txtFind.Width
    picFind.Left = txtFind.Left - 100 - picFind.Width
    
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 5000 Then
            Me.Height = 5000
        End If
        If Me.Width < 4500 Then
            Me.Width = 4500
        End If
    End If
    If cbrThis.Bands(1).MinHeight <> tlbThis.Height Then cbrThis.Bands(1).MinHeight = tlbThis.Height
    
    cbrThis.Move 0, 0, Me.ScaleWidth
    
    If lblHsc.Left > Me.ScaleWidth - 2000 Then lblHsc.Left = Me.ScaleWidth - 2000
    
    lblHsc.Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
    lblHsc.Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - lblHsc.Top
    
    tvwList.Move 0, lblHsc.Top, lblHsc.Left, lblHsc.Height
    lvwList.Move lblHsc.Left + lblHsc.Width, lblHsc.Top, Me.ScaleWidth - (lblHsc.Left + lblHsc.Width), lblHsc.Height
    lvwTemp.Move lvwList.Left, lvwList.Top, lvwList.Width, lvwList.Height

    mnuViewToolButton.Checked = cbrThis.Visible
    mnuViewStatus.Checked = stbThis.Visible
    mnuViewToolText.Checked = tlbThis.Buttons(1).Caption <> ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lvwList.Sorted = False
    mstrFindValue = ""
    Set mrsFind = Nothing
    SaveWinState Me, App.ProductName
End Sub

Private Sub lblHsc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownX = X
End Sub

Private Sub lblHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With lblHsc
            If .Left + X - msngDownX < 2000 Then Exit Sub
            If .Left + X - msngDownX > ScaleWidth - 2000 Then Exit Sub
            .Left = .Left + X - msngDownX
        End With
        Call Form_Resize
    End If
End Sub

Private Sub lvwList_Click()
    If Me.lvwList.SelectedItem Is Nothing Then
        SetEnabled
    End If
End Sub

Private Sub lvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwList.Sorted = True
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwList.SortOrder = IIf(lvwList.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwList.SortKey = mintColumn
        lvwList.SortOrder = lvwAscending
    End If
'    lvwList.Refresh
End Sub

Private Sub lvwList_DblClick()
    Dim bln权限 As Boolean
    If lvwList.SelectedItem Is Nothing Then Exit Sub
    bln权限 = SetEditPro(Split(lvwList.SelectedItem.Tag, "|")(1))
    If Me.mnuEdit.Visible = False Or Me.mnuEditModify.Visible = False Or bln权限 = False Then
        '可查看
        frm供应商编辑.编辑单位 Me, Val(lvwList.SelectedItem.Tag), g查看, Mid(lvwList.SelectedItem.Key, 2), True
    Else
        mnuEditModify_Click
    End If
End Sub

Private Sub lvwList_GotFocus()
    mFocus = 2
    SetEnabled
End Sub

Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    SetEnabled
End Sub

Private Sub lvwList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownX = X
    msngDownY = Y
    If Button = 2 Then
        PopupMenu mnuFast
    End If
End Sub

Private Sub lvwList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call 权限控制
'    If Button = 2 Then
'        mnuFastLine1.Visible = True
'        mnuFastLine2.Visible = True
'        mnuFastStop.Visible = True
'        mnuFastRestore.Visible = True
'        mnuFastIcon(0).Visible = True
'        mnuFastIcon(1).Visible = True
'        mnuFastIcon(2).Visible = True
'        mnuFastIcon(3).Visible = True
'        Me.PopupMenu mnuFast
'    End If
End Sub

Private Sub mnuEditAdd_Click()
    Dim blnReturn As Boolean
    Dim strLstKey As String
    
    blnReturn = frm供应商编辑.编辑单位(Me, Val(Mid(tvwList.SelectedItem.Key, 2)), g新增, "", True, mstrPrivs)
    If blnReturn = False Then Exit Sub
    
    Err = 0
    On Error Resume Next
    If lvwList.SelectedItem Is Nothing Then
        strLstKey = ""
    Else
        strLstKey = lvwList.SelectedItem.Key
    End If
    '恢复选择
    If tvwList.SelectedItem Is Nothing Then
        tvwList.Nodes("Root").Selected = True
        tvwList.Nodes("Root").Expanded = True
    End If
    
    mSaveKey = ""
    tvwList_NodeClick tvwList.SelectedItem
    '恢复历史选择数据
    lvwList.ListItems(strLstKey).Selected = True
    lvwList.ListItems(strLstKey).EnsureVisible
    
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditAddP_Click()
    Dim blnReturn As Boolean
    Dim strSaveKey As String
    Dim strLstKey As String
    blnReturn = frm供应商编辑.编辑单位(Me, Val(Mid(tvwList.SelectedItem.Key, 2)), g新增, "", False, mstrPrivs)
    If blnReturn = False Then Exit Sub
    
    Err = 0
    On Error Resume Next
    strSaveKey = tvwList.SelectedItem.Key
    If lvwList.SelectedItem Is Nothing Then
        strLstKey = ""
    Else
        strLstKey = lvwList.SelectedItem.Key
    End If
    mSaveKey = ""
    '重新装入数据
    Call FullType
    '恢复选择
    tvwList.Nodes(strSaveKey).Selected = True
    If tvwList.SelectedItem Is Nothing Then
        tvwList.Nodes("Root").Selected = True
        tvwList.Nodes("Root").Expanded = True
    End If
    tvwList.SelectedItem.Expanded = True
    tvwList_NodeClick tvwList.SelectedItem
    '恢复历史选择数据
    lvwList.ListItems(strLstKey).Selected = True
    lvwList.ListItems(strLstKey).EnsureVisible
    
End Sub

Private Sub mnuEditDelete_Click()
    Dim intIndex As Long
    Dim strSQL As String
    Dim blnYes As Boolean
    Dim blnActTree As Boolean
    Dim mstrKey As String
    blnActTree = Me.ActiveControl Is tvwList
    
    If blnActTree Then
        If Me.tvwList.SelectedItem Is Nothing Then Exit Sub
        ShowMsgbox "你确认要删除分类(包含下级项目)为" & vbCrLf & "“" & Me.tvwList.SelectedItem.Text & "”的记录吗？", True, blnYes
        mstrKey = Me.tvwList.SelectedItem.Key
    Else
        If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        If SetEditPro(Split(Me.lvwList.SelectedItem.Tag, "|")(1)) = False Then Exit Sub
        
        ShowMsgbox "你确认要删除供应商为" & vbCrLf & "“" & Me.tvwList.SelectedItem.Text & "”的记录吗？", True, blnYes
        mstrKey = Me.lvwList.SelectedItem.Key
    End If
    If blnYes = False Then Exit Sub
    
    If ActiveControl Is tvwList Then
        strSQL = "zl_供应商_delete(" & Mid(tvwList.SelectedItem.Key, 2) & ")"
    Else
        strSQL = "zl_供应商_delete(" & Mid(lvwList.SelectedItem.Key, 2) & ")"
    End If
    Err = 0
    On Error GoTo errHandle:
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
    If blnActTree Then
        If tvwList.SelectedItem.Next Is Nothing Then
            If tvwList.SelectedItem.Previous Is Nothing Then
                tvwList.Nodes.Remove mstrKey
            Else
                Set tvwList.SelectedItem = tvwList.SelectedItem.Previous
                tvwList.Nodes.Remove mstrKey
            End If
        Else
            tvwList.SelectedItem.Next.Selected = True
            tvwList.Nodes.Remove mstrKey
        End If
        mSaveKey = tvwList.SelectedItem.Key
        FullList
    Else
        With lvwList
            '再删除ListView中对应节点
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
            Else
                .SetFocus
            End If
        End With
    End If
    
    SetEnabled
    Exit Sub
errHandle:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub mnuEditModify_Click()
    Dim blnReturn  As Boolean
    Dim strLstKey As String
    Dim bln末级 As Boolean
    Dim strID As String
    Dim lng上级id As Long
    
    bln末级 = Me.ActiveControl Is lvwList
    If bln末级 Then
        If lvwList.SelectedItem Is Nothing Then Exit Sub
        strID = Mid(Me.lvwList.SelectedItem.Key, 2)
        lng上级id = Val(Split(Me.lvwList.SelectedItem.Tag, "|")(0))
        
        If SetEditPro(Split(Me.lvwList.SelectedItem.Tag, "|")(1)) = False Then Exit Sub
        
        blnReturn = frm供应商编辑.编辑单位(Me, lng上级id, g修改, strID, bln末级, mstrPrivs)
        
        If blnReturn = False Then Exit Sub
        
        Err = 0
        On Error Resume Next
        If lvwList.SelectedItem Is Nothing Then
            strLstKey = ""
        Else
            strLstKey = lvwList.SelectedItem.Key
        End If
        '恢复选择
        If tvwList.SelectedItem Is Nothing Then
            tvwList.Nodes("Root").Selected = True
            tvwList.Nodes("Root").Expanded = True
        End If
        mSaveKey = ""
        tvwList_NodeClick tvwList.SelectedItem
        '恢复历史选择数据
        lvwList.ListItems(strLstKey).Selected = True
        lvwList.ListItems(strLstKey).EnsureVisible
        Err = 0
        On Error GoTo 0
        Call mnuViewRefresh_Click
        Exit Sub
    End If
    
    If tvwList.SelectedItem.Key = "Root" Then Exit Sub
    If tvwList.SelectedItem Is Nothing Then Exit Sub
    strID = Mid(tvwList.SelectedItem.Key, 2)
    blnReturn = frm供应商编辑.编辑单位(Me, Val(Mid(tvwList.SelectedItem.Parent.Key, 2)), g修改, strID, bln末级, mstrPrivs)
    If blnReturn = False Then Exit Sub
    
    Dim strSaveKey  As String
    
    Err = 0
    On Error Resume Next
    strSaveKey = tvwList.SelectedItem.Key
    If lvwList.SelectedItem Is Nothing Then
        strLstKey = ""
    Else
        strLstKey = lvwList.SelectedItem.Key
    End If
    mSaveKey = ""
    '重新装入数据
    Call FullType
    '恢复选择
    tvwList.Nodes(strSaveKey).Selected = True
    If tvwList.SelectedItem Is Nothing Then
        tvwList.Nodes("Root").Selected = True
        tvwList.Nodes("Root").Expanded = True
    End If
    tvwList_NodeClick tvwList.SelectedItem
    '恢复历史选择数据
    
    lvwList.ListItems(strLstKey).Selected = True
    lvwList.ListItems(strLstKey).EnsureVisible
    
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditRestore_Click()
    Dim strSQL As String
    
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    strSQL = "zl_供应商_reuse (" & Mid(lvwList.SelectedItem.Key, 2) & ")"
        
    On Error GoTo errHandle:
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    lvwList.SelectedItem.Icon = 2
    lvwList.SelectedItem.SmallIcon = 2
    lvwList.SelectedItem.SubItems(18) = ""

    SetEnabled
    Exit Sub
errHandle:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    Dim strSQL As String
    Dim intIndex As Integer
    
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        
    strSQL = "zl_供应商_stop(" & Mid(lvwList.SelectedItem.Key, 2) & ")"
        
    On Error GoTo errHandle:
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If mnuViewHide.Checked Then
        lvwList.SelectedItem.Icon = 3
        lvwList.SelectedItem.SmallIcon = 3
        lvwList.SelectedItem.SubItems(18) = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
        
    Else
        With lvwList
            '再删除ListView中对应节点
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
            Else
                .SetFocus
            End If
        End With
    End If
    SetEnabled
    Exit Sub
errHandle:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub mnuFastAdd_Click()
    mnuEditAdd_Click
End Sub

Private Sub mnuFastChild_Click()
    mnuEditAddP_Click
End Sub

Private Sub mnuFastDelete_Click()
    mnuEditDelete_Click
End Sub

Private Sub mnuFastIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuFastModify_Click()
    mnuEditModify_Click
End Sub

Private Sub mnuFastRestore_Click()
    mnuEditRestore_Click
End Sub

Private Sub mnuFastStop_Click()
    mnuEditStop_Click
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub mnuFilePrintView_Click()
    subPrint 2
End Sub

Private Sub mnuViewFiltrate_Click()
    Dim blnCancel As Boolean
    Dim strFilter As String
    Dim cllFilter As Collection
    
    Dim intOldCondition As Integer
    intOldCondition = mintFilt
    Call frm供应商过滤.GetFiler(Me, blnCancel, mstrFilt, cllFilter, mstrPrivs)
    If blnCancel = True Then Exit Sub
    If mstrFilt <> "" Then
        mstrFilt = "(" & mstrFilt & ")"
    End If
    Set mcllFilter = cllFilter
    FullList
End Sub

Private Sub mnuViewFind_Click()
    '查找功能
    Dim strSQL As String
    Dim strOthers() As String

    strSQL = frm供应商定位.getSql(strOthers)
    If strSQL = "" Then Exit Sub
    strSQL = strSQL & IIf(mnuViewHide.Checked, "", " And (撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))") & IIf(mstrFilt <> "", " And " & mstrFilt, "")
    Set mrstFind = New ADODB.Recordset
'    zlDatabase.OpenRecordset mrstFind, strSql, Me.Caption
    On Error GoTo errHandle
    Set mrstFind = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(mcllFilter("编码")(0)), CStr(mcllFilter("编码")(1)), _
                            CStr(mcllFilter("名称")), CLng(mcllFilter("信用期")(0)), CLng(mcllFilter("信用期")(1)), _
                            CDbl(mcllFilter("信用额")(0)), CDbl(mcllFilter("信用额")(1)), strOthers(0), strOthers(1), strOthers(2))


    mrstFind.Sort = "上级ID,名称,ID"
    If mrstFind.EOF Then
        MsgBox "没有满足定位条件的数据！", vbInformation, Me.Caption
        Exit Sub
    End If
    mrstFind.MoveFirst
    mFirstID = mrstFind("ID")
    mrstFind.MoveLast
    mLastID = mrstFind("ID")
    Unload frm供应商定位
    frmToolBarWin.ShowBar "供应商定位", Me
    subFirst
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Find供应商()
    Dim strKeytvw As String, strKeylvw As String, blnUP As Boolean, blnDown As Boolean
    Dim rstTemp As New ADODB.Recordset
    If mrstFind.EOF Then Exit Sub
    frmToolBarWin.屏蔽 0, False
    frmToolBarWin.屏蔽 1, False
    On Error GoTo errHandle
    If Not IsNull(mrstFind("上级ID")) Then
        'by lesfeng 2009-12-2 性能优化
        Set rstTemp = zlDatabase.OpenSQLRecord("Select id From 供应商 Where ID=[1]", Me.Caption, Val(mrstFind!上级ID))
        If Not rstTemp.EOF Then
            strKeytvw = "K" & rstTemp!ID
        End If
        rstTemp.Close
    Else
        strKeytvw = "Root"
    End If
    strKeylvw = "K" & mrstFind("ID")
    If strKeytvw = mSaveKey Then
        Set lvwList.SelectedItem = lvwList.ListItems(strKeylvw)
    Else
        Set tvwList.SelectedItem = tvwList.Nodes(strKeytvw)
        FullList
        Set lvwList.SelectedItem = lvwList.ListItems(strKeylvw)
    End If
    blnUP = (mrstFind("ID") <> mFirstID)
    blnDown = (mrstFind("ID") <> mLastID)
    frmToolBarWin.屏蔽 0, blnUP
    frmToolBarWin.屏蔽 1, blnDown
    tvwList.SelectedItem.EnsureVisible
    lvwList.SelectedItem.EnsureVisible
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub subFirst()
    mrstFind.MoveFirst
    Find供应商
End Sub

Public Sub subPrevious()
    mrstFind.MovePrevious
    Find供应商
End Sub

Public Sub subNext()
    mrstFind.MoveNext
    Find供应商
End Sub

Public Sub subLast()
    mrstFind.MoveLast
    Find供应商
End Sub

Private Sub mnuViewHide_Click()
    mnuViewHide.Checked = Not mnuViewHide.Checked
    FullList
    Set mrsFind = Nothing
    mstrFindValue = ""
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim intTemp As Integer
    For intTemp = 0 To 3
        mnuViewIcon(intTemp).Checked = False
        mnuFastIcon(intTemp).Checked = False
    Next
    
    mnuViewIcon(Index).Checked = True
    mnuFastIcon(Index).Checked = True
    lvwList.View = Index
    lvwList.Refresh
End Sub

Private Sub mnuViewRefresh_Click()
    Dim strKey As String
    mSaveKey = ""
    If Me.tvwList.SelectedItem Is Nothing Then
        strKey = "Root"
    Else
        strKey = Me.tvwList.SelectedItem.Key
    End If
    Call FullType
    Err = 0
    On Error Resume Next
    tvwList.Nodes(strKey).Selected = True
    If tvwList.SelectedItem Is Nothing Then
        tvwList.Nodes("Root").Selected = True
        tvwList.Nodes("Root").Expanded = True
    Else
        tvwList.SelectedItem.Expanded = True
    End If
    tvwList_NodeClick tvwList.SelectedItem
    
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    cbrThis.Bands(1).MinHeight = tlbThis.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In tlbThis.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    cbrThis.Bands(1).MinHeight = tlbThis.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
       ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hwnd)
End Sub



Private Sub tlbthis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lvwTemp As ListView
    Select Case Button.Key
        Case "filtrate"
            mnuViewFiltrate_Click
        Case "Add"
            If Me.ActiveControl Is tvwList Then
                mnuEditAddP_Click
            Else
                mnuEditAdd_Click
            End If
        Case "Modify"
            mnuEditModify_Click
        Case "Print"
            mnuFilePrint_Click
        Case "PrintView"
            mnuFilePrintView_Click
        Case "Find"
            mnuViewFind_Click
        Case "Refresh"
            mnuViewRefresh_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Restore"
            mnuEditRestore_Click
        Case "Stop"
            mnuEditStop_Click
        Case "View"
            Set lvwTemp = lvwList
            mnuViewIcon(lvwTemp.View).Checked = False
            mnuFastIcon(lvwTemp.View).Checked = False
            If lvwTemp.View = 3 Then
                mnuViewIcon(0).Checked = True
                mnuFastIcon(0).Checked = True
                lvwTemp.View = 0
            Else
                mnuViewIcon(lvwTemp.View + 1).Checked = True
                mnuFastIcon(lvwTemp.View + 1).Checked = True
                lvwTemp.View = lvwTemp.View + 1
            End If
        Case "Exit"
            mnuFileExit_Click
    End Select
End Sub

Private Sub tlbThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Call mnuViewIcon_Click(ButtonMenu.Index - 1)
End Sub

Private Sub tlbthis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu mnuViewTool
End Sub

Private Sub tvwList_Collapse(ByVal Node As MSComctlLib.Node)
    If Node.Parent Is Nothing Then
        Node.Expanded = True
        Exit Sub
    End If
    If InStr(tvwList.SelectedItem.Key, Node.Key) > 0 Then
        Set tvwList.SelectedItem = Node
        tvwList_NodeClick Node
    End If
End Sub

Private Sub tvwList_GotFocus()
    mFocus = 1
    SetEnabled
End Sub

Private Sub tvwList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call 权限控制
    If mnuEdit.Visible = False Then Exit Sub
    If Button = 2 Then
        PopupMenu mnuEdit
    End If
End Sub

Private Sub tvwList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    DoEvents
'    If Button = 2 Then
'        mnuFastLine1.Visible = False
'        mnuFastLine2.Visible = False
'        mnuFastStop.Visible = False
'        mnuFastRestore.Visible = False
'        mnuFastIcon(0).Visible = False
'        mnuFastIcon(1).Visible = False
'        mnuFastIcon(2).Visible = False
'        mnuFastIcon(3).Visible = False
'        Me.PopupMenu mnuFast
'    End If
End Sub

Private Sub tvwList_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = mSaveKey Then Exit Sub
    Call FullList
End Sub

Public Sub FullList(Optional strCon As String = "")
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:加载明细数据
    '--入参数:strCon -条件
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lstItem As ListItem, strTempKey As String
    Dim strWhere As String
    
    Dim strTvwKey As String
    
    strTvwKey = tvwList.SelectedItem.Key
    
    strWhere = ""
    If strCon <> "" Then
        strWhere = " and (" & strCon & ") "
    End If
    'by lesfeng 2009-12-2 性能优化
    If mnuViewHide.Checked = False Then
        If strTvwKey = "Root" Then
            strWhere = strWhere & "  and (to_char(撤档时间,'yyyy-MM-DD') = '3000-01-01' or 撤档时间 is null)"
        Else
            strWhere = strWhere & "  and (to_char(撤档时间,'yyyy-MM-DD') = '3000-01-01' or 撤档时间 is null)" & " start with  上级ID = [8] connect by prior id=上级id  "
        End If
    Else
        If strTvwKey <> "Root" Then
            strWhere = strWhere & " start with  上级ID = [8] connect by prior id=上级id  "
        End If
    End If
    Err = 0
    On Error GoTo ErrHand:
    'by lesfeng 2009-12-2 性能优化
    gstrSQL = "" & _
        "   Select ID,上级ID,编码,名称,简码,末级,许可证号,许可证效期,执照号,执照效期,税务登记号,地址,电话,开户银行," & _
        "          帐号,联系人,建档时间,撤档时间,类型,信用期,信用额,销售委托人,销售委托日期,质量认证号,质量认证日期," & _
        "          药监局备案号,药监局备案日期,授权号,授权期,站点" & _
        "    from 供应商  where 末级=1  " & IIf(mstrFilt = "", "", " And " & mstrFilt) & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(mcllFilter("编码")(0)), CStr(mcllFilter("编码")(1)), _
                            CStr(mcllFilter("名称")), CLng(mcllFilter("信用期")(0)), CLng(mcllFilter("信用期")(1)), _
                            CDbl(mcllFilter("信用额")(0)), CDbl(mcllFilter("信用额")(1)), Val(Mid(strTvwKey, 2)))
    
    Dim strTmp As String
    Dim i As Integer
    Dim str类型 As String
    
    If lvwList.SelectedItem Is Nothing Then
        strTempKey = ""
    Else
        strTempKey = lvwList.SelectedItem.Key
    End If
    lvwList.ListItems.Clear
    
    With rsTemp
        Do While Not .EOF
            If Format(!撤档时间, "yyyy-mm-dd") = "3000-01-01" Or IsNull(!撤档时间) Then
                Set lstItem = lvwList.ListItems.Add(, "K" & rsTemp("ID"), rsTemp("名称"), 2, 2)
            Else
                Set lstItem = lvwList.ListItems.Add(, "K" & rsTemp("ID"), rsTemp("名称"), 3, 3)
            End If
            lstItem.Tag = Nvl(!上级ID, 0) & "|" & Nvl(!类型)
            
            strTmp = Nvl(!类型) 'Right(Dec2Bin(Nvl(!类型, 0)), 4)
            str类型 = ""
            For i = 1 To Len(strTmp)
                If Mid(strTmp, i, 1) = 1 Then
                    str类型 = str类型 & "," & Switch(i = 1, "药品", i = 2, "物资", i = 3, "设备", i = 4, "其它", i = 5, "卫生材料")
                End If
            Next
            If str类型 <> "" Then
                str类型 = Mid(str类型, 2)
            End If
            
            lstItem.ListSubItems.Add , , Nvl(!编码)
            lstItem.ListSubItems.Add , , Nvl(!简码)
            lstItem.ListSubItems.Add , , str类型
            
            lstItem.ListSubItems.Add , , Nvl(!许可证号)
            lstItem.ListSubItems.Add , , Nvl(!许可证效期)
            lstItem.ListSubItems.Add , , Nvl(!执照号)
            lstItem.ListSubItems.Add , , Nvl(!执照效期)
            lstItem.ListSubItems.Add , , Nvl(!地址)
            lstItem.ListSubItems.Add , , Nvl(!电话)
            lstItem.ListSubItems.Add , , Nvl(!税务登记号)
            lstItem.ListSubItems.Add , , Nvl(!开户银行)
            lstItem.ListSubItems.Add , , Nvl(!帐号)
            lstItem.ListSubItems.Add , , Nvl(!联系人)
            lstItem.ListSubItems.Add , , Nvl(!信用期, " ")
            lstItem.ListSubItems.Add , , Nvl(!信用额, " ")
            lstItem.ListSubItems.Add , , Nvl(!站点, " ")
            lstItem.ListSubItems.Add , , Format(!建档时间, "yyyy-mm-dd")
            If Format(!撤档时间, "yyyy-mm-dd") = "3000-01-01" Or IsNull(!撤档时间) Then
                lstItem.ListSubItems.Add , , " "
            Else
                lstItem.ListSubItems.Add , , Format(!撤档时间, "yyyy-mm-dd")
            End If
            rsTemp.MoveNext
        Loop
    End With
    mSaveKey = tvwList.SelectedItem.Key
    
    If lvwList.ListItems.Count > 0 Then
        If strTempKey <> "" Then
            On Error Resume Next
            Set lvwList.SelectedItem = lvwList.ListItems(strTempKey)
        End If
        If lvwList.SelectedItem Is Nothing Then
            Set lvwList.SelectedItem = lvwList.ListItems(1)
        End If
        Err = 0
        On Error GoTo 0
    End If
    SetEnabled
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetEnabled()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:设置壮态
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim blnData As Boolean '存在数据
    Dim blnActTree As Boolean   '当前激活控件为树
    Dim blnRoot As Boolean      '是否选择的根
    Dim blnItmSel As Boolean    '是否选中
    Dim blnStop As Boolean      '停用部份
    Dim blnChild As Boolean     '有只子目录
    Dim str类型 As String
    Dim bln权限 As Boolean
    blnActTree = Me.ActiveControl Is tvwList
    blnData = Me.lvwList.ListItems.Count <> 0
    blnRoot = Me.tvwList.SelectedItem.Key = "Root"
    bln权限 = False
    If Not Me.lvwList.SelectedItem Is Nothing Then
        str类型 = Split(lvwList.SelectedItem.Tag, "|")(1)
        bln权限 = SetEditPro(str类型)
        blnStop = Trim(lvwList.SelectedItem.SubItems(18)) <> "" And bln权限
        blnItmSel = bln权限
    Else
        blnItmSel = False
    End If
    
    If blnRoot Then
        mnuEditAddP.Enabled = (blnRoot) And (blnActTree)
    Else
        mnuEditAddP.Enabled = (blnActTree)
    End If
    mnuEditUpdateP.Enabled = (Not blnRoot) And (blnActTree)
    mnuEditDeleteP.Enabled = (Not blnRoot) And (blnActTree)
    mnuEditAdd.Enabled = Not blnActTree
    mnuEditModify.Enabled = Not blnActTree
    mnuEditUpdate.Enabled = Not blnActTree
    mnuEditDelete.Enabled = Not blnActTree
    mnuEditDel.Enabled = Not blnActTree
    mnuEditStop.Enabled = (Not blnActTree) And blnItmSel And (Not blnStop)
    mnuEditRestore.Enabled = (Not blnActTree) And blnItmSel And blnStop
    
    mnuFastAdd.Enabled = mnuEditAdd.Enabled
    mnuFastDelete.Enabled = mnuEditDelete.Enabled
    mnuFastModify.Enabled = mnuEditModify.Enabled
    mnuFastStop.Enabled = mnuEditStop.Enabled
    mnuFastRestore.Enabled = mnuEditRestore.Enabled
    
    If blnActTree Then
        tlbThis.Buttons("Add").Enabled = mnuEditAddP.Enabled: tlbThis.Buttons("Add").Caption = "增加分类"
        tlbThis.Buttons("Modify").Enabled = mnuEditUpdateP.Enabled: tlbThis.Buttons("Modify").Caption = "修改分类"
        tlbThis.Buttons("Delete").Enabled = mnuEditDeleteP.Enabled: tlbThis.Buttons("Delete").Caption = "删除分类"
    Else
        tlbThis.Buttons("Add").Enabled = mnuEditAdd.Enabled: tlbThis.Buttons("Add").Caption = "增加项目"
        tlbThis.Buttons("Modify").Enabled = mnuEditModify.Enabled: tlbThis.Buttons("Modify").Caption = "修改项目"
        tlbThis.Buttons("Delete").Enabled = mnuEditDelete.Enabled: tlbThis.Buttons("Delete").Caption = "删除项目"
    End If
'    tlbThis.Buttons("Add").Visible = False
'    tlbThis.Buttons("Modify").Visible = False
'    tlbThis.Buttons("Delete").Visible = False
    tlbThis.Buttons("PrintSeparate").Visible = True
    
    tlbThis.Buttons("Restore").Enabled = mnuEditRestore.Enabled
    tlbThis.Buttons("Stop").Enabled = mnuEditStop.Enabled
    
    
    mnuFilePrint.Enabled = blnData
    mnuFilePrintView.Enabled = blnData
    mnuFileExcel.Enabled = blnData
        
    tlbThis.Buttons("Print").Enabled = blnData
    tlbThis.Buttons("PrintView").Enabled = blnData
End Sub

Private Sub subPrint(ByVal bytMode As Byte)
    Dim objPrint As New zlPrintLvw
    objPrint.Title.Text = "供应商列表"
    Set objPrint.Body.objData = lvwList
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")

    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
        End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Private Sub 权限控制()
    Dim blnAdd As Boolean
    Dim blnModify As Boolean
    Dim blnDelete As Boolean
    Dim blnStart As Boolean
    Dim blnStop As Boolean
    blnAdd = InStr(mstrPrivs, ";增加;") <> 0
    blnModify = InStr(mstrPrivs, ";修改;") <> 0
    blnDelete = InStr(mstrPrivs, ";删除;") <> 0
    blnStart = InStr(mstrPrivs, ";启用;") <> 0
    blnStop = InStr(mstrPrivs, ";停用;") <> 0
    
    If blnAdd = False And blnModify = False And blnStart = False And blnDelete = False And blnStop = False Then
        mnuEdit.Visible = False
        '快捷菜单
        mnuFastAdd.Visible = False
        mnuFastModify.Visible = False
        mnuFastDelete.Visible = False
        mnuFastLine1.Visible = False
        mnuFastRestore.Visible = False
        mnuFastStop.Visible = False
        mnuFastLine2.Visible = False
    Else
        mnuEditAdd.Visible = blnAdd: mnuEditAddP.Visible = blnAdd   '增加
        mnuEditUpdate.Visible = blnModify: mnuEditUpdateP.Visible = blnModify   '修改
        mnuEditDel.Visible = blnDelete: mnuEditDeleteP.Visible = blnDelete   '删除
        mnuEditLine1.Visible = (blnAdd Or blnModify Or blnDelete) And (blnStop Or blnStart)
        mnuEditRestore.Visible = blnStart
        mnuEditStop.Visible = blnStop
        '快捷菜单
        mnuFastAdd.Visible = blnAdd
        mnuFastModify.Visible = blnModify
        mnuFastDelete.Visible = blnDelete
        If blnAdd = False And blnModify = False And blnDelete = False Then
            mnuFastLine1.Visible = False
        End If
        mnuFastRestore.Visible = blnStart
        mnuFastStop.Visible = blnStop
        If blnStart = False And blnStop = False Then
            mnuFastLine2.Visible = False
        End If
    End If
    tlbThis.Buttons("Add").Visible = blnAdd
    tlbThis.Buttons("Modify").Visible = blnModify
    tlbThis.Buttons("Delete").Visible = blnDelete
    tlbThis.Buttons("EditSeparate").Visible = blnAdd Or blnModify Or blnDelete
    tlbThis.Buttons("StateSeparate").Visible = blnStart Or blnStop
    tlbThis.Buttons("Restore").Visible = blnStart
    tlbThis.Buttons("Stop").Visible = blnStop
    
    mnuEditModify.Visible = False
    mnuEditDelete.Visible = False
End Sub
Private Function SetEditPro(ByVal str类型 As String) As Boolean
    '设置编辑权限
    
    Dim bln药品 As Boolean
    Dim bln物资 As Boolean
    Dim bln设备 As Boolean
    Dim bln其他 As Boolean
    Dim bln卫材 As Boolean
    
    bln药品 = InStr(1, mstrPrivs, "药品供应商") <> 0
    bln物资 = InStr(1, mstrPrivs, "物资供应商") <> 0
    bln设备 = InStr(1, mstrPrivs, "设备供应商") <> 0
    bln其他 = InStr(1, mstrPrivs, "其他供应商") <> 0
    bln卫材 = InStr(1, mstrPrivs, "卫材供应商") <> 0
    
    Err = 0: On Error GoTo ErrHand:
    
    SetEditPro = False
    If bln药品 = False And bln物资 = False And bln设备 = False And bln其他 = False And bln卫材 = False Then
            Exit Function
    End If
    If Mid(str类型, 1, 1) = 1 Then
        If Not bln药品 Then
            Exit Function
        End If
    End If
    
    If Mid(str类型, 2, 1) = 1 Then
        If Not bln物资 Then
            Exit Function
        End If
    End If
    If Mid(str类型, 3, 1) = 1 Then
        If Not bln设备 Then
            Exit Function
        End If
    End If
    
    If Mid(str类型, 4, 1) = 1 Then
        If Not bln其他 Then
            Exit Function
        End If
    End If
    If Mid(str类型, 5, 1) = 1 Then
        If Not bln卫材 Then
            Exit Function
        End If
    End If
    
    
    SetEditPro = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Private Function GetDefault类型() As String
    '设置编辑权限
    
    Dim bln药品 As Boolean
    Dim bln物资 As Boolean
    Dim bln设备 As Boolean
    Dim bln其他 As Boolean
    Dim bln卫材 As Boolean
    Dim strTemp As String
    
    bln药品 = InStr(1, mstrPrivs, "药品供应商") <> 0
    bln物资 = InStr(1, mstrPrivs, "物资供应商") <> 0
    bln设备 = InStr(1, mstrPrivs, "设备供应商") <> 0
    bln其他 = InStr(1, mstrPrivs, "其他供应商") <> 0
    bln卫材 = InStr(1, mstrPrivs, "卫材供应商") <> 0
    
    strTemp = ""
    strTemp = strTemp & IIf(bln药品, "1", "0")
    strTemp = strTemp & IIf(bln物资, "1", "0")
    strTemp = strTemp & IIf(bln设备, "1", "0")
    strTemp = strTemp & IIf(bln其他, "1", "0")
    strTemp = strTemp & IIf(bln卫材, "1", "0")
    GetDefault类型 = strTemp
    
End Function


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
    zlCommFun.OpenIme True
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    
    If KeyAscii = vbKeyReturn Then
        If txtFind.Text = "" Then Exit Sub
        On Error GoTo errHandle
        If mstrFindValue <> txtFind.Text And txtFind.Text <> "" Then
            mstrFindValue = txtFind.Text
            Set mrsFind = Nothing
            strTemp = " and (撤档时间 = to_date('3000-01-01','YYYY-MM-DD') or 撤档时间 is null ) "
            gstrSQL = "select id,上级id from 供应商 where 编码 like [1] or 名称 like [1] or 简码 like [1] and 末级=1"
            
            If mnuViewHide.Checked = False Then
                gstrSQL = gstrSQL & strTemp
            End If
            Set mrsFind = zlDatabase.OpenSQLRecord(gstrSQL, "供应商查询", UCase(txtFind.Text) & "%")
            Call LocateItem
        Else
            If Not mrsFind.EOF Then
                mrsFind.MoveNext
                Call LocateItem
            ElseIf mrsFind.RecordCount <> 0 And mrsFind.EOF Then
                mrsFind.MoveFirst
                Call LocateItem
            End If
        End If
    End If
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub LocateItem()
    Dim strTemp As String
    
    txtFind.SetFocus
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
    If mrsFind.RecordCount = 0 Then
        MsgBox " 没有找到符合条件的信息！", vbInformation, gstrSysName
        txtFind.SetFocus
        txtFind.SelStart = 0
        txtFind.SelLength = Len(txtFind.Text)
        Exit Sub
    End If
    If mrsFind.EOF = True Then
        MsgBox " 已经定位完所有找到的信息，请重新输入条件！", vbInformation, gstrSysName
        txtFind.SetFocus
        txtFind.SelStart = 0
        txtFind.SelLength = Len(txtFind.Text)
        Exit Sub
    End If
    
    With tvwList
        If IsNull(mrsFind("上级ID")) = False Then
            .Nodes("K" & mrsFind("上级ID")).Selected = True
        Else
            .Nodes("Root").Selected = True
        End If
        .SelectedItem.EnsureVisible
    End With
        
    With lvwList
        .ListItems("K" & mrsFind("id")).Selected = True
        .SelectedItem.EnsureVisible
    End With
End Sub
